"""
main.py - Backend del SaaS de verificación de emails (Fase 1, 100% gratis)

Stack:
    - FastAPI (backend/API)
    - Supabase (Postgres gratis) para usuarios/créditos/historial
    - Verificación: sintaxis + MX + toxicidad (sin SMTP real, ver verifier.py)

Variables de entorno necesarias (ver .env.example):
    SUPABASE_URL, SUPABASE_KEY, FRONTEND_ORIGIN
"""

import os
import io
import csv
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import openpyxl

from verifier import verify_email
from supabase import create_client, Client

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")
FRONTEND_ORIGIN = os.getenv("FRONTEND_ORIGIN", "*")
FREE_SIGNUP_CREDITS = int(os.getenv("FREE_SIGNUP_CREDITS", "20"))
MAX_EMAILS_PER_JOB = int(os.getenv("MAX_EMAILS_PER_JOB", "300"))  # límite por request (evita timeouts en tier gratis)
WORKERS = int(os.getenv("VERIFY_WORKERS", "20"))

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None

app = FastAPI(title="Email Verifier SaaS - API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[FRONTEND_ORIGIN] if FRONTEND_ORIGIN != "*" else ["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def db():
    if supabase is None:
        raise HTTPException(500, "Backend no configurado: faltan SUPABASE_URL / SUPABASE_KEY")
    return supabase


# ---------------------------------------------------------------------------
# Usuarios / créditos
# ---------------------------------------------------------------------------

@app.post("/api/user/login")
def login(email: str = Form(...)):
    """Crea el usuario si no existe (con créditos gratis) y devuelve su saldo."""
    email = email.strip().lower()
    if not email or "@" not in email:
        raise HTTPException(400, "Email inválido")

    existing = db().table("users").select("*").eq("email", email).execute()
    if existing.data:
        return existing.data[0]

    new_user = {"email": email, "credits": FREE_SIGNUP_CREDITS}
    result = db().table("users").insert(new_user).execute()
    return result.data[0]


@app.get("/api/user/{email}")
def get_user(email: str):
    email = email.strip().lower()
    res = db().table("users").select("*").eq("email", email).execute()
    if not res.data:
        raise HTTPException(404, "Usuario no encontrado")
    return res.data[0]


@app.post("/api/credits/add")
def add_credits(email: str = Form(...), amount: int = Form(...)):
    """
    PLACEHOLDER de compra de créditos para pruebas de Fase 1.
    En Fase 2 esto se reemplaza por un webhook de Stripe que llama a esta
    misma lógica tras confirmar el pago (no lo dejes público/sin proteger
    cuando conectes pagos reales).
    """
    email = email.strip().lower()
    user_res = db().table("users").select("*").eq("email", email).execute()
    if not user_res.data:
        raise HTTPException(404, "Usuario no encontrado")
    user = user_res.data[0]
    new_balance = user["credits"] + amount
    db().table("users").update({"credits": new_balance}).eq("email", email).execute()
    return {"email": email, "credits": new_balance}


# ---------------------------------------------------------------------------
# Verificación
# ---------------------------------------------------------------------------

def extract_emails_from_upload(filename: str, content: bytes):
    emails = []
    if filename.lower().endswith(".csv"):
        text = None
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                text = content.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            raise HTTPException(400, "No se pudo leer la codificación del CSV")
        reader = csv.reader(io.StringIO(text))
        for row in reader:
            for cell in row:
                cell = cell.strip()
                if "@" in cell and "." in cell.split("@")[-1]:
                    emails.append(cell)
                    break
    elif filename.lower().endswith((".xlsx", ".xls")):
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is None:
                        continue
                    cell = str(cell).strip()
                    if "@" in cell and "." in cell.split("@")[-1]:
                        emails.append(cell)
                        break
    else:
        raise HTTPException(400, "Formato no soportado. Sube un .csv o .xlsx")

    seen = set()
    unique = []
    for e in emails:
        key = e.lower()
        if key not in seen:
            seen.add(key)
            unique.append(e)
    return unique


@app.post("/api/verify")
def verify_batch(email: str = Form(...), file: UploadFile = File(...)):
    user_email = email.strip().lower()
    user_res = db().table("users").select("*").eq("email", user_email).execute()
    if not user_res.data:
        raise HTTPException(404, "Usuario no encontrado. Inicia sesión primero.")
    user = user_res.data[0]

    content = file.file.read()
    emails = extract_emails_from_upload(file.filename, content)

    if not emails:
        raise HTTPException(400, "No se encontraron direcciones de email válidas en el archivo")

    if len(emails) > MAX_EMAILS_PER_JOB:
        raise HTTPException(
            400,
            f"El archivo tiene {len(emails)} emails. El límite por lote en esta fase es "
            f"{MAX_EMAILS_PER_JOB} (para evitar timeouts en hosting gratuito). Divide el archivo."
        )

    if user["credits"] < len(emails):
        raise HTTPException(
            402,
            f"Créditos insuficientes. Necesitas {len(emails)}, tienes {user['credits']}."
        )

    results = []
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(verify_email, e): e for e in emails}
        for future in as_completed(futures):
            results.append(future.result())

    new_balance = user["credits"] - len(emails)
    db().table("users").update({"credits": new_balance}).eq("email", user_email).execute()

    job = {
        "id": str(uuid.uuid4()),
        "user_email": user_email,
        "filename": file.filename,
        "total_emails": len(emails),
        "results": results,
        "source": "web",
    }
    db().table("jobs").insert(job).execute()

    return {"credits_restantes": new_balance, "total": len(results), "resultados": results}


@app.post("/api/jobs/submit")
def submit_local_job(
    email: str = Form(...),
    filename: str = Form(...),
    total_emails: int = Form(...),
    status_counts_json: str = Form(...),
):
    """
    Usado por la app de escritorio (Windows). La verificación SMTP real
    ocurre EN LA MÁQUINA DEL USUARIO — este endpoint nunca recibe ni
    almacena los emails ni sus resultados individuales, solo un conteo
    agregado por estado (ej. {"Accepted": 40, "Rejected": 5}), para poder
    descontar créditos y mostrar un resumen en el historial.

    El detalle completo (qué email dio qué resultado) se queda únicamente
    en el CSV que la app exporta localmente en la computadora del usuario.
    """
    import json as _json
    user_email = email.strip().lower()

    user_res = db().table("users").select("*").eq("email", user_email).execute()
    if not user_res.data:
        raise HTTPException(404, "Usuario no encontrado. Inicia sesión primero.")
    user = user_res.data[0]

    try:
        status_counts = _json.loads(status_counts_json)
    except Exception:
        raise HTTPException(400, "status_counts_json inválido")

    if not isinstance(status_counts, dict) or total_emails <= 0:
        raise HTTPException(400, "Datos de resumen inválidos")

    if user["credits"] < total_emails:
        raise HTTPException(
            402, f"Créditos insuficientes. Necesitas {total_emails}, tienes {user['credits']}."
        )

    new_balance = user["credits"] - total_emails
    db().table("users").update({"credits": new_balance}).eq("email", user_email).execute()

    job = {
        "id": str(uuid.uuid4()),
        "user_email": user_email,
        "filename": filename,
        "total_emails": total_emails,
        "results": None,
        "status_counts": status_counts,
        "source": "desktop",
    }
    db().table("jobs").insert(job).execute()

    return {"credits_restantes": new_balance, "total": total_emails}


@app.get("/api/jobs/{email}")
def list_jobs(email: str):
    email = email.strip().lower()
    res = (
        db().table("jobs").select("id,filename,total_emails,status_counts,source,created_at")
        .eq("user_email", email).order("created_at", desc=True).execute()
    )
    return res.data


@app.get("/api/jobs/{email}/{job_id}/download")
def download_job(email: str, job_id: str):
    res = db().table("jobs").select("*").eq("id", job_id).eq("user_email", email.lower()).execute()
    if not res.data:
        raise HTTPException(404, "Job no encontrado")
    job = res.data[0]

    if not job.get("results"):
        raise HTTPException(
            409,
            "Esta verificación se hizo localmente en la app de escritorio (SMTP real) y, por "
            "diseño, el detalle nunca se envía ni se guarda en el servidor. El CSV completo con "
            "cada email y su resultado quedó exportado en tu computadora en el momento en que "
            "terminó la verificación."
        )

    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(["email", "status", "detalle", "toxicidad", "señales_toxicidad"])
    for r in job["results"]:
        writer.writerow([r["email"], r["status"], r["detalle"], r["toxicidad"], r["señales_toxicidad"]])
    buf.seek(0)

    return StreamingResponse(
        iter([buf.getvalue()]),
        media_type="text/csv",
        headers={"Content-Disposition": f"attachment; filename={job['filename']}_resultados.csv"},
    )


@app.get("/api/health")
def health():
    return {"status": "ok", "supabase_configured": supabase is not None}
