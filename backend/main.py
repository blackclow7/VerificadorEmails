"""
main.py - Backend SaaS verificador de emails
Con autenticación real via Supabase Auth (email + contraseña + JWT)
"""

import os
import io
import csv
import uuid
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import openpyxl

from verifier import verify_email
from supabase import create_client, Client

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

SUPABASE_URL        = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY        = os.getenv("SUPABASE_KEY", "")
FRONTEND_ORIGIN     = os.getenv("FRONTEND_ORIGIN", "*")
FREE_SIGNUP_CREDITS = int(os.getenv("FREE_SIGNUP_CREDITS", "20"))
MAX_EMAILS_PER_JOB  = int(os.getenv("MAX_EMAILS_PER_JOB", "300"))
WORKERS             = int(os.getenv("VERIFY_WORKERS", "20"))

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
        raise HTTPException(500, "Backend no configurado")
    return supabase


# ---------------------------------------------------------------------------
# Autenticación — helpers
# ---------------------------------------------------------------------------

def get_user_from_token(authorization: str) -> dict:
    """
    Verifica el JWT de Supabase Auth que viene en el header Authorization.
    Devuelve el usuario de la tabla 'users' si el token es válido.
    """
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(401, "Token de autenticación requerido")
    token = authorization.split(" ", 1)[1]
    try:
        # Supabase valida el JWT y devuelve el usuario autenticado
        resp = db().auth.get_user(token)
        if not resp or not resp.user:
            raise HTTPException(401, "Token inválido o expirado")
        email = resp.user.email.lower()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(401, f"Error de autenticación: {e}")

    # Buscar en nuestra tabla de usuarios (créditos, etc.)
    user_res = db().table("users").select("*").eq("email", email).execute()
    if not user_res.data:
        raise HTTPException(404, "Usuario no encontrado en el sistema")
    return user_res.data[0]


# ---------------------------------------------------------------------------
# Auth endpoints
# ---------------------------------------------------------------------------

@app.post("/api/auth/register")
def register(email: str = Form(...), password: str = Form(...)):
    """Registra un usuario nuevo. Supabase manda email de verificación."""
    email = email.strip().lower()
    if not email or "@" not in email:
        raise HTTPException(400, "Email inválido")
    if len(password) < 8:
        raise HTTPException(400, "La contraseña debe tener al menos 8 caracteres")
    try:
        resp = db().auth.sign_up({"email": email, "password": password})
        if not resp.user:
            raise HTTPException(400, "No se pudo crear la cuenta")
    except HTTPException:
        raise
    except Exception as e:
        msg = str(e).lower()
        if "already" in msg or "registered" in msg:
            raise HTTPException(409, "Este email ya está registrado")
        raise HTTPException(400, f"Error al registrar: {e}")

    # Crear registro en nuestra tabla de usuarios con créditos gratis
    existing = db().table("users").select("id").eq("email", email).execute()
    if not existing.data:
        db().table("users").insert({
            "email": email,
            "credits": FREE_SIGNUP_CREDITS
        }).execute()

    return {
        "message": "Cuenta creada. Revisa tu email para confirmar tu cuenta antes de iniciar sesión.",
        "email": email
    }


@app.post("/api/auth/login")
def login(email: str = Form(...), password: str = Form(...)):
    """Inicia sesión y devuelve el token JWT + datos del usuario."""
    email = email.strip().lower()
    try:
        resp = db().auth.sign_in_with_password({"email": email, "password": password})
        if not resp.user or not resp.session:
            raise HTTPException(401, "Credenciales incorrectas")
    except HTTPException:
        raise
    except Exception as e:
        msg = str(e).lower()
        if "invalid" in msg or "credentials" in msg or "password" in msg:
            raise HTTPException(401, "Email o contraseña incorrectos")
        if "not confirmed" in msg or "confirm" in msg:
            raise HTTPException(403, "Debes confirmar tu email antes de iniciar sesión. Revisa tu bandeja de entrada.")
        raise HTTPException(401, f"Error al iniciar sesión: {e}")

    # Asegurar que el usuario existe en nuestra tabla
    existing = db().table("users").select("*").eq("email", email).execute()
    if not existing.data:
        db().table("users").insert({
            "email": email,
            "credits": FREE_SIGNUP_CREDITS
        }).execute()
        user_data = db().table("users").select("*").eq("email", email).execute().data[0]
    else:
        user_data = existing.data[0]

    return {
        "access_token": resp.session.access_token,
        "token_type":   "bearer",
        "email":        email,
        "credits":      user_data["credits"],
    }


@app.post("/api/auth/logout")
def logout(authorization: str = Header(None)):
    """Invalida el token actual."""
    if authorization and authorization.startswith("Bearer "):
        token = authorization.split(" ", 1)[1]
        try:
            db().auth.sign_out()
        except Exception:
            pass
    return {"message": "Sesión cerrada"}


@app.post("/api/auth/forgot-password")
def forgot_password(email: str = Form(...)):
    """Envía email para resetear contraseña."""
    email = email.strip().lower()
    try:
        db().auth.reset_password_email(email)
    except Exception:
        pass  # No revelar si el email existe o no
    return {"message": "Si el email existe, recibirás instrucciones para restablecer tu contraseña."}


# ---------------------------------------------------------------------------
# Usuario / créditos (ahora protegidos con JWT)
# ---------------------------------------------------------------------------

@app.get("/api/user/me")
def get_me(authorization: str = Header(None)):
    """Devuelve datos del usuario autenticado."""
    user = get_user_from_token(authorization)
    return user


@app.post("/api/credits/add")
def add_credits(
    amount: int = Form(...),
    authorization: str = Header(None)
):
    """
    PLACEHOLDER de compra — en Fase 2 esto lo llama el webhook de Stripe.
    Por ahora solo para pruebas internas.
    """
    user = get_user_from_token(authorization)
    new_balance = user["credits"] + amount
    db().table("users").update({"credits": new_balance}).eq("email", user["email"]).execute()
    return {"email": user["email"], "credits": new_balance}


# ---------------------------------------------------------------------------
# Verificación (protegida con JWT)
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
def verify_batch(
    file: UploadFile = File(...),
    authorization: str = Header(None)
):
    user = get_user_from_token(authorization)

    content = file.file.read()
    emails  = extract_emails_from_upload(file.filename, content)

    if not emails:
        raise HTTPException(400, "No se encontraron emails válidos en el archivo")
    if len(emails) > MAX_EMAILS_PER_JOB:
        raise HTTPException(400, f"Límite de {MAX_EMAILS_PER_JOB} emails por lote. Divide el archivo.")
    if user["credits"] < len(emails):
        raise HTTPException(402, f"Créditos insuficientes. Necesitas {len(emails)}, tienes {user['credits']}.")

    results = []
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        futures = {pool.submit(verify_email, e): e for e in emails}
        for future in as_completed(futures):
            results.append(future.result())

    new_balance = user["credits"] - len(emails)
    db().table("users").update({"credits": new_balance}).eq("email", user["email"]).execute()

    # Privacidad: no persistimos los emails ni resultados individuales en la
    # base de datos, solo un conteo agregado por estado (igual que la app de
    # escritorio). Los resultados completos SÍ se devuelven en esta respuesta
    # para que el usuario pueda descargar su CSV justo después de verificar,
    # pero nunca quedan guardados en Supabase.
    status_counts = dict(Counter(r["status"] for r in results))

    job = {
        "id":            str(uuid.uuid4()),
        "user_email":    user["email"],
        "filename":      file.filename,
        "total_emails":  len(emails),
        "results":       None,
        "status_counts": status_counts,
        "source":        "web",
    }
    db().table("jobs").insert(job).execute()

    return {"credits_restantes": new_balance, "total": len(results), "resultados": results}


@app.post("/api/jobs/submit")
def submit_local_job(
    filename:           str = Form(...),
    total_emails:       int = Form(...),
    status_counts_json: str = Form(...),
    authorization:      str = Header(None)
):
    """Recibe conteo agregado de la app de escritorio. Nunca almacena emails individuales."""
    import json as _json
    user = get_user_from_token(authorization)

    try:
        status_counts = _json.loads(status_counts_json)
    except Exception:
        raise HTTPException(400, "status_counts_json inválido")

    if not isinstance(status_counts, dict) or total_emails <= 0:
        raise HTTPException(400, "Datos de resumen inválidos")
    if user["credits"] < total_emails:
        raise HTTPException(402, f"Créditos insuficientes. Necesitas {total_emails}, tienes {user['credits']}.")

    new_balance = user["credits"] - total_emails
    db().table("users").update({"credits": new_balance}).eq("email", user["email"]).execute()

    job = {
        "id":             str(uuid.uuid4()),
        "user_email":     user["email"],
        "filename":       filename,
        "total_emails":   total_emails,
        "results":        None,
        "status_counts":  status_counts,
        "source":         "desktop",
    }
    db().table("jobs").insert(job).execute()

    return {"credits_restantes": new_balance, "total": total_emails}


@app.get("/api/jobs/history")
def list_jobs(authorization: str = Header(None)):
    user = get_user_from_token(authorization)
    res  = (
        db().table("jobs")
        .select("id,filename,total_emails,status_counts,source,created_at")
        .eq("user_email", user["email"])
        .order("created_at", desc=True)
        .execute()
    )
    return res.data


@app.get("/api/jobs/{job_id}/download")
def download_job(job_id: str, authorization: str = Header(None)):
    user = get_user_from_token(authorization)
    res  = db().table("jobs").select("*").eq("id", job_id).eq("user_email", user["email"]).execute()
    if not res.data:
        raise HTTPException(404, "Job no encontrado")
    job = res.data[0]

    if not job.get("results"):
        raise HTTPException(
            409,
            "Este lote no tiene detalle disponible: por privacidad, solo se guarda el conteo "
            "agregado, no los emails ni resultados individuales. Descarga el CSV justo después "
            "de verificar, desde la pantalla de resultados."
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
        headers={"Content-Disposition": f"attachment; filename=resultados.csv"},
    )


@app.get("/api/health")
def health():
    return {"status": "ok", "supabase_configured": supabase is not None}
