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

from datetime import datetime, timezone

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Header, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import openpyxl
import stripe

from verifier import (
    verify_email,
    EMAIL_REGEX,
    DISPOSABLE_DOMAINS,
    TYPO_DOMAINS,
    get_mx_records,
)
from supabase import create_client, Client

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

SUPABASE_URL        = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY        = os.getenv("SUPABASE_KEY", "")
FRONTEND_ORIGIN     = os.getenv("FRONTEND_ORIGIN", "*")
FREE_SIGNUP_CREDITS = int(os.getenv("FREE_SIGNUP_CREDITS", "100"))
MAX_EMAILS_PER_JOB  = int(os.getenv("MAX_EMAILS_PER_JOB", "300"))
WORKERS             = int(os.getenv("VERIFY_WORKERS", "20"))

# ---------------------------------------------------------------------------
# Estimado rápido de bounce (quick-estimate) — sin verificación SMTP real,
# solo formato + dominio desechable/typo + MX + reputación histórica
# acumulada en la tabla domain_reputation. Pensado para dar un resultado
# instantáneo antes de que el usuario gaste créditos en la verificación real.
# ---------------------------------------------------------------------------
MAX_EMAILS_QUICK_ESTIMATE = int(os.getenv("MAX_EMAILS_QUICK_ESTIMATE", "50000"))

# Proveedores grandes y conocidos: cuando no hay datos históricos propios
# todavía para un dominio, se usa esta tasa base de bounce estimada en vez
# de la genérica "dominio desconocido".
KNOWN_SAFE_DOMAINS = {
    "gmail.com": 0.01, "outlook.com": 0.02, "hotmail.com": 0.03,
    "yahoo.com": 0.03, "icloud.com": 0.02, "aol.com": 0.04,
    "live.com": 0.03, "msn.com": 0.03, "protonmail.com": 0.02,
}
# Tasa asumida para un dominio corporativo/desconocido sin historial propio
# ni coincidencia con la lista de arriba, pero con MX válido.
DEFAULT_UNKNOWN_DOMAIN_RATE = 0.12

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY) if SUPABASE_URL and SUPABASE_KEY else None

# ---------------------------------------------------------------------------
# Stripe — planes de suscripción y paquetes de créditos sueltos
# ---------------------------------------------------------------------------

stripe.api_key       = os.getenv("STRIPE_SECRET_KEY", "")
STRIPE_WEBHOOK_SECRET = os.getenv("STRIPE_WEBHOOK_SECRET", "")

# Créditos que otorga cada plan por cada ciclo de facturación (mensual o anual)
PLAN_CREDITS = {
    "esencial":   1500,
    "starter":    2800,
    "pro":        5000,
    "growth":     12000,
    "business":   22000,
    "enterprise": 75000,
}

# Créditos que otorga cada paquete de pago único (sin suscripción)
PACK_CREDITS = {
    "chico":   100,
    "mediano": 500,
    "grande":  2000,
}


def get_price_id(plan: str = None, interval: str = None, pack: str = None) -> str:
    """Busca el Price ID real de Stripe en variables de entorno.
    Nunca se hardcodean IDs de Stripe en el código: viven en el dashboard
    de Render, así se pueden actualizar sin tocar código."""
    if pack:
        env_name = f"STRIPE_PRICE_PACK_{pack.upper()}"
    else:
        env_name = f"STRIPE_PRICE_{plan.upper()}_{interval.upper()}"
    price_id = os.getenv(env_name)
    if not price_id:
        raise HTTPException(500, f"Precio no configurado en el servidor: {env_name}")
    return price_id


def ensure_stripe_customer(user: dict) -> str:
    """Devuelve el stripe_customer_id del usuario, creándolo en Stripe la
    primera vez que lo necesite."""
    if user.get("stripe_customer_id"):
        return user["stripe_customer_id"]
    customer = stripe.Customer.create(email=user["email"])
    db().table("users").update({"stripe_customer_id": customer.id}).eq("email", user["email"]).execute()
    return customer.id


def add_credits_to_email(user_email: str, amount: int):
    res = db().table("users").select("credits").eq("email", user_email).execute()
    if not res.data:
        return
    new_balance = res.data[0]["credits"] + amount
    db().table("users").update({"credits": new_balance}).eq("email", user_email).execute()

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
            "credits": FREE_SIGNUP_CREDITS,
            "plan": "free",
            "plan_status": "active",
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
            "credits": FREE_SIGNUP_CREDITS,
            "plan": "free",
            "plan_status": "active",
        }).execute()
        user_data = db().table("users").select("*").eq("email", email).execute().data[0]
    else:
        user_data = existing.data[0]

    return {
        "access_token": resp.session.access_token,
        "token_type":   "bearer",
        "email":        email,
        "credits":      user_data["credits"],
        "plan":         user_data.get("plan"),
        "plan_status":  user_data.get("plan_status"),
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


@app.post("/api/billing/checkout")
def create_checkout(
    plan:          str = Form(None),
    interval:      str = Form(None),
    pack:          str = Form(None),
    authorization: str = Header(None),
):
    """
    Crea una sesión de Stripe Checkout.
    - Para suscripción: mandar plan (esencial/starter/pro/growth/business/enterprise)
      + interval (monthly/annual)
    - Para créditos sueltos: mandar pack (chico/mediano/grande)
    Devuelve la URL a la que el frontend debe redirigir al usuario.
    """
    user = get_user_from_token(authorization)

    if pack:
        if pack not in PACK_CREDITS:
            raise HTTPException(400, "Paquete de créditos inválido")
        price_id = get_price_id(pack=pack)
        mode = "payment"
        metadata = {"user_email": user["email"], "pack": pack}
        subscription_data = None
    elif plan:
        if plan not in PLAN_CREDITS or interval not in ("monthly", "annual"):
            raise HTTPException(400, "Plan o periodicidad inválidos")
        price_id = get_price_id(plan=plan, interval=interval)
        mode = "subscription"
        metadata = {"user_email": user["email"], "plan": plan}
        subscription_data = {"metadata": {"user_email": user["email"], "plan": plan}}
    else:
        raise HTTPException(400, "Debes indicar un plan (con interval) o un pack")

    customer_id = ensure_stripe_customer(user)

    session_args = dict(
        mode=mode,
        customer=customer_id,
        line_items=[{"price": price_id, "quantity": 1}],
        success_url=f"{FRONTEND_ORIGIN}/?checkout=success",
        cancel_url=f"{FRONTEND_ORIGIN}/?checkout=cancel",
        metadata=metadata,
    )
    if subscription_data:
        session_args["subscription_data"] = subscription_data

    session = stripe.checkout.Session.create(**session_args)
    return {"checkout_url": session.url}


@app.post("/api/billing/portal")
def billing_portal(authorization: str = Header(None)):
    """Sesión del Customer Portal de Stripe: permite al usuario ver facturas,
    cambiar de plan o cancelar su suscripción sin que tengamos que construir
    esa UI nosotros."""
    user = get_user_from_token(authorization)
    if not user.get("stripe_customer_id"):
        raise HTTPException(400, "Aún no tienes una suscripción o compra registrada")
    session = stripe.billing_portal.Session.create(
        customer=user["stripe_customer_id"],
        return_url=FRONTEND_ORIGIN,
    )
    return {"portal_url": session.url}


@app.post("/api/billing/webhook")
async def stripe_webhook(request: Request):
    """
    Endpoint que Stripe llama directamente (no lo llama el frontend).
    NO lleva JWT — se autentica verificando la firma con STRIPE_WEBHOOK_SECRET.
    """
    payload    = await request.body()
    sig_header = request.headers.get("stripe-signature")
    try:
        event = stripe.Webhook.construct_event(payload, sig_header, STRIPE_WEBHOOK_SECRET)
    except Exception as e:
        raise HTTPException(400, f"Firma de webhook inválida: {e}")

    # Idempotencia: Stripe puede reintentar el mismo evento varias veces.
    event_id = event["id"]
    already  = db().table("billing_events").select("id").eq("id", event_id).execute()
    if already.data:
        return {"received": True}
    db().table("billing_events").insert({"id": event_id}).execute()

    etype = event["type"]
    obj   = event["data"]["object"]

    if etype == "checkout.session.completed" and obj.get("mode") == "payment":
        # Compra de créditos sueltos (pago único)
        meta       = obj.get("metadata") or {}
        pack       = meta.get("pack")
        user_email = meta.get("user_email")
        if pack in PACK_CREDITS and user_email:
            add_credits_to_email(user_email, PACK_CREDITS[pack])

    elif etype == "invoice.paid":
        # Cubre tanto el primer cobro de una suscripción como cada renovación
        sub_id = obj.get("subscription")
        if sub_id:
            sub        = stripe.Subscription.retrieve(sub_id)
            meta       = sub.get("metadata") or {}
            plan       = meta.get("plan")
            user_email = meta.get("user_email")
            if plan in PLAN_CREDITS and user_email:
                add_credits_to_email(user_email, PLAN_CREDITS[plan])
                period_end = datetime.fromtimestamp(
                    sub["current_period_end"], tz=timezone.utc
                ).isoformat()
                db().table("users").update({
                    "plan":            plan,
                    "plan_status":     "active",
                    "plan_period_end": period_end,
                }).eq("email", user_email).execute()

    elif etype == "customer.subscription.deleted":
        meta       = obj.get("metadata") or {}
        user_email = meta.get("user_email")
        if user_email:
            db().table("users").update({"plan_status": "canceled"}).eq("email", user_email).execute()

    return {"received": True}


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


def extract_emails_from_text(text: str):
    """Extrae emails de texto plano (pegado directo), igual que en el cliente."""
    import re as _re
    found = _re.findall(r"[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+", text or "")
    seen = set()
    unique = []
    for e in found:
        key = e.lower()
        if key not in seen:
            seen.add(key)
            unique.append(e)
    return unique


def get_domain_reputation_map(domains: list):
    """Trae de Supabase las estadísticas históricas ya acumuladas para estos
    dominios (de verificaciones reales previas de cualquier usuario)."""
    if not domains:
        return {}
    try:
        res = (
            db().table("domain_reputation")
            .select("domain,total_checked,total_bounced")
            .in_("domain", domains)
            .execute()
        )
        return {row["domain"]: row for row in res.data}
    except Exception:
        # Si la tabla no existe todavía (no se corrió la migración) o falla
        # la consulta, seguimos sin datos históricos en vez de tumbar el
        # endpoint completo.
        return {}


@app.post("/api/quick-estimate")
def quick_estimate(
    file: UploadFile = File(None),
    emails_text: str = Form(None),
    authorization: str = Header(None),
):
    """
    Estimado de bounce INSTANTÁNEO, sin verificación SMTP real y SIN costo
    de créditos. Agrupa por dominio único (no por email) para que sea rápido
    incluso con listas grandes, usando:
      1. Formato de email válido
      2. Dominios desechables/temporales conocidos
      3. Typos de dominio conocidos (gmial.com, etc.)
      4. Registro MX (existe o no, sin conectar por SMTP)
      5. Reputación histórica acumulada en domain_reputation, si ya se vio
         ese dominio antes en una verificación real de cualquier usuario.
    Pensado para mostrarse ANTES de que el usuario gaste créditos en la
    verificación real.
    """
    user = get_user_from_token(authorization)

    if file is not None:
        content = file.file.read()
        emails = extract_emails_from_upload(file.filename, content)
    elif emails_text is not None:
        emails = extract_emails_from_text(emails_text)
    else:
        raise HTTPException(400, "Envía un archivo o una lista de emails pegados")

    if not emails:
        raise HTTPException(400, "No se encontraron emails válidos")
    if len(emails) > MAX_EMAILS_QUICK_ESTIMATE:
        raise HTTPException(400, f"Límite de {MAX_EMAILS_QUICK_ESTIMATE} emails para el estimado rápido")

    invalid_format = 0
    by_domain = {}  # domain -> {"count": n, "emails_sample": [...]}

    for email in emails:
        e = email.strip()
        if not EMAIL_REGEX.match(e):
            invalid_format += 1
            continue
        domain = e.split("@", 1)[1].lower()
        by_domain.setdefault(domain, 0)
        by_domain[domain] += 1

    unique_domains = list(by_domain.keys())
    reputation_map = get_domain_reputation_map(unique_domains)

    domains_out = []
    no_mx_count = 0
    disposable_count = 0
    typo_count = 0
    estimated_bounces = 0.0

    # MX en paralelo — es la parte más lenta (una consulta DNS por dominio
    # único, no por email), pero sigue siendo rápida comparado con SMTP real.
    with ThreadPoolExecutor(max_workers=min(20, max(1, len(unique_domains)))) as pool:
        mx_futures = {pool.submit(get_mx_records, d): d for d in unique_domains}
        mx_results = {}
        for future in as_completed(mx_futures):
            d = mx_futures[future]
            try:
                mx_results[d] = future.result()
            except Exception:
                mx_results[d] = []

    for domain, count in by_domain.items():
        is_disposable = domain in DISPOSABLE_DOMAINS
        is_typo = domain in TYPO_DOMAINS
        mx = mx_results.get(domain)
        has_mx = bool(mx)  # None (NXDOMAIN) o [] (sin MX) -> False

        if is_disposable:
            disposable_count += count
        if is_typo:
            typo_count += count
        if not has_mx:
            no_mx_count += count

        rep = reputation_map.get(domain)
        historical_rate = None
        if rep and rep.get("total_checked", 0) >= 5:
            historical_rate = rep["total_bounced"] / rep["total_checked"]

        if not has_mx:
            rate = 1.0
            risk = "muy alto"
            reason = "El dominio no tiene registros MX (no existe o no recibe correo)"
        elif is_disposable:
            rate = 0.9
            risk = "muy alto"
            reason = "Dominio de correo desechable/temporal"
        elif is_typo:
            rate = 0.8
            risk = "alto"
            reason = "Parece un typo de un dominio conocido (gmial.com, etc.)"
        elif historical_rate is not None:
            rate = historical_rate
            risk = "bajo" if rate < 0.05 else ("medio" if rate < 0.2 else "alto")
            reason = f"Basado en {rep['total_checked']} verificaciones reales previas de este dominio"
        elif domain in KNOWN_SAFE_DOMAINS:
            rate = KNOWN_SAFE_DOMAINS[domain]
            risk = "bajo"
            reason = "Proveedor de correo grande y conocido"
        else:
            rate = DEFAULT_UNKNOWN_DOMAIN_RATE
            risk = "medio"
            reason = "Dominio con MX válido pero sin historial propio todavía — estimado genérico"

        estimated_bounces += rate * count
        domains_out.append({
            "domain": domain,
            "count": count,
            "has_mx": has_mx,
            "disposable": is_disposable,
            "typo": is_typo,
            "historical_bounce_rate": historical_rate,
            "estimated_bounce_rate": round(rate, 3),
            "risk": risk,
            "reason": reason,
        })

    domains_out.sort(key=lambda d: d["count"], reverse=True)
    total = len(emails)
    estimated_bounces += invalid_format  # el formato inválido siempre rebota

    return {
        "total_emails": total,
        "unique_domains": len(unique_domains),
        "invalid_format": invalid_format,
        "disposable": disposable_count,
        "typo_domain": typo_count,
        "no_mx": no_mx_count,
        "estimated_bounce_rate": round(estimated_bounces / total, 4) if total else 0,
        "estimated_bounces": round(estimated_bounces),
        "domains": domains_out,
    }



def update_domain_reputation_from_results(results: list):
    """
    Toma resultados de una verificación REAL (con status por email) y
    actualiza la tabla domain_reputation con conteos agregados por dominio.
    Nunca guarda los emails individuales, solo el dominio y cuántos de
    'bounce' vs 'aceptados' se vieron.

    Se llama siempre en segundo plano (try/except) para que un fallo aquí
    nunca tumbe la respuesta real de la verificación al usuario.
    """
    BOUNCE_STATUSES = {"Rejected", "No MX", "SPAM Block", "Formato inválido"}
    ACCEPTED_STATUSES = {"Accepted", "Catch-All"}

    per_domain = {}
    for r in results:
        email = r.get("email", "")
        if "@" not in email:
            continue
        domain = email.split("@", 1)[1].lower()
        entry = per_domain.setdefault(domain, {"checked": 0, "bounced": 0, "accepted": 0})
        entry["checked"] += 1
        if r.get("status") in BOUNCE_STATUSES:
            entry["bounced"] += 1
        elif r.get("status") in ACCEPTED_STATUSES:
            entry["accepted"] += 1

    for domain, stats in per_domain.items():
        try:
            db().rpc("upsert_domain_reputation", {
                "p_domain": domain,
                "p_checked": stats["checked"],
                "p_bounced": stats["bounced"],
                "p_accepted": stats["accepted"],
            }).execute()
        except Exception:
            # Si la función/tabla no existe todavía (falta correr la
            # migración) simplemente no se acumula reputación esta vez.
            pass


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

    try:
        update_domain_reputation_from_results(results)
    except Exception:
        pass  # nunca debe afectar la respuesta al usuario

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
