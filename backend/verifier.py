"""
verifier.py
===========
Lógica de verificación de emails para el backend web.

Adaptado de email_verifier.py (script local del cliente). Diferencia clave:
en hosting gratuito (Render, Railway, Vercel, Fly free, etc.) el puerto 25
saliente está bloqueado por defecto, así que la verificación SMTP real
(RCPT TO) NO funciona ahí. Este módulo hace por defecto una verificación
"ligera" (sintaxis + MX + toxicidad) que no depende del puerto 25.

Si en el futuro despliegas esto en un servidor/VPS con el puerto 25 abierto
en salida, pon ENABLE_SMTP_CHECK=true en las variables de entorno y se
activará la verificación SMTP real (idéntica a la del script original).
"""

import os
import re
import smtplib
import socket
import random
import string
import threading
from collections import defaultdict

import dns.resolver

ENABLE_SMTP_CHECK = os.getenv("ENABLE_SMTP_CHECK", "false").lower() == "true"

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

MX_CACHE = {}
MX_CACHE_LOCK = threading.Lock()

DOMAIN_SEMAPHORES = defaultdict(lambda: threading.Semaphore(4))
DOMAIN_SEM_LOCK = threading.Lock()

# ---------------------------------------------------------------------------
# Heurística de toxicidad (portada tal cual de tu script)
# ---------------------------------------------------------------------------

DISPOSABLE_DOMAINS = {
    "mailinator.com", "10minutemail.com", "guerrillamail.com", "tempmail.com",
    "temp-mail.org", "yopmail.com", "trashmail.com", "fakeinbox.com",
    "getnada.com", "throwawaymail.com", "maildrop.cc", "sharklasers.com",
    "dispostable.com", "mytemp.email", "moakt.com", "mohmal.com",
    "emailondeck.com", "mintemail.com", "spamgourmet.com", "mailnesia.com",
    "correotemporal.org", "tempinbox.com", "burnermail.io", "guerrillamailblock.com",
    "trbvm.com", "spam4.me", "tempail.com", "discard.email", "mailcatch.com",
}

ROLE_PREFIXES = {
    "admin", "administrator", "info", "contact", "contacto", "sales", "ventas",
    "support", "soporte", "noreply", "no-reply", "webmaster", "postmaster",
    "abuse", "help", "ayuda", "hello", "hola", "office", "hr", "rh",
    "marketing", "billing", "facturacion", "compras", "purchasing",
    "sistemas", "it", "sistema", "notificaciones", "notifications",
}

TYPO_DOMAINS = {
    "gmial.com", "gmal.com", "gamil.com", "gmail.co", "gmailcom",
    "hotmial.com", "hotmal.com", "hotmai.com", "hotmailcom",
    "yaho.com", "yahooo.com", "yahoo.co", "yahoocom",
    "outlok.com", "outloo.com", "outlookcom", "outlook.co",
}

TOXICITY_MAX_SCORE = 5


def _looks_random(local_part):
    lp = local_part.lower()
    if len(lp) < 10:
        return False
    if not lp.replace(".", "").replace("_", "").replace("-", "").isalnum():
        return False
    digits = sum(c.isdigit() for c in lp)
    letters = sum(c.isalpha() for c in lp)
    vowels = sum(c in "aeiou" for c in lp if c.isalpha())
    if digits >= 4 and letters >= 4 and (vowels / max(letters, 1)) < 0.25:
        return True
    return False


def assess_toxicity(email):
    reasons = []
    score = 0
    if "@" not in email:
        return 0, "formato inválido, no evaluado"
    local, domain = email.split("@", 1)
    local = local.strip().lower()
    domain = domain.strip().lower()

    if domain in DISPOSABLE_DOMAINS:
        score += 3
        reasons.append("dominio desechable/temporal")
    if domain in TYPO_DOMAINS:
        score += 2
        reasons.append("typo de dominio conocido (gmail/hotmail/yahoo/outlook)")
    if local in ROLE_PREFIXES:
        score += 1
        reasons.append("dirección de rol genérica")
    if _looks_random(local):
        score += 1
        reasons.append("local-part con patrón aleatorio (posible spam trap)")

    score = min(score, TOXICITY_MAX_SCORE)
    return score, ("; ".join(reasons) if reasons else "sin señales detectadas")


# ---------------------------------------------------------------------------
# MX lookup (DNS, funciona en cualquier hosting - no requiere puerto 25)
# ---------------------------------------------------------------------------

def get_mx_records(domain, timeout=6):
    with MX_CACHE_LOCK:
        if domain in MX_CACHE:
            return MX_CACHE[domain]
    try:
        resolver = dns.resolver.Resolver()
        resolver.timeout = timeout
        resolver.lifetime = timeout
        answers = resolver.resolve(domain, "MX")
        hosts = sorted(
            [(r.preference, str(r.exchange).rstrip(".")) for r in answers]
        )
        result = [h for _, h in hosts]
    except dns.resolver.NXDOMAIN:
        result = None
    except (dns.resolver.NoAnswer, Exception):
        result = []
    with MX_CACHE_LOCK:
        MX_CACHE[domain] = result
    return result


# ---------------------------------------------------------------------------
# Verificación SMTP real (solo se ejecuta si ENABLE_SMTP_CHECK=true, es decir,
# si esto corre en un servidor con el puerto 25 saliente abierto)
# ---------------------------------------------------------------------------

def _random_probe_localpart():
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=12))


def smtp_check(mx_host, email, from_address, timeout, helo_domain):
    try:
        with smtplib.SMTP(mx_host, 25, timeout=timeout) as server:
            server.set_debuglevel(0)
            server.ehlo(helo_domain)
            if server.has_extn("starttls"):
                try:
                    server.starttls()
                    server.ehlo(helo_domain)
                except Exception:
                    pass
            server.mail(from_address)
            code, _ = server.rcpt(email)
            if code in (250, 251):
                return "Accepted", "El servidor confirmó el buzón"
            if code in (450, 451, 452):
                return "Greylisted", f"Código SMTP {code}"
            if code in (550, 551, 552, 553):
                return "Rejected", f"Código SMTP {code}"
            return "Limited", f"Código SMTP {code}"
    except socket.timeout:
        return "Timeout", "El servidor no respondió a tiempo"
    except smtplib.SMTPServerDisconnected:
        return "SPAM Block", "El servidor cortó la conexión (posible bloqueo por reputación)"
    except Exception as e:
        return "MX Error", str(e)


def check_catch_all(mx_host, domain, from_address, timeout, helo_domain):
    probe = f"{_random_probe_localpart()}@{domain}"
    status, _ = smtp_check(mx_host, probe, from_address, timeout, helo_domain)
    return status == "Accepted"


def real_smtp_verify(email, domain, mx_hosts, from_address, helo_domain, timeout=8):
    key = domain
    with DOMAIN_SEM_LOCK:
        sem = DOMAIN_SEMAPHORES[key]
    with sem:
        for mx_host in mx_hosts:
            status, detail = smtp_check(mx_host, email, from_address, timeout, helo_domain)
            if status == "MX Error":
                continue
            if status == "Accepted" and check_catch_all(mx_host, domain, from_address, timeout, helo_domain):
                return "Catch-All", "El dominio acepta cualquier dirección"
            return status, detail
    return "MX Error", "No se pudo conectar a ningún servidor MX"


# ---------------------------------------------------------------------------
# Función pública usada por el backend
# ---------------------------------------------------------------------------

def verify_email(email, from_address="verify@example.com", helo_domain="example.com", timeout=6):
    """Verifica un email y devuelve un dict con status, detalle y toxicidad.

    Modo por defecto (ENABLE_SMTP_CHECK=false, recomendado en hosting gratis):
        status ∈ {"Formato inválido", "Dominio inexistente", "Sin MX", "MX válido"}

    Modo SMTP real (ENABLE_SMTP_CHECK=true, requiere puerto 25 abierto):
        status ∈ {"Accepted", "Catch-All", "Rejected", "Greylisted",
                   "Timeout", "SPAM Block", "Limited", "MX Error"}
    """
    email = email.strip()
    tox_score, tox_reasons = assess_toxicity(email)

    if not EMAIL_REGEX.match(email):
        return {
            "email": email, "status": "Formato inválido", "detalle": "No cumple formato usuario@dominio",
            "toxicidad": tox_score, "señales_toxicidad": tox_reasons,
        }

    domain = email.split("@", 1)[1].lower()
    mx_hosts = get_mx_records(domain, timeout=timeout)

    if mx_hosts is None:
        status, detail = "Dominio inexistente", "El dominio no existe (NXDOMAIN)"
    elif not mx_hosts:
        status, detail = "Sin MX", "El dominio no tiene registros MX"
    elif ENABLE_SMTP_CHECK:
        status, detail = real_smtp_verify(email, domain, mx_hosts, from_address, helo_domain, timeout)
    else:
        status, detail = "MX válido", "El dominio puede recibir correo (no verificado a nivel de buzón)"

    return {
        "email": email, "status": status, "detalle": detail,
        "toxicidad": tox_score, "señales_toxicidad": tox_reasons,
    }
