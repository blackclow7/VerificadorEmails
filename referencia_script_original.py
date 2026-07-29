#!/usr/bin/env python3
"""
Verificador de Emails en Lote (SMTP Batch Checker)
====================================================

Verifica el estado de entregabilidad de una lista de direcciones de email
consultando directamente los servidores SMTP de destino (sin enviar
correos reales), y clasifica cada dirección en uno de estos estados:

    Accepted     - El servidor confirma que el buzón existe
    SPAM Block   - El servidor bloqueó la consulta por reputación/spam
    Catch-All    - El dominio acepta cualquier dirección (no verificable)
    Timeout      - El servidor no respondió a tiempo
    Rejected     - El servidor rechazó explícitamente la dirección
    No MX        - El dominio no tiene registros MX (no puede recibir correo)
    Limited      - Límite de conexiones/consultas alcanzado (rate limit)
    MX Error     - No se pudo conectar/resolver el servidor MX
    Greylisted   - El servidor difirió temporalmente la respuesta (antispam)

IMPORTANTE - LIMITACIONES A TENER EN CUENTA:
    - Muchos proveedores de internet y de hosting/cloud (AWS, Azure, GCP,
      la mayoría de ISPs residenciales) BLOQUEAN el puerto 25 saliente por
      defecto para prevenir spam. Si el script se queda "colgado" en
      Timeout/MX Error para todo, muy probablemente sea por esto y no por
      un fallo del script. Se necesita una red/servidor donde el puerto 25
      esté abierto en salida.
    - Muchos servidores de correo grandes (Gmail, Outlook, Yahoo...) hoy en
      día NO permiten verificación real vía RCPT TO y siempre devuelven
      "Accepted" o cortan la conexión tras varios intentos (para eso existe
      el estado "Limited"/"SPAM Block").
    - Usa esto solo con listas de las que tengas permiso/derecho a verificar.
      Hacer barridos masivos de dominios ajenos puede considerarse abuso y
      hacer que tu IP acabe en listas negras.

Uso:
    python email_verifier.py lista.csv
    python email_verifier.py lista.csv -o resultados.csv --timeout 8 --workers 50

Dependencias:
    pip install dnspython colorama
"""

import csv
import smtplib
import socket
import sys
import time
import random
import string
import argparse
import threading
from pathlib import Path
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed

import dns.resolver

try:
    from colorama import init, Fore, Style
    init(autoreset=True)
except ImportError:
    print("Falta la dependencia 'colorama'. Instala con: pip install colorama")
    sys.exit(1)

try:
    import dns.resolver  # noqa: F811 (re-import intencional para mensaje de error claro)
except ImportError:
    print("Falta la dependencia 'dnspython'. Instala con: pip install dnspython")
    sys.exit(1)


STATUS_COLORS = {
    "Accepted":    Fore.GREEN,
    "SPAM Block":  Fore.MAGENTA,
    "Catch-All":   Fore.CYAN,
    "Timeout":     Fore.YELLOW,
    "Rejected":    Fore.RED,
    "No MX":       Fore.RED,
    "Limited":     Fore.YELLOW,
    "MX Error":    Fore.RED,
    "Greylisted":  Fore.YELLOW,
}

# Orden fijo para el resumen final
STATUS_ORDER = [
    "Accepted", "Catch-All", "Greylisted", "Limited",
    "Timeout", "SPAM Block", "Rejected", "MX Error", "No MX",
]

MX_CACHE = {}
MX_CACHE_LOCK = threading.Lock()
MX_ERROR_DETAIL = {}
PRINT_LOCK = threading.Lock()

# ---------------------------------------------------------------------------
# Datos para el análisis heurístico de "toxicidad" (score 0-5, estilo Bouncer)
# ---------------------------------------------------------------------------

# Dominios de correo desechable/temporal más comunes (lista no exhaustiva)
DISPOSABLE_DOMAINS = {
    "mailinator.com", "10minutemail.com", "guerrillamail.com", "tempmail.com",
    "temp-mail.org", "yopmail.com", "trashmail.com", "fakeinbox.com",
    "getnada.com", "throwawaymail.com", "maildrop.cc", "sharklasers.com",
    "dispostable.com", "mytemp.email", "moakt.com", "mohmal.com",
    "emailondeck.com", "mintemail.com", "spamgourmet.com", "mailnesia.com",
    "correotemporal.org", "tempinbox.com", "burnermail.io", "guerrillamailblock.com",
    "trbvm.com", "spam4.me", "tempail.com", "discard.email", "mailcatch.com",
}

# Prefijos genéricos de rol (no representan a una persona específica)
ROLE_PREFIXES = {
    "admin", "administrator", "info", "contact", "contacto", "sales", "ventas",
    "support", "soporte", "noreply", "no-reply", "webmaster", "postmaster",
    "abuse", "help", "ayuda", "hello", "hola", "office", "hr", "rh",
    "marketing", "billing", "facturacion", "compras", "purchasing",
    "sistemas", "it", "sistema", "notificaciones", "notifications",
}

# Typos comunes de proveedores grandes (indicio de dirección inválida/trampa)
TYPO_DOMAINS = {
    "gmial.com", "gmal.com", "gamil.com", "gmail.co", "gmailcom",
    "hotmial.com", "hotmal.com", "hotmai.com", "hotmailcom",
    "yaho.com", "yahooo.com", "yahoo.co", "yahoocom",
    "outlok.com", "outloo.com", "outlookcom", "outlook.co",
}

TOXICITY_MAX_SCORE = 5


def _looks_random(local_part):
    """Heurística simple: local-parts que parecen generados al azar
    (típico de listas compradas o direcciones de spam trap)."""
    lp = local_part.lower()
    if len(lp) < 10:
        return False
    if not lp.replace(".", "").replace("_", "").replace("-", "").isalnum():
        return False
    digits = sum(c.isdigit() for c in lp)
    letters = sum(c.isalpha() for c in lp)
    vowels = sum(c in "aeiou" for c in lp if c.isalpha())
    # Mucho digito+letra mezclado y pocas vocales relativas a letras -> pinta random
    if digits >= 4 and letters >= 4 and (vowels / max(letters, 1)) < 0.25:
        return True
    return False


def assess_toxicity(email, check_blocklist=False, dnsbl_timeout=4):
    """Devuelve (score 0-5, texto con las señales detectadas).

    Esto es una heurística propia basada en señales públicas, NO una
    réplica de la base de datos propietaria de toxicidad de Bouncer u
    otras herramientas comerciales (esa información es privada de cada
    proveedor). Sirve como primera capa de filtro, no como sustituto.
    """
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

    if check_blocklist:
        listed, detail = check_spamhaus_dbl(domain, timeout=dnsbl_timeout)
        if listed:
            score = TOXICITY_MAX_SCORE
            reasons.append(f"dominio listado en Spamhaus DBL ({detail})")

    score = min(score, TOXICITY_MAX_SCORE)
    return score, ("; ".join(reasons) if reasons else "sin señales detectadas")


def check_spamhaus_dbl(domain, timeout=4):
    """Consulta el Spamhaus Domain Block List (DBL) vía DNS público.
    Es un servicio gratuito de uso abierto para consultas de bajo volumen
    (uso no comercial/individual); respeta sus políticas de uso si vas a
    hacer volúmenes grandes: https://www.spamhaus.org/dbl/
    Devuelve (listado: bool, detalle: str)."""
    query = f"{domain}.dbl.spamhaus.org"
    try:
        resolver = dns.resolver.Resolver()
        resolver.nameservers = ["8.8.8.8", "1.1.1.1"]
        resolver.timeout = timeout
        resolver.lifetime = timeout
        answers = resolver.resolve(query, "A")
        codes = [str(a) for a in answers]
        return True, ",".join(codes)
    except dns.resolver.NXDOMAIN:
        return False, ""
    except Exception:
        return False, "no verificado (error de consulta)"


# ---------------------------------------------------------------------------
# Resolución DNS / MX
# ---------------------------------------------------------------------------

def get_mx_records(domain, timeout=5):
    """Devuelve lista de hosts MX ordenados por prioridad, None si NXDOMAIN,
    'ERROR' si hubo un fallo de resolución, o [] si no hay registros MX.

    Primero intenta con el resolver configurado en el sistema; si falla
    (p.ej. por bloqueo de firewall a consultas DNS directas), reintenta
    usando servidores DNS públicos (Google y Cloudflare) como respaldo.
    """
    with MX_CACHE_LOCK:
        if domain in MX_CACHE:
            return MX_CACHE[domain]

    nameserver_sets = [
        None,  # usar el resolver / nameservers del sistema
        ["8.8.8.8", "8.8.4.4"],       # Google DNS
        ["1.1.1.1", "1.0.0.1"],       # Cloudflare DNS
    ]

    result = "ERROR"
    last_exc = None

    for nameservers in nameserver_sets:
        try:
            resolver = dns.resolver.Resolver(configure=(nameservers is None))
            if nameservers:
                resolver.nameservers = nameservers
            resolver.timeout = timeout
            resolver.lifetime = timeout
            answers = resolver.resolve(domain, "MX")
            records = sorted(answers, key=lambda r: r.preference)
            result = [str(r.exchange).rstrip(".") for r in records]
            break
        except dns.resolver.NoAnswer:
            result = []
            break
        except dns.resolver.NXDOMAIN:
            result = None
            break
        except Exception as e:
            last_exc = e
            result = "ERROR"
            continue  # probar el siguiente set de nameservers

    with MX_CACHE_LOCK:
        MX_CACHE[domain] = result
        if result == "ERROR" and last_exc is not None:
            MX_ERROR_DETAIL[domain] = f"{type(last_exc).__name__}: {last_exc}"
    return result


def random_local_part(length=14):
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=length))


# ---------------------------------------------------------------------------
# Verificación SMTP
# ---------------------------------------------------------------------------

def smtp_check(mx_host, email, from_address, timeout, helo_domain):
    """Realiza el diálogo SMTP (HELO / MAIL FROM / RCPT TO) y clasifica
    la respuesta. Devuelve (status, detalle)."""
    try:
        server = smtplib.SMTP(timeout=timeout)
        server.connect(mx_host, 25)
        server.helo(helo_domain)
        server.mail(from_address)
        code, message = server.rcpt(email)
        try:
            server.quit()
        except Exception:
            pass

        msg_text = message.decode(errors="ignore") if isinstance(message, bytes) else str(message)
        lowered = msg_text.lower()

        if code == 250:
            return "Accepted", msg_text
        if code in (450, 451, 452):
            return "Greylisted", msg_text
        if code == 421:
            return "Limited", msg_text
        if code in (550, 551, 552, 553, 554):
            if any(k in lowered for k in ("spam", "blocked", "block", "reputation", "denied", "blacklist")):
                return "SPAM Block", msg_text
            return "Rejected", msg_text
        return "Rejected", msg_text

    except (socket.timeout, TimeoutError):
        return "Timeout", "La conexión superó el tiempo límite"
    except smtplib.SMTPServerDisconnected:
        return "Rejected", "El servidor cerró la conexión"
    except smtplib.SMTPConnectError as e:
        return "MX Error", f"No se pudo conectar: {e}"
    except ConnectionRefusedError:
        return "MX Error", "Conexión rechazada por el servidor"
    except OSError as e:
        lowered = str(e).lower()
        if any(k in lowered for k in ("blocked", "spam", "reputation", "blacklist")):
            return "SPAM Block", str(e)
        return "MX Error", str(e)
    except Exception as e:
        return "MX Error", str(e)


def check_catch_all(mx_host, domain, from_address, timeout, helo_domain):
    fake_email = f"{random_local_part()}@{domain}"
    status, _ = smtp_check(mx_host, fake_email, from_address, timeout, helo_domain)
    return status == "Accepted"


DOMAIN_LOCKS = {}
DOMAIN_LOCKS_GUARD = threading.Lock()
MAX_PER_DOMAIN = 3  # conexiones simultáneas máximas contra el mismo dominio


def get_domain_semaphore(domain):
    with DOMAIN_LOCKS_GUARD:
        if domain not in DOMAIN_LOCKS:
            DOMAIN_LOCKS[domain] = threading.Semaphore(MAX_PER_DOMAIN)
        return DOMAIN_LOCKS[domain]


def verify_email(email, from_address, timeout, helo_domain, check_catchall=True):
    if "@" not in email or email.count("@") != 1:
        return "Rejected", "Formato inválido"

    local, domain = email.split("@", 1)
    domain = domain.strip().lower()
    if not local.strip() or "." not in domain:
        return "Rejected", "Formato inválido"

    mx_hosts = get_mx_records(domain, timeout=timeout)

    if mx_hosts is None:
        return "No MX", "El dominio no existe (NXDOMAIN)"
    if mx_hosts == "ERROR":
        detail = MX_ERROR_DETAIL.get(domain, "Error al resolver registros MX")
        return "MX Error", detail
    if not mx_hosts:
        return "No MX", "El dominio no tiene registros MX"

    last_error = None
    domain_sem = get_domain_semaphore(domain)
    for mx_host in mx_hosts:
        with domain_sem:
            status, detail = smtp_check(mx_host, email, from_address, timeout, helo_domain)

            if status == "MX Error":
                last_error = detail
                continue  # probar el siguiente MX de la lista

            if status == "Accepted" and check_catchall:
                if check_catch_all(mx_host, domain, from_address, timeout, helo_domain):
                    return "Catch-All", "El dominio acepta cualquier dirección"

            return status, detail

    return "MX Error", last_error or "No se pudo conectar a ningún servidor MX"


# ---------------------------------------------------------------------------
# Entrada / salida de datos
# ---------------------------------------------------------------------------

def read_emails_from_csv(path):
    rows = None
    last_error = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            with open(path, newline="", encoding=encoding) as f:
                reader = csv.reader(f)
                rows = list(reader)
            break
        except UnicodeDecodeError as e:
            last_error = e
            continue

    if rows is None:
        raise last_error

    emails = []
    for row in rows:
        for cell in row:
            cell = cell.strip()
            if "@" in cell and "." in cell.split("@")[-1]:
                emails.append(cell)
                break  # solo una dirección por fila

    seen = set()
    unique_emails = []
    for e in emails:
        key = e.lower()
        if key not in seen:
            seen.add(key)
            unique_emails.append(e)
    return unique_emails


def write_results_csv(path, results):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["email", "status", "detalle", "toxicidad", "señales_toxicidad"])
        for email, status, detail, tox_score, tox_reasons in results:
            writer.writerow([email, status, detail, tox_score, tox_reasons])


# ---------------------------------------------------------------------------
# Visualización en consola (estilo terminal / PowerShell)
# ---------------------------------------------------------------------------

def print_banner():
    line = "=" * 64
    print(Fore.CYAN + Style.BRIGHT + line)
    print(Fore.CYAN + Style.BRIGHT + "   VERIFICADOR DE EMAILS EN LOTE - SMTP BATCH CHECKER")
    print(Fore.CYAN + Style.BRIGHT + line + Style.RESET_ALL)


def print_status_line(index, total, email, status, elapsed, tox_score=0):
    color = STATUS_COLORS.get(status, "")
    bar_len = 28
    filled = int(bar_len * index / total)
    bar = "#" * filled + "-" * (bar_len - filled)
    pct = (index / total) * 100
    tox_color = Fore.WHITE
    if tox_score >= 4:
        tox_color = Fore.RED
    elif tox_score >= 1:
        tox_color = Fore.YELLOW
    tox_label = f"Tox:{tox_score}"
    tox_str = f"{tox_color}{tox_label:<6}{Style.RESET_ALL}"
    line = (
        f"{Fore.WHITE}[{bar}] {pct:5.1f}% {Fore.WHITE}({index:>4}/{total:<4}) "
        f"{Style.BRIGHT}{email:<32}{Style.RESET_ALL} "
        f"{color}{status:<12}{Style.RESET_ALL} "
        f"{tox_str} "
        f"{Fore.LIGHTBLACK_EX}{elapsed:5.2f}s"
    )
    with PRINT_LOCK:
        print(line)


def print_summary(results):
    counts = Counter(r[1] for r in results)
    tox_counts = Counter(r[3] for r in results)
    line = "-" * 64
    print()
    print(Fore.CYAN + Style.BRIGHT + line)
    print(Fore.CYAN + Style.BRIGHT + "   RESUMEN")
    print(Fore.CYAN + Style.BRIGHT + line + Style.RESET_ALL)
    for status in STATUS_ORDER:
        n = counts.get(status, 0)
        if n:
            color = STATUS_COLORS[status]
            print(f"   {color}{status:<15}{Style.RESET_ALL} {n}")
    print(Fore.CYAN + Style.BRIGHT + line + Style.RESET_ALL)
    print(f"   {Style.BRIGHT}Toxicidad (0-5):{Style.RESET_ALL}")
    for score in range(0, TOXICITY_MAX_SCORE + 1):
        n = tox_counts.get(score, 0)
        if n:
            color = Fore.RED if score >= 4 else (Fore.YELLOW if score >= 1 else Fore.GREEN)
            print(f"     {color}Score {score}{Style.RESET_ALL}: {n}")
    print(Fore.CYAN + Style.BRIGHT + line + Style.RESET_ALL)
    print(f"   Total procesados: {len(results)}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Verificador de emails en lote vía SMTP")
    parser.add_argument("input_csv", help="Ruta al CSV/lista con las direcciones de email")
    parser.add_argument("-o", "--output", default="resultados.csv", help="Archivo CSV de salida")
    parser.add_argument("--from-address", default="verify@example.com",
                         help="Dirección usada en el comando MAIL FROM")
    parser.add_argument("--helo-domain", default="example.com",
                         help="Dominio usado en el comando HELO/EHLO")
    parser.add_argument("--timeout", type=int, default=10, help="Timeout de conexión en segundos")
    parser.add_argument("--no-catchall", action="store_true", help="Desactiva la detección de catch-all")
    parser.add_argument("--workers", type=int, default=30,
                         help="Cantidad de verificaciones simultáneas (default: 30)")
    parser.add_argument("--check-blocklist", action="store_true",
                         help="Consulta Spamhaus DBL por dominio para reforzar el score de toxicidad (más lento)")
    args = parser.parse_args()

    if not Path(args.input_csv).exists():
        print(Fore.RED + f"Error: no se encontró el archivo '{args.input_csv}'")
        sys.exit(1)

    print_banner()
    emails = read_emails_from_csv(args.input_csv)
    total = len(emails)
    print(Fore.WHITE + f"\n  Emails encontrados: {total}")
    print(Fore.WHITE + f"  Hilos simultáneos: {args.workers}\n")

    if total == 0:
        print(Fore.YELLOW + "  No se encontraron direcciones de correo válidas en el archivo.")
        sys.exit(0)

    results = [None] * total
    completed = 0
    t_start = time.time()

    def worker(email):
        start = time.time()
        status, detail = verify_email(
            email,
            from_address=args.from_address,
            timeout=args.timeout,
            helo_domain=args.helo_domain,
            check_catchall=not args.no_catchall,
        )
        tox_score, tox_reasons = assess_toxicity(email, check_blocklist=args.check_blocklist)
        return email, status, detail, tox_score, tox_reasons, time.time() - start

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = {pool.submit(worker, email): i for i, email in enumerate(emails)}
        for future in as_completed(futures):
            idx = futures[future]
            try:
                email, status, detail, tox_score, tox_reasons, elapsed = future.result()
            except Exception as e:
                email, status, detail, tox_score, tox_reasons, elapsed = emails[idx], "MX Error", str(e), 0, "", 0.0
            results[idx] = (email, status, detail, tox_score, tox_reasons)
            completed += 1
            print_status_line(completed, total, email, status, elapsed, tox_score)

    total_elapsed = time.time() - t_start
    print_summary(results)
    write_results_csv(args.output, results)
    print(Fore.WHITE + f"\n  Tiempo total: {total_elapsed:.1f}s")
    print(Fore.GREEN + f"  Resultados guardados en: {args.output}\n")


if __name__ == "__main__":
    main()