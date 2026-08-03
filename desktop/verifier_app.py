"""
verifier_app.py — Green Email Verifier (App de Escritorio, Windows)
=====================================================================
Esta app ya NO usa Tkinter para la interfaz: carga el mismo HTML/CSS/JS
que la web (desktop_index.html) dentro de una ventana con WebView2
(vía pywebview), para que ambas plataformas se vean y funcionen igual.

Lo único que sigue siendo nativo/Python es lo que TIENE que serlo:
  - La verificación SMTP real (HELO/MAIL FROM/RCPT TO), que requiere el
    puerto 25 y no puede hacerse desde JavaScript en un navegador.
  - La prueba de puerto 25.
  - Los diálogos nativos de "elegir archivo" / "guardar archivo".
  - El almacenamiento seguro del token de sesión (Windows Credential Manager).

Ese puente se expone al HTML a través de `js_api` de pywebview
(disponible en el HTML como `window.pywebview.api.<método>`).

Requisitos para correr desde código fuente:
    pip install pywebview requests openpyxl keyring

Para compilar a un .exe de Windows (ver desktop/BUILD.md):
    pip install pyinstaller
    pyinstaller --onefile --windowed --name VerificadorEmails
        --add-data "desktop_index.html;." verifier_app.py
"""

import os
import re
import sys
import csv
import json
import random
import string
import smtplib
import socket
import pathlib
import tempfile
import threading
import time
import webbrowser
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
import keyring  # Windows Credential Manager
import webview

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
API_BASE    = "https://verificadoremails-api.onrender.com"

FROM_ADDRESS = "verify@tudominio.com"
HELO_DOMAIN  = "tudominio.com"
SMTP_TIMEOUT = 10
WORKERS      = 8

KEYRING_SERVICE = "GreenEmailVerifier"

# Carpeta local donde se guarda automáticamente el detalle completo de cada
# lote verificado — son archivos .csv reales en el disco del usuario, nunca
# suben al servidor.
LOCAL_HISTORY_DIR = os.path.join(
    os.path.expanduser("~"), "GreenEmailVerifier", "Historial"
)

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ---------------------------------------------------------------------------
# Toxicidad (idéntica a la del backend / versión anterior de escritorio)
# ---------------------------------------------------------------------------

DISPOSABLE_DOMAINS = {
    "mailinator.com", "10minutemail.com", "guerrillamail.com", "tempmail.com",
    "temp-mail.org", "yopmail.com", "trashmail.com", "fakeinbox.com",
    "getnada.com", "throwawaymail.com", "maildrop.cc", "sharklasers.com",
    "dispostable.com", "mytemp.email", "moakt.com", "mohmal.com",
    "emailondeck.com", "mintemail.com", "spamgourmet.com", "mailnesia.com",
    "correotemporal.org", "tempinbox.com", "burnermail.io",
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


def _looks_random(local_part):
    lp = local_part.lower()
    if len(lp) < 10:
        return False
    if not lp.replace(".", "").replace("_", "").replace("-", "").isalnum():
        return False
    digits  = sum(c.isdigit() for c in lp)
    letters = sum(c.isalpha() for c in lp)
    vowels  = sum(c in "aeiou" for c in lp if c.isalpha())
    if digits >= 4 and letters >= 4 and (vowels / max(letters, 1)) < 0.25:
        return True
    return False


def assess_toxicity(email):
    if "@" not in email:
        return 0, "formato inválido"
    local, domain = email.split("@", 1)
    local, domain = local.lower(), domain.lower()
    score, reasons = 0, []
    if domain in DISPOSABLE_DOMAINS:
        score += 3; reasons.append("dominio desechable/temporal")
    if domain in TYPO_DOMAINS:
        score += 2; reasons.append("typo de dominio conocido")
    if local in ROLE_PREFIXES:
        score += 1; reasons.append("dirección de rol genérica")
    if _looks_random(local):
        score += 1; reasons.append("local-part con patrón aleatorio")
    return min(score, 5), ("; ".join(reasons) if reasons else "sin señales detectadas")


# ---------------------------------------------------------------------------
# DNS / MX (con fallback a DNS públicos, vía nslookup para evitar problemas
# de dnspython con PyInstaller)
# ---------------------------------------------------------------------------

MX_CACHE      = {}
MX_CACHE_LOCK = threading.Lock()


def get_mx_records(domain, timeout=6):
    import subprocess

    with MX_CACHE_LOCK:
        if domain in MX_CACHE:
            return MX_CACHE[domain]

    def _run_nslookup(domain, server=None):
        cmd = ["nslookup", "-type=MX", domain]
        if server:
            cmd.append(server)
        try:
            out = subprocess.run(
                cmd, capture_output=True, text=True,
                timeout=timeout, creationflags=0x08000000  # NO_WINDOW en Windows
            )
            return out.stdout + out.stderr
        except Exception as e:
            return f"ERROR: {e}"

    result = ("ERROR", "No se pudo resolver MX")

    for server in [None, "8.8.8.8", "1.1.1.1"]:
        output = _run_nslookup(domain, server)

        if any(x in output.lower() for x in [
            "non-existent domain", "nxdomain", "can't find",
            "no existe", "no encontrado"
        ]):
            result = None
            break

        mx_lines = [l for l in output.splitlines()
                    if "mail exchanger" in l.lower() or "MX preference" in l]

        if mx_lines:
            hosts = []
            for line in mx_lines:
                parts = line.strip().split()
                if parts:
                    hosts.append(parts[-1].rstrip("."))
            if hosts:
                result = hosts
                break

        if "no records" in output.lower() or (
            "answer" not in output.lower() and "exchanger" not in output.lower()
            and "error" not in output.lower() and len(output) > 10
        ):
            result = []
            break

    with MX_CACHE_LOCK:
        MX_CACHE[domain] = result
    return result


# ---------------------------------------------------------------------------
# Prueba de puerto 25
# ---------------------------------------------------------------------------

PORT25_TEST_HOSTS = [
    "gmail-smtp-in.l.google.com",
    "smtp.mail.yahoo.com",
    "outlook-com.olc.protection.outlook.com",
]


def _try_port25_host(host, timeout):
    try:
        with socket.create_connection((host, 25), timeout=timeout) as s:
            s.settimeout(timeout)
            banner = s.recv(256)
            if banner.startswith(b"220"):
                return True, host
            return False, f"Respuesta inesperada de {host}: {banner[:60]!r}"
    except socket.timeout:
        return False, f"Timeout conectando a {host}"
    except ConnectionRefusedError:
        return False, f"Conexión rechazada por {host}"
    except OSError as e:
        return False, f"Error conectando a {host}: {e}"


def check_port25_open(timeout=6):
    """Prueba los 3 servidores EN PARALELO (no uno tras otro) para que el
    peor caso sea ~timeout segundos en vez de ~3×timeout."""
    last_error = "No se pudo conectar a ningún servidor de prueba"
    with ThreadPoolExecutor(max_workers=len(PORT25_TEST_HOSTS)) as pool:
        futures = {pool.submit(_try_port25_host, h, timeout): h for h in PORT25_TEST_HOSTS}
        for future in as_completed(futures):
            ok, info = future.result()
            if ok:
                return True, info
            last_error = info
    return False, last_error


# ---------------------------------------------------------------------------
# Verificación SMTP real
# ---------------------------------------------------------------------------

def _random_probe_localpart():
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=14))


def smtp_check(mx_host, email, timeout):
    try:
        server = smtplib.SMTP(timeout=timeout)
        server.connect(mx_host, 25)
        server.helo(HELO_DOMAIN)
        server.mail(FROM_ADDRESS)
        code, message = server.rcpt(email)
        try:
            server.quit()
        except Exception:
            pass

        msg = message.decode(errors="ignore") if isinstance(message, bytes) else str(message)
        low = msg.lower()

        if code == 250:
            return "Accepted", msg
        if code in (450, 451, 452):
            return "Greylisted", msg
        if code == 421:
            return "Limited", msg
        if code in (550, 551, 552, 553, 554):
            if any(k in low for k in ("spam", "blocked", "block", "reputation", "denied", "blacklist")):
                return "SPAM Block", msg
            return "Rejected", msg
        return "Rejected", msg

    except (socket.timeout, TimeoutError):
        return "Timeout", "La conexión superó el tiempo límite"
    except smtplib.SMTPServerDisconnected:
        return "SPAM Block", "El servidor cortó la conexión (posible bloqueo por reputación)"
    except smtplib.SMTPConnectError as e:
        return "MX Error", f"No se pudo conectar: {e}"
    except ConnectionRefusedError:
        return "MX Error", "Conexión rechazada por el servidor"
    except OSError as e:
        low = str(e).lower()
        if any(k in low for k in ("blocked", "spam", "reputation", "blacklist")):
            return "SPAM Block", str(e)
        return "MX Error", str(e)
    except Exception as e:
        return "MX Error", str(e)


def verify_email_local(email):
    email = email.strip()
    tox_score, tox_reasons = assess_toxicity(email)

    if not EMAIL_REGEX.match(email):
        return {"email": email, "status": "Formato inválido", "detalle": "No cumple formato usuario@dominio",
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}

    domain    = email.split("@", 1)[1].lower()
    mx_result = get_mx_records(domain)

    if mx_result is None:
        return {"email": email, "status": "No MX", "detalle": "El dominio no existe (NXDOMAIN)",
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}
    if isinstance(mx_result, tuple) and mx_result[0] == "ERROR":
        return {"email": email, "status": "MX Error", "detalle": f"Error DNS: {mx_result[1]}",
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}
    if not mx_result:
        return {"email": email, "status": "No MX", "detalle": "El dominio no tiene registros MX",
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}

    last_detail = "No se pudo conectar a ningún servidor MX"
    for mx_host in mx_result:
        status, detail = smtp_check(mx_host, email, SMTP_TIMEOUT)
        if status == "MX Error":
            last_detail = detail
            continue
        if status == "Accepted":
            probe = f"{_random_probe_localpart()}@{domain}"
            ps, _ = smtp_check(mx_host, probe, SMTP_TIMEOUT)
            if ps == "Accepted":
                return {"email": email, "status": "Catch-All",
                        "detalle": "El dominio acepta cualquier dirección",
                        "toxicidad": tox_score, "señales_toxicidad": tox_reasons}
        return {"email": email, "status": status, "detalle": detail,
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}

    return {"email": email, "status": "MX Error", "detalle": last_detail,
            "toxicidad": tox_score, "señales_toxicidad": tox_reasons}


# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------

def extract_emails_from_file(path):
    emails = []
    if path.lower().endswith(".csv"):
        for enc in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                with open(path, "r", encoding=enc, errors="replace") as f:
                    reader = csv.reader(f)
                    for row in reader:
                        for cell in row:
                            cell = cell.strip()
                            if "@" in cell and "." in cell.split("@")[-1]:
                                emails.append(cell)
                                break
                break
            except UnicodeDecodeError:
                continue
    elif path.lower().endswith((".xlsx", ".xls")):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                for cell in row:
                    if cell is None:
                        continue
                    cell = str(cell).strip()
                    if "@" in cell and "." in cell.split("@")[-1]:
                        emails.append(cell)
                        break

    seen, unique = set(), []
    for e in emails:
        k = e.lower()
        if k not in seen:
            seen.add(k)
            unique.append(e)
    return unique


# ---------------------------------------------------------------------------
# Puente Python <-> HTML (expuesto como window.pywebview.api en el JS)
# ---------------------------------------------------------------------------

class Api:
    def __init__(self):
        self.window = None  # se asigna después de crear la ventana
        self.cancel_requested = False  # bandera para detener verify_batch en curso

    # ── Puerto 25 ────────────────────────────────────────────────────────
    def check_port25(self, context="default"):
        """
        No bloquea: corre la prueba en un hilo aparte y le avisa al HTML
        cuando termina (window.onPort25Result). Si esto se ejecutara de
        forma síncrona aquí, toda la ventana se congelaría durante los
        varios segundos que tarda la prueba de red (pywebview corre las
        funciones expuestas en el mismo hilo de la interfaz).
        """
        def worker():
            is_open, info = check_port25_open()
            if self.window:
                try:
                    payload = json.dumps({"context": context, "open": is_open, "info": info})
                    self.window.evaluate_js(f"window.onPort25Result({payload})")
                except Exception:
                    pass
        threading.Thread(target=worker, daemon=True).start()
        return {"started": True}

    # ── Verificación SMTP real, con progreso en vivo empujado al HTML ───
    def save_pasted_emails(self, text):
        """
        Guarda una lista de emails pegados directamente en la app en un
        archivo temporal .txt, para poder reutilizar exactamente el mismo
        flujo (count_emails_in_file + verify_batch) que con un archivo real.
        """
        try:
            emails = re.findall(r"[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+", text)
            emails = list(dict.fromkeys(emails))  # quita duplicados preservando orden
            if not emails:
                return {"path": None, "total": 0}
            fd, path = tempfile.mkstemp(suffix=".txt", prefix="pegado_")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write("\n".join(emails))
            return {"path": path, "total": len(emails)}
        except Exception as e:
            return {"path": None, "total": 0, "error": str(e)}

    def count_emails_in_file(self, file_path):
        """
        Cuenta cuántos emails hay en el archivo ANTES de verificar, para que
        el frontend pueda comparar contra los créditos disponibles y avisar
        al usuario si no le van a alcanzar, sin gastar tiempo verificando.
        También devuelve la lista completa de emails, para poder mandarla
        al endpoint /api/quick-estimate (estimado rápido, gratis, sin SMTP).
        """
        try:
            emails = extract_emails_from_file(file_path)
            return {"total": len(emails), "emails": emails}
        except Exception as e:
            return {"total": 0, "emails": [], "error": str(e)}

    def cancel_verification(self):
        """
        Llamado desde el botón 'Detener' del HTML. No detiene instantáneamente
        (los hilos ya en vuelo terminan su verificación actual), pero hace que
        el worker deje de esperar más resultados y entregue lo que ya se
        verificó hasta ese momento, en vez de descartarlo.
        """
        self.cancel_requested = True
        return {"ok": True}

    def verify_batch(self, file_path):
        """
        Tampoco bloquea: arranca la verificación completa en un hilo aparte
        y va empujando cada resultado (window.onVerifyProgress) y al final
        el arreglo completo (window.onVerifyComplete). Un lote de miles de
        emails podía tardar minutos — bloquear la interfaz ese tiempo sería
        inaceptable.
        """
        self.cancel_requested = False

        def worker():
            emails = extract_emails_from_file(file_path)
            total  = len(emails)
            results = []

            def emit(done, result):
                if self.window:
                    try:
                        payload = json.dumps({"done": done, "total": total, "result": result})
                        self.window.evaluate_js(f"window.onVerifyProgress({payload})")
                    except Exception:
                        pass

            with ThreadPoolExecutor(max_workers=WORKERS) as pool:
                futures = {pool.submit(verify_email_local, e): e for e in emails}
                done = 0
                for future in as_completed(futures):
                    if self.cancel_requested:
                        # Cancela lo que aún no arrancó; lo ya en vuelo se
                        # deja terminar solo (ThreadPoolExecutor no permite
                        # interrumpir hilos ya corriendo en Python).
                        for f in futures:
                            f.cancel()
                        break
                    r = future.result()
                    results.append(r)
                    done += 1
                    emit(done, r)

            if self.window:
                try:
                    payload = json.dumps(results)
                    self.window.evaluate_js(f"window.onVerifyComplete({payload})")
                except Exception:
                    pass

        threading.Thread(target=worker, daemon=True).start()
        return {"started": True}

    # ── Diálogos nativos de archivo ──────────────────────────────────────
    def choose_file(self):
        result = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("Archivos CSV y Excel (*.csv;*.xlsx;*.xls)", "Todos los archivos (*.*)"),
        )
        return result[0] if result else None

    def save_csv(self, content, suggested_name):
        result = self.window.create_file_dialog(
            webview.SAVE_DIALOG, save_filename=suggested_name or "resultados.csv"
        )
        if result:
            path = result if isinstance(result, str) else result[0]
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                f.write(content)
            return True
        return False

    # ── Credenciales seguras (Windows Credential Manager) ────────────────
    def get_credential(self, key):
        try:
            return keyring.get_password(KEYRING_SERVICE, key)
        except Exception:
            return None

    def set_credential(self, key, value):
        try:
            keyring.set_password(KEYRING_SERVICE, key, value)
            return True
        except Exception:
            return False

    def delete_credential(self, key):
        try:
            keyring.delete_password(KEYRING_SERVICE, key)
            return True
        except Exception:
            return False

    # ── Abrir enlaces externos (checkout de Stripe, crear cuenta, etc.) ──
    def open_external(self, url):
        webbrowser.open(url)
        return True

    # ── Historial local en disco (archivos .csv reales, no navegador) ────
    def save_local_history(self, filename, csv_content):
        """Guarda automáticamente el detalle completo de un lote verificado
        como un .csv real en el disco del usuario. Nunca sube al servidor."""
        try:
            os.makedirs(LOCAL_HISTORY_DIR, exist_ok=True)
            ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
            base = os.path.splitext(os.path.basename(filename or "resultados"))[0]
            path = os.path.join(LOCAL_HISTORY_DIR, f"{ts}_{base}.csv")
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                f.write(csv_content)
            return path
        except Exception:
            return None

    def list_local_history(self):
        """Devuelve los archivos guardados en la carpeta de historial local,
        más recientes primero."""
        os.makedirs(LOCAL_HISTORY_DIR, exist_ok=True)
        files = sorted(pathlib.Path(LOCAL_HISTORY_DIR).glob("*.csv"),
                        key=lambda p: p.stat().st_mtime, reverse=True)
        return [
            {"name": f.name, "path": str(f), "mtime": f.stat().st_mtime}
            for f in files
        ]

    def open_path(self, path):
        """Abre un archivo o carpeta con el programa predeterminado de Windows."""
        try:
            os.startfile(path)
            return True
        except Exception:
            return False

    def open_history_folder(self):
        os.makedirs(LOCAL_HISTORY_DIR, exist_ok=True)
        return self.open_path(LOCAL_HISTORY_DIR)

    def get_history_folder_path(self):
        os.makedirs(LOCAL_HISTORY_DIR, exist_ok=True)
        return LOCAL_HISTORY_DIR


def resource_path(relative_path):
    """Encuentra el archivo tanto corriendo desde código fuente como
    empaquetado con PyInstaller (que descomprime todo en sys._MEIPASS)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)


# Pantalla de carga: se muestra de inmediato (es HTML mínimo, sin puente
# js_api, así que renderiza rápido) mientras la ventana real y WebView2
# terminan de inicializarse por debajo. Sin esto, el usuario ve una ventana
# en blanco o "No responde" durante ese arranque, aunque no sea un error.
SPLASH_HTML = """<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>
  html,body{ margin:0; height:100%; }
  body{
    display:flex; flex-direction:column; align-items:center; justify-content:center;
    background:#157A42; font-family:Segoe UI, Arial, sans-serif; color:#fff;
  }
  .spinner{
    width:34px; height:34px; border:4px solid rgba(255,255,255,.3);
    border-top-color:#fff; border-radius:50%; margin-bottom:16px;
    animation:spin .8s linear infinite;
  }
  @keyframes spin{ to{ transform:rotate(360deg); } }
  .label{ font-size:14px; opacity:.9; }
</style></head>
<body>
  <div class="spinner"></div>
  <div class="label">Cargando Green Email Verifier…</div>
</body></html>
"""


if __name__ == "__main__":
    api = Api()

    splash = webview.create_window(
        "Green Email Verifier",
        html=SPLASH_HTML,
        width=380,
        height=220,
        resizable=False,
        frameless=True,
        on_top=True,
    )

    main_window = webview.create_window(
        "Green Email Verifier — Verificador de Escritorio",
        resource_path("desktop_index.html"),
        js_api=api,
        width=1280,
        height=820,
        min_size=(1024, 700),
        hidden=True,  # se revela solo cuando ya está lista
    )
    api.window = main_window

    _revealed = threading.Event()

    def _reveal_main():
        if _revealed.is_set():
            return
        _revealed.set()
        try:
            main_window.show()
        except Exception:
            pass
        try:
            splash.destroy()
        except Exception:
            pass

    def _on_main_loaded():
        # Pequeño margen extra para que el puente js_api termine de asentarse
        # antes de mostrar la ventana (el JS igual reintenta por su cuenta,
        # esto solo mejora la primera impresión).
        time.sleep(0.4)
        _reveal_main()

    def _fallback_timeout():
        # Si por lo que sea el evento 'loaded' nunca dispara, no dejamos al
        # usuario viendo el splash para siempre.
        time.sleep(8)
        _reveal_main()

    main_window.events.loaded += _on_main_loaded
    threading.Thread(target=_fallback_timeout, daemon=True).start()

    # Se fuerza explícitamente el motor EdgeChromium (WebView2). Sin esto,
    # dentro de un .exe empaquetado a veces pywebview cae en silencio a un
    # motor viejo (Trident/IE) que no soporta bien el puente js_api, y los
    # métodos de Python (como check_port25) nunca aparecen del lado del HTML.
    webview.start(gui="edgechromium")
