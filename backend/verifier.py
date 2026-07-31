"""
verifier_app.py — Verificador de Emails (App de Escritorio, Windows)
=====================================================================
Corre la verificación SMTP REAL (puerto 25) usando la conexión de internet
del propio usuario, y descuenta créditos contra el backend web (el mismo
sistema de cuentas que la página).

Flujo:
    1. El usuario inicia sesión con su email (misma cuenta que en la web).
    2. Elige un .csv o .xlsx con emails.
    3. La app verifica LOCALMENTE por SMTP real (no pasa por el servidor).
    4. Al terminar, envía solo un conteo agregado al backend para
       descontar créditos (los emails y resultados individuales NUNCA
       salen de esta computadora).

Requisitos para correr desde código fuente:
    pip install requests dnspython openpyxl

Para compilar a un .exe de Windows (ver desktop/BUILD.md):
    pip install pyinstaller
    pyinstaller --onefile --windowed --name VerificadorEmails verifier_app.py
"""

import os
import re
import csv
import json
import random
import string
import smtplib
import socket
import pathlib
import webbrowser
import threading
import queue
import tkinter as tk
from datetime import datetime
from tkinter import ttk, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import openpyxl
import keyring  # para guardar el token de forma segura en Windows Credential Manager

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
API_BASE = "https://verificadoremails-api.onrender.com"
WEB_APP_URL = "https://verificador-emails.vercel.app"

FROM_ADDRESS = "verify@tudominio.com"
HELO_DOMAIN  = "tudominio.com"
SMTP_TIMEOUT = 10
WORKERS      = 8

# Carpeta local donde se guarda automáticamente el detalle completo de cada
# lote verificado, para poder consultarlo o descargarlo después sin depender
# del servidor (los resultados individuales nunca se suben, por privacidad).
LOCAL_HISTORY_DIR = os.path.join(
    os.path.expanduser("~"), "CorreoCertificado", "Historial"
)

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ---------------------------------------------------------------------------
# Términos y Condiciones (BORRADOR — revisar con un abogado antes de publicar)
# ---------------------------------------------------------------------------
TERMS_AND_CONDITIONS = """TÉRMINOS Y CONDICIONES DE USO — CORREO CERTIFICADO
(Borrador — pendiente de revisión legal)

1. OBJETO DEL SERVICIO
Correo Certificado ("el Software") es una herramienta técnica que permite validar
si una dirección de correo electrónico tiene un formato correcto, si su dominio
existe, si cuenta con registros MX, y opcionalmente si el servidor de destino
confirma la existencia del buzón (verificación SMTP). El Software NO envía
correos electrónicos ni realiza campañas de comunicación de ningún tipo.

2. USO PERMITIDO
El usuario se compromete a utilizar el Software únicamente para verificar
direcciones de correo sobre las cuales tiene una base legítima de tratamiento
(por ejemplo: listas propias de clientes, registros con consentimiento, o
contactos obtenidos de forma lícita). Queda expresamente prohibido usar el
Software para:
  a) Validar listas de correos obtenidas de forma ilícita, robada o comprada
     sin consentimiento de los titulares.
  b) Facilitar el envío de correo no solicitado (spam) o campañas que violen
     leyes de protección de datos o anti-spam aplicables (incluyendo, sin
     limitarse a, la Ley Federal de Protección de Datos Personales en Posesión
     de los Particulares en México, el CAN-SPAM Act en EE.UU., o el RGPD en
     la Unión Europea, según corresponda al usuario).
  c) Cualquier actividad que infrinja derechos de terceros o la legislación
     vigente en la jurisdicción del usuario.

3. RESPONSABILIDAD DEL USUARIO
El usuario es el único responsable de:
  a) El origen, legalidad y consentimiento asociado a las direcciones de
     correo que decida verificar.
  b) El uso que dé a los resultados obtenidos del Software.
  c) Cumplir con toda ley, reglamento o normativa aplicable a su actividad,
     incluyendo las relativas a protección de datos personales y comunicaciones
     electrónicas.
Correo Certificado actúa únicamente como proveedor de una herramienta técnica
de validación y no participa, controla ni supervisa el uso que el usuario haga
de las direcciones verificadas ni de los resultados obtenidos.

4. PRECISIÓN DE LOS RESULTADOS
Los resultados del Software (incluyendo estados como "MX válido", "Accepted",
"Catch-All", etc.) son de carácter informativo y no constituyen una garantía
absoluta de entregabilidad real del correo. Factores fuera del control de
Correo Certificado (políticas de los servidores de destino, greylisting,
bloqueos temporales, cambios de configuración del dominio, entre otros) pueden
afectar la precisión del resultado en el momento de un envío real posterior.

5. LIMITACIÓN DE RESPONSABILIDAD
En la máxima medida permitida por la ley aplicable, Correo Certificado no será
responsable por daños directos, indirectos, incidentales o consecuentes
derivados de: (a) el uso indebido del Software por parte del usuario, (b)
decisiones tomadas con base en los resultados de verificación, o (c) el
incumplimiento por parte del usuario de leyes de protección de datos o
anti-spam aplicables a su actividad.

6. PRIVACIDAD Y MANEJO DE DATOS
Esta aplicación de escritorio realiza la verificación de forma local en el
equipo del usuario. Los correos individuales y sus resultados detallados NO
se transmiten ni almacenan en los servidores de Correo Certificado — solo se
reporta un conteo agregado por estado, necesario para descontar créditos de
la cuenta del usuario. Ver la Política de Privacidad completa para más detalle.

7. CRÉDITOS Y PAGOS
El acceso a la verificación se rige por un sistema de créditos, adquiribles
mediante planes de suscripción o paquetes de pago único. Los cargos se
procesan a través de Stripe, Inc. Las políticas de reembolso, en su caso, se
detallan por separado en la sección correspondiente de la plataforma web.

8. REQUISITOS TÉCNICOS
El uso de la verificación SMTP real requiere que el puerto 25 saliente esté
disponible en la red del usuario. Correo Certificado no controla ni garantiza
la disponibilidad de dicho puerto, que depende del proveedor de internet o de
la red del usuario.

9. MODIFICACIONES
Correo Certificado podrá actualizar estos Términos y Condiciones en cualquier
momento. El uso continuado del Software tras una actualización constituye la
aceptación de los nuevos términos.

10. LEGISLACIÓN APLICABLE
[Pendiente de definir con asesoría legal: país/jurisdicción y mecanismo de
resolución de controversias aplicable.]

Al marcar la casilla "Acepto los Términos y Condiciones", el usuario declara
haber leído, entendido y aceptado íntegramente este documento.
"""

# ---------------------------------------------------------------------------
# Toxicidad (portada completa del script original)
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
# DNS / MX  (con fallback a DNS públicos, igual que el script original)
# ---------------------------------------------------------------------------

MX_CACHE      = {}
MX_CACHE_LOCK = threading.Lock()


def get_mx_records(domain, timeout=6):
    """
    Resuelve registros MX usando nslookup (disponible en todo Windows sin
    dependencias externas). Esto evita los problemas de dnspython con
    PyInstaller. Prueba primero con el DNS del sistema, luego con 8.8.8.8.
    Devuelve: lista de hosts | None (NXDOMAIN) | [] (sin MX) | ("ERROR", detalle)
    """
    import subprocess, re as _re

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

        # Dominio no existe
        if any(x in output.lower() for x in [
            "non-existent domain", "nxdomain", "can't find",
            "no existe", "no encontrado"
        ]):
            result = None
            break

        # Extraer hosts MX de la salida de nslookup
        mx_lines = [l for l in output.splitlines()
                    if "mail exchanger" in l.lower() or "MX preference" in l]

        if mx_lines:
            hosts = []
            for line in mx_lines:
                # Formato: "domain MX preference = N, mail exchanger = host"
                # o: "mail exchanger = N host"
                parts = line.strip().split()
                if parts:
                    hosts.append(parts[-1].rstrip("."))
            if hosts:
                result = hosts
                break

        # Sin registros MX pero dominio existe
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


def check_port25_open(timeout=6):
    last_error = "No se pudo conectar a ningún servidor de prueba"
    for host in PORT25_TEST_HOSTS:
        try:
            with socket.create_connection((host, 25), timeout=timeout) as s:
                s.settimeout(timeout)
                banner = s.recv(256)
                if banner.startswith(b"220"):
                    return True, host
                last_error = f"Respuesta inesperada de {host}: {banner[:60]!r}"
        except socket.timeout:
            last_error = f"Timeout conectando a {host}"
        except ConnectionRefusedError:
            last_error = f"Conexión rechazada por {host}"
        except OSError as e:
            last_error = f"Error conectando a {host}: {e}"
    return False, last_error


# ---------------------------------------------------------------------------
# Verificación SMTP (lógica completa del script original)
# ---------------------------------------------------------------------------

def _random_probe_localpart():
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=14))


def smtp_check(mx_host, email, timeout):
    """
    Realiza HELO / MAIL FROM / RCPT TO y clasifica la respuesta.
    Estados posibles: Accepted, Greylisted, Limited, SPAM Block, Rejected, MX Error, Timeout
    """
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
        return {"email": email, "status": "Rejected", "detalle": "Formato inválido",
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

    # Intentar con cada servidor MX hasta obtener una respuesta no-error
    last_detail = "No se pudo conectar a ningún servidor MX"
    for mx_host in mx_result:
        status, detail = smtp_check(mx_host, email, SMTP_TIMEOUT)
        if status == "MX Error":
            last_detail = detail
            continue
        # Detección de catch-all
        if status == "Accepted":
            probe      = f"{_random_probe_localpart()}@{domain}"
            ps, _      = smtp_check(mx_host, probe, SMTP_TIMEOUT)
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
# GUI
# ---------------------------------------------------------------------------

# Colores por estado (igual que STATUS_COLORS del script original)
STATUS_COLORS_TK = {
    "Accepted":   "#3B7A57",   # verde
    "Catch-All":  "#1B7A8C",   # cian oscuro
    "Greylisted": "#B08D57",   # ámbar
    "Limited":    "#B08D57",
    "Timeout":    "#B08D57",
    "SPAM Block": "#8B3A8B",   # magenta
    "Rejected":   "#C1272D",   # rojo
    "MX Error":   "#C1272D",
    "No MX":      "#C1272D",
}


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Correo Certificado — Verificador de Escritorio")
        self.geometry("900x640")
        self.configure(bg="#12213B")

        self.user_email    = None
        self.auth_token    = None
        self.selected_file = None
        self.results       = []
        self.progress_queue = queue.Queue()
        self.port25_queue   = queue.Queue()
        self.port25_open    = None

        self._build_login_frame()
        self._build_main_frame()
        self.main_frame.pack_forget()

    # ── Login ──────────────────────────────────────────────────────────────

    def _build_login_frame(self):
        self.login_frame = tk.Frame(self, bg="#12213B", padx=40, pady=40)
        self.login_frame.pack(expand=True)

        tk.Label(self.login_frame, text="Correo Certificado",
                 font=("Segoe UI", 22, "bold"), fg="#F6EFE2", bg="#12213B").pack(pady=(0, 4))
        tk.Label(self.login_frame,
                 text="Verificador de escritorio — SMTP real, puerto 25 local",
                 font=("Segoe UI", 10), fg="#B08D57", bg="#12213B").pack(pady=(0, 24))

        tk.Label(self.login_frame, text="Email:", font=("Segoe UI", 11),
                 fg="#F6EFE2", bg="#12213B").pack(anchor="w")
        self.email_entry = tk.Entry(self.login_frame, font=("Consolas", 12), width=36)
        self.email_entry.pack(pady=(4, 10))

        tk.Label(self.login_frame, text="Contraseña:", font=("Segoe UI", 11),
                 fg="#F6EFE2", bg="#12213B").pack(anchor="w")
        self.pass_entry = tk.Entry(self.login_frame, font=("Consolas", 12),
                                    width=36, show="●")
        self.pass_entry.pack(pady=(4, 16))
        self.pass_entry.bind("<Return>", lambda e: self._login())

        # ── Requisito: puerto 25 abierto en esta red ────────────────────────
        self.login_port25_open = None  # None = probando, True/False = resultado
        port_frame = tk.Frame(self.login_frame, bg="#12213B")
        port_frame.pack(fill="x", pady=(0, 10))
        self.login_port_label = tk.Label(
            port_frame, text="Verificando puerto 25 de tu red…",
            font=("Consolas", 9, "bold"), fg="#B08D57", bg="#12213B", wraplength=340, justify="left")
        self.login_port_label.pack(anchor="w")
        tk.Button(port_frame, text="Volver a probar", command=self._test_port25_login,
                  bg="#1B3055", fg="white", relief="flat", font=("Segoe UI", 8),
                  padx=8, pady=3).pack(anchor="w", pady=(4, 0))

        # ── Requisito: aceptar Términos y Condiciones ───────────────────────
        self.terms_accepted = tk.BooleanVar(value=False)
        self.terms_accepted.trace_add("write", lambda *a: self._update_login_btn_state())
        terms_frame = tk.Frame(self.login_frame, bg="#12213B")
        terms_frame.pack(fill="x", pady=(4, 16))
        tk.Checkbutton(
            terms_frame, variable=self.terms_accepted, bg="#12213B",
            activebackground="#12213B", selectcolor="#1B3055",
        ).pack(side="left")
        tk.Label(terms_frame, text="Acepto los", font=("Segoe UI", 9),
                 fg="#F6EFE2", bg="#12213B").pack(side="left")
        terms_link = tk.Label(terms_frame, text="Términos y Condiciones",
                               font=("Segoe UI", 9, "underline"),
                               fg="#B08D57", bg="#12213B", cursor="hand2")
        terms_link.pack(side="left", padx=(4, 0))
        terms_link.bind("<Button-1>", lambda e: self._show_terms())

        btn_frame = tk.Frame(self.login_frame, bg="#12213B")
        btn_frame.pack()
        self.login_btn = tk.Button(btn_frame, text="Entrar", command=self._login,
                  bg="#C1272D", fg="white", font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=16, pady=8, state="disabled")
        self.login_btn.pack(side="left")
        tk.Button(btn_frame, text="Crear cuenta", command=self._open_register,
                  bg="#1B3055", fg="white", font=("Segoe UI", 10),
                  relief="flat", padx=12, pady=8).pack(side="left", padx=(10, 0))

        self.login_status = tk.Label(self.login_frame, text="",
                                      font=("Segoe UI", 10), fg="#C1272D", bg="#12213B")
        self.login_status.pack(pady=(12, 0))

        # Intentar restaurar sesión guardada
        self._restore_session()
        # Probar el puerto 25 en cuanto se abre la pantalla de login
        self._test_port25_login()

    def _show_terms(self):
        win = tk.Toplevel(self)
        win.title("Términos y Condiciones")
        win.geometry("640x520")
        win.configure(bg="#12213B")
        text = tk.Text(win, wrap="word", bg="#F6EFE2", fg="#12213B",
                        font=("Segoe UI", 10), padx=14, pady=14)
        text.insert("1.0", TERMS_AND_CONDITIONS)
        text.config(state="disabled")
        text.pack(fill="both", expand=True, padx=10, pady=10)
        tk.Button(win, text="Cerrar", command=win.destroy,
                  bg="#1B3055", fg="white", relief="flat", padx=12, pady=6).pack(pady=(0, 10))

    def _test_port25_login(self):
        """Prueba el puerto 25 desde la pantalla de login. Es un requisito para
        poder entrar: si está bloqueado, el usuario no podrá verificar nada
        real con esta app, así que se le avisa ANTES de loguearse."""
        self.login_port25_open = None
        self.login_port_label.config(text="Verificando puerto 25 de tu red…", fg="#B08D57")
        self.login_btn.config(state="disabled")

        q = queue.Queue()

        def worker():
            q.put(check_port25_open())

        threading.Thread(target=worker, daemon=True).start()

        def poll():
            try:
                is_open, info = q.get_nowait()
            except queue.Empty:
                self.after(200, poll)
                return
            self.login_port25_open = is_open
            if is_open:
                self.login_port_label.config(
                    text=f"✅ Puerto 25 abierto (respondió {info}). Puedes iniciar sesión.",
                    fg="#3B7A57")
            else:
                self.login_port_label.config(
                    text=f"❌ Puerto 25 bloqueado en esta red — {info}\n"
                         "No podrás verificar correos por SMTP real desde aquí. "
                         "Prueba desde otra red o contacta a tu proveedor de internet.",
                    fg="#C1272D")
            self._update_login_btn_state()

        self.after(200, poll)

    def _update_login_btn_state(self):
        can_login = bool(self.login_port25_open) and self.terms_accepted.get()
        self.login_btn.config(state="normal" if can_login else "disabled")

    def _open_register(self):
        import webbrowser
        webbrowser.open("https://verificador-emails.vercel.app")

    def _restore_session(self):
        """Intenta restaurar el token guardado en el Credential Manager de Windows."""
        try:
            token = keyring.get_password("CorreoCertificado", "access_token")
            email = keyring.get_password("CorreoCertificado", "user_email")
            if token and email:
                self.login_status.config(text="Restaurando sesión...", fg="#B08D57")
                self.update()
                res = requests.get(f"{API_BASE}/api/user/me",
                                    headers={"Authorization": f"Bearer {token}"},
                                    timeout=15)
                if res.ok:
                    user_data = res.json()
                    self._after_login(email, token, user_data["credits"], user_data.get("plan"))
                    return
        except Exception:
            pass
        try:
            keyring.delete_password("CorreoCertificado", "access_token")
            keyring.delete_password("CorreoCertificado", "user_email")
        except Exception:
            pass

    def _login(self):
        if not self.login_port25_open:
            self.login_status.config(
                text="No puedes iniciar sesión: el puerto 25 debe estar abierto en tu red.",
                fg="#C1272D")
            return
        if not self.terms_accepted.get():
            self.login_status.config(
                text="Debes aceptar los Términos y Condiciones para continuar.",
                fg="#C1272D")
            return
        email    = self.email_entry.get().strip().lower()
        password = self.pass_entry.get()
        if not email or "@" not in email:
            self.login_status.config(text="Ingresa un email válido")
            return
        if not password:
            self.login_status.config(text="Ingresa tu contraseña")
            return
        self.login_status.config(text="Conectando...", fg="#B08D57")
        self.update()
        try:
            res = requests.post(f"{API_BASE}/api/auth/login",
                                data={"email": email, "password": password},
                                timeout=20)
            if not res.ok:
                detail = res.json().get("detail", res.text)
                self.login_status.config(text=f"Error: {detail}", fg="#C1272D")
                return
            data  = res.json()
            token = data["access_token"]
            # Guardar token de forma segura
            try:
                keyring.set_password("CorreoCertificado", "access_token", token)
                keyring.set_password("CorreoCertificado", "user_email", data["email"])
            except Exception:
                pass
            self._after_login(data["email"], token, data["credits"], data.get("plan"))
        except Exception as e:
            self.login_status.config(text=f"No se pudo conectar: {e}", fg="#C1272D")

    def _after_login(self, email, token, credits, plan=None):
        self.user_email = email
        self.auth_token = token
        self.login_frame.pack_forget()
        self.main_frame.pack(fill="both", expand=True)
        self._refresh_credits(credits)
        self._refresh_plan(plan)
        self._test_port25()

    # ── Main ───────────────────────────────────────────────────────────────

    def _build_main_frame(self):
        self.main_frame = tk.Frame(self, bg="#12213B", padx=20, pady=16)

        # Cabecera
        top = tk.Frame(self.main_frame, bg="#12213B")
        top.pack(fill="x", pady=(0, 4))
        tk.Label(top, text="Correo Certificado", font=("Segoe UI", 16, "bold"),
                 fg="#F6EFE2", bg="#12213B").pack(side="left")

        tk.Button(top, text="Cerrar sesión", command=self._logout,
                  bg="#1B3055", fg="white", relief="flat", padx=10, pady=6,
                  font=("Segoe UI", 9)).pack(side="right", padx=(8, 0))
        tk.Button(top, text="Historial", command=self._open_history,
                  bg="#1B3055", fg="white", relief="flat", padx=10, pady=6,
                  font=("Segoe UI", 9)).pack(side="right", padx=(8, 0))
        tk.Button(top, text="Comprar créditos", command=self._open_buy_credits,
                  bg="#B08D57", fg="white", relief="flat", padx=10, pady=6,
                  font=("Segoe UI", 9, "bold")).pack(side="right", padx=(8, 0))

        self.credits_label = tk.Label(top, text="Créditos: —",
                                       font=("Consolas", 12, "bold"),
                                       fg="#C1272D", bg="#F6EFE2", padx=12, pady=6)
        self.credits_label.pack(side="right")

        self.plan_label = tk.Label(top, text="Sin plan activo",
                                    font=("Segoe UI", 10, "bold"),
                                    fg="#12213B", bg="#B08D57", padx=10, pady=6)
        self.plan_label.pack(side="right", padx=(0, 8))

        # Aviso de privacidad
        tk.Label(self.main_frame,
                 text="100% local: los emails y resultados individuales nunca salen de esta "
                      "computadora. Solo se reporta un conteo agregado para descontar créditos.",
                 font=("Segoe UI", 9, "italic"), fg="#B08D57", bg="#12213B",
                 wraplength=840, justify="left").pack(anchor="w", pady=(0, 8))

        # Estado puerto 25
        port_frame = tk.Frame(self.main_frame, bg="#12213B")
        port_frame.pack(fill="x", pady=(0, 10))
        self.port_status_label = tk.Label(port_frame, text="Puerto 25: sin probar",
                                           font=("Consolas", 10, "bold"),
                                           fg="#B08D57", bg="#12213B")
        self.port_status_label.pack(side="left")
        tk.Button(port_frame, text="Probar puerto 25", command=self._test_port25,
                  bg="#1B3055", fg="white", relief="flat", padx=10, pady=4,
                  font=("Segoe UI", 9)).pack(side="left", padx=(12, 0))

        # Selector de archivo
        file_frame = tk.Frame(self.main_frame, bg="#F6EFE2", padx=16, pady=12)
        file_frame.pack(fill="x", pady=(0, 10))
        self.file_label = tk.Label(file_frame, text="Ningún archivo seleccionado",
                                    font=("Consolas", 10), fg="#12213B", bg="#F6EFE2")
        self.file_label.pack(side="left")
        tk.Button(file_frame, text="Elegir archivo (.csv / .xlsx)",
                  command=self._choose_file,
                  bg="#12213B", fg="white", relief="flat",
                  padx=10, pady=6).pack(side="right")

        # Botones de acción
        action_frame = tk.Frame(self.main_frame, bg="#12213B")
        action_frame.pack(fill="x", pady=(0, 10))
        self.verify_btn = tk.Button(action_frame, text="Verificar (SMTP real)",
                                     command=self._start_verification, state="disabled",
                                     bg="#C1272D", fg="white",
                                     font=("Segoe UI", 11, "bold"),
                                     relief="flat", padx=14, pady=8)
        self.verify_btn.pack(side="left")
        self.export_btn = tk.Button(action_frame, text="Exportar CSV",
                                     command=self._export_csv, state="disabled",
                                     bg="#B08D57", fg="white", relief="flat",
                                     padx=14, pady=8)
        self.export_btn.pack(side="left", padx=(10, 0))

        # Barra de progreso y estado
        self.progress = ttk.Progressbar(self.main_frame, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 4))

        # Panel de estadísticas en vivo
        stats_frame = tk.Frame(self.main_frame, bg="#1B3055", padx=12, pady=8)
        stats_frame.pack(fill="x", pady=(0, 6))

        # Fila 1: progreso + tiempo
        row1 = tk.Frame(stats_frame, bg="#1B3055")
        row1.pack(fill="x")
        self.lbl_progreso   = tk.Label(row1, text="—", font=("Consolas", 10, "bold"),
                                        fg="#F6EFE2", bg="#1B3055")
        self.lbl_progreso.pack(side="left")
        self.lbl_velocidad  = tk.Label(row1, text="", font=("Consolas", 10),
                                        fg="#B08D57", bg="#1B3055")
        self.lbl_velocidad.pack(side="left", padx=(16, 0))
        self.lbl_tiempo     = tk.Label(row1, text="", font=("Consolas", 10),
                                        fg="#B08D57", bg="#1B3055")
        self.lbl_tiempo.pack(side="right")
        self.lbl_eta        = tk.Label(row1, text="", font=("Consolas", 10),
                                        fg="#B08D57", bg="#1B3055")
        self.lbl_eta.pack(side="right", padx=(0, 16))

        # Fila 2: contadores por estado
        row2 = tk.Frame(stats_frame, bg="#1B3055")
        row2.pack(fill="x", pady=(4, 0))
        self.lbl_counts = {}
        for status, color in [
            ("Accepted", "#3B7A57"), ("Catch-All", "#1B7A8C"),
            ("Rejected", "#C1272D"), ("No MX",     "#C1272D"),
            ("SPAM Block","#8B3A8B"), ("Timeout",  "#B08D57"),
            ("Greylisted","#B08D57"), ("MX Error", "#C1272D"),
        ]:
            lbl = tk.Label(row2, text=f"{status}: 0",
                           font=("Consolas", 9), fg=color, bg="#1B3055", padx=6)
            lbl.pack(side="left")
            self.lbl_counts[status] = lbl

        self.status_label = tk.Label(self.main_frame, text="",
                                      font=("Consolas", 10), fg="#B08D57", bg="#12213B")
        self.status_label.pack(anchor="w", pady=(0, 4))

        # Tabla de resultados
        columns = ("email", "status", "detalle", "toxicidad")
        self.tree = ttk.Treeview(self.main_frame, columns=columns,
                                  show="headings", height=15)
        for col, w in zip(columns, (280, 110, 320, 80)):
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=w)

        # Tags de color por estado
        for status, color in STATUS_COLORS_TK.items():
            self.tree.tag_configure(status, foreground=color)

        vsb = ttk.Scrollbar(self.main_frame, orient="vertical",
                             command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

    # ── Créditos / puerto 25 ───────────────────────────────────────────────

    def _refresh_credits(self, credits):
        self.credits_label.config(text=f"Créditos: {credits}")

    def _refresh_plan(self, plan):
        if plan:
            self.plan_label.config(text=f"Plan {plan.capitalize()}", bg="#3B7A57", fg="white")
        else:
            self.plan_label.config(text="Sin plan activo", bg="#B08D57", fg="#12213B")

    def _logout(self):
        """Cierra sesión: borra el token guardado y regresa a la pantalla de login."""
        try:
            keyring.delete_password("CorreoCertificado", "access_token")
            keyring.delete_password("CorreoCertificado", "user_email")
        except Exception:
            pass
        self.user_email    = None
        self.auth_token    = None
        self.selected_file = None
        self.results       = []
        self.main_frame.pack_forget()
        self._refresh_plan(None)
        self.email_entry.delete(0, "end")
        self.pass_entry.delete(0, "end")
        self.login_status.config(text="")
        self.login_frame.pack(expand=True)

    def _open_buy_credits(self):
        """
        Abre en el navegador la página de planes/créditos de la web.
        La compra en sí se hace ahí con Stripe Checkout — este botón ya
        queda funcional en cuanto esa pantalla esté publicada, sin tener
        que tocar la app de escritorio de nuevo.
        """
        webbrowser.open(f"{WEB_APP_URL}/#planes")

    def _open_history(self):
        """
        Ventana con dos pestañas:
          - 'Servidor': conteo agregado de todos tus lotes (igual en cualquier
            dispositivo donde inicies sesión).
          - 'Este equipo': el detalle completo (cada email, su resultado) que
            se guarda automáticamente aquí después de cada verificación —
            nunca se sube al servidor, así que solo existe en esta máquina.
        """
        win = tk.Toplevel(self)
        win.title("Historial de lotes")
        win.geometry("780x480")
        win.configure(bg="#12213B")

        nb = ttk.Notebook(win)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Pestaña: historial del servidor (solo conteos agregados) ---
        server_tab = tk.Frame(nb, bg="#12213B")
        nb.add(server_tab, text="Servidor (conteos)")

        cols = ("filename", "total", "fecha")
        server_tree = ttk.Treeview(server_tab, columns=cols, show="headings", height=16)
        for col, w, txt in [("filename", 320, "Archivo"), ("total", 90, "Emails"), ("fecha", 200, "Fecha")]:
            server_tree.heading(col, text=txt)
            server_tree.column(col, width=w)
        server_tree.pack(fill="both", expand=True)

        try:
            res = requests.get(
                f"{API_BASE}/api/jobs/history",
                headers={"Authorization": f"Bearer {self.auth_token}"},
                timeout=15,
            )
            res.raise_for_status()
            for job in res.json():
                server_tree.insert("", "end", values=(
                    job.get("filename", "—"),
                    job.get("total_emails", "—"),
                    job.get("created_at", "—"),
                ))
        except Exception as e:
            tk.Label(server_tab, text=f"No se pudo cargar el historial: {e}",
                      fg="#C1272D", bg="#12213B").pack(pady=8)

        # --- Pestaña: historial local (detalle completo, descargable) ---
        local_tab = tk.Frame(nb, bg="#12213B", padx=10, pady=10)
        nb.add(local_tab, text="Este equipo (detalle completo)")

        tk.Label(
            local_tab,
            text="Estos archivos viven únicamente en esta computadora — "
                 "puedes abrirlos o copiarlos cuando quieras.",
            fg="#B08D57", bg="#12213B", wraplength=700, justify="left",
        ).pack(anchor="w", pady=(0, 8))

        list_frame = tk.Frame(local_tab, bg="#12213B")
        list_frame.pack(fill="both", expand=True)

        local_list = tk.Listbox(list_frame, bg="#F6EFE2", font=("Consolas", 10))
        local_list.pack(side="left", fill="both", expand=True)
        vsb2 = ttk.Scrollbar(list_frame, orient="vertical", command=local_list.yview)
        local_list.configure(yscrollcommand=vsb2.set)
        vsb2.pack(side="left", fill="y")

        os.makedirs(LOCAL_HISTORY_DIR, exist_ok=True)
        files = sorted(pathlib.Path(LOCAL_HISTORY_DIR).glob("*.csv"), reverse=True)
        for f in files:
            local_list.insert("end", f.name)
        if not files:
            local_list.insert("end", "(aún no hay lotes verificados desde esta computadora)")

        def _open_selected():
            sel = local_list.curselection()
            if not sel or not files:
                return
            idx = sel[0]
            if idx >= len(files):
                return
            path = str(files[idx])
            try:
                os.startfile(path)  # Windows
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")

        def _open_folder():
            os.makedirs(LOCAL_HISTORY_DIR, exist_ok=True)
            try:
                os.startfile(LOCAL_HISTORY_DIR)  # Windows
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo abrir la carpeta: {e}")

        btn_row = tk.Frame(local_tab, bg="#12213B")
        btn_row.pack(fill="x", pady=(8, 0))
        tk.Button(btn_row, text="Abrir archivo seleccionado", command=_open_selected,
                  bg="#1B3055", fg="white", relief="flat", padx=10, pady=6).pack(side="left")
        tk.Button(btn_row, text="Abrir carpeta", command=_open_folder,
                  bg="#1B3055", fg="white", relief="flat", padx=10, pady=6).pack(side="left", padx=(8, 0))

    def _test_port25(self):
        self.port_status_label.config(text="Puerto 25: probando...", fg="#B08D57")
        self.port25_open = None

        def worker():
            result = check_port25_open()
            self.port25_queue.put(result)

        threading.Thread(target=worker, daemon=True).start()
        self.after(200, self._poll_port25)

    def _poll_port25(self):
        try:
            is_open, info = self.port25_queue.get_nowait()
        except queue.Empty:
            self.after(200, self._poll_port25)
            return
        self.port25_open = is_open
        if is_open:
            self.port_status_label.config(
                text=f"Puerto 25: abierto ✅ (respondió {info})", fg="#3B7A57")
        else:
            self.port_status_label.config(
                text=f"Puerto 25: bloqueado ❌ — {info}", fg="#C1272D")

    # ── Selección de archivo ───────────────────────────────────────────────

    def _choose_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV / Excel", "*.csv *.xlsx *.xls")])
        if path:
            self.selected_file = path
            self.file_label.config(text=os.path.basename(path))
            self.verify_btn.config(state="normal")

    # ── Verificación ───────────────────────────────────────────────────────

    def _start_verification(self):
        if not self.selected_file or not self.user_email:
            return

        if self.port25_open is False:
            if not messagebox.askyesno(
                "Puerto 25 bloqueado",
                "La prueba indica que el puerto 25 está bloqueado en esta red.\n\n"
                "Si continúas, la mayoría de emails saldrán como 'MX Error' y de "
                "todas formas se descontarán créditos.\n\n¿Continuar de todas formas?"
            ):
                return
        elif self.port25_open is None:
            if not messagebox.askyesno(
                "Puerto 25 no probado",
                "No se ha comprobado el puerto 25.\n\n"
                "Se recomienda usar el botón 'Probar puerto 25' primero.\n\n"
                "¿Continuar de todas formas?"
            ):
                return

        self.verify_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.results       = []
        self._t_start      = None   # se setea cuando llega el primer resultado
        self._live_counts  = {s: 0 for s in self.lbl_counts}
        self._total_emails = 0

        # Resetear panel de stats
        self.lbl_progreso.config(text="Iniciando...")
        self.lbl_velocidad.config(text="")
        self.lbl_tiempo.config(text="")
        self.lbl_eta.config(text="")
        for s, lbl in self.lbl_counts.items():
            lbl.config(text=f"{s}: 0")

        emails = extract_emails_from_file(self.selected_file)
        if not emails:
            messagebox.showerror("Error",
                                  "No se encontraron emails válidos en el archivo")
            self.verify_btn.config(state="normal")
            return

        try:
            res = requests.get(f"{API_BASE}/api/user/me",
                               headers={"Authorization": f"Bearer {self.auth_token}"},
                               timeout=15)
            res.raise_for_status()
            credits = res.json()["credits"]
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo consultar créditos: {e}")
            self.verify_btn.config(state="normal")
            return

        if credits < len(emails):
            messagebox.showerror("Créditos insuficientes",
                                  f"Necesitas {len(emails)} créditos, tienes {credits}.")
            self.verify_btn.config(state="normal")
            return

        self._total_emails = len(emails)
        self.progress.config(maximum=len(emails), value=0)
        self.status_label.config(
            text=f"Verificando {len(emails)} emails por SMTP real...")

        threading.Thread(target=self._run_verification,
                          args=(emails,), daemon=True).start()
        self.after(100, self._poll_progress)

    def _run_verification(self, emails):
        done = 0
        with ThreadPoolExecutor(max_workers=WORKERS) as pool:
            futures = {pool.submit(verify_email_local, e): e for e in emails}
            for future in as_completed(futures):
                result = future.result()
                self.results.append(result)
                done += 1
                self.progress_queue.put(("progress", done, len(emails), result))
        self.progress_queue.put(("done", None, None, None))

    def _poll_progress(self):
        import time as _time
        try:
            while True:
                kind, done, total, result = self.progress_queue.get_nowait()
                if kind == "progress":
                    # Iniciar timer en el primer resultado
                    if self._t_start is None:
                        self._t_start = _time.time()

                    self.progress.config(value=done)

                    # Tiempo transcurrido
                    elapsed   = _time.time() - self._t_start
                    vel       = done / elapsed if elapsed > 0 else 0
                    restantes = total - done
                    eta_seg   = restantes / vel if vel > 0 else 0

                    def _fmt(s):
                        s = int(s)
                        m, sec = divmod(s, 60)
                        return f"{m}m {sec:02d}s" if m else f"{sec}s"

                    pct = (done / total) * 100
                    self.lbl_progreso.config(
                        text=f"{done}/{total}  ({pct:.1f}%)")
                    self.lbl_velocidad.config(
                        text=f"⚡ {vel:.1f} emails/s")
                    self.lbl_tiempo.config(
                        text=f"⏱ {_fmt(elapsed)}")
                    self.lbl_eta.config(
                        text=f"ETA {_fmt(eta_seg)}" if done < total else "ETA —")

                    # Contador por estado
                    st = result["status"]
                    if st in self._live_counts:
                        self._live_counts[st] += 1
                        self.lbl_counts[st].config(
                            text=f"{st}: {self._live_counts[st]}")

                    # Insertar en tabla
                    tag = result["status"]
                    self.tree.insert("", "end",
                                      values=(result["email"], result["status"],
                                              result["detalle"], result["toxicidad"]),
                                      tags=(tag,))
                elif kind == "done":
                    self._finish_verification()
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll_progress)

    def _finish_verification(self):
        import time as _time
        elapsed = _time.time() - self._t_start if self._t_start else 0
        m, s    = divmod(int(elapsed), 60)
        tiempo  = f"{m}m {s:02d}s" if m else f"{s}s"
        total   = len(self.results)
        vel     = total / elapsed if elapsed > 0 else 0

        self.lbl_progreso.config(text=f"{total}/{total}  (100%)")
        self.lbl_velocidad.config(text=f"⚡ {vel:.1f} emails/s")
        self.lbl_tiempo.config(text=f"⏱ {tiempo}")
        self.lbl_eta.config(text="✅ Completado")
        self.status_label.config(
            text=f"Verificación completada — {total} emails en {tiempo}")

        status_counts = {}
        for r in self.results:
            status_counts[r["status"]] = status_counts.get(r["status"], 0) + 1

        try:
            res = requests.post(
                f"{API_BASE}/api/jobs/submit",
                data={
                    "filename":           os.path.basename(self.selected_file),
                    "total_emails":       len(self.results),
                    "status_counts_json": json.dumps(status_counts),
                },
                headers={"Authorization": f"Bearer {self.auth_token}"},
                timeout=30,
            )
            res.raise_for_status()
            data = res.json()
            self._refresh_credits(data["credits_restantes"])
            self.status_label.config(
                text=f"Listo. Créditos restantes: {data['credits_restantes']}")
        except Exception as e:
            messagebox.showwarning(
                "Aviso",
                f"La verificación terminó pero no se pudo reportar al servidor: {e}\n"
                "Tus resultados están abajo y puedes exportarlos."
            )

        self.verify_btn.config(state="normal")
        self.export_btn.config(state="normal")

        # Guardado automático local (detalle completo) — nunca sube al servidor.
        try:
            os.makedirs(LOCAL_HISTORY_DIR, exist_ok=True)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            base = os.path.splitext(os.path.basename(self.selected_file))[0]
            auto_path = os.path.join(LOCAL_HISTORY_DIR, f"{ts}_{base}.csv")
            with open(auto_path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(["email", "status", "detalle", "toxicidad", "señales_toxicidad"])
                for r in self.results:
                    writer.writerow([r["email"], r["status"], r["detalle"],
                                      r["toxicidad"], r["señales_toxicidad"]])
        except Exception:
            pass  # el guardado automático no debe interrumpir el flujo si falla

        messagebox.showinfo(
            "Recuerda exportar",
            "El detalle completo (cada email con su resultado) solo existe aquí, "
            "en esta computadora — no quedó copia en el servidor.\n\n"
            "Se guardó automáticamente en tu historial local (botón 'Historial' "
            "arriba). También puedes usar 'Exportar CSV' para guardarlo en otra ubicación."
        )

    # ── Exportar CSV ───────────────────────────────────────────────────────

    def _export_csv(self):
        if not self.results:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["email", "status", "detalle",
                              "toxicidad", "señales_toxicidad"])
            for r in self.results:
                writer.writerow([r["email"], r["status"], r["detalle"],
                                  r["toxicidad"], r["señales_toxicidad"]])
        messagebox.showinfo("Exportado", f"Guardado en:\n{path}")


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
