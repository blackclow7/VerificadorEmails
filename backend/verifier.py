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
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import openpyxl

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------
API_BASE = "https://verificadoremails-api.onrender.com"

FROM_ADDRESS = "verify@tudominio.com"
HELO_DOMAIN  = "tudominio.com"
SMTP_TIMEOUT = 10
WORKERS      = 8

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

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

        tk.Label(self.login_frame, text="Tu email:", font=("Segoe UI", 11),
                 fg="#F6EFE2", bg="#12213B").pack(anchor="w")
        self.email_entry = tk.Entry(self.login_frame, font=("Consolas", 12), width=36)
        self.email_entry.pack(pady=(4, 16))
        self.email_entry.bind("<Return>", lambda e: self._login())

        tk.Button(self.login_frame, text="Entrar / Crear cuenta", command=self._login,
                  bg="#C1272D", fg="white", font=("Segoe UI", 11, "bold"),
                  relief="flat", padx=16, pady=8).pack()

        self.login_status = tk.Label(self.login_frame, text="",
                                      font=("Segoe UI", 10), fg="#C1272D", bg="#12213B")
        self.login_status.pack(pady=(12, 0))

    def _login(self):
        email = self.email_entry.get().strip().lower()
        if not email or "@" not in email:
            self.login_status.config(text="Ingresa un email válido")
            return
        self.login_status.config(text="Conectando...", fg="#B08D57")
        self.update()
        try:
            res = requests.post(f"{API_BASE}/api/user/login",
                                data={"email": email}, timeout=20)
            res.raise_for_status()
            data = res.json()
            self.user_email = data["email"]
            self.login_frame.pack_forget()
            self.main_frame.pack(fill="both", expand=True)
            self._refresh_credits(data["credits"])
            self._test_port25()
        except Exception as e:
            self.login_status.config(text=f"No se pudo conectar: {e}", fg="#C1272D")

    # ── Main ───────────────────────────────────────────────────────────────

    def _build_main_frame(self):
        self.main_frame = tk.Frame(self, bg="#12213B", padx=20, pady=16)

        # Cabecera
        top = tk.Frame(self.main_frame, bg="#12213B")
        top.pack(fill="x", pady=(0, 4))
        tk.Label(top, text="Correo Certificado", font=("Segoe UI", 16, "bold"),
                 fg="#F6EFE2", bg="#12213B").pack(side="left")
        self.credits_label = tk.Label(top, text="Créditos: —",
                                       font=("Consolas", 12, "bold"),
                                       fg="#C1272D", bg="#F6EFE2", padx=12, pady=6)
        self.credits_label.pack(side="right")

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
            res = requests.get(f"{API_BASE}/api/user/{self.user_email}", timeout=15)
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
                    "email":              self.user_email,
                    "filename":           os.path.basename(self.selected_file),
                    "total_emails":       len(self.results),
                    "status_counts_json": json.dumps(status_counts),
                },
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
        messagebox.showinfo(
            "Recuerda exportar",
            "El detalle completo (cada email con su resultado) solo existe aquí, "
            "en esta computadora — no quedó copia en el servidor.\n\n"
            "Usa 'Exportar CSV' si quieres conservarlo."
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
