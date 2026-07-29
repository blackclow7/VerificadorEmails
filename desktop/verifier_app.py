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
    4. Al terminar, envía solo el resumen de resultados al backend para
       descontar créditos y guardarlo en el historial (visible en la web).

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
import dns.resolver
import openpyxl

# ---------------------------------------------------------------------------
# Configuración — apunta esto a tu backend real desplegado
# ---------------------------------------------------------------------------
API_BASE = "https://verificadoremails-api.onrender.com"

FROM_ADDRESS = "verify@tudominio.com"   # cámbialo por un dominio que controles
HELO_DOMAIN = "tudominio.com"
SMTP_TIMEOUT = 8
WORKERS = 8   # verificación local: no conviene tanta concurrencia como en un servidor

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ---------------------------------------------------------------------------
# Lógica de verificación (adaptada de email_verifier.py / verifier.py)
# ---------------------------------------------------------------------------

DISPOSABLE_DOMAINS = {
    "mailinator.com", "10minutemail.com", "guerrillamail.com", "tempmail.com",
    "temp-mail.org", "yopmail.com", "trashmail.com", "fakeinbox.com",
    "getnada.com", "throwawaymail.com", "maildrop.cc", "sharklasers.com",
    "dispostable.com", "mytemp.email", "moakt.com", "mohmal.com",
}
ROLE_PREFIXES = {
    "admin", "administrator", "info", "contact", "contacto", "sales", "ventas",
    "support", "soporte", "noreply", "no-reply", "webmaster", "postmaster",
}
TYPO_DOMAINS = {
    "gmial.com", "gmal.com", "gamil.com", "hotmial.com", "hotmal.com",
    "yaho.com", "yahooo.com", "outlok.com", "outloo.com",
}


def assess_toxicity(email):
    if "@" not in email:
        return 0, "formato inválido"
    local, domain = email.split("@", 1)
    local, domain = local.lower(), domain.lower()
    score, reasons = 0, []
    if domain in DISPOSABLE_DOMAINS:
        score += 3; reasons.append("dominio desechable")
    if domain in TYPO_DOMAINS:
        score += 2; reasons.append("typo de dominio conocido")
    if local in ROLE_PREFIXES:
        score += 1; reasons.append("dirección de rol")
    return min(score, 5), ("; ".join(reasons) if reasons else "sin señales")


def get_mx_records(domain, timeout=6):
    try:
        resolver = dns.resolver.Resolver()
        resolver.timeout = timeout
        resolver.lifetime = timeout
        answers = resolver.resolve(domain, "MX")
        hosts = sorted([(r.preference, str(r.exchange).rstrip(".")) for r in answers])
        return [h for _, h in hosts]
    except dns.resolver.NXDOMAIN:
        return None
    except Exception:
        return []


# Servidores MX conocidos y estables, usados solo para probar si el puerto
# 25 saliente está abierto en la red del usuario (no se envía nada, solo
# se abre y cierra la conexión).
PORT25_TEST_HOSTS = [
    "gmail-smtp-in.l.google.com",
    "smtp.mail.yahoo.com",
    "outlook-com.olc.protection.outlook.com",
]


def check_port25_open(timeout=6):
    """
    Prueba si el puerto 25 saliente está abierto, intentando conectar (sin
    enviar nada) a un par de servidores MX grandes y estables.
    Devuelve (True, host) si al menos uno respondió, o (False, motivo) si no.
    """
    last_error = "No se pudo conectar a ningún servidor de prueba"
    for host in PORT25_TEST_HOSTS:
        try:
            with socket.create_connection((host, 25), timeout=timeout) as s:
                s.settimeout(timeout)
                banner = s.recv(256)  # el servidor SMTP saluda primero (código 220)
                if banner.startswith(b"220"):
                    return True, host
                last_error = f"Respuesta inesperada de {host}: {banner[:60]!r}"
        except socket.timeout:
            last_error = f"Timeout conectando a {host} (probable bloqueo del ISP/firewall)"
        except ConnectionRefusedError:
            last_error = f"Conexión rechazada por {host} (el puerto 25 está bloqueado)"
        except OSError as e:
            last_error = f"No se pudo conectar a {host}: {e}"
    return False, last_error


def _random_probe_localpart():
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=12))


def smtp_check(mx_host, email, timeout):
    try:
        with smtplib.SMTP(mx_host, 25, timeout=timeout) as server:
            server.ehlo(HELO_DOMAIN)
            if server.has_extn("starttls"):
                try:
                    server.starttls()
                    server.ehlo(HELO_DOMAIN)
                except Exception:
                    pass
            server.mail(FROM_ADDRESS)
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
        return "SPAM Block", "El servidor cortó la conexión"
    except (ConnectionRefusedError, OSError) as e:
        return "MX Error", f"No se pudo conectar (¿tu red bloquea el puerto 25? {e})"
    except Exception as e:
        return "MX Error", str(e)


def verify_email_local(email):
    email = email.strip()
    tox_score, tox_reasons = assess_toxicity(email)

    if not EMAIL_REGEX.match(email):
        return {"email": email, "status": "Formato inválido", "detalle": "No cumple formato",
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}

    domain = email.split("@", 1)[1].lower()
    mx_hosts = get_mx_records(domain)

    if mx_hosts is None:
        return {"email": email, "status": "Dominio inexistente", "detalle": "NXDOMAIN",
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}
    if not mx_hosts:
        return {"email": email, "status": "Sin MX", "detalle": "Sin registros MX",
                "toxicidad": tox_score, "señales_toxicidad": tox_reasons}

    status, detail = "MX Error", "No se pudo conectar a ningún servidor MX"
    for mx_host in mx_hosts:
        s, d = smtp_check(mx_host, email, SMTP_TIMEOUT)
        if s == "MX Error":
            continue
        status, detail = s, d
        break

    if status == "Accepted":
        probe = f"{_random_probe_localpart()}@{domain}"
        for mx_host in mx_hosts:
            ps, _ = smtp_check(mx_host, probe, SMTP_TIMEOUT)
            if ps == "Accepted":
                status, detail = "Catch-All", "El dominio acepta cualquier dirección"
            break

    return {"email": email, "status": status, "detalle": detail,
            "toxicidad": tox_score, "señales_toxicidad": tox_reasons}


# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------

def extract_emails_from_file(path):
    emails = []
    if path.lower().endswith(".csv"):
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            reader = csv.reader(f)
            for row in reader:
                for cell in row:
                    cell = cell.strip()
                    if "@" in cell and "." in cell.split("@")[-1]:
                        emails.append(cell)
                        break
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

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Correo Certificado — Verificador de Escritorio")
        self.geometry("820x600")
        self.configure(bg="#12213B")

        self.user_email = None
        self.selected_file = None
        self.results = []
        self.progress_queue = queue.Queue()
        self.port25_queue = queue.Queue()
        self.port25_open = None

        self._build_login_frame()
        self._build_main_frame()
        self.main_frame.pack_forget()

    # -- Login --
    def _build_login_frame(self):
        self.login_frame = tk.Frame(self, bg="#12213B", padx=40, pady=40)
        self.login_frame.pack(expand=True)

        tk.Label(self.login_frame, text="Correo Certificado", font=("Segoe UI", 22, "bold"),
                 fg="#F6EFE2", bg="#12213B").pack(pady=(0, 4))
        tk.Label(self.login_frame, text="Verificador de escritorio (SMTP real, puerto 25 propio)",
                 font=("Segoe UI", 10), fg="#B08D57", bg="#12213B").pack(pady=(0, 24))

        tk.Label(self.login_frame, text="Tu email:", font=("Segoe UI", 11),
                 fg="#F6EFE2", bg="#12213B").pack(anchor="w")
        self.email_entry = tk.Entry(self.login_frame, font=("Consolas", 12), width=35)
        self.email_entry.pack(pady=(4, 16))

        self.login_btn = tk.Button(self.login_frame, text="Entrar / Crear cuenta",
                                    command=self._login, bg="#C1272D", fg="white",
                                    font=("Segoe UI", 11, "bold"), relief="flat", padx=16, pady=8)
        self.login_btn.pack()

        self.login_status = tk.Label(self.login_frame, text="", font=("Segoe UI", 10),
                                      fg="#C1272D", bg="#12213B")
        self.login_status.pack(pady=(12, 0))

    def _login(self):
        email = self.email_entry.get().strip().lower()
        if not email or "@" not in email:
            self.login_status.config(text="Ingresa un email válido")
            return
        self.login_status.config(text="Conectando...", fg="#B08D57")
        self.update()
        try:
            res = requests.post(f"{API_BASE}/api/user/login", data={"email": email}, timeout=15)
            res.raise_for_status()
            data = res.json()
            self.user_email = data["email"]
            self.login_frame.pack_forget()
            self.main_frame.pack(fill="both", expand=True)
            self._refresh_credits(data["credits"])
            self._test_port25()
        except Exception as e:
            self.login_status.config(text=f"No se pudo conectar: {e}", fg="#C1272D")

    # -- Main --
    def _build_main_frame(self):
        self.main_frame = tk.Frame(self, bg="#12213B", padx=20, pady=20)

        top = tk.Frame(self.main_frame, bg="#12213B")
        top.pack(fill="x", pady=(0, 16))
        tk.Label(top, text="Correo Certificado", font=("Segoe UI", 16, "bold"),
                 fg="#F6EFE2", bg="#12213B").pack(side="left")
        tk.Label(top, text="", bg="#12213B").pack(side="left", expand=True)  # spacer
        self.credits_label = tk.Label(top, text="Créditos: —", font=("Consolas", 12, "bold"),
                                       fg="#C1272D", bg="#F6EFE2", padx=12, pady=6)
        self.credits_label.pack(side="right")

        tk.Label(self.main_frame,
                 text="100% local: los emails y resultados nunca salen de esta computadora. "
                      "Solo se reporta un conteo agregado para descontar créditos.",
                 font=("Segoe UI", 9, "italic"), fg="#B08D57", bg="#12213B",
                 wraplength=760, justify="left").pack(anchor="w", pady=(0, 10))

        port_frame = tk.Frame(self.main_frame, bg="#12213B")
        port_frame.pack(fill="x", pady=(0, 12))
        self.port_status_label = tk.Label(port_frame, text="Puerto 25: sin probar",
                                           font=("Consolas", 10, "bold"), fg="#B08D57", bg="#12213B")
        self.port_status_label.pack(side="left")
        tk.Button(port_frame, text="Probar puerto 25", command=self._test_port25,
                  bg="#1B3055", fg="white", relief="flat", padx=10, pady=4,
                  font=("Segoe UI", 9)).pack(side="left", padx=(12, 0))

        file_frame = tk.Frame(self.main_frame, bg="#F6EFE2", padx=16, pady=16)
        file_frame.pack(fill="x", pady=(0, 12))
        self.file_label = tk.Label(file_frame, text="Ningún archivo seleccionado",
                                    font=("Consolas", 10), fg="#12213B", bg="#F6EFE2")
        self.file_label.pack(side="left")
        tk.Button(file_frame, text="Elegir archivo (.csv / .xlsx)", command=self._choose_file,
                  bg="#12213B", fg="white", relief="flat", padx=10, pady=6).pack(side="right")

        action_frame = tk.Frame(self.main_frame, bg="#12213B")
        action_frame.pack(fill="x", pady=(0, 12))
        self.verify_btn = tk.Button(action_frame, text="Verificar (SMTP real)",
                                     command=self._start_verification, state="disabled",
                                     bg="#C1272D", fg="white", font=("Segoe UI", 11, "bold"),
                                     relief="flat", padx=14, pady=8)
        self.verify_btn.pack(side="left")
        self.export_btn = tk.Button(action_frame, text="Exportar CSV", command=self._export_csv,
                                     state="disabled", bg="#B08D57", fg="white", relief="flat",
                                     padx=14, pady=8)
        self.export_btn.pack(side="left", padx=(10, 0))

        self.progress = ttk.Progressbar(self.main_frame, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 12))

        self.status_label = tk.Label(self.main_frame, text="", font=("Consolas", 10),
                                      fg="#B08D57", bg="#12213B")
        self.status_label.pack(anchor="w", pady=(0, 8))

        columns = ("email", "status", "detalle", "toxicidad")
        self.tree = ttk.Treeview(self.main_frame, columns=columns, show="headings", height=16)
        for col, w in zip(columns, (260, 130, 300, 90)):
            self.tree.heading(col, text=col.capitalize())
            self.tree.column(col, width=w)
        self.tree.pack(fill="both", expand=True)

    def _refresh_credits(self, credits):
        self.credits_label.config(text=f"Créditos: {credits}")

    def _test_port25(self):
        self.port_status_label.config(text="Puerto 25: probando...", fg="#B08D57")
        self.port25_open = None

        def worker():
            is_open, info = check_port25_open()
            self.port25_queue.put((is_open, info))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_port25)

    def _poll_port25(self):
        try:
            is_open, info = self.port25_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_port25)
            return
        self.port25_open = is_open
        if is_open:
            self.port_status_label.config(
                text=f"Puerto 25: abierto ✅ (respondió {info})", fg="#3B7A57")
        else:
            self.port_status_label.config(
                text="Puerto 25: bloqueado ❌ — la verificación real no funcionará en esta red",
                fg="#C1272D")

    def _choose_file(self):
        path = filedialog.askopenfilename(filetypes=[("CSV / Excel", "*.csv *.xlsx *.xls")])
        if path:
            self.selected_file = path
            self.file_label.config(text=os.path.basename(path))
            self.verify_btn.config(state="normal")

    def _start_verification(self):
        if not self.selected_file or not self.user_email:
            return

        if self.port25_open is False:
            proceed = messagebox.askyesno(
                "Puerto 25 bloqueado",
                "La prueba de conexión indica que el puerto 25 está bloqueado en esta red.\n\n"
                "Si continúas, todos los emails van a salir como 'MX Error' (no se podrá "
                "confirmar el buzón) y de todas formas se descontarán créditos.\n\n"
                "¿Quieres continuar de todas formas?"
            )
            if not proceed:
                return
        elif self.port25_open is None:
            proceed = messagebox.askyesno(
                "Puerto 25 no probado",
                "Aún no se ha probado si el puerto 25 está abierto en esta red.\n\n"
                "¿Quieres continuar sin probarlo? (recomendado: cancela y usa el botón "
                "'Probar puerto 25' primero)"
            )
            if not proceed:
                return

        self.verify_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.tree.delete(*self.tree.get_children())
        self.results = []

        emails = extract_emails_from_file(self.selected_file)
        if not emails:
            messagebox.showerror("Error", "No se encontraron emails válidos en el archivo")
            self.verify_btn.config(state="normal")
            return

        # Verifica créditos antes de gastar tiempo verificando
        try:
            res = requests.get(f"{API_BASE}/api/user/{self.user_email}", timeout=15)
            res.raise_for_status()
            credits = res.json()["credits"]
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo consultar créditos: {e}")
            self.verify_btn.config(state="normal")
            return

        if credits < len(emails):
            messagebox.showerror(
                "Créditos insuficientes",
                f"Necesitas {len(emails)} créditos, tienes {credits}."
            )
            self.verify_btn.config(state="normal")
            return

        self.progress.config(maximum=len(emails), value=0)
        self.status_label.config(text=f"Verificando {len(emails)} emails por SMTP real...")

        thread = threading.Thread(target=self._run_verification, args=(emails,), daemon=True)
        thread.start()
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
        try:
            while True:
                kind, done, total, result = self.progress_queue.get_nowait()
                if kind == "progress":
                    self.progress.config(value=done)
                    self.status_label.config(text=f"Verificando... {done}/{total}")
                    self.tree.insert("", "end", values=(
                        result["email"], result["status"], result["detalle"], result["toxicidad"]
                    ))
                elif kind == "done":
                    self._finish_verification()
                    return
        except queue.Empty:
            pass
        self.after(100, self._poll_progress)

    def _finish_verification(self):
        self.status_label.config(text=f"Listo. {len(self.results)} emails verificados. Enviando resumen...")

        # Solo se manda un CONTEO agregado por estado — nunca los emails ni
        # sus resultados individuales. Eso se queda 100% en esta computadora.
        status_counts = {}
        for r in self.results:
            status_counts[r["status"]] = status_counts.get(r["status"], 0) + 1

        try:
            res = requests.post(
                f"{API_BASE}/api/jobs/submit",
                data={
                    "email": self.user_email,
                    "filename": os.path.basename(self.selected_file),
                    "total_emails": len(self.results),
                    "status_counts_json": json.dumps(status_counts),
                },
                timeout=30,
            )
            res.raise_for_status()
            data = res.json()
            self._refresh_credits(data["credits_restantes"])
            self.status_label.config(text=f"Listo. Créditos restantes: {data['credits_restantes']}")
        except Exception as e:
            messagebox.showwarning(
                "Aviso",
                f"La verificación terminó pero no se pudo reportar al servidor: {e}\n"
                "Tus resultados igual quedaron abajo y puedes exportarlos."
            )
        self.verify_btn.config(state="normal")
        self.export_btn.config(state="normal")
        messagebox.showinfo(
            "Recuerda exportar",
            "El detalle completo (cada email con su resultado) solo existe aquí, en esta "
            "computadora — no quedó copia en el servidor. Usa 'Exportar CSV' ahora si quieres "
            "conservarlo."
        )

    def _export_csv(self):
        if not self.results:
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV", "*.csv")])
        if not path:
            return
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["email", "status", "detalle", "toxicidad", "señales_toxicidad"])
            for r in self.results:
                writer.writerow([r["email"], r["status"], r["detalle"], r["toxicidad"], r["señales_toxicidad"]])
        messagebox.showinfo("Exportado", f"Guardado en:\n{path}")


if __name__ == "__main__":
    app = App()
    app.mainloop()
