# Cómo compilar la app de escritorio a un .exe de Windows

## 1. Configura la URL de tu backend

Antes de compilar, abre `verifier_app.py` y edita esta línea con la URL real de tu backend en Render:

```python
API_BASE = "https://testeointenso.onrender.com"
```

También puedes ajustar (opcional, pero recomendado):

```python
FROM_ADDRESS = "verify@tudominio.com"   # idealmente un dominio que controles tú
HELO_DOMAIN = "tudominio.com"
```

Usar un dominio propio en vez de un genérico ayuda a que los servidores de destino confíen un poco más en la conexión (menos probabilidad de "SPAM Block").

## 2. Compilar (esto se hace en una PC con Windows)

Necesitas Python instalado en la PC donde compiles (no en la del usuario final — el `.exe` ya incluye todo).

```bash
# Instala dependencias
pip install requests dnspython openpyxl pyinstaller

# Compila a un solo .exe, sin consola visible (ventana solo con la GUI)
pyinstaller --onefile --windowed --name VerificadorEmails verifier_app.py
```

El ejecutable queda en `dist/VerificadorEmails.exe`. Ese es el archivo que le compartes a tus usuarios (por descarga desde tu web, por ejemplo).

## 3. Notas importantes

- **Antivirus/SmartScreen:** los `.exe` generados con PyInstaller y sin firma digital suelen disparar advertencias de Windows Defender/SmartScreen ("Windows protegió su PC"). Es normal para software nuevo sin reputación aún — no significa que tenga virus. Para reducir esto a futuro, se puede firmar el ejecutable con un certificado de firma de código (tiene costo, ~$100-300 USD/año), lo cual además genera confianza inmediata en vez de la advertencia.
- **Puerto 25 en la red del usuario:** como hablamos, algunos ISPs residenciales también bloquean el puerto 25 saliente. Si un usuario no puede verificar, es casi seguro esta la causa — no hay nada que la app pueda hacer en ese caso salvo mostrar el error "MX Error: ¿tu red bloquea el puerto 25?" (ya está contemplado en el código).
- **Actualizaciones:** si cambias `verifier_app.py`, tienes que volver a compilar y redistribuir el `.exe`. Este MVP no tiene auto-actualización — para fase 2, se podría añadir una verificación de versión contra tu API al iniciar.
