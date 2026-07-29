# Guía de despliegue — Fase 1 (100% gratis)

Este proyecto tiene 3 partes, todas desplegables gratis:

| Parte | Qué hace | Dónde va gratis |
|---|---|---|
| **Backend** (`/backend`) | API que verifica emails y controla créditos | Render (Free Web Service) |
| **Base de datos** | Usuarios, créditos, historial de lotes | Supabase (Free Postgres) |
| **Frontend** (`/frontend/index.html`) | Página web que sube el archivo y sube resultados | Vercel o Netlify (Free Static) |

---

## 1. Crear la base de datos (Supabase)

1. Ve a https://supabase.com → crea cuenta gratis → **New Project**.
2. Cuando esté listo, entra a **SQL Editor** → **New query**, pega el contenido de `backend/schema.sql` y dale **Run**.
3. Ve a **Project Settings → API**. Copia:
   - `Project URL` → esto es tu `SUPABASE_URL`
   - `service_role` key (no la `anon`) → esto es tu `SUPABASE_KEY`

   ⚠️ La `service_role` key tiene permisos totales. Nunca la pongas en el frontend, solo en el backend (variables de entorno de Render).

---

## 2. Desplegar el backend (Render)

1. Sube la carpeta `backend/` a un repositorio de GitHub (puede ser solo esa carpeta, o el proyecto completo).
2. Ve a https://render.com → crea cuenta gratis → **New → Web Service** → conecta tu repo.
3. Configuración:
   - **Root Directory:** `backend` (si subiste todo el proyecto junto)
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - **Instance Type:** Free
4. En **Environment**, agrega las variables de `backend/.env.example` con tus valores reales:
   - `SUPABASE_URL`
   - `SUPABASE_KEY`
   - `FRONTEND_ORIGIN` (déjalo en `*` por ahora, luego ponle la URL de tu Vercel)
   - `FREE_SIGNUP_CREDITS`, `MAX_EMAILS_PER_JOB`, `VERIFY_WORKERS`, `ENABLE_SMTP_CHECK=false`
5. Deploy. Al terminar tendrás una URL tipo `https://tu-app.onrender.com`.
6. Prueba en el navegador: `https://tu-app.onrender.com/api/health` → debe responder `{"status":"ok","supabase_configured":true}`.

**Nota sobre el tier gratis de Render:** el servicio "se duerme" tras ~15 min sin uso, y tarda ~30-50s en despertar en la siguiente petición. Es normal en fase de pruebas; para producción real conviene el plan pago ($7/mes) que no se duerme.

---

## 3. Desplegar el frontend (Vercel o Netlify)

1. Abre `frontend/index.html` y edita esta línea con la URL real de tu backend de Render:
   ```js
   const API_BASE = "https://tu-app.onrender.com";
   ```
2. Sube la carpeta `frontend/` a GitHub (o arrastra el archivo directo si usas Netlify Drop: https://app.netlify.com/drop).
3. En Vercel: **New Project** → importa el repo → Root Directory `frontend` → Deploy.
4. Copia la URL que te da Vercel/Netlify (ej. `https://tu-verificador.vercel.app`) y vuelve a Render para poner esa URL en `FRONTEND_ORIGIN` (así solo tu web puede llamar a tu API).

---

## 4. Probar todo

1. Abre tu URL de frontend.
2. Ingresa un email de prueba → deberías ver créditos gratis asignados.
3. Sube un `.csv` o `.xlsx` con emails de prueba (puedes usar el mismo formato que ya usabas con tu script local).
4. Verifica que los resultados aparezcan y que los créditos bajen.
5. Descarga el CSV de resultados.

---

## Qué SÍ hace esta Fase 1 (y qué no)

✅ Verifica: formato del email, si el dominio existe, si tiene registros MX, y un score de toxicidad (dominios desechables, typos comunes, direcciones de rol, patrones de spam-trap).

❌ NO confirma si el buzón específico existe de verdad (eso requiere conexión SMTP real por puerto 25, bloqueado en todo hosting gratuito). El estado que verás es `MX válido` en vez de `Accepted`/`Rejected`.

Esto es una limitación real de la infraestructura gratuita, no del código — tu script original (`email_verifier.py`) sigue teniendo la lógica SMTP completa portada en `verifier.py`, apagada por defecto (`ENABLE_SMTP_CHECK=false`). El día que quieras verificación real a nivel de buzón, la Fase 2 lo resuelve (ver abajo).

---

## Fase 2 — cuando quieras monetizar en serio

1. **Créditos con pago real:** reemplazar `/api/credits/add` (hoy es un botón de prueba sin cobro) por un webhook de Stripe Checkout. Stripe no cobra por integrarlo, solo comisión por transacción.
2. **Verificación SMTP real:** mover el backend (o solo la función de verificación) a un VPS barato (~$5-6/mes: DigitalOcean, Hetzner, Vultr) y solicitarles habilitar el puerto 25 saliente (todos los proveedores, incluso de pago, lo bloquean por defecto y hay que pedirlo por ticket). Ahí pones `ENABLE_SMTP_CHECK=true`.
3. **Procesamiento en segundo plano:** para lotes grandes (miles de emails), pasar de "espera síncrona" a una cola de trabajos (ej. Supabase + un worker separado) para no depender de los límites de tiempo de una sola petición HTTP.
4. **Autenticación real:** hoy el "login" es solo un email sin contraseña (suficiente para probar el producto). Antes de cobrar dinero real, conviene pasar a Supabase Auth (magic link o contraseña) para que nadie use créditos de otra persona con solo saber su correo.
