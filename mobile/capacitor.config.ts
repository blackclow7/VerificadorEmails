import type { CapacitorConfig } from '@capacitor/cli';

const config: CapacitorConfig = {
  appId: 'com.greenemailverifier.app',
  appName: 'Green Email Verifier',
  webDir: 'www'
  // Para la versión de prueba de puerto 25, cargamos la página local (www/)
  // donde el plugin nativo Port25 sí puede ser llamado desde JavaScript.
  // Una vez confirmado que el puerto funciona, restauramos `server.url`
  // apuntando a la web en Vercel y construimos el plugin SMTP completo.
};

export default config;
