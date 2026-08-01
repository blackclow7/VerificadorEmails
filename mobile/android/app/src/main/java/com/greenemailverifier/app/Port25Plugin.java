package com.greenemailverifier.app;

import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

import java.io.InputStream;
import java.net.InetSocketAddress;
import java.net.Socket;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * Port25Plugin
 * ============
 * Prueba mínima: intenta abrir una conexión TCP cruda al puerto 25 de un
 * puñado de servidores de correo conocidos y revisa si responden con el
 * saludo SMTP estándar ("220 ..."). No envía ningún correo ni hace la
 * verificación completa (HELO/MAIL FROM/RCPT TO) — solo confirma si la red
 * del dispositivo permite salir por ese puerto.
 *
 * Se ejecuta en un hilo aparte porque Android prohíbe hacer llamadas de red
 * en el hilo principal (lanzaría NetworkOnMainThreadException).
 */
@CapacitorPlugin(name = "Port25")
public class Port25Plugin extends Plugin {

    private static final String[] TEST_HOSTS = {
        "gmail-smtp-in.l.google.com",
        "smtp.mail.yahoo.com",
        "outlook-com.olc.protection.outlook.com"
    };
    private static final int TIMEOUT_MS = 6000;

    private final ExecutorService executor = Executors.newSingleThreadExecutor();

    @PluginMethod
    public void check(PluginCall call) {
        executor.submit(() -> {
            String lastError = "No se pudo conectar a ningún servidor de prueba";
            for (String host : TEST_HOSTS) {
                try (Socket socket = new Socket()) {
                    socket.connect(new InetSocketAddress(host, 25), TIMEOUT_MS);
                    socket.setSoTimeout(TIMEOUT_MS);
                    InputStream in = socket.getInputStream();
                    byte[] buffer = new byte[256];
                    int read = in.read(buffer);
                    String banner = read > 0 ? new String(buffer, 0, read) : "";
                    if (banner.startsWith("220")) {
                        JSObject ret = new JSObject();
                        ret.put("open", true);
                        ret.put("info", host);
                        call.resolve(ret);
                        return;
                    }
                    lastError = "Respuesta inesperada de " + host + ": " + banner.trim();
                } catch (java.net.SocketTimeoutException e) {
                    lastError = "Timeout conectando a " + host;
                } catch (java.net.ConnectException e) {
                    lastError = "Conexión rechazada por " + host;
                } catch (Exception e) {
                    lastError = "Error conectando a " + host + ": " + e.getMessage();
                }
            }
            JSObject ret = new JSObject();
            ret.put("open", false);
            ret.put("info", lastError);
            call.resolve(ret);
        });
    }
}
