package com.greenemailverifier.app;

import android.webkit.JavascriptInterface;
import android.webkit.WebView;
import com.getcapacitor.BridgeActivity;

/**
 * AndroidBridge
 * =============
 * JavaScript Interface directo — no depende de Capacitor bridge.
 * Se expone como window.AndroidBridge en el WebView.
 * Delega la verificación SMTP al SmtpVerifierPlugin existente.
 */
public class AndroidBridge {

    private final BridgeActivity activity;
    private final WebView webView;
    private final SmtpVerifierPlugin smtpPlugin;
    private final Port25Plugin port25Plugin;

    public AndroidBridge(BridgeActivity activity, WebView webView) {
        this.activity = activity;
        this.webView = webView;
        this.smtpPlugin = new SmtpVerifierPlugin();
        this.port25Plugin = new Port25Plugin();
    }

    // Llamado desde JS: window.AndroidBridge.isReady()
    @JavascriptInterface
    public boolean isReady() {
        return true;
    }

    // Llamado desde JS: window.AndroidBridge.checkPort25(context)
    // Resultado llega vía window.onPort25Result(...)
    @JavascriptInterface
    public void checkPort25(final String context) {
        new Thread(() -> {
            String[] hosts = {
                "gmail-smtp-in.l.google.com",
                "smtp.mail.yahoo.com",
                "outlook-com.olc.protection.outlook.com"
            };
            boolean open = false;
            String info = "No se pudo conectar";
            for (String host : hosts) {
                try {
                    java.net.Socket s = new java.net.Socket();
                    s.connect(new java.net.InetSocketAddress(host, 25), 8000);
                    s.setSoTimeout(8000);
                    byte[] buf = new byte[256];
                    int n = s.getInputStream().read(buf);
                    String banner = n > 0 ? new String(buf, 0, n) : "";
                    s.close();
                    if (banner.startsWith("220")) {
                        open = true;
                        info = host;
                        break;
                    }
                } catch (Exception e) {
                    info = e.getMessage();
                }
            }
            final boolean finalOpen = open;
            final String finalInfo = info;
            activity.runOnUiThread(() -> {
                String js = "window.onPort25Result({context:'" + context +
                    "',open:" + finalOpen +
                    ",info:'" + finalInfo.replace("'", "\\'") + "'})";
                webView.evaluateJavascript(js, null);
            });
        }).start();
    }

    // Llamado desde JS: window.AndroidBridge.verifyBatch(emailsJson)
    // Resultados llegan vía window.onVerifyProgress y window.onVerifyComplete
    @JavascriptInterface
    public void verifyBatch(final String emailsJson) {
        new Thread(() -> {
            try {
                // Parsear array JSON simple
                String clean = emailsJson.trim();
                if (clean.startsWith("[")) clean = clean.substring(1);
                if (clean.endsWith("]")) clean = clean.substring(0, clean.length()-1);
                String[] parts = clean.split("\",\"");
                java.util.List<String> emails = new java.util.ArrayList<>();
                for (String p : parts) {
                    String e = p.replace("\"","").trim();
                    if (!e.isEmpty()) emails.add(e);
                }
                int total = emails.size();
                java.util.concurrent.ExecutorService pool =
                    java.util.concurrent.Executors.newFixedThreadPool(6);
                java.util.List<java.util.concurrent.Future<String>> futures = new java.util.ArrayList<>();
                for (String email : emails) {
                    final String e = email;
                    futures.add(pool.submit(() -> smtpPlugin.verifyEmailPublic(e)));
                }
                java.util.List<String> results = new java.util.ArrayList<>();
                int done = 0;
                for (java.util.concurrent.Future<String> f : futures) {
                    String resultJson = f.get();
                    results.add(resultJson);
                    done++;
                    final int d = done;
                    final String rj = resultJson;
                    activity.runOnUiThread(() -> {
                        String js = "window.onVerifyProgress({done:" + d +
                            ",total:" + total + ",result:" + rj + "})";
                        webView.evaluateJavascript(js, null);
                    });
                }
                pool.shutdown();
                StringBuilder all = new StringBuilder("[");
                for (int i = 0; i < results.size(); i++) {
                    all.append(results.get(i));
                    if (i < results.size()-1) all.append(",");
                }
                all.append("]");
                final String allJson = all.toString();
                activity.runOnUiThread(() ->
                    webView.evaluateJavascript("window.onVerifyComplete(" + allJson + ")", null));
            } catch (Exception ex) {
                final String err = ex.getMessage();
                activity.runOnUiThread(() ->
                    webView.evaluateJavascript("window.onVerifyComplete([])", null));
            }
        }).start();
    }
}
