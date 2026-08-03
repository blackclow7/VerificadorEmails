package com.greenemailverifier.app

import android.app.Activity
import android.webkit.JavascriptInterface
import android.webkit.WebView
import org.json.JSONArray
import org.json.JSONObject
import org.xbill.DNS.Lookup
import org.xbill.DNS.MXRecord
import org.xbill.DNS.Type
import java.net.InetSocketAddress
import java.net.Socket
import java.util.concurrent.ExecutorService
import java.util.concurrent.Executors
import java.util.concurrent.Future

/**
 * AndroidBridge
 * =============
 * JavaScript Interface directo — NO depende del bridge de Capacitor.
 * Se expone como window.AndroidBridge en el WebView desde MainActivity.
 *
 * Se usó este enfoque tras confirmar (con diagnóstico en vivo) que
 * window.Capacitor.registerPlugin nunca queda disponible en este build,
 * a pesar de que window.Capacitor y el evento deviceready sí existen.
 *
 * Toda la lógica de verificación es autónoma — no reutiliza clases de
 * Capacitor (Plugin, PluginCall, JSObject) porque esas requieren un
 * contexto de plugin inicializado por el framework que no tenemos aquí.
 */
class AndroidBridge(private val activity: Activity, private val webView: WebView) {

    companion object {
        private const val HELO_DOMAIN  = "gmail.com"
        private const val FROM_ADDRESS = "verify@gmail.com"
        private const val TIMEOUT_MS   = 8000
        private const val WORKERS      = 6

        private val PORT25_TEST_HOSTS = arrayOf(
            "gmail-smtp-in.l.google.com",
            "smtp.mail.yahoo.com",
            "outlook-com.olc.protection.outlook.com"
        )

        private val DISPOSABLE_DOMAINS = setOf(
            "mailinator.com","10minutemail.com","guerrillamail.com","tempmail.com",
            "temp-mail.org","yopmail.com","trashmail.com","fakeinbox.com",
            "getnada.com","throwawaymail.com","maildrop.cc","sharklasers.com",
            "dispostable.com","mytemp.email","moakt.com","mohmal.com",
            "emailondeck.com","mintemail.com","spamgourmet.com","mailnesia.com",
            "correotemporal.org","tempinbox.com","burnermail.io",
            "guerrillamailblock.com","trbvm.com","spam4.me","tempail.com",
            "discard.email","mailcatch.com"
        )
        private val ROLE_PREFIXES = setOf(
            "admin","administrator","info","contact","contacto","sales","ventas",
            "support","soporte","noreply","no-reply","webmaster","postmaster",
            "abuse","help","ayuda","hello","hola","office","hr","rh",
            "marketing","billing","facturacion","compras","purchasing",
            "sistemas","it","sistema","notificaciones","notifications"
        )
        private val TYPO_DOMAINS = setOf(
            "gmial.com","gmal.com","gamil.com","gmail.co","gmailcom",
            "hotmial.com","hotmal.com","hotmai.com","hotmailcom",
            "yaho.com","yahooo.com","yahoo.co","yahoocom",
            "outlok.com","outloo.com","outlookcom","outlook.co"
        )
        private val EMAIL_REGEX = Regex("""^[^\s@]+@[^\s@]+\.[^\s@]+$""")
    }

    private val executor: ExecutorService = Executors.newFixedThreadPool(WORKERS + 2)
    private val mxCache = java.util.Collections.synchronizedMap(HashMap<String, List<String>?>())

    private fun runJs(js: String) {
        activity.runOnUiThread { webView.evaluateJavascript(js, null) }
    }

    // =========================================================================
    // shareXlsx — llamado desde JS: window.AndroidBridge.shareXlsx(base64Content, filename)
    // Decodifica el .xlsx (generado con SheetJS en el lado JS) desde base64,
    // lo guarda en el almacenamiento de la app y abre el selector nativo de
    // Android para compartir o guardar el archivo (Drive, WhatsApp, etc).
    // =========================================================================
    @JavascriptInterface
    fun shareXlsx(base64Content: String, filename: String) {
        activity.runOnUiThread {
            try {
                val bytes = android.util.Base64.decode(base64Content, android.util.Base64.DEFAULT)
                val dir = java.io.File(activity.cacheDir, "exports")
                if (!dir.exists()) dir.mkdirs()
                val file = java.io.File(dir, filename)
                file.writeBytes(bytes)

                val uri = androidx.core.content.FileProvider.getUriForFile(
                    activity,
                    "${activity.packageName}.fileprovider",
                    file
                )

                val shareIntent = android.content.Intent(android.content.Intent.ACTION_SEND).apply {
                    type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    putExtra(android.content.Intent.EXTRA_STREAM, uri)
                    addFlags(android.content.Intent.FLAG_GRANT_READ_URI_PERMISSION)
                }
                activity.startActivity(
                    android.content.Intent.createChooser(shareIntent, "Guardar o compartir Excel")
                )
            } catch (e: Exception) {
                runJs("window.onShareCsvError && window.onShareCsvError(${JSONObject.quote(e.message ?: "error")})")
            }
        }
    }

    // =========================================================================
    // checkPort25 — llamado desde JS: window.AndroidBridge.checkPort25(context)
    // Resultado vía window.onPort25Result({context, open, info})
    // =========================================================================
    @JavascriptInterface
    fun checkPort25(context: String) {
        executor.submit {
            var open = false
            var info = "No se pudo conectar a ningun servidor de prueba"
            for (host in PORT25_TEST_HOSTS) {
                try {
                    Socket().use { s ->
                        s.connect(InetSocketAddress(host, 25), TIMEOUT_MS)
                        s.soTimeout = TIMEOUT_MS
                        val buf = ByteArray(256)
                        val n = s.getInputStream().read(buf)
                        val banner = if (n > 0) String(buf, 0, n) else ""
                        if (banner.startsWith("220")) {
                            open = true
                            info = host
                        } else {
                            info = "Respuesta inesperada de $host"
                        }
                    }
                    if (open) break
                } catch (e: Exception) {
                    info = "Error conectando a $host: ${e.message}"
                }
            }
            val json = JSONObject()
                .put("context", context)
                .put("open", open)
                .put("info", info)
            runJs("window.onPort25Result(${json})")
        }
    }

    // =========================================================================
    // verifyBatch — llamado desde JS: window.AndroidBridge.verifyBatch(jsonArrayString)
    // Progreso vía window.onVerifyProgress({done,total,result})
    // Final vía window.onVerifyComplete([...])
    // =========================================================================
    @JavascriptInterface
    fun verifyBatch(emailsJson: String) {
        executor.submit {
            try {
                val arr = JSONArray(emailsJson)
                val emails = (0 until arr.length()).map { arr.getString(it) }
                val total = emails.size

                val futures: List<Future<JSONObject>> = emails.map { email ->
                    executor.submit<JSONObject> { verifyEmail(email) }
                }

                val results = mutableListOf<JSONObject>()
                var done = 0
                for (f in futures) {
                    val result = try { f.get() } catch (e: Exception) {
                        JSONObject().put("email","?").put("status","MX Error")
                            .put("detalle", e.message ?: "error").put("toxicidad", 0)
                    }
                    results.add(result)
                    done++
                    val progress = JSONObject()
                        .put("done", done)
                        .put("total", total)
                        .put("result", result)
                    runJs("window.onVerifyProgress(${progress})")
                }

                val all = JSONArray()
                results.forEach { all.put(it) }
                runJs("window.onVerifyComplete(${all})")
            } catch (e: Exception) {
                runJs("window.onVerifyComplete([])")
            }
        }
    }

    // =========================================================================
    // Verificación de un email — MX + SMTP + toxicidad
    // =========================================================================
    private fun verifyEmail(email: String): JSONObject {
        val (toxScore, toxReasons) = assessToxicity(email)

        if (!EMAIL_REGEX.matches(email)) {
            return result(email, "Formato inválido", "No cumple formato usuario@dominio", toxScore, toxReasons)
        }

        val domain = email.substringAfter("@").lowercase()
        val mxHosts = getMxRecords(domain)
            ?: return result(email, "No MX", "El dominio no existe (NXDOMAIN)", toxScore, toxReasons)

        if (mxHosts.isEmpty()) {
            return result(email, "No MX", "El dominio no tiene registros MX", toxScore, toxReasons)
        }

        var lastDetail = "No se pudo conectar a ningun servidor MX"
        for (mx in mxHosts) {
            val (status, detail) = smtpCheck(mx, email)
            if (status == "MX Error") { lastDetail = detail; continue }
            if (status == "Accepted") {
                val probe = "${randomLocalPart()}@$domain"
                val (ps, _) = smtpCheck(mx, probe)
                if (ps == "Accepted") {
                    return result(email, "Catch-All", "El dominio acepta cualquier direccion", toxScore, toxReasons)
                }
            }
            return result(email, status, detail, toxScore, toxReasons)
        }
        return result(email, "MX Error", lastDetail, toxScore, toxReasons)
    }

    private fun smtpCheck(mxHost: String, email: String): Pair<String, String> {
        return try {
            Socket().use { s ->
                s.connect(InetSocketAddress(mxHost, 25), TIMEOUT_MS)
                s.soTimeout = TIMEOUT_MS
                val inp = s.getInputStream()
                val out = s.getOutputStream()

                fun readLine(): String {
                    val sb = StringBuilder()
                    var prev = 0
                    while (true) {
                        val b = inp.read()
                        if (b == -1 || (prev == '\r'.code && b == '\n'.code)) break
                        if (b != '\r'.code) sb.append(b.toChar())
                        prev = b
                    }
                    return sb.toString()
                }
                fun readResponse(): Pair<Int, String> {
                    var code = 0; var msg = ""
                    while (true) {
                        val line = readLine()
                        if (line.length < 3) break
                        code = line.substring(0, 3).toIntOrNull() ?: 0
                        msg = if (line.length > 4) line.substring(4) else ""
                        if (line.length <= 3 || line[3] == ' ') break
                    }
                    return Pair(code, msg)
                }
                fun send(cmd: String) { out.write("$cmd\r\n".toByteArray()); out.flush() }

                val (bannerCode, _) = readResponse()
                if (bannerCode != 220) return@use Pair("MX Error", "Banner inesperado: $bannerCode")

                send("HELO $HELO_DOMAIN")
                val (heloCode, _) = readResponse()
                if (heloCode !in 200..299) return@use Pair("MX Error", "HELO rechazado: $heloCode")

                send("MAIL FROM:<$FROM_ADDRESS>")
                val (fromCode, _) = readResponse()
                if (fromCode !in 200..299) return@use Pair("MX Error", "MAIL FROM rechazado: $fromCode")

                send("RCPT TO:<$email>")
                val (rcptCode, rcptMsg) = readResponse()
                val low = rcptMsg.lowercase()
                try { send("QUIT"); readResponse() } catch (_: Exception) {}

                when {
                    rcptCode == 250 || rcptCode == 251 -> Pair("Accepted", rcptMsg)
                    rcptCode in 450..452 -> Pair("Greylisted", "Codigo SMTP $rcptCode")
                    rcptCode == 421 -> Pair("Limited", rcptMsg)
                    rcptCode in 550..554 -> {
                        if (listOf("spam","blocked","block","reputation","denied","blacklist").any { it in low })
                            Pair("SPAM Block", rcptMsg) else Pair("Rejected", rcptMsg)
                    }
                    else -> Pair("Rejected", "Codigo SMTP $rcptCode: $rcptMsg")
                }
            }
        } catch (e: java.net.SocketTimeoutException) {
            Pair("Timeout", "La conexion supero el tiempo limite")
        } catch (e: java.net.ConnectException) {
            Pair("MX Error", "Conexion rechazada por $mxHost")
        } catch (e: Exception) {
            Pair("MX Error", e.message ?: "error")
        }
    }

    // Android no expone /etc/resolv.conf como Linux/Windows, así que dnsjava
    // no puede autodetectar servidores DNS del sistema. Se configuran
    // explícitamente resolvers públicos (Google + Cloudflare) para que
    // Lookup.run() realmente pueda resolver los registros MX.
    private var resolversConfigured = false
    private fun ensureResolvers() {
        if (resolversConfigured) return
        try {
            val r1 = org.xbill.DNS.SimpleResolver("8.8.8.8")
            r1.setTimeout(java.time.Duration.ofSeconds(6))
            val r2 = org.xbill.DNS.SimpleResolver("1.1.1.1")
            r2.setTimeout(java.time.Duration.ofSeconds(6))
            val r3 = org.xbill.DNS.SimpleResolver("8.8.4.4")
            r3.setTimeout(java.time.Duration.ofSeconds(6))
            val extResolver = org.xbill.DNS.ExtendedResolver(arrayOf(r1, r2, r3))
            Lookup.setDefaultResolver(extResolver)
        } catch (e: Exception) {
            // Si falla la configuración, Lookup intentará su comportamiento por defecto
        }
        resolversConfigured = true
    }

    private fun getMxRecords(domain: String): List<String>? {
        mxCache[domain]?.let { return it }
        ensureResolvers()
        return try {
            val lookup = Lookup(domain, Type.MX)
            val records = lookup.run()
            when {
                lookup.result == Lookup.HOST_NOT_FOUND || lookup.result == Lookup.TYPE_NOT_FOUND -> {
                    mxCache[domain] = null; null
                }
                records == null || records.isEmpty() -> {
                    mxCache[domain] = emptyList(); emptyList()
                }
                else -> {
                    val sorted = records.filterIsInstance<MXRecord>()
                        .sortedBy { it.priority }
                        .map { it.target.toString().trimEnd('.') }
                    mxCache[domain] = sorted; sorted
                }
            }
        } catch (e: Exception) {
            mxCache[domain] = emptyList(); emptyList()
        }
    }

    private fun assessToxicity(email: String): Pair<Int, String> {
        if ("@" !in email) return Pair(0, "formato invalido, no evaluado")
        val local  = email.substringBefore("@").lowercase()
        val domain = email.substringAfter("@").lowercase()
        val reasons = mutableListOf<String>()
        var score = 0
        if (domain in DISPOSABLE_DOMAINS) { score += 3; reasons += "dominio desechable/temporal" }
        if (domain in TYPO_DOMAINS)        { score += 2; reasons += "typo de dominio conocido" }
        if (local  in ROLE_PREFIXES)       { score += 1; reasons += "direccion de rol generica" }
        if (looksRandom(local))            { score += 1; reasons += "local-part con patron aleatorio" }
        score = minOf(score, 5)
        return Pair(score, if (reasons.isEmpty()) "sin senales detectadas" else reasons.joinToString("; "))
    }

    private fun looksRandom(lp: String): Boolean {
        if (lp.length < 10) return false
        val clean = lp.replace(Regex("[._\\-]"), "")
        if (!clean.all { it.isLetterOrDigit() }) return false
        val digits  = clean.count { it.isDigit() }
        val letters = clean.count { it.isLetter() }
        val vowels  = clean.count { it in "aeiou" }
        return digits >= 4 && letters >= 4 && (vowels.toDouble() / maxOf(letters, 1)) < 0.25
    }

    private fun randomLocalPart(len: Int = 14) =
        (1..len).map { ('a'..'z').toList()[java.util.Random().nextInt(26)] }.joinToString("")

    private fun result(email: String, status: String, detalle: String, toxicidad: Int, senales: String) =
        JSONObject()
            .put("email", email)
            .put("status", status)
            .put("detalle", detalle)
            .put("toxicidad", toxicidad)
            .put("señales_toxicidad", senales)
}
