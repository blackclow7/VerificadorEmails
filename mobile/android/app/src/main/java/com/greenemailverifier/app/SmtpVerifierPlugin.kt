package com.greenemailverifier.app

import com.getcapacitor.JSArray
import com.getcapacitor.JSObject
import com.getcapacitor.Plugin
import com.getcapacitor.PluginCall
import com.getcapacitor.PluginMethod
import com.getcapacitor.annotation.CapacitorPlugin
import org.xbill.DNS.Lookup
import org.xbill.DNS.MXRecord
import org.xbill.DNS.Type
import java.io.InputStream
import java.net.InetSocketAddress
import java.net.Socket
import java.util.concurrent.ExecutorService
import java.util.concurrent.Executors
import java.util.concurrent.Future
import java.util.concurrent.TimeUnit

/**
 * SmtpVerifierPlugin
 * ==================
 * Plugin Capacitor que replica en Android toda la lógica de verificación de
 * emails que en escritorio hace verifier_app.py / verifier.py.
 *
 * Métodos expuestos a JavaScript:
 *
 *   checkPort25({ context })
 *     Prueba si el puerto 25 saliente está abierto en la red actual.
 *     Resultado asíncrono vía window.onPort25Result({ context, open, info }).
 *
 *   verifyBatch({ emails: string[] })
 *     Verifica una lista de emails con DNS MX + SMTP RCPT TO + análisis de
 *     toxicidad. Emite progreso en tiempo real vía:
 *       window.onVerifyProgress({ done, total, result })
 *     y al terminar:
 *       window.onVerifyComplete(results[])
 *
 * Todo el trabajo de red ocurre en hilos de fondo (Android prohíbe red en
 * el hilo principal). Los resultados se devuelven llamando evaluateJavascript
 * desde el bridge de Capacitor, igual que pywebview hace window.evaluate_js.
 */

@CapacitorPlugin(name = "SmtpVerifier")
class SmtpVerifierPlugin : Plugin() {

    // ── Configuración ────────────────────────────────────────────────────────

    companion object {
        private const val HELO_DOMAIN   = "gmail.com"
        private const val FROM_ADDRESS  = "verify@gmail.com"
        private const val TIMEOUT_MS    = 8_000
        private const val WORKERS       = 6   // hilos paralelos para verifyBatch

        // Mismos hosts de prueba que en Python / Port25Plugin.java
        private val PORT25_TEST_HOSTS = arrayOf(
            "gmail-smtp-in.l.google.com",
            "smtp.mail.yahoo.com",
            "outlook-com.olc.protection.outlook.com"
        )

        // ── Toxicidad ────────────────────────────────────────────────────────
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

    // Pool de hilos para no tocar el hilo principal de UI
    private val executor: ExecutorService = Executors.newFixedThreadPool(WORKERS + 2)

    // Caché MX por dominio (hilo-seguro con synchronizedMap)
    private val mxCache = java.util.Collections.synchronizedMap(HashMap<String, List<String>?>())

    // =========================================================================
    // checkPort25
    // =========================================================================

    @PluginMethod
    fun checkPort25(call: PluginCall) {
        val context = call.getString("context", "default")!!
        call.resolve(JSObject().put("started", true))   // respuesta inmediata al JS

        executor.submit {
            var open = false
            var info = "No se pudo conectar a ningún servidor de prueba"

            // Prueba los 3 hosts EN PARALELO (igual que Python)
            val futures: List<Future<Pair<Boolean, String>>> = PORT25_TEST_HOSTS.map { host ->
                executor.submit<Pair<Boolean, String>> { tryPort25(host) }
            }
            outer@ for (f in futures) {
                try {
                    val (ok, msg) = f.get(TIMEOUT_MS.toLong() + 1000, TimeUnit.MILLISECONDS)
                    if (ok) { open = true; info = msg; break@outer }
                    info = msg
                } catch (_: Exception) { /* continúa */ }
            }
            futures.forEach { it.cancel(true) }

            // Empujar resultado al JS igual que Python: window.onPort25Result(...)
            val payload = JSObject()
                .put("context", context)
                .put("open", open)
                .put("info", info)
            notifyJs("onPort25Result", payload)
        }
    }

    private fun tryPort25(host: String): Pair<Boolean, String> {
        return try {
            Socket().use { s ->
                s.connect(InetSocketAddress(host, 25), TIMEOUT_MS)
                s.soTimeout = TIMEOUT_MS
                val buf = ByteArray(256)
                val n = s.getInputStream().read(buf)
                val banner = if (n > 0) String(buf, 0, n) else ""
                if (banner.startsWith("220"))
                    Pair(true, host)
                else
                    Pair(false, "Respuesta inesperada de $host: ${banner.take(60).trim()}")
            }
        } catch (e: java.net.SocketTimeoutException) {
            Pair(false, "Timeout conectando a $host")
        } catch (e: java.net.ConnectException) {
            Pair(false, "Conexión rechazada por $host")
        } catch (e: Exception) {
            Pair(false, "Error conectando a $host: ${e.message}")
        }
    }

    // =========================================================================
    // verifyBatch
    // =========================================================================

    @PluginMethod
    fun verifyBatch(call: PluginCall) {
        val emailsArr = call.getArray("emails") ?: run {
            call.reject("Se requiere 'emails' (array de strings)")
            return
        }
        val emails = (0 until emailsArr.length()).mapNotNull { emailsArr.getString(it)?.trim() }
        val total = emails.size
        call.resolve(JSObject().put("started", true).put("total", total))

        executor.submit {
            val results = java.util.Collections.synchronizedList(mutableListOf<JSObject>())
            var done = 0

            // Verifica en paralelo con WORKERS hilos, igual que Python
            val futures: List<java.util.concurrent.Future<JSObject>> = emails.map { email ->
                executor.submit<JSObject> { verifyEmail(email) }
            }
            for (f in futures) {
                val result: JSObject = try {
                    f.get()
                } catch (e: Exception) {
                    JSObject().put("email", "?").put("status", "MX Error")
                        .put("detalle", e.message ?: "error").put("toxicidad", 0) as JSObject
                }
                results.add(result)
                done++
                // Progreso en vivo → window.onVerifyProgress(...)
                val progress = JSObject()
                    .put("done", done)
                    .put("total", total)
                    .put("result", result)
                notifyJs("onVerifyProgress", progress)
            }

            // Resultado final → window.onVerifyComplete([...])
            val all = JSArray()
            results.forEach { all.put(it) }
            notifyJs("onVerifyComplete", all)
        }
    }

    // =========================================================================
    // Lógica de verificación (equivalente a verify_email_local de Python)
    // =========================================================================

    private fun verifyEmail(email: String): JSObject {
        val (toxScore, toxReasons) = assessToxicity(email)

        if (!EMAIL_REGEX.matches(email)) {
            return result(email, "Rejected", "Formato inválido", toxScore, toxReasons)
        }

        val domain = email.substringAfter("@").lowercase()
        val mxHosts = getMxRecords(domain)
            ?: return result(email, "No MX", "El dominio no existe (NXDOMAIN)", toxScore, toxReasons)

        if (mxHosts.isEmpty()) {
            return result(email, "No MX", "El dominio no tiene registros MX", toxScore, toxReasons)
        }

        var lastDetail = "No se pudo conectar a ningún servidor MX"
        for (mx in mxHosts) {
            val (status, detail) = smtpCheck(mx, email)
            if (status == "MX Error") { lastDetail = detail; continue }
            if (status == "Accepted") {
                // Prueba catch-all con dirección aleatoria
                val probe = "${randomLocalPart()}@$domain"
                val (ps, _) = smtpCheck(mx, probe)
                if (ps == "Accepted") {
                    return result(email, "Catch-All", "El dominio acepta cualquier dirección", toxScore, toxReasons)
                }
            }
            return result(email, status, detail, toxScore, toxReasons)
        }
        return result(email, "MX Error", lastDetail, toxScore, toxReasons)
    }

    // =========================================================================
    // SMTP: HELO / MAIL FROM / RCPT TO (equivalente a smtp_check de Python)
    // =========================================================================

    private fun smtpCheck(mxHost: String, email: String): Pair<String, String> {
        return try {
            Socket().use { s ->
                s.connect(InetSocketAddress(mxHost, 25), TIMEOUT_MS)
                s.soTimeout = TIMEOUT_MS
                val inp: InputStream = s.getInputStream()
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
                    // Lee todas las líneas del banner (puede ser multilínea: "250-...\r\n250 ...")
                    var code = 0; var msg = ""
                    while (true) {
                        val line = readLine()
                        if (line.length < 3) break
                        code = line.substring(0, 3).toIntOrNull() ?: 0
                        msg = if (line.length > 4) line.substring(4) else ""
                        if (line.length <= 3 || line[3] == ' ') break  // última línea
                    }
                    return Pair(code, msg)
                }

                fun send(cmd: String) {
                    out.write("$cmd\r\n".toByteArray())
                    out.flush()
                }

                // Banner 220
                val (bannerCode, _) = readResponse()
                if (bannerCode != 220) return@use Pair("MX Error", "Banner inesperado: $bannerCode")

                // HELO
                send("HELO $HELO_DOMAIN")
                val (heloCode, _) = readResponse()
                if (heloCode !in 200..299) return@use Pair("MX Error", "HELO rechazado: $heloCode")

                // MAIL FROM
                send("MAIL FROM:<$FROM_ADDRESS>")
                val (fromCode, _) = readResponse()
                if (fromCode !in 200..299) return@use Pair("MX Error", "MAIL FROM rechazado: $fromCode")

                // RCPT TO — aquí está la verificación real
                send("RCPT TO:<$email>")
                val (rcptCode, rcptMsg) = readResponse()
                val low = rcptMsg.lowercase()

                try { send("QUIT"); readResponse() } catch (_: Exception) {}

                when {
                    rcptCode == 250 || rcptCode == 251 -> Pair("Accepted", rcptMsg)
                    rcptCode in 450..452              -> Pair("Greylisted", "Código SMTP $rcptCode")
                    rcptCode == 421                   -> Pair("Limited", rcptMsg)
                    rcptCode in 550..554 -> {
                        if (listOf("spam","blocked","block","reputation","denied","blacklist")
                                .any { it in low })
                            Pair("SPAM Block", rcptMsg)
                        else
                            Pair("Rejected", rcptMsg)
                    }
                    else -> Pair("Rejected", "Código SMTP $rcptCode: $rcptMsg")
                }
            }
        } catch (e: java.net.SocketTimeoutException) {
            Pair("Timeout", "La conexión superó el tiempo límite")
        } catch (e: java.net.ConnectException) {
            Pair("MX Error", "Conexión rechazada por $mxHost")
        } catch (e: Exception) {
            val low = e.message?.lowercase() ?: ""
            if (listOf("spam","blocked","reputation","blacklist").any { it in low })
                Pair("SPAM Block", e.message ?: "error")
            else
                Pair("MX Error", e.message ?: "error")
        }
    }

    // =========================================================================
    // Resolución MX con dnsjava (equivalente a get_mx_records de Python)
    // =========================================================================

    private fun getMxRecords(domain: String): List<String>? {
        mxCache[domain]?.let { return it }

        return try {
            val lookup = Lookup(domain, Type.MX)
            val records = lookup.run()
            when {
                lookup.result == Lookup.HOST_NOT_FOUND ||
                lookup.result == Lookup.TYPE_NOT_FOUND -> {
                    mxCache[domain] = null
                    null   // NXDOMAIN → dominio no existe
                }
                records == null || records.isEmpty() -> {
                    mxCache[domain] = emptyList()
                    emptyList()
                }
                else -> {
                    val sorted = records
                        .filterIsInstance<MXRecord>()
                        .sortedBy { it.priority }
                        .map { it.target.toString().trimEnd('.') }
                    mxCache[domain] = sorted
                    sorted
                }
            }
        } catch (e: Exception) {
            mxCache[domain] = emptyList()
            emptyList()
        }
    }

    // =========================================================================
    // Toxicidad (equivalente exacto a assess_toxicity de Python)
    // =========================================================================

    private fun assessToxicity(email: String): Pair<Int, String> {
        if ("@" !in email) return Pair(0, "formato inválido, no evaluado")
        val local  = email.substringBefore("@").lowercase()
        val domain = email.substringAfter("@").lowercase()
        val reasons = mutableListOf<String>()
        var score = 0

        if (domain in DISPOSABLE_DOMAINS) { score += 3; reasons += "dominio desechable/temporal" }
        if (domain in TYPO_DOMAINS)        { score += 2; reasons += "typo de dominio conocido" }
        if (local  in ROLE_PREFIXES)       { score += 1; reasons += "dirección de rol genérica" }
        if (looksRandom(local))            { score += 1; reasons += "local-part con patrón aleatorio (posible spam trap)" }

        score = minOf(score, 5)
        return Pair(score, if (reasons.isEmpty()) "sin señales detectadas" else reasons.joinToString("; "))
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

    // =========================================================================
    // Utilidades
    // =========================================================================

    private fun randomLocalPart(len: Int = 14) =
        (1..len).map { ('a'..'z').toList()[java.util.Random().nextInt(26)] }.joinToString("")

    /** Crea el JSObject de resultado igual que el dict de Python */
    private fun result(
        email: String, status: String, detalle: String,
        toxicidad: Int, señales: String
    ) = JSObject()
        .put("email",    email)
        .put("status",   status)
        .put("detalle",  detalle)
        .put("toxicidad", toxicidad)
        .put("señales_toxicidad", señales)

    /**
     * Empuja un evento al JavaScript del WebView igual que pywebview hace
     * window.evaluate_js("window.onXxx(payload)").
     * Capacitor expone bridge.eval() para esto.
     */
    private fun notifyJs(fnName: String, payload: Any) {
        try {
            val json = payload.toString()
            bridge.eval("window.$fnName($json)", null)
        } catch (_: Exception) {}
    }
}
