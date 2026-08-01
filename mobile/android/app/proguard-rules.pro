# ── Green Email Verifier — ProGuard rules ──────────────────────────────────

# dnsjava: mantener todas las clases públicas para que la resolución MX funcione
-keep class org.xbill.DNS.** { *; }
-dontwarn org.xbill.DNS.**

# Capacitor: mantener los plugins y sus anotaciones
-keep class com.getcapacitor.** { *; }
-keepclassmembers class * extends com.getcapacitor.Plugin {
    @com.getcapacitor.annotation.CapacitorPlugin *;
    @com.getcapacitor.PluginMethod *;
}

# SmtpVerifierPlugin y Port25Plugin — no ofuscar nombres (Capacitor los busca por nombre)
-keep class com.greenemailverifier.app.SmtpVerifierPlugin { *; }
-keep class com.greenemailverifier.app.Port25Plugin { *; }
-keep class com.greenemailverifier.app.MainActivity { *; }

# Kotlin stdlib
-dontwarn kotlin.**
-keep class kotlin.** { *; }
-keep class kotlin.Metadata { *; }

# JSObject de Capacitor (serialización JSON nativa)
-keepclassmembers class * {
    @com.getcapacitor.annotation.* <methods>;
}
