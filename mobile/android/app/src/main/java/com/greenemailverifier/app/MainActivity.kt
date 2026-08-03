package com.greenemailverifier.app

import android.os.Bundle
import android.webkit.WebView
import com.getcapacitor.BridgeActivity

class MainActivity : BridgeActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        registerPlugin(Port25Plugin::class.java)
        registerPlugin(SmtpVerifierPlugin::class.java)
        super.onCreate(savedInstanceState)

        // Exponer AndroidBridge como window.AndroidBridge en el WebView
        // Esto NO depende de Capacitor bridge — funciona siempre
        val webView: WebView = bridge.webView
        val androidBridge = AndroidBridge(this, webView)
        webView.addJavascriptInterface(androidBridge, "AndroidBridge")
    }
}
