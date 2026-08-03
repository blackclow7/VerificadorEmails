package com.greenemailverifier.app

import android.os.Bundle
import com.getcapacitor.BridgeActivity

class MainActivity : BridgeActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        registerPlugin(Port25Plugin::class.java)
        super.onCreate(savedInstanceState)

        // AndroidBridge: JavaScript Interface directo, independiente de Capacitor.
        // Se registra DESPUÉS de super.onCreate() porque ahí es cuando el
        // bridge (y por tanto bridge.webView) ya está inicializado.
        val webView = bridge.webView
        val androidBridge = AndroidBridge(this, webView)
        webView.addJavascriptInterface(androidBridge, "AndroidBridge")
    }
}
