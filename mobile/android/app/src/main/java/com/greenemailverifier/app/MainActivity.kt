package com.greenemailverifier.app

import android.os.Bundle
import com.getcapacitor.BridgeActivity

class MainActivity : BridgeActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        // Registra los dos plugins antes de super.onCreate
        registerPlugin(Port25Plugin::class.java)        // prueba legada (puedes quitar si ya no la necesitas)
        registerPlugin(SmtpVerifierPlugin::class.java)  // verificador completo nuevo
        super.onCreate(savedInstanceState)
    }
}
