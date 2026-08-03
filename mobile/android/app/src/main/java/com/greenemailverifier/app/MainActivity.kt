package com.greenemailverifier.app

import android.os.Bundle
import com.getcapacitor.BridgeActivity

class MainActivity : BridgeActivity() {
    override fun onCreate(savedInstanceState: Bundle?) {
        registerPlugin(Port25Plugin::class.java)
        registerPlugin(SmtpVerifierPlugin::class.java)
        super.onCreate(savedInstanceState)
    }
}
