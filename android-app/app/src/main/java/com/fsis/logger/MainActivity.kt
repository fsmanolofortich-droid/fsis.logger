package com.fsis.logger

import android.annotation.SuppressLint
import android.app.Activity
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Bundle
import android.webkit.GeolocationPermissions
import android.webkit.ValueCallback
import android.webkit.WebChromeClient
import android.webkit.WebResourceError
import android.webkit.WebResourceRequest
import android.webkit.WebSettings
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.Toast
import androidx.activity.OnBackPressedCallback
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.content.ContextCompat
import androidx.core.view.WindowCompat

/**
 * WebView shell for the FSIS web app. Host your static files (home.html, bfp.css, home.js, …)
 * over HTTPS, set [R.string.app_web_url] to that page, then build a release APK/AAB in Android Studio.
 */
class MainActivity : AppCompatActivity() {

    private lateinit var webView: WebView
    private var filePathCallback: ValueCallback<Array<Uri>>? = null
    private var geoPending: Pair<String, GeolocationPermissions.Callback>? = null

    private val fileChooserLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        val cb = filePathCallback ?: return@registerForActivityResult
        filePathCallback = null
        when {
            result.resultCode != Activity.RESULT_OK -> cb.onReceiveValue(null)
            result.data == null -> cb.onReceiveValue(null)
            else -> {
                val data = result.data!!
                val single = data.data
                if (single != null) {
                    cb.onReceiveValue(arrayOf(single))
                    return@registerForActivityResult
                }
                val clip = data.clipData
                if (clip != null && clip.itemCount > 0) {
                    val uris = Array(clip.itemCount) { clip.getItemAt(it).uri }
                    cb.onReceiveValue(uris)
                } else {
                    cb.onReceiveValue(null)
                }
            }
        }
    }

    private val locationPermissionLauncher = registerForActivityResult(
        ActivityResultContracts.RequestPermission()
    ) { granted ->
        val pending = geoPending ?: return@registerForActivityResult
        geoPending = null
        val (origin, callback) = pending
        callback.invoke(origin, granted, false)
    }

    @SuppressLint("SetJavaScriptEnabled")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        WindowCompat.setDecorFitsSystemWindows(window, true)
        webView = WebView(this)
        setContentView(webView)

        WebView.setWebContentsDebuggingEnabled(BuildConfig.DEBUG)

        webView.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            @Suppress("DEPRECATION")
            databaseEnabled = true
            loadWithOverviewMode = true
            useWideViewPort = true
            builtInZoomControls = true
            displayZoomControls = false
            mixedContentMode = WebSettings.MIXED_CONTENT_NEVER_ALLOW
            setGeolocationEnabled(true)
            allowFileAccess = true
            allowContentAccess = true
        }

        webView.webViewClient = object : WebViewClient() {
            override fun shouldOverrideUrlLoading(view: WebView, request: WebResourceRequest): Boolean {
                val url = request.url.toString()
                if (url.startsWith("tel:") || url.startsWith("mailto:")) {
                    startActivity(Intent(Intent.ACTION_VIEW, Uri.parse(url)))
                    return true
                }
                if (url.contains("google.com/maps") || url.contains("maps.google.com")) {
                    startActivity(Intent(Intent.ACTION_VIEW, Uri.parse(url)))
                    return true
                }
                return false
            }

            override fun onReceivedError(
                view: WebView,
                request: WebResourceRequest,
                error: WebResourceError
            ) {
                super.onReceivedError(view, request, error)
                if (request.isForMainFrame) {
                    Toast.makeText(this@MainActivity, getString(R.string.load_error), Toast.LENGTH_LONG).show()
                }
            }
        }

        webView.webChromeClient = object : WebChromeClient() {
            override fun onShowFileChooser(
                webView: WebView?,
                filePathCallback: ValueCallback<Array<Uri>>?,
                fileChooserParams: FileChooserParams?
            ): Boolean {
                this@MainActivity.filePathCallback?.onReceiveValue(null)
                this@MainActivity.filePathCallback = filePathCallback
                val intent = Intent(Intent.ACTION_GET_CONTENT).apply {
                    addCategory(Intent.CATEGORY_OPENABLE)
                    type = "image/*"
                    if (fileChooserParams?.mode == FileChooserParams.MODE_OPEN_MULTIPLE) {
                        putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true)
                    }
                }
                return try {
                    fileChooserLauncher.launch(
                        Intent.createChooser(intent, getString(R.string.choose_image))
                    )
                    true
                } catch (_: Exception) {
                    filePathCallback?.onReceiveValue(null)
                    this@MainActivity.filePathCallback = null
                    false
                }
            }

            override fun onGeolocationPermissionsShowPrompt(
                origin: String?,
                callback: GeolocationPermissions.Callback?
            ) {
                if (origin == null || callback == null) return
                val fine = android.Manifest.permission.ACCESS_FINE_LOCATION
                when {
                    ContextCompat.checkSelfPermission(this@MainActivity, fine) ==
                        PackageManager.PERMISSION_GRANTED -> callback.invoke(origin, true, false)
                    else -> {
                        geoPending = origin to callback
                        locationPermissionLauncher.launch(fine)
                    }
                }
            }
        }

        onBackPressedDispatcher.addCallback(
            this,
            object : OnBackPressedCallback(true) {
                override fun handleOnBackPressed() {
                    if (webView.canGoBack()) {
                        webView.goBack()
                    } else {
                        finish()
                    }
                }
            }
        )

        val url = getString(R.string.app_web_url).trim()
        val placeholder = "REPLACE_WITH_YOUR_HTTPS_URL/home.html"
        if (url.isEmpty() || url == placeholder || url.contains("REPLACE_WITH_YOUR_HTTPS_URL")) {
            Toast.makeText(this, getString(R.string.set_app_url), Toast.LENGTH_LONG).show()
        }
        webView.loadUrl(url.ifEmpty { "about:blank" })
    }
}
