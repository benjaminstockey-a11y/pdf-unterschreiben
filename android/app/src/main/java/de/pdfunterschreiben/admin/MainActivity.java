package de.pdfunterschreiben.admin;

import android.annotation.SuppressLint;
import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Bundle;
import android.webkit.JavascriptInterface;
import android.webkit.WebChromeClient;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import org.json.JSONObject;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;

/**
 * Admin-App für PDF Unterschreiben: zeigt die Oberfläche aus assets/index.html und stellt ihr
 * eine kleine Brücke ins Netz bereit (Anfragen an den Anmeldeserver, Links öffnen).
 */
public class MainActivity extends Activity {
    private WebView web;

    @SuppressLint({"SetJavaScriptEnabled", "AddJavascriptInterface"})
    @Override
    protected void onCreate(Bundle state) {
        super.onCreate(state);
        web = new WebView(this);
        WebSettings s = web.getSettings();
        s.setJavaScriptEnabled(true);
        s.setDomStorageEnabled(true);
        s.setAllowFileAccess(false);
        s.setAllowContentAccess(false);
        web.setWebChromeClient(new WebChromeClient());
        web.setWebViewClient(new WebViewClient() {
            @Override
            public boolean shouldOverrideUrlLoading(WebView view, String url) {
                if (url.startsWith("file:///android_asset/")) return false;
                openExternal(url);
                return true;
            }
        });
        web.addJavascriptInterface(new Bridge(), "Native");
        setContentView(web);
        if (state != null) web.restoreState(state);
        else web.loadUrl("file:///android_asset/index.html");
    }

    @Override
    protected void onSaveInstanceState(Bundle out) {
        super.onSaveInstanceState(out);
        web.saveState(out);
    }

    @Override
    public void onBackPressed() {
        web.evaluateJavascript("window.goBack && window.goBack()", v -> {
            if (!"true".equals(v)) MainActivity.super.onBackPressed();
        });
    }

    private void openExternal(String url) {
        if (!url.startsWith("https://")) return;
        try { startActivity(new Intent(Intent.ACTION_VIEW, Uri.parse(url))); } catch (Exception ignored) { }
    }

    private void reply(String cb, String result) {
        String js = "window.__reply(" + JSONObject.quote(cb) + "," + JSONObject.quote(result) + ")";
        runOnUiThread(() -> web.evaluateJavascript(js, null));
    }

    private static String read(InputStream in) throws Exception {
        ByteArrayOutputStream b = new ByteArrayOutputStream();
        byte[] buf = new byte[8192];
        int n;
        while ((n = in.read(buf)) > 0) b.write(buf, 0, n);
        return b.toString("UTF-8");
    }

    private class Bridge {
        @JavascriptInterface
        public String version() { return BuildConfig.VERSION_NAME; }

        @JavascriptInterface
        public String repo() { return BuildConfig.REPO; }

        @JavascriptInterface
        public boolean configured() { return BuildConfig.REPORT_URL.startsWith("https://script.google.com/"); }

        /** Anfrage an den Anmeldeserver; der Schlüssel wird hier ergänzt und bleibt in der App. */
        @JavascriptInterface
        public void post(String json, String cb) {
            new Thread(() -> {
                try {
                    JSONObject o = new JSONObject(json);
                    o.put("key", BuildConfig.REPORT_KEY);
                    o.put("version", "Android " + BuildConfig.VERSION_NAME);
                    HttpURLConnection c = (HttpURLConnection) new URL(BuildConfig.REPORT_URL).openConnection();
                    c.setConnectTimeout(20000);
                    c.setReadTimeout(40000);
                    c.setInstanceFollowRedirects(true);
                    c.setRequestMethod("POST");
                    c.setDoOutput(true);
                    c.setRequestProperty("Content-Type", "application/json");
                    try (OutputStream os = c.getOutputStream()) { os.write(o.toString().getBytes(StandardCharsets.UTF_8)); }
                    int code = c.getResponseCode();
                    // Apps Script antwortet mit einer Weiterleitung, die als GET geholt werden muss
                    if (code == 301 || code == 302 || code == 303 || code == 307) {
                        String loc = c.getHeaderField("Location");
                        c.disconnect();
                        c = (HttpURLConnection) new URL(loc).openConnection();
                        c.setConnectTimeout(20000);
                        c.setReadTimeout(40000);
                        code = c.getResponseCode();
                    }
                    String body = read(code >= 400 ? c.getErrorStream() : c.getInputStream());
                    c.disconnect();
                    reply(cb, body);
                } catch (Exception e) {
                    reply(cb, "{\"ok\":false,\"offline\":true,\"error\":\"Keine Verbindung zum Server.\"}");
                }
            }).start();
        }

        /** Einfache GET-Anfrage (für die Update-Prüfung). */
        @JavascriptInterface
        public void get(String url, String cb) {
            new Thread(() -> {
                try {
                    if (!url.startsWith("https://github.com/")) throw new Exception("nicht erlaubt");
                    HttpURLConnection c = (HttpURLConnection) new URL(url).openConnection();
                    c.setConnectTimeout(15000);
                    c.setReadTimeout(20000);
                    c.setInstanceFollowRedirects(true);
                    int code = c.getResponseCode();
                    String body = code == 200 ? read(c.getInputStream()) : "";
                    c.disconnect();
                    reply(cb, body);
                } catch (Exception e) {
                    reply(cb, "");
                }
            }).start();
        }

        @JavascriptInterface
        public void open(String url) { runOnUiThread(() -> openExternal(url)); }
    }
}
