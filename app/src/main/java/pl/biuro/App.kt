package pl.biuro

import android.app.Activity
import android.app.AlertDialog
import android.content.Context
import android.content.Intent
import android.graphics.Color
import android.graphics.Rect
import android.graphics.Typeface
import android.graphics.drawable.GradientDrawable
import android.net.Uri
import android.os.Bundle
import android.os.Handler
import android.os.Looper
import android.print.PrintAttributes
import android.print.PrintManager
import android.provider.OpenableColumns
import android.util.Base64
import android.view.Gravity
import android.view.View
import android.webkit.JavascriptInterface
import android.webkit.WebChromeClient
import android.webkit.WebResourceRequest
import android.webkit.WebResourceResponse
import android.webkit.WebView
import android.webkit.WebViewClient
import android.widget.LinearLayout
import android.widget.PopupMenu
import android.widget.ScrollView
import android.widget.TextView
import android.widget.Toast
import androidx.core.content.FileProvider
import androidx.webkit.WebViewAssetLoader
import com.google.mlkit.vision.common.InputImage
import com.google.mlkit.vision.text.TextRecognition
import com.google.mlkit.vision.text.latin.TextRecognizerOptions
import org.json.JSONArray
import org.json.JSONObject
import java.io.ByteArrayInputStream
import java.io.File
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

/** Rodzaj pliku po rozszerzeniu. */
object Kind {
    val SHEET = setOf("xlsx", "xlsm", "xltx", "xltm", "xls", "xlsb", "ods", "csv", "tsv")
    val DOC = setOf("docx", "docm", "dotx", "txt", "md")
    val NOT_YET = setOf("doc", "odt", "rtf", "pages")

    fun ext(name: String) = name.substringAfterLast('.', "").lowercase()

    fun of(name: String): String = when (ext(name)) {
        in SHEET -> "sheet"
        in DOC -> "doc"
        else -> ""
    }

    val MIMES = arrayOf(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "application/vnd.ms-excel",
        "application/vnd.oasis.opendocument.spreadsheet", "text/csv", "text/comma-separated-values", "text/tab-separated-values",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword",
        "application/vnd.oasis.opendocument.text", "text/plain", "application/rtf", "application/octet-stream"
    )
}

/** Ostatnie pliki. */
object Recent {
    fun list(c: Context): List<JSONObject> {
        val a = try { JSONArray(c.getSharedPreferences("b", 0).getString("recent", "[]")) } catch (e: Exception) { JSONArray() }
        return (0 until a.length()).map { a.getJSONObject(it) }
    }

    fun add(c: Context, uri: Uri, name: String) {
        val l = list(c).filter { it.optString("u") != uri.toString() }.toMutableList()
        l.add(0, JSONObject().put("u", uri.toString()).put("n", name).put("t", System.currentTimeMillis()))
        c.getSharedPreferences("b", 0).edit().putString("recent", JSONArray(l.take(40)).toString()).apply()
    }

    fun remove(c: Context, uri: String) {
        val l = list(c).filter { it.optString("u") != uri }
        c.getSharedPreferences("b", 0).edit().putString("recent", JSONArray(l).toString()).apply()
    }

    var autoSave: Boolean = true
}

fun Context.dp(v: Int) = (v * resources.displayMetrics.density).toInt()

fun nameOf(c: Context, u: Uri): String {
    try {
        c.contentResolver.query(u, arrayOf(OpenableColumns.DISPLAY_NAME), null, null, null)?.use { cur ->
            if (cur.moveToFirst()) return cur.getString(0) ?: "plik"
        }
    } catch (e: Exception) {
    }
    return u.lastPathSegment?.substringAfterLast('/') ?: "plik"
}

// ====================================================================== ekran startowy
class MainActivity : Activity() {
    private lateinit var col: LinearLayout
    private val REQ_OPEN = 1
    private val REQ_PHOTO = 2
    private val REQ_PICK_IMG = 3
    private var photoUri: Uri? = null

    private fun tv(t: String, size: Float = 15f, bold: Boolean = false, color: Int = 0xFF202124.toInt()) = TextView(this).apply {
        text = t; textSize = size; setTextColor(color); if (bold) typeface = Typeface.DEFAULT_BOLD
    }

    private fun card(title: String, sub: String, f: () -> Unit) = LinearLayout(this).apply {
        orientation = LinearLayout.VERTICAL
        background = GradientDrawable().apply { setColor(0xFFF1F3F4.toInt()); cornerRadius = dp(16).toFloat() }
        setPadding(dp(16), dp(14), dp(16), dp(14))
        addView(tv(title, 17f, true))
        if (sub.isNotEmpty()) addView(tv(sub, 13f, color = 0xFF5F6368.toInt()))
        setOnClickListener { f() }
    }

    private fun lp(top: Int = 10) = LinearLayout.LayoutParams(-1, -2).apply { topMargin = dp(top) }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        Recent.autoSave = getSharedPreferences("b", 0).getBoolean("auto", true)
        val sv = ScrollView(this)
        col = LinearLayout(this).apply { orientation = LinearLayout.VERTICAL; setPadding(dp(20), dp(20), dp(20), dp(40)) }
        sv.addView(col)
        setContentView(sv)
        // plik otwarty z innej aplikacji
        if (intent?.action == Intent.ACTION_VIEW && intent.data != null) openUri(intent.data!!, false)
    }

    override fun onResume() {
        super.onResume()
        render()
    }

    private fun render() {
        col.removeAllViews()
        col.addView(tv("Biuro", 32f, true))
        col.addView(tv("Arkusze i dokumenty offline · zapis w xlsx i docx", 14f, color = 0xFF5F6368.toInt()))
        col.addView(card("📂  Otwórz plik", "xlsx, xls, ods, csv · docx, txt") {
            startActivityForResult(Intent(Intent.ACTION_OPEN_DOCUMENT).addCategory(Intent.CATEGORY_OPENABLE).setType("*/*")
                .putExtra(Intent.EXTRA_MIME_TYPES, Kind.MIMES), REQ_OPEN)
        }, lp(18))
        val r = LinearLayout(this).apply { orientation = LinearLayout.HORIZONTAL }
        r.addView(card("＋ Arkusz", "") { edit("sheet", null) }, LinearLayout.LayoutParams(0, -2, 1f).apply { rightMargin = dp(8) })
        r.addView(card("＋ Dokument", "") { edit("doc", null) }, LinearLayout.LayoutParams(0, -2, 1f))
        col.addView(r, lp())
        col.addView(card("📷  Zdjęcie tabeli → arkusz", "Zrób zdjęcie wydruku lub kartki – dane trafią do komórek") { photo() }, lp())
        col.addView(tv("Szablony", 16f, true), lp(22))
        for ((k, n, s) in listOf(Triple("grafik", "Grafik zmian", "pracownicy × dni, suma godzin"),
            Triple("raport", "Raport dzienny zakładu", "obsada, wydajność, plan, realizacja %"),
            Triple("rezerwacje", "Rezerwacje pensjonatu", "przyjazd, wyjazd, noce, kwoty, zaliczka"))) {
            col.addView(card(n, s) { startActivity(Intent(this, EditorActivity::class.java).putExtra("kind", "sheet").putExtra("template", k)) }, lp(8))
        }
        val rec = Recent.list(this)
        if (rec.isNotEmpty()) col.addView(tv("Ostatnie", 16f, true), lp(22))
        val df = SimpleDateFormat("d.MM HH:mm", Locale("pl"))
        for (o in rec) {
            val n = o.optString("n")
            val icon = if (Kind.of(n) == "sheet") "▦" else "≡"
            val c = card("$icon  $n", df.format(Date(o.optLong("t")))) { openUri(Uri.parse(o.optString("u")), true) }
            c.setOnLongClickListener {
                AlertDialog.Builder(this).setTitle(n).setItems(arrayOf("Usuń z listy")) { _, _ -> Recent.remove(this, o.optString("u")); render() }.show()
                true
            }
            col.addView(c, lp(8))
        }
        val auto = tv((if (Recent.autoSave) "☑" else "☐") + "  Zapisuj automatycznie co 30 s", 15f).apply {
            setPadding(0, dp(18), 0, dp(8))
            setOnClickListener {
                Recent.autoSave = !Recent.autoSave
                getSharedPreferences("b", 0).edit().putBoolean("auto", Recent.autoSave).apply(); render()
            }
        }
        col.addView(auto)
    }

    private fun edit(kind: String, uri: Uri?) {
        startActivity(Intent(this, EditorActivity::class.java).putExtra("kind", kind).also { if (uri != null) it.data = uri })
    }

    private fun openUri(u: Uri, fromRecent: Boolean) {
        val name = nameOf(this, u)
        val kind = Kind.of(name)
        if (kind.isEmpty()) {
            val ext = Kind.ext(name)
            Toast.makeText(this, if (ext in Kind.NOT_YET) "Format .$ext nie jest jeszcze obsługiwany – zapisz go jako .docx" else "Nieobsługiwany plik: $name", Toast.LENGTH_LONG).show()
            return
        }
        try {
            contentResolver.takePersistableUriPermission(u, Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION)
        } catch (e: Exception) {
            if (fromRecent) try { contentResolver.openInputStream(u)?.close() } catch (x: Exception) {
                Toast.makeText(this, "Brak dostępu do pliku – otwórz go ponownie", Toast.LENGTH_LONG).show(); Recent.remove(this, u.toString()); render(); return
            }
        }
        Recent.add(this, u, name)
        edit(kind, u)
    }

    private fun photo() {
        AlertDialog.Builder(this).setTitle("Zdjęcie tabeli").setItems(arrayOf("Zrób zdjęcie", "Wybierz z galerii")) { _, w ->
            if (w == 0) {
                val f = File(cacheDir, "tabela.jpg")
                photoUri = FileProvider.getUriForFile(this, "$packageName.files", f)
                startActivityForResult(Intent(android.provider.MediaStore.ACTION_IMAGE_CAPTURE).putExtra(android.provider.MediaStore.EXTRA_OUTPUT, photoUri)
                    .addFlags(Intent.FLAG_GRANT_WRITE_URI_PERMISSION), REQ_PHOTO)
            } else startActivityForResult(Intent(Intent.ACTION_GET_CONTENT).setType("image/*"), REQ_PICK_IMG)
        }.show()
    }

    @Deprecated("Deprecated in Java")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        @Suppress("DEPRECATION")
        super.onActivityResult(requestCode, resultCode, data)
        if (resultCode != RESULT_OK) return
        when (requestCode) {
            REQ_OPEN -> data?.data?.let { openUri(it, false) }
            REQ_PHOTO -> photoUri?.let { scanTable(it) }
            REQ_PICK_IMG -> data?.data?.let { scanTable(it) }
        }
    }

    /** Rozpoznaje tekst i układa go w wiersze i kolumny (TSV). */
    private fun scanTable(u: Uri) {
        Toast.makeText(this, "Rozpoznaję tabelę…", Toast.LENGTH_SHORT).show()
        val img = try { InputImage.fromFilePath(this, u) } catch (e: Exception) { Toast.makeText(this, "Nie mogę wczytać zdjęcia", Toast.LENGTH_LONG).show(); return }
        TextRecognition.getClient(TextRecognizerOptions.DEFAULT_OPTIONS).process(img)
            .addOnSuccessListener { res ->
                class W(val t: String, val r: Rect)
                val words = ArrayList<W>()
                for (b in res.textBlocks) for (l in b.lines) for (e in l.elements) e.boundingBox?.let { words.add(W(e.text, it)) }
                if (words.isEmpty()) { Toast.makeText(this, "Nie znaleziono tekstu", Toast.LENGTH_LONG).show(); return@addOnSuccessListener }
                val hMed = words.map { it.r.height() }.sorted()[words.size / 2].coerceAtLeast(8)
                // wiersze po środku w pionie
                val rows = ArrayList<MutableList<W>>()
                for (w in words.sortedBy { it.r.centerY() }) {
                    val row = rows.lastOrNull()
                    if (row != null && Math.abs(row.map { it.r.centerY() }.average() - w.r.centerY()) < hMed * 0.6) row.add(w) else rows.add(mutableListOf(w))
                }
                val sb = StringBuilder()
                for (row in rows) {
                    val cells = ArrayList<String>()
                    var cur = StringBuilder()
                    var lastRight = -1
                    for (w in row.sortedBy { it.r.left }) {
                        if (lastRight >= 0 && w.r.left - lastRight > hMed * 1.2) { cells.add(cur.toString()); cur = StringBuilder() }
                        if (cur.isNotEmpty()) cur.append(' ')
                        cur.append(w.t)
                        lastRight = w.r.right
                    }
                    cells.add(cur.toString())
                    sb.append(cells.joinToString("\t") { it.replace("\t", " ") }).append('\n')
                }
                startActivity(Intent(this, EditorActivity::class.java).putExtra("kind", "sheet").putExtra("tsv", sb.toString()))
            }
            .addOnFailureListener { Toast.makeText(this, "Rozpoznawanie nieudane: ${it.message}", Toast.LENGTH_LONG).show() }
    }
}

// ====================================================================== edytor
class EditorActivity : Activity() {
    private val h = Handler(Looper.getMainLooper())
    private lateinit var web: WebView
    private lateinit var titleV: TextView
    private lateinit var saveBtn: TextView
    private var uri: Uri? = null
    private var name = ""
    private var kind = "sheet"
    private var tsv: String? = null
    private var template: String? = null
    private var dirty = false
    private var complex = ""
    private var pending: ByteArray? = null
    private var pendingExt = "xlsx"
    private var view = "pc"
    private var closeAfterSave = false
    private val REQ_CREATE = 10

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        uri = intent.data
        kind = intent.getStringExtra("kind") ?: "sheet"
        tsv = intent.getStringExtra("tsv")
        template = intent.getStringExtra("template")
        name = uri?.let { nameOf(this, it) } ?: when {
            template != null -> mapOf("grafik" to "Grafik zmian.xlsx", "raport" to "Raport dzienny.xlsx", "rezerwacje" to "Rezerwacje.xlsx")[template] ?: "Arkusz.xlsx"
            tsv != null -> "Tabela ze zdjęcia.xlsx"
            kind == "doc" -> "Dokument.docx"
            else -> "Arkusz.xlsx"
        }
        val root = LinearLayout(this).apply { orientation = LinearLayout.VERTICAL }
        root.addView(bar())
        web = WebView(this)
        root.addView(web, LinearLayout.LayoutParams(-1, 0, 1f))
        setContentView(root)
        setupWeb()
    }

    private fun bar(): View {
        val b = LinearLayout(this).apply {
            orientation = LinearLayout.HORIZONTAL; gravity = Gravity.CENTER_VERTICAL
            setPadding(dp(6), dp(4), dp(6), dp(4)); setBackgroundColor(0xFFF8F9FA.toInt())
        }
        fun btn(t: String, f: (View) -> Unit) = TextView(this).apply {
            text = t; textSize = 16f; setTextColor(0xFF202124.toInt()); gravity = Gravity.CENTER
            setPadding(dp(12), dp(8), dp(12), dp(8)); setOnClickListener { f(it) }
        }
        b.addView(btn("←") { onBackPressed() })
        titleV = TextView(this).apply { text = name; textSize = 15f; setTextColor(0xFF202124.toInt()); maxLines = 1; ellipsize = android.text.TextUtils.TruncateAt.MIDDLE }
        b.addView(titleV, LinearLayout.LayoutParams(0, -2, 1f))
        saveBtn = btn("Zapisz") { save(false) }.apply { setTypeface(typeface, Typeface.BOLD) }
        b.addView(saveBtn)
        b.addView(btn("⋮") { menu(it) })
        return b
    }

    private fun menu(anchor: View) {
        val p = PopupMenu(this, anchor)
        p.menu.add(0, 1, 0, "Zapisz jako kopię…")
        p.menu.add(0, 2, 0, if (view == "pc") "Widok: większe elementy" else "Widok: jak na komputerze")
        p.menu.add(0, 3, 0, "Udostępnij")
        p.menu.add(0, 4, 0, "Drukuj / PDF")
        p.setOnMenuItemClickListener {
            when (it.itemId) {
                1 -> save(true)
                2 -> { view = if (view == "pc") "big" else "pc"; web.evaluateJavascript("window.biuroView && window.biuroView('$view')", null) }
                3 -> share()
                4 -> print()
            }
            true
        }
        p.show()
    }

    private fun setupWeb() {
        val loader = WebViewAssetLoader.Builder()
            .addPathHandler("/assets/", WebViewAssetLoader.AssetsPathHandler(this))
            .addPathHandler("/file/", object : WebViewAssetLoader.PathHandler {
                override fun handle(path: String): WebResourceResponse? {
                    val data: java.io.InputStream? = when {
                        tsv != null -> ByteArrayInputStream(tsv!!.toByteArray())
                        uri != null -> try { contentResolver.openInputStream(uri!!) } catch (e: Exception) { null }
                        else -> null
                    }
                    return WebResourceResponse("application/octet-stream", null, data ?: ByteArrayInputStream(ByteArray(0)))
                }
            })
            .build()
        web.settings.apply {
            javaScriptEnabled = true
            domStorageEnabled = true
            useWideViewPort = true
            loadWithOverviewMode = true
            builtInZoomControls = true
            displayZoomControls = false
            setSupportZoom(true)
            allowFileAccess = false
            allowContentAccess = false
            textZoom = 100
        }
        web.webViewClient = object : WebViewClient() {
            override fun shouldInterceptRequest(view: WebView, request: WebResourceRequest): WebResourceResponse? =
                loader.shouldInterceptRequest(request.url)

            override fun shouldOverrideUrlLoading(view: WebView, request: WebResourceRequest): Boolean {
                val u = request.url
                if (u.host == "appassets.androidplatform.net") return false
                try { startActivity(Intent(Intent.ACTION_VIEW, u)) } catch (e: Exception) {}
                return true
            }
        }
        web.webChromeClient = WebChromeClient()
        web.addJavascriptInterface(Bridge(), "Android")
        web.loadUrl("https://appassets.androidplatform.net/assets/www/" + (if (kind == "doc") "doc.html" else "sheet.html"))
    }

    inner class Bridge {
        @JavascriptInterface
        fun fileInfo(): String {
            val ext = if (tsv != null) "tsv" else Kind.ext(name)
            return JSONObject().put("hasFile", uri != null || tsv != null).put("name", name).put("ext", ext)
                .put("template", template ?: "").toString()
        }

        @JavascriptInterface
        fun save(b64: String, ext: String, asCopy: Boolean) {
            val bytes = Base64.decode(b64, Base64.DEFAULT)
            h.post { write(bytes, ext, asCopy) }
        }

        @JavascriptInterface
        fun error(m: String) { h.post { Toast.makeText(this@EditorActivity, m, Toast.LENGTH_LONG).show() } }

        @JavascriptInterface
        fun complex(m: String) {
            h.post {
                complex = m
                AlertDialog.Builder(this@EditorActivity).setTitle("Plik zawiera: $m")
                    .setMessage("Edytor pokaże i zachowa dane, formuły i formatowanie, ale $m mogą nie przetrwać zapisu. Dlatego zmiany zapiszą się jako KOPIA – oryginał zostanie nietknięty.")
                    .setPositiveButton("Rozumiem", null).show()
            }
        }

        @JavascriptInterface
        fun dirty(d: Boolean) { h.post { dirty = d; saveBtn.text = if (d) "Zapisz •" else "Zapisz" } }

        @JavascriptInterface
        fun autoSave(): Boolean = Recent.autoSave && uri != null && complex.isEmpty() && Kind.ext(name) in setOf("xlsx", "docx")
    }

    private fun save(asCopy: Boolean) {
        web.evaluateJavascript("window.biuroSave && window.biuroSave(${if (asCopy) "true" else "false"})", null)
    }

    /** Zapis: do tego samego pliku, gdy format się zgadza; inaczej „Zapisz jako” (xlsx/docx). */
    private fun write(bytes: ByteArray, ext: String, asCopy: Boolean) {
        val same = uri != null && Kind.ext(name) == ext && !asCopy
        if (same) {
            try {
                contentResolver.openOutputStream(uri!!, "wt")?.use { it.write(bytes) } ?: throw Exception("brak dostępu")
                done()
                return
            } catch (e: Exception) {
                Toast.makeText(this, "Nie mogę nadpisać pliku (${e.message}) – wybierz miejsce zapisu", Toast.LENGTH_LONG).show()
            }
        }
        pending = bytes
        pendingExt = ext
        val base = name.substringBeforeLast('.').ifBlank { "Plik" } + if (asCopy || (uri != null && Kind.ext(name) == ext)) " (kopia)" else ""
        val mime = if (ext == "xlsx") "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        startActivityForResult(Intent(Intent.ACTION_CREATE_DOCUMENT).addCategory(Intent.CATEGORY_OPENABLE).setType(mime).putExtra(Intent.EXTRA_TITLE, "$base.$ext"), REQ_CREATE)
    }

    private fun done() {
        dirty = false
        saveBtn.text = "Zapisz"
        Toast.makeText(this, "Zapisano", Toast.LENGTH_SHORT).show()
        if (closeAfterSave) finish()
    }

    @Deprecated("Deprecated in Java")
    override fun onActivityResult(requestCode: Int, resultCode: Int, data: Intent?) {
        @Suppress("DEPRECATION")
        super.onActivityResult(requestCode, resultCode, data)
        if (requestCode != REQ_CREATE) return
        val u = data?.data
        val b = pending
        pending = null
        if (resultCode != RESULT_OK || u == null || b == null) { closeAfterSave = false; return }
        try {
            contentResolver.openOutputStream(u, "wt")?.use { it.write(b) }
            try { contentResolver.takePersistableUriPermission(u, Intent.FLAG_GRANT_READ_URI_PERMISSION or Intent.FLAG_GRANT_WRITE_URI_PERMISSION) } catch (e: Exception) {}
            // dalsza praca na nowym pliku (oryginał zostaje nietknięty)
            uri = u
            name = nameOf(this, u)
            tsv = null
            complex = ""
            titleV.text = name
            Recent.add(this, u, name)
            done()
        } catch (e: Exception) {
            Toast.makeText(this, "Zapis nieudany: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun share() {
        val u = uri
        if (u == null || dirty) { Toast.makeText(this, "Najpierw zapisz plik", Toast.LENGTH_SHORT).show(); return }
        val mime = contentResolver.getType(u) ?: "*/*"
        startActivity(Intent.createChooser(Intent(Intent.ACTION_SEND).setType(mime).putExtra(Intent.EXTRA_STREAM, u)
            .addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION), "Udostępnij"))
    }

    private fun print() {
        val pm = getSystemService(PrintManager::class.java) ?: return
        pm.print(name, web.createPrintDocumentAdapter(name), PrintAttributes.Builder().setMediaSize(PrintAttributes.MediaSize.ISO_A4).build())
    }

    @Deprecated("Deprecated in Java")
    override fun onBackPressed() {
        if (!dirty) { @Suppress("DEPRECATION") super.onBackPressed(); return }
        AlertDialog.Builder(this).setTitle("Zapisać zmiany?")
            .setPositiveButton("Zapisz") { _, _ -> closeAfterSave = true; save(false) }
            .setNegativeButton("Nie zapisuj") { _, _ -> finish() }
            .setNeutralButton("Anuluj", null).show()
    }

    override fun onPause() {
        super.onPause()
        // przy wyjściu z aplikacji – szybki zapis, jeśli włączony autozapis
        if (dirty && Recent.autoSave && uri != null && complex.isEmpty()) save(false)
    }

    override fun onDestroy() {
        web.destroy()
        super.onDestroy()
    }
}

@Suppress("unused")
private val keepColor = Color.BLACK
