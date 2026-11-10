// Most do aplikacji Android (+ zaślepka do testów w przeglądarce)
const A = window.Android
export const bridge = {
  fileInfo: () => (A ? A.fileInfo() : JSON.stringify({ hasFile: false, name: 'test', ext: '' })),
  save: (b64, ext, asCopy) => A && A.save(b64, ext, asCopy),
  error: (m) => (A ? A.error(m) : alert(m)),
  complex: (m) => A && A.complex(m),
  dirty: (d) => A && A.dirty(d),
  autoSave: () => (A ? A.autoSave() : false),
}

export async function toB64(data) {
  const blob = data instanceof Blob ? data : new Blob([data])
  return await new Promise((ok, fail) => {
    const r = new FileReader()
    r.onload = () => ok(String(r.result).split(',')[1] || '')
    r.onerror = () => fail(r.error)
    r.readAsDataURL(blob)
  })
}

let box = null
export function status(t) {
  if (!box) {
    box = document.createElement('div')
    box.style.cssText = 'position:fixed;left:50%;top:40%;transform:translate(-50%,-50%);background:rgba(32,33,36,.92);color:#fff;padding:14px 22px;border-radius:12px;font:16px sans-serif;z-index:99999;display:none'
    document.body.appendChild(box)
  }
  box.textContent = t
  box.style.display = t ? 'block' : 'none'
}

/** Widok: „pc” – szeroka strona jak na komputerze, „big” – większe elementy. */
export function setView(mode) {
  const m = document.querySelector('meta[name=viewport]')
  if (!m) return
  m.setAttribute('content', mode === 'big'
    ? 'width=device-width, initial-scale=1, minimum-scale=0.3, maximum-scale=5, user-scalable=yes'
    : 'width=1280, minimum-scale=0.2, maximum-scale=5, user-scalable=yes')
  window.dispatchEvent(new Event('resize'))
}

/** Prosty parser CSV/TSV (cudzysłowy, separator wykrywany sam). */
export function csvParse(text) {
  text = text.replace(/^\uFEFF/, '')
  const first = text.split(/\r?\n/, 1)[0] || ''
  const cnt = (ch) => first.split(ch).length - 1
  const sep = [['\t', cnt('\t')], [';', cnt(';')], [',', cnt(',')]].sort((a, b) => b[1] - a[1])[0][0]
  const rows = []
  let row = [], cur = '', q = false
  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++ } else q = false }
      else cur += ch
    } else if (ch === '"') q = true
    else if (ch === sep) { row.push(cur); cur = '' }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++
      row.push(cur); rows.push(row); row = []; cur = ''
    } else cur += ch
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row) }
  return rows
}
