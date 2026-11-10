import { createUniver, LocaleType, mergeLocales } from '@univerjs/presets'
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import '@univerjs/preset-sheets-core/lib/index.css'
import ExcelJS from 'exceljs'
import * as XLSX from 'xlsx'
import JSZip from 'jszip'
import { bridge, toB64, status, setView, csvParse } from './common.js'

let univerAPI = null
let lastHash = ''
let complex = []

// ------------------------------------------------------------ kolory, style
function hex(c) {
  if (!c) return null
  const a = c.argb || ''
  if (a.length === 8) return '#' + a.slice(2)
  if (a.length === 6) return '#' + a
  return null
}

const HA = { left: 1, center: 2, right: 3, justify: 4, centerContinuous: 2, distributed: 4, fill: 1 }
const VA = { top: 1, middle: 2, bottom: 3, distributed: 2, justify: 2 }
const BS = { thin: 1, hair: 2, dotted: 3, dashed: 4, dashDot: 5, dashDotDot: 6, double: 7, medium: 8, mediumDashed: 9, mediumDashDot: 10, mediumDashDotDot: 11, slantDashDot: 12, thick: 13 }
const BS_REV = Object.fromEntries(Object.entries(BS).map(([k, v]) => [v, k]))
const HA_REV = { 1: 'left', 2: 'center', 3: 'right', 4: 'justify' }
const VA_REV = { 1: 'top', 2: 'middle', 3: 'bottom' }

function styleFromExcel(cell) {
  const s = {}
  const f = cell.font || {}
  if (f.bold) s.bl = 1
  if (f.italic) s.it = 1
  if (f.underline) s.ul = { s: 1 }
  if (f.strike) s.st = { s: 1 }
  if (f.size) s.fs = f.size
  if (f.name) s.ff = f.name
  const fc = hex(f.color)
  if (fc) s.cl = { rgb: fc }
  const fill = cell.fill
  if (fill && fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = hex(fill.fgColor) || hex(fill.bgColor)
    if (bg) s.bg = { rgb: bg }
  }
  const al = cell.alignment || {}
  if (HA[al.horizontal]) s.ht = HA[al.horizontal]
  if (VA[al.vertical]) s.vt = VA[al.vertical]
  if (al.wrapText) s.tb = 3
  if (al.textRotation && typeof al.textRotation === 'number') s.tr = { a: al.textRotation }
  if (cell.numFmt && cell.numFmt !== 'General') s.n = { pattern: cell.numFmt }
  const b = cell.border || {}
  const bd = {}
  for (const [k, u] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']]) {
    const x = b[k]
    if (x && x.style) bd[u] = { s: BS[x.style] || 1, cl: { rgb: hex(x.color) || '#000000' } }
  }
  if (Object.keys(bd).length) s.bd = bd
  return Object.keys(s).length ? s : null
}

function excelSerial(d) {
  return (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000
}

// ------------------------------------------------------------ odczyt xlsx (ExcelJS – ze stylami)
async function fromXlsx(buf) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(buf)
  const styles = {}
  const styleIds = new Map()
  const styleId = (s) => {
    const k = JSON.stringify(s)
    let id = styleIds.get(k)
    if (!id) { id = 's' + styleIds.size; styleIds.set(k, id); styles[id] = s }
    return id
  }
  const sheets = {}
  const order = []
  wb.eachSheet((ws, idx) => {
    const id = 'sheet' + idx
    order.push(id)
    const cellData = {}
    let maxR = 0, maxC = 0
    ws.eachRow({ includeEmpty: false }, (row, r) => {
      row.eachCell({ includeEmpty: true }, (cell, c) => {
        const R = r - 1, C = c - 1
        const out = {}
        const st = styleFromExcel(cell)
        const v = cell.value
        if (cell.type === ExcelJS.ValueType.Formula || (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v))) {
          const fm = cell.formula || (v && v.formula)
          if (fm) out.f = '=' + fm
          let res = v && v.result
          if (res instanceof Date) res = excelSerial(res)
          if (res && typeof res === 'object' && 'error' in res) res = res.error
          if (res !== undefined && res !== null) out.v = res
        } else if (v instanceof Date) {
          out.v = excelSerial(v); out.t = 2
          if (!st || !st.n) { const s2 = Object.assign({}, st || {}); s2.n = { pattern: 'yyyy-mm-dd' }; out.s = styleId(s2) }
        } else if (typeof v === 'number') { out.v = v; out.t = 2 }
        else if (typeof v === 'boolean') { out.v = v ? 1 : 0; out.t = 3 }
        else if (typeof v === 'string') { out.v = v; out.t = 1 }
        else if (v && typeof v === 'object') {
          if (v.richText) { out.v = v.richText.map(x => x.text).join(''); out.t = 1 }
          else if ('text' in v) { out.v = String(v.text); out.t = 1 }
          else if ('error' in v) { out.v = v.error; out.t = 1 }
        }
        if (st && !out.s) out.s = styleId(st)
        if (out.v === undefined && out.f === undefined && out.s === undefined) return
        ;(cellData[R] = cellData[R] || {})[C] = out
        maxR = Math.max(maxR, R); maxC = Math.max(maxC, C)
      })
    })
    const columnData = {}
    for (let c = 1; c <= Math.max(ws.columnCount, maxC + 1); c++) {
      const col = ws.getColumn(c)
      const e = {}
      if (col.width) e.w = Math.round(col.width * 7 + 5)
      if (col.hidden) e.hd = 1
      if (Object.keys(e).length) columnData[c - 1] = e
    }
    const rowData = {}
    ws.eachRow({ includeEmpty: true }, (row, r) => {
      const e = {}
      if (row.height) e.h = Math.round(row.height * 4 / 3)
      if (row.hidden) e.hd = 1
      if (Object.keys(e).length) rowData[r - 1] = e
    })
    const mergeData = []
    for (const m of (ws.model.merges || [])) {
      const d = XLSX.utils.decode_range(m)
      mergeData.push({ startRow: d.s.r, endRow: d.e.r, startColumn: d.s.c, endColumn: d.e.c })
    }
    const sheet = {
      id, name: ws.name, cellData, mergeData, columnData, rowData,
      rowCount: Math.max(maxR + 100, 1000), columnCount: Math.max(maxC + 10, 26),
      defaultColumnWidth: 88, defaultRowHeight: 24,
      hidden: ws.state === 'hidden' ? 1 : 0,
    }
    const view = (ws.views || [])[0]
    if (view && view.state === 'frozen' && (view.xSplit || view.ySplit)) {
      sheet.freeze = { xSplit: view.xSplit || 0, ySplit: view.ySplit || 0, startRow: view.ySplit || 0, startColumn: view.xSplit || 0 }
    }
    if (ws.properties && ws.properties.tabColor) { const tc = hex(ws.properties.tabColor); if (tc) sheet.tabColor = tc }
    sheets[id] = sheet
  })
  return { id: 'wb', name: 'Skoroszyt', appVersion: '', locale: LocaleType.EN_US, styles, sheetOrder: order, sheets }
}

// ------------------------------------------------------------ odczyt xls / ods / xlsb (SheetJS – wartości i formuły)
function fromOther(buf) {
  const wb = XLSX.read(buf, { type: 'array', cellFormula: true, cellNF: true, cellDates: false })
  return fromSheetJS(wb)
}

function fromSheetJS(wb) {
  const sheets = {}
  const order = []
  const styles = {}
  wb.SheetNames.forEach((name, i) => {
    const ws = wb.Sheets[name]
    const id = 'sheet' + (i + 1)
    order.push(id)
    const cellData = {}
    const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } }
    for (const addr of Object.keys(ws)) {
      if (addr[0] === '!') continue
      const cell = ws[addr]
      const p = XLSX.utils.decode_cell(addr)
      const out = {}
      if (cell.f) out.f = '=' + cell.f
      if (cell.t === 'n') { out.v = cell.v; out.t = 2 }
      else if (cell.t === 'b') { out.v = cell.v ? 1 : 0; out.t = 3 }
      else if (cell.t === 'd') { out.v = excelSerial(new Date(cell.v)); out.t = 2 }
      else if (cell.t === 'e') { out.v = cell.w || '#ERR'; out.t = 1 }
      else if (cell.v !== undefined) { out.v = String(cell.v); out.t = 1 }
      if (cell.z && cell.z !== 'General') { const k = 'n' + Object.keys(styles).length; styles[k] = { n: { pattern: cell.z } }; out.s = k }
      ;(cellData[p.r] = cellData[p.r] || {})[p.c] = out
    }
    const columnData = {}
    ;(ws['!cols'] || []).forEach((c, k) => { if (c && (c.wpx || c.wch)) columnData[k] = { w: Math.round(c.wpx || c.wch * 7 + 5) } })
    const rowData = {}
    ;(ws['!rows'] || []).forEach((r, k) => { if (r && (r.hpx || r.hpt)) rowData[k] = { h: Math.round(r.hpx || r.hpt * 4 / 3) } })
    const mergeData = (ws['!merges'] || []).map(m => ({ startRow: m.s.r, endRow: m.e.r, startColumn: m.s.c, endColumn: m.e.c }))
    sheets[id] = { id, name, cellData, columnData, rowData, mergeData, rowCount: Math.max(ref.e.r + 100, 1000), columnCount: Math.max(ref.e.c + 10, 26), defaultColumnWidth: 88, defaultRowHeight: 24 }
  })
  return { id: 'wb', name: 'Skoroszyt', appVersion: '', locale: LocaleType.EN_US, styles, sheetOrder: order, sheets }
}

function fromCsv(text) {
  const rows = csvParse(text)
  const cellData = {}
  let maxC = 0
  rows.forEach((r, i) => {
    r.forEach((v, c) => {
      if (v === '') return
      const num = Number(String(v).replace(/\s/g, '').replace(',', '.'))
      const isNum = /^-?[\d\s]+([.,]\d+)?$/.test(String(v).trim()) && !isNaN(num)
      ;(cellData[i] = cellData[i] || {})[c] = isNum ? { v: num, t: 2 } : (String(v).startsWith('=') ? { f: v } : { v, t: 1 })
      maxC = Math.max(maxC, c)
    })
  })
  return {
    id: 'wb', name: 'Skoroszyt', appVersion: '', locale: LocaleType.EN_US, styles: {}, sheetOrder: ['sheet1'],
    sheets: { sheet1: { id: 'sheet1', name: 'Arkusz1', cellData, rowCount: Math.max(rows.length + 100, 1000), columnCount: Math.max(maxC + 10, 26), defaultColumnWidth: 100, defaultRowHeight: 24 } },
  }
}

// ------------------------------------------------------------ szablony
function template(kind) {
  const b = { bl: 1, bg: { rgb: '#E8EAED' }, bd: { b: { s: 1, cl: { rgb: '#9AA0A6' } } } }
  const head = (cols) => Object.fromEntries(cols.map((t, i) => [i, { v: t, t: 1, s: 'h' }]))
  let name = 'Arkusz1', cols = [], extra = {}, widths = {}
  if (kind === 'grafik') {
    name = 'Grafik'; cols = ['Pracownik', 'Pon', 'Wt', 'Śr', 'Czw', 'Pt', 'Sob', 'Nd', 'Godzin razem']
    for (let r = 1; r <= 20; r++) extra[r] = { 8: { f: `=COUNTA(B${r + 1}:H${r + 1})*8` } }
    widths = { 0: { w: 180 }, 8: { w: 110 } }
  } else if (kind === 'raport') {
    name = 'Raport'; cols = ['Data', 'Zmiana', 'Obsada', 'Wydajność [szt]', 'Plan [szt]', 'Realizacja %', 'Przestoje [min]', 'Uwagi']
    for (let r = 1; r <= 31; r++) extra[r] = { 5: { f: `=IF(E${r + 1}=0,"",D${r + 1}/E${r + 1})`, s: 'p' } }
    widths = { 0: { w: 100 }, 7: { w: 260 } }
  } else if (kind === 'rezerwacje') {
    name = 'Rezerwacje'; cols = ['Gość', 'Telefon', 'Przyjazd', 'Wyjazd', 'Noce', 'Pokój', 'Cena / noc', 'Razem', 'Zaliczka', 'Do zapłaty']
    for (let r = 1; r <= 50; r++) extra[r] = { 4: { f: `=IF(D${r + 1}="","",D${r + 1}-C${r + 1})` }, 7: { f: `=IF(E${r + 1}="","",E${r + 1}*G${r + 1})` }, 9: { f: `=IF(H${r + 1}="","",H${r + 1}-I${r + 1})` } }
    widths = { 0: { w: 180 }, 1: { w: 120 } }
  }
  const cellData = { 0: head(cols), ...extra }
  return {
    id: 'wb', name, appVersion: '', locale: LocaleType.EN_US, styles: { h: b, p: { n: { pattern: '0%' } } }, sheetOrder: ['sheet1'],
    sheets: { sheet1: { id: 'sheet1', name, cellData, columnData: widths, rowCount: 1000, columnCount: 26, defaultColumnWidth: 100, defaultRowHeight: 26, freeze: { xSplit: 0, ySplit: 1, startRow: 1, startColumn: 0 } } },
  }
}

// ------------------------------------------------------------ zapis xlsx
function colLetters(n) { let s = ''; n++; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) } return s }
function colNum(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1 }

/** Przesuwa względne adresy w formule (formuły współdzielone). */
function shiftFormula(f, dr, dc) {
  let out = '', inStr = false
  const re = /(\$?)([A-Z]{1,3})(\$?)(\d+)/y
  for (let i = 0; i < f.length;) {
    const ch = f[i]
    if (ch === '"') { inStr = !inStr; out += ch; i++; continue }
    if (!inStr && /[A-Z$]/.test(ch) && (i === 0 || !/[A-Za-z0-9_.]/.test(f[i - 1]))) {
      re.lastIndex = i
      const m = re.exec(f)
      if (m && !/[A-Za-z0-9_(]/.test(f[re.lastIndex] || '')) {
        const c = m[1] ? colNum(m[2]) : colNum(m[2]) + dc
        const r = m[3] ? +m[4] : +m[4] + dr
        out += m[1] + colLetters(Math.max(0, c)) + m[3] + Math.max(1, r)
        i = re.lastIndex
        continue
      }
    }
    out += ch; i++
  }
  return out
}

function excelStyle(s, cell) {
  if (!s) return
  const font = {}
  if (s.bl) font.bold = true
  if (s.it) font.italic = true
  if (s.ul && s.ul.s) font.underline = true
  if (s.st && s.st.s) font.strike = true
  if (s.fs) font.size = s.fs
  if (s.ff) font.name = s.ff
  if (s.cl && s.cl.rgb) font.color = { argb: 'FF' + s.cl.rgb.replace('#', '').toUpperCase() }
  if (Object.keys(font).length) cell.font = font
  if (s.bg && s.bg.rgb) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + s.bg.rgb.replace('#', '').toUpperCase() } }
  const al = {}
  if (HA_REV[s.ht]) al.horizontal = HA_REV[s.ht]
  if (VA_REV[s.vt]) al.vertical = VA_REV[s.vt]
  if (s.tb === 3) al.wrapText = true
  if (s.tr && typeof s.tr.a === 'number') al.textRotation = s.tr.a
  if (Object.keys(al).length) cell.alignment = al
  if (s.n && s.n.pattern) cell.numFmt = s.n.pattern
  if (s.bd) {
    const bd = {}
    for (const [u, k] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) {
      const x = s.bd[u]
      if (x && x.s) bd[k] = { style: BS_REV[x.s] || 'thin', color: { argb: 'FF' + ((x.cl && x.cl.rgb) || '#000000').replace('#', '').toUpperCase() } }
    }
    if (Object.keys(bd).length) cell.border = bd
  }
}

async function toXlsx(snap) {
  const wb = new ExcelJS.Workbook()
  wb.creator = 'Biuro'
  const styles = snap.styles || {}
  for (const sid of snap.sheetOrder) {
    const sh = snap.sheets[sid]
    const ws = wb.addWorksheet((sh.name || 'Arkusz').slice(0, 31))
    if (sh.hidden) ws.state = 'hidden'
    if (sh.tabColor) ws.properties.tabColor = { argb: 'FF' + String(sh.tabColor).replace('#', '').toUpperCase() }
    const cd = sh.cellData || {}
    // formuły współdzielone: si → (formuła, wiersz, kolumna źródła)
    const shared = {}
    for (const r of Object.keys(cd)) for (const c of Object.keys(cd[r])) {
      const x = cd[r][c]
      if (x && x.si && x.f) shared[x.si] = { f: x.f, r: +r, c: +c }
    }
    for (const r of Object.keys(cd)) {
      for (const c of Object.keys(cd[r])) {
        const x = cd[r][c]
        if (!x) continue
        const cell = ws.getCell(+r + 1, +c + 1)
        let f = x.f
        if (!f && x.si && shared[x.si]) { const s0 = shared[x.si]; f = shiftFormula(s0.f, +r - s0.r, +c - s0.c) }
        const raw = x.p && x.p.body ? (x.p.body.dataStream || '').replace(/[\r\n]+$/, '').replace(/\r/g, '\n') : x.v
        if (f) {
          const res = raw === undefined || raw === null ? undefined : (x.t === 2 || typeof raw === 'number' ? Number(raw) : raw)
          cell.value = { formula: String(f).replace(/^=/, ''), result: res }
        } else if (raw !== undefined && raw !== null && raw !== '') {
          if (x.t === 3) cell.value = !!(raw === true || raw === 1 || raw === '1' || raw === 'TRUE')
          else if (x.t === 2 || (typeof raw === 'number')) cell.value = Number(raw)
          else cell.value = String(raw)
        }
        excelStyle(typeof x.s === 'string' ? styles[x.s] : x.s, cell)
      }
    }
    for (const c of Object.keys(sh.columnData || {})) {
      const e = sh.columnData[c]
      const col = ws.getColumn(+c + 1)
      if (e.w) col.width = Math.max(1, (e.w - 5) / 7)
      if (e.hd) col.hidden = true
    }
    for (const r of Object.keys(sh.rowData || {})) {
      const e = sh.rowData[r]
      const row = ws.getRow(+r + 1)
      if (e.h) row.height = e.h * 0.75
      if (e.hd) row.hidden = true
    }
    for (const m of (sh.mergeData || [])) {
      try { ws.mergeCells(m.startRow + 1, m.startColumn + 1, m.endRow + 1, m.endColumn + 1) } catch (e) {}
    }
    const fz = sh.freeze
    if (fz && (fz.xSplit > 0 || fz.ySplit > 0)) ws.views = [{ state: 'frozen', xSplit: fz.xSplit || 0, ySplit: fz.ySplit || 0 }]
  }
  return await wb.xlsx.writeBuffer()
}

// ------------------------------------------------------------ start i zapis
function snapshot() {
  const wb = univerAPI.getActiveWorkbook()
  return wb ? wb.save() : null
}

function hash(s) { let h = 0; for (let i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0; return h + ':' + s.length }

window.biuroSave = async (asCopy) => {
  try {
    status('Zapisuję…')
    const snap = snapshot()
    if (!snap) throw new Error('Brak arkusza')
    const buf = await toXlsx(snap)
    lastHash = hash(JSON.stringify(snap))
    bridge.save(await toB64(buf), 'xlsx', !!asCopy || complex.length > 0)
    status('')
  } catch (e) {
    status('')
    bridge.error('Zapis nieudany: ' + (e && e.message ? e.message : e))
  }
}

window.biuroView = setView

// autozapis co 30 s, jeśli coś się zmieniło
setInterval(() => {
  if (!univerAPI || !bridge.autoSave() || complex.length) return
  try {
    const s = snapshot()
    if (!s) return
    const h = hash(JSON.stringify(s))
    if (lastHash && h !== lastHash) { bridge.dirty(true); window.biuroSave(false) }
  } catch (e) {}
}, 30000)

setInterval(() => {
  try {
    const s = snapshot()
    if (s && lastHash && hash(JSON.stringify(s)) !== lastHash) bridge.dirty(true)
  } catch (e) {}
}, 4000)

document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); window.biuroSave(false) }
})

async function main() {
  status('Otwieram…')
  const info = JSON.parse(bridge.fileInfo())
  const res = createUniver({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS) },
    presets: [UniverSheetsCorePreset({ container: 'app' })],
  })
  univerAPI = res.univerAPI
  let data = null
  try {
    if (info.hasFile) {
      const buf = await (await fetch('/file/current')).arrayBuffer()
      const ext = (info.ext || '').toLowerCase()
      if (ext === 'xlsx' || ext === 'xlsm' || ext === 'xltx' || ext === 'xltm') {
        try {
          const zip = await JSZip.loadAsync(buf)
          const names = Object.keys(zip.files)
          if (names.some(n => n.startsWith('xl/charts/'))) complex.push('wykresy')
          if (names.some(n => n.startsWith('xl/pivotTables/'))) complex.push('tabele przestawne')
          if (names.some(n => n.endsWith('vbaProject.bin'))) complex.push('makra')
          if (names.some(n => n.startsWith('xl/media/'))) complex.push('obrazy')
        } catch (e) {}
        data = await fromXlsx(buf)
      } else if (ext === 'csv' || ext === 'tsv' || ext === 'txt') {
        data = fromCsv(new TextDecoder('utf-8').decode(buf))
      } else {
        data = fromOther(buf)
      }
    } else if (info.template) {
      data = template(info.template)
    }
  } catch (e) {
    bridge.error('Nie mogę odczytać pliku: ' + (e && e.message ? e.message : e))
  }
  univerAPI.createWorkbook(data || { sheetOrder: ['sheet1'], sheets: { sheet1: { id: 'sheet1', name: 'Arkusz1', rowCount: 1000, columnCount: 26 } } })
  setTimeout(() => { try { lastHash = hash(JSON.stringify(snapshot())) } catch (e) {} }, 1500)
  if (complex.length) bridge.complex(complex.join(', '))
  status('')
}

main()
