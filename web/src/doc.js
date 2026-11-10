import { SuperDoc } from 'superdoc'
import 'superdoc/style.css'
import { bridge, toB64, status, setView } from './common.js'

let sd = null
let dirty = false
let pendingText = null

function esc(s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') }

async function exportDocx() {
  // różne wersje SuperDoc mają różne metody eksportu – próbujemy po kolei
  if (sd && typeof sd.export === 'function') {
    try {
      const b = await sd.export({ exportType: ['docx'], triggerDownload: false, isFinalDoc: false })
      if (b instanceof Blob) return b
      if (b && b.docx instanceof Blob) return b.docx
    } catch (e) {}
  }
  const ed = sd && (sd.activeEditor || (sd.editors && sd.editors[0]))
  if (ed && typeof ed.exportDocx === 'function') {
    const b = await ed.exportDocx({ isFinalDoc: false })
    if (b) return b instanceof Blob ? b : new Blob([b])
  }
  throw new Error('Ta wersja edytora nie udostępnia eksportu DOCX')
}

window.biuroSave = async (asCopy) => {
  try {
    status('Zapisuję…')
    const blob = await exportDocx()
    bridge.save(await toB64(blob), 'docx', !!asCopy)
    dirty = false
    bridge.dirty(false)
    status('')
  } catch (e) {
    status('')
    bridge.error('Zapis nieudany: ' + (e && e.message ? e.message : e))
  }
}

window.biuroView = setView

setInterval(() => { if (dirty && bridge.autoSave()) window.biuroSave(false) }, 30000)

document.addEventListener('keydown', (e) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); window.biuroSave(false) }
})

async function main() {
  status('Otwieram…')
  const info = JSON.parse(bridge.fileInfo())
  let file = null
  try {
    if (info.hasFile) {
      const buf = await (await fetch('/file/current')).arrayBuffer()
      const ext = (info.ext || '').toLowerCase()
      if (ext === 'docx' || ext === 'docm' || ext === 'dotx') {
        file = new File([buf], info.name || 'dokument.docx', { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' })
      } else {
        pendingText = new TextDecoder('utf-8').decode(buf)
      }
    }
  } catch (e) {
    bridge.error('Nie mogę odczytać pliku: ' + (e && e.message ? e.message : e))
  }
  const cfg = {
    selector: '#editor',
    toolbar: '#toolbar',
    documentMode: 'editing',
    pagination: true,
    rulers: true,
    onReady: () => {
      status('')
      if (pendingText) {
        try {
          const ed = sd.activeEditor || (sd.editors && sd.editors[0])
          const html = pendingText.split(/\r?\n/).map(l => '<p>' + esc(l) + '</p>').join('')
          ed.commands.insertContent(html)
          dirty = true; bridge.dirty(true)
        } catch (e) {}
      }
    },
    onEditorUpdate: () => { dirty = true; bridge.dirty(true) },
    onContentError: (e) => bridge.error('Problem z plikiem: ' + (e && e.error ? e.error.message : '')),
  }
  if (file) cfg.document = file
  sd = new SuperDoc(cfg)
  setTimeout(() => status(''), 8000)
}

main()
