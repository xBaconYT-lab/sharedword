// Shared Word – a real-time collaborative document for GitHub Pages.
// Sync: Yjs CRDT over a public websocket relay + WebRTC peer-to-peer, with a
// local IndexedDB copy so the document survives reloads and works offline.
import * as Y from 'yjs'
import { WebsocketProvider } from 'y-websocket'
import { WebrtcProvider } from 'y-webrtc'
import { IndexeddbPersistence } from 'y-indexeddb'
import Quill from 'quill'
import QuillCursors from 'quill-cursors'
import { QuillBinding } from 'y-quill'
import mammoth from 'mammoth'
import {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, ImageRun,
  HeadingLevel, AlignmentType, BorderStyle, LevelFormat, ShadingType
} from 'docx'
import 'quill/dist/quill.snow.css'
import 'quill-cursors/css'
import './style.css'

// ---------------------------------------------------------------- helpers
const $ = (sel) => document.querySelector(sel)
const toastEl = $('#toast')
let toastTimer
function toast (msg, ms = 2400) {
  toastEl.textContent = msg
  toastEl.classList.add('show')
  clearTimeout(toastTimer)
  toastTimer = setTimeout(() => toastEl.classList.remove('show'), ms)
}

const COLORS = ['#e11d48', '#db2777', '#9333ea', '#4f46e5', '#2563eb', '#0891b2', '#059669', '#65a30d', '#d97706', '#ea580c']
const pick = (arr) => arr[Math.floor(Math.random() * arr.length)]
const store = {
  get (k) { try { return localStorage.getItem(k) } catch { return null } },
  set (k, v) { try { localStorage.setItem(k, v) } catch {} }
}
function loadUser () {
  let name = store.get('sharedword:name') || ''
  let color = store.get('sharedword:color') || ''
  if (!name) name = 'Guest ' + Math.floor(10 + Math.random() * 90)
  if (!color) color = pick(COLORS)
  store.set('sharedword:name', name)
  store.set('sharedword:color', color)
  return { name, color }
}
const initials = (n) => (n || '?').trim().split(/\s+/).map((w) => w[0]).join('').slice(0, 2).toUpperCase()

// ---------------------------------------------------------------- room
const CFG = Object.assign(
  { room: 'sharedword', websocket: 'wss://demos.yjs.dev/ws', signaling: ['wss://y-webrtc-eu.fly.dev'] },
  window.SHAREDWORD_CONFIG || {}
)
const hashName = decodeURIComponent(location.hash.replace(/^#/, '')).trim().replace(/[^\w฀-๿.-]+/g, '-')
const room = hashName ? `${CFG.room}--${encodeURIComponent(hashName)}` : CFG.room

// ---------------------------------------------------------------- Yjs
const ydoc = new Y.Doc()
const ytext = ydoc.getText('quill')
const persistence = new IndexeddbPersistence(room, ydoc)
const ws = new WebsocketProvider(CFG.websocket, room, ydoc)
const awareness = ws.awareness
let rtc = null
try {
  rtc = new WebrtcProvider(room, ydoc, { awareness, signaling: CFG.signaling })
} catch (e) {
  console.warn('WebRTC provider unavailable', e)
}

const user = loadUser()
awareness.setLocalStateField('user', user)

// ---------------------------------------------------------------- Quill
const Font = Quill.import('formats/font')
Font.whitelist = ['sarabun', 'arial', 'times', 'courier']
Quill.register(Font, true)
const SIZES = ['10pt', '12pt', '14pt', '16pt', '18pt', '20pt', '24pt', '28pt', '32pt', '36pt', '48pt']
const Size = Quill.import('attributors/style/size')
Size.whitelist = SIZES
Quill.register(Size, true)
Quill.register('modules/cursors', QuillCursors)

const quill = new Quill('#editor', {
  theme: 'snow',
  placeholder: 'Start typing… everyone on this page sees your changes live. / พิมพ์ได้เลย ทุกคนจะเห็นทันที',
  modules: {
    cursors: { transformOnTextChange: true },
    toolbar: [
      [{ font: [false, 'arial', 'times', 'courier'] }, { size: ['10pt', '12pt', '14pt', false, '18pt', '20pt', '24pt', '28pt', '32pt', '36pt', '48pt'] }],
      [{ header: [1, 2, 3, false] }],
      ['bold', 'italic', 'underline', 'strike'],
      [{ color: [] }, { background: [] }],
      [{ list: 'ordered' }, { list: 'bullet' }, { list: 'check' }],
      [{ indent: '-1' }, { indent: '+1' }],
      [{ align: [] }],
      ['blockquote', 'code-block', 'link', 'image'],
      ['clean']
    ],
    history: { userOnly: true }
  }
})
// eslint-disable-next-line no-unused-vars
const binding = new QuillBinding(ytext, quill, awareness)

// ---------------------------------------------------------------- status UI
let wsStatus = 'connecting'
let rtcPeers = 0
const dot = $('#dot')
const statusText = $('#status-text')
function renderStatus () {
  const n = awareness.getStates().size
  const people = `${n} ${n === 1 ? 'person' : 'people'} here`
  let text, cls
  if (wsStatus === 'connected') { text = `Live · ${people}`; cls = 'ok' } else if (rtcPeers > 0) { text = `Live (peer-to-peer) · ${people}`; cls = 'ok' } else if (wsStatus === 'connecting') { text = 'Connecting…'; cls = '' } else { text = 'Offline · saved on this device, will sync when back online'; cls = 'bad' }
  statusText.textContent = text
  dot.className = 'dot ' + cls
}
function renderPeople () {
  const states = [...awareness.getStates().values()].filter((s) => s && s.user)
  const el = $('#people')
  el.innerHTML = ''
  states.slice(0, 6).forEach((s) => {
    const d = document.createElement('span')
    d.className = 'avatar'
    d.style.background = s.user.color
    d.textContent = initials(s.user.name)
    d.title = s.user.name
    el.appendChild(d)
  })
  if (states.length > 6) {
    const d = document.createElement('span')
    d.className = 'avatar more'
    d.textContent = '+' + (states.length - 6)
    el.appendChild(d)
  }
  renderStatus()
}
awareness.on('change', renderPeople)
ws.on('status', (e) => { wsStatus = e.status; renderStatus() })
if (rtc) rtc.on('peers', (e) => { rtcPeers = e.webrtcPeers.length; renderStatus() })
renderPeople()

let titleTimer
function updateTitle () {
  clearTimeout(titleTimer)
  titleTimer = setTimeout(() => {
    const first = ytext.toString().split('\n').find((l) => l.trim())
    document.title = first ? first.trim().slice(0, 60) + ' – Shared Word' : 'Shared Word'
  }, 300)
}
ytext.observe(updateTitle)
persistence.whenSynced.then(updateTitle)

// name
const nameInput = $('#name')
nameInput.value = user.name
nameInput.addEventListener('input', () => {
  const name = nameInput.value.trim() || 'Guest'
  user.name = name
  store.set('sharedword:name', name)
  awareness.setLocalStateField('user', { ...user })
})

// share link
$('#share').addEventListener('click', async () => {
  const url = location.href
  try {
    await navigator.clipboard.writeText(url)
    toast('Link copied – send it to anyone, they can edit right away')
  } catch {
    prompt('Copy this link:', url)
  }
})

// new blank document
$('#new-doc').addEventListener('click', (e) => {
  e.preventDefault()
  const id = Math.random().toString(36).slice(2, 8)
  location.hash = 'doc-' + id
})
window.addEventListener('hashchange', () => location.reload())

// ---------------------------------------------------------------- Word import (.docx -> editor)
const fileInput = document.createElement('input')
fileInput.type = 'file'
fileInput.accept = '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document'
fileInput.style.display = 'none'
document.body.appendChild(fileInput)
$('#upload').addEventListener('click', () => fileInput.click())
fileInput.addEventListener('change', () => {
  const f = fileInput.files && fileInput.files[0]
  if (f) importDocx(f)
  fileInput.value = ''
})

async function importDocx (file) {
  if (!/\.docx$/i.test(file.name)) {
    toast('Please choose a Word .docx file (older .doc files must be saved as .docx first)', 4000)
    return
  }
  toast(`Reading ${file.name}…`)
  try {
    const arrayBuffer = await file.arrayBuffer()
    const result = await mammoth.convertToHtml({ arrayBuffer }, { styleMap: ['u => u', 'strike => s'] })
    const delta = quill.clipboard.convert({ html: result.value })
    if (!delta.ops || !delta.ops.length) throw new Error('the file looks empty')
    const hasContent = quill.getLength() > 1
    let replace = true
    if (hasContent) {
      replace = confirm(
        `This shared document already has text.\n\n` +
        `OK = replace everything with "${file.name}" (everyone will see this)\n` +
        `Cancel = add the file's content at the end instead`
      )
    }
    if (replace) {
      quill.setContents(delta, 'user')
    } else {
      const Delta = Quill.import('delta')
      quill.updateContents(new Delta().retain(quill.getLength() - 1).insert('\n').concat(delta), 'user')
    }
    quill.setSelection(0, 0, 'silent')
    toast(`Loaded ${file.name} – everyone on this page can edit it now`, 3500)
  } catch (e) {
    console.error(e)
    toast('Could not read that file: ' + (e && e.message ? e.message : e), 5000)
  }
}

// drag & drop a .docx anywhere on the page
let dragDepth = 0
document.addEventListener('dragenter', (e) => {
  if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) {
    dragDepth++
    document.body.classList.add('dragging')
  }
})
document.addEventListener('dragleave', () => {
  if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove('dragging') }
})
document.addEventListener('dragover', (e) => {
  if (document.body.classList.contains('dragging')) e.preventDefault()
})
document.addEventListener('drop', (e) => {
  dragDepth = 0
  document.body.classList.remove('dragging')
  const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]
  if (f && /\.docx$/i.test(f.name)) {
    e.preventDefault()
    e.stopPropagation()
    importDocx(f)
  }
}, true)

// ---------------------------------------------------------------- Word export (editor -> .docx)
const DOCX_FONTS = { sarabun: 'TH SarabunPSK', arial: 'Arial', times: 'Times New Roman', courier: 'Courier New' }
const DEFAULT_FONT = 'TH SarabunPSK'
const DEFAULT_PT = 16
const HEADINGS = {
  1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4, 5: HeadingLevel.HEADING_5, 6: HeadingLevel.HEADING_6
}

function ptFromSize (v) {
  if (!v) return null
  const m = /^([\d.]+)\s*(pt|px)?$/i.exec(String(v))
  if (!m) return null
  const n = parseFloat(m[1])
  return (m[2] || '').toLowerCase() === 'px' ? n * 0.75 : n
}
function hex (c) {
  if (!c) return undefined
  c = String(c).trim()
  let m = /^#([0-9a-f]{6})$/i.exec(c)
  if (m) return m[1].toUpperCase()
  m = /^#([0-9a-f]{3})$/i.exec(c)
  if (m) return m[1].split('').map((x) => x + x).join('').toUpperCase()
  m = /^rgba?\((\d+)\D+(\d+)\D+(\d+)/i.exec(c)
  if (m) return [m[1], m[2], m[3]].map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase()
  return undefined
}
function runOpts (text, a, block, extra = {}) {
  const pt = ptFromSize(a.size)
  const o = {
    text,
    bold: !!a.bold,
    italics: !!a.italic,
    strike: !!a.strike,
    font: a.font && DOCX_FONTS[a.font] ? DOCX_FONTS[a.font] : DEFAULT_FONT,
    ...extra
  }
  if (a.underline) o.underline = {}
  if (pt) { o.size = Math.round(pt * 2); o.sizeComplexScript = Math.round(pt * 2) }
  const color = hex(a.color)
  if (color) o.color = color
  const bg = hex(a.background)
  if (bg) o.shading = { type: ShadingType.CLEAR, fill: bg }
  if (a.script === 'super') o.superScript = true
  if (a.script === 'sub') o.subScript = true
  if (a.code || block['code-block']) o.font = 'Courier New'
  if (a.code) o.shading = { type: ShadingType.CLEAR, fill: 'F3F4F6' }
  return o
}
function makeRun (seg, block) {
  if (seg.run) return seg.run
  const { text, a } = seg
  if (a.link) {
    return new ExternalHyperlink({
      link: a.link,
      children: [new TextRun(runOpts(text, a, block, { style: 'Hyperlink', color: '0563C1', underline: {} }))]
    })
  }
  return new TextRun(runOpts(text, a, block))
}
function makeParagraph (segments, a) {
  const o = { children: segments.map((s) => makeRun(s, a)) }
  if (HEADINGS[a.header]) o.heading = HEADINGS[a.header]
  const level = Math.min(Math.max(a.indent || 0, 0), 8)
  if (a.list === 'bullet') {
    o.bullet = { level }
  } else if (a.list === 'ordered') {
    o.numbering = { reference: 'numbers', level }
  } else if (a.list === 'checked' || a.list === 'unchecked') {
    o.children.unshift(new TextRun({ text: a.list === 'checked' ? '☑ ' : '☐ ', font: 'Segoe UI Symbol' }))
    if (level) o.indent = { left: 720 * level }
  } else if (level) {
    o.indent = { left: 720 * level }
  }
  if (a.align === 'center') o.alignment = AlignmentType.CENTER
  else if (a.align === 'right') o.alignment = AlignmentType.RIGHT
  else if (a.align === 'justify') o.alignment = AlignmentType.JUSTIFIED
  if (a.blockquote) {
    o.indent = { left: 720 }
    o.border = { left: { style: BorderStyle.SINGLE, size: 12, color: 'CCCCCC', space: 8 } }
  }
  if (a['code-block']) o.shading = { type: ShadingType.CLEAR, fill: 'F3F4F6' }
  if (a.direction === 'rtl') o.bidirectional = true
  return new Paragraph(o)
}
async function imageRun (src) {
  const m = /^data:(image\/(png|jpeg|jpg|gif|bmp));base64,(.+)$/i.exec(src || '')
  if (!m) return null
  let type = m[2].toLowerCase()
  if (type === 'jpeg') type = 'jpg'
  const bin = atob(m[3])
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  const dims = await new Promise((resolve) => {
    const im = new Image()
    im.onload = () => resolve({ w: im.naturalWidth || 400, h: im.naturalHeight || 300 })
    im.onerror = () => resolve({ w: 400, h: 300 })
    im.src = src
  })
  const maxW = 600
  const scale = dims.w > maxW ? maxW / dims.w : 1
  return new ImageRun({ type, data: bytes, transformation: { width: Math.round(dims.w * scale), height: Math.round(dims.h * scale) } })
}
async function deltaToParagraphs (ops) {
  const paragraphs = []
  let segments = []
  const flush = (a) => { paragraphs.push(makeParagraph(segments, a)); segments = [] }
  for (const op of ops) {
    const a = op.attributes || {}
    if (typeof op.insert === 'string') {
      const parts = op.insert.split('\n')
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) segments.push({ text: parts[i], a })
        if (i < parts.length - 1) flush(a)
      }
    } else if (op.insert && op.insert.image) {
      const run = await imageRun(op.insert.image)
      if (run) segments.push({ run })
    }
  }
  if (segments.length) flush({})
  return paragraphs
}
function saveBlob (blob, filename) {
  const a = document.createElement('a')
  a.href = URL.createObjectURL(blob)
  a.download = filename
  document.body.appendChild(a)
  a.click()
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove() }, 1500)
}
async function buildDocxBlob () {
    const ops = quill.getContents().ops
    const paragraphs = await deltaToParagraphs(ops)
    const firstLine = (quill.getText(0, 400).split('\n').find((l) => l.trim()) || 'shared-document').trim()
    const title = firstLine.slice(0, 60)
    const run = (size, bold) => ({ font: DEFAULT_FONT, size, sizeComplexScript: size, bold })
    const doc = new Document({
      creator: 'Shared Word',
      title,
      styles: {
        default: {
          document: { run: { font: DEFAULT_FONT, size: DEFAULT_PT * 2, sizeComplexScript: DEFAULT_PT * 2 } },
          heading1: { run: run(56, true), paragraph: { spacing: { before: 240, after: 120 } } },
          heading2: { run: run(44, true), paragraph: { spacing: { before: 200, after: 100 } } },
          heading3: { run: run(36, true), paragraph: { spacing: { before: 160, after: 80 } } },
          heading4: { run: run(32, true) },
          heading5: { run: run(32, true) },
          heading6: { run: run(32, true) }
        }
      },
      numbering: {
        config: [{
          reference: 'numbers',
          levels: Array.from({ length: 9 }, (_, l) => ({
            level: l,
            format: LevelFormat.DECIMAL,
            text: `%${l + 1}.`,
            alignment: AlignmentType.START,
            style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } }
          }))
        }]
      },
      sections: [{ properties: {}, children: paragraphs.length ? paragraphs : [new Paragraph('')] }]
    })
    const blob = await Packer.toBlob(doc)
    return { blob, filename: title.replace(/[\\/:*?"<>|]+/g, '-') + '.docx' }
}
async function downloadDocx () {
  const btn = $('#download')
  btn.disabled = true
  toast('Preparing Word file…')
  try {
    const { blob, filename } = await buildDocxBlob()
    saveBlob(blob, filename)
    toast('Word file downloaded')
  } catch (e) {
    console.error(e)
    toast('Could not create the Word file: ' + (e && e.message ? e.message : e), 5000)
  } finally {
    btn.disabled = false
  }
}
$('#download').addEventListener('click', downloadDocx)

// expose a little for debugging in the console
window.sharedword = { ydoc, ytext, ws, rtc, awareness, quill, room, importDocx, buildDocxBlob, downloadDocx }
