// Shared Word – a real-time collaborative Word-style document for GitHub Pages.
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
import TableUp, { defaultCustomSelect, TableSelection, TableMenuContextmenu, TableMenuSelect, TableResizeLine, TableResizeScale, TableAlign } from 'quill-table-up'
import { FONTS, DEFAULT_DOCX_FONT, mapFontName, docxFontName, injectFontCss, fillFontSelect } from './fonts.js'
import { docxToDelta } from './docx-import.js'
import { deltaToDocxBlob, normalizePage } from './docx-export.js'
import 'quill/dist/quill.snow.css'
import 'quill-cursors/css'
import 'quill-table-up/index.css'
import 'quill-table-up/table-creator.css'
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
const DEFAULT_PT = 16
const FIRSTLINE_DEFAULT = '2.5cm' // ย่อหน้า used in Thai official documents
const UA = navigator.userAgent || ''
const IS_LINE = /\bLine\//i.test(UA)
const IN_APP = IS_LINE || /FBAN|FBAV|FB_IAB|FBIOS|Instagram|MicroMessenger|KAKAOTALK|TikTok|musical_ly|BytedanceWebview|Twitter|Snapchat|LinkedInApp|\bGSA\//i.test(UA)
const IS_IOS = /iPad|iPhone|iPod/.test(UA) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)
const MOBILE = IS_IOS || /Android/i.test(UA)

// ---------------------------------------------------------------- room
const CFG = Object.assign(
  { room: 'sharedword', websocket: 'wss://demos.yjs.dev/ws', signaling: [] },
  window.SHAREDWORD_CONFIG || {}
)
const hashName = decodeURIComponent(location.hash.replace(/^#/, '')).trim().replace(/[^\w฀-๿.-]+/g, '-')
const room = hashName ? `${CFG.room}--${encodeURIComponent(hashName)}` : CFG.room

// ---------------------------------------------------------------- Yjs
const ydoc = new Y.Doc()
const ytext = ydoc.getText('quill')
const ypage = ydoc.getMap('page') // page size & margins (twips), shared like the text
const persistence = new IndexeddbPersistence(room, ydoc)
const ws = new WebsocketProvider(CFG.websocket, room, ydoc)
const awareness = ws.awareness
let rtc = null
try {
  if (CFG.signaling && CFG.signaling.length) rtc = new WebrtcProvider(room, ydoc, { awareness, signaling: CFG.signaling })
} catch (e) {
  console.warn('WebRTC provider unavailable', e)
}

const user = loadUser()
awareness.setLocalStateField('user', user)

// ---------------------------------------------------------------- Quill formats
const Parchment = Quill.import('parchment')
const blockStyle = (name, key) => new Parchment.StyleAttributor(name, key, { scope: Parchment.Scope.BLOCK })
// Word-like paragraph layout: first-line indent (ย่อหน้า), left/right indent, spacing
const FirstLine = blockStyle('firstline', 'text-indent')
const LeftIndent = blockStyle('leftindent', 'margin-left')
const RightIndent = blockStyle('rightindent', 'margin-right')
const SpaceBefore = blockStyle('spacebefore', 'margin-top')
const SpaceAfter = blockStyle('spaceafter', 'margin-bottom')
const LineSpacing = blockStyle('linespacing', 'line-height')
Quill.register({
  'formats/firstline': FirstLine,
  'formats/leftindent': LeftIndent,
  'formats/rightindent': RightIndent,
  'formats/spacebefore': SpaceBefore,
  'formats/spaceafter': SpaceAfter,
  'formats/linespacing': LineSpacing
}, true)

const Font = Quill.import('formats/font')
Font.whitelist = FONTS.map((f) => f.id)
Quill.register(Font, true)
const Size = Quill.import('attributors/style/size')
Size.whitelist = null // any "NNpt" value, like Word
Quill.register(Size, true)
Quill.register('modules/cursors', QuillCursors)
Quill.register({ [`modules/${TableUp.moduleName}`]: TableUp }, true)

injectFontCss()
fillFontSelect($('#toolbar select.ql-font'))

// ---------------------------------------------------------------- Quill editor
const TABLE_TEXTS = {
  fullCheckboxText: 'ตารางเต็มความกว้าง · Full width',
  customBtnText: 'กำหนดขนาดเอง · Custom size',
  confirmText: 'ตกลง · OK',
  cancelText: 'ยกเลิก · Cancel',
  rowText: 'แถว · Rows',
  colText: 'คอลัมน์ · Columns',
  notPositiveNumberError: 'กรุณาใส่จำนวนเต็มบวก · Please enter a positive number',
  custom: 'กำหนดเอง · Custom',
  clear: 'ล้าง · Clear',
  transparent: 'โปร่งใส · Transparent',
  perWidthInsufficient: 'ความกว้างไม่พอ ต้องเปลี่ยนเป็นความกว้างคงที่ ดำเนินการต่อไหม? · Not enough width; switch the table to a fixed width?',
  InsertTop: 'แทรกแถวด้านบน · Insert row above',
  InsertRight: 'แทรกคอลัมน์ด้านขวา · Insert column right',
  InsertBottom: 'แทรกแถวด้านล่าง · Insert row below',
  InsertLeft: 'แทรกคอลัมน์ด้านซ้าย · Insert column left',
  MergeCell: 'ผสานเซลล์ · Merge cells',
  SplitCell: 'แยกเซลล์ · Split cell',
  DeleteRow: 'ลบแถว · Delete row',
  DeleteColumn: 'ลบคอลัมน์ · Delete column',
  DeleteTable: 'ลบตาราง · Delete table',
  BackgroundColor: 'สีพื้นเซลล์ · Cell colour',
  BorderColor: 'สีเส้นขอบ · Border colour',
  FreezeRow: 'ตรึงถึงแถวนี้ · Freeze to this row',
  UnfreezeRow: 'เลิกตรึงแถว · Unfreeze rows',
  FreezeCol: 'ตรึงถึงคอลัมน์นี้ · Freeze to this column',
  UnfreezeCol: 'เลิกตรึงคอลัมน์ · Unfreeze columns',
  SwitchWidth: 'สลับความกว้างตาราง · Switch table width',
  InsertCaption: 'เพิ่มคำอธิบายตาราง · Add table caption',
  ToggleTdBetweenTh: 'สลับเป็นหัวตาราง · Toggle header cell'
}
let quill // assigned below; toolbar handlers run later
const toolbarHandlers = {
  firstline () {
    const cur = quill.getFormat().firstline
    quill.format('firstline', cur ? false : FIRSTLINE_DEFAULT, 'user')
  }
}

quill = new Quill('#editor', {
  theme: 'snow',
  placeholder: 'Start typing… everyone on this page sees your changes live. / พิมพ์ได้เลย ทุกคนจะเห็นทันที',
  modules: {
    cursors: { transformOnTextChange: true },
    toolbar: { container: '#toolbar', handlers: toolbarHandlers },
    [TableUp.moduleName]: {
      full: true,
      fullSwitch: false,
      customSelect: defaultCustomSelect,
      customBtn: true,
      texts: TABLE_TEXTS,
      modules: [
        { module: TableSelection },
        // phones/tablets have no right-click: show the table menu when a cell is tapped
        { module: MOBILE ? TableMenuSelect : TableMenuContextmenu },
        { module: TableResizeLine },
        { module: TableResizeScale },
        { module: TableAlign }
      ]
    },
    history: { userOnly: true }
  }
})
// eslint-disable-next-line no-unused-vars
const binding = new QuillBinding(ytext, quill, awareness)

// right-clicking a table cell opens the table menu even if the cell was not selected first
quill.root.addEventListener('contextmenu', (e) => {
  const td = e.target && e.target.closest && e.target.closest('td.ql-table-cell, th.ql-table-cell')
  if (!td) return
  const tableUp = quill.getModule(TableUp.moduleName)
  const table = td.closest('table')
  // the menu only learns its table from a left mouse-down; tell it directly
  const menu = tableUp.getModule('table-menu-contextmenu')
  if (menu && menu.table !== table) menu.setSelectionTable(table)
  const selection = tableUp.getModule('table-selection')
  if (!selection) return
  const already = (selection.selectedTds || []).some((c) => c.domNode && c.domNode.closest('td, th') === td)
  if (already) return
  // move the text cursor into the clicked cell (like Word), otherwise the table
  // module re-selects whichever cell the cursor was in
  const range = quill.getSelection()
  const [line] = range ? quill.getLine(range.index) : [null]
  if (!line || !td.contains(line.domNode)) {
    const inner = Quill.find(td.querySelector('.ql-table-cell-inner'))
    if (inner) quill.setSelection(quill.getIndex(inner), 0, 'silent')
  }
  const point = { x: e.clientX, y: e.clientY }
  selection.setSelectionTable(table)
  if (typeof selection.recordScrollPosition === 'function') selection.recordScrollPosition()
  selection.setSelectedTds(selection.computeSelectedTds(point, point))
  selection.show()
}, true)

// ---------------------------------------------------------------- page size & margins (like Word's Page Setup)
const pageStyle = document.createElement('style')
document.head.appendChild(pageStyle)
const cm = (twips) => (Math.round((twips / 567) * 100) / 100) + 'cm'
function currentPage () { return normalizePage(ypage.toJSON()) }
function applyPage () {
  const p = currentPage()
  const paper = document.querySelector('.paper')
  paper.style.setProperty('--page-width', cm(p.width))
  paper.style.setProperty('--page-height', cm(p.height))
  paper.style.setProperty('--m-top', cm(p.top))
  paper.style.setProperty('--m-right', cm(p.right))
  paper.style.setProperty('--m-bottom', cm(p.bottom))
  paper.style.setProperty('--m-left', cm(p.left))
  pageStyle.textContent = `@media print { @page { size: ${cm(p.width)} ${cm(p.height)}; margin: ${cm(p.top)} ${cm(p.right)} ${cm(p.bottom)} ${cm(p.left)}; } }`
}
ypage.observe(applyPage)
persistence.whenSynced.then(applyPage)
applyPage()

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

// export menu
const exportBtn = $('#export-btn')
const exportMenu = $('#export-menu')
function closeExportMenu () { exportMenu.hidden = true; exportBtn.setAttribute('aria-expanded', 'false') }
exportBtn.addEventListener('click', (e) => {
  e.stopPropagation()
  exportMenu.hidden = !exportMenu.hidden
  exportBtn.setAttribute('aria-expanded', String(!exportMenu.hidden))
})
document.addEventListener('click', (e) => { if (!exportMenu.contains(e.target)) closeExportMenu() })
exportMenu.addEventListener('click', (e) => {
  const kind = e.target.closest('button') && e.target.closest('button').dataset.export
  if (!kind) return
  closeExportMenu()
  if (kind === 'docx') exportDocx()
  else if (kind === 'pdf') exportPdf()
  else if (kind === 'txt') exportText()
})

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

function trimTrailingNewline (ops) {
  ops = ops.slice()
  const last = ops[ops.length - 1]
  if (last && typeof last.insert === 'string' && !last.attributes && last.insert.endsWith('\n')) {
    last.insert = last.insert.slice(0, -1)
    if (!last.insert) ops.pop()
  }
  return ops
}

async function importDocx (file) {
  if (!/\.docx$/i.test(file.name)) {
    toast('Please choose a Word .docx file (older .doc files must be saved as .docx first)', 4000)
    return
  }
  toast(`Reading ${file.name}…`)
  try {
    const arrayBuffer = await file.arrayBuffer()
    let delta
    try {
      delta = await docxToDelta(arrayBuffer, { mapFont: mapFontName, defaultSize: DEFAULT_PT + 'pt' })
    } catch (err) {
      console.warn('Native importer failed, falling back to Mammoth', err)
      const result = await mammoth.convertToHtml({ arrayBuffer }, { styleMap: ['u => u', 'strike => s'] })
      delta = quill.clipboard.convert({ html: result.value })
    }
    if (!delta.ops || !delta.ops.length) throw new Error('the file looks empty')
    const hasContent = quill.getLength() > 1
    let replace = true
    if (hasContent) {
      replace = confirm(
        'This shared document already has text.\n\n' +
        `OK = replace everything with "${file.name}" (everyone will see this)\n` +
        'Cancel = add the file\'s content at the end instead'
      )
    }
    const Delta = Quill.import('delta')
    if (replace) {
      quill.setContents(new Delta(delta.ops), 'user')
      const page = delta.page && Object.keys(delta.page).length ? normalizePage(delta.page) : null
      ydoc.transact(() => {
        ypage.clear()
        if (page) for (const [k, v] of Object.entries(page)) ypage.set(k, v)
      })
    } else {
      quill.updateContents(new Delta().retain(quill.getLength() - 1).insert('\n').concat(new Delta(trimTrailingNewline(delta.ops))), 'user')
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

// ---------------------------------------------------------------- export (Word / PDF / text)
const IMG_TYPES = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp' }

function externalBrowserUrl () {
  const u = new URL(location.href)
  u.searchParams.set('openExternalBrowser', '1') // LINE opens this in Safari / Chrome
  return u.toString()
}
function docName () {
  const first = (quill.getText(0, 600).split('\n').find((l) => l.trim()) || 'shared-document').trim()
  return first.slice(0, 80).replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-').replace(/[\s.]+$/, '') || 'shared-document'
}

async function loadImage (src, attrs, maxWidthPx = 602) {
  let blob
  try { blob = await (await fetch(src)).blob() } catch { return null }
  const url = URL.createObjectURL(blob)
  try {
    const img = await new Promise((resolve, reject) => {
      const im = new Image()
      im.onload = () => resolve(im)
      im.onerror = reject
      im.src = url
    })
    const w = img.naturalWidth || 400
    const h = img.naturalHeight || 300
    let type = IMG_TYPES[blob.type]
    let data
    if (type) {
      data = new Uint8Array(await blob.arrayBuffer())
    } else { // webp, svg, avif… -> PNG so Word can show it
      const c = document.createElement('canvas')
      c.width = w; c.height = h
      c.getContext('2d').drawImage(img, 0, 0)
      const png = await new Promise((resolve) => c.toBlob(resolve, 'image/png'))
      if (!png) return null
      data = new Uint8Array(await png.arrayBuffer())
      type = 'png'
    }
    const shown = parseFloat(attrs && attrs.width) || w
    const width = Math.min(shown, maxWidthPx)
    return { type, data, width: Math.round(width), height: Math.round(width * h / w) }
  } catch {
    return null
  } finally {
    URL.revokeObjectURL(url)
  }
}

async function buildDocxBlob () {
  const title = docName()
  const { blob, stats } = await deltaToDocxBlob(quill.getContents().ops, {
    defaultFont: DEFAULT_DOCX_FONT,
    defaultPt: DEFAULT_PT,
    fontName: docxFontName,
    loadImage,
    title,
    page: ypage.toJSON()
  })
  return { blob, filename: title + '.docx', stats }
}

// "file ready" sheet: a real link + share button works on phones, tablets and in-app browsers
const sheet = $('#sheet')
const sheetDownload = $('#sheet-download')
const sheetShare = $('#sheet-share')
const sheetExternal = $('#sheet-external')
let sheetUrl = null
let sheetFile = null
function closeSheet () { sheet.hidden = true }
$('#sheet-close').addEventListener('click', closeSheet)
sheet.addEventListener('click', (e) => { if (e.target === sheet) closeSheet() })
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !sheet.hidden) closeSheet() })
sheetShare.addEventListener('click', async () => {
  if (!sheetFile) return
  try {
    await navigator.share({ files: [sheetFile], title: sheetFile.name })
  } catch (e) {
    if (e && e.name !== 'AbortError') toast('Sharing is not available here – use Download instead')
  }
})
sheetExternal.href = externalBrowserUrl()
function showSheet ({ title, file, note, blob, filename }) {
  $('#sheet-title').textContent = title
  $('#sheet-file').textContent = file || ''
  $('#sheet-note').textContent = note || ''
  if (sheetUrl) { URL.revokeObjectURL(sheetUrl); sheetUrl = null }
  sheetFile = null
  if (blob) {
    sheetUrl = URL.createObjectURL(blob)
    sheetDownload.href = sheetUrl
    sheetDownload.download = filename
    sheetDownload.hidden = false
    try {
      const f = new File([blob], filename, { type: blob.type })
      if (navigator.canShare && navigator.canShare({ files: [f] })) sheetFile = f
    } catch {}
  } else {
    sheetDownload.hidden = true
  }
  sheetShare.hidden = !sheetFile
  sheetExternal.hidden = !IN_APP
  sheetExternal.textContent = IS_IOS ? 'เปิดใน Safari · Open in Safari' : 'เปิดใน Chrome · Open in browser'
  sheet.hidden = false
}

async function deliver (blob, filename) {
  const inAppNote = IN_APP
    ? (IS_LINE
        ? 'LINE’s built-in browser can’t save files. Tap “Open in Safari/Chrome”, then export again. · เบราว์เซอร์ใน LINE บันทึกไฟล์ไม่ได้ กด “เปิดใน Safari/Chrome” แล้วส่งออกอีกครั้ง'
        : 'This app’s built-in browser may not save files. Use Share, or open this page in Safari/Chrome (menu ⋯ → Open in browser).')
    : ''
  if (!MOBILE && !IN_APP) {
    // desktop: start the download right away, keep the sheet as a fallback
    const a = document.createElement('a')
    a.href = URL.createObjectURL(blob)
    a.download = filename
    document.body.appendChild(a)
    a.click()
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove() }, 4000)
  }
  showSheet({
    title: 'ไฟล์พร้อมแล้ว · Your file is ready',
    file: filename,
    note: inAppNote || (!MOBILE ? 'Download started. If nothing happened, click Download. · เริ่มดาวน์โหลดแล้ว ถ้าไม่มีอะไรเกิดขึ้น กดปุ่มดาวน์โหลด' : 'Tap Download, or Share to save it to Files / send it in LINE. · กดดาวน์โหลด หรือกดส่งต่อเพื่อบันทึก/ส่งทาง LINE'),
    blob,
    filename
  })
}

let exporting = false
async function exportDocx () {
  if (exporting) return
  exporting = true
  const btn = $('#export-btn')
  btn.disabled = true
  toast('กำลังสร้างไฟล์ Word… · Preparing Word file…', 10000)
  try {
    const { blob, filename, stats } = await buildDocxBlob()
    await deliver(blob, filename)
    toast(stats.skippedImages ? `Word file ready · ${stats.skippedImages} image(s) could not be included` : 'Word file ready · ไฟล์ Word พร้อมแล้ว', 3000)
  } catch (e) {
    console.error(e)
    toast('Could not create the Word file: ' + (e && e.message ? e.message : e), 6000)
  } finally {
    btn.disabled = false
    exporting = false
  }
}
function exportPdf () {
  if (IN_APP) {
    showSheet({
      title: 'เปิดในเบราว์เซอร์เพื่อบันทึก PDF · Open in a browser to save a PDF',
      note: 'In-app browsers (LINE, Facebook, Instagram…) can’t print or save PDFs. Open this page in Safari/Chrome, then Export → PDF. · เปิดใน Safari/Chrome แล้วเลือก ส่งออก → PDF'
    })
    return
  }
  toast(IS_IOS ? 'Tap Share (⬆︎) in the print screen → Save to Files · กดแชร์ → บันทึกไปยังไฟล์' : 'Choose “Save as PDF” as the printer · เลือก “บันทึกเป็น PDF”', 6000)
  quill.blur()
  setTimeout(() => window.print(), 350)
}
function exportText () {
  const blob = new Blob(['﻿' + quill.getText()], { type: 'text/plain;charset=utf-8' })
  deliver(blob, docName() + '.txt')
}
const downloadDocx = exportDocx

// expose a little for debugging in the console
window.sharedword = { Y, ydoc, ypage, ytext, ws, rtc, awareness, quill, room, importDocx, buildDocxBlob, downloadDocx, exportDocx, docxToDelta, FONTS, TableUp }
