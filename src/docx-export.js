// Quill delta -> Word (.docx) that opens cleanly in Microsoft Word and keeps
// the layout: A4 page, TH Sarabun PSK, first-line/hanging indents (ย่อหน้า),
// spacing, real tabs, lists that restart, Thai distributed justification,
// images and tables (multi-line cells, merged cells, widths, colours).
import {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, ImageRun, Tab,
  Table, TableRow, TableCell, WidthType, HeightRule, TableLayoutType,
  HeadingLevel, AlignmentType, BorderStyle, LevelFormat, ShadingType, LineRuleType
} from 'docx'

export const TABLE_CELL = 'table-up-cell-inner'
export const TABLE_COL = 'table-up-col'
export const TABLE_CAPTION = 'table-up-caption'

// default page: A4 with 2.54 cm margins (twips); a document opened from Word keeps its own
export const DEFAULT_PAGE = { width: 11906, height: 16838, top: 1440, right: 1440, bottom: 1440, left: 1440 }
export function normalizePage (page) {
  const p = { ...DEFAULT_PAGE }
  for (const k of Object.keys(DEFAULT_PAGE)) {
    const v = Number(page && page[k])
    if (Number.isFinite(v) && v >= 0 && v < 40000) p[k] = Math.round(v)
  }
  if (p.width - p.left - p.right < 2000) Object.assign(p, { left: DEFAULT_PAGE.left, right: DEFAULT_PAGE.right })
  return p
}
export const textWidthOf = (page) => page.width - page.left - page.right

const THAI = /[฀-๿]/
// characters that are illegal in XML 1.0 make Word refuse to open the file
const XML_BAD = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g
export const cleanText = (s) => String(s).replace(XML_BAD, '')

const HEADINGS = {
  1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4, 5: HeadingLevel.HEADING_5, 6: HeadingLevel.HEADING_6
}
const LIST_FORMATS = [LevelFormat.DECIMAL, LevelFormat.LOWER_LETTER, LevelFormat.LOWER_ROMAN]

// ---------------------------------------------------------------- units & colours
export function ptFromSize (v) {
  if (!v) return null
  const m = /^([\d.]+)\s*(pt|px)?$/i.exec(String(v))
  if (!m) return null
  const n = parseFloat(m[1])
  return (m[2] || '').toLowerCase() === 'px' ? n * 0.75 : n
}
export function twips (v, basePt = 16) {
  if (v == null || v === '') return 0
  const m = /^(-?[\d.]+)\s*(cm|mm|in|pt|px|em)?$/i.exec(String(v).trim())
  if (!m) return 0
  const n = parseFloat(m[1])
  switch ((m[2] || 'px').toLowerCase()) {
    case 'cm': return Math.round(n * 567)
    case 'mm': return Math.round(n * 56.7)
    case 'in': return Math.round(n * 1440)
    case 'pt': return Math.round(n * 20)
    case 'em': return Math.round(n * basePt * 20)
    default: return Math.round(n * 15)
  }
}
const NAMED = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080' }
export function hex (c) {
  if (!c) return undefined
  c = String(c).trim().toLowerCase()
  if (c === 'transparent' || c === 'none' || c === 'initial' || c === 'inherit') return undefined
  if (NAMED[c]) return NAMED[c]
  let m = /^#([0-9a-f]{6})$/i.exec(c)
  if (m) return m[1].toUpperCase()
  m = /^#([0-9a-f]{3})$/i.exec(c)
  if (m) return m[1].split('').map((x) => x + x).join('').toUpperCase()
  m = /^rgba?\(\s*(\d+)\D+(\d+)\D+(\d+)(?:\D+([\d.]+))?/i.exec(c)
  if (m) {
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return undefined
    return [m[1], m[2], m[3]].map((n) => (+n).toString(16).padStart(2, '0')).join('').toUpperCase()
  }
  return undefined
}
function parseStyle (s) {
  const out = {}
  for (const part of String(s || '').split(';')) {
    const i = part.indexOf(':')
    if (i > 0) out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim()
  }
  return out
}

// ---------------------------------------------------------------- delta -> lines
function toLines (ops) {
  const lines = []
  const cols = {}
  let segs = []
  for (const op of ops) {
    const a = op.attributes || {}
    if (typeof op.insert === 'string') {
      const parts = op.insert.split('\n')
      for (let i = 0; i < parts.length; i++) {
        if (parts[i]) segs.push({ text: cleanText(parts[i]), a })
        if (i < parts.length - 1) { lines.push({ segs, a }); segs = [] }
      }
    } else if (op.insert && op.insert[TABLE_COL]) {
      const c = op.insert[TABLE_COL]
      ;(cols[c.tableId] = cols[c.tableId] || []).push(c)
    } else if (op.insert && op.insert.image) {
      segs.push({ image: op.insert.image, a })
    }
  }
  if (segs.length) lines.push({ segs, a: {} })
  return { lines, cols }
}

// ---------------------------------------------------------------- runs & paragraphs
function runOptions (text, a, block, o) {
  const opt = { font: a.font ? o.fontName(a.font) : o.defaultFont }
  if (text.includes('\t')) {
    const children = []
    text.split('\t').forEach((part, i) => { if (i > 0) children.push(new Tab()); if (part) children.push(part) })
    opt.children = children
  } else {
    opt.text = text
  }
  // only switch formatting ON, so Word styles (e.g. bold headings) still apply
  if (a.bold) opt.bold = true
  if (a.italic) opt.italics = true
  if (a.strike) opt.strike = true
  if (a.underline) opt.underline = {}
  const pt = ptFromSize(a.size)
  if (pt) { opt.size = Math.round(pt * 2); opt.sizeComplexScript = Math.round(pt * 2) }
  const color = hex(a.color)
  if (color) opt.color = color
  const bg = hex(a.background)
  if (bg) opt.shading = { type: ShadingType.CLEAR, color: 'auto', fill: bg }
  if (a.script === 'super') opt.superScript = true
  if (a.script === 'sub') opt.subScript = true
  if (a.code || block['code-block']) opt.font = 'Courier New'
  if (a.code) opt.shading = { type: ShadingType.CLEAR, color: 'auto', fill: 'F3F4F6' }
  return opt
}
function makeRuns (segs, block, o) {
  const out = []
  for (const s of segs) {
    if (s.image !== undefined) {
      if (s.img) out.push(new ImageRun({ type: s.img.type, data: s.img.data, transformation: { width: s.img.width, height: s.img.height } }))
      continue
    }
    if (s.a.link) {
      out.push(new ExternalHyperlink({
        link: s.a.link,
        children: [new TextRun({ ...runOptions(s.text, s.a, block, o), style: 'Hyperlink', color: '0563C1', underline: {} })]
      }))
    } else {
      out.push(new TextRun(runOptions(s.text, s.a, block, o)))
    }
  }
  return out
}
function makeParagraph (line, ctx, o) {
  const a = line.a
  const p = { children: makeRuns(line.segs, a, o) }
  if (HEADINGS[a.header]) p.heading = HEADINGS[a.header]
  const level = Math.min(Math.max(a.indent || 0, 0), 8)
  const indent = {}
  if (a.list === 'ordered') {
    if (ctx.prevList !== 'ordered') ctx.listInstance++
    p.numbering = { reference: 'numbers', level, instance: ctx.listInstance }
  } else if (a.list === 'bullet') {
    p.bullet = { level }
  } else if (a.list === 'checked' || a.list === 'unchecked') {
    p.children.unshift(new TextRun({ text: a.list === 'checked' ? '☑ ' : '☐ ', font: 'Segoe UI Symbol' }))
    if (level) indent.left = 720 * level
  } else if (level) {
    indent.left = 720 * level
  }
  ctx.prevList = a.list === 'ordered' ? 'ordered' : (a.list || null)
  if (a.leftindent) indent.left = (indent.left || 0) + twips(a.leftindent, o.defaultPt)
  if (a.rightindent) indent.right = twips(a.rightindent, o.defaultPt)
  if (a.firstline) {
    const fl = twips(a.firstline, o.defaultPt)
    if (fl >= 0) indent.firstLine = fl
    else { indent.hanging = -fl; indent.left = Math.max(indent.left || 0, -fl) }
  }
  if (a.blockquote) indent.left = (indent.left || 0) + 720
  if (Object.keys(indent).length) p.indent = indent
  const spacing = {}
  if (a.spacebefore) spacing.before = twips(a.spacebefore, o.defaultPt)
  if (a.spaceafter) spacing.after = twips(a.spaceafter, o.defaultPt)
  if (a.linespacing) {
    const v = String(a.linespacing).trim()
    if (/pt$/i.test(v)) { spacing.line = twips(v); spacing.lineRule = LineRuleType.AT_LEAST } else if (!isNaN(parseFloat(v))) { spacing.line = Math.round(parseFloat(v) * 240); spacing.lineRule = LineRuleType.AUTO }
  }
  if (Object.keys(spacing).length) p.spacing = spacing
  const thai = line.segs.some((s) => s.text && THAI.test(s.text))
  if (a.align === 'center') p.alignment = AlignmentType.CENTER
  else if (a.align === 'right') p.alignment = AlignmentType.RIGHT
  else if (a.align === 'justify') p.alignment = thai ? AlignmentType.THAI_DISTRIBUTE : AlignmentType.JUSTIFIED
  if (a.blockquote) p.border = { left: { style: BorderStyle.SINGLE, size: 12, color: 'CCCCCC', space: 8 } }
  if (a['code-block']) p.shading = { type: ShadingType.CLEAR, color: 'auto', fill: 'F3F4F6' }
  if (a.direction === 'rtl') p.bidirectional = true
  return new Paragraph(p)
}

// ---------------------------------------------------------------- tables
const LINE = { style: BorderStyle.SINGLE, size: 4, color: '000000' }
const NONE = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
const ALL_LINES = { top: LINE, bottom: LINE, left: LINE, right: LINE, insideHorizontal: LINE, insideVertical: LINE }

// on-screen guide lines for borderless Word tables (see docx-import.js); exported as "no border"
export const GUIDE_BORDER = 'border: 1px dashed rgb(203, 213, 225)'
function cellBorders (st) {
  const b = st.border || st['border-color'] || st['border-style']
  if (!b) return undefined
  if (/\bnone\b|\bhidden\b|transparent|\b0px\b/i.test(b) || (/dashed/i.test(b) && /203,\s*213,\s*225|#cbd5e1/i.test(b))) {
    return { top: NONE, bottom: NONE, left: NONE, right: NONE }
  }
  const color = hex((/(#[0-9a-f]{3,6}|rgba?\([^)]*\)|\b[a-z]+\b)\s*$/i.exec(b) || [])[1]) || '000000'
  const style = /dashed/i.test(b) ? BorderStyle.DASHED : /dotted/i.test(b) ? BorderStyle.DOTTED : /double/i.test(b) ? BorderStyle.DOUBLE : BorderStyle.SINGLE
  const line = { style, size: 4, color }
  return { top: line, bottom: line, left: line, right: line }
}
function scaleWidths (widths, textWidth) {
  const total = widths.reduce((s, w) => s + w, 0) || 1
  const scale = total > textWidth ? textWidth / total : 1
  return widths.map((w) => Math.max(200, Math.round(w * scale)))
}
function tableUpToDocx (group, colDefs, o) {
  // rows and cells in document order
  const rowOrder = []
  const rows = new Map()
  for (const line of group) {
    const c = line.a[TABLE_CELL]
    if (!rows.has(c.rowId)) { rows.set(c.rowId, new Map()); rowOrder.push(c.rowId) }
    const row = rows.get(c.rowId)
    if (!row.has(c.colId)) row.set(c.colId, { c, lines: [] })
    const { [TABLE_CELL]: _ignored, ...blockAttrs } = line.a
    row.get(c.colId).lines.push({ segs: line.segs, a: blockAttrs })
  }
  // rows that are completely covered by a row-span are stored as "emptyRow"
  for (const rowId of [...rowOrder]) {
    let at = rowOrder.indexOf(rowId) + 1
    for (const cell of rows.get(rowId).values()) {
      for (const er of cell.c.emptyRow || []) {
        if (!rows.has(er)) { rows.set(er, new Map()); rowOrder.splice(at++, 0, er) }
      }
    }
  }
  let colIds = colDefs.map((d) => d.colId)
  if (!colIds.length) colIds = [...new Set(group.map((l) => l.a[TABLE_CELL].colId))]
  const full = colDefs.length ? colDefs.every((d) => d.full) : true
  let widths = colIds.map((id) => {
    const d = colDefs.find((x) => x.colId === id)
    const w = d ? parseFloat(d.width) : NaN
    if (isNaN(w)) return o.textWidth / colIds.length
    return full ? (w / 100) * o.textWidth : w * 15
  })
  widths = scaleWidths(widths, o.textWidth)
  const total = widths.reduce((s, w) => s + w, 0)
  const docRows = rowOrder.map((rowId) => {
    const cells = [...rows.get(rowId).values()]
      .sort((x, y) => colIds.indexOf(x.c.colId) - colIds.indexOf(y.c.colId))
    let height
    const children = cells.map(({ c, lines }) => {
      const st = parseStyle(c.style)
      const start = Math.max(0, colIds.indexOf(c.colId))
      const span = Math.max(1, Number(c.colspan) || 1)
      const width = widths.slice(start, start + span).reduce((s, w) => s + w, 0)
      if (st.height) height = Math.max(height || 0, twips(st.height))
      const ctx = { listInstance: o.ctx.listInstance, prevList: null }
      const paragraphs = lines.map((l) => makeParagraph(l, ctx, o))
      o.ctx.listInstance = ctx.listInstance
      const cell = {
        children: paragraphs.length ? paragraphs : [new Paragraph('')],
        width: { size: width, type: WidthType.DXA }
      }
      if (span > 1) cell.columnSpan = span
      const rs = Math.max(1, Number(c.rowspan) || 1)
      if (rs > 1) cell.rowSpan = rs
      const bg = hex(st['background-color'] || st.background)
      if (bg) cell.shading = { type: ShadingType.CLEAR, color: 'auto', fill: bg }
      const borders = cellBorders(st)
      if (borders) cell.borders = borders
      return new TableCell(cell)
    })
    const row = { children }
    if (height) row.height = { value: height, rule: HeightRule.ATLEAST }
    if (cells.some(({ c }) => c.wrapTag === 'thead')) row.tableHeader = true
    return new TableRow(row)
  })
  const align = colDefs[0] && colDefs[0].align
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    alignment: align === 'center' ? AlignmentType.CENTER : align === 'right' ? AlignmentType.RIGHT : undefined,
    borders: ALL_LINES,
    rows: docRows
  })
}
// tables made with the older built-in table (one line per cell, attribute `table: rowId`)
function legacyTableToDocx (group, o) {
  const rows = []
  let cur = null
  for (const line of group) {
    if (!cur || cur.id !== line.a.table) { cur = { id: line.a.table, cells: [] }; rows.push(cur) }
    const { table, ...rest } = line.a
    cur.cells.push({ segs: line.segs, a: rest })
  }
  const cols = Math.max(1, ...rows.map((r) => r.cells.length))
  const widths = scaleWidths(Array(cols).fill(o.textWidth / cols), o.textWidth)
  return new Table({
    width: { size: o.textWidth, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    borders: ALL_LINES,
    rows: rows.map((r) => new TableRow({
      children: Array.from({ length: cols }, (_, i) => new TableCell({
        width: { size: widths[i], type: WidthType.DXA },
        children: [r.cells[i] ? makeParagraph(r.cells[i], { listInstance: 0, prevList: null }, o) : new Paragraph('')]
      }))
    }))
  })
}

// ---------------------------------------------------------------- document
/**
 * @param {Array} ops Quill delta ops
 * @param {object} opts { defaultFont, defaultPt, fontName(id), loadImage(src, attrs, maxWidthPx) -> {type,data,width,height}|null, title, page }
 * @returns {Promise<{doc: Document, stats: {skippedImages:number}}>}
 */
export async function deltaToDocument (ops, opts) {
  const page = normalizePage(opts.page)
  const o = {
    page,
    textWidth: textWidthOf(page),
    defaultFont: opts.defaultFont || 'TH SarabunPSK',
    defaultPt: opts.defaultPt || 16,
    fontName: opts.fontName || ((x) => x),
    ctx: { listInstance: 0, prevList: null }
  }
  const { lines, cols } = toLines(ops)
  let skippedImages = 0
  for (const line of lines) {
    for (const s of line.segs) {
      if (s.image === undefined) continue
      s.img = opts.loadImage ? await opts.loadImage(s.image, s.a, Math.round(o.textWidth / 15)).catch(() => null) : null
      if (!s.img) skippedImages++
    }
  }
  const blocks = []
  let i = 0
  while (i < lines.length) {
    const a = lines[i].a
    if (a[TABLE_CELL]) {
      const id = a[TABLE_CELL].tableId
      const group = []
      while (i < lines.length && lines[i].a[TABLE_CELL] && lines[i].a[TABLE_CELL].tableId === id) group.push(lines[i++])
      blocks.push(tableUpToDocx(group, cols[id] || [], o))
      o.ctx.prevList = null
      continue
    }
    if (a.table) {
      const group = []
      while (i < lines.length && lines[i].a.table) group.push(lines[i++])
      blocks.push(legacyTableToDocx(group, o))
      o.ctx.prevList = null
      continue
    }
    if (a[TABLE_CAPTION]) {
      const { [TABLE_CAPTION]: _c, ...rest } = a
      blocks.push(makeParagraph({ segs: lines[i].segs, a: { align: 'center', ...rest } }, o.ctx, o))
      i++
      continue
    }
    blocks.push(makeParagraph(lines[i], o.ctx, o))
    i++
  }
  // Word needs a paragraph after a table at the end of the document
  if (!blocks.length || blocks[blocks.length - 1] instanceof Table) blocks.push(new Paragraph(''))

  const F = o.defaultFont
  const hRun = (size) => ({ font: F, size, sizeComplexScript: size, bold: true, color: '000000' })
  const doc = new Document({
    creator: 'Shared Word',
    title: opts.title || 'Document',
    styles: {
      default: {
        document: { run: { font: F, size: o.defaultPt * 2, sizeComplexScript: o.defaultPt * 2 } },
        heading1: { run: hRun(48), paragraph: { spacing: { before: 240, after: 120 } } },
        heading2: { run: hRun(40), paragraph: { spacing: { before: 200, after: 100 } } },
        heading3: { run: hRun(36), paragraph: { spacing: { before: 160, after: 80 } } },
        heading4: { run: hRun(32) },
        heading5: { run: hRun(32) },
        heading6: { run: hRun(32) },
        hyperlink: { run: { color: '0563C1', underline: {} } }
      }
    },
    numbering: {
      config: [{
        reference: 'numbers',
        levels: Array.from({ length: 9 }, (_, l) => ({
          level: l,
          format: LIST_FORMATS[l % 3],
          text: `%${l + 1}.`,
          alignment: AlignmentType.START,
          style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } }
        }))
      }]
    },
    sections: [{
      properties: {
        page: {
          size: { width: page.width, height: page.height },
          margin: { top: page.top, right: page.right, bottom: page.bottom, left: page.left }
        }
      },
      children: blocks
    }]
  })
  return { doc, stats: { skippedImages } }
}

export async function deltaToDocxBlob (ops, opts) {
  const { doc, stats } = await deltaToDocument(ops, opts)
  return { blob: await Packer.toBlob(doc), stats }
}
export async function deltaToDocxBuffer (ops, opts) {
  const { doc, stats } = await deltaToDocument(ops, opts)
  return { buffer: await Packer.toBuffer(doc), stats }
}
