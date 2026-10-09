// Word (.docx) -> Quill delta importer that keeps paragraph layout:
// first-line indents (ย่อหน้า), left/right indents, alignment (incl. Thai
// distributed), spacing, fonts, sizes, colours, lists, headings, tabs,
// images and tables. Resolves style inheritance (docDefaults -> style chain
// -> paragraph/run) so style-defined formatting is kept too.
import JSZip from 'jszip'
import Delta from 'quill-delta'

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const V = 'urn:schemas-microsoft-com:vml'

const THAI = /[฀-๿]/
const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000'
}

const parseXml = (s) => new DOMParser().parseFromString(s, 'application/xml')
const kids = (el, name) => el ? Array.from(el.children).filter((c) => c.localName === name) : []
const child = (el, name) => kids(el, name)[0] || null
function attr (el, name) {
  if (!el) return null
  const v = el.getAttributeNS(W, name)
  return v != null ? v : el.getAttribute('w:' + name)
}
const num = (v) => (v == null || v === '' || isNaN(+v) ? null : +v)
const onOff = (el, name) => {
  const e = child(el, name)
  if (!e) return undefined
  const v = attr(e, 'val')
  return !(v === '0' || v === 'false' || v === 'off')
}
const trim = (n) => String(Math.round(n * 100) / 100)
const cm = (twips) => trim(twips / 567) + 'cm'
const pt = (twips) => trim(twips / 20) + 'pt'

// ---------------------------------------------------------------- properties
function readPPr (pPr) {
  if (!pPr) return {}
  const o = {}
  o.styleId = attr(child(pPr, 'pStyle'), 'val') || undefined
  o.jc = attr(child(pPr, 'jc'), 'val') || undefined
  const ind = child(pPr, 'ind')
  if (ind) {
    o.indLeft = num(attr(ind, 'left') ?? attr(ind, 'start'))
    o.indRight = num(attr(ind, 'right') ?? attr(ind, 'end'))
    o.firstLine = num(attr(ind, 'firstLine'))
    o.hanging = num(attr(ind, 'hanging'))
  }
  const sp = child(pPr, 'spacing')
  if (sp) {
    o.before = num(attr(sp, 'before'))
    o.after = num(attr(sp, 'after'))
    o.line = num(attr(sp, 'line'))
    o.lineRule = attr(sp, 'lineRule') || undefined
  }
  const np = child(pPr, 'numPr')
  if (np) {
    o.numId = attr(child(np, 'numId'), 'val') ?? undefined
    o.ilvl = attr(child(np, 'ilvl'), 'val') ?? undefined
  }
  o.outlineLvl = num(attr(child(pPr, 'outlineLvl'), 'val'))
  for (const k of Object.keys(o)) if (o[k] === undefined || o[k] === null) delete o[k]
  return o
}

function readRPr (rPr) {
  if (!rPr) return {}
  const o = {}
  o.rStyle = attr(child(rPr, 'rStyle'), 'val') || undefined
  o.b = onOff(rPr, 'b'); o.bCs = onOff(rPr, 'bCs')
  o.i = onOff(rPr, 'i'); o.iCs = onOff(rPr, 'iCs')
  const u = child(rPr, 'u')
  if (u) o.u = (attr(u, 'val') || 'single') !== 'none'
  if (child(rPr, 'strike')) o.strike = onOff(rPr, 'strike')
  if (child(rPr, 'dstrike')) o.strike = onOff(rPr, 'dstrike')
  const color = attr(child(rPr, 'color'), 'val')
  if (color) o.color = color
  o.sz = num(attr(child(rPr, 'sz'), 'val')); o.szCs = num(attr(child(rPr, 'szCs'), 'val'))
  const f = child(rPr, 'rFonts')
  if (f) {
    o.font = attr(f, 'ascii') || attr(f, 'hAnsi') || undefined
    o.fontCs = attr(f, 'cs') || undefined
  }
  const hl = attr(child(rPr, 'highlight'), 'val')
  if (hl && hl !== 'none') o.highlight = hl
  const shd = child(rPr, 'shd')
  if (shd) { const fill = attr(shd, 'fill'); if (fill && fill !== 'auto') o.shd = fill }
  const va = attr(child(rPr, 'vertAlign'), 'val')
  if (va && va !== 'baseline') o.vertAlign = va
  for (const k of Object.keys(o)) if (o[k] === undefined || o[k] === null) delete o[k]
  return o
}

const merge = (...objs) => Object.assign({}, ...objs.map((o) => o || {}))

// ---------------------------------------------------------------- styles / numbering / rels
function parseStyles (xml) {
  const out = { byId: {}, defaults: { pPr: {}, rPr: {} }, defaultParagraphStyle: null }
  if (!xml) return out
  const root = parseXml(xml).documentElement
  const dd = child(root, 'docDefaults')
  if (dd) {
    out.defaults.pPr = readPPr(child(child(dd, 'pPrDefault'), 'pPr'))
    out.defaults.rPr = readRPr(child(child(dd, 'rPrDefault'), 'rPr'))
  }
  for (const s of kids(root, 'style')) {
    const id = attr(s, 'styleId')
    if (!id) continue
    const st = {
      id,
      type: attr(s, 'type') || 'paragraph',
      name: attr(child(s, 'name'), 'val') || '',
      basedOn: attr(child(s, 'basedOn'), 'val') || null,
      pPr: readPPr(child(s, 'pPr')),
      rPr: readRPr(child(s, 'rPr'))
    }
    out.byId[id] = st
    if (st.type === 'paragraph' && (attr(s, 'default') === '1' || attr(s, 'default') === 'true')) out.defaultParagraphStyle = id
  }
  return out
}
function styleChain (styles, id) {
  const chain = []
  let cur = id ? styles.byId[id] : null
  let guard = 0
  while (cur && guard++ < 12) { chain.unshift(cur); cur = cur.basedOn ? styles.byId[cur.basedOn] : null }
  return chain
}

function parseNumbering (xml) {
  const out = { nums: {}, abstracts: {} }
  if (!xml) return out
  const root = parseXml(xml).documentElement
  for (const an of kids(root, 'abstractNum')) {
    const lvls = {}
    for (const l of kids(an, 'lvl')) lvls[attr(l, 'ilvl')] = { fmt: attr(child(l, 'numFmt'), 'val') || 'decimal' }
    out.abstracts[attr(an, 'abstractNumId')] = lvls
  }
  for (const n of kids(root, 'num')) out.nums[attr(n, 'numId')] = attr(child(n, 'abstractNumId'), 'val')
  return out
}
function listInfo (numbering, numId, ilvl) {
  if (!numId || numId === '0') return null
  const abs = numbering.abstracts[numbering.nums[numId]]
  if (!abs) return null
  const lvl = abs[ilvl || '0'] || abs['0']
  if (!lvl || lvl.fmt === 'none') return null
  return { type: lvl.fmt === 'bullet' ? 'bullet' : 'ordered', level: parseInt(ilvl || '0', 10) || 0 }
}

function parseRels (xml) {
  const out = {}
  if (!xml) return out
  const root = parseXml(xml).documentElement
  for (const rel of Array.from(root.children)) {
    if (rel.localName !== 'Relationship') continue
    out[rel.getAttribute('Id')] = { target: rel.getAttribute('Target'), external: rel.getAttribute('TargetMode') === 'External' }
  }
  return out
}

// ---------------------------------------------------------------- images
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml' }
async function imageFromRel (ctx, rId) {
  const rel = rId && ctx.rels[rId]
  if (!rel || rel.external) return null
  let path = rel.target.replace(/^\//, '')
  if (!/^word\//.test(path)) path = 'word/' + path
  const ext = (path.split('.').pop() || '').toLowerCase()
  const mime = MIME[ext]
  if (!mime) return null
  const f = ctx.zip.file(path)
  if (!f) return null
  if (!ctx.imageCache[path]) {
    const b64 = await f.async('base64')
    if (b64.length > 6 * 1024 * 1024) return null // skip very large images
    ctx.imageCache[path] = `data:${mime};base64,${b64}`
  }
  return ctx.imageCache[path]
}
async function findImage (el, ctx) {
  const blips = el.getElementsByTagNameNS(A, 'blip')
  for (const b of Array.from(blips)) {
    const rId = b.getAttributeNS(R, 'embed') || b.getAttributeNS(R, 'link')
    const url = await imageFromRel(ctx, rId)
    if (url) return url
  }
  const datas = el.getElementsByTagNameNS(V, 'imagedata')
  for (const d of Array.from(datas)) {
    const url = await imageFromRel(ctx, d.getAttributeNS(R, 'id'))
    if (url) return url
  }
  return null
}

// ---------------------------------------------------------------- runs
async function readRun (r, ctx, extra, runs) {
  const rPr = readRPr(child(r, 'rPr'))
  let text = ''
  const flush = () => { if (text) { runs.push({ text, rPr, extra }); text = '' } }
  for (const c of Array.from(r.children)) {
    switch (c.localName) {
      case 't': text += c.textContent; break
      case 'tab': text += '\t'; break
      case 'br': case 'cr': text += '\n'; break
      case 'noBreakHyphen': text += '-'; break
      case 'softHyphen': break
      case 'drawing': case 'pict': case 'object': {
        const url = await findImage(c, ctx)
        if (url) { flush(); runs.push({ image: url }) }
        break
      }
      default: break
    }
  }
  flush()
}
async function paragraphRuns (p, ctx, inherited = {}) {
  const runs = []
  const walk = async (el, extra) => {
    for (const c of Array.from(el.children)) {
      switch (c.localName) {
        case 'r': await readRun(c, ctx, extra, runs); break
        case 'hyperlink': {
          const rel = ctx.rels[c.getAttributeNS(R, 'id')]
          await walk(c, rel && rel.external ? { ...extra, link: rel.target } : extra)
          break
        }
        case 'sdt': await walk(child(c, 'sdtContent') || c, extra); break
        case 'ins': case 'smartTag': case 'fldSimple': case 'customXml': case 'dir': case 'bdo': case 'moveTo':
          await walk(c, extra); break
        default: break
      }
    }
  }
  await walk(p, inherited)
  return runs
}

function chainRPr (ctx, styleId) {
  return merge(...styleChain(ctx.styles, styleId).map((s) => s.rPr))
}
function chainPPr (ctx, styleId) {
  return merge(...styleChain(ctx.styles, styleId).map((s) => s.pPr))
}
function runAttrs (run, ctx, pPr) {
  const pStyle = pPr.styleId || ctx.styles.defaultParagraphStyle
  const eff = merge(ctx.styles.defaults.rPr, chainRPr(ctx, pStyle), chainRPr(ctx, run.rPr.rStyle), run.rPr)
  const thai = THAI.test(run.text)
  const pick = (latin, cs) => (thai ? (cs !== undefined ? cs : latin) : (latin !== undefined ? latin : cs))
  const a = {}
  if (pick(eff.b, eff.bCs)) a.bold = true
  if (pick(eff.i, eff.iCs)) a.italic = true
  if (eff.u) a.underline = true
  if (eff.strike) a.strike = true
  const half = pick(eff.sz, eff.szCs)
  if (half) {
    const size = trim(half / 2) + 'pt'
    if (size !== ctx.defaultSize) a.size = size
  }
  const fontId = ctx.mapFont(pick(eff.font, eff.fontCs))
  if (fontId) a.font = fontId
  if (eff.color && eff.color !== 'auto' && /^[0-9a-f]{6}$/i.test(eff.color)) a.color = '#' + eff.color.toLowerCase()
  if (eff.shd && /^[0-9a-f]{6}$/i.test(eff.shd)) a.background = '#' + eff.shd.toLowerCase()
  else if (eff.highlight && HIGHLIGHT[eff.highlight]) a.background = HIGHLIGHT[eff.highlight]
  if (eff.vertAlign === 'superscript') a.script = 'super'
  else if (eff.vertAlign === 'subscript') a.script = 'sub'
  if (run.extra && run.extra.link) a.link = run.extra.link
  return a
}

function headingLevel (ctx, pPr, eff) {
  const chain = styleChain(ctx.styles, pPr.styleId)
  for (let k = chain.length - 1; k >= 0; k--) {
    const s = chain[k]
    const m = /^heading\s*(\d)$/i.exec(s.id) || /^heading\s*(\d)$/i.exec(s.name) || /^หัวเรื่อง\s*(\d)$/.exec(s.name)
    if (m) return Math.min(+m[1], 6)
    if (/^title$/i.test(s.id) || /^title$/i.test(s.name)) return 1
  }
  if (eff.outlineLvl != null && eff.outlineLvl >= 0 && eff.outlineLvl < 6) return eff.outlineLvl + 1
  return null
}

function blockAttrs (pPr, ctx) {
  const pStyle = pPr.styleId || ctx.styles.defaultParagraphStyle
  const eff = merge(ctx.styles.defaults.pPr, chainPPr(ctx, pStyle), pPr)
  const a = {}
  const h = headingLevel(ctx, pPr, eff)
  if (h) a.header = h
  const li = listInfo(ctx.numbering, eff.numId, eff.ilvl)
  if (li) { a.list = li.type; if (li.level > 0) a.indent = Math.min(li.level, 8) }
  const jc = eff.jc
  if (jc === 'center') a.align = 'center'
  else if (jc === 'right' || jc === 'end') a.align = 'right'
  else if (jc === 'both' || jc === 'justify' || jc === 'distribute' || jc === 'thaiDistribute') a.align = 'justify'
  if (!li) {
    const left = eff.indLeft || 0
    let first = eff.firstLine || 0
    if (eff.hanging) first = -eff.hanging
    if (left) a.leftindent = cm(left)
    if (first) a.firstline = cm(first)
    if (eff.indRight) a.rightindent = cm(eff.indRight)
  }
  if (eff.before) a.spacebefore = pt(eff.before)
  if (eff.after) a.spaceafter = pt(eff.after)
  if (eff.line) {
    const rule = eff.lineRule || 'auto'
    if (rule === 'auto') {
      const m = Math.round((eff.line / 240) * 100) / 100
      if (Math.abs(m - 1) > 0.01) a.linespacing = String(m)
    } else {
      a.linespacing = pt(eff.line)
    }
  }
  return a
}

// ---------------------------------------------------------------- paragraphs & tables
function pushRuns (delta, runs, inlineCtx, pPr, ctx, block, { singleLine = false } = {}) {
  for (const run of runs) {
    if (run.image) { delta.insert({ image: run.image }); continue }
    const attrs = runAttrs(run, ctx, pPr)
    const text = singleLine ? run.text.replace(/\n/g, ' ') : run.text
    const parts = text.split('\n')
    parts.forEach((part, i) => {
      if (i > 0) delta.insert('\n', block)
      if (part) delta.insert(part, attrs)
    })
  }
}
async function pushParagraph (delta, p, ctx) {
  const pPr = readPPr(child(p, 'pPr'))
  const runs = await paragraphRuns(p, ctx)
  const block = blockAttrs(pPr, ctx)
  pushRuns(delta, runs, null, pPr, ctx, block)
  delta.insert('\n', block)
}
// ---------------------------------------------------------------- tables (quill-table-up format)
export const TABLE_CELL = 'table-up-cell-inner'
export const TABLE_COL = 'table-up-col'
const rid = () => Math.random().toString(36).slice(2, 10)

function cellsOf (tr) {
  const out = []
  for (const c of Array.from(tr.children)) {
    if (c.localName === 'tc') out.push(c)
    else if (c.localName === 'sdt' || c.localName === 'customXml') out.push(...cellsOf(child(c, 'sdtContent') || c))
  }
  return out
}
function rowsOf (tbl) {
  const out = []
  for (const c of Array.from(tbl.children)) {
    if (c.localName === 'tr') out.push(c)
    else if (c.localName === 'sdt' || c.localName === 'customXml') out.push(...rowsOf(child(c, 'sdtContent') || c))
  }
  return out
}
// paragraphs inside a cell, in order; nested tables are flattened row by row
function cellParagraphs (el) {
  const out = []
  for (const c of Array.from(el.children)) {
    if (c.localName === 'p') out.push(c)
    else if (c.localName === 'tbl') for (const tr of rowsOf(c)) for (const tc of cellsOf(tr)) out.push(...cellParagraphs(tc))
    else if (c.localName === 'sdt' || c.localName === 'customXml') out.push(...cellParagraphs(child(c, 'sdtContent') || c))
  }
  return out
}
function noBorders (tblPr, ctx) {
  const b = child(tblPr, 'tblBorders')
  const sides = b ? Array.from(b.children) : []
  if (sides.length && sides.every((e) => ['nil', 'none'].includes(attr(e, 'val')))) return true
  // a table without borders and without a bordered style (e.g. "Normal Table") has no lines in Word
  if (!b) {
    const styleId = attr(child(tblPr, 'tblStyle'), 'val')
    const st = styleId && ctx.tableStyles[styleId]
    return st ? !st.hasBorders : true
  }
  return false
}
async function pushTable (delta, tbl, ctx) {
  const tableId = rid()
  const tblPr = child(tbl, 'tblPr')
  const grid = kids(child(tbl, 'tblGrid'), 'gridCol').map((g) => num(attr(g, 'w')) || 0)
  const rows = rowsOf(tbl).map((tr) => {
    const trPr = child(tr, 'trPr')
    let col = num(attr(child(trPr, 'gridBefore'), 'val')) || 0
    const cells = []
    for (const tc of cellsOf(tr)) {
      const tcPr = child(tc, 'tcPr')
      const span = Math.max(1, num(attr(child(tcPr, 'gridSpan'), 'val')) || 1)
      const vm = child(tcPr, 'vMerge')
      const hm = child(tcPr, 'hMerge')
      cells.push({
        tc, tcPr, col, span,
        vMerge: vm ? (attr(vm, 'val') || 'continue') : null,
        hMergeContinue: hm ? (attr(hm, 'val') || 'continue') === 'continue' : false,
        width: num(attr(child(tcPr, 'tcW'), 'w'))
      })
      col += span
    }
    return {
      header: onOff(trPr, 'tblHeader') === true,
      height: num(attr(child(trPr, 'trHeight'), 'val')),
      cells
    }
  }).filter((r) => r.cells.length)
  if (!rows.length) return
  // legacy horizontal merges: fold "continue" cells into the cell on their left
  for (const r of rows) {
    for (let k = r.cells.length - 1; k > 0; k--) {
      if (r.cells[k].hMergeContinue) { r.cells[k - 1].span += r.cells[k].span; r.cells.splice(k, 1) }
    }
  }
  const ncols = Math.max(grid.length, ...rows.map((r) => r.cells.reduce((m, c) => Math.max(m, c.col + c.span), 0)), 1)
  let widths = Array.from({ length: ncols }, (_, i) => grid[i] || 0)
  if (widths.some((w) => !w)) {
    // no usable grid: use cell widths of the first row, else equal widths
    const fromCells = Array(ncols).fill(0)
    for (const c of rows[0].cells) if (c.width && c.span === 1) fromCells[c.col] = c.width
    widths = widths.map((w, i) => w || fromCells[i] || Math.round(ctx.textWidth / ncols))
  }
  const total = widths.reduce((s, w) => s + w, 0)
  const full = total >= 0.9 * ctx.textWidth
  const colIds = widths.map(() => rid())
  const rowIds = rows.map(() => rid())
  widths.forEach((w, i) => {
    delta.insert({
      [TABLE_COL]: {
        tableId, colId: colIds[i], full,
        width: full ? Math.round((w / total) * 10000) / 100 : Math.max(40, Math.round(w / 15))
      }
    })
  })
  // vertical merges -> rowspan
  rows.forEach((r, ri) => {
    for (const c of r.cells) {
      if (c.vMerge === 'continue') { c.covered = true; continue }
      let rs = 1
      if (c.vMerge === 'restart') {
        for (let k = ri + 1; k < rows.length; k++) {
          const below = rows[k].cells.find((x) => x.col === c.col)
          if (below && below.vMerge === 'continue') rs++
          else break
        }
      }
      c.rowspan = rs
    }
  })
  // rows whose cells are all covered by row-spans must be recorded on the spanning cell
  rows.forEach((r, ri) => {
    if (r.cells.some((c) => !c.covered)) return
    for (let k = ri - 1; k >= 0; k--) {
      const owner = rows[k].cells.find((c) => !c.covered && k + c.rowspan > ri)
      if (owner) { (owner.emptyRow = owner.emptyRow || []).push(rowIds[ri]); break }
    }
  })
  const bordersOff = noBorders(tblPr, ctx)
  // repeating header rows are the ones at the top of the table
  let headerRows = 0
  while (headerRows < rows.length && rows[headerRows].header) headerRows++
  if (headerRows === rows.length) headerRows = 0
  for (let ri = 0; ri < rows.length; ri++) {
    const r = rows[ri]
    for (const c of r.cells) {
      if (c.covered) continue
      const style = []
      const shd = child(c.tcPr, 'shd')
      const fill = shd && attr(shd, 'fill')
      if (fill && fill !== 'auto' && /^[0-9a-f]{6}$/i.test(fill)) style.push(`background-color: #${fill.toLowerCase()}`)
      if (r.height) style.push(`height: ${Math.round(r.height / 15)}px`)
      if (bordersOff) style.push('border: 1px dashed rgb(203, 213, 225)')
      const cellAttr = {
        tableId,
        rowId: rowIds[ri],
        colId: colIds[Math.min(c.col, ncols - 1)],
        rowspan: c.rowspan || 1,
        colspan: c.span,
        tag: 'td',
        wrapTag: ri < headerRows ? 'thead' : 'tbody'
      }
      if (style.length) cellAttr.style = style.join('; ')
      if (c.emptyRow) cellAttr.emptyRow = c.emptyRow
      const paras = cellParagraphs(c.tc)
      let wrote = false
      for (const p of paras) {
        const pPr = readPPr(child(p, 'pPr'))
        const runs = await paragraphRuns(p, ctx)
        const { list, indent, header, ...ba } = blockAttrs(pPr, ctx)
        const block = { ...ba, [TABLE_CELL]: cellAttr }
        if (list) block.list = list
        if (indent) block.indent = indent
        pushRuns(delta, runs, null, pPr, ctx, block)
        delta.insert('\n', block)
        wrote = true
      }
      if (!wrote) delta.insert('\n', { [TABLE_CELL]: cellAttr })
    }
  }
}
function parseTableStyles (xml) {
  const out = {}
  if (!xml) return out
  const root = parseXml(xml).documentElement
  for (const s of kids(root, 'style')) {
    if (attr(s, 'type') !== 'table') continue
    const b = child(child(s, 'tblPr'), 'tblBorders')
    const sides = b ? Array.from(b.children) : []
    out[attr(s, 'styleId')] = { hasBorders: sides.some((e) => !['nil', 'none'].includes(attr(e, 'val'))) }
  }
  return out
}
function pageOf (body) {
  const sect = Array.from(body.children).reverse().find((c) => c.localName === 'sectPr')
  const sz = child(sect, 'pgSz')
  const mar = child(sect, 'pgMar')
  const page = {
    width: num(attr(sz, 'w')),
    height: num(attr(sz, 'h')),
    top: num(attr(mar, 'top')),
    right: num(attr(mar, 'right')) ?? num(attr(mar, 'end')),
    bottom: num(attr(mar, 'bottom')),
    left: num(attr(mar, 'left')) ?? num(attr(mar, 'start'))
  }
  for (const k of Object.keys(page)) if (page[k] == null) delete page[k]
  if (page.top != null) page.top = Math.abs(page.top)
  if (page.bottom != null) page.bottom = Math.abs(page.bottom)
  return page
}
function textWidthOf (body) {
  const sect = Array.from(body.children).reverse().find((c) => c.localName === 'sectPr')
  const w = num(attr(child(sect, 'pgSz'), 'w'))
  const mar = child(sect, 'pgMar')
  const l = num(attr(mar, 'left')) ?? num(attr(mar, 'start'))
  const r = num(attr(mar, 'right')) ?? num(attr(mar, 'end'))
  if (w && l != null && r != null && w - l - r > 1000) return w - l - r
  return 9026
}

/**
 * Convert a .docx ArrayBuffer to a Quill Delta.
 * opts.mapFont(name) -> quill font id or null; opts.defaultSize e.g. '16pt'
 */
export async function docxToDelta (arrayBuffer, opts = {}) {
  const zip = await JSZip.loadAsync(arrayBuffer)
  const read = async (p) => { const f = zip.file(p); return f ? f.async('string') : null }
  const docXml = await read('word/document.xml')
  if (!docXml) throw new Error('not a Word .docx file')
  const ctx = {
    zip,
    styles: parseStyles(await read('word/styles.xml')),
    numbering: parseNumbering(await read('word/numbering.xml')),
    rels: parseRels(await read('word/_rels/document.xml.rels')),
    imageCache: {},
    mapFont: opts.mapFont || (() => null),
    defaultSize: opts.defaultSize || '16pt',
    tableStyles: {},
    textWidth: 9026
  }
  ctx.tableStyles = parseTableStyles(await read('word/styles.xml'))
  const body = parseXml(docXml).documentElement.getElementsByTagNameNS(W, 'body')[0]
  if (!body) throw new Error('document body missing')
  ctx.textWidth = textWidthOf(body)
  const delta = new Delta()
  const walkBody = async (el) => {
    for (const c of Array.from(el.children)) {
      if (c.localName === 'p') await pushParagraph(delta, c, ctx)
      else if (c.localName === 'tbl') await pushTable(delta, c, ctx)
      else if (c.localName === 'sdt') await walkBody(child(c, 'sdtContent') || c)
    }
  }
  await walkBody(body)
  delta.page = pageOf(body)
  // a table at the very end needs a paragraph after it so people can type below
  const last = delta.ops[delta.ops.length - 1]
  if (last && last.attributes && last.attributes[TABLE_CELL]) delta.insert('\n')
  return delta
}
