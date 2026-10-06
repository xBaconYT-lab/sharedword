// Font catalogue. `docx` is the exact family name written into Word files.
// bundled = served from ./fonts, google = loaded from Google Fonts (links in
// index.html), system = only renders if installed (Windows/Office fonts), but
// the exported Word file still uses the right name.
const slug = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '')
const SANS = ', "TH Sarabun PSK", sans-serif'
const SERIF = ', "TH Sarabun PSK", serif'

const raw = [
  // --- Thai national / official fonts (bundled)
  { name: 'TH Sarabun PSK', docx: 'TH SarabunPSK', css: '"TH Sarabun PSK", "Sarabun", sans-serif', source: 'bundled', isDefault: true },
  { name: 'TH Sarabun New', docx: 'TH Sarabun New', css: '"TH Sarabun New"' + SANS, source: 'system' },
  // --- Thai fonts from Google Fonts (free)
  ...['Sarabun', 'Kanit', 'Prompt', 'Mitr', 'Pridi', 'Taviraj', 'Trirong', 'Athiti', 'Maitree', 'Thasadith',
    'Chonburi', 'Sriracha', 'Itim', 'Pattaya', 'Noto Sans Thai', 'Noto Serif Thai', 'Noto Sans Thai Looped',
    'IBM Plex Sans Thai', 'Anuphan', 'Bai Jamjuree', 'Chakra Petch', 'Charm', 'Charmonman', 'Fahkwang', 'K2D',
    'Kodchasan', 'KoHo', 'Krub', 'Mali', 'Niramit', 'Srisakdi'
  ].map((name) => ({ name, docx: name, css: `"${name}"` + SANS, source: 'google' })),
  // --- Thai fonts that ship with Windows / Microsoft Office (not bundled)
  ...[['Angsana New', SERIF], ['AngsanaUPC', SERIF], ['Browallia New', SANS], ['BrowalliaUPC', SANS], ['Cordia New', SANS],
    ['CordiaUPC', SANS], ['DilleniaUPC', SERIF], ['EucrosiaUPC', SERIF], ['FreesiaUPC', SANS], ['IrisUPC', SANS],
    ['JasmineUPC', SERIF], ['KodchiangUPC', SERIF], ['LilyUPC', SANS], ['Leelawadee', SANS], ['Leelawadee UI', SANS],
    ['Tahoma', SANS], ['Microsoft Sans Serif', SANS]
  ].map(([name, fb]) => ({ name, docx: name, css: `"${name}"` + fb, source: 'system' })),
  // --- Common Latin fonts
  ...[['Arial', ', Helvetica, sans-serif'], ['Arial Black', ', Impact, sans-serif'], ['Calibri', ', Carlito, sans-serif'],
    ['Cambria', ', Caladea, serif'], ['Candara', ', sans-serif'], ['Century Gothic', ', sans-serif'],
    ['Comic Sans MS', ', cursive'], ['Consolas', ', Menlo, monospace'], ['Courier New', ', Courier, monospace'],
    ['Garamond', ', serif'], ['Georgia', ', serif'], ['Helvetica', ', Arial, sans-serif'], ['Impact', ', sans-serif'],
    ['Lucida Console', ', monospace'], ['Palatino Linotype', ', Palatino, serif'], ['Segoe UI', ', sans-serif'],
    ['Times New Roman', ', Times, serif'], ['Trebuchet MS', ', sans-serif'], ['Verdana', ', sans-serif']
  ].map(([name, fb]) => ({ name, docx: name, css: `"${name}"` + fb, source: 'system' }))
]

export const FONTS = raw.map((f) => ({ ...f, id: slug(f.name) }))
export const DEFAULT_FONT = FONTS.find((f) => f.isDefault)
export const DEFAULT_DOCX_FONT = DEFAULT_FONT.docx

const byId = Object.fromEntries(FONTS.map((f) => [f.id, f]))
const norm = (s) => String(s || '').toLowerCase().replace(/[^a-z0-9]/g, '')
const byName = {}
for (const f of FONTS) { byName[norm(f.name)] = f.id; byName[norm(f.docx)] = f.id }
// extra spellings seen in Word files
Object.assign(byName, { thsarabun: DEFAULT_FONT.id, thsarabunpsk: DEFAULT_FONT.id, sarabunpsk: DEFAULT_FONT.id })

/** Word font name -> quill font id, or null for the default font / unknown fonts. */
export function mapFontName (name) {
  const id = byName[norm(name)]
  if (!id || id === DEFAULT_FONT.id) return null
  return id
}
/** quill font id -> Word font name */
export function docxFontName (id) {
  return (id && byId[id] ? byId[id].docx : DEFAULT_DOCX_FONT)
}
/** CSS for .ql-font-<id> classes and font previews in the toolbar picker. */
export function injectFontCss () {
  let css = ''
  for (const f of FONTS) {
    css += `.ql-font-${f.id}{font-family:${f.css};}\n`
    css += `.ql-snow .ql-picker.ql-font .ql-picker-item[data-value="${f.id}"]::before{font-family:${f.css};}\n`
  }
  const style = document.createElement('style')
  style.textContent = css
  document.head.appendChild(style)
}
/** Fill the toolbar <select class="ql-font"> (blank value = default font). */
export function fillFontSelect (select) {
  select.innerHTML = ''
  const def = document.createElement('option')
  def.textContent = DEFAULT_FONT.name
  def.setAttribute('selected', '')
  select.appendChild(def)
  for (const f of FONTS) {
    if (f.isDefault) continue
    const o = document.createElement('option')
    o.value = f.id
    o.textContent = f.name
    select.appendChild(o)
  }
}
