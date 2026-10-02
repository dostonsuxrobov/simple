'use strict'

// Simple's own conversions for Combine, used when no local office engine is
// available (and always for CSV). Word documents (.docx) go through mammoth;
// spreadsheets (.xlsx, .ods) through SheetJS, and CSV through a small reader
// that keeps every value as written. Each file becomes one or more print jobs
// (complete HTML documents plus page options); the launcher main process
// prints them with the shared, hardened html-to-pdf, because printing needs a
// browser window and this module runs in the Combine worker thread.
//
// Like combine-service.cjs, this requires workspace node_modules read-only;
// the Combine worker bundle (scripts/sync-build.cjs) includes them.

const path = require('node:path')
const { decodeText } = require('../shared/electron/text-codec.cjs')
const { CombineError } = require('./combine-policy.cjs')

const MAX_CELLS = 200_000
const MAX_COLUMNS = 1_000
const PX_PER_INCH = 96
const DEFAULT_COLUMN_PX = 64
const A4 = Object.freeze({ width: 8.27, height: 11.69 })
/** Excel paperSize codes Simple prints, in inches (portrait). */
const EXCEL_PAPER = Object.freeze({
  1: { width: 8.5, height: 11 }, 2: { width: 8.5, height: 11 }, 3: { width: 11, height: 17 }, 4: { width: 17, height: 11 },
  5: { width: 8.5, height: 14 }, 6: { width: 5.5, height: 8.5 }, 7: { width: 7.25, height: 10.5 }, 8: { width: 11.69, height: 16.54 },
  9: { width: 8.27, height: 11.69 }, 10: { width: 8.27, height: 11.69 }, 11: { width: 5.83, height: 8.27 }, 12: { width: 10.12, height: 14.33 },
  13: { width: 7.17, height: 10.12 }, 14: { width: 8.5, height: 13 }, 66: { width: 16.54, height: 23.39 }, 70: { width: 4.13, height: 5.83 },
})
const EXCEL_MARGINS = Object.freeze({ left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 })
const DOCX_STYLE_MAP = Object.freeze(['u => u', "br[type='page'] => br.page-break"])
const SAFE_LINK = /^(?:https?:|mailto:|#)/i
const COMBINE_PRINT = Object.freeze({ generateTaggedPDF: false, generateDocumentOutline: false, printBackground: true })

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function decodeXml(value) {
  return String(value ?? '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_match, hex) => String.fromCodePoint(Number.parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_match, digits) => String.fromCodePoint(Number(digits)))
    .replace(/&amp;/g, '&')
}

/** Value of an XML attribute in one tag's text, decoded; undefined when absent. */
function attribute(tag, name) {
  const match = new RegExp(`\\s${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*=\\s*("([^"]*)"|'([^']*)')`).exec(tag)
  if (!match) return undefined
  return decodeXml(match[2] !== undefined ? match[2] : match[3])
}

function clamp(value, minimum, maximum) {
  return Math.min(maximum, Math.max(minimum, value))
}

function finite(value, fallback) {
  const number = Number(value)
  return Number.isFinite(number) ? number : fallback
}

function stem(name) {
  const base = path.basename(String(name || 'Document'))
  return base.replace(/\.[^.]+$/, '') || base
}

function damaged(kind, technical) {
  return new CombineError('DAMAGED', `Simple couldn't read this ${kind}. It may be damaged, protected or not really a ${kind}.`, { technical })
}

function tooLarge(kind) {
  return new CombineError('TOO_LARGE', `This ${kind} has more cells than Simple can combine at once (${MAX_CELLS.toLocaleString('en-US')}). Combine a smaller copy, or set a print area and save it first.`)
}

/** Font family names are only ever copied into CSS when they are plain words. */
function cssFontName(value) {
  const name = String(value || '').trim()
  return /^[A-Za-z0-9][A-Za-z0-9 ._-]{0,62}$/.test(name) ? `"${name}"` : null
}

async function zipText(zip, name) {
  const file = zip.file(name)
  return file ? file.async('string') : null
}

/**
 * Elements with the exact tag name (not longer names sharing its prefix),
 * self-closing or with content. These elements never nest in the files read here.
 * @returns {Array<{tag: string, content: string}>}
 */
function elements(xml, name) {
  const found = []
  const opening = new RegExp(`<${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(?=[\\s/>])[^>]*>`, 'g')
  const closing = `</${name}>`
  let match
  while ((match = opening.exec(xml))) {
    if (match[0].endsWith('/>')) { found.push({ tag: match[0], content: '' }); continue }
    const end = xml.indexOf(closing, opening.lastIndex)
    if (end < 0) { found.push({ tag: match[0], content: '' }); continue }
    found.push({ tag: match[0], content: xml.slice(opening.lastIndex, end) })
    opening.lastIndex = end + closing.length
  }
  return found
}

/**
 * What the bytes really are, whatever the extension says.
 * @param {Uint8Array} bytes
 * @returns {'zip'|'cfb'|'pdf'|'other'}
 */
function sniffContainer(bytes) {
  if (!bytes || bytes.length < 8) return 'other'
  if (bytes[0] === 0x50 && bytes[1] === 0x4b && (bytes[2] === 0x03 || bytes[2] === 0x05 || bytes[2] === 0x07)) return 'zip'
  if (Buffer.from(bytes.buffer, bytes.byteOffset, 8).toString('hex') === 'd0cf11e0a1b11ae1') return 'cfb'
  if (Buffer.from(bytes.buffer, bytes.byteOffset, Math.min(bytes.length, 1024)).includes('%PDF-')) return 'pdf'
  return 'other'
}

// ---------------------------------------------------------------------------
// Word (.docx)
// ---------------------------------------------------------------------------

/** First section's page size and margins, the default font and paragraph spacing. */
async function readDocxLayout(zip) {
  const layout = { width: A4.width, height: A4.height, margins: { top: 1, right: 1, bottom: 1, left: 1 }, font: null, size: 11, after: 8, line: 1.08 }
  const documentXml = await zipText(zip, 'word/document.xml') || ''
  const section = /<w:sectPr\b[^>]*>([\s\S]*?)<\/w:sectPr>/.exec(documentXml)
  if (section) {
    const size = /<w:pgSz\b[^>]*>/.exec(section[1])
    if (size) {
      const width = finite(attribute(size[0], 'w:w'), 0) / 1440
      const height = finite(attribute(size[0], 'w:h'), 0) / 1440
      if (width >= 1 && height >= 1 && width <= 60 && height <= 60) {
        const landscape = attribute(size[0], 'w:orient') === 'landscape'
        layout.width = landscape ? Math.max(width, height) : width
        layout.height = landscape ? Math.min(width, height) : height
      }
    }
    const margins = /<w:pgMar\b[^>]*>/.exec(section[1])
    if (margins) {
      for (const side of ['top', 'right', 'bottom', 'left']) {
        const value = Math.abs(finite(attribute(margins[0], `w:${side}`), 1440)) / 1440
        const limit = (side === 'top' || side === 'bottom' ? layout.height : layout.width) / 3
        layout.margins[side] = clamp(value, 0, limit)
      }
    }
  }
  const stylesXml = await zipText(zip, 'word/styles.xml') || ''
  const defaults = /<w:docDefaults>([\s\S]*?)<\/w:docDefaults>/.exec(stylesXml)?.[1] || ''
  const normal = /<w:style\b[^>]*w:styleId="Normal"[^>]*>([\s\S]*?)<\/w:style>/.exec(stylesXml)?.[1] || ''
  let fonts = null
  for (const source of [normal, defaults]) {
    const size = /<w:sz\b[^>]*w:val="(\d+)"/.exec(source)
    if (size && layout.size === 11) layout.size = clamp(Number(size[1]) / 2, 6, 36)
    const tag = /<w:rFonts\b[^>]*>/.exec(source)?.[0]
    if (tag && !fonts) fonts = tag
  }
  const spacing = /<w:spacing\b[^>]*>/.exec(/<w:pPrDefault>([\s\S]*?)<\/w:pPrDefault>/.exec(defaults)?.[1] || '')?.[0]
  if (spacing) {
    const after = attribute(spacing, 'w:after')
    if (after !== undefined) layout.after = clamp(finite(after, 160) / 20, 0, 48)
    const line = attribute(spacing, 'w:line')
    if (line !== undefined && (attribute(spacing, 'w:lineRule') || 'auto') === 'auto') layout.line = clamp(finite(line, 259) / 240, 0.8, 3)
  }
  if (fonts) {
    let name = attribute(fonts, 'w:ascii') || attribute(fonts, 'w:hAnsi')
    const theme = attribute(fonts, 'w:asciiTheme') || attribute(fonts, 'w:hAnsiTheme')
    if (!name && theme) {
      const themeXml = await zipText(zip, 'word/theme/theme1.xml') || ''
      const group = /^major/i.test(theme) ? 'majorFont' : 'minorFont'
      name = attribute(new RegExp(`<a:${group}>[\\s\\S]*?(<a:latin\\b[^>]*>)`).exec(themeXml)?.[1] || '', 'typeface')
    }
    layout.font = cssFontName(name)
  }
  return layout
}

/** Neutralizes link targets other than web, mail and in-document anchors. */
function sanitizeLinks(html) {
  return html.replace(/(<a\b[^>]*?\shref=")([^"]*)(")/gi, (match, before, value, after) => (SAFE_LINK.test(decodeXml(value).trim()) ? match : `${before}#${after}`))
}

const INLINE_WRAPPER = /<(strong|em|u|s|sup|sub|span|a)(?:\s[^>]*)?>\s*<br class="page-break" \/>\s*<\/\1>/g
const BREAK = '<br class="page-break" />'
const VOID_TAG = /^(?:br|img|hr|wbr|input|col|meta|link)$/i

/** Splits block paragraphs at Word page breaks so each break starts a new printed page. */
function applyPageBreaks(html) {
  let text = html
  // Unwrap breaks that sit alone inside formatting, e.g. <strong><br class="page-break" /></strong>.
  for (let pass = 0; pass < 4; pass += 1) {
    const unwrapped = text.replace(INLINE_WRAPPER, BREAK)
    if (unwrapped === text) break
    text = unwrapped
  }
  text = text.replace(/<(p|h[1-6])(\s[^>]*)?>([\s\S]*?)<\/\1>/g, (match, tag, attributes = '', inner) => {
    if (!inner.includes(BREAK)) return match
    const parts = []
    let depth = 0
    let start = 0
    const scanner = /<br class="page-break" \/>|<\/?([a-zA-Z0-9]+)(?:\s[^>]*)?\/?>/g
    let token
    while ((token = scanner.exec(inner))) {
      if (token[0] === BREAK) {
        if (depth === 0) {
          parts.push(inner.slice(start, token.index))
          start = token.index + BREAK.length
        }
        continue
      }
      if (VOID_TAG.test(token[1]) || token[0].endsWith('/>')) continue
      depth += token[0].startsWith('</') ? -1 : 1
    }
    parts.push(inner.slice(start))
    const blocks = parts.map((part) => part.split(BREAK).join(''))
    return blocks
      .map((part, index) => (part.trim() || (index > 0 && index < blocks.length - 1) ? `<${tag}${attributes}>${part}</${tag}>` : ''))
      .join('<div class="page-break"></div>')
  })
  return text.split(BREAK).join('')
}

function docxHtml(body, layout, title) {
  const font = [layout.font, 'Calibri', 'Carlito', '"Segoe UI"', 'Arial', 'sans-serif'].filter(Boolean).join(', ')
  const size = layout.size
  const inches = (value) => `${Number(value.toFixed(3))}in`
  return `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>
@page { size: ${inches(layout.width)} ${inches(layout.height)}; margin: ${inches(layout.margins.top)} ${inches(layout.margins.right)} ${inches(layout.margins.bottom)} ${inches(layout.margins.left)}; }
html { background: #fff; color: #000; }
body { margin: 0; font-family: ${font}; font-size: ${size}pt; line-height: ${Number(layout.line.toFixed(3))}; overflow-wrap: break-word; }
p, li, h1, h2, h3, h4, h5, h6, td, th { unicode-bidi: plaintext; text-align: start; }
p { margin: 0 0 ${layout.after}pt; }
p:empty::before { content: "\\00a0"; }
h1, h2, h3, h4, h5, h6 { margin: ${Math.round(size)}pt 0 ${Math.round(size / 3)}pt; line-height: 1.2; break-after: avoid; }
h1 { font-size: ${Number((size * 1.45).toFixed(1))}pt; } h2 { font-size: ${Number((size * 1.2).toFixed(1))}pt; } h3 { font-size: ${Number((size * 1.1).toFixed(1))}pt; }
h4, h5, h6 { font-size: ${size}pt; }
ul, ol { margin: 0 0 ${layout.after}pt; padding-left: 0.35in; }
li { margin: 0; }
table { border-collapse: collapse; margin: 0 0 ${layout.after}pt; width: 100%; }
td, th { border: 0.5pt solid #a6a6a6; padding: 1pt 5pt; vertical-align: top; text-align: start; }
td > :last-child, th > :last-child { margin-bottom: 0; }
img { max-width: 100%; height: auto; }
a { color: #0563c1; }
sup, sub { line-height: 0; }
.page-break { break-after: page; height: 0; margin: 0; }
</style></head><body>${body}</body></html>`
}

/**
 * Print jobs for a Word document (.docx), laid out by Simple: text, headings,
 * lists, tables, images, links and footnotes; the first section's page size
 * and margins; page breaks.
 * @param {Uint8Array} bytes
 * @param {{name?: string}} [context]
 * @returns {Promise<{jobs: Array<{html: string, options: object}>}>}
 * @throws {CombineError} DAMAGED when the file is not a readable Word document
 */
async function docxPrintJobs(bytes, context = {}) {
  const JSZip = require('../../simple_pdf_source/node_modules/jszip')
  const mammoth = require('../../simple_pdf_source/node_modules/mammoth')
  const buffer = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes)
  let zip
  try { zip = await JSZip.loadAsync(buffer) } catch (error) { throw damaged('Word document', error.message) }
  if (!zip.file('word/document.xml')) throw damaged('Word document', 'word/document.xml is missing.')
  const layout = await readDocxLayout(zip)
  let result
  try {
    result = await mammoth.convertToHtml({ buffer }, { ignoreEmptyParagraphs: false, styleMap: [...DOCX_STYLE_MAP] })
  } catch (error) { throw damaged('Word document', error.message) }
  const title = stem(context.name)
  const body = applyPageBreaks(sanitizeLinks(String(result.value || '')))
  return { jobs: [{ html: docxHtml(body, layout, title), options: { ...COMBINE_PRINT, preferCSSPageSize: true, title } }] }
}

// ---------------------------------------------------------------------------
// Spreadsheets
// ---------------------------------------------------------------------------

/** Ranges of hidden indices, e.g. [{start: 4, count: 2}]. */
function hiddenIn(ranges, index) {
  for (const range of ranges || []) if (index >= range.start && index < range.start + range.count) return true
  return false
}

function lengthInInches(value) {
  const match = /^\s*(-?[0-9.]+)\s*(in|cm|mm|pt|pc|px)?\s*$/i.exec(String(value || ''))
  if (!match) return null
  const number = Number(match[1])
  if (!Number.isFinite(number)) return null
  const factor = { in: 1, cm: 1 / 2.54, mm: 1 / 25.4, pt: 1 / 72, pc: 1 / 6, px: 1 / 96 }[(match[2] || 'in').toLowerCase()]
  return number * factor
}

/** Page setup, header and footer of every sheet of an .xlsx, in workbook order. */
async function readXlsxSetups(zip) {
  const workbookXml = await zipText(zip, 'xl/workbook.xml') || ''
  const relsXml = await zipText(zip, 'xl/_rels/workbook.xml.rels') || ''
  const targets = new Map()
  for (const tag of relsXml.match(/<Relationship\b[^>]*>/g) || []) targets.set(attribute(tag, 'Id'), attribute(tag, 'Target'))
  const setups = []
  for (const tag of workbookXml.match(/<sheet\b[^>]*>/g) || []) {
    const relationship = /\s[\w]+:id\s*=\s*"([^"]*)"/.exec(tag)?.[1]
    let target = relationship ? targets.get(relationship) : undefined
    const setup = { hidden: /^(hidden|veryHidden)$/.test(attribute(tag, 'state') || ''), paper: null, landscape: false, scale: null, fitToWidth: false, header: null, footer: null }
    if (target) {
      target = target.startsWith('/') ? target.slice(1) : path.posix.normalize(`xl/${target}`)
      const sheetXml = await zipText(zip, target) || ''
      const pageSetup = /<pageSetup\b[^>]*>/.exec(sheetXml)?.[0] || ''
      const fitToPage = /<pageSetUpPr\b[^>]*\sfitToPage\s*=\s*"(1|true)"/.test(sheetXml)
      setup.paper = EXCEL_PAPER[Number(attribute(pageSetup, 'paperSize') || 0)] || null
      setup.landscape = attribute(pageSetup, 'orientation') === 'landscape'
      if (fitToPage) setup.fitToWidth = Number(attribute(pageSetup, 'fitToWidth') ?? 1) >= 1 || Number(attribute(pageSetup, 'fitToHeight') ?? 1) >= 1
      else if (attribute(pageSetup, 'scale')) setup.scale = clamp(finite(attribute(pageSetup, 'scale'), 100), 10, 400) / 100
      const headerFooter = /<headerFooter\b[^>]*>([\s\S]*?)<\/headerFooter>/.exec(sheetXml)?.[1] || ''
      const header = /<oddHeader>([\s\S]*?)<\/oddHeader>/.exec(headerFooter)?.[1]
      const footer = /<oddFooter>([\s\S]*?)<\/oddFooter>/.exec(headerFooter)?.[1]
      if (header) setup.header = decodeXml(header)
      if (footer) setup.footer = decodeXml(footer)
    }
    setups.push(setup)
  }
  return setups
}

/** Visibility, hidden rows and columns, and page layout of every table of an .ods, in order. */
async function readOdsSetups(zip) {
  const contentXml = await zipText(zip, 'content.xml') || ''
  const stylesXml = await zipText(zip, 'styles.xml') || ''
  const tableStyles = new Map()
  for (const { tag, content } of elements(contentXml, 'style:style')) {
    if (attribute(tag, 'style:family') !== 'table') continue
    const properties = elements(content, 'style:table-properties')[0]?.tag || ''
    tableStyles.set(attribute(tag, 'style:name'), { master: attribute(tag, 'style:master-page-name'), hidden: attribute(properties, 'table:display') === 'false' })
  }
  const layouts = new Map()
  for (const { tag, content } of elements(stylesXml, 'style:page-layout')) {
    layouts.set(attribute(tag, 'style:name'), elements(content, 'style:page-layout-properties')[0]?.tag || '')
  }
  const masters = new Map()
  for (const { tag } of elements(stylesXml, 'style:master-page')) masters.set(attribute(tag, 'style:name'), attribute(tag, 'style:page-layout-name'))
  const setups = []
  for (const { tag, content } of elements(contentXml, 'table:table')) {
    const style = tableStyles.get(attribute(tag, 'table:style-name')) || {}
    const properties = layouts.get(masters.get(style.master || 'Default')) || ''
    const width = lengthInInches(attribute(properties, 'fo:page-width'))
    const height = lengthInInches(attribute(properties, 'fo:page-height'))
    const setup = { hidden: Boolean(style.hidden), paper: null, landscape: attribute(properties, 'style:print-orientation') === 'landscape', scale: null, fitToWidth: false, margins: null, hiddenRows: [], hiddenColumns: [] }
    if (width && height && width >= 1 && height >= 1 && width <= 60 && height <= 60) setup.paper = { width: Math.min(width, height), height: Math.max(width, height) }
    const margin = (side) => lengthInInches(attribute(properties, `fo:margin-${side}`))
    if (['top', 'right', 'bottom', 'left'].some((side) => margin(side) !== null)) {
      setup.margins = { top: margin('top') ?? 0.79, right: margin('right') ?? 0.79, bottom: margin('bottom') ?? 0.79, left: margin('left') ?? 0.79, header: 0, footer: 0 }
    }
    const scaleTo = /^([0-9.]+)%$/.exec(attribute(properties, 'style:scale-to') || '')
    if (scaleTo) setup.scale = clamp(Number(scaleTo[1]), 10, 400) / 100
    if (attribute(properties, 'style:scale-to-X') || attribute(properties, 'style:scale-to-pages')) setup.fitToWidth = true
    let column = 0
    for (const columnTag of content.match(/<table:table-column(?=[\s/>])[^>]*>/g) || []) {
      const count = clamp(Math.floor(finite(attribute(columnTag, 'table:number-columns-repeated'), 1)), 1, 1_048_576)
      if (/^(collapse|filter)$/.test(attribute(columnTag, 'table:visibility') || '')) setup.hiddenColumns.push({ start: column, count })
      column += count
    }
    let row = 0
    for (const rowTag of content.match(/<table:table-row(?=[\s/>])[^>]*>/g) || []) {
      const count = clamp(Math.floor(finite(attribute(rowTag, 'table:number-rows-repeated'), 1)), 1, 1_048_576)
      if (/^(collapse|filter)$/.test(attribute(rowTag, 'table:visibility') || '')) setup.hiddenRows.push({ start: row, count })
      row += count
    }
    setups.push(setup)
  }
  return setups
}

/**
 * Cell ranges of a defined name such as "'Sheet 1'!$A$1:$F$20,'Sheet 1'!$H:$H".
 * Whole columns or rows are bounded by `used`.
 */
function parseAreas(XLSX, reference, used) {
  const areas = []
  const parts = decodeXml(reference).match(/(?:'[^']*(?:''[^']*)*'|[^,'])+/g) || []
  for (const part of parts) {
    const local = part.slice(part.lastIndexOf('!') + 1).replace(/\$/g, '').trim().toUpperCase()
    let range = null
    if (/^[A-Z]+\d+(:[A-Z]+\d+)?$/.test(local)) range = XLSX.utils.decode_range(local.includes(':') ? local : `${local}:${local}`)
    else if (/^[A-Z]+:[A-Z]+$/.test(local)) {
      const [first, last] = local.split(':').map((column) => XLSX.utils.decode_col(column))
      range = { s: { r: used.s.r, c: first }, e: { r: used.e.r, c: last } }
    } else if (/^\d+:\d+$/.test(local)) {
      const [first, last] = local.split(':').map((row) => Number(row) - 1)
      range = { s: { r: first, c: used.s.c }, e: { r: last, c: used.e.c } }
    }
    if (range && range.s.r >= 0 && range.s.c >= 0 && range.e.r >= range.s.r && range.e.c >= range.s.c) areas.push(range)
  }
  return areas
}

/** Title rows ("$3:$3") of a Print_Titles name, as {start, end} zero-based, or null. */
function parseTitleRows(reference) {
  for (const part of decodeXml(reference).split(',')) {
    const match = /^\s*\$?(\d+):\$?(\d+)\s*$/.exec(part.slice(part.lastIndexOf('!') + 1))
    if (match) {
      const start = Number(match[1]) - 1
      const end = Number(match[2]) - 1
      if (start >= 0 && end >= start) return { start, end }
    }
  }
  return null
}

function cellText(cell) {
  if (!cell || cell.t === 'z') return ''
  if (cell.w !== undefined && cell.w !== null) return String(cell.w)
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE'
  if (cell.t === 'e') return '#ERROR!'
  if (cell.v instanceof Date) return cell.v.toISOString().slice(0, 10)
  return cell.v === undefined || cell.v === null ? '' : String(cell.v)
}

/** Excel's default alignment: numbers and dates right, logical values and errors centered, text at the start. */
function cellAlign(cell) {
  if (!cell) return ''
  if (cell.t === 'n' || cell.t === 'd') return 'n'
  if (cell.t === 'b' || cell.t === 'e') return 'c'
  return ''
}

function cellFill(cell) {
  const style = cell && cell.s
  if (!style || style.patternType !== 'solid') return null
  const rgb = String(style.fgColor?.rgb || '')
  if (!/^(?:[0-9a-f]{2})?[0-9a-f]{6}$/i.test(rgb)) return null
  const color = rgb.slice(-6).toLowerCase()
  return color === 'ffffff' ? null : `#${color}`
}

/**
 * Text color for a fill. Font colors are not available here, and dark fills
 * almost always carry light text, so dark fills get white text.
 * @param {string|null} fill "#rrggbb"
 * @returns {string|null}
 */
function textOnFill(fill) {
  if (!fill) return null
  const channel = (offset) => {
    const value = Number.parseInt(fill.slice(offset, offset + 2), 16) / 255
    return value <= 0.03928 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4
  }
  const luminance = 0.2126 * channel(1) + 0.7152 * channel(3) + 0.0722 * channel(5)
  return luminance < 0.18 ? '#ffffff' : null
}

function columnPx(column) {
  if (!column) return null
  if (Number.isFinite(column.width) && column.width > 0) return Math.round(column.width * 7)
  if (Number.isFinite(column.wpx) && column.wpx > 0) return Math.round(column.wpx)
  if (Number.isFinite(column.wch) && column.wch > 0) return Math.round(column.wch * 7 + 5)
  return null
}

/** Widths from content for files that store none (CSV, ODS). */
function contentWidths(rows, columnCount) {
  const widths = new Array(columnCount).fill(0)
  for (const row of rows.slice(0, 2000)) {
    for (let index = 0; index < columnCount; index += 1) {
      const text = row[index]?.text || ''
      if (!text || row[index].colspan > 1) continue
      const longest = text.split('\n').reduce((maximum, line) => Math.max(maximum, line.length), 0)
      widths[index] = Math.max(widths[index], longest)
    }
  }
  return widths.map((length) => clamp(Math.round(length * 7.2 + 12), 48, 420))
}

/**
 * Builds the printable rows of one area. Hidden rows and columns are left out
 * and merged cells become row and column spans.
 * @returns {Array<Array<object|null>>}
 */
function sectionRows(rowIndices, columnIndices, readCell, merges) {
  const covered = new Set()
  const spans = new Map()
  for (const merge of merges) {
    const rows = rowIndices.filter((row) => row >= merge.s.r && row <= merge.e.r)
    const columns = columnIndices.filter((column) => column >= merge.s.c && column <= merge.e.c)
    if (!rows.length || !columns.length || (rows.length === 1 && columns.length === 1)) continue
    spans.set(`${rows[0]},${columns[0]}`, { rowspan: rows.length, colspan: columns.length, source: merge.s })
    for (const row of rows) for (const column of columns) if (row !== rows[0] || column !== columns[0]) covered.add(`${row},${column}`)
  }
  return rowIndices.map((row) => columnIndices.map((column) => {
    const key = `${row},${column}`
    if (covered.has(key)) return null
    const span = spans.get(key)
    const cell = span ? readCell(span.source.r, span.source.c) : readCell(row, column)
    const text = cellText(cell)
    return { text, align: cellAlign(cell), fill: cellFill(cell), rowspan: span?.rowspan || 1, colspan: span?.colspan || 1 }
  }))
}

function rowsHtml(rows) {
  return rows.map((cells) => `<tr>${cells.map((cell) => {
    if (!cell) return ''
    const classes = [cell.align, cell.text.includes('\n') ? 'w' : ''].filter(Boolean).join(' ')
    const attributes = [
      classes ? ` class="${classes}"` : '',
      cell.colspan > 1 ? ` colspan="${cell.colspan}"` : '',
      cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : '',
      cell.fill ? ` style="background:${cell.fill}${textOnFill(cell.fill) ? `;color:${textOnFill(cell.fill)}` : ''}"` : '',
    ].join('')
    return `<td${attributes}>${escapeHtml(cell.text)}</td>`
  }).join('')}</tr>`).join('')
}

/** One table per print area: an optional lead block, then rows with repeated title rows. */
function tableHtml(area, widths) {
  const columns = `<colgroup>${widths.map((width) => `<col style="width:${width}px">`).join('')}</colgroup>`
  const width = widths.reduce((sum, value) => sum + value, 0)
  const parts = []
  if (area.lead.length) parts.push(`<table style="width:${width}px">${columns}<tbody>${rowsHtml(area.lead)}</tbody></table>`)
  if (area.titles.length || area.body.length) {
    parts.push(`<table style="width:${width}px">${columns}${area.titles.length ? `<thead>${rowsHtml(area.titles)}</thead>` : ''}<tbody>${rowsHtml(area.body)}</tbody></table>`)
  }
  return parts.join('')
}

const SHEET_CSS = `html { background: #fff; }
body { margin: 0; color: #000; font-family: Calibri, Carlito, "Segoe UI", Arial, sans-serif; font-size: 11pt; }
table { border-collapse: collapse; table-layout: fixed; }
td { border: 0.75px solid #d4d4d4; padding: 1px 4px; height: 18px; vertical-align: bottom; overflow-wrap: anywhere; line-height: 1.25; text-align: start; unicode-bidi: plaintext; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
td.n { text-align: right; } td.c { text-align: center; } td.w { white-space: pre-wrap; }
thead { display: table-header-group; }
tr { break-inside: avoid; }
.area + .area { break-before: page; }`

/**
 * Parses Excel header/footer codes (&L &C &R sections, &P page, &N pages,
 * &A sheet, &F file, &Z folder, &D date, &T time; formatting codes are
 * dropped) into the three sections as HTML.
 * @param {string|null} code
 * @param {{sheetName?: string, fileName?: string, folder?: string, now?: Date}} context
 * @returns {{left: string, center: string, right: string}|null}
 */
function excelHeaderSections(code, context = {}) {
  if (!code) return null
  const text = String(code)
  const sections = { left: '', center: '', right: '' }
  const now = context.now || new Date()
  let current = 'center'
  let index = 0
  while (index < text.length) {
    const character = text[index]
    if (character !== '&') { sections[current] += escapeHtml(character); index += 1; continue }
    const next = text[index + 1]
    index += 2
    if (next === undefined) break
    if (next === '&') sections[current] += '&amp;'
    else if (next === 'L') current = 'left'
    else if (next === 'C') current = 'center'
    else if (next === 'R') current = 'right'
    else if (next === 'P') {
      sections[current] += '<span class="pageNumber"></span>'
      const offset = /^[+-]\d+/.exec(text.slice(index))
      if (offset) index += offset[0].length
    } else if (next === 'N') sections[current] += '<span class="totalPages"></span>'
    else if (next === 'A') sections[current] += escapeHtml(context.sheetName || '')
    else if (next === 'F') sections[current] += escapeHtml(context.fileName || '')
    else if (next === 'Z') sections[current] += escapeHtml(context.folder ? `${context.folder}${path.sep}` : '')
    else if (next === 'D') sections[current] += escapeHtml(now.toLocaleDateString())
    else if (next === 'T') sections[current] += escapeHtml(now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }))
    else if (next === '"') {
      const end = text.indexOf('"', index)
      index = end < 0 ? text.length : end + 1
    } else if (next === 'K') {
      const color = /^(?:[0-9A-Fa-f]{6}|\d\d[+-]\d{3})/.exec(text.slice(index))
      if (color) index += color[0].length
    } else if (/\d/.test(next)) {
      while (/\d/.test(text[index] || '')) index += 1
    }
  }
  return sections.left || sections.center || sections.right ? sections : null
}

function headerTemplate(sections, page, position) {
  if (!sections) return null
  const offset = Number(clamp(position === 'header' ? page.margins.header : page.margins.footer, 0, 2).toFixed(3))
  const edge = position === 'header' ? `padding-top:${offset}in` : `padding-bottom:${offset}in`
  return `<div style="box-sizing:border-box;width:100%;${edge};padding-left:${page.margins.left}in;padding-right:${page.margins.right}in;font-family:Calibri,Carlito,'Segoe UI',Arial,sans-serif;font-size:8pt;color:#000;display:flex;align-items:${position === 'header' ? 'flex-start' : 'flex-end'};gap:8px">`
    + `<span style="flex:1;text-align:left">${sections.left}</span><span style="flex:1;text-align:center">${sections.center}</span><span style="flex:1;text-align:right">${sections.right}</span></div>`
}

/**
 * Turns one sheet's printable areas into a print job: page size and
 * orientation, margins, header and footer, and a scale that keeps every
 * column on the page (Chromium cannot continue a table on a page to the right).
 */
function sheetJob(sheet, context) {
  const paper = sheet.page.paper || A4
  const orientedWidth = sheet.page.landscape ? Math.max(paper.width, paper.height) : Math.min(paper.width, paper.height)
  const orientedHeight = sheet.page.landscape ? Math.min(paper.width, paper.height) : Math.max(paper.width, paper.height)
  // Margins from a damaged or unusual file must still leave most of the page for content.
  const margins = { ...sheet.page.margins }
  for (const side of ['left', 'right']) margins[side] = clamp(finite(margins[side], 0), 0, orientedWidth * 0.3)
  for (const side of ['top', 'bottom', 'header', 'footer']) margins[side] = clamp(finite(margins[side], 0), 0, orientedHeight * 0.3)
  const page = { ...sheet.page, margins }
  const tableWidth = Math.max(...sheet.areas.map((area) => area.widths.reduce((sum, value) => sum + value, 0)), 1)
  const printable = Math.max(1, (orientedWidth - page.margins.left - page.margins.right) * PX_PER_INCH)
  const fit = printable / tableWidth
  let scale = page.fitToWidth ? Math.min(1, fit) : (page.scale || 1)
  if (tableWidth * scale > printable) scale = fit
  scale = Number(clamp(scale, 0.1, 2).toFixed(3))
  const title = sheet.name && sheet.name !== stem(context.name) ? `${stem(context.name)} - ${sheet.name}` : stem(context.name)
  const html = `<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>${SHEET_CSS}</style></head><body>${sheet.areas.map((area) => `<section class="area">${tableHtml(area, area.widths)}</section>`).join('')}</body></html>`
  const headerContext = { sheetName: sheet.name, fileName: context.name ? path.basename(context.name) : '', folder: context.path ? path.dirname(context.path) : '', now: context.now }
  const header = headerTemplate(excelHeaderSections(sheet.header, headerContext), page, 'header')
  const footer = headerTemplate(excelHeaderSections(sheet.footer, headerContext), page, 'footer')
  const options = {
    ...COMBINE_PRINT,
    title,
    preferCSSPageSize: false,
    pageSize: { width: Number(Math.min(paper.width, paper.height).toFixed(3)), height: Number(Math.max(paper.width, paper.height).toFixed(3)) },
    landscape: page.landscape,
    margins: {
      top: Number(page.margins.top.toFixed(3)), right: Number(page.margins.right.toFixed(3)),
      bottom: Number(page.margins.bottom.toFixed(3)), left: Number(page.margins.left.toFixed(3)),
    },
    scale,
  }
  if (header) options.headerTemplate = header
  if (footer) options.footerTemplate = footer
  return { html, options }
}

function sheetMargins(worksheet, setup) {
  if (setup?.margins) return setup.margins
  const margins = { ...EXCEL_MARGINS, ...(worksheet && worksheet['!margins'] ? worksheet['!margins'] : {}) }
  for (const side of Object.keys(EXCEL_MARGINS)) margins[side] = clamp(finite(margins[side], EXCEL_MARGINS[side]), 0, 10)
  return margins
}

/** Every printable area of every visible sheet in a SheetJS workbook. */
function workbookSheets(XLSX, workbook, setups, kind) {
  const names = workbook.Workbook?.Names || []
  const sheets = []
  let cells = 0
  workbook.SheetNames.forEach((sheetName, index) => {
    const worksheet = workbook.Sheets[sheetName]
    const setup = setups[index] || {}
    const hiddenSheet = setup.hidden || Number(workbook.Workbook?.Sheets?.[index]?.Hidden || 0) > 0
    if (!worksheet || hiddenSheet || !worksheet['!ref']) return
    const used = XLSX.utils.decode_range(worksheet['!ref'])
    const sheetNames = names.filter((name) => Number(name.Sheet) === index && typeof name.Ref === 'string')
    let areas = sheetNames.filter((name) => name.Name === '_xlnm.Print_Area').flatMap((name) => parseAreas(XLSX, name.Ref, used))
    const seen = new Set()
    areas = areas.filter((area) => {
      const key = XLSX.utils.encode_range(area)
      if (seen.has(key)) return false
      seen.add(key)
      return true
    })
    if (!areas.length) areas = [used]
    const titleName = sheetNames.find((name) => name.Name === '_xlnm.Print_Titles')
    const titleRows = titleName ? parseTitleRows(titleName.Ref) : null
    const dense = Array.isArray(worksheet['!data'])
    const readCell = (row, column) => (dense ? worksheet['!data'][row]?.[column] : worksheet[XLSX.utils.encode_cell({ r: row, c: column })])
    const rowInfo = worksheet['!rows'] || []
    const columnInfo = worksheet['!cols'] || []
    const rowHidden = (row) => Boolean(rowInfo[row]?.hidden) || hiddenIn(setup.hiddenRows, row)
    const columnHidden = (column) => Boolean(columnInfo[column]?.hidden) || hiddenIn(setup.hiddenColumns, column)
    const merges = worksheet['!merges'] || []
    const printableAreas = []
    for (const candidate of areas) {
      const area = {
        s: { r: Math.max(candidate.s.r, used.s.r), c: Math.max(candidate.s.c, used.s.c) },
        e: { r: Math.min(candidate.e.r, used.e.r), c: Math.min(candidate.e.c, used.e.c) },
      }
      if (area.e.r < area.s.r || area.e.c < area.s.c) continue
      const columnIndices = []
      for (let column = area.s.c; column <= area.e.c && columnIndices.length <= MAX_COLUMNS; column += 1) if (!columnHidden(column)) columnIndices.push(column)
      if (columnIndices.length > MAX_COLUMNS) throw tooLarge('spreadsheet')
      const visibleRows = []
      for (let row = area.s.r; row <= area.e.r; row += 1) {
        if (rowHidden(row)) continue
        visibleRows.push(row)
        if (visibleRows.length * columnIndices.length + cells > MAX_CELLS) throw tooLarge('spreadsheet')
      }
      if (!columnIndices.length || !visibleRows.length) continue
      let lead = []
      let titles = []
      let body = visibleRows
      if (titleRows && titleRows.start <= area.e.r) {
        const titleIndices = []
        for (let row = titleRows.start; row <= titleRows.end; row += 1) if (!rowHidden(row) && row <= used.e.r) titleIndices.push(row)
        if (titleIndices.length) {
          lead = visibleRows.filter((row) => row < titleRows.start)
          body = visibleRows.filter((row) => row > titleRows.end)
          titles = titleIndices
        }
      }
      cells += (lead.length + titles.length + body.length) * columnIndices.length
      if (cells > MAX_CELLS) throw tooLarge('spreadsheet')
      const section = (rows) => (rows.length ? sectionRows(rows, columnIndices, readCell, merges) : [])
      const printable = { lead: section(lead), titles: section(titles), body: section(body) }
      const hasText = [...printable.lead, ...printable.titles, ...printable.body].some((row) => row.some((cell) => cell && (cell.text.trim() || cell.fill)))
      if (!hasText) continue
      const stored = columnIndices.map((column) => columnPx(columnInfo[column]))
      const estimated = stored.some((width) => width === null) ? contentWidths([...printable.lead, ...printable.titles, ...printable.body], columnIndices.length) : []
      printable.widths = stored.map((width, position) => width ?? (kind === 'xlsx' ? DEFAULT_COLUMN_PX : estimated[position]))
      printableAreas.push(printable)
    }
    if (!printableAreas.length) return
    sheets.push({
      // SheetJS returns OpenDocument sheet names still XML-escaped.
      name: kind === 'ods' ? decodeXml(sheetName) : sheetName,
      areas: printableAreas,
      header: setup.header || null,
      footer: setup.footer || null,
      page: { paper: setup.paper, landscape: Boolean(setup.landscape), scale: setup.scale || null, fitToWidth: Boolean(setup.fitToWidth), margins: sheetMargins(worksheet, setup) },
    })
  })
  return sheets
}

/**
 * Splits CSV text into rows of fields. Honors an Excel "sep=" first line,
 * otherwise picks the delimiter (comma, semicolon, tab or bar) that splits
 * the first lines most consistently. Values stay exactly as written.
 * @param {string} text
 * @returns {{rows: string[][], delimiter: string}}
 */
function parseCsv(text) {
  let source = String(text || '')
  let delimiter = null
  const hint = /^sep=(.)\r?\n/i.exec(source)
  if (hint) {
    delimiter = hint[1]
    source = source.slice(hint[0].length)
  }
  if (!delimiter) {
    const lines = source.split(/\r\n|\n|\r/).filter((line) => line.trim()).slice(0, 25)
    let best = { delimiter: ',', score: 0 }
    for (const candidate of [',', ';', '\t', '|']) {
      const counts = lines.map((line) => {
        let count = 0
        let quoted = false
        for (const character of line) {
          if (character === '"') quoted = !quoted
          else if (!quoted && character === candidate) count += 1
        }
        return count
      })
      const frequency = new Map()
      for (const count of counts) if (count > 0) frequency.set(count, (frequency.get(count) || 0) + 1)
      let score = 0
      for (const [count, lineCount] of frequency) score = Math.max(score, lineCount * 1000 + count)
      if (score > best.score) best = { delimiter: candidate, score }
    }
    delimiter = best.delimiter
  }
  const rows = []
  let row = []
  let field = ''
  let quoted = false
  let cells = 0
  const pushField = () => {
    row.push(field)
    field = ''
    cells += 1
    if (cells > MAX_CELLS) throw tooLarge('CSV file')
  }
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index]
    if (quoted) {
      if (character === '"') {
        if (source[index + 1] === '"') { field += '"'; index += 1 } else quoted = false
      } else field += character
      continue
    }
    if (character === '"' && field === '') quoted = true
    else if (character === delimiter) pushField()
    else if (character === '\r' || character === '\n') {
      pushField()
      rows.push(row)
      row = []
      if (character === '\r' && source[index + 1] === '\n') index += 1
    } else field += character
  }
  if (field !== '' || row.length) {
    pushField()
    rows.push(row)
  }
  return { rows, delimiter }
}

function csvSheet(text, name) {
  const { rows } = parseCsv(text)
  const columnCount = rows.reduce((maximum, row) => Math.max(maximum, row.length), 0)
  if (columnCount > MAX_COLUMNS) throw tooLarge('CSV file')
  const printableRows = rows.map((row) => Array.from({ length: columnCount }, (_value, index) => {
    const value = row[index] ?? ''
    return { text: value, align: /^\s*[-+]?[$€£¥]?\s*\d[\d,.\s]*%?\s*$/.test(value) && value.trim().length < 20 ? 'n' : '', fill: null, rowspan: 1, colspan: 1 }
  }))
  while (printableRows.length && printableRows[printableRows.length - 1].every((cell) => !cell.text.trim())) printableRows.pop()
  if (!printableRows.length || !columnCount) return []
  const widths = contentWidths(printableRows, columnCount)
  const width = widths.reduce((sum, value) => sum + value, 0)
  const margins = { ...EXCEL_MARGINS }
  const landscape = width > (A4.width - margins.left - margins.right) * PX_PER_INCH * 1.05
  return [{ name: stem(name), areas: [{ lead: [], titles: [], body: printableRows, widths }], header: null, footer: null, page: { paper: A4, landscape, scale: null, fitToWidth: true, margins } }]
}

function quietly(operation) {
  // SheetJS reports harmless number-format notes on the console while reading some ODS files.
  const original = { error: console.error, warn: console.warn, log: console.log }
  console.error = () => {}
  console.warn = () => {}
  console.log = () => {}
  try { return operation() } finally { Object.assign(console, original) }
}

/**
 * Print jobs for a spreadsheet, laid out by Simple: one job per visible sheet,
 * honoring print areas, repeated title rows, page size, orientation, margins,
 * scaling, Excel headers and footers, merged cells and solid fills. Hidden
 * sheets, rows and columns are never printed.
 * @param {Uint8Array} bytes
 * @param {{kind: 'xlsx'|'ods'|'csv', name?: string, path?: string, now?: Date}} context
 * @returns {Promise<{jobs: Array<{html: string, options: object}>}>}
 * @throws {CombineError} DAMAGED, TOO_LARGE or NOTHING_TO_PRINT
 */
async function spreadsheetPrintJobs(bytes, context) {
  const kind = context.kind
  const buffer = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes)
  let sheets
  if (kind === 'csv') {
    sheets = csvSheet(decodeText(buffer).text, context.name)
  } else {
    const XLSX = require('../../simple_calc_source/node_modules/xlsx')
    const JSZip = require('../../simple_pdf_source/node_modules/jszip')
    const label = kind === 'ods' ? 'OpenDocument spreadsheet' : 'Excel workbook'
    let workbook
    let setups
    try {
      const zip = await JSZip.loadAsync(buffer)
      setups = kind === 'ods' ? await readOdsSetups(zip) : await readXlsxSetups(zip)
      workbook = quietly(() => XLSX.read(buffer, { type: 'buffer', cellStyles: true, cellNF: true, cellDates: false, dense: true }))
    } catch (error) {
      if (error instanceof CombineError) throw error
      throw damaged(label, error.message)
    }
    if (!workbook || !Array.isArray(workbook.SheetNames)) throw damaged(label, 'No sheets.')
    sheets = workbookSheets(XLSX, workbook, setups, kind)
  }
  if (!sheets.length) throw new CombineError('NOTHING_TO_PRINT', 'This spreadsheet has no visible cells to print.')
  return { jobs: sheets.map((sheet) => sheetJob(sheet, context)) }
}

module.exports = {
  MAX_CELLS,
  applyPageBreaks,
  docxPrintJobs,
  excelHeaderSections,
  parseCsv,
  readDocxLayout,
  sanitizeLinks,
  sniffContainer,
  spreadsheetPrintJobs,
}
