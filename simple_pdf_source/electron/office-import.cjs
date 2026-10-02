'use strict'

// Turns everything Simple PDF can open into PDF pages without an office
// suite. Word (.docx) goes through mammoth, spreadsheets through SheetJS, and
// text, Markdown, HTML, RTF and OpenDocument text through HTML that Chromium
// lays out and prints in a hidden, script-free window, so the main process
// never blocks on layout. PNG, JPEG and TIFF are placed by image-to-pdf.cjs;
// WebP, GIF, BMP, AVIF and ICO are decoded by Chromium first, and SVG is
// printed as vector graphics.
//
// A local LibreOffice, when present, is only an optional fidelity upgrade for
// office formats (SIMPLE_FORCE_NO_OFFICE=1 makes Simple ignore it). The
// formats that still need it (.doc, .ppt, .pptx, .odp) fail with a coded
// NEEDS_OFFICE_ENGINE error whose message says what to do instead.
//
// The pure helpers (sniffing, decoding, HTML building) load without Electron
// so they can be tested in Node; printing needs the Electron main process.

const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { fileURLToPath, pathToFileURL } = require('node:url')
const { codedError } = require('./pdf-problems.cjs')
const { imageToPdfBytes, sniffImageType, svgDimensions, tiffFirstImage } = require('./image-to-pdf.cjs')

const PARTITION = 'simple-pdf-convert'
const JOB_PREFIX = 'simple-pdf-convert-'
const MAX_DOCUMENT_BYTES = 256 * 1024 * 1024
const MAX_HTML_BYTES = 192 * 1024 * 1024
const MAX_CONCURRENT_JOBS = 2
const MAX_SHEET_CELLS = 400_000
const POINTS_PER_TWIP = 1 / 20
const POINTS_PER_EMU = 1 / 12700
const CONTENT_SECURITY_POLICY = "default-src 'none'; img-src data: file:; style-src 'unsafe-inline' data: file:; font-src data: file:; media-src data: file:; base-uri file:; form-action 'none'; frame-src 'none'; object-src 'none'"

/**
 * Every importable format. `engine` formats convert only through a local
 * office engine and are left out of pickers when there is none; dropping one
 * still explains what to do instead.
 */
const IMPORT_FORMATS = Object.freeze([
  { kind: 'pdf', label: 'PDF files', extensions: ['pdf'] },
  { kind: 'word', label: 'Word documents', extensions: ['docx', 'docm', 'dotx', 'dotm'] },
  { kind: 'word-legacy', label: 'Word 97–2003 documents', extensions: ['doc'], engine: true },
  { kind: 'odt', label: 'OpenDocument text', extensions: ['odt'] },
  { kind: 'rtf', label: 'Rich Text', extensions: ['rtf'] },
  { kind: 'spreadsheet', label: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xlsb', 'xls', 'ods', 'csv', 'tsv'] },
  { kind: 'markdown', label: 'Markdown', extensions: ['md', 'markdown', 'mdown', 'mkd'] },
  { kind: 'html', label: 'Web pages', extensions: ['html', 'htm', 'xhtml'] },
  { kind: 'text', label: 'Text files', extensions: ['txt', 'text', 'log'] },
  { kind: 'image', label: 'Images', extensions: ['png', 'jpg', 'jpeg', 'jpe', 'jfif', 'gif', 'webp', 'bmp', 'dib', 'tif', 'tiff', 'svg', 'avif', 'ico'] },
  { kind: 'presentation', label: 'Presentations', extensions: ['pptx', 'ppt', 'odp'], engine: true },
].map((entry) => Object.freeze({ ...entry, extensions: Object.freeze(entry.extensions) })))

const KIND_BY_EXTENSION = new Map(IMPORT_FORMATS.flatMap((entry) => entry.extensions.map((extension) => [extension, entry.kind])))
/** Formats a local office engine renders more faithfully than the built-in path. */
const ENGINE_INPUTS = new Set(['doc', 'docx', 'xls', 'xlsx', 'odt', 'ods', 'rtf', 'ppt', 'pptx'])
const IMAGE_MIME = {
  png: 'image/png', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp',
  svg: 'image/svg+xml', avif: 'image/avif', ico: 'image/x-icon', tiff: 'image/tiff', heic: 'image/heic',
}
const LETTER_REGIONS = new Set(['US', 'CA', 'MX', 'PH', 'CL', 'CO', 'VE', 'CR', 'GT', 'PA', 'PR', 'DO', 'SV', 'NI', 'HN', 'BZ'])

function toBuffer(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return Buffer.from(value || [])
}

function extensionOf(name) {
  return path.extname(String(name || '')).slice(1).toLowerCase()
}

function baseTitle(name) {
  return path.basename(String(name || 'Document'), path.extname(String(name || '')))
    .replace(/[\u0000-\u001f]/g, ' ').trim() || 'Document'
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
}

function cssString(value) {
  return `"${String(value || '').replace(/["\\\n\r]/g, ' ').trim()}"`
}

function round(value, digits = 2) {
  const factor = 10 ** digits
  return Math.round(Number(value) * factor) / factor
}

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

function namedMessage(name, message) {
  return name ? `${path.basename(String(name))}: ${message}` : message
}

/** Coded error for formats that only a local office engine can convert. */
function needsOfficeEngineError(kind, name) {
  const extension = extensionOf(name)
  const message = kind === 'presentation'
    ? "Presentations can't be turned into PDF pages on this PC yet. In PowerPoint, save the presentation as a PDF (File > Save As > PDF), then open that PDF here."
    : "Older Word documents (.doc) can't be turned into PDF pages directly. Open the file in Simple Documents or Word, save it as a .docx, then use the .docx here."
  const error = codedError('NEEDS_OFFICE_ENGINE', namedMessage(name, message), {
    kind,
    altExt: kind === 'presentation' ? '.pdf' : '.docx',
    extension,
  })
  return error
}

function unsupportedFormatError(name, detail = '') {
  const extension = extensionOf(name)
  const what = extension ? `.${extension} files` : 'this kind of file'
  return codedError('UNSUPPORTED_FORMAT', namedMessage(name, `Simple can't open ${what} as a PDF.${detail ? ` ${detail}` : ''} It can open PDFs, Word (.docx), spreadsheets, text, Markdown, web pages, RTF, OpenDocument text and images.`))
}

function conversionFailedError(name, error) {
  if (error && typeof error.code === 'string' && /^[A-Z_]+$/.test(error.code) && error.name === error.code) return error
  const result = codedError('CONVERSION_FAILED', namedMessage(name, "Simple couldn't turn this file into PDF pages. The file may be damaged or use features Simple can't read. Your file is unchanged."))
  result.technical = error?.message || String(error)
  return result
}

// ---------------------------------------------------------------------------
// Format detection
// ---------------------------------------------------------------------------

function utf16Bytes(text) {
  return Buffer.from(text, 'utf16le')
}

const OLE_STREAMS = [
  ['doc', utf16Bytes('WordDocument')],
  ['xls', utf16Bytes('Workbook')],
  ['xls', utf16Bytes('Book\u0000')],
  ['ppt', utf16Bytes('PowerPoint Document')],
]

function zipKind(bytes) {
  // ODF packages start with an uncompressed "mimetype" entry.
  if (bytes.length > 38 && bytes.toString('latin1', 30, 38) === 'mimetype') {
    const mime = bytes.toString('latin1', 38, 38 + 64)
    if (mime.startsWith('application/vnd.oasis.opendocument.text')) return 'odt'
    if (mime.startsWith('application/vnd.oasis.opendocument.spreadsheet')) return 'ods'
    if (mime.startsWith('application/vnd.oasis.opendocument.presentation')) return 'odp'
  }
  // Entry names are repeated in the central directory at the end.
  const sample = bytes.length <= 4 * 1024 * 1024
    ? bytes
    : Buffer.concat([bytes.subarray(0, 512 * 1024), bytes.subarray(bytes.length - 2 * 1024 * 1024)])
  if (sample.includes('word/document') || sample.includes('word/_rels')) return 'docx'
  if (sample.includes('xl/workbook') || sample.includes('xl/_rels')) return 'xlsx'
  if (sample.includes('ppt/presentation') || sample.includes('ppt/_rels')) return 'pptx'
  return 'zip'
}

function textStart(bytes) {
  let start = 0
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) start = 3
  return bytes.toString('latin1', start, Math.min(bytes.length, start + 4096))
    .replace(/^\s+/, '')
}

function looksLikeHtml(head) {
  const lower = head.slice(0, 1024).toLowerCase()
  return /^<!doctype\s+html/.test(lower) || /^(<!--[\s\S]*?-->\s*)*<(html|head|body)[\s>]/.test(lower)
}

/**
 * The real type of a file from its first bytes, or null when nothing is
 * recognised. Binary containers win over the extension, so a PNG named .jpg
 * or a PDF without an extension still opens.
 * @returns {string|null} pdf, png, jpeg, gif, webp, bmp, tiff, ico, avif, heic,
 *   svg, docx, xlsx, pptx, odt, ods, odp, zip, doc, xls, ppt, ole, rtf or html
 */
function sniffType(input) {
  const bytes = toBuffer(input)
  if (bytes.length < 4) return null
  const pdfAt = bytes.subarray(0, 1024).indexOf('%PDF-')
  if (pdfAt >= 0) return 'pdf'
  const image = sniffImageType(bytes)
  if (image) return image
  if (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) return zipKind(bytes)
  if (bytes.subarray(0, 8).toString('hex') === 'd0cf11e0a1b11ae1') {
    for (const [kind, needle] of OLE_STREAMS) if (bytes.includes(needle)) return kind
    return 'ole'
  }
  const head = textStart(bytes)
  if (head.startsWith('{\\rtf')) return 'rtf'
  if (looksLikeHtml(head)) return 'html'
  return null
}

function looksLikeText(bytes) {
  const sample = bytes.subarray(0, 8192)
  if (!sample.length) return true
  if ((sample[0] === 0xff && sample[1] === 0xfe) || (sample[0] === 0xfe && sample[1] === 0xff)) return true
  const utf16 = detectUtf16(sample)
  if (utf16) return sample.length >= 8 && controlShare(new TextDecoder(utf16).decode(sample.subarray(0, sample.length & ~1))) < 0.02
  let control = 0
  for (const byte of sample) {
    if (byte === 0) return false
    if (byte < 0x09 || (byte > 0x0d && byte < 0x20 && byte !== 0x1b)) control += 1
  }
  return control / sample.length < 0.02
}

const SNIFF_KINDS = {
  pdf: 'pdf', png: 'image', jpeg: 'image', gif: 'image', webp: 'image', bmp: 'image', tiff: 'image',
  ico: 'image', avif: 'image', svg: 'image', heic: 'heic', docx: 'word', xlsx: 'spreadsheet', ods: 'spreadsheet',
  odt: 'odt', pptx: 'presentation', odp: 'presentation', doc: 'word-legacy', xls: 'spreadsheet', ppt: 'presentation',
  rtf: 'rtf',
}

/**
 * What a file is and how it converts.
 * @param {Buffer|Uint8Array} bytes
 * @param {string} name file name (only its extension is used)
 * @returns {{kind: string|null, sniffed: string|null, extension: string}}
 */
function detectImportKind(bytes, name) {
  const buffer = toBuffer(bytes)
  const extension = extensionOf(name)
  const byExtension = KIND_BY_EXTENSION.get(extension) || null
  const sniffed = sniffType(buffer)
  if (sniffed && SNIFF_KINDS[sniffed]) {
    // An SVG or HTML file is text: keep a text-like extension's meaning, so a
    // .txt holding markup is shown as text, not rendered.
    if (sniffed === 'svg' && ['text', 'markdown', 'html'].includes(byExtension)) return { kind: byExtension, sniffed, extension }
    if (sniffed === 'rtf' && ['text', 'markdown'].includes(byExtension)) return { kind: byExtension, sniffed, extension }
    return { kind: SNIFF_KINDS[sniffed], sniffed, extension }
  }
  if (sniffed === 'zip' || sniffed === 'ole') {
    return { kind: byExtension && byExtension !== 'text' && byExtension !== 'markdown' ? byExtension : null, sniffed, extension }
  }
  // Word can save a web page with a .doc name; it opens as the page it is.
  if (sniffed === 'html' && byExtension === 'word-legacy') return { kind: 'html', sniffed, extension }
  if (byExtension) return { kind: byExtension, sniffed, extension }
  if (sniffed === 'html') return { kind: 'html', sniffed, extension }
  if (looksLikeText(buffer)) return { kind: 'text', sniffed: 'text', extension }
  return { kind: null, sniffed, extension }
}

/** True for a name Simple may try to open (by extension). */
function isImportableName(name) {
  return KIND_BY_EXTENSION.has(extensionOf(name))
}

function importExtensions({ engine = false } = {}) {
  return IMPORT_FORMATS.filter((entry) => engine || !entry.engine).flatMap((entry) => entry.extensions)
}

/**
 * File-picker filters. Formats that need the office engine are offered only
 * when this PC has one.
 * @param {{engine?: boolean, imagesOnly?: boolean}} [options]
 */
function importDialogFilters({ engine = false, imagesOnly = false } = {}) {
  const images = IMPORT_FORMATS.find((entry) => entry.kind === 'image').extensions
  if (imagesOnly) return [{ name: 'Images', extensions: [...images] }]
  const groups = [
    { name: 'PDF files', extensions: ['pdf'] },
    { name: 'Word documents', extensions: ['docx', 'docm', 'dotx', 'dotm', ...(engine ? ['doc'] : []), 'odt', 'rtf'] },
    { name: 'Spreadsheets', extensions: ['xlsx', 'xlsm', 'xltx', 'xltm', 'xlsb', 'xls', 'ods', 'csv', 'tsv'] },
    { name: 'Text and web pages', extensions: ['txt', 'text', 'log', 'md', 'markdown', 'html', 'htm', 'xhtml'] },
    { name: 'Images', extensions: [...images] },
    ...(engine ? [{ name: 'Presentations', extensions: ['pptx', 'ppt'] }] : []),
  ]
  return [{ name: 'All supported files', extensions: importExtensions({ engine }) }, ...groups]
}

// ---------------------------------------------------------------------------
// Text decoding
// ---------------------------------------------------------------------------

function detectUtf16(bytes) {
  const length = Math.min(bytes.length, 4096) & ~1
  if (length < 4) return null
  let evenZeros = 0
  let oddZeros = 0
  for (let index = 0; index < length; index += 2) {
    if (bytes[index] === 0) evenZeros += 1
    if (bytes[index + 1] === 0) oddZeros += 1
  }
  const pairs = length / 2
  if (oddZeros / pairs > 0.3 && evenZeros / pairs < 0.05) return 'utf-16le'
  if (evenZeros / pairs > 0.3 && oddZeros / pairs < 0.05) return 'utf-16be'
  return null
}

/**
 * BOM-less UTF-16 of a non-Latin script has few zero bytes, but nearly every
 * high byte is the same block (0x04 for Cyrillic, 0x05 Hebrew, 0x06 Arabic).
 * Used only after strict UTF-8 failed.
 */
function detectUtf16ByBlock(bytes) {
  const length = Math.min(bytes.length, 8192) & ~1
  if (length < 8 || bytes.length % 2) return null
  for (const [label, highOffset] of [['utf-16le', 1], ['utf-16be', 0]]) {
    const counts = new Map()
    for (let index = highOffset; index < length; index += 2) counts.set(bytes[index], (counts.get(bytes[index]) || 0) + 1)
    const [block, count] = [...counts].sort((left, right) => right[1] - left[1])[0]
    const share = count / (length / 2)
    const zeroShare = (counts.get(0) || 0) / (length / 2)
    if (block > 0 && block < 0x30 && share + zeroShare > 0.9 && share > 0.5) {
      const text = new TextDecoder(label).decode(bytes)
      if (!/[\u0000-\u0008\u000e-\u001f�]/.test(text.slice(0, 4096))) return label
    }
  }
  return null
}

/** Share of C0 control characters other than tab, line breaks and form feed. */
function controlShare(text) {
  const sample = text.slice(0, 8192)
  if (!sample.length) return 0
  let count = 0
  for (let index = 0; index < sample.length; index += 1) {
    const code = sample.charCodeAt(index)
    if (code < 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d && code !== 0x0c) count += 1
  }
  return count / sample.length
}

/** Length of a truncated UTF-8 sequence at the very end, if any. */
function incompleteUtf8Tail(bytes) {
  for (let back = 1; back <= Math.min(3, bytes.length); back += 1) {
    const byte = bytes[bytes.length - back]
    if ((byte & 0xc0) === 0x80) continue
    const needed = byte >= 0xf0 ? 4 : byte >= 0xe0 ? 3 : byte >= 0xc0 ? 2 : 1
    return needed > back ? back : 0
  }
  return 0
}

function strictUtf8(bytes) {
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch {
    const tail = incompleteUtf8Tail(bytes)
    if (!tail) return null
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(0, bytes.length - tail)) + '�'
    } catch { return null }
  }
}

/**
 * True when high bytes read as Cyrillic letters (windows-1251) rather than
 * the occasional accented Latin letter (windows-1252).
 */
function looksCyrillic(bytes) {
  let ascii = 0
  let high = 0
  const sample = bytes.subarray(0, 256 * 1024)
  for (const byte of sample) {
    if ((byte >= 0x41 && byte <= 0x5a) || (byte >= 0x61 && byte <= 0x7a)) ascii += 1
    else if (byte >= 0xc0 || byte === 0xa8 || byte === 0xb8) high += 1
  }
  return high >= 3 && high > ascii * 0.6
}

const CHARSET_ALIASES = { 'x-user-defined': 'windows-1252', 'iso-8859-1': 'windows-1252', latin1: 'windows-1252', ascii: 'windows-1252', 'us-ascii': 'windows-1252' }

function supportedLabel(label) {
  const value = String(label || '').trim().toLowerCase()
  if (!value) return null
  const alias = CHARSET_ALIASES[value] || value
  try {
    return new TextDecoder(alias).encoding
  } catch { return null }
}

/**
 * Decodes a text file the way Notepad would: a byte order mark wins, then
 * BOM-less UTF-16 (by its zero bytes), then strict UTF-8, then the declared
 * charset (HTML), then windows-1251 for Cyrillic-looking bytes, else
 * windows-1252.
 * @param {Buffer|Uint8Array} input
 * @param {{declared?: string}} [options]
 * @returns {{text: string, encoding: string, bom: boolean}}
 */
function decodeText(input, options = {}) {
  const bytes = toBuffer(input)
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return { text: new TextDecoder('utf-8').decode(bytes.subarray(3)), encoding: 'utf-8', bom: true }
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe && !(bytes[2] === 0 && bytes[3] === 0)) {
    return { text: new TextDecoder('utf-16le').decode(bytes.subarray(2)), encoding: 'utf-16le', bom: true }
  }
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return { text: new TextDecoder('utf-16be').decode(bytes.subarray(2)), encoding: 'utf-16be', bom: true }
  }
  const utf16 = detectUtf16(bytes)
  if (utf16) return { text: new TextDecoder(utf16).decode(bytes), encoding: utf16, bom: false }
  const utf8 = strictUtf8(bytes)
  // Russian UTF-16 without a BOM is all bytes below 0x80, so it is valid
  // UTF-8 too, but made of control characters.
  if (utf8 !== null && controlShare(utf8) < 0.02) return { text: utf8, encoding: 'utf-8', bom: false }
  const blockUtf16 = detectUtf16ByBlock(bytes)
  if (blockUtf16) return { text: new TextDecoder(blockUtf16).decode(bytes), encoding: blockUtf16, bom: false }
  if (utf8 !== null) return { text: utf8, encoding: 'utf-8', bom: false }
  const declared = supportedLabel(options.declared)
  if (declared && declared !== 'utf-8' && !declared.startsWith('utf-16')) {
    return { text: new TextDecoder(declared).decode(bytes), encoding: declared, bom: false }
  }
  const label = looksCyrillic(bytes) ? 'windows-1251' : 'windows-1252'
  return { text: new TextDecoder(label).decode(bytes), encoding: label, bom: false }
}

/** The charset an HTML file declares in its first bytes, if any. */
function declaredHtmlCharset(bytes) {
  const head = toBuffer(bytes).toString('latin1', 0, Math.min(bytes.length, 8192))
  const match = /<meta[^>]+charset\s*=\s*["']?\s*([\w:.-]+)/i.exec(head)
  return match ? match[1] : null
}

// ---------------------------------------------------------------------------
// Page set-up and shared print CSS
// ---------------------------------------------------------------------------

function regionCode() {
  try {
    const electron = require('electron')
    const code = electron?.app?.getLocaleCountryCode?.()
    if (code) return String(code).toUpperCase()
  } catch { /* plain Node */ }
  try {
    const locale = new Intl.DateTimeFormat().resolvedOptions().locale
    const region = new Intl.Locale(locale).maximize().region
    if (region) return region.toUpperCase()
  } catch { /* unknown */ }
  return ''
}

/** Letter in the Americas that use it, A4 elsewhere; sizes in points. */
function defaultPaper() {
  return LETTER_REGIONS.has(regionCode())
    ? { name: 'Letter', width: 612, height: 792 }
    : { name: 'A4', width: 595.28, height: 841.89 }
}

function pageRule({ width, height, margins }) {
  const [top, right, bottom, left] = margins
  return `@page { size: ${round(width)}pt ${round(height)}pt; margin: ${round(top)}pt ${round(right)}pt ${round(bottom)}pt ${round(left)}pt; }`
}

const FONT_FALLBACKS = '"Segoe UI", "Segoe UI Emoji", "Segoe UI Symbol", "Nirmala UI", "Microsoft YaHei UI", "Yu Gothic UI", "Malgun Gothic", "Leelawadee UI", Arial, sans-serif'

const BASE_CSS = `
html { color: #111113; background: #fff; font-family: ${FONT_FALLBACKS}; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { margin: 0; font-size: 11pt; line-height: 1.45; overflow-wrap: break-word; }
h1, h2, h3, h4, h5, h6 { line-height: 1.22; break-after: avoid; page-break-after: avoid; }
h1 { font-size: 22pt; margin: 0 0 10pt; }
h2 { font-size: 17pt; margin: 16pt 0 7pt; }
h3 { font-size: 13.5pt; margin: 13pt 0 6pt; }
h4, h5, h6 { font-size: 11.5pt; margin: 11pt 0 5pt; }
p { margin: 0 0 8pt; }
ul, ol { margin: 0 0 8pt; padding-left: 22pt; }
li { margin: 0 0 3pt; }
li > p { margin: 0; }
table { margin: 6pt 0 10pt; border-collapse: collapse; }
th, td { padding: 3pt 5pt; border: 0.6pt solid #a1a1aa; vertical-align: top; }
th { background: #f4f4f5; font-weight: 600; }
thead { display: table-header-group; }
tr, img, pre, blockquote { break-inside: avoid; }
img { max-width: 100%; height: auto; }
blockquote { margin: 8pt 0; padding: 2pt 0 2pt 12pt; border-left: 2.5pt solid #d4d4d8; color: #3f3f46; }
a { color: #1d4ed8; text-decoration: underline; }
code { font-family: Consolas, "Cascadia Mono", "Courier New", monospace; font-size: 0.92em; background: #f4f4f5; padding: 0 2pt; border-radius: 2pt; }
pre { font-family: Consolas, "Cascadia Mono", "Courier New", monospace; font-size: 9.5pt; line-height: 1.38; white-space: pre-wrap; overflow-wrap: anywhere; background: #fafafa; border: 0.6pt solid #e4e4e7; padding: 7pt 9pt; margin: 0 0 9pt; }
pre code { background: none; padding: 0; font-size: inherit; }
hr { border: 0; border-top: 0.75pt solid #d4d4d8; margin: 12pt 0; }
hr.page-break, div.page-break { break-after: page; page-break-after: always; border: 0; height: 0; margin: 0; }
mark { background: #fff59d; color: inherit; }
.small-caps { font-variant: small-caps; }
.all-caps { text-transform: uppercase; }
`

function htmlDocument({ title, css, body, base, lang }) {
  return `<!doctype html>
<html${lang ? ` lang="${escapeHtml(lang)}"` : ''}><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${CONTENT_SECURITY_POLICY}">${base ? `<base href="${escapeHtml(base)}">` : ''}<title>${escapeHtml(title)}</title><style>${css}</style></head><body>${body}</body></html>`
}

// ---------------------------------------------------------------------------
// Plain text, Markdown and HTML
// ---------------------------------------------------------------------------

/** Plain text as preformatted blocks; form feeds become page breaks. */
function plainTextToHtml(text, { title = 'Document', paper = defaultPaper() } = {}) {
  const normalized = String(text || '').replace(/\r\n?/g, '\n').replace(/\u0000/g, '')
  const pages = normalized.split('\f')
  const body = pages.map((pageText, pageIndex) => {
    // Many short blocks lay out and paginate much faster than one huge text node.
    const lines = pageText.split('\n')
    const blocks = []
    for (let index = 0; index < lines.length; index += 120) {
      const chunk = lines.slice(index, index + 120).join('\n')
      blocks.push(`<pre class="plain">${escapeHtml(chunk)}${index + 120 < lines.length ? '\n' : ''}</pre>`)
    }
    return `${pageIndex ? '<div class="page-break"></div>' : ''}${blocks.join('')}`
  }).join('')
  const css = `${pageRule({ ...paper, margins: [56.7, 56.7, 56.7, 56.7] })}${BASE_CSS}
pre.plain { margin: 0; padding: 0; border: 0; background: none; font-size: 10pt; line-height: 1.32; tab-size: 4; white-space: pre-wrap; overflow-wrap: anywhere; break-inside: auto; }`
  return htmlDocument({ title, css, body: body || '<pre class="plain"></pre>' })
}

let markdownRenderer = null
function markdownIt() {
  if (!markdownRenderer) {
    const MarkdownIt = require('markdown-it')
    // Inline HTML (line breaks, <sup>, aligned pictures in READMEs) is kept and
    // passes through the same sanitizer as web pages; the print window runs no
    // scripts and has no network either way.
    markdownRenderer = new MarkdownIt({ html: true, linkify: true, typographer: false, breaks: false })
  }
  return markdownRenderer
}

/** YAML front matter (as a small preformatted block) and the Markdown after it. */
function splitFrontMatter(text) {
  const source = String(text || '').replace(/\r\n?/g, '\n').replace(/^\uFEFF/, '')
  const front = /^---\n([\s\S]*?)\n(---|\.\.\.)\n/.exec(source)
  if (!front) return { frontMatter: '', source }
  return { frontMatter: `<pre class="front-matter">${escapeHtml(front[1])}</pre>`, source: source.slice(front[0].length) }
}

/** Markdown to HTML. */
function markdownToHtmlBody(text) {
  const { frontMatter, source } = splitFrontMatter(text)
  return `${frontMatter}${sanitizeMarkup(markdownIt().render(source))}`
}

function markdownCss(paper) {
  return `${pageRule({ ...paper, margins: [56.7, 56.7, 56.7, 56.7] })}${BASE_CSS}
h1, h2 { padding-bottom: 3pt; border-bottom: 0.6pt solid #e4e4e7; }
pre.front-matter { color: #52525b; }
td, th { text-align: left; }
input[type=checkbox] { margin: 0 4pt 0 0; }`
}

function markdownToHtml(text, { title = 'Document', paper = defaultPaper(), base } = {}) {
  return htmlDocument({ title, css: markdownCss(paper), body: markdownToHtmlBody(text), base })
}

/**
 * markdownToHtml for very large files: rendered in slices (split at blank
 * lines outside code fences) so the main process keeps answering other
 * windows. Link reference definitions are shared by every slice.
 */
async function markdownToHtmlChunked(text, { title = 'Document', paper = defaultPaper(), base } = {}) {
  const { frontMatter, source } = splitFrontMatter(text)
  const references = source.match(/^ {0,3}\[[^\]\n]+\]:[^\n]*$/gm) || []
  const parts = []
  let fenced = false
  let current = []
  let size = 0
  for (const line of source.split('\n')) {
    if (/^ {0,3}(```|~~~)/.test(line)) fenced = !fenced
    current.push(line)
    size += line.length + 1
    if (!fenced && size > 200_000 && line.trim() === '') {
      parts.push(current.join('\n'))
      current = []
      size = 0
    }
  }
  if (current.length) parts.push(current.join('\n'))
  const bodies = [frontMatter]
  for (const part of parts) {
    bodies.push(sanitizeMarkup(markdownIt().render(`${part}\n\n${references.join('\n')}`)))
    await new Promise((resolve) => setImmediate(resolve))
  }
  return htmlDocument({ title, css: markdownCss(paper), body: bodies.join(''), base })
}

/**
 * A web page made safe to print: scripts, frames, plug-ins, refreshes and
 * event handlers are removed (scripts could not run anyway, and the network
 * is blocked), its charset is declared as UTF-8, and a default page size is
 * added when the page has none.
 */
/** Markup without scripts, frames, plug-ins, refreshes or event handlers. */
function sanitizeMarkup(html) {
  return String(html || '')
    .replace(/<script\b[\s\S]*?<\/script\s*>/gi, '')
    .replace(/<script\b[^>]*>/gi, '')
    .replace(/<noscript\b[\s\S]*?<\/noscript\s*>/gi, '')
    .replace(/<(iframe|frame|frameset|object|embed|applet|portal|template)\b[\s\S]*?<\/\1\s*>/gi, '')
    .replace(/<(iframe|frame|object|embed|applet|portal|base|link\b[^>]*rel\s*=\s*["']?(?:preload|prefetch|import|manifest)[^>]*)\b[^>]*>/gi, '')
    .replace(/<meta\b[^>]*(http-equiv|charset)[^>]*>/gi, '')
    .replace(/\s(on[a-z]+)\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/\s(href|src|action|formaction|xlink:href|srcset)\s*=\s*(["']?)\s*(javascript|vbscript):[^"'>\s]*\2/gi, ' $1=$2#$2')
}

function sanitizeHtmlDocument(html, { title = 'Document', paper = defaultPaper(), base } = {}) {
  let text = sanitizeMarkup(html)
  const hasPageRule = /@page\b/i.test(text)
  const head = `<meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="${CONTENT_SECURITY_POLICY}">${base ? `<base href="${escapeHtml(base)}">` : ''}${hasPageRule ? '' : `<style>${pageRule({ ...paper, margins: [42.5, 42.5, 42.5, 42.5] })} html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }</style>`}`
  if (!/<html[\s>]/i.test(text) && !/^\s*<!doctype/i.test(text)) {
    return `<!doctype html><html><head>${head}<title>${escapeHtml(title)}</title><style>${BASE_CSS}</style></head><body>${text}</body></html>`
  }
  if (/<head[\s>]/i.test(text)) return text.replace(/<head(\s[^>]*)?>/i, (match) => `${match}${head}`)
  if (/<html[\s>]/i.test(text)) return text.replace(/<html(\s[^>]*)?>/i, (match) => `${match}<head>${head}</head>`)
  return `<!doctype html><html><head>${head}</head><body>${text.replace(/^\s*<!doctype[^>]*>/i, '')}</body></html>`
}

// ---------------------------------------------------------------------------
// A small, non-validating XML reader (ODT and Word package parts). It never
// expands DTD entities, so external entities cannot be reached.
// ---------------------------------------------------------------------------

const XML_ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" }

function decodeXmlEntities(text) {
  if (text.indexOf('&') < 0) return text
  return text.replace(/&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);/g, (_match, entity) => {
    if (entity[0] !== '#') return XML_ENTITIES[entity]
    const code = entity[1] === 'x' || entity[1] === 'X' ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10)
    return Number.isFinite(code) && code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : ''
  })
}

function tagEnd(text, start) {
  let quote = ''
  for (let index = start; index < text.length; index += 1) {
    const character = text[index]
    if (quote) { if (character === quote) quote = '' }
    else if (character === '"' || character === "'") quote = character
    else if (character === '>') return index
  }
  return -1
}

/**
 * Parses XML into `{name, attributes, children}` elements; text children are
 * strings. Tolerates unbalanced end tags.
 */
function parseXml(source) {
  const text = String(source || '')
  const root = { name: '#document', attributes: {}, children: [] }
  const stack = [root]
  const attributePattern = /([^\s=/>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g
  let index = 0
  const addText = (value) => { if (value) stack[stack.length - 1].children.push(decodeXmlEntities(value)) }
  while (index < text.length) {
    const open = text.indexOf('<', index)
    if (open < 0) { addText(text.slice(index)); break }
    if (open > index) addText(text.slice(index, open))
    if (text.startsWith('<!--', open)) {
      const end = text.indexOf('-->', open + 4)
      index = end < 0 ? text.length : end + 3
      continue
    }
    if (text.startsWith('<![CDATA[', open)) {
      const end = text.indexOf(']]>', open + 9)
      stack[stack.length - 1].children.push(text.slice(open + 9, end < 0 ? text.length : end))
      index = end < 0 ? text.length : end + 3
      continue
    }
    if (text.startsWith('<?', open)) {
      const end = text.indexOf('?>', open + 2)
      index = end < 0 ? text.length : end + 2
      continue
    }
    if (text.startsWith('<!', open)) {
      // DOCTYPE, possibly with an internal subset; skipped, never expanded.
      let depth = 0
      let cursor = open + 2
      for (; cursor < text.length; cursor += 1) {
        if (text[cursor] === '[') depth += 1
        else if (text[cursor] === ']') depth -= 1
        else if (text[cursor] === '>' && depth <= 0) break
      }
      index = cursor + 1
      continue
    }
    const close = tagEnd(text, open + 1)
    if (close < 0) break
    const raw = text.slice(open + 1, close)
    index = close + 1
    if (raw[0] === '/') {
      const name = raw.slice(1).trim()
      for (let depth = stack.length - 1; depth > 0; depth -= 1) {
        if (stack[depth].name === name) { stack.length = depth; break }
      }
      continue
    }
    const selfClosing = raw.endsWith('/')
    const body = selfClosing ? raw.slice(0, -1) : raw
    const nameMatch = /^[^\s/>]+/.exec(body)
    if (!nameMatch) continue
    const attributes = {}
    attributePattern.lastIndex = 0
    const rest = body.slice(nameMatch[0].length)
    let match
    while ((match = attributePattern.exec(rest))) attributes[match[1]] = decodeXmlEntities(match[2] ?? match[3] ?? '')
    const element = { name: nameMatch[0], attributes, children: [] }
    stack[stack.length - 1].children.push(element)
    if (!selfClosing) stack.push(element)
  }
  return root
}

function elementChildren(element, name) {
  return (element?.children || []).filter((child) => typeof child !== 'string' && (!name || child.name === name))
}

function firstChild(element, name) {
  return (element?.children || []).find((child) => typeof child !== 'string' && child.name === name) || null
}

function findFirst(element, name) {
  if (!element || typeof element === 'string') return null
  for (const child of element.children || []) {
    if (typeof child === 'string') continue
    if (child.name === name) return child
    const nested = findFirst(child, name)
    if (nested) return nested
  }
  return null
}

function findAll(element, name, output = []) {
  for (const child of element?.children || []) {
    if (typeof child === 'string') continue
    if (child.name === name) output.push(child)
    findAll(child, name, output)
  }
  return output
}

function textOf(element) {
  if (typeof element === 'string') return element
  return (element?.children || []).map(textOf).join('')
}

async function zipText(zip, name) {
  const file = zip.file(name)
  return file ? file.async('string') : ''
}

// ---------------------------------------------------------------------------
// Word (.docx) through mammoth
// ---------------------------------------------------------------------------

const PARAGRAPH_MARK = '\uE000\uE00A'
const RUN_MARK = '\uE001\uE00A'
const MARK_END = '\uE00B'
const RUN_CLOSE = '\uE002\uE00C'

const DOCX_STYLE_MAP = [
  "p[style-name='Title'] => h1.doc-title:fresh",
  "p[style-name='Subtitle'] => p.doc-subtitle:fresh",
  "p[style-name='Quote'] => blockquote > p:fresh",
  "p[style-name='Intense Quote'] => blockquote.intense > p:fresh",
  "p[style-name='Caption'] => p.caption:fresh",
  "p[style-name='No Spacing'] => p.no-spacing:fresh",
  "r[style-name='Strong'] => strong",
  "r[style-name='Emphasis'] => em",
  'u => u',
  'strike => s',
  'highlight => mark',
  'small-caps => span.small-caps',
  'all-caps => span.all-caps',
  "br[type='page'] => hr.page-break",
]

function twips(value) {
  const number = Number.parseFloat(value)
  return Number.isFinite(number) ? number * POINTS_PER_TWIP : null
}

function attribute(element, name) {
  return element?.attributes?.[name]
}

function onOff(element) {
  if (!element) return null
  const value = attribute(element, 'w:val')
  return !(value === 'false' || value === '0' || value === 'off')
}

function themeFonts(themeXml) {
  if (!themeXml) return {}
  const theme = parseXml(themeXml)
  const major = findFirst(findFirst(theme, 'a:majorFont'), 'a:latin')
  const minor = findFirst(findFirst(theme, 'a:minorFont'), 'a:latin')
  return { major: attribute(major, 'typeface') || null, minor: attribute(minor, 'typeface') || null }
}

function runFont(rPr, theme) {
  const fonts = firstChild(rPr, 'w:rFonts')
  if (!fonts) return null
  const direct = attribute(fonts, 'w:ascii') || attribute(fonts, 'w:hAnsi')
  if (direct) return direct
  const themed = attribute(fonts, 'w:asciiTheme') || attribute(fonts, 'w:hAnsiTheme') || ''
  if (/^major/i.test(themed)) return theme.major || null
  if (/^minor/i.test(themed)) return theme.minor || null
  return null
}

function readRunProperties(rPr, theme) {
  if (!rPr) return {}
  const size = Number.parseFloat(attribute(firstChild(rPr, 'w:sz'), 'w:val'))
  const color = attribute(firstChild(rPr, 'w:color'), 'w:val')
  return {
    font: runFont(rPr, theme) || undefined,
    size: Number.isFinite(size) ? size / 2 : undefined,
    color: color && /^[0-9a-f]{6}$/i.test(color) ? color : undefined,
    bold: onOff(firstChild(rPr, 'w:b')) ?? undefined,
    italic: onOff(firstChild(rPr, 'w:i')) ?? undefined,
    caps: onOff(firstChild(rPr, 'w:caps')) ?? undefined,
  }
}

function readParagraphProperties(pPr) {
  if (!pPr) return {}
  const spacing = firstChild(pPr, 'w:spacing')
  const line = Number.parseFloat(attribute(spacing, 'w:line'))
  const rule = attribute(spacing, 'w:lineRule') || 'auto'
  return {
    before: twips(attribute(spacing, 'w:before')) ?? undefined,
    after: twips(attribute(spacing, 'w:after')) ?? undefined,
    line: Number.isFinite(line) ? (rule === 'auto' ? { multiple: line / 240 } : { points: line / 20 }) : undefined,
    align: attribute(firstChild(pPr, 'w:jc'), 'w:val') || undefined,
  }
}

function defined(object) {
  return Object.fromEntries(Object.entries(object).filter(([, value]) => value !== undefined && value !== null))
}

/** Page size and margins (points) of a Word section, or null when unusable. */
function sectionGeometry(section) {
  const size = firstChild(section, 'w:pgSz')
  const margin = firstChild(section, 'w:pgMar')
  const width = twips(attribute(size, 'w:w'))
  const height = twips(attribute(size, 'w:h'))
  if (!(width > 72 && height > 72 && width < 14400 && height < 14400)) return null
  const side = (name, fallback) => {
    const value = twips(attribute(margin, name))
    return value === null ? fallback : Math.min(Math.abs(value), Math.min(width, height) / 3)
  }
  return { width, height, margins: [side('w:top', 72), side('w:right', 72), side('w:bottom', 72), side('w:left', 72)] }
}

/**
 * Page set-up of every section (in order, the body's own section last),
 * default text and heading styles from a Word package.
 */
async function readDocxLayout(zip, documentXml) {
  const [stylesXml, themeXml] = await Promise.all([
    zipText(zip, 'word/styles.xml'),
    zipText(zip, 'word/theme/theme1.xml'),
  ])
  const theme = themeFonts(themeXml)
  const layout = {
    page: null,
    sections: [],
    text: { font: theme.minor || 'Calibri', size: 11, after: 8, line: { multiple: 1.08 } },
    styles: {},
    drawings: [],
  }
  const document = parseXml(documentXml)
  const body = findFirst(document, 'w:body')
  let previous = null
  for (const section of findAll(body, 'w:sectPr')) {
    previous = sectionGeometry(section) || previous
    layout.sections.push(previous)
  }
  layout.page = layout.sections.find(Boolean) || null
  layout.drawings = docxDrawingSizes(body)

  if (stylesXml) {
    const styles = parseXml(stylesXml)
    const defaults = findFirst(styles, 'w:docDefaults')
    const defaultRun = readRunProperties(findFirst(findFirst(defaults, 'w:rPrDefault'), 'w:rPr'), theme)
    const defaultParagraph = readParagraphProperties(findFirst(findFirst(defaults, 'w:pPrDefault'), 'w:pPr'))
    const byId = new Map()
    for (const style of findAll(styles, 'w:style')) byId.set(attribute(style, 'w:styleId'), style)
    const resolve = (style, depth = 0) => {
      if (!style || depth > 8) return { run: {}, paragraph: {} }
      const parent = resolve(byId.get(attribute(firstChild(style, 'w:basedOn'), 'w:val')), depth + 1)
      return {
        run: { ...parent.run, ...defined(readRunProperties(firstChild(style, 'w:rPr'), theme)) },
        paragraph: { ...parent.paragraph, ...defined(readParagraphProperties(firstChild(style, 'w:pPr'))) },
      }
    }
    const normalStyle = findAll(styles, 'w:style').find((style) => attribute(style, 'w:type') === 'paragraph' && onOff({ attributes: { 'w:val': attribute(style, 'w:default') } }) && attribute(style, 'w:default'))
      || byId.get('Normal')
    const normal = resolve(normalStyle)
    const run = { ...defined(defaultRun), ...normal.run }
    const paragraph = { ...defined(defaultParagraph), ...normal.paragraph }
    layout.text = {
      font: run.font || layout.text.font,
      size: run.size || 10,
      color: run.color,
      after: paragraph.after ?? 0,
      before: paragraph.before ?? 0,
      line: paragraph.line || { multiple: 1 },
    }
    for (const [styleId, tag] of [['Heading1', 'h1'], ['Heading2', 'h2'], ['Heading3', 'h3'], ['Heading4', 'h4'], ['Heading5', 'h5'], ['Heading6', 'h6'], ['Title', 'h1.doc-title'], ['Subtitle', 'p.doc-subtitle']]) {
      const style = byId.get(styleId)
        || findAll(styles, 'w:style').find((candidate) => attribute(firstChild(candidate, 'w:name'), 'w:val')?.toLowerCase() === styleId.replace(/(\d)$/, ' $1').toLowerCase())
      if (style) layout.styles[tag] = resolve(style)
    }
  }
  return layout
}

/**
 * Display sizes (points) of the pictures mammoth will convert, in the order it
 * converts them: DrawingML pictures (inline or anchored) and VML image data,
 * skipping deleted runs and the Choice branch of alternate content.
 */
function docxDrawingSizes(body) {
  const sizes = []
  const visit = (element, extent, shapeSize) => {
    for (const child of element?.children || []) {
      if (typeof child === 'string') continue
      if (child.name === 'w:del' || child.name === 'mc:Choice' || child.name === 'w:pPr' || child.name === 'w:rPr') continue
      if (child.name === 'wp:inline' || child.name === 'wp:anchor') {
        const box = firstChild(child, 'wp:extent')
        const cx = Number(attribute(box, 'cx'))
        const cy = Number(attribute(box, 'cy'))
        visit(child, cx > 0 && cy > 0 ? { width: cx * POINTS_PER_EMU, height: cy * POINTS_PER_EMU } : null, shapeSize)
        continue
      }
      if (child.name === 'a:blip' && extent !== undefined) {
        sizes.push(extent)
        continue
      }
      if (child.name === 'v:shape' || child.name === 'v:rect') {
        visit(child, extent, vmlSize(attribute(child, 'style')))
        continue
      }
      if (child.name === 'v:imagedata' && attribute(child, 'r:id')) {
        sizes.push(shapeSize || null)
        continue
      }
      // Only pictures (pic:pic) become images; charts and shapes do not.
      if (child.name === 'a:graphicData' && !findFirst(child, 'pic:pic')) continue
      visit(child, extent, shapeSize)
    }
  }
  visit(body, undefined, null)
  return sizes
}

function cssLengthToPoints(value) {
  const match = /^\s*(-?[\d.]+)\s*(pt|px|in|cm|mm|pc|em)?\s*$/i.exec(String(value || ''))
  if (!match) return null
  const number = Number.parseFloat(match[1])
  const unit = (match[2] || 'px').toLowerCase()
  const factor = { pt: 1, px: 0.75, in: 72, cm: 72 / 2.54, mm: 72 / 25.4, pc: 12, em: 12 }[unit]
  return Number.isFinite(number) && factor ? number * factor : null
}

function vmlSize(style) {
  const values = Object.fromEntries(String(style || '').split(';').map((part) => part.split(':').map((item) => item.trim().toLowerCase())).filter((pair) => pair.length === 2))
  const width = cssLengthToPoints(values.width)
  const height = cssLengthToPoints(values.height)
  return width > 0 && height > 0 ? { width, height } : null
}

function lineHeightCss(line) {
  if (!line) return '1.2'
  if (line.points) return `${round(line.points)}pt`
  // Word's "single" spacing is about 1.17 of the font size for common faces.
  return String(round(Math.max(0.8, Math.min(4, line.multiple * 1.17)), 3))
}

function styleCss(selector, style, text) {
  if (!style) return ''
  const run = style.run || {}
  const paragraph = style.paragraph || {}
  const declarations = []
  if (run.font) declarations.push(`font-family: ${cssString(run.font)}, ${cssString(text.font)}, ${FONT_FALLBACKS}`)
  if (run.size) declarations.push(`font-size: ${round(run.size)}pt`)
  if (run.color && run.color.toLowerCase() !== 'auto') declarations.push(`color: #${run.color}`)
  declarations.push(`font-weight: ${run.bold === false ? 400 : run.bold ? 700 : 'inherit'}`)
  if (run.italic) declarations.push('font-style: italic')
  if (run.caps) declarations.push('text-transform: uppercase')
  declarations.push(`margin: ${round(paragraph.before ?? 12)}pt 0 ${round(paragraph.after ?? 4)}pt`)
  if (paragraph.line) declarations.push(`line-height: ${lineHeightCss(paragraph.line)}`)
  if (paragraph.align === 'center') declarations.push('text-align: center')
  if (paragraph.align === 'right' || paragraph.align === 'end') declarations.push('text-align: right')
  return `${selector} { ${declarations.join('; ')}; }\n`
}

function docxCss(layout, paper) {
  const text = layout.text
  const page = layout.page || { ...paper, margins: [56.7, 56.7, 56.7, 56.7] }
  let css = `${pageRule(page)}${BASE_CSS}
body { font-family: ${cssString(text.font)}, Calibri, ${FONT_FALLBACKS}; font-size: ${round(text.size)}pt; line-height: ${lineHeightCss(text.line)};${text.color && text.color !== 'auto' ? ` color: #${text.color};` : ''} }
p, li { margin: ${round(text.before || 0)}pt 0 ${round(text.after || 0)}pt; }
li > p { margin: 0; }
p:empty::after { content: "\\00a0"; }
p.no-spacing { margin: 0; }
p.caption { font-size: 0.82em; font-style: italic; color: #44546a; }
p.doc-subtitle { color: #5a5a5a; font-size: 1.15em; }
table { border-collapse: collapse; margin: 0 0 ${round(Math.max(text.after || 0, 4))}pt; }
th, td { border: 0.5pt solid #7f7f7f; padding: 0 5.4pt; vertical-align: top; }
th { background: none; font-weight: inherit; }
td > p:last-child, th > p:last-child { margin-bottom: 0; }
img.missing-image { display: none; }
ol { list-style: decimal; } ol ol { list-style: lower-alpha; } ol ol ol { list-style: lower-roman; }
a[href^="#footnote-ref-"], a[href^="#endnote-ref-"] { display: none; }
`
  for (const [tag, style] of Object.entries(layout.styles)) css += styleCss(tag, style, text)
  return css
}

function sanitizeFontName(value) {
  return String(value || '').replace(/[^\p{L}\p{N} _.-]/gu, '').trim().slice(0, 64)
}

/** Marks mammoth cannot express (alignment, indents, sizes, fonts) as private-use text. */
function docxTransform(mammoth, layout) {
  const defaultSize = layout.text.size
  const defaultFont = String(layout.text.font || '').toLowerCase()
  const markParagraph = (paragraph) => {
    const parts = []
    const align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[paragraph.alignment]
    if (align) parts.push(`a=${align}`)
    if (!paragraph.numbering) {
      const start = twips(paragraph.indent?.start)
      const firstLine = twips(paragraph.indent?.firstLine)
      const hanging = twips(paragraph.indent?.hanging)
      if (start && start > 0) parts.push(`l=${round(start)}`)
      if (firstLine) parts.push(`f=${round(firstLine)}`)
      else if (hanging) parts.push(`f=${round(-hanging)}`)
    }
    if (!parts.length) return paragraph
    return { ...paragraph, children: [{ type: 'text', value: `${PARAGRAPH_MARK}${parts.join(';')}${MARK_END}` }, ...paragraph.children] }
  }
  const markRun = (run) => {
    const parts = []
    if (run.fontSize && Math.abs(run.fontSize - defaultSize) >= 0.5) parts.push(`s=${round(run.fontSize, 1)}`)
    const font = sanitizeFontName(run.font)
    if (font && font.toLowerCase() !== defaultFont) parts.push(`f=${font}`)
    if (!parts.length || !run.children?.length) return run
    return {
      ...run,
      children: [
        { type: 'text', value: `${RUN_MARK}${parts.join(';')}${MARK_END}` },
        ...run.children,
        { type: 'text', value: RUN_CLOSE },
      ],
    }
  }
  const walk = (element) => {
    let next = element
    if (Array.isArray(next.children)) next = { ...next, children: next.children.map(walk) }
    if (next.type === 'paragraph') return markParagraph(next)
    if (next.type === 'run') return markRun(next)
    return next
  }
  return walk
}

function markerDeclarations(payload, kind) {
  const declarations = []
  for (const part of String(payload || '').split(';')) {
    const [key, ...rest] = part.split('=')
    const value = rest.join('=')
    if (kind === 'p') {
      if (key === 'a' && /^(center|right|justify)$/.test(value)) declarations.push(`text-align: ${value}`)
      if (key === 'l' && Number.isFinite(Number(value))) declarations.push(`margin-left: ${Number(value)}pt`)
      if (key === 'f' && Number.isFinite(Number(value))) declarations.push(`text-indent: ${Number(value)}pt`)
    } else {
      if (key === 's' && Number.isFinite(Number(value)) && Number(value) > 0) declarations.push(`font-size: ${Number(value)}pt`)
      if (key === 'f' && value) declarations.push(`font-family: ${cssString(value)}, inherit`)
    }
  }
  return declarations.join('; ')
}

function applyDocxMarkers(html) {
  return html
    .replace(/<p>\uE006(\d+)\uE00B<\/p>/g, '<!--docx-section:$1-->')
    .replace(/(<(p|h[1-6]|li|td|th|pre)\b[^>]*?)>\uE000\uE00A([^\uE00B<]*)\uE00B/g, (_match, openTag, _tag, payload) => {
      const css = markerDeclarations(payload, 'p')
      return css ? `${openTag} style="${escapeHtml(css)}">` : `${openTag}>`
    })
    .replace(/\uE000\uE00A[^\uE00B]*\uE00B/g, '')
    .replace(/\uE001\uE00A([^\uE00B<]*)\uE00B/g, (_match, payload) => {
      const css = markerDeclarations(payload, 'r')
      return css ? `<span style="${escapeHtml(css)}">` : '<span>'
    })
    .replace(/\uE002\uE00C/g, '</span>')
    .replace(/\uE004([0-9A-Fa-f]{6})\uE00B/g, (_match, color) => `<span style="color: #${color.toLowerCase()}">`)
    .replace(/\uE005/g, '</span>')
    .replace(/\uE004[^\uE00B<]{0,8}\uE00B?|\uE006\d*\uE00B?/g, '')
    .replace(/[\uE000-\uE006][\uE00A-\uE00C]?/g, '')
}

const SECTION_MARK = '\uE006'
const COLOR_MARK = '\uE004'
const COLOR_CLOSE = '\uE005'

/** Index just past the `</w:p>` that closes the paragraph open at `from`. */
function paragraphEnd(xml, from) {
  const tags = /<(\/?)w:p(?=[\s>/])([^>]*)>/g
  tags.lastIndex = from
  let depth = 0
  let match
  while ((match = tags.exec(xml))) {
    if (match[1]) {
      if (depth === 0) return match.index + match[0].length
      depth -= 1
    } else if (!match[2].endsWith('/')) depth += 1
  }
  return -1
}

/**
 * Mammoth drops section breaks and run colours, so they travel through it as
 * private-use marker text: a marker paragraph after each paragraph that ends
 * a section, and colour markers around the text of coloured runs. Runs that
 * hold drawings or objects (which can nest runs) keep their own colour.
 * @returns {{xml: string, changed: boolean}}
 */
function prepareDocxXml(documentXml) {
  let xml = String(documentXml || '')
  let changed = false
  xml = xml.replace(/<w:r(?=[\s>])[^>]*>[\s\S]*?<\/w:r>/g, (run) => {
    if (/<w:(drawing|pict|object)\b|<mc:AlternateContent\b/.test(run)) return run
    const properties = /<w:rPr(?:\s[^>]*)?>([\s\S]*?)<\/w:rPr>/.exec(run)
    const color = properties && /<w:color\b[^>]*\bw:val="([0-9A-Fa-f]{6})"/.exec(properties[1])
    if (!color) return run
    const first = /<w:t(?:\s[^>]*[^/])?>/.exec(run)
    const last = run.lastIndexOf('</w:t>')
    if (!first || last < first.index) return run
    changed = true
    const open = first.index + first[0].length
    return `${run.slice(0, open)}${COLOR_MARK}${color[1]}${MARK_END}${run.slice(open, last)}${COLOR_CLOSE}${run.slice(last)}`
  })
  const inserts = []
  const sectionProperties = /<w:pPr\b[^>]*>(?:(?!<\/w:pPr>)[\s\S])*?<w:sectPr\b/g
  let match
  while ((match = sectionProperties.exec(xml))) {
    const end = paragraphEnd(xml, match.index)
    if (end > 0) inserts.push({ at: end, index: inserts.length + 1 })
  }
  for (const insert of inserts.reverse()) {
    xml = `${xml.slice(0, insert.at)}<w:p><w:r><w:t>${SECTION_MARK}${insert.index}${MARK_END}</w:t></w:r></w:p>${xml.slice(insert.at)}`
    changed = true
  }
  return { xml, changed }
}

/**
 * Wraps each Word section in a block with its own CSS named page, so mixed
 * portrait and landscape sections print on pages of their own size.
 * @returns {{html: string, css: string}}
 */
function applyDocxSections(html, sections, fallback) {
  const geometries = sections.map((section) => section || fallback)
  if (geometries.length < 2) return { html: html.replace(/<!--docx-section:\d+-->/g, ''), css: '' }
  const names = new Map()
  let css = ''
  const pageName = (geometry) => {
    const key = JSON.stringify(geometry)
    if (!names.has(key)) {
      const name = `section${names.size + 1}`
      names.set(key, name)
      css += pageRule(geometry).replace('@page {', `@page ${name} {`)
    }
    return names.get(key)
  }
  const open = (index) => `<div class="docx-section" style="page: ${pageName(geometries[Math.min(index, geometries.length - 1)])}">`
  const body = html.replace(/<!--docx-section:(\d+)-->/g, (_match, index) => `</div>${open(Number(index))}`)
  return { html: `${open(0)}${body}</div>`, css }
}

const BROWSER_IMAGE_TYPES = new Set(['image/png', 'image/jpeg', 'image/gif', 'image/bmp', 'image/webp', 'image/svg+xml', 'image/x-icon', 'image/avif'])
const TRANSPARENT_PIXEL = 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7'

/**
 * Converts a Word document to printable HTML. Pictures are written as files
 * next to the HTML (returned in `files`) at their Word display size.
 * @returns {Promise<{html: string, files: Map<string, Buffer>, warnings: string[], layout: object}>}
 */
async function docxToHtml(bytes, { title = 'Document', paper = defaultPaper() } = {}) {
  const mammoth = require('mammoth')
  const JSZip = require('jszip')
  let buffer = toBuffer(bytes)
  let zip
  try {
    zip = await JSZip.loadAsync(buffer)
  } catch (error) {
    throw codedError('FILE_DAMAGED', "This Word document is damaged or isn't really a .docx file.", { technical: error?.message })
  }
  if (!zip.file('word/document.xml')) throw codedError('FILE_DAMAGED', "This Word document is damaged or isn't really a .docx file.")
  const documentXml = await zipText(zip, 'word/document.xml')
  const layout = await readDocxLayout(zip, documentXml)
  const prepared = prepareDocxXml(documentXml)
  if (prepared.changed) {
    zip.file('word/document.xml', prepared.xml)
    buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'STORE' })
  }
  const files = new Map()
  const warnings = new Set()
  let pictureCount = 0
  const result = await mammoth.convertToHtml({ buffer }, {
    styleMap: DOCX_STYLE_MAP,
    includeDefaultStyleMap: true,
    ignoreEmptyParagraphs: false,
    transformDocument: docxTransform(mammoth, layout),
    convertImage: mammoth.images.imgElement(async (image) => {
      const index = pictureCount++
      const contentType = String(image.contentType || '').toLowerCase()
      if (!BROWSER_IMAGE_TYPES.has(contentType)) {
        warnings.add('Some pictures (for example EMF or WMF drawings) could not be shown.')
        return { src: TRANSPARENT_PIXEL, class: 'missing-image' }
      }
      const data = await image.read()
      const extension = { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/webp': 'webp', 'image/svg+xml': 'svg', 'image/x-icon': 'ico', 'image/avif': 'avif' }[contentType]
      const name = `picture-${index + 1}.${extension}`
      files.set(name, toBuffer(data))
      return { src: name, 'data-picture': String(index) }
    }),
  })
  let html = applyDocxMarkers(result.value)
  if (layout.drawings.length === pictureCount) {
    html = html.replace(/<img\b([^>]*?)\sdata-picture="(\d+)"([^>]*)>/g, (match, before, index, after) => {
      const size = layout.drawings[Number(index)]
      if (!size) return `<img${before}${after}>`
      return `<img${before}${after} style="width: ${round(size.width)}pt; height: ${round(size.height)}pt">`
    })
  }
  html = html.replace(/\sdata-picture="\d+"/g, '')
  for (const message of result.messages || []) {
    if (message.type === 'warning' && /image|picture/i.test(message.message)) warnings.add('Some pictures could not be shown.')
  }
  const sections = applyDocxSections(html, layout.sections, layout.page || { ...paper, margins: [56.7, 56.7, 56.7, 56.7] })
  const css = `${docxCss(layout, paper)}${sections.css}`
  return { html: htmlDocument({ title, css, body: sections.html }), files, warnings: [...warnings], layout }
}

// ---------------------------------------------------------------------------
// Spreadsheets through SheetJS
// ---------------------------------------------------------------------------

function quietly(task) {
  // SheetJS reports recoverable format quirks on the console.
  const { error, warn } = console
  console.error = () => {}
  console.warn = () => {}
  try { return task() } finally {
    console.error = error
    console.warn = warn
  }
}

function sheetCell(sheet, row, column, XLSX) {
  if (sheet['!data']) return sheet['!data'][row]?.[column]
  return sheet[XLSX.utils.encode_cell({ r: row, c: column })]
}

function definedRange(workbook, sheetIndex, name, XLSX) {
  const names = workbook.Workbook?.Names || []
  const entry = names.find((item) => item.Name === name && Number(item.Sheet) === sheetIndex)
  if (!entry?.Ref) return null
  const reference = decodeXmlEntities(String(entry.Ref)).replace(/^.*!/, '').split(',')[0].replace(/\$/g, '')
  try {
    if (/^\d+:\d+$/.test(reference)) {
      const [first, last] = reference.split(':').map((value) => Number(value) - 1)
      return { rows: [first, last] }
    }
    return XLSX.utils.decode_range(reference)
  } catch { return null }
}

function cellText(cell, XLSX) {
  if (!cell || cell.t === 'z') return ''
  let text = cell.w
  if (text === undefined) {
    try { text = XLSX.utils.format_cell(cell) } catch { text = cell.v === undefined ? '' : String(cell.v) }
  }
  return String(text ?? '').replace(/\r\n?/g, '\n').trim()
}

/**
 * Turns a workbook (or CSV) into one printable HTML table per visible sheet,
 * honouring print areas, print-title rows, merged cells, hidden rows and
 * columns, and column widths. Wide sheets are turned sideways and scaled to
 * fit the page width.
 * @returns {{html: string, warnings: string[], sheets: number}}
 */
function spreadsheetToHtml(bytes, { title = 'Workbook', extension = 'xlsx', paper = defaultPaper(), hasDrawings = false } = {}) {
  const XLSX = require('xlsx')
  const buffer = toBuffer(bytes)
  const delimited = extension === 'csv' || extension === 'tsv'
  let workbook
  try {
    workbook = quietly(() => delimited
      ? XLSX.read(decodeText(buffer).text, { type: 'string', raw: true, dense: true, ...(extension === 'tsv' ? { FS: '\t' } : {}) })
      : XLSX.read(buffer, { type: 'buffer', cellDates: true, cellStyles: true, cellNF: true, dense: true }))
  } catch (error) {
    throw codedError('FILE_DAMAGED', "This spreadsheet is damaged or isn't in a format Simple can read.", { technical: error?.message })
  }
  const warnings = []
  if (hasDrawings) warnings.push("Charts and pictures in the workbook aren't included.")
  const sections = []
  let totalCells = 0
  let truncated = false
  workbook.SheetNames.forEach((sheetName, sheetIndex) => {
    const sheet = workbook.Sheets[sheetName]
    if (!sheet || !sheet['!ref']) return
    if (Number(workbook.Workbook?.Sheets?.[sheetIndex]?.Hidden) > 0) return
    let range
    try { range = XLSX.utils.decode_range(sheet['!ref']) } catch { return }
    const printArea = definedRange(workbook, sheetIndex, '_xlnm.Print_Area', XLSX)
    if (printArea?.s && printArea?.e) {
      range = {
        s: { r: Math.max(range.s.r, printArea.s.r), c: Math.max(range.s.c, printArea.s.c) },
        e: { r: Math.min(range.e.r, printArea.e.r), c: Math.min(range.e.c, printArea.e.c) },
      }
    }
    const titles = definedRange(workbook, sheetIndex, '_xlnm.Print_Titles', XLSX)
    const titleRows = titles?.rows || (titles?.s && titles.s.c === 0 && titles.e.c >= 16383 ? [titles.s.r, titles.e.r] : null)
    const rowsInfo = sheet['!rows'] || []
    const columnsInfo = sheet['!cols'] || []
    // Trim empty trailing rows and columns inside the range.
    let lastRow = range.s.r - 1
    let lastColumn = range.s.c - 1
    for (let row = range.s.r; row <= range.e.r; row += 1) {
      for (let column = range.s.c; column <= range.e.c; column += 1) {
        if (cellText(sheetCell(sheet, row, column, XLSX), XLSX)) {
          if (row > lastRow) lastRow = row
          if (column > lastColumn) lastColumn = column
        }
      }
    }
    if (lastRow < range.s.r || lastColumn < range.s.c) return
    const merges = new Map()
    const covered = new Set()
    for (const merge of sheet['!merges'] || []) {
      if (merge.s.r > lastRow || merge.s.c > lastColumn) continue
      merges.set(`${merge.s.r}:${merge.s.c}`, merge)
      for (let row = merge.s.r; row <= merge.e.r; row += 1) {
        for (let column = merge.s.c; column <= merge.e.c; column += 1) {
          if (row !== merge.s.r || column !== merge.s.c) covered.add(`${row}:${column}`)
        }
      }
    }
    const columns = []
    for (let column = range.s.c; column <= lastColumn; column += 1) {
      const info = columnsInfo[column] || {}
      if (info.hidden) continue
      const pixels = Number(info.wpx) || (Number(info.wch) ? Number(info.wch) * 7 + 5 : 0) || (Number(info.width) ? Number(info.width) * 7 + 5 : 0) || 64
      columns.push({ column, width: Math.max(12, Math.min(900, pixels)) * 0.75 })
    }
    const tableWidth = columns.reduce((sum, item) => sum + item.width, 0)
    const rowHtml = (row, cellTag) => {
      const cells = []
      for (const { column } of columns) {
        const key = `${row}:${column}`
        if (covered.has(key)) continue
        const cell = sheetCell(sheet, row, column, XLSX)
        const merge = merges.get(key)
        let colspan = 1
        let rowspan = 1
        if (merge) {
          colspan = columns.filter((item) => item.column >= merge.s.c && item.column <= Math.min(merge.e.c, lastColumn)).length || 1
          rowspan = Math.max(1, Math.min(merge.e.r, lastRow) - merge.s.r + 1)
        }
        const text = cellText(cell, XLSX)
        const align = cell?.t === 'n' ? 'num' : cell?.t === 'b' || cell?.t === 'e' ? 'mid' : ''
        cells.push(`<${cellTag}${colspan > 1 ? ` colspan="${colspan}"` : ''}${rowspan > 1 ? ` rowspan="${rowspan}"` : ''}${align ? ` class="${align}"` : ''}>${escapeHtml(text).replace(/\n/g, '<br>')}</${cellTag}>`)
      }
      return `<tr>${cells.join('')}</tr>`
    }
    const head = []
    const bodyRows = []
    for (let row = range.s.r; row <= lastRow; row += 1) {
      if (rowsInfo[row]?.hidden) continue
      if (totalCells > MAX_SHEET_CELLS) { truncated = true; break }
      totalCells += columns.length
      if (titleRows && row >= titleRows[0] && row <= titleRows[1]) head.push(rowHtml(row, 'th'))
      else bodyRows.push(rowHtml(row, 'td'))
    }
    sections.push({ name: decodeXmlEntities(sheetName), tableWidth, html: `<table style="width: ${round(tableWidth)}pt"><colgroup>${columns.map((item) => `<col style="width: ${round(item.width)}pt">`).join('')}</colgroup>${head.length ? `<thead>${head.join('')}</thead>` : ''}<tbody>${bodyRows.join('')}</tbody></table>` })
  })
  if (truncated) warnings.push('This workbook is very large; only the first part of it was turned into pages.')
  if (!sections.length) {
    sections.push({ name: '', tableWidth: 0, html: '<p class="empty">This workbook has no visible cells.</p>' })
  }
  // One page orientation for the document; sheets wider than the page are scaled to fit.
  const widest = Math.max(...sections.map((section) => section.tableWidth))
  const margins = [43.2, 28.8, 43.2, 28.8]
  const portraitWidth = paper.width - margins[1] - margins[3]
  const landscape = widest > portraitWidth
  const page = landscape ? { width: paper.height, height: paper.width, margins } : { width: paper.width, height: paper.height, margins }
  const printable = page.width - margins[1] - margins[3]
  const body = sections.map((section, index) => {
    const zoom = section.tableWidth > printable ? Math.max(0.3, printable / section.tableWidth) : 1
    const heading = sections.length > 1 && section.name ? `<h2 class="sheet-name">${escapeHtml(section.name)}</h2>` : ''
    return `<section class="sheet"${index ? ' style="break-before: page"' : ''}>${heading}<div${zoom < 1 ? ` style="zoom: ${round(zoom, 4)}"` : ''}>${section.html}</div></section>`
  }).join('')
  const css = `${pageRule(page)}${BASE_CSS}
body { font-family: Calibri, ${FONT_FALLBACKS}; font-size: 10pt; line-height: 1.25; }
h2.sheet-name { font-size: 10pt; font-weight: 600; color: #52525b; margin: 0 0 6pt; border: 0; }
table { table-layout: fixed; border-collapse: collapse; margin: 0; }
th, td { border: 0.5pt solid #d4d4d8; padding: 1.5pt 3pt; overflow: hidden; vertical-align: bottom; white-space: pre-wrap; overflow-wrap: anywhere; }
th { background: #f4f4f5; font-weight: 600; text-align: left; }
td.num { text-align: right; white-space: nowrap; }
td.mid { text-align: center; }
tr { break-inside: avoid; }
p.empty { color: #71717a; }`
  return { html: htmlDocument({ title, css, body }), warnings, sheets: sections.length }
}

// ---------------------------------------------------------------------------
// RTF (text, basic formatting, tables and PNG/JPEG pictures)
// ---------------------------------------------------------------------------

const RTF_CHARSET_CODEPAGES = { 0: 1252, 77: 10000, 128: 932, 129: 949, 130: 1361, 134: 936, 136: 950, 161: 1253, 162: 1254, 163: 1258, 177: 1255, 178: 1256, 186: 1257, 204: 1251, 222: 874, 238: 1250, 255: 437 }
const CODEPAGE_LABELS = { 437: 'ibm866', 874: 'windows-874', 932: 'shift_jis', 936: 'gbk', 949: 'euc-kr', 950: 'big5', 1250: 'windows-1250', 1251: 'windows-1251', 1252: 'windows-1252', 1253: 'windows-1253', 1254: 'windows-1254', 1255: 'windows-1255', 1256: 'windows-1256', 1257: 'windows-1257', 1258: 'windows-1258', 10000: 'macintosh', 20866: 'koi8-r', 65001: 'utf-8' }
const RTF_SKIPPED = new Set([
  'fonttbl', 'colortbl', 'stylesheet', 'info', 'listtable', 'listoverridetable', 'revtbl', 'rsidtbl', 'generator',
  'xmlnstbl', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'filetbl', 'pgdsctbl', 'mmathPr',
  'header', 'headerl', 'headerr', 'headerf', 'footer', 'footerl', 'footerr', 'footerf', 'fldinst', 'pn', 'pntxta',
  'pntxtb', 'bkmkstart', 'bkmkend', 'object', 'objdata', 'nonshppict', 'shpinst', 'sp', 'sn', 'sv', 'annotation',
  'atnid', 'atnauthor', 'atnref', 'atrfstart', 'atrfend', 'template', 'docvar', 'xe', 'tc', 'blipuid', 'wgrffmtfilter',
  'aftnsep', 'aftnsepc', 'aftncn', 'ftnsep', 'ftnsepc', 'ftncn', 'background', 'falt', 'panose', 'fname', 'listtext',
  'ud', 'upr', 'oldcprops', 'oldpprops', 'oldtprops', 'oldsprops', 'pgp', 'protusertbl', 'mhtmltag', 'htmltag',
])
const RTF_SYMBOLS = { emdash: '—', endash: '–', bullet: '•', lquote: '‘', rquote: '’', ldblquote: '“', rdblquote: '”', emspace: ' ', enspace: ' ', qmspace: ' ', zwj: '\u200D', zwnj: '\u200C', ltrmark: '\u200E', rtlmark: '\u200F' }

function rtfColor(colors, index) {
  const color = colors[index]
  return color ? `#${color.map((value) => Math.max(0, Math.min(255, value)).toString(16).padStart(2, '0')).join('')}` : null
}

/**
 * Converts RTF to HTML: paragraphs, bold/italic/underline/strike, sizes,
 * colours, fonts, alignment, indents, tables, page breaks, footnotes (at the
 * end) and PNG/JPEG pictures. Code pages follow \ansicpg and each font's
 * \fcharset.
 * @returns {{body: string, warnings: string[]}}
 */
function rtfToHtmlBody(input) {
  const source = toBuffer(input).toString('latin1')
  const fonts = new Map()
  const colors = []
  const blocks = []
  const notes = []
  const warnings = new Set()
  let defaultCodepage = 1252
  let defaultFont = null
  const initial = {
    skip: false, destination: null, uc: 1,
    b: false, i: false, ul: false, strike: false, sup: false, sub: false, fs: 24, f: null, cf: 0, hidden: false,
    align: '', li: 0, fi: 0, intbl: false,
  }
  let state = { ...initial }
  const stack = []
  let paragraph = []
  let pendingBytes = []
  let table = null
  let row = null
  let cell = null
  let pict = null
  let fontEntry = null
  let colorEntry = null
  let noteTarget = null
  let skipFallback = 0

  const codepageFor = () => {
    const font = fonts.get(state.f ?? defaultFont)
    return font?.codepage || defaultCodepage
  }
  const flushBytes = () => {
    if (!pendingBytes.length) return
    const label = CODEPAGE_LABELS[codepageFor()] || 'windows-1252'
    let text
    try { text = new TextDecoder(label).decode(Uint8Array.from(pendingBytes)) } catch { text = Buffer.from(pendingBytes).toString('latin1') }
    pendingBytes = []
    emitText(text)
  }
  const runStyle = () => {
    const css = []
    if (state.b) css.push('font-weight: 700')
    if (state.i) css.push('font-style: italic')
    const decorations = [state.ul ? 'underline' : '', state.strike ? 'line-through' : ''].filter(Boolean)
    if (decorations.length) css.push(`text-decoration: ${decorations.join(' ')}`)
    if (state.fs && state.fs !== 24) css.push(`font-size: ${round(state.fs / 2, 1)}pt`)
    const color = state.cf ? rtfColor(colors, state.cf) : null
    if (color) css.push(`color: ${color}`)
    const font = fonts.get(state.f ?? -1)
    if (font?.name && state.f !== defaultFont) css.push(`font-family: ${cssString(font.name)}, inherit`)
    if (state.sup) css.push('vertical-align: super; font-size: smaller')
    if (state.sub) css.push('vertical-align: sub; font-size: smaller')
    return css.join('; ')
  }
  const target = () => noteTarget || paragraph
  const emitText = (text) => {
    if (!text || state.hidden) return
    if (state.destination === 'fonttbl' && fontEntry) { fontEntry.name += text; return }
    if (state.skip) return
    const runs = target()
    const style = runStyle()
    const last = runs[runs.length - 1]
    if (last && last.style === style && last.text !== null) last.text += text
    else runs.push({ text, style })
  }
  const emitRaw = (html) => {
    if (state.skip) return
    target().push({ text: null, html })
  }
  const paragraphHtml = (runs, props) => {
    const content = runs.map((run) => run.text === null ? run.html : (run.style ? `<span style="${escapeHtml(run.style)}">${escapeHtml(run.text)}</span>` : escapeHtml(run.text))).join('')
    const css = []
    if (props.align) css.push(`text-align: ${props.align}`)
    if (props.li) css.push(`margin-left: ${round(props.li / 20)}pt`)
    if (props.fi) css.push(`text-indent: ${round(props.fi / 20)}pt`)
    return `<p${css.length ? ` style="${css.join('; ')}"` : ''}>${content || '&nbsp;'}</p>`
  }
  const endParagraph = () => {
    flushBytes()
    const html = paragraphHtml(paragraph, state)
    paragraph = []
    if (state.intbl) {
      table ||= { rows: [], widths: [] }
      row ||= []
      cell ||= []
      cell.push(html)
      return
    }
    closeTable()
    blocks.push(html)
  }
  const closeTable = () => {
    if (!table) return
    if (cell && cell.length) { row ||= []; row.push(cell) }
    if (row && row.length) table.rows.push(row)
    const widths = table.widths
    const columns = Math.max(...table.rows.map((item) => item.length), 1)
    const html = `<table class="rtf-table">${widths.length === columns ? `<colgroup>${widths.map((width) => `<col style="width: ${round(width)}pt">`).join('')}</colgroup>` : ''}${table.rows.map((cells) => `<tr>${cells.map((content) => `<td>${content.join('')}</td>`).join('')}</tr>`).join('')}</table>`
    blocks.push(html)
    table = null
    row = null
    cell = null
  }
  const endCell = () => {
    flushBytes()
    if (paragraph.length) cell = [...(cell || []), paragraphHtml(paragraph, state)]
    paragraph = []
    table ||= { rows: [], widths: [] }
    row ||= []
    row.push(cell || [])
    cell = null
  }
  const endRow = () => {
    flushBytes()
    if (paragraph.length) endCell()
    table ||= { rows: [], widths: [] }
    if (row) table.rows.push(row)
    row = null
    cell = null
  }

  const controlWord = (word, parameter) => {
    const value = parameter === null ? null : Number(parameter)
    if (state.destination === 'fonttbl') {
      if (word === 'f') { fontEntry = { id: value, name: '', codepage: null }; fonts.set(value, fontEntry); return }
      if (word === 'fcharset' && fontEntry) { fontEntry.codepage = RTF_CHARSET_CODEPAGES[value] || fontEntry.codepage; return }
      if (word === 'cpg' && fontEntry) { fontEntry.codepage = value; return }
    }
    if (state.destination === 'colortbl') {
      colorEntry ||= [0, 0, 0]
      if (word === 'red') colorEntry[0] = value
      if (word === 'green') colorEntry[1] = value
      if (word === 'blue') colorEntry[2] = value
      return
    }
    if (state.destination === 'pict' && pict) {
      if (word === 'pngblip') pict.type = 'image/png'
      else if (word === 'jpegblip') pict.type = 'image/jpeg'
      else if (word === 'picwgoal') pict.width = value / 20
      else if (word === 'pichgoal') pict.height = value / 20
      else if (word === 'picscalex') pict.scaleX = value / 100
      else if (word === 'picscaley') pict.scaleY = value / 100
      else if (word === 'bin' && value > 0) pict.binary = value
      return
    }
    switch (word) {
      case 'ansicpg': defaultCodepage = value || 1252; return
      case 'deff': defaultFont = value; return
      case 'uc': state.uc = Math.max(0, value ?? 1); return
      case 'u': {
        flushBytes()
        const code = value < 0 ? value + 65536 : value
        emitText(String.fromCharCode(code))
        skipFallback = state.uc
        return
      }
      case 'par': endParagraph(); return
      case 'pard': state.align = ''; state.li = 0; state.fi = 0; state.intbl = false; return
      case 'plain': Object.assign(state, { b: false, i: false, ul: false, strike: false, sup: false, sub: false, fs: 24, f: defaultFont, cf: 0, hidden: false }); return
      case 'b': flushBytes(); state.b = value !== 0; return
      case 'i': flushBytes(); state.i = value !== 0; return
      case 'ul': case 'uld': case 'uldb': case 'ulw': case 'ulth': flushBytes(); state.ul = value !== 0; return
      case 'ulnone': flushBytes(); state.ul = false; return
      case 'strike': case 'striked': flushBytes(); state.strike = value !== 0; return
      case 'super': flushBytes(); state.sup = true; state.sub = false; return
      case 'sub': flushBytes(); state.sub = true; state.sup = false; return
      case 'nosupersub': flushBytes(); state.sup = false; state.sub = false; return
      case 'v': flushBytes(); state.hidden = value !== 0; return
      case 'fs': flushBytes(); state.fs = value || 24; return
      case 'f': flushBytes(); state.f = value; return
      case 'cf': flushBytes(); state.cf = value || 0; return
      case 'qc': state.align = 'center'; return
      case 'qr': state.align = 'right'; return
      case 'qj': state.align = 'justify'; return
      case 'ql': state.align = ''; return
      case 'li': state.li = value || 0; return
      case 'fi': state.fi = value || 0; return
      case 'intbl': state.intbl = true; return
      case 'trowd': flushBytes(); table ||= { rows: [], widths: [] }; if (!table.rows.length) table.widths = []; table.lastX = 0; return
      case 'cellx': if (table && !table.rows.length) { table.widths.push(Math.max(1, (value - (table.lastX || 0)) / 20)); table.lastX = value } return
      case 'cell': case 'nestcell': endCell(); return
      case 'row': case 'nestrow': endRow(); return
      case 'line': flushBytes(); emitRaw('<br>'); return
      case 'tab': emitText('\t'); return
      case 'page': endParagraph(); closeTable(); blocks.push('<div class="page-break"></div>'); return
      case 'sect': endParagraph(); return
      case 'chftn': if (!noteTarget) emitRaw(`<sup>${notes.length + 1}</sup>`); return
      default:
        if (RTF_SYMBOLS[word]) emitText(RTF_SYMBOLS[word])
    }
  }

  const beginDestination = (word, ignorable) => {
    if (word === 'fonttbl' || word === 'colortbl') { state.destination = word; return true }
    if (word === 'pict') { state.destination = 'pict'; pict = { type: null, hex: [], width: 0, height: 0, scaleX: 1, scaleY: 1 }; return true }
    if (word === 'shppict' || word === 'fldrslt' || word === 'field') return true
    if (word === 'footnote') {
      flushBytes()
      state.destination = 'footnote'
      noteTarget = []
      notes.push(noteTarget)
      emitRawToParagraph(`<sup>${notes.length}</sup>`)
      return true
    }
    if (RTF_SKIPPED.has(word) || ignorable) { state.skip = true; state.destination = word; return true }
    return false
  }
  const emitRawToParagraph = (html) => { paragraph.push({ text: null, html }) }

  let index = 0
  let ignorableNext = false
  let groupStart = false
  while (index < source.length) {
    const character = source[index]
    if (character === '{') {
      flushBytes()
      stack.push({ state: { ...state }, pict, noteTarget })
      groupStart = true
      ignorableNext = false
      index += 1
      continue
    }
    if (character === '}') {
      flushBytes()
      if (state.destination === 'pict' && pict && !state.skip) {
        if (pict.type && pict.hex.length) {
          const data = Buffer.from(pict.hex.join(''), 'hex')
          const width = pict.width * pict.scaleX
          const height = pict.height * pict.scaleY
          const size = width > 0 && height > 0 ? ` style="width: ${round(width)}pt; height: ${round(height)}pt"` : ''
          stack.length && (state = stack[stack.length - 1].state)
          target().push({ text: null, html: `<img src="data:${pict.type};base64,${data.toString('base64')}"${size}>` })
        } else warnings.add('Some pictures in this RTF file could not be shown.')
      }
      if (state.destination === 'colortbl' && colorEntry) { colors.push(colorEntry); colorEntry = null }
      const previous = stack.pop()
      if (previous) {
        if (state.destination === 'footnote' && previous.noteTarget !== noteTarget) noteTarget = previous.noteTarget
        // Paragraph formatting is kept across groups only until \pard; restore the rest.
        state = previous.state
        pict = previous.pict
      }
      groupStart = false
      index += 1
      continue
    }
    if (character === '\\') {
      const next = source[index + 1]
      if (/[a-zA-Z]/.test(next || '')) {
        let end = index + 1
        while (end < source.length && /[a-zA-Z]/.test(source[end])) end += 1
        const word = source.slice(index + 1, end)
        let parameter = null
        const numberMatch = /^-?\d+/.exec(source.slice(end, end + 12))
        if (numberMatch) { parameter = numberMatch[0]; end += parameter.length }
        if (source[end] === ' ') end += 1
        index = end
        if (word === 'bin') {
          const length = Number(parameter) || 0
          if (state.destination === 'pict' && pict) pict.hex.push(Buffer.from(source.slice(index, index + length), 'latin1').toString('hex'))
          index += length
          continue
        }
        if (skipFallback > 0) { skipFallback -= 1; continue }
        if (groupStart) {
          groupStart = false
          if (beginDestination(word, ignorableNext)) { ignorableNext = false; continue }
        }
        ignorableNext = false
        if (state.skip) continue
        if (state.destination === 'colortbl' && word !== 'red' && word !== 'green' && word !== 'blue') continue
        controlWord(word, parameter)
        continue
      }
      if (next === '*') { ignorableNext = true; index += 2; continue }
      if (next === "'") {
        const byte = Number.parseInt(source.slice(index + 2, index + 4), 16)
        index += 4
        if (skipFallback > 0) { skipFallback -= 1; continue }
        if (state.skip || state.destination === 'pict') continue
        if (Number.isFinite(byte)) {
          if (state.destination === 'fonttbl') { if (fontEntry) fontEntry.name += String.fromCharCode(byte) } else pendingBytes.push(byte)
        }
        continue
      }
      index += 2
      groupStart = false
      if (skipFallback > 0) { skipFallback -= 1; continue }
      if (state.skip) continue
      if (next === '\\' || next === '{' || next === '}') emitText(next)
      else if (next === '~') emitText(' ')
      else if (next === '_') emitText('‑')
      else if (next === '\n' || next === '\r') endParagraph()
      continue
    }
    if (character === '\r' || character === '\n') { index += 1; continue }
    groupStart = false
    if (skipFallback > 0) { skipFallback -= 1; index += 1; continue }
    if (state.destination === 'pict' && pict) {
      if (/[0-9a-fA-F]/.test(character)) pict.hex.push(character)
      index += 1
      continue
    }
    if (state.destination === 'colortbl') {
      if (character === ';') { colors.push(colorEntry || null); colorEntry = null }
      index += 1
      continue
    }
    if (state.destination === 'fonttbl') {
      if (character === ';') { if (fontEntry) fontEntry.name = fontEntry.name.trim(); fontEntry = null } else if (fontEntry) fontEntry.name += character
      index += 1
      continue
    }
    if (state.skip) { index += 1; continue }
    const code = character.charCodeAt(0)
    if (code >= 0x80) pendingBytes.push(code)
    else { flushBytes(); emitText(character) }
    index += 1
  }
  flushBytes()
  if (paragraph.length) endParagraph()
  closeTable()
  // Decode font names written in the font's own code page.
  if (notes.length) {
    blocks.push(`<hr><ol class="notes">${notes.map((runs) => `<li>${runs.map((run) => run.text === null ? run.html : escapeHtml(run.text)).join('')}</li>`).join('')}</ol>`)
  }
  const defaultFace = fonts.get(defaultFont)?.name
  return { body: blocks.join('\n'), warnings: [...warnings], defaultFont: defaultFace || null }
}

function rtfToHtml(bytes, { title = 'Document', paper = defaultPaper() } = {}) {
  const { body, warnings, defaultFont } = rtfToHtmlBody(bytes)
  const css = `${pageRule({ ...paper, margins: [72, 72, 72, 72] })}${BASE_CSS}
body { font-family: ${defaultFont ? `${cssString(defaultFont)}, ` : ''}"Times New Roman", ${FONT_FALLBACKS}; font-size: 12pt; line-height: 1.2; }
p { margin: 0; white-space: pre-wrap; }
table.rtf-table { border-collapse: collapse; margin: 4pt 0; }
table.rtf-table td { border: 0.5pt solid #7f7f7f; padding: 1pt 5.4pt; }
ol.notes { font-size: 10pt; }`
  return { html: htmlDocument({ title, css, body: body || '<p>&nbsp;</p>' }), warnings }
}

// ---------------------------------------------------------------------------
// OpenDocument text (.odt)
// ---------------------------------------------------------------------------

function odfLengthToPoints(value) {
  return cssLengthToPoints(String(value || '').replace(/^([\d.]+)(?:\s*)$/, '$1pt'))
}

function odfTextProperties(element) {
  const properties = firstChild(element, 'style:text-properties')
  const paragraph = firstChild(element, 'style:paragraph-properties')
  const result = {}
  if (properties) {
    const attributes = properties.attributes
    if (attributes['fo:font-weight']) result.bold = /bold|[6-9]00/.test(attributes['fo:font-weight'])
    if (attributes['fo:font-style']) result.italic = /italic|oblique/.test(attributes['fo:font-style'])
    if (attributes['style:text-underline-style']) result.underline = attributes['style:text-underline-style'] !== 'none'
    if (attributes['style:text-line-through-style']) result.strike = attributes['style:text-line-through-style'] !== 'none'
    if (attributes['fo:font-size'] && /pt$/.test(attributes['fo:font-size'])) result.size = Number.parseFloat(attributes['fo:font-size'])
    if (/^#[0-9a-f]{6}$/i.test(attributes['fo:color'] || '')) result.color = attributes['fo:color']
    if (/^#[0-9a-f]{6}$/i.test(attributes['fo:background-color'] || '')) result.background = attributes['fo:background-color']
    if (attributes['style:font-name']) result.font = attributes['style:font-name']
    if (attributes['style:text-position'] && /^super/.test(attributes['style:text-position'])) result.position = 'super'
    if (attributes['style:text-position'] && /^sub/.test(attributes['style:text-position'])) result.position = 'sub'
  }
  if (paragraph) {
    const attributes = paragraph.attributes
    if (attributes['fo:text-align']) result.align = { start: 'left', end: 'right', left: 'left', right: 'right', center: 'center', justify: 'justify' }[attributes['fo:text-align']]
    const left = odfLengthToPoints(attributes['fo:margin-left'])
    if (left) result.marginLeft = left
    const indent = odfLengthToPoints(attributes['fo:text-indent'])
    if (indent) result.textIndent = indent
    const top = odfLengthToPoints(attributes['fo:margin-top'])
    if (top !== null) result.marginTop = top
    const bottom = odfLengthToPoints(attributes['fo:margin-bottom'])
    if (bottom !== null) result.marginBottom = bottom
    if (attributes['fo:break-before'] === 'page') result.breakBefore = true
    if (attributes['fo:break-after'] === 'page') result.breakAfter = true
  }
  return result
}

function odfStyles(...roots) {
  const raw = new Map()
  const fonts = new Map()
  for (const root of roots) {
    for (const face of findAll(root, 'style:font-face')) fonts.set(face.attributes['style:name'], String(face.attributes['svg:font-family'] || '').replace(/^['"]|['"]$/g, ''))
    for (const style of findAll(root, 'style:style')) raw.set(style.attributes['style:name'], style)
  }
  const cache = new Map()
  const resolve = (name, depth = 0) => {
    if (!name || depth > 10) return {}
    if (cache.has(name)) return cache.get(name)
    const style = raw.get(name)
    if (!style) return {}
    const parent = resolve(style.attributes['style:parent-style-name'], depth + 1)
    const own = odfTextProperties(style)
    const value = { ...parent, ...own }
    if (value.font && fonts.has(value.font)) value.fontFamily = fonts.get(value.font)
    // Page breaks belong to the style itself, not to styles based on it.
    if (!own.breakBefore) delete value.breakBefore
    if (!own.breakAfter) delete value.breakAfter
    cache.set(name, value)
    return value
  }
  return resolve
}

function odfInlineCss(style) {
  const css = []
  if (style.bold) css.push('font-weight: 700')
  if (style.italic) css.push('font-style: italic')
  const decorations = [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean)
  if (decorations.length) css.push(`text-decoration: ${decorations.join(' ')}`)
  if (style.size) css.push(`font-size: ${round(style.size)}pt`)
  if (style.color) css.push(`color: ${style.color}`)
  if (style.background) css.push(`background: ${style.background}`)
  if (style.fontFamily) css.push(`font-family: ${cssString(style.fontFamily)}, inherit`)
  if (style.position === 'super') css.push('vertical-align: super; font-size: smaller')
  if (style.position === 'sub') css.push('vertical-align: sub; font-size: smaller')
  return css.join('; ')
}

function odfBlockCss(style) {
  const css = []
  if (style.align && style.align !== 'left') css.push(`text-align: ${style.align}`)
  if (style.marginLeft) css.push(`margin-left: ${round(style.marginLeft)}pt`)
  if (style.textIndent) css.push(`text-indent: ${round(style.textIndent)}pt`)
  if (style.marginTop !== undefined) css.push(`margin-top: ${round(style.marginTop)}pt`)
  if (style.marginBottom !== undefined) css.push(`margin-bottom: ${round(style.marginBottom)}pt`)
  if (style.breakBefore) css.push('break-before: page')
  if (style.breakAfter) css.push('break-after: page')
  return css.join('; ')
}

/**
 * Converts OpenDocument text to HTML: headings, paragraphs, spans, lists,
 * tables (with spans), links, footnotes (at the end), page breaks and
 * pictures (written next to the HTML).
 * @returns {Promise<{html: string, files: Map<string, Buffer>, warnings: string[]}>}
 */
async function odtToHtml(bytes, { title = 'Document', paper = defaultPaper() } = {}) {
  const JSZip = require('jszip')
  let zip
  try {
    zip = await JSZip.loadAsync(toBuffer(bytes))
  } catch (error) {
    throw codedError('FILE_DAMAGED', "This OpenDocument file is damaged or isn't really an .odt file.", { technical: error?.message })
  }
  const contentXml = await zipText(zip, 'content.xml')
  if (!contentXml) throw codedError('FILE_DAMAGED', "This OpenDocument file is damaged or isn't really an .odt file.")
  const stylesXml = await zipText(zip, 'styles.xml')
  const content = parseXml(contentXml)
  const styles = stylesXml ? parseXml(stylesXml) : null
  const resolve = odfStyles(...(styles ? [styles] : []), content)
  const files = new Map()
  const warnings = new Set()
  const notes = []
  let pictureCount = 0

  // Page size from the default master page's layout.
  let page = null
  if (styles) {
    const master = findAll(styles, 'style:master-page')[0]
    const layoutName = master?.attributes['style:page-layout-name']
    const layout = findAll(styles, 'style:page-layout').find((item) => item.attributes['style:name'] === layoutName)
    const properties = firstChild(layout, 'style:page-layout-properties')
    if (properties) {
      const attributes = properties.attributes
      const width = odfLengthToPoints(attributes['fo:page-width'])
      const height = odfLengthToPoints(attributes['fo:page-height'])
      if (width > 72 && height > 72) {
        const side = (name) => odfLengthToPoints(attributes[name]) ?? 56.7
        page = { width, height, margins: [side('fo:margin-top'), side('fo:margin-right'), side('fo:margin-bottom'), side('fo:margin-left')] }
      }
    }
  }
  const defaultStyle = styles ? findAll(styles, 'style:default-style').find((item) => item.attributes['style:family'] === 'paragraph') : null
  const defaults = defaultStyle ? odfTextProperties(defaultStyle) : {}

  const inline = async (element) => {
    let html = ''
    for (const child of element.children || []) {
      if (typeof child === 'string') { html += escapeHtml(child); continue }
      html += await node(child)
    }
    return html
  }
  const node = async (element) => {
    switch (element.name) {
      case 'text:h': {
        const level = Math.max(1, Math.min(6, Number(element.attributes['text:outline-level']) || 1))
        const style = resolve(element.attributes['text:style-name'])
        const css = [odfBlockCss(style), odfInlineCss({ ...style, bold: style.bold === false ? false : undefined })].filter(Boolean).join('; ')
        return `<h${level}${css ? ` style="${escapeHtml(css)}"` : ''}>${await inline(element)}</h${level}>`
      }
      case 'text:p': {
        const style = resolve(element.attributes['text:style-name'])
        const css = [odfBlockCss(style), odfInlineCss(style)].filter(Boolean).join('; ')
        const content = await inline(element)
        return `<p${css ? ` style="${escapeHtml(css)}"` : ''}>${content || '&nbsp;'}</p>`
      }
      case 'text:span': {
        const css = odfInlineCss(resolve(element.attributes['text:style-name']))
        const content = await inline(element)
        return css ? `<span style="${escapeHtml(css)}">${content}</span>` : content
      }
      case 'text:a': {
        const href = String(element.attributes['xlink:href'] || '')
        const content = await inline(element)
        return /^(https?:|mailto:)/i.test(href) ? `<a href="${escapeHtml(href)}">${content}</a>` : content
      }
      case 'text:s': return ' '.repeat(Math.max(1, Math.min(200, Number(element.attributes['text:c']) || 1)))
      case 'text:tab': return '<span class="tab">\t</span>'
      case 'text:line-break': return '<br>'
      case 'text:soft-page-break': case 'text:bookmark': case 'text:bookmark-start': case 'text:bookmark-end':
      case 'office:annotation': case 'office:annotation-end': case 'text:sequence-decls': case 'text:variable-decls':
      case 'text:user-field-decls': case 'office:forms': case 'text:tracked-changes': case 'draw:shape':
        return ''
      case 'text:list': {
        const items = []
        for (const item of elementChildren(element)) {
          if (item.name === 'text:list-item' || item.name === 'text:list-header') items.push(`<li${item.name === 'text:list-header' ? ' class="header"' : ''}>${await inline(item)}</li>`)
        }
        return `<ul>${items.join('')}</ul>`
      }
      case 'text:note': {
        const citation = textOf(firstChild(element, 'text:note-citation')) || String(notes.length + 1)
        const body = firstChild(element, 'text:note-body')
        notes.push({ citation, html: body ? await inline(body) : '' })
        return `<sup>${escapeHtml(citation)}</sup>`
      }
      case 'table:table': return table(element)
      case 'draw:frame': return frame(element)
      case 'draw:a': case 'text:section': case 'text:index-body': case 'text:table-of-content': case 'text:alphabetical-index':
      case 'text:illustration-index': case 'text:table-index': case 'text:bibliography': case 'text:user-index':
      case 'office:text': case 'text:list-item': case 'text:list-header':
        return inline(element)
      case 'text:index-title': case 'text:table-of-content-source': case 'text:alphabetical-index-source':
        return element.name === 'text:index-title' ? inline(element) : ''
      default:
        return inline(element)
    }
  }
  const frame = async (element) => {
    const width = odfLengthToPoints(element.attributes['svg:width'])
    const height = odfLengthToPoints(element.attributes['svg:height'])
    const image = firstChild(element, 'draw:image')
    if (image) {
      const href = String(image.attributes['xlink:href'] || '').replace(/^\.\//, '')
      const file = href && !/^[a-z]+:/i.test(href) ? zip.file(href) : null
      const extension = extensionOf(href)
      if (file && ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'webp', 'svg'].includes(extension)) {
        const name = `picture-${++pictureCount}.${extension}`
        files.set(name, await file.async('nodebuffer'))
        const size = width > 0 && height > 0 ? ` style="width: ${round(width)}pt; height: ${round(height)}pt"` : ''
        return `<img src="${name}"${size}>`
      }
      warnings.add('Some pictures could not be shown.')
      return ''
    }
    const textBox = firstChild(element, 'draw:text-box')
    return textBox ? `<div class="text-box">${await inline(textBox)}</div>` : ''
  }
  const table = async (element) => {
    const columns = []
    const header = []
    const rows = []
    const collectRows = async (container, target) => {
      for (const child of elementChildren(container)) {
        if (child.name === 'table:table-header-rows') await collectRows(child, header)
        else if (child.name === 'table:table-rows' || child.name === 'table:table-row-group') await collectRows(child, target)
        else if (child.name === 'table:table-row') {
          const repeat = Math.max(1, Math.min(50, Number(child.attributes['table:number-rows-repeated']) || 1))
          const cells = []
          for (const cell of elementChildren(child)) {
            if (cell.name === 'table:covered-table-cell') continue
            if (cell.name !== 'table:table-cell') continue
            const span = Number(cell.attributes['table:number-columns-spanned']) || 1
            const rowSpan = Number(cell.attributes['table:number-rows-spanned']) || 1
            const cellRepeat = Math.max(1, Math.min(64, Number(cell.attributes['table:number-columns-repeated']) || 1))
            const html = await inline(cell)
            for (let copy = 0; copy < cellRepeat; copy += 1) {
              cells.push(`<td${span > 1 ? ` colspan="${span}"` : ''}${rowSpan > 1 ? ` rowspan="${rowSpan}"` : ''}>${html}</td>`)
            }
          }
          for (let copy = 0; copy < repeat; copy += 1) target.push(`<tr>${cells.join('')}</tr>`)
        } else if (child.name === 'table:table-column' || child.name === 'table:table-columns' || child.name === 'table:table-column-group') {
          for (const column of child.name === 'table:table-column' ? [child] : findAll(child, 'table:table-column')) {
            const repeat = Math.max(1, Math.min(64, Number(column.attributes['table:number-columns-repeated']) || 1))
            for (let copy = 0; copy < repeat; copy += 1) columns.push(column)
          }
        }
      }
    }
    await collectRows(element, rows)
    return `<table class="odt-table">${header.length ? `<thead>${header.join('')}</thead>` : ''}<tbody>${rows.join('')}</tbody></table>`
  }

  const text = findFirst(content, 'office:text')
  let body = text ? await inline(text) : ''
  if (notes.length) body += `<hr><ol class="notes">${notes.map((note) => `<li value="${Number.parseInt(note.citation, 10) || ''}">${note.html}</li>`).join('')}</ol>`
  const css = `${pageRule(page || { ...paper, margins: [56.7, 56.7, 56.7, 56.7] })}${BASE_CSS}
body { font-family: ${defaults.fontFamily ? `${cssString(defaults.fontFamily)}, ` : ''}"Liberation Serif", "Times New Roman", ${FONT_FALLBACKS}; font-size: ${round(defaults.size || 12)}pt; line-height: 1.2; }
p { margin: 0 0 4pt; white-space: pre-wrap; }
span.tab { display: inline-block; min-width: 2em; }
table.odt-table { border-collapse: collapse; margin: 4pt 0; }
table.odt-table td { border: 0.5pt solid #7f7f7f; padding: 1.5pt 4pt; }
table.odt-table td > p { margin: 0; }
ol.notes { font-size: 10pt; }`
  return { html: htmlDocument({ title, css, body: body || '<p>&nbsp;</p>' }), files, warnings: [...warnings] }
}

// ---------------------------------------------------------------------------
// Chromium: printing HTML and decoding images in hidden, isolated windows
// ---------------------------------------------------------------------------

const allowedRootsByContents = new Map()
let convertSession = null
let activeJobs = 0
const waitingJobs = []

function insideFolder(child, folder) {
  const relative = path.relative(folder, child)
  return relative === '' || (Boolean(relative) && !relative.startsWith('..') && !path.isAbsolute(relative))
}

function comparablePath(value) {
  const resolved = path.resolve(value)
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}

/**
 * The request policy of conversion windows: data: and blob: URLs, and local
 * file: URLs (never a network share) inside one of `roots`.
 */
function isAllowedRequest(url, roots) {
  let parsed
  try { parsed = new URL(url) } catch { return false }
  if (parsed.protocol === 'data:' || parsed.protocol === 'blob:') return true
  if (parsed.protocol !== 'file:' || parsed.host !== '') return false
  let filePath
  try { filePath = fileURLToPath(parsed) } catch { return false }
  if (/^[\\/]{2}/.test(filePath)) return false
  const target = comparablePath(filePath)
  return (roots || []).some((root) => typeof root === 'string' && root && insideFolder(target, comparablePath(root)))
}

function conversionSession(electron) {
  if (convertSession) return convertSession
  const session = electron.session.fromPartition(PARTITION, { cache: false })
  session.webRequest.onBeforeRequest((details, callback) => {
    const roots = allowedRootsByContents.get(details.webContentsId)
    callback({ cancel: !(roots && isAllowedRequest(details.url, roots)) })
  })
  session.setPermissionRequestHandler((_contents, _permission, callback) => callback(false))
  session.setPermissionCheckHandler(() => false)
  session.on('will-download', (event) => event.preventDefault())
  if (typeof session.setSpellCheckerEnabled === 'function') session.setSpellCheckerEnabled(false)
  convertSession = session
  return session
}

function acquireSlot() {
  if (activeJobs < MAX_CONCURRENT_JOBS) {
    activeJobs += 1
    return Promise.resolve()
  }
  return new Promise((resolve) => waitingJobs.push(resolve))
}

function releaseSlot() {
  const next = waitingJobs.shift()
  if (next) next()
  else activeJobs -= 1
}

async function removeJobDirectory(directory) {
  const resolved = path.resolve(directory)
  const relative = path.relative(path.resolve(os.tmpdir()), resolved)
  if (relative && !relative.startsWith('..') && !path.isAbsolute(relative) && path.basename(resolved).startsWith(JOB_PREFIX)) {
    await fs.rm(resolved, { recursive: true, force: true, maxRetries: 3 }).catch(() => {})
  }
}

function safeAssetName(name) {
  const clean = path.basename(String(name || '')).replace(/[^\w.-]/g, '_')
  if (!clean || clean === 'document.html' || clean.startsWith('.')) throw new Error('Invalid conversion asset name.')
  return clean
}

/**
 * Runs `task(window, directory)` in a hidden window whose session can only
 * load data: URLs and files inside the job folder (plus `fileRoots`).
 */
async function withConversionWindow({ javascript = false, fileRoots = [], timeoutMs = 60_000, timeoutMessage }, task) {
  const electron = require('electron')
  await electron.app.whenReady()
  await acquireSlot()
  let directory = null
  let window = null
  let contentsId = null
  let timer = null
  try {
    directory = await fs.mkdtemp(path.join(os.tmpdir(), JOB_PREFIX))
    window = new electron.BrowserWindow({
      show: false,
      width: 1000,
      height: 1300,
      skipTaskbar: true,
      focusable: false,
      paintWhenInitiallyHidden: true,
      backgroundColor: '#ffffff',
      webPreferences: {
        session: conversionSession(electron),
        sandbox: true,
        contextIsolation: true,
        nodeIntegration: false,
        nodeIntegrationInSubFrames: false,
        javascript,
        webSecurity: true,
        allowRunningInsecureContent: false,
        webgl: false,
        plugins: false,
        experimentalFeatures: false,
        spellcheck: false,
        devTools: false,
        disableDialogs: true,
        navigateOnDragDrop: false,
        backgroundThrottling: false,
        autoplayPolicy: 'document-user-activation-required',
      },
    })
    contentsId = window.webContents.id
    const roots = [directory, ...fileRoots.filter((root) => typeof root === 'string' && path.isAbsolute(root) && !/^[\\/]{2}/.test(root))]
    allowedRootsByContents.set(contentsId, roots)
    const contents = window.webContents
    contents.setWindowOpenHandler(() => ({ action: 'deny' }))
    contents.on('will-redirect', (event) => event.preventDefault())
    contents.on('will-attach-webview', (event) => event.preventDefault())
    contents.setAudioMuted(true)
    const job = Promise.resolve().then(() => task(window, directory))
    job.catch(() => {})
    const timeout = new Promise((_resolve, reject) => {
      timer = setTimeout(() => reject(codedError('CONVERSION_TIMEOUT', timeoutMessage || 'Making the PDF took too long, so Simple stopped. Your file is unchanged.')), timeoutMs)
    })
    let crashed
    const gone = new Promise((_resolve, reject) => {
      crashed = (_event, details) => reject(codedError('CONVERSION_FAILED', "Simple couldn't make the PDF. Your file is unchanged.", { technical: `renderer ${details?.reason || 'gone'}` }))
      contents.once('render-process-gone', crashed)
    })
    gone.catch(() => {})
    return await Promise.race([job, timeout, gone])
  } finally {
    clearTimeout(timer)
    if (contentsId !== null) allowedRootsByContents.delete(contentsId)
    if (window && !window.isDestroyed()) window.destroy()
    if (directory) await removeJobDirectory(directory)
    releaseSlot()
  }
}

/**
 * Prints an HTML document to PDF in a hidden, script-free window.
 * @param {string} html a complete document
 * @param {object} [options]
 * @param {Map<string, Buffer>|Array<[string, Buffer]>} [options.files] assets written next to the HTML
 * @param {string[]} [options.fileRoots] extra local folders the page may load files from
 * @param {number} [options.timeoutMs] defaults to 60 s plus 15 s per MB of HTML (at most 10 min)
 * @param {boolean} [options.outline=true] make bookmarks from headings
 * @returns {Promise<Buffer>} PDF bytes
 */
async function htmlToPdfBytes(html, options = {}) {
  const document = String(html || '')
  const size = Buffer.byteLength(document, 'utf8')
  if (size > MAX_HTML_BYTES) throw codedError('TOO_LARGE', 'This file is too large to turn into PDF pages in one piece.')
  const timeoutMs = options.timeoutMs || Math.min(600_000, 60_000 + Math.ceil(size / (1024 * 1024)) * 15_000)
  const outline = options.outline !== false
  return withConversionWindow({ javascript: false, fileRoots: options.fileRoots || [], timeoutMs }, async (window, directory) => {
    const htmlPath = path.join(directory, 'document.html')
    await fs.writeFile(htmlPath, document, 'utf8')
    for (const [name, data] of options.files || []) await fs.writeFile(path.join(directory, safeAssetName(name)), toBuffer(data))
    const allowedUrl = pathToFileURL(htmlPath).href
    window.webContents.on('will-navigate', (event, url) => { if (url !== allowedUrl) event.preventDefault() })
    await window.webContents.loadFile(htmlPath)
    // Let late layout (fonts, decoded images) settle before printing.
    await new Promise((resolve) => setTimeout(resolve, 150))
    const data = await window.webContents.printToPDF({
      printBackground: true,
      preferCSSPageSize: true,
      generateTaggedPDF: outline,
      generateDocumentOutline: outline,
    })
    const bytes = Buffer.from(data)
    if (bytes.length < 8 || bytes.subarray(0, 5).toString('latin1') !== '%PDF-') {
      throw codedError('CONVERSION_FAILED', "Simple couldn't make the PDF. Your file is unchanged.")
    }
    return bytes
  })
}

function webpIsLossy(bytes) {
  if (bytes.toString('latin1', 12, 16) === 'VP8 ') return true
  if (bytes.toString('latin1', 12, 16) === 'VP8X') return bytes.subarray(0, 4096).includes('VP8 ') && !bytes.subarray(0, 4096).includes('VP8L')
  return false
}

/**
 * Decodes any image Chromium understands (WebP, GIF, BMP, AVIF, ICO, SVG, and
 * PNG or JPEG) and re-encodes it as PNG, or as JPEG for opaque photos from a
 * lossy source. Animated images give their first frame.
 * @param {Buffer|Uint8Array} bytes
 * @param {{type?: string, maxPixels?: number, scale?: number}} [options]
 * @returns {Promise<{data: Buffer, mime: 'image/png'|'image/jpeg', width: number, height: number}>}
 */
async function rasterizeImage(bytes, options = {}) {
  const buffer = toBuffer(bytes)
  const type = options.type || sniffImageType(buffer)
  if (!type || !IMAGE_MIME[type] || type === 'tiff' || type === 'heic') throw unsupportedImageError(type)
  const lossy = type === 'jpeg' || type === 'avif' || (type === 'webp' && webpIsLossy(buffer))
  const maxPixels = Math.max(1, Math.min(60e6, Number(options.maxPixels) || 40e6))
  const scale = Number(options.scale) > 0 ? Number(options.scale) : 1
  const result = await withConversionWindow({ javascript: true, timeoutMs: 60_000, timeoutMessage: 'Reading this picture took too long, so Simple stopped.' }, async (window) => {
    window.webContents.on('will-navigate', (event) => event.preventDefault())
    await window.webContents.loadURL('data:text/html;charset=utf-8,<!doctype html><meta charset="utf-8"><title>image</title>')
    return window.webContents.executeJavaScript(`(async () => {
      const bytes = Uint8Array.from(atob(${JSON.stringify(buffer.toString('base64'))}), (character) => character.charCodeAt(0))
      const url = URL.createObjectURL(new Blob([bytes], { type: ${JSON.stringify(IMAGE_MIME[type])} }))
      try {
        const image = new Image()
        image.src = url
        await image.decode()
        let width = image.naturalWidth || 300
        let height = image.naturalHeight || 150
        const scale = Math.min(${scale}, 16384 / Math.max(width * ${scale}, height * ${scale}) * ${scale}, Math.sqrt(${maxPixels} / (width * height)))
        width = Math.max(1, Math.round(width * scale))
        height = Math.max(1, Math.round(height * scale))
        const canvas = new OffscreenCanvas(width, height)
        const context = canvas.getContext('2d')
        context.drawImage(image, 0, 0, width, height)
        const pixels = context.getImageData(0, 0, width, height).data
        let opaque = true
        for (let index = 3; index < pixels.length; index += 4) if (pixels[index] !== 255) { opaque = false; break }
        const type = opaque && ${lossy} ? 'image/jpeg' : 'image/png'
        const blob = await canvas.convertToBlob({ type, quality: 0.92 })
        return { width, height, type, data: new Uint8Array(await blob.arrayBuffer()) }
      } finally {
        URL.revokeObjectURL(url)
      }
    })()`, true)
  }).catch((error) => {
    if (error?.code === 'CONVERSION_TIMEOUT') throw error
    throw codedError('IMAGE_DAMAGED', "This picture is damaged or uses a format Simple can't read.", { technical: error?.message })
  })
  const data = toBuffer(result.data)
  return { data, mime: result.type, width: result.width, height: result.height }
}

/**
 * A picture chosen for Add image or Replace image, as PNG or JPEG bytes. PNG
 * and JPEG pass through untouched, TIFF gives its first page, and everything
 * else the browser engine decodes is converted (SVG drawn at up to 4x so
 * small vector art stays sharp).
 * @returns {Promise<{data: Buffer, mime: 'image/png'|'image/jpeg'}>}
 */
async function pickedImageData(bytes, name = 'image') {
  const buffer = toBuffer(bytes)
  const type = sniffImageType(buffer)
  try {
    if (type === 'png' || type === 'jpeg') return { data: buffer, mime: `image/${type}` }
    if (type === 'tiff') {
      const image = tiffFirstImage(buffer)
      return { data: image.data, mime: image.mime }
    }
    if (!type || type === 'heic') throw unsupportedImageError(type)
    let scale = 1
    if (type === 'svg') {
      const size = svgDimensions(decodeText(buffer).text)
      const longest = Math.max(size.width, size.height) / 0.75
      scale = Math.max(1, Math.min(4, 1600 / Math.max(1, longest)))
    }
    const raster = await rasterizeImage(buffer, { type, scale })
    return { data: raster.data, mime: raster.mime }
  } catch (error) {
    if (error && typeof error.code === 'string' && error.name === error.code && !String(error.message).startsWith(`${path.basename(name)}:`)) {
      error.message = namedMessage(name, error.message)
    }
    throw error
  }
}

function unsupportedImageError(type) {
  if (type === 'heic') return codedError('UNSUPPORTED_IMAGE', "HEIC photos can't be opened yet. In Photos, save or export the picture as JPEG, then open the JPEG.")
  return codedError('UNSUPPORTED_IMAGE', "This picture uses a format Simple can't read. Save it as PNG or JPEG, then try again.")
}

/** An SVG printed as vector graphics on a page of its own size. */
async function svgToPdfBytes(bytes, { title = 'Image' } = {}) {
  const buffer = toBuffer(bytes)
  const size = svgDimensions(decodeText(buffer).text)
  const largest = Math.max(size.width, size.height)
  const factor = largest > 14400 ? 14400 / largest : 1
  const width = Math.max(1, size.width * factor)
  const height = Math.max(1, size.height * factor)
  const html = htmlDocument({
    title,
    css: `@page { size: ${round(width, 3)}pt ${round(height, 3)}pt; margin: 0; } html, body { margin: 0; padding: 0; background: transparent; overflow: hidden; } img { display: block; width: ${round(width, 3)}pt; height: ${round(height, 3)}pt; }`,
    body: '<img src="image.svg" alt="">',
  })
  const pdf = await htmlToPdfBytes(html, { files: [['image.svg', buffer]], outline: false, timeoutMs: 60_000 })
  const { PDFDocument } = require('pdf-lib')
  const document = await PDFDocument.load(pdf)
  if (document.getPageCount() <= 1) return pdf
  // A rounding overflow can add an empty second page.
  for (let index = document.getPageCount() - 1; index > 0; index -= 1) document.removePage(index)
  return Buffer.from(await document.save())
}

// ---------------------------------------------------------------------------
// Office engine (optional)
// ---------------------------------------------------------------------------

async function officeEngine() {
  if (process.env.SIMPLE_FORCE_NO_OFFICE === '1') return null
  try {
    return await require('./office-converter.cjs').findOfficeConverter()
  } catch { return null }
}

async function officeEngineAvailable() {
  return Boolean(await officeEngine())
}

async function convertWithEngine(bytes, extension, kind) {
  const { convertOfficeBytes } = require('./office-converter.cjs')
  return convertOfficeBytes({
    bytes,
    inputExtension: extension,
    outputExtension: 'pdf',
    filter: kind === 'spreadsheet' ? 'calc_pdf_Export' : kind === 'presentation' ? 'impress_pdf_Export' : 'writer_pdf_Export',
  })
}

// ---------------------------------------------------------------------------
// The one entry point
// ---------------------------------------------------------------------------

function sourceBase(sourcePath) {
  if (!sourcePath || !path.isAbsolute(sourcePath) || /^[\\/]{2}/.test(sourcePath)) return null
  const directory = path.dirname(sourcePath)
  return { root: directory, href: pathToFileURL(directory + path.sep).href }
}

async function xlsxHasDrawings(bytes) {
  try {
    const zip = await require('jszip').loadAsync(toBuffer(bytes))
    return Object.keys(zip.files).some((name) => /^xl\/(drawings|charts)\//.test(name))
  } catch { return false }
}

/**
 * Word (.docx/.doc) to PDF. A local office engine is used when present (and
 * not disabled with SIMPLE_FORCE_NO_OFFICE=1); otherwise .docx is laid out
 * by Simple and .doc fails with NEEDS_OFFICE_ENGINE.
 */
async function convertWordToPdf(bytes, extension = 'docx', options = {}) {
  const ext = String(extension || 'docx').replace(/^\./, '').toLowerCase()
  const name = options.name || `document.${ext}`
  const legacy = ext === 'doc' || sniffType(bytes) === 'doc'
  return (await convertToPdf(bytes, { ...options, name: legacy && ext !== 'doc' ? `${baseTitle(name)}.doc` : name, kind: legacy ? 'word-legacy' : 'word' })).data
}

/**
 * Converts any importable file to PDF bytes.
 * @param {Buffer|Uint8Array} input
 * @param {{name?: string, sourcePath?: string, kind?: string}} [options]
 * @returns {Promise<{data: Buffer|Uint8Array, warnings: string[], kind: string, converted: boolean}>}
 */
async function convertToPdf(input, options = {}) {
  const bytes = toBuffer(input)
  const name = String(options.name || (options.sourcePath ? path.basename(options.sourcePath) : 'document'))
  if (!bytes.length) throw codedError('EMPTY_FILE', namedMessage(name, 'This file is empty.'))
  const detected = detectImportKind(bytes, name)
  const kind = options.kind || detected.kind
  if (kind === 'pdf') return { data: bytes, warnings: [], kind, converted: false }
  if (bytes.length > MAX_DOCUMENT_BYTES) throw codedError('TOO_LARGE', namedMessage(name, 'This file is too large to turn into PDF pages (the limit is 256 MB).'))
  const extension = detected.sniffed && ['docx', 'xlsx', 'ods', 'odt', 'pptx', 'odp', 'doc', 'xls', 'ppt'].includes(detected.sniffed) ? detected.sniffed : detected.extension
  const title = baseTitle(name)
  const paper = defaultPaper()
  const base = sourceBase(options.sourcePath)
  const fileRoots = base ? [base.root] : []
  if (!kind) throw unsupportedFormatError(name)
  try {
    if (kind === 'heic') throw unsupportedImageError('heic')
    if (kind === 'image') {
      const type = detected.sniffed && IMAGE_MIME[detected.sniffed] ? detected.sniffed : sniffImageType(bytes)
      if (!type) throw codedError('IMAGE_DAMAGED', namedMessage(name, "This picture is damaged or isn't really an image file."))
      if (type === 'svg') return { data: await svgToPdfBytes(bytes, { title }), warnings: [], kind, converted: true }
      if (type === 'png' || type === 'jpeg' || type === 'tiff') {
        return { data: await imageToPdfBytes(bytes, `.${type === 'jpeg' ? 'jpg' : type}`, title), warnings: [], kind, converted: true }
      }
      if (type === 'heic') throw unsupportedImageError('heic')
      const raster = await rasterizeImage(bytes, { type })
      return { data: await imageToPdfBytes(raster.data, raster.mime === 'image/png' ? '.png' : '.jpg', title), warnings: [], kind, converted: true }
    }
    const engineCandidate = ENGINE_INPUTS.has(extension) && ['word', 'word-legacy', 'spreadsheet', 'rtf', 'odt', 'presentation'].includes(kind)
    const engine = engineCandidate ? await officeEngine() : null
    if (engine) {
      try {
        return { data: await convertWithEngine(bytes, extension, kind), warnings: [], kind, converted: true }
      } catch (error) {
        if (kind === 'word-legacy' || kind === 'presentation') throw error
        // Fall back to Simple's own layout below.
      }
    }
    if (kind === 'word-legacy' || kind === 'presentation') throw needsOfficeEngineError(kind, name)
    if (kind === 'word') {
      const result = await docxToHtml(bytes, { title, paper })
      return { data: await htmlToPdfBytes(result.html, { files: result.files }), warnings: result.warnings, kind, converted: true }
    }
    if (kind === 'spreadsheet') {
      const hasDrawings = ['xlsx', 'xlsm', 'xltx', 'xltm'].includes(extension) ? await xlsxHasDrawings(bytes) : false
      const result = spreadsheetToHtml(bytes, { title, extension, paper, hasDrawings })
      return { data: await htmlToPdfBytes(result.html, { outline: result.sheets > 1 }), warnings: result.warnings, kind, converted: true }
    }
    if (kind === 'rtf') {
      const result = rtfToHtml(bytes, { title, paper })
      return { data: await htmlToPdfBytes(result.html), warnings: result.warnings, kind, converted: true }
    }
    if (kind === 'odt') {
      const result = await odtToHtml(bytes, { title, paper })
      return { data: await htmlToPdfBytes(result.html, { files: result.files }), warnings: result.warnings, kind, converted: true }
    }
    if (kind === 'markdown') {
      const { text } = decodeText(bytes)
      // markdown-it is fast, but a very large file is rendered in slices so the
      // main process keeps answering other windows.
      const html = text.length > 1_000_000 ? await markdownToHtmlChunked(text, { title, paper, base: base?.href }) : markdownToHtml(text, { title, paper, base: base?.href })
      return { data: await htmlToPdfBytes(html, { fileRoots }), warnings: [], kind, converted: true }
    }
    if (kind === 'html') {
      const { text } = decodeText(bytes, { declared: declaredHtmlCharset(bytes) })
      return { data: await htmlToPdfBytes(sanitizeHtmlDocument(text, { title, paper, base: base?.href }), { fileRoots }), warnings: [], kind, converted: true }
    }
    if (kind === 'text') {
      const { text } = decodeText(bytes)
      return { data: await htmlToPdfBytes(plainTextToHtml(text, { title, paper }), { outline: false }), warnings: [], kind, converted: true }
    }
  } catch (error) {
    if (error && typeof error.code === 'string' && error.name === error.code) {
      if (!String(error.message).startsWith(`${path.basename(name)}:`) && ['FILE_DAMAGED', 'IMAGE_DAMAGED', 'UNSUPPORTED_IMAGE', 'CONVERSION_TIMEOUT', 'TOO_LARGE'].includes(error.code)) {
        error.message = namedMessage(name, error.message)
      }
      throw error
    }
    throw conversionFailedError(name, error)
  }
  throw unsupportedFormatError(name)
}

module.exports = {
  CONTENT_SECURITY_POLICY,
  IMPORT_FORMATS,
  PARTITION,
  convertToPdf,
  convertWordToPdf,
  decodeText,
  defaultPaper,
  detectImportKind,
  docxDrawingSizes,
  docxToHtml,
  htmlToPdfBytes,
  importDialogFilters,
  importExtensions,
  isAllowedRequest,
  isImportableName,
  markdownToHtml,
  needsOfficeEngineError,
  officeEngineAvailable,
  odtToHtml,
  parseXml,
  pickedImageData,
  plainTextToHtml,
  rasterizeImage,
  rtfToHtml,
  sanitizeHtmlDocument,
  sniffType,
  spreadsheetToHtml,
  svgToPdfBytes,
  unsupportedFormatError,
}
