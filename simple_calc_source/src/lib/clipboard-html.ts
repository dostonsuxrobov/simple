/*
 * Rich clipboard interoperability.
 *
 * Copy: `serializeSelectionToClipboard` produces the two flavours spreadsheet apps exchange —
 * `text/plain` (TSV of the *displayed* values, like Excel) and `text/html` (a table with inline
 * styles, Excel `mso-*` / `x:*` hints and Google Sheets `data-sheets-*` attributes).
 *
 * Paste: `parseClipboardHtml` reads tables written by Excel, Google Sheets, LibreOffice, Word,
 * simple_calc itself and ordinary web pages; `parsePastedText` reads TSV / single-column text.
 * Both produce a matrix of `CellData`. Parsing never executes anything: scripts, event handlers,
 * images and unsafe links are dropped, and the HTML is only ever walked as inert data.
 */
import type {
  CellAlignment,
  CellBorder,
  CellBorderSide,
  CellData,
  CellFill,
  CellFont,
  CellScalar,
  CellStyle,
} from '../spreadsheet-types'

export const CLIPBOARD_MAX_CELLS = 100_000
export const CLIPBOARD_TOO_LARGE_MESSAGE = 'That clipboard content is too large to paste at once.'
const SHEET_ROWS = 1_048_576
const SHEET_COLS = 16_384
const MAX_TEXT_LENGTH = 32_767
const MAX_FORMULA_LENGTH = 8_192
const DEFAULT_COLUMN_PX = 100
const DEFAULT_ROW_PX = 20

export class ClipboardTooLargeError extends RangeError {
  constructor(message = CLIPBOARD_TOO_LARGE_MESSAGE) {
    super(message)
    this.name = 'ClipboardTooLargeError'
  }
}

export function isClipboardTooLarge(error: unknown): error is ClipboardTooLargeError {
  return error instanceof ClipboardTooLargeError
}

export interface ClipboardCoord { row: number; col: number }
export interface ClipboardBounds { top: number; left: number; bottom: number; right: number }

export type ClipboardSource = 'simple-calc' | 'excel' | 'google-sheets' | 'libreoffice' | 'word' | 'html' | 'text'

export interface ParsedClipboard {
  /** Rectangular matrix; covered merge cells and gaps are `{}`. */
  cells: CellData[][]
  /** Merged areas as A1 ranges relative to `cells[0][0]` (A1 = first cell). */
  merges: string[]
  /** Column widths in CSS pixels when the source declared them. */
  columnWidths?: Array<number | undefined>
  /** Row heights in CSS pixels when the source declared them. */
  rowHeights?: Array<number | undefined>
  source: ClipboardSource
  /**
   * Absolute sheet coordinate the formulas in `cells[0][0]` are written against. For
   * simple_calc copies it is the source origin; for other apps it is `options.origin`
   * (the paste destination), so shifting from here to the final cell is always correct.
   */
  formulaOrigin: ClipboardCoord
  /** Worksheet the content was copied from, when the source says so (simple_calc). */
  sourceSheetName?: string
}

export interface ClipboardParseOptions {
  /** Top-left destination cell; R1C1 formulas (Google Sheets) are resolved against it. */
  origin?: ClipboardCoord
  /** Hard cap; larger content throws `ClipboardTooLargeError`. Default 100 000. */
  maxCells?: number
  /** Font the grid uses by default; matching font names/sizes are not stored per cell. */
  defaultFont?: { name?: string; size?: number }
  /** `auto` uses DOMParser when available (renderer) and the built-in tokenizer otherwise. */
  parser?: 'auto' | 'dom' | 'builtin'
  decimalSeparator?: '.' | ','
  /** Read ambiguous numeric dates as d/m/y. */
  dayFirst?: boolean
  /** Plain text only: field delimiter (default tab). */
  delimiter?: string
  /** Plain text only: treat "=..." as a formula (default true, like typing it). */
  formulas?: boolean
}

// ---------------------------------------------------------------------------------------------
// Addresses
// ---------------------------------------------------------------------------------------------

export function columnLabel(col: number): string {
  let n = Math.floor(col) + 1
  let label = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    label = String.fromCharCode(65 + rem) + label
    n = Math.floor((n - 1) / 26)
  }
  return label
}

function columnNumber(label: string): number {
  let n = 0
  for (let index = 0; index < label.length; index += 1) n = n * 26 + (label.charCodeAt(index) & ~32) - 64
  return n - 1
}

export function cellAddress(row: number, col: number): string {
  return `${columnLabel(col)}${row + 1}`
}

export function boundsToRange(bounds: ClipboardBounds): string {
  const start = cellAddress(bounds.top, bounds.left)
  const end = cellAddress(bounds.bottom, bounds.right)
  return start === end ? start : `${start}:${end}`
}

export function rangeToBounds(range: string): ClipboardBounds | null {
  const match = /^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})(?::\$?([A-Za-z]{1,3})\$?(\d{1,7}))?\s*$/.exec(range)
  if (!match) return null
  const r1 = Number(match[2]) - 1
  const c1 = columnNumber(match[1])
  const r2 = match[4] ? Number(match[4]) - 1 : r1
  const c2 = match[3] ? columnNumber(match[3]) : c1
  if (r1 < 0 || r2 < 0) return null
  return { top: Math.min(r1, r2), bottom: Math.max(r1, r2), left: Math.min(c1, c2), right: Math.max(c1, c2) }
}

const hasOwn = (object: object, key: string) => Object.prototype.hasOwnProperty.call(object, key)

function parseJson(raw: string | undefined): unknown {
  if (!raw || raw.length > 1_000_000) return undefined
  try {
    return JSON.parse(raw)
  } catch {
    return undefined
  }
}

// ---------------------------------------------------------------------------------------------
// Tolerant HTML parser (used in Node and as the renderer fallback)
// ---------------------------------------------------------------------------------------------

export interface HtmlElement {
  tag: string
  attrs: Record<string, string>
  children: HtmlNode[]
}
export type HtmlNode = HtmlElement | string

const VOID_TAGS = new Set(['area', 'base', 'basefont', 'bgsound', 'br', 'col', 'embed', 'frame', 'hr', 'img', 'input', 'keygen', 'link', 'meta', 'param', 'source', 'track', 'wbr'])
const RAW_TEXT_TAGS = new Set(['script', 'style', 'textarea', 'title', 'xmp', 'iframe', 'noembed', 'noframes', 'noscript'])
const DISCARDED_TAGS = new Set(['script', 'iframe', 'noembed', 'noframes', 'noscript', 'template', 'object', 'embed', 'applet'])
const P_CLOSERS = new Set(['address', 'article', 'aside', 'blockquote', 'center', 'details', 'dialog', 'dir', 'div', 'dl', 'fieldset', 'figcaption', 'figure', 'footer', 'form', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'header', 'hgroup', 'hr', 'main', 'menu', 'nav', 'ol', 'p', 'pre', 'section', 'summary', 'ul'])
const BUTTON_SCOPE = new Set(['td', 'th', 'table', 'caption', 'button', 'html', 'applet', 'marquee', 'object'])
const TABLE_SECTIONS = new Set(['tbody', 'thead', 'tfoot'])
const TABLE_STRUCTURE = new Set(['td', 'th', 'tr', 'tbody', 'thead', 'tfoot', 'colgroup', 'caption'])
const INLINE_END_BOUNDARY = new Set(['td', 'th', 'table', 'caption'])
const HEAD_TAGS = new Set(['head', 'meta', 'link', 'style', 'title', 'base', 'script', 'noscript', 'template', 'basefont', 'bgsound', 'object'])

const START_TAG = /<([A-Za-z][^\s/>]*)/y
const END_TAG = /<\/([A-Za-z][^\s/>]*)[^>]*>/y
const ATTRIBUTE_NAME = /[^\s"'>/=]+/y
const UNQUOTED_VALUE = /[^\s>]+/y

const LATIN1_ENTITY_NAMES = 'nbsp iexcl cent pound curren yen brvbar sect uml copy ordf laquo not shy reg macr deg plusmn sup2 sup3 acute micro para middot cedil sup1 ordm raquo frac14 frac12 frac34 iquest Agrave Aacute Acirc Atilde Auml Aring AElig Ccedil Egrave Eacute Ecirc Euml Igrave Iacute Icirc Iuml ETH Ntilde Ograve Oacute Ocirc Otilde Ouml times Oslash Ugrave Uacute Ucirc Uuml Yacute THORN szlig agrave aacute acirc atilde auml aring aelig ccedil egrave eacute ecirc euml igrave iacute icirc iuml eth ntilde ograve oacute ocirc otilde ouml divide oslash ugrave uacute ucirc uuml yacute thorn yuml'
const NAMED_ENTITIES: Record<string, string> = Object.create(null)
LATIN1_ENTITY_NAMES.split(' ').forEach((name, index) => { NAMED_ENTITIES[name] = String.fromCharCode(160 + index) })
Object.assign(NAMED_ENTITIES, {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", euro: '\u20AC', ndash: '\u2013', mdash: '\u2014',
  lsquo: '\u2018', rsquo: '\u2019', sbquo: '\u201A', ldquo: '\u201C', rdquo: '\u201D', bdquo: '\u201E',
  bull: '\u2022', hellip: '\u2026', trade: '\u2122', ensp: '\u2002', emsp: '\u2003', thinsp: '\u2009',
  zwnj: '\u200C', zwj: '\u200D', lrm: '\u200E', rlm: '\u200F', permil: '\u2030', prime: '\u2032',
  Prime: '\u2033', minus: '\u2212', dagger: '\u2020', Dagger: '\u2021', lsaquo: '\u2039', rsaquo: '\u203A',
  larr: '\u2190', rarr: '\u2192', uarr: '\u2191', darr: '\u2193', harr: '\u2194', le: '\u2264', ge: '\u2265',
  ne: '\u2260', asymp: '\u2248', infin: '\u221E', sum: '\u2211', radic: '\u221A', OElig: '\u0152',
  oelig: '\u0153', Scaron: '\u0160', scaron: '\u0161', Yuml: '\u0178', fnof: '\u0192', circ: '\u02C6',
  tilde: '\u02DC', alpha: '\u03B1', beta: '\u03B2', gamma: '\u03B3', delta: '\u03B4', pi: '\u03C0',
  sigma: '\u03C3', mu: '\u03BC', omega: '\u03C9', lambda: '\u03BB', theta: '\u03B8', Delta: '\u0394',
  Omega: '\u03A9', Sigma: '\u03A3', check: '\u2713',
})
const LEGACY_ENTITIES = new Set(['amp', 'lt', 'gt', 'quot', 'nbsp', 'copy', 'reg'])
const WINDOWS_1252 = [0x20AC, 0x81, 0x201A, 0x0192, 0x201E, 0x2026, 0x2020, 0x2021, 0x02C6, 0x2030, 0x0160, 0x2039, 0x0152, 0x8D, 0x017D, 0x8F, 0x90, 0x2018, 0x2019, 0x201C, 0x201D, 0x2022, 0x2013, 0x2014, 0x02DC, 0x2122, 0x0161, 0x203A, 0x0153, 0x9D, 0x017E, 0x0178]
const ENTITY_AT = /&(?:#(\d{1,8})|#[xX]([0-9a-fA-F]{1,7})|([A-Za-z][A-Za-z0-9]{0,31}))(;?)/y

function codePointEntity(code: number) {
  let value = code
  if (value >= 0x80 && value <= 0x9f) value = WINDOWS_1252[value - 0x80]
  if (!value || value > 0x10ffff || (value >= 0xd800 && value <= 0xdfff)) return '\uFFFD'
  return String.fromCodePoint(value)
}

/** Decodes HTML character references in one pass (so "&amp;lt;" stays "&lt;"). */
export function decodeEntities(text: string, attribute = false): string {
  let amp = text.indexOf('&')
  if (amp === -1) return text
  let out = ''
  let last = 0
  while (amp !== -1) {
    ENTITY_AT.lastIndex = amp
    const match = ENTITY_AT.exec(text)
    let decoded: string | undefined
    let end = amp + 1
    if (match) {
      const [whole, dec, hex, name, semicolon] = match
      if (dec || hex) {
        decoded = codePointEntity(dec ? Number.parseInt(dec, 10) : Number.parseInt(hex, 16))
        end = amp + whole.length
      } else {
        const named = NAMED_ENTITIES[name]
        const next = text[amp + whole.length] ?? ''
        if (named !== undefined && (semicolon || (LEGACY_ENTITIES.has(name) && !(attribute && (next === '=' || /[A-Za-z0-9]/.test(next)))))) {
          decoded = named
          end = amp + whole.length
        }
      }
    }
    if (decoded !== undefined) {
      out += text.slice(last, amp) + decoded
      last = end
    }
    amp = text.indexOf('&', end)
  }
  return out + text.slice(last)
}

function isSpaceCode(code: number) {
  return code === 32 || code === 9 || code === 10 || code === 13 || code === 12
}

/**
 * Parses HTML into a light element tree with the HTML5 implied-end-tag rules that matter for
 * tables (unclosed td/tr/p, implicit rows, void elements, raw-text elements). `xml` mode honours
 * self-closing tags and never implies end tags.
 */
export function parseHtml(html: string, options: { xml?: boolean } = {}): HtmlElement {
  const xml = Boolean(options.xml)
  const root: HtmlElement = { tag: '#document', attrs: Object.create(null), children: [] }
  const stack: HtmlElement[] = [root]
  const length = html.length
  let index = 0

  const current = () => stack[stack.length - 1]
  const appendText = (text: string) => {
    if (!text) return
    const children = current().children
    const last = children.length - 1
    if (last >= 0 && typeof children[last] === 'string') children[last] = (children[last] as string) + text
    else children.push(text)
  }
  const findOpen = (match: (tag: string) => boolean, boundary: (tag: string) => boolean) => {
    for (let k = stack.length - 1; k > 0; k -= 1) {
      const tag = stack[k].tag
      if (match(tag)) return k
      if (boundary(tag)) return -1
    }
    return -1
  }
  const closeAt = (k: number) => { if (k > 0) stack.length = k }

  let headOpen = false
  const implicitClose = (tag: string) => {
    if (headOpen && !HEAD_TAGS.has(tag)) {
      closeAt(findOpen((t) => t === 'head', () => false))
      headOpen = false
    }
    if (tag === 'td' || tag === 'th') {
      closeAt(findOpen((t) => t === 'td' || t === 'th', (t) => t === 'tr' || t === 'table' || TABLE_SECTIONS.has(t)))
      if (current().tag === 'colgroup' || current().tag === 'caption') stack.pop()
      const parent = current().tag
      if (parent === 'table' || TABLE_SECTIONS.has(parent)) open('tr', Object.create(null), false)
    } else if (tag === 'tr') {
      closeAt(findOpen((t) => t === 'tr' || t === 'colgroup' || t === 'caption', (t) => t === 'table' || TABLE_SECTIONS.has(t)))
    } else if (TABLE_SECTIONS.has(tag) || tag === 'colgroup' || tag === 'caption') {
      closeAt(findOpen((t) => TABLE_SECTIONS.has(t) || t === 'colgroup' || t === 'caption', (t) => t === 'table'))
    } else if (P_CLOSERS.has(tag)) {
      closeAt(findOpen((t) => t === 'p', (t) => BUTTON_SCOPE.has(t)))
    } else if (tag === 'li') {
      closeAt(findOpen((t) => t === 'li', (t) => t === 'ul' || t === 'ol' || BUTTON_SCOPE.has(t)))
    } else if (tag === 'dt' || tag === 'dd') {
      closeAt(findOpen((t) => t === 'dt' || t === 'dd', (t) => t === 'dl' || BUTTON_SCOPE.has(t)))
    } else if (tag === 'option' && current().tag === 'option') {
      stack.pop()
    }
  }

  function open(tag: string, attrs: Record<string, string>, selfClosing: boolean): HtmlElement {
    if (!xml) implicitClose(tag)
    const element: HtmlElement = { tag, attrs, children: [] }
    current().children.push(element)
    const isVoid = xml ? selfClosing : VOID_TAGS.has(tag) || (selfClosing && tag.includes(':'))
    if (!isVoid) stack.push(element)
    if (!xml && tag === 'head') headOpen = true
    return element
  }

  const close = (tag: string) => {
    if (xml) { closeAt(findOpen((t) => t === tag, () => false)); return }
    if (VOID_TAGS.has(tag)) { if (tag === 'br') open('br', Object.create(null), true); return }
    if (tag === 'head') {
      closeAt(findOpen((t) => t === 'head', () => false))
      headOpen = false
      return
    }
    if (tag === 'html' || tag === 'body') return
    let k: number
    if (TABLE_STRUCTURE.has(tag)) k = findOpen((t) => t === tag, (t) => t === 'table')
    else if (tag === 'table') k = findOpen((t) => t === 'table', () => false)
    else k = findOpen((t) => t === tag, (t) => INLINE_END_BOUNDARY.has(t))
    closeAt(k)
  }

  while (index < length) {
    const lt = html.indexOf('<', index)
    if (lt === -1) { appendText(decodeEntities(html.slice(index))); break }
    if (lt > index) appendText(decodeEntities(html.slice(index, lt)))
    index = lt
    const next = html.charCodeAt(lt + 1)
    if (next === 33) { // "!"
      if (html.startsWith('<!--', lt)) {
        const end = html.indexOf('-->', lt + 4)
        index = end === -1 ? length : end + 3
        continue
      }
      if (html.startsWith('<![CDATA[', lt)) {
        const end = html.indexOf(']]>', lt + 9)
        if (xml) appendText(html.slice(lt + 9, end === -1 ? length : end))
        index = end === -1 ? length : end + 3
        continue
      }
      const end = html.indexOf('>', lt + 2)
      index = end === -1 ? length : end + 1
      continue
    }
    if (next === 63) { // "?"
      const end = html.indexOf('>', lt + 2)
      index = end === -1 ? length : end + 1
      continue
    }
    if (next === 47) { // "/"
      END_TAG.lastIndex = lt
      const match = END_TAG.exec(html)
      if (match) {
        close(match[1].toLowerCase())
        index = END_TAG.lastIndex
      } else {
        const end = html.indexOf('>', lt + 2)
        index = end === -1 ? length : end + 1
      }
      continue
    }
    if ((next >= 65 && next <= 90) || (next >= 97 && next <= 122)) {
      START_TAG.lastIndex = lt
      const nameMatch = START_TAG.exec(html) as RegExpExecArray
      const tag = nameMatch[1].toLowerCase()
      let position = START_TAG.lastIndex
      const attrs: Record<string, string> = Object.create(null)
      let selfClosing = false
      let closed = false
      while (position < length) {
        const code = html.charCodeAt(position)
        if (code === 62) { position += 1; closed = true; break }
        if (isSpaceCode(code)) { position += 1; continue }
        if (code === 47) {
          if (html.charCodeAt(position + 1) === 62) { selfClosing = true; position += 2; closed = true; break }
          position += 1
          continue
        }
        ATTRIBUTE_NAME.lastIndex = position
        const attributeMatch = ATTRIBUTE_NAME.exec(html)
        if (!attributeMatch) { position += 1; continue }
        const name = attributeMatch[0].toLowerCase()
        position = ATTRIBUTE_NAME.lastIndex
        while (position < length && isSpaceCode(html.charCodeAt(position))) position += 1
        let value = ''
        if (html.charCodeAt(position) === 61) {
          position += 1
          while (position < length && isSpaceCode(html.charCodeAt(position))) position += 1
          const quote = html.charCodeAt(position)
          if (quote === 34 || quote === 39) {
            const end = html.indexOf(quote === 34 ? '"' : "'", position + 1)
            value = decodeEntities(html.slice(position + 1, end === -1 ? length : end), true)
            position = end === -1 ? length : end + 1
          } else {
            UNQUOTED_VALUE.lastIndex = position
            const valueMatch = UNQUOTED_VALUE.exec(html)
            if (valueMatch) {
              value = decodeEntities(valueMatch[0], true)
              position = UNQUOTED_VALUE.lastIndex
            }
          }
        }
        if (!hasOwn(attrs, name)) attrs[name] = value
      }
      if (!closed) break
      index = position
      if (!xml && RAW_TEXT_TAGS.has(tag)) {
        const closePattern = new RegExp(`</${tag}[\\s/>]`, 'ig')
        closePattern.lastIndex = index
        const closeMatch = closePattern.exec(html)
        const end = closeMatch ? closeMatch.index : length
        if (!DISCARDED_TAGS.has(tag)) {
          const raw = html.slice(index, end)
          const element = open(tag, attrs, false)
          element.children.push(tag === 'style' ? raw : decodeEntities(raw))
          stack.pop()
        }
        if (!closeMatch) { index = length; break }
        const gt = html.indexOf('>', end)
        index = gt === -1 ? length : gt + 1
        continue
      }
      open(tag, attrs, selfClosing)
      continue
    }
    appendText('<')
    index = lt + 1
  }
  return root
}

function convertDomElement(element: Element): HtmlElement {
  const tag = (element.localName || element.nodeName || '').toLowerCase()
  const attrs: Record<string, string> = Object.create(null)
  const attributes = element.attributes
  for (let index = 0; index < attributes.length; index += 1) {
    const attribute = attributes[index]
    const name = attribute.name.toLowerCase()
    if (!hasOwn(attrs, name)) attrs[name] = attribute.value
  }
  const children: HtmlNode[] = []
  for (let child = element.firstChild; child; child = child.nextSibling) {
    if (child.nodeType === 3 || child.nodeType === 4) children.push(child.nodeValue ?? '')
    else if (child.nodeType === 1) {
      const childTag = ((child as Element).localName || '').toLowerCase()
      if (DISCARDED_TAGS.has(childTag)) continue
      children.push(convertDomElement(child as Element))
    }
  }
  return { tag, attrs, children }
}

function parseHtmlWithDom(html: string): HtmlElement | null {
  const Parser = (globalThis as { DOMParser?: typeof DOMParser }).DOMParser
  if (typeof Parser !== 'function') return null
  try {
    // DOMParser documents are inert: scripts never run and images never load.
    const documentNode = new Parser().parseFromString(html, 'text/html')
    const root: HtmlElement = { tag: '#document', attrs: Object.create(null), children: [] }
    if (documentNode.documentElement) root.children.push(convertDomElement(documentNode.documentElement))
    return root
  } catch {
    return null
  }
}

function elementChildren(element: HtmlElement): HtmlElement[] {
  const result: HtmlElement[] = []
  for (const child of element.children) if (typeof child !== 'string') result.push(child)
  return result
}

function walkElements(element: HtmlElement, visit: (element: HtmlElement) => boolean | void) {
  const pending: HtmlElement[] = [element]
  while (pending.length) {
    const next = pending.pop() as HtmlElement
    if (visit(next) === false) continue
    for (let index = next.children.length - 1; index >= 0; index -= 1) {
      const child = next.children[index]
      if (typeof child !== 'string') pending.push(child)
    }
  }
}

// ---------------------------------------------------------------------------------------------
// CSS
// ---------------------------------------------------------------------------------------------

type Declarations = Record<string, string>
const SIDES = ['top', 'right', 'bottom', 'left'] as const
type Side = typeof SIDES[number]

/** Splits a declaration block into [property, value] pairs, respecting quotes and escapes. */
function declarationPairs(text: string): Array<[string, string]> {
  const pairs: Array<[string, string]> = []
  let start = 0
  let quote = ''
  let depth = 0
  const flush = (end: number) => {
    const part = text.slice(start, end)
    start = end + 1
    const colon = part.indexOf(':')
    if (colon <= 0) return
    const property = part.slice(0, colon).trim().toLowerCase()
    const value = part.slice(colon + 1).trim().replace(/\s*!important\s*$/i, '')
    if (property && value && property.length < 64) pairs.push([property, value])
  }
  for (let index = 0; index < text.length; index += 1) {
    const ch = text[index]
    if (ch === '\\') { index += 1; continue }
    if (quote) { if (ch === quote) quote = ''; continue }
    if (ch === '"' || ch === "'") { quote = ch; continue }
    if (ch === '(') depth += 1
    else if (ch === ')') depth = Math.max(0, depth - 1)
    else if (ch === ';' && !depth) flush(index)
  }
  flush(text.length)
  return pairs
}

function splitCssValues(value: string): string[] {
  const parts: string[] = []
  let depth = 0
  let quote = ''
  let current = ''
  for (const ch of value) {
    if (quote) { current += ch; if (ch === quote) quote = ''; continue }
    if (ch === '"' || ch === "'") { quote = ch; current += ch; continue }
    if (ch === '(') depth += 1
    if (ch === ')') depth = Math.max(0, depth - 1)
    if (!depth && (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r')) {
      if (current) parts.push(current)
      current = ''
      continue
    }
    current += ch
  }
  if (current) parts.push(current)
  return parts
}

const BORDER_STYLE_KEYWORDS = new Set(['none', 'hidden', 'solid', 'dotted', 'dashed', 'double', 'groove', 'ridge', 'inset', 'outset', 'hairline', 'dot-dash', 'dot-dot-dash', 'slanted-dot-dash'])

function isBorderWidthToken(token: string) {
  return token === 'thin' || token === 'medium' || token === 'thick' || /^[+-]?(?:\d+\.?\d*|\.\d+)(?:px|pt|em|rem|in|cm|mm|pc)?$/.test(token)
}

function parseBorderShorthand(value: string) {
  let style = ''
  let width = ''
  let color = ''
  for (const raw of splitCssValues(value)) {
    const token = raw.toLowerCase()
    if (BORDER_STYLE_KEYWORDS.has(token)) style = token
    else if (isBorderWidthToken(token)) width = token
    else color = raw
  }
  return { style: style || 'none', width, color }
}

function expandBoxValues(value: string): [string, string, string, string] {
  const parts = splitCssValues(value)
  const [top = '', right = top, bottom = top, left = right] = parts
  return [top, right, bottom, left]
}

/** Expands the shorthands we care about so the cascade can merge longhands correctly. */
function normalizeDeclarations(pairs: Array<[string, string]>): Declarations {
  const out: Declarations = Object.create(null)
  const setSide = (side: Side, parsed: { style: string; width: string; color: string }) => {
    out[`border-${side}-style`] = parsed.style
    out[`border-${side}-width`] = parsed.width
    out[`border-${side}-color`] = parsed.color
  }
  for (const [property, value] of pairs) {
    if (property === 'border') {
      const parsed = parseBorderShorthand(value)
      SIDES.forEach((side) => setSide(side, parsed))
    } else if (property === 'border-top' || property === 'border-right' || property === 'border-bottom' || property === 'border-left') {
      setSide(property.slice(7) as Side, parseBorderShorthand(value))
    } else if (property === 'border-width' || property === 'border-style' || property === 'border-color') {
      const kind = property.slice(7)
      expandBoxValues(value).forEach((part, index) => { out[`border-${SIDES[index]}-${kind}`] = kind === 'color' ? part : part.toLowerCase() })
    } else if (property === 'background') {
      let color = ''
      for (const token of splitCssValues(value)) {
        const parsed = parseCssColor(token)
        if (parsed !== undefined) { color = parsed === null ? 'transparent' : token; break }
      }
      out['background-color'] = color || 'transparent'
    } else if (property === 'text-decoration-line') {
      out['text-decoration'] = value
    } else {
      out[property] = value
    }
  }
  return out
}

interface StyleRule { tag: string; classes: string[]; decls: Declarations; specificity: number; order: number }
interface StyleSheetIndex { count: number; byTag: Map<string, StyleRule[]>; byClass: Map<string, StyleRule[]> }

function emptyStyleSheet(): StyleSheetIndex {
  return { count: 0, byTag: new Map(), byClass: new Map() }
}

function withoutBorders(decls: Declarations): Declarations {
  const out: Declarations = Object.create(null)
  for (const key of Object.keys(decls)) if (!key.startsWith('border')) out[key] = decls[key]
  return out
}

function addStyleSheet(index: StyleSheetIndex, css: string, dropTagBorders: boolean, depth = 0) {
  const text = depth ? css : css.replace(/\/\*[\s\S]*?\*\//g, ' ').replace(/<!--|-->/g, ' ').replace(/@(?:import|charset|namespace)\b[^;{]*;/gi, ' ')
  let position = 0
  while (position < text.length) {
    const openBrace = text.indexOf('{', position)
    if (openBrace === -1) break
    const selectorText = text.slice(position, openBrace).trim()
    let depthCount = 1
    let cursor = openBrace + 1
    let quote = ''
    while (cursor < text.length && depthCount > 0) {
      const ch = text[cursor]
      if (ch === '\\') { cursor += 2; continue }
      if (quote) { if (ch === quote) quote = '' }
      else if (ch === '"' || ch === "'") quote = ch
      else if (ch === '{') depthCount += 1
      else if (ch === '}') depthCount -= 1
      cursor += 1
    }
    const body = text.slice(openBrace + 1, Math.max(openBrace + 1, cursor - 1))
    position = cursor
    if (selectorText.startsWith('@')) {
      if (/^@media\b/i.test(selectorText) && !/\bprint\b/i.test(selectorText) && depth < 3) addStyleSheet(index, body, dropTagBorders, depth + 1)
      continue
    }
    const decls = normalizeDeclarations(declarationPairs(body))
    for (const raw of selectorText.split(',')) {
      const selector = raw.trim().toLowerCase()
      const match = /^([a-z][a-z0-9-]*|\*)?((?:\.[a-z0-9_-]+)*)$/.exec(selector)
      if (!match || (!match[1] && !match[2])) continue
      const tag = match[1] && match[1] !== '*' ? match[1] : ''
      const classes = match[2] ? match[2].slice(1).split('.') : []
      const rule: StyleRule = {
        tag,
        classes,
        decls: dropTagBorders && !classes.length ? withoutBorders(decls) : decls,
        specificity: classes.length * 10 + (tag ? 1 : 0),
        order: index.count,
      }
      index.count += 1
      if (classes.length) {
        const list = index.byClass.get(classes[0]) ?? []
        list.push(rule)
        index.byClass.set(classes[0], list)
      } else {
        const key = tag || '*'
        const list = index.byTag.get(key) ?? []
        list.push(rule)
        index.byTag.set(key, list)
      }
    }
  }
}

function matchedRuleDecls(sheet: StyleSheetIndex, tag: string, classes: string[]): Declarations | null {
  if (!sheet.count) return null
  const rules: StyleRule[] = []
  for (const rule of sheet.byTag.get(tag) ?? []) rules.push(rule)
  for (const rule of sheet.byTag.get('*') ?? []) rules.push(rule)
  for (const name of classes) {
    for (const rule of sheet.byClass.get(name) ?? []) {
      if (rules.includes(rule)) continue
      if (rule.tag && rule.tag !== tag) continue
      if (!rule.classes.every((item) => classes.includes(item))) continue
      rules.push(rule)
    }
  }
  if (!rules.length) return null
  rules.sort((a, b) => a.specificity - b.specificity || a.order - b.order)
  const merged: Declarations = Object.create(null)
  for (const rule of rules) Object.assign(merged, rule.decls)
  return merged
}

/** Decodes CSS escapes (`\0022`, `\#`) and strips one level of surrounding quotes. */
export function cssUnescape(raw: string): string {
  let value = raw.trim()
  if (value.length >= 2 && (value[0] === '"' || value[0] === "'") && value[value.length - 1] === value[0]) value = value.slice(1, -1)
  if (value.indexOf('\\') === -1) return value
  let out = ''
  for (let index = 0; index < value.length; index += 1) {
    const ch = value[index]
    if (ch !== '\\') { out += ch; continue }
    const rest = value.slice(index + 1)
    // Excel writes four-digit escapes such as \0022 with no terminator.
    const excelHex = /^00[0-9a-fA-F]{2}/.exec(rest)
    const hex = excelHex ?? /^[0-9a-fA-F]{1,6}/.exec(rest)
    if (hex) {
      const code = Number.parseInt(hex[0], 16)
      out += code && code <= 0x10ffff && (code < 0xd800 || code > 0xdfff) ? String.fromCodePoint(code) : '\uFFFD'
      index += hex[0].length
      if (!excelHex && value[index + 1] === ' ') index += 1
      continue
    }
    if (index + 1 < value.length) { out += value[index + 1]; index += 1 }
  }
  return out
}

const NAMED_COLORS: Record<string, string> = {
  black: '000000', white: 'FFFFFF', red: 'FF0000', lime: '00FF00', green: '008000', blue: '0000FF',
  yellow: 'FFFF00', aqua: '00FFFF', cyan: '00FFFF', fuchsia: 'FF00FF', magenta: 'FF00FF', silver: 'C0C0C0',
  gray: '808080', grey: '808080', maroon: '800000', olive: '808000', purple: '800080', teal: '008080',
  navy: '000080', orange: 'FFA500', darkred: '8B0000', darkblue: '00008B', darkgreen: '006400',
  lightgray: 'D3D3D3', lightgrey: 'D3D3D3', darkgray: 'A9A9A9', darkgrey: 'A9A9A9', dimgray: '696969',
  dimgrey: '696969', gainsboro: 'DCDCDC', whitesmoke: 'F5F5F5', gold: 'FFD700', pink: 'FFC0CB',
  brown: 'A52A2A', violet: 'EE82EE', indigo: '4B0082', lightblue: 'ADD8E6', lightgreen: '90EE90',
  lightyellow: 'FFFFE0', coral: 'FF7F50', tomato: 'FF6347', salmon: 'FA8072', crimson: 'DC143C',
  khaki: 'F0E68C', beige: 'F5F5DC', ivory: 'FFFFF0', lavender: 'E6E6FA', turquoise: '40E0D0', tan: 'D2B48C',
  orchid: 'DA70D6', plum: 'DDA0DD', skyblue: '87CEEB', steelblue: '4682B4', slategray: '708090',
  slategrey: '708090', darkorange: 'FF8C00', forestgreen: '228B22', seagreen: '2E8B57', royalblue: '4169E1',
  midnightblue: '191970', firebrick: 'B22222', chocolate: 'D2691E', sienna: 'A0522D', peru: 'CD853F',
  goldenrod: 'DAA520', olivedrab: '6B8E23', yellowgreen: '9ACD32', limegreen: '32CD32', springgreen: '00FF7F',
  aquamarine: '7FFFD4', cadetblue: '5F9EA0', cornflowerblue: '6495ED', dodgerblue: '1E90FF',
  deepskyblue: '00BFFF', hotpink: 'FF69B4', deeppink: 'FF1493', mediumpurple: '9370DB', darkviolet: '9400D3',
  rebeccapurple: '663399', lightpink: 'FFB6C1', lightcyan: 'E0FFFF', honeydew: 'F0FFF0', mintcream: 'F5FFFA',
  aliceblue: 'F0F8FF', azure: 'F0FFFF', linen: 'FAF0E6', seashell: 'FFF5EE', snow: 'FFFAFA',
  wheat: 'F5DEB3', moccasin: 'FFE4B5', peachpuff: 'FFDAB9', mistyrose: 'FFE4E1', lemonchiffon: 'FFFACD',
  windowtext: '000000', window: 'FFFFFF', buttonface: 'F0F0F0', buttontext: '000000', graytext: '6D6D6D',
}

/** Returns RRGGBB, `null` for transparent, `undefined` when not a (useful) color. */
export function parseCssColor(raw: string | undefined): string | null | undefined {
  if (!raw) return undefined
  const value = raw.trim().toLowerCase()
  if (!value) return undefined
  if (value === 'transparent' || value === 'none') return null
  const hex = /^#([0-9a-f]{3,8})$/.exec(value)
  if (hex) {
    const digits = hex[1]
    if (digits.length === 3 || digits.length === 4) {
      if (digits.length === 4 && digits[3] === '0') return null
      return (digits[0] + digits[0] + digits[1] + digits[1] + digits[2] + digits[2]).toUpperCase()
    }
    if (digits.length === 6) return digits.toUpperCase()
    if (digits.length === 8) return digits.slice(6) === '00' ? null : digits.slice(0, 6).toUpperCase()
    return undefined
  }
  const functional = /^rgba?\(([^)]*)\)$/.exec(value)
  if (functional) {
    const parts = functional[1].split(/[\s,/]+/).filter(Boolean)
    if (parts.length < 3) return undefined
    const channel = (part: string) => {
      const number = Number.parseFloat(part)
      if (!Number.isFinite(number)) return NaN
      return Math.round(Math.min(255, Math.max(0, part.endsWith('%') ? number * 2.55 : number)))
    }
    const rgb = parts.slice(0, 3).map(channel)
    if (rgb.some((part) => Number.isNaN(part))) return undefined
    if (parts[3] !== undefined) {
      const alpha = Number.parseFloat(parts[3]) / (parts[3].endsWith('%') ? 100 : 1)
      if (alpha === 0) return null
    }
    return rgb.map((part) => part.toString(16).padStart(2, '0')).join('').toUpperCase()
  }
  return NAMED_COLORS[value]
}

function parseLengthPt(raw: string | undefined, basePt: number): number | null {
  if (!raw) return null
  const match = /^([+-]?(?:\d+\.?\d*|\.\d+))\s*(pt|px|in|cm|mm|pc|em|rem|%|q)?$/i.exec(raw.trim())
  if (!match) return null
  const value = Number.parseFloat(match[1])
  if (!Number.isFinite(value)) return null
  switch ((match[2] || 'px').toLowerCase()) {
    case 'pt': return value
    case 'px': return value * 0.75
    case 'in': return value * 72
    case 'cm': return (value * 72) / 2.54
    case 'mm': return (value * 72) / 25.4
    case 'q': return (value * 72) / 101.6
    case 'pc': return value * 12
    case 'em': case 'rem': return value * basePt
    case '%': return (value / 100) * basePt
    default: return null
  }
}

function parseLengthPx(raw: string | undefined): number | undefined {
  if (!raw || /%$/.test(raw.trim())) return undefined
  const pt = parseLengthPt(raw, 11)
  if (pt === null || pt <= 0) return undefined
  return Math.round((pt / 0.75) * 100) / 100
}

const BROWSER_FONT_KEYWORDS: Record<string, number> = { 'xx-small': 7, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, 'xxx-large': 36 }
// LibreOffice writes its 10pt default as `x-small` (its own 8/10/12/14/18/24/36 ladder).
const LIBREOFFICE_FONT_KEYWORDS: Record<string, number> = { 'xx-small': 8, 'x-small': 10, small: 12, medium: 14, large: 18, 'x-large': 24, 'xx-large': 36 }
const HTML_FONT_SIZES = [8, 10, 12, 14, 18, 24, 36]
const GENERIC_FONT_FAMILIES = new Set(['serif', 'sans-serif', 'monospace', 'cursive', 'fantasy', 'system-ui', '-apple-system', 'blinkmacsystemfont', 'ui-sans-serif', 'ui-serif', 'ui-monospace', 'inherit', 'initial', 'auto'])

function primaryFontFamily(raw: string | undefined): string | undefined {
  if (!raw) return undefined
  for (const part of raw.split(',')) {
    const name = cssUnescape(part.trim()).replace(/["'\\\u0000-\u001f]/g, '').trim()
    if (!name || GENERIC_FONT_FAMILIES.has(name.toLowerCase())) continue
    return name.slice(0, 100)
  }
  return undefined
}

function roundHalf(value: number) {
  return Math.round(value * 2) / 2
}

// ---------------------------------------------------------------------------------------------
// Number formats
// ---------------------------------------------------------------------------------------------

function sanitizeNumFmt(raw: string | undefined): string | undefined {
  if (typeof raw !== 'string') return undefined
  const value = raw.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '').slice(0, 255)
  if (!value.trim() || /^general$/i.test(value.trim())) return undefined
  return value
}

const EXCEL_NAMED_FORMATS: Record<string, string> = {
  general: '',
  fixed: '0.00',
  standard: '#,##0.00',
  percent: '0.00%',
  scientific: '0.00E+00',
  currency: '"$"#,##0.00',
  'short date': 'm/d/yyyy',
  'medium date': 'd-mmm-yy',
  'long date': 'dddd, mmmm d, yyyy',
  'short time': 'h:mm',
  'medium time': 'h:mm AM/PM',
  'long time': 'h:mm:ss AM/PM',
  '@': '@',
  'yes/no': '"Yes";"Yes";"No"',
  'true/false': '"TRUE";"TRUE";"FALSE"',
  'on/off': '"On";"On";"Off"',
}

function msoNumberFormat(raw: string | undefined): string | undefined {
  if (!raw) return undefined
  const value = cssUnescape(raw)
  const named = EXCEL_NAMED_FORMATS[value.trim().toLowerCase()]
  if (named !== undefined) return named || undefined
  return sanitizeNumFmt(value)
}

/** Removes quoted literals, escapes, fill/pad characters and [..] sections (except elapsed time). */
function formatSkeleton(format: string): string {
  return format
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/[_*]./g, '')
    .replace(/\[(?![hms]+\])[^\]]*\]/gi, '')
}

export function isDateNumberFormat(format: string | undefined): boolean {
  if (!format) return false
  const skeleton = formatSkeleton(format.split(';')[0])
  return /[dmyhs]/i.test(skeleton) && !/^general$/i.test(format.trim())
}

function isTimeOnlyFormat(format: string) {
  const skeleton = formatSkeleton(format.split(';')[0])
  return isDateNumberFormat(format) && !/[dy]/i.test(skeleton) && /[hs]/i.test(skeleton)
}

function isDayFirstFormat(format: string | undefined) {
  if (!format || !isDateNumberFormat(format)) return undefined
  const skeleton = formatSkeleton(format.split(';')[0]).toLowerCase()
  const day = skeleton.indexOf('d')
  const month = skeleton.search(/m(?![^h]*s)/)
  if (day === -1 || month === -1) return undefined
  return day < month
}

function sheetsNumberFormat(raw: string | undefined): string | undefined {
  const json = parseJson(raw) as Record<string, unknown> | undefined
  if (!json || typeof json !== 'object') return undefined
  const pattern = json['2']
  if (typeof pattern === 'string' && pattern.trim()) return sanitizeNumFmt(pattern)
  switch (Number(json['1'])) {
    case 1: return '@'
    case 3: return '0.00%'
    case 4: return '"$"#,##0.00'
    case 5: return 'm/d/yyyy'
    case 6: return 'h:mm:ss AM/PM'
    case 7: return 'm/d/yyyy h:mm:ss'
    case 8: return '0.00E+00'
    default: return undefined
  }
}

function sheetsNumberFormatType(format: string): number {
  const skeleton = formatSkeleton(format.split(';')[0])
  if (format.trim() === '@') return 1
  if (isDateNumberFormat(format)) {
    const date = /[dy]/i.test(skeleton) || /m(?![^h]*s)/i.test(skeleton.replace(/h+[^a-z]*m+/gi, ''))
    const time = /[hs]/i.test(skeleton)
    return date && time ? 7 : time ? 6 : 5
  }
  if (skeleton.includes('%')) return 3
  if (/[$€£¥₹₩₽]/.test(format)) return 4
  if (/E[+-]/i.test(skeleton)) return 8
  return 2
}

function libreOfficeNumberFormat(raw: string | undefined): string | undefined {
  if (!raw) return undefined
  const parts = raw.split(';')
  if (parts.length < 3) return undefined
  const format = parts.slice(2).join(';')
  if (/^(standard|general)$/i.test(format.trim())) return undefined
  if (/^boolean$/i.test(format.trim())) return undefined
  return sanitizeNumFmt(format)
}

// ---------------------------------------------------------------------------------------------
// Smart value parsing (mirrors the grid's typing rules, plus currency / grouping / month names)
// ---------------------------------------------------------------------------------------------

export interface SmartParseOptions {
  decimalSeparator?: '.' | ','
  dayFirst?: boolean
}

export interface SmartValue {
  value: CellScalar
  numFmt?: string
  type?: string
}

const MONTH_NAMES = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
const WEEKDAY_NAMES = ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday']

function monthFromName(raw: string): number | null {
  const name = raw.toLowerCase().replace(/\.$/, '')
  if (name.length < 3) return null
  const index = MONTH_NAMES.findIndex((month) => month.startsWith(name))
  return index === -1 ? null : index + 1
}

function isWeekdayName(raw: string) {
  const name = raw.toLowerCase().replace(/\.$/, '')
  return name.length >= 3 && WEEKDAY_NAMES.some((day) => day.startsWith(name))
}

function serialFromDate(year: number, month: number, day: number): number | null {
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null
  const timestamp = Date.UTC(year, month - 1, day)
  const date = new Date(timestamp)
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null
  return (timestamp - Date.UTC(1899, 11, 30)) / 86_400_000
}

function expandYear(raw: string) {
  const year = Number(raw)
  if (raw.length === 4) return year
  return year < 30 ? 2000 + year : 1900 + year
}

const TIME_SOURCE = '(\\d{1,2}):(\\d{2})(?::(\\d{2})(?:\\.(\\d{1,3}))?)?\\s*([AaPp]\\.?[Mm]\\.?)?'
const TIME_ONLY = new RegExp(`^${TIME_SOURCE}$`)
const DATE_ISO = new RegExp(`^(\\d{4})-(\\d{1,2})-(\\d{1,2})(?:(?:T|\\s+)${TIME_SOURCE})?$`)
const DATE_YMD_SLASH = new RegExp(`^(\\d{4})/(\\d{1,2})/(\\d{1,2})(?:\\s+${TIME_SOURCE})?$`)
const DATE_NUMERIC = new RegExp(`^(\\d{1,2})([/.-])(\\d{1,2})\\2(\\d{4}|\\d{2})(?:\\s+${TIME_SOURCE})?$`)
const DATE_DAY_MONTH = new RegExp(`^(\\d{1,2})(?:st|nd|rd|th)?([-\\s])([A-Za-z]{3,9}\\.?)[-\\s,]+(\\d{4}|\\d{2})(?:\\s+${TIME_SOURCE})?$`)
const DATE_MONTH_DAY = new RegExp(`^(?:([A-Za-z]{3,9}\\.?),?\\s+)?([A-Za-z]{3,9}\\.?)\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})(?:\\s+${TIME_SOURCE})?$`)
const DATE_MONTH_YEAR = /^([A-Za-z]{3,9})[-\s](\d{4}|\d{2})$/

function timeParts(match: RegExpExecArray, offset: number) {
  if (match[offset] === undefined) return null
  let hours = Number(match[offset])
  const minutes = Number(match[offset + 1])
  const seconds = match[offset + 2] === undefined ? 0 : Number(match[offset + 2])
  const millis = match[offset + 3] === undefined ? 0 : Number(`0.${match[offset + 3]}`)
  const meridiem = match[offset + 4]?.replace(/\./g, '').toLowerCase()
  if (minutes > 59 || seconds > 59) return null
  if (meridiem) {
    if (hours < 1 || hours > 12) return null
    if (meridiem === 'pm' && hours !== 12) hours += 12
    if (meridiem === 'am' && hours === 12) hours = 0
  }
  const fraction = (hours * 3600 + minutes * 60 + seconds + millis) / 86_400
  const withSeconds = match[offset + 2] !== undefined
  const format = meridiem
    ? (withSeconds ? 'h:mm:ss AM/PM' : 'h:mm AM/PM')
    : hours >= 24 ? (withSeconds ? '[h]:mm:ss' : '[h]:mm') : (withSeconds ? 'h:mm:ss' : 'h:mm')
  return { fraction, format, hours }
}

function parseDateTime(text: string, dayFirst: boolean): { serial: number; numFmt: string; date: boolean } | null {
  let match = TIME_ONLY.exec(text)
  if (match) {
    const time = timeParts(match, 1)
    if (!time || (time.hours > 9999)) return null
    return { serial: time.fraction, numFmt: time.format, date: false }
  }
  const finish = (serial: number | null, format: string, groups: RegExpExecArray, timeOffset: number) => {
    if (serial === null) return null
    if (groups[timeOffset] === undefined) return { serial, numFmt: format, date: true }
    const time = timeParts(groups, timeOffset)
    if (!time || time.hours >= 24) return null
    return { serial: serial + time.fraction, numFmt: `${format} ${time.format}`, date: true }
  }
  if ((match = DATE_ISO.exec(text))) return finish(serialFromDate(Number(match[1]), Number(match[2]), Number(match[3])), dayFirst ? 'd/m/yyyy' : 'm/d/yyyy', match, 4)
  if ((match = DATE_YMD_SLASH.exec(text))) return finish(serialFromDate(Number(match[1]), Number(match[2]), Number(match[3])), 'yyyy/m/d', match, 4)
  if ((match = DATE_NUMERIC.exec(text))) {
    if (match[2] === '.' && !dayFirst) return null
    const first = Number(match[1])
    const second = Number(match[3])
    const year = expandYear(match[4])
    const serial = dayFirst ? serialFromDate(year, second, first) : serialFromDate(year, first, second)
    return finish(serial, dayFirst ? 'd/m/yyyy' : 'm/d/yyyy', match, 5)
  }
  if ((match = DATE_DAY_MONTH.exec(text))) {
    const month = monthFromName(match[3])
    if (!month) return null
    const fullName = match[3].replace(/\.$/, '').length > 3
    const format = match[2] === '-'
      ? (match[4].length === 2 ? 'd-mmm-yy' : 'd-mmm-yyyy')
      : (fullName ? 'd mmmm yyyy' : 'd mmm yyyy')
    return finish(serialFromDate(expandYear(match[4]), month, Number(match[1])), format, match, 5)
  }
  if ((match = DATE_MONTH_DAY.exec(text))) {
    if (match[1] && !isWeekdayName(match[1])) return null
    const month = monthFromName(match[2])
    if (!month) return null
    const fullName = match[2].replace(/\.$/, '').length > 3
    const format = match[1] ? 'dddd, mmmm d, yyyy' : fullName ? 'mmmm d, yyyy' : 'mmm d, yyyy'
    return finish(serialFromDate(Number(match[4]), month, Number(match[3])), format, match, 5)
  }
  if ((match = DATE_MONTH_YEAR.exec(text))) {
    const month = monthFromName(match[1])
    if (!month) return null
    const serial = serialFromDate(expandYear(match[2]), month, 1)
    return serial === null ? null : { serial, numFmt: 'mmm-yy', date: true }
  }
  return null
}

function parseNumberCore(raw: string, decimal: '.' | ',') {
  const pattern = decimal === '.'
    ? /^(\d{1,3}(?:,\d{3})+|\d+)?(?:\.(\d*))?(?:[eE]([+-]?\d{1,3}))?$/
    : /^(\d{1,3}(?:[. \u202f]\d{3})+|\d+)?(?:,(\d*))?(?:[eE]([+-]?\d{1,3}))?$/
  const match = pattern.exec(raw)
  if (!match || (!match[1] && !match[2])) return null
  const integer = (match[1] || '').replace(/[,. \u202f]/g, '')
  if (integer.length > 1 && integer[0] === '0') return null
  const value = Number(`${integer || '0'}.${match[2] || '0'}${match[3] ? `e${match[3]}` : ''}`)
  if (!Number.isFinite(value)) return null
  return {
    value,
    grouped: Boolean(match[1] && match[1].length !== integer.length),
    decimals: match[2]?.length ?? 0,
    exponent: Boolean(match[3]),
  }
}

const CURRENCY_PREFIX = /^(US\$|C\$|A\$|NZ\$|HK\$|R\$|[$€£¥₹₩₽¢₺₴₦₱₫₪])\s*/
const CURRENCY_SUFFIX = /(\s*)([$€£¥₹₩₽¢₺₴₦₱₫₪]|kr|zł|Kč|CHF|lei|Ft)$/

function decimalPattern(decimals: number, minimum = 0) {
  const count = Math.min(10, Math.max(decimals, decimals ? minimum : 0))
  return count ? `.${'0'.repeat(count)}` : ''
}

/**
 * Types a displayed/pasted string the way a spreadsheet would when it is typed: booleans,
 * percentages, dates and times, grouped numbers, currency amounts and (1,234) negatives.
 * Leading-zero digit strings ("007") stay text. Returns null for blank input.
 */
export function parseSmartValue(input: string, options: SmartParseOptions = {}): SmartValue | null {
  const text = input.replace(/[\u00A0\u2007\u202F]/g, ' ').trim()
  if (!text) return null
  if (/^(true|false)$/i.test(text)) return { value: text.toLowerCase() === 'true' }
  const decimal = options.decimalSeparator === ',' ? ',' : '.'

  if (text.endsWith('%')) {
    let inner = text.slice(0, -1).trim()
    let negative = false
    if (/^[+\-\u2212]/.test(inner)) { negative = inner[0] !== '+'; inner = inner.slice(1).trim() }
    const core = parseNumberCore(inner, decimal)
    if (core && !core.exponent) {
      const value = Number.parseFloat(((negative ? -core.value : core.value) / 100).toPrecision(15))
      return { value, numFmt: `${core.grouped ? '#,##0' : '0'}${decimalPattern(core.decimals)}%` }
    }
    return { value: input }
  }

  const date = parseDateTime(text, Boolean(options.dayFirst))
  if (date) return date.date ? { value: date.serial, numFmt: date.numFmt, type: 'date' } : { value: date.serial, numFmt: date.numFmt }

  let rest = text
  let negative = false
  let parentheses = false
  if (rest.length > 2 && rest.startsWith('(') && rest.endsWith(')')) { parentheses = true; negative = true; rest = rest.slice(1, -1).trim() }
  if (/^[+\-\u2212]/.test(rest)) { if (rest[0] !== '+') negative = !negative; rest = rest.slice(1).trim() }
  let symbol = ''
  let suffix = false
  let suffixSpace = false
  const prefix = CURRENCY_PREFIX.exec(rest)
  if (prefix) {
    symbol = prefix[1]
    rest = rest.slice(prefix[0].length)
    if (/^[+\-\u2212]/.test(rest)) { if (rest[0] !== '+') negative = !negative; rest = rest.slice(1).trim() }
  } else {
    const trailing = CURRENCY_SUFFIX.exec(rest)
    if (trailing && trailing.index > 0) {
      symbol = trailing[2]
      suffix = true
      suffixSpace = Boolean(trailing[1])
      rest = rest.slice(0, trailing.index)
    }
  }
  if (rest.endsWith('-') && !negative && symbol) { negative = true; rest = rest.slice(0, -1).trim() }
  const core = parseNumberCore(rest, decimal)
  if (!core) return { value: input }
  const value = negative ? -core.value : core.value
  let numFmt: string | undefined
  if (symbol) {
    const body = `#,##0${decimalPattern(core.decimals, 2)}`
    const positive = suffix ? `${body}${suffixSpace ? ' ' : ''}"${symbol}"` : `"${symbol}"${body}`
    numFmt = parentheses ? `${positive};(${positive})` : positive
  } else if (core.grouped) {
    const body = `#,##0${decimalPattern(core.decimals)}`
    numFmt = parentheses ? `${body};(${body})` : body
  }
  return numFmt ? { value, numFmt } : { value }
}

/**
 * Converts one pasted text field into a cell, mirroring what typing it would do:
 * a leading apostrophe keeps text, "=..." becomes a formula, everything else is typed by
 * `parseSmartValue`.
 */
export function parsePastedValue(draft: string, options: ClipboardParseOptions = {}): CellData {
  const text = draft.length > MAX_TEXT_LENGTH ? draft.slice(0, MAX_TEXT_LENGTH) : draft
  if (text.startsWith("'")) return { value: text.slice(1) }
  if (options.formulas !== false && text.startsWith('=') && text.trim().length > 1) {
    return { formula: text.trim().replace(/^=/, '').slice(0, MAX_FORMULA_LENGTH) }
  }
  if (!text.length) return {}
  const smart = parseSmartValue(text, options)
  if (!smart) return { value: text }
  const cell: CellData = { value: typeof smart.value === 'string' ? text : smart.value }
  if (smart.numFmt) cell.numFmt = smart.numFmt
  if (smart.type) cell.type = smart.type
  return cell
}

// ---------------------------------------------------------------------------------------------
// Plain text (TSV / single column)
// ---------------------------------------------------------------------------------------------

/**
 * Excel-compatible delimited text parser: a field is quoted only when it starts with `"`,
 * `""` escapes a quote, quoted fields may contain delimiters and line breaks, and a stray
 * quote inside an unquoted field is literal. A single trailing line break is ignored.
 */
export function parseDelimitedText(input: string, delimiter = '\t'): string[][] {
  const text = input.charCodeAt(0) === 0xfeff ? input.slice(1) : input
  const rows: string[][] = []
  const length = text.length
  if (!length) return rows
  let row: string[] = []
  let index = 0
  const unquotedEnd = (from: number) => {
    let cursor = from
    while (cursor < length) {
      const ch = text[cursor]
      if (ch === delimiter || ch === '\n' || ch === '\r') break
      cursor += 1
    }
    return cursor
  }
  for (;;) {
    let value: string
    if (text[index] === '"') {
      let cursor = index + 1
      let buffer = ''
      let closed = false
      while (cursor < length) {
        const quote = text.indexOf('"', cursor)
        if (quote === -1) { buffer += text.slice(cursor); cursor = length; break }
        buffer += text.slice(cursor, quote)
        if (text[quote + 1] === '"') { buffer += '"'; cursor = quote + 2; continue }
        closed = true
        cursor = quote + 1
        break
      }
      const after = text[cursor]
      if (closed && (cursor >= length || after === delimiter || after === '\n' || after === '\r')) {
        value = buffer
        index = cursor
      } else {
        const end = unquotedEnd(index)
        value = text.slice(index, end)
        index = end
      }
    } else {
      const end = unquotedEnd(index)
      value = text.slice(index, end)
      index = end
    }
    row.push(value)
    if (index >= length) { rows.push(row); break }
    const ch = text[index]
    if (ch === delimiter) {
      index += 1
      if (index >= length) { row.push(''); rows.push(row); break }
      continue
    }
    index += ch === '\r' && text[index + 1] === '\n' ? 2 : 1
    rows.push(row)
    row = []
    if (index >= length) break
  }
  return rows
}

/**
 * Parses clipboard text into typed cells: TSV when it contains tabs, otherwise one value
 * per line (quoted multi-line fields stay in one cell). The matrix is padded to a rectangle.
 */
export function parsePastedText(text: string, options: ClipboardParseOptions = {}): CellData[][] {
  const maxCells = options.maxCells ?? CLIPBOARD_MAX_CELLS
  const rows = parseDelimitedText(text, options.delimiter ?? '\t')
  if (!rows.length) return []
  let width = 0
  for (const row of rows) if (row.length > width) width = row.length
  if (rows.length * width > maxCells || rows.length > SHEET_ROWS || width > SHEET_COLS) throw new ClipboardTooLargeError()
  return rows.map((row) => {
    const cells: CellData[] = new Array(width)
    for (let col = 0; col < width; col += 1) cells[col] = col < row.length ? parsePastedValue(row[col], options) : {}
    return cells
  })
}

// ---------------------------------------------------------------------------------------------
// A1 / R1C1 formula references
// ---------------------------------------------------------------------------------------------

export interface FormulaRefPoint {
  /** 0-based absolute row, or null for a whole-column reference. */
  row: number | null
  /** 0-based absolute column, or null for a whole-row reference. */
  col: number | null
  rowAbs: boolean
  colAbs: boolean
}

export interface FormulaReference {
  start: FormulaRefPoint
  end: FormulaRefPoint | null
}

interface ReferenceHit { ref: FormulaReference; end: number }
type ReferenceReader = (formula: string, position: number) => ReferenceHit | null

const IDENT_CHAR = /[A-Za-z0-9_.?\\\u00C0-\uFFFF]/
const IDENT_START = /[A-Za-z_\\\u00C0-\uFFFF]/
const A1_CELL = /(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/y
const A1_COLUMNS = /(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})/y
const A1_ROWS = /(\$?)(\d{1,7}):(\$?)(\d{1,7})/y
const R1C1_CELL = /R(?:\[([+-]?\d{1,7})\]|(\d{1,7}))?C(?:\[([+-]?\d{1,5})\]|(\d{1,5}))?/y
const R1C1_ROWS = /R(?:\[([+-]?\d{1,7})\]|(\d{1,7}))?:R(?:\[([+-]?\d{1,7})\]|(\d{1,7}))?/y
const R1C1_COLUMNS = /C(?:\[([+-]?\d{1,5})\]|(\d{1,5}))?:C(?:\[([+-]?\d{1,5})\]|(\d{1,5}))?/y
const SHEET_PREFIX = /[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*!/y
const NUMBER_TOKEN = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y
const IDENT_TOKEN = /[A-Za-z_\\\u00C0-\uFFFF][A-Za-z0-9_.?\\\u00C0-\uFFFF]*/y

const referenceBoundary = (formula: string, end: number) => {
  const next = formula[end]
  return next === undefined || (!IDENT_CHAR.test(next) && next !== '(')
}

const readA1Reference: ReferenceReader = (formula, position) => {
  A1_CELL.lastIndex = position
  const cell = A1_CELL.exec(formula)
  if (cell) {
    const col = columnNumber(cell[2])
    const row = Number(cell[4]) - 1
    const end = position + cell[0].length
    const inSheet = col < SHEET_COLS && row >= 0 && row < SHEET_ROWS
    const start: FormulaRefPoint = { row, col, rowAbs: cell[3] === '$', colAbs: cell[1] === '$' }
    if (inSheet && formula[end] === ':') {
      A1_CELL.lastIndex = end + 1
      const second = A1_CELL.exec(formula)
      if (second) {
        const col2 = columnNumber(second[2])
        const row2 = Number(second[4]) - 1
        const end2 = end + 1 + second[0].length
        if (col2 < SHEET_COLS && row2 >= 0 && row2 < SHEET_ROWS && referenceBoundary(formula, end2)) {
          return { ref: { start, end: { row: row2, col: col2, rowAbs: second[3] === '$', colAbs: second[1] === '$' } }, end: end2 }
        }
      }
      return { ref: { start, end: null }, end }
    }
    if (inSheet && referenceBoundary(formula, end)) return { ref: { start, end: null }, end }
  }
  A1_COLUMNS.lastIndex = position
  const columns = A1_COLUMNS.exec(formula)
  if (columns) {
    const c1 = columnNumber(columns[2])
    const c2 = columnNumber(columns[4])
    const end = position + columns[0].length
    if (c1 < SHEET_COLS && c2 < SHEET_COLS && referenceBoundary(formula, end)) {
      return { ref: { start: { row: null, col: c1, rowAbs: false, colAbs: columns[1] === '$' }, end: { row: null, col: c2, rowAbs: false, colAbs: columns[3] === '$' } }, end }
    }
  }
  A1_ROWS.lastIndex = position
  const rows = A1_ROWS.exec(formula)
  if (rows) {
    const r1 = Number(rows[2]) - 1
    const r2 = Number(rows[4]) - 1
    const end = position + rows[0].length
    if (r1 >= 0 && r2 >= 0 && r1 < SHEET_ROWS && r2 < SHEET_ROWS && referenceBoundary(formula, end) && formula[end] !== '.') {
      return { ref: { start: { row: r1, col: null, rowAbs: rows[1] === '$', colAbs: false }, end: { row: r2, col: null, rowAbs: rows[3] === '$', colAbs: false } }, end }
    }
  }
  return null
}

function r1c1Component(relative: string | undefined, absolute: string | undefined, base: number): { value: number; abs: boolean } {
  if (absolute !== undefined) return { value: Number(absolute) - 1, abs: true }
  return { value: base + (relative === undefined ? 0 : Number(relative)), abs: false }
}

function makeR1C1Reader(baseRow: number, baseCol: number): ReferenceReader {
  const point = (match: RegExpExecArray, offset: number): FormulaRefPoint => {
    const row = r1c1Component(match[offset], match[offset + 1], baseRow)
    const col = r1c1Component(match[offset + 2], match[offset + 3], baseCol)
    return { row: row.value, col: col.value, rowAbs: row.abs, colAbs: col.abs }
  }
  return (formula, position) => {
    R1C1_CELL.lastIndex = position
    const cell = R1C1_CELL.exec(formula)
    if (cell) {
      const end = position + cell[0].length
      if (formula[end] === ':') {
        R1C1_CELL.lastIndex = end + 1
        const second = R1C1_CELL.exec(formula)
        if (second && referenceBoundary(formula, end + 1 + second[0].length)) {
          return { ref: { start: point(cell, 1), end: point(second, 1) }, end: end + 1 + second[0].length }
        }
      }
      if (referenceBoundary(formula, end)) return { ref: { start: point(cell, 1), end: null }, end }
    }
    R1C1_ROWS.lastIndex = position
    const rows = R1C1_ROWS.exec(formula)
    if (rows && referenceBoundary(formula, position + rows[0].length)) {
      const a = r1c1Component(rows[1], rows[2], baseRow)
      const b = r1c1Component(rows[3], rows[4], baseRow)
      return { ref: { start: { row: a.value, col: null, rowAbs: a.abs, colAbs: false }, end: { row: b.value, col: null, rowAbs: b.abs, colAbs: false } }, end: position + rows[0].length }
    }
    R1C1_COLUMNS.lastIndex = position
    const columns = R1C1_COLUMNS.exec(formula)
    if (columns && referenceBoundary(formula, position + columns[0].length)) {
      const a = r1c1Component(columns[1], columns[2], baseCol)
      const b = r1c1Component(columns[3], columns[4], baseCol)
      return { ref: { start: { row: null, col: a.value, rowAbs: false, colAbs: a.abs }, end: { row: null, col: b.value, rowAbs: false, colAbs: b.abs } }, end: position + columns[0].length }
    }
    return null
  }
}

function skipStringLiteral(formula: string, start: number) {
  let index = start + 1
  while (index < formula.length) {
    if (formula[index] === '"') {
      if (formula[index + 1] === '"') { index += 2; continue }
      return index + 1
    }
    index += 1
  }
  return formula.length
}

function skipBrackets(formula: string, start: number) {
  let depth = 0
  for (let index = start; index < formula.length; index += 1) {
    if (formula[index] === '[') depth += 1
    else if (formula[index] === ']') { depth -= 1; if (!depth) return index + 1 }
  }
  return formula.length
}

function quotedSheetEnd(formula: string, start: number) {
  let index = start + 1
  while (index < formula.length) {
    if (formula[index] === "'") {
      if (formula[index + 1] === "'") { index += 2; continue }
      return index + 1
    }
    index += 1
  }
  return -1
}

/** Walks a formula and rewrites every cell/range reference; strings and names are untouched. */
function rewriteReferences(formula: string, read: ReferenceReader, emit: (ref: FormulaReference) => string): string {
  let out = ''
  let index = 0
  const length = formula.length
  while (index < length) {
    const ch = formula[index]
    if (ch === '"') {
      const end = skipStringLiteral(formula, index)
      out += formula.slice(index, end)
      index = end
      continue
    }
    if (ch === '[') {
      const end = skipBrackets(formula, index)
      out += formula.slice(index, end)
      index = end
      continue
    }
    if (ch === "'") {
      const close = quotedSheetEnd(formula, index)
      if (close !== -1 && formula[close] === '!') {
        const prefix = formula.slice(index, close + 1)
        const hit = read(formula, close + 1)
        if (hit) { out += prefix + emit(hit.ref); index = hit.end; continue }
        out += prefix
        index = close + 1
        continue
      }
      out += ch
      index += 1
      continue
    }
    const code = formula.charCodeAt(index)
    const digit = code >= 48 && code <= 57
    if (digit || ch === '$' || IDENT_START.test(ch)) {
      if (!digit && ch !== '$') {
        SHEET_PREFIX.lastIndex = index
        const sheet = SHEET_PREFIX.exec(formula)
        if (sheet) {
          const hit = read(formula, index + sheet[0].length)
          if (hit) { out += sheet[0] + emit(hit.ref); index = hit.end; continue }
        }
      }
      const hit = read(formula, index)
      if (hit) { out += emit(hit.ref); index = hit.end; continue }
      if (digit) {
        NUMBER_TOKEN.lastIndex = index
        const number = NUMBER_TOKEN.exec(formula)
        const end = number ? index + number[0].length : index + 1
        out += formula.slice(index, end)
        index = end
      } else if (ch === '$') {
        out += ch
        index += 1
      } else {
        IDENT_TOKEN.lastIndex = index
        const ident = IDENT_TOKEN.exec(formula)
        const end = ident ? index + ident[0].length : index + 1
        out += formula.slice(index, end)
        index = end
      }
      continue
    }
    out += ch
    index += 1
  }
  return out
}

const pointInSheet = (point: FormulaRefPoint) => (point.row === null || (point.row >= 0 && point.row < SHEET_ROWS)) && (point.col === null || (point.col >= 0 && point.col < SHEET_COLS))

function formatA1Point(point: FormulaRefPoint) {
  if (point.row === null) return `${point.colAbs ? '$' : ''}${columnLabel(point.col as number)}`
  if (point.col === null) return `${point.rowAbs ? '$' : ''}${point.row + 1}`
  return `${point.colAbs ? '$' : ''}${columnLabel(point.col)}${point.rowAbs ? '$' : ''}${point.row + 1}`
}

export function formatA1Reference(ref: FormulaReference): string {
  if (!pointInSheet(ref.start) || (ref.end && !pointInSheet(ref.end))) return '#REF!'
  return ref.end ? `${formatA1Point(ref.start)}:${formatA1Point(ref.end)}` : formatA1Point(ref.start)
}

function formatR1C1Point(point: FormulaRefPoint, baseRow: number, baseCol: number) {
  const row = point.row === null ? '' : point.rowAbs ? `R${point.row + 1}` : `R[${point.row - baseRow}]`
  const col = point.col === null ? '' : point.colAbs ? `C${point.col + 1}` : `C[${point.col - baseCol}]`
  return row + col
}

export function formatR1C1Reference(ref: FormulaReference, baseRow: number, baseCol: number): string {
  if (!pointInSheet(ref.start) || (ref.end && !pointInSheet(ref.end))) return '#REF!'
  const start = formatR1C1Point(ref.start, baseRow, baseCol)
  return ref.end ? `${start}:${formatR1C1Point(ref.end, baseRow, baseCol)}` : start
}

/** Visits each A1 reference of a formula (0-based absolute coordinates) and rewrites it. */
export function mapA1References(formula: string, map: (ref: FormulaReference) => string): string {
  return rewriteReferences(formula, readA1Reference, map)
}

/** Converts an A1 formula written in cell (row, col) to Google Sheets style R1C1 (`R[0]C[-1]`). */
export function a1ToR1C1(formula: string, row: number, col: number): string {
  return rewriteReferences(formula, readA1Reference, (ref) => formatR1C1Reference(ref, row, col))
}

/** Converts an R1C1 formula (Sheets `R[0]C[-1]`, Excel `RC[-1]`, absolute `R1C1`) in cell (row, col) to A1. */
export function r1c1ToA1(formula: string, row: number, col: number): string {
  return rewriteReferences(formula, makeR1C1Reader(row, col), formatA1Reference)
}

/** Shifts relative A1 references by a row/column delta (absolute parts stay; out of range → #REF!). */
export function shiftFormulaA1(formula: string, rowDelta: number, colDelta: number): string {
  const shift = (point: FormulaRefPoint): FormulaRefPoint => ({
    ...point,
    row: point.row === null || point.rowAbs ? point.row : point.row + rowDelta,
    col: point.col === null || point.colAbs ? point.col : point.col + colDelta,
  })
  return rewriteReferences(formula, readA1Reference, (ref) => formatA1Reference({ start: shift(ref.start), end: ref.end ? shift(ref.end) : null }))
}

/**
 * Moves a formula from source cell to destination cell of a transposed paste (Excel semantics):
 * fully relative references swap their row/column offsets, absolute references stay, and
 * mixed references shift like a normal paste.
 */
export function transposeFormula(formula: string, sourceRow: number, sourceCol: number, destRow: number, destCol: number): string {
  const move = (point: FormulaRefPoint): FormulaRefPoint => {
    const rowRelative = point.row !== null && !point.rowAbs
    const colRelative = point.col !== null && !point.colAbs
    if (point.row === null) {
      return colRelative ? { row: destRow + ((point.col as number) - sourceCol), col: null, rowAbs: false, colAbs: false } : point
    }
    if (point.col === null) {
      return rowRelative ? { row: null, col: destCol + (point.row - sourceRow), rowAbs: false, colAbs: false } : point
    }
    if (rowRelative && colRelative) {
      return { row: destRow + (point.col - sourceCol), col: destCol + (point.row - sourceRow), rowAbs: false, colAbs: false }
    }
    return {
      ...point,
      row: rowRelative ? point.row + (destRow - sourceRow) : point.row,
      col: colRelative ? point.col + (destCol - sourceCol) : point.col,
    }
  }
  return rewriteReferences(formula, readA1Reference, (ref) => {
    const start = move(ref.start)
    const end = ref.end ? move(ref.end) : null
    if (end && start.row !== null && end.row !== null && start.row > end.row) [start.row, end.row] = [end.row, start.row]
    if (end && start.col !== null && end.col !== null && start.col > end.col) [start.col, end.col] = [end.col, start.col]
    return formatA1Reference({ start, end })
  })
}

// ---------------------------------------------------------------------------------------------
// HTML → CellData
// ---------------------------------------------------------------------------------------------

const INHERITED_PROPERTIES = ['color', 'font-family', 'font-size', 'font-weight', 'font-style', 'text-align', 'white-space', 'direction', 'text-decoration', 'text-underline-style', 'vertical-align', '-attr-align', 'mso-number-format']
const ROW_INHERITED_PROPERTIES = [...INHERITED_PROPERTIES, 'background-color']
const SKIP_CONTENT_TAGS = new Set(['img', 'script', 'style', 'svg', 'math', 'object', 'embed', 'iframe', 'input', 'select', 'button', 'textarea', 'template', 'head', 'title', 'noscript', 'canvas', 'video', 'audio', 'picture', 'map', 'meta', 'link', 'datalist', 'colgroup', 'col'])
const BLOCK_TAGS = new Set(['p', 'div', 'li', 'ul', 'ol', 'dl', 'dt', 'dd', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'blockquote', 'pre', 'table', 'tr', 'section', 'article', 'header', 'footer', 'aside', 'nav', 'figure', 'figcaption', 'address', 'hr', 'caption', 'center', 'main', 'form', 'fieldset', 'tbody', 'thead', 'tfoot'])

/** Style facts derived once per distinct (inherited + own) declaration set. */
interface CellDeclInfo {
  font: CellFont
  finalFont?: CellFont
  fill?: CellFill
  border?: CellBorder
  /** null when alignment depends on the cell value (legacy align attribute). */
  alignment?: CellAlignment | null
  msoFormat?: string
}

interface ParseContext {
  source: ClipboardSource
  sheet: StyleSheetIndex
  cache: Map<string, Declarations>
  mergeCache: WeakMap<Declarations, Map<Declarations, Declarations>>
  pickCache: Map<string[], WeakMap<Declarations, Declarations>>
  infoCache: WeakMap<Declarations, CellDeclInfo>
  defaultFont: { name?: string; size: number }
  formulaOrigin: ClipboardCoord
  smart: SmartParseOptions
  keywordSizes: Record<string, number>
  maxCells: number
}

function pick(decls: Declarations, keys: string[]): Declarations {
  const out: Declarations = Object.create(null)
  for (const key of keys) if (decls[key] !== undefined) out[key] = decls[key]
  return out
}

// Declaration objects are memoised, so identical rows/cells share them and the caches below
// turn per-cell style derivation into a lookup.
function pickCached(ctx: ParseContext, decls: Declarations, keys: string[]): Declarations {
  let byDecls = ctx.pickCache.get(keys)
  if (!byDecls) { byDecls = new WeakMap(); ctx.pickCache.set(keys, byDecls) }
  let picked = byDecls.get(decls)
  if (!picked) { picked = pick(decls, keys); byDecls.set(decls, picked) }
  return picked
}

function mergeDecls(ctx: ParseContext, base: Declarations, own: Declarations): Declarations {
  let byOwn = ctx.mergeCache.get(base)
  if (!byOwn) { byOwn = new Map(); ctx.mergeCache.set(base, byOwn) }
  let merged = byOwn.get(own)
  if (!merged) { merged = Object.assign(Object.create(null), base, own) as Declarations; byOwn.set(own, merged) }
  return merged
}

function cloneColor<T>(color: T): T {
  return color && typeof color === 'object' ? { ...color } : color
}

function cloneFont(font: CellFont): CellFont {
  return font.color && typeof font.color === 'object' ? { ...font, color: { ...font.color } } : { ...font }
}

function cloneFill(fill: CellFill): CellFill {
  const out: CellFill = { ...fill }
  if (fill.fgColor) out.fgColor = cloneColor(fill.fgColor)
  if (fill.bgColor) out.bgColor = cloneColor(fill.bgColor)
  return out
}

function cloneBorder(border: CellBorder): CellBorder {
  const out: CellBorder = {}
  for (const side of SIDES) {
    const value = border[side] as CellBorderSide | undefined
    if (value) out[side] = { ...value, color: cloneColor(value.color) }
  }
  return out
}

function presentationalPairs(element: HtmlElement, source: ClipboardSource): Array<[string, string]> {
  const { tag, attrs } = element
  const pairs: Array<[string, string]> = []
  if (attrs.bgcolor) pairs.push(['background-color', attrs.bgcolor])
  if (attrs.valign) pairs.push(['vertical-align', attrs.valign])
  if (tag === 'td' || tag === 'th' || tag === 'tr' || TABLE_SECTIONS.has(tag)) {
    if (attrs.align) pairs.push(['-attr-align', attrs.align])
    if (attrs.nowrap !== undefined) pairs.push(['white-space', 'nowrap'])
  }
  switch (tag) {
    case 'b': case 'strong': case 'h1': case 'h2': case 'h3': case 'h4': case 'h5': case 'h6':
      pairs.push(['font-weight', 'bold'])
      break
    case 'th':
      if (source === 'html' || source === 'word') pairs.push(['font-weight', 'bold'])
      break
    case 'i': case 'em': case 'cite': case 'dfn': case 'var':
      pairs.push(['font-style', 'italic'])
      break
    case 'u': case 'ins':
      pairs.push(['text-decoration', 'underline'])
      break
    case 's': case 'strike': case 'del':
      pairs.push(['text-decoration', 'line-through'])
      break
    case 'sup':
      pairs.push(['vertical-align', 'super'])
      break
    case 'sub':
      pairs.push(['vertical-align', 'sub'])
      break
    case 'font': {
      if (attrs.color) pairs.push(['color', attrs.color])
      if (attrs.face) pairs.push(['font-family', attrs.face])
      const size = Number.parseInt(attrs.size ?? '', 10)
      if (Number.isFinite(size)) {
        const index = /^[+-]/.test(attrs.size ?? '') ? 2 + size : size - 1
        pairs.push(['font-size', `${HTML_FONT_SIZES[Math.min(6, Math.max(0, index))]}pt`])
      }
      break
    }
    default:
      break
  }
  return pairs
}

/** Author-level declarations of one element: presentational hints < stylesheet rules < inline. */
function elementDecls(element: HtmlElement, ctx: ParseContext): Declarations {
  const { attrs } = element
  const key = `${element.tag}\u0001${attrs.class ?? ''}\u0001${attrs.style ?? ''}\u0001${attrs.bgcolor ?? ''}\u0001${attrs.align ?? ''}\u0001${attrs.valign ?? ''}\u0001${attrs.nowrap === undefined ? '' : 'n'}\u0001${attrs.color ?? ''}\u0001${attrs.face ?? ''}\u0001${attrs.size ?? ''}`
  const cached = ctx.cache.get(key)
  if (cached) return cached
  const out: Declarations = Object.create(null)
  const presentational = presentationalPairs(element, ctx.source)
  if (presentational.length) Object.assign(out, normalizeDeclarations(presentational))
  const classes = (attrs.class ?? '').toLowerCase().split(/\s+/).filter(Boolean)
  const ruled = matchedRuleDecls(ctx.sheet, element.tag, classes)
  if (ruled) Object.assign(out, ruled)
  if (attrs.style) Object.assign(out, normalizeDeclarations(declarationPairs(decodeEntities(attrs.style, true))))
  ctx.cache.set(key, out)
  return out
}

function applyFontDecls(decls: Declarations, base: CellFont, ctx: ParseContext, inline: boolean): CellFont {
  let font: CellFont | null = null
  const set = <K extends keyof CellFont>(key: K, value: CellFont[K]) => {
    if (!font) font = { ...base }
    font[key] = value
  }
  const family = primaryFontFamily(decls['font-family'])
  if (family && family !== base.name) set('name', family)
  const rawSize = decls['font-size']?.trim().toLowerCase()
  if (rawSize) {
    const basePt = base.size ?? ctx.defaultFont.size
    let size: number | null = ctx.keywordSizes[rawSize] ?? null
    if (size === null && rawSize === 'larger') size = basePt * 1.2
    if (size === null && rawSize === 'smaller') size = basePt / 1.2
    if (size === null) size = parseLengthPt(rawSize, basePt)
    if (size !== null && size > 0) {
      const rounded = roundHalf(Math.min(409, Math.max(1, size)))
      if (rounded !== base.size) set('size', rounded)
    }
  }
  const weight = decls['font-weight']?.trim().toLowerCase()
  if (weight) {
    const bold = weight === 'bold' || weight === 'bolder' || Number(weight) >= 600
    if (bold !== Boolean(base.bold)) set('bold', bold)
  }
  const style = decls['font-style']?.trim().toLowerCase()
  if (style) {
    const italic = style === 'italic' || style === 'oblique'
    if (italic !== Boolean(base.italic)) set('italic', italic)
  }
  const decoration = decls['text-decoration']?.toLowerCase()
  if (decoration) {
    if (decoration.includes('underline')) {
      const kind = (decls['text-underline-style'] ?? decls['text-decoration-style'] ?? '').toLowerCase()
      const underline: CellFont['underline'] = kind.includes('double-accounting') ? 'doubleAccounting'
        : kind.includes('single-accounting') ? 'singleAccounting'
          : kind.includes('double') ? 'double' : true
      if (underline !== base.underline) set('underline', underline)
    }
    if (decoration.includes('line-through') && !base.strike) set('strike', true)
  }
  const color = parseCssColor(decls.color)
  if (color) {
    const current = typeof base.color === 'object' ? base.color?.argb : undefined
    if (color === '000000' && !base.color) {
      // Black is the grid default; storing it would only add noise.
    } else if (current !== `FF${color}`) set('color', { argb: `FF${color}` })
  }
  if (inline) {
    const vertical = decls['vertical-align']?.trim().toLowerCase()
    if (vertical === 'super' || vertical === 'superscript') set('vertAlign', 'superscript')
    else if (vertical === 'sub' || vertical === 'subscript') set('vertAlign', 'subscript')
  }
  return font ?? base
}

function finalizeFont(font: CellFont | undefined, defaultFont: { name?: string; size: number }): CellFont | undefined {
  if (!font) return undefined
  const out: CellFont = {}
  if (font.name && !(defaultFont.name && font.name.toLowerCase() === defaultFont.name.toLowerCase())) out.name = font.name
  if (font.size && font.size !== defaultFont.size) out.size = font.size
  if (font.bold) out.bold = true
  if (font.italic) out.italic = true
  if (font.underline) out.underline = font.underline
  if (font.strike) out.strike = true
  if (font.color) out.color = font.color
  if (font.vertAlign) out.vertAlign = font.vertAlign
  return Object.keys(out).length ? out : undefined
}

function fontKey(font: CellFont) {
  return JSON.stringify([font.name, font.size, font.bold, font.italic, font.underline, font.strike, typeof font.color === 'object' ? font.color?.argb : font.color, font.vertAlign])
}

function fillFromDecls(decls: Declarations): CellFill | undefined {
  const color = parseCssColor(decls['background-color'])
  if (!color) return undefined
  return { type: 'pattern', pattern: 'solid', fgColor: { argb: `FF${color}` } }
}

function borderSide(decls: Declarations, side: Side): CellBorderSide | undefined {
  const style = (decls[`border-${side}-style`] ?? '').toLowerCase()
  if (!style || style === 'none' || style === 'hidden') return undefined
  const widthRaw = (decls[`border-${side}-width`] ?? '').toLowerCase()
  let weight = 1
  if (widthRaw === 'medium') weight = 2
  else if (widthRaw === 'thick') weight = 3
  else if (widthRaw && widthRaw !== 'thin') {
    const match = /^([+-]?(?:\d+\.?\d*|\.\d+))(px|pt|em|rem|in|cm|mm|pc)?$/.exec(widthRaw)
    if (match) {
      const value = Number.parseFloat(match[1])
      if (value <= 0) return undefined
      if ((match[2] ?? 'px') === 'px') weight = value <= 1.5 ? 1 : value < 2.5 ? 2 : 3
      else {
        const pt = parseLengthPt(widthRaw, 11) ?? 0.75
        weight = pt < 0.9 ? 1 : pt < 1.4 ? 2 : 3
      }
    }
  }
  let modelStyle: string
  switch (style) {
    case 'hairline': modelStyle = 'hair'; break
    case 'dotted': modelStyle = 'dotted'; break
    case 'dashed': modelStyle = weight >= 2 ? 'mediumDashed' : 'dashed'; break
    case 'dot-dash': modelStyle = weight >= 2 ? 'mediumDashDot' : 'dashDot'; break
    case 'dot-dot-dash': modelStyle = weight >= 2 ? 'mediumDashDotDot' : 'dashDotDot'; break
    case 'slanted-dot-dash': modelStyle = 'slantDashDot'; break
    case 'double': modelStyle = 'double'; break
    default: modelStyle = weight >= 3 ? 'thick' : weight === 2 ? 'medium' : 'thin'
  }
  const color = parseCssColor(decls[`border-${side}-color`])
  return { style: modelStyle, color: { argb: `FF${color ?? '000000'}` } }
}

function borderFromDecls(decls: Declarations): CellBorder | undefined {
  let border: CellBorder | undefined
  for (const side of SIDES) {
    const value = borderSide(decls, side)
    if (!value) continue
    border ??= {}
    border[side] = value
  }
  return border
}

function horizontalAlignment(raw: string | undefined): string | undefined {
  switch ((raw ?? '').trim().toLowerCase()) {
    case 'left': case 'start': return 'left'
    case 'right': case 'end': return 'right'
    case 'center': case 'middle': case '-webkit-center': case '-moz-center': return 'center'
    case 'justify': return 'justify'
    case 'center-across': case 'centre-across': return 'centerContinuous'
    case 'fill': return 'fill'
    case 'distributed': return 'distributed'
    default: return undefined
  }
}

function alignmentFromDecls(decls: Declarations, cell: CellData, ctx: ParseContext): CellAlignment | undefined {
  const alignment: CellAlignment = {}
  let horizontal = horizontalAlignment(decls['text-align'])
  const attributeAlign = decls['-attr-align']?.trim().toLowerCase()
  if (!decls['text-align'] && attributeAlign && ctx.source !== 'excel' && ctx.source !== 'simple-calc') {
    const value = cell.formula ? cell.result : cell.value
    // Apps write align="right" for numbers / "left" for text as their general alignment.
    const generalDefault = (attributeAlign === 'right' && typeof value === 'number')
      || (attributeAlign === 'left' && (typeof value === 'string' || value === undefined))
      || (attributeAlign === 'center' && typeof value === 'boolean')
    if (!generalDefault) horizontal = horizontalAlignment(attributeAlign)
  }
  if (horizontal) alignment.horizontal = horizontal
  switch ((decls['vertical-align'] ?? '').trim().toLowerCase()) {
    case 'top': case 'text-top': alignment.vertical = 'top'; break
    case 'middle': case 'center': alignment.vertical = 'middle'; break
    case 'justify': alignment.vertical = 'justify'; break
    case 'distributed': alignment.vertical = 'distributed'; break
    default: break
  }
  const whiteSpace = (decls['white-space'] ?? '').trim().toLowerCase()
  const breakWord = /break-word|anywhere/.test(`${decls['overflow-wrap'] ?? ''} ${decls['word-wrap'] ?? ''}`.toLowerCase())
  if (whiteSpace === 'normal' || whiteSpace === 'pre-wrap' || whiteSpace === 'pre-line' || whiteSpace === 'break-spaces' || (breakWord && whiteSpace !== 'nowrap')) alignment.wrapText = true
  const indent = Number.parseInt(decls['mso-char-indent-count'] ?? '', 10)
  if (Number.isFinite(indent) && indent > 0) alignment.indent = Math.min(250, indent)
  const rotation = Number.parseFloat(decls['mso-rotate'] ?? '')
  if (Number.isFinite(rotation) && rotation !== 0) alignment.textRotation = Math.max(-90, Math.min(90, Math.round(rotation)))
  if (/vertical/.test((decls['layout-flow'] ?? '').toLowerCase())) alignment.textRotation = 'vertical'
  if (/shrinktofit/.test((decls['mso-text-control'] ?? '').toLowerCase())) alignment.shrinkToFit = true
  return Object.keys(alignment).length ? alignment : undefined
}

function sanitizeHyperlink(raw: string | undefined): string | undefined {
  if (!raw) return undefined
  const value = raw.trim().replace(/[\u0000-\u001f\u007f]/g, '')
  if (!value || value.length > 2083) return undefined
  if (/^(https?:|mailto:|ftp:|tel:)/i.test(value)) return value
  if (/^#'?[^!]+'?![A-Z$]+\$?\d+/i.test(value)) return value
  return undefined
}

function notePlainText(note: CellData['note']): string {
  if (!note) return ''
  if (typeof note === 'string') return note
  if (note.text) return note.text
  if (note.texts?.length) return note.texts.map((part) => part.text ?? '').join('')
  if (note.comments?.length) return note.comments.map((comment) => comment.text ?? '').join('\n')
  return ''
}

interface TextRun { text: string; font: CellFont }

class RunBuilder {
  runs: TextRun[] = []
  private lineStart = true
  private pendingSpace = false

  text(raw: string, font: CellFont, preserve: boolean) {
    if (!raw) return
    if (preserve) {
      const text = raw.replace(/\r\n?/g, '\n').replace(/\u00A0/g, ' ')
      if (!text) return
      if (this.pendingSpace) { this.push(' ', font); this.pendingSpace = false }
      this.push(text, font)
      this.lineStart = text.endsWith('\n')
      return
    }
    let text = raw.replace(/[\t\n\r\f ]+/g, ' ')
    if (text.startsWith(' ')) {
      text = text.slice(1)
      if (!this.lineStart) this.pendingSpace = true
    }
    if (!text) return
    const trailing = text.endsWith(' ')
    if (trailing) text = text.slice(0, -1)
    if (!text) return
    if (this.pendingSpace) {
      // A collapsed space belongs to the run it followed (that is how it renders).
      const last = this.runs[this.runs.length - 1]
      if (last && !last.text.endsWith('\n')) last.text += ' '
      else text = ` ${text}`
    }
    this.push(text.replace(/\u00A0/g, ' '), font)
    this.lineStart = false
    this.pendingSpace = trailing
  }

  newline(font: CellFont) {
    this.pendingSpace = false
    this.push('\n', font)
    this.lineStart = true
  }

  block(font: CellFont) {
    if (!this.lineStart) this.newline(font)
  }

  space() {
    if (!this.lineStart) this.pendingSpace = true
  }

  private push(text: string, font: CellFont) {
    const last = this.runs[this.runs.length - 1]
    if (last && last.font === font) last.text += text
    else this.runs.push({ text, font })
  }
}

interface CellContent {
  text: string
  runs: TextRun[]
  uniformFont?: CellFont
  hyperlink?: string
  note?: string
}

function collectCellContent(cellElement: HtmlElement, baseFont: CellFont, ctx: ParseContext, preserveEdges: boolean): CellContent {
  const builder = new RunBuilder()
  let hyperlink: string | undefined
  let note: string | undefined
  const baseDecls = elementDecls(cellElement, ctx)
  const basePreserve = /^pre/.test((baseDecls['white-space'] ?? '').toLowerCase())

  const visit = (node: HtmlElement, font: CellFont, preserve: boolean, depth: number) => {
    for (const child of node.children) {
      if (typeof child === 'string') { builder.text(child, font, preserve); continue }
      const tag = child.tag
      if (tag === 'br') { builder.newline(font); continue }
      if (tag === 'comment') {
        const text = textContent(child).trim()
        if (text && !note) note = text
        continue
      }
      if (SKIP_CONTENT_TAGS.has(tag) || depth > 64) continue
      if ((tag === 'td' || tag === 'th') && depth > 0) builder.space()
      const decls = elementDecls(child, ctx)
      if (/none/.test(decls.display ?? '') || decls['mso-hide'] === 'all') continue
      const childFont = applyFontDecls(decls, font, ctx, true)
      const whiteSpace = (decls['white-space'] ?? '').toLowerCase()
      const childPreserve = whiteSpace ? /^pre/.test(whiteSpace) : preserve || /yes/i.test(decls['mso-spacerun'] ?? '')
      if (tag === 'a' && !hyperlink) hyperlink = sanitizeHyperlink(child.attrs.href)
      const block = BLOCK_TAGS.has(tag)
      if (block) builder.block(font)
      visit(child, childFont, childPreserve, depth + 1)
      if (block) builder.block(font)
    }
  }
  visit(cellElement, baseFont, basePreserve, 0)

  let runs = builder.runs.filter((run) => run.text.length > 0)
  const edge = preserveEdges ? /^\n+/ : /^\s+/
  const edgeEnd = preserveEdges ? /\n+$/ : /\s+$/
  while (runs.length && !runs[0].text.replace(edge, '')) runs.shift()
  while (runs.length && !runs[runs.length - 1].text.replace(edgeEnd, '')) runs.pop()
  if (runs.length) {
    runs[0] = { ...runs[0], text: runs[0].text.replace(edge, '') }
    const last = runs.length - 1
    runs[last] = { ...runs[last], text: runs[last].text.replace(edgeEnd, '') }
  }
  let text = runs.map((run) => run.text).join('')
  if (!text.trim()) { text = ''; runs = [] }
  if (text.length > MAX_TEXT_LENGTH) text = text.slice(0, MAX_TEXT_LENGTH)

  const distinct = new Map<string, CellFont>()
  for (const run of runs) if (run.text.trim()) distinct.set(fontKey(run.font), run.font)
  if (distinct.size <= 1) {
    return { text, runs: [], uniformFont: distinct.values().next().value ?? baseFont, hyperlink, note }
  }
  const merged: TextRun[] = []
  for (const run of runs) {
    const last = merged[merged.length - 1]
    if (last && fontKey(last.font) === fontKey(run.font)) last.text += run.text
    else merged.push({ text: run.text, font: run.font })
  }
  return { text, runs: merged, hyperlink, note }
}

function textContent(element: HtmlElement): string {
  let out = ''
  for (const child of element.children) {
    if (typeof child === 'string') out += child
    else if (child.tag === 'br') out += '\n'
    else if (!SKIP_CONTENT_TAGS.has(child.tag)) {
      const inner = textContent(child)
      out += BLOCK_TAGS.has(child.tag) && out && !out.endsWith('\n') ? `\n${inner}` : inner
    }
  }
  return out
}

const INTERNAL_CELL_KEYS = new Set(['value', 'formula', 'formulaType', 'formulaRange', 'dynamicFormula', 'result', 'resultType', 'richText', 'style', 'numFmt', 'hyperlink', 'hyperlinkTooltip', 'note', 'type'])
const UNSAFE_KEYS = new Set(['__proto__', 'constructor', 'prototype'])

function cloneJson(value: unknown, depth = 0): unknown {
  if (depth > 8) return undefined
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined
  if (Array.isArray(value)) return value.slice(0, 4096).map((item) => cloneJson(item, depth + 1))
  if (typeof value === 'object') {
    const out: Record<string, unknown> = {}
    for (const [key, item] of Object.entries(value as Record<string, unknown>)) {
      if (UNSAFE_KEYS.has(key)) continue
      const cloned = cloneJson(item, depth + 1)
      if (cloned !== undefined) out[key] = cloned
    }
    return out
  }
  return undefined
}

function isScalar(value: unknown): value is CellScalar {
  return value === null || typeof value === 'string' || typeof value === 'boolean' || (typeof value === 'number' && Number.isFinite(value))
}

/** Validates the JSON cell simple_calc embeds in its own clipboard HTML. */
function sanitizeInternalCell(raw: unknown): CellData | null {
  if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return null
  const source = raw as Record<string, unknown>
  const cell: CellData = {}
  for (const key of Object.keys(source)) {
    if (!INTERNAL_CELL_KEYS.has(key)) continue
    const value = source[key]
    switch (key) {
      case 'value': case 'result':
        if (isScalar(value)) cell[key] = typeof value === 'string' ? value.slice(0, MAX_TEXT_LENGTH) : value
        break
      case 'formula':
        if (typeof value === 'string' && value.length <= MAX_FORMULA_LENGTH) cell.formula = value
        break
      case 'dynamicFormula':
        if (typeof value === 'boolean') cell.dynamicFormula = value
        break
      case 'hyperlink': {
        const link = typeof value === 'string' ? sanitizeHyperlink(value) : undefined
        if (link) cell.hyperlink = link
        break
      }
      case 'note':
        if (typeof value === 'string') cell.note = value.slice(0, MAX_TEXT_LENGTH)
        else if (value && typeof value === 'object') cell.note = cloneJson(value) as CellData['note']
        break
      case 'richText':
        if (Array.isArray(value)) {
          const runs = value.filter((run) => run && typeof run === 'object' && typeof (run as { text?: unknown }).text === 'string')
            .map((run) => cloneJson(run) as { text: string; font?: CellFont })
          if (runs.length) cell.richText = runs
        }
        break
      case 'style':
        if (value && typeof value === 'object' && !Array.isArray(value)) cell.style = cloneJson(value) as CellStyle
        break
      default:
        if (typeof value === 'string' && value.length < 512) (cell as Record<string, unknown>)[key] = value
    }
  }
  return cell
}

interface TypedValue { value?: CellScalar; numFmt?: string; type?: string; error?: boolean }

function typedValue(attrs: Record<string, string>, text: string, numFmt: string | undefined, ctx: ParseContext): TypedValue {
  const dateType = numFmt && isDateNumberFormat(numFmt) && !isTimeOnlyFormat(numFmt) ? 'date' : undefined
  const sheets = parseJson(attrs['data-sheets-value']) as Record<string, unknown> | undefined
  if (sheets && typeof sheets === 'object') {
    const kind = Number(sheets['1'])
    if (kind === 2 && typeof sheets['2'] === 'string') return { value: sheets['2'].slice(0, MAX_TEXT_LENGTH) }
    if (kind === 3 && typeof sheets['3'] === 'number' && Number.isFinite(sheets['3'])) return { value: sheets['3'], type: dateType }
    if (kind === 4) return { value: sheets['4'] === true || Number(sheets['4']) === 1 }
  }
  if (attrs.sdval !== undefined && attrs.sdval.trim() !== '') {
    const number = Number(attrs.sdval)
    if (Number.isFinite(number)) {
      if (/;\s*boolean\s*$/i.test(attrs.sdnum ?? '')) return { value: number !== 0 }
      return { value: number, type: dateType }
    }
  }
  if (hasOwn(attrs, 'x:bool')) return { value: /^(true|1)$/i.test((attrs['x:bool'] || text).trim()) }
  if (hasOwn(attrs, 'x:err')) return { value: (attrs['x:err'] || text).trim(), error: true }
  if (hasOwn(attrs, 'x:num')) {
    const raw = attrs['x:num'].trim()
    const number = raw ? Number(raw) : (() => {
      const smart = parseSmartValue(text, ctx.smart)
      return typeof smart?.value === 'number' ? smart.value : NaN
    })()
    if (Number.isFinite(number)) return { value: number, type: dateType }
  }
  if (hasOwn(attrs, 'x:str')) return { value: (attrs['x:str'] || text).slice(0, MAX_TEXT_LENGTH) }
  if (numFmt === '@') return text ? { value: text } : {}
  if (!text) return {}
  const dayFirst = isDayFirstFormat(numFmt) ?? ctx.smart.dayFirst
  const smart = parseSmartValue(text, { ...ctx.smart, dayFirst })
  if (!smart || typeof smart.value === 'string') return { value: text }
  return { value: smart.value, numFmt: smart.numFmt, type: smart.type ?? (typeof smart.value === 'number' ? dateType : undefined) }
}

function buildCell(td: HtmlElement, inherited: Declarations, row: number, col: number, ctx: ParseContext): CellData {
  const attrs = td.attrs
  if (ctx.source === 'simple-calc' && attrs['data-sc']) {
    const internal = sanitizeInternalCell(parseJson(attrs['data-sc']))
    if (internal) return internal
  }
  const decls = mergeDecls(ctx, inherited, elementDecls(td, ctx))
  let info = ctx.infoCache.get(decls)
  if (!info) {
    const font = applyFontDecls(decls, {}, ctx, false)
    info = {
      font,
      finalFont: finalizeFont(font, ctx.defaultFont),
      fill: fillFromDecls(decls),
      border: borderFromDecls(decls),
      alignment: decls['-attr-align'] && ctx.source !== 'excel' && ctx.source !== 'simple-calc' ? null : alignmentFromDecls(decls, {}, ctx),
      msoFormat: msoNumberFormat(decls['mso-number-format']),
    }
    ctx.infoCache.set(decls, info)
  }
  const cellFont = info.font
  const content = collectCellContent(td, cellFont, ctx, ctx.source === 'excel' || ctx.source === 'simple-calc')
  const cell: CellData = {}

  const numFmt = sheetsNumberFormat(attrs['data-sheets-numberformat'])
    ?? info.msoFormat
    ?? libreOfficeNumberFormat(attrs.sdnum)
  const typed = typedValue(attrs, content.text, numFmt, ctx)

  let formula: string | undefined
  const sheetsFormula = attrs['data-sheets-formula']?.trim()
  const excelFormula = attrs['x:fmla']?.trim()
  if (sheetsFormula && sheetsFormula.startsWith('=') && sheetsFormula.length <= MAX_FORMULA_LENGTH) {
    formula = r1c1ToA1(sheetsFormula, ctx.formulaOrigin.row + row, ctx.formulaOrigin.col + col).slice(1)
  } else if (excelFormula && excelFormula.startsWith('=') && excelFormula.length <= MAX_FORMULA_LENGTH) {
    formula = excelFormula.slice(1)
  }
  if (formula) {
    cell.formula = formula
    if (typed.value !== undefined) cell.result = typed.value
    if (typed.error) cell.resultType = 'error'
  } else if (typed.value !== undefined) {
    cell.value = typed.value
    if (typed.type) cell.type = typed.type
  }
  const format = numFmt ?? typed.numFmt
  if (format) cell.numFmt = format

  let finalFont = info.finalFont
  if (content.runs.length > 1 && typeof cell.value === 'string') {
    cell.richText = content.runs.map((run) => {
      const runFont = finalizeFont(run.font, ctx.defaultFont)
      return runFont ? { text: run.text, font: runFont } : { text: run.text }
    })
  } else if (content.uniformFont && content.uniformFont !== cellFont) {
    finalFont = finalizeFont(content.uniformFont, ctx.defaultFont)
  }
  const style: CellStyle = {}
  if (finalFont) style.font = cloneFont(finalFont)
  if (info.fill) style.fill = cloneFill(info.fill)
  if (info.border) style.border = cloneBorder(info.border)
  const alignment = info.alignment === null ? alignmentFromDecls(decls, cell, ctx) : info.alignment ? { ...info.alignment } : undefined
  if (alignment) style.alignment = alignment
  if (content.text.includes('\n') && !alignment?.wrapText) style.alignment = { ...(style.alignment ?? {}), wrapText: true }
  if (Object.keys(style).length) cell.style = style

  const hyperlink = content.hyperlink ?? sanitizeHyperlink(attrs['data-sheets-hyperlink'])
  if (hyperlink) cell.hyperlink = hyperlink
  const note = attrs['data-sheets-note'] ?? content.note
  if (note && note.trim()) cell.note = note.slice(0, MAX_TEXT_LENGTH)
  return cell
}

function clampSpan(raw: string | undefined, max: number) {
  const value = Number.parseInt(raw ?? '', 10)
  if (!Number.isFinite(value) || value < 1) return 1
  return Math.min(value, max)
}

function elementWidthPx(element: HtmlElement, decls: Declarations | null) {
  const attribute = element.attrs.width
  if (attribute && !attribute.trim().endsWith('%')) {
    const value = Number.parseFloat(attribute)
    if (Number.isFinite(value) && value > 0) return Math.round(value * 100) / 100
  }
  return parseLengthPx(decls?.width)
}

function elementHeightPx(element: HtmlElement, decls: Declarations | null) {
  const attribute = element.attrs.height
  if (attribute) {
    const value = Number.parseFloat(attribute)
    if (Number.isFinite(value) && value > 0) return Math.round(value * 100) / 100
  }
  return parseLengthPx(decls?.height)
}

interface TableLayout {
  grid: Array<Array<CellData | null | undefined>>
  merges: ClipboardBounds[]
  widths: Array<number | undefined>
  heights: Array<number | undefined>
  width: number
}

function layoutTable(table: HtmlElement, ctx: ParseContext, rowOffset: number, cellBudget: { used: number }): TableLayout {
  const tableDecls = pickCached(ctx, elementDecls(table, ctx), INHERITED_PROPERTIES)
  const rows: Array<{ tr: HtmlElement; inherited: Declarations }> = []
  const widths: Array<number | undefined> = []
  let columnCursor = 0
  let widthsFromColumns = false
  const addColumn = (col: HtmlElement, fallback?: HtmlElement) => {
    const span = clampSpan(col.attrs.span ?? fallback?.attrs.span, 1000)
    const width = elementWidthPx(col, elementDecls(col, ctx)) ?? (fallback ? elementWidthPx(fallback, elementDecls(fallback, ctx)) : undefined)
    for (let index = 0; index < span && columnCursor < SHEET_COLS; index += 1) {
      widths[columnCursor] = width
      columnCursor += 1
    }
    if (width !== undefined) widthsFromColumns = true
  }
  const visitSection = (section: HtmlElement, inherited: Declarations) => {
    for (const child of elementChildren(section)) {
      if (child.tag === 'tr') rows.push({ tr: child, inherited })
      else if (TABLE_SECTIONS.has(child.tag)) visitSection(child, mergeDecls(ctx, inherited, pickCached(ctx, elementDecls(child, ctx), ROW_INHERITED_PROPERTIES)))
      else if (child.tag === 'colgroup') {
        const cols = elementChildren(child).filter((item) => item.tag === 'col')
        if (cols.length) cols.forEach((col) => addColumn(col, child))
        else addColumn(child)
      } else if (child.tag === 'col') addColumn(child)
      else if (child.tag === 'form' || child.tag === 'google-sheets-html-origin') visitSection(child, inherited)
    }
  }
  visitSection(table, tableDecls)

  const grid: Array<Array<CellData | null | undefined>> = []
  const heights: Array<number | undefined> = []
  const merges: ClipboardBounds[] = []
  let width = 0
  const ensureRow = (row: number) => { while (grid.length <= row) grid.push([]) }
  rows.forEach(({ tr, inherited }, rowIndex) => {
    if (rowOffset + rowIndex >= SHEET_ROWS) throw new ClipboardTooLargeError()
    ensureRow(rowIndex)
    const trDecls = elementDecls(tr, ctx)
    heights[rowIndex] = elementHeightPx(tr, trDecls)
    const rowInherited = mergeDecls(ctx, inherited, pickCached(ctx, trDecls, ROW_INHERITED_PROPERTIES))
    const line = grid[rowIndex]
    let col = 0
    for (const td of elementChildren(tr)) {
      if (td.tag !== 'td' && td.tag !== 'th') continue
      while (line[col] !== undefined) col += 1
      const colSpan = clampSpan(td.attrs.colspan, 1000)
      const rowSpan = clampSpan(td.attrs.rowspan, Math.max(1, rows.length - rowIndex))
      const right = col + colSpan
      if (right > SHEET_COLS) throw new ClipboardTooLargeError()
      const nextWidth = Math.max(width, right)
      const areaRows = Math.max(grid.length, rowIndex + rowSpan)
      if (cellBudget.used + areaRows * nextWidth > ctx.maxCells) throw new ClipboardTooLargeError()
      line[col] = buildCell(td, rowInherited, rowOffset + rowIndex, col, ctx)
      for (let r = rowIndex; r < rowIndex + rowSpan; r += 1) {
        ensureRow(r)
        for (let c = col; c < right; c += 1) if (r !== rowIndex || c !== col) grid[r][c] = null
      }
      // Excel writes overflowing text as colspan + mso-ignore:colspan; that is not a merge.
      const ignored = (elementDecls(td, ctx)['mso-ignore'] ?? '').toLowerCase()
      const mergeCols = ignored.includes('colspan') ? 1 : colSpan
      const mergeRows = ignored.includes('rowspan') ? 1 : rowSpan
      if (mergeCols > 1 || mergeRows > 1) merges.push({ top: rowIndex, left: col, bottom: rowIndex + mergeRows - 1, right: col + mergeCols - 1 })
      if (!widthsFromColumns && colSpan === 1 && widths[col] === undefined) widths[col] = elementWidthPx(td, elementDecls(td, ctx))
      width = nextWidth
      col = right
    }
  })
  return { grid, merges, widths, heights, width }
}

function detectSource(root: HtmlElement): ClipboardSource {
  let source: ClipboardSource = 'html'
  walkElements(root, (element) => {
    const { tag, attrs } = element
    if (tag === 'table' && attrs['data-sc-clip'] !== undefined) { source = 'simple-calc'; return false }
    if (tag === 'google-sheets-html-origin' || attrs['data-sheets-root'] !== undefined) { if (source === 'html') source = 'google-sheets' }
    if (tag === 'meta') {
      const name = (attrs.name ?? '').toLowerCase()
      const content = (attrs.content ?? '').toLowerCase()
      if (name === 'generator' || name === 'progid') {
        if (/excel/.test(content)) source = 'excel'
        else if (/libreoffice|openoffice|staroffice/.test(content)) source = 'libreoffice'
        else if (/word/.test(content) && source === 'html') source = 'word'
      }
    }
    if (tag === 'html' && attrs['xmlns:x'] && /office:excel/i.test(attrs['xmlns:x']) && source === 'html') source = 'excel'
    return undefined
  })
  return source
}

interface Segment { kind: 'table'; table: HtmlElement }
interface TextSegment { kind: 'text'; lines: string[] }

function collectSegments(root: HtmlElement): Array<Segment | TextSegment> {
  const segments: Array<Segment | TextSegment> = []
  let buffer = ''
  const flush = () => {
    const lines = buffer.split('\n').map((line) => line.replace(/[ \t\u00A0]+/g, ' ').trim()).filter(Boolean)
    if (lines.length) segments.push({ kind: 'text', lines })
    buffer = ''
  }
  const containsTable = (element: HtmlElement) => {
    let found = false
    walkElements(element, (item) => {
      if (found) return false
      if (item !== element && item.tag === 'table') { found = true; return false }
      return undefined
    })
    return found
  }
  const visit = (element: HtmlElement) => {
    for (const child of element.children) {
      if (typeof child === 'string') { buffer += child.replace(/[\r\n\t]+/g, ' '); continue }
      if (child.tag === 'table') { flush(); segments.push({ kind: 'table', table: child }); continue }
      if (child.tag === 'br') { buffer += '\n'; continue }
      if (SKIP_CONTENT_TAGS.has(child.tag) && child.tag !== 'head') continue
      if (child.tag === 'head') continue
      const block = BLOCK_TAGS.has(child.tag)
      if (block) buffer += '\n'
      if (containsTable(child)) visit(child)
      else buffer += textContent(child).replace(/[\r\t]+/g, ' ')
      if (block) buffer += '\n'
    }
  }
  visit(root)
  flush()
  return segments
}

/**
 * Parses clipboard HTML into cells. Returns null when the HTML holds no table (paste the
 * text flavour instead). Throws `ClipboardTooLargeError` beyond `maxCells`.
 */
export function parseClipboardHtml(html: string, options: ClipboardParseOptions = {}): ParsedClipboard | null {
  if (!html || !/<table[\s>]/i.test(html)) return null
  const mode = options.parser ?? 'auto'
  const root = (mode !== 'builtin' ? parseHtmlWithDom(html) : null) ?? (mode === 'dom' ? null : parseHtml(html))
  if (!root) return null

  const source = detectSource(root)
  const sheet = emptyStyleSheet()
  walkElements(root, (element) => {
    if (element.tag === 'style') addStyleSheet(sheet, element.children.filter((child): child is string => typeof child === 'string').join(''), source === 'google-sheets')
  })
  const segments = collectSegments(root)
  const tables = segments.filter((segment): segment is Segment => segment.kind === 'table')
  if (!tables.length) return null

  let formulaOrigin = options.origin ?? { row: 0, col: 0 }
  let sourceSheetName: string | undefined
  if (source === 'simple-calc') {
    const table = tables[0].table
    const origin = rangeToBounds(table.attrs['data-sc-origin'] ?? '')
    if (origin) formulaOrigin = { row: origin.top, col: origin.left }
    if (table.attrs['data-sc-sheet']) sourceSheetName = table.attrs['data-sc-sheet'].slice(0, 255)
  }
  const defaultSize = options.defaultFont?.size ?? 11
  const ctx: ParseContext = {
    source,
    sheet,
    cache: new Map(),
    mergeCache: new WeakMap(),
    pickCache: new Map(),
    infoCache: new WeakMap(),
    defaultFont: { name: options.defaultFont?.name ?? 'Calibri', size: defaultSize },
    formulaOrigin,
    smart: { decimalSeparator: options.decimalSeparator, dayFirst: options.dayFirst },
    keywordSizes: source === 'libreoffice' ? LIBREOFFICE_FONT_KEYWORDS : BROWSER_FONT_KEYWORDS,
    maxCells: options.maxCells ?? CLIPBOARD_MAX_CELLS,
  }

  const matrix: Array<Array<CellData | null | undefined>> = []
  const merges: string[] = []
  let widths: Array<number | undefined> = []
  const heights: Array<number | undefined> = []
  let width = 0
  const budget = { used: 0 }
  const onlyTable = segments.length === 1
  for (const segment of segments) {
    const top = matrix.length
    if (segment.kind === 'text') {
      if (budget.used + segment.lines.length * Math.max(1, width) > ctx.maxCells) throw new ClipboardTooLargeError()
      for (const line of segment.lines) {
        matrix.push([parsePastedValue(line, { ...options, formulas: false })])
        heights.push(undefined)
      }
      width = Math.max(width, 1)
      budget.used = matrix.length * width
      continue
    }
    const layout = layoutTable(segment.table, ctx, top, budget)
    layout.grid.forEach((line) => matrix.push(line))
    layout.heights.forEach((height) => heights.push(height))
    for (let index = layout.heights.length; index < layout.grid.length; index += 1) heights.push(undefined)
    layout.merges.forEach((merge) => merges.push(boundsToRange({ top: merge.top + top, bottom: merge.bottom + top, left: merge.left, right: merge.right })))
    if (!widths.length || onlyTable) widths = layout.widths
    width = Math.max(width, layout.width)
    budget.used = matrix.length * width
  }
  if (!matrix.length || !width) return null
  if (matrix.length * width > ctx.maxCells) throw new ClipboardTooLargeError()

  const cells: CellData[][] = matrix.map((line) => {
    const out: CellData[] = new Array(width)
    for (let col = 0; col < width; col += 1) out[col] = line[col] ?? {}
    return out
  })
  const result: ParsedClipboard = { cells, merges, source, formulaOrigin }
  const columnWidths = Array.from({ length: width }, (_, index) => widths[index])
  if (columnWidths.some((value) => value !== undefined)) result.columnWidths = columnWidths
  const rowHeights = Array.from({ length: cells.length }, (_, index) => heights[index])
  if (rowHeights.some((value) => value !== undefined)) result.rowHeights = rowHeights
  if (sourceSheetName) result.sourceSheetName = sourceSheetName
  return result
}

// ---------------------------------------------------------------------------------------------
// Excel "XML Spreadsheet" clipboard format (SpreadsheetML 2003) — carries formulas
// ---------------------------------------------------------------------------------------------

const localName = (name: string) => name.slice(name.lastIndexOf(':') + 1)

function xmlAttr(element: HtmlElement, name: string): string | undefined {
  const { attrs } = element
  return attrs[`ss:${name}`] ?? attrs[name] ?? attrs[`x:${name}`] ?? attrs[`html:${name}`]
}

function xmlChildren(element: HtmlElement, name: string): HtmlElement[] {
  const out: HtmlElement[] = []
  for (const child of element.children) if (typeof child !== 'string' && localName(child.tag) === name) out.push(child)
  return out
}

function xmlText(element: HtmlElement): string {
  let out = ''
  for (const child of element.children) out += typeof child === 'string' ? child : xmlText(child)
  return out
}

function xmlColor(raw: string | undefined): SpreadsheetColorArgb | undefined {
  const hex = parseCssColor(raw)
  return hex ? { argb: `FF${hex}` } : undefined
}
type SpreadsheetColorArgb = { argb: string }

const XML_HORIZONTAL: Record<string, string> = { left: 'left', center: 'center', right: 'right', fill: 'fill', justify: 'justify', centeracrossselection: 'centerContinuous', distributed: 'distributed', justifydistributed: 'distributed' }
const XML_VERTICAL: Record<string, string> = { top: 'top', center: 'middle', justify: 'justify', distributed: 'distributed', justifydistributed: 'distributed' }
const XML_UNDERLINE: Record<string, CellFont['underline']> = { single: true, double: 'double', singleaccounting: 'singleAccounting', doubleaccounting: 'doubleAccounting' }

function xmlBorderStyle(lineStyle: string, weight: number): string | undefined {
  switch (lineStyle.toLowerCase()) {
    case 'continuous': return weight <= 0 ? 'hair' : weight === 1 ? 'thin' : weight === 2 ? 'medium' : 'thick'
    case 'dash': return weight >= 2 ? 'mediumDashed' : 'dashed'
    case 'dot': return 'dotted'
    case 'dashdot': return weight >= 2 ? 'mediumDashDot' : 'dashDot'
    case 'dashdotdot': return weight >= 2 ? 'mediumDashDotDot' : 'dashDotDot'
    case 'slantdashdot': return 'slantDashDot'
    case 'double': return 'double'
    default: return undefined
  }
}

interface XmlStyle { font: CellFont; fill?: CellFill; border?: CellBorder; alignment?: CellAlignment; numFmt?: string }

function applyXmlFont(element: HtmlElement, base: CellFont): CellFont {
  const font: CellFont = { ...base }
  const name = xmlAttr(element, 'fontname') ?? xmlAttr(element, 'face')
  if (name) font.name = name.slice(0, 100)
  const size = Number.parseFloat(xmlAttr(element, 'size') ?? '')
  if (Number.isFinite(size) && size > 0) font.size = roundHalf(Math.min(409, size))
  const flag = (key: string) => {
    const value = xmlAttr(element, key)
    return value === undefined ? undefined : value === '1' || value.toLowerCase() === 'true'
  }
  const bold = flag('bold')
  if (bold !== undefined) font.bold = bold
  const italic = flag('italic')
  if (italic !== undefined) font.italic = italic
  const strike = flag('strikethrough')
  if (strike !== undefined) font.strike = strike
  const underline = xmlAttr(element, 'underline')
  if (underline) font.underline = XML_UNDERLINE[underline.toLowerCase()] ?? false
  const vertical = (xmlAttr(element, 'verticalalign') ?? '').toLowerCase()
  if (vertical === 'superscript' || vertical === 'subscript') font.vertAlign = vertical
  const color = xmlAttr(element, 'color')
  if (color) {
    const parsed = xmlColor(color)
    if (parsed && parsed.argb !== 'FF000000') font.color = parsed
    else if (parsed) delete font.color
  }
  return font
}

function resolveXmlStyles(workbook: HtmlElement): Map<string, XmlStyle> {
  const raw = new Map<string, HtmlElement>()
  for (const styles of xmlChildren(workbook, 'styles')) {
    for (const style of xmlChildren(styles, 'style')) {
      const id = xmlAttr(style, 'id')
      if (id) raw.set(id, style)
    }
  }
  const resolved = new Map<string, XmlStyle>()
  const resolve = (id: string, depth: number): XmlStyle => {
    const done = resolved.get(id)
    if (done) return done
    const element = raw.get(id)
    const parentId = element ? xmlAttr(element, 'parent') ?? (id === 'Default' ? undefined : 'Default') : undefined
    const parent: XmlStyle = parentId && parentId !== id && depth < 16 && raw.has(parentId) ? resolve(parentId, depth + 1) : { font: {} }
    const style: XmlStyle = {
      font: { ...parent.font },
      fill: parent.fill,
      border: parent.border ? { ...parent.border } : undefined,
      alignment: parent.alignment ? { ...parent.alignment } : undefined,
      numFmt: parent.numFmt,
    }
    if (element) {
      for (const child of element.children) {
        if (typeof child === 'string') continue
        switch (localName(child.tag)) {
          case 'font':
            style.font = applyXmlFont(child, style.font)
            break
          case 'interior': {
            const color = xmlColor(xmlAttr(child, 'color'))
            const pattern = (xmlAttr(child, 'pattern') ?? '').toLowerCase()
            style.fill = color && pattern !== 'none' ? { type: 'pattern', pattern: 'solid', fgColor: color } : undefined
            break
          }
          case 'borders': {
            const border: CellBorder = {}
            for (const item of xmlChildren(child, 'border')) {
              const position = (xmlAttr(item, 'position') ?? '').toLowerCase()
              if (!['top', 'bottom', 'left', 'right'].includes(position)) continue
              const modelStyle = xmlBorderStyle(xmlAttr(item, 'linestyle') ?? '', Number(xmlAttr(item, 'weight') ?? 1))
              if (modelStyle) border[position] = { style: modelStyle, color: xmlColor(xmlAttr(item, 'color')) ?? { argb: 'FF000000' } }
            }
            style.border = Object.keys(border).length ? border : undefined
            break
          }
          case 'alignment': {
            const alignment: CellAlignment = {}
            const horizontal = XML_HORIZONTAL[(xmlAttr(child, 'horizontal') ?? '').toLowerCase()]
            if (horizontal) alignment.horizontal = horizontal
            const vertical = XML_VERTICAL[(xmlAttr(child, 'vertical') ?? '').toLowerCase()]
            if (vertical) alignment.vertical = vertical
            if (xmlAttr(child, 'wraptext') === '1') alignment.wrapText = true
            if (xmlAttr(child, 'shrinktofit') === '1') alignment.shrinkToFit = true
            const indent = Number(xmlAttr(child, 'indent'))
            if (Number.isFinite(indent) && indent > 0) alignment.indent = Math.min(250, indent)
            const rotate = Number(xmlAttr(child, 'rotate'))
            if (Number.isFinite(rotate) && rotate !== 0) alignment.textRotation = Math.max(-90, Math.min(90, rotate))
            if (xmlAttr(child, 'verticaltext') === '1') alignment.textRotation = 'vertical'
            style.alignment = Object.keys(alignment).length ? alignment : undefined
            break
          }
          case 'numberformat': {
            const format = xmlAttr(child, 'format')
            if (format === undefined) break
            const named = EXCEL_NAMED_FORMATS[format.trim().toLowerCase()]
            style.numFmt = format.trim().toLowerCase() === 'general number' ? undefined : named !== undefined ? named || undefined : sanitizeNumFmt(format)
            break
          }
          default:
            break
        }
      }
    }
    resolved.set(id, style)
    return style
  }
  for (const id of raw.keys()) resolve(id, 0)
  if (!resolved.has('Default')) resolved.set('Default', { font: {} })
  return resolved
}

function xmlDateTime(text: string): number | undefined {
  const match = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?)?/.exec(text.trim())
  if (!match) return undefined
  const [year, month, day] = [Number(match[1]), Number(match[2]), Number(match[3])]
  let serial: number | null
  if (year === 1899 && month === 12 && day === 31) serial = 0
  else {
    serial = serialFromDate(year, month, day)
    // Excel's serials before 1 Mar 1900 are one lower (the 1900 leap-year bug).
    if (serial !== null && serial < 61) serial -= 1
  }
  if (serial === null) return undefined
  const seconds = Number(match[4] ?? 0) * 3600 + Number(match[5] ?? 0) * 60 + Number(match[6] ?? 0) + Number(`0.${match[7] ?? '0'}`)
  return serial + seconds / 86_400
}

function xmlRichRuns(data: HtmlElement, base: CellFont): TextRun[] {
  const runs: TextRun[] = []
  const visit = (element: HtmlElement, font: CellFont) => {
    for (const child of element.children) {
      if (typeof child === 'string') {
        if (!child) continue
        const last = runs[runs.length - 1]
        if (last && last.font === font) last.text += child
        else runs.push({ text: child, font })
        continue
      }
      let next: CellFont = font
      switch (localName(child.tag)) {
        case 'b': next = { ...font, bold: true }; break
        case 'i': next = { ...font, italic: true }; break
        case 'u': next = { ...font, underline: true }; break
        case 's': next = { ...font, strike: true }; break
        case 'sup': next = { ...font, vertAlign: 'superscript' }; break
        case 'sub': next = { ...font, vertAlign: 'subscript' }; break
        case 'font': next = applyXmlFont(child, font); break
        default: break
      }
      visit(child, next)
    }
  }
  visit(data, base)
  return runs
}

/**
 * Parses Excel's "XML Spreadsheet" clipboard flavour (only reachable through Electron's
 * clipboard.readBuffer). Unlike Excel's HTML it carries formulas (R1C1, resolved at
 * `options.origin`), exact values, notes and full styles.
 */
export function parseExcelXmlSpreadsheet(xml: string, options: ClipboardParseOptions = {}): ParsedClipboard | null {
  if (!xml || !/<(?:ss:)?Workbook[\s>]/i.test(xml)) return null
  const root = parseHtml(xml, { xml: true })
  let workbook: HtmlElement | undefined
  walkElements(root, (element) => {
    if (workbook) return false
    if (localName(element.tag) === 'workbook') { workbook = element; return false }
    return undefined
  })
  if (!workbook) return null
  const worksheet = xmlChildren(workbook, 'worksheet')[0]
  const table = worksheet ? xmlChildren(worksheet, 'table')[0] : undefined
  if (!table) return null
  const maxCells = options.maxCells ?? CLIPBOARD_MAX_CELLS
  const expandedRows = Number(xmlAttr(table, 'expandedrowcount'))
  const expandedCols = Number(xmlAttr(table, 'expandedcolumncount'))
  if (Number.isFinite(expandedRows) && Number.isFinite(expandedCols) && expandedRows * expandedCols > maxCells) throw new ClipboardTooLargeError()

  const origin = options.origin ?? { row: 0, col: 0 }
  const defaultFont = { name: options.defaultFont?.name ?? 'Calibri', size: options.defaultFont?.size ?? 11 }
  const styles = resolveXmlStyles(workbook)
  const grid: Array<Array<CellData | null | undefined>> = []
  const merges: string[] = []
  const widths: Array<number | undefined> = []
  const heights: Array<number | undefined> = []
  let width = 0
  let columnCursor = -1
  for (const column of xmlChildren(table, 'column')) {
    const index = Number(xmlAttr(column, 'index'))
    columnCursor = Number.isFinite(index) && index > 0 ? index - 1 : columnCursor + 1
    const span = Math.min(1000, Math.max(0, Number(xmlAttr(column, 'span') ?? 0) || 0))
    const pt = Number.parseFloat(xmlAttr(column, 'width') ?? '')
    for (let offset = 0; offset <= span && columnCursor + offset < SHEET_COLS; offset += 1) {
      if (Number.isFinite(pt) && pt > 0) widths[columnCursor + offset] = Math.round((pt * 4) / 3 * 100) / 100
    }
    columnCursor += span
  }

  let rowIndex = -1
  for (const row of xmlChildren(table, 'row')) {
    const explicitRow = Number(xmlAttr(row, 'index'))
    rowIndex = Number.isFinite(explicitRow) && explicitRow > 0 ? explicitRow - 1 : rowIndex + 1
    if (rowIndex >= SHEET_ROWS) throw new ClipboardTooLargeError()
    while (grid.length <= rowIndex) grid.push([])
    const heightPt = Number.parseFloat(xmlAttr(row, 'height') ?? '')
    if (Number.isFinite(heightPt) && heightPt > 0) heights[rowIndex] = Math.round((heightPt * 4) / 3 * 100) / 100
    let colIndex = -1
    for (const cellElement of xmlChildren(row, 'cell')) {
      const explicitCol = Number(xmlAttr(cellElement, 'index'))
      colIndex = Number.isFinite(explicitCol) && explicitCol > 0 ? explicitCol - 1 : colIndex + 1
      const across = Math.min(999, Math.max(0, Number(xmlAttr(cellElement, 'mergeacross') ?? 0) || 0))
      const down = Math.min(65_533, Math.max(0, Number(xmlAttr(cellElement, 'mergedown') ?? 0) || 0))
      const right = colIndex + across
      if (right >= SHEET_COLS) throw new ClipboardTooLargeError()
      if ((Math.max(grid.length, rowIndex + down + 1)) * Math.max(width, right + 1) > maxCells) throw new ClipboardTooLargeError()

      const style = styles.get(xmlAttr(cellElement, 'styleid') ?? 'Default') ?? styles.get('Default') as XmlStyle
      const cell: CellData = {}
      const data = xmlChildren(cellElement, 'data')[0]
      const type = (data ? xmlAttr(data, 'type') ?? '' : '').toLowerCase()
      const text = data ? xmlText(data) : ''
      let value: CellScalar | undefined
      let error = false
      if (data) {
        if (type === 'number') { const number = Number(text); if (Number.isFinite(number)) value = number }
        else if (type === 'boolean') value = text.trim() === '1' || text.trim().toLowerCase() === 'true'
        else if (type === 'datetime') value = xmlDateTime(text)
        else if (type === 'error') { value = text.trim(); error = true }
        else value = text.slice(0, MAX_TEXT_LENGTH)
      }
      const formula = xmlAttr(cellElement, 'formula')?.trim()
      if (formula && formula.startsWith('=') && formula.length <= MAX_FORMULA_LENGTH) {
        cell.formula = r1c1ToA1(formula, origin.row + rowIndex, origin.col + colIndex).slice(1)
        if (value !== undefined) cell.result = value
        if (error) cell.resultType = 'error'
      } else if (value !== undefined) {
        cell.value = value
        if (type === 'datetime' && style.numFmt && isDateNumberFormat(style.numFmt) && !isTimeOnlyFormat(style.numFmt)) cell.type = 'date'
      }
      if (style.numFmt) cell.numFmt = style.numFmt
      if (data && typeof cell.value === 'string' && data.children.some((child) => typeof child !== 'string')) {
        const runs = xmlRichRuns(data, style.font).filter((run) => run.text)
        const keys = new Set(runs.map((run) => fontKey(run.font)))
        if (runs.length > 1 && keys.size > 1) {
          cell.richText = runs.map((run) => {
            const runFont = finalizeFont(run.font, defaultFont)
            return runFont ? { text: run.text, font: runFont } : { text: run.text }
          })
        }
      }
      const cellStyle: CellStyle = {}
      const font = finalizeFont(style.font, defaultFont)
      if (font) cellStyle.font = cloneFont(font)
      if (style.fill) cellStyle.fill = cloneFill(style.fill)
      if (style.border) cellStyle.border = cloneBorder(style.border)
      if (style.alignment) cellStyle.alignment = { ...style.alignment }
      if (Object.keys(cellStyle).length) cell.style = cellStyle
      const hyperlink = sanitizeHyperlink(xmlAttr(cellElement, 'href'))
      if (hyperlink) cell.hyperlink = hyperlink
      const comment = xmlChildren(cellElement, 'comment')[0]
      if (comment) {
        const note = xmlText(comment).trim()
        if (note) cell.note = note.slice(0, MAX_TEXT_LENGTH)
      }

      for (let r = rowIndex; r <= rowIndex + down; r += 1) {
        while (grid.length <= r) grid.push([])
        for (let c = colIndex; c <= right; c += 1) if (r !== rowIndex || c !== colIndex) grid[r][c] = null
      }
      grid[rowIndex][colIndex] = cell
      if (across || down) merges.push(boundsToRange({ top: rowIndex, left: colIndex, bottom: rowIndex + down, right }))
      width = Math.max(width, right + 1)
      colIndex = right
    }
  }
  if (!grid.length || !width) return null
  const cells: CellData[][] = grid.map((line) => {
    const out: CellData[] = new Array(width)
    for (let col = 0; col < width; col += 1) out[col] = line[col] ?? {}
    return out
  })
  const result: ParsedClipboard = { cells, merges, source: 'excel', formulaOrigin: origin }
  const columnWidths = Array.from({ length: width }, (_, index) => widths[index])
  if (columnWidths.some((value) => value !== undefined)) result.columnWidths = columnWidths
  const rowHeights = Array.from({ length: cells.length }, (_, index) => heights[index])
  if (rowHeights.some((value) => value !== undefined)) result.rowHeights = rowHeights
  const sheetName = worksheet ? xmlAttr(worksheet, 'name') : undefined
  if (sheetName) result.sourceSheetName = sheetName.slice(0, 255)
  return result
}

/**
 * Parses whatever the clipboard holds, best flavour first: Excel's XML Spreadsheet (formulas),
 * then HTML (formatting, merges, typed values), then plain text.
 */
export function parseClipboardPayload(payload: { text?: string; html?: string; excelXml?: string }, options: ClipboardParseOptions = {}): ParsedClipboard | null {
  if (payload.excelXml) {
    try {
      const parsed = parseExcelXmlSpreadsheet(payload.excelXml, options)
      if (parsed && parsed.cells.length) return parsed
    } catch (error) {
      if (error instanceof ClipboardTooLargeError) throw error
      // A malformed XML flavour must never block the HTML / text fallback.
    }
  }
  if (payload.html) {
    const parsed = parseClipboardHtml(payload.html, options)
    if (parsed && parsed.cells.length) return parsed
  }
  if (payload.text) {
    const cells = parsePastedText(payload.text, options)
    if (cells.length) return { cells, merges: [], source: 'text', formulaOrigin: options.origin ?? { row: 0, col: 0 } }
  }
  return null
}

export function isSimpleCalcClipboardHtml(html: string | undefined): boolean {
  return Boolean(html && /data-sc-clip\s*=/.test(html))
}

// ---------------------------------------------------------------------------------------------
// CellData → HTML / TSV
// ---------------------------------------------------------------------------------------------

export interface SerializeClipboardOptions {
  /** Formatted text of the cell at (row, col) relative to the matrix. */
  displayAt?: (row: number, col: number) => string
  /** Theme-aware CSS color for a model color; '' when none. */
  cssColor?: (color: unknown) => string
  columnWidthsPx?: ArrayLike<number | undefined>
  rowHeightsPx?: ArrayLike<number | undefined>
  /** Merged areas as A1 ranges relative to the matrix (A1 = first selected cell). */
  merges?: string[]
  sheetName?: string
  /** Absolute sheet position of cells[0][0]; used for R1C1 formulas and internal re-paste. */
  origin?: ClipboardCoord
  defaultFont?: { name?: string; size?: number }
  /** Embed full CellData JSON for lossless paste between simple_calc windows (default ≤ 20k cells). */
  includeInternalData?: boolean
}

export interface ClipboardSerialized { text: string; html: string }

function escapeHtml(text: string) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function escapeAttribute(text: string) {
  return text.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

/** Excel's text flavour quotes a field only when it holds a tab / line break or starts with a quote. */
export function escapeTsvField(value: string): string {
  return /[\t\n\r]/.test(value) || value.startsWith('"') ? `"${value.replace(/"/g, '""')}"` : value
}

function fallbackCssColor(color: unknown): string {
  let value = ''
  if (typeof color === 'string') value = color
  else if (color && typeof color === 'object') {
    const object = color as { argb?: string; rgb?: string }
    value = object.argb || object.rgb || ''
  }
  value = value.replace(/^#/, '')
  if (value.length === 8) value = value.slice(2)
  return /^[0-9a-f]{6}$/i.test(value) ? `#${value.toUpperCase()}` : ''
}

function fallbackDisplay(cell: CellData | undefined) {
  if (!cell) return ''
  const value = cell.formula ? cell.result : cell.value
  if (value === null || value === undefined) return cell.display ?? ''
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'object') return cell.display ?? ''
  return String(value)
}

function cssFontFamily(name: string, family?: number | string) {
  const generic = Number(family) === 1 ? 'serif' : Number(family) === 3 ? 'monospace' : 'sans-serif'
  return `'${name.replace(/['"\\\u0000-\u001f]/g, '')}',${generic}`
}

function cssString(value: string) {
  return `'${value.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/[\r\n]+/g, ' ')}'`
}

const BORDER_CSS: Record<string, string> = {
  thin: '.5pt solid', medium: '1pt solid', thick: '1.5pt solid', dashed: '.5pt dashed', mediumDashed: '1pt dashed',
  dotted: '.5pt dotted', hair: '.5pt dotted', double: '2pt double', dashDot: '.5pt dashed', mediumDashDot: '1pt dashed',
  dashDotDot: '.5pt dotted', mediumDashDotDot: '1pt dotted', slantDashDot: '1pt dashed',
}

function fillCssColor(fill: CellFill | undefined, toCss: (color: unknown) => string) {
  if (!fill) return ''
  if (String(fill.type).toLowerCase() === 'gradient' && fill.stops?.length) return toCss(fill.stops[0].color)
  const pattern = String(fill.pattern ?? '').toLowerCase()
  const foreground = toCss(fill.color ?? fill.fgColor)
  if (pattern === 'none' && !foreground) return ''
  return foreground || toCss(fill.bgColor)
}

function fontCss(font: CellFont | undefined, toCss: (color: unknown) => string, defaultFont: { name: string; size: number }, css: string[], full = false) {
  if (!font) return
  if (font.name && (full || font.name.toLowerCase() !== defaultFont.name.toLowerCase())) css.push(`font-family:${cssFontFamily(font.name, font.family)}`)
  if (font.size && (full || font.size !== defaultFont.size)) css.push(`font-size:${font.size}pt`)
  if (font.bold) css.push('font-weight:700')
  if (font.italic) css.push('font-style:italic')
  const decoration = [font.underline ? 'underline' : '', font.strike ? 'line-through' : ''].filter(Boolean).join(' ')
  if (decoration) css.push(`text-decoration:${decoration}`)
  if (typeof font.underline === 'string' && /double/i.test(font.underline)) css.push('text-underline-style:double', 'text-decoration-style:double')
  const color = toCss(font.color)
  if (color) css.push(`color:${color}`)
  if (font.vertAlign === 'superscript') css.push('vertical-align:super')
  else if (font.vertAlign === 'subscript') css.push('vertical-align:sub')
}

function cellCss(cell: CellData | undefined, toCss: (color: unknown) => string, defaultFont: { name: string; size: number }): string[] {
  const css: string[] = []
  const style = cell?.style
  if (!cell) return css
  const font = style?.font
  if (font) {
    const { vertAlign: _ignored, ...cellFont } = font
    fontCss(cellFont, toCss, defaultFont, css)
  }
  const background = fillCssColor(style?.fill, toCss)
  if (background) css.push(`background-color:${background}`)
  for (const side of SIDES) {
    const border = style?.border?.[side] as CellBorderSide | undefined
    if (!border?.style || border.style === 'none') continue
    css.push(`border-${side}:${BORDER_CSS[border.style] ?? '.5pt solid'} ${toCss(border.color) || '#000000'}`)
  }
  const alignment = style?.alignment
  if (alignment) {
    const horizontal = ({ left: 'left', center: 'center', right: 'right', justify: 'justify', centerContinuous: 'center', fill: 'left', distributed: 'justify' } as Record<string, string>)[alignment.horizontal ?? '']
    if (horizontal) css.push(`text-align:${horizontal}`)
    const vertical = ({ top: 'top', middle: 'middle', center: 'middle', justify: 'middle', distributed: 'middle' } as Record<string, string>)[alignment.vertical ?? '']
    if (vertical) css.push(`vertical-align:${vertical}`)
    if (alignment.wrapText) css.push('white-space:normal', 'overflow-wrap:break-word')
    if (alignment.indent && alignment.indent > 0) css.push(`padding-left:${Math.round(alignment.indent * 12) + 3}px`, `mso-char-indent-count:${Math.round(alignment.indent)}`)
    if (alignment.textRotation === 'vertical') css.push('layout-flow:vertical')
    else if (typeof alignment.textRotation === 'number' && alignment.textRotation !== 0) css.push(`mso-rotate:${alignment.textRotation}`)
    if (alignment.shrinkToFit) css.push('mso-text-control:shrinktofit')
  }
  const numFmt = cell.numFmt || style?.numFmt
  if (numFmt && numFmt !== 'General') css.push(`mso-number-format:${cssString(numFmt)}`)
  return css
}

function textToHtml(text: string) {
  return escapeHtml(text)
    .replace(/^ /gm, '&nbsp;')
    .replace(/ {2,}/g, (spaces) => ` ${'&nbsp;'.repeat(spaces.length - 1)}`)
    .replace(/\r\n|\r|\n/g, '<br>')
}

const FORMULA_ERROR = /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|ERROR!)$/

function compactInternalCell(cell: CellData): CellData {
  const copy: CellData = { ...cell }
  delete copy.display
  delete copy.arrayMember
  return copy
}

/**
 * Serialises a selection for the system clipboard: `text` is TSV of the displayed values
 * (what Excel puts on the clipboard) and `html` is a styled table that Excel, Google Sheets,
 * Word, Outlook and simple_calc itself read with formatting, merges and formulas.
 */
export function serializeSelectionToClipboard(cells: ReadonlyArray<ReadonlyArray<CellData | undefined> | undefined>, options: SerializeClipboardOptions = {}): ClipboardSerialized {
  const rowCount = cells.length
  let colCount = 0
  for (const row of cells) if (row && row.length > colCount) colCount = row.length
  const origin = options.origin ?? { row: 0, col: 0 }
  const toCss = options.cssColor ?? fallbackCssColor
  const defaultFont = { name: options.defaultFont?.name ?? 'Calibri', size: options.defaultFont?.size ?? 11 }
  const includeInternal = options.includeInternalData ?? rowCount * colCount <= 20_000
  const displayAt = options.displayAt ?? ((row: number, col: number) => fallbackDisplay(cells[row]?.[col]))
  const stride = colCount + 1

  const masters = new Map<number, { rowSpan: number; colSpan: number }>()
  const covered = new Set<number>()
  for (const range of options.merges ?? []) {
    const bounds = rangeToBounds(range)
    if (!bounds || bounds.top >= rowCount || bounds.left >= colCount) continue
    const bottom = Math.min(bounds.bottom, rowCount - 1)
    const right = Math.min(bounds.right, colCount - 1)
    if (bottom === bounds.top && right === bounds.left) continue
    const key = bounds.top * stride + bounds.left
    if (covered.has(key) || masters.has(key)) continue
    masters.set(key, { rowSpan: bottom - bounds.top + 1, colSpan: right - bounds.left + 1 })
    for (let r = bounds.top; r <= bottom; r += 1) {
      for (let c = bounds.left; c <= right; c += 1) if (r !== bounds.top || c !== bounds.left) covered.add(r * stride + c)
    }
  }

  const widths: number[] = []
  let tableWidth = 0
  for (let col = 0; col < colCount; col += 1) {
    const raw = Number(options.columnWidthsPx?.[col])
    const width = Number.isFinite(raw) && raw > 0 ? Math.round(raw) : DEFAULT_COLUMN_PX
    widths.push(width)
    tableWidth += width
  }

  const textLines: string[] = []
  const html: string[] = []
  html.push('<meta charset="utf-8"><meta name="generator" content="simple_calc">')
  html.push('<style type="text/css"><!--td{border:none;padding:1px 3px;vertical-align:bottom;white-space:nowrap;overflow:hidden;}br{mso-data-placement:same-cell;}--></style>')
  html.push(`<table xmlns:x="urn:schemas-microsoft-com:office:excel" cellspacing="0" cellpadding="0" dir="ltr" border="0" style="table-layout:fixed;border-collapse:collapse;width:${tableWidth}px;font-family:${escapeAttribute(cssFontFamily(defaultFont.name))};font-size:${defaultFont.size}pt" data-sc-clip="1" data-sc-origin="${cellAddress(origin.row, origin.col)}"${options.sheetName ? ` data-sc-sheet="${escapeAttribute(options.sheetName)}"` : ''}>`)
  html.push('<colgroup>')
  for (const width of widths) html.push(`<col width="${width}" style="width:${width}px">`)
  html.push('</colgroup><tbody>')

  for (let row = 0; row < rowCount; row += 1) {
    const rawHeight = Number(options.rowHeightsPx?.[row])
    const height = Number.isFinite(rawHeight) && rawHeight > 0 ? Math.round(rawHeight) : DEFAULT_ROW_PX
    html.push(`<tr height="${height}" style="height:${height}px">`)
    const fields: string[] = []
    const line = cells[row]
    for (let col = 0; col < colCount; col += 1) {
      const key = row * stride + col
      if (covered.has(key)) { fields.push(''); continue }
      const cell = line?.[col]
      const display = String(displayAt(row, col) ?? '')
      fields.push(escapeTsvField(display))

      const attrs: string[] = []
      const css = cellCss(cell, toCss, defaultFont)
      const merge = masters.get(key)
      if (merge?.colSpan && merge.colSpan > 1) attrs.push(`colspan="${merge.colSpan}"`)
      if (merge?.rowSpan && merge.rowSpan > 1) attrs.push(`rowspan="${merge.rowSpan}"`)
      const explicitHorizontal = Boolean(cell?.style?.alignment?.horizontal && cell.style.alignment.horizontal !== 'general')
      const scalar = cell?.formula ? cell.result : cell?.value
      let sheetsValue: Record<string, unknown> | undefined
      const isError = typeof scalar === 'string' && (cell?.resultType === 'error' || (Boolean(cell?.formula) && FORMULA_ERROR.test(scalar)))
      if (typeof scalar === 'number' && Number.isFinite(scalar)) {
        attrs.push(`x:num="${scalar}"`)
        sheetsValue = { 1: 3, 3: scalar }
        if (!explicitHorizontal) attrs.push('align="right"')
      } else if (typeof scalar === 'boolean') {
        attrs.push(`x:bool="${scalar ? 'TRUE' : 'FALSE'}"`)
        sheetsValue = { 1: 4, 4: scalar ? 1 : 0 }
        if (!explicitHorizontal) attrs.push('align="center"')
      } else if (isError) {
        attrs.push(`x:err="${escapeAttribute(scalar)}"`)
        if (!explicitHorizontal) attrs.push('align="center"')
      } else if (typeof scalar === 'string' && scalar !== '') {
        sheetsValue = { 1: 2, 2: scalar }
        const typed = parseSmartValue(scalar)
        if (typed && typeof typed.value !== 'string') attrs.push('x:str')
      }
      if (cell?.formula) {
        const formula = `=${cell.formula.replace(/^=/, '')}`
        attrs.push(`x:fmla="${escapeAttribute(formula)}"`)
        attrs.push(`data-sheets-formula="${escapeAttribute(a1ToR1C1(formula, origin.row + row, origin.col + col))}"`)
      }
      if (sheetsValue) attrs.push(`data-sheets-value="${escapeAttribute(JSON.stringify(sheetsValue))}"`)
      const numFmt = cell?.numFmt || cell?.style?.numFmt
      if (numFmt && numFmt !== 'General') attrs.push(`data-sheets-numberformat="${escapeAttribute(JSON.stringify({ 1: sheetsNumberFormatType(numFmt), 2: numFmt, 3: 1 }))}"`)
      const noteText = notePlainText(cell?.note)
      if (noteText) attrs.push(`data-sheets-note="${escapeAttribute(noteText)}"`)
      const hyperlink = sanitizeHyperlink(cell?.hyperlink)
      if (hyperlink) attrs.push(`data-sheets-hyperlink="${escapeAttribute(hyperlink)}"`)
      if (includeInternal && cell && Object.keys(cell).length) attrs.push(`data-sc="${escapeAttribute(JSON.stringify(compactInternalCell(cell)))}"`)
      if (css.length) attrs.unshift(`style="${escapeAttribute(css.join(';'))}"`)

      let inner: string
      if (cell?.richText?.length && !cell.formula) {
        inner = cell.richText.map((run) => {
          const runCss: string[] = []
          fontCss(run.font, toCss, defaultFont, runCss)
          const text = textToHtml(String(run.text ?? ''))
          return runCss.length ? `<span style="${escapeAttribute(runCss.join(';'))}">${text}</span>` : text
        }).join('')
      } else inner = textToHtml(display)
      if (hyperlink) inner = `<a href="${escapeAttribute(hyperlink)}">${inner}</a>`
      html.push(`<td${attrs.length ? ` ${attrs.join(' ')}` : ''}>${inner}</td>`)
    }
    html.push('</tr>')
    textLines.push(fields.join('\t'))
  }
  html.push('</tbody></table>')
  return { text: textLines.join('\r\n'), html: html.join('') }
}
