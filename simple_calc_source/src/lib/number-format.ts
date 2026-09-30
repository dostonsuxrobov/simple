import { SSF } from 'xlsx'
import type { CellScalar } from '../spreadsheet-types'

const CURRENCY_SYMBOL = /[$€£¥₹₩₽]/

/** A value rendered through an Excel number format. */
export interface FormattedScalar {
  text: string
  /** CSS colour requested by the chosen section's colour tag ([Red], [Blue], [Color10] …). */
  color?: string
  /**
   * True when Excel cannot show the value and fills the cell with "#" instead — a negative
   * or out-of-range serial under a date/time format. `text` then holds a short "#" run.
   */
  overflow?: boolean
}

/** Excel's eight named format colours. */
const FORMAT_COLOR_NAMES: Record<string, string> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#00FF00',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
}

/** Excel's default 56-colour palette: `[ColorN]` resolves to entry N-1 (indexed colours 8-63). */
export const EXCEL_FORMAT_PALETTE: readonly string[] = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
  '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
  '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
]

// ---------------------------------------------------------------------------------------
// Format-code tokenizer (shared with format-codes.ts)
// ---------------------------------------------------------------------------------------

export type FormatTokenKind = 'quoted' | 'escaped' | 'skip' | 'fill' | 'bracket' | 'char'

export interface FormatToken {
  kind: FormatTokenKind
  /** Exact source text of the token (`"kg"`, `\-`, `_)`, `*-`, `[Red]`, `0`). */
  text: string
  /** Literal payload: the quoted text, escaped/skip/fill character, bracket body, or the char itself. */
  value: string
}

/** Splits one format section into tokens. Quotes, escapes, `_x`, `*x` and `[...]` stay whole. */
export function tokenizeFormat(section: string): FormatToken[] {
  const tokens: FormatToken[] = []
  let index = 0
  while (index < section.length) {
    const char = section[index]
    if (char === '"') {
      const end = section.indexOf('"', index + 1)
      const stop = end < 0 ? section.length : end + 1
      tokens.push({ kind: 'quoted', text: section.slice(index, stop), value: section.slice(index + 1, end < 0 ? section.length : end) })
      index = stop
    } else if (char === '\\' || char === '_' || char === '*') {
      const next = section[index + 1] ?? ''
      tokens.push({ kind: char === '\\' ? 'escaped' : char === '_' ? 'skip' : 'fill', text: char + next, value: next })
      index += next ? 2 : 1
    } else if (char === '[') {
      const end = section.indexOf(']', index + 1)
      const stop = end < 0 ? section.length : end + 1
      tokens.push({ kind: 'bracket', text: section.slice(index, stop), value: section.slice(index + 1, end < 0 ? section.length : end) })
      index = stop
    } else {
      tokens.push({ kind: 'char', text: char, value: char })
      index += 1
    }
  }
  return tokens
}

/** Splits a format code into its `;` sections, ignoring semicolons inside quotes, escapes and brackets. */
export function splitFormatSections(code: string): string[] {
  const sections: string[] = []
  let current = ''
  for (const token of tokenizeFormat(code)) {
    if (token.kind === 'char' && token.value === ';') {
      sections.push(current)
      current = ''
    } else current += token.text
  }
  sections.push(current)
  return sections
}

const ELAPSED_BRACKET = /^(h+|m+|s+)$/i
const CONDITION_BRACKET = /^(=|>=?|<[>=]?)\s*(-?\d+(?:\.\d*)?(?:e[+-]?\d+)?)$/i
const COLOR_BRACKET = /^(black|blue|cyan|green|magenta|red|white|yellow|color\s*(\d{1,2}))$/i

interface SectionInfo {
  raw: string
  color?: string
  condition?: { operator: string; threshold: number }
  /** Contains date or time placeholders (including elapsed [h]/[m]/[s]). */
  date: boolean
  general: boolean
  /** The section without bracket codes, for Excel behaviours SSF misses. */
  body: string
  /** An engineering-notation mantissa (`##0.0E+0`) that SSF renders wrongly for small values. */
  engineering: boolean
}

interface FormatInfo {
  sections: SectionInfo[]
}

function colorFromBracket(body: string) {
  const match = COLOR_BRACKET.exec(body.trim())
  if (!match) return undefined
  if (match[2] !== undefined) {
    const entry = EXCEL_FORMAT_PALETTE[Number(match[2]) - 1]
    return entry ? `#${entry}` : undefined
  }
  return FORMAT_COLOR_NAMES[match[1].toLowerCase()]
}

function analyzeSection(raw: string): SectionInfo {
  const info: SectionInfo = { raw, date: false, general: false, body: '', engineering: false }
  let bare = ''
  for (const token of tokenizeFormat(raw)) {
    if (token.kind !== 'bracket') info.body += token.text
    if (token.kind === 'bracket') {
      const body = token.value.trim()
      if (ELAPSED_BRACKET.test(body)) { info.date = true; continue }
      const color = colorFromBracket(body)
      if (color) { info.color ??= color; continue }
      const condition = CONDITION_BRACKET.exec(body)
      if (condition && !info.condition) info.condition = { operator: condition[1], threshold: Number(condition[2]) }
      continue
    }
    if (token.kind === 'char') bare += token.value
  }
  if (/general/i.test(bare)) info.general = true
  const withoutGeneral = bare.replace(/general/gi, '')
  if (/[ymdhs]/i.test(withoutGeneral.replace(/e[+-]/gi, '')) || /am\/pm|a\/p/i.test(withoutGeneral)) info.date = true
  const engineering = ENGINEERING.exec(info.body.trim())
  info.engineering = Boolean(engineering && engineering[1].length > 1)
  return info
}

const formatInfoCache = new Map<string, FormatInfo | null>()

function formatInfo(code: string): FormatInfo | null {
  const cached = formatInfoCache.get(code)
  if (cached !== undefined) return cached
  const sections = splitFormatSections(code)
  const info = sections.length > 4 ? null : { sections: sections.map(analyzeSection) }
  if (formatInfoCache.size > 4000) formatInfoCache.clear()
  formatInfoCache.set(code, info)
  return info
}

function conditionHolds(value: number, condition: SectionInfo['condition']) {
  if (!condition) return false
  const { operator, threshold } = condition
  switch (operator) {
    case '=': return value === threshold
    case '>': return value > threshold
    case '<': return value < threshold
    case '>=': return value >= threshold
    case '<=': return value <= threshold
    case '<>': return value !== threshold
    default: return false
  }
}

/**
 * Index of the section Excel (and SheetJS, whose text we display) uses for a number, or -1
 * when a text-only format falls back to General. Mirrors SSF's `choose_fmt`.
 */
function numberSectionIndex(info: FormatInfo, value: number) {
  const sections = info.sections
  const count = sections.length
  const hasText = sections[count - 1].raw.includes('@')
  const map = count === 1 ? (hasText ? [-1, -1, -1, 0] : [0, 0, 0, -1])
    : count === 2 ? (hasText ? [0, 0, 0, 1] : [0, 1, 0, -1])
      : count === 3 ? (hasText ? [0, 1, 0, 2] : [0, 1, 2, -1])
        : [0, 1, 2, 3]
  const v = Number.isFinite(value) ? value : 0
  const byValue = map[v > 0 ? 0 : v < 0 ? 1 : 2]
  const first = sections[map[0]]
  const second = sections[map[1]]
  if (!first?.raw.includes('[') && !second?.raw.includes('[')) return byValue
  if (first?.condition || second?.condition) {
    if (conditionHolds(v, first?.condition)) return map[0]
    if (conditionHolds(v, second?.condition)) return map[1]
    return map[first?.condition && second?.condition ? 2 : 1]
  }
  return byValue
}

/** Index of the section applied to text, or -1 when the format has no text section. */
function textSectionIndex(info: FormatInfo) {
  const count = info.sections.length
  return count === 4 || info.sections[count - 1].raw.includes('@') ? count - 1 : -1
}

// ---------------------------------------------------------------------------------------
// Excel behaviours SheetJS SSF does not reproduce
// ---------------------------------------------------------------------------------------

const ENGINEERING = /^([#0?]+)(?:\.([#0?]*))?[eE]([+-])([0#?]+)$/

/**
 * Engineering notation (`##0.0E+0`): the exponent is a multiple of the integer placeholder
 * count. SSF gets negative exponents wrong (0.00012345 → "1.23E-04" instead of "123.45E-06").
 */
function formatEngineering(body: string, value: number): string | null {
  const match = ENGINEERING.exec(body)
  if (!match) return null
  const integerPattern = match[1]
  const width = integerPattern.length
  if (width < 2) return null
  const fractionPattern = match[2] ?? ''
  const required = fractionPattern.replace(/[#?]+$/, '').length
  const decimals = fractionPattern.length
  const magnitude = Math.abs(value)
  let exponent = 0
  if (magnitude > 0) exponent = Math.floor(Math.floor(Math.log10(magnitude)) / width) * width
  let mantissa = magnitude / 10 ** exponent
  let fixed = mantissa.toFixed(decimals)
  if (Number(fixed) >= 10 ** width) {
    exponent += width
    mantissa = magnitude / 10 ** exponent
    fixed = mantissa.toFixed(decimals)
  }
  let [whole, fraction = ''] = fixed.split('.')
  const minimumWhole = (integerPattern.match(/0/g) || []).length
  if (whole === '0' && minimumWhole === 0) whole = ''
  whole = whole.padStart(minimumWhole, '0')
  while (fraction.length > required && fraction.endsWith('0')) fraction = fraction.slice(0, -1)
  const exponentDigits = (match[4].match(/0/g) || []).length || 1
  const sign = exponent < 0 ? '-' : match[3] === '+' ? '+' : ''
  const exponentText = `${sign}${String(Math.abs(exponent)).padStart(exponentDigits, '0')}`
  const numberText = `${whole}${decimals ? `.${fraction}` : ''}E${exponentText}`
  return value < 0 ? `-${numberText}` : numberText
}

/** Quotes "." separators in date sections, keeping the seconds fraction of `ss.00`. */
function quoteDateDots(code: string) {
  return splitFormatSections(code).map((section) => {
    if (!analyzeSection(section).date) return section
    const tokens = tokenizeFormat(section)
    return tokens.map((token, index) => {
      if (token.kind !== 'char' || token.value !== '.') return token.text
      const previous = tokens[index - 1]
      const next = tokens[index + 1]
      const afterSeconds = previous && ((previous.kind === 'char' && /s/i.test(previous.value)) || (previous.kind === 'bracket' && /^s+$/i.test(previous.value)))
      if (afterSeconds && next?.kind === 'char' && next.value === '0') return token.text
      return '"."'
    }).join('')
  }).join(';')
}

/** Formats with SSF, retrying the few Excel-valid spellings SSF rejects. */
function ssfFormat(code: string, value: number | string): string | null {
  try {
    const formatted = SSF.format(code, value)
    return typeof formatted === 'string' ? formatted : null
  } catch (error) {
    const message = error instanceof Error ? String(error.message) : String(error)
    // "dd.mm.yyyy" — SSF rejects "." between date parts; quoting it renders identically.
    if (code.includes('.')) {
      const quoted = quoteDateDots(code)
      if (quoted !== code) return ssfFormat(quoted, value)
    }
    // "0 ?/?" — a zero integer placeholder in front of a fraction.
    if (typeof value === 'number' && /\//.test(code) && /unsupported format/i.test(message)) {
      const alternate = code.replace(/(^|[^#0?.,])0(?=\s+[?#0]+\/)/g, '$1#')
      if (alternate !== code) {
        try {
          const formatted = SSF.format(alternate, value)
          if (typeof formatted === 'string') {
            if (Math.trunc(Math.abs(value)) === 0 && value !== 0 && /\//.test(formatted)) {
              return formatted.replace(/^\s*(-?)\s?/, '$10 ')
            }
            return formatted
          }
        } catch {
          // Fall through to the generic retries.
        }
      }
    }
    // An unquoted literal character Excel tolerates in files written by other tools.
    const unrecognized = /unrecognized character (.) in/i.exec(message)
    if (unrecognized) {
      const tokens = tokenizeFormat(code)
      const position = tokens.findIndex((token) => token.kind === 'char' && token.value === unrecognized[1])
      if (position >= 0) {
        tokens[position] = { kind: 'escaped', text: `\\${unrecognized[1]}`, value: unrecognized[1] }
        const escaped = tokens.map((token) => token.text).join('')
        if (escaped !== code) return ssfFormat(escaped, value)
      }
    }
    // "???.???" alignment placeholders outside fractions.
    if (/\?/.test(code) && !/\//.test(code)) {
      const alternate = code.replace(/\?/g, '#')
      if (alternate !== code) return ssfFormat(alternate, value)
    }
    return null
  }
}

const MAX_DATE_SERIAL = 2958466 // 1/1/10000

function output(text: string, color?: string, overflow = false): FormattedScalar {
  const result: FormattedScalar = { text }
  if (color) result.color = color
  if (overflow) result.overflow = true
  return result
}

function formatText(value: string, format: string, fallbackDisplay?: string): FormattedScalar {
  if (!format) return { text: fallbackDisplay ?? value }
  const info = formatInfo(format)
  if (!info) return { text: fallbackDisplay ?? value }
  const index = textSectionIndex(info)
  if (index < 0) return { text: fallbackDisplay ?? value }
  const color = info.sections[index].color
  const formatted = ssfFormat(format, value)
  if (formatted === null || formatted.trim() === value.trim()) return output(fallbackDisplay ?? value, color)
  return output(formatted, color)
}

/**
 * Formats a cell value like Excel: picks the positive/negative/zero/text section (honouring
 * `[>100]`-style conditions), applies the section's colour tag and returns the display text.
 */
export function formatScalarDetailed(value: CellScalar | undefined, numFmt?: string, fallbackDisplay?: string): FormattedScalar {
  if (value === undefined || value === null) return { text: '' }
  // A structured value must never reach the grid as "[object Object]".
  if (typeof value === 'object') return { text: fallbackDisplay ?? '' }
  const format = String(numFmt || '').trim()
  if (typeof value === 'string') return formatText(value, format, fallbackDisplay)
  if (typeof value !== 'number') return { text: fallbackDisplay ?? String(value) }
  if (!format) return { text: fallbackDisplay ?? String(value) }

  const info = formatInfo(format)
  const index = info ? numberSectionIndex(info, value) : -1
  const section = info && index >= 0 ? info.sections[index] : undefined
  const color = section?.color

  if (section?.date && (value < 0 || value >= MAX_DATE_SERIAL)) return output('########', color, true)

  if (section && info) {
    // Engineering notation with a negative exponent.
    if (section.engineering && (info.sections.length === 1 || (index === 0 && value >= 0))) {
      const engineering = formatEngineering(section.body.trim(), value)
      if (engineering !== null) return output(engineering, color)
    }
    // SSF doubles the sign when a negative section spells General ("-General" → "--3").
    if (section.general && index === 1 && value < 0 && !section.condition) {
      const formatted = ssfFormat(section.body, Math.abs(value))
      if (formatted !== null) return output(formatted, color)
    }
  }

  const formatted = ssfFormat(format, value)
  if (formatted !== null) return output(formatted, color)
  // Keep the workbook usable when a vendor-specific format is not supported.
  return { text: fallbackDisplay ?? String(value) }
}

export function formatScalar(value: CellScalar | undefined, numFmt?: string, fallbackDisplay?: string) {
  return formatScalarDetailed(value, numFmt, fallbackDisplay).text
}

/** The CSS colour a format assigns to a value, if its chosen section carries a colour tag. */
export function formatColor(value: CellScalar | undefined, numFmt?: string) {
  const format = String(numFmt || '').trim()
  if (!format || value === undefined || value === null || typeof value === 'object' || typeof value === 'boolean') return undefined
  const info = formatInfo(format)
  if (!info) return undefined
  if (typeof value === 'string') {
    const index = textSectionIndex(info)
    return index >= 0 ? info.sections[index].color : undefined
  }
  const index = numberSectionIndex(info, value)
  return index >= 0 ? info.sections[index].color : undefined
}

export function accountingCurrencySymbol(numFmt?: string) {
  const format = String(numFmt || '')
  if (!isAccountingNumberFormat(format)) return ''
  // Only a symbol written before the fill (`*`) sits at the cell's left edge; a trailing
  // symbol ("1,234.00 €") stays attached to the number.
  const first = splitFormatSections(format)[0] || ''
  let head = ''
  for (const token of tokenizeFormat(first)) {
    if (token.kind === 'fill') break
    head += token.kind === 'quoted' ? token.value : token.kind === 'bracket' ? `[${token.value}]` : token.value
  }
  const bracket = /\[\$([^\]-]+)(?:-[^\]]*)?\]/.exec(head)
  if (bracket && bracket[1].trim()) return bracket[1].trim()
  return head.match(CURRENCY_SYMBOL)?.[0] || ''
}

export function isAccountingNumberFormat(numFmt?: string) {
  return String(numFmt || '').includes('*')
}

export function accountingDisplayParts(display: string, numFmt?: string) {
  const symbol = accountingCurrencySymbol(numFmt)
  if (!symbol) return null
  const text = String(display || '').trim()
  const symbolIndex = text.indexOf(symbol)
  if (symbolIndex < 0) return null
  const amount = `${text.slice(0, symbolIndex)}${text.slice(symbolIndex + symbol.length)}`.trim()
  return { symbol, amount }
}
