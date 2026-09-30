/**
 * Excel "Format Cells → Number" logic: categories, builders for every category, detection of
 * the category/options behind an existing code, Excel's custom preset list, and decimal
 * increase/decrease that works on any code.
 */
import { formatScalarDetailed, splitFormatSections, tokenizeFormat } from './number-format'
import type { FormatToken } from './number-format'

export type NumberFormatCategory =
  | 'general' | 'number' | 'currency' | 'accounting' | 'date' | 'time'
  | 'percentage' | 'fraction' | 'scientific' | 'text' | 'special' | 'custom'

export interface NumberFormatCategoryInfo {
  id: NumberFormatCategory
  label: string
  description: string
}

export const NUMBER_FORMAT_CATEGORIES: readonly NumberFormatCategoryInfo[] = [
  { id: 'general', label: 'General', description: 'General format cells have no specific number format.' },
  { id: 'number', label: 'Number', description: 'Number is used for general display of numbers. Currency and Accounting offer specialized formatting for monetary value.' },
  { id: 'currency', label: 'Currency', description: 'Currency formats are used for general monetary values. Use Accounting formats to align decimal points in a column.' },
  { id: 'accounting', label: 'Accounting', description: 'Accounting formats line up the currency symbols and decimal points in a column.' },
  { id: 'date', label: 'Date', description: 'Date formats display date and time serial numbers as date values. Formats that begin with an asterisk (*) follow the system’s regional settings.' },
  { id: 'time', label: 'Time', description: 'Time formats display date and time serial numbers as time values. Formats that begin with an asterisk (*) follow the system’s regional settings.' },
  { id: 'percentage', label: 'Percentage', description: 'Percentage formats multiply the cell value by 100 and display the result with a percent symbol.' },
  { id: 'fraction', label: 'Fraction', description: 'Fraction formats show the fractional part of a number as the closest fraction of the chosen precision.' },
  { id: 'scientific', label: 'Scientific', description: 'Scientific formats display numbers in exponential notation, replacing part of the number with E+n.' },
  { id: 'text', label: 'Text', description: 'Text format cells are treated as text even when a number is in the cell. The cell is displayed exactly as entered.' },
  { id: 'special', label: 'Special', description: 'Special formats are useful for tracking list and database values.' },
  { id: 'custom', label: 'Custom', description: 'Type the number format code, using one of the existing codes as a starting point.' },
]

export type NegativeNumberStyle = 'minus' | 'red' | 'parens' | 'red-parens' | 'red-minus'

export const NEGATIVE_NUMBER_STYLES: readonly { id: NegativeNumberStyle; label: string; red: boolean }[] = [
  { id: 'minus', label: 'Minus sign', red: false },
  { id: 'red', label: 'Red', red: true },
  { id: 'parens', label: 'Parentheses', red: false },
  { id: 'red-parens', label: 'Red parentheses', red: true },
  { id: 'red-minus', label: 'Red minus sign', red: true },
]

export interface CurrencySymbol {
  id: string
  /** Label shown in the Symbol list. */
  label: string
  /** The characters the symbol displays as. */
  symbol: string
  /** Code written before the number, e.g. `$` or `[$€-x-euro2] `. */
  prefix?: string
  /** Code written after the number, e.g. ` [$€-x-euro1]`. */
  suffix?: string
}

export const CURRENCY_SYMBOLS: readonly CurrencySymbol[] = [
  { id: 'none', label: 'None', symbol: '' },
  { id: 'usd', label: '$ English (United States)', symbol: '$', prefix: '$' },
  { id: 'eur-prefix', label: '€ Euro (€ 123)', symbol: '€', prefix: '[$€-x-euro2] ' },
  { id: 'eur-suffix', label: '€ Euro (123 €)', symbol: '€', suffix: ' [$€-x-euro1]' },
  { id: 'gbp', label: '£ English (United Kingdom)', symbol: '£', prefix: '[$£-809]' },
  { id: 'jpy', label: '¥ Japanese', symbol: '¥', prefix: '[$¥-411]' },
  { id: 'cny', label: '¥ Chinese (PRC)', symbol: '¥', prefix: '[$¥-804]' },
  { id: 'inr', label: '₹ English (India)', symbol: '₹', prefix: '[$₹-4009] ' },
  { id: 'krw', label: '₩ Korean', symbol: '₩', prefix: '[$₩-412]' },
  { id: 'rub', label: '₽ Russian', symbol: '₽', suffix: ' [$₽-419]' },
  { id: 'chf', label: 'CHF German (Switzerland)', symbol: 'CHF', prefix: '[$CHF-807] ' },
  { id: 'cad', label: '$ English (Canada)', symbol: '$', prefix: '[$$-1009]' },
  { id: 'aud', label: '$ English (Australia)', symbol: '$', prefix: '[$$-C09]' },
  { id: 'mxn', label: '$ Spanish (Mexico)', symbol: '$', prefix: '[$$-80A]' },
  { id: 'brl', label: 'R$ Portuguese (Brazil)', symbol: 'R$', prefix: '[$R$-416] ' },
  { id: 'sek', label: 'kr Swedish (Sweden)', symbol: 'kr', suffix: ' [$kr-41D]' },
  { id: 'nok', label: 'kr Norwegian (Bokmål)', symbol: 'kr', prefix: '[$kr-414] ' },
  { id: 'dkk', label: 'kr. Danish', symbol: 'kr.', prefix: '[$kr.-406] ' },
  { id: 'pln', label: 'zł Polish', symbol: 'zł', suffix: ' [$zł-415]' },
  { id: 'try', label: '₺ Turkish', symbol: '₺', prefix: '[$₺-41F]' },
  { id: 'ils', label: '₪ Hebrew', symbol: '₪', prefix: '[$₪-40D] ' },
  { id: 'zar', label: 'R English (South Africa)', symbol: 'R', prefix: '[$R-1C09]' },
  { id: 'uah', label: '₴ Ukrainian', symbol: '₴', suffix: ' [$₴-422]' },
  { id: 'vnd', label: '₫ Vietnamese', symbol: '₫', suffix: ' [$₫-42A]' },
  { id: 'thb', label: '฿ Thai', symbol: '฿', prefix: '[$฿-41E]' },
  { id: 'php', label: '₱ Filipino', symbol: '₱', prefix: '[$₱-3409]' },
  { id: 'usd-iso', label: 'USD', symbol: 'USD', prefix: '[$USD] ' },
  { id: 'eur-iso', label: 'EUR', symbol: 'EUR', prefix: '[$EUR] ' },
  { id: 'gbp-iso', label: 'GBP', symbol: 'GBP', prefix: '[$GBP] ' },
  { id: 'jpy-iso', label: 'JPY', symbol: 'JPY', prefix: '[$JPY] ' },
  { id: 'cny-iso', label: 'CNY', symbol: 'CNY', prefix: '[$CNY] ' },
  { id: 'inr-iso', label: 'INR', symbol: 'INR', prefix: '[$INR] ' },
]

export interface FormatTypeOption {
  id: string
  code: string
  /** Fixed label (fractions, special); date/time lists render `sampleText` instead. */
  label: string
  /** Follows the operating system's regional settings (shown with a leading asterisk). */
  system?: boolean
}

/** Excel's sample instant for the Date and Time type lists: 3/14/2012 1:30:55.2 PM. */
export const SAMPLE_DATE_SERIAL = 40982.5631388889

export const DATE_FORMATS: readonly FormatTypeOption[] = [
  { id: 'm/d/yyyy', code: 'm/d/yyyy', label: '3/14/2012', system: true },
  { id: '[$-x-sysdate]dddd, mmmm dd, yyyy', code: '[$-x-sysdate]dddd, mmmm dd, yyyy', label: 'Wednesday, March 14, 2012', system: true },
  { id: 'm/d;@', code: 'm/d;@', label: '3/14' },
  { id: 'm/d/yy;@', code: 'm/d/yy;@', label: '3/14/12' },
  { id: 'mm/dd/yy;@', code: 'mm/dd/yy;@', label: '03/14/12' },
  { id: '[$-409]d-mmm;@', code: '[$-409]d-mmm;@', label: '14-Mar' },
  { id: '[$-409]d-mmm-yy;@', code: '[$-409]d-mmm-yy;@', label: '14-Mar-12' },
  { id: '[$-409]dd-mmm-yy;@', code: '[$-409]dd-mmm-yy;@', label: '14-Mar-12' },
  { id: '[$-409]mmm-yy;@', code: '[$-409]mmm-yy;@', label: 'Mar-12' },
  { id: '[$-409]mmmm-yy;@', code: '[$-409]mmmm-yy;@', label: 'March-12' },
  { id: '[$-409]mmmm d, yyyy;@', code: '[$-409]mmmm d, yyyy;@', label: 'March 14, 2012' },
  { id: '[$-409]dddd, mmmm d, yyyy;@', code: '[$-409]dddd, mmmm d, yyyy;@', label: 'Wednesday, March 14, 2012' },
  { id: '[$-409]m/d/yy h:mm AM/PM;@', code: '[$-409]m/d/yy h:mm AM/PM;@', label: '3/14/12 1:30 PM' },
  { id: 'm/d/yy h:mm;@', code: 'm/d/yy h:mm;@', label: '3/14/12 13:30' },
  { id: '[$-409]mmmmm;@', code: '[$-409]mmmmm;@', label: 'M' },
  { id: '[$-409]mmmmm-yy;@', code: '[$-409]mmmmm-yy;@', label: 'M-12' },
  { id: 'm/d/yyyy;@', code: 'm/d/yyyy;@', label: '3/14/2012' },
  { id: '[$-409]d-mmm-yyyy;@', code: '[$-409]d-mmm-yyyy;@', label: '14-Mar-2012' },
  { id: 'yyyy-mm-dd;@', code: 'yyyy-mm-dd;@', label: '2012-03-14' },
  { id: 'yyyy-mm-dd h:mm;@', code: 'yyyy-mm-dd h:mm;@', label: '2012-03-14 13:30' },
  { id: 'dd/mm/yyyy;@', code: 'dd/mm/yyyy;@', label: '14/03/2012' },
  { id: 'dd.mm.yyyy;@', code: 'dd.mm.yyyy;@', label: '14.03.2012' },
]

export const TIME_FORMATS: readonly FormatTypeOption[] = [
  { id: '[$-x-systime]h:mm:ss AM/PM', code: '[$-x-systime]h:mm:ss AM/PM', label: '1:30:55 PM', system: true },
  { id: 'h:mm;@', code: 'h:mm;@', label: '13:30' },
  { id: '[$-409]h:mm AM/PM;@', code: '[$-409]h:mm AM/PM;@', label: '1:30 PM' },
  { id: 'h:mm:ss;@', code: 'h:mm:ss;@', label: '13:30:55' },
  { id: '[$-409]h:mm:ss AM/PM;@', code: '[$-409]h:mm:ss AM/PM;@', label: '1:30:55 PM' },
  { id: 'mm:ss.0;@', code: 'mm:ss.0;@', label: '30:55.2' },
  { id: '[h]:mm:ss;@', code: '[h]:mm:ss;@', label: '37:30:55' },
  { id: '[$-409]m/d/yy h:mm AM/PM;@', code: '[$-409]m/d/yy h:mm AM/PM;@', label: '3/14/12 1:30 PM' },
  { id: 'm/d/yy h:mm;@', code: 'm/d/yy h:mm;@', label: '3/14/12 13:30' },
]

export const FRACTION_FORMATS: readonly FormatTypeOption[] = [
  { id: 'up-to-1', code: '# ?/?', label: 'Up to one digit (1/4)' },
  { id: 'up-to-2', code: '# ??/??', label: 'Up to two digits (21/25)' },
  { id: 'up-to-3', code: '# ???/???', label: 'Up to three digits (312/943)' },
  { id: 'halves', code: '# ?/2', label: 'As halves (1/2)' },
  { id: 'quarters', code: '# ?/4', label: 'As quarters (2/4)' },
  { id: 'eighths', code: '# ?/8', label: 'As eighths (4/8)' },
  { id: 'sixteenths', code: '# ??/16', label: 'As sixteenths (8/16)' },
  { id: 'tenths', code: '# ?/10', label: 'As tenths (3/10)' },
  { id: 'hundredths', code: '# ??/100', label: 'As hundredths (30/100)' },
]

export const SPECIAL_FORMATS: readonly FormatTypeOption[] = [
  { id: 'zip', code: '00000', label: 'Zip Code' },
  { id: 'zip4', code: '00000-0000', label: 'Zip Code + 4' },
  { id: 'phone', code: '[<=9999999]###-####;(###) ###-####', label: 'Phone Number' },
  { id: 'ssn', code: '000-00-0000', label: 'Social Security Number' },
]

/** Excel's built-in Custom list (en-US), followed by a few widely used additions. */
export const CUSTOM_FORMAT_PRESETS: readonly string[] = [
  'General',
  '0',
  '0.00',
  '#,##0',
  '#,##0.00',
  '#,##0_);(#,##0)',
  '#,##0_);[Red](#,##0)',
  '#,##0.00_);(#,##0.00)',
  '#,##0.00_);[Red](#,##0.00)',
  '$#,##0_);($#,##0)',
  '$#,##0_);[Red]($#,##0)',
  '$#,##0.00_);($#,##0.00)',
  '$#,##0.00_);[Red]($#,##0.00)',
  '0%',
  '0.00%',
  '0.00E+00',
  '##0.0E+0',
  '# ?/?',
  '# ??/??',
  'm/d/yyyy',
  'd-mmm-yy',
  'd-mmm',
  'mmm-yy',
  'h:mm AM/PM',
  'h:mm:ss AM/PM',
  'h:mm',
  'h:mm:ss',
  'm/d/yyyy h:mm',
  'mm:ss',
  'mm:ss.0',
  '@',
  '[h]:mm:ss',
  '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
  '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
  '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
  '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
  'yyyy-mm-dd',
  'yyyy-mm-dd hh:mm:ss',
  'dddd, mmmm d, yyyy',
  '#,##0,"K"',
  '#,##0.0,,"M"',
  '0.0%;[Red]-0.0%',
  '[Blue]#,##0;[Red]-#,##0;0',
  '#,##0;-#,##0;"–"',
  ';;;',
]

export const MAX_DECIMALS = 30

export interface NumberFormatOptions {
  /** Number, Currency, Accounting, Percentage, Scientific: 0-30. */
  decimals?: number
  /** Number: use the 1000 separator. */
  thousands?: boolean
  /** Number, Currency. */
  negative?: NegativeNumberStyle
  /** Currency, Accounting: a CURRENCY_SYMBOLS id. */
  symbol?: string
  /** Date, Time, Fraction, Special: the chosen type's id. */
  type?: string
  /** Custom: the literal code. */
  code?: string
}

export interface DetectedNumberFormat {
  category: NumberFormatCategory
  options: Required<Pick<NumberFormatOptions, 'decimals' | 'thousands' | 'negative' | 'symbol'>> & NumberFormatOptions
}

export const DEFAULT_NUMBER_FORMAT_OPTIONS: DetectedNumberFormat['options'] = {
  decimals: 2,
  thousands: false,
  negative: 'minus',
  symbol: 'usd',
}

const clampDecimals = (value: number | undefined, fallback = 2) => {
  const number = Number.isFinite(value) ? Math.trunc(Number(value)) : fallback
  return Math.min(MAX_DECIMALS, Math.max(0, number))
}

const decimalPart = (decimals: number) => (decimals > 0 ? `.${'0'.repeat(decimals)}` : '')

export function currencySymbol(id: string | undefined): CurrencySymbol {
  return CURRENCY_SYMBOLS.find((symbol) => symbol.id === id) || CURRENCY_SYMBOLS[1]
}

function withNegative(positive: string, style: NegativeNumberStyle | undefined) {
  switch (style) {
    case 'red': return `${positive};[Red]${positive}`
    case 'parens': return `${positive}_);(${positive})`
    case 'red-parens': return `${positive}_);[Red](${positive})`
    case 'red-minus': return `${positive};[Red]-${positive}`
    default: return positive
  }
}

function accountingCode(decimals: number, symbol: CurrencySymbol) {
  const fraction = decimalPart(decimals)
  const pad = '?'.repeat(decimals)
  if (symbol.id === 'none') return `_(* #,##0${fraction}_);_(* (#,##0${fraction});_(* "-"${pad}_);_(@_)`
  if (symbol.id === 'usd') return `_($* #,##0${fraction}_);_($* (#,##0${fraction});_($* "-"${pad}_);_(@_)`
  if (symbol.prefix) {
    const prefix = symbol.prefix
    return `_-${prefix}* #,##0${fraction}_-;-${prefix}* #,##0${fraction}_-;_-${prefix}* "-"${pad}_-;_-@_-`
  }
  const suffix = symbol.suffix || ''
  return `_-* #,##0${fraction}${suffix}_-;-* #,##0${fraction}${suffix}_-;_-* "-"${pad}${suffix}_-;_-@_-`
}

const typeCode = (list: readonly FormatTypeOption[], id: string | undefined) => (list.find((entry) => entry.id === id) || list[0]).code

/** Builds the format code Excel's Format Cells dialog writes for a category and its options. */
export function buildNumberFormat(category: NumberFormatCategory, options: NumberFormatOptions = {}): string {
  const decimals = clampDecimals(options.decimals)
  switch (category) {
    case 'general': return 'General'
    case 'number': return withNegative(`${options.thousands ? '#,##0' : '0'}${decimalPart(decimals)}`, options.negative)
    case 'currency': {
      const symbol = currencySymbol(options.symbol)
      return withNegative(`${symbol.prefix || ''}#,##0${decimalPart(decimals)}${symbol.suffix || ''}`, options.negative)
    }
    case 'accounting': return accountingCode(decimals, currencySymbol(options.symbol))
    case 'date': return typeCode(DATE_FORMATS, options.type)
    case 'time': return typeCode(TIME_FORMATS, options.type)
    case 'percentage': return `0${decimalPart(decimals)}%`
    case 'fraction': return typeCode(FRACTION_FORMATS, options.type)
    case 'scientific': return `0${decimalPart(decimals)}E+00`
    case 'text': return '@'
    case 'special': return typeCode(SPECIAL_FORMATS, options.type)
    case 'custom': return String(options.code ?? '').trim() || 'General'
  }
}

// ---------------------------------------------------------------------------------------
// Detection
// ---------------------------------------------------------------------------------------

/** Spelling-insensitive key: drops escapes and quotes around single characters, lowercases. */
function comparisonKey(code: string) {
  return splitFormatSections(code.trim()).map((section) => tokenizeFormat(section).map((token) => {
    if (token.kind === 'escaped') return token.value
    if (token.kind === 'quoted' && token.value.length === 1) return token.value
    if (token.kind === 'bracket') return `[${token.value.replace(/\s+/g, '')}]`
    return token.text
  }).join('')).join(';').toLowerCase()
}

/** Looser key for date/time lists: also ignores [$-409]-style locale tags and a trailing ";@". */
function dateKey(code: string) {
  return comparisonKey(code)
    .replace(/\[\$-(?:409|en-us|x-sysdate|x-systime|f800|f400)\]/g, '')
    .replace(/;@$/, '')
}

function firstSectionDecimals(code: string) {
  const tokens = tokenizeFormat(splitFormatSections(code)[0] || '')
  const dot = tokens.findIndex((token) => token.kind === 'char' && token.value === '.')
  if (dot < 0) return 0
  let count = 0
  for (let index = dot + 1; index < tokens.length; index += 1) {
    const token = tokens[index]
    if (token.kind === 'char' && /[0#?]/.test(token.value)) count += 1
    else break
  }
  return count
}

const detectionCache = new Map<string, DetectedNumberFormat>()

/** Recovers the Format Cells category and options behind a code; anything unrecognized is Custom. */
export function detectNumberFormat(code: string | undefined): DetectedNumberFormat {
  const format = String(code ?? '').trim()
  const cached = detectionCache.get(format)
  if (cached) return { category: cached.category, options: { ...cached.options } }
  const result = detectUncached(format)
  if (detectionCache.size > 500) detectionCache.clear()
  detectionCache.set(format, result)
  return { category: result.category, options: { ...result.options } }
}

function detectUncached(format: string): DetectedNumberFormat {
  const base = { ...DEFAULT_NUMBER_FORMAT_OPTIONS }
  if (!format || /^general$/i.test(format)) return { category: 'general', options: base }
  if (format === '@') return { category: 'text', options: base }
  const key = comparisonKey(format)
  const loose = dateKey(format)

  const findType = (list: readonly FormatTypeOption[], compare: (code: string) => string, target: string) => (
    list.find((entry) => compare(entry.code) === target)
  )
  const special = findType(SPECIAL_FORMATS, comparisonKey, key)
  if (special) return { category: 'special', options: { ...base, type: special.id } }
  const fraction = findType(FRACTION_FORMATS, comparisonKey, key)
  if (fraction) return { category: 'fraction', options: { ...base, type: fraction.id } }

  if (isDateTimeFormat(format)) {
    const date = findType(DATE_FORMATS, comparisonKey, key) || findType(DATE_FORMATS, dateKey, loose)
    if (date) return { category: 'date', options: { ...base, type: date.id } }
    const time = findType(TIME_FORMATS, comparisonKey, key) || findType(TIME_FORMATS, dateKey, loose)
    if (time) return { category: 'time', options: { ...base, type: time.id } }
    return { category: 'custom', options: { ...base, code: format } }
  }

  const decimals = firstSectionDecimals(format)
  if (decimals <= MAX_DECIMALS) {
    if (key === comparisonKey(buildNumberFormat('percentage', { decimals }))) return { category: 'percentage', options: { ...base, decimals } }
    if (key === comparisonKey(buildNumberFormat('scientific', { decimals }))) return { category: 'scientific', options: { ...base, decimals } }
    for (const thousands of [false, true]) {
      for (const negative of NEGATIVE_NUMBER_STYLES) {
        if (key === comparisonKey(buildNumberFormat('number', { decimals, thousands, negative: negative.id }))) {
          return { category: 'number', options: { ...base, decimals, thousands, negative: negative.id } }
        }
      }
    }
    for (const symbol of CURRENCY_SYMBOLS) {
      if (symbol.id === 'none') continue
      for (const negative of NEGATIVE_NUMBER_STYLES) {
        if (key === comparisonKey(buildNumberFormat('currency', { decimals, symbol: symbol.id, negative: negative.id }))) {
          return { category: 'currency', options: { ...base, decimals, symbol: symbol.id, negative: negative.id } }
        }
      }
    }
    for (const symbol of CURRENCY_SYMBOLS) {
      if (key === comparisonKey(buildNumberFormat('accounting', { decimals, symbol: symbol.id }))) {
        return { category: 'accounting', options: { ...base, decimals, symbol: symbol.id } }
      }
    }
  }
  return { category: 'custom', options: { ...base, code: format } }
}

// ---------------------------------------------------------------------------------------
// Classification helpers
// ---------------------------------------------------------------------------------------

function bareCharacters(section: string) {
  let bare = ''
  let elapsed = false
  for (const token of tokenizeFormat(section)) {
    if (token.kind === 'char') bare += token.value
    else if (token.kind === 'bracket' && /^(h+|m+|s+)$/i.test(token.value.trim())) elapsed = true
  }
  return { bare: bare.replace(/general/gi, ''), elapsed }
}

/** True when any numeric section of the code renders a date or time. */
export function isDateTimeFormat(code: string | undefined) {
  const sections = splitFormatSections(String(code ?? ''))
  return sections.some((section) => {
    if (section.includes('@') && !/[ymdhs]/i.test(bareCharacters(section).bare)) return false
    const { bare, elapsed } = bareCharacters(section)
    return elapsed || /[ymdhs]/i.test(bare.replace(/e[+-]/gi, '')) || /am\/pm|a\/p/i.test(bare)
  })
}

export function isTextFormat(code: string | undefined) {
  return String(code ?? '').trim() === '@'
}

/** Whether a code is usable: SheetJS can render it for a number and for text. */
export function validateNumberFormat(code: string): { valid: boolean; message?: string } {
  const format = code.trim()
  if (!format) return { valid: false, message: 'Type a number format code.' }
  if (splitFormatSections(format).length > 4) return { valid: false, message: 'A number format can have at most four sections.' }
  if ((format.match(/"/g) || []).length % 2) return { valid: false, message: 'A quoted literal is missing its closing quote.' }
  const opened = (format.match(/\[/g) || []).length
  const closed = (format.match(/]/g) || []).length
  if (opened !== closed) return { valid: false, message: 'A bracketed code such as [Red] is not closed.' }
  const sample = formatScalarDetailed(1234.5, format, '\u0000')
  if (sample.text === '\u0000') return { valid: false, message: 'This number format code isn’t supported.' }
  return { valid: true }
}

// ---------------------------------------------------------------------------------------
// Increase / decrease decimal
// ---------------------------------------------------------------------------------------

const isPlaceholder = (token: FormatToken | undefined) => Boolean(token && token.kind === 'char' && /[0#?]/.test(token.value))
const charToken = (value: string): FormatToken => ({ kind: 'char', text: value, value })

function adjustSection(section: string, delta: number): string {
  const tokens = tokenizeFormat(section)
  const hasPlaceholder = tokens.some(isPlaceholder)
  if (!hasPlaceholder) {
    // Accounting zero section at 0 decimals (`"-"`) gains its first "?" pad.
    const dash = tokens.findIndex((token) => token.kind === 'quoted' && token.value.trim() === '-')
    if (delta > 0 && dash >= 0 && tokens.some((token) => token.kind === 'fill')) {
      tokens.splice(dash + 1, 0, charToken('?'))
      return tokens.map((token) => token.text).join('')
    }
    return section
  }
  // Scientific: only the mantissa (before E+/E-) carries decimals.
  let limit = tokens.length
  for (let index = 0; index < tokens.length - 1; index += 1) {
    const token = tokens[index]
    const next = tokens[index + 1]
    if (token.kind === 'char' && /e/i.test(token.value) && next.kind === 'char' && /[+-]/.test(next.value)) { limit = index; break }
  }
  const dot = tokens.findIndex((token, index) => index < limit && token.kind === 'char' && token.value === '.')
  if (dot >= 0) {
    let end = dot + 1
    while (end < limit && isPlaceholder(tokens[end])) end += 1
    const count = end - dot - 1
    if (delta > 0) {
      if (count >= MAX_DECIMALS) return section
      tokens.splice(end, 0, charToken('0'))
    } else if (count > 1) tokens.splice(end - 1, 1)
    else if (count === 1) tokens.splice(dot, 2)
    else tokens.splice(dot, 1)
    return tokens.map((token) => token.text).join('')
  }
  const placeholders = tokens.slice(0, limit).map((token, index) => (isPlaceholder(token) ? index : -1)).filter((index) => index >= 0)
  if (!placeholders.length) return section
  // Accounting zero section: `"-"??` pads with one "?" per decimal.
  if (placeholders.every((index) => tokens[index].value === '?')) {
    const last = placeholders[placeholders.length - 1]
    if (delta > 0) tokens.splice(last + 1, 0, charToken('?'))
    else tokens.splice(last, 1)
    return tokens.map((token) => token.text).join('')
  }
  if (delta < 0) return section
  const last = placeholders[placeholders.length - 1]
  tokens.splice(last + 1, 0, charToken('.'), charToken('0'))
  return tokens.map((token) => token.text).join('')
}

function generalDecimals(value: unknown) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return { decimals: 0, scientific: false }
  const text = formatScalarDetailed(value, 'General').text
  const scientific = /e[+-]/i.test(text)
  const mantissa = scientific ? text.split(/e/i)[0] : text
  const fraction = mantissa.split('.')[1] || ''
  return { decimals: fraction.length, scientific }
}

/**
 * Excel's Increase / Decrease Decimal on any code — every numeric section of multi-section,
 * currency, accounting, percentage and scientific codes moves together. General starts from
 * the decimals the sample value currently shows. Returns null where decimals don't apply
 * (dates, times, fractions, text).
 */
export function adjustDecimals(code: string | undefined, delta: 1 | -1 | number, sampleValue?: unknown): string | null {
  const format = String(code ?? '').trim()
  const step = delta > 0 ? 1 : -1
  if (!format || /^general$/i.test(format)) {
    const { decimals, scientific } = generalDecimals(sampleValue)
    const next = clampDecimals(decimals + step, 0)
    return scientific ? `0${decimalPart(next)}E+00` : `0${decimalPart(next)}`
  }
  if (isTextFormat(format) || isDateTimeFormat(format)) return null
  const sections = splitFormatSections(format)
  if (sections.some((section) => tokenizeFormat(section).some((token) => token.kind === 'char' && token.value === '/'))) return null
  if (!sections.some((section) => tokenizeFormat(section).some(isPlaceholder))) return null
  return sections.map((section) => (section.includes('@') && !tokenizeFormat(section).some(isPlaceholder) ? section : adjustSection(section, step))).join(';')
}

/** Display text of a type-list entry: date/time entries render the shared sample instant. */
export function formatTypeSample(category: NumberFormatCategory, entry: FormatTypeOption) {
  if (category === 'date' || category === 'time') {
    // Elapsed-time entries show one day plus the sample time ("37:30:55"), as Excel's list does.
    const elapsed = /\[(h+|m+|s+)\]/i.test(entry.code)
    const text = formatScalarDetailed(elapsed ? 1 + (SAMPLE_DATE_SERIAL % 1) : SAMPLE_DATE_SERIAL, entry.code).text
    return entry.system ? `*${text}` : text
  }
  return entry.label
}
