/**
 * Excel-style interpretation of typed cell input: numbers with thousands separators,
 * currency, accounting negatives, percentages, fractions, dates, times, date-times,
 * booleans, and error literals. Returns the stored value and, when Excel would apply
 * one, the number format the entry implies.
 */

import { excelSerialFromUtcTime, parseFormula } from './formulas'
import type { FormulaNode } from './formulas'

export interface ParsedInput {
  value: string | number | boolean
  numFmt?: string
  type?: 'date' | 'error'
}

const ERROR_LITERALS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#SPILL!', '#CALC!'])
const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december']
const CURRENCY = '[$€£¥₹₩₽]'

function monthIndex(token: string): number | null {
  const lower = token.toLowerCase().replace(/\.$/, '')
  if (lower.length < 3) return null
  const index = MONTHS.findIndex((month) => month.startsWith(lower))
  if (index < 0) return null
  // "Mar" is March, "Ma" is ambiguous and rejected above; full names must match exactly.
  if (lower.length > 3 && !MONTHS[index].startsWith(lower)) return null
  return index
}

function serial(year: number, month: number, day: number): number | null {
  const fullYear = year < 100 ? (year < 30 ? 2000 + year : 1900 + year) : year
  if (fullYear < 1900 || fullYear > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null
  // Excel's 1900 date system has a 1900-02-29 (serial 60); 1900-01-01 is serial 1.
  if (fullYear === 1900 && month === 2 && day === 29) return 60
  const timestamp = Date.UTC(fullYear, month - 1, day)
  const date = new Date(timestamp)
  if (date.getUTCFullYear() !== fullYear || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null
  return excelSerialFromUtcTime(timestamp)
}

function timeFraction(hours: number, minutes: number, seconds: number, meridiem?: string): number | null {
  let hour = hours
  if (meridiem) {
    if (hour < 1 || hour > 12) return null
    const pm = /^p/i.test(meridiem)
    if (pm && hour !== 12) hour += 12
    if (!pm && hour === 12) hour = 0
  }
  if (hour > 23 || minutes > 59 || seconds >= 60) return null
  return (hour * 3600 + minutes * 60 + seconds) / 86_400
}

interface ParsedDate { serial: number; numFmt: string }

function parseDate(text: string, now = new Date()): ParsedDate | null {
  const currentYear = now.getFullYear()
  let match = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(text) || /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/.exec(text)
  if (match) {
    const value = serial(Number(match[1]), Number(match[2]), Number(match[3]))
    return value === null ? null : { serial: value, numFmt: 'yyyy-mm-dd' }
  }
  match = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(text) || /^(\d{1,2})-(\d{1,2})-(\d{2}|\d{4})$/.exec(text)
  if (match) {
    const value = serial(Number(match[3]), Number(match[1]), Number(match[2]))
    return value === null ? null : { serial: value, numFmt: 'm/d/yyyy' }
  }
  match = /^(\d{1,2})\/(\d{1,2})$/.exec(text)
  if (match) {
    const value = serial(currentYear, Number(match[1]), Number(match[2]))
    return value === null ? null : { serial: value, numFmt: 'd-mmm' }
  }
  match = /^(\d{1,2})[-\s]([A-Za-z]{3,9}\.?)[-\s,]*(\d{2}|\d{4})?$/.exec(text)
  if (match) {
    const month = monthIndex(match[2])
    if (month === null) return null
    const value = serial(match[3] ? Number(match[3]) : currentYear, month + 1, Number(match[1]))
    return value === null ? null : { serial: value, numFmt: match[3] ? 'd-mmm-yy' : 'd-mmm' }
  }
  match = /^([A-Za-z]{3,9}\.?)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s*(\d{4})?$/.exec(text)
  if (match) {
    const month = monthIndex(match[1])
    if (month === null) return null
    const value = serial(match[3] ? Number(match[3]) : currentYear, month + 1, Number(match[2]))
    return value === null ? null : { serial: value, numFmt: match[3] ? 'mmm d, yyyy' : 'd-mmm' }
  }
  match = /^([A-Za-z]{3,9}\.?)[-\s](\d{2}|\d{4})$/.exec(text)
  if (match) {
    const month = monthIndex(match[1])
    if (month === null) return null
    const value = serial(Number(match[2]), month + 1, 1)
    return value === null ? null : { serial: value, numFmt: 'mmm-yy' }
  }
  return null
}

function parseTime(text: string): { fraction: number; numFmt: string } | null {
  const match = /^(\d{1,2}):(\d{2})(?::(\d{2}(?:\.\d+)?))?\s*([AaPp]\.?[Mm]?\.?)?$/.exec(text) || /^(\d{1,2})()()\s*([AaPp]\.?[Mm]\.?)$/.exec(text)
  if (!match) return null
  const fraction = timeFraction(Number(match[1]), Number(match[2] || 0), Number(match[3] || 0), match[4])
  if (fraction === null) return null
  const numFmt = match[4] ? (match[3] ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : match[3] ? 'h:mm:ss' : 'h:mm'
  return { fraction, numFmt }
}

/** Interpret typed text the way Excel does. Returns null for plain text. */
export function parseCellInput(raw: string, now = new Date()): ParsedInput | null {
  const text = raw.trim()
  if (!text) return null
  const upper = text.toUpperCase()
  if (ERROR_LITERALS.has(upper)) return { value: upper, type: 'error' }
  if (upper === 'TRUE' || upper === 'FALSE') return { value: upper === 'TRUE' }

  // Percentages: 15%, -2.5 %, 1,200%
  let match = /^([+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|[+-]?\.\d+)(?:e[+-]?\d+)?\s*%$/i.exec(text)
  if (match) {
    const number = Number(text.replace(/[,%\s]/g, ''))
    if (Number.isFinite(number)) {
      const decimals = /\.(\d+)/.exec(match[1])?.[1].length || 0
      return { value: number / 100, numFmt: decimals ? `0.${'0'.repeat(decimals)}%` : '0%' }
    }
  }

  // Currency: $1,234.50  -$5  ($5)  $-5  €12
  match = new RegExp(`^(\\()?([+-])?\\s*(${CURRENCY})\\s*([+-])?((?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?|\\.\\d+)(\\))?$`).exec(text)
  if (match && Boolean(match[1]) === Boolean(match[6])) {
    const magnitude = Number(match[5].replace(/,/g, ''))
    if (Number.isFinite(magnitude)) {
      const negative = Boolean(match[1]) || match[2] === '-' || match[4] === '-'
      const decimals = /\.(\d+)$/.exec(match[5])?.[1].length || 0
      const symbol = match[3]
      const body = `#,##0${decimals ? `.${'0'.repeat(Math.min(decimals, 2) || decimals)}` : ''}`
      return { value: negative ? -magnitude : magnitude, numFmt: `${symbol === '$' ? '$' : `"${symbol}"`}${body}` }
    }
  }

  // Thousands separators: 1,234  -12,345.67
  match = /^([+-])?(\d{1,3}(?:,\d{3})+)(\.\d+)?$/.exec(text)
  if (match) {
    const number = Number(text.replace(/,/g, ''))
    if (Number.isFinite(number)) {
      const decimals = match[3] ? match[3].length - 1 : 0
      return { value: number, numFmt: decimals ? `#,##0.${'0'.repeat(decimals)}` : '#,##0' }
    }
  }

  // Accounting negative: (123.45)
  match = /^\((\d+(?:\.\d+)?)\)$/.exec(text)
  if (match) return { value: -Number(match[1]) }

  // Fractions: 1 1/2, 0 3/4 (a bare 1/2 is a date, as in Excel)
  match = /^([+-])?(\d+)\s+(\d+)\/(\d+)$/.exec(text)
  if (match && Number(match[4]) !== 0 && Number(match[3]) < Number(match[4])) {
    const value = Number(match[2]) + Number(match[3]) / Number(match[4])
    const denominatorDigits = match[4].length
    return { value: match[1] === '-' ? -value : value, numFmt: denominatorDigits > 1 ? '# ??/??' : '# ?/?' }
  }

  // Date and time: 1/15/2024 10:30, 2024-01-15 14:05:00, Jan 5, 2024 9:00 PM
  match = /^(.+?)[\sT]+(\d{1,2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:\s*[AaPp]\.?[Mm]?\.?)?)$/.exec(text)
  if (match) {
    const date = parseDate(match[1].trim(), now)
    const time = parseTime(match[2].trim())
    if (date && time) return { value: date.serial + time.fraction, numFmt: 'm/d/yyyy h:mm', type: 'date' }
  }
  const date = parseDate(text, now)
  if (date) return { value: date.serial, numFmt: date.numFmt, type: 'date' }
  const time = parseTime(text)
  if (time) return { value: time.fraction, numFmt: time.numFmt }

  // Plain numbers. Leading zeros ("007") stay text so identifiers survive.
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i.test(text) && !/^[-+]?0\d/.test(text)) {
    const number = Number(text)
    if (Number.isFinite(number)) return { value: number }
  }
  return null
}

function readsCells(node: FormulaNode): boolean {
  switch (node.kind) {
    case 'reference':
    case 'range':
    case 'wholeRange':
    case 'spill':
    case 'structured':
    case 'rangeOp':
    case 'call':
    case 'invoke':
      return true
    case 'unary':
      return readsCells(node.operand)
    case 'binary':
      return readsCells(node.left) || readsCells(node.right)
    default:
      return false
  }
}

/**
 * The formula an entry typed with a leading "+" or "-" stands for, as in Excel's numeric-keypad
 * and Lotus-style entry: "+B1+C1" is =+B1+C1 and "-A1*2" is =-A1*2. Only entries that read a
 * cell or call a function become formulas, so "- item" and "+1-555-0100" stay text. Returns the
 * formula without "=", or null.
 */
export function formulaFromSignedEntry(raw: string): string | null {
  const text = raw.trim()
  if (!/^[+-]/.test(text) || parseCellInput(text) !== null) return null
  const node = parseFormula(text)
  return typeof node !== 'string' && readsCells(node) ? text : null
}

/**
 * Whether stored text would turn into something else (a number, date, logical, error, or
 * formula) if it were typed again as shown, so the editor must show it with Excel's leading
 * apostrophe to keep it text.
 */
export function textNeedsQuotePrefix(text: string): boolean {
  if (text.startsWith("'") || text.startsWith('=')) return true
  return parseCellInput(text) !== null || formulaFromSignedEntry(text) !== null
}

const DATE_FUNCTIONS: Record<string, string> = {
  TODAY: 'm/d/yyyy',
  DATE: 'm/d/yyyy',
  EDATE: 'm/d/yyyy',
  EOMONTH: 'm/d/yyyy',
  WORKDAY: 'm/d/yyyy',
  'WORKDAY.INTL': 'm/d/yyyy',
  DATEVALUE: 'm/d/yyyy',
  NOW: 'm/d/yyyy h:mm',
  TIME: 'h:mm AM/PM',
  TIMEVALUE: 'h:mm AM/PM',
}

/**
 * The number format Excel gives a newly entered formula in a General cell: date/time
 * functions get a date/time format, and simple formulas inherit the format of their first
 * referenced cell (=A1+7 on a date stays a date; =B2*1.1 on currency stays currency).
 */
export function inferFormulaNumberFormat(formula: string, formatOf: (address: string, sheet?: string) => string | undefined): string | undefined {
  const source = formula.replace(/^=/, '').trim()
  const call = /^(?:_xlfn\.)?([A-Z][A-Z0-9.]*)\s*\(/i.exec(source)
  if (call) {
    const name = call[1].toUpperCase()
    if (DATE_FUNCTIONS[name]) return DATE_FUNCTIONS[name]
    if (!['SUM', 'AVERAGE', 'MIN', 'MAX', 'MEDIAN', 'ROUND', 'ROUNDUP', 'ROUNDDOWN', 'SUBTOTAL', 'SUMIF', 'SUMIFS', 'AVERAGEIF', 'AVERAGEIFS', 'MINIFS', 'MAXIFS', 'ABS', 'INDEX', 'VLOOKUP', 'HLOOKUP', 'XLOOKUP', 'LOOKUP', 'IF', 'IFERROR', 'MROUND', 'CEILING', 'FLOOR', 'LARGE', 'SMALL'].includes(name)) return undefined
  }
  const reference = /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!)?\$?([A-Za-z]{1,3})\$?(\d+)/.exec(source)
  if (!reference) return undefined
  // Counting functions and comparisons never inherit a format.
  if (/^(?:COUNT|COUNTA|COUNTIF|COUNTIFS|ROWS|COLUMNS|LEN|MATCH)\b/i.test(source) || /[<>=]/.test(source)) return undefined
  // Only inherit when the referenced value is used arithmetically, not divided/multiplied by another reference.
  if (/[*/^]\s*\$?[A-Za-z]{1,3}\$?\d/.test(source.slice(reference.index + reference[0].length))) return undefined
  const sheet = reference[1]?.replace(/''/g, "'") ?? reference[2]
  const format = formatOf(`${reference[3].toUpperCase()}${reference[4]}`, sheet)
  if (!format || format === 'General' || format === '@') return undefined
  return format
}
