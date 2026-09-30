/**
 * Shared, dependency-free helpers for the data tools (filter, sort, validation, cleanup):
 * A1 addressing, the host abstraction the app supplies, Excel scalar classification,
 * date/time/number parsing, wildcard matching and cell colours.
 */
import type { CellData, SheetData } from '../spreadsheet-types'

export type Scalar = string | number | boolean | null

/** 0-based inclusive rectangle. */
export interface Bounds { top: number; bottom: number; left: number; right: number }

/** A1 address -> new cell (or null to delete). Apply with {@link applyCellChanges}. */
export type CellChanges = Record<string, CellData | null>

/** Reference shifting supplied by the app (the formula engine's shiftFormulaReferences). */
export type ShiftFormula = (formula: string, rowDelta: number, colDelta: number) => string

/**
 * What the data tools need from the app. Coordinates are 0-based.
 * Only `valueAt` and `displayAt` are required; see {@link createSheetHost} for defaults.
 */
export interface DataHost {
  /** Calculated value (formulas resolved). Blank cells return null, undefined or ''. */
  valueAt(row: number, col: number): Scalar | undefined
  /** Formatted text exactly as shown in the grid. */
  displayAt(row: number, col: number): string
  /** Model cell (style, numFmt, formula flags). */
  cellAt?(row: number, col: number): CellData | undefined
  /** Resolved solid fill as '#RRGGBB', or null for no fill. */
  fillColorAt?(row: number, col: number): string | null
  /** Resolved font colour as '#RRGGBB', or null for automatic. */
  fontColorAt?(row: number, col: number): string | null
  /** True when the cell's number format is a date/time format. */
  isDateAt?(row: number, col: number): boolean
  /**
   * Evaluate a formula (no leading '=') as if it lived in (row, col). `override` asks the
   * engine to treat one cell as holding a candidate value (custom validation of a pending edit).
   */
  evaluate?(formula: string, row: number, col: number, override?: { row: number; col: number; value: Scalar }): Scalar | Scalar[][]
  /**
   * Resolve a list-validation source: an A1 range (optionally sheet-qualified), a defined name or
   * a formula such as OFFSET/INDIRECT. Returns the cells row-major, or null when unresolvable.
   */
  resolveReference?(reference: string, origin: { row: number; col: number }): Array<{ value: Scalar; text: string }> | null
  /** Reference shifting (for custom formulas written relative to a range's top-left). */
  shiftFormula?: ShiftFormula
  /** Clock used by relative date filters (tests pin it). */
  now?(): Date
  date1904?: boolean
}

// ---------------------------------------------------------------------------------------------
// A1 addressing
// ---------------------------------------------------------------------------------------------

export const MAX_ROWS = 1_048_576
export const MAX_COLUMNS = 16_384

const columnLabelCache: string[] = []

/** 0-based column index -> 'A'.. 'XFD'. Cached; cheap enough for 100k-row loops. */
export function columnLabel(index: number): string {
  const cached = columnLabelCache[index]
  if (cached !== undefined) return cached
  let value = index + 1
  let result = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    value = Math.floor((value - 1) / 26)
  }
  if (index >= 0 && index < MAX_COLUMNS) columnLabelCache[index] = result
  return result
}

export function columnIndexFromLabel(label: string): number {
  let value = 0
  for (let index = 0; index < label.length; index += 1) value = value * 26 + (label.charCodeAt(index) & 0xdf) - 64
  return value - 1
}

export function addressOf(row: number, col: number): string {
  return columnLabel(col) + (row + 1)
}

const ADDRESS_PATTERN = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/

export function parseAddress(address: string): { row: number; col: number } | null {
  const match = ADDRESS_PATTERN.exec(address.trim())
  if (!match) return null
  const row = Number(match[2]) - 1
  const col = columnIndexFromLabel(match[1])
  if (row < 0 || row >= MAX_ROWS || col < 0 || col >= MAX_COLUMNS) return null
  return { row, col }
}

/** Strip an optional sheet prefix ('Sheet 2'!A1:B3 or Sheet2!A1). */
export function splitSheetPrefix(reference: string): { sheet: string | null; ref: string } {
  const text = reference.trim().replace(/^=/, '')
  const bang = text.lastIndexOf('!')
  if (bang < 0) return { sheet: null, ref: text }
  let sheet = text.slice(0, bang)
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'")
  return { sheet, ref: text.slice(bang + 1) }
}

/**
 * Parse 'A1', 'A1:F200', '$A$1:$F$200', 'A:C' or '2:5' (whole columns/rows are clipped to
 * `limits`). A sheet prefix is ignored. Returns null for anything else.
 */
export function parseRange(reference: string, limits: { rows?: number; cols?: number } = {}): Bounds | null {
  if (typeof reference !== 'string') return null
  const { ref } = splitSheetPrefix(reference)
  const [first, second = first] = ref.split(':')
  const a = parseAddress(first)
  const b = parseAddress(second)
  if (a && b) return { top: Math.min(a.row, b.row), bottom: Math.max(a.row, b.row), left: Math.min(a.col, b.col), right: Math.max(a.col, b.col) }
  const columns = /^\$?([A-Za-z]{1,3})$/.exec(first.trim()) && /^\$?([A-Za-z]{1,3})$/.exec(second.trim())
  if (columns) {
    const left = columnIndexFromLabel(first.replace(/\$/g, '').trim())
    const right = columnIndexFromLabel(second.replace(/\$/g, '').trim())
    return { top: 0, bottom: (limits.rows ?? MAX_ROWS) - 1, left: Math.min(left, right), right: Math.max(left, right) }
  }
  const rows = /^\$?(\d{1,7})$/.exec(first.trim()) && /^\$?(\d{1,7})$/.exec(second.trim())
  if (rows) {
    const top = Number(first.replace(/\$/g, '')) - 1
    const bottom = Number(second.replace(/\$/g, '')) - 1
    return { top: Math.min(top, bottom), bottom: Math.max(top, bottom), left: 0, right: (limits.cols ?? MAX_COLUMNS) - 1 }
  }
  return null
}

export function formatRange(bounds: Bounds): string {
  const start = addressOf(bounds.top, bounds.left)
  const end = addressOf(bounds.bottom, bounds.right)
  return start === end ? start : `${start}:${end}`
}

export function boundsIntersect(a: Bounds, b: Bounds): boolean {
  return a.bottom >= b.top && a.top <= b.bottom && a.right >= b.left && a.left <= b.right
}

export function boundsContain(outer: Bounds, inner: Bounds): boolean {
  return inner.top >= outer.top && inner.bottom <= outer.bottom && inner.left >= outer.left && inner.right <= outer.right
}

/** Parse an ExcelJS/SheetJS autoFilter model ('A1:F9', { ref }, { from, to }). */
export function autoFilterBounds(autoFilter: unknown): Bounds | null {
  if (typeof autoFilter === 'string') return parseRange(autoFilter)
  if (!autoFilter || typeof autoFilter !== 'object') return null
  const model = autoFilter as { ref?: unknown; Ref?: unknown; from?: unknown; to?: unknown }
  if (typeof model.ref === 'string') return parseRange(model.ref)
  if (typeof model.Ref === 'string') return parseRange(model.Ref)
  const endpoint = (value: unknown) => {
    if (typeof value === 'string') return parseAddress(value)
    if (!value || typeof value !== 'object') return null
    const item = value as { row?: unknown; column?: unknown }
    const row = Number(item.row)
    const col = Number(item.column)
    return Number.isInteger(row) && Number.isInteger(col) && row >= 1 && col >= 1 ? { row: row - 1, col: col - 1 } : null
  }
  const from = endpoint(model.from)
  const to = endpoint(model.to)
  if (!from || !to) return null
  return { top: Math.min(from.row, to.row), bottom: Math.max(from.row, to.row), left: Math.min(from.col, to.col), right: Math.max(from.col, to.col) }
}

// ---------------------------------------------------------------------------------------------
// Cell access
// ---------------------------------------------------------------------------------------------

export function isBlankScalar(value: unknown): boolean {
  return value === null || value === undefined || value === ''
}

export function cellHasContent(cell: CellData | undefined): boolean {
  if (!cell) return false
  return Boolean(cell.formula) || !isBlankScalar(cell.value)
}

/** Apply a change set produced by the data tools to a cell map (mutates `cells`). */
export function applyCellChanges(cells: Record<string, CellData>, changes: CellChanges): void {
  for (const address of Object.keys(changes)) {
    const cell = changes[address]
    if (cell) cells[address] = cell
    else delete cells[address]
  }
}

/** Remove cached calculation state from a cell whose formula has been rewritten or moved. */
export function stripFormulaCache(cell: CellData): CellData {
  delete cell.result
  delete cell.resultType
  delete cell.display
  delete cell.formulaType
  delete cell.formulaRange
  delete cell.dynamicFormula
  delete (cell as CellData & { sharedFormulaMaster?: string }).sharedFormulaMaster
  return cell
}

/**
 * The cell to store at a new location. A cell without a formula is returned as-is (moves are
 * permutations, so each object still ends up at exactly one address, and the model treats cells
 * as immutable); a formula cell is copied with its references shifted by the move distance.
 */
export function moveCell(cell: CellData, rowDelta: number, colDelta: number, shiftFormula?: ShiftFormula): CellData {
  if (!cell.formula || (!rowDelta && !colDelta)) return cell
  const copy: CellData = { ...cell }
  if (shiftFormula) copy.formula = shiftFormula(cell.formula, rowDelta, colDelta)
  // The engine recalculates moved formulas; a stale cached result must not be saved.
  return stripFormulaCache(copy)
}

/** A host for one sheet backed by the stored values/results (for tests and fallbacks). */
export function createSheetHost(sheet: Pick<SheetData, 'cells'>, overrides: Partial<DataHost> = {}): DataHost {
  const cellAt = (row: number, col: number) => sheet.cells[addressOf(row, col)]
  const valueAt = (row: number, col: number): Scalar | undefined => {
    const cell = cellAt(row, col)
    if (!cell) return null
    const value = cell.formula ? cell.result : cell.value
    return value === undefined ? null : typeof value === 'object' && value !== null ? cell.display ?? null : value
  }
  return {
    cellAt,
    valueAt,
    displayAt: (row, col) => {
      const cell = cellAt(row, col)
      if (!cell) return ''
      if (typeof cell.display === 'string') return cell.display
      const value = valueAt(row, col)
      if (value === null || value === undefined) return ''
      if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
      return String(value)
    },
    fillColorAt: (row, col) => cellFillColor(cellAt(row, col)),
    fontColorAt: (row, col) => cellFontColor(cellAt(row, col)),
    isDateAt: (row, col) => {
      const cell = cellAt(row, col)
      return Boolean(cell && (cell.type === 'date' || cell.resultType === 'date' || isDateNumberFormat(cell.numFmt || cell.style?.numFmt)))
    },
    ...overrides,
  }
}

export function hostIsDate(host: DataHost, row: number, col: number): boolean {
  if (host.isDateAt) return host.isDateAt(row, col)
  const cell = host.cellAt?.(row, col)
  return Boolean(cell && (cell.type === 'date' || cell.resultType === 'date' || isDateNumberFormat(cell.numFmt || cell.style?.numFmt)))
}

export function hostFillColor(host: DataHost, row: number, col: number): string | null {
  if (host.fillColorAt) return normalizeColor(host.fillColorAt(row, col))
  return cellFillColor(host.cellAt?.(row, col))
}

export function hostFontColor(host: DataHost, row: number, col: number): string | null {
  if (host.fontColorAt) return normalizeColor(host.fontColorAt(row, col))
  return cellFontColor(host.cellAt?.(row, col))
}

// ---------------------------------------------------------------------------------------------
// Scalar classification (Excel sort / filter semantics)
// ---------------------------------------------------------------------------------------------

export const ERROR_VALUES = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#CIRC!', '#PARSE!', '#SPILL!', '#CALC!',
  '#GETTING_DATA', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!',
])

export type ScalarKind = 'blank' | 'number' | 'text' | 'boolean' | 'error'

export function scalarKind(value: unknown): ScalarKind {
  if (value === null || value === undefined || value === '') return 'blank'
  if (typeof value === 'number') return Number.isFinite(value) ? 'number' : 'error'
  if (typeof value === 'boolean') return 'boolean'
  if (typeof value === 'string') return ERROR_VALUES.has(value) ? 'error' : 'text'
  return 'text'
}

/** Accent-sensitive, case-insensitive (Excel's default sort / comparison). */
export const TEXT_COLLATOR = new Intl.Collator(undefined, { sensitivity: 'accent', numeric: false })
/** Case-sensitive with lowercase first (Excel's "Case sensitive" sort option). */
export const CASE_COLLATOR = new Intl.Collator(undefined, { sensitivity: 'variant', caseFirst: 'lower', numeric: false })

export function foldText(text: string): string {
  return text.toLocaleLowerCase()
}

const FAST_TEXT = /^[a-z0-9 ]*$/
let fastTextOrder: boolean | null = null

/**
 * Whether plain [A-Za-z0-9 ] strings collate exactly like their lowercased code units under
 * TEXT_COLLATOR in this runtime's locale. Probed once: locales that tailor basic Latin
 * (Danish "aa", Czech "ch", Lithuanian "y", Turkish dotted I, …) fail the probe and keep the
 * collator for everything.
 */
function fastTextOrderValid(): boolean {
  if (fastTextOrder !== null) return fastTextOrder
  const alphabet = ' 0123456789abcdefghijklmnopqrstuvwxyz'
  const probes: string[] = ['a b', 'ab', 'a  b', 'a0', 'a 0', 'aab', 'aa b']
  for (const first of alphabet) {
    probes.push(first)
    for (const second of alphabet) probes.push(first + second)
  }
  const byCode = [...probes].sort((a, b) => (a < b ? -1 : a > b ? 1 : 0))
  const byCollator = [...probes].sort(TEXT_COLLATOR.compare)
  let valid = byCode.every((value, index) => value === byCollator[index])
  for (const letter of 'abcdefghijklmnopqrstuvwxyz') {
    if (TEXT_COLLATOR.compare(letter, letter.toUpperCase()) !== 0) valid = false
  }
  fastTextOrder = valid
  return valid
}

/**
 * Precomputed keys for {@link compareCollated}: the lowercased string for plain alphanumeric
 * text (compared by code unit, ~10x faster than Intl.Collator) or null to use the collator.
 */
export function collationKeys(texts: readonly string[], caseSensitive = false): Array<string | null> {
  const keys: Array<string | null> = new Array(texts.length)
  const fast = !caseSensitive && fastTextOrderValid()
  for (let index = 0; index < texts.length; index += 1) {
    if (!fast) { keys[index] = null; continue }
    const lowered = texts[index].toLowerCase()
    keys[index] = FAST_TEXT.test(lowered) ? lowered : null
  }
  return keys
}

/** Collator-equivalent comparison using keys from {@link collationKeys}. */
export function compareCollated(a: string, keyA: string | null, b: string, keyB: string | null, caseSensitive = false): number {
  if (keyA !== null && keyB !== null) return keyA < keyB ? -1 : keyA > keyB ? 1 : 0
  return (caseSensitive ? CASE_COLLATOR : TEXT_COLLATOR).compare(a, b)
}

/** Sort strings in place in Excel's text order (accent-sensitive; case-insensitive unless asked). */
export function sortTexts(texts: string[], caseSensitive = false): string[] {
  const keys = collationKeys(texts, caseSensitive)
  const indices = texts.map((_, index) => index)
  indices.sort((a, b) => compareCollated(texts[a], keys[a], texts[b], keys[b], caseSensitive) || a - b)
  const sorted = indices.map((index) => texts[index])
  for (let index = 0; index < sorted.length; index += 1) texts[index] = sorted[index]
  return texts
}

// ---------------------------------------------------------------------------------------------
// Dates, times and numbers
// ---------------------------------------------------------------------------------------------

const MS_PER_DAY = 86_400_000
const EPOCH_1900 = Date.UTC(1899, 11, 30)
const OFFSET_1904 = 1462

export function serialFromYMD(year: number, month: number, day: number, date1904 = false): number {
  return (Date.UTC(year, month - 1, day) - EPOCH_1900) / MS_PER_DAY - (date1904 ? OFFSET_1904 : 0)
}

export function ymdFromSerial(serial: number, date1904 = false): { year: number; month: number; day: number; weekday: number } {
  const date = new Date(EPOCH_1900 + Math.floor(serial + (date1904 ? OFFSET_1904 : 0)) * MS_PER_DAY)
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate(), weekday: date.getUTCDay() }
}

/** Serial of a JS Date's local calendar day (+ time of day as a fraction). */
export function serialFromLocalDate(date: Date, date1904 = false, withTime = false): number {
  const base = serialFromYMD(date.getFullYear(), date.getMonth() + 1, date.getDate(), date1904)
  if (!withTime) return base
  return base + (date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds()) / 86_400
}

/** Serial for a UTC instant (how ExcelJS stores dates). */
export function serialFromUtcDate(date: Date, date1904 = false): number {
  return (date.getTime() - EPOCH_1900) / MS_PER_DAY - (date1904 ? OFFSET_1904 : 0)
}

export function utcDateFromSerial(serial: number, date1904 = false): Date {
  return new Date(EPOCH_1900 + (serial + (date1904 ? OFFSET_1904 : 0)) * MS_PER_DAY)
}

export function formatSerialDate(serial: number, date1904 = false): string {
  const { year, month, day } = ymdFromSerial(serial, date1904)
  return `${month}/${day}/${year}`
}

export function formatSerialTime(fraction: number): string {
  let seconds = Math.round((((fraction % 1) + 1) % 1) * 86_400)
  if (seconds >= 86_400) seconds = 0
  const hours = Math.floor(seconds / 3600)
  const minutes = Math.floor((seconds % 3600) / 60)
  const secs = seconds % 60
  const hour12 = hours % 12 === 0 ? 12 : hours % 12
  return `${hour12}:${String(minutes).padStart(2, '0')}${secs ? `:${String(secs).padStart(2, '0')}` : ''} ${hours < 12 ? 'AM' : 'PM'}`
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']

function monthFromName(name: string): number {
  const key = name.toLocaleLowerCase().slice(0, 3)
  const index = MONTHS.indexOf(key)
  return index < 0 ? 0 : index + 1
}

function validYMD(year: number, month: number, day: number): boolean {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return false
  if (month < 1 || month > 12 || day < 1 || day > 31 || year < 1900 || year > 9999) return false
  const date = new Date(Date.UTC(year, month - 1, day))
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day
}

function expandYear(year: number, digits: number): number {
  if (digits > 2) return year
  return year < 30 ? 2000 + year : 1900 + year
}

/** Parse a time of day ('9:30', '9:30:15 PM', '9 am', '21:05:00.5') into a day fraction. */
export function parseTimeText(input: string): number | null {
  const text = input.trim()
  const match = /^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}(?:\.\d+)?))?\s*([ap]\.?m?\.?)?$/i.exec(text)
  if (!match || (match[2] === undefined && !match[4])) return null
  let hours = Number(match[1])
  const minutes = Number(match[2] || 0)
  const seconds = Number(match[3] || 0)
  const meridiem = match[4]?.toLocaleLowerCase()
  if (meridiem) {
    if (hours < 1 || hours > 12) return null
    if (meridiem.startsWith('p') && hours !== 12) hours += 12
    if (meridiem.startsWith('a') && hours === 12) hours = 0
  }
  if (hours > 23 || minutes > 59 || seconds >= 60) return null
  return (hours * 3600 + minutes * 60 + seconds) / 86_400
}

/**
 * Parse a date typed by a user (US order): 2024-01-15, 1/15/2024, 1/15/24, 1-15-2024,
 * 15-Jan-2024, 15 Jan 2024, Jan 15, 2024, January 15 2024 — each optionally followed by a time.
 * Returns an Excel serial (with fraction when a time is present) or null.
 */
export function parseDateText(input: string, date1904 = false): number | null {
  const text = input.trim().replace(/\s+/g, ' ')
  if (!text) return null
  let rest = ''
  let year = 0
  let month = 0
  let day = 0
  let match: RegExpExecArray | null
  if ((match = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](.*))?$/.exec(text))) {
    year = Number(match[1]); month = Number(match[2]); day = Number(match[3]); rest = match[4] || ''
  } else if ((match = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2}|\d{4})(?: (.*))?$/.exec(text))) {
    month = Number(match[1]); day = Number(match[2]); year = expandYear(Number(match[3]), match[3].length); rest = match[4] || ''
  } else if ((match = /^(\d{4})\/(\d{1,2})\/(\d{1,2})(?: (.*))?$/.exec(text))) {
    year = Number(match[1]); month = Number(match[2]); day = Number(match[3]); rest = match[4] || ''
  } else if ((match = /^(\d{1,2})[- ]([A-Za-z]{3,9})\.?[- ,]*(\d{2}|\d{4})(?: (.*))?$/.exec(text))) {
    day = Number(match[1]); month = monthFromName(match[2]); year = expandYear(Number(match[3]), match[3].length); rest = match[4] || ''
  } else if ((match = /^([A-Za-z]{3,9})\.? (\d{1,2})(?:st|nd|rd|th)?,? (\d{2}|\d{4})(?: (.*))?$/.exec(text))) {
    month = monthFromName(match[1]); day = Number(match[2]); year = expandYear(Number(match[3]), match[3].length); rest = match[4] || ''
  } else {
    return null
  }
  if (!validYMD(year, month, day)) return null
  let serial = serialFromYMD(year, month, day, date1904)
  if (rest.trim()) {
    const time = parseTimeText(rest.replace(/^T/, '').replace(/(?:Z|[+-]\d{2}:?\d{2})$/i, ''))
    if (time === null) return null
    serial += time
  }
  return serial
}

/**
 * Parse a number the way a user types it: '1,234.5', '$1,234', '(12)', '-5', '12%', '1e3'.
 * Returns null for anything that is not wholly numeric.
 */
export function parseNumberText(input: string): number | null {
  let text = input.trim()
  if (!text) return null
  let negative = false
  if (/^\(.*\)$/.test(text)) { negative = true; text = text.slice(1, -1).trim() }
  let percent = false
  if (text.endsWith('%')) { percent = true; text = text.slice(0, -1).trim() }
  text = text.replace(/^([+-]?)\s*[$€£¥₹]\s*/, '$1').replace(/\s*[$€£¥₹]$/, '')
  if (/^[+-]?(\d{1,3}(,\d{3})+)(\.\d*)?$/.test(text)) text = text.replace(/,/g, '')
  if (!/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(text)) return null
  let value = Number(text)
  if (!Number.isFinite(value)) return null
  if (percent) value /= 100
  return negative ? -value : value
}

/** Whether an Excel number format shows a date and/or time. */
export function isDateNumberFormat(format: string | undefined): boolean {
  if (!format || format === 'General' || format === '@') return false
  const stripped = format
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/\[(?:h+|m+|s+)\]/gi, 'h')
    .replace(/\[[^\]]*\]/g, '')
    .replace(/_.|\*./g, '')
  return /[dmyhs]/i.test(stripped.split(';')[0] || '') && !/^[#0.,%E+\- ]*$/i.test(stripped)
}

// ---------------------------------------------------------------------------------------------
// Wildcards
// ---------------------------------------------------------------------------------------------

export function hasWildcards(pattern: string): boolean {
  return /(^|[^~])[*?]/.test(pattern)
}

/** Excel wildcard pattern (* ? with ~ escape) -> case-insensitive RegExp. */
export function wildcardRegExp(pattern: string, anchorStart = true, anchorEnd = true): RegExp {
  let source = ''
  for (let index = 0; index < pattern.length; index += 1) {
    const character = pattern[index]
    if (character === '~' && index + 1 < pattern.length && '*?~'.includes(pattern[index + 1])) {
      source += `\\${pattern[index + 1]}`
      index += 1
    } else if (character === '*') source += '[\\s\\S]*'
    else if (character === '?') source += '[\\s\\S]'
    else source += character.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')
  }
  return new RegExp(`${anchorStart ? '^' : ''}${source}${anchorEnd ? '$' : ''}`, 'i')
}

// ---------------------------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------------------------

const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00',
  '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
  '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF',
  '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300',
  '993300', '993366', '333399', '333333',
]

function tint(hex: string, amount: number): string {
  if (!Number.isFinite(amount) || amount === 0) return hex
  const clamped = Math.max(-1, Math.min(1, amount))
  return [0, 2, 4].map((offset) => {
    const channel = Number.parseInt(hex.slice(offset, offset + 2), 16)
    const next = clamped < 0 ? channel * (1 + clamped) : channel + (255 - channel) * clamped
    return Math.round(Math.max(0, Math.min(255, next))).toString(16).padStart(2, '0')
  }).join('').toUpperCase()
}

/** Normalise '#rgb', 'RRGGBB', 'AARRGGBB' or an ExcelJS colour object to '#RRGGBB' (or null). */
export function normalizeColor(value: unknown): string | null {
  if (value === null || value === undefined || value === '') return null
  let hex = ''
  let amount = 0
  if (typeof value === 'string') hex = value
  else if (typeof value === 'object') {
    const color = value as { argb?: string; rgb?: string; theme?: number; indexed?: number; tint?: number; auto?: boolean }
    hex = color.argb || color.rgb || ''
    amount = Number(color.tint) || 0
    if (!hex && Number.isInteger(color.theme)) hex = OFFICE_THEME[Number(color.theme)] || ''
    if (!hex && Number.isInteger(color.indexed)) hex = INDEXED[Number(color.indexed)] || ''
  }
  hex = hex.trim().replace(/^#/, '')
  if (/^[0-9a-f]{3}$/i.test(hex)) hex = hex.split('').map((character) => character + character).join('')
  if (hex.length === 8) hex = hex.slice(2)
  if (!/^[0-9a-f]{6}$/i.test(hex)) return null
  return `#${tint(hex.toUpperCase(), amount)}`
}

export function cellFillColor(cell: CellData | undefined): string | null {
  const fill = cell?.style?.fill
  if (!fill) return null
  const pattern = String(fill.pattern || '').toLocaleLowerCase()
  if (pattern === 'none') return null
  if (String(fill.type || '').toLocaleLowerCase() === 'gradient') return normalizeColor(fill.stops?.[0]?.color)
  return normalizeColor(fill.color || fill.fgColor) || (pattern && pattern !== 'solid' ? normalizeColor(fill.bgColor) : null)
}

export function cellFontColor(cell: CellData | undefined): string | null {
  const color = cell?.style?.font?.color
  if (!color || (typeof color === 'object' && color.auto)) return null
  const normalized = normalizeColor(color)
  return normalized === '#000000' ? null : normalized
}
