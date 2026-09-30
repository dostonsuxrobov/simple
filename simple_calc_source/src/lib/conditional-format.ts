/**
 * Conditional formatting: an Excel-compatible evaluation engine plus pure helpers that author
 * and manage rules in the ExcelJS conditional-formatting model (the shape stored in
 * `SheetData.conditionalFormattings` and written back to .xlsx by ExcelJS).
 *
 * Model: `[{ ref: 'A1:C10 E1:E5', rules: [{ type, priority, stopIfTrue?, style?, ... }] }]`.
 * Rule shapes follow ExcelJS 4.4 (as patched by electron/conditional-format-exceljs.cjs):
 *   expression   { formulae: [f] }
 *   cellIs       { operator, formulae: [f1, f2?] }
 *   containsText { operator: 'containsText' | 'containsBlanks' | 'notContainsBlanks' | 'containsErrors' | 'notContainsErrors', text?, formulae }
 *   notContainsText { operator: 'notContains', text, formulae } / beginsWith / endsWith { operator, text, formulae }
 *   timePeriod   { timePeriod, formulae }
 *   top10        { rank, percent?, bottom? }
 *   aboveAverage { aboveAverage? (false = below), equalAverage?, stdDev? }
 *   duplicateValues / uniqueValues {}
 *   colorScale   { cfvo: [{ type, value? }] x2|3, color: [{ argb }] }
 *   dataBar      { cfvo: [lo, hi], color, gradient?, border?, borderColor?, negativeFillColor?, ... }
 *   iconSet      { iconSet, reverse?, showValue?, cfvo: [{ type, value?, gte? }], icons? }
 * Formulas are stored without "=" and are relative to the top-left cell of the FIRST range in `ref`.
 *
 * Everything here is framework-free and never mutates its inputs.
 */
import type { CellStyle } from '../spreadsheet-types'

// ---------------------------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------------------------

export type ConditionalScalar = string | number | boolean | null

export interface ConditionalFormatHost {
  sheetId: string
  /** Resolved (calculated) value of a cell, 0-based row/col. Errors are strings such as "#N/A". */
  valueAt(row: number, col: number): ConditionalScalar
  /**
   * Evaluate a formula (no leading "=") as if entered at (row, col) on this sheet. Relative
   * references have already been shifted for that cell by the engine.
   */
  evaluate(formula: string, row: number, col: number): string | number | boolean
  /** Serial "today" for timePeriod rules (Excel 1900 date serial). */
  today: number
  /** Resolve a model colour ({ argb } / { theme, tint } / { indexed }) to CSS. Defaults to argb + Office theme. */
  cssColor?: (color: unknown, fallback?: string) => string
  /** Used-range extents (counts). Clamp whole-column/row rules and range statistics. */
  rowCount?: number
  colCount?: number
}

/** 0-based inclusive bounds. */
export interface CellBounds {
  top: number
  bottom: number
  left: number
  right: number
}

export interface ConditionalDataBar {
  /** Bar length as a fraction (0..1) of the cell width, measured from the axis (or the left edge). */
  fraction: number
  color: string
  /** The value is negative: the bar extends LEFT from `axis`. */
  negative?: boolean
  /** Axis position as a fraction (0..1) of the cell width; undefined when there is no axis. */
  axis?: number
  axisColor?: string
  gradient: boolean
  border?: string
  showValue: boolean
  /** Bars grow right-to-left (mirror the geometry). */
  rtl?: boolean
}

export interface ConditionalCellFormat {
  font?: { color?: string; bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean }
  fill?: string
  border?: { top?: string; right?: string; bottom?: string; left?: string }
  numFmt?: string
  dataBar?: ConditionalDataBar
  icon?: { set: string; index: number; showValue: boolean }
  colorScale?: string
}

export type CfvoType = 'min' | 'max' | 'autoMin' | 'autoMax' | 'num' | 'percent' | 'percentile' | 'formula'

export interface ConditionalValueObject {
  type: CfvoType | string
  value?: number | string | null
  /** Icon-set thresholds: false means ">" rather than ">=". */
  gte?: boolean
}

export interface ConditionalColor {
  argb?: string
  theme?: number
  tint?: number
  indexed?: number
  auto?: boolean
}

export interface ConditionalRule {
  type: string
  priority?: number
  stopIfTrue?: boolean
  style?: CellStyle
  operator?: string
  formulae?: string[]
  text?: string
  rank?: number
  percent?: boolean
  bottom?: boolean
  aboveAverage?: boolean
  equalAverage?: boolean
  stdDev?: number
  timePeriod?: string
  cfvo?: ConditionalValueObject[]
  color?: ConditionalColor | ConditionalColor[]
  iconSet?: string
  reverse?: boolean
  showValue?: boolean
  icons?: Array<{ iconSet: string; iconId: number }>
  gradient?: boolean
  border?: boolean
  borderColor?: ConditionalColor
  negativeFillColor?: ConditionalColor
  negativeBorderColor?: ConditionalColor
  axisColor?: ConditionalColor
  axisPosition?: string
  direction?: string
  minLength?: number
  maxLength?: number
  negativeBarColorSameAsPositive?: boolean
  negativeBarBorderColorSameAsPositive?: boolean
  [key: string]: unknown
}

export interface ConditionalFormattingBlock {
  ref: string
  rules: ConditionalRule[]
  [key: string]: unknown
}

export type ConditionalRuleKind =
  | 'expression'
  | 'cellIs'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors'
  | 'timePeriod'
  | 'top10'
  | 'aboveAverage'
  | 'duplicateValues'
  | 'uniqueValues'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'

export type CellIsOperator =
  | 'between'
  | 'notBetween'
  | 'equal'
  | 'notEqual'
  | 'greaterThan'
  | 'lessThan'
  | 'greaterThanOrEqual'
  | 'lessThanOrEqual'

export type TextRuleOperator = 'containsText' | 'notContainsText' | 'beginsWith' | 'endsWith'

export type TimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth'

export const TIME_PERIODS: readonly TimePeriod[] = [
  'today', 'yesterday', 'tomorrow', 'last7Days', 'thisWeek', 'lastWeek', 'nextWeek', 'thisMonth', 'lastMonth', 'nextMonth',
]

export const CELL_IS_OPERATORS: readonly CellIsOperator[] = [
  'greaterThan', 'greaterThanOrEqual', 'lessThan', 'lessThanOrEqual', 'equal', 'notEqual', 'between', 'notBetween',
]

export type IconSetName =
  | '3Arrows' | '3ArrowsGray' | '3Flags' | '3TrafficLights1' | '3TrafficLights2' | '3Signs' | '3Symbols' | '3Symbols2'
  | '3Stars' | '3Triangles'
  | '4Arrows' | '4ArrowsGray' | '4RedToBlack' | '4Rating' | '4TrafficLights'
  | '5Arrows' | '5ArrowsGray' | '5Rating' | '5Quarters' | '5Boxes'

export const ICON_SET_NAMES: readonly IconSetName[] = [
  '3Arrows', '3ArrowsGray', '3Triangles', '3Flags', '3TrafficLights1', '3TrafficLights2', '3Signs', '3Symbols', '3Symbols2', '3Stars',
  '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights',
  '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '5Boxes',
]

/** Number of icons in a set ("3Arrows" -> 3). Unknown names fall back to 3. */
export function iconSetSize(name: string | undefined): number {
  const size = Number(String(name || '').charAt(0))
  return size === 4 || size === 5 ? size : 3
}

/** Canonical icon-set name (ExcelJS reads a missing attribute as "3TrafficLights"). */
export function normalizeIconSetName(name: unknown): string {
  const text = typeof name === 'string' && name ? name : '3TrafficLights1'
  return text === '3TrafficLights' ? '3TrafficLights1' : text
}

// ---------------------------------------------------------------------------------------------
// Addresses and ranges
// ---------------------------------------------------------------------------------------------

export const MAX_SHEET_ROWS = 1_048_576
export const MAX_SHEET_COLUMNS = 16_384
/** Range statistics (top10, averages, duplicates, scales, bars, icons) stop scanning here. */
const MAX_STATISTIC_CELLS = 1_000_000
/** Per-cell application is capped at this many cells per rule. */
const MAX_APPLY_CELLS = 1_000_000

const COLUMN_LABELS: string[] = []

/** 0-based column index to its label ("A", "XFD"). */
export function columnLabel(col: number): string {
  let label = COLUMN_LABELS[col]
  if (label !== undefined) return label
  let value = col + 1
  label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  if (col >= 0 && col < MAX_SHEET_COLUMNS) COLUMN_LABELS[col] = label
  return label
}

/** Column label to a 0-based index, or -1. */
export function columnIndex(label: string): number {
  if (!label || label.length > 3) return -1
  let value = 0
  for (let index = 0; index < label.length; index += 1) {
    const code = label.charCodeAt(index) & ~32
    if (code < 65 || code > 90) return -1
    value = value * 26 + code - 64
  }
  return value >= 1 && value <= MAX_SHEET_COLUMNS ? value - 1 : -1
}

/** 0-based row/col to "A1". */
export function cellAddress(row: number, col: number): string {
  return columnLabel(col) + (row + 1)
}

const CELL_REF = /^\$?([A-Za-z]{1,3})\$?(\d{1,7})$/
const COLUMN_SPAN = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/
const ROW_SPAN = /^\$?(\d{1,7}):\$?(\d{1,7})$/

/** Parse "A1" (optionally with `$`) to 0-based coordinates. */
export function parseCellAddress(text: string): { row: number; col: number } | null {
  const match = CELL_REF.exec(text.trim())
  if (!match) return null
  const col = columnIndex(match[1])
  const row = Number(match[2]) - 1
  if (col < 0 || row < 0 || row >= MAX_SHEET_ROWS) return null
  return { row, col }
}

function stripSheetPrefix(text: string): string {
  const bang = text.lastIndexOf('!')
  return bang >= 0 ? text.slice(bang + 1) : text
}

/** Parse one range ("A1", "A1:C9", "A:C", "2:5"; `$` and a sheet prefix are ignored). */
export function parseRangeAddress(input: string): CellBounds | null {
  const text = stripSheetPrefix(input.trim())
  if (!text) return null
  const columns = COLUMN_SPAN.exec(text)
  if (columns) {
    const a = columnIndex(columns[1])
    const b = columnIndex(columns[2])
    if (a < 0 || b < 0) return null
    return { top: 0, bottom: MAX_SHEET_ROWS - 1, left: Math.min(a, b), right: Math.max(a, b) }
  }
  const rows = ROW_SPAN.exec(text)
  if (rows) {
    const a = Number(rows[1]) - 1
    const b = Number(rows[2]) - 1
    if (a < 0 || b < 0 || a >= MAX_SHEET_ROWS || b >= MAX_SHEET_ROWS) return null
    return { top: Math.min(a, b), bottom: Math.max(a, b), left: 0, right: MAX_SHEET_COLUMNS - 1 }
  }
  const parts = text.split(':')
  if (parts.length > 2) return null
  const start = parseCellAddress(parts[0])
  const end = parts.length === 2 ? parseCellAddress(parts[1]) : start
  if (!start || !end) return null
  return {
    top: Math.min(start.row, end.row),
    bottom: Math.max(start.row, end.row),
    left: Math.min(start.col, end.col),
    right: Math.max(start.col, end.col),
  }
}

/** Parse an sqref / user range list ("A1:B2 D4", "A1:B2, D4"). Invalid pieces are skipped. */
export function parseRangeList(ref: unknown): CellBounds[] {
  if (typeof ref !== 'string') return []
  const result: CellBounds[] = []
  for (const piece of ref.split(/[\s,;]+/)) {
    if (!piece) continue
    const bounds = parseRangeAddress(piece)
    if (bounds) result.push(bounds)
  }
  return result
}

/** Format bounds as an sqref piece ("B2", "B2:D9"). */
export function formatRangeAddress(bounds: CellBounds): string {
  const start = cellAddress(bounds.top, bounds.left)
  if (bounds.top === bounds.bottom && bounds.left === bounds.right) return start
  return `${start}:${cellAddress(bounds.bottom, bounds.right)}`
}

/** Format a range list as an Excel sqref (space separated). */
export function formatRangeList(list: readonly CellBounds[]): string {
  return list.map(formatRangeAddress).join(' ')
}

/**
 * Validate a user-entered "Apply to range" (comma/space separated). Returns the normalized
 * sqref or an error message.
 */
export function parseRangeInput(input: string): { ref: string; ranges: CellBounds[] } | { error: string } {
  const pieces = input.split(/[\s,;]+/).filter(Boolean)
  if (!pieces.length) return { error: 'Enter a range such as A1:D20.' }
  const ranges: CellBounds[] = []
  for (const piece of pieces) {
    const bounds = parseRangeAddress(piece)
    if (!bounds) return { error: `"${piece}" is not a valid range.` }
    ranges.push(bounds)
  }
  return { ref: formatRangeList(ranges), ranges }
}

function intersect(a: CellBounds, b: CellBounds): CellBounds | null {
  const top = Math.max(a.top, b.top)
  const bottom = Math.min(a.bottom, b.bottom)
  const left = Math.max(a.left, b.left)
  const right = Math.min(a.right, b.right)
  return top <= bottom && left <= right ? { top, bottom, left, right } : null
}

function boundsArea(bounds: CellBounds): number {
  return (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
}

/** Remove `cut` from `source`, returning up to four remaining rectangles (top, bottom, left, right bands). */
export function subtractBounds(source: CellBounds, cut: CellBounds): CellBounds[] {
  const overlap = intersect(source, cut)
  if (!overlap) return [source]
  const pieces: CellBounds[] = []
  if (source.top < overlap.top) pieces.push({ top: source.top, bottom: overlap.top - 1, left: source.left, right: source.right })
  if (overlap.left > source.left) pieces.push({ top: overlap.top, bottom: overlap.bottom, left: source.left, right: overlap.left - 1 })
  if (overlap.right < source.right) pieces.push({ top: overlap.top, bottom: overlap.bottom, left: overlap.right + 1, right: source.right })
  if (overlap.bottom < source.bottom) pieces.push({ top: overlap.bottom + 1, bottom: source.bottom, left: source.left, right: source.right })
  return pieces
}

// ---------------------------------------------------------------------------------------------
// Relative formula shifting (compiled once per rule, rendered per cell)
// ---------------------------------------------------------------------------------------------

interface CellRefSegment {
  kind: 0
  row: number
  col: number
  rowAbsolute: boolean
  colAbsolute: boolean
  lower: boolean
}

interface ColumnSpanSegment {
  kind: 1
  first: number
  firstAbsolute: boolean
  firstLower: boolean
  second: number
  secondAbsolute: boolean
  secondLower: boolean
}

interface RowSpanSegment {
  kind: 2
  first: number
  firstAbsolute: boolean
  second: number
  secondAbsolute: boolean
}

type FormulaSegment = string | CellRefSegment | ColumnSpanSegment | RowSpanSegment

export interface CompiledFormula {
  source: string
  /** True when at least one reference moves with the cell. */
  relative: boolean
  /** True when the result depends on the evaluating cell even without relative refs (ROW(), RAND()...). */
  positional: boolean
  segments: FormulaSegment[]
}

const CELL_TOKEN = /^(\$?)([A-Za-z]{1,3})(\$?)([1-9]\d*)/
const QUALIFIED_CELL_TOKEN = /^([A-Za-z_\\][A-Za-z0-9_.]*!)(\$?[A-Za-z]{1,3}\$?[1-9]\d*)/
const QUALIFIED_PREFIX = /^[A-Za-z_\\][A-Za-z0-9_.]*!/
const COLUMN_SPAN_TOKEN = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_.$])/
const ROW_SPAN_TOKEN = /^(\$?)([1-9]\d*):(\$?)([1-9]\d*)(?![A-Za-z0-9_.$])/
const POSITIONAL_FUNCTIONS = /\b(?:ROW|COLUMN|CELL|INDIRECT|OFFSET|RAND|RANDBETWEEN|RANDARRAY|ISFORMULA|FORMULATEXT)\s*\(/i

function boundaryBefore(source: string, position: number): boolean {
  if (position === 0) return true
  return !/[A-Za-z0-9_.$]/.test(source[position - 1])
}

function boundaryAfter(source: string, position: number): boolean {
  return !/[A-Za-z0-9_.]/.test(source[position] ?? '')
}

function nextNonSpace(source: string, start: number): string {
  let position = start
  while (/\s/.test(source[position] ?? '')) position += 1
  return source[position] ?? ''
}

function quotedStringEnd(source: string, start: number): number {
  let position = start + 1
  while (position < source.length) {
    if (source[position] !== '"') position += 1
    else if (source[position + 1] === '"') position += 2
    else return position + 1
  }
  return source.length
}

function quotedSheetEnd(source: string, start: number): number | null {
  let position = start + 1
  while (position < source.length) {
    if (source[position] !== "'") position += 1
    else if (source[position + 1] === "'") position += 2
    else return source[position + 1] === '!' ? position + 2 : null
  }
  return null
}

function cellSegment(text: string): CellRefSegment | null {
  const match = CELL_TOKEN.exec(text)
  if (!match || match[0].length !== text.length) return null
  const col = columnIndex(match[2])
  const row = Number(match[4]) - 1
  if (col < 0 || row < 0 || row >= MAX_SHEET_ROWS) return null
  return { kind: 0, row, col, rowAbsolute: match[3] === '$', colAbsolute: match[1] === '$', lower: match[2] === match[2].toLowerCase() }
}

function spanSegment(text: string): { segment: ColumnSpanSegment | RowSpanSegment; length: number } | null {
  const columns = COLUMN_SPAN_TOKEN.exec(text)
  if (columns) {
    const first = columnIndex(columns[2])
    const second = columnIndex(columns[4])
    if (first < 0 || second < 0) return null
    return {
      length: columns[0].length,
      segment: {
        kind: 1,
        first,
        firstAbsolute: columns[1] === '$',
        firstLower: columns[2] === columns[2].toLowerCase(),
        second,
        secondAbsolute: columns[3] === '$',
        secondLower: columns[4] === columns[4].toLowerCase(),
      },
    }
  }
  const rows = ROW_SPAN_TOKEN.exec(text)
  if (rows) {
    const first = Number(rows[2]) - 1
    const second = Number(rows[4]) - 1
    if (first >= MAX_SHEET_ROWS || second >= MAX_SHEET_ROWS) return null
    return {
      length: rows[0].length,
      segment: { kind: 2, first, firstAbsolute: rows[1] === '$', second, secondAbsolute: rows[3] === '$' },
    }
  }
  return null
}

const COMPILED_CACHE = new Map<string, CompiledFormula>()

/**
 * Split a formula into literal text and references so it can be re-rendered for any cell
 * offset cheaply. Mirrors `shiftFormulaReferences` in formulas.ts (strings, structured
 * references, quoted/unquoted sheet prefixes, whole rows/columns, function names).
 */
export function compileRelativeFormula(formula: string): CompiledFormula {
  const cached = COMPILED_CACHE.get(formula)
  if (cached) return cached
  const segments: FormulaSegment[] = []
  let literal = ''
  let relative = false
  const push = (segment: CellRefSegment | ColumnSpanSegment | RowSpanSegment) => {
    if (literal) segments.push(literal)
    literal = ''
    segments.push(segment)
    if (segment.kind === 0) relative ||= !segment.rowAbsolute || !segment.colAbsolute
    else relative ||= !segment.firstAbsolute || !segment.secondAbsolute
  }
  let position = 0
  while (position < formula.length) {
    const char = formula[position]
    if (char === '"') {
      const end = quotedStringEnd(formula, position)
      literal += formula.slice(position, end)
      position = end
      continue
    }
    if (char === '[') {
      const end = formula.indexOf(']', position + 1)
      if (end >= 0) {
        literal += formula.slice(position, end + 1)
        position = end + 1
        continue
      }
    }
    if (char === "'" && boundaryBefore(formula, position)) {
      const prefixEnd = quotedSheetEnd(formula, position)
      if (prefixEnd !== null) {
        const rest = formula.slice(prefixEnd)
        const match = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*/.exec(rest)
        if (match && boundaryAfter(formula, prefixEnd + match[0].length)) {
          const segment = cellSegment(match[0])
          literal += formula.slice(position, prefixEnd)
          if (segment) push(segment)
          else literal += match[0]
          position = prefixEnd + match[0].length
          continue
        }
        const span = spanSegment(rest)
        if (span) {
          literal += formula.slice(position, prefixEnd)
          push(span.segment)
          position = prefixEnd + span.length
          continue
        }
      }
    }
    if (boundaryBefore(formula, position)) {
      const rest = formula.slice(position)
      const qualified = QUALIFIED_CELL_TOKEN.exec(rest)
      if (qualified && boundaryAfter(formula, position + qualified[0].length)) {
        const segment = cellSegment(qualified[2])
        literal += qualified[1]
        if (segment) push(segment)
        else literal += qualified[2]
        position += qualified[0].length
        continue
      }
      const prefix = QUALIFIED_PREFIX.exec(rest)
      if (prefix) {
        const span = spanSegment(rest.slice(prefix[0].length))
        if (span) {
          literal += prefix[0]
          push(span.segment)
          position += prefix[0].length + span.length
          continue
        }
      }
      const reference = /^\$?[A-Za-z]{1,3}\$?[1-9]\d*/.exec(rest)
      if (reference && boundaryAfter(formula, position + reference[0].length) && nextNonSpace(formula, position + reference[0].length) !== '(') {
        const segment = cellSegment(reference[0])
        if (segment) push(segment)
        else literal += reference[0]
        position += reference[0].length
        continue
      }
      const span = spanSegment(rest)
      if (span) {
        push(span.segment)
        position += span.length
        continue
      }
    }
    literal += char
    position += 1
  }
  if (literal) segments.push(literal)
  const compiled: CompiledFormula = { source: formula, relative, positional: POSITIONAL_FUNCTIONS.test(formula), segments }
  if (COMPILED_CACHE.size > 2_000) COMPILED_CACHE.clear()
  COMPILED_CACHE.set(formula, compiled)
  return compiled
}

function wrap(value: number, size: number): number {
  const result = value % size
  return result < 0 ? result + size : result
}

function labelCase(label: string, lower: boolean) {
  return lower ? label.toLowerCase() : label
}

/**
 * Render a compiled formula shifted by (rowDelta, colDelta). Relative references wrap around
 * the sheet edges like Excel's conditional-format and defined-name references.
 */
export function renderRelativeFormula(compiled: CompiledFormula, rowDelta: number, colDelta: number): string {
  if (!compiled.relative || (rowDelta === 0 && colDelta === 0)) return compiled.source
  let output = ''
  for (const segment of compiled.segments) {
    if (typeof segment === 'string') {
      output += segment
    } else if (segment.kind === 0) {
      const col = segment.colAbsolute ? segment.col : wrap(segment.col + colDelta, MAX_SHEET_COLUMNS)
      const row = segment.rowAbsolute ? segment.row : wrap(segment.row + rowDelta, MAX_SHEET_ROWS)
      output += `${segment.colAbsolute ? '$' : ''}${labelCase(columnLabel(col), segment.lower)}${segment.rowAbsolute ? '$' : ''}${row + 1}`
    } else if (segment.kind === 1) {
      const first = segment.firstAbsolute ? segment.first : wrap(segment.first + colDelta, MAX_SHEET_COLUMNS)
      const second = segment.secondAbsolute ? segment.second : wrap(segment.second + colDelta, MAX_SHEET_COLUMNS)
      output += `${segment.firstAbsolute ? '$' : ''}${labelCase(columnLabel(first), segment.firstLower)}:${segment.secondAbsolute ? '$' : ''}${labelCase(columnLabel(second), segment.secondLower)}`
    } else {
      const first = segment.firstAbsolute ? segment.first : wrap(segment.first + rowDelta, MAX_SHEET_ROWS)
      const second = segment.secondAbsolute ? segment.second : wrap(segment.second + rowDelta, MAX_SHEET_ROWS)
      output += `${segment.firstAbsolute ? '$' : ''}${first + 1}:${segment.secondAbsolute ? '$' : ''}${second + 1}`
    }
  }
  return output
}

/** Shift the relative references of a conditional-format formula by a row/column offset. */
export function shiftConditionalFormula(formula: string, rowDelta: number, colDelta: number): string {
  return renderRelativeFormula(compileRelativeFormula(formula), rowDelta, colDelta)
}

// ---------------------------------------------------------------------------------------------
// Scalars
// ---------------------------------------------------------------------------------------------

const ERROR_VALUES = new Set([
  '#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!',
  '#CIRC!', '#PARSE!', '#FIELD!', '#BLOCKED!', '#CONNECT!', '#BUSY!', '#UNKNOWN!',
])

export function isConditionalError(value: unknown): boolean {
  return typeof value === 'string' && value.charCodeAt(0) === 35 && ERROR_VALUES.has(value)
}

function typeRank(value: string | number | boolean): number {
  return typeof value === 'number' ? 1 : typeof value === 'string' ? 2 : 3
}

/** Excel comparison ordering (numbers < text < logicals; text case-insensitive). Null on error. */
function compareScalars(left: ConditionalScalar, right: ConditionalScalar): number | null {
  if (isConditionalError(left) || isConditionalError(right)) return null
  let a = left
  let b = right
  if (a === null && b === null) return 0
  if (a === null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0
  if (b === null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0
  const ranks = typeRank(a) - typeRank(b as string | number | boolean)
  if (ranks !== 0) return ranks
  if (typeof a === 'number') {
    const y = b as number
    if (a === y) return 0
    const scale = Math.max(Math.abs(a), Math.abs(y))
    if (Math.abs(a - y) <= scale * 4e-15) return 0
    return a < y ? -1 : 1
  }
  if (typeof a === 'string') {
    const x = a.toLowerCase()
    const y = (b as string).toLowerCase()
    return x < y ? -1 : x > y ? 1 : 0
  }
  return Number(a) - Number(b)
}

function numberText(value: number): string {
  if (!Number.isFinite(value)) return String(value)
  const precise = Number(value.toPrecision(15))
  return String(precise).replace('e', 'E')
}

/** Cell value as SEARCH/LEFT/RIGHT see it, or null for errors. */
function cellText(value: ConditionalScalar): string | null {
  if (value === null) return ''
  if (typeof value === 'number') return numberText(value)
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  return isConditionalError(value) ? null : value
}

function truthy(value: unknown): boolean {
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') return value !== 0 && Number.isFinite(value)
  return false
}

function numericValue(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'string' && value.trim() && !isConditionalError(value)) {
    const number = Number(value)
    return Number.isFinite(number) ? number : null
  }
  return null
}

function stripEquals(formula: string): string {
  const text = formula.trim()
  return text.startsWith('=') ? text.slice(1).trim() : text
}

type Literal = { value: ConditionalScalar } | null

function formulaLiteral(formula: string): Literal {
  const text = formula.trim()
  if (/^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/.test(text)) return { value: Number(text) }
  if (/^"(?:[^"]|"")*"$/.test(text)) return { value: text.slice(1, -1).replace(/""/g, '"') }
  if (/^true$/i.test(text)) return { value: true }
  if (/^false$/i.test(text)) return { value: false }
  return null
}

// ---------------------------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------------------------

const OFFICE_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']

function tintChannel(channel: number, tint: number): number {
  const adjusted = tint < 0 ? channel * (1 + tint) : channel + (255 - channel) * tint
  return Math.round(Math.min(255, Math.max(0, adjusted)))
}

/** Default colour resolver: argb/rgb strings, `{ argb }`, and Office-theme `{ theme, tint }`. */
export function defaultConditionalCssColor(color: unknown, fallback = ''): string {
  let hex = ''
  let tint = 0
  if (typeof color === 'string') hex = color
  else if (color && typeof color === 'object') {
    const record = color as ConditionalColor & { rgb?: string }
    hex = record.argb || record.rgb || ''
    tint = Number(record.tint) || 0
    if (!hex && Number.isInteger(record.theme)) hex = OFFICE_THEME[Number(record.theme)] || ''
  }
  hex = hex.replace(/^#/, '')
  if (hex.length === 8) hex = hex.slice(2)
  if (!/^[0-9a-f]{6}$/i.test(hex)) return fallback
  if (!tint) return `#${hex.toUpperCase()}`
  const channels = [0, 2, 4].map((offset) => tintChannel(Number.parseInt(hex.slice(offset, offset + 2), 16), Math.max(-1, Math.min(1, tint))))
  return `#${channels.map((channel) => channel.toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

/** Parse "#RGB", "#RRGGBB", "#AARRGGBB"-less CSS or "rgb(r, g, b)" into channels. */
export function parseCssColor(css: string): [number, number, number] | null {
  const text = css.trim()
  let match = /^#?([0-9a-f]{6})$/i.exec(text)
  if (match) {
    const value = Number.parseInt(match[1], 16)
    return [(value >> 16) & 255, (value >> 8) & 255, value & 255]
  }
  match = /^#([0-9a-f]{3})$/i.exec(text)
  if (match) {
    const [r, g, b] = match[1].split('').map((digit) => Number.parseInt(digit + digit, 16))
    return [r, g, b]
  }
  const rgb = /^rgba?\(\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)\s*,\s*(\d+(?:\.\d+)?)/i.exec(text)
  if (rgb) return [Number(rgb[1]), Number(rgb[2]), Number(rgb[3])]
  return null
}

function rgbCss(channels: readonly number[]): string {
  return `#${channels.map((channel) => Math.round(Math.min(255, Math.max(0, channel))).toString(16).padStart(2, '0')).join('').toUpperCase()}`
}

/** Mix two CSS colours (amount 0 = a, 1 = b). */
export function mixCssColors(a: string, b: string, amount: number): string {
  const left = parseCssColor(a)
  const right = parseCssColor(b)
  if (!left || !right) return amount < 0.5 ? a : b
  return rgbCss(left.map((channel, index) => channel + (right[index] - channel) * amount))
}

/** "#RRGGBB" / "RRGGBB" / "AARRGGBB" -> { argb: 'FFRRGGBB' }. */
export function toArgbColor(color: string): ConditionalColor {
  const hex = color.trim().replace(/^#/, '').toUpperCase()
  if (/^[0-9A-F]{8}$/.test(hex)) return { argb: hex }
  if (/^[0-9A-F]{6}$/.test(hex)) return { argb: `FF${hex}` }
  if (/^[0-9A-F]{3}$/.test(hex)) return { argb: `FF${hex.split('').map((digit) => digit + digit).join('')}` }
  return { argb: 'FF000000' }
}

// ---------------------------------------------------------------------------------------------
// Rule classification
// ---------------------------------------------------------------------------------------------

const TEXT_OPERATOR_KINDS: Record<string, ConditionalRuleKind> = {
  containsText: 'containsText',
  notContains: 'notContainsText',
  notContainsText: 'notContainsText',
  beginsWith: 'beginsWith',
  endsWith: 'endsWith',
  containsBlanks: 'containsBlanks',
  notContainsBlanks: 'notContainsBlanks',
  containsErrors: 'containsErrors',
  notContainsErrors: 'notContainsErrors',
}

const DIRECT_KINDS = new Set<string>([
  'expression', 'cellIs', 'notContainsText', 'beginsWith', 'endsWith', 'containsBlanks', 'notContainsBlanks',
  'containsErrors', 'notContainsErrors', 'timePeriod', 'top10', 'aboveAverage', 'duplicateValues', 'uniqueValues',
  'colorScale', 'dataBar', 'iconSet',
])

/** Canonical kind of a rule (ExcelJS stores blank/error rules as type "containsText" + operator). */
export function conditionalRuleKind(rule: unknown): ConditionalRuleKind | null {
  if (!rule || typeof rule !== 'object') return null
  const record = rule as ConditionalRule
  if (record.type === 'containsText') return TEXT_OPERATOR_KINDS[String(record.operator || 'containsText')] || 'containsText'
  return DIRECT_KINDS.has(record.type) ? (record.type as ConditionalRuleKind) : null
}

/** Rule kinds that apply a dxf style (the others draw scales, bars or icons). */
export function ruleUsesStyle(kind: ConditionalRuleKind | null): boolean {
  return kind !== null && kind !== 'colorScale' && kind !== 'dataBar' && kind !== 'iconSet'
}

// ---------------------------------------------------------------------------------------------
// Style fragments (dxf -> css)
// ---------------------------------------------------------------------------------------------

interface StyleFragment {
  font?: NonNullable<ConditionalCellFormat['font']>
  fill?: string
  border?: NonNullable<ConditionalCellFormat['border']>
  numFmt?: string
}

const BORDER_WIDTHS: Record<string, string> = {
  thin: '1px solid',
  hair: '1px dotted',
  dotted: '1px dotted',
  dashed: '1px dashed',
  dashDot: '1px dashed',
  dashDotDot: '1px dotted',
  medium: '2px solid',
  mediumDashed: '2px dashed',
  mediumDashDot: '2px dashed',
  mediumDashDotDot: '2px dotted',
  slantDashDot: '2px dashed',
  thick: '3px solid',
  double: '3px double',
}

function styleFragment(style: unknown, color: (value: unknown, fallback?: string) => string): StyleFragment | null {
  if (!style || typeof style !== 'object') return null
  const record = style as CellStyle & { numFmt?: unknown }
  const fragment: StyleFragment = {}
  const font = record.font
  if (font && typeof font === 'object') {
    const result: NonNullable<ConditionalCellFormat['font']> = {}
    const fontColor = font.color !== undefined ? color(font.color) : ''
    if (fontColor) result.color = fontColor
    if (typeof font.bold === 'boolean') result.bold = font.bold
    if (typeof font.italic === 'boolean') result.italic = font.italic
    if (font.underline !== undefined) result.underline = Boolean(font.underline) && font.underline !== 'none'
    if (typeof font.strike === 'boolean') result.strike = font.strike
    if (Object.keys(result).length) fragment.font = result
  }
  const fill = record.fill
  if (fill && typeof fill === 'object') {
    let css = ''
    if (fill.type === 'gradient' && Array.isArray(fill.stops) && fill.stops.length) {
      css = color(fill.stops[0]?.color)
    } else if (fill.pattern !== 'none') {
      // Differential (dxf) solid fills carry their colour in bgColor; fall back to fgColor.
      css = color(fill.bgColor) || color(fill.fgColor) || color(fill.color)
    }
    if (css) fragment.fill = css
  }
  const border = record.border
  if (border && typeof border === 'object') {
    const result: NonNullable<ConditionalCellFormat['border']> = {}
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const item = border[side]
      if (!item || typeof item !== 'object' || !item.style || item.style === 'none') continue
      result[side] = `${BORDER_WIDTHS[item.style] || '1px solid'} ${color(item.color, '#000000') || '#000000'}`
    }
    if (Object.keys(result).length) fragment.border = result
  }
  const numFmt = typeof record.numFmt === 'string'
    ? record.numFmt
    : record.numFmt && typeof record.numFmt === 'object' && typeof (record.numFmt as { formatCode?: unknown }).formatCode === 'string'
      ? String((record.numFmt as { formatCode: string }).formatCode)
      : ''
  if (numFmt) fragment.numFmt = numFmt
  return fragment.font || fragment.fill || fragment.border || fragment.numFmt ? fragment : null
}

function mergeFragment(target: ConditionalCellFormat, fragment: StyleFragment) {
  if (fragment.font) {
    const font = target.font ?? (target.font = {})
    const source = fragment.font
    if (source.color !== undefined && font.color === undefined) font.color = source.color
    if (source.bold !== undefined && font.bold === undefined) font.bold = source.bold
    if (source.italic !== undefined && font.italic === undefined) font.italic = source.italic
    if (source.underline !== undefined && font.underline === undefined) font.underline = source.underline
    if (source.strike !== undefined && font.strike === undefined) font.strike = source.strike
  }
  if (fragment.fill !== undefined && target.fill === undefined) target.fill = fragment.fill
  if (fragment.border) {
    const border = target.border ?? (target.border = {})
    const source = fragment.border
    if (source.top !== undefined && border.top === undefined) border.top = source.top
    if (source.right !== undefined && border.right === undefined) border.right = source.right
    if (source.bottom !== undefined && border.bottom === undefined) border.bottom = source.bottom
    if (source.left !== undefined && border.left === undefined) border.left = source.left
  }
  if (fragment.numFmt !== undefined && target.numFmt === undefined) target.numFmt = fragment.numFmt
}

// ---------------------------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------------------------

interface PreparedRule {
  rule: ConditionalRule
  kind: ConditionalRuleKind
  ranges: CellBounds[]
  rangeKey: string
  anchorRow: number
  anchorCol: number
  priority: number
  order: number
}

interface RangeStatistics {
  values: number[]
  count: number
  min: number
  max: number
  sum: number
  mean: number
  deviation: number
  sorted?: Float64Array
}

function flattenRules(input: unknown[] | undefined): PreparedRule[] {
  const prepared: PreparedRule[] = []
  if (!Array.isArray(input)) return prepared
  let order = 0
  for (const block of input) {
    if (!block || typeof block !== 'object') continue
    const record = block as ConditionalFormattingBlock
    const ranges = parseRangeList(record.ref)
    if (!ranges.length || !Array.isArray(record.rules)) continue
    for (const rule of record.rules) {
      order += 1
      const kind = conditionalRuleKind(rule)
      if (!kind) continue
      const priority = Number((rule as ConditionalRule).priority)
      prepared.push({
        rule: rule as ConditionalRule,
        kind,
        ranges,
        rangeKey: formatRangeList(ranges),
        anchorRow: ranges[0].top,
        anchorCol: ranges[0].left,
        priority: Number.isFinite(priority) && priority > 0 ? priority : Number.POSITIVE_INFINITY,
        order,
      })
    }
  }
  prepared.sort((a, b) => a.priority - b.priority || a.order - b.order)
  return prepared
}

function rangesOverlap(ranges: readonly CellBounds[]): boolean {
  for (let a = 0; a < ranges.length; a += 1) {
    for (let b = a + 1; b < ranges.length; b += 1) if (intersect(ranges[a], ranges[b])) return true
  }
  return false
}

/** Visit each cell of the clipped ranges once (row-major per range), up to `limit` cells. */
function forEachCell(ranges: readonly CellBounds[], limit: number, visit: (row: number, col: number) => void) {
  const dedupe = ranges.length > 1 && rangesOverlap(ranges) ? new Set<number>() : null
  let remaining = limit
  for (const range of ranges) {
    for (let row = range.top; row <= range.bottom; row += 1) {
      for (let col = range.left; col <= range.right; col += 1) {
        if (dedupe) {
          const key = row * MAX_SHEET_COLUMNS + col
          if (dedupe.has(key)) continue
          dedupe.add(key)
        }
        if (remaining-- <= 0) return
        visit(row, col)
      }
    }
  }
}

function percentileOf(sorted: Float64Array, fraction: number): number {
  if (!sorted.length) return Number.NaN
  const clamped = Math.min(1, Math.max(0, fraction))
  const rank = clamped * (sorted.length - 1)
  const low = Math.floor(rank)
  const high = Math.min(sorted.length - 1, low + 1)
  return sorted[low] + (sorted[high] - sorted[low]) * (rank - low)
}

function serialWeekday(serial: number): number {
  return (((Math.floor(serial) - 1) % 7) + 7) % 7 + 1
}

function serialMonthIndex(serial: number): number {
  const day = Math.floor(serial)
  if (day === 60) return 1900 * 12 + 1
  const date = new Date(Date.UTC(1899, 11, 31) + (day < 60 ? day : day - 1) * 86_400_000)
  return date.getUTCFullYear() * 12 + date.getUTCMonth()
}

function matchesTimePeriod(period: string, value: number, today: number): boolean {
  const day = Math.floor(value)
  const now = Math.floor(today)
  switch (period) {
    case 'today': return day === now
    case 'yesterday': return day === now - 1
    case 'tomorrow': return day === now + 1
    case 'last7Days': return day <= now && day >= now - 6
    case 'thisWeek': {
      const weekday = serialWeekday(now)
      return day >= now - (weekday - 1) && day <= now + (7 - weekday)
    }
    case 'lastWeek': {
      const weekday = serialWeekday(now)
      return day >= now - weekday - 6 && day <= now - weekday
    }
    case 'nextWeek': {
      const weekday = serialWeekday(now)
      return day >= now + 8 - weekday && day <= now + 14 - weekday
    }
    case 'thisMonth': return serialMonthIndex(day) === serialMonthIndex(now)
    case 'lastMonth': return serialMonthIndex(day) === serialMonthIndex(now) - 1
    case 'nextMonth': return serialMonthIndex(day) === serialMonthIndex(now) + 1
    default: return false
  }
}

function wildcardMatcher(pattern: string): (text: string) => boolean {
  if (!/[*?~]/.test(pattern)) {
    const needle = pattern.toLowerCase()
    return (text) => text.toLowerCase().includes(needle)
  }
  let source = ''
  for (let index = 0; index < pattern.length; index += 1) {
    const char = pattern[index]
    if (char === '~' && index + 1 < pattern.length) {
      index += 1
      source += pattern[index].replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    } else if (char === '*') source += '[\\s\\S]*'
    else if (char === '?') source += '[\\s\\S]'
    else source += char.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }
  const expression = new RegExp(source, 'i')
  return (text) => expression.test(text)
}

class Computation {
  readonly color: (value: unknown, fallback?: string) => string
  readonly used: CellBounds | null
  private readonly statistics = new Map<string, RangeStatistics>()
  private readonly duplicates = new Map<string, Map<string, number>>()

  constructor(readonly host: ConditionalFormatHost, readonly visible: CellBounds | null) {
    const resolver = host.cssColor
    this.color = resolver
      ? (value, fallback = '') => {
        try {
          return resolver(value, fallback) || fallback
        } catch {
          return fallback
        }
      }
      : defaultConditionalCssColor
    const rows = Math.floor(Number(host.rowCount))
    const cols = Math.floor(Number(host.colCount))
    this.used = rows > 0 && cols > 0
      ? { top: 0, bottom: Math.min(rows, MAX_SHEET_ROWS) - 1, left: 0, right: Math.min(cols, MAX_SHEET_COLUMNS) - 1 }
      : null
  }

  value(row: number, col: number): ConditionalScalar {
    try {
      const value = this.host.valueAt(row, col)
      return value === undefined ? null : value
    } catch {
      return '#VALUE!'
    }
  }

  evaluate(formula: string, row: number, col: number): ConditionalScalar {
    try {
      const value = this.host.evaluate(formula, row, col)
      return value === undefined ? null : value
    } catch {
      return '#VALUE!'
    }
  }

  /** Ranges clipped to the used range (statistics). */
  statisticRanges(ranges: readonly CellBounds[]): CellBounds[] {
    if (!this.used) return ranges.slice()
    const result: CellBounds[] = []
    for (const range of ranges) {
      const clipped = intersect(range, this.used)
      if (clipped) result.push(clipped)
    }
    return result
  }

  /** Ranges clipped to the visible window (or the used range) for per-cell application. */
  applyRanges(ranges: readonly CellBounds[]): CellBounds[] {
    const clip = this.visible ?? this.used
    if (!clip) return ranges.slice()
    const result: CellBounds[] = []
    for (const range of ranges) {
      const clipped = intersect(range, clip)
      if (clipped) result.push(clipped)
    }
    return result
  }

  stats(entry: PreparedRule): RangeStatistics {
    const cached = this.statistics.get(entry.rangeKey)
    if (cached) return cached
    const values: number[] = []
    let min = Number.POSITIVE_INFINITY
    let max = Number.NEGATIVE_INFINITY
    let sum = 0
    forEachCell(this.statisticRanges(entry.ranges), MAX_STATISTIC_CELLS, (row, col) => {
      const value = this.value(row, col)
      if (typeof value !== 'number' || !Number.isFinite(value)) return
      values.push(value)
      if (value < min) min = value
      if (value > max) max = value
      sum += value
    })
    const count = values.length
    const mean = count ? sum / count : Number.NaN
    let squares = 0
    for (const value of values) squares += (value - mean) * (value - mean)
    const result: RangeStatistics = {
      values,
      count,
      min: count ? min : Number.NaN,
      max: count ? max : Number.NaN,
      sum,
      mean,
      deviation: count ? Math.sqrt(squares / count) : Number.NaN,
    }
    this.statistics.set(entry.rangeKey, result)
    return result
  }

  sorted(stats: RangeStatistics): Float64Array {
    if (!stats.sorted) stats.sorted = Float64Array.from(stats.values).sort()
    return stats.sorted
  }

  duplicateCounts(entry: PreparedRule): Map<string, number> {
    const cached = this.duplicates.get(entry.rangeKey)
    if (cached) return cached
    const counts = new Map<string, number>()
    forEachCell(this.statisticRanges(entry.ranges), MAX_STATISTIC_CELLS, (row, col) => {
      const key = duplicateKey(this.value(row, col))
      if (key !== null) counts.set(key, (counts.get(key) ?? 0) + 1)
    })
    this.duplicates.set(entry.rangeKey, counts)
    return counts
  }
}

function duplicateKey(value: ConditionalScalar): string | null {
  if (value === null || value === '') return null
  if (typeof value === 'number') return Number.isFinite(value) ? `n${value}` : null
  if (typeof value === 'boolean') return value ? 'bT' : 'bF'
  if (isConditionalError(value)) return `e${value}`
  return `s${value.toLowerCase()}`
}

type Matcher = (row: number, col: number) => boolean

/** Operand of a cellIs rule: a literal, a per-rule constant, or a per-cell relative formula. */
function operandResolver(formula: unknown, entry: PreparedRule, computation: Computation): (row: number, col: number) => ConditionalScalar {
  if (typeof formula === 'number') return () => formula
  if (typeof formula !== 'string' || !formula.trim()) return () => '#N/A'
  const text = stripEquals(formula)
  const literal = formulaLiteral(text)
  if (literal) return () => literal.value
  return formulaResolver(text, entry, computation)
}

function formulaResolver(formula: string, entry: PreparedRule, computation: Computation): (row: number, col: number) => ConditionalScalar {
  const compiled = compileRelativeFormula(formula)
  if (!compiled.relative && !compiled.positional) {
    let computed = false
    let constant: ConditionalScalar = null
    return () => {
      if (!computed) {
        constant = computation.evaluate(formula, entry.anchorRow, entry.anchorCol)
        computed = true
      }
      return constant
    }
  }
  const cache = compiled.positional ? null : new Map<string, ConditionalScalar>()
  return (row, col) => {
    const shifted = renderRelativeFormula(compiled, row - entry.anchorRow, col - entry.anchorCol)
    if (cache) {
      const hit = cache.get(shifted)
      if (hit !== undefined || cache.has(shifted)) return hit as ConditionalScalar
      const value = computation.evaluate(shifted, row, col)
      cache.set(shifted, value)
      return value
    }
    return computation.evaluate(shifted, row, col)
  }
}

function cellIsMatcher(entry: PreparedRule, computation: Computation): Matcher | null {
  const formulae = Array.isArray(entry.rule.formulae) ? entry.rule.formulae : []
  if (!formulae.length) return null
  const operator = String(entry.rule.operator || 'equal')
  const first = operandResolver(formulae[0], entry, computation)
  const second = formulae.length > 1 ? operandResolver(formulae[1], entry, computation) : null
  const compare = (row: number, col: number, operand: (row: number, col: number) => ConditionalScalar) => {
    return compareScalars(computation.value(row, col), operand(row, col))
  }
  switch (operator) {
    case 'between':
    case 'notBetween': {
      if (!second) return null
      const negate = operator === 'notBetween'
      return (row, col) => {
        const low = compare(row, col, first)
        const high = compare(row, col, second)
        if (low === null || high === null) return false
        const inside = (low >= 0 && high <= 0) || (low <= 0 && high >= 0)
        return negate ? !inside : inside
      }
    }
    case 'equal': return (row, col) => compare(row, col, first) === 0
    case 'notEqual': return (row, col) => { const result = compare(row, col, first); return result !== null && result !== 0 }
    case 'greaterThan': return (row, col) => { const result = compare(row, col, first); return result !== null && result > 0 }
    case 'lessThan': return (row, col) => { const result = compare(row, col, first); return result !== null && result < 0 }
    case 'greaterThanOrEqual': return (row, col) => { const result = compare(row, col, first); return result !== null && result >= 0 }
    case 'lessThanOrEqual': return (row, col) => { const result = compare(row, col, first); return result !== null && result <= 0 }
    default: return null
  }
}

function expressionMatcher(entry: PreparedRule, computation: Computation): Matcher | null {
  const formula = Array.isArray(entry.rule.formulae) ? entry.rule.formulae[0] : undefined
  if (typeof formula !== 'string' || !stripEquals(formula)) return null
  const resolve = formulaResolver(stripEquals(formula), entry, computation)
  return (row, col) => truthy(resolve(row, col))
}

function textMatcher(entry: PreparedRule, computation: Computation): Matcher | null {
  const text = entry.rule.text
  if (typeof text !== 'string') return expressionMatcher(entry, computation)
  switch (entry.kind) {
    case 'containsText':
    case 'notContainsText': {
      const matches = wildcardMatcher(text)
      const negate = entry.kind === 'notContainsText'
      return (row, col) => {
        const value = cellText(computation.value(row, col))
        if (value === null) return negate
        return negate ? !matches(value) : matches(value)
      }
    }
    case 'beginsWith':
    case 'endsWith': {
      const needle = text.toLowerCase()
      const begins = entry.kind === 'beginsWith'
      return (row, col) => {
        const value = cellText(computation.value(row, col))
        if (value === null) return false
        const lower = value.toLowerCase()
        return begins ? lower.startsWith(needle) : lower.endsWith(needle)
      }
    }
    default: return null
  }
}

function isBlankValue(value: ConditionalScalar): boolean {
  return value === null || (typeof value === 'string' && /^ *$/.test(value))
}

function styleMatcher(entry: PreparedRule, computation: Computation): Matcher | null {
  const rule = entry.rule
  switch (entry.kind) {
    case 'expression': return expressionMatcher(entry, computation)
    case 'cellIs': return cellIsMatcher(entry, computation)
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return textMatcher(entry, computation)
    case 'containsBlanks': return (row, col) => { const value = computation.value(row, col); return !isConditionalError(value) && isBlankValue(value) }
    case 'notContainsBlanks': return (row, col) => { const value = computation.value(row, col); return !isConditionalError(value) && !isBlankValue(value) }
    case 'containsErrors': return (row, col) => isConditionalError(computation.value(row, col))
    case 'notContainsErrors': return (row, col) => !isConditionalError(computation.value(row, col))
    case 'timePeriod': {
      const period = String(rule.timePeriod || '')
      const today = Number(computation.host.today)
      if (!Number.isFinite(today)) return null
      return (row, col) => {
        const value = computation.value(row, col)
        return typeof value === 'number' && Number.isFinite(value) && matchesTimePeriod(period, value, today)
      }
    }
    case 'top10': {
      const stats = computation.stats(entry)
      if (!stats.count) return null
      const rank = Math.floor(Number(rule.rank ?? 10))
      if (!Number.isFinite(rank) || rank < 1) return null
      const count = rule.percent
        ? Math.max(1, Math.floor((stats.count * Math.min(rank, 100)) / 100))
        : Math.min(rank, stats.count)
      const sorted = computation.sorted(stats)
      if (rule.bottom) {
        const threshold = sorted[count - 1]
        return (row, col) => { const value = computation.value(row, col); return typeof value === 'number' && value <= threshold }
      }
      const threshold = sorted[sorted.length - count]
      return (row, col) => { const value = computation.value(row, col); return typeof value === 'number' && value >= threshold }
    }
    case 'aboveAverage': {
      const stats = computation.stats(entry)
      if (!stats.count) return null
      const above = rule.aboveAverage !== false
      const equal = rule.equalAverage === true
      const deviations = Math.max(0, Math.floor(Number(rule.stdDev) || 0))
      const threshold = above ? stats.mean + deviations * stats.deviation : stats.mean - deviations * stats.deviation
      return (row, col) => {
        const value = computation.value(row, col)
        if (typeof value !== 'number' || !Number.isFinite(value)) return false
        if (above) return equal ? value >= threshold : value > threshold
        return equal ? value <= threshold : value < threshold
      }
    }
    case 'duplicateValues':
    case 'uniqueValues': {
      const counts = computation.duplicateCounts(entry)
      const duplicate = entry.kind === 'duplicateValues'
      return (row, col) => {
        const key = duplicateKey(computation.value(row, col))
        if (key === null) return false
        const count = counts.get(key) ?? 0
        return duplicate ? count > 1 : count === 1
      }
    }
    default: return null
  }
}

/** Resolve a cfvo threshold to a number (null when it cannot be resolved). */
function cfvoNumber(cfvo: ConditionalValueObject | undefined, stats: RangeStatistics, entry: PreparedRule, computation: Computation, fallback: 'min' | 'max'): number | null {
  const type = cfvo && typeof cfvo === 'object' ? String(cfvo.type || fallback) : fallback
  const raw = cfvo && typeof cfvo === 'object' ? cfvo.value : undefined
  const scalar = (): number | null => {
    if (typeof raw === 'number') return Number.isFinite(raw) ? raw : null
    if (typeof raw !== 'string' || !raw.trim()) return null
    const direct = numericValue(raw)
    if (direct !== null) return direct
    return numericValue(computation.evaluate(stripEquals(raw), entry.anchorRow, entry.anchorCol))
  }
  switch (type) {
    case 'min': return stats.min
    case 'max': return stats.max
    case 'autoMin': return Math.min(0, stats.min)
    case 'autoMax': return Math.max(0, stats.max)
    case 'num':
    case 'formula':
      return scalar()
    case 'percent': {
      const percent = scalar()
      return percent === null ? null : stats.min + ((stats.max - stats.min) * percent) / 100
    }
    case 'percentile': {
      const percentile = scalar()
      return percentile === null ? null : percentileOf(computation.sorted(stats), percentile / 100)
    }
    default: return fallback === 'min' ? stats.min : stats.max
  }
}

type VisualEffect = (row: number, col: number, target: ConditionalCellFormat) => boolean

function colorScaleEffect(entry: PreparedRule, computation: Computation): VisualEffect | null {
  const rule = entry.rule
  const cfvo = Array.isArray(rule.cfvo) ? rule.cfvo : []
  const colors = Array.isArray(rule.color) ? rule.color : []
  const size = Math.min(cfvo.length, colors.length)
  if (size < 2) return null
  const stats = computation.stats(entry)
  if (!stats.count) return null
  const points: Array<{ at: number; css: string; rgb: [number, number, number] | null }> = []
  for (let index = 0; index < size; index += 1) {
    const at = cfvoNumber(cfvo[index], stats, entry, computation, index === 0 ? 'min' : 'max')
    if (at === null || !Number.isFinite(at)) return null
    const css = computation.color(colors[index], '#FFFFFF')
    points.push({ at, css, rgb: parseCssColor(css) })
  }
  const cache = new Map<number, string>()
  const colorFor = (value: number): string => {
    if (value <= points[0].at) return points[0].css
    const last = points[points.length - 1]
    if (value >= last.at) return last.css
    for (let index = 0; index < points.length - 1; index += 1) {
      const low = points[index]
      const high = points[index + 1]
      if (value > high.at) continue
      if (high.at <= low.at) return high.css
      const amount = (value - low.at) / (high.at - low.at)
      if (!low.rgb || !high.rgb) return amount < 0.5 ? low.css : high.css
      return rgbCss([
        low.rgb[0] + (high.rgb[0] - low.rgb[0]) * amount,
        low.rgb[1] + (high.rgb[1] - low.rgb[1]) * amount,
        low.rgb[2] + (high.rgb[2] - low.rgb[2]) * amount,
      ])
    }
    return last.css
  }
  return (row, col, target) => {
    const value = computation.value(row, col)
    if (typeof value !== 'number' || !Number.isFinite(value)) return false
    if (target.fill !== undefined) return true
    let css = cache.get(value)
    if (css === undefined) {
      css = colorFor(value)
      if (cache.size < 50_000) cache.set(value, css)
    }
    target.fill = css
    target.colorScale = css
    return true
  }
}

function dataBarEffect(entry: PreparedRule, computation: Computation): VisualEffect | null {
  const rule = entry.rule
  const stats = computation.stats(entry)
  if (!stats.count) return null
  const cfvo = Array.isArray(rule.cfvo) ? rule.cfvo : []
  let low = cfvoNumber(cfvo[0], stats, entry, computation, 'min')
  let high = cfvoNumber(cfvo[1], stats, entry, computation, 'max')
  if (low === null || high === null || !Number.isFinite(low) || !Number.isFinite(high)) return null
  if (high < low) [low, high] = [high, low]
  const singleColor = !Array.isArray(rule.color) ? rule.color : rule.color[0]
  const color = computation.color(singleColor, '#638EC6') || '#638EC6'
  const negativeColor = rule.negativeBarColorSameAsPositive === true ? color : computation.color(rule.negativeFillColor, '#FF0000') || '#FF0000'
  const hasBorder = rule.border === true
  const borderColor = hasBorder ? computation.color(rule.borderColor, color) || color : undefined
  const negativeBorder = hasBorder
    ? rule.negativeBarBorderColorSameAsPositive === false ? computation.color(rule.negativeBorderColor, '#FF0000') || '#FF0000' : borderColor
    : undefined
  const gradient = rule.gradient !== false
  const showValue = rule.showValue !== false
  const rtl = rule.direction === 'rightToLeft'
  const minLength = Math.min(100, Math.max(0, Number.isFinite(Number(rule.minLength)) && rule.minLength !== undefined ? Number(rule.minLength) : 10)) / 100
  const maxLength = Math.min(100, Math.max(0, Number.isFinite(Number(rule.maxLength)) && rule.maxLength !== undefined ? Number(rule.maxLength) : 90)) / 100
  const axisMode = rule.axisPosition === 'middle' ? 'middle' : rule.axisPosition === 'none' ? 'none' : 'automatic'
  const axisColor = computation.color(rule.axisColor, '#000000') || '#000000'
  let axis: number | undefined
  if (axisMode === 'middle') axis = 0.5
  else if (axisMode === 'automatic' && low < 0) axis = high <= 0 ? 1 : -low / (high - low)
  const span = high - low
  return (row, col, target) => {
    const value = computation.value(row, col)
    if (typeof value !== 'number' || !Number.isFinite(value)) return false
    if (target.dataBar !== undefined) return true
    const clamped = Math.min(high, Math.max(low, value))
    const negative = value < 0
    let fraction: number
    if (axis === undefined) {
      fraction = span > 0 ? minLength + ((clamped - low) / span) * (maxLength - minLength) : maxLength
    } else if (negative) {
      const extreme = Math.min(low, 0)
      fraction = extreme < 0 ? (Math.max(value, extreme) / extreme) * axis : 0
    } else {
      const extreme = Math.max(high, 0)
      fraction = extreme > 0 ? (Math.min(value, extreme) / extreme) * (1 - axis) : 0
    }
    target.dataBar = {
      fraction: Math.min(1, Math.max(0, fraction)),
      color: negative ? negativeColor : color,
      ...(negative ? { negative: true } : {}),
      ...(axis !== undefined ? { axis, axisColor } : {}),
      gradient,
      ...(hasBorder ? { border: negative ? negativeBorder : borderColor } : {}),
      showValue,
      ...(rtl ? { rtl: true } : {}),
    }
    return true
  }
}

/** Default icon thresholds (percent) for a set of `size` icons. */
export function defaultIconThresholds(size: number): number[] {
  if (size === 4) return [0, 25, 50, 75]
  if (size === 5) return [0, 20, 40, 60, 80]
  return [0, 33, 67]
}

function iconSetEffect(entry: PreparedRule, computation: Computation): VisualEffect | null {
  const rule = entry.rule
  const set = normalizeIconSetName(rule.iconSet)
  const size = iconSetSize(set)
  const stats = computation.stats(entry)
  if (!stats.count) return null
  const source: ConditionalValueObject[] = Array.isArray(rule.cfvo) && rule.cfvo.length === size
    ? rule.cfvo
    : defaultIconThresholds(size).map((value) => ({ type: 'percent', value }))
  const thresholds: Array<{ at: number; inclusive: boolean }> = []
  for (let index = 0; index < size; index += 1) {
    const at = cfvoNumber(source[index], stats, entry, computation, 'min')
    thresholds.push({ at: at === null ? Number.NaN : at, inclusive: source[index]?.gte !== false })
  }
  const reverse = rule.reverse === true
  const showValue = rule.showValue !== false
  const custom = Array.isArray(rule.icons) && rule.icons.length ? rule.icons : null
  return (row, col, target) => {
    const value = computation.value(row, col)
    if (typeof value !== 'number' || !Number.isFinite(value)) return false
    if (target.icon !== undefined) return true
    let index = 0
    for (let position = size - 1; position >= 1; position -= 1) {
      const threshold = thresholds[position]
      if (Number.isNaN(threshold.at)) continue
      if (threshold.inclusive ? value >= threshold.at : value > threshold.at) {
        index = position
        break
      }
    }
    if (custom) {
      const icon = custom[index]
      if (!icon || !icon.iconSet || icon.iconSet === 'NoIcons') return true
      target.icon = { set: normalizeIconSetName(icon.iconSet), index: Math.max(0, Math.floor(Number(icon.iconId) || 0)), showValue }
      return true
    }
    target.icon = { set, index: reverse ? size - 1 - index : index, showValue }
    return true
  }
}

/**
 * Evaluate a sheet's conditional formats. Returns the formatting for every affected cell,
 * keyed by A1 address. Rules apply in Excel priority order; `stopIfTrue` halts lower-priority
 * rules for that cell and a lower-priority rule only fills properties still unset.
 *
 * With `visible`, per-cell conditions are evaluated only inside those bounds; range statistics
 * (top/bottom, averages, duplicates, scales, bars, icons) still cover each rule's whole range.
 */
export function computeConditionalFormats(
  rules: unknown[] | undefined,
  host: ConditionalFormatHost,
  visible?: CellBounds | null,
): Map<string, ConditionalCellFormat> {
  const output = new Map<string, ConditionalCellFormat>()
  const prepared = flattenRules(rules)
  if (!prepared.length) return output
  const window = visible && Number.isFinite(visible.top) && Number.isFinite(visible.left)
    ? {
      top: Math.max(0, Math.floor(Math.min(visible.top, visible.bottom))),
      bottom: Math.min(MAX_SHEET_ROWS - 1, Math.floor(Math.max(visible.top, visible.bottom))),
      left: Math.max(0, Math.floor(Math.min(visible.left, visible.right))),
      right: Math.min(MAX_SHEET_COLUMNS - 1, Math.floor(Math.max(visible.left, visible.right))),
    }
    : null
  const computation = new Computation(host, window)
  const formats = new Map<number, ConditionalCellFormat>()
  const stopped = new Set<number>()

  for (const entry of prepared) {
    const ranges = computation.applyRanges(entry.ranges)
    if (!ranges.length) continue
    const stop = entry.rule.stopIfTrue === true
    if (ruleUsesStyle(entry.kind)) {
      const fragment = styleFragment(entry.rule.style, computation.color)
      // A rule without formatting still matters when it stops lower-priority rules.
      if (!fragment && !stop) continue
      const matcher = styleMatcher(entry, computation)
      if (!matcher) continue
      forEachCell(ranges, MAX_APPLY_CELLS, (row, col) => {
        const key = row * MAX_SHEET_COLUMNS + col
        if (stopped.size && stopped.has(key)) return
        if (!matcher(row, col)) return
        if (fragment) {
          let target = formats.get(key)
          if (!target) {
            target = {}
            formats.set(key, target)
          }
          mergeFragment(target, fragment)
        }
        if (stop) stopped.add(key)
      })
      continue
    }
    const effect = entry.kind === 'colorScale'
      ? colorScaleEffect(entry, computation)
      : entry.kind === 'dataBar'
        ? dataBarEffect(entry, computation)
        : iconSetEffect(entry, computation)
    if (!effect) continue
    forEachCell(ranges, MAX_APPLY_CELLS, (row, col) => {
      const key = row * MAX_SHEET_COLUMNS + col
      if (stopped.size && stopped.has(key)) return
      let target = formats.get(key)
      const created = !target
      if (!target) target = {}
      const applied = effect(row, col, target)
      if (!applied) return
      if (created && Object.keys(target).length) formats.set(key, target)
      if (stop) stopped.add(key)
    })
  }

  for (const [key, format] of formats) {
    const row = Math.floor(key / MAX_SHEET_COLUMNS)
    const col = key - row * MAX_SHEET_COLUMNS
    output.set(columnLabel(col) + (row + 1), format)
  }
  return output
}

/**
 * CSS background for a data bar (paint it under the cell text). Returns `background-image`,
 * `background-size`, `background-position` and `background-repeat` values.
 */
export function dataBarBackground(bar: ConditionalDataBar): { backgroundImage: string; backgroundSize: string; backgroundPosition: string; backgroundRepeat: string } {
  const fraction = Math.min(1, Math.max(0, bar.fraction))
  const axis = bar.axis
  let start = axis === undefined ? 0 : bar.negative ? axis - fraction : axis
  if (bar.rtl) start = 1 - start - fraction
  start = Math.min(1, Math.max(0, start))
  // A percentage background-position p places an image of width w at p * (100% - w);
  // convert the desired left edge so the bar starts exactly at `start` of the cell.
  const position = (left: number, width: number) => `${(width < 0.9999 ? (left / (1 - width)) * 100 : 0).toFixed(3)}%`
  const width = `${(fraction * 100).toFixed(3)}%`
  const layers: Array<{ image: string; size: string; position: string }> = []
  if (axis !== undefined && axis > 0 && axis < 1) {
    const axisAt = bar.rtl ? 1 - axis : axis
    layers.push({
      image: `repeating-linear-gradient(180deg, ${bar.axisColor || '#000000'} 0 2px, transparent 2px 4px)`,
      size: '1px 100%',
      position: `${(axisAt * 100).toFixed(3)}% 0`,
    })
  }
  const towardsLeft = bar.negative ? !bar.rtl : Boolean(bar.rtl)
  layers.push({
    image: bar.gradient
      ? `linear-gradient(${towardsLeft ? 270 : 90}deg, ${bar.color}, ${mixCssColors(bar.color, '#FFFFFF', 0.85)})`
      : `linear-gradient(${bar.color}, ${bar.color})`,
    size: `${width} calc(100% - 6px)`,
    position: `${position(start, fraction)} 50%`,
  })
  if (bar.border) {
    layers.push({
      image: `linear-gradient(${bar.border}, ${bar.border})`,
      size: `${width} calc(100% - 4px)`,
      position: `${position(start, fraction)} 50%`,
    })
  }
  return {
    backgroundImage: layers.map((layer) => layer.image).join(', '),
    backgroundSize: layers.map((layer) => layer.size).join(', '),
    backgroundPosition: layers.map((layer) => layer.position).join(', '),
    backgroundRepeat: layers.map(() => 'no-repeat').join(', '),
  }
}

// ---------------------------------------------------------------------------------------------
// Authoring helpers (exact ExcelJS model shapes)
// ---------------------------------------------------------------------------------------------

export interface RuleOptions {
  style?: CellStyle
  stopIfTrue?: boolean
  priority?: number
}

function baseRule(type: string, options: RuleOptions | undefined): ConditionalRule {
  const rule: ConditionalRule = { type }
  if (options?.priority !== undefined) rule.priority = options.priority
  if (options?.stopIfTrue) rule.stopIfTrue = true
  if (options?.style) rule.style = cloneValue(options.style)
  return rule
}

function operandFormula(value: string | number): string {
  if (typeof value === 'number') return Number.isFinite(value) ? String(value) : '0'
  return stripEquals(String(value))
}

function quoteFormulaText(text: string): string {
  return `"${text.replace(/"/g, '""')}"`
}

/** Top-left cell (relative A1) of the first range of an sqref. */
export function ruleAnchorAddress(ref: string): string {
  const first = parseRangeList(ref)[0]
  return first ? cellAddress(first.top, first.left) : 'A1'
}

/** The formula Excel stores for a text/blank/error/date rule anchored at `anchor`. */
export function derivedRuleFormula(rule: ConditionalRule, anchor: string): string | null {
  const kind = conditionalRuleKind(rule)
  const textKind = kind === 'containsText' || kind === 'notContainsText' || kind === 'beginsWith' || kind === 'endsWith'
  // Without `text` the stored formula is the only definition of a text rule: keep it.
  if (textKind && typeof rule.text !== 'string') return null
  const text = quoteFormulaText(String(rule.text ?? ''))
  switch (kind) {
    case 'containsText': return `NOT(ISERROR(SEARCH(${text},${anchor})))`
    case 'notContainsText': return `ISERROR(SEARCH(${text},${anchor}))`
    case 'beginsWith': return `LEFT(${anchor},LEN(${text}))=${text}`
    case 'endsWith': return `RIGHT(${anchor},LEN(${text}))=${text}`
    case 'containsBlanks': return `LEN(TRIM(${anchor}))=0`
    case 'notContainsBlanks': return `LEN(TRIM(${anchor}))>0`
    case 'containsErrors': return `ISERROR(${anchor})`
    case 'notContainsErrors': return `NOT(ISERROR(${anchor}))`
    case 'timePeriod':
      switch (rule.timePeriod) {
        case 'today': return `FLOOR(${anchor},1)=TODAY()`
        case 'yesterday': return `FLOOR(${anchor},1)=TODAY()-1`
        case 'tomorrow': return `FLOOR(${anchor},1)=TODAY()+1`
        case 'last7Days': return `AND(TODAY()-FLOOR(${anchor},1)<=6,FLOOR(${anchor},1)<=TODAY())`
        case 'thisWeek': return `AND(TODAY()-ROUNDDOWN(${anchor},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${anchor},0)-TODAY()<=7-WEEKDAY(TODAY()))`
        case 'lastWeek': return `AND(TODAY()-ROUNDDOWN(${anchor},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${anchor},0)<(WEEKDAY(TODAY())+7))`
        case 'nextWeek': return `AND(ROUNDDOWN(${anchor},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${anchor},0)-TODAY()<(15-WEEKDAY(TODAY())))`
        case 'thisMonth': return `AND(MONTH(${anchor})=MONTH(TODAY()),YEAR(${anchor})=YEAR(TODAY()))`
        case 'lastMonth': return `AND(MONTH(${anchor})=MONTH(EDATE(TODAY(),0-1)),YEAR(${anchor})=YEAR(EDATE(TODAY(),0-1)))`
        case 'nextMonth': return `AND(MONTH(${anchor})=MONTH(EDATE(TODAY(),0+1)),YEAR(${anchor})=YEAR(EDATE(TODAY(),0+1)))`
        default: return null
      }
    default: return null
  }
}

/** Refresh the Excel-generated formula of text/blank/error/date rules for their range. */
export function withDerivedFormula(rule: ConditionalRule, ref: string): ConditionalRule {
  const formula = derivedRuleFormula(rule, ruleAnchorAddress(ref))
  return formula === null ? rule : { ...rule, formulae: [formula] }
}

/** "Cell value <operator> ..." — values are formulas (numbers become literals, a leading "=" is dropped). */
export function createCellIsRule(operator: CellIsOperator, values: Array<string | number>, options?: RuleOptions): ConditionalRule {
  const count = operator === 'between' || operator === 'notBetween' ? 2 : 1
  const formulae = values.slice(0, count).map(operandFormula)
  while (formulae.length < count) formulae.push('0')
  return { ...baseRule('cellIs', options), operator, formulae }
}

/** Text contains / does not contain / begins with / ends with (case-insensitive; contains supports * ? ~). */
export function createTextRule(operator: TextRuleOperator, text: string, options?: RuleOptions & { ref?: string }): ConditionalRule {
  const rule: ConditionalRule = operator === 'containsText'
    ? { ...baseRule('containsText', options), operator: 'containsText', text }
    : operator === 'notContainsText'
      ? { ...baseRule('notContainsText', options), operator: 'notContains', text }
      : { ...baseRule(operator, options), operator, text }
  return withDerivedFormula(rule, options?.ref || 'A1')
}

/** Blank cells (`blanks = true`) or non-blank cells. */
export function createBlankRule(blanks: boolean, options?: RuleOptions & { ref?: string }): ConditionalRule {
  const rule = { ...baseRule('containsText', options), operator: blanks ? 'containsBlanks' : 'notContainsBlanks' }
  return withDerivedFormula(rule, options?.ref || 'A1')
}

/** Error cells (`errors = true`) or cells without errors. */
export function createErrorRule(errors: boolean, options?: RuleOptions & { ref?: string }): ConditionalRule {
  const rule = { ...baseRule('containsText', options), operator: errors ? 'containsErrors' : 'notContainsErrors' }
  return withDerivedFormula(rule, options?.ref || 'A1')
}

/** "A date occurring ..." */
export function createDateRule(timePeriod: TimePeriod, options?: RuleOptions & { ref?: string }): ConditionalRule {
  return withDerivedFormula({ ...baseRule('timePeriod', options), timePeriod }, options?.ref || 'A1')
}

/** Top/bottom N items or N percent. */
export function createTopBottomRule(config: { rank?: number; percent?: boolean; bottom?: boolean }, options?: RuleOptions): ConditionalRule {
  const rank = Math.max(1, Math.min(config.percent ? 100 : 1000, Math.floor(Number(config.rank ?? 10)) || 10))
  const rule: ConditionalRule = { ...baseRule('top10', options), rank }
  if (config.percent) rule.percent = true
  if (config.bottom) rule.bottom = true
  return rule
}

/** Above/below average, optionally "or equal" or N standard deviations. */
export function createAverageRule(config: { below?: boolean; equal?: boolean; stdDev?: number }, options?: RuleOptions): ConditionalRule {
  const rule: ConditionalRule = baseRule('aboveAverage', options)
  if (config.below) rule.aboveAverage = false
  const deviations = Math.floor(Number(config.stdDev) || 0)
  if (deviations > 0) rule.stdDev = Math.min(3, deviations)
  else if (config.equal) rule.equalAverage = true
  return rule
}

/** Duplicate (default) or unique values — case-insensitive, blanks ignored. */
export function createDuplicateRule(unique = false, options?: RuleOptions): ConditionalRule {
  return baseRule(unique ? 'uniqueValues' : 'duplicateValues', options)
}

/** "Use a formula to determine which cells to format" (relative to the range's top-left cell). */
export function createFormulaRule(formula: string, options?: RuleOptions): ConditionalRule {
  return { ...baseRule('expression', options), formulae: [stripEquals(formula)] }
}

export interface ScalePoint {
  type: CfvoType
  value?: number | string
  /** CSS hex ("#63BE7B") or ARGB. */
  color: string
}

/** 2- or 3-colour scale. */
export function createColorScaleRule(points: ScalePoint[], options?: Omit<RuleOptions, 'style'>): ConditionalRule {
  const used = points.slice(0, 3)
  return {
    ...baseRule('colorScale', options),
    cfvo: used.map((point) => cfvoFor(point.type, point.value)),
    color: used.map((point) => toArgbColor(point.color)),
  }
}

function cfvoFor(type: CfvoType | string, value?: number | string, gte?: boolean): ConditionalValueObject {
  const cfvo: ConditionalValueObject = { type }
  if (type !== 'min' && type !== 'max' && type !== 'autoMin' && type !== 'autoMax') {
    const numeric = typeof value === 'number' ? value : numericValue(value)
    cfvo.value = numeric !== null ? numeric : typeof value === 'string' ? stripEquals(value) : 0
  }
  if (gte === false) cfvo.gte = false
  return cfvo
}

export interface DataBarOptions {
  color?: string
  gradient?: boolean
  showValue?: boolean
  border?: boolean
  min?: { type: CfvoType; value?: number | string }
  max?: { type: CfvoType; value?: number | string }
  negativeColor?: string
  axisPosition?: 'automatic' | 'middle' | 'none'
  axisColor?: string
  direction?: 'context' | 'leftToRight' | 'rightToLeft'
}

/** Excel 2010+ data bar (writes the x14 extension, like Excel's gallery bars). */
export function createDataBarRule(config: DataBarOptions = {}, options?: Omit<RuleOptions, 'style'>): ConditionalRule {
  const color = toArgbColor(config.color || '#638EC6')
  const gradient = config.gradient !== false
  const border = config.border ?? gradient
  const negative = toArgbColor(config.negativeColor || '#FF0000')
  const rule: ConditionalRule = {
    ...baseRule('dataBar', options),
    cfvo: [cfvoFor(config.min?.type || 'autoMin', config.min?.value), cfvoFor(config.max?.type || 'autoMax', config.max?.value)],
    color,
    minLength: 0,
    maxLength: 100,
    gradient,
    border,
    negativeFillColor: negative,
    axisColor: toArgbColor(config.axisColor || '#000000'),
  }
  if (border) {
    rule.borderColor = { ...color }
    rule.negativeBorderColor = { ...negative }
    rule.negativeBarBorderColorSameAsPositive = false
  }
  if (config.showValue === false) rule.showValue = false
  if (config.axisPosition && config.axisPosition !== 'automatic') rule.axisPosition = config.axisPosition
  if (config.direction && config.direction !== 'context') rule.direction = config.direction
  return rule
}

export interface IconThreshold {
  type: CfvoType
  value: number | string
  /** false = ">" instead of ">=". */
  gte?: boolean
}

/** Icon set; `thresholds` are the N-1 lower bounds for icons 1..N-1 (icon 0 takes the rest). */
export function createIconSetRule(iconSet: IconSetName, config: { reverse?: boolean; showValue?: boolean; thresholds?: IconThreshold[] } = {}, options?: Omit<RuleOptions, 'style'>): ConditionalRule {
  const size = iconSetSize(iconSet)
  const defaults = defaultIconThresholds(size)
  const thresholds: IconThreshold[] = config.thresholds && config.thresholds.length === size - 1
    ? config.thresholds
    : defaults.slice(1).map((value) => ({ type: 'percent' as CfvoType, value }))
  const rule: ConditionalRule = {
    ...baseRule('iconSet', options),
    iconSet,
    cfvo: [{ type: 'percent', value: 0 }, ...thresholds.map((threshold) => cfvoFor(threshold.type, threshold.value, threshold.gte))],
  }
  if (config.reverse) rule.reverse = true
  if (config.showValue === false) rule.showValue = false
  return rule
}

// ---------------------------------------------------------------------------------------------
// Preset styles
// ---------------------------------------------------------------------------------------------

export type ConditionalStylePresetId =
  | 'lightRedFillDarkRedText'
  | 'yellowFillDarkYellowText'
  | 'greenFillDarkGreenText'
  | 'lightRedFill'
  | 'redText'
  | 'redBorder'
  | 'custom'

function solidFill(argb: string) {
  return { type: 'pattern', pattern: 'solid', fgColor: { argb }, bgColor: { argb } }
}

const RED_BORDER_SIDE = { style: 'thin', color: { argb: 'FF9C0006' } }

export const CONDITIONAL_STYLE_PRESETS: ReadonlyArray<{ id: Exclude<ConditionalStylePresetId, 'custom'>; label: string; style: CellStyle }> = [
  { id: 'lightRedFillDarkRedText', label: 'Light red fill with dark red text', style: { font: { color: { argb: 'FF9C0006' } }, fill: solidFill('FFFFC7CE') } },
  { id: 'yellowFillDarkYellowText', label: 'Yellow fill with dark yellow text', style: { font: { color: { argb: 'FF9C5700' } }, fill: solidFill('FFFFEB9C') } },
  { id: 'greenFillDarkGreenText', label: 'Green fill with dark green text', style: { font: { color: { argb: 'FF006100' } }, fill: solidFill('FFC6EFCE') } },
  { id: 'lightRedFill', label: 'Light red fill', style: { fill: solidFill('FFFFC7CE') } },
  { id: 'redText', label: 'Red text', style: { font: { color: { argb: 'FF9C0006' } } } },
  {
    id: 'redBorder',
    label: 'Red border',
    style: { border: { top: { ...RED_BORDER_SIDE }, left: { ...RED_BORDER_SIDE }, bottom: { ...RED_BORDER_SIDE }, right: { ...RED_BORDER_SIDE } } },
  },
]

/** A fresh copy of a preset style. */
export function conditionalPresetStyle(id: Exclude<ConditionalStylePresetId, 'custom'>): CellStyle {
  const preset = CONDITIONAL_STYLE_PRESETS.find((item) => item.id === id) ?? CONDITIONAL_STYLE_PRESETS[0]
  return cloneValue(preset.style)
}

export interface SimpleStyleSpec {
  /** CSS hex colours; omit for "no change". */
  fontColor?: string | null
  fillColor?: string | null
  borderColor?: string | null
  bold?: boolean
  italic?: boolean
  underline?: boolean
  strike?: boolean
}

/** Build a dxf CellStyle (argb colours) from simple settings. */
export function createConditionalStyle(spec: SimpleStyleSpec): CellStyle {
  const style: CellStyle = {}
  const font: NonNullable<CellStyle['font']> = {}
  if (spec.fontColor) font.color = toArgbColor(spec.fontColor)
  if (spec.bold) font.bold = true
  if (spec.italic) font.italic = true
  if (spec.underline) font.underline = true
  if (spec.strike) font.strike = true
  if (Object.keys(font).length) style.font = font
  if (spec.fillColor) style.fill = solidFill(toArgbColor(spec.fillColor).argb || 'FFFFFFFF')
  if (spec.borderColor) {
    const side = { style: 'thin', color: toArgbColor(spec.borderColor) }
    style.border = { top: { ...side }, left: { ...side }, bottom: { ...side }, right: { ...side } }
  }
  return style
}

/** Read simple settings back from a dxf style (for editors). */
export function describeConditionalStyle(style: unknown, cssColor: (color: unknown, fallback?: string) => string = defaultConditionalCssColor): Required<Pick<SimpleStyleSpec, 'bold' | 'italic' | 'underline' | 'strike'>> & { fontColor: string | null; fillColor: string | null; borderColor: string | null } {
  const fragment = styleFragment(style, cssColor)
  const borderSide = fragment?.border?.top || fragment?.border?.left || fragment?.border?.bottom || fragment?.border?.right
  return {
    fontColor: fragment?.font?.color ?? null,
    fillColor: fragment?.fill ?? null,
    borderColor: borderSide ? borderSide.split(' ').pop() || null : null,
    bold: fragment?.font?.bold === true,
    italic: fragment?.font?.italic === true,
    underline: fragment?.font?.underline === true,
    strike: fragment?.font?.strike === true,
  }
}

/** Preset whose style matches, or "custom". */
export function matchConditionalStylePreset(style: unknown): ConditionalStylePresetId {
  const current = describeConditionalStyle(style)
  for (const preset of CONDITIONAL_STYLE_PRESETS) {
    const candidate = describeConditionalStyle(preset.style)
    if (
      candidate.fontColor === current.fontColor && candidate.fillColor === current.fillColor && candidate.borderColor === current.borderColor
      && candidate.bold === current.bold && candidate.italic === current.italic && candidate.underline === current.underline && candidate.strike === current.strike
    ) return preset.id
  }
  return 'custom'
}

/** Resolve a rule's dxf into the same shape the grid receives (for previews). */
export function conditionalStylePreview(style: unknown, cssColor: (color: unknown, fallback?: string) => string = defaultConditionalCssColor): ConditionalCellFormat {
  const fragment = styleFragment(style, cssColor)
  const format: ConditionalCellFormat = {}
  if (fragment) mergeFragment(format, fragment)
  return format
}

// ---------------------------------------------------------------------------------------------
// Rule management (pure; never mutates the input)
// ---------------------------------------------------------------------------------------------

function cloneValue<T>(value: T): T {
  if (value === null || typeof value !== 'object') return value
  if (Array.isArray(value)) return value.map((item) => cloneValue(item)) as unknown as T
  const result: Record<string, unknown> = {}
  for (const [key, item] of Object.entries(value as Record<string, unknown>)) {
    if (item !== undefined) result[key] = cloneValue(item)
  }
  return result as T
}

interface FlatRule {
  blockIndex: number
  ruleIndex: number
  ref: string
  rule: ConditionalRule
  priority: number
}

function flatRules(blocks: readonly ConditionalFormattingBlock[]): FlatRule[] {
  const items: Array<FlatRule & { order: number }> = []
  let order = 0
  blocks.forEach((block, blockIndex) => {
    block.rules.forEach((rule, ruleIndex) => {
      const priority = Number(rule.priority)
      items.push({
        blockIndex,
        ruleIndex,
        ref: block.ref,
        rule,
        priority: Number.isFinite(priority) && priority > 0 ? priority : Number.POSITIVE_INFINITY,
        order: order++,
      })
    })
  })
  items.sort((a, b) => a.priority - b.priority || a.order - b.order)
  return items.map((item, index) => ({ ...item, priority: index + 1 }))
}

/**
 * Deep-cloned, validated blocks with dense unique priorities (1 = highest) in current
 * priority order. Every management helper starts from this, so rule ids stay consistent.
 */
export function normalizeConditionalFormattings(input: unknown): ConditionalFormattingBlock[] {
  const blocks: ConditionalFormattingBlock[] = []
  if (Array.isArray(input)) {
    for (const item of input) {
      if (!item || typeof item !== 'object') continue
      const record = item as ConditionalFormattingBlock
      if (typeof record.ref !== 'string' || !Array.isArray(record.rules)) continue
      const rules = record.rules.filter((rule) => rule && typeof rule === 'object' && typeof rule.type === 'string')
      if (!rules.length) continue
      blocks.push(cloneValue({ ...record, rules }))
    }
  }
  for (const item of flatRules(blocks)) blocks[item.blockIndex].rules[item.ruleIndex].priority = item.priority
  return blocks
}

export interface ConditionalRuleEntry {
  /** Rule id: its normalized priority (1 = highest). Stable until the next edit. */
  id: number
  ref: string
  rule: ConditionalRule
  kind: ConditionalRuleKind | null
}

function rangesIntersect(a: readonly CellBounds[], b: readonly CellBounds[]): boolean {
  for (const left of a) for (const right of b) if (intersect(left, right)) return true
  return false
}

/** Rules (priority order) whose range intersects `range` (an sqref string or bounds); all rules when omitted. */
export function listRulesForRange(conditionalFormattings: unknown, range?: string | CellBounds | null): ConditionalRuleEntry[] {
  const blocks = normalizeConditionalFormattings(conditionalFormattings)
  const filter = range == null ? null : typeof range === 'string' ? parseRangeList(range) : [range]
  return flatRules(blocks)
    .filter((item) => !filter || rangesIntersect(parseRangeList(item.ref), filter))
    .map((item) => ({ id: item.priority, ref: item.ref, rule: item.rule, kind: conditionalRuleKind(item.rule) }))
}

/** Find a rule by id. */
export function findRule(conditionalFormattings: unknown, id: number): ConditionalRuleEntry | null {
  return listRulesForRange(conditionalFormattings).find((entry) => entry.id === id) ?? null
}

function reprioritize(blocks: ConditionalFormattingBlock[], ordered: ConditionalRule[]): ConditionalFormattingBlock[] {
  ordered.forEach((rule, index) => { rule.priority = index + 1 })
  return blocks.filter((block) => block.rules.length > 0)
}

function orderedRules(blocks: ConditionalFormattingBlock[]): ConditionalRule[] {
  return flatRules(blocks).map((item) => blocks[item.blockIndex].rules[item.ruleIndex])
}

/** Add a rule for `ref` at the top of the priority list (like Excel). Returns the new model. */
export function addRule(conditionalFormattings: unknown, ref: string, rule: ConditionalRule): ConditionalFormattingBlock[] {
  const parsed = parseRangeInput(ref)
  if ('error' in parsed) throw new Error(parsed.error)
  const blocks = normalizeConditionalFormattings(conditionalFormattings)
  const ordered = orderedRules(blocks)
  const added = withDerivedFormula(cloneValue(rule), parsed.ref)
  blocks.unshift({ ref: parsed.ref, rules: [added] })
  return reprioritize(blocks, [added, ...ordered])
}

/** Replace a rule and/or its range. A rule sharing a block is split into its own block when its range changes. */
export function updateRule(conditionalFormattings: unknown, id: number, next: { ref?: string; rule?: ConditionalRule }): ConditionalFormattingBlock[] {
  const blocks = normalizeConditionalFormattings(conditionalFormattings)
  const flat = flatRules(blocks)
  const target = flat.find((item) => item.priority === id)
  if (!target) return blocks
  const block = blocks[target.blockIndex]
  let ref = block.ref
  if (next.ref !== undefined) {
    const parsed = parseRangeInput(next.ref)
    if ('error' in parsed) throw new Error(parsed.error)
    ref = parsed.ref
  }
  const replacement = withDerivedFormula(cloneValue(next.rule ?? block.rules[target.ruleIndex]), ref)
  replacement.priority = id
  if (ref === block.ref) {
    block.rules[target.ruleIndex] = replacement
  } else if (block.rules.length === 1) {
    block.ref = ref
    block.rules[0] = replacement
  } else {
    block.rules.splice(target.ruleIndex, 1)
    blocks.splice(target.blockIndex + 1, 0, { ref, rules: [replacement] })
  }
  return blocks.filter((item) => item.rules.length > 0)
}

/** Delete a rule. */
export function deleteRule(conditionalFormattings: unknown, id: number): ConditionalFormattingBlock[] {
  const blocks = normalizeConditionalFormattings(conditionalFormattings)
  const target = flatRules(blocks).find((item) => item.priority === id)
  if (!target) return blocks
  blocks[target.blockIndex].rules.splice(target.ruleIndex, 1)
  return reprioritize(blocks, orderedRules(blocks.filter((block) => block.rules.length > 0)))
}

/** Move a rule `delta` places in priority (negative = higher priority). Returns the model; the rule's new id is clamp(id + delta). */
export function moveRulePriority(conditionalFormattings: unknown, id: number, delta: number): ConditionalFormattingBlock[] {
  const blocks = normalizeConditionalFormattings(conditionalFormattings)
  const ordered = orderedRules(blocks)
  const from = id - 1
  if (from < 0 || from >= ordered.length || !Number.isFinite(delta)) return blocks
  const to = Math.max(0, Math.min(ordered.length - 1, from + Math.trunc(delta)))
  const [moved] = ordered.splice(from, 1)
  ordered.splice(to, 0, moved)
  return reprioritize(blocks, ordered)
}

/** Shift relative references in a rule's formulas so they keep meaning after its anchor moves. */
function reanchorRule(rule: ConditionalRule, from: { row: number; col: number }, to: { row: number; col: number }, ref: string): ConditionalRule {
  const derived = derivedRuleFormula(rule, ruleAnchorAddress(ref))
  if (derived !== null) return { ...rule, formulae: [derived] }
  if (!Array.isArray(rule.formulae) || (from.row === to.row && from.col === to.col)) return rule
  return {
    ...rule,
    formulae: rule.formulae.map((formula) => typeof formula === 'string' ? shiftConditionalFormula(formula, to.row - from.row, to.col - from.col) : formula),
  }
}

/**
 * Excel "Clear Rules from Selected Cells": remove `bounds` from every rule's range, splitting
 * ranges into the remaining rectangles and dropping rules left with no cells.
 */
export function clearRulesInRange(conditionalFormattings: unknown, bounds: CellBounds | string): ConditionalFormattingBlock[] {
  const cuts = typeof bounds === 'string' ? parseRangeList(bounds) : [bounds]
  const blocks = normalizeConditionalFormattings(conditionalFormattings)
  if (!cuts.length) return blocks
  const result: ConditionalFormattingBlock[] = []
  for (const block of blocks) {
    const original = parseRangeList(block.ref)
    if (!original.length) continue
    let remaining = original
    for (const cut of cuts) remaining = remaining.flatMap((range) => subtractBounds(range, cut))
    if (!remaining.length) continue
    const ref = formatRangeList(remaining)
    if (ref === formatRangeList(original)) {
      result.push(block)
      continue
    }
    const from = { row: original[0].top, col: original[0].left }
    const to = { row: remaining[0].top, col: remaining[0].left }
    result.push({ ...block, ref, rules: block.rules.map((rule) => reanchorRule(rule, from, to, ref)) })
  }
  return reprioritize(result, orderedRules(result))
}

/** Number of cells covered by an sqref (for UI hints). */
export function rangeListCellCount(ref: string): number {
  return parseRangeList(ref).reduce((total, range) => total + boundsArea(range), 0)
}
