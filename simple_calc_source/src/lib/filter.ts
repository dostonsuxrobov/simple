/**
 * AutoFilter engine (pure): evaluates SheetFilterState criteria against calculated values and
 * display text, builds the per-column checklist, and maintains the filter range and the
 * filter-hidden rows independently of rows the user hid by hand.
 */
import type { CellData, SheetData, SheetFilterCriteria, SheetFilterState } from '../spreadsheet-types'
import type { SheetStructureOperation } from './sheet-operations'
import {
  MAX_COLUMNS,
  MAX_ROWS,
  TEXT_COLLATOR,
  addressOf,
  autoFilterBounds,
  cellHasContent,
  collationKeys,
  compareCollated,
  foldText,
  formatRange,
  formatSerialDate,
  hasWildcards,
  hostFillColor,
  hostFontColor,
  hostIsDate,
  normalizeColor,
  parseDateText,
  parseNumberText,
  parseRange,
  scalarKind,
  serialFromLocalDate,
  wildcardRegExp,
  ymdFromSerial,
} from './data-tools-core'
import type { Bounds, DataHost, Scalar, ScalarKind } from './data-tools-core'

export type FilterOperator =
  | 'equals' | 'notEquals'
  | 'beginsWith' | 'notBeginsWith' | 'endsWith' | 'notEndsWith' | 'contains' | 'notContains'
  | 'greaterThan' | 'greaterThanOrEqual' | 'lessThan' | 'lessThanOrEqual' | 'between' | 'notBetween'
  | 'isEmpty' | 'isNotEmpty'
  | 'top' | 'bottom' | 'topPercent' | 'bottomPercent' | 'aboveAverage' | 'belowAverage'
  | 'dateEquals' | 'dateBefore' | 'dateAfter' | 'dateBetween'
  | 'today' | 'yesterday' | 'tomorrow'
  | 'thisWeek' | 'lastWeek' | 'nextWeek'
  | 'thisMonth' | 'lastMonth' | 'nextMonth'
  | 'thisQuarter' | 'lastQuarter' | 'nextQuarter'
  | 'thisYear' | 'lastYear' | 'nextYear' | 'yearToDate'
  | 'Q1' | 'Q2' | 'Q3' | 'Q4'
  | 'M1' | 'M2' | 'M3' | 'M4' | 'M5' | 'M6' | 'M7' | 'M8' | 'M9' | 'M10' | 'M11' | 'M12'
  | 'customFormula'

export type FilterOperatorGroup = 'general' | 'text' | 'number' | 'date'

export interface FilterOperatorInfo {
  id: FilterOperator
  label: string
  group: FilterOperatorGroup
  /** Number of inputs the operator reads: `value` (1) or `value` and `value2` (2). */
  inputs: 0 | 1 | 2
  input?: 'text' | 'number' | 'date' | 'formula'
  /** Operators that can be combined with a second condition (And / Or). */
  combinable: boolean
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']

export const FILTER_OPERATORS: FilterOperatorInfo[] = [
  { id: 'isEmpty', label: 'Is empty', group: 'general', inputs: 0, combinable: true },
  { id: 'isNotEmpty', label: 'Is not empty', group: 'general', inputs: 0, combinable: true },
  { id: 'equals', label: 'Equals', group: 'general', inputs: 1, input: 'text', combinable: true },
  { id: 'notEquals', label: 'Does not equal', group: 'general', inputs: 1, input: 'text', combinable: true },
  { id: 'contains', label: 'Text contains', group: 'text', inputs: 1, input: 'text', combinable: true },
  { id: 'notContains', label: 'Text does not contain', group: 'text', inputs: 1, input: 'text', combinable: true },
  { id: 'beginsWith', label: 'Text begins with', group: 'text', inputs: 1, input: 'text', combinable: true },
  { id: 'notBeginsWith', label: 'Text does not begin with', group: 'text', inputs: 1, input: 'text', combinable: true },
  { id: 'endsWith', label: 'Text ends with', group: 'text', inputs: 1, input: 'text', combinable: true },
  { id: 'notEndsWith', label: 'Text does not end with', group: 'text', inputs: 1, input: 'text', combinable: true },
  { id: 'greaterThan', label: 'Greater than', group: 'number', inputs: 1, input: 'number', combinable: true },
  { id: 'greaterThanOrEqual', label: 'Greater than or equal to', group: 'number', inputs: 1, input: 'number', combinable: true },
  { id: 'lessThan', label: 'Less than', group: 'number', inputs: 1, input: 'number', combinable: true },
  { id: 'lessThanOrEqual', label: 'Less than or equal to', group: 'number', inputs: 1, input: 'number', combinable: true },
  { id: 'between', label: 'Is between', group: 'number', inputs: 2, input: 'number', combinable: false },
  { id: 'notBetween', label: 'Is not between', group: 'number', inputs: 2, input: 'number', combinable: false },
  { id: 'top', label: 'Top N items', group: 'number', inputs: 1, input: 'number', combinable: false },
  { id: 'bottom', label: 'Bottom N items', group: 'number', inputs: 1, input: 'number', combinable: false },
  { id: 'topPercent', label: 'Top N percent', group: 'number', inputs: 1, input: 'number', combinable: false },
  { id: 'bottomPercent', label: 'Bottom N percent', group: 'number', inputs: 1, input: 'number', combinable: false },
  { id: 'aboveAverage', label: 'Above average', group: 'number', inputs: 0, combinable: false },
  { id: 'belowAverage', label: 'Below average', group: 'number', inputs: 0, combinable: false },
  { id: 'dateEquals', label: 'Date is', group: 'date', inputs: 1, input: 'date', combinable: true },
  { id: 'dateBefore', label: 'Date is before', group: 'date', inputs: 1, input: 'date', combinable: true },
  { id: 'dateAfter', label: 'Date is after', group: 'date', inputs: 1, input: 'date', combinable: true },
  { id: 'dateBetween', label: 'Date is between', group: 'date', inputs: 2, input: 'date', combinable: false },
  { id: 'today', label: 'Today', group: 'date', inputs: 0, combinable: false },
  { id: 'yesterday', label: 'Yesterday', group: 'date', inputs: 0, combinable: false },
  { id: 'tomorrow', label: 'Tomorrow', group: 'date', inputs: 0, combinable: false },
  { id: 'thisWeek', label: 'This week', group: 'date', inputs: 0, combinable: false },
  { id: 'lastWeek', label: 'Last week', group: 'date', inputs: 0, combinable: false },
  { id: 'nextWeek', label: 'Next week', group: 'date', inputs: 0, combinable: false },
  { id: 'thisMonth', label: 'This month', group: 'date', inputs: 0, combinable: false },
  { id: 'lastMonth', label: 'Last month', group: 'date', inputs: 0, combinable: false },
  { id: 'nextMonth', label: 'Next month', group: 'date', inputs: 0, combinable: false },
  { id: 'thisQuarter', label: 'This quarter', group: 'date', inputs: 0, combinable: false },
  { id: 'lastQuarter', label: 'Last quarter', group: 'date', inputs: 0, combinable: false },
  { id: 'nextQuarter', label: 'Next quarter', group: 'date', inputs: 0, combinable: false },
  { id: 'thisYear', label: 'This year', group: 'date', inputs: 0, combinable: false },
  { id: 'lastYear', label: 'Last year', group: 'date', inputs: 0, combinable: false },
  { id: 'nextYear', label: 'Next year', group: 'date', inputs: 0, combinable: false },
  { id: 'yearToDate', label: 'Year to date', group: 'date', inputs: 0, combinable: false },
  ...(['Q1', 'Q2', 'Q3', 'Q4'] as const).map((id, index) => ({ id, label: `Quarter ${index + 1} (any year)`, group: 'date' as const, inputs: 0 as const, combinable: false })),
  ...MONTH_NAMES.map((name, index) => ({ id: `M${index + 1}` as FilterOperator, label: `${name} (any year)`, group: 'date' as const, inputs: 0 as const, combinable: false })),
  { id: 'customFormula', label: 'Custom formula is', group: 'general', inputs: 1, input: 'formula', combinable: false },
]

const OPERATOR_INDEX = new Map(FILTER_OPERATORS.map((info) => [info.id, info]))

export function filterOperatorInfo(id: string | undefined): FilterOperatorInfo | undefined {
  return id ? OPERATOR_INDEX.get(id as FilterOperator) : undefined
}

// ---------------------------------------------------------------------------------------------
// Criteria helpers
// ---------------------------------------------------------------------------------------------

export function criteriaIsActive(criteria: SheetFilterCriteria | null | undefined): boolean {
  if (!criteria) return false
  return Array.isArray(criteria.values) || criteria.blanks === false || Boolean(criteria.condition?.operator) || Boolean(criteria.fillColor) || Boolean(criteria.fontColor)
}

export function filterIsActive(state: SheetFilterState | null | undefined): boolean {
  return Boolean(state && Object.values(state.columns || {}).some(criteriaIsActive))
}

/** Human-readable summary for a filter button tooltip / status bar. */
export function describeCriteria(criteria: SheetFilterCriteria | null | undefined): string {
  if (!criteriaIsActive(criteria)) return 'Showing all'
  const parts: string[] = []
  if (criteria!.values) {
    const shown = criteria!.values.slice(0, 3).map((value) => value === '' ? '(Blanks)' : `"${value}"`)
    const extra = criteria!.values.length - shown.length
    if (criteria!.blanks) shown.push('(Blanks)')
    parts.push(shown.length ? `Showing ${shown.join(', ')}${extra > 0 ? ` and ${extra.toLocaleString()} more` : ''}` : 'Showing nothing')
  } else if (criteria!.blanks === false) parts.push('Hiding blanks')
  const condition = criteria!.condition
  if (condition?.operator) {
    const describe = (operator: string, value: unknown, value2?: unknown) => {
      const info = filterOperatorInfo(operator)
      const label = info?.label || operator
      if (!info || info.inputs === 0) return label
      if (info.inputs === 2) return `${label} ${String(value ?? '')} and ${String(value2 ?? '')}`
      return `${label} ${info.input === 'text' ? `"${String(value ?? '')}"` : String(value ?? '')}`
    }
    const first = filterOperatorInfo(condition.operator)
    let text = describe(condition.operator, condition.value, condition.value2)
    if (first?.combinable && condition.operator2) text += ` ${condition.join === 'or' ? 'or' : 'and'} ${describe(condition.operator2, condition.value2).toLocaleLowerCase()}`
    parts.push(text)
  }
  if (criteria!.fillColor) parts.push(criteria!.fillColor === 'none' ? 'No fill' : `Fill ${criteria!.fillColor}`)
  if (criteria!.fontColor) parts.push(criteria!.fontColor === 'none' ? 'Automatic font colour' : `Font ${criteria!.fontColor}`)
  return parts.join('; ')
}

// ---------------------------------------------------------------------------------------------
// Evaluation
// ---------------------------------------------------------------------------------------------

type RowTest = (row: number, value: Scalar | undefined, display: string, kind: ScalarKind) => boolean

interface ColumnContext {
  col: number
  top: number
  bottom: number
  host: DataHost
  todaySerial: number
  stats: { numbers: Float64Array; mean: number } | null
}

function columnStats(context: ColumnContext) {
  if (context.stats) return context.stats
  const values: number[] = []
  let sum = 0
  for (let row = context.top; row <= context.bottom; row += 1) {
    const value = context.host.valueAt(row, context.col)
    if (typeof value === 'number' && Number.isFinite(value)) { values.push(value); sum += value }
  }
  const numbers = Float64Array.from(values).sort()
  context.stats = { numbers, mean: values.length ? sum / values.length : NaN }
  return context.stats
}

function criterionNumber(value: unknown, date1904: boolean): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value !== 'string') return null
  return parseNumberText(value) ?? parseDateText(value, date1904)
}

function criterionDate(value: unknown, date1904: boolean): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? Math.floor(value) : null
  if (typeof value !== 'string') return null
  const serial = parseDateText(value, date1904) ?? parseNumberText(value)
  return serial === null ? null : Math.floor(serial)
}

/** [start, end) serial interval for a dynamic date period. */
export function datePeriodInterval(operator: string, todaySerial: number, date1904 = false): [number, number] | null {
  const today = Math.floor(todaySerial)
  const { year, month, weekday } = ymdFromSerial(today, date1904)
  const monthStart = (y: number, m: number) => {
    const normalizedYear = y + Math.floor((m - 1) / 12)
    const normalizedMonth = ((m - 1) % 12 + 12) % 12 + 1
    return serialFromYMDLocal(normalizedYear, normalizedMonth, 1, date1904)
  }
  const quarterStartMonth = Math.floor((month - 1) / 3) * 3 + 1
  switch (operator) {
    case 'today': return [today, today + 1]
    case 'yesterday': return [today - 1, today]
    case 'tomorrow': return [today + 1, today + 2]
    case 'thisWeek': return [today - weekday, today - weekday + 7]
    case 'lastWeek': return [today - weekday - 7, today - weekday]
    case 'nextWeek': return [today - weekday + 7, today - weekday + 14]
    case 'thisMonth': return [monthStart(year, month), monthStart(year, month + 1)]
    case 'lastMonth': return [monthStart(year, month - 1), monthStart(year, month)]
    case 'nextMonth': return [monthStart(year, month + 1), monthStart(year, month + 2)]
    case 'thisQuarter': return [monthStart(year, quarterStartMonth), monthStart(year, quarterStartMonth + 3)]
    case 'lastQuarter': return [monthStart(year, quarterStartMonth - 3), monthStart(year, quarterStartMonth)]
    case 'nextQuarter': return [monthStart(year, quarterStartMonth + 3), monthStart(year, quarterStartMonth + 6)]
    case 'thisYear': return [monthStart(year, 1), monthStart(year + 1, 1)]
    case 'lastYear': return [monthStart(year - 1, 1), monthStart(year, 1)]
    case 'nextYear': return [monthStart(year + 1, 1), monthStart(year + 2, 1)]
    case 'yearToDate': return [monthStart(year, 1), today + 1]
    default: return null
  }
}

function serialFromYMDLocal(year: number, month: number, day: number, date1904: boolean) {
  const base = Date.UTC(year, month - 1, day)
  return (base - Date.UTC(1899, 11, 30)) / 86_400_000 - (date1904 ? 1462 : 0)
}

function buildTest(operator: string, value: unknown, value2: unknown, context: ColumnContext): RowTest | null {
  const date1904 = Boolean(context.host.date1904)
  const text = value === undefined || value === null ? '' : String(value)
  switch (operator) {
    case 'isEmpty': return (_row, _value, _display, kind) => kind === 'blank'
    case 'isNotEmpty': return (_row, _value, _display, kind) => kind !== 'blank'
    case 'equals':
    case 'notEquals': {
      const negate = operator === 'notEquals'
      let test: RowTest
      if (text === '') test = (_row, _value, _display, kind) => kind === 'blank'
      else if (hasWildcards(text) || /~[*?~]/.test(text)) {
        const pattern = wildcardRegExp(text)
        test = (_row, _value, display, kind) => kind !== 'blank' && pattern.test(display)
      } else {
        const number = criterionNumber(value, date1904)
        const folded = foldText(text)
        test = (_row, cellValue, display, kind) => {
          if (kind === 'blank') return false
          if (number !== null && kind === 'number' && Math.abs((cellValue as number) - number) < 1e-9) return true
          return foldText(display) === folded || (kind !== 'number' && foldText(String(cellValue)) === folded)
        }
      }
      return negate ? (row, cellValue, display, kind) => !test(row, cellValue, display, kind) : test
    }
    case 'beginsWith':
    case 'notBeginsWith':
    case 'endsWith':
    case 'notEndsWith':
    case 'contains':
    case 'notContains': {
      if (!text) return null
      const pattern = wildcardRegExp(text, operator.endsWith('eginsWith'), operator.endsWith('ndsWith'))
      const negate = operator.startsWith('not')
      return negate
        ? (_row, _value, display, kind) => kind === 'blank' || !pattern.test(display)
        : (_row, _value, display, kind) => kind !== 'blank' && pattern.test(display)
    }
    case 'greaterThan':
    case 'greaterThanOrEqual':
    case 'lessThan':
    case 'lessThanOrEqual': {
      const number = criterionNumber(value, date1904)
      const compare = (difference: number) => operator === 'greaterThan' ? difference > 0
        : operator === 'greaterThanOrEqual' ? difference >= 0
          : operator === 'lessThan' ? difference < 0 : difference <= 0
      if (number !== null) return (_row, cellValue, _display, kind) => kind === 'number' && compare((cellValue as number) - number)
      if (!text) return null
      return (_row, cellValue, _display, kind) => kind === 'text' && compare(TEXT_COLLATOR.compare(String(cellValue), text))
    }
    case 'between':
    case 'notBetween': {
      const a = criterionNumber(value, date1904)
      const b = criterionNumber(value2, date1904)
      if (a === null || b === null) return null
      const low = Math.min(a, b)
      const high = Math.max(a, b)
      const inside: RowTest = (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) >= low && (cellValue as number) <= high
      return operator === 'between' ? inside : (row, cellValue, display, kind) => kind !== 'blank' && !inside(row, cellValue, display, kind)
    }
    case 'top':
    case 'bottom':
    case 'topPercent':
    case 'bottomPercent': {
      const { numbers } = columnStats(context)
      if (!numbers.length) return () => false
      const requested = criterionNumber(value, date1904) ?? 10
      const percent = operator.endsWith('Percent')
      const n = percent
        ? Math.max(1, Math.floor(numbers.length * Math.min(100, Math.max(0, requested)) / 100))
        : Math.max(1, Math.min(numbers.length, Math.floor(requested)))
      const top = operator.startsWith('top')
      const threshold = top ? numbers[numbers.length - n] : numbers[n - 1]
      return top
        ? (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) >= threshold
        : (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) <= threshold
    }
    case 'aboveAverage':
    case 'belowAverage': {
      const { mean } = columnStats(context)
      if (!Number.isFinite(mean)) return () => false
      return operator === 'aboveAverage'
        ? (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) > mean
        : (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) < mean
    }
    case 'dateEquals':
    case 'dateBefore':
    case 'dateAfter': {
      const day = criterionDate(value, date1904)
      if (day === null) return null
      if (operator === 'dateEquals') return (_row, cellValue, _display, kind) => kind === 'number' && Math.floor(cellValue as number) === day
      if (operator === 'dateBefore') return (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) < day
      return (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) >= day + 1
    }
    case 'dateBetween': {
      const a = criterionDate(value, date1904)
      const b = criterionDate(value2, date1904)
      if (a === null || b === null) return null
      const low = Math.min(a, b)
      const high = Math.max(a, b) + 1
      return (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) >= low && (cellValue as number) < high
    }
    case 'customFormula': {
      const formula = text.trim().replace(/^=/, '')
      const { host } = context
      if (!formula || !host.evaluate) return null
      return (row) => {
        const delta = row - context.top
        const shifted = delta && host.shiftFormula ? host.shiftFormula(formula, delta, 0) : formula
        const result = host.evaluate!(shifted, row, context.col)
        const scalar = Array.isArray(result) ? result[0]?.[0] ?? null : result
        return scalar === true || (typeof scalar === 'number' && scalar !== 0)
      }
    }
    default: {
      const quarter = /^Q([1-4])$/.exec(operator)
      if (quarter) {
        const q = Number(quarter[1])
        return (_row, cellValue, _display, kind) => kind === 'number' && Math.ceil(ymdFromSerial(cellValue as number, date1904).month / 3) === q
      }
      const month = /^M([1-9]|1[0-2])$/.exec(operator)
      if (month) {
        const m = Number(month[1])
        return (_row, cellValue, _display, kind) => kind === 'number' && ymdFromSerial(cellValue as number, date1904).month === m
      }
      const interval = datePeriodInterval(operator, context.todaySerial, date1904)
      if (interval) {
        const [start, end] = interval
        return (_row, cellValue, _display, kind) => kind === 'number' && (cellValue as number) >= start && (cellValue as number) < end
      }
      return null
    }
  }
}

function conditionTest(criteria: SheetFilterCriteria, context: ColumnContext): RowTest | null {
  const condition = criteria.condition
  if (!condition?.operator) return null
  const info = filterOperatorInfo(condition.operator)
  const first = buildTest(condition.operator, condition.value, condition.value2, context)
  const second = info?.combinable !== false && condition.operator2
    ? buildTest(condition.operator2, condition.value2, undefined, context)
    : null
  if (!first) return second
  if (!second) return first
  return condition.join === 'or'
    ? (row, value, display, kind) => first(row, value, display, kind) || second(row, value, display, kind)
    : (row, value, display, kind) => first(row, value, display, kind) && second(row, value, display, kind)
}

interface CompiledCriteria {
  needsDisplay: boolean
  test: (row: number, value: Scalar | undefined, display: string, kind: ScalarKind) => boolean
}

function compileCriteria(criteria: SheetFilterCriteria, context: ColumnContext): CompiledCriteria | null {
  const tests: RowTest[] = []
  let needsDisplay = false
  if (Array.isArray(criteria.values)) {
    const allowed = new Set(criteria.values.map((item) => foldText(String(item))))
    const blanks = criteria.blanks === true || allowed.has('')
    needsDisplay = true
    tests.push((_row, _value, display, kind) => kind === 'blank' ? blanks : allowed.has(foldText(display)))
  } else if (criteria.blanks === false) {
    tests.push((_row, _value, _display, kind) => kind !== 'blank')
  }
  const condition = conditionTest(criteria, context)
  if (condition) { tests.push(condition); needsDisplay = true }
  const { host, col } = context
  if (criteria.fillColor) {
    const wanted = criteria.fillColor === 'none' ? null : normalizeColor(criteria.fillColor)
    tests.push((row) => hostFillColor(host, row, col) === wanted)
  }
  if (criteria.fontColor) {
    const wanted = criteria.fontColor === 'none' || criteria.fontColor === 'auto' ? null : normalizeColor(criteria.fontColor)
    tests.push((row) => hostFontColor(host, row, col) === wanted)
  }
  if (!tests.length) return null
  if (tests.length === 1) return { needsDisplay, test: tests[0] }
  return { needsDisplay, test: (row, value, display, kind) => tests.every((test) => test(row, value, display, kind)) }
}

function cellKind(value: Scalar | undefined, display: string): ScalarKind {
  const kind = scalarKind(value)
  // A formula that yields "" and an empty-looking cell both count as (Blanks).
  if (kind === 'text' && display === '' && String(value).trim() === '') return 'blank'
  return kind
}

function todaySerialFor(host: DataHost) {
  return serialFromLocalDate(host.now ? host.now() : new Date(), Boolean(host.date1904))
}

function dataBounds(state: SheetFilterState): Bounds | null {
  const bounds = parseRange(state.ref)
  if (!bounds) return null
  return bounds
}

/**
 * Evaluate a column's criteria over the filter range's data rows (header excluded).
 * Returns a Uint8Array indexed by row - (top + 1): 1 = passes.
 */
export function evaluateColumnCriteria(state: SheetFilterState, column: number, criteria: SheetFilterCriteria, host: DataHost, todaySerial = todaySerialFor(host)): Uint8Array {
  const bounds = dataBounds(state)
  if (!bounds) return new Uint8Array(0)
  const top = bounds.top + 1
  const length = Math.max(0, bounds.bottom - top + 1)
  const pass = new Uint8Array(length).fill(1)
  const col = bounds.left + column
  const compiled = compileCriteria(criteria, { col, top, bottom: bounds.bottom, host, todaySerial, stats: null })
  if (!compiled) return pass
  for (let index = 0; index < length; index += 1) {
    const row = top + index
    const value = host.valueAt(row, col)
    const display = compiled.needsDisplay || typeof value === 'string' ? host.displayAt(row, col) : ''
    if (!compiled.test(row, value, display, cellKind(value, display))) pass[index] = 0
  }
  return pass
}

/**
 * Visibility of every data row given all criteria except `excludeColumn` (used by the checklist
 * so it lists values from rows the other columns' filters leave visible, like Excel and Sheets).
 */
export function computeRowVisibility(state: SheetFilterState, host: DataHost, excludeColumn?: number): { top: number; visible: Uint8Array } {
  const bounds = dataBounds(state)
  if (!bounds) return { top: 0, visible: new Uint8Array(0) }
  const top = bounds.top + 1
  const length = Math.max(0, bounds.bottom - top + 1)
  const visible = new Uint8Array(length).fill(1)
  const todaySerial = todaySerialFor(host)
  const width = bounds.right - bounds.left + 1
  for (const [key, criteria] of Object.entries(state.columns || {})) {
    const column = Number(key)
    if (!Number.isInteger(column) || column < 0 || column >= width || column === excludeColumn || !criteriaIsActive(criteria)) continue
    const col = bounds.left + column
    const compiled = compileCriteria(criteria, { col, top, bottom: bounds.bottom, host, todaySerial, stats: null })
    if (!compiled) continue
    for (let index = 0; index < length; index += 1) {
      if (!visible[index]) continue
      const row = top + index
      const value = host.valueAt(row, col)
      const display = compiled.needsDisplay || typeof value === 'string' ? host.displayAt(row, col) : ''
      if (!compiled.test(row, value, display, cellKind(value, display))) visible[index] = 0
    }
  }
  return { top, visible }
}

/** 1-based rows the filter hides (the header row is never hidden). */
export function computeFilteredRows(state: SheetFilterState, host: DataHost): number[] {
  const { top, visible } = computeRowVisibility(state, host)
  const hidden: number[] = []
  for (let index = 0; index < visible.length; index += 1) if (!visible[index]) hidden.push(top + index + 1)
  return hidden
}

// ---------------------------------------------------------------------------------------------
// Checklist values
// ---------------------------------------------------------------------------------------------

export type DistinctValueKind = 'number' | 'date' | 'text' | 'boolean' | 'error'

export interface DistinctValue {
  /** Case-folded display text: the identity used by `criteria.values`. */
  key: string
  /** Display text of the first occurrence. */
  text: string
  count: number
  kind: DistinctValueKind
  /** Numeric value for numbers/dates (for sorting and date grouping). */
  number?: number
  date?: { year: number; month: number; day: number }
  /** Whether the value passes this column's current criteria (checkbox state). */
  checked: boolean
}

export interface ColorCount { color: string | null; count: number }

export interface DistinctValuesResult {
  items: DistinctValue[]
  blanks: { count: number; checked: boolean } | null
  /** Data rows considered (visible under the other columns' filters). */
  rows: number
  /** Dominant data type, for choosing Text / Number / Date filter menus. */
  kind: 'text' | 'number' | 'date' | 'empty'
  fillColors: ColorCount[]
  fontColors: ColorCount[]
}

const KIND_ORDER: Record<DistinctValueKind, number> = { number: 0, date: 0, text: 1, boolean: 2, error: 3 }

/**
 * Distinct display values of `column` (0-based offset within the filter range) among rows
 * visible under the other columns' filters, sorted like Excel's checklist:
 * numbers/dates ascending, then text A→Z, then FALSE/TRUE, then errors; (Blanks) reported apart.
 */
export function distinctColumnValues(state: SheetFilterState, column: number, host: DataHost): DistinctValuesResult {
  const bounds = dataBounds(state)
  const empty: DistinctValuesResult = { items: [], blanks: null, rows: 0, kind: 'empty', fillColors: [], fontColors: [] }
  if (!bounds || column < 0 || column > bounds.right - bounds.left) return empty
  const { top, visible } = computeRowVisibility(state, host, column)
  const col = bounds.left + column
  const criteria = state.columns?.[column]
  const todaySerial = todaySerialFor(host)
  const compiled = criteria && criteriaIsActive(criteria)
    ? compileCriteria(criteria, { col, top, bottom: bounds.bottom, host, todaySerial, stats: null })
    : null
  const date1904 = Boolean(host.date1904)
  const map = new Map<string, DistinctValue>()
  const fills = new Map<string | null, number>()
  const fonts = new Map<string | null, number>()
  let blankCount = 0
  let blankChecked = false
  let rows = 0
  let numbers = 0
  let dates = 0
  let texts = 0
  for (let index = 0; index < visible.length; index += 1) {
    if (!visible[index]) continue
    rows += 1
    const row = top + index
    const value = host.valueAt(row, col)
    const display = host.displayAt(row, col)
    const kind = cellKind(value, display)
    const passes = compiled ? compiled.test(row, value, display, kind) : true
    const fill = hostFillColor(host, row, col)
    fills.set(fill, (fills.get(fill) || 0) + 1)
    const font = hostFontColor(host, row, col)
    fonts.set(font, (fonts.get(font) || 0) + 1)
    if (kind === 'blank') {
      blankCount += 1
      if (passes) blankChecked = true
      continue
    }
    const key = foldText(display)
    const existing = map.get(key)
    if (existing) {
      existing.count += 1
      if (passes) existing.checked = true
      continue
    }
    let itemKind: DistinctValueKind = kind === 'number' ? 'number' : kind === 'boolean' ? 'boolean' : kind === 'error' ? 'error' : 'text'
    const item: DistinctValue = { key, text: display, count: 1, kind: itemKind, checked: passes }
    if (kind === 'number') {
      item.number = value as number
      if (hostIsDate(host, row, col)) {
        itemKind = 'date'
        item.kind = 'date'
        const { year, month, day } = ymdFromSerial(value as number, date1904)
        item.date = { year, month, day }
      }
    } else if (kind === 'boolean') item.number = value ? 1 : 0
    if (itemKind === 'date') dates += 1
    else if (itemKind === 'number') numbers += 1
    else if (itemKind === 'text') texts += 1
    map.set(key, item)
  }
  const items = [...map.values()]
  const keys = collationKeys(items.map((item) => item.text))
  const order = items.map((_, index) => index)
  order.sort((a, b) => {
    const left = items[a]
    const right = items[b]
    const rank = KIND_ORDER[left.kind] - KIND_ORDER[right.kind]
    if (rank) return rank
    if (left.number !== undefined && right.number !== undefined && left.number !== right.number) return left.number - right.number
    return compareCollated(left.text, keys[a], right.text, keys[b])
  })
  const sortedItems = order.map((index) => items[index])
  const byCount = (entries: Map<string | null, number>) => [...entries.entries()]
    .map(([color, count]) => ({ color, count }))
    .sort((a, b) => b.count - a.count || String(a.color).localeCompare(String(b.color)))
  const dominant = dates + numbers + texts === 0 ? 'empty'
    : dates >= numbers && dates >= texts ? 'date' : numbers >= texts ? 'number' : 'text'
  return {
    items: sortedItems,
    blanks: blankCount ? { count: blankCount, checked: blankChecked } : null,
    rows,
    kind: dominant,
    fillColors: byCount(fills),
    fontColors: byCount(fonts),
  }
}

/**
 * Criteria for a checklist selection. Returns null when everything is selected (no filter).
 * `selectedKeys` are DistinctValue keys; `includeBlanks` is the (Blanks) checkbox.
 */
export function valuesCriteria(result: DistinctValuesResult, selectedKeys: Iterable<string>, includeBlanks: boolean): SheetFilterCriteria | null {
  const selected = new Set(selectedKeys)
  const values = result.items.filter((item) => selected.has(item.key)).map((item) => item.text)
  const allValues = values.length === result.items.length
  const allBlanks = !result.blanks || includeBlanks
  if (allValues && allBlanks) return null
  return { values, blanks: Boolean(result.blanks) && includeBlanks }
}

// ---------------------------------------------------------------------------------------------
// Range management
// ---------------------------------------------------------------------------------------------

type CellMap = Pick<SheetData, 'cells'>

function occupied(sheet: CellMap, row: number, col: number): boolean {
  if (row < 0 || col < 0 || row >= MAX_ROWS || col >= MAX_COLUMNS) return false
  return cellHasContent(sheet.cells[addressOf(row, col)] as CellData | undefined)
}

function rowHasContent(sheet: CellMap, row: number, left: number, right: number) {
  if (row < 0 || row >= MAX_ROWS) return false
  for (let col = Math.max(0, left); col <= Math.min(MAX_COLUMNS - 1, right); col += 1) if (occupied(sheet, row, col)) return true
  return false
}

/**
 * Excel's "current region" (Ctrl+Shift+8 / Ctrl+A): the rectangle around (row, col) bounded by
 * blank rows and columns. Returns null when the cell is empty and has no filled neighbour.
 */
export function detectCurrentRegion(sheet: CellMap, row: number, col: number): Bounds | null {
  const region = currentRegionAround(sheet, { top: row, bottom: row, left: col, right: col })
  if (region.top === region.bottom && region.left === region.right && !occupied(sheet, row, col)) return null
  return region
}

/**
 * The current region grown from a whole rectangle (a selection) rather than one cell: the
 * rectangle plus every block of data touching it, bounded by blank rows and columns.
 */
export function currentRegionAround(sheet: CellMap, start: Bounds): Bounds {
  const bounds: Bounds = { ...start }
  // Each side remembers the span of its outer line already found empty, so growing the region
  // in one direction never rescans the other sides' lines (linear in the region's perimeter).
  const memos: Record<'top' | 'bottom' | 'left' | 'right', { line: number; from: number; to: number } | null> = { top: null, bottom: null, left: null, right: null }
  const sideHasContent = (side: keyof typeof memos, line: number, from: number, to: number) => {
    const horizontal = side === 'top' || side === 'bottom'
    const limit = (horizontal ? MAX_COLUMNS : MAX_ROWS) - 1
    const low = Math.max(0, from)
    const high = Math.min(limit, to)
    const memo = memos[side] && memos[side]!.line === line ? memos[side]! : null
    const spans: Array<[number, number]> = memo ? [[low, Math.min(high, memo.from - 1)], [Math.max(low, memo.to + 1), high]] : [[low, high]]
    for (const [a, b] of spans) {
      for (let index = a; index <= b; index += 1) {
        if (horizontal ? occupied(sheet, line, index) : occupied(sheet, index, line)) return true
      }
    }
    memos[side] = { line, from: memo ? Math.min(low, memo.from) : low, to: memo ? Math.max(high, memo.to) : high }
    return false
  }
  let changed = true
  while (changed) {
    changed = false
    while (bounds.top > 0 && sideHasContent('top', bounds.top - 1, bounds.left - 1, bounds.right + 1)) { bounds.top -= 1; changed = true }
    while (bounds.bottom < MAX_ROWS - 1 && sideHasContent('bottom', bounds.bottom + 1, bounds.left - 1, bounds.right + 1)) { bounds.bottom += 1; changed = true }
    while (bounds.left > 0 && sideHasContent('left', bounds.left - 1, bounds.top - 1, bounds.bottom + 1)) { bounds.left -= 1; changed = true }
    while (bounds.right < MAX_COLUMNS - 1 && sideHasContent('right', bounds.right + 1, bounds.top - 1, bounds.bottom + 1)) { bounds.right += 1; changed = true }
  }
  return bounds
}

/**
 * The range Ctrl+Shift+L filters: a single cell expands to its current region; a one-row
 * selection (a header) extends down over the data beneath; a larger selection is used as-is.
 */
export function filterRangeForSelection(sheet: CellMap, selection: Bounds, active: { row: number; col: number } = { row: selection.top, col: selection.left }): Bounds | null {
  if (selection.top === selection.bottom && selection.left === selection.right) return detectCurrentRegion(sheet, active.row, active.col)
  if (selection.top === selection.bottom) {
    let bottom = selection.top
    while (bottom < MAX_ROWS - 1 && rowHasContent(sheet, bottom + 1, selection.left, selection.right)) bottom += 1
    return { ...selection, bottom }
  }
  return { ...selection }
}

export function createFilterState(bounds: Bounds): SheetFilterState {
  return { ref: formatRange(bounds), columns: {} }
}

/** The live filter for a sheet, falling back to an imported autoFilter range. */
export function filterStateFromSheet(sheet: Pick<SheetData, 'filter' | 'autoFilter'>): SheetFilterState | null {
  if (sheet.filter && parseRange(sheet.filter.ref)) return sheet.filter
  const bounds = autoFilterBounds(sheet.autoFilter)
  return bounds ? createFilterState(bounds) : null
}

export function setColumnCriteria(state: SheetFilterState, column: number, criteria: SheetFilterCriteria | null): SheetFilterState {
  const columns: Record<number, SheetFilterCriteria> = { ...state.columns }
  if (criteria && criteriaIsActive(criteria)) columns[column] = criteria
  else delete columns[column]
  return { ...state, columns }
}

export function clearFilterCriteria(state: SheetFilterState): SheetFilterState {
  return { ...state, columns: {} }
}

/** Grow the range over rows added directly beneath it (Excel re-applies over new data). */
export function extendFilterRange(state: SheetFilterState, sheet: CellMap): SheetFilterState {
  const bounds = parseRange(state.ref)
  if (!bounds) return state
  let bottom = bounds.bottom
  while (bottom < MAX_ROWS - 1 && rowHasContent(sheet, bottom + 1, bounds.left, bounds.right)) bottom += 1
  return bottom === bounds.bottom ? state : { ...state, ref: formatRange({ ...bounds, bottom }) }
}

/**
 * Merge a new filter result into the sheet's hidden rows. Rows the user hid by hand (hidden but
 * not previously filter-hidden) stay hidden and are never claimed by the filter.
 * Returns the new `hiddenRows` and `filteredRows` (1-based, ascending).
 */
export function applyFilterResult(sheet: Pick<SheetData, 'hiddenRows' | 'filteredRows'>, rowsToHide: number[]): { hiddenRows: number[]; filteredRows: number[] } {
  const previous = new Set(sheet.filteredRows || [])
  const manual = new Set<number>()
  for (const row of sheet.hiddenRows || []) if (!previous.has(row)) manual.add(row)
  const filteredRows: number[] = []
  for (const row of rowsToHide) if (!manual.has(row)) filteredRows.push(row)
  filteredRows.sort((a, b) => a - b)
  const hidden = new Set(manual)
  for (const row of filteredRows) hidden.add(row)
  return { hiddenRows: [...hidden].sort((a, b) => a - b), filteredRows }
}

/** Hidden rows after removing the filter (manually hidden rows are kept). */
export function hiddenRowsWithoutFilter(sheet: Pick<SheetData, 'hiddenRows' | 'filteredRows'>): number[] {
  const filtered = new Set(sheet.filteredRows || [])
  return (sheet.hiddenRows || []).filter((row) => !filtered.has(row))
}

/**
 * 0-based rows an active filter (the sheet's AutoFilter or a table's) currently hides. Excel's
 * range commands (Delete, Fill Down/Right, the fill handle, Ctrl+Enter, Copy, formatting, the
 * status-bar sums) act on visible cells only in a filtered list, so they skip these rows. Rows
 * the user hid by hand are not included: Excel still clears, fills and copies those.
 */
export function filterHiddenRowSet(sheet: Pick<SheetData, 'hiddenRows' | 'filteredRows'>): Set<number> {
  const rows = new Set<number>()
  if (!sheet.filteredRows?.length || !sheet.hiddenRows?.length) return rows
  const hidden = new Set(sheet.hiddenRows.map(Number))
  for (const value of sheet.filteredRows) {
    const row = Number(value)
    // A filtered row the user unhid by hand is visible again, so commands act on it.
    if (Number.isInteger(row) && row >= 1 && hidden.has(row)) rows.add(row - 1)
  }
  return rows
}

/**
 * Convenience for the app: compute and merge in one call. Returns the fields to assign on the
 * sheet (`filter`, `autoFilter`, `hiddenRows`, `filteredRows`).
 */
export function refreshSheetFilter(sheet: Pick<SheetData, 'hiddenRows' | 'filteredRows' | 'cells'>, state: SheetFilterState, host: DataHost, options: { extend?: boolean } = {}) {
  const next = options.extend ? extendFilterRange(state, sheet) : state
  const { hiddenRows, filteredRows } = applyFilterResult(sheet, computeFilteredRows(next, host))
  return { filter: next, autoFilter: next.ref, hiddenRows, filteredRows }
}

// ---------------------------------------------------------------------------------------------
// Structural edits
// ---------------------------------------------------------------------------------------------

function transformIndex(index: number, operation: SheetStructureOperation): number | null {
  if (operation.kind === 'insert') return index >= operation.index ? index + operation.count : index
  if (index < operation.index) return index
  if (index < operation.index + operation.count) return null
  return index - operation.count
}

/** Shift 1-based row numbers (e.g. `filteredRows`) for an inserted/deleted block of rows. */
export function adjustRowNumbersForStructure(rows: number[] | undefined, operation: SheetStructureOperation): number[] | undefined {
  if (!rows || operation.axis !== 'row') return rows
  const next: number[] = []
  for (const row of rows) {
    const transformed = transformIndex(row - 1, operation)
    if (transformed !== null) next.push(transformed + 1)
  }
  return next
}

/**
 * Keep the filter range and per-column criteria aligned when rows/columns are inserted or
 * deleted (0-based operation, as in sheet-operations). Returns null when the header row or
 * every filtered column is deleted.
 */
export function adjustFilterForStructure(state: SheetFilterState, operation: SheetStructureOperation): SheetFilterState | null {
  const bounds = parseRange(state.ref)
  if (!bounds || operation.count <= 0) return state
  const vertical = operation.axis === 'row'
  const start = vertical ? bounds.top : bounds.left
  const end = vertical ? bounds.bottom : bounds.right
  let nextStart = start
  let nextEnd = end
  if (operation.kind === 'insert') {
    if (operation.index <= start) { nextStart += operation.count; nextEnd += operation.count }
    else if (operation.index <= end) nextEnd += operation.count
  } else {
    const deleteEnd = operation.index + operation.count - 1
    if (deleteEnd < start) { nextStart -= operation.count; nextEnd -= operation.count }
    else if (operation.index <= end) {
      if (vertical && operation.index <= start) return null
      const removedInside = Math.min(end, deleteEnd) - Math.max(start, operation.index) + 1
      if (!vertical && removedInside >= end - start + 1) return null
      if (operation.index < start) nextStart = operation.index
      nextEnd = nextStart + (end - start + 1 - removedInside) - 1
    }
  }
  const nextBounds: Bounds = vertical
    ? { ...bounds, top: nextStart, bottom: nextEnd }
    : { ...bounds, left: nextStart, right: nextEnd }
  let columns = state.columns
  let sort = state.sort
  if (!vertical) {
    const remap = (offset: number) => {
      const transformed = transformIndex(start + offset, operation)
      return transformed === null ? null : transformed - nextStart
    }
    columns = {}
    for (const [key, criteria] of Object.entries(state.columns || {})) {
      const mapped = remap(Number(key))
      if (mapped !== null && mapped >= 0 && mapped <= nextEnd - nextStart) columns[mapped] = criteria
    }
    if (sort) {
      const mapped = remap(sort.column)
      sort = mapped === null ? undefined : { ...sort, column: mapped }
    }
  }
  const next: SheetFilterState = { ...state, ref: formatRange(nextBounds), columns }
  if (sort) next.sort = sort
  else delete next.sort
  return next
}

/** Label for a date checklist group, e.g. "January 2024". */
export function monthGroupLabel(year: number, month: number): string {
  return `${MONTH_NAMES[month - 1] || month} ${year}`
}

/** Display a serial as m/d/yyyy (used by the condition editor for date inputs). */
export function serialToDateInput(serial: number, date1904 = false): string {
  return formatSerialDate(serial, date1904)
}
