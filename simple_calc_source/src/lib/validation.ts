/**
 * Data validation (pure) over ExcelJS's validation model: checking entries, list sources
 * (literals, ranges, names), human descriptions, dialog <-> model conversion, range upkeep and
 * "Circle Invalid Data".
 */
import {
  addressOf,
  foldText,
  formatRange,
  formatSerialDate,
  formatSerialTime,
  parseDateText,
  parseNumberText,
  parseRange,
  parseTimeText,
  scalarKind,
  serialFromUtcDate,
  splitSheetPrefix,
  utcDateFromSerial,
} from './data-tools-core'
import type { Bounds, DataHost, Scalar } from './data-tools-core'

export type ValidationType = 'any' | 'whole' | 'decimal' | 'list' | 'date' | 'time' | 'textLength' | 'custom'
export type ValidationOperator = 'between' | 'notBetween' | 'equal' | 'notEqual' | 'greaterThan' | 'lessThan' | 'greaterThanOrEqual' | 'lessThanOrEqual'
export type ValidationErrorStyle = 'stop' | 'warning' | 'information'
/** How ExcelJS dates survive the renderer <-> main IPC boundary (see workbooks.cjs plainClone). */
export interface SerializedDate { type: 'date'; value: string }
export type ValidationFormula = string | number | Date | SerializedDate

/** ExcelJS's data validation model (the values of `SheetData.dataValidations`). */
export interface DataValidationModel {
  type: ValidationType
  operator?: ValidationOperator
  allowBlank?: boolean
  formulae?: ValidationFormula[]
  showInputMessage?: boolean
  promptTitle?: string
  prompt?: string
  showErrorMessage?: boolean
  errorStyle?: ValidationErrorStyle
  errorTitle?: string
  error?: string
  /** OOXML semantics: true HIDES the in-cell dropdown arrow of a list. */
  showDropDown?: boolean
}

export interface ValidationMessage { title: string; text: string; style: ValidationErrorStyle }

export interface ValidationOutcome {
  ok: boolean
  /** Alert to show when !ok. Absent when ok. */
  message?: ValidationMessage
  /**
   * Reaction when !ok: 'stop' rejects (Retry / Cancel), 'warning' asks "Continue?" (Yes / No /
   * Cancel), 'information' informs (OK / Cancel), 'none' accepts silently because the rule's
   * error alert is switched off (the cell still shows under Circle Invalid Data).
   */
  alert?: ValidationErrorStyle | 'none'
  /** The value to store when the entry was coerced (e.g. "9:30 AM" -> 0.395833 for time rules). */
  value?: Scalar
}

export interface ValidationTarget {
  /** Cell being validated (0-based). */
  row: number
  col: number
  /** Top-left of the validation's range: relative references in formulae are written for it. */
  anchor?: { row: number; col: number }
  /** Convert typed text (dates, times, numbers) before checking. True for edits, false for existing cells. */
  coerce?: boolean
}

export const VALIDATION_TYPES: Array<{ id: ValidationType; label: string }> = [
  { id: 'any', label: 'Any value' },
  { id: 'whole', label: 'Whole number' },
  { id: 'decimal', label: 'Decimal' },
  { id: 'list', label: 'List' },
  { id: 'date', label: 'Date' },
  { id: 'time', label: 'Time' },
  { id: 'textLength', label: 'Text length' },
  { id: 'custom', label: 'Custom' },
]

export const VALIDATION_OPERATORS: Array<{ id: ValidationOperator; label: string; inputs: 1 | 2 }> = [
  { id: 'between', label: 'between', inputs: 2 },
  { id: 'notBetween', label: 'not between', inputs: 2 },
  { id: 'equal', label: 'equal to', inputs: 1 },
  { id: 'notEqual', label: 'not equal to', inputs: 1 },
  { id: 'greaterThan', label: 'greater than', inputs: 1 },
  { id: 'lessThan', label: 'less than', inputs: 1 },
  { id: 'greaterThanOrEqual', label: 'greater than or equal to', inputs: 1 },
  { id: 'lessThanOrEqual', label: 'less than or equal to', inputs: 1 },
]

export const DEFAULT_VALIDATION_ERROR = "This value doesn't match the data validation restrictions defined for this cell."
export const MAX_LIST_OPTIONS = 1000
const MAX_REFERENCE_CELLS = 100_000

/** Narrow an unknown model value (as stored on the sheet) to a validation, or null. */
export function asValidation(value: unknown): DataValidationModel | null {
  if (!value || typeof value !== 'object') return null
  const model = value as DataValidationModel
  const type = typeof model.type === 'string' ? model.type : 'any'
  return VALIDATION_TYPES.some((item) => item.id === type) ? { ...model, type } : null
}

function operatorOf(validation: DataValidationModel): ValidationOperator {
  return validation.operator && VALIDATION_OPERATORS.some((item) => item.id === validation.operator) ? validation.operator : 'between'
}

function operatorInputs(operator: ValidationOperator) {
  return operator === 'between' || operator === 'notBetween' ? 2 : 1
}

function isSerializedDate(value: unknown): value is SerializedDate {
  return Boolean(value && typeof value === 'object' && (value as SerializedDate).type === 'date' && typeof (value as SerializedDate).value === 'string')
}

function formulaText(formula: ValidationFormula | undefined): string {
  if (formula === undefined || formula === null) return ''
  if (typeof formula === 'string') return formula.trim().replace(/^=/, '')
  if (typeof formula === 'number') return String(formula)
  return ''
}

function shiftedFormula(formula: string, host: DataHost, target: ValidationTarget): string {
  const anchor = target.anchor
  if (!anchor || !host.shiftFormula) return formula
  const rowDelta = target.row - anchor.row
  const colDelta = target.col - anchor.col
  return rowDelta || colDelta ? host.shiftFormula(formula, rowDelta, colDelta) : formula
}

function firstScalar(value: Scalar | Scalar[][] | undefined): Scalar {
  if (Array.isArray(value)) return value[0]?.[0] ?? null
  return value ?? null
}

/** Numeric bound for whole/decimal/date/time/textLength rules (serials for dates). */
function resolveBound(formula: ValidationFormula | undefined, type: ValidationType, host: DataHost, target: ValidationTarget): number | null {
  const date1904 = Boolean(host.date1904)
  if (formula === undefined || formula === null) return null
  if (typeof formula === 'number') return Number.isFinite(formula) ? formula : null
  if (formula instanceof Date) return Number.isNaN(formula.getTime()) ? null : serialFromUtcDate(formula, date1904)
  if (isSerializedDate(formula)) {
    const date = new Date(formula.value)
    return Number.isNaN(date.getTime()) ? null : serialFromUtcDate(date, date1904)
  }
  const text = formulaText(formula)
  if (!text) return null
  const number = parseNumberText(text)
  if (number !== null) return number
  if (type === 'date' || type === 'time') {
    const literal = type === 'time' ? parseTimeText(text) ?? parseDateText(text, date1904) : parseDateText(text, date1904) ?? parseTimeText(text)
    if (literal !== null) return literal
    const iso = /^\d{4}-\d{2}-\d{2}T/.test(text) ? new Date(text) : null
    if (iso && !Number.isNaN(iso.getTime())) return serialFromUtcDate(iso, date1904)
  }
  if (!host.evaluate) return null
  const result = firstScalar(host.evaluate(shiftedFormula(text, host, target), target.row, target.col))
  if (typeof result === 'number') return result
  if (typeof result === 'boolean') return result ? 1 : 0
  if (typeof result === 'string') return parseNumberText(result) ?? (type === 'date' ? parseDateText(result, date1904) : null)
  return null
}

function compare(operator: ValidationOperator, value: number, a: number, b: number | null): boolean {
  const epsilon = 1e-10
  switch (operator) {
    case 'between': return b !== null && value >= Math.min(a, b) - epsilon && value <= Math.max(a, b) + epsilon
    case 'notBetween': return b !== null && (value < Math.min(a, b) - epsilon || value > Math.max(a, b) + epsilon)
    case 'equal': return Math.abs(value - a) <= epsilon
    case 'notEqual': return Math.abs(value - a) > epsilon
    case 'greaterThan': return value > a + epsilon
    case 'lessThan': return value < a - epsilon
    case 'greaterThanOrEqual': return value >= a - epsilon
    case 'lessThanOrEqual': return value <= a + epsilon
    default: return false
  }
}

// ---------------------------------------------------------------------------------------------
// List sources
// ---------------------------------------------------------------------------------------------

/** Parse a literal list formula ('"a,b,c"'); null when the source is a reference. */
export function literalListItems(formula: unknown): string[] | null {
  if (typeof formula !== 'string') return null
  const text = formula.trim()
  if (!(text.length >= 2 && text.startsWith('"') && text.endsWith('"'))) return null
  return text.slice(1, -1).replace(/""/g, '"').split(',').map((item) => item.trim()).filter((item) => item.length > 0)
}

/** Every entry of a list rule's source, in order (literal items, or the referenced cells). */
export function listSourceEntries(validation: DataValidationModel, host: DataHost, target: ValidationTarget = { row: 0, col: 0 }): Array<{ value: Scalar; text: string }> {
  if (validation.type !== 'list') return []
  const source = validation.formulae?.[0]
  const literal = literalListItems(source)
  if (literal) return literal.map((text) => ({ value: parseNumberText(text) ?? text, text }))
  const reference = formulaText(source as ValidationFormula)
  if (!reference) return []
  const shifted = shiftedFormula(reference, host, target)
  if (host.resolveReference) {
    const resolved = host.resolveReference(shifted, { row: target.row, col: target.col })
    if (resolved) return resolved
  }
  const { sheet, ref } = splitSheetPrefix(shifted)
  const bounds = sheet === null ? parseRange(ref) : null
  if (!bounds) return []
  const entries: Array<{ value: Scalar; text: string }> = []
  const cells = (bounds.bottom - bounds.top + 1) * (bounds.right - bounds.left + 1)
  if (cells > MAX_REFERENCE_CELLS) return []
  for (let row = bounds.top; row <= bounds.bottom; row += 1) {
    for (let col = bounds.left; col <= bounds.right; col += 1) {
      const value = host.valueAt(row, col)
      entries.push({ value: value ?? null, text: host.displayAt(row, col) })
    }
  }
  return entries
}

/**
 * Dropdown choices for a list rule: literal items or referenced/named cells, blanks skipped,
 * de-duplicated (case-insensitively, first spelling wins), at most 1000.
 */
export function listOptionsForValidation(validation: unknown, host: DataHost, target?: ValidationTarget): string[] {
  const model = asValidation(validation)
  if (!model || model.type !== 'list') return []
  const seen = new Set<string>()
  const options: string[] = []
  for (const entry of listSourceEntries(model, host, target)) {
    const text = entry.text.trim() === '' && scalarKind(entry.value) === 'blank' ? '' : entry.text
    if (!text) continue
    const key = foldText(text)
    if (seen.has(key)) continue
    seen.add(key)
    options.push(text)
    if (options.length >= MAX_LIST_OPTIONS) break
  }
  return options
}

export function hasInCellDropdown(validation: unknown): boolean {
  const model = asValidation(validation)
  return Boolean(model && model.type === 'list' && model.showDropDown !== true)
}

// ---------------------------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------------------------

function isBlankInput(raw: string | undefined, value: Scalar | undefined) {
  return (raw === undefined || raw === '') && scalarKind(value) === 'blank'
}

function failure(validation: DataValidationModel, host?: DataHost): ValidationOutcome {
  const style: ValidationErrorStyle = validation.errorStyle === 'warning' || validation.errorStyle === 'information' ? validation.errorStyle : 'stop'
  const enforced = validation.showErrorMessage === true
  return {
    ok: false,
    alert: enforced ? style : 'none',
    message: {
      title: validation.errorTitle?.trim() || (style === 'stop' ? 'Invalid entry' : style === 'warning' ? 'Check this entry' : 'Please note'),
      text: validation.error?.trim() || requirementText(validation, host) || DEFAULT_VALIDATION_ERROR,
      style,
    },
  }
}

/** Sentence describing what the rule accepts, e.g. "Enter a whole number between 1 and 10." */
export function requirementText(validation: DataValidationModel, host?: DataHost): string {
  const description = describeValidation(validation, host)
  switch (validation.type) {
    case 'any': return ''
    case 'list': return 'Choose a value from the list.'
    case 'custom': return DEFAULT_VALIDATION_ERROR
    default: return `Enter ${description.charAt(0).toLocaleLowerCase()}${description.slice(1)}.`.replace(/^Enter text length/, 'Enter text with a length')
  }
}

/**
 * Check an entry against a rule. `rawInput` is the text typed (or undefined for a programmatic
 * value); `parsedValue` is what the app would store. Blank entries pass when `allowBlank`.
 */
export function validateValue(validation: unknown, rawInput: string | undefined, parsedValue: Scalar | undefined, host: DataHost, target: ValidationTarget = { row: 0, col: 0 }): ValidationOutcome {
  const model = asValidation(validation)
  if (!model || model.type === 'any') return { ok: true }
  if (isBlankInput(rawInput, parsedValue)) return model.allowBlank ? { ok: true } : failure(model, host)
  const coerce = target.coerce !== false
  const date1904 = Boolean(host.date1904)
  const raw = rawInput ?? (parsedValue === null || parsedValue === undefined ? '' : typeof parsedValue === 'boolean' ? (parsedValue ? 'TRUE' : 'FALSE') : String(parsedValue))
  const operator = operatorOf(model)
  const bounds = () => {
    const a = resolveBound(model.formulae?.[0], model.type, host, target)
    const b = operatorInputs(operator) === 2 ? resolveBound(model.formulae?.[1], model.type, host, target) : null
    return { a, b }
  }
  const numeric = (value: number, coerced?: Scalar): ValidationOutcome => {
    const { a, b } = bounds()
    if (a === null || (operatorInputs(operator) === 2 && b === null)) return failure(model, host)
    if (!compare(operator, value, a, b)) return failure(model, host)
    return coerced === undefined ? { ok: true } : { ok: true, value: coerced }
  }
  switch (model.type) {
    case 'whole':
    case 'decimal': {
      let value = typeof parsedValue === 'number' ? parsedValue : null
      let coerced: Scalar | undefined
      if (value === null && coerce && typeof parsedValue === 'string') {
        value = parseNumberText(parsedValue)
        if (value !== null) coerced = value
      }
      if (value === null || !Number.isFinite(value)) return failure(model, host)
      if (model.type === 'whole' && Math.abs(value - Math.round(value)) > 1e-9) return failure(model, host)
      return numeric(value, coerced)
    }
    case 'date': {
      let value = typeof parsedValue === 'number' ? parsedValue : null
      let coerced: Scalar | undefined
      if (value === null && coerce && typeof parsedValue === 'string') {
        value = parseDateText(parsedValue, date1904) ?? parseNumberText(parsedValue)
        if (value !== null) coerced = value
      }
      if (value === null || !Number.isFinite(value)) return failure(model, host)
      return numeric(value, coerced)
    }
    case 'time': {
      let value = typeof parsedValue === 'number' ? parsedValue : null
      let coerced: Scalar | undefined
      if (value === null && coerce && typeof parsedValue === 'string') {
        value = parseTimeText(parsedValue) ?? parseDateText(parsedValue, date1904) ?? parseNumberText(parsedValue)
        if (value !== null) coerced = value
      }
      if (value === null || !Number.isFinite(value) || value < 0) return failure(model, host)
      const { a, b } = bounds()
      // A date-time is compared on its time of day when the rule's bounds are times.
      const compared = value >= 1 && (a === null || a < 1) && (b === null || b < 1) ? value % 1 : value
      return numeric(compared, coerced)
    }
    case 'textLength': {
      const text = typeof parsedValue === 'string' ? parsedValue : raw
      return numeric(text.length)
    }
    case 'list': {
      const entries = listSourceEntries(model, host, target)
      const valueKind = scalarKind(parsedValue)
      const folded = foldText(valueKind === 'text' ? String(parsedValue) : raw)
      const number = valueKind === 'number' ? parsedValue as number : null
      for (const entry of entries) {
        if (number !== null && typeof entry.value === 'number' && Math.abs(entry.value - number) < 1e-9) return { ok: true }
        if (foldText(entry.text) === folded) return { ok: true }
        if (typeof entry.value === 'string' && foldText(entry.value) === folded) return { ok: true }
        if (typeof entry.value === 'boolean' && folded === (entry.value ? 'true' : 'false')) return { ok: true }
      }
      return failure(model, host)
    }
    case 'custom': {
      const formula = formulaText(model.formulae?.[0])
      if (!formula || !host.evaluate) return { ok: true }
      const result = firstScalar(host.evaluate(shiftedFormula(formula, host, target), target.row, target.col, { row: target.row, col: target.col, value: parsedValue ?? null }))
      const ok = result === true || (typeof result === 'number' && result !== 0) || (typeof result === 'string' && foldText(result) === 'true')
      return ok ? { ok: true } : failure(model, host)
    }
    default:
      return { ok: true }
  }
}

// ---------------------------------------------------------------------------------------------
// Descriptions
// ---------------------------------------------------------------------------------------------

function boundText(formula: ValidationFormula | undefined, type: ValidationType, host?: DataHost): string {
  const date1904 = Boolean(host?.date1904)
  if (formula === undefined || formula === null) return '?'
  if (formula instanceof Date || isSerializedDate(formula)) {
    const date = formula instanceof Date ? formula : new Date(formula.value)
    return Number.isNaN(date.getTime()) ? '?' : formatSerialDate(serialFromUtcDate(date, date1904), date1904)
  }
  if (typeof formula === 'number') {
    if (type === 'date') return formatSerialDate(formula, date1904)
    if (type === 'time') return formatSerialTime(formula)
    return formula.toLocaleString(undefined, { maximumFractionDigits: 10 })
  }
  const text = formulaText(formula)
  const number = parseNumberText(text)
  if (number !== null) {
    if (type === 'time') return formatSerialTime(number)
    if (type === 'date') return formatSerialDate(number, date1904)
    return text
  }
  if (type === 'date' && /^\d{4}-\d{2}-\d{2}T/.test(text)) {
    const date = new Date(text)
    if (!Number.isNaN(date.getTime())) return formatSerialDate(serialFromUtcDate(date, date1904), date1904)
  }
  return `=${text}`
}

const TYPE_NOUNS: Record<ValidationType, string> = {
  any: 'Any value', whole: 'A whole number', decimal: 'A decimal number', list: 'A value from the list',
  date: 'A date', time: 'A time', textLength: 'Text length', custom: 'Custom formula',
}

/** Short description, e.g. "A whole number between 1 and 10", "List: Yes, No", "List from =$A$1:$A$9". */
export function describeValidation(validation: unknown, host?: DataHost): string {
  const model = asValidation(validation)
  if (!model) return 'No validation'
  switch (model.type) {
    case 'any': return 'Any value'
    case 'list': {
      const literal = literalListItems(model.formulae?.[0])
      if (literal) {
        const shown = literal.slice(0, 6).join(', ')
        return `List: ${shown}${literal.length > 6 ? `, … (${literal.length} items)` : ''}`
      }
      return `List from =${formulaText(model.formulae?.[0])}`
    }
    case 'custom': return `Custom formula =${formulaText(model.formulae?.[0])}`
    default: {
      const operator = operatorOf(model)
      const a = boundText(model.formulae?.[0], model.type, host)
      const label = VALIDATION_OPERATORS.find((item) => item.id === operator)?.label || operator
      const noun = model.type === 'date' && (operator === 'greaterThan' || operator === 'lessThan')
        ? `A date ${operator === 'greaterThan' ? 'after' : 'before'}`
        : `${TYPE_NOUNS[model.type]} ${label}`
      if (model.type === 'date' && (operator === 'greaterThan' || operator === 'lessThan')) return `${noun} ${a}`
      if (operatorInputs(operator) === 2) return `${noun} ${a} and ${boundText(model.formulae?.[1], model.type, host)}`
      return `${noun} ${a}`
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Dialog state
// ---------------------------------------------------------------------------------------------

export interface ValidationDialogState {
  type: ValidationType
  operator: ValidationOperator
  /** Minimum / value / list source / custom formula, as typed. */
  value1: string
  /** Maximum (between / not between). */
  value2: string
  allowBlank: boolean
  inCellDropdown: boolean
  showInputMessage: boolean
  promptTitle: string
  prompt: string
  showErrorMessage: boolean
  errorStyle: ValidationErrorStyle
  errorTitle: string
  error: string
}

export const DEFAULT_VALIDATION_DIALOG_STATE: ValidationDialogState = {
  type: 'any',
  operator: 'between',
  value1: '',
  value2: '',
  allowBlank: true,
  inCellDropdown: true,
  showInputMessage: true,
  promptTitle: '',
  prompt: '',
  showErrorMessage: true,
  errorStyle: 'stop',
  errorTitle: '',
  error: '',
}

function inputText(formula: ValidationFormula | undefined, type: ValidationType, date1904: boolean): string {
  if (formula === undefined || formula === null) return ''
  if (formula instanceof Date || isSerializedDate(formula)) {
    const date = formula instanceof Date ? formula : new Date(formula.value)
    return Number.isNaN(date.getTime()) ? '' : formatSerialDate(serialFromUtcDate(date, date1904), date1904)
  }
  const text = typeof formula === 'number' ? String(formula) : formulaText(formula)
  const number = parseNumberText(text)
  if (number !== null) {
    if (type === 'date') return formatSerialDate(number, date1904)
    if (type === 'time') return formatSerialTime(number)
    return text
  }
  if (type === 'date' && /^\d{4}-\d{2}-\d{2}T/.test(text)) {
    const date = new Date(text)
    if (!Number.isNaN(date.getTime())) return formatSerialDate(serialFromUtcDate(date, date1904), date1904)
  }
  return text ? `=${text}` : ''
}

/** Dialog fields for an existing rule (or Excel's defaults for none). */
export function validationToDialogState(validation: unknown, options: { date1904?: boolean } = {}): ValidationDialogState {
  const model = asValidation(validation)
  if (!model) return { ...DEFAULT_VALIDATION_DIALOG_STATE }
  const date1904 = Boolean(options.date1904)
  let value1 = ''
  let value2 = ''
  if (model.type === 'list') {
    const literal = literalListItems(model.formulae?.[0])
    value1 = literal ? literal.join(',') : formulaText(model.formulae?.[0]) ? `=${formulaText(model.formulae?.[0])}` : ''
  } else if (model.type === 'custom') {
    value1 = formulaText(model.formulae?.[0]) ? `=${formulaText(model.formulae?.[0])}` : ''
  } else if (model.type !== 'any') {
    value1 = inputText(model.formulae?.[0], model.type, date1904)
    value2 = inputText(model.formulae?.[1], model.type, date1904)
  }
  return {
    type: model.type,
    operator: operatorOf(model),
    value1,
    value2,
    allowBlank: model.allowBlank !== false,
    inCellDropdown: model.showDropDown !== true,
    showInputMessage: model.showInputMessage !== false,
    promptTitle: model.promptTitle || '',
    prompt: model.prompt || '',
    showErrorMessage: model.showErrorMessage !== false,
    errorStyle: model.errorStyle === 'warning' || model.errorStyle === 'information' ? model.errorStyle : 'stop',
    errorTitle: model.errorTitle || '',
    error: model.error || '',
  }
}

export type DialogConversion =
  | { ok: true; validation: DataValidationModel }
  | { ok: false; error: string; field: 'value1' | 'value2' }

function serializedDate(serial: number, date1904: boolean): SerializedDate {
  return { type: 'date', value: utcDateFromSerial(serial, date1904).toISOString() }
}

/**
 * Build the ExcelJS model from dialog fields, with Excel's input checks. Date bounds are stored
 * as `{ type: 'date', value: ISO }` (how dates cross the IPC boundary); times as day fractions.
 */
export function dialogStateToValidation(state: ValidationDialogState, options: { date1904?: boolean } = {}): DialogConversion {
  const date1904 = Boolean(options.date1904)
  const validation: DataValidationModel = { type: state.type }
  const needsTwo = operatorInputs(state.operator) === 2
  const fieldLabel = (field: 'value1' | 'value2') => needsTwo ? (field === 'value1' ? 'Minimum' : 'Maximum') : 'Value'
  const parseBound = (field: 'value1' | 'value2'): { ok: true; formula: ValidationFormula; number: number | null } | { ok: false; error: string; field: 'value1' | 'value2' } => {
    const text = state[field].trim()
    if (!text) return { ok: false, error: `Enter a ${fieldLabel(field).toLocaleLowerCase()}.`, field }
    if (text.startsWith('=')) {
      const formula = text.slice(1).trim()
      if (!formula) return { ok: false, error: `Enter a ${fieldLabel(field).toLocaleLowerCase()}.`, field }
      return { ok: true, formula, number: null }
    }
    switch (state.type) {
      case 'whole':
      case 'textLength': {
        const number = parseNumberText(text)
        if (number === null || !Number.isInteger(number)) return { ok: false, error: `The ${fieldLabel(field)} must be a whole number.`, field }
        if (state.type === 'textLength' && number < 0) return { ok: false, error: `The ${fieldLabel(field)} must be zero or more.`, field }
        return { ok: true, formula: number, number }
      }
      case 'decimal': {
        const number = parseNumberText(text)
        if (number === null) return { ok: false, error: `The ${fieldLabel(field)} must be a number.`, field }
        return { ok: true, formula: number, number }
      }
      case 'date': {
        const serial = parseDateText(text, date1904)
        if (serial === null) return { ok: false, error: `The ${fieldLabel(field)} must be a date, such as 1/31/2025.`, field }
        return { ok: true, formula: serializedDate(serial, date1904), number: serial }
      }
      case 'time': {
        const fraction = parseTimeText(text) ?? parseNumberText(text)
        if (fraction === null || fraction < 0 || fraction >= 1) return { ok: false, error: `The ${fieldLabel(field)} must be a time, such as 9:30 AM.`, field }
        const rounded = Math.round(fraction * 1e10) / 1e10
        return { ok: true, formula: String(rounded), number: rounded }
      }
      default:
        return { ok: true, formula: text, number: null }
    }
  }
  switch (state.type) {
    case 'any':
      break
    case 'list': {
      const source = state.value1.trim()
      if (!source) return { ok: false, error: 'Enter the list source: items separated by commas, or a reference such as =$A$1:$A$10.', field: 'value1' }
      if (source.startsWith('=')) {
        const reference = source.slice(1).trim()
        if (!reference) return { ok: false, error: 'Enter a reference after "=".', field: 'value1' }
        const { ref } = splitSheetPrefix(reference)
        const bounds = parseRange(ref)
        if (bounds && bounds.top !== bounds.bottom && bounds.left !== bounds.right) {
          return { ok: false, error: 'The list source must be a delimited list, or a reference to a single row or column.', field: 'value1' }
        }
        validation.formulae = [reference]
      } else {
        const items = source.split(',').map((item) => item.trim()).filter(Boolean)
        if (!items.length) return { ok: false, error: 'Enter at least one list item.', field: 'value1' }
        const literal = items.join(',')
        if (literal.length > 255) return { ok: false, error: 'A typed list can hold at most 255 characters. Put longer lists in cells and refer to them (=$A$1:$A$50).', field: 'value1' }
        validation.formulae = [`"${literal.replace(/"/g, '""')}"`]
      }
      if (!state.inCellDropdown) validation.showDropDown = true
      break
    }
    case 'custom': {
      const formula = state.value1.trim().replace(/^=/, '').trim()
      if (!formula) return { ok: false, error: 'Enter a formula that returns TRUE for valid entries, such as =ISNUMBER(A1).', field: 'value1' }
      validation.formulae = [formula]
      break
    }
    default: {
      validation.operator = state.operator
      const first = parseBound('value1')
      if (!first.ok) return first
      const formulae: ValidationFormula[] = [first.formula]
      if (needsTwo) {
        const second = parseBound('value2')
        if (!second.ok) return second
        if (first.number !== null && second.number !== null && second.number < first.number) {
          return { ok: false, error: 'The Maximum must be greater than or equal to the Minimum.', field: 'value2' }
        }
        formulae.push(second.formula)
      }
      validation.formulae = formulae
    }
  }
  if (state.type !== 'any') validation.allowBlank = state.allowBlank
  validation.showInputMessage = state.showInputMessage
  if (state.promptTitle.trim()) validation.promptTitle = state.promptTitle.trim().slice(0, 32)
  if (state.prompt.trim()) validation.prompt = state.prompt.slice(0, 255)
  validation.showErrorMessage = state.showErrorMessage
  if (state.errorStyle !== 'stop') validation.errorStyle = state.errorStyle
  if (state.errorTitle.trim()) validation.errorTitle = state.errorTitle.trim().slice(0, 32)
  if (state.error.trim()) validation.error = state.error.slice(0, 255)
  return { ok: true, validation }
}

// ---------------------------------------------------------------------------------------------
// Rule ranges
// ---------------------------------------------------------------------------------------------

interface IndexedRule { key: string; bounds: Bounds; validation: DataValidationModel }
const ruleIndexCache = new WeakMap<object, IndexedRule[]>()

function indexedRules(validations: Record<string, unknown>): IndexedRule[] {
  const cached = ruleIndexCache.get(validations)
  if (cached) return cached
  const rules: IndexedRule[] = []
  for (const [key, value] of Object.entries(validations)) {
    const validation = asValidation(value)
    if (!validation) continue
    for (const part of key.trim().split(/\s+/)) {
      const bounds = parseRange(part)
      if (bounds) rules.push({ key, bounds, validation })
    }
  }
  ruleIndexCache.set(validations, rules)
  return rules
}

/** The rule covering a cell, with the anchor (top-left) its relative formulae are written for. */
export function findValidation(validations: Record<string, unknown> | undefined, row: number, col: number): { key: string; validation: DataValidationModel; anchor: { row: number; col: number } } | null {
  if (!validations) return null
  const direct = asValidation(validations[addressOf(row, col)])
  if (direct) return { key: addressOf(row, col), validation: direct, anchor: { row, col } }
  for (const rule of indexedRules(validations)) {
    const { bounds } = rule
    if (row >= bounds.top && row <= bounds.bottom && col >= bounds.left && col <= bounds.right) {
      return { key: rule.key, validation: rule.validation, anchor: { row: bounds.top, col: bounds.left } }
    }
  }
  return null
}

function stable(value: unknown): unknown {
  if (value instanceof Date) return { type: 'date', value: value.toISOString() }
  if (Array.isArray(value)) return value.map(stable)
  if (value && typeof value === 'object') {
    const entries = Object.entries(value as Record<string, unknown>)
      .filter(([, item]) => item !== undefined && item !== false && item !== '')
      .sort(([a], [b]) => a.localeCompare(b))
    return Object.fromEntries(entries.map(([key, item]) => [key, stable(item)]))
  }
  return value
}

export function validationsEqual(a: unknown, b: unknown): boolean {
  return JSON.stringify(stable(a)) === JSON.stringify(stable(b))
}

/** Keys of every range whose rule matches `validation` ("Apply these changes to all other cells with the same settings"). */
export function rangesWithSameValidation(validations: Record<string, unknown> | undefined, validation: unknown): string[] {
  if (!validations || !validation) return []
  return Object.keys(validations).filter((key) => validationsEqual(validations[key], validation))
}

/**
 * Set (or clear, with null) the rule for a rectangle, splitting any overlapped rules so the
 * uncovered parts keep theirs. Returns a new record.
 */
export function withValidation(validations: Record<string, unknown> | undefined, target: Bounds, validation: DataValidationModel | null): Record<string, unknown> {
  const next: Record<string, unknown> = {}
  for (const [key, existing] of Object.entries(validations || {})) {
    const parts = key.trim().split(/\s+/)
    const kept: string[] = []
    for (const part of parts) {
      const bounds = parseRange(part)
      if (!bounds || bounds.bottom < target.top || bounds.top > target.bottom || bounds.right < target.left || bounds.left > target.right) {
        kept.push(part)
        continue
      }
      const pieces: Bounds[] = []
      if (bounds.top < target.top) pieces.push({ top: bounds.top, bottom: target.top - 1, left: bounds.left, right: bounds.right })
      if (bounds.bottom > target.bottom) pieces.push({ top: target.bottom + 1, bottom: bounds.bottom, left: bounds.left, right: bounds.right })
      const top = Math.max(bounds.top, target.top)
      const bottom = Math.min(bounds.bottom, target.bottom)
      if (bounds.left < target.left) pieces.push({ top, bottom, left: bounds.left, right: target.left - 1 })
      if (bounds.right > target.right) pieces.push({ top, bottom, left: target.right + 1, right: bounds.right })
      for (const piece of pieces) kept.push(formatRange(piece))
    }
    for (const part of kept) next[part] = existing
  }
  if (validation) next[formatRange(target)] = validation
  return next
}

/**
 * Addresses of non-blank cells that break their rule (Excel's "Circle Invalid Data"). Whole-
 * column rules are clipped to `extent` (the used range); `within` limits the scan.
 */
export function findInvalidCells(
  validations: Record<string, unknown> | undefined,
  host: DataHost,
  options: { extent: { rows: number; cols: number }; within?: Bounds; limit?: number },
): string[] {
  if (!validations) return []
  const limit = options.limit ?? 10_000
  const invalid: Array<{ row: number; col: number }> = []
  const seen = new Set<string>()
  scan: for (const rule of indexedRules(validations)) {
    if (rule.validation.type === 'any') continue
    const area: Bounds = {
      top: Math.max(rule.bounds.top, options.within?.top ?? 0),
      bottom: Math.min(rule.bounds.bottom, options.within?.bottom ?? Infinity, options.extent.rows - 1),
      left: Math.max(rule.bounds.left, options.within?.left ?? 0),
      right: Math.min(rule.bounds.right, options.within?.right ?? Infinity, options.extent.cols - 1),
    }
    const anchor = { row: rule.bounds.top, col: rule.bounds.left }
    for (let row = area.top; row <= area.bottom; row += 1) {
      for (let col = area.left; col <= area.right; col += 1) {
        const value = host.valueAt(row, col)
        if (scalarKind(value) === 'blank') continue
        const address = addressOf(row, col)
        if (seen.has(address)) continue
        seen.add(address)
        const outcome = validateValue(rule.validation, undefined, value ?? null, host, { row, col, anchor, coerce: false })
        if (!outcome.ok) {
          invalid.push({ row, col })
          if (invalid.length >= limit) break scan
        }
      }
    }
  }
  return invalid
    .sort((a, b) => a.row - b.row || a.col - b.col)
    .slice(0, limit)
    .map((cell) => addressOf(cell.row, cell.col))
}
