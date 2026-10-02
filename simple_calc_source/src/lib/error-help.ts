/**
 * Plain-language help for formula error values, for the hover tooltip and the card shown next
 * to a selected error cell (Excel's error smart tag, Sheets' error card).
 */
import { describeFormulaError, isFormulaError } from './formulas'

const ERROR_TITLES: Record<string, string> = {
  '#DIV/0!': 'Divide by zero',
  '#N/A': 'Value not available',
  '#NAME?': 'Unknown name',
  '#NULL!': 'Ranges don’t intersect',
  '#NUM!': 'Invalid number',
  '#REF!': 'Invalid cell reference',
  '#VALUE!': 'Wrong type of value',
  '#SPILL!': 'Spill range isn’t empty',
  '#CALC!': 'Calculation error',
  '#CIRC!': 'Circular reference',
  '#PARSE!': 'Formula parse error',
  '#ERROR!': 'Formula error',
  '#GETTING_DATA': 'Getting data',
}

export interface FormulaErrorHelp {
  /** The error value, e.g. "#DIV/0!". */
  code: string
  /** Short name of the error ("Divide by zero"). */
  title: string
  /** One-sentence explanation; `detail` (a diagnostic) replaces it when given. */
  text: string
}

/** Help for an error value, or null when `value` is not an error. */
export function formulaErrorHelp(value: unknown, detail?: string | null): FormulaErrorHelp | null {
  if (typeof value !== 'string' || !isFormulaError(value)) return null
  const code = value.toUpperCase()
  const text = (detail && detail.trim()) || describeFormulaError(code) || 'The formula can’t be calculated.'
  return { code, title: ERROR_TITLES[code] || 'Formula error', text }
}

/** "#DIV/0! – Divide by zero. The formula divides by zero or by an empty cell." for a tooltip. */
export function formulaErrorTooltip(help: FormulaErrorHelp): string {
  return `${help.code} – ${help.title}. ${help.text}`
}
