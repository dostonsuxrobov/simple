/**
 * Editor model for conditional-format rules: a flat, form-friendly draft that converts to and
 * from the ExcelJS rule model, a catalog of conditions, presets and human-readable rule
 * descriptions. Pure and framework-free (used by ConditionalFormatPanel and the QA script).
 */
import {
  CONDITIONAL_STYLE_PRESETS,
  conditionalPresetStyle,
  conditionalRuleKind,
  createAverageRule,
  createBlankRule,
  createCellIsRule,
  createColorScaleRule,
  createConditionalStyle,
  createDataBarRule,
  createDateRule,
  createDuplicateRule,
  createErrorRule,
  createFormulaRule,
  createIconSetRule,
  createTextRule,
  createTopBottomRule,
  defaultConditionalCssColor,
  defaultIconThresholds,
  describeConditionalStyle,
  iconSetSize,
  matchConditionalStylePreset,
  normalizeIconSetName,
  parseRangeInput,
} from './conditional-format'
import type {
  CellIsOperator,
  CfvoType,
  ConditionalRule,
  ConditionalStylePresetId,
  IconSetName,
  TimePeriod,
} from './conditional-format'
import type { CellStyle } from '../spreadsheet-types'

export type RuleFormatKind = 'single' | 'colorScale' | 'dataBar' | 'iconSet'

export type ConditionId =
  | `cellIs:${CellIsOperator}`
  | 'text:containsText' | 'text:notContainsText' | 'text:beginsWith' | 'text:endsWith'
  | `date:${TimePeriod}`
  | 'top:top' | 'top:topPercent' | 'top:bottom' | 'top:bottomPercent'
  | 'avg:above' | 'avg:below' | 'avg:equalOrAbove' | 'avg:equalOrBelow' | 'avg:stdDevAbove' | 'avg:stdDevBelow'
  | 'dup:duplicate' | 'dup:unique'
  | 'blank:blanks' | 'blank:noBlanks' | 'error:errors' | 'error:noErrors'
  | 'formula'

export type ConditionInput = 'none' | 'value' | 'values' | 'text' | 'rank' | 'percentRank' | 'stdDev' | 'formula'

export interface ConditionOption {
  id: ConditionId
  label: string
  group: string
  input: ConditionInput
}

export const CONDITION_OPTIONS: readonly ConditionOption[] = [
  { id: 'cellIs:greaterThan', label: 'Greater than', group: 'Cell value', input: 'value' },
  { id: 'cellIs:greaterThanOrEqual', label: 'Greater than or equal to', group: 'Cell value', input: 'value' },
  { id: 'cellIs:lessThan', label: 'Less than', group: 'Cell value', input: 'value' },
  { id: 'cellIs:lessThanOrEqual', label: 'Less than or equal to', group: 'Cell value', input: 'value' },
  { id: 'cellIs:equal', label: 'Equal to', group: 'Cell value', input: 'value' },
  { id: 'cellIs:notEqual', label: 'Not equal to', group: 'Cell value', input: 'value' },
  { id: 'cellIs:between', label: 'Between', group: 'Cell value', input: 'values' },
  { id: 'cellIs:notBetween', label: 'Not between', group: 'Cell value', input: 'values' },
  { id: 'text:containsText', label: 'Text contains', group: 'Text', input: 'text' },
  { id: 'text:notContainsText', label: 'Text does not contain', group: 'Text', input: 'text' },
  { id: 'text:beginsWith', label: 'Text begins with', group: 'Text', input: 'text' },
  { id: 'text:endsWith', label: 'Text ends with', group: 'Text', input: 'text' },
  { id: 'date:today', label: 'Date is today', group: 'Date', input: 'none' },
  { id: 'date:yesterday', label: 'Date is yesterday', group: 'Date', input: 'none' },
  { id: 'date:tomorrow', label: 'Date is tomorrow', group: 'Date', input: 'none' },
  { id: 'date:last7Days', label: 'Date is in the last 7 days', group: 'Date', input: 'none' },
  { id: 'date:thisWeek', label: 'Date is this week', group: 'Date', input: 'none' },
  { id: 'date:lastWeek', label: 'Date is last week', group: 'Date', input: 'none' },
  { id: 'date:nextWeek', label: 'Date is next week', group: 'Date', input: 'none' },
  { id: 'date:thisMonth', label: 'Date is this month', group: 'Date', input: 'none' },
  { id: 'date:lastMonth', label: 'Date is last month', group: 'Date', input: 'none' },
  { id: 'date:nextMonth', label: 'Date is next month', group: 'Date', input: 'none' },
  { id: 'top:top', label: 'Top items', group: 'Top / bottom', input: 'rank' },
  { id: 'top:topPercent', label: 'Top percent', group: 'Top / bottom', input: 'percentRank' },
  { id: 'top:bottom', label: 'Bottom items', group: 'Top / bottom', input: 'rank' },
  { id: 'top:bottomPercent', label: 'Bottom percent', group: 'Top / bottom', input: 'percentRank' },
  { id: 'avg:above', label: 'Above average', group: 'Average', input: 'none' },
  { id: 'avg:below', label: 'Below average', group: 'Average', input: 'none' },
  { id: 'avg:equalOrAbove', label: 'Equal to or above average', group: 'Average', input: 'none' },
  { id: 'avg:equalOrBelow', label: 'Equal to or below average', group: 'Average', input: 'none' },
  { id: 'avg:stdDevAbove', label: 'Standard deviations above average', group: 'Average', input: 'stdDev' },
  { id: 'avg:stdDevBelow', label: 'Standard deviations below average', group: 'Average', input: 'stdDev' },
  { id: 'dup:duplicate', label: 'Duplicate values', group: 'Duplicates', input: 'none' },
  { id: 'dup:unique', label: 'Unique values', group: 'Duplicates', input: 'none' },
  { id: 'blank:blanks', label: 'Cell is empty', group: 'Blanks and errors', input: 'none' },
  { id: 'blank:noBlanks', label: 'Cell is not empty', group: 'Blanks and errors', input: 'none' },
  { id: 'error:errors', label: 'Cell contains an error', group: 'Blanks and errors', input: 'none' },
  { id: 'error:noErrors', label: 'Cell has no error', group: 'Blanks and errors', input: 'none' },
  { id: 'formula', label: 'Custom formula is', group: 'Formula', input: 'formula' },
]

export function conditionOption(id: ConditionId): ConditionOption {
  return CONDITION_OPTIONS.find((option) => option.id === id) ?? CONDITION_OPTIONS[0]
}

export interface StyleDraft {
  preset: ConditionalStylePresetId
  bold: boolean
  italic: boolean
  underline: boolean
  strike: boolean
  /** CSS hex colours or null for "no change". */
  fontColor: string | null
  fillColor: string | null
  borderColor: string | null
}

export type ScalePointType = 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula'
export type BarPointType = 'autoMin' | 'autoMax' | ScalePointType

export interface ScalePointDraft {
  type: ScalePointType
  value: string
  color: string
}

export interface ThresholdDraft {
  type: 'num' | 'percent' | 'percentile' | 'formula'
  value: string
  gte: boolean
}

export interface RuleDraft {
  ref: string
  format: RuleFormatKind
  condition: ConditionId
  value1: string
  value2: string
  stdDev: number
  stopIfTrue: boolean
  style: StyleDraft
  scale: { mid: boolean; points: [ScalePointDraft, ScalePointDraft, ScalePointDraft] }
  bar: {
    color: string
    gradient: boolean
    showValue: boolean
    negativeColor: string
    axisPosition: 'automatic' | 'middle' | 'none'
    direction: 'context' | 'leftToRight' | 'rightToLeft'
    min: { type: BarPointType; value: string }
    max: { type: BarPointType; value: string }
  }
  icons: {
    set: IconSetName
    reverse: boolean
    showValue: boolean
    /** Thresholds for icons 1..N-1 (index 0 = second-lowest icon). */
    thresholds: ThresholdDraft[]
  }
}

export interface ColorScalePreset {
  id: string
  label: string
  /** Low -> (mid) -> high. */
  colors: string[]
}

/** Excel's gallery scales, named top (highest value) to bottom like Excel. */
export const COLOR_SCALE_PRESETS: readonly ColorScalePreset[] = [
  { id: 'gyr', label: 'Green - Yellow - Red', colors: ['#F8696B', '#FFEB84', '#63BE7B'] },
  { id: 'ryg', label: 'Red - Yellow - Green', colors: ['#63BE7B', '#FFEB84', '#F8696B'] },
  { id: 'gwr', label: 'Green - White - Red', colors: ['#F8696B', '#FCFCFF', '#63BE7B'] },
  { id: 'rwg', label: 'Red - White - Green', colors: ['#63BE7B', '#FCFCFF', '#F8696B'] },
  { id: 'bwr', label: 'Blue - White - Red', colors: ['#F8696B', '#FCFCFF', '#5A8AC6'] },
  { id: 'rwb', label: 'Red - White - Blue', colors: ['#5A8AC6', '#FCFCFF', '#F8696B'] },
  { id: 'gw', label: 'Green - White', colors: ['#FCFCFF', '#63BE7B'] },
  { id: 'wg', label: 'White - Green', colors: ['#63BE7B', '#FCFCFF'] },
  { id: 'rw', label: 'Red - White', colors: ['#FCFCFF', '#F8696B'] },
  { id: 'wr', label: 'White - Red', colors: ['#F8696B', '#FCFCFF'] },
  { id: 'gy', label: 'Green - Yellow', colors: ['#FFEF9C', '#63BE7B'] },
  { id: 'yg', label: 'Yellow - Green', colors: ['#63BE7B', '#FFEF9C'] },
]

/** Excel's gallery data-bar colours. */
export const DATA_BAR_COLORS: readonly { label: string; color: string }[] = [
  { label: 'Blue', color: '#638EC6' },
  { label: 'Green', color: '#63C384' },
  { label: 'Red', color: '#FF555A' },
  { label: 'Orange', color: '#FFB628' },
  { label: 'Light blue', color: '#008AEF' },
  { label: 'Purple', color: '#D6007B' },
]

const ICON_SET_LABELS: Record<IconSetName, string> = {
  '3Arrows': '3 Arrows (colored)',
  '3ArrowsGray': '3 Arrows (gray)',
  '3Triangles': '3 Triangles',
  '3Flags': '3 Flags',
  '3TrafficLights1': '3 Traffic lights (unrimmed)',
  '3TrafficLights2': '3 Traffic lights (rimmed)',
  '3Signs': '3 Signs',
  '3Symbols': '3 Symbols (circled)',
  '3Symbols2': '3 Symbols (uncircled)',
  '3Stars': '3 Stars',
  '4Arrows': '4 Arrows (colored)',
  '4ArrowsGray': '4 Arrows (gray)',
  '4RedToBlack': 'Red to black',
  '4Rating': '4 Ratings',
  '4TrafficLights': '4 Traffic lights',
  '5Arrows': '5 Arrows (colored)',
  '5ArrowsGray': '5 Arrows (gray)',
  '5Rating': '5 Ratings',
  '5Quarters': '5 Quarters',
  '5Boxes': '5 Boxes',
}

export function iconSetLabel(name: string): string {
  return ICON_SET_LABELS[normalizeIconSetName(name) as IconSetName] ?? name
}

const TIME_PERIOD_LABELS: Record<TimePeriod, string> = {
  today: 'today',
  yesterday: 'yesterday',
  tomorrow: 'tomorrow',
  last7Days: 'in the last 7 days',
  thisWeek: 'this week',
  lastWeek: 'last week',
  nextWeek: 'next week',
  thisMonth: 'this month',
  lastMonth: 'last month',
  nextMonth: 'next month',
}

const OPERATOR_SYMBOLS: Record<CellIsOperator, string> = {
  greaterThan: '>',
  greaterThanOrEqual: '≥',
  lessThan: '<',
  lessThanOrEqual: '≤',
  equal: '=',
  notEqual: '≠',
  between: 'between',
  notBetween: 'not between',
}

// ---------------------------------------------------------------------------------------------
// Operand conversion (what the user types <-> the stored formula)
// ---------------------------------------------------------------------------------------------

const NUMERIC = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/

/** User input -> cellIs formula: "=A1" is a formula, numbers and TRUE/FALSE are literals, other text is quoted. */
export function operandFromInput(input: string): string {
  const text = input.trim()
  if (text.startsWith('=')) return text.slice(1).trim()
  if (NUMERIC.test(text)) return text
  if (/^(true|false)$/i.test(text)) return text.toUpperCase()
  return `"${text.replace(/"/g, '""')}"`
}

/** Stored cellIs formula -> what the user sees. */
export function inputFromOperand(formula: unknown): string {
  if (typeof formula === 'number') return String(formula)
  if (typeof formula !== 'string') return ''
  const text = formula.trim()
  if (NUMERIC.test(text) || /^(true|false)$/i.test(text)) return text
  if (/^"(?:[^"]|"")*"$/.test(text)) return text.slice(1, -1).replace(/""/g, '"')
  return `=${text}`
}

// ---------------------------------------------------------------------------------------------
// Drafts
// ---------------------------------------------------------------------------------------------

function defaultStyleDraft(): StyleDraft {
  return { ...styleDraftFrom(conditionalPresetStyle('lightRedFillDarkRedText')), preset: 'lightRedFillDarkRedText' }
}

function styleDraftFrom(style: unknown, cssColor = defaultConditionalCssColor): StyleDraft {
  const described = describeConditionalStyle(style, cssColor)
  return {
    preset: matchConditionalStylePreset(style),
    bold: described.bold,
    italic: described.italic,
    underline: described.underline,
    strike: described.strike,
    fontColor: described.fontColor,
    fillColor: described.fillColor,
    borderColor: described.borderColor,
  }
}

/** Apply a preset's colours to a style draft (keeps font toggles). */
export function applyStylePreset(draft: StyleDraft, preset: ConditionalStylePresetId): StyleDraft {
  if (preset === 'custom') return { ...draft, preset }
  const colors = styleDraftFrom(conditionalPresetStyle(preset))
  return { ...draft, preset, fontColor: colors.fontColor, fillColor: colors.fillColor, borderColor: colors.borderColor }
}

export function styleFromDraft(draft: StyleDraft, base?: CellStyle): CellStyle {
  const style = createConditionalStyle(draft)
  if (base && typeof base === 'object' && typeof base.numFmt === 'string' && base.numFmt) style.numFmt = base.numFmt
  return style
}

function defaultScale(): RuleDraft['scale'] {
  return {
    mid: true,
    points: [
      { type: 'min', value: '', color: '#F8696B' },
      { type: 'percentile', value: '50', color: '#FFEB84' },
      { type: 'max', value: '', color: '#63BE7B' },
    ],
  }
}

function defaultThresholds(set: string): ThresholdDraft[] {
  return defaultIconThresholds(iconSetSize(set)).slice(1).map((value) => ({ type: 'percent', value: String(value), gte: true }))
}

/** A new rule draft for `ref` (Excel's default: cell value greater than, light red fill). */
export function defaultRuleDraft(ref: string): RuleDraft {
  return {
    ref,
    format: 'single',
    condition: 'cellIs:greaterThan',
    value1: '',
    value2: '',
    stdDev: 1,
    stopIfTrue: false,
    style: defaultStyleDraft(),
    scale: defaultScale(),
    bar: {
      color: '#638EC6',
      gradient: true,
      showValue: true,
      negativeColor: '#FF0000',
      axisPosition: 'automatic',
      direction: 'context',
      min: { type: 'autoMin', value: '' },
      max: { type: 'autoMax', value: '' },
    },
    icons: { set: '3Arrows', reverse: false, showValue: true, thresholds: defaultThresholds('3Arrows') },
  }
}

/** Reset icon thresholds when the set size changes. */
export function withIconSet(draft: RuleDraft, set: IconSetName): RuleDraft {
  const size = iconSetSize(set)
  const thresholds = draft.icons.thresholds.length === size - 1 ? draft.icons.thresholds : defaultThresholds(set)
  return { ...draft, icons: { ...draft.icons, set, thresholds } }
}

function valueText(value: unknown): string {
  if (value === undefined || value === null) return ''
  if (typeof value === 'number') return String(value)
  const text = String(value)
  return NUMERIC.test(text.trim()) ? text.trim() : `=${text.replace(/^=/, '')}`
}

function hex(cssColor: (color: unknown, fallback?: string) => string, color: unknown, fallback: string): string {
  const css = cssColor(color, fallback) || fallback
  return /^#[0-9a-f]{6}$/i.test(css) ? css.toUpperCase() : fallback
}

/** Convert a stored rule to an editor draft. */
export function ruleToDraft(rule: ConditionalRule, ref: string, cssColor: (color: unknown, fallback?: string) => string = defaultConditionalCssColor): RuleDraft {
  const draft = defaultRuleDraft(ref)
  const kind = conditionalRuleKind(rule)
  draft.stopIfTrue = rule.stopIfTrue === true
  if (rule.style) draft.style = styleDraftFrom(rule.style, cssColor)
  const formulae = Array.isArray(rule.formulae) ? rule.formulae : []
  switch (kind) {
    case 'cellIs': {
      const operator = (String(rule.operator || 'equal') as CellIsOperator)
      draft.condition = `cellIs:${operator}` as ConditionId
      draft.value1 = inputFromOperand(formulae[0])
      draft.value2 = inputFromOperand(formulae[1])
      break
    }
    case 'containsText':
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      if (typeof rule.text === 'string') {
        draft.condition = `text:${kind}` as ConditionId
        draft.value1 = rule.text
      } else {
        draft.condition = 'formula'
        draft.value1 = formulae[0] ? `=${formulae[0]}` : ''
      }
      break
    case 'timePeriod':
      draft.condition = `date:${rule.timePeriod || 'today'}` as ConditionId
      if (!CONDITION_OPTIONS.some((option) => option.id === draft.condition)) draft.condition = 'date:today'
      break
    case 'top10': {
      const bottom = rule.bottom === true
      const percent = rule.percent === true
      draft.condition = bottom ? (percent ? 'top:bottomPercent' : 'top:bottom') : percent ? 'top:topPercent' : 'top:top'
      draft.value1 = String(rule.rank ?? 10)
      break
    }
    case 'aboveAverage': {
      const below = rule.aboveAverage === false
      const deviations = Math.floor(Number(rule.stdDev) || 0)
      if (deviations > 0) {
        draft.condition = below ? 'avg:stdDevBelow' : 'avg:stdDevAbove'
        draft.stdDev = Math.min(3, deviations)
      } else if (rule.equalAverage) draft.condition = below ? 'avg:equalOrBelow' : 'avg:equalOrAbove'
      else draft.condition = below ? 'avg:below' : 'avg:above'
      break
    }
    case 'duplicateValues': draft.condition = 'dup:duplicate'; break
    case 'uniqueValues': draft.condition = 'dup:unique'; break
    case 'containsBlanks': draft.condition = 'blank:blanks'; break
    case 'notContainsBlanks': draft.condition = 'blank:noBlanks'; break
    case 'containsErrors': draft.condition = 'error:errors'; break
    case 'notContainsErrors': draft.condition = 'error:noErrors'; break
    case 'expression':
      draft.condition = 'formula'
      draft.value1 = formulae[0] ? `=${String(formulae[0]).replace(/^=/, '')}` : ''
      break
    case 'colorScale': {
      draft.format = 'colorScale'
      const cfvo = Array.isArray(rule.cfvo) ? rule.cfvo : []
      const colors = Array.isArray(rule.color) ? rule.color : []
      const point = (index: number, fallbackType: ScalePointType, fallbackColor: string): ScalePointDraft => {
        const item = cfvo[index]
        const type = (item && ['min', 'max', 'num', 'percent', 'percentile', 'formula'].includes(String(item.type)) ? item.type : fallbackType) as ScalePointType
        return { type, value: valueText(item?.value).replace(/^=(?=[-+.\d])/, ''), color: hex(cssColor, colors[index], fallbackColor) }
      }
      if (cfvo.length >= 3) {
        draft.scale = { mid: true, points: [point(0, 'min', '#F8696B'), point(1, 'percentile', '#FFEB84'), point(2, 'max', '#63BE7B')] }
      } else {
        const base = defaultScale()
        draft.scale = { mid: false, points: [point(0, 'min', '#F8696B'), base.points[1], point(1, 'max', '#63BE7B')] }
      }
      break
    }
    case 'dataBar': {
      draft.format = 'dataBar'
      const cfvo = Array.isArray(rule.cfvo) ? rule.cfvo : []
      const barPoint = (index: number, fallback: BarPointType) => {
        const item = cfvo[index]
        const type = (item && ['autoMin', 'autoMax', 'min', 'max', 'num', 'percent', 'percentile', 'formula'].includes(String(item.type)) ? item.type : fallback) as BarPointType
        return { type, value: valueText(item?.value) }
      }
      const color = Array.isArray(rule.color) ? rule.color[0] : rule.color
      draft.bar = {
        color: hex(cssColor, color, '#638EC6'),
        gradient: rule.gradient !== false,
        showValue: rule.showValue !== false,
        negativeColor: hex(cssColor, rule.negativeFillColor, '#FF0000'),
        axisPosition: rule.axisPosition === 'middle' || rule.axisPosition === 'none' ? rule.axisPosition : 'automatic',
        direction: rule.direction === 'leftToRight' || rule.direction === 'rightToLeft' ? rule.direction : 'context',
        min: barPoint(0, 'autoMin'),
        max: barPoint(1, 'autoMax'),
      }
      break
    }
    case 'iconSet': {
      draft.format = 'iconSet'
      const set = normalizeIconSetName(rule.iconSet) as IconSetName
      const size = iconSetSize(set)
      const cfvo = Array.isArray(rule.cfvo) ? rule.cfvo : []
      draft.icons = {
        set,
        reverse: rule.reverse === true,
        showValue: rule.showValue !== false,
        thresholds: cfvo.length === size
          ? cfvo.slice(1).map((item) => ({
            type: (['num', 'percent', 'percentile', 'formula'].includes(String(item.type)) ? item.type : 'percent') as ThresholdDraft['type'],
            value: valueText(item.value).replace(/^=(?=[-+.\d])/, ''),
            gte: item.gte !== false,
          }))
          : defaultThresholds(set),
      }
      break
    }
    default:
      break
  }
  return draft
}

function cfvoInput(type: string, value: string): { type: CfvoType; value?: number | string } {
  if (type === 'min' || type === 'max' || type === 'autoMin' || type === 'autoMax') return { type: type as CfvoType }
  const text = value.trim()
  if (NUMERIC.test(text)) return { type: type as CfvoType, value: Number(text) }
  return { type: type as CfvoType, value: text.replace(/^=/, '') }
}

function pointError(label: string, type: string, value: string): string | null {
  if (type === 'min' || type === 'max' || type === 'autoMin' || type === 'autoMax') return null
  const text = value.trim()
  if (!text) return `Enter a value for the ${label}.`
  if (type === 'formula') return null
  if (!NUMERIC.test(text) && !text.startsWith('=')) return `The ${label} must be a number or a formula starting with "=".`
  if ((type === 'percent' || type === 'percentile') && NUMERIC.test(text) && (Number(text) < 0 || Number(text) > 100)) {
    return `The ${label} ${type} must be between 0 and 100.`
  }
  return null
}

/** Validate a draft; returns the first problem or null. */
export function validateDraft(draft: RuleDraft): string | null {
  const range = parseRangeInput(draft.ref)
  if ('error' in range) return range.error
  if (draft.format === 'single') {
    const option = conditionOption(draft.condition)
    switch (option.input) {
      case 'value':
        if (!draft.value1.trim()) return 'Enter a value.'
        break
      case 'values':
        if (!draft.value1.trim() || !draft.value2.trim()) return 'Enter both values.'
        break
      case 'text':
        if (!draft.value1) return 'Enter the text to look for.'
        break
      case 'rank':
      case 'percentRank': {
        const rank = Number(draft.value1)
        const limit = option.input === 'percentRank' ? 100 : 1000
        if (!Number.isInteger(rank) || rank < 1 || rank > limit) return `Enter a whole number from 1 to ${limit}.`
        break
      }
      case 'formula':
        if (!draft.value1.trim().replace(/^=/, '').trim()) return 'Enter a formula, for example =A1>10.'
        break
      default:
        break
    }
    return null
  }
  if (draft.format === 'colorScale') {
    const [min, mid, max] = draft.scale.points
    return pointError('minimum', min.type, min.value)
      ?? (draft.scale.mid ? pointError('midpoint', mid.type, mid.value) : null)
      ?? pointError('maximum', max.type, max.value)
  }
  if (draft.format === 'dataBar') {
    return pointError('minimum', draft.bar.min.type, draft.bar.min.value) ?? pointError('maximum', draft.bar.max.type, draft.bar.max.value)
  }
  for (const [index, threshold] of draft.icons.thresholds.entries()) {
    const problem = pointError(`threshold ${index + 1}`, threshold.type, threshold.value)
    if (problem) return problem
  }
  return null
}

/**
 * Convert a draft to a rule (ExcelJS model). `base` is the rule being edited: its priority and
 * any properties the editor does not expose (number format, custom icons, bar lengths) are kept.
 */
export function draftToRule(draft: RuleDraft, base?: ConditionalRule): { rule: ConditionalRule; ref: string } | { error: string } {
  const problem = validateDraft(draft)
  if (problem) return { error: problem }
  const range = parseRangeInput(draft.ref)
  if ('error' in range) return { error: range.error }
  const ref = range.ref
  const baseKind = base ? conditionalRuleKind(base) : null
  let rule: ConditionalRule
  if (draft.format === 'single') {
    const options = { style: styleFromDraft(draft.style, base?.style), stopIfTrue: draft.stopIfTrue, ref }
    const [group, variant] = draft.condition.split(':') as [string, string | undefined]
    switch (group) {
      case 'cellIs':
        rule = createCellIsRule(variant as CellIsOperator, [operandFromInput(draft.value1), operandFromInput(draft.value2)], options)
        break
      case 'text':
        rule = createTextRule(variant as 'containsText', draft.value1, options)
        break
      case 'date':
        rule = createDateRule(variant as TimePeriod, options)
        break
      case 'top':
        rule = createTopBottomRule({ rank: Number(draft.value1), percent: variant?.endsWith('Percent'), bottom: variant?.startsWith('bottom') }, options)
        break
      case 'avg':
        rule = createAverageRule({
          below: variant === 'below' || variant === 'equalOrBelow' || variant === 'stdDevBelow',
          equal: variant === 'equalOrAbove' || variant === 'equalOrBelow',
          stdDev: variant === 'stdDevAbove' || variant === 'stdDevBelow' ? draft.stdDev : 0,
        }, options)
        break
      case 'dup':
        rule = createDuplicateRule(variant === 'unique', options)
        break
      case 'blank':
        rule = createBlankRule(variant === 'blanks', options)
        break
      case 'error':
        rule = createErrorRule(variant === 'errors', options)
        break
      default:
        rule = createFormulaRule(draft.value1, options)
        break
    }
  } else if (draft.format === 'colorScale') {
    const [min, mid, max] = draft.scale.points
    const points = draft.scale.mid ? [min, mid, max] : [min, max]
    rule = createColorScaleRule(points.map((point) => ({ ...cfvoInput(point.type, point.value), color: point.color })))
  } else if (draft.format === 'dataBar') {
    rule = createDataBarRule({
      color: draft.bar.color,
      gradient: draft.bar.gradient,
      showValue: draft.bar.showValue,
      negativeColor: draft.bar.negativeColor,
      axisPosition: draft.bar.axisPosition,
      direction: draft.bar.direction,
      min: cfvoInput(draft.bar.min.type, draft.bar.min.value),
      max: cfvoInput(draft.bar.max.type, draft.bar.max.value),
    })
    if (baseKind === 'dataBar' && base) {
      if (base.minLength !== undefined) rule.minLength = base.minLength
      if (base.maxLength !== undefined) rule.maxLength = base.maxLength
    }
  } else {
    rule = createIconSetRule(draft.icons.set, {
      reverse: draft.icons.reverse,
      showValue: draft.icons.showValue,
      thresholds: draft.icons.thresholds.map((threshold) => ({ ...cfvoInput(threshold.type, threshold.value), value: cfvoInput(threshold.type, threshold.value).value ?? 0, gte: threshold.gte })),
    })
    if (baseKind === 'iconSet' && base && Array.isArray(base.icons) && base.icons.length && normalizeIconSetName(base.iconSet) === draft.icons.set) {
      rule.icons = base.icons.map((icon) => ({ ...icon }))
    }
  }
  if (base?.priority !== undefined) rule.priority = base.priority
  return { rule, ref }
}

// ---------------------------------------------------------------------------------------------
// Descriptions
// ---------------------------------------------------------------------------------------------

function quoted(text: unknown) {
  return `“${String(text ?? '')}”`
}

/** Human-readable summary of a rule (Rules Manager "Rule" column). */
export function describeRule(rule: ConditionalRule): string {
  const kind = conditionalRuleKind(rule)
  const formulae = Array.isArray(rule.formulae) ? rule.formulae : []
  switch (kind) {
    case 'cellIs': {
      const operator = String(rule.operator || 'equal') as CellIsOperator
      if (operator === 'between' || operator === 'notBetween') {
        return `Cell value ${OPERATOR_SYMBOLS[operator]} ${inputFromOperand(formulae[0])} and ${inputFromOperand(formulae[1])}`
      }
      return `Cell value ${OPERATOR_SYMBOLS[operator] ?? operator} ${inputFromOperand(formulae[0])}`
    }
    case 'containsText': return typeof rule.text === 'string' ? `Text contains ${quoted(rule.text)}` : `Formula: =${formulae[0] ?? ''}`
    case 'notContainsText': return `Text does not contain ${quoted(rule.text)}`
    case 'beginsWith': return `Text begins with ${quoted(rule.text)}`
    case 'endsWith': return `Text ends with ${quoted(rule.text)}`
    case 'containsBlanks': return 'Cell is empty'
    case 'notContainsBlanks': return 'Cell is not empty'
    case 'containsErrors': return 'Cell contains an error'
    case 'notContainsErrors': return 'Cell has no error'
    case 'timePeriod': return `Date is ${TIME_PERIOD_LABELS[rule.timePeriod as TimePeriod] ?? String(rule.timePeriod || '')}`
    case 'top10': return `${rule.bottom ? 'Bottom' : 'Top'} ${rule.rank ?? 10}${rule.percent ? '%' : ''}`
    case 'aboveAverage': {
      const below = rule.aboveAverage === false
      const deviations = Math.floor(Number(rule.stdDev) || 0)
      if (deviations > 0) return `${deviations} std dev ${below ? 'below' : 'above'} average`
      if (rule.equalAverage) return `Equal to or ${below ? 'below' : 'above'} average`
      return below ? 'Below average' : 'Above average'
    }
    case 'duplicateValues': return 'Duplicate values'
    case 'uniqueValues': return 'Unique values'
    case 'expression': return `Formula: =${String(formulae[0] ?? '').replace(/^=/, '')}`
    case 'colorScale': return `${Array.isArray(rule.cfvo) && rule.cfvo.length >= 3 ? '3' : '2'}-color scale`
    case 'dataBar': return rule.showValue === false ? 'Data bar (bar only)' : 'Data bar'
    case 'iconSet': return `Icon set: ${iconSetLabel(String(rule.iconSet || ''))}${rule.showValue === false ? ' (icon only)' : ''}`
    default: return `Unsupported rule (${String(rule.type)})`
  }
}

export { CONDITIONAL_STYLE_PRESETS }
