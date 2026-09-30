import { useCallback, useEffect, useId, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react'
import {
  ArrowLeft,
  Bold,
  Check,
  ChevronDown,
  ChevronUp,
  CircleAlert,
  Eraser,
  Italic,
  Paintbrush,
  Plus,
  Strikethrough,
  Trash2,
  Underline,
  X,
} from 'lucide-react'
import {
  CONDITIONAL_STYLE_PRESETS,
  ICON_SET_NAMES,
  addRule,
  clearRulesInRange,
  conditionalRuleKind,
  conditionalStylePreview,
  createConditionalStyle,
  dataBarBackground,
  deleteRule,
  iconSetSize,
  listRulesForRange,
  moveRulePriority,
  normalizeIconSetName,
  parseRangeInput,
  ruleAnchorAddress,
  ruleUsesStyle,
  updateRule,
} from '../lib/conditional-format'
import type { ConditionalCellFormat, ConditionalRule, ConditionalRuleEntry, ConditionalStylePresetId, IconSetName } from '../lib/conditional-format'
import {
  COLOR_SCALE_PRESETS,
  CONDITION_OPTIONS,
  DATA_BAR_COLORS,
  applyStylePreset,
  conditionOption,
  defaultRuleDraft,
  describeRule,
  draftToRule,
  iconSetLabel,
  ruleToDraft,
  validateDraft,
  withIconSet,
} from '../lib/conditional-format-editor'
import type { BarPointType, ConditionId, RuleDraft, RuleFormatKind, ScalePointType, StyleDraft, ThresholdDraft } from '../lib/conditional-format-editor'
import { ConditionalIcon } from './ConditionalIcon'
import './conditional-format.css'

export { ConditionalIcon } from './ConditionalIcon'

export interface ConditionalFormatPanelProps {
  sheetName: string
  /** Current selection, e.g. "B2:D20". */
  selectionRef: string
  /** The sheet's current model (SheetData.conditionalFormattings). */
  conditionalFormattings: unknown[]
  /** Receives the complete next model; the host applies it as one undoable change. */
  onChange: (next: unknown[]) => void
  onClose: () => void
  cssColor: (color: unknown, fallback?: string) => string
}

type Scope = 'selection' | 'sheet'

interface EditingState {
  id: number | null
  base?: ConditionalRule
  draft: RuleDraft
}

const FOCUSABLE = 'button:not([disabled]), select:not([disabled]), input:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'

const FORMAT_TABS: ReadonlyArray<{ id: RuleFormatKind; label: string }> = [
  { id: 'single', label: 'Single color' },
  { id: 'colorScale', label: 'Color scale' },
  { id: 'dataBar', label: 'Data bar' },
  { id: 'iconSet', label: 'Icon set' },
]

const CONDITION_GROUPS = Array.from(new Set(CONDITION_OPTIONS.map((option) => option.group)))

function styleFormatCss(format: ConditionalCellFormat): CSSProperties {
  const decorations = [format.font?.underline ? 'underline' : '', format.font?.strike ? 'line-through' : ''].filter(Boolean).join(' ')
  const border = format.border
  return {
    color: format.font?.color,
    backgroundColor: format.fill,
    fontWeight: format.font?.bold ? 700 : undefined,
    fontStyle: format.font?.italic ? 'italic' : undefined,
    textDecorationLine: decorations || undefined,
    borderTop: border?.top,
    borderRight: border?.right,
    borderBottom: border?.bottom,
    borderLeft: border?.left,
  }
}

function draftStyleCss(style: StyleDraft): CSSProperties {
  return styleFormatCss(conditionalStylePreview(createConditionalStyle(style)))
}

function gradientCss(colors: readonly string[]): string {
  return `linear-gradient(90deg, ${colors.join(', ')})`
}

function ruleScaleColors(rule: ConditionalRule, cssColor: ConditionalFormatPanelProps['cssColor']): string[] {
  const colors = Array.isArray(rule.color) ? rule.color : []
  return colors.map((color) => cssColor(color, '#FFFFFF') || '#FFFFFF')
}

/** Small swatch that previews a rule's effect in the list. */
function RulePreview({ rule, cssColor }: { rule: ConditionalRule; cssColor: ConditionalFormatPanelProps['cssColor'] }) {
  const kind = conditionalRuleKind(rule)
  if (kind === 'colorScale') {
    return <span className="cf-swatch cf-swatch-fill" style={{ backgroundImage: gradientCss(ruleScaleColors(rule, cssColor)) }} aria-hidden="true" />
  }
  if (kind === 'dataBar') {
    const color = cssColor(Array.isArray(rule.color) ? rule.color[0] : rule.color, '#638EC6') || '#638EC6'
    const bar = dataBarBackground({ fraction: 0.68, color, gradient: rule.gradient !== false, showValue: true, border: rule.border ? cssColor(rule.borderColor, color) || color : undefined })
    return <span className="cf-swatch" style={bar} aria-hidden="true" />
  }
  if (kind === 'iconSet') {
    const set = normalizeIconSetName(rule.iconSet)
    const size = iconSetSize(set)
    const indexes = Array.from({ length: Math.min(size, 3) }, (_, offset) => (size === 3 ? 2 - offset : [size - 1, Math.floor(size / 2), 0][offset]))
    return (
      <span className="cf-swatch cf-swatch-icons" aria-hidden="true">
        {indexes.map((index) => <ConditionalIcon key={index} set={set} index={rule.reverse ? size - 1 - index : index} size={12} />)}
      </span>
    )
  }
  const format = conditionalStylePreview(rule.style, cssColor)
  const empty = !format.fill && !format.font && !format.border
  return <span className={`cf-swatch cf-swatch-text${empty ? ' is-empty' : ''}`} style={styleFormatCss(format)} aria-hidden="true">{empty ? '—' : '123'}</span>
}

interface ColorFieldProps {
  label: string
  value: string | null
  onChange: (value: string | null) => void
  allowNone?: boolean
  children?: ReactNode
}

/** Swatch + native colour picker, optionally clearable ("no colour"). */
function ColorField({ label, value, onChange, allowNone, children }: ColorFieldProps) {
  return (
    <span className={`cf-color-field${value ? '' : ' is-none'}`}>
      <label className={`cf-color-swatch${children ? '' : ' is-plain'}`} title={label}>
        {children}
        <span className="cf-color-chip" style={{ backgroundColor: value ?? 'transparent' }} aria-hidden="true" />
        <input type="color" aria-label={label} value={value ?? '#000000'} onChange={(event) => onChange(event.target.value.toUpperCase())} />
      </label>
      {allowNone && value && (
        <button type="button" className="cf-color-clear" aria-label={`Remove ${label.toLowerCase()}`} title={`Remove ${label.toLowerCase()}`} onClick={() => onChange(null)}>
          <X size={11} aria-hidden="true" />
        </button>
      )}
    </span>
  )
}

function ToggleButton({ label, pressed, onToggle, children }: { label: string; pressed: boolean; onToggle: () => void; children: ReactNode }) {
  return (
    <button type="button" className={`cf-toggle${pressed ? ' is-on' : ''}`} aria-label={label} title={label} aria-pressed={pressed} onClick={onToggle}>
      {children}
    </button>
  )
}

const SCALE_TYPE_LABELS: Record<ScalePointType, string> = {
  min: 'Lowest value',
  max: 'Highest value',
  num: 'Number',
  percent: 'Percent',
  percentile: 'Percentile',
  formula: 'Formula',
}

const BAR_TYPE_LABELS: Record<BarPointType, string> = {
  autoMin: 'Automatic',
  autoMax: 'Automatic',
  ...SCALE_TYPE_LABELS,
}

function needsValue(type: string) {
  return type !== 'min' && type !== 'max' && type !== 'autoMin' && type !== 'autoMax'
}

// ---------------------------------------------------------------------------------------------
// Editors per format
// ---------------------------------------------------------------------------------------------

interface EditorProps {
  draft: RuleDraft
  update: (patch: Partial<RuleDraft>) => void
  anchor: string
}

function SingleColorEditor({ draft, update, anchor }: EditorProps) {
  const option = conditionOption(draft.condition)
  const style = draft.style
  const setStyle = (patch: Partial<StyleDraft>) => update({ style: { ...style, ...patch, preset: patch.preset ?? 'custom' } })
  const valueId = useId()
  return (
    <>
      <label className="cf-field">
        <span>Format cells if…</span>
        <select value={draft.condition} onChange={(event) => update({ condition: event.target.value as ConditionId })}>
          {CONDITION_GROUPS.map((group) => (
            <optgroup key={group} label={group}>
              {CONDITION_OPTIONS.filter((item) => item.group === group).map((item) => <option key={item.id} value={item.id}>{item.label}</option>)}
            </optgroup>
          ))}
        </select>
      </label>

      {option.input === 'value' && (
        <label className="cf-field">
          <span>Value</span>
          <input value={draft.value1} placeholder="A number, text, or a formula such as =$B$1" spellCheck={false} onChange={(event) => update({ value1: event.target.value })} />
        </label>
      )}
      {option.input === 'values' && (
        <div className="cf-field">
          <span id={valueId}>Values</span>
          <div className="cf-between" role="group" aria-labelledby={valueId}>
            <input aria-label="From" value={draft.value1} placeholder="Minimum" spellCheck={false} onChange={(event) => update({ value1: event.target.value })} />
            <span>and</span>
            <input aria-label="To" value={draft.value2} placeholder="Maximum" spellCheck={false} onChange={(event) => update({ value2: event.target.value })} />
          </div>
        </div>
      )}
      {option.input === 'text' && (
        <label className="cf-field">
          <span>Text</span>
          <input value={draft.value1} placeholder="Case-insensitive; * and ? are wildcards" spellCheck={false} onChange={(event) => update({ value1: event.target.value })} />
        </label>
      )}
      {(option.input === 'rank' || option.input === 'percentRank') && (
        <label className="cf-field">
          <span>{option.input === 'rank' ? 'Number of items' : 'Percent of items'}</span>
          <span className="cf-suffixed">
            <input type="number" min={1} max={option.input === 'rank' ? 1000 : 100} value={draft.value1} placeholder="10" onChange={(event) => update({ value1: event.target.value })} />
            <em>{option.input === 'rank' ? 'items' : '%'}</em>
          </span>
        </label>
      )}
      {option.input === 'stdDev' && (
        <label className="cf-field">
          <span>Standard deviations</span>
          <select value={draft.stdDev} onChange={(event) => update({ stdDev: Number(event.target.value) })}>
            <option value={1}>1 std dev</option>
            <option value={2}>2 std dev</option>
            <option value={3}>3 std dev</option>
          </select>
        </label>
      )}
      {option.input === 'formula' && (
        <label className="cf-field">
          <span>Formula</span>
          <input className="cf-mono" value={draft.value1} placeholder={`=${anchor}>100`} spellCheck={false} onChange={(event) => update({ value1: event.target.value })} />
          <small>Write it for the top-left cell ({anchor}); relative references move with each cell.</small>
        </label>
      )}

      <fieldset className="cf-group">
        <legend>Formatting style</legend>
        <div className="cf-style-row">
          <select
            aria-label="Preset style"
            value={style.preset}
            onChange={(event) => update({ style: applyStylePreset(style, event.target.value as ConditionalStylePresetId) })}
          >
            {CONDITIONAL_STYLE_PRESETS.map((preset) => <option key={preset.id} value={preset.id}>{preset.label}</option>)}
            <option value="custom">Custom format</option>
          </select>
          <span className="cf-style-preview" style={draftStyleCss(style)}>AaBbCc 123</span>
        </div>
        <div className="cf-toolbar" role="toolbar" aria-label="Font and colors">
          <ToggleButton label="Bold" pressed={style.bold} onToggle={() => setStyle({ bold: !style.bold })}><Bold size={14} aria-hidden="true" /></ToggleButton>
          <ToggleButton label="Italic" pressed={style.italic} onToggle={() => setStyle({ italic: !style.italic })}><Italic size={14} aria-hidden="true" /></ToggleButton>
          <ToggleButton label="Underline" pressed={style.underline} onToggle={() => setStyle({ underline: !style.underline })}><Underline size={14} aria-hidden="true" /></ToggleButton>
          <ToggleButton label="Strikethrough" pressed={style.strike} onToggle={() => setStyle({ strike: !style.strike })}><Strikethrough size={14} aria-hidden="true" /></ToggleButton>
          <span className="cf-toolbar-separator" aria-hidden="true" />
          <ColorField label="Text color" value={style.fontColor} allowNone onChange={(fontColor) => setStyle({ fontColor })}><span className="cf-color-glyph" aria-hidden="true">A</span></ColorField>
          <ColorField label="Fill color" value={style.fillColor} allowNone onChange={(fillColor) => setStyle({ fillColor })}><Paintbrush size={12} aria-hidden="true" /></ColorField>
          <ColorField label="Border color" value={style.borderColor} allowNone onChange={(borderColor) => setStyle({ borderColor })}><span className="cf-border-glyph" aria-hidden="true" /></ColorField>
        </div>
      </fieldset>

      <label className="cf-check">
        <input type="checkbox" checked={draft.stopIfTrue} onChange={(event) => update({ stopIfTrue: event.target.checked })} />
        <span><strong>Stop if true</strong><small>Cells this rule formats skip the rules below it.</small></span>
      </label>
    </>
  )
}

function PointRow({
  label,
  types,
  labels,
  type,
  value,
  color,
  onType,
  onValue,
  onColor,
  disabled,
}: {
  label: string
  types: readonly string[]
  labels: Record<string, string>
  type: string
  value: string
  color?: string
  onType: (type: string) => void
  onValue: (value: string) => void
  onColor?: (color: string) => void
  disabled?: boolean
}) {
  return (
    <div className={`cf-point-row${disabled ? ' is-disabled' : ''}`} role="group" aria-label={label}>
      <span className="cf-point-label">{label}</span>
      <select aria-label={`${label} type`} value={type} disabled={disabled} onChange={(event) => onType(event.target.value)}>
        {types.map((item) => <option key={item} value={item}>{labels[item]}</option>)}
      </select>
      <input
        aria-label={`${label} value`}
        className="cf-mono"
        value={needsValue(type) ? value : ''}
        disabled={disabled || !needsValue(type)}
        placeholder={needsValue(type) ? (type === 'formula' ? '=$A$1' : '0') : ''}
        spellCheck={false}
        onChange={(event) => onValue(event.target.value)}
      />
      {onColor && color !== undefined && (
        <span className={disabled ? 'cf-muted-color' : undefined}>
          <ColorField label={`${label} color`} value={color} onChange={(next) => next && onColor(next)} />
        </span>
      )}
    </div>
  )
}

function ColorScaleEditor({ draft, update }: EditorProps) {
  const { scale } = draft
  const setPoint = (index: 0 | 1 | 2, patch: Partial<RuleDraft['scale']['points'][number]>) => {
    const points = scale.points.map((point, position) => (position === index ? { ...point, ...patch } : point)) as RuleDraft['scale']['points']
    update({ scale: { ...scale, points } })
  }
  const colors = scale.mid ? scale.points.map((point) => point.color) : [scale.points[0].color, scale.points[2].color]
  return (
    <>
      <div className="cf-field">
        <span>Presets</span>
        <div className="cf-scale-presets" role="group" aria-label="Preset color scales">
          {COLOR_SCALE_PRESETS.map((preset) => (
            <button
              key={preset.id}
              type="button"
              className="cf-scale-chip"
              title={preset.label}
              aria-label={preset.label}
              style={{ backgroundImage: gradientCss(preset.colors) }}
              onClick={() => {
                const three = preset.colors.length === 3
                const points = [
                  { ...scale.points[0], color: preset.colors[0] },
                  { ...scale.points[1], color: three ? preset.colors[1] : scale.points[1].color },
                  { ...scale.points[2], color: preset.colors[preset.colors.length - 1] },
                ] as RuleDraft['scale']['points']
                update({ scale: { mid: three, points } })
              }}
            />
          ))}
        </div>
      </div>
      <div className="cf-gradient-preview" style={{ backgroundImage: gradientCss(colors) }} aria-hidden="true">
        <span>Low</span>
        <span>High</span>
      </div>
      <fieldset className="cf-group cf-points">
        <legend>Points</legend>
        <PointRow
          label="Minpoint"
          types={['min', 'num', 'percent', 'percentile', 'formula']}
          labels={SCALE_TYPE_LABELS}
          type={scale.points[0].type}
          value={scale.points[0].value}
          color={scale.points[0].color}
          onType={(type) => setPoint(0, { type: type as ScalePointType })}
          onValue={(value) => setPoint(0, { value })}
          onColor={(color) => setPoint(0, { color })}
        />
        <label className="cf-check cf-check-compact">
          <input type="checkbox" checked={scale.mid} onChange={(event) => update({ scale: { ...scale, mid: event.target.checked } })} />
          <span><strong>Use a midpoint</strong></span>
        </label>
        <PointRow
          label="Midpoint"
          types={['num', 'percent', 'percentile', 'formula']}
          labels={SCALE_TYPE_LABELS}
          type={scale.points[1].type}
          value={scale.points[1].value}
          color={scale.points[1].color}
          disabled={!scale.mid}
          onType={(type) => setPoint(1, { type: type as ScalePointType })}
          onValue={(value) => setPoint(1, { value })}
          onColor={(color) => setPoint(1, { color })}
        />
        <PointRow
          label="Maxpoint"
          types={['max', 'num', 'percent', 'percentile', 'formula']}
          labels={SCALE_TYPE_LABELS}
          type={scale.points[2].type}
          value={scale.points[2].value}
          color={scale.points[2].color}
          onType={(type) => setPoint(2, { type: type as ScalePointType })}
          onValue={(value) => setPoint(2, { value })}
          onColor={(color) => setPoint(2, { color })}
        />
      </fieldset>
    </>
  )
}

function DataBarEditor({ draft, update }: EditorProps) {
  const { bar } = draft
  const setBar = (patch: Partial<RuleDraft['bar']>) => update({ bar: { ...bar, ...patch } })
  const samples = [
    { label: '82', fraction: 0.82, negative: false },
    { label: '35', fraction: 0.35, negative: false },
    { label: '-40', fraction: 0.4, negative: true },
  ]
  const axis = bar.axisPosition === 'middle' ? 0.5 : bar.axisPosition === 'none' ? undefined : 0.3
  return (
    <>
      <div className="cf-bar-preview" aria-hidden="true">
        {samples.map((sample) => {
          const negative = sample.negative && axis !== undefined
          const background = dataBarBackground({
            fraction: negative ? sample.fraction * (axis ?? 0) : axis === undefined ? sample.fraction : sample.fraction * (1 - axis),
            color: sample.negative ? bar.negativeColor : bar.color,
            negative,
            axis,
            axisColor: '#000000',
            gradient: bar.gradient,
            border: bar.gradient ? (sample.negative ? bar.negativeColor : bar.color) : undefined,
            showValue: bar.showValue,
            rtl: bar.direction === 'rightToLeft',
          })
          return <span key={sample.label} className="cf-bar-sample" style={background}>{bar.showValue ? sample.label : ''}</span>
        })}
      </div>
      <div className="cf-field">
        <span>Bar color</span>
        <div className="cf-swatches" role="group" aria-label="Bar color">
          {DATA_BAR_COLORS.map((item) => (
            <button
              key={item.color}
              type="button"
              className={`cf-swatch-button${bar.color.toUpperCase() === item.color ? ' is-selected' : ''}`}
              aria-label={item.label}
              aria-pressed={bar.color.toUpperCase() === item.color}
              title={item.label}
              style={{ backgroundColor: item.color }}
              onClick={() => setBar({ color: item.color })}
            />
          ))}
          <ColorField label="Custom bar color" value={bar.color} onChange={(color) => color && setBar({ color })} />
        </div>
      </div>
      <div className="cf-field">
        <span>Fill</span>
        <div className="cf-segmented" role="radiogroup" aria-label="Bar fill">
          <button type="button" role="radio" aria-checked={bar.gradient} className={bar.gradient ? 'is-selected' : ''} onClick={() => setBar({ gradient: true })}>Gradient</button>
          <button type="button" role="radio" aria-checked={!bar.gradient} className={!bar.gradient ? 'is-selected' : ''} onClick={() => setBar({ gradient: false })}>Solid</button>
        </div>
      </div>
      <label className="cf-check">
        <input type="checkbox" checked={!bar.showValue} onChange={(event) => setBar({ showValue: !event.target.checked })} />
        <span><strong>Show bar only</strong><small>Hide the cell value behind the bar.</small></span>
      </label>
      <fieldset className="cf-group cf-points">
        <legend>Bar length</legend>
        <PointRow
          label="Minimum"
          types={['autoMin', 'min', 'num', 'percent', 'percentile', 'formula']}
          labels={BAR_TYPE_LABELS}
          type={bar.min.type}
          value={bar.min.value}
          onType={(type) => setBar({ min: { ...bar.min, type: type as BarPointType } })}
          onValue={(value) => setBar({ min: { ...bar.min, value } })}
        />
        <PointRow
          label="Maximum"
          types={['autoMax', 'max', 'num', 'percent', 'percentile', 'formula']}
          labels={BAR_TYPE_LABELS}
          type={bar.max.type}
          value={bar.max.value}
          onType={(type) => setBar({ max: { ...bar.max, type: type as BarPointType } })}
          onValue={(value) => setBar({ max: { ...bar.max, value } })}
        />
      </fieldset>
      <fieldset className="cf-group cf-grid-2">
        <legend>Negative values and axis</legend>
        <div className="cf-field">
          <span>Negative bar</span>
          <ColorField label="Negative bar color" value={bar.negativeColor} onChange={(color) => color && setBar({ negativeColor: color })} />
        </div>
        <label className="cf-field">
          <span>Axis</span>
          <select value={bar.axisPosition} onChange={(event) => setBar({ axisPosition: event.target.value as RuleDraft['bar']['axisPosition'] })}>
            <option value="automatic">Automatic</option>
            <option value="middle">Cell midpoint</option>
            <option value="none">None</option>
          </select>
        </label>
        <label className="cf-field cf-span-2">
          <span>Bar direction</span>
          <select value={bar.direction} onChange={(event) => setBar({ direction: event.target.value as RuleDraft['bar']['direction'] })}>
            <option value="context">Context</option>
            <option value="leftToRight">Left to right</option>
            <option value="rightToLeft">Right to left</option>
          </select>
        </label>
      </fieldset>
    </>
  )
}

function IconSetEditor({ draft, update }: EditorProps) {
  const { icons } = draft
  const size = iconSetSize(icons.set)
  const setIcons = (patch: Partial<RuleDraft['icons']>) => update({ icons: { ...icons, ...patch } })
  const setThreshold = (index: number, patch: Partial<ThresholdDraft>) => {
    setIcons({ thresholds: icons.thresholds.map((threshold, position) => (position === index ? { ...threshold, ...patch } : threshold)) })
  }
  const shown = (index: number) => (icons.reverse ? size - 1 - index : index)
  const gridRef = useRef<HTMLDivElement>(null)
  const moveSelection = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const keys: Record<string, number> = { ArrowRight: 1, ArrowDown: 2, ArrowLeft: -1, ArrowUp: -2 }
    const step = keys[event.key]
    if (!step) return
    event.preventDefault()
    const current = ICON_SET_NAMES.indexOf(icons.set)
    const next = ICON_SET_NAMES[(current + step + ICON_SET_NAMES.length) % ICON_SET_NAMES.length]
    update(withIconSet(draft, next))
    window.requestAnimationFrame(() => gridRef.current?.querySelector<HTMLElement>(`[data-set="${next}"]`)?.focus())
  }
  return (
    <>
      <div className="cf-field">
        <span>Icon style</span>
        <div ref={gridRef} className="cf-icon-grid" role="radiogroup" aria-label="Icon style" onKeyDown={moveSelection}>
          {ICON_SET_NAMES.map((set) => {
            const count = iconSetSize(set)
            const selected = set === icons.set
            return (
              <button
                key={set}
                type="button"
                role="radio"
                data-set={set}
                aria-checked={selected}
                tabIndex={selected ? 0 : -1}
                className={`cf-icon-set${selected ? ' is-selected' : ''}`}
                title={iconSetLabel(set)}
                aria-label={iconSetLabel(set)}
                onClick={() => update(withIconSet(draft, set))}
              >
                {Array.from({ length: count }, (_, offset) => count - 1 - offset).map((index) => <ConditionalIcon key={index} set={set} index={index} size={15} />)}
              </button>
            )
          })}
        </div>
      </div>
      <div className="cf-inline-checks">
        <label className="cf-check cf-check-compact">
          <input type="checkbox" checked={icons.reverse} onChange={(event) => setIcons({ reverse: event.target.checked })} />
          <span><strong>Reverse icon order</strong></span>
        </label>
        <label className="cf-check cf-check-compact">
          <input type="checkbox" checked={!icons.showValue} onChange={(event) => setIcons({ showValue: !event.target.checked })} />
          <span><strong>Show icon only</strong></span>
        </label>
      </div>
      <fieldset className="cf-group cf-thresholds">
        <legend>Display each icon when the value is</legend>
        {Array.from({ length: size - 1 }, (_, offset) => size - 1 - offset).map((index) => {
          const threshold = icons.thresholds[index - 1] ?? { type: 'percent', value: '', gte: true }
          return (
            <div key={index} className="cf-threshold-row" role="group" aria-label={`Icon ${size - index}`}>
              <ConditionalIcon set={icons.set} index={shown(index)} size={16} />
              <select aria-label="Comparison" value={threshold.gte ? 'gte' : 'gt'} onChange={(event) => setThreshold(index - 1, { gte: event.target.value === 'gte' })}>
                <option value="gte">≥</option>
                <option value="gt">&gt;</option>
              </select>
              <input aria-label="Threshold value" className="cf-mono" value={threshold.value} spellCheck={false} onChange={(event) => setThreshold(index - 1, { value: event.target.value })} />
              <select aria-label="Threshold type" value={threshold.type} onChange={(event) => setThreshold(index - 1, { type: event.target.value as ThresholdDraft['type'] })}>
                <option value="percent">Percent</option>
                <option value="num">Number</option>
                <option value="percentile">Percentile</option>
                <option value="formula">Formula</option>
              </select>
            </div>
          )
        })}
        <div className="cf-threshold-row is-rest">
          <ConditionalIcon set={icons.set} index={shown(0)} size={16} />
          <span>Otherwise (below {icons.thresholds[0]?.value || '…'}{icons.thresholds[0]?.type === 'percent' ? '%' : ''})</span>
        </div>
      </fieldset>
    </>
  )
}

// ---------------------------------------------------------------------------------------------
// Panel
// ---------------------------------------------------------------------------------------------

/**
 * Conditional-format rules manager (Excel's Rules Manager / Google Sheets' side panel): lists
 * the rules for the selection or the whole sheet in priority order, and edits one rule at a
 * time. Every committed change is reported through `onChange` with the complete next model.
 */
export function ConditionalFormatPanel({ sheetName, selectionRef, conditionalFormattings, onChange, onClose, cssColor }: ConditionalFormatPanelProps) {
  const [scope, setScope] = useState<Scope>('selection')
  const [editing, setEditing] = useState<EditingState | null>(null)
  const [attempted, setAttempted] = useState(false)
  const panelRef = useRef<HTMLElement>(null)
  const previousFocusRef = useRef<HTMLElement | null>(null)
  const pendingFocusRef = useRef<string | null>(null)
  const titleId = useId()
  const tabsId = useId()

  const model = useMemo(() => (Array.isArray(conditionalFormattings) ? conditionalFormattings : []), [conditionalFormattings])
  const allRules = useMemo(() => listRulesForRange(model), [model])
  const entries = useMemo(
    () => (scope === 'selection' ? listRulesForRange(model, selectionRef) : allRules),
    [allRules, model, scope, selectionRef],
  )

  // Initial focus on the first rule (or "Add rule"); restore the previous focus on close.
  useEffect(() => {
    previousFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const panel = panelRef.current
    const first = panel?.querySelector<HTMLElement>('[data-cf-action="edit"]:not(:disabled)') ?? panel?.querySelector<HTMLElement>('[data-cf-action="add"]')
    first?.focus()
    return () => {
      const previous = previousFocusRef.current
      if (previous?.isConnected) previous.focus()
    }
  }, [])

  // Move focus after list/editor transitions and list edits (targets are set by the actions).
  useLayoutEffect(() => {
    const target = pendingFocusRef.current
    const panel = panelRef.current
    if (!target || !panel) return
    pendingFocusRef.current = null
    let element = panel.querySelector<HTMLElement>(target)
    if (element && (element as HTMLButtonElement).disabled) element = element.closest('li')?.querySelector<HTMLElement>('[data-cf-action="edit"]:not(:disabled)') ?? null
    element ??= panel.querySelector<HTMLElement>('[data-cf-action="edit"]:not(:disabled)') ?? panel.querySelector<HTMLElement>('[data-cf-action="add"]')
    element?.focus()
  }, [editing, model, scope])

  const commit = useCallback((next: unknown[]) => onChange(next), [onChange])

  const startAdd = () => {
    const ref = parseRangeInput(selectionRef)
    setAttempted(false)
    pendingFocusRef.current = '[data-cf-field="range"]'
    setEditing({ id: null, draft: defaultRuleDraft('error' in ref ? selectionRef : ref.ref) })
  }

  const startEdit = (entry: ConditionalRuleEntry) => {
    setAttempted(false)
    pendingFocusRef.current = '[data-cf-field="range"]'
    setEditing({ id: entry.id, base: entry.rule, draft: ruleToDraft(entry.rule, entry.ref, cssColor) })
  }

  const leaveEditor = (focusId: number | null) => {
    pendingFocusRef.current = focusId !== null ? `[data-cf-rule="${focusId}"] [data-cf-action="edit"]` : '[data-cf-action="add"]'
    setEditing(null)
    setAttempted(false)
  }

  const save = () => {
    if (!editing) return
    setAttempted(true)
    const result = draftToRule(editing.draft, editing.base)
    if ('error' in result) return
    try {
      if (editing.id === null) {
        commit(addRule(model, result.ref, result.rule))
        leaveEditor(1)
      } else {
        commit(updateRule(model, editing.id, { ref: result.ref, rule: result.rule }))
        leaveEditor(editing.id)
      }
    } catch {
      // Validation already covers range errors; keep the editor open on anything unexpected.
    }
  }

  const move = (entry: ConditionalRuleEntry, neighbour: ConditionalRuleEntry | undefined, direction: 'up' | 'down') => {
    if (!neighbour) return
    pendingFocusRef.current = `[data-cf-rule="${neighbour.id}"] [data-cf-action="${direction}"]`
    commit(moveRulePriority(model, entry.id, neighbour.id - entry.id))
  }

  const remove = (entry: ConditionalRuleEntry, index: number) => {
    const next = entries[index + 1] ?? entries[index - 1]
    // Ids above the deleted rule shift down by one.
    const nextId = next ? (next.id > entry.id ? next.id - 1 : next.id) : null
    pendingFocusRef.current = nextId !== null ? `[data-cf-rule="${nextId}"] [data-cf-action="edit"]` : '[data-cf-action="add"]'
    commit(deleteRule(model, entry.id))
  }

  const clearRules = () => {
    pendingFocusRef.current = '[data-cf-action="add"]'
    if (scope === 'selection') {
      const range = parseRangeInput(selectionRef)
      if (!('error' in range)) commit(clearRulesInRange(model, range.ref))
    } else {
      commit([])
    }
  }

  const updateDraft = (patch: Partial<RuleDraft>) => {
    setEditing((current) => (current ? { ...current, draft: { ...current.draft, ...patch } } : current))
  }

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      if (editing) leaveEditor(editing.id)
      else onClose()
      return
    }
    if (editing && event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
      event.preventDefault()
      event.stopPropagation()
      save()
      return
    }
    if (event.key !== 'Tab' || !panelRef.current) return
    const controls = Array.from(panelRef.current.querySelectorAll<HTMLElement>(FOCUSABLE)).filter((element) => element.getClientRects().length > 0)
    if (!controls.length) return
    const first = controls[0]
    const last = controls[controls.length - 1]
    if (event.shiftKey && document.activeElement === first) {
      event.preventDefault()
      last.focus()
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault()
      first.focus()
    }
  }

  const onTabKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (!editing) return
    const order = FORMAT_TABS.map((tab) => tab.id)
    const index = order.indexOf(editing.draft.format)
    let next = -1
    if (event.key === 'ArrowRight') next = (index + 1) % order.length
    else if (event.key === 'ArrowLeft') next = (index - 1 + order.length) % order.length
    else if (event.key === 'Home') next = 0
    else if (event.key === 'End') next = order.length - 1
    if (next < 0) return
    event.preventDefault()
    updateDraft({ format: order[next] })
    window.requestAnimationFrame(() => panelRef.current?.querySelector<HTMLElement>(`[data-cf-tab="${order[next]}"]`)?.focus())
  }

  const draft = editing?.draft
  const liveError = draft ? validateDraft(draft) : null
  const rangeError = draft ? (() => { const parsed = parseRangeInput(draft.ref); return 'error' in parsed ? parsed.error : null })() : null
  const anchor = draft ? ruleAnchorAddress(('error' in parseRangeInput(draft.ref)) ? selectionRef : draft.ref) : 'A1'

  return (
    <div className="cf-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget && !editing) onClose() }}>
      <section ref={panelRef} className="cf-panel" role="dialog" aria-modal="true" aria-labelledby={titleId} onKeyDown={handleKeyDown}>
        <header className="cf-header">
          {editing ? (
            <button type="button" className="cf-icon-button" aria-label="Back to rules" title="Back to rules" onClick={() => leaveEditor(editing.id)}>
              <ArrowLeft size={15} aria-hidden="true" />
            </button>
          ) : (
            <span className="cf-header-icon" aria-hidden="true"><Paintbrush size={16} /></span>
          )}
          <div className="cf-header-copy">
            <h2 id={titleId}>{editing ? (editing.id === null ? 'New formatting rule' : 'Edit formatting rule') : 'Conditional formatting'}</h2>
            <p>{editing ? `${sheetName} · applies to ${draft?.ref || '…'}` : `${sheetName} · first rule wins where rules conflict`}</p>
          </div>
          <button type="button" className="cf-icon-button" aria-label="Close conditional formatting" title="Close" onClick={onClose}>
            <X size={15} aria-hidden="true" />
          </button>
        </header>

        {!editing && (
          <>
            <div className="cf-body">
              <div className="cf-scope" role="radiogroup" aria-label="Show rules for">
                <button type="button" role="radio" aria-checked={scope === 'selection'} className={scope === 'selection' ? 'is-selected' : ''} onClick={() => setScope('selection')}>
                  Selection <span className="cf-mono">{selectionRef}</span>
                </button>
                <button type="button" role="radio" aria-checked={scope === 'sheet'} className={scope === 'sheet' ? 'is-selected' : ''} onClick={() => setScope('sheet')}>
                  This sheet <span>{allRules.length}</span>
                </button>
              </div>

              {entries.length === 0 ? (
                <div className="cf-empty">
                  <span className="cf-empty-art" aria-hidden="true">
                    <span style={{ background: '#FFC7CE' }} />
                    <span style={{ background: '#FFEB9C' }} />
                    <span style={{ background: '#C6EFCE' }} />
                  </span>
                  <strong>{scope === 'selection' ? 'No rules apply to this selection' : 'This sheet has no rules yet'}</strong>
                  <p>Highlight values, show color scales, data bars or icons that update as the data changes.</p>
                </div>
              ) : (
                <ol className="cf-rule-list" aria-label="Rules in priority order">
                  {entries.map((entry, index) => {
                    const description = describeRule(entry.rule)
                    const supported = entry.kind !== null
                    return (
                      <li key={`${entry.id}:${entry.ref}`} className="cf-rule" data-cf-rule={entry.id}>
                        <button
                          type="button"
                          className="cf-rule-main"
                          data-cf-action="edit"
                          disabled={!supported}
                          aria-label={`Edit rule ${entry.id}: ${description}, applies to ${entry.ref}`}
                          onClick={() => startEdit(entry)}
                        >
                          <RulePreview rule={entry.rule} cssColor={cssColor} />
                          <span className="cf-rule-copy">
                            <strong>{description}</strong>
                            <small>
                              <span className="cf-mono">{entry.ref}</span>
                              {entry.rule.stopIfTrue && ruleUsesStyle(entry.kind) ? <span className="cf-tag">Stop if true</span> : null}
                            </small>
                          </span>
                        </button>
                        <div className="cf-rule-actions">
                          <button type="button" className="cf-icon-button" data-cf-action="up" aria-label="Move rule up" title="Move up (higher priority)" disabled={index === 0} onClick={() => move(entry, entries[index - 1], 'up')}>
                            <ChevronUp size={14} aria-hidden="true" />
                          </button>
                          <button type="button" className="cf-icon-button" data-cf-action="down" aria-label="Move rule down" title="Move down (lower priority)" disabled={index === entries.length - 1} onClick={() => move(entry, entries[index + 1], 'down')}>
                            <ChevronDown size={14} aria-hidden="true" />
                          </button>
                          <button type="button" className="cf-icon-button cf-danger" data-cf-action="delete" aria-label="Delete rule" title="Delete rule" onClick={() => remove(entry, index)}>
                            <Trash2 size={14} aria-hidden="true" />
                          </button>
                        </div>
                      </li>
                    )
                  })}
                </ol>
              )}
            </div>
            <footer className="cf-footer">
              <button type="button" className="secondary-action cf-footer-button" disabled={!entries.length} onClick={clearRules}>
                <Eraser size={14} aria-hidden="true" />
                {scope === 'selection' ? 'Clear from selection' : 'Clear all rules'}
              </button>
              <button type="button" className="primary-action cf-footer-button" data-cf-action="add" onClick={startAdd}>
                <Plus size={14} aria-hidden="true" />
                Add rule
              </button>
            </footer>
          </>
        )}

        {editing && draft && (
          <>
            <div className="cf-body cf-editor">
              <label className="cf-field">
                <span>Apply to range</span>
                <span className="cf-inline">
                  <input
                    data-cf-field="range"
                    className="cf-mono"
                    value={draft.ref}
                    spellCheck={false}
                    aria-invalid={rangeError ? true : undefined}
                    placeholder="A1:D20, F1:F20"
                    onChange={(event) => updateDraft({ ref: event.target.value })}
                  />
                  <button type="button" className="cf-chip-button" title="Use the current selection" onClick={() => updateDraft({ ref: selectionRef })}>
                    Use selection
                  </button>
                </span>
              </label>

              <div className="cf-tabs" role="tablist" aria-label="Format style" onKeyDown={onTabKey}>
                {FORMAT_TABS.map((tab) => (
                  <button
                    key={tab.id}
                    type="button"
                    role="tab"
                    id={`${tabsId}-${tab.id}`}
                    data-cf-tab={tab.id}
                    aria-selected={draft.format === tab.id}
                    aria-controls={`${tabsId}-panel`}
                    tabIndex={draft.format === tab.id ? 0 : -1}
                    className={draft.format === tab.id ? 'is-selected' : ''}
                    onClick={() => updateDraft({ format: tab.id })}
                  >
                    {tab.label}
                  </button>
                ))}
              </div>

              <div className="cf-tabpanel" role="tabpanel" id={`${tabsId}-panel`} aria-labelledby={`${tabsId}-${draft.format}`}>
                {draft.format === 'single' && <SingleColorEditor draft={draft} update={updateDraft} anchor={anchor} />}
                {draft.format === 'colorScale' && <ColorScaleEditor draft={draft} update={updateDraft} anchor={anchor} />}
                {draft.format === 'dataBar' && <DataBarEditor draft={draft} update={updateDraft} anchor={anchor} />}
                {draft.format === 'iconSet' && <IconSetEditor draft={draft} update={updateDraft} anchor={anchor} />}
              </div>

              {attempted && liveError && (
                <p className="cf-error" role="alert">
                  <CircleAlert size={13} aria-hidden="true" />
                  {liveError}
                </p>
              )}
            </div>
            <footer className="cf-footer">
              <span className="cf-hint">Ctrl+Enter to save</span>
              <button type="button" className="secondary-action cf-footer-button" onClick={() => leaveEditor(editing.id)}>Cancel</button>
              <button type="button" className="primary-action cf-footer-button" onClick={save}>
                <Check size={14} aria-hidden="true" />
                Done
              </button>
            </footer>
          </>
        )}
      </section>
    </div>
  )
}

export default ConditionalFormatPanel

export type { IconSetName }
