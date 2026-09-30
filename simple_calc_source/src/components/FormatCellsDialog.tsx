import { useEffect, useId, useMemo, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, MouseEvent as ReactMouseEvent, PointerEvent as ReactPointerEvent, ReactNode, RefObject } from 'react'
import { CircleAlert, X } from 'lucide-react'
import type { CellAlignment, CellBorderSide, CellFill, CellFont, CellStyle, SpreadsheetColor } from '../spreadsheet-types'
import { formatScalarDetailed } from '../lib/number-format'
import {
  CURRENCY_SYMBOLS, CUSTOM_FORMAT_PRESETS, DATE_FORMATS, FRACTION_FORMATS, MAX_DECIMALS, NEGATIVE_NUMBER_STYLES,
  NUMBER_FORMAT_CATEGORIES, SAMPLE_DATE_SERIAL, SPECIAL_FORMATS, TIME_FORMATS, buildNumberFormat, detectNumberFormat,
  formatTypeSample, validateNumberFormat,
} from '../lib/format-codes'
import type { FormatTypeOption, NegativeNumberStyle, NumberFormatCategory } from '../lib/format-codes'
import {
  DEFAULT_THEME_COLORS, PATTERN_FILLS, argbColor, fillPreviewCss, resolveColorHex, sameColor,
} from '../lib/cell-styles'
import type { ColorValue, FormatBorderChange, FormatCellsChange, NullablePartial } from '../lib/cell-styles'
import { DEFAULT_FONT, FONT_FAMILIES, FONT_SIZES, fontInfo, fontStack, isFontInstalled, isSymbolFont, parseFontSize } from '../lib/fonts'
import { ColorDropdownButton, ColorPicker } from './ColorPicker'
import { BorderLine, LINE_STYLE_GALLERY, LINE_STYLE_SPECS, LineStyleSample } from './BorderPicker'
import './format-ui.css'

export type { CellPosition, FormatBorderChange, FormatCellsChange, NullablePartial } from '../lib/cell-styles'
export { applyFormatChange, applyFormatChangeToCell, applyNeighborBorderChange, cellPosition } from '../lib/cell-styles'

export type FormatCellsTab = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection'

export interface FormatCellsDialogProps {
  initialTab?: FormatCellsTab
  /** Active cell's current style, with the effective number format merged into `numFmt`. */
  style: CellStyle
  /** Active cell's value, used for the Number tab's live sample. */
  sampleValue: string | number | boolean | null
  /** Selection spans several rows / columns: enables the inside borders. */
  multiRow: boolean
  multiColumn: boolean
  /** Workbook theme palette, 12 hex strings in OOXML theme order. */
  themeColors: string[]
  /** Called on OK with only what the user changed (skipped when nothing changed), then onClose runs. */
  onApply: (change: FormatCellsChange) => void
  onClose: () => void
  /** Whether the selection is merged now (Alignment → Merge cells). */
  merged?: boolean
  /** Normal-style font for "Normal font" and for cells without an explicit font. */
  defaultFont?: { name: string; size: number }
}

const TABS: readonly { id: FormatCellsTab; label: string }[] = [
  { id: 'number', label: 'Number' },
  { id: 'alignment', label: 'Alignment' },
  { id: 'font', label: 'Font' },
  { id: 'border', label: 'Border' },
  { id: 'fill', label: 'Fill' },
  { id: 'protection', label: 'Protection' },
]

const FOCUSABLE = 'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'
const clamp = (value: number, min: number, max: number) => Math.min(max, Math.max(min, value))
const AUTOMATIC_BLACK: SpreadsheetColor = { argb: 'FF000000' }

function toColorObject(value: ColorValue): SpreadsheetColor {
  return typeof value === 'string' ? argbColor(value) : value
}

// ---------------------------------------------------------------------------------------
// Listbox
// ---------------------------------------------------------------------------------------

interface ListboxItem {
  key: string
  label: ReactNode
  /** Plain text for type-ahead. */
  text: string
  style?: CSSProperties
  title?: string
  muted?: boolean
}

function Listbox({ label, items, selectedKey, scrollKey, onSelect, className, labelledBy, id: providedId }: {
  label?: string
  labelledBy?: string
  items: ListboxItem[]
  selectedKey: string | null
  /** Scrolled into view when nothing is selected (e.g. a typed prefix match). */
  scrollKey?: string | null
  onSelect: (key: string) => void
  className?: string
  id?: string
}) {
  const generated = useId()
  const id = providedId || generated
  const ref = useRef<HTMLDivElement>(null)
  const typeahead = useRef({ text: '', time: 0 })
  const index = items.findIndex((item) => item.key === selectedKey)
  const scrollIndex = index >= 0 ? index : items.findIndex((item) => item.key === scrollKey)

  useEffect(() => {
    const list = ref.current
    if (!list || scrollIndex < 0) return
    const option = list.querySelector<HTMLElement>(`[data-index="${scrollIndex}"]`)
    if (!option) return
    const top = option.offsetTop
    const bottom = top + option.offsetHeight
    if (top < list.scrollTop) list.scrollTop = top
    else if (bottom > list.scrollTop + list.clientHeight) list.scrollTop = bottom - list.clientHeight
  }, [scrollIndex])

  const move = (next: number) => {
    if (!items.length) return
    const item = items[clamp(next, 0, items.length - 1)]
    if (item) onSelect(item.key)
  }

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const page = Math.max(1, Math.floor((ref.current?.clientHeight || 160) / 20) - 1)
    switch (event.key) {
      case 'ArrowDown': event.preventDefault(); move(index < 0 ? 0 : index + 1); return
      case 'ArrowUp': event.preventDefault(); move(index < 0 ? 0 : index - 1); return
      case 'Home': event.preventDefault(); move(0); return
      case 'End': event.preventDefault(); move(items.length - 1); return
      case 'PageDown': event.preventDefault(); move((index < 0 ? 0 : index) + page); return
      case 'PageUp': event.preventDefault(); move((index < 0 ? 0 : index) - page); return
      default: break
    }
    if (event.key.length !== 1 || event.ctrlKey || event.metaKey || event.altKey || event.key === ' ') return
    const now = Date.now()
    const buffer = now - typeahead.current.time < 700 ? typeahead.current.text + event.key : event.key
    typeahead.current = { text: buffer, time: now }
    const needle = buffer.toLocaleLowerCase()
    const start = buffer.length === 1 ? index + 1 : Math.max(0, index)
    for (let offset = 0; offset < items.length; offset += 1) {
      const candidate = items[(start + offset) % items.length]
      if (candidate.text.toLocaleLowerCase().startsWith(needle)) {
        event.preventDefault()
        onSelect(candidate.key)
        return
      }
    }
  }

  return (
    <div
      ref={ref}
      id={id}
      className={`fmt-listbox${className ? ` ${className}` : ''}`}
      role="listbox"
      tabIndex={0}
      aria-label={labelledBy ? undefined : label}
      aria-labelledby={labelledBy}
      aria-activedescendant={index >= 0 ? `${id}-option-${index}` : undefined}
      onKeyDown={handleKeyDown}
    >
      {items.map((item, itemIndex) => (
        <div
          key={item.key}
          id={`${id}-option-${itemIndex}`}
          data-index={itemIndex}
          role="option"
          aria-selected={itemIndex === index}
          className={`fmt-option${itemIndex === index ? ' is-selected' : ''}${item.muted ? ' is-muted' : ''}${itemIndex === scrollIndex && index < 0 ? ' is-hint' : ''}`}
          style={item.style}
          title={item.title}
          onMouseDown={(event) => {
            event.preventDefault()
            ref.current?.focus()
            onSelect(item.key)
          }}
        >
          {item.label}
        </div>
      ))}
    </div>
  )
}

function Group({ legend, children, className }: { legend: string; children: ReactNode; className?: string }) {
  return (
    <fieldset className={`fmt-group${className ? ` ${className}` : ''}`}>
      <legend>{legend}</legend>
      {children}
    </fieldset>
  )
}

function NumberInput({ label, value, min, max, step = 1, onChange, disabled, id, suffix }: {
  label: string; value: number; min: number; max: number; step?: number; onChange: (value: number) => void; disabled?: boolean; id?: string; suffix?: string
}) {
  const [draft, setDraft] = useState(String(value))
  useEffect(() => { setDraft(String(value)) }, [value])
  const commit = (text: string) => {
    const number = Number(text)
    if (text.trim() === '' || !Number.isFinite(number)) { setDraft(String(value)); return }
    const next = clamp(Math.round(number / step) * step, min, max)
    setDraft(String(next))
    if (next !== value) onChange(next)
  }
  return (
    <label className={`fmt-field fmt-number-field${disabled ? ' is-disabled' : ''}`}>
      <span>{label}</span>
      <span className="fmt-number-wrap">
        <input
          id={id}
          type="number"
          min={min}
          max={max}
          step={step}
          value={draft}
          disabled={disabled}
          onChange={(event) => {
            setDraft(event.target.value)
            const number = Number(event.target.value)
            if (event.target.value.trim() !== '' && Number.isFinite(number) && number >= min && number <= max) onChange(number)
          }}
          onBlur={(event) => commit(event.target.value)}
        />
        {suffix && <small>{suffix}</small>}
      </span>
    </label>
  )
}

// ---------------------------------------------------------------------------------------
// Dialog state
// ---------------------------------------------------------------------------------------

type EdgeKey = 'top' | 'bottom' | 'left' | 'right' | 'insideHorizontal' | 'insideVertical' | 'diagonalUp' | 'diagonalDown'
const EDGE_KEYS: readonly EdgeKey[] = ['top', 'bottom', 'left', 'right', 'insideHorizontal', 'insideVertical', 'diagonalUp', 'diagonalDown']

type FontStyleKey = 'regular' | 'italic' | 'bold' | 'boldItalic'
const FONT_STYLES: readonly { id: FontStyleKey; label: string; css: CSSProperties }[] = [
  { id: 'regular', label: 'Regular', css: {} },
  { id: 'italic', label: 'Italic', css: { fontStyle: 'italic' } },
  { id: 'bold', label: 'Bold', css: { fontWeight: 700 } },
  { id: 'boldItalic', label: 'Bold Italic', css: { fontWeight: 700, fontStyle: 'italic' } },
]

type UnderlineKey = 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting'
const UNDERLINES: readonly { id: UnderlineKey; label: string }[] = [
  { id: 'none', label: 'None' },
  { id: 'single', label: 'Single' },
  { id: 'double', label: 'Double' },
  { id: 'singleAccounting', label: 'Single Accounting' },
  { id: 'doubleAccounting', label: 'Double Accounting' },
]

const HORIZONTAL_OPTIONS = [
  { id: 'general', label: 'General' },
  { id: 'left', label: 'Left (Indent)' },
  { id: 'center', label: 'Center' },
  { id: 'right', label: 'Right (Indent)' },
  { id: 'fill', label: 'Fill' },
  { id: 'justify', label: 'Justify' },
  { id: 'centerContinuous', label: 'Center Across Selection' },
  { id: 'distributed', label: 'Distributed (Indent)' },
]

const VERTICAL_OPTIONS = [
  { id: 'top', label: 'Top' },
  { id: 'middle', label: 'Center' },
  { id: 'bottom', label: 'Bottom' },
  { id: 'justify', label: 'Justify' },
  { id: 'distributed', label: 'Distributed' },
]

type ReadingOrder = 'context' | 'ltr' | 'rtl'

interface AlignmentState {
  horizontal: string
  vertical: string
  indent: number
  wrapText: boolean
  shrinkToFit: boolean
  justifyLastLine: boolean
  degrees: number
  verticalText: boolean
  readingOrder: ReadingOrder
  merge: boolean
}

interface FontState {
  name: string
  style: FontStyleKey
  sizeText: string
  underline: UnderlineKey
  color: ColorValue | null
  strike: boolean
  vertAlign: '' | 'superscript' | 'subscript'
}

interface NumberState {
  category: NumberFormatCategory
  decimals: number
  thousands: boolean
  negative: NegativeNumberStyle
  symbol: string
  types: Record<'date' | 'time' | 'fraction' | 'special', string>
  custom: string
}

function initialAlignment(alignment: CellAlignment | undefined, merged: boolean): AlignmentState {
  const horizontal = String(alignment?.horizontal || 'general')
  const vertical = String(alignment?.vertical || 'bottom')
  const rotation = alignment?.textRotation as number | 'vertical' | undefined
  const verticalText = rotation === 'vertical' || rotation === 255
  let degrees = typeof rotation === 'number' && !verticalText ? rotation : 0
  if (degrees > 90 && degrees <= 180) degrees = 90 - degrees
  const order = alignment?.readingOrder
  return {
    horizontal: HORIZONTAL_OPTIONS.some((option) => option.id === horizontal) ? horizontal : 'general',
    vertical: vertical === 'center' ? 'middle' : VERTICAL_OPTIONS.some((option) => option.id === vertical) ? vertical : 'bottom',
    indent: Math.max(0, Number(alignment?.indent) || 0),
    wrapText: Boolean(alignment?.wrapText),
    shrinkToFit: Boolean(alignment?.shrinkToFit),
    justifyLastLine: Boolean(alignment?.justifyLastLine),
    degrees: clamp(Math.round(degrees), -90, 90),
    verticalText,
    readingOrder: order === 'rtl' || order === 2 ? 'rtl' : order === 'ltr' || order === 1 ? 'ltr' : 'context',
    merge: merged,
  }
}

function initialFont(font: CellFont | undefined, fallback: { name: string; size: number }): FontState {
  const underline = font?.underline
  return {
    name: String(font?.name || fallback.name),
    style: font?.bold && font?.italic ? 'boldItalic' : font?.bold ? 'bold' : font?.italic ? 'italic' : 'regular',
    sizeText: String(font?.size || fallback.size),
    underline: underline === true ? 'single' : typeof underline === 'string' && UNDERLINES.some((option) => option.id === underline) ? underline as UnderlineKey : 'none',
    color: font?.color ?? null,
    strike: Boolean(font?.strike),
    vertAlign: font?.vertAlign === 'superscript' || font?.vertAlign === 'subscript' ? font.vertAlign : '',
  }
}

function initialNumber(code: string): NumberState {
  const detected = detectNumberFormat(code)
  const typeFor = (category: NumberFormatCategory, list: readonly FormatTypeOption[]) => (
    detected.category === category && detected.options.type ? detected.options.type : list[0].id
  )
  return {
    category: detected.category,
    decimals: detected.options.decimals,
    thousands: detected.options.thousands,
    negative: detected.options.negative,
    symbol: detected.options.symbol,
    types: {
      date: typeFor('date', DATE_FORMATS),
      time: typeFor('time', TIME_FORMATS),
      fraction: typeFor('fraction', FRACTION_FORMATS),
      special: typeFor('special', SPECIAL_FORMATS),
    },
    custom: code,
  }
}

function numberCode(state: NumberState) {
  if (state.category === 'custom') return state.custom
  const type = state.category === 'date' || state.category === 'time' || state.category === 'fraction' || state.category === 'special'
    ? state.types[state.category] : undefined
  return buildNumberFormat(state.category, { decimals: state.decimals, thousands: state.thousands, negative: state.negative, symbol: state.symbol, type })
}

interface FillState {
  background: ColorValue | null
  pattern: string
  patternColor: ColorValue | null
  gradient: boolean
}

function initialFill(fill: CellFill | undefined): FillState {
  if (!fill) return { background: null, pattern: 'none', patternColor: null, gradient: false }
  if (fill.type === 'gradient') return { background: null, pattern: 'none', patternColor: null, gradient: true }
  const pattern = String(fill.pattern || (fill.fgColor || fill.color ? 'solid' : 'none'))
  if (pattern === 'none') return { background: null, pattern: 'none', patternColor: null, gradient: false }
  if (pattern === 'solid') return { background: fill.fgColor ?? fill.color ?? null, pattern: 'none', patternColor: null, gradient: false }
  return {
    background: fill.bgColor ?? null,
    pattern: PATTERN_FILLS.some((entry) => entry.id === pattern) ? pattern : 'none',
    patternColor: fill.fgColor ?? null,
    gradient: false,
  }
}

function buildFill(state: FillState): CellFill | null {
  if (state.pattern === 'none') return state.background ? { type: 'pattern', pattern: 'solid', fgColor: toColorObject(state.background) } : null
  return {
    type: 'pattern',
    pattern: state.pattern,
    fgColor: state.patternColor ? toColorObject(state.patternColor) : AUTOMATIC_BLACK,
    ...(state.background ? { bgColor: toColorObject(state.background) } : {}),
  }
}

function sameSide(a: CellBorderSide | null | undefined, b: CellBorderSide | null | undefined, themeColors: readonly string[]) {
  if (!a || !b) return !a && !b
  return a.style === b.style && resolveColorHex(a.color ?? AUTOMATIC_BLACK, themeColors) === resolveColorHex(b.color ?? AUTOMATIC_BLACK, themeColors)
}

const drawable = (side: CellBorderSide | undefined): CellBorderSide | null => (side?.style && side.style !== 'none' ? side : null)

// ---------------------------------------------------------------------------------------
// Dialog
// ---------------------------------------------------------------------------------------

export function FormatCellsDialog({
  initialTab = 'number',
  style,
  sampleValue,
  multiRow,
  multiColumn,
  themeColors: providedTheme,
  onApply,
  onClose,
  merged = false,
  defaultFont = DEFAULT_FONT,
}: FormatCellsDialogProps) {
  const themeColors = providedTheme?.length ? providedTheme : DEFAULT_THEME_COLORS
  const baseId = useId()
  const [tab, setTab] = useState<FormatCellsTab>(initialTab)
  const dialogRef = useRef<HTMLElement>(null)
  const panelRef = useRef<HTMLDivElement>(null)
  const customInputRef = useRef<HTMLInputElement>(null)
  const sizeInputRef = useRef<HTMLInputElement>(null)
  const previousFocusRef = useRef<HTMLElement | null>(null)
  const [error, setError] = useState('')

  // Snapshot of the incoming style: later prop changes must not reset the user's edits.
  const initial = useMemo(() => {
    const code = String(style.numFmt || 'General').trim() || 'General'
    const border = style.border || {}
    const edges: Record<EdgeKey, CellBorderSide | null> = {
      top: drawable(border.top),
      bottom: drawable(border.bottom),
      left: drawable(border.left),
      right: drawable(border.right),
      insideHorizontal: multiRow ? drawable(border.bottom) : null,
      insideVertical: multiColumn ? drawable(border.right) : null,
      diagonalUp: border.diagonalUp ? drawable(border.diagonal) : null,
      diagonalDown: border.diagonalDown ? drawable(border.diagonal) : null,
    }
    return {
      code,
      number: initialNumber(code),
      alignment: initialAlignment(style.alignment, merged),
      font: initialFont(style.font, defaultFont),
      fill: initialFill(style.fill),
      edges,
      protection: { locked: style.protection?.locked !== false, hidden: Boolean(style.protection?.hidden) },
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const [numberState, setNumberState] = useState<NumberState>(initial.number)
  const [numberTouched, setNumberTouched] = useState(false)
  const [alignment, setAlignment] = useState<AlignmentState>(initial.alignment)
  const [font, setFont] = useState<FontState>(initial.font)
  const [edges, setEdges] = useState(initial.edges)
  const [dirtyEdges, setDirtyEdges] = useState<ReadonlySet<EdgeKey>>(new Set())
  const [lineStyle, setLineStyle] = useState<string | null>('thin')
  const [lineColor, setLineColor] = useState<ColorValue | null>(null)
  const [fill, setFill] = useState<FillState>(initial.fill)
  const [fillTouched, setFillTouched] = useState(false)
  const [protection, setProtection] = useState(initial.protection)

  const code = numberCode(numberState)
  const validation = useMemo(() => (numberState.category === 'custom' ? validateNumberFormat(code) : { valid: true }), [code, numberState.category])

  useEffect(() => {
    previousFocusRef.current = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const frame = window.requestAnimationFrame(() => {
      const target = panelRef.current?.querySelector<HTMLElement>(FOCUSABLE)
      target?.focus()
    })
    return () => {
      window.cancelAnimationFrame(frame)
      const previous = previousFocusRef.current
      if (previous?.isConnected) previous.focus()
    }
  }, [])

  const updateNumber = (patch: Partial<NumberState>) => {
    setNumberTouched(true)
    setError('')
    setNumberState((current) => ({ ...current, ...patch }))
  }

  const selectCategory = (category: NumberFormatCategory) => {
    setNumberTouched(true)
    setError('')
    setNumberState((current) => {
      if (category === current.category) return current
      // Custom starts from the code the previous category produced, as in Excel.
      return category === 'custom' ? { ...current, category, custom: numberCode(current) } : { ...current, category }
    })
  }

  const updateAlignment = (patch: Partial<AlignmentState>) => setAlignment((current) => ({ ...current, ...patch }))
  const updateFont = (patch: Partial<FontState>) => { setError(''); setFont((current) => ({ ...current, ...patch })) }
  const updateFill = (patch: Partial<FillState>) => { setFillTouched(true); setFill((current) => ({ ...current, ...patch, gradient: false })) }

  const activeSide = (): CellBorderSide | null => (lineStyle ? { style: lineStyle, color: lineColor ? toColorObject(lineColor) : AUTOMATIC_BLACK } : null)

  const setEdgeValues = (values: Partial<Record<EdgeKey, CellBorderSide | null>>) => {
    setEdges((current) => ({ ...current, ...values }))
    setDirtyEdges((current) => new Set([...current, ...(Object.keys(values) as EdgeKey[])]))
  }

  const edgeEnabled = (edge: EdgeKey) => (edge === 'insideHorizontal' ? multiRow : edge === 'insideVertical' ? multiColumn : true)

  const toggleEdge = (edge: EdgeKey) => {
    if (!edgeEnabled(edge)) return
    const side = activeSide()
    const current = edges[edge]
    setEdgeValues({ [edge]: current && (!side || sameSide(current, side, themeColors)) ? null : side })
  }

  const applyBorderPresetButton = (preset: 'none' | 'outline' | 'inside') => {
    const side = activeSide()
    if (preset === 'none') {
      setEdgeValues(Object.fromEntries(EDGE_KEYS.filter(edgeEnabled).map((edge) => [edge, null])))
    } else if (preset === 'outline') {
      setEdgeValues({ top: side, bottom: side, left: side, right: side })
    } else {
      const values: Partial<Record<EdgeKey, CellBorderSide | null>> = {}
      if (multiRow) values.insideHorizontal = side
      if (multiColumn) values.insideVertical = side
      setEdgeValues(values)
    }
  }

  const computeChange = (): FormatCellsChange | null => {
    const change: FormatCellsChange = {}

    if (numberTouched && code.trim() !== initial.code) change.numFmt = code.trim()

    const a = alignment
    const i = initial.alignment
    const align: NullablePartial<CellAlignment> = {}
    if (a.horizontal !== i.horizontal) align.horizontal = a.horizontal === 'general' ? null : a.horizontal
    if (a.vertical !== i.vertical) align.vertical = a.vertical === 'bottom' ? null : a.vertical
    if (a.indent !== i.indent) align.indent = a.indent > 0 ? a.indent : null
    if (a.wrapText !== i.wrapText) align.wrapText = a.wrapText || null
    if (a.shrinkToFit !== i.shrinkToFit) align.shrinkToFit = a.shrinkToFit || null
    if (a.justifyLastLine !== i.justifyLastLine) align.justifyLastLine = a.justifyLastLine || null
    if (a.verticalText !== i.verticalText || (!a.verticalText && a.degrees !== i.degrees)) {
      align.textRotation = a.verticalText ? 'vertical' : a.degrees !== 0 ? a.degrees : null
    }
    if (a.readingOrder !== i.readingOrder) align.readingOrder = a.readingOrder === 'context' ? null : a.readingOrder
    if (Object.keys(align).length) change.alignment = align
    if (a.merge !== i.merge) change.merge = a.merge

    const f = font
    const fi = initial.font
    const fontChange: NullablePartial<CellFont> = {}
    const typedName = f.name.trim()
    if (typedName && typedName.toLocaleLowerCase() !== fi.name.toLocaleLowerCase()) fontChange.name = fontInfo(typedName)?.name || typedName
    const size = parseFontSize(f.sizeText)
    if (size === null) return null
    if (size !== parseFontSize(fi.sizeText)) fontChange.size = size
    const bold = f.style === 'bold' || f.style === 'boldItalic'
    const italic = f.style === 'italic' || f.style === 'boldItalic'
    if (bold !== (fi.style === 'bold' || fi.style === 'boldItalic')) fontChange.bold = bold || null
    if (italic !== (fi.style === 'italic' || fi.style === 'boldItalic')) fontChange.italic = italic || null
    if (f.underline !== fi.underline) fontChange.underline = f.underline === 'none' ? null : f.underline === 'single' ? true : f.underline
    if (!sameColor(f.color, fi.color, themeColors)) fontChange.color = f.color ? toColorObject(f.color) : null
    if (f.strike !== fi.strike) fontChange.strike = f.strike || null
    if (f.vertAlign !== fi.vertAlign) fontChange.vertAlign = f.vertAlign || null
    if (Object.keys(fontChange).length) change.font = fontChange

    if (dirtyEdges.size) {
      const borders: FormatBorderChange = {}
      // For a range the dialog shows the active cell's borders as a guess, so every edge the
      // user touched is written; for one cell, edges that end up as they started are skipped.
      const guessed = multiRow || multiColumn
      dirtyEdges.forEach((edge) => {
        if (!edgeEnabled(edge)) return
        const value = edges[edge]
        if (!guessed && sameSide(value, initial.edges[edge], themeColors)) return
        borders[edge] = value
      })
      if (Object.keys(borders).length) change.borders = borders
    }

    if (fillTouched) change.fill = buildFill(fill)

    if (protection.locked !== initial.protection.locked || protection.hidden !== initial.protection.hidden) {
      change.protection = {}
      if (protection.locked !== initial.protection.locked) change.protection.locked = protection.locked
      if (protection.hidden !== initial.protection.hidden) change.protection.hidden = protection.hidden
    }
    return change
  }

  const submit = () => {
    if (numberTouched && !validation.valid) {
      setTab('number')
      setError(validation.message || 'This number format code isn’t valid.')
      window.requestAnimationFrame(() => customInputRef.current?.focus())
      return
    }
    const change = computeChange()
    if (!change) {
      setTab('font')
      setError('Font size must be a number between 1 and 409.')
      window.requestAnimationFrame(() => sizeInputRef.current?.focus())
      return
    }
    if (Object.keys(change).length) onApply(change)
    onClose()
  }

  const switchTab = (next: FormatCellsTab, focusTab = false) => {
    setTab(next)
    if (focusTab) window.requestAnimationFrame(() => document.getElementById(`${baseId}-tab-${next}`)?.focus())
  }

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLElement>) => {
    const control = event.ctrlKey || event.metaKey
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (event.key === 'Enter' && !event.shiftKey && !event.altKey && !control) {
      const target = event.target as HTMLElement
      if (target.tagName === 'BUTTON' || target.tagName === 'TEXTAREA') return
      event.preventDefault()
      event.stopPropagation()
      submit()
      return
    }
    if (control && (event.key === 'Tab' || event.key === 'PageDown' || event.key === 'PageUp')) {
      event.preventDefault()
      event.stopPropagation()
      const index = TABS.findIndex((entry) => entry.id === tab)
      const backwards = event.key === 'PageUp' || (event.key === 'Tab' && event.shiftKey)
      switchTab(TABS[(index + (backwards ? TABS.length - 1 : 1)) % TABS.length].id, true)
      return
    }
    if (control && ['o', 's', 'p', 'n', 'e'].includes(event.key.toLocaleLowerCase())) {
      event.preventDefault()
      return
    }
    if (event.key !== 'Tab' || !dialogRef.current) return
    const controls = Array.from(dialogRef.current.querySelectorAll<HTMLElement>(FOCUSABLE)).filter((element) => element.getClientRects().length > 0)
    if (!controls.length) return
    const first = controls[0]
    const last = controls[controls.length - 1]
    if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
    else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
  }

  const handleTabKey = (event: ReactKeyboardEvent<HTMLButtonElement>) => {
    const index = TABS.findIndex((entry) => entry.id === tab)
    let next = -1
    if (event.key === 'ArrowRight') next = (index + 1) % TABS.length
    else if (event.key === 'ArrowLeft') next = (index + TABS.length - 1) % TABS.length
    else if (event.key === 'Home') next = 0
    else if (event.key === 'End') next = TABS.length - 1
    if (next < 0) return
    event.preventDefault()
    switchTab(TABS[next].id, true)
  }

  return (
    <div className="fmt-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onClose() }}>
      <section
        ref={dialogRef}
        className="fmt-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${baseId}-title`}
        onKeyDown={handleKeyDown}
      >
        <header className="fmt-dialog-header">
          <h2 id={`${baseId}-title`}>Format Cells</h2>
          <button type="button" className="fmt-icon-button fmt-close" aria-label="Close" onClick={onClose}><X size={15} /></button>
        </header>
        <div className="fmt-tabs" role="tablist" aria-label="Format categories">
          {TABS.map((entry) => (
            <button
              key={entry.id}
              id={`${baseId}-tab-${entry.id}`}
              type="button"
              role="tab"
              className={`fmt-tab${tab === entry.id ? ' is-active' : ''}`}
              aria-selected={tab === entry.id}
              aria-controls={`${baseId}-panel`}
              tabIndex={tab === entry.id ? 0 : -1}
              onClick={() => switchTab(entry.id)}
              onKeyDown={handleTabKey}
            >
              {entry.label}
            </button>
          ))}
        </div>
        <div ref={panelRef} id={`${baseId}-panel`} className="fmt-panel" role="tabpanel" aria-labelledby={`${baseId}-tab-${tab}`}>
          {tab === 'number' && (
            <NumberPanel
              state={numberState}
              code={code}
              validation={validation}
              sampleValue={sampleValue}
              initialCode={initial.code}
              customInputRef={customInputRef}
              onCategory={selectCategory}
              onChange={updateNumber}
            />
          )}
          {tab === 'alignment' && (
            <AlignmentPanel state={alignment} onChange={updateAlignment} canMerge={multiRow || multiColumn || initial.alignment.merge} />
          )}
          {tab === 'font' && (
            <FontPanel state={font} initial={initial.font} defaultFont={defaultFont} themeColors={themeColors} sizeInputRef={sizeInputRef} onChange={updateFont} />
          )}
          {tab === 'border' && (
            <BorderPanel
              edges={edges}
              lineStyle={lineStyle}
              lineColor={lineColor}
              themeColors={themeColors}
              multiRow={multiRow}
              multiColumn={multiColumn}
              onLineStyle={setLineStyle}
              onLineColor={setLineColor}
              onToggle={toggleEdge}
              onPreset={applyBorderPresetButton}
            />
          )}
          {tab === 'fill' && <FillPanel state={fill} themeColors={themeColors} onChange={updateFill} />}
          {tab === 'protection' && <ProtectionPanel state={protection} onChange={(patch) => setProtection((current) => ({ ...current, ...patch }))} />}
        </div>
        <footer className="fmt-dialog-footer">
          {error ? <span className="fmt-error" role="alert"><CircleAlert size={14} aria-hidden="true" />{error}</span> : <span className="fmt-hint">Ctrl+Tab switches tabs · Enter applies</span>}
          <button type="button" className="fmt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="fmt-button is-primary" onClick={submit}>OK</button>
        </footer>
      </section>
    </div>
  )
}

// ---------------------------------------------------------------------------------------
// Number
// ---------------------------------------------------------------------------------------

function NumberPanel({ state, code, validation, sampleValue, initialCode, customInputRef, onCategory, onChange }: {
  state: NumberState
  code: string
  validation: { valid: boolean; message?: string }
  sampleValue: string | number | boolean | null
  initialCode: string
  customInputRef: RefObject<HTMLInputElement>
  onCategory: (category: NumberFormatCategory) => void
  onChange: (patch: Partial<NumberState>) => void
}) {
  const categoryId = useId()
  const category = state.category
  const info = NUMBER_FORMAT_CATEGORIES.find((entry) => entry.id === category)!
  const hasValue = sampleValue !== null && sampleValue !== ''
  const sample = useMemo(() => {
    if (!validation.valid) return { text: '', color: undefined as string | undefined, example: false }
    if (typeof sampleValue === 'boolean') return { text: sampleValue ? 'TRUE' : 'FALSE', color: undefined, example: false }
    if (hasValue) return { ...formatScalarDetailed(sampleValue as string | number, code), example: false }
    const example = category === 'date' || category === 'time' ? SAMPLE_DATE_SERIAL : category === 'text' ? 'Text' : 1234.5678
    return { ...formatScalarDetailed(example, code), example: true }
  }, [category, code, hasValue, sampleValue, validation.valid])

  const categoryItems = NUMBER_FORMAT_CATEGORIES.map((entry) => ({ key: entry.id, label: entry.label, text: entry.label }))
  const decimals = (
    <NumberInput label="Decimal places" value={state.decimals} min={0} max={MAX_DECIMALS} onChange={(value) => onChange({ decimals: value })} />
  )
  const symbol = (
    <label className="fmt-field">
      <span>Symbol</span>
      <select value={state.symbol} onChange={(event) => onChange({ symbol: event.target.value })}>
        {CURRENCY_SYMBOLS.map((entry) => <option key={entry.id} value={entry.id}>{entry.label}</option>)}
      </select>
    </label>
  )
  const negativeList = (
    <div className="fmt-stack">
      <span className="fmt-label" id={`${categoryId}-negative`}>Negative numbers</span>
      <Listbox
        labelledBy={`${categoryId}-negative`}
        className="fmt-list-negative"
        selectedKey={state.negative}
        onSelect={(key) => onChange({ negative: key as NegativeNumberStyle })}
        items={NEGATIVE_NUMBER_STYLES.map((entry) => {
          const rendered = formatScalarDetailed(-1234.1, buildNumberFormat(category, { decimals: state.decimals, thousands: state.thousands, symbol: state.symbol, negative: entry.id }))
          return { key: entry.id, text: rendered.text, title: entry.label, label: <span style={{ color: rendered.color }}>{rendered.text}</span> }
        })}
      />
    </div>
  )
  const typeList = (list: readonly FormatTypeOption[], key: 'date' | 'time' | 'fraction' | 'special') => (
    <div className="fmt-stack fmt-grow">
      <span className="fmt-label" id={`${categoryId}-type`}>Type</span>
      <Listbox
        labelledBy={`${categoryId}-type`}
        className="fmt-list-type"
        selectedKey={state.types[key]}
        onSelect={(id) => onChange({ types: { ...state.types, [key]: id } })}
        items={list.map((entry) => {
          const text = formatTypeSample(category, entry)
          return { key: entry.id, text, label: text, title: entry.code }
        })}
      />
    </div>
  )

  const customItems = useMemo(() => {
    const presets = [...CUSTOM_FORMAT_PRESETS]
    if (initialCode && !presets.includes(initialCode)) presets.push(initialCode)
    return presets.map((preset) => ({ key: preset, text: preset, label: preset }))
  }, [initialCode])

  return (
    <div className="fmt-number">
      <div className="fmt-stack fmt-category-column">
        <span className="fmt-label" id={`${categoryId}-category`}>Category</span>
        <Listbox labelledBy={`${categoryId}-category`} className="fmt-list-category" items={categoryItems} selectedKey={category} onSelect={(key) => onCategory(key as NumberFormatCategory)} />
      </div>
      <div className="fmt-number-options">
        <Group legend={sample.example ? 'Sample (example value)' : 'Sample'} className="fmt-sample-group">
          <div className={`fmt-sample${sample.example ? ' is-example' : ''}`} style={{ color: sample.color }} aria-live="polite">
            {sample.text || ' '}
          </div>
        </Group>
        {category === 'number' && (
          <div className="fmt-options-body">
            <div className="fmt-row">
              {decimals}
              <label className="fmt-check">
                <input type="checkbox" checked={state.thousands} onChange={(event) => onChange({ thousands: event.target.checked })} />
                <span>Use 1000 Separator (,)</span>
              </label>
            </div>
            {negativeList}
          </div>
        )}
        {category === 'currency' && (
          <div className="fmt-options-body">
            <div className="fmt-row">{decimals}{symbol}</div>
            {negativeList}
          </div>
        )}
        {category === 'accounting' && <div className="fmt-options-body"><div className="fmt-row">{decimals}{symbol}</div></div>}
        {(category === 'percentage' || category === 'scientific') && <div className="fmt-options-body"><div className="fmt-row">{decimals}</div></div>}
        {category === 'date' && <div className="fmt-options-body">{typeList(DATE_FORMATS, 'date')}<div className="fmt-locale">Locale (location): English (United States)</div></div>}
        {category === 'time' && <div className="fmt-options-body">{typeList(TIME_FORMATS, 'time')}<div className="fmt-locale">Locale (location): English (United States)</div></div>}
        {category === 'fraction' && <div className="fmt-options-body">{typeList(FRACTION_FORMATS, 'fraction')}</div>}
        {category === 'special' && <div className="fmt-options-body">{typeList(SPECIAL_FORMATS, 'special')}</div>}
        {category === 'custom' && (
          <div className="fmt-options-body">
            <label className="fmt-field fmt-custom-code">
              <span>Type</span>
              <input
                ref={customInputRef}
                value={state.custom}
                spellCheck={false}
                aria-invalid={!validation.valid}
                onChange={(event) => onChange({ custom: event.target.value })}
              />
            </label>
            {!validation.valid && <div className="fmt-inline-error" role="status">{validation.message}</div>}
            <Listbox
              label="Format codes"
              className="fmt-list-custom"
              items={customItems}
              selectedKey={customItems.some((item) => item.key === state.custom) ? state.custom : null}
              onSelect={(key) => onChange({ custom: key })}
            />
          </div>
        )}
        <p className="fmt-description">{info.description}</p>
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------------------
// Alignment
// ---------------------------------------------------------------------------------------

function AlignmentPanel({ state, onChange, canMerge }: { state: AlignmentState; onChange: (patch: Partial<AlignmentState>) => void; canMerge: boolean }) {
  const indentEnabled = ['left', 'right', 'distributed'].includes(state.horizontal)
  return (
    <div className="fmt-alignment">
      <div className="fmt-alignment-main">
        <Group legend="Text alignment">
          <div className="fmt-row">
            <label className="fmt-field fmt-grow">
              <span>Horizontal</span>
              <select value={state.horizontal} onChange={(event) => onChange({ horizontal: event.target.value, ...(['left', 'right', 'distributed'].includes(event.target.value) ? {} : { indent: 0 }) })}>
                {HORIZONTAL_OPTIONS.map((option) => <option key={option.id} value={option.id}>{option.label}</option>)}
              </select>
            </label>
            <NumberInput
              label="Indent"
              value={state.indent}
              min={0}
              max={250}
              disabled={!indentEnabled && state.horizontal !== 'general'}
              onChange={(value) => onChange({ indent: value, ...(value > 0 && !indentEnabled ? { horizontal: 'left' } : {}) })}
            />
          </div>
          <label className="fmt-field">
            <span>Vertical</span>
            <select value={state.vertical} onChange={(event) => onChange({ vertical: event.target.value })}>
              {VERTICAL_OPTIONS.map((option) => <option key={option.id} value={option.id}>{option.label}</option>)}
            </select>
          </label>
          <label className={`fmt-check${state.horizontal !== 'distributed' ? ' is-disabled' : ''}`}>
            <input type="checkbox" disabled={state.horizontal !== 'distributed'} checked={state.justifyLastLine} onChange={(event) => onChange({ justifyLastLine: event.target.checked })} />
            <span>Justify distributed</span>
          </label>
        </Group>
        <Group legend="Text control">
          <label className="fmt-check">
            <input type="checkbox" checked={state.wrapText} onChange={(event) => onChange({ wrapText: event.target.checked, ...(event.target.checked ? { shrinkToFit: false } : {}) })} />
            <span>Wrap text</span>
          </label>
          <label className={`fmt-check${state.wrapText ? ' is-disabled' : ''}`}>
            <input type="checkbox" disabled={state.wrapText} checked={state.shrinkToFit} onChange={(event) => onChange({ shrinkToFit: event.target.checked })} />
            <span>Shrink to fit</span>
          </label>
          <label className={`fmt-check${!canMerge ? ' is-disabled' : ''}`}>
            <input type="checkbox" disabled={!canMerge} checked={state.merge} onChange={(event) => onChange({ merge: event.target.checked })} />
            <span>Merge cells</span>
          </label>
        </Group>
        <Group legend="Right-to-left">
          <label className="fmt-field">
            <span>Text direction</span>
            <select value={state.readingOrder} onChange={(event) => onChange({ readingOrder: event.target.value as ReadingOrder })}>
              <option value="context">Context</option>
              <option value="ltr">Left-to-Right</option>
              <option value="rtl">Right-to-Left</option>
            </select>
          </label>
        </Group>
      </div>
      <Group legend="Orientation" className="fmt-orientation-group">
        <div className="fmt-orientation">
          <button
            type="button"
            className={`fmt-vertical-text${state.verticalText ? ' is-active' : ''}`}
            aria-pressed={state.verticalText}
            aria-label="Vertical text"
            title="Vertical text (stacked letters)"
            onClick={() => onChange({ verticalText: !state.verticalText })}
          >
            {'Text'.split('').map((letter, index) => <span key={index}>{letter}</span>)}
          </button>
          <OrientationDial degrees={state.verticalText ? 0 : state.degrees} disabled={state.verticalText} onChange={(degrees) => onChange({ degrees, verticalText: false })} />
        </div>
        <NumberInput label="Degrees" value={state.verticalText ? 0 : state.degrees} min={-90} max={90} onChange={(degrees) => onChange({ degrees, verticalText: false })} />
      </Group>
    </div>
  )
}

function OrientationDial({ degrees, disabled, onChange }: { degrees: number; disabled: boolean; onChange: (degrees: number) => void }) {
  const svgRef = useRef<SVGSVGElement>(null)
  const width = 112
  const height = 136
  const cx = 12
  const cy = 68
  const radius = 56
  const point = (angle: number, r: number) => {
    const radians = (angle * Math.PI) / 180
    return [cx + r * Math.cos(radians), cy - r * Math.sin(radians)] as const
  }
  const fromPointer = (event: ReactPointerEvent<SVGSVGElement>) => {
    const rect = svgRef.current?.getBoundingClientRect()
    if (!rect) return
    const x = ((event.clientX - rect.left) / rect.width) * width
    const y = ((event.clientY - rect.top) / rect.height) * height
    const angle = Math.round((Math.atan2(cy - y, Math.max(0.001, x - cx)) * 180) / Math.PI)
    onChange(clamp(angle, -90, 90))
  }
  const handleKey = (event: ReactKeyboardEvent<SVGSVGElement>) => {
    const steps: Record<string, number> = { ArrowUp: 1, ArrowRight: 1, ArrowDown: -1, ArrowLeft: -1, PageUp: 15, PageDown: -15 }
    if (event.key in steps) {
      event.preventDefault()
      onChange(clamp(degrees + steps[event.key], -90, 90))
    } else if (event.key === 'Home' || event.key === 'End') {
      event.preventDefault()
      onChange(event.key === 'Home' ? 90 : -90)
    }
  }
  const ticks = []
  for (let angle = -90; angle <= 90; angle += 15) {
    const [x, y] = point(angle, radius)
    ticks.push(<circle key={angle} cx={x} cy={y} r={angle % 45 === 0 ? 2.2 : 1.5} className={angle === degrees && !disabled ? 'is-active' : ''} />)
  }
  const [hx, hy] = point(degrees, radius)
  const [lx1, ly1] = point(degrees, 34)
  const [lx2, ly2] = point(degrees, radius - 5)
  return (
    <svg
      ref={svgRef}
      className={`fmt-dial${disabled ? ' is-disabled' : ''}`}
      width={width}
      height={height}
      viewBox={`0 0 ${width} ${height}`}
      role="slider"
      tabIndex={0}
      aria-label="Text orientation"
      aria-valuemin={-90}
      aria-valuemax={90}
      aria-valuenow={degrees}
      aria-valuetext={`${degrees} degrees`}
      onKeyDown={handleKey}
      onPointerDown={(event) => { event.preventDefault(); event.currentTarget.setPointerCapture(event.pointerId); event.currentTarget.focus(); fromPointer(event) }}
      onPointerMove={(event) => { if (event.currentTarget.hasPointerCapture(event.pointerId)) fromPointer(event) }}
    >
      <path d={`M ${cx} ${cy - radius} A ${radius} ${radius} 0 0 1 ${cx} ${cy + radius}`} className="fmt-dial-arc" />
      <g className="fmt-dial-ticks">{ticks}</g>
      <g transform={`rotate(${-degrees} ${cx} ${cy})`}>
        <text x={cx + 4} y={cy + 4} className="fmt-dial-text">Text</text>
      </g>
      <line x1={lx1} y1={ly1} x2={lx2} y2={ly2} className="fmt-dial-line" />
      <rect x={hx - 4} y={hy - 4} width={8} height={8} transform={`rotate(45 ${hx} ${hy})`} className="fmt-dial-handle" />
    </svg>
  )
}

// ---------------------------------------------------------------------------------------
// Font
// ---------------------------------------------------------------------------------------

function FontPanel({ state, initial, defaultFont, themeColors, sizeInputRef, onChange }: {
  state: FontState
  initial: FontState
  defaultFont: { name: string; size: number }
  themeColors: readonly string[]
  sizeInputRef: RefObject<HTMLInputElement>
  onChange: (patch: Partial<FontState>) => void
}) {
  const ids = useId()
  const installed = useMemo(() => new Map(FONT_FAMILIES.map((entry) => [entry.name, isFontInstalled(entry.name)])), [])
  const fontItems = useMemo<ListboxItem[]>(() => {
    const names = FONT_FAMILIES.map((entry) => entry.name)
    if (initial.name && !names.some((name) => name.toLocaleLowerCase() === initial.name.toLocaleLowerCase())) names.unshift(initial.name)
    return names.map((name) => {
      const available = installed.get(name)
      const symbol = isSymbolFont(name)
      return {
        key: name,
        text: name,
        label: <>{name}{symbol && <small className="fmt-option-note">symbol</small>}</>,
        style: symbol || available === false ? undefined : { fontFamily: fontStack(name) },
        muted: available === false,
        title: available === false ? `${name} isn’t installed on this computer; a similar font will be shown.` : name,
      }
    })
  }, [initial.name, installed])
  const typed = state.name.trim().toLocaleLowerCase()
  const exact = fontItems.find((item) => item.key.toLocaleLowerCase() === typed)?.key ?? null
  const prefix = exact ? null : fontItems.find((item) => item.key.toLocaleLowerCase().startsWith(typed))?.key ?? null
  const size = parseFontSize(state.sizeText)
  const sizeKey = size !== null && FONT_SIZES.includes(size) ? String(size) : null
  const isNormal = state.name.toLocaleLowerCase() === defaultFont.name.toLocaleLowerCase() && size === defaultFont.size && state.style === 'regular'
    && state.underline === 'none' && !state.color && !state.strike && !state.vertAlign
  const available = isFontInstalled(state.name)

  const previewStyle: CSSProperties = {
    fontFamily: fontStack(state.name),
    fontWeight: state.style === 'bold' || state.style === 'boldItalic' ? 700 : 400,
    fontStyle: state.style === 'italic' || state.style === 'boldItalic' ? 'italic' : 'normal',
    fontSize: `${clamp((size ?? defaultFont.size) * (4 / 3), 8, 44)}px`,
    color: resolveColorHex(state.color, themeColors) || '#000000',
    textDecorationLine: [state.underline !== 'none' ? 'underline' : '', state.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none',
    textDecorationStyle: state.underline === 'double' || state.underline === 'doubleAccounting' ? 'double' : 'solid',
    textUnderlineOffset: state.underline.endsWith('Accounting') ? '0.28em' : undefined,
  }

  return (
    <div className="fmt-font">
      <div className="fmt-font-lists">
        <div className="fmt-stack fmt-font-name">
          <label className="fmt-label" htmlFor={`${ids}-name`}>Font</label>
          <input id={`${ids}-name`} value={state.name} spellCheck={false} autoComplete="off" onChange={(event) => onChange({ name: event.target.value })} />
          <Listbox label="Fonts" className="fmt-list-font" items={fontItems} selectedKey={exact} scrollKey={prefix} onSelect={(key) => onChange({ name: key })} />
        </div>
        <div className="fmt-stack fmt-font-style">
          <span className="fmt-label" id={`${ids}-style`}>Font style</span>
          <input value={FONT_STYLES.find((entry) => entry.id === state.style)?.label || 'Regular'} readOnly tabIndex={-1} aria-hidden="true" />
          <Listbox
            labelledBy={`${ids}-style`}
            className="fmt-list-font"
            items={FONT_STYLES.map((entry) => ({ key: entry.id, text: entry.label, label: entry.label, style: entry.css }))}
            selectedKey={state.style}
            onSelect={(key) => onChange({ style: key as FontStyleKey })}
          />
        </div>
        <div className="fmt-stack fmt-font-size">
          <label className="fmt-label" htmlFor={`${ids}-size`}>Size</label>
          <input
            ref={sizeInputRef}
            id={`${ids}-size`}
            value={state.sizeText}
            inputMode="decimal"
            aria-invalid={size === null}
            onChange={(event) => onChange({ sizeText: event.target.value })}
          />
          <Listbox
            label="Font sizes"
            className="fmt-list-font"
            items={FONT_SIZES.map((value) => ({ key: String(value), text: String(value), label: String(value) }))}
            selectedKey={sizeKey}
            onSelect={(key) => onChange({ sizeText: key })}
          />
        </div>
      </div>
      <div className="fmt-row fmt-font-row">
        <label className="fmt-field fmt-grow">
          <span>Underline</span>
          <select value={state.underline} onChange={(event) => onChange({ underline: event.target.value as UnderlineKey })}>
            {UNDERLINES.map((option) => <option key={option.id} value={option.id}>{option.label}</option>)}
          </select>
        </label>
        <div className="fmt-field fmt-grow">
          <span id={`${ids}-color`}>Color</span>
          <ColorDropdownButton label="Font color" value={state.color} themeColors={themeColors} mode="text" onChange={(color) => onChange({ color })} />
        </div>
        <label className="fmt-check fmt-normal-font">
          <input
            type="checkbox"
            checked={isNormal}
            onChange={() => {
              if (!isNormal) onChange({ name: defaultFont.name, sizeText: String(defaultFont.size), style: 'regular', underline: 'none', color: null, strike: false, vertAlign: '' })
            }}
          />
          <span>Normal font</span>
        </label>
      </div>
      <div className="fmt-row fmt-font-row">
        <Group legend="Effects" className="fmt-effects">
          <label className="fmt-check">
            <input type="checkbox" checked={state.strike} onChange={(event) => onChange({ strike: event.target.checked })} />
            <span>Strikethrough</span>
          </label>
          <label className="fmt-check">
            <input type="checkbox" checked={state.vertAlign === 'superscript'} onChange={(event) => onChange({ vertAlign: event.target.checked ? 'superscript' : '' })} />
            <span>Superscript</span>
          </label>
          <label className="fmt-check">
            <input type="checkbox" checked={state.vertAlign === 'subscript'} onChange={(event) => onChange({ vertAlign: event.target.checked ? 'subscript' : '' })} />
            <span>Subscript</span>
          </label>
        </Group>
        <Group legend="Preview" className="fmt-font-preview-group">
          <div className="fmt-font-preview" aria-label="Font preview">
            <span style={previewStyle}>
              {state.vertAlign ? <>Aa<span className={`fmt-vert-${state.vertAlign}`}>BbCc</span>YyZz</> : 'AaBbCcYyZz'}
            </span>
          </div>
        </Group>
      </div>
      <p className="fmt-description">
        {available === false
          ? `“${state.name}” isn’t installed on this computer. The preview and grid show a similar font; the name is kept in the file.`
          : 'This font is installed on this computer. The same font will be used on screen and when printing.'}
      </p>
    </div>
  )
}

// ---------------------------------------------------------------------------------------
// Border
// ---------------------------------------------------------------------------------------

type EdgeIconKind = EdgeKey
const EDGE_LABELS: Record<EdgeKey, string> = {
  top: 'Top border',
  bottom: 'Bottom border',
  left: 'Left border',
  right: 'Right border',
  insideHorizontal: 'Inside horizontal border',
  insideVertical: 'Inside vertical border',
  diagonalUp: 'Diagonal up border',
  diagonalDown: 'Diagonal down border',
}

function EdgeIcon({ edge }: { edge: EdgeIconKind }) {
  const lines: Record<EdgeKey, [number, number, number, number]> = {
    top: [3, 3, 17, 3],
    bottom: [3, 17, 17, 17],
    left: [3, 3, 3, 17],
    right: [17, 3, 17, 17],
    insideHorizontal: [3, 10, 17, 10],
    insideVertical: [10, 3, 10, 17],
    diagonalUp: [3, 17, 17, 3],
    diagonalDown: [3, 3, 17, 17],
  }
  const [x1, y1, x2, y2] = lines[edge]
  return (
    <svg width={20} height={20} viewBox="0 0 20 20" aria-hidden="true">
      <rect x={3} y={3} width={14} height={14} fill="none" stroke="currentColor" strokeOpacity={0.3} strokeDasharray="1 1.5" />
      <line x1={x1} y1={y1} x2={x2} y2={y2} stroke="currentColor" strokeWidth={1.8} />
    </svg>
  )
}

function BorderPanel({ edges, lineStyle, lineColor, themeColors, multiRow, multiColumn, onLineStyle, onLineColor, onToggle, onPreset }: {
  edges: Record<EdgeKey, CellBorderSide | null>
  lineStyle: string | null
  lineColor: ColorValue | null
  themeColors: readonly string[]
  multiRow: boolean
  multiColumn: boolean
  onLineStyle: (style: string | null) => void
  onLineColor: (color: ColorValue | null) => void
  onToggle: (edge: EdgeKey) => void
  onPreset: (preset: 'none' | 'outline' | 'inside') => void
}) {
  const galleryRef = useRef<HTMLDivElement>(null)
  const lineHex = resolveColorHex(lineColor, themeColors) || '#000000'
  const width = 196
  const height = 132
  const x0 = 18, y0 = 14, x1 = width - 18, y1 = height - 14
  const midX = (x0 + x1) / 2
  const midY = (y0 + y1) / 2
  const columns = multiColumn ? [[x0, midX], [midX, x1]] : [[x0, x1]]
  const rows = multiRow ? [[y0, midY], [midY, y1]] : [[y0, y1]]

  const segments: Array<{ edge: EdgeKey; x1: number; y1: number; x2: number; y2: number }> = [
    { edge: 'top', x1: x0, y1: y0, x2: x1, y2: y0 },
    { edge: 'bottom', x1: x0, y1: y1, x2: x1, y2: y1 },
    { edge: 'left', x1: x0, y1: y0, x2: x0, y2: y1 },
    { edge: 'right', x1: x1, y1: y0, x2: x1, y2: y1 },
  ]
  if (multiRow) segments.push({ edge: 'insideHorizontal', x1: x0, y1: midY, x2: x1, y2: midY })
  if (multiColumn) segments.push({ edge: 'insideVertical', x1: midX, y1: y0, x2: midX, y2: y1 })

  const handleDiagramClick = (event: ReactMouseEvent<SVGSVGElement>) => {
    const rect = event.currentTarget.getBoundingClientRect()
    const x = ((event.clientX - rect.left) / rect.width) * width
    const y = ((event.clientY - rect.top) / rect.height) * height
    let best: EdgeKey | null = null
    let bestDistance = 9
    for (const segment of segments) {
      const horizontal = segment.y1 === segment.y2
      const distance = horizontal
        ? (x >= segment.x1 - 4 && x <= segment.x2 + 4 ? Math.abs(y - segment.y1) : Infinity)
        : (y >= segment.y1 - 4 && y <= segment.y2 + 4 ? Math.abs(x - segment.x1) : Infinity)
      if (distance < bestDistance) { bestDistance = distance; best = segment.edge }
    }
    if (!best) {
      // Diagonals: distance to either diagonal of the cell under the pointer.
      const column = columns.find(([start, end]) => x >= start && x <= end)
      const row = rows.find(([start, end]) => y >= start && y <= end)
      if (column && row) {
        const [cx0, cx1] = column
        const [cy0, cy1] = row
        const u = (x - cx0) / (cx1 - cx0)
        const v = (y - cy0) / (cy1 - cy0)
        const cellHeight = cy1 - cy0
        const downDistance = Math.abs(v - u) * cellHeight
        const upDistance = Math.abs(1 - v - u) * cellHeight
        if (Math.min(downDistance, upDistance) < 10) best = downDistance < upDistance ? 'diagonalDown' : 'diagonalUp'
      }
    }
    if (best) onToggle(best)
  }

  const handleGalleryKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const index = LINE_STYLE_GALLERY.indexOf(lineStyle)
    const perColumn = 7
    let next = index
    if (event.key === 'ArrowDown') next = index + 1
    else if (event.key === 'ArrowUp') next = index - 1
    else if (event.key === 'ArrowRight') next = index + perColumn
    else if (event.key === 'ArrowLeft') next = index - perColumn
    else if (event.key === 'Home') next = 0
    else if (event.key === 'End') next = LINE_STYLE_GALLERY.length - 1
    else return
    event.preventDefault()
    next = clamp(next, 0, LINE_STYLE_GALLERY.length - 1)
    onLineStyle(LINE_STYLE_GALLERY[next])
    window.requestAnimationFrame(() => galleryRef.current?.querySelector<HTMLElement>('[aria-checked="true"]')?.focus())
  }

  const toggleButton = (edge: EdgeKey) => (
    <button
      type="button"
      className={`fmt-edge-button${edges[edge] ? ' is-active' : ''}`}
      aria-pressed={Boolean(edges[edge])}
      aria-label={EDGE_LABELS[edge]}
      title={EDGE_LABELS[edge]}
      disabled={(edge === 'insideHorizontal' && !multiRow) || (edge === 'insideVertical' && !multiColumn)}
      onClick={() => onToggle(edge)}
    >
      <EdgeIcon edge={edge} />
    </button>
  )

  return (
    <div className="fmt-border">
      <Group legend="Line" className="fmt-line-group">
        <span className="fmt-label">Style</span>
        <div ref={galleryRef} className="fmt-line-gallery" role="radiogroup" aria-label="Line style" onKeyDown={handleGalleryKey}>
          {LINE_STYLE_GALLERY.map((style) => (
            <button
              key={style || 'none'}
              type="button"
              role="radio"
              aria-checked={style === lineStyle}
              tabIndex={style === lineStyle ? 0 : -1}
              className={`fmt-line-option${style === lineStyle ? ' is-selected' : ''}`}
              title={style ? LINE_STYLE_SPECS[style].label : 'None'}
              onClick={() => onLineStyle(style)}
            >
              <LineStyleSample style={style} color={lineHex} width={52} />
            </button>
          ))}
        </div>
        <div className="fmt-field">
          <span>Color</span>
          <ColorDropdownButton label="Line color" value={lineColor} themeColors={themeColors} mode="text" onChange={(color) => onLineColor(color)} />
        </div>
      </Group>
      <div className="fmt-border-main-column">
        <Group legend="Presets">
          <div className="fmt-border-presets">
            {([
              { id: 'none', label: 'None', segments: [] as EdgeKey[] },
              { id: 'outline', label: 'Outline', segments: ['top', 'bottom', 'left', 'right'] as EdgeKey[] },
              { id: 'inside', label: 'Inside', segments: ['insideHorizontal', 'insideVertical'] as EdgeKey[] },
            ] as const).map((preset) => (
              <button
                key={preset.id}
                type="button"
                className="fmt-border-preset-button"
                disabled={preset.id === 'inside' && !multiRow && !multiColumn}
                onClick={() => onPreset(preset.id)}
              >
                <svg width={34} height={34} viewBox="0 0 34 34" aria-hidden="true">
                  <rect x={4} y={4} width={26} height={26} fill="#fff" stroke="currentColor" strokeOpacity={0.28} strokeDasharray="1 1.6" />
                  <line x1={17} y1={4} x2={17} y2={30} stroke="currentColor" strokeOpacity={0.28} strokeDasharray="1 1.6" />
                  <line x1={4} y1={17} x2={30} y2={17} stroke="currentColor" strokeOpacity={0.28} strokeDasharray="1 1.6" />
                  {preset.id === 'outline' && <rect x={4} y={4} width={26} height={26} fill="none" stroke="currentColor" strokeWidth={1.8} />}
                  {preset.id === 'inside' && <><line x1={17} y1={4} x2={17} y2={30} stroke="currentColor" strokeWidth={1.8} /><line x1={4} y1={17} x2={30} y2={17} stroke="currentColor" strokeWidth={1.8} /></>}
                </svg>
                <span>{preset.label}</span>
              </button>
            ))}
          </div>
        </Group>
        <Group legend="Border">
          <div className="fmt-border-editor">
            <div className="fmt-edge-column">
              {toggleButton('top')}
              {toggleButton('insideHorizontal')}
              {toggleButton('bottom')}
            </div>
            <div className="fmt-border-diagram-wrap">
              <svg
                className="fmt-border-diagram"
                width={width}
                height={height}
                viewBox={`0 0 ${width} ${height}`}
                role="img"
                aria-label="Border preview. Click an edge to toggle it."
                onClick={handleDiagramClick}
              >
                <rect x={x0} y={y0} width={x1 - x0} height={y1 - y0} fill="#fff" />
                {/* Corner and midpoint guides, as in Excel. */}
                <g className="fmt-diagram-guides">
                  {[[x0, y0, -1, -1], [x1, y0, 1, -1], [x0, y1, -1, 1], [x1, y1, 1, 1]].map(([gx, gy, dx, dy], index) => (
                    <g key={index}>
                      <line x1={gx + dx * 3} y1={gy} x2={gx + dx * 9} y2={gy} />
                      <line x1={gx} y1={gy + dy * 3} x2={gx} y2={gy + dy * 9} />
                    </g>
                  ))}
                  {multiColumn && <><line x1={midX} y1={y0 - 3} x2={midX} y2={y0 - 9} /><line x1={midX} y1={y1 + 3} x2={midX} y2={y1 + 9} /></>}
                  {multiRow && <><line x1={x0 - 3} y1={midY} x2={x0 - 9} y2={midY} /><line x1={x1 + 3} y1={midY} x2={x1 + 9} y2={midY} /></>}
                </g>
                {rows.map(([ry0, ry1], rowIndex) => columns.map(([cx0, cx1], columnIndex) => (
                  <g key={`${rowIndex}-${columnIndex}`}>
                    <text x={(cx0 + cx1) / 2} y={(ry0 + ry1) / 2 + 4} className="fmt-diagram-text">Text</text>
                    {edges.diagonalUp && <BorderLine side={edges.diagonalUp} x1={cx0} y1={ry1} x2={cx1} y2={ry0} themeColors={themeColors} />}
                    {edges.diagonalDown && <BorderLine side={edges.diagonalDown} x1={cx0} y1={ry0} x2={cx1} y2={ry1} themeColors={themeColors} />}
                  </g>
                )))}
                {segments.map((segment) => {
                  const side = edges[segment.edge]
                  return side ? <BorderLine key={segment.edge} side={side} x1={segment.x1} y1={segment.y1} x2={segment.x2} y2={segment.y2} themeColors={themeColors} /> : null
                })}
              </svg>
            </div>
            <div className="fmt-edge-row">
              {toggleButton('diagonalUp')}
              {toggleButton('left')}
              {toggleButton('insideVertical')}
              {toggleButton('right')}
              {toggleButton('diagonalDown')}
            </div>
          </div>
        </Group>
        <p className="fmt-description">Choose a line style and color, then click a preset, the preview, or the buttons around it.</p>
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------------------
// Fill
// ---------------------------------------------------------------------------------------

function FillPanel({ state, themeColors, onChange }: { state: FillState; themeColors: readonly string[]; onChange: (patch: Partial<FillState>) => void }) {
  const patternRef = useRef<HTMLDivElement>(null)
  const inkHex = resolveColorHex(state.patternColor, themeColors) || '#000000'
  const sample = fillPreviewCss(buildFill(state), themeColors)
  const handlePatternKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const index = PATTERN_FILLS.findIndex((entry) => entry.id === state.pattern)
    const perRow = 6
    let next = index
    if (event.key === 'ArrowRight') next = index + 1
    else if (event.key === 'ArrowLeft') next = index - 1
    else if (event.key === 'ArrowDown') next = index + perRow
    else if (event.key === 'ArrowUp') next = index - perRow
    else if (event.key === 'Home') next = 0
    else if (event.key === 'End') next = PATTERN_FILLS.length - 1
    else return
    event.preventDefault()
    onChange({ pattern: PATTERN_FILLS[clamp(next, 0, PATTERN_FILLS.length - 1)].id })
    window.requestAnimationFrame(() => patternRef.current?.querySelector<HTMLElement>('[aria-checked="true"]')?.focus())
  }
  return (
    <div className="fmt-fill">
      <Group legend="Background Color" className="fmt-fill-background">
        <ColorPicker value={state.background} themeColors={themeColors} mode="fill" nullLabel="No Color" label="Background color" onChange={(color) => onChange({ background: color })} />
      </Group>
      <div className="fmt-fill-side">
        <Group legend="Pattern">
          <div className="fmt-field">
            <span>Pattern Color</span>
            <ColorDropdownButton label="Pattern color" value={state.patternColor} themeColors={themeColors} mode="text" onChange={(color) => onChange({ patternColor: color })} />
          </div>
          <span className="fmt-label">Pattern Style</span>
          <div ref={patternRef} className="fmt-pattern-grid" role="radiogroup" aria-label="Pattern style" onKeyDown={handlePatternKey}>
            {PATTERN_FILLS.map((entry) => (
              <button
                key={entry.id}
                type="button"
                role="radio"
                aria-checked={state.pattern === entry.id}
                tabIndex={state.pattern === entry.id ? 0 : -1}
                className={`fmt-pattern${state.pattern === entry.id ? ' is-selected' : ''}${entry.id === 'none' ? ' is-none' : ''}`}
                title={entry.label}
                aria-label={entry.label}
                style={entry.id === 'none' ? undefined : fillPreviewCss({ type: 'pattern', pattern: entry.id, fgColor: { argb: `FF${inkHex.slice(1)}` }, bgColor: { argb: 'FFFFFFFF' } }, themeColors)}
                onClick={() => onChange({ pattern: entry.id })}
              />
            ))}
          </div>
        </Group>
        <Group legend="Sample" className="fmt-fill-sample-group">
          <div className="fmt-fill-sample" style={sample} aria-label="Fill sample" />
        </Group>
        {state.gradient && <p className="fmt-description">This cell has a gradient fill. Choosing a color or pattern replaces it.</p>}
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------------------
// Protection
// ---------------------------------------------------------------------------------------

function ProtectionPanel({ state, onChange }: { state: { locked: boolean; hidden: boolean }; onChange: (patch: Partial<{ locked: boolean; hidden: boolean }>) => void }) {
  return (
    <div className="fmt-protection">
      <label className="fmt-check">
        <input type="checkbox" checked={state.locked} onChange={(event) => onChange({ locked: event.target.checked })} />
        <span>Locked</span>
      </label>
      <label className="fmt-check">
        <input type="checkbox" checked={state.hidden} onChange={(event) => onChange({ hidden: event.target.checked })} />
        <span>Hidden</span>
      </label>
      <p className="fmt-description">
        Locking cells or hiding formulas has no effect until you protect the worksheet. Locked cells can’t be edited on a
        protected sheet; hidden cells don’t show their formula in the formula bar.
      </p>
    </div>
  )
}
