// src/advanced/panels/AdjustmentControls.tsx (WP6)
// The settings of every adjustment (design 5.10, 5.14), shared by the Properties panel (adjustment layers,
// non-destructive) and the Image > Adjustments dialogs (destructive, with a live canvas preview):
//   - AdjustmentControls renders the controls for one AdjustmentSpec and reports every change as
//     onChange(next, phase): 'live' while a slider or handle is dragged, 'commit' when a drag ends or a
//     value is typed. onInteractive(true / false) brackets drags (the compositor shows a coarser proxy
//     meanwhile, so dragging stays smooth on large images).
//   - Levels and Curves have their own editors (LevelsEditor.tsx, CurvesEditor.tsx); Gradient Map has a
//     stop editor below; Hue/Saturation edits Master or one of the six colour ranges; Photo Filter offers
//     Photoshop's filter presets.
//   - The small form controls (NumberInput, SliderField, CheckField, SelectField, Segmented) are exported
//     for the other panels and dialogs. Components here are function declarations, so the editors that
//     import them back (LevelsEditor, CurvesEditor) form a safe import cycle.
//   - histogramBelow / histogramOfTarget compute the histograms Levels, Curves and Threshold show, on a
//     small pyramid proxy (at most 512 px on the long side), so they stay cheap on 50 MP documents.
import { useEffect, useId, useRef, useState } from 'react'
import type { CSSProperties, PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import { Trash2 } from 'lucide-react'
import type {
  AdjustmentSpec,
  ColorBalanceShift,
  GradientStop,
  Histogram,
  HslShift,
  HueRange,
  IntRect,
  MaskBuffer,
  PixelBuffer,
  Rgb8,
} from '../../imaging/types.ts'
import type { DocumentState, LayerId, Selection } from '../types.ts'
import { normalizeStops } from '../../imaging/adjustments.ts'
import { computeHistogram } from '../../imaging/histogram.ts'
import { gradientColorAt } from '../../imaging/gradient.ts'
import { compositeRect } from '../composite.ts'
import { levelOffset, maxPyramidLevel, readMaskLevel, readSurfaceLevel } from '../pyramid.ts'
import { activeLayerOf, canvasRect, isBackgroundLayer, layerPixelBounds, pixelSourceOf } from '../tools/shared.ts'
import { HistogramView, LevelsEditor } from './LevelsEditor.tsx'
import { CurvesEditor } from './CurvesEditor.tsx'
import { ColorButton, cssColor, sameRgb } from './ColorPicker.tsx'

export type ChangePhase = 'live' | 'commit'

// ---------------------------------------------------------------------------------------------
// Small form controls
// ---------------------------------------------------------------------------------------------

/** Scrolls only `list` (not the dock or the page) so that `item`, a child of it, is fully visible. */
export function revealInList(list: HTMLElement | null, item: Element | null): void {
  if (!list || !(item instanceof HTMLElement)) return
  const top = item.offsetTop
  const bottom = top + item.offsetHeight
  if (top < list.scrollTop) list.scrollTop = top
  else if (bottom > list.scrollTop + list.clientHeight) list.scrollTop = bottom - list.clientHeight
}

export function clampNumber(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value
}

/** "1.5" for 1.50 (digits = decimals kept at most); fixed keeps trailing zeros ("1.00"). */
export function formatNumber(value: number, digits = 0, fixed = false): string {
  if (!Number.isFinite(value)) return ''
  const text = value.toFixed(digits)
  return fixed || digits === 0 ? text : text.replace(/\.?0+$/, '')
}

function roundTo(value: number, digits: number): number {
  return digits > 0 ? Number(value.toFixed(digits)) : Math.round(value)
}

export interface NumberInputProps {
  readonly value: number
  readonly min: number
  readonly max: number
  readonly step?: number
  readonly digits?: number
  readonly fixed?: boolean
  /** Accessible name. */
  readonly label: string
  readonly id?: string
  readonly className?: string
  readonly disabled?: boolean
  readonly title?: string
  /** Called with the clamped value on Enter, blur or an arrow key. */
  readonly onChange: (value: number) => void
}

/**
 * A text box for a number: type and press Enter (or leave the field), or use the arrow keys (Shift = 10
 * steps). Escape restores the shown value. The new value is reported synchronously, before the key event
 * reaches a dialog, so Enter in a field applies exactly what was typed.
 */
export function NumberInput({ value, min, max, step = 1, digits = 0, fixed = false, label, id, className, disabled, title, onChange }: NumberInputProps) {
  const format = (model: number) => formatNumber(model, digits, fixed)
  const [text, setText] = useState(() => format(value))
  const editing = useRef(false)
  useEffect(() => {
    if (!editing.current) setText(format(value))
  }, [value, digits, fixed])
  const commit = () => {
    editing.current = false
    const parsed = Number(text.replace(',', '.').trim())
    if (text.trim() !== '' && Number.isFinite(parsed)) {
      const next = clampNumber(roundTo(parsed, digits), min, max)
      setText(format(next))
      if (next !== value) onChange(next)
    } else {
      setText(format(value))
    }
  }
  return (
    <input
      id={id}
      className={`ae-num ${className ?? ''}`}
      type="text"
      inputMode="decimal"
      aria-label={label}
      title={title}
      value={text}
      disabled={disabled}
      onFocus={(event) => {
        editing.current = true
        event.currentTarget.select()
      }}
      onChange={(event) => setText(event.currentTarget.value)}
      onBlur={commit}
      onKeyDown={(event) => {
        if (event.key === 'Enter') {
          commit()
        } else if (event.key === 'Escape') {
          editing.current = false
          setText(format(value))
        } else if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
          event.preventDefault()
          const base = Number(text.replace(',', '.'))
          const from = Number.isFinite(base) && text.trim() !== '' ? base : value
          const next = clampNumber(roundTo(from + (event.key === 'ArrowUp' ? step : -step) * (event.shiftKey ? 10 : 1), digits), min, max)
          setText(format(next))
          if (next !== value) onChange(next)
        }
      }}
    />
  )
}

export interface SliderFieldProps {
  readonly label: string
  readonly value: number
  readonly min: number
  readonly max: number
  readonly step?: number
  readonly digits?: number
  readonly fixed?: boolean
  readonly unit?: string
  readonly disabled?: boolean
  /** CSS background of the track (colour sliders show what they do). */
  readonly track?: string
  readonly title?: string
  /** One row: label, slider, value (compact panels such as Color). */
  readonly inline?: boolean
  readonly onChange: (value: number, phase: ChangePhase) => void
  readonly onInteractive?: (active: boolean) => void
}

/** Label and value box on one line, the slider under them (Photoshop's Properties layout), or all in one row. */
export function SliderField({ label, value, min, max, step = 1, digits = 0, fixed, unit, disabled, track, title, inline, onChange, onInteractive }: SliderFieldProps) {
  const id = useId()
  const dragging = useRef(false)
  const last = useRef(value)
  last.current = value
  const end = () => {
    if (!dragging.current) return
    dragging.current = false
    onInteractive?.(false)
    onChange(last.current, 'commit')
  }
  const shown = Number.isFinite(value) ? clampNumber(value, min, max) : min
  const box = <NumberInput id={id} label={label} value={value} min={min} max={max} step={step} digits={digits} fixed={fixed} disabled={disabled} onChange={(next) => onChange(next, 'commit')} />
  return (
    <div className={`ae-sf ${inline ? 'is-inline' : ''}`} title={title}>
      {inline ? <label htmlFor={id} className="ae-sf-label">{label}</label> : (
        <div className="ae-sf-head">
          <label htmlFor={id}>{label}</label>
          {box}
          {unit && <span className="ae-sf-unit">{unit}</span>}
        </div>
      )}
      <input
        type="range"
        className="ae-slider"
        aria-label={label}
        min={min}
        max={max}
        step={step}
        value={shown}
        disabled={disabled}
        style={track ? { '--ae-track': track } as CSSProperties : undefined}
        onPointerDown={(event: ReactPointerEvent<HTMLInputElement>) => {
          if (event.button !== 0) return
          dragging.current = true
          onInteractive?.(true)
        }}
        onPointerUp={end}
        onPointerCancel={end}
        onLostPointerCapture={end}
        onBlur={end}
        onChange={(event) => {
          const next = roundTo(Number(event.currentTarget.value), digits)
          last.current = next
          onChange(next, dragging.current ? 'live' : 'commit')
        }}
      />
      {inline && box}
      {inline && <span className="ae-sf-unit">{unit ?? ''}</span>}
    </div>
  )
}

export function CheckField({ label, checked, disabled, title, onChange }: {
  readonly label: string
  readonly checked: boolean
  readonly disabled?: boolean
  readonly title?: string
  readonly onChange: (checked: boolean) => void
}) {
  return (
    <label className="ae-check" title={title}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(event) => onChange(event.currentTarget.checked)} />
      <span>{label}</span>
    </label>
  )
}

export function SelectField<T extends string | number>({ label, value, options, disabled, title, inline, onChange }: {
  readonly label: string
  readonly value: T
  readonly options: readonly (readonly [T, string] | '-')[]
  readonly disabled?: boolean
  readonly title?: string
  /** Label and select on one line. */
  readonly inline?: boolean
  readonly onChange: (value: T) => void
}) {
  const id = useId()
  return (
    <div className={`ae-field ${inline ? 'is-inline' : ''}`} title={title}>
      <label htmlFor={id}>{label}</label>
      <select
        id={id}
        value={String(value)}
        disabled={disabled}
        onChange={(event) => {
          const raw = event.currentTarget.value
          const match = options.find((option) => option !== '-' && String(option[0]) === raw)
          if (match && match !== '-') onChange(match[0])
        }}
      >
        {options.map((option, index) => (option === '-'
          ? <option key={`sep-${index}`} disabled>──────────</option>
          : <option key={String(option[0])} value={String(option[0])}>{option[1]}</option>))}
      </select>
    </div>
  )
}

export function Segmented<T extends string>({ label, value, options, disabled, onChange }: {
  readonly label: string
  readonly value: T
  readonly options: readonly (readonly [T, string])[]
  readonly disabled?: boolean
  readonly onChange: (value: T) => void
}) {
  return (
    <div className="ae-seg" role="radiogroup" aria-label={label}>
      {options.map(([option, text]) => (
        <button
          key={option}
          type="button"
          role="radio"
          aria-checked={value === option}
          className={value === option ? 'is-active' : ''}
          disabled={disabled}
          onClick={() => onChange(option)}
        >{text}</button>
      ))}
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// Histograms (proxy level, cheap on big documents)
// ---------------------------------------------------------------------------------------------

/** Pyramid level at which a width x height area is at most `target` px on its long side. */
export function proxyLevel(width: number, height: number, target = 512, maxLevel = maxPyramidLevel(width, height)): number {
  const long = Math.max(1, width, height)
  let level = 0
  while (long / 2 ** level > target && level < maxLevel) level += 1
  return level
}

/** What an adjustment layer receives: the composite of the layers below it. */
export function histogramBelow(state: DocumentState, layerId: LayerId): Histogram | null {
  try {
    const level = proxyLevel(state.width, state.height)
    const width = Math.max(1, Math.ceil(state.width / 2 ** level))
    const height = Math.max(1, Math.ceil(state.height / 2 ** level))
    const pixels = compositeRect(state, { x: 0, y: 0, width, height }, { level, belowLayerId: layerId })
    return computeHistogram(pixels)
  } catch (error) {
    console.error(error)
    return null
  }
}

/** Selection coverage sampled at the centres of level pixels (0 outside the document). */
export function sampleSelection(selection: Selection, rect: IntRect, level: number): MaskBuffer {
  const scale = 2 ** level
  const { mask } = selection
  const data = new Uint8Array(rect.width * rect.height)
  for (let y = 0; y < rect.height; y += 1) {
    const dy = Math.floor((rect.y + y) * scale + scale / 2)
    if (dy < 0 || dy >= mask.height) continue
    for (let x = 0; x < rect.width; x += 1) {
      const dx = Math.floor((rect.x + x) * scale + scale / 2)
      if (dx < 0 || dx >= mask.width) continue
      data[y * rect.width + x] = mask.data[dy * mask.width + dx]
    }
  }
  return { width: rect.width, height: rect.height, data }
}

function levelRectOf(rect: IntRect, level: number): IntRect {
  const scale = 2 ** level
  const x0 = Math.floor(rect.x / scale)
  const y0 = Math.floor(rect.y / scale)
  const x1 = Math.ceil((rect.x + rect.width) / scale)
  const y1 = Math.ceil((rect.y + rect.height) / scale)
  return { x: x0, y: y0, width: Math.max(1, x1 - x0), height: Math.max(1, y1 - y0) }
}

export interface TargetSample {
  /** The active layer's pixels (or its targeted mask as grey) at a proxy level. */
  readonly pixels: PixelBuffer
  /** Selection coverage over the same pixels, or null without a selection. */
  readonly mask: MaskBuffer | null
}

/**
 * A small proxy of what a destructive adjustment changes: the active layer (or its targeted mask, as
 * grey) inside the selection. Null when there is nothing to measure.
 */
export function sampleOfTarget(state: DocumentState): TargetSample | null {
  const layer = activeLayerOf(state)
  if (!layer) return null
  try {
    const selection = state.selection
    if (state.editTarget === 'mask' && layer.mask) {
      const mask = layer.mask
      const area = selection ? selection.bounds : canvasRect(state)
      const level = proxyLevel(area.width, area.height, 512, maxPyramidLevel(state.width, state.height))
      const rect = levelRectOf(area, level)
      const values = readMaskLevel(mask.surface, level, {
        x: rect.x - levelOffset(mask.offsetX, level),
        y: rect.y - levelOffset(mask.offsetY, level),
        width: rect.width,
        height: rect.height,
      })
      const pixels: PixelBuffer = { width: rect.width, height: rect.height, data: new Uint8ClampedArray(rect.width * rect.height * 4) }
      for (let i = 0, p = 0; i < values.data.length; i += 1, p += 4) {
        const v = values.data[i]
        pixels.data[p] = v
        pixels.data[p + 1] = v
        pixels.data[p + 2] = v
        pixels.data[p + 3] = 255
      }
      return { pixels, mask: selection ? sampleSelection(selection, rect, level) : null }
    }
    const source = pixelSourceOf(layer)
    if (!source) return null
    const area = selection ? selection.bounds : isBackgroundLayer(layer) ? canvasRect(state) : layerPixelBounds(layer)
    if (!area) return null
    const level = proxyLevel(area.width, area.height, 512, maxPyramidLevel(state.width, state.height))
    const rect = levelRectOf(area, level)
    const pixels = readSurfaceLevel(source.surface, level, {
      x: rect.x - levelOffset(source.offsetX, level),
      y: rect.y - levelOffset(source.offsetY, level),
      width: rect.width,
      height: rect.height,
    })
    return { pixels, mask: selection ? sampleSelection(selection, rect, level) : null }
  } catch (error) {
    console.error(error)
    return null
  }
}

export function histogramOfTarget(state: DocumentState): Histogram | null {
  const sample = sampleOfTarget(state)
  return sample ? computeHistogram(sample.pixels, sample.mask) : null
}

/** The sample with unselected pixels made transparent (the auto adjustments skip transparent pixels). */
export function selectedPixels(sample: TargetSample): PixelBuffer {
  if (!sample.mask) return sample.pixels
  const data = new Uint8ClampedArray(sample.pixels.data)
  for (let i = 0, p = 3; i < sample.mask.data.length; i += 1, p += 4) if (sample.mask.data[i] === 0) data[p] = 0
  return { width: sample.pixels.width, height: sample.pixels.height, data }
}

// ---------------------------------------------------------------------------------------------
// Track backgrounds
// ---------------------------------------------------------------------------------------------

const HUE_SHIFT_TRACK = 'linear-gradient(to right, #00e5e5, #0000e5, #e500e5, #e50000, #e5e500, #00e500, #00e5e5)'
const HUE_ABSOLUTE_TRACK = 'linear-gradient(to right, #e50000, #e5e500, #00e500, #00e5e5, #0000e5, #e500e5, #e50000)'
const SATURATION_TRACK = 'linear-gradient(to right, #8c8c8c, #c45a5a 50%, #ff1f1f)'
const LIGHTNESS_TRACK = 'linear-gradient(to right, #000, #808080 50%, #fff)'
const DARK_LIGHT_TRACK = 'linear-gradient(to right, #111, #f2f2f2)'
const VIVID_TRACK = 'linear-gradient(to right, #8c8c8c, #2f9be8)'

function towards(color: string): string {
  return `linear-gradient(to right, #2b2b2b, ${color})`
}

// ---------------------------------------------------------------------------------------------
// Per-type controls
// ---------------------------------------------------------------------------------------------

type Spec<T extends AdjustmentSpec['type']> = Extract<AdjustmentSpec, { type: T }>

interface TypeProps<T extends AdjustmentSpec['type']> {
  readonly spec: Spec<T>
  readonly onChange: (next: AdjustmentSpec, phase: ChangePhase) => void
  readonly disabled?: boolean
  readonly onInteractive?: (active: boolean) => void
  readonly histogram?: Histogram | null
}

const ZERO_SHIFT: HslShift = Object.freeze({ hue: 0, saturation: 0, lightness: 0 })

const HUE_RANGES: readonly (readonly ['master' | HueRange, string])[] = [
  ['master', 'Master'], ['reds', 'Reds'], ['yellows', 'Yellows'], ['greens', 'Greens'], ['cyans', 'Cyans'], ['blues', 'Blues'], ['magentas', 'Magentas'],
]

function HueSaturationControls({ spec, onChange, disabled, onInteractive }: TypeProps<'hue-saturation'>) {
  const [range, setRange] = useState<'master' | HueRange>('master')
  const active = spec.colorize ? 'master' : range
  const shift: HslShift = active === 'master' ? spec.master : spec.ranges[active] ?? ZERO_SHIFT
  const set = (patch: Partial<HslShift>, phase: ChangePhase) => {
    const next = { ...shift, ...patch }
    if (active === 'master') onChange({ ...spec, master: next }, phase)
    else onChange({ ...spec, ranges: { ...spec.ranges, [active]: next } }, phase)
  }
  const toggleColorize = (colorize: boolean) => {
    if (colorize) {
      // Photoshop starts Colorize with saturation 25 so the effect is visible at once.
      const hue = ((spec.master.hue % 360) + 360) % 360
      onChange({ ...spec, colorize, master: { hue, saturation: spec.master.saturation > 0 ? spec.master.saturation : 25, lightness: spec.master.lightness } }, 'commit')
    } else {
      const hue = spec.master.hue > 180 ? spec.master.hue - 360 : spec.master.hue
      onChange({ ...spec, colorize, master: { ...spec.master, hue } }, 'commit')
    }
  }
  return (
    <>
      <SelectField label="Range" inline value={active} options={HUE_RANGES} disabled={disabled || spec.colorize} onChange={setRange} />
      {spec.colorize ? (
        <SliderField label="Hue" value={((spec.master.hue % 360) + 360) % 360} min={0} max={359} unit="°" track={HUE_ABSOLUTE_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(hue, phase) => set({ hue }, phase)} />
      ) : (
        <SliderField label="Hue" value={shift.hue} min={-180} max={180} unit="°" track={HUE_SHIFT_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(hue, phase) => set({ hue }, phase)} />
      )}
      <SliderField label="Saturation" value={shift.saturation} min={spec.colorize ? 0 : -100} max={100} track={SATURATION_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(saturation, phase) => set({ saturation }, phase)} />
      <SliderField label="Lightness" value={shift.lightness} min={-100} max={100} track={LIGHTNESS_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(lightness, phase) => set({ lightness }, phase)} />
      <CheckField label="Colorize" checked={spec.colorize} disabled={disabled} onChange={toggleColorize} />
    </>
  )
}

type Tone = 'shadows' | 'midtones' | 'highlights'

function ColorBalanceControls({ spec, onChange, disabled, onInteractive }: TypeProps<'color-balance'>) {
  const [tone, setTone] = useState<Tone>('midtones')
  const shift = spec[tone]
  const set = (patch: Partial<ColorBalanceShift>, phase: ChangePhase) => onChange({ ...spec, [tone]: { ...shift, ...patch } }, phase)
  return (
    <>
      <Segmented label="Tone" value={tone} options={[['shadows', 'Shadows'], ['midtones', 'Midtones'], ['highlights', 'Highlights']]} disabled={disabled} onChange={setTone} />
      <SliderField label="Cyan – Red" value={shift.cyanRed} min={-100} max={100} track="linear-gradient(to right, #00d2d2, #b8b8b8 50%, #e21b1b)" disabled={disabled} onInteractive={onInteractive} onChange={(cyanRed, phase) => set({ cyanRed }, phase)} />
      <SliderField label="Magenta – Green" value={shift.magentaGreen} min={-100} max={100} track="linear-gradient(to right, #d21bd2, #b8b8b8 50%, #1ba51b)" disabled={disabled} onInteractive={onInteractive} onChange={(magentaGreen, phase) => set({ magentaGreen }, phase)} />
      <SliderField label="Yellow – Blue" value={shift.yellowBlue} min={-100} max={100} track="linear-gradient(to right, #e0d21b, #b8b8b8 50%, #1b3ae2)" disabled={disabled} onInteractive={onInteractive} onChange={(yellowBlue, phase) => set({ yellowBlue }, phase)} />
      <CheckField label="Preserve Luminosity" checked={spec.preserveLuminosity} disabled={disabled} onChange={(preserveLuminosity) => onChange({ ...spec, preserveLuminosity }, 'commit')} />
    </>
  )
}

const BW_SLIDERS: readonly (readonly ['reds' | 'yellows' | 'greens' | 'cyans' | 'blues' | 'magentas', string, string])[] = [
  ['reds', 'Reds', '#e21b1b'], ['yellows', 'Yellows', '#e0d21b'], ['greens', 'Greens', '#1ba51b'],
  ['cyans', 'Cyans', '#00c4c4'], ['blues', 'Blues', '#1b3ae2'], ['magentas', 'Magentas', '#d21bd2'],
]

/** Photoshop's default Black & White tint. */
const DEFAULT_TINT: Rgb8 = Object.freeze({ r: 225, g: 211, b: 179 })

function BlackWhiteControls({ spec, onChange, disabled, onInteractive }: TypeProps<'black-white'>) {
  const lastTint = useRef<Rgb8>(spec.tint ?? DEFAULT_TINT)
  if (spec.tint) lastTint.current = spec.tint
  return (
    <>
      {BW_SLIDERS.map(([key, label, color]) => (
        <SliderField key={key} label={label} value={spec[key]} min={-200} max={300} unit="%" track={towards(color)} disabled={disabled} onInteractive={onInteractive} onChange={(value, phase) => onChange({ ...spec, [key]: value }, phase)} />
      ))}
      <div className="ae-inline-row">
        <CheckField label="Tint" checked={spec.tint !== null} disabled={disabled} onChange={(on) => onChange({ ...spec, tint: on ? lastTint.current : null }, 'commit')} />
        <ColorButton label="Tint colour" value={spec.tint ?? lastTint.current} disabled={disabled || spec.tint === null} onChange={(tint, phase) => onChange({ ...spec, tint }, phase)} />
      </div>
    </>
  )
}

/** Photoshop's Photo Filter presets (approximate colours). */
const PHOTO_FILTERS: readonly (readonly [string, string, Rgb8])[] = [
  ['warming-85', 'Warming Filter (85)', { r: 236, g: 138, b: 0 }],
  ['warming-lba', 'Warming Filter (LBA)', { r: 250, g: 150, b: 0 }],
  ['warming-81', 'Warming Filter (81)', { r: 235, g: 177, b: 19 }],
  ['cooling-80', 'Cooling Filter (80)', { r: 0, g: 109, b: 255 }],
  ['cooling-lbb', 'Cooling Filter (LBB)', { r: 0, g: 93, b: 255 }],
  ['cooling-82', 'Cooling Filter (82)', { r: 0, g: 181, b: 255 }],
  ['red', 'Red', { r: 234, g: 26, b: 26 }],
  ['orange', 'Orange', { r: 243, g: 132, b: 23 }],
  ['yellow', 'Yellow', { r: 249, g: 227, b: 28 }],
  ['green', 'Green', { r: 25, g: 201, b: 25 }],
  ['cyan', 'Cyan', { r: 29, g: 203, b: 234 }],
  ['blue', 'Blue', { r: 29, g: 53, b: 234 }],
  ['violet', 'Violet', { r: 155, g: 29, b: 234 }],
  ['magenta', 'Magenta', { r: 227, g: 24, b: 227 }],
  ['sepia', 'Sepia', { r: 172, g: 122, b: 51 }],
  ['deep-red', 'Deep Red', { r: 255, g: 0, b: 0 }],
  ['deep-blue', 'Deep Blue', { r: 0, g: 34, b: 205 }],
  ['deep-emerald', 'Deep Emerald', { r: 0, g: 140, b: 0 }],
  ['deep-yellow', 'Deep Yellow', { r: 255, g: 213, b: 0 }],
  ['underwater', 'Underwater', { r: 0, g: 193, b: 177 }],
]

function PhotoFilterControls({ spec, onChange, disabled, onInteractive }: TypeProps<'photo-filter'>) {
  const preset = PHOTO_FILTERS.find(([, , color]) => sameRgb(color, spec.color))?.[0] ?? 'custom'
  const options: (readonly [string, string])[] = [...PHOTO_FILTERS.map(([id, label]) => [id, label] as const), ['custom', 'Custom colour']]
  return (
    <>
      <SelectField
        label="Filter"
        value={preset}
        options={options}
        disabled={disabled}
        onChange={(id) => {
          const found = PHOTO_FILTERS.find(([key]) => key === id)
          if (found) onChange({ ...spec, color: found[2] }, 'commit')
        }}
      />
      <div className="ae-inline-row">
        <span className="ae-inline-label">Colour</span>
        <ColorButton label="Filter colour" value={spec.color} disabled={disabled} onChange={(color, phase) => onChange({ ...spec, color }, phase)} />
      </div>
      <SliderField label="Density" value={spec.density} min={0} max={100} unit="%" track={`linear-gradient(to right, #d8d8d8, ${cssColor(spec.color)})`} disabled={disabled} onInteractive={onInteractive} onChange={(density, phase) => onChange({ ...spec, density }, phase)} />
      <CheckField label="Preserve Luminosity" checked={spec.preserveLuminosity} disabled={disabled} onChange={(preserveLuminosity) => onChange({ ...spec, preserveLuminosity }, 'commit')} />
    </>
  )
}

function ThresholdControls({ spec, onChange, disabled, onInteractive, histogram }: TypeProps<'threshold'>) {
  return (
    <>
      <HistogramView histogram={histogram ?? null} channel="luma" height={64} marker={spec.level} />
      <SliderField label="Threshold Level" value={spec.level} min={1} max={255} track={DARK_LIGHT_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(level, phase) => onChange({ ...spec, level }, phase)} />
    </>
  )
}

// ----- gradient stops -------------------------------------------------------------------------

/** CSS preview of gradient stops (midpoints become CSS colour hints). */
export function gradientCss(stops: readonly GradientStop[]): string {
  const list = normalizeStops(stops)
  const parts: string[] = []
  list.forEach((stop, index) => {
    parts.push(`${cssColor(stop.color)} ${(stop.position * 100).toFixed(2)}%`)
    const next = list[index + 1]
    if (next) {
      const midpoint = stop.midpoint ?? 0.5
      if (Math.abs(midpoint - 0.5) > 0.001) parts.push(`${((stop.position + (next.position - stop.position) * midpoint) * 100).toFixed(2)}%`)
    }
  })
  if (parts.length === 1) parts.push(parts[0])
  return `linear-gradient(to right, ${parts.join(', ')})`
}

interface StopDrag {
  readonly pointerId: number
  readonly startY: number
  removing: boolean
}

function GradientStopsEditor({ stops, onChange, disabled, onInteractive }: {
  readonly stops: readonly GradientStop[]
  readonly onChange: (stops: GradientStop[], phase: ChangePhase) => void
  readonly disabled?: boolean
  readonly onInteractive?: (active: boolean) => void
}) {
  const [selected, setSelected] = useState(0)
  const [hidden, setHidden] = useState<number | null>(null)
  const stripRef = useRef<HTMLDivElement>(null)
  const dragRef = useRef<StopDrag | null>(null)
  const list = stops.length ? stops : normalizeStops(stops)
  const index = Math.min(selected, list.length - 1)
  const current = list[index]

  const sortAndSelect = (next: GradientStop[], keep: GradientStop | null): GradientStop[] => {
    const sorted = [...next].sort((a, b) => a.position - b.position)
    if (keep) setSelected(Math.max(0, sorted.indexOf(keep)))
    return sorted
  }

  const positionAt = (clientX: number) => {
    const rect = stripRef.current?.getBoundingClientRect()
    if (!rect || rect.width <= 0) return 0
    return clampNumber((clientX - rect.left) / rect.width, 0, 1)
  }

  const startDrag = (event: ReactPointerEvent<HTMLElement>, stopIndex: number) => {
    if (disabled || event.button !== 0) return
    event.preventDefault()
    event.stopPropagation()
    setSelected(stopIndex)
    const element = event.currentTarget
    try {
      element.setPointerCapture(event.pointerId)
    } catch {
      // ignore
    }
    dragRef.current = { pointerId: event.pointerId, startY: event.clientY, removing: false }
    onInteractive?.(true)
    let working = [...list]
    const move = (moveEvent: PointerEvent) => {
      const drag = dragRef.current
      if (!drag) return
      const removing = working.length > 2 && moveEvent.clientY - drag.startY > 28
      drag.removing = removing
      setHidden(removing ? stopIndex : null)
      if (removing) {
        onChange(working.filter((_, i) => i !== stopIndex), 'live')
        return
      }
      working = working.map((stop, i) => (i === stopIndex ? { ...stop, position: Number(positionAt(moveEvent.clientX).toFixed(4)) } : stop))
      onChange(working, 'live')
    }
    const end = () => {
      element.removeEventListener('pointermove', move)
      element.removeEventListener('pointerup', end)
      element.removeEventListener('pointercancel', end)
      const drag = dragRef.current
      dragRef.current = null
      setHidden(null)
      onInteractive?.(false)
      if (drag?.removing) {
        setSelected(0)
        onChange(working.filter((_, i) => i !== stopIndex), 'commit')
      } else {
        const moved = working[stopIndex]
        onChange(sortAndSelect(working, moved), 'commit')
      }
    }
    element.addEventListener('pointermove', move)
    element.addEventListener('pointerup', end)
    element.addEventListener('pointercancel', end)
  }

  const addAt = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (disabled || event.button !== 0 || event.target !== event.currentTarget) return
    const position = Number(positionAt(event.clientX).toFixed(4))
    const sample = gradientColorAt({ stops: list, opacityStops: [] }, position)
    const stop: GradientStop = { position, color: { r: Math.round(sample.r), g: Math.round(sample.g), b: Math.round(sample.b) }, midpoint: 0.5 }
    onChange(sortAndSelect([...list, stop], stop), 'commit')
  }

  const update = (patch: Partial<GradientStop>) => {
    const next = list.map((stop, i) => (i === index ? { ...stop, ...patch } : stop))
    const kept = next[index]
    onChange(patch.position !== undefined ? sortAndSelect(next, kept) : next, 'commit')
  }

  return (
    <div className="ae-stops">
      <div className="ae-stops-bar" style={{ background: gradientCss(list) }} aria-hidden="true" />
      <div ref={stripRef} className="ae-stops-strip" title="Click to add a colour stop; drag a stop down to remove it" onPointerDown={addAt}>
        {list.map((stop, i) => (
          <button
            key={i}
            type="button"
            role="slider"
            aria-label={`Colour stop ${i + 1}`}
            aria-valuemin={0}
            aria-valuemax={100}
            aria-valuenow={Math.round(stop.position * 100)}
            className={`ae-stop ${i === index ? 'is-selected' : ''}`}
            style={{ left: `${stop.position * 100}%`, visibility: hidden === i ? 'hidden' : undefined, '--ae-stop': cssColor(stop.color) } as CSSProperties}
            disabled={disabled}
            onPointerDown={(event) => startDrag(event, i)}
            onKeyDown={(event) => {
              if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
                event.preventDefault()
                const delta = (event.key === 'ArrowRight' ? 0.01 : -0.01) * (event.shiftKey ? 10 : 1)
                setSelected(i)
                const next = list.map((item, j) => (j === i ? { ...item, position: Number(clampNumber(item.position + delta, 0, 1).toFixed(4)) } : item))
                onChange(sortAndSelect(next, next[i]), 'commit')
              } else if ((event.key === 'Delete' || event.key === 'Backspace') && list.length > 2) {
                event.preventDefault()
                setSelected(0)
                onChange(list.filter((_, j) => j !== i), 'commit')
              }
            }}
          />
        ))}
      </div>
      {current && (
        <div className="ae-stop-edit">
          <ColorButton label="Stop colour" value={current.color} disabled={disabled} onChange={(color, phase) => onChange(list.map((stop, i) => (i === index ? { ...stop, color } : stop)), phase)} />
          <label className="ae-mini-field" title="Location">
            <span>Location</span>
            <NumberInput label="Stop location in percent" value={Math.round(current.position * 1000) / 10} min={0} max={100} step={1} digits={1} disabled={disabled} onChange={(value) => update({ position: value / 100 })} />
            <i>%</i>
          </label>
          {index < list.length - 1 && (
            <label className="ae-mini-field" title="Midpoint between this stop and the next">
              <span>Mid</span>
              <NumberInput label="Midpoint in percent" value={Math.round((current.midpoint ?? 0.5) * 100)} min={5} max={95} disabled={disabled} onChange={(value) => update({ midpoint: value / 100 })} />
              <i>%</i>
            </label>
          )}
          <button type="button" className="ae-icon-mini" title="Delete this stop" aria-label="Delete stop" disabled={disabled || list.length <= 2} onClick={() => { setSelected(0); onChange(list.filter((_, i) => i !== index), 'commit') }}>
            <Trash2 aria-hidden="true" />
          </button>
        </div>
      )}
    </div>
  )
}

function GradientMapControls({ spec, onChange, disabled, onInteractive }: TypeProps<'gradient-map'>) {
  return (
    <>
      <GradientStopsEditor stops={spec.stops} disabled={disabled} onInteractive={onInteractive} onChange={(stops, phase) => onChange({ ...spec, stops }, phase)} />
      <div className="ae-inline-row">
        <CheckField label="Reverse" checked={spec.reverse} disabled={disabled} onChange={(reverse) => onChange({ ...spec, reverse }, 'commit')} />
        <CheckField label="Dither" checked={spec.dither} disabled={disabled} onChange={(dither) => onChange({ ...spec, dither }, 'commit')} />
      </div>
    </>
  )
}

// ---------------------------------------------------------------------------------------------
// The switch
// ---------------------------------------------------------------------------------------------

export interface AdjustmentControlsProps {
  readonly spec: AdjustmentSpec
  readonly onChange: (next: AdjustmentSpec, phase: ChangePhase) => void
  /** Levels, Curves and Threshold draw it (null hides the graph's histogram). */
  readonly histogram?: Histogram | null
  readonly disabled?: boolean
  readonly onInteractive?: (active: boolean) => void
}

export function AdjustmentControls({ spec, onChange, histogram, disabled, onInteractive }: AdjustmentControlsProps): ReactNode {
  const common = { onChange, disabled, onInteractive, histogram }
  let body: ReactNode
  switch (spec.type) {
    case 'brightness-contrast':
      body = (
        <>
          <SliderField label="Brightness" value={spec.brightness} min={-150} max={150} track={DARK_LIGHT_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(brightness, phase) => onChange({ ...spec, brightness }, phase)} />
          <SliderField label="Contrast" value={spec.contrast} min={-50} max={100} disabled={disabled} onInteractive={onInteractive} onChange={(contrast, phase) => onChange({ ...spec, contrast }, phase)} />
          <CheckField label="Use Legacy" title="Photoshop CS2 behaviour: shifts and stretches values linearly" checked={spec.legacy} disabled={disabled} onChange={(legacy) => onChange({ ...spec, legacy }, 'commit')} />
        </>
      )
      break
    case 'levels':
      body = <LevelsEditor spec={spec} {...common} />
      break
    case 'curves':
      body = <CurvesEditor spec={spec} {...common} />
      break
    case 'exposure':
      body = (
        <>
          <SliderField label="Exposure" value={spec.exposure} min={-20} max={20} step={0.01} digits={2} fixed track={DARK_LIGHT_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(exposure, phase) => onChange({ ...spec, exposure }, phase)} />
          <SliderField label="Offset" value={spec.offset} min={-0.5} max={0.5} step={0.0001} digits={4} fixed disabled={disabled} onInteractive={onInteractive} onChange={(offset, phase) => onChange({ ...spec, offset }, phase)} />
          <SliderField label="Gamma Correction" value={spec.gamma} min={0.01} max={9.99} step={0.01} digits={2} fixed disabled={disabled} onInteractive={onInteractive} onChange={(gamma, phase) => onChange({ ...spec, gamma }, phase)} />
        </>
      )
      break
    case 'vibrance':
      body = (
        <>
          <SliderField label="Vibrance" value={spec.vibrance} min={-100} max={100} track={VIVID_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(vibrance, phase) => onChange({ ...spec, vibrance }, phase)} />
          <SliderField label="Saturation" value={spec.saturation} min={-100} max={100} track={SATURATION_TRACK} disabled={disabled} onInteractive={onInteractive} onChange={(saturation, phase) => onChange({ ...spec, saturation }, phase)} />
        </>
      )
      break
    case 'hue-saturation':
      body = <HueSaturationControls spec={spec} {...common} />
      break
    case 'color-balance':
      body = <ColorBalanceControls spec={spec} {...common} />
      break
    case 'black-white':
      body = <BlackWhiteControls spec={spec} {...common} />
      break
    case 'photo-filter':
      body = <PhotoFilterControls spec={spec} {...common} />
      break
    case 'invert':
      body = <p className="ae-note">Invert has no settings: every colour becomes its opposite.</p>
      break
    case 'posterize':
      body = <SliderField label="Levels" value={spec.levels} min={2} max={255} disabled={disabled} onInteractive={onInteractive} onChange={(levels, phase) => onChange({ ...spec, levels }, phase)} />
      break
    case 'threshold':
      body = <ThresholdControls spec={spec} {...common} />
      break
    case 'gradient-map':
      body = <GradientMapControls spec={spec} {...common} />
      break
    default:
      body = <p className="ae-note">This adjustment has no settings here.</p>
  }
  return <div className="ae-adjust">{body}</div>
}
