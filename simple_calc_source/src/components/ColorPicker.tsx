import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { Ban, Check, ChevronDown, Palette, Pipette } from 'lucide-react'
import type { SpreadsheetColor } from '../spreadsheet-types'
import {
  DEFAULT_THEME_COLORS, argbColor, colorLabel, hexToRgb, hslToRgb, hsvToRgb, normalizeHex, resolveColorHex,
  rgbToHex, rgbToHsl, rgbToHsv, sameColor, standardPalette, themePalette,
} from '../lib/cell-styles'
import type { ColorValue, Hsv, PaletteSwatch } from '../lib/cell-styles'
import './format-ui.css'

// ---------------------------------------------------------------------------------------
// Shared popover primitive
// ---------------------------------------------------------------------------------------

export interface AnchorRect {
  left: number
  top: number
  right: number
  bottom: number
}

const FOCUSABLE = 'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'

// Open popovers, oldest first. A press inside a later (nested) popover must not close an
// earlier one, even though the nested popover is portalled outside it.
const popoverStack: HTMLElement[] = []

export interface FormatPopoverProps {
  anchor: AnchorRect
  onClose: () => void
  children: ReactNode
  label: string
  className?: string
  /** Presses on this element (usually the toggle button) don't count as outside presses. */
  ignoreElement?: HTMLElement | null
  /** Initial focus target inside the popover (selector); defaults to the first focusable control. */
  initialFocus?: string
  /** Align the popover's right edge with the anchor's instead of its left edge. */
  align?: 'start' | 'end'
}

/** A floating panel anchored to a rect: portalled, clamped to the viewport, closes on Escape / outside press. */
export function FormatPopover({ anchor, onClose, children, label, className, ignoreElement, initialFocus, align = 'start' }: FormatPopoverProps) {
  const ref = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState<{ left: number; top: number } | null>(null)
  const onCloseRef = useRef(onClose)
  onCloseRef.current = onClose

  const place = useCallback(() => {
    const element = ref.current
    if (!element) return
    const rect = element.getBoundingClientRect()
    const margin = 6
    const width = window.innerWidth
    const height = window.innerHeight
    let left = align === 'end' ? anchor.right - rect.width : anchor.left
    left = Math.min(Math.max(margin, left), Math.max(margin, width - margin - rect.width))
    let top = anchor.bottom + 3
    if (top + rect.height > height - margin) {
      const above = anchor.top - 3 - rect.height
      top = above >= margin ? above : Math.max(margin, height - margin - rect.height)
    }
    setPosition((current) => (current && current.left === left && current.top === top ? current : { left, top }))
  }, [align, anchor.bottom, anchor.left, anchor.right, anchor.top])

  useLayoutEffect(() => {
    place()
    const element = ref.current
    if (!element || typeof ResizeObserver === 'undefined') return
    const observer = new ResizeObserver(() => place())
    observer.observe(element)
    return () => observer.disconnect()
  }, [place])

  useEffect(() => {
    const element = ref.current
    if (!element) return
    popoverStack.push(element)
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const frame = window.requestAnimationFrame(() => {
      const target = (initialFocus && element.querySelector<HTMLElement>(initialFocus)) || element.querySelector<HTMLElement>(FOCUSABLE)
      target?.focus()
    })
    const handlePointer = (event: PointerEvent) => {
      const target = event.target as Node | null
      if (!target) return
      const index = popoverStack.indexOf(element)
      if (popoverStack.slice(Math.max(0, index)).some((popover) => popover.contains(target))) return
      if (ignoreElement?.contains(target)) return
      onCloseRef.current()
    }
    const handleResize = () => onCloseRef.current()
    document.addEventListener('pointerdown', handlePointer, true)
    window.addEventListener('resize', handleResize)
    return () => {
      window.cancelAnimationFrame(frame)
      document.removeEventListener('pointerdown', handlePointer, true)
      window.removeEventListener('resize', handleResize)
      const index = popoverStack.indexOf(element)
      if (index >= 0) popoverStack.splice(index, 1)
      if (previous?.isConnected) previous.focus()
    }
    // Mount-only: focus handling and listeners belong to this popover instance.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [])

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onClose()
      return
    }
    if (event.key === 'Tab' && ref.current) {
      const controls = Array.from(ref.current.querySelectorAll<HTMLElement>(FOCUSABLE)).filter((control) => control.getClientRects().length > 0)
      if (controls.length) {
        const first = controls[0]
        const last = controls[controls.length - 1]
        if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
        else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
      }
    }
    // Keys typed in a popover never reach the grid or the dialog underneath.
    event.stopPropagation()
  }

  // Until measured, the panel is laid out transparent (not visibility:hidden, whose removal
  // Chromium can fail to propagate to deep descendants).
  const style: CSSProperties = position ? { left: position.left, top: position.top } : { left: anchor.left, top: anchor.bottom + 3, opacity: 0, pointerEvents: 'none' }
  return createPortal(
    <div ref={ref} className={`fmt-popover${className ? ` ${className}` : ''}`} role="dialog" aria-label={label} style={style} onKeyDown={handleKeyDown}>
      {children}
    </div>,
    document.body,
  )
}

/** Moves focus among elements with arrow keys by on-screen geometry (works for ragged grids). */
export function moveFocusByArrow(event: ReactKeyboardEvent, items: HTMLElement[]): HTMLElement | null {
  const current = items.indexOf(document.activeElement as HTMLElement)
  if (current < 0) return null
  const key = event.key
  if (key === 'Home' || key === 'End') {
    event.preventDefault()
    const target = key === 'Home' ? items[0] : items[items.length - 1]
    target.focus()
    return target
  }
  if (!['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(key)) return null
  event.preventDefault()
  const from = items[current].getBoundingClientRect()
  const cx = from.left + from.width / 2
  const cy = from.top + from.height / 2
  let best: HTMLElement | null = null
  let bestScore = Infinity
  for (const item of items) {
    if (item === items[current]) continue
    const rect = item.getBoundingClientRect()
    const x = rect.left + rect.width / 2
    const y = rect.top + rect.height / 2
    const dx = x - cx
    const dy = y - cy
    let primary: number
    let secondary: number
    if (key === 'ArrowRight') { if (dx <= 1 || Math.abs(dy) > from.height / 2) continue; primary = dx; secondary = Math.abs(dy) }
    else if (key === 'ArrowLeft') { if (dx >= -1 || Math.abs(dy) > from.height / 2) continue; primary = -dx; secondary = Math.abs(dy) }
    else if (key === 'ArrowDown') { if (dy <= 1) continue; primary = dy; secondary = Math.abs(dx) }
    else { if (dy >= -1) continue; primary = -dy; secondary = Math.abs(dx) }
    const score = primary * 1000 + secondary * (key === 'ArrowDown' || key === 'ArrowUp' ? 10 : 1)
    if (score < bestScore) { bestScore = score; best = item }
  }
  // Wrap horizontally to the next / previous row, like Excel's palettes.
  if (!best && (key === 'ArrowRight' || key === 'ArrowLeft')) {
    best = items[key === 'ArrowRight' ? Math.min(items.length - 1, current + 1) : Math.max(0, current - 1)]
  }
  best?.focus()
  return best
}

// ---------------------------------------------------------------------------------------
// Recent colours
// ---------------------------------------------------------------------------------------

const RECENT_KEY = 'simple-calc.recent-colors'
const RECENT_LIMIT = 10

export function readRecentColors(): string[] {
  try {
    const parsed = JSON.parse(window.localStorage.getItem(RECENT_KEY) || '[]')
    return Array.isArray(parsed) ? parsed.map((value) => normalizeHex(String(value))).filter((value): value is string => Boolean(value)).slice(0, RECENT_LIMIT) : []
  } catch {
    return []
  }
}

export function rememberRecentColor(hex: string) {
  const value = normalizeHex(hex)
  if (!value) return
  try {
    const next = [value, ...readRecentColors().filter((entry) => entry !== value)].slice(0, RECENT_LIMIT)
    window.localStorage.setItem(RECENT_KEY, JSON.stringify(next))
  } catch {
    // Storage can be unavailable (private mode, quota); recents are a convenience only.
  }
}

interface EyeDropperResult { sRGBHex: string }
interface EyeDropperConstructor { new(): { open: () => Promise<EyeDropperResult> } }

function eyeDropper(): EyeDropperConstructor | null {
  const candidate = (window as unknown as { EyeDropper?: EyeDropperConstructor }).EyeDropper
  return typeof candidate === 'function' ? candidate : null
}

async function pickScreenColor(): Promise<string | null> {
  const Dropper = eyeDropper()
  if (!Dropper) return null
  try {
    const result = await new Dropper().open()
    return normalizeHex(result.sRGBHex)
  } catch {
    return null
  }
}

// ---------------------------------------------------------------------------------------
// Colour picker
// ---------------------------------------------------------------------------------------

export type ColorPickerMode = 'text' | 'fill'

export interface ColorPickerProps {
  value?: ColorValue | null
  /** A theme swatch yields `{ theme, tint }`; standard/custom colours `{ argb }`; Automatic / No Fill yields null. */
  onChange: (color: SpreadsheetColor | null) => void
  /** Workbook theme palette: 12 hex strings (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink). */
  themeColors?: readonly string[]
  /** 'text' offers "Automatic"; 'fill' offers "No Fill". */
  mode?: ColorPickerMode
  /** Popover anchor. Omit (or pass null) to render inline. */
  anchor?: AnchorRect | null
  /** Popover only: called on Escape, outside press, and after a colour is picked. */
  onClose?: () => void
  ignoreElement?: HTMLElement | null
  /** CSS colour drawn in the Automatic swatch (text mode). */
  automaticColor?: string
  /** Accessible name for the palette. */
  label?: string
  /** Label for the null option; defaults to "Automatic" / "No Fill". */
  nullLabel?: string
}

/** Excel-style colour picker: Automatic / No Fill, theme grid with tints, standard colours, recent colours, More Colors… */
export function ColorPicker(props: ColorPickerProps) {
  const { anchor, onClose, ignoreElement, label, mode = 'text' } = props
  const popover = Boolean(anchor)
  const panel = (
    <ColorPickerPanel
      {...props}
      onChange={(color) => {
        props.onChange(color)
        if (popover) onClose?.()
      }}
    />
  )
  if (!anchor) return panel
  return (
    <FormatPopover anchor={anchor} onClose={() => onClose?.()} ignoreElement={ignoreElement} label={label || (mode === 'fill' ? 'Fill color' : 'Font color')} initialFocus='[data-fmt-swatch][tabindex="0"]'>
      {panel}
    </FormatPopover>
  )
}

function ColorPickerPanel({ value, onChange, themeColors = DEFAULT_THEME_COLORS, mode = 'text', automaticColor = '#000000', label, nullLabel }: ColorPickerProps) {
  const [view, setView] = useState<'palette' | 'custom'>('palette')
  const [recent, setRecent] = useState<string[]>(() => readRecentColors())
  const palette = useMemo(() => themePalette(themeColors), [themeColors])
  const standard = useMemo(() => standardPalette(), [])
  const recentSwatches = useMemo<PaletteSwatch[]>(() => recent.map((hex) => ({ color: argbColor(hex), hex: `#${hex}`, label: `Recent color #${hex}` })), [recent])
  const gridRef = useRef<HTMLDivElement>(null)
  const currentHex = resolveColorHex(value ?? null, themeColors)
  const canDrop = typeof window !== 'undefined' && Boolean(eyeDropper())

  const pickCustom = (hex: string) => {
    rememberRecentColor(hex)
    setRecent(readRecentColors())
    onChange(argbColor(hex))
  }

  if (view === 'custom') {
    return (
      <CustomColorEditor
        initialHex={currentHex || (mode === 'fill' ? '#FFFFFF' : '#000000')}
        onCancel={() => setView('palette')}
        onApply={(hex) => { setView('palette'); pickCustom(hex) }}
      />
    )
  }

  const rows: Array<{ key: string; label?: string; swatches: PaletteSwatch[]; gapBefore?: boolean }> = [
    { key: 'theme-0', label: 'Theme Colors', swatches: palette[0] },
    ...palette.slice(1).map((swatches, index) => ({ key: `theme-${index + 1}`, swatches, gapBefore: index === 0 })),
    { key: 'standard', label: 'Standard Colors', swatches: standard },
    ...(recentSwatches.length ? [{ key: 'recent', label: 'Recent Colors', swatches: recentSwatches }] : []),
  ]
  const flat = rows.flatMap((row) => row.swatches)
  let selectedIndex = flat.findIndex((swatch) => sameColor(swatch.color, value ?? null, themeColors))
  if (selectedIndex < 0 && currentHex) selectedIndex = flat.findIndex((swatch) => swatch.hex === currentHex)
  const focusIndex = selectedIndex >= 0 ? selectedIndex : 0

  const handleGridKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const items = Array.from(gridRef.current?.querySelectorAll<HTMLElement>('[data-fmt-swatch]') || [])
    const moved = moveFocusByArrow(event, items)
    if (moved) {
      items.forEach((item) => item.setAttribute('tabindex', item === moved ? '0' : '-1'))
    }
  }

  let flatIndex = -1
  const nullText = nullLabel || (mode === 'fill' ? 'No Fill' : 'Automatic')
  return (
    <div className="fmt-color-panel" aria-label={label}>
      <button type="button" className={`fmt-color-auto${!value ? ' is-selected' : ''}`} aria-pressed={!value} onClick={() => onChange(null)}>
        {mode === 'fill'
          ? <span className="fmt-color-auto-swatch is-none" aria-hidden="true"><Ban size={12} /></span>
          : <span className="fmt-color-auto-swatch" aria-hidden="true" style={{ background: automaticColor }} />}
        <span>{nullText}</span>
      </button>
      <div ref={gridRef} className="fmt-color-grid" role="group" aria-label="Colors" onKeyDown={handleGridKey}>
        {rows.map((row) => (
          <div key={row.key} className={`fmt-color-row${row.gapBefore ? ' has-gap' : ''}${row.key.startsWith('theme-') && row.key !== 'theme-0' ? ' is-tint' : ''}`}>
            {row.label && <div className="fmt-color-heading">{row.label}</div>}
            <div className="fmt-color-swatches">
              {row.swatches.map((swatch) => {
                flatIndex += 1
                const selected = flatIndex === selectedIndex
                return (
                  <button
                    key={`${row.key}-${swatch.label}-${swatch.hex}`}
                    type="button"
                    data-fmt-swatch=""
                    tabIndex={flatIndex === focusIndex ? 0 : -1}
                    className={`fmt-swatch${selected ? ' is-selected' : ''}`}
                    style={{ background: swatch.hex }}
                    title={`${swatch.label} (${swatch.hex})`}
                    aria-label={swatch.label}
                    aria-pressed={selected}
                    onClick={() => onChange(swatch.color)}
                  />
                )
              })}
            </div>
          </div>
        ))}
      </div>
      <div className="fmt-color-footer">
        <button type="button" className="fmt-color-more" onClick={() => setView('custom')}>
          <Palette size={14} aria-hidden="true" />
          <span>More Colors…</span>
        </button>
        {canDrop && (
          <button type="button" className="fmt-icon-button" title="Pick a color from the screen" aria-label="Pick a color from the screen" onClick={() => { void pickScreenColor().then((hex) => { if (hex) pickCustom(hex) }) }}>
            <Pipette size={14} aria-hidden="true" />
          </button>
        )}
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------------------
// More Colors… editor
// ---------------------------------------------------------------------------------------

interface CustomColorEditorProps {
  initialHex: string
  onApply: (hex: string) => void
  onCancel: () => void
}

const round = (value: number) => Math.round(value)

export function CustomColorEditor({ initialHex, onApply, onCancel }: CustomColorEditorProps) {
  const initial = normalizeHex(initialHex) || '000000'
  const [hsv, setHsv] = useState<Hsv>(() => rgbToHsv(hexToRgb(initial)))
  const [hexDraft, setHexDraft] = useState(`#${initial}`)
  const hex = rgbToHex(hsvToRgb(hsv))
  const rgb = hexToRgb(hex)
  const hsl = rgbToHsl(rgb)
  const svRef = useRef<HTMLDivElement>(null)
  const hueRef = useRef<HTMLDivElement>(null)
  const canDrop = typeof window !== 'undefined' && Boolean(eyeDropper())

  useEffect(() => { setHexDraft(`#${hex}`) }, [hex])

  const setFromHex = (value: string) => {
    const next = normalizeHex(value)
    if (!next) return false
    const nextHsv = rgbToHsv(hexToRgb(next))
    // Keep the hue for greys so the square doesn't jump back to red.
    setHsv((current) => ({ ...nextHsv, h: nextHsv.s === 0 ? current.h : nextHsv.h }))
    return true
  }

  const setRgbChannel = (channel: 'r' | 'g' | 'b', raw: string) => {
    const value = Math.min(255, Math.max(0, Math.round(Number(raw) || 0)))
    setFromHex(rgbToHex({ ...rgb, [channel]: value }))
  }

  const setHslChannel = (channel: 'h' | 's' | 'l', raw: string) => {
    const number = Number(raw) || 0
    const next = { ...hsl }
    if (channel === 'h') next.h = ((number % 360) + 360) % 360
    else next[channel] = Math.min(100, Math.max(0, number)) / 100
    const nextHsv = rgbToHsv(hslToRgb(next))
    setHsv({ ...nextHsv, h: channel === 'h' ? next.h : nextHsv.s > 0 ? nextHsv.h : hsv.h })
  }

  const dragSv = (event: ReactPointerEvent<HTMLDivElement>) => {
    const rect = svRef.current?.getBoundingClientRect()
    if (!rect) return
    const s = Math.min(1, Math.max(0, (event.clientX - rect.left) / rect.width))
    const v = Math.min(1, Math.max(0, 1 - (event.clientY - rect.top) / rect.height))
    setHsv((current) => ({ ...current, s, v }))
  }

  const dragHue = (event: ReactPointerEvent<HTMLDivElement>) => {
    const rect = hueRef.current?.getBoundingClientRect()
    if (!rect) return
    const h = Math.min(359.9, Math.max(0, ((event.clientX - rect.left) / rect.width) * 360))
    setHsv((current) => ({ ...current, h }))
  }

  const pointerHandlers = (drag: (event: ReactPointerEvent<HTMLDivElement>) => void) => ({
    onPointerDown: (event: ReactPointerEvent<HTMLDivElement>) => {
      event.preventDefault()
      event.currentTarget.setPointerCapture(event.pointerId)
      event.currentTarget.focus()
      drag(event)
    },
    onPointerMove: (event: ReactPointerEvent<HTMLDivElement>) => {
      if (event.currentTarget.hasPointerCapture(event.pointerId)) drag(event)
    },
  })

  const svKeys = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const step = event.shiftKey ? 0.1 : 0.01
    const delta: Record<string, [number, number]> = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, step], ArrowDown: [0, -step] }
    const move = delta[event.key]
    if (!move) return
    event.preventDefault()
    event.stopPropagation()
    setHsv((current) => ({ ...current, s: Math.min(1, Math.max(0, current.s + move[0])), v: Math.min(1, Math.max(0, current.v + move[1])) }))
  }

  const hueKeys = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const step = event.shiftKey ? 15 : 1
    let next: number | null = null
    if (event.key === 'ArrowRight' || event.key === 'ArrowUp') next = hsv.h + step
    else if (event.key === 'ArrowLeft' || event.key === 'ArrowDown') next = hsv.h - step
    else if (event.key === 'Home') next = 0
    else if (event.key === 'End') next = 359
    if (next === null) return
    event.preventDefault()
    event.stopPropagation()
    setHsv((current) => ({ ...current, h: Math.min(359.9, Math.max(0, next!)) }))
  }

  const submit = () => onApply(hex)

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
      event.preventDefault()
      event.stopPropagation()
      const typed = event.target instanceof HTMLInputElement && event.target.name === 'hex' ? normalizeHex(hexDraft) : null
      onApply(typed || hex)
    } else if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onCancel()
    }
  }

  const hueColor = `hsl(${round(hsv.h)} 100% 50%)`
  return (
    <div className="fmt-custom-color" onKeyDown={handleKeyDown}>
      <div className="fmt-custom-title">Custom color</div>
      <div
        ref={svRef}
        className="fmt-sv"
        style={{ backgroundColor: hueColor }}
        role="slider"
        tabIndex={0}
        aria-label="Saturation and brightness"
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuenow={round(hsv.s * 100)}
        aria-valuetext={`Saturation ${round(hsv.s * 100)}%, brightness ${round(hsv.v * 100)}%`}
        onKeyDown={svKeys}
        {...pointerHandlers(dragSv)}
      >
        <span className="fmt-sv-handle" style={{ left: `${hsv.s * 100}%`, top: `${(1 - hsv.v) * 100}%`, background: `#${hex}` }} />
      </div>
      <div
        ref={hueRef}
        className="fmt-hue"
        role="slider"
        tabIndex={0}
        aria-label="Hue"
        aria-valuemin={0}
        aria-valuemax={359}
        aria-valuenow={round(hsv.h)}
        onKeyDown={hueKeys}
        {...pointerHandlers(dragHue)}
      >
        <span className="fmt-hue-handle" style={{ left: `${(hsv.h / 360) * 100}%`, background: hueColor }} />
      </div>
      <div className="fmt-custom-fields">
        <label className="fmt-field fmt-field-hex">
          <span>Hex</span>
          <input
            name="hex"
            value={hexDraft}
            spellCheck={false}
            maxLength={7}
            onChange={(event) => { setHexDraft(event.target.value); if (/^#?[0-9a-f]{6}$/i.test(event.target.value.trim())) setFromHex(event.target.value) }}
            onBlur={(event) => { if (!setFromHex(event.target.value)) setHexDraft(`#${hex}`) }}
          />
        </label>
        {(['r', 'g', 'b'] as const).map((channel) => (
          <label key={channel} className="fmt-field fmt-field-channel">
            <span>{channel.toUpperCase()}</span>
            <input type="number" min={0} max={255} value={rgb[channel]} onChange={(event) => setRgbChannel(channel, event.target.value)} aria-label={{ r: 'Red', g: 'Green', b: 'Blue' }[channel]} />
          </label>
        ))}
        <span className="fmt-field-spacer" />
        {(['h', 's', 'l'] as const).map((channel) => (
          <label key={channel} className="fmt-field fmt-field-channel">
            <span>{channel.toUpperCase()}</span>
            <input
              type="number"
              min={0}
              max={channel === 'h' ? 359 : 100}
              value={channel === 'h' ? round(hsl.h) : round(hsl[channel] * 100)}
              onChange={(event) => setHslChannel(channel, event.target.value)}
              aria-label={{ h: 'Hue (degrees)', s: 'Saturation (percent)', l: 'Lightness (percent)' }[channel]}
            />
          </label>
        ))}
      </div>
      <div className="fmt-custom-compare" aria-hidden="true">
        <span style={{ background: `#${hex}` }}><small>New</small></span>
        <span style={{ background: `#${initial}` }}><small>Current</small></span>
      </div>
      <div className="fmt-custom-actions">
        {canDrop && (
          <button type="button" className="fmt-icon-button" title="Pick a color from the screen" aria-label="Pick a color from the screen" onClick={() => { void pickScreenColor().then((picked) => { if (picked) setFromHex(picked) }) }}>
            <Pipette size={14} aria-hidden="true" />
          </button>
        )}
        <span className="fmt-grow" />
        <button type="button" className="fmt-button" onClick={onCancel}>Cancel</button>
        <button type="button" className="fmt-button is-primary" onClick={submit}><Check size={13} aria-hidden="true" />OK</button>
      </div>
    </div>
  )
}

// ---------------------------------------------------------------------------------------
// Dropdown button
// ---------------------------------------------------------------------------------------

export interface ColorDropdownButtonProps {
  value?: ColorValue | null
  onChange: (color: SpreadsheetColor | null) => void
  themeColors?: readonly string[]
  mode?: ColorPickerMode
  /** Accessible name, e.g. "Font color". */
  label: string
  id?: string
  disabled?: boolean
  automaticColor?: string
  nullLabel?: string
  className?: string
  /** Show only the swatch bar and chevron (toolbar density). */
  compact?: boolean
  icon?: ReactNode
}

/** A button showing the current colour that opens a ColorPicker popover beneath it. */
export function ColorDropdownButton({ value, onChange, themeColors = DEFAULT_THEME_COLORS, mode = 'text', label, id, disabled, automaticColor = '#000000', nullLabel, className, compact, icon }: ColorDropdownButtonProps) {
  const [anchor, setAnchor] = useState<AnchorRect | null>(null)
  const buttonRef = useRef<HTMLButtonElement>(null)
  const hex = resolveColorHex(value ?? null, themeColors)
  const text = value ? colorLabel(value, themeColors) : (nullLabel || (mode === 'fill' ? 'No Color' : 'Automatic'))
  const toggle = () => {
    if (anchor) { setAnchor(null); return }
    const rect = buttonRef.current?.getBoundingClientRect()
    if (rect) setAnchor({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom })
  }
  return (
    <>
      <button
        ref={buttonRef}
        id={id}
        type="button"
        className={`fmt-color-button${compact ? ' is-compact' : ''}${className ? ` ${className}` : ''}`}
        disabled={disabled}
        aria-haspopup="dialog"
        aria-expanded={Boolean(anchor)}
        aria-label={`${label}: ${text}`}
        title={`${label}: ${text}`}
        onClick={toggle}
        onKeyDown={(event) => { if (event.key === 'ArrowDown' && (event.altKey || !anchor)) { event.preventDefault(); if (!anchor) toggle() } }}
      >
        {icon}
        <span
          className={`fmt-color-button-swatch${!hex && mode === 'fill' ? ' is-none' : ''}`}
          style={{ background: hex || (mode === 'fill' ? undefined : automaticColor) }}
          aria-hidden="true"
        />
        {!compact && <span className="fmt-color-button-label">{text}</span>}
        <ChevronDown size={12} aria-hidden="true" className="fmt-chevron" />
      </button>
      {anchor && (
        <ColorPicker
          anchor={anchor}
          ignoreElement={buttonRef.current}
          value={value}
          onChange={onChange}
          onClose={() => setAnchor(null)}
          themeColors={themeColors}
          mode={mode}
          automaticColor={automaticColor}
          nullLabel={nullLabel}
          label={label}
        />
      )}
    </>
  )
}
