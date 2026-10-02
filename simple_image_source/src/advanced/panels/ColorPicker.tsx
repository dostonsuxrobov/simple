// src/advanced/panels/ColorPicker.tsx (WP6)
// Photoshop-style colour picking for the Advanced editor (design 5.14):
//   - ColorPicker: a saturation / brightness field and a hue strip, H S B / R G B / hex fields, the new and
//     the current colour side by side (click the current one to go back), and "pick from screen".
//   - ColorPickerPopover: the picker anchored to a control. Changes apply live ('live' while dragging,
//     'commit' when a drag or a typed value ends); Enter or a click outside keeps the colour, Escape puts
//     the original back. Keys typed in it never reach the editor's shortcuts.
//   - ColorButton: a swatch button that opens the popover (text colour, shape fill, gradient stops ...).
//   - pickScreenColor(): the EyeDropper API (samples anywhere on screen, also outside the window); it
//     needs a click to start and resolves null when cancelled with Escape or unavailable.
// A leaf module (imports only shared helpers), so every panel and dialog can use it without import cycles.
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { Pipette } from 'lucide-react'
import type { Rgb8 } from '../../imaging/types.ts'
import { hsvToRgb, parseHex, rgbToHsv, toHex } from '../../imaging/color.ts'

export type ColorPhase = 'live' | 'commit'

// ---------------------------------------------------------------------------------------------
// Screen colour (EyeDropper API)
// ---------------------------------------------------------------------------------------------

interface EyeDropperInstance {
  open(options?: { readonly signal?: AbortSignal }): Promise<{ readonly sRGBHex: string }>
}
type EyeDropperConstructor = new () => EyeDropperInstance

function eyeDropperConstructor(): EyeDropperConstructor | null {
  const ctor = (globalThis as { EyeDropper?: unknown }).EyeDropper
  return typeof ctor === 'function' ? ctor as EyeDropperConstructor : null
}

/** True when "pick from screen" can work here (Chromium's EyeDropper API). */
export function eyeDropperAvailable(): boolean {
  return eyeDropperConstructor() !== null
}

/** "#rrggbb" or "rgb(r, g, b)" -> Rgb8. */
function parseCssColor(text: string): Rgb8 | null {
  const hex = parseHex(text)
  if (hex) return hex
  const match = /rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)/i.exec(text)
  if (!match) return null
  const channel = (value: string) => Math.max(0, Math.min(255, Math.round(Number(value))))
  return { r: channel(match[1]), g: channel(match[2]), b: channel(match[3]) }
}

/** Samples a colour anywhere on screen. Call from a click handler. Null when cancelled or unavailable. */
export async function pickScreenColor(): Promise<Rgb8 | null> {
  const Ctor = eyeDropperConstructor()
  if (!Ctor) return null
  try {
    const result = await new Ctor().open()
    return result && typeof result.sRGBHex === 'string' ? parseCssColor(result.sRGBHex) : null
  } catch {
    // AbortError: the user pressed Escape. Anything else: the picker could not start.
    return null
  }
}

// ---------------------------------------------------------------------------------------------
// Colour maths
// ---------------------------------------------------------------------------------------------

interface Hsv {
  /** 0..360 */
  readonly h: number
  /** 0..1 */
  readonly s: number
  /** 0..1 */
  readonly v: number
}

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value
}

function hsvOf(color: Rgb8, keepHue?: number): Hsv {
  const [h, s, v] = rgbToHsv(color.r / 255, color.g / 255, color.b / 255)
  // Grey has no hue: keep the hue the user chose so the field does not jump back to red.
  return { h: s === 0 && keepHue !== undefined ? keepHue : h, s, v }
}

function rgbOf(hsv: Hsv): Rgb8 {
  const [r, g, b] = hsvToRgb(hsv.h, hsv.s, hsv.v)
  return { r: Math.round(r * 255), g: Math.round(g * 255), b: Math.round(b * 255) }
}

export function sameRgb(a: Rgb8 | null | undefined, b: Rgb8 | null | undefined): boolean {
  return Boolean(a && b && a.r === b.r && a.g === b.g && a.b === b.b)
}

export function cssColor(color: Rgb8): string {
  return `rgb(${color.r}, ${color.g}, ${color.b})`
}

// ---------------------------------------------------------------------------------------------
// A small numeric field (kept local: this module must not import the panel controls)
// ---------------------------------------------------------------------------------------------

function ChannelInput({ label, value, min, max, unit, disabled, onCommit }: {
  readonly label: string
  readonly value: number
  readonly min: number
  readonly max: number
  readonly unit?: string
  readonly disabled?: boolean
  readonly onCommit: (value: number) => void
}) {
  const [text, setText] = useState(String(Math.round(value)))
  const editing = useRef(false)
  useEffect(() => {
    if (!editing.current) setText(String(Math.round(value)))
  }, [value])
  const commit = () => {
    editing.current = false
    const parsed = Number(text.replace(',', '.'))
    if (text.trim() !== '' && Number.isFinite(parsed)) onCommit(clamp(Math.round(parsed), min, max))
    else setText(String(Math.round(value)))
  }
  return (
    <label className="ae-cp-field">
      <span>{label}</span>
      <input
        type="text"
        inputMode="numeric"
        aria-label={label}
        value={text}
        disabled={disabled}
        onFocus={(event) => {
          editing.current = true
          event.currentTarget.select()
        }}
        onChange={(event) => setText(event.currentTarget.value)}
        onBlur={commit}
        onKeyDown={(event) => {
          if (event.key === 'Enter') commit()
          else if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
            event.preventDefault()
            const next = clamp(Math.round(value) + (event.key === 'ArrowUp' ? 1 : -1) * (event.shiftKey ? 10 : 1), min, max)
            setText(String(next))
            onCommit(next)
          }
        }}
      />
      {unit && <i>{unit}</i>}
    </label>
  )
}

// ---------------------------------------------------------------------------------------------
// The picker
// ---------------------------------------------------------------------------------------------

export interface ColorPickerProps {
  readonly value: Rgb8
  /** The colour before this edit (shown next to the new one; clicking it goes back). */
  readonly original?: Rgb8 | null
  readonly disabled?: boolean
  readonly onChange: (color: Rgb8, phase: ColorPhase) => void
  /** Optional row under the fields. */
  readonly footer?: ReactNode
}

/** Drags inside `element`, reporting positions as 0..1 fractions (x right, y down). */
function useFractionDrag(onMove: (fx: number, fy: number, phase: ColorPhase) => void, disabled?: boolean) {
  const moveRef = useRef(onMove)
  moveRef.current = onMove
  return useCallback((event: ReactPointerEvent<HTMLElement>) => {
    if (disabled || event.button !== 0) return
    event.preventDefault()
    const element = event.currentTarget
    element.focus({ preventScroll: true })
    const rect = element.getBoundingClientRect()
    const report = (clientX: number, clientY: number, phase: ColorPhase) => {
      const fx = rect.width > 0 ? clamp((clientX - rect.left) / rect.width, 0, 1) : 0
      const fy = rect.height > 0 ? clamp((clientY - rect.top) / rect.height, 0, 1) : 0
      moveRef.current(fx, fy, phase)
    }
    try {
      element.setPointerCapture(event.pointerId)
    } catch {
      // The pointer may already be gone.
    }
    report(event.clientX, event.clientY, 'live')
    const move = (moveEvent: PointerEvent) => report(moveEvent.clientX, moveEvent.clientY, 'live')
    const end = (upEvent: PointerEvent) => {
      element.removeEventListener('pointermove', move)
      element.removeEventListener('pointerup', end)
      element.removeEventListener('pointercancel', end)
      report(upEvent.clientX, upEvent.clientY, 'commit')
    }
    element.addEventListener('pointermove', move)
    element.addEventListener('pointerup', end)
    element.addEventListener('pointercancel', end)
  }, [disabled])
}

export function ColorPicker({ value, original, disabled, onChange, footer }: ColorPickerProps) {
  const [hsv, setHsv] = useState<Hsv>(() => hsvOf(value))
  const hsvRef = useRef(hsv)
  hsvRef.current = hsv
  const [hex, setHex] = useState(() => toHex(value).slice(1).toUpperCase())
  const hexEditing = useRef(false)

  // Follow colour changes made elsewhere (another panel, Undo), keeping the chosen hue for greys.
  useEffect(() => {
    setHsv((current) => (sameRgb(rgbOf(current), value) ? current : hsvOf(value, current.h)))
    if (!hexEditing.current) setHex(toHex(value).slice(1).toUpperCase())
  }, [value.r, value.g, value.b])

  const emit = (next: Hsv, phase: ColorPhase) => {
    hsvRef.current = next
    setHsv(next)
    onChange(rgbOf(next), phase)
  }
  const emitRgb = (color: Rgb8) => {
    const next = hsvOf(color, hsvRef.current.h)
    hsvRef.current = next
    setHsv(next)
    onChange(color, 'commit')
  }

  const dragField = useFractionDrag((fx, fy, phase) => emit({ h: hsvRef.current.h, s: fx, v: 1 - fy }, phase), disabled)
  const dragHue = useFractionDrag((_fx, fy, phase) => emit({ ...hsvRef.current, h: clamp((1 - fy) * 360, 0, 359.999) }, phase), disabled)

  const onFieldKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const step = event.shiftKey ? 0.1 : 0.01
    const current = hsvRef.current
    let next: Hsv | null = null
    if (event.key === 'ArrowLeft') next = { ...current, s: clamp(current.s - step, 0, 1) }
    else if (event.key === 'ArrowRight') next = { ...current, s: clamp(current.s + step, 0, 1) }
    else if (event.key === 'ArrowUp') next = { ...current, v: clamp(current.v + step, 0, 1) }
    else if (event.key === 'ArrowDown') next = { ...current, v: clamp(current.v - step, 0, 1) }
    if (next) {
      event.preventDefault()
      emit(next, 'commit')
    }
  }
  const onHueKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    const step = event.shiftKey ? 10 : 1
    const current = hsvRef.current
    if (event.key === 'ArrowUp' || event.key === 'ArrowRight') {
      event.preventDefault()
      emit({ ...current, h: clamp(current.h + step, 0, 359.999) }, 'commit')
    } else if (event.key === 'ArrowDown' || event.key === 'ArrowLeft') {
      event.preventDefault()
      emit({ ...current, h: clamp(current.h - step, 0, 359.999) }, 'commit')
    }
  }

  const commitHex = () => {
    hexEditing.current = false
    const parsed = parseHex(hex)
    if (parsed) emitRgb(parsed)
    else setHex(toHex(value).slice(1).toUpperCase())
  }

  const pure = rgbOf({ h: hsv.h, s: 1, v: 1 })
  const pick = async () => {
    const picked = await pickScreenColor()
    if (picked) emitRgb(picked)
  }
  const rounded = { h: Math.round(hsv.h), s: Math.round(hsv.s * 100), v: Math.round(hsv.v * 100) }
  return (
    <div className="ae-cp" aria-disabled={disabled || undefined}>
      <div className="ae-cp-picker">
        <div
          className="ae-cp-field-area"
          role="slider"
          tabIndex={disabled ? -1 : 0}
          aria-label="Saturation and brightness"
          aria-valuetext={`Saturation ${rounded.s}%, brightness ${rounded.v}%`}
          aria-valuenow={rounded.s}
          style={{ background: `linear-gradient(to top, #000, rgba(0, 0, 0, 0)), linear-gradient(to right, #fff, ${cssColor(pure)})` }}
          onPointerDown={dragField}
          onKeyDown={onFieldKey}
        >
          <span className="ae-cp-ring" style={{ left: `${hsv.s * 100}%`, top: `${(1 - hsv.v) * 100}%`, borderColor: hsv.v > 0.55 && hsv.s < 0.45 ? '#111' : '#fff' }} />
        </div>
        <div
          className="ae-cp-hue"
          role="slider"
          tabIndex={disabled ? -1 : 0}
          aria-label="Hue"
          aria-valuemin={0}
          aria-valuemax={360}
          aria-valuenow={rounded.h}
          onPointerDown={dragHue}
          onKeyDown={onHueKey}
        >
          <span className="ae-cp-hue-mark" style={{ top: `${(1 - hsv.h / 360) * 100}%` }} />
        </div>
      </div>
      <div className="ae-cp-row">
        <div className="ae-cp-compare" aria-label="New and current colour">
          <span className="ae-cp-new" style={{ background: cssColor(value) }} title="New colour" />
          {original && (
            <button
              type="button"
              className="ae-cp-old"
              style={{ background: cssColor(original) }}
              title="Current colour (click to go back to it)"
              aria-label="Go back to the current colour"
              disabled={disabled}
              onClick={() => emitRgb(original)}
            />
          )}
        </div>
        <label className="ae-cp-hex">
          <span>#</span>
          <input
            type="text"
            aria-label="Hex colour"
            maxLength={7}
            spellCheck={false}
            value={hex}
            disabled={disabled}
            onFocus={(event) => {
              hexEditing.current = true
              event.currentTarget.select()
            }}
            onChange={(event) => setHex(event.currentTarget.value.replace(/^#/, '').toUpperCase())}
            onBlur={commitHex}
            onKeyDown={(event) => {
              if (event.key === 'Enter') commitHex()
            }}
          />
        </label>
        {eyeDropperAvailable() && (
          <button type="button" className="ae-cp-pick" title="Pick a colour from anywhere on the screen" aria-label="Pick from screen" disabled={disabled} onClick={() => void pick()}>
            <Pipette aria-hidden="true" />
          </button>
        )}
      </div>
      <div className="ae-cp-fields">
        <ChannelInput label="H" value={rounded.h} min={0} max={360} unit="°" disabled={disabled} onCommit={(h) => emit({ ...hsvRef.current, h: h % 360 }, 'commit')} />
        <ChannelInput label="S" value={rounded.s} min={0} max={100} unit="%" disabled={disabled} onCommit={(s) => emit({ ...hsvRef.current, s: s / 100 }, 'commit')} />
        <ChannelInput label="B" value={rounded.v} min={0} max={100} unit="%" disabled={disabled} onCommit={(v) => emit({ ...hsvRef.current, v: v / 100 }, 'commit')} />
        <ChannelInput label="R" value={value.r} min={0} max={255} disabled={disabled} onCommit={(r) => emitRgb({ ...value, r })} />
        <ChannelInput label="G" value={value.g} min={0} max={255} disabled={disabled} onCommit={(g) => emitRgb({ ...value, g })} />
        <ChannelInput label="B" value={value.b} min={0} max={255} disabled={disabled} onCommit={(b) => emitRgb({ ...value, b })} />
      </div>
      {footer}
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// Popover
// ---------------------------------------------------------------------------------------------

export interface ColorPickerPopoverProps {
  readonly anchor: HTMLElement
  readonly value: Rgb8
  readonly title?: string
  readonly onChange: (color: Rgb8, phase: ColorPhase) => void
  /** Called once when the popover closes (after Escape restored the original, if it did). */
  readonly onClose: () => void
}

const POPOVER_GAP = 6

export function ColorPickerPopover({ anchor, value, title, onChange, onClose }: ColorPickerPopoverProps) {
  const original = useRef(value).current
  const latest = useRef(value)
  latest.current = value
  const rootRef = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState<{ left: number; top: number } | null>(null)
  const closedRef = useRef(false)
  const onCloseRef = useRef(onClose)
  onCloseRef.current = onClose
  const onChangeRef = useRef(onChange)
  onChangeRef.current = onChange

  const close = useCallback((restore: boolean) => {
    if (closedRef.current) return
    closedRef.current = true
    if (restore && !sameRgb(latest.current, original)) onChangeRef.current(original, 'commit')
    onCloseRef.current()
  }, [original])

  useLayoutEffect(() => {
    const root = rootRef.current
    if (!root) return
    const place = () => {
      const box = anchor.getBoundingClientRect()
      const width = root.offsetWidth
      const height = root.offsetHeight
      let left = box.left
      let top = box.bottom + POPOVER_GAP
      if (top + height > window.innerHeight - 8) top = box.top - height - POPOVER_GAP
      if (left + width > window.innerWidth - 8) left = window.innerWidth - width - 8
      setPosition({ left: Math.max(8, Math.round(left)), top: Math.max(8, Math.round(top)) })
    }
    place()
    window.addEventListener('resize', place)
    return () => window.removeEventListener('resize', place)
  }, [anchor])

  useEffect(() => {
    // A click anywhere else keeps the colour and closes (the anchor toggles itself).
    const onPointerDown = (event: PointerEvent) => {
      const target = event.target
      if (!(target instanceof Node)) return
      if (rootRef.current?.contains(target) || anchor.contains(target)) return
      close(false)
    }
    document.addEventListener('pointerdown', onPointerDown, true)
    return () => document.removeEventListener('pointerdown', onPointerDown, true)
  }, [anchor, close])

  useLayoutEffect(() => {
    rootRef.current?.querySelector<HTMLElement>('.ae-cp-field-area')?.focus({ preventScroll: true })
  }, [])

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      close(true)
      anchor.focus?.({ preventScroll: true })
    } else if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
      // Let a text field commit its value first (its own Enter handler ran already).
      event.preventDefault()
      close(false)
      anchor.focus?.({ preventScroll: true })
    }
    // The popover owns the keyboard: nothing typed here becomes an editor shortcut.
    event.stopPropagation()
  }

  const host = anchor.closest('.ae-shell') ?? document.body
  const style: CSSProperties = position ? { left: position.left, top: position.top } : { left: -9999, top: -9999 }
  return createPortal(
    <div ref={rootRef} className="ae-cp-popover" role="dialog" aria-label={title ?? 'Color picker'} style={style} onKeyDown={onKeyDown}>
      {title && <div className="ae-cp-title">{title}</div>}
      <ColorPicker value={value} original={original} onChange={(color, phase) => onChangeRef.current(color, phase)} />
    </div>,
    host,
  )
}

// ---------------------------------------------------------------------------------------------
// Swatch button
// ---------------------------------------------------------------------------------------------

export interface ColorButtonProps {
  readonly value: Rgb8
  readonly label: string
  readonly title?: string
  readonly disabled?: boolean
  readonly className?: string
  readonly onChange: (color: Rgb8, phase: ColorPhase) => void
}

export function ColorButton({ value, label, title, disabled, className, onChange }: ColorButtonProps) {
  const [anchor, setAnchor] = useState<HTMLButtonElement | null>(null)
  const ref = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    if (disabled) setAnchor(null)
  }, [disabled])
  return (
    <>
      <button
        ref={ref}
        type="button"
        className={`ae-color-button ${className ?? ''}`}
        aria-label={`${label}: ${toHex(value)}`}
        aria-haspopup="dialog"
        aria-expanded={Boolean(anchor)}
        title={title ?? `${label} (${toHex(value).toUpperCase()})`}
        disabled={disabled}
        onClick={() => setAnchor((current) => (current ? null : ref.current))}
      ><span style={{ background: cssColor(value) }} /></button>
      {anchor && <ColorPickerPopover anchor={anchor} value={value} title={label} onChange={onChange} onClose={() => setAnchor(null)} />}
    </>
  )
}
