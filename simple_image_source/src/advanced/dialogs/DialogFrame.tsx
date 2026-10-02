// src/advanced/dialogs/DialogFrame.tsx (WP6)
// The frame every Advanced dialog uses (design 5.14): a compact, movable window over the editor (drag the
// title bar; the position is remembered per dialog for the session) so a live preview on the canvas stays
// visible. It owns the keyboard while open: Enter = OK (unless a button or a multi-line field has focus),
// Escape = Cancel, Tab cycles inside the dialog, and no key reaches the editor's shortcuts. Holding Alt
// turns Cancel into Reset, as in Photoshop. Not a registered dialog itself (the registry only loads
// files named *Dialog.tsx).
import { useEffect, useId, useLayoutEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import { LoaderCircle } from 'lucide-react'

const positions = new Map<string, { readonly left: number; readonly top: number }>()

export interface DialogFrameProps {
  /** Remembers where the user moved this dialog (per kind, for the session). */
  readonly id: string
  readonly title: string
  readonly width?: number
  /** Work in progress: buttons disabled, OK shows a spinner. */
  readonly busy?: boolean
  readonly okLabel?: string
  readonly okDisabled?: boolean
  /** Left side of the footer (Auto, Preview ...). */
  readonly footerStart?: ReactNode
  readonly children: ReactNode
  readonly onOk: () => void
  readonly onCancel: () => void
  /** When given, holding Alt turns Cancel into Reset. */
  readonly onReset?: () => void
  readonly className?: string
}

const FOCUSABLE = 'button:not([disabled]), input:not([disabled]), select:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'

export function DialogFrame({ id, title, width = 360, busy, okLabel = 'OK', okDisabled, footerStart, children, onOk, onCancel, onReset, className }: DialogFrameProps) {
  const titleId = useId()
  const rootRef = useRef<HTMLDivElement>(null)
  const [position, setPosition] = useState<{ left: number; top: number } | null>(() => positions.get(id) ?? null)
  const [alt, setAlt] = useState(false)

  // First placement: centred near the top of the editor (or where the user left it), inside the layer.
  useLayoutEffect(() => {
    const root = rootRef.current
    const layer = root?.parentElement
    if (!root || !layer) return
    const bounds = layer.getBoundingClientRect()
    const fit = (left: number, top: number) => ({
      left: Math.round(Math.max(8, Math.min(left, bounds.width - root.offsetWidth - 8))),
      top: Math.round(Math.max(8, Math.min(top, bounds.height - Math.min(root.offsetHeight, bounds.height - 16) - 8))),
    })
    const remembered = positions.get(id)
    setPosition(remembered ? fit(remembered.left, remembered.top) : fit((bounds.width - root.offsetWidth) / 2, 52))
  }, [id])

  useLayoutEffect(() => {
    const root = rootRef.current
    if (!root) return
    const first = Array.from(root.querySelectorAll<HTMLElement>('input:not([type="range"]):not([type="checkbox"]):not([disabled]), select:not([disabled])'))[0]
      ?? root.querySelector<HTMLElement>('.ae-dframe-ok')
    first?.focus({ preventScroll: true })
    if (first instanceof HTMLInputElement) first.select()
  }, [])

  useEffect(() => {
    if (!onReset) return
    const up = (event: KeyboardEvent) => { if (event.key === 'Alt') setAlt(false) }
    const blur = () => setAlt(false)
    window.addEventListener('keyup', up, true)
    window.addEventListener('blur', blur)
    return () => {
      window.removeEventListener('keyup', up, true)
      window.removeEventListener('blur', blur)
    }
  }, [onReset])

  const drag = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (event.button !== 0 || !position) return
    if (event.target instanceof Element && event.target.closest('button, input, select')) return
    event.preventDefault()
    const header = event.currentTarget
    const root = rootRef.current
    const layer = root?.parentElement
    if (!root || !layer) return
    const bounds = layer.getBoundingClientRect()
    const startX = event.clientX
    const startY = event.clientY
    const start = position
    try {
      header.setPointerCapture(event.pointerId)
    } catch {
      // ignore
    }
    const move = (moveEvent: PointerEvent) => {
      const next = {
        left: Math.round(Math.max(-root.offsetWidth + 80, Math.min(start.left + moveEvent.clientX - startX, bounds.width - 80))),
        top: Math.round(Math.max(0, Math.min(start.top + moveEvent.clientY - startY, bounds.height - 40))),
      }
      positions.set(id, next)
      setPosition(next)
    }
    const end = () => {
      header.removeEventListener('pointermove', move)
      header.removeEventListener('pointerup', end)
      header.removeEventListener('pointercancel', end)
    }
    header.addEventListener('pointermove', move)
    header.addEventListener('pointerup', end)
    header.addEventListener('pointercancel', end)
  }

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Alt' && onReset) {
      setAlt(true)
      event.preventDefault()
    } else if (event.key === 'Escape') {
      event.preventDefault()
      if (!busy) onCancel()
    } else if (event.key === 'Enter' && !event.altKey && !event.ctrlKey && !event.metaKey) {
      const target = event.target
      if (!(target instanceof HTMLButtonElement) && !(target instanceof HTMLTextAreaElement) && !(target instanceof Element && target.closest('[role="slider"], svg'))) {
        event.preventDefault()
        if (!busy && !okDisabled) onOk()
      }
    } else if (event.key === 'Tab') {
      const root = rootRef.current
      const focusables = root ? Array.from(root.querySelectorAll<HTMLElement>(FOCUSABLE)).filter((element) => element.offsetParent !== null) : []
      if (focusables.length) {
        const first = focusables[0]
        const last = focusables[focusables.length - 1]
        if (event.shiftKey && document.activeElement === first) {
          event.preventDefault()
          last.focus()
        } else if (!event.shiftKey && document.activeElement === last) {
          event.preventDefault()
          first.focus()
        }
      }
    }
    // The dialog owns the keyboard: nothing typed here becomes an editor shortcut.
    event.stopPropagation()
  }

  const resetMode = Boolean(alt && onReset)
  return (
    <div
      ref={rootRef}
      className={`ae-dframe ${className ?? ''}`}
      role="dialog"
      aria-modal="true"
      aria-labelledby={titleId}
      aria-busy={busy || undefined}
      style={{ width: `min(${width}px, calc(100% - 16px))`, left: position?.left ?? 0, top: position?.top ?? 0, visibility: position ? undefined : 'hidden' }}
      onKeyDown={onKeyDown}
    >
      <div className="ae-dframe-head" onPointerDown={drag} title="Drag to move">
        <h2 id={titleId}>{title}</h2>
      </div>
      <div className="ae-dframe-body">{children}</div>
      <div className="ae-dframe-foot">
        <div className="ae-dframe-start">{footerStart}</div>
        <button
          type="button"
          className="ae-dframe-cancel"
          disabled={busy}
          title={onReset ? 'Hold Alt to reset the settings instead' : undefined}
          onClick={() => {
            if (resetMode) onReset?.()
            else onCancel()
          }}
        >{resetMode ? 'Reset' : 'Cancel'}</button>
        <button type="button" className="ae-dframe-ok" disabled={busy || okDisabled} onClick={onOk}>
          {busy ? <><LoaderCircle className="ae-spin" aria-hidden="true" /><span>Working…</span></> : okLabel}
        </button>
      </div>
    </div>
  )
}

/** Unit conversions shared by the size dialogs. */
export type LengthUnit = 'px' | 'percent' | 'in' | 'cm'

export const LENGTH_UNITS: readonly (readonly [LengthUnit, string])[] = [['px', 'Pixels'], ['percent', 'Percent'], ['in', 'Inches'], ['cm', 'Centimeters']]

/** A pixel length shown in `unit` (percent of `reference` px; inches / cm at `ppi`). */
export function toUnit(pixels: number, unit: LengthUnit, reference: number, ppi: number): number {
  if (unit === 'percent') return reference > 0 ? (pixels / reference) * 100 : 100
  if (unit === 'in') return pixels / ppi
  if (unit === 'cm') return (pixels / ppi) * 2.54
  return pixels
}

/** Pixels for a value in `unit` (rounded to whole pixels). */
export function fromUnit(value: number, unit: LengthUnit, reference: number, ppi: number): number {
  if (unit === 'percent') return Math.round((reference * value) / 100)
  if (unit === 'in') return Math.round(value * ppi)
  if (unit === 'cm') return Math.round((value / 2.54) * ppi)
  return Math.round(value)
}

export function unitDigits(unit: LengthUnit): number {
  return unit === 'px' ? 0 : unit === 'percent' ? 1 : 2
}
