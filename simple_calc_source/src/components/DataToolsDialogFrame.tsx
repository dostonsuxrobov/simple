import { useEffect, useRef } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, ReactNode, RefObject } from 'react'
import { X } from 'lucide-react'
import './data-tools.css'

export const DT_FOCUSABLE = 'button:not([disabled]), select:not([disabled]), input:not([disabled]), textarea:not([disabled]), [tabindex]:not([tabindex="-1"])'

/** Tab / Shift+Tab stay inside `container`. Returns true when the event was handled. */
export function trapTab(event: ReactKeyboardEvent<HTMLElement>, container: HTMLElement | null): boolean {
  if (event.key !== 'Tab' || !container) return false
  const controls = Array.from(container.querySelectorAll<HTMLElement>(DT_FOCUSABLE)).filter((element) => element.getClientRects().length > 0)
  if (!controls.length) return false
  const first = controls[0]
  const last = controls[controls.length - 1]
  const active = document.activeElement
  if (event.shiftKey && (active === first || !container.contains(active))) {
    event.preventDefault()
    last.focus()
    return true
  }
  if (!event.shiftKey && (active === last || !container.contains(active))) {
    event.preventDefault()
    first.focus()
    return true
  }
  return false
}

/** Focus the first control (or `initial`) on mount; restore the previously focused element on unmount. */
export function useInitialFocus(container: RefObject<HTMLElement>, initial?: RefObject<HTMLElement>) {
  useEffect(() => {
    const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null
    const frame = window.requestAnimationFrame(() => {
      const target = initial?.current || container.current?.querySelector<HTMLElement>('[data-autofocus]') || container.current?.querySelector<HTMLElement>(DT_FOCUSABLE)
      target?.focus()
      if (target instanceof HTMLInputElement && target.type === 'text') target.select()
    })
    return () => {
      window.cancelAnimationFrame(frame)
      if (previous?.isConnected) previous.focus()
    }
  }, [container, initial])
}

function enterConfirms(event: ReactKeyboardEvent<HTMLElement>) {
  if (event.key !== 'Enter' || event.shiftKey || event.altKey || event.nativeEvent.isComposing) return false
  const target = event.target as HTMLElement
  if (target instanceof HTMLTextAreaElement && !(event.ctrlKey || event.metaKey)) return false
  if (target instanceof HTMLButtonElement || target instanceof HTMLAnchorElement) return false
  if (target instanceof HTMLInputElement && (target.type === 'checkbox' || target.type === 'radio') && !(event.ctrlKey || event.metaKey)) return false
  return true
}

export interface DataToolsDialogProps {
  title: string
  subtitle?: ReactNode
  icon: ReactNode
  width?: number
  className?: string
  children: ReactNode
  footer: ReactNode
  onClose: () => void
  /** Enter (outside buttons / textareas) and Ctrl+Enter anywhere. */
  onConfirm?: () => void
  initialFocus?: RefObject<HTMLElement>
  /** Close when the backdrop is pressed (off by default: dialogs hold unsaved settings). */
  closeOnBackdrop?: boolean
  busy?: boolean
}

/** Shared modal chrome: header, scrollable body, footer, focus trap, Escape / Enter handling. */
export function DataToolsDialog({ title, subtitle, icon, width = 520, className = '', children, footer, onClose, onConfirm, initialFocus, closeOnBackdrop = false, busy = false }: DataToolsDialogProps) {
  const dialogRef = useRef<HTMLElement>(null)
  const titleId = useRef(`dt-title-${Math.random().toString(36).slice(2, 9)}`).current
  useInitialFocus(dialogRef, initialFocus)

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      if (!busy) onClose()
      return
    }
    if (onConfirm && enterConfirms(event)) {
      event.preventDefault()
      event.stopPropagation()
      if (!busy) onConfirm()
      return
    }
    if (trapTab(event, dialogRef.current)) event.stopPropagation()
    // Keep app-level shortcuts (Ctrl+S, Ctrl+Z…) from acting behind the dialog.
    if ((event.ctrlKey || event.metaKey) && ['s', 'o', 'n', 'p', 'z', 'y'].includes(event.key.toLocaleLowerCase())) {
      const editable = event.target instanceof HTMLInputElement || event.target instanceof HTMLTextAreaElement
      if (!editable || ['s', 'o', 'n', 'p'].includes(event.key.toLocaleLowerCase())) event.preventDefault()
      event.stopPropagation()
    }
  }

  return (
    <div
      className="dt-overlay"
      onMouseDown={(event) => {
        if (event.target !== event.currentTarget) return
        if (closeOnBackdrop && !busy) onClose()
        else {
          event.preventDefault()
          dialogRef.current?.classList.remove('dt-nudge')
          void dialogRef.current?.offsetWidth
          dialogRef.current?.classList.add('dt-nudge')
        }
      }}
    >
      <section
        ref={dialogRef}
        className={`dt-dialog ${className}`.trim()}
        style={{ width: `min(${width}px, calc(100vw - 48px))` }}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        onKeyDown={handleKeyDown}
      >
        <header className="dt-header">
          <span className="dt-header-icon" aria-hidden="true">{icon}</span>
          <div className="dt-title">
            <h2 id={titleId}>{title}</h2>
            {subtitle ? <p>{subtitle}</p> : null}
          </div>
          <button type="button" className="dt-close" aria-label={`Close ${title}`} disabled={busy} onClick={onClose}><X size={15} /></button>
        </header>
        <div className="dt-body">{children}</div>
        <footer className="dt-footer">{footer}</footer>
      </section>
    </div>
  )
}

export function Swatch({ color, label }: { color: string | null; label?: string }) {
  return (
    <span
      className={`dt-swatch${color ? '' : ' is-none'}`}
      style={color ? { background: color } : undefined}
      aria-label={label}
      role={label ? 'img' : undefined}
    />
  )
}

export function colorLabel(color: string | null, kind: 'fill' | 'font'): string {
  if (!color) return kind === 'fill' ? 'No fill' : 'Automatic'
  return color.toUpperCase()
}
