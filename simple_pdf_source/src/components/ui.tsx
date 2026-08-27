import type { ButtonHTMLAttributes, ReactNode } from 'react'
import type { LucideIcon } from 'lucide-react'
import { LoaderCircle, X } from 'lucide-react'
import { cx } from '../lib/utils'

interface IconButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  icon: LucideIcon
  label: string
  active?: boolean
  compact?: boolean
}

export function IconButton({ icon: Icon, label, active, compact, className, ...props }: IconButtonProps) {
  return (
    <button
      type="button"
      className={cx('icon-button', active && 'is-active', compact && 'is-compact', className)}
      aria-label={label}
      title={label}
      {...props}
    >
      <Icon size={16} strokeWidth={1.8} aria-hidden="true" />
    </button>
  )
}

interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  icon?: LucideIcon
  variant?: 'primary' | 'secondary' | 'ghost' | 'danger'
  children: ReactNode
}

export function Button({ icon: Icon, variant = 'secondary', className, children, ...props }: ButtonProps) {
  return (
    <button type="button" className={cx('button', `button-${variant}`, className)} {...props}>
      {Icon && <Icon size={16} strokeWidth={1.8} aria-hidden="true" />}
      <span>{children}</span>
    </button>
  )
}

export function Separator({ vertical = true }: { vertical?: boolean }) {
  return <span className={vertical ? 'separator vertical' : 'separator'} aria-hidden="true" />
}

export function BusyOverlay({ label }: { label: string }) {
  return (
    <div className="busy-overlay" role="status" aria-live="polite">
      <div className="busy-card">
        <LoaderCircle className="spin" size={18} />
        <span>{label}</span>
      </div>
    </div>
  )
}

export function Toast({ message, action, onAction, onClose }: { message: string; action?: string; onAction?: () => void; onClose: () => void }) {
  return (
    <div className="toast" role="status">
      <span>{message}</span>
      {action && onAction && <button type="button" onClick={onAction}>{action}</button>}
      <button type="button" className="toast-close" aria-label="Dismiss" onClick={onClose}>
        <X size={14} />
      </button>
    </div>
  )
}

export function EmptyState({ icon: Icon, title, detail }: { icon: LucideIcon; title: string; detail: string }) {
  return (
    <div className="empty-state">
      <span className="empty-icon"><Icon size={18} /></span>
      <strong>{title}</strong>
      <p>{detail}</p>
    </div>
  )
}
