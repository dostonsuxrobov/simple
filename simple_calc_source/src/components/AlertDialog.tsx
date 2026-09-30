import { useEffect, useRef } from 'react'
import { AlertTriangle, CircleX, Info } from 'lucide-react'

export interface AlertButton {
  id: string
  label: string
  primary?: boolean
}

export interface AlertRequest {
  title: string
  message: string
  tone: 'stop' | 'warning' | 'information'
  buttons: AlertButton[]
}

/** A small modal message with Excel-style choices (Retry/Cancel, Yes/No/Cancel, OK/Cancel). */
export function AlertDialog({ request, onClose }: { request: AlertRequest; onClose: (buttonId: string) => void }) {
  const primaryRef = useRef<HTMLButtonElement>(null)
  useEffect(() => {
    primaryRef.current?.focus()
    const handle = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return
      event.preventDefault()
      event.stopPropagation()
      onClose(request.buttons.some((button) => button.id === 'cancel') ? 'cancel' : request.buttons[request.buttons.length - 1].id)
    }
    window.addEventListener('keydown', handle, true)
    return () => window.removeEventListener('keydown', handle, true)
  }, [onClose, request.buttons])
  const Icon = request.tone === 'stop' ? CircleX : request.tone === 'warning' ? AlertTriangle : Info
  return (
    <div className="prompt-overlay">
      <div className={`alert-card tone-${request.tone}`} role="alertdialog" aria-modal="true" aria-label={request.title}>
        <div className="alert-body">
          <Icon size={22} className="alert-icon" aria-hidden="true" />
          <div>
            <h2>{request.title}</h2>
            <p>{request.message}</p>
          </div>
        </div>
        <div className="prompt-actions">
          {request.buttons.map((button) => (
            <button
              key={button.id}
              ref={button.primary ? primaryRef : undefined}
              type="button"
              className={button.primary ? 'primary-action' : 'secondary-action'}
              onClick={() => onClose(button.id)}
            >{button.label}</button>
          ))}
        </div>
      </div>
    </div>
  )
}
