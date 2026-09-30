import { useEffect, useRef, useState, type KeyboardEvent as ReactKeyboardEvent } from 'react'
import { ClipboardPaste, Link2, X } from 'lucide-react'
import {
  DEFAULT_PASTE_SPECIAL_OPTIONS,
  PASTE_SPECIAL_MENU,
  pasteTypeSupportsLink,
  pasteTypeSupportsOperation,
  resolvePasteSpecialOptions,
  type PasteSpecialOperation,
  type PasteSpecialOptions,
  type PasteSpecialPasteType,
} from '../lib/paste-special'
import './paste-special.css'

export { PASTE_SPECIAL_MENU }

interface PasteSpecialDialogProps {
  initialOptions?: Partial<PasteSpecialOptions>
  /** Short description of what is on the clipboard, e.g. "B2:D9 on Sheet1". */
  sourceLabel?: string
  /** Paste Link needs cells copied inside this workbook. */
  canPasteLink?: boolean
  /** Paste types that make no sense for the current clipboard (e.g. validation from another app). */
  disabledPasteTypes?: readonly PasteSpecialPasteType[]
  onCancel: () => void
  onPaste: (options: PasteSpecialOptions) => void
}

interface Choice<T extends string> { value: T; label: string; key: string }

// Two columns, filled top-to-bottom like Excel's dialog; `key` is the underlined access key.
const PASTE_CHOICES: Array<Choice<PasteSpecialPasteType>> = [
  { value: 'all', label: 'All', key: 'a' },
  { value: 'formulas', label: 'Formulas', key: 'f' },
  { value: 'values', label: 'Values', key: 'v' },
  { value: 'formats', label: 'Formats', key: 't' },
  { value: 'comments', label: 'Comments and notes', key: 'c' },
  { value: 'validation', label: 'Validation', key: 'n' },
  { value: 'allExceptBorders', label: 'All except borders', key: 'x' },
  { value: 'columnWidths', label: 'Column widths', key: 'w' },
  { value: 'formulasAndNumberFormats', label: 'Formulas and number formats', key: 'r' },
  { value: 'valuesAndNumberFormats', label: 'Values and number formats', key: 'u' },
]

const OPERATION_CHOICES: Array<Choice<PasteSpecialOperation>> = [
  { value: 'none', label: 'None', key: 'o' },
  { value: 'add', label: 'Add', key: 'd' },
  { value: 'subtract', label: 'Subtract', key: 's' },
  { value: 'multiply', label: 'Multiply', key: 'm' },
  { value: 'divide', label: 'Divide', key: 'i' },
]

const SKIP_BLANKS_KEY = 'b'
const TRANSPOSE_KEY = 'e'
const PASTE_LINK_KEY = 'l'

function AccessLabel({ label, accessKey }: { label: string; accessKey: string }) {
  const index = label.toLowerCase().indexOf(accessKey)
  if (index === -1) return <>{label}</>
  return <>{label.slice(0, index)}<span className="ps-key">{label[index]}</span>{label.slice(index + 1)}</>
}

export function PasteSpecialDialog({
  initialOptions,
  sourceLabel,
  canPasteLink = true,
  disabledPasteTypes = [],
  onCancel,
  onPaste,
}: PasteSpecialDialogProps) {
  const [options, setOptions] = useState<PasteSpecialOptions>(() => {
    const resolved = resolvePasteSpecialOptions({ ...initialOptions, pasteLink: false })
    return disabledPasteTypes.includes(resolved.paste) ? { ...resolved, paste: DEFAULT_PASTE_SPECIAL_OPTIONS.paste } : resolved
  })
  const cardRef = useRef<HTMLElement>(null)
  const checkedPasteRef = useRef<HTMLInputElement>(null)
  const operationEnabled = pasteTypeSupportsOperation(options.paste)
  const linkEnabled = canPasteLink && pasteTypeSupportsLink(options)

  useEffect(() => {
    checkedPasteRef.current?.focus()
  }, [])

  const update = (patch: Partial<PasteSpecialOptions>) => setOptions((current) => {
    const next = { ...current, ...patch }
    if (!pasteTypeSupportsOperation(next.paste)) next.operation = 'none'
    return next
  })

  const submit = () => onPaste({ ...options, pasteLink: false })
  const pasteLink = () => { if (linkEnabled) onPaste({ ...options, operation: 'none', skipBlanks: false, pasteLink: true }) }

  const handleKeyDown = (event: ReactKeyboardEvent<HTMLElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      onCancel()
      return
    }
    if (event.key === 'Tab') {
      const focusable = [...event.currentTarget.querySelectorAll<HTMLElement>('button:not([disabled]), input:not([disabled])')]
        .filter((element) => element.getClientRects().length > 0 && (element.getAttribute('type') !== 'radio' || (element as HTMLInputElement).checked))
      if (!focusable.length) return
      const first = focusable[0]
      const last = focusable[focusable.length - 1]
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last.focus() }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first.focus() }
      return
    }
    if (event.key === 'Enter') {
      const target = event.target as HTMLElement
      if (target.tagName === 'BUTTON') return
      event.preventDefault()
      submit()
      return
    }
    if (event.ctrlKey || event.metaKey || event.key.length !== 1) return
    const key = event.key.toLowerCase()
    const paste = PASTE_CHOICES.find((choice) => choice.key === key)
    const operation = OPERATION_CHOICES.find((choice) => choice.key === key)
    const focusValue = (name: string, value: string) => {
      requestAnimationFrame(() => cardRef.current?.querySelector<HTMLInputElement>(`input[name="${name}"][value="${value}"]`)?.focus())
    }
    if (paste && !disabledPasteTypes.includes(paste.value)) {
      event.preventDefault()
      update({ paste: paste.value })
      focusValue('ps-paste', paste.value)
    } else if (operation && operationEnabled) {
      event.preventDefault()
      update({ operation: operation.value })
      focusValue('ps-operation', operation.value)
    } else if (key === SKIP_BLANKS_KEY) {
      event.preventDefault()
      update({ skipBlanks: !options.skipBlanks })
    } else if (key === TRANSPOSE_KEY) {
      event.preventDefault()
      update({ transpose: !options.transpose })
    } else if (key === PASTE_LINK_KEY && linkEnabled) {
      event.preventDefault()
      pasteLink()
    }
  }

  return (
    <div className="ps-overlay" onMouseDown={(event) => { if (event.target === event.currentTarget) onCancel() }}>
      <section
        ref={cardRef}
        className="ps-card"
        role="dialog"
        aria-modal="true"
        aria-labelledby="ps-title"
        aria-describedby={sourceLabel ? 'ps-source' : undefined}
        onKeyDown={handleKeyDown}
      >
        <header className="ps-header">
          <span className="ps-icon"><ClipboardPaste size={16} aria-hidden="true" /></span>
          <div>
            <h2 id="ps-title">Paste Special</h2>
            {sourceLabel && <p id="ps-source">{sourceLabel}</p>}
          </div>
          <button type="button" className="ps-close" aria-label="Close paste special" onClick={onCancel}><X size={15} /></button>
        </header>

        <form className="ps-body" onSubmit={(event) => { event.preventDefault(); submit() }}>
          <fieldset className="ps-group">
            <legend>Paste</legend>
            <div className="ps-grid ps-grid-paste" role="radiogroup" aria-label="Paste">
              {PASTE_CHOICES.map((choice) => {
                const disabled = disabledPasteTypes.includes(choice.value)
                const checked = options.paste === choice.value
                return (
                  <label key={choice.value} className={`ps-option${checked ? ' is-checked' : ''}${disabled ? ' is-disabled' : ''}`}>
                    <input
                      ref={checked ? checkedPasteRef : undefined}
                      type="radio"
                      name="ps-paste"
                      value={choice.value}
                      checked={checked}
                      disabled={disabled}
                      aria-keyshortcuts={choice.key.toUpperCase()}
                      onChange={() => update({ paste: choice.value })}
                    />
                    <span><AccessLabel label={choice.label} accessKey={choice.key} /></span>
                  </label>
                )
              })}
            </div>
          </fieldset>

          <fieldset className="ps-group" disabled={!operationEnabled}>
            <legend>Operation</legend>
            <div className="ps-grid ps-grid-operation" role="radiogroup" aria-label="Operation">
              {OPERATION_CHOICES.map((choice) => {
                const checked = options.operation === choice.value
                return (
                  <label key={choice.value} className={`ps-option${checked ? ' is-checked' : ''}${operationEnabled ? '' : ' is-disabled'}`}>
                    <input
                      type="radio"
                      name="ps-operation"
                      value={choice.value}
                      checked={checked}
                      aria-keyshortcuts={choice.key.toUpperCase()}
                      onChange={() => update({ operation: choice.value })}
                    />
                    <span><AccessLabel label={choice.label} accessKey={choice.key} /></span>
                  </label>
                )
              })}
            </div>
          </fieldset>

          <div className="ps-toggles">
            <label className="ps-option">
              <input type="checkbox" checked={options.skipBlanks} aria-keyshortcuts="B" onChange={(event) => update({ skipBlanks: event.target.checked })} />
              <span><AccessLabel label="Skip blanks" accessKey={SKIP_BLANKS_KEY} /></span>
            </label>
            <label className="ps-option">
              <input type="checkbox" checked={options.transpose} aria-keyshortcuts="E" onChange={(event) => update({ transpose: event.target.checked })} />
              <span><AccessLabel label="Transpose" accessKey={TRANSPOSE_KEY} /></span>
            </label>
          </div>

          <footer className="ps-actions">
            <button type="button" className="secondary-action ps-link" disabled={!linkEnabled} aria-keyshortcuts="L" onClick={pasteLink}>
              <Link2 size={13} aria-hidden="true" /><span><AccessLabel label="Paste Link" accessKey={PASTE_LINK_KEY} /></span>
            </button>
            <span className="ps-spacer" />
            <button type="button" className="secondary-action" onClick={onCancel}>Cancel</button>
            <button type="submit" className="primary-action">OK</button>
          </footer>
        </form>
      </section>
    </div>
  )
}
