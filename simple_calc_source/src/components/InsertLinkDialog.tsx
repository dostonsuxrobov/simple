import { useEffect, useRef, useState } from 'react'
import { Globe, Table2, Trash2 } from 'lucide-react'
import { classifyLinkTarget, internalLinkTarget, isCellReference, isWebLink, normalizeWebAddress, splitLinkLocation } from '../lib/hyperlinks'
import './app-polish.css'

export interface InsertLinkResult {
  /** "https://…", "mailto:…" or a place in this workbook ("#Sheet2!A1", "#MyName"). */
  target: string
  /** Text the cell shows; null keeps the cell's current content (a formula cell). */
  text: string | null
  tooltip?: string
}

export interface InsertLinkDialogProps {
  /** The cell's current link, when editing one. */
  initialTarget: string | null
  initialText: string
  /** The cell holds a formula: its text stays as it is. */
  textLocked?: boolean
  /** Visible sheets, in tab order. */
  sheets: string[]
  activeSheet: string
  /** Defined names and table names a link can go to. */
  names: string[]
  /** Proposed cell for a new in-workbook link ("A1"). */
  defaultReference: string
  workbookName?: string
  onApply: (result: InsertLinkResult) => void
  onRemove?: () => void
  onClose: () => void
}

type Mode = 'web' | 'place'

/**
 * Ctrl+K: link a cell to a web or email address, or to a place in this workbook (a sheet and
 * cell, or a defined name), like Excel's Insert Hyperlink and Sheets' link card.
 */
export function InsertLinkDialog({ initialTarget, initialText, textLocked, sheets, activeSheet, names, defaultReference, workbookName, onApply, onRemove, onClose }: InsertLinkDialogProps) {
  const parsed = initialTarget ? classifyLinkTarget(initialTarget, workbookName) : null
  const location = parsed?.kind === 'internal' ? splitLinkLocation(parsed.location) : null
  const knownSheet = (name?: string) => sheets.find((sheet) => sheet.toLocaleLowerCase() === String(name || '').toLocaleLowerCase())
  const knownName = (name?: string) => names.find((item) => item.toLocaleLowerCase() === String(name || '').toLocaleLowerCase())
  const initialPlace = location
    ? (location.sheet && knownSheet(location.sheet) ? `sheet:${knownSheet(location.sheet)}` : !location.sheet && knownName(location.reference) ? `name:${knownName(location.reference)}` : `sheet:${activeSheet}`)
    : `sheet:${activeSheet}`
  const [mode, setMode] = useState<Mode>(parsed?.kind === 'internal' ? 'place' : 'web')
  const [address, setAddress] = useState(parsed?.kind === 'external' ? parsed.url.replace(/^mailto:/i, '') : parsed?.kind === 'file' ? initialTarget || '' : '')
  const [place, setPlace] = useState(initialPlace)
  const [reference, setReference] = useState(location && isCellReference(location.reference) ? location.reference : defaultReference)
  const [text, setText] = useState(initialText)
  const [error, setError] = useState('')
  const addressRef = useRef<HTMLInputElement>(null)
  const referenceRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    const field = mode === 'web' ? addressRef.current : referenceRef.current
    field?.focus()
    field?.select()
  }, [mode])

  const placeIsName = place.startsWith('name:')
  const placeValue = place.slice(place.indexOf(':') + 1)

  const submit = () => {
    if (mode === 'web') {
      const url = normalizeWebAddress(address)
      if (!url || !isWebLink(url)) {
        setError('Enter a web address (https://…) or an email address.')
        addressRef.current?.focus()
        return
      }
      onApply({ target: url, text: textLocked ? null : text.trim() ? text : address.trim(), tooltip: url })
      return
    }
    if (placeIsName) {
      onApply({ target: internalLinkTarget(null, placeValue), text: textLocked ? null : text.trim() ? text : placeValue })
      return
    }
    const cell = reference.trim().replace(/\$/g, '').toUpperCase()
    if (!isCellReference(cell)) {
      setError('Enter a cell reference such as A1 or B2:D10.')
      referenceRef.current?.focus()
      return
    }
    const target = internalLinkTarget(placeValue, cell)
    onApply({ target, text: textLocked ? null : text.trim() ? text : `${placeValue}!${cell}` })
  }

  return (
    <div className="prompt-overlay" onPointerDown={(event) => { if (event.target === event.currentTarget) onClose() }}>
      <form
        className="prompt-card link-dialog"
        role="dialog"
        aria-modal="true"
        aria-label={initialTarget ? 'Edit link' : 'Insert link'}
        onSubmit={(event) => { event.preventDefault(); submit() }}
        onKeyDown={(event) => { if (event.key === 'Escape') { event.preventDefault(); event.stopPropagation(); onClose() } }}
      >
        <h2>{initialTarget ? 'Edit link' : 'Insert link'}</h2>
        <label>
          <span>Text to display</span>
          <input value={textLocked ? '(the formula’s result)' : text} disabled={textLocked} aria-label="Text to display" onChange={(event) => setText(event.target.value)} />
        </label>
        <div className="link-mode" role="radiogroup" aria-label="Link to">
          <button type="button" role="radio" aria-checked={mode === 'web'} onClick={() => { setMode('web'); setError('') }}><Globe size={13} aria-hidden="true" />Web address</button>
          <button type="button" role="radio" aria-checked={mode === 'place'} onClick={() => { setMode('place'); setError('') }}><Table2 size={13} aria-hidden="true" />Place in this workbook</button>
        </div>
        {mode === 'web' ? (
          <label>
            <span>Address</span>
            <input
              ref={addressRef}
              aria-label="Link address"
              placeholder="https://example.com or name@example.com"
              spellCheck={false}
              value={address}
              onChange={(event) => { setAddress(event.target.value); setError('') }}
            />
          </label>
        ) : (
          <div className="link-place">
            <label>
              <span>Sheet or name</span>
              <select aria-label="Place in this workbook" value={place} onChange={(event) => { setPlace(event.target.value); setError('') }}>
                <optgroup label="Sheets">
                  {sheets.map((sheet) => <option key={`sheet-${sheet}`} value={`sheet:${sheet}`}>{sheet}</option>)}
                </optgroup>
                {names.length > 0 && (
                  <optgroup label="Defined names">
                    {names.map((name) => <option key={`name-${name}`} value={`name:${name}`}>{name}</option>)}
                  </optgroup>
                )}
              </select>
            </label>
            <label>
              <span>Cell reference</span>
              <input
                ref={referenceRef}
                aria-label="Cell reference"
                spellCheck={false}
                disabled={placeIsName}
                value={placeIsName ? '' : reference}
                placeholder={placeIsName ? 'Whole name' : 'A1'}
                onChange={(event) => { setReference(event.target.value); setError('') }}
              />
            </label>
          </div>
        )}
        {error && <p className="link-error" role="alert">{error}</p>}
        <div className="prompt-actions">
          {onRemove && <button type="button" className="secondary-action link-remove" onClick={onRemove}><Trash2 size={13} aria-hidden="true" />Remove link</button>}
          <button type="button" className="secondary-action" onClick={onClose}>Cancel</button>
          <button type="submit" className="primary-action">OK</button>
        </div>
      </form>
    </div>
  )
}
