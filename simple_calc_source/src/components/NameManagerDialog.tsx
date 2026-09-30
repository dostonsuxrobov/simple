import { useEffect, useMemo, useRef, useState } from 'react'
import { Pencil, Plus, Tag, Trash2, X } from 'lucide-react'
import type { DefinedName } from '../spreadsheet-types'
import './name-manager.css'

export interface NameDraft {
  name: string
  refersTo: string
  /** Sheet index for a sheet-scoped name, or null for workbook scope. */
  scope: number | null
  comment?: string
}

interface NameManagerDialogProps {
  names: DefinedName[]
  sheetNames: string[]
  /** Reference text for the current selection, e.g. Sheet1!$A$1:$B$5 */
  selectionReference: string
  evaluate: (refersTo: string) => string
  onSave: (names: NameDraft[]) => void
  onClose: () => void
}

const RESERVED = /^(?:[A-Za-z]{1,3}\d+|R\d*C\d*|TRUE|FALSE)$/i

export function validateDefinedName(name: string, existing: NameDraft[], editingIndex: number | null, scope: number | null): string | null {
  const trimmed = name.trim()
  if (!trimmed) return 'Enter a name.'
  if (trimmed.length > 255) return 'Names can be at most 255 characters.'
  if (!/^[A-Za-z_\\À-￿][A-Za-z0-9_.?\\À-￿]*$/.test(trimmed)) return 'Names start with a letter, underscore, or backslash and cannot contain spaces.'
  if (RESERVED.test(trimmed)) return 'That name looks like a cell reference.'
  const clash = existing.findIndex((item, index) => index !== editingIndex && item.scope === scope && item.name.toLocaleLowerCase() === trimmed.toLocaleLowerCase())
  return clash >= 0 ? 'A name with that scope already exists.' : null
}

export function NameManagerDialog({ names, sheetNames, selectionReference, evaluate, onSave, onClose }: NameManagerDialogProps) {
  const initial = useMemo<NameDraft[]>(() => names
    .filter((item) => !/^_xlnm\./i.test(item.name))
    .map((item) => ({
      name: item.name,
      refersTo: `=${(item.ranges?.[0] || item.ref || '').replace(/^=/, '')}`,
      scope: item.localSheetIndex ?? item.localSheetId ?? null,
      comment: item.comment,
    })), [names])
  const [items, setItems] = useState<NameDraft[]>(initial)
  const [editing, setEditing] = useState<{ index: number | null; draft: NameDraft } | null>(null)
  const [filter, setFilter] = useState('')
  const [error, setError] = useState('')
  const rootRef = useRef<HTMLDivElement>(null)
  const nameFieldRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    const handle = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return
      event.preventDefault()
      if (editing) setEditing(null)
      else onClose()
    }
    window.addEventListener('keydown', handle)
    return () => window.removeEventListener('keydown', handle)
  }, [editing, onClose])

  useEffect(() => {
    if (editing) window.requestAnimationFrame(() => nameFieldRef.current?.focus())
  }, [editing?.index])

  const visible = items
    .map((item, index) => ({ item, index }))
    .filter(({ item }) => !filter || item.name.toLocaleLowerCase().includes(filter.toLocaleLowerCase()))

  const commitEdit = () => {
    if (!editing) return
    const message = validateDefinedName(editing.draft.name, items, editing.index, editing.draft.scope)
    if (message) { setError(message); return }
    if (!editing.draft.refersTo.replace(/^=/, '').trim()) { setError('Enter what the name refers to.'); return }
    const draft = { ...editing.draft, name: editing.draft.name.trim(), refersTo: `=${editing.draft.refersTo.trim().replace(/^=/, '')}` }
    setItems((current) => editing.index === null ? [...current, draft] : current.map((item, index) => index === editing.index ? draft : item))
    setEditing(null)
    setError('')
  }

  return (
    <div className="prompt-overlay" onPointerDown={(event) => { if (event.target === event.currentTarget) onClose() }}>
      <div ref={rootRef} className="nm-dialog" role="dialog" aria-modal="true" aria-label="Name Manager">
        <header className="nm-header">
          <h2><Tag size={15} /> Name Manager</h2>
          <button type="button" className="nm-icon-button" aria-label="Close" onClick={onClose}><X size={15} /></button>
        </header>
        <div className="nm-toolbar">
          <button type="button" className="secondary-action" onClick={() => { setEditing({ index: null, draft: { name: '', refersTo: selectionReference, scope: null } }); setError('') }}><Plus size={14} /> New…</button>
          <input className="nm-filter" placeholder="Filter names" value={filter} onChange={(event) => setFilter(event.target.value)} aria-label="Filter names" />
        </div>
        <div className="nm-table" role="table" aria-label="Defined names">
          <div className="nm-row nm-head" role="row">
            <span role="columnheader">Name</span><span role="columnheader">Value</span><span role="columnheader">Refers to</span><span role="columnheader">Scope</span><span />
          </div>
          {visible.map(({ item, index }) => (
            <div className="nm-row" role="row" key={`${item.name}-${index}`} onDoubleClick={() => { setEditing({ index, draft: { ...item } }); setError('') }}>
              <span className="nm-name" role="cell">{item.name}</span>
              <span className="nm-value" role="cell" title={evaluate(item.refersTo)}>{evaluate(item.refersTo)}</span>
              <span className="nm-refers" role="cell" title={item.refersTo}>{item.refersTo}</span>
              <span role="cell">{item.scope === null ? 'Workbook' : sheetNames[item.scope] ?? `Sheet ${item.scope + 1}`}</span>
              <span className="nm-actions" role="cell">
                <button type="button" className="nm-icon-button" aria-label={`Edit ${item.name}`} onClick={() => { setEditing({ index, draft: { ...item } }); setError('') }}><Pencil size={13} /></button>
                <button type="button" className="nm-icon-button" aria-label={`Delete ${item.name}`} onClick={() => setItems((current) => current.filter((_, position) => position !== index))}><Trash2 size={13} /></button>
              </span>
            </div>
          ))}
          {!visible.length && <p className="nm-empty">{items.length ? 'No names match the filter.' : 'No defined names yet. Select cells and choose New… (or type a name in the Name Box).'}</p>}
        </div>
        {editing && (
          <form className="nm-editor" onSubmit={(event) => { event.preventDefault(); commitEdit() }}>
            <label><span>Name</span><input ref={nameFieldRef} value={editing.draft.name} onChange={(event) => setEditing({ ...editing, draft: { ...editing.draft, name: event.target.value } })} /></label>
            <label><span>Scope</span>
              <select value={editing.draft.scope === null ? '' : String(editing.draft.scope)} onChange={(event) => setEditing({ ...editing, draft: { ...editing.draft, scope: event.target.value === '' ? null : Number(event.target.value) } })}>
                <option value="">Workbook</option>
                {sheetNames.map((name, index) => <option key={name} value={index}>{name}</option>)}
              </select>
            </label>
            <label className="nm-wide"><span>Refers to</span><input value={editing.draft.refersTo} spellCheck={false} onChange={(event) => setEditing({ ...editing, draft: { ...editing.draft, refersTo: event.target.value } })} /></label>
            <label className="nm-wide"><span>Comment</span><input value={editing.draft.comment || ''} onChange={(event) => setEditing({ ...editing, draft: { ...editing.draft, comment: event.target.value } })} /></label>
            {error && <p className="nm-error" role="alert">{error}</p>}
            <div className="nm-editor-actions">
              <button type="button" className="secondary-action" onClick={() => { setEditing(null); setError('') }}>Cancel</button>
              <button type="submit" className="primary-action">{editing.index === null ? 'Add name' : 'Update'}</button>
            </div>
          </form>
        )}
        <footer className="nm-footer">
          <button type="button" className="secondary-action" onClick={onClose}>Cancel</button>
          <button type="button" className="primary-action" disabled={Boolean(editing)} onClick={() => onSave(items)}>Save names</button>
        </footer>
      </div>
    </div>
  )
}
