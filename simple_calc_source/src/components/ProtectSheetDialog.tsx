import { useState } from 'react'
import { Lock } from 'lucide-react'
import { PROTECTION_OPTIONS } from '../lib/protection'
import { DataToolsDialog } from './DataToolsDialogFrame'

export interface ProtectSheetResult {
  password: string
  allow: Record<string, boolean>
}

interface ProtectSheetDialogProps {
  sheetName: string
  onProtect: (result: ProtectSheetResult) => Promise<string | null>
  onClose: () => void
}

/** Excel's Protect Sheet dialog: an optional password and what users may still do. */
export function ProtectSheetDialog({ sheetName, onProtect, onClose }: ProtectSheetDialogProps) {
  const [password, setPassword] = useState('')
  const [confirm, setConfirm] = useState('')
  const [allow, setAllow] = useState<Record<string, boolean>>(() => Object.fromEntries(PROTECTION_OPTIONS.map((option) => [option.id, option.defaultOn])))
  const [error, setError] = useState<string | null>(null)
  const [busy, setBusy] = useState(false)
  const submit = async () => {
    if (password !== confirm) { setError('The passwords do not match.'); return }
    setBusy(true)
    const problem = await onProtect({ password, allow })
    setBusy(false)
    if (problem) setError(problem)
  }
  return (
    <DataToolsDialog
      title="Protect sheet"
      subtitle={`Lock the cells of “${sheetName}” against changes. Unlocked cells stay editable.`}
      icon={<Lock size={16} />}
      width={400}
      className="protect-sheet-dialog"
      busy={busy}
      onClose={onClose}
      onConfirm={() => { void submit() }}
      footer={(
        <>
          <span className="dt-footer-start" />
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" disabled={busy} onClick={() => { void submit() }}>Protect</button>
        </>
      )}
    >
      <label className="dt-field">
        <span>Password to unprotect (optional)</span>
        <input className="dt-input" type="password" aria-label="Password" value={password} autoComplete="new-password" data-autofocus onChange={(event) => { setPassword(event.target.value); setError(null) }} />
      </label>
      {password && (
        <label className="dt-field">
          <span>Re-enter password</span>
          <input className="dt-input" type="password" aria-label="Confirm password" value={confirm} autoComplete="new-password" onChange={(event) => { setConfirm(event.target.value); setError(null) }} />
        </label>
      )}
      <fieldset className="dt-fieldset protect-options" aria-label="Allow all users of this sheet to">
        <legend>Allow all users of this sheet to</legend>
        {PROTECTION_OPTIONS.map((option) => (
          <label key={option.id} className="dt-check">
            <input type="checkbox" checked={allow[option.id]} onChange={(event) => setAllow((current) => ({ ...current, [option.id]: event.target.checked }))} />
            {option.label}
          </label>
        ))}
      </fieldset>
      {error && <p className="dt-error" role="alert">{error}</p>}
    </DataToolsDialog>
  )
}
