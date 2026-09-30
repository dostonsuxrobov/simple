import { useState } from 'react'
import { Target } from 'lucide-react'
import { DataToolsDialog } from './DataToolsDialogFrame'

interface GoalSeekDialogProps {
  initialSetCell: string
  initialChangingCell: string
  /** Runs the search; returns an error or status message to keep the dialog open. */
  onSeek: (setCell: string, toValue: string, changingCell: string) => string | null
  onClose: () => void
}

/** Excel's Goal Seek: make a formula reach a value by changing one input cell. */
export function GoalSeekDialog({ initialSetCell, initialChangingCell, onSeek, onClose }: GoalSeekDialogProps) {
  const [setCell, setSetCell] = useState(initialSetCell)
  const [toValue, setToValue] = useState('')
  const [changingCell, setChangingCell] = useState(initialChangingCell)
  const [message, setMessage] = useState<string | null>(null)
  const seek = () => setMessage(onSeek(setCell, toValue, changingCell))
  return (
    <DataToolsDialog
      title="Goal seek"
      subtitle="Find the input value that makes a formula reach a target."
      icon={<Target size={16} />}
      width={380}
      className="goal-seek-dialog"
      onClose={onClose}
      onConfirm={seek}
      footer={(
        <>
          <span className="dt-footer-start" />
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" onClick={seek}>OK</button>
        </>
      )}
    >
      <label className="dt-field">
        <span>Set cell</span>
        <input className="dt-input dt-mono" aria-label="Set cell" value={setCell} spellCheck={false} onChange={(event) => { setSetCell(event.target.value); setMessage(null) }} />
      </label>
      <label className="dt-field">
        <span>To value</span>
        <input className="dt-input dt-mono" aria-label="To value" value={toValue} inputMode="decimal" data-autofocus onChange={(event) => { setToValue(event.target.value); setMessage(null) }} />
      </label>
      <label className="dt-field">
        <span>By changing cell</span>
        <input className="dt-input dt-mono" aria-label="By changing cell" value={changingCell} spellCheck={false} onChange={(event) => { setChangingCell(event.target.value); setMessage(null) }} />
      </label>
      {message && <p className="dt-error" role="alert">{message}</p>}
    </DataToolsDialog>
  )
}
