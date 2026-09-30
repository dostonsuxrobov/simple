import { useState } from 'react'
import { ChartSpline } from 'lucide-react'
import { DataToolsDialog } from './DataToolsDialogFrame'

export type SparklineKind = 'line' | 'column' | 'stacked'

interface CreateSparklinesDialogProps {
  kind: SparklineKind
  initialData: string
  initialLocation: string
  onCreate: (data: string, location: string, kind: SparklineKind) => string | null
  onClose: () => void
}

const TITLES: Record<SparklineKind, string> = { line: 'Line sparklines', column: 'Column sparklines', stacked: 'Win/Loss sparklines' }

/** Excel's Create Sparklines dialog: a data range and the cells that show one sparkline each. */
export function CreateSparklinesDialog({ kind, initialData, initialLocation, onCreate, onClose }: CreateSparklinesDialogProps) {
  const [data, setData] = useState(initialData)
  const [location, setLocation] = useState(initialLocation)
  const [error, setError] = useState<string | null>(null)
  const create = () => setError(onCreate(data.trim().replace(/^=/, ''), location.trim().replace(/^=/, ''), kind))
  return (
    <DataToolsDialog
      title={TITLES[kind]}
      subtitle="Draw a small chart inside each location cell, one per row (or column) of the data."
      icon={<ChartSpline size={16} />}
      width={400}
      className="create-sparklines-dialog"
      onClose={onClose}
      onConfirm={create}
      footer={(
        <>
          <span className="dt-footer-start" />
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" onClick={create}>OK</button>
        </>
      )}
    >
      <label className="dt-field">
        <span>Data range</span>
        <input className="dt-input dt-mono" aria-label="Sparkline data range" value={data} spellCheck={false} data-autofocus onChange={(event) => { setData(event.target.value); setError(null) }} />
      </label>
      <label className="dt-field">
        <span>Location range</span>
        <input className="dt-input dt-mono" aria-label="Sparkline location range" value={location} spellCheck={false} onChange={(event) => { setLocation(event.target.value); setError(null) }} />
      </label>
      {error && <p className="dt-error" role="alert">{error}</p>}
    </DataToolsDialog>
  )
}
