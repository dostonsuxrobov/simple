import { useState } from 'react'
import { Grid2X2 } from 'lucide-react'
import { DataToolsDialog } from './DataToolsDialogFrame'

export type ShiftCellsChoice = 'shift-down' | 'shift-right' | 'shift-up' | 'shift-left' | 'entire-row' | 'entire-column'

interface ShiftCellsDialogProps {
  mode: 'insert' | 'delete'
  /** Excel's default: shift along the selection's shorter side. */
  initial: ShiftCellsChoice
  onApply: (choice: ShiftCellsChoice) => void
  onClose: () => void
}

/** Excel's Insert / Delete dialog for partial ranges. */
export function ShiftCellsDialog({ mode, initial, onApply, onClose }: ShiftCellsDialogProps) {
  const [choice, setChoice] = useState<ShiftCellsChoice>(initial)
  const options: Array<{ id: ShiftCellsChoice; label: string }> = mode === 'insert'
    ? [{ id: 'shift-right', label: 'Shift cells right' }, { id: 'shift-down', label: 'Shift cells down' }, { id: 'entire-row', label: 'Entire row' }, { id: 'entire-column', label: 'Entire column' }]
    : [{ id: 'shift-left', label: 'Shift cells left' }, { id: 'shift-up', label: 'Shift cells up' }, { id: 'entire-row', label: 'Entire row' }, { id: 'entire-column', label: 'Entire column' }]
  const apply = () => onApply(choice)
  return (
    <DataToolsDialog
      title={mode === 'insert' ? 'Insert cells' : 'Delete cells'}
      icon={<Grid2X2 size={16} />}
      width={320}
      className="shift-cells-dialog"
      onClose={onClose}
      onConfirm={apply}
      footer={(
        <>
          <span className="dt-footer-start" />
          <button type="button" className="dt-button" onClick={onClose}>Cancel</button>
          <button type="button" className="dt-button is-primary" onClick={apply}>OK</button>
        </>
      )}
    >
      <fieldset className="dt-fieldset" role="radiogroup" aria-label={mode === 'insert' ? 'Insert' : 'Delete'}>
        {options.map((option) => (
          <label key={option.id} className="dt-check">
            <input
              type="radio"
              name="shift-cells"
              checked={choice === option.id}
              data-autofocus={choice === option.id ? true : undefined}
              onChange={() => setChoice(option.id)}
            />
            {option.label}
          </label>
        ))}
      </fieldset>
    </DataToolsDialog>
  )
}
