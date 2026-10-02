// src/advanced/dialogs/ModifySelectionDialog.tsx (WP6)
// Select > Modify > Feather (Shift+F6) / Expand / Contract (design 5.8, 5.14): one amount in pixels.
// Feather softens the edge (Gaussian, sigma = radius / 2); Expand and Contract move the edge by an exact
// Euclidean distance. The work runs in the imaging worker; the result is one selection step (it does not
// mark the image as modified). Shows how big the selection is so the amount can be judged.
import { useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { modifySelection } from '../commands.ts'
import type { ModifySelectionOperation } from '../commands.ts'
import type { DialogDefinition, DialogProps } from '../panels/PanelHost.tsx'
import { NumberInput } from '../panels/AdjustmentControls.tsx'
import { DialogFrame } from './DialogFrame.tsx'

const TITLES: Readonly<Record<ModifySelectionOperation, string>> = Object.freeze({ feather: 'Feather Selection', expand: 'Expand Selection', contract: 'Contract Selection' })
const LABELS: Readonly<Record<ModifySelectionOperation, string>> = Object.freeze({ feather: 'Feather Radius', expand: 'Expand By', contract: 'Contract By' })

/** Last amounts used this session (Photoshop remembers them). */
const lastAmount: Record<ModifySelectionOperation, number> = { feather: 2, expand: 1, contract: 1 }

function ModifySelectionDialog({ context, request, close }: DialogProps): ReactNode {
  const operation: ModifySelectionOperation = request.kind === 'modify-selection' ? request.operation : 'feather'
  const selection = context.store.getState().selection
  const [amount, setAmountState] = useState(lastAmount[operation])
  const latest = useRef(amount)
  const setAmount = (value: number) => {
    latest.current = value
    setAmountState(value)
  }
  const [working, setWorking] = useState(false)
  const feather = operation === 'feather'

  const apply = async () => {
    if (working) return
    const value = latest.current
    setWorking(true)
    let ok = false
    try {
      ok = await modifySelection(context.commands, operation, value)
    } finally {
      setWorking(false)
    }
    if (ok) {
      lastAmount[operation] = value
      close()
      context.focusCanvas()
    }
  }

  return (
    <DialogFrame
      id={`modify-${operation}`}
      title={TITLES[operation]}
      width={300}
      busy={working}
      okDisabled={!selection || !(amount > 0)}
      onOk={() => void apply()}
      onCancel={() => { close(); context.focusCanvas() }}
    >
      <div className="ae-size-grid">
        <label htmlFor="ae-modify-amount">{LABELS[operation]}</label>
        <NumberInput
          id="ae-modify-amount"
          label={LABELS[operation]}
          value={amount}
          min={feather ? 0.1 : 1}
          max={feather ? 1000 : 500}
          step={feather ? 0.5 : 1}
          digits={feather ? 1 : 0}
          disabled={working}
          onChange={setAmount}
        />
        <span className="ae-unit-label">pixels</span>
      </div>
      <p className="ae-dialog-note">
        {selection
          ? `The selection is ${selection.bounds.width.toLocaleString()} × ${selection.bounds.height.toLocaleString()} px.`
          : 'There is no selection to change.'}
        {feather ? ' A larger radius makes a softer edge.' : ''}
      </p>
    </DialogFrame>
  )
}

export const dialog: DialogDefinition = {
  handles: (request) => request.kind === 'modify-selection',
  component: ModifySelectionDialog,
}
