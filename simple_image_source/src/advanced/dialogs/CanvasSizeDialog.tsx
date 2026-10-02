// src/advanced/dialogs/CanvasSizeDialog.tsx (WP6)
// Image > Canvas Size (Ctrl+Alt+C), Photoshop's dialog (design 5.12, 5.14): the new width and height in
// pixels, percent, inches or centimetres, absolute or Relative (added to the current size), and the 3 x 3
// anchor that says where the image stays. Layers keep their pixels outside a smaller canvas (they are not
// cut); the Background is extended with the background colour. OK is one history step
// (commands.resizeCanvas).
import { useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { ArrowDown, ArrowDownLeft, ArrowDownRight, ArrowLeft, ArrowRight, ArrowUp, ArrowUpLeft, ArrowUpRight, Dot } from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import { LIMITS } from '../types.ts'
import type { CanvasAnchor } from '../commands.ts'
import { CANVAS_ANCHORS, resizeCanvas } from '../commands.ts'
import { formatBytes } from '../memory.ts'
import { isBackgroundLayer } from '../tools/shared.ts'
import type { DialogDefinition, DialogProps } from '../panels/PanelHost.tsx'
import { useEditorState } from '../panels/PanelHost.tsx'
import { CheckField, NumberInput } from '../panels/AdjustmentControls.tsx'
import { cssColor } from '../panels/ColorPicker.tsx'
import type { LengthUnit } from './DialogFrame.tsx'
import { DialogFrame, LENGTH_UNITS, fromUnit, toUnit, unitDigits } from './DialogFrame.tsx'

const ANCHOR_LABELS: Readonly<Record<CanvasAnchor, string>> = Object.freeze({
  'top-left': 'Top left', top: 'Top', 'top-right': 'Top right', left: 'Left', center: 'Centre', right: 'Right',
  'bottom-left': 'Bottom left', bottom: 'Bottom', 'bottom-right': 'Bottom right',
})

const ANCHOR_COLUMN: Readonly<Record<CanvasAnchor, number>> = Object.freeze({
  'top-left': 0, top: 1, 'top-right': 2, left: 0, center: 1, right: 2, 'bottom-left': 0, bottom: 1, 'bottom-right': 2,
})
const ANCHOR_ROW: Readonly<Record<CanvasAnchor, number>> = Object.freeze({
  'top-left': 0, top: 0, 'top-right': 0, left: 1, center: 1, right: 1, 'bottom-left': 2, bottom: 2, 'bottom-right': 2,
})

/** Photoshop's anchor grid: arrows point away from the anchor (where the canvas grows). */
function arrowFor(cell: CanvasAnchor, anchor: CanvasAnchor): LucideIcon | null {
  const dx = ANCHOR_COLUMN[cell] - ANCHOR_COLUMN[anchor]
  const dy = ANCHOR_ROW[cell] - ANCHOR_ROW[anchor]
  if (dx === 0 && dy === 0) return Dot
  if (Math.abs(dx) > 1 || Math.abs(dy) > 1) return null
  if (dx === 0) return dy < 0 ? ArrowUp : ArrowDown
  if (dy === 0) return dx < 0 ? ArrowLeft : ArrowRight
  if (dx < 0) return dy < 0 ? ArrowUpLeft : ArrowDownLeft
  return dy < 0 ? ArrowUpRight : ArrowDownRight
}

interface Values {
  /** New size in pixels. */
  readonly width: number
  readonly height: number
  readonly relative: boolean
  readonly unit: LengthUnit
  readonly anchor: CanvasAnchor
}

function CanvasSizeDialog({ context, close }: DialogProps): ReactNode {
  const start = useRef(context.store.getState()).current
  const editor = useEditorState(context.editor)
  const [values, setValuesState] = useState<Values>({ width: start.width, height: start.height, relative: false, unit: 'px', anchor: 'center' })
  const valuesRef = useRef(values)
  const set = (next: Values) => {
    valuesRef.current = next
    setValuesState(next)
  }
  const [working, setWorking] = useState(false)
  const { width, height, unit, relative, anchor } = values
  const digits = unitDigits(unit)

  /** The number shown for a pixel size (relative mode shows the change). */
  const shown = (pixels: number, reference: number) => {
    const value = relative ? toUnit(pixels - reference, unit, reference, start.ppi) : toUnit(pixels, unit, reference, start.ppi)
    return Number(value.toFixed(digits))
  }
  const pixelsFor = (value: number, reference: number) => {
    const current = valuesRef.current
    if (current.relative) return reference + (current.unit === 'percent' ? Math.round((reference * value) / 100) : fromUnit(value, current.unit, reference, start.ppi))
    return fromUnit(value, current.unit, reference, start.ppi)
  }

  const tooBig = width < 1 || height < 1 || width > LIMITS.maxDimension || height > LIMITS.maxDimension || width * height > LIMITS.maxPixels
  const hasBackground = isBackgroundLayer(start.layers[0])

  const apply = () => {
    if (working) return
    const current = valuesRef.current
    if (current.width === start.width && current.height === start.height) {
      close()
      context.focusCanvas()
      return
    }
    setWorking(true)
    let ok = false
    try {
      ok = resizeCanvas(context.commands, current.width, current.height, current.anchor)
    } finally {
      setWorking(false)
    }
    if (ok) {
      close()
      context.focusCanvas()
    }
  }

  const min = relative ? -100000 : unit === 'px' ? 1 : 0.01
  return (
    <DialogFrame
      id="canvas-size"
      title="Canvas Size"
      width={340}
      busy={working}
      okDisabled={tooBig}
      onOk={apply}
      onCancel={() => { close(); context.focusCanvas() }}
      onReset={() => set({ ...valuesRef.current, width: start.width, height: start.height, anchor: 'center' })}
    >
      <p className="ae-dialog-summary">Current size: <strong>{start.width.toLocaleString()} × {start.height.toLocaleString()} px</strong> ({formatBytes(start.width * start.height * 4)})</p>
      <div className="ae-size-grid">
        <label htmlFor="ae-cs-width">{relative ? 'Add width' : 'Width'}</label>
        <NumberInput id="ae-cs-width" label="Width" value={shown(width, start.width)} min={min} max={100000} digits={digits} disabled={working} onChange={(value) => set({ ...valuesRef.current, width: pixelsFor(value, start.width) })} />
        <span />
        <label htmlFor="ae-cs-height">{relative ? 'Add height' : 'Height'}</label>
        <NumberInput id="ae-cs-height" label="Height" value={shown(height, start.height)} min={min} max={100000} digits={digits} disabled={working} onChange={(value) => set({ ...valuesRef.current, height: pixelsFor(value, start.height) })} />
        <span />
        <label htmlFor="ae-cs-unit">Unit</label>
        <select id="ae-cs-unit" value={unit} disabled={working} onChange={(event) => set({ ...valuesRef.current, unit: event.currentTarget.value as LengthUnit })}>
          {LENGTH_UNITS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
        </select>
        <span />
      </div>
      <CheckField label="Relative" title="Add to (or take from) the current size" checked={relative} disabled={working} onChange={(next) => set({ ...valuesRef.current, relative: next })} />
      <div className="ae-anchor-row">
        <span className="ae-anchor-caption">Anchor</span>
        <div className="ae-anchor" role="radiogroup" aria-label="Anchor">
          {CANVAS_ANCHORS.map((cell) => {
            const Icon = arrowFor(cell, anchor)
            return (
              <button
                key={cell}
                type="button"
                role="radio"
                aria-checked={cell === anchor}
                aria-label={ANCHOR_LABELS[cell]}
                title={ANCHOR_LABELS[cell]}
                className={cell === anchor ? 'is-active' : ''}
                disabled={working}
                onClick={() => set({ ...valuesRef.current, anchor: cell })}
              >{Icon ? <Icon aria-hidden="true" /> : null}</button>
            )
          })}
        </div>
        <p className="ae-dialog-note">
          New size: <strong>{width.toLocaleString()} × {height.toLocaleString()} px</strong>.{' '}
          {hasBackground
            ? <>New Background area: <i className="ae-color-dot" style={{ background: cssColor(editor.background) }} /> background colour.</>
            : 'New areas are transparent.'}
          {' '}Pixels outside a smaller canvas are kept on their layers.
        </p>
      </div>
      {tooBig && <p className="ae-dialog-warning" role="alert">The canvas must be 1 to {LIMITS.maxDimension.toLocaleString()} px per side and at most {LIMITS.maxPixels / 1_000_000} megapixels.</p>}
    </DialogFrame>
  )
}

export const dialog: DialogDefinition = {
  handles: (request) => request.kind === 'canvas-size',
  component: CanvasSizeDialog,
}
