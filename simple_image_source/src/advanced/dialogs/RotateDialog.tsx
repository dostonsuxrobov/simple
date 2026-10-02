// src/advanced/dialogs/RotateDialog.tsx (WP6)
// Image > Image Rotation > Arbitrary (design 5.12, 5.14): an angle and its direction (clockwise or
// counter-clockwise, as in Photoshop). The canvas grows to hold the whole turned image; the Background's
// new corners take the background colour, other layers stay transparent there. Multiples of 90 degrees
// turn exactly (no resampling). Shows the resulting canvas size before anything changes.
import { useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { RotateCcw, RotateCw } from 'lucide-react'
import { LIMITS } from '../types.ts'
import { rotateCanvas } from '../commands.ts'
import type { DialogDefinition, DialogProps } from '../panels/PanelHost.tsx'
import { NumberInput, Segmented } from '../panels/AdjustmentControls.tsx'
import { DialogFrame } from './DialogFrame.tsx'

type Direction = 'cw' | 'ccw'

/** Canvas size after turning a width x height canvas by `degrees` (the bounds of the turned image). */
export function rotatedCanvasSize(width: number, height: number, degrees: number): { readonly width: number; readonly height: number } {
  const normalized = ((degrees % 360) + 360) % 360
  if (normalized === 0 || normalized === 180) return { width, height }
  if (normalized === 90 || normalized === 270) return { width: height, height: width }
  const radians = (normalized * Math.PI) / 180
  const cos = Math.abs(Math.cos(radians))
  const sin = Math.abs(Math.sin(radians))
  return { width: Math.max(1, Math.round(width * cos + height * sin)), height: Math.max(1, Math.round(width * sin + height * cos)) }
}

function RotateDialog({ context, close }: DialogProps): ReactNode {
  const start = useRef(context.store.getState()).current
  const [angle, setAngleState] = useState(0)
  const [direction, setDirectionState] = useState<Direction>('cw')
  const latest = useRef({ angle: 0, direction: 'cw' as Direction })
  const setAngle = (value: number) => {
    latest.current = { ...latest.current, angle: value }
    setAngleState(value)
  }
  const setDirection = (value: Direction) => {
    latest.current = { ...latest.current, direction: value }
    setDirectionState(value)
  }
  const [working, setWorking] = useState(false)
  const degrees = direction === 'cw' ? angle : -angle
  const size = rotatedCanvasSize(start.width, start.height, degrees)
  const tooBig = size.width > LIMITS.maxDimension || size.height > LIMITS.maxDimension || size.width * size.height > LIMITS.maxPixels

  const apply = async () => {
    if (working) return
    const { angle: value, direction: way } = latest.current
    if (!value) {
      close()
      context.focusCanvas()
      return
    }
    setWorking(true)
    let ok = false
    try {
      ok = await rotateCanvas(context.commands, way === 'cw' ? value : -value)
    } finally {
      setWorking(false)
    }
    if (ok) {
      close()
      context.focusCanvas()
    }
  }

  return (
    <DialogFrame
      id="rotate-canvas"
      title="Rotate Canvas"
      width={300}
      busy={working}
      okDisabled={tooBig}
      onOk={() => void apply()}
      onCancel={() => { close(); context.focusCanvas() }}
      onReset={() => { setAngle(0); setDirection('cw') }}
    >
      <div className="ae-size-grid">
        <label htmlFor="ae-rot-angle">Angle</label>
        <NumberInput id="ae-rot-angle" label="Angle in degrees" value={angle} min={-359.99} max={359.99} step={1} digits={2} disabled={working} onChange={setAngle} />
        <span className="ae-unit-label">°</span>
      </div>
      <div className="ae-inline-row">
        <Segmented label="Direction" value={direction} options={[['cw', '° Clockwise'], ['ccw', '° Counter Clockwise']]} disabled={working} onChange={setDirection} />
        {direction === 'cw' ? <RotateCw className="ae-dialog-icon" aria-hidden="true" /> : <RotateCcw className="ae-dialog-icon" aria-hidden="true" />}
      </div>
      <p className="ae-dialog-note">New canvas: <strong>{size.width.toLocaleString()} × {size.height.toLocaleString()} px</strong>{Math.abs(angle) % 90 === 0 ? ' (exact turn, no resampling)' : '. New corners are transparent, or the background colour on the Background.'}</p>
      {tooBig && <p className="ae-dialog-warning" role="alert">The turned image would be larger than the {LIMITS.maxPixels / 1_000_000} megapixel limit.</p>}
    </DialogFrame>
  )
}

export const dialog: DialogDefinition = {
  handles: (request) => request.kind === 'rotate-canvas',
  component: RotateDialog,
}
