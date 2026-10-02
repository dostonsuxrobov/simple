// src/advanced/dialogs/ImageSizeDialog.tsx (WP6)
// Image > Image Size (Ctrl+Alt+I), Photoshop's dialog (design 5.12, 5.14):
//   - width and height in pixels, percent, inches or centimetres, linked by "Constrain proportions";
//   - resolution in pixels per inch; with Resample off the pixels stay as they are and only the print
//     size / resolution change (like Photoshop);
//   - the resampling method (Automatic picks area averaging when shrinking a lot, Lanczos / bicubic
//     otherwise); the new size, megapixels and memory are shown before anything changes;
//   - OK resizes every layer and mask in the imaging worker as one history step (commands.resizeImage).
import { useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { Link2, Link2Off } from 'lucide-react'
import type { ResampleMethod } from '../../imaging/types.ts'
import { LIMITS } from '../types.ts'
import { resizeImage } from '../commands.ts'
import { formatBytes } from '../memory.ts'
import type { DialogDefinition, DialogProps } from '../panels/PanelHost.tsx'
import { CheckField, NumberInput, SelectField } from '../panels/AdjustmentControls.tsx'
import type { LengthUnit } from './DialogFrame.tsx'
import { DialogFrame, LENGTH_UNITS, fromUnit, toUnit, unitDigits } from './DialogFrame.tsx'

const METHODS: readonly (readonly [ResampleMethod, string])[] = [
  ['auto', 'Automatic'],
  ['bicubic', 'Bicubic (smooth gradients)'],
  ['lanczos3', 'Lanczos (sharper)'],
  ['bilinear', 'Bilinear'],
  ['area', 'Area average (reduction)'],
  ['nearest', 'Nearest Neighbor (hard edges)'],
]

interface Values {
  readonly width: number
  readonly height: number
  readonly ppi: number
  readonly resample: boolean
  readonly constrain: boolean
  readonly method: ResampleMethod
  readonly unit: LengthUnit
}

function ImageSizeDialog({ context, close }: DialogProps): ReactNode {
  const start = useRef(context.store.getState()).current
  const [values, setValuesState] = useState<Values>({
    width: start.width,
    height: start.height,
    ppi: start.ppi,
    resample: true,
    constrain: true,
    method: 'auto',
    unit: 'px',
  })
  // Enter in a field applies in the same key event: read the newest values from the ref.
  const valuesRef = useRef(values)
  const [working, setWorking] = useState(false)
  const set = (next: Values) => {
    valuesRef.current = next
    setValuesState(next)
  }
  const ratio = start.width / start.height

  const setWidth = (shown: number) => {
    const current = valuesRef.current
    if (!current.resample) {
      // Print size only: the pixels stay, the resolution follows.
      const ppi = current.unit === 'in' || current.unit === 'cm' ? start.width / (current.unit === 'in' ? shown : shown / 2.54) : current.ppi
      if (Number.isFinite(ppi) && ppi > 0) set({ ...current, ppi: Number(Math.min(10000, Math.max(1, ppi)).toFixed(3)) })
      return
    }
    const width = Math.max(1, fromUnit(shown, current.unit, start.width, current.ppi))
    const height = current.constrain ? Math.max(1, Math.round(width / ratio)) : current.height
    set({ ...current, width, height })
  }
  const setHeight = (shown: number) => {
    const current = valuesRef.current
    if (!current.resample) {
      const ppi = current.unit === 'in' || current.unit === 'cm' ? start.height / (current.unit === 'in' ? shown : shown / 2.54) : current.ppi
      if (Number.isFinite(ppi) && ppi > 0) set({ ...current, ppi: Number(Math.min(10000, Math.max(1, ppi)).toFixed(3)) })
      return
    }
    const height = Math.max(1, fromUnit(shown, current.unit, start.height, current.ppi))
    const width = current.constrain ? Math.max(1, Math.round(height * ratio)) : current.width
    set({ ...current, width, height })
  }
  const setPpi = (ppi: number) => {
    const current = valuesRef.current
    if (current.resample && (current.unit === 'in' || current.unit === 'cm')) {
      // Keeping the print size while changing the resolution changes the pixels (Photoshop).
      const width = Math.max(1, Math.round((current.width / current.ppi) * ppi))
      const height = current.constrain ? Math.max(1, Math.round(width / ratio)) : Math.max(1, Math.round((current.height / current.ppi) * ppi))
      set({ ...current, ppi, width, height })
    } else {
      set({ ...current, ppi })
    }
  }

  const { width, height, ppi, unit } = values
  const pixels = width * height
  const tooBig = width > LIMITS.maxDimension || height > LIMITS.maxDimension || pixels > LIMITS.maxPixels
  const changed = width !== start.width || height !== start.height || Math.abs(ppi - start.ppi) > 1e-6
  const digits = unitDigits(unit)
  // Without resampling the pixels cannot change; only a print size (inches, cm) can be typed.
  const pixelsFixed = !values.resample && (unit === 'px' || unit === 'percent')
  const layersWithPixels = start.layers.filter((layer) => layer.kind !== 'adjustment').length || 1

  const apply = async () => {
    if (working) return
    const current = valuesRef.current
    if (current.width === start.width && current.height === start.height && Math.abs(current.ppi - start.ppi) < 1e-6) {
      close()
      context.focusCanvas()
      return
    }
    setWorking(true)
    let ok = false
    try {
      ok = await resizeImage(context.commands, current.width, current.height, { method: current.method, ppi: current.ppi })
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
      id="image-size"
      title="Image Size"
      width={340}
      busy={working}
      okDisabled={tooBig}
      onOk={() => void apply()}
      onCancel={() => { close(); context.focusCanvas() }}
      onReset={() => set({ ...valuesRef.current, width: start.width, height: start.height, ppi: start.ppi })}
    >
      <p className="ae-dialog-summary">
        <strong>{formatBytes(pixels * 4)}</strong> per full layer{changed ? ` (was ${formatBytes(start.width * start.height * 4)})` : ''} · {(pixels / 1_000_000).toFixed(pixels < 10_000_000 ? 2 : 1)} MP
        {layersWithPixels > 1 ? ` · ${layersWithPixels} layers` : ''}
      </p>
      <div className="ae-size-grid">
        <label htmlFor="ae-is-width">Width</label>
        <NumberInput id="ae-is-width" label="Width" value={Number(toUnit(width, unit, start.width, ppi).toFixed(digits))} min={unit === 'px' ? 1 : 0.01} max={unit === 'px' ? LIMITS.maxDimension : 100000} step={1} digits={digits} disabled={working || pixelsFixed} onChange={setWidth} />
        <button
          type="button"
          className={`ae-chain ${values.constrain || !values.resample ? 'is-on' : ''}`}
          title={values.constrain ? 'Proportions are linked (click to unlink)' : 'Link width and height'}
          aria-label="Constrain proportions"
          aria-pressed={values.constrain || !values.resample}
          disabled={working || !values.resample}
          onClick={() => {
            const current = valuesRef.current
            const constrain = !current.constrain
            set({ ...current, constrain, height: constrain ? Math.max(1, Math.round(current.width / ratio)) : current.height })
          }}
        >{values.constrain || !values.resample ? <Link2 aria-hidden="true" /> : <Link2Off aria-hidden="true" />}</button>
        <label htmlFor="ae-is-height">Height</label>
        <NumberInput id="ae-is-height" label="Height" value={Number(toUnit(height, unit, start.height, ppi).toFixed(digits))} min={unit === 'px' ? 1 : 0.01} max={unit === 'px' ? LIMITS.maxDimension : 100000} step={1} digits={digits} disabled={working || pixelsFixed} onChange={setHeight} />
        <span />
        <label htmlFor="ae-is-unit">Unit</label>
        <select id="ae-is-unit" value={unit} disabled={working} onChange={(event) => set({ ...valuesRef.current, unit: event.currentTarget.value as LengthUnit })}>
          {LENGTH_UNITS.map(([id, label]) => <option key={id} value={id}>{label}</option>)}
        </select>
        <span />
        <label htmlFor="ae-is-ppi">Resolution</label>
        <NumberInput id="ae-is-ppi" label="Resolution in pixels per inch" value={ppi} min={1} max={10000} step={1} digits={3} disabled={working} onChange={setPpi} />
        <span className="ae-unit-label">ppi</span>
      </div>
      <p className="ae-dialog-note">{width.toLocaleString()} × {height.toLocaleString()} px · prints at {(width / ppi).toFixed(2)} × {(height / ppi).toFixed(2)} in</p>
      <div className="ae-inline-row">
        <CheckField
          label="Resample"
          title="Off: keep every pixel and only change the resolution (print size)"
          checked={values.resample}
          disabled={working}
          onChange={(resample) => {
            const current = valuesRef.current
            set(resample ? { ...current, resample } : { ...current, resample, width: start.width, height: start.height, constrain: true })
          }}
        />
      </div>
      <SelectField label="Resampling" value={values.method} options={METHODS} disabled={working || !values.resample} onChange={(method) => set({ ...valuesRef.current, method })} />
      {tooBig && <p className="ae-dialog-warning" role="alert">That is larger than {LIMITS.maxDimension.toLocaleString()} px per side or {LIMITS.maxPixels / 1_000_000} megapixels.</p>}
    </DialogFrame>
  )
}

export const dialog: DialogDefinition = {
  handles: (request) => request.kind === 'image-size',
  component: ImageSizeDialog,
}
