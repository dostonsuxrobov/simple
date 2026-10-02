// src/advanced/dialogs/AdjustmentDialog.tsx (WP6)
// Image > Adjustments (destructive), Photoshop's dialogs (design 5.10, 5.14): the same controls as an
// adjustment layer's Properties, applied to the active layer (or its targeted mask) inside the selection.
//   - Preview (on by default) shows the result live on the canvas through the compositor's 'adjustment'
//     preview: nothing changes in the document until OK. Slider drags use a coarser proxy meanwhile.
//   - Levels / Curves / Threshold show the histogram of the pixels being changed; Levels has Auto.
//   - OK runs the adjustment in the imaging worker as one history step; the preview stays up until the
//     result lands, so the canvas never flashes the old image. Cancel (or Escape) leaves the layer as it
//     was; holding Alt turns Cancel into Reset.
//   - Masks have no canvas preview (the preview paints layer pixels); the dialog says so.
import { useEffect, useMemo, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import type { AdjustmentSpec, AdjustmentType } from '../../imaging/types.ts'
import { adjustmentLabel, defaultAdjustment } from '../../imaging/adjustments.ts'
import { autoTone } from '../../imaging/autoEnhance.ts'
import { computeHistogram } from '../../imaging/histogram.ts'
import { applyAdjustmentToLayer } from '../commands.ts'
import { activeLayerOf, setInteractive } from '../tools/shared.ts'
import type { DialogDefinition, DialogProps } from '../panels/PanelHost.tsx'
import { AdjustmentControls, CheckField, sampleOfTarget, selectedPixels } from '../panels/AdjustmentControls.tsx'
import { DialogFrame } from './DialogFrame.tsx'

const WIDTHS: Readonly<Partial<Record<AdjustmentType, number>>> = Object.freeze({ levels: 340, curves: 340, 'gradient-map': 340, 'black-white': 330 })

function AdjustmentDialog({ context, request, close }: DialogProps): ReactNode {
  const type: AdjustmentType = request.kind === 'adjustment' ? request.type : 'levels'
  const start = useRef(context.store.getState()).current
  const layer = activeLayerOf(start)
  const maskTarget = start.editTarget === 'mask' && Boolean(layer?.mask)
  const sample = useMemo(() => sampleOfTarget(start), [start])
  const histogram = useMemo(() => (sample ? computeHistogram(sample.pixels, sample.mask) : null), [sample])
  const [spec, setSpecState] = useState<AdjustmentSpec>(() => defaultAdjustment(type))
  const specRef = useRef(spec)
  const [preview, setPreviewOn] = useState(true)
  const [working, setWorking] = useState(false)
  const interactiveRef = useRef(false)
  const canPreview = Boolean(layer && !maskTarget)

  const setSpec = (next: AdjustmentSpec) => {
    specRef.current = next
    setSpecState(next)
  }

  const showPreview = (on: boolean, current: AdjustmentSpec) => {
    if (!layer || !canPreview) return
    context.compositor.setPreview(on ? { kind: 'adjustment', layerId: layer.id, spec: current, selection: start.selection } : null)
  }

  useEffect(() => {
    showPreview(preview, spec)
  }, [preview, spec])

  useEffect(() => () => {
    context.compositor.setPreview(null)
    if (interactiveRef.current) setInteractive(context.commands, false)
  }, [context])

  const onInteractive = (active: boolean) => {
    interactiveRef.current = active
    setInteractive(context.commands, active)
  }

  const apply = async () => {
    if (working) return
    setWorking(true)
    let ok = false
    try {
      ok = await applyAdjustmentToLayer(context.commands, specRef.current)
    } finally {
      // The committed pixels replace the preview before the next frame is drawn.
      if (canPreview) context.compositor.setPreview(null)
      setWorking(false)
    }
    if (ok) {
      close()
      context.focusCanvas()
    } else {
      showPreview(preview, specRef.current)
    }
  }

  const auto = () => {
    if (!sample) return
    try {
      setSpec(autoTone(selectedPixels(sample)))
    } catch (error) {
      context.host.notify(error instanceof Error ? error.message : 'Auto could not analyse the image.', 'error')
    }
  }

  const footer = (
    <>
      <CheckField
        label="Preview"
        title={canPreview ? 'Show the result on the canvas' : 'There is no canvas preview while a layer mask is the target'}
        checked={canPreview && preview}
        disabled={!canPreview || working}
        onChange={setPreviewOn}
      />
      {type === 'levels' && <button type="button" className="ae-text-button" disabled={!sample || working} title="Stretch each channel to the full range" onClick={auto}>Auto</button>}
    </>
  )

  return (
    <DialogFrame
      id={`adjust-${type}`}
      title={adjustmentLabel(type)}
      width={WIDTHS[type] ?? 310}
      busy={working}
      footerStart={footer}
      onOk={() => void apply()}
      onCancel={() => { close(); context.focusCanvas() }}
      onReset={() => setSpec(defaultAdjustment(type))}
    >
      {!layer && <p className="ae-dialog-warning" role="alert">Select a layer first.</p>}
      <AdjustmentControls spec={spec} histogram={histogram} disabled={working || !layer} onInteractive={onInteractive} onChange={(next) => setSpec(next)} />
      {maskTarget && <p className="ae-dialog-note">This changes the layer mask of “{layer?.name}”.</p>}
      {start.selection && <p className="ae-dialog-note">Only the selected area changes.</p>}
    </DialogFrame>
  )
}

export const dialog: DialogDefinition = {
  handles: (request) => request.kind === 'adjustment',
  component: AdjustmentDialog,
}
