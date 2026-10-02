// src/advanced/dialogs/FilterDialog.tsx (WP6)
// The Filter menu's dialogs (design 5.11, 5.14), one component for every filter with settings:
//   - a 100% preview window computed in the imaging worker: drag it to look at another spot, press and
//     hold to compare with the original (Photoshop's filter preview);
//   - Preview (on by default) also shows the result on the canvas at the displayed zoom: the visible part
//     of the layer is filtered at the compositor's proxy level (radii scaled to match) and shown through
//     the 'layer-pixels' preview, so nothing in the document changes until OK;
//   - previews are debounced and stale worker jobs are aborted while the settings change;
//   - OK applies the filter at full resolution (commands.applyFilterToLayer: inside the selection, edges
//     read their neighbours, one history step, remembered for Filter > Last Filter). The settings start
//     from the last use of the same filter.
// Masks have no canvas preview (it paints layer pixels); the 100% window shows the mask as grey.
import { useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import { LoaderCircle } from 'lucide-react'
import type { FilterSpec, FilterType, IntRect, MaskBuffer, PixelBuffer, Point } from '../../imaging/types.ts'
import type { DocumentState, Layer } from '../types.ts'
import { defaultFilter, filterLabel, filterMargin } from '../../imaging/filters.ts'
import { applyFilterToLayer } from '../commands.ts'
import { levelOffset, readSurfaceLevel } from '../pyramid.ts'
import { selectionCoverage } from '../selection.ts'
import {
  activeLayerOf,
  canvasRect,
  isAbortError,
  isBackgroundLayer,
  layerPixelBounds,
  pixelSourceOf,
  readLayerRect,
  readMaskRect,
  visibleLevelRect,
} from '../tools/shared.ts'
import type { DialogDefinition, DialogProps, PanelContext } from '../panels/PanelHost.tsx'
import { CheckField, SelectField, SliderField, sampleSelection } from '../panels/AdjustmentControls.tsx'
import { DialogFrame } from './DialogFrame.tsx'

// ---------------------------------------------------------------------------------------------
// Settings per filter
// ---------------------------------------------------------------------------------------------

type Field =
  | { readonly kind: 'number'; readonly key: string; readonly label: string; readonly min: number; readonly max: number; readonly step?: number; readonly digits?: number; readonly unit?: string }
  | { readonly kind: 'select'; readonly key: string; readonly label: string; readonly options: readonly (readonly [string, string])[] }
  | { readonly kind: 'boolean'; readonly key: string; readonly label: string }

const num = (key: string, label: string, min: number, max: number, extra: { step?: number; digits?: number; unit?: string } = {}): Field => ({ kind: 'number', key, label, min, max, ...extra })

export const FILTER_FIELDS: Readonly<Record<FilterType, readonly Field[]>> = Object.freeze({
  'gaussian-blur': [num('radius', 'Radius', 0.1, 250, { step: 0.1, digits: 1, unit: 'px' })],
  'motion-blur': [num('angle', 'Angle', -90, 90, { unit: '°' }), num('distance', 'Distance', 1, 999, { unit: 'px' })],
  'unsharp-mask': [
    num('amount', 'Amount', 1, 500, { unit: '%' }),
    num('radius', 'Radius', 0.1, 250, { step: 0.1, digits: 1, unit: 'px' }),
    num('threshold', 'Threshold', 0, 255, { unit: 'levels' }),
  ],
  sharpen: [{ kind: 'select', key: 'strength', label: 'Strength', options: [['normal', 'Sharpen'], ['more', 'Sharpen More']] }],
  'add-noise': [
    num('amount', 'Amount', 0.1, 400, { step: 0.1, digits: 1, unit: '%' }),
    { kind: 'select', key: 'distribution', label: 'Distribution', options: [['uniform', 'Uniform'], ['gaussian', 'Gaussian']] },
    { kind: 'boolean', key: 'monochromatic', label: 'Monochromatic' },
  ],
  median: [num('radius', 'Radius', 1, 100, { unit: 'px' })],
  'reduce-noise': [num('strength', 'Strength', 0, 10), num('preserveDetails', 'Preserve Details', 0, 100, { unit: '%' })],
  pixelate: [num('cellSize', 'Cell Size', 2, 200, { unit: 'square' })],
  emboss: [num('angle', 'Angle', -180, 180, { unit: '°' }), num('height', 'Height', 1, 10, { unit: 'px' }), num('amount', 'Amount', 1, 500, { unit: '%' })],
  'find-edges': [],
})

/** The same filter for a proxy level: distances shrink with the image (a display approximation). */
export function scaleFilterForLevel(spec: FilterSpec, level: number): FilterSpec {
  if (level <= 0) return spec
  const scale = 2 ** level
  switch (spec.type) {
    case 'gaussian-blur': return { ...spec, radius: Math.max(0.1, spec.radius / scale) }
    case 'motion-blur': return { ...spec, distance: Math.max(1, Math.round(spec.distance / scale)) }
    case 'unsharp-mask': return { ...spec, radius: Math.max(0.1, spec.radius / scale) }
    case 'median': return { ...spec, radius: Math.max(1, Math.round(spec.radius / scale)) }
    case 'pixelate': return { ...spec, cellSize: Math.max(2, Math.round(spec.cellSize / scale)) }
    case 'emboss': return { ...spec, height: Math.max(1, Math.round(spec.height / scale)) }
    default: return spec
  }
}

// ---------------------------------------------------------------------------------------------
// Reading what the filter works on
// ---------------------------------------------------------------------------------------------

function grayToPixels(values: MaskBuffer): PixelBuffer {
  const data = new Uint8ClampedArray(values.width * values.height * 4)
  for (let i = 0, p = 0; i < values.data.length; i += 1, p += 4) {
    const v = values.data[i]
    data[p] = v
    data[p + 1] = v
    data[p + 2] = v
    data[p + 3] = 255
  }
  return { width: values.width, height: values.height, data }
}

/** Document pixels of the filter target (layer, or targeted mask as grey) over `rect`. */
function readTarget(state: DocumentState, layer: Layer, rect: IntRect): PixelBuffer {
  if (state.editTarget === 'mask' && layer.mask) return grayToPixels(readMaskRect(layer.mask, rect))
  return readLayerRect(layer, rect)
}

/** An `outer`-sized mask: selection coverage inside `inner` (offset by the margin), 0 in the margin. */
function marginMask(outerWidth: number, outerHeight: number, margin: number, coverage: MaskBuffer | null, innerWidth: number, innerHeight: number): MaskBuffer | null {
  if (!coverage && margin === 0) return null
  const data = new Uint8Array(outerWidth * outerHeight)
  for (let y = 0; y < innerHeight; y += 1) {
    const row = (y + margin) * outerWidth + margin
    if (coverage) data.set(coverage.data.subarray(y * innerWidth, (y + 1) * innerWidth), row)
    else data.fill(255, row, row + innerWidth)
  }
  return { width: outerWidth, height: outerHeight, data }
}

function cropBuffer(src: PixelBuffer, x: number, y: number, width: number, height: number): PixelBuffer {
  const out = new Uint8ClampedArray(width * height * 4)
  for (let row = 0; row < height; row += 1) {
    const from = ((y + row) * src.width + x) * 4
    out.set(src.data.subarray(from, from + width * 4), row * width * 4)
  }
  return { width, height, data: out }
}

function putPixels(canvas: HTMLCanvasElement | null, pixels: PixelBuffer | null): void {
  const context = canvas?.getContext('2d')
  if (!canvas || !context) return
  if (!pixels) {
    context.clearRect(0, 0, canvas.width, canvas.height)
    return
  }
  context.putImageData(new ImageData(pixels.data as Uint8ClampedArray<ArrayBuffer>, pixels.width, pixels.height), 0, 0)
}

const BOX_WIDTH = 248
const BOX_HEIGHT = 150

function deviceScale(): number {
  return typeof window !== 'undefined' && window.devicePixelRatio > 0 ? Math.min(3, window.devicePixelRatio) : 1
}

function initialCenter(context: PanelContext, state: DocumentState, layer: Layer | null): Point {
  if (state.selection) {
    const b = state.selection.bounds
    return { x: b.x + b.width / 2, y: b.y + b.height / 2 }
  }
  const size = context.view.getViewportSize()
  const visible = size.width > 0 && size.height > 0 ? context.view.screenToDoc({ x: size.width / 2, y: size.height / 2 }) : { x: state.width / 2, y: state.height / 2 }
  const bounds = layer && !isBackgroundLayer(layer) ? layerPixelBounds(layer) ?? canvasRect(state) : canvasRect(state)
  return {
    x: Math.min(bounds.x + bounds.width, Math.max(bounds.x, visible.x)),
    y: Math.min(bounds.y + bounds.height, Math.max(bounds.y, visible.y)),
  }
}

// ---------------------------------------------------------------------------------------------
// The dialog
// ---------------------------------------------------------------------------------------------

function FilterDialog({ context, request, close }: DialogProps): ReactNode {
  const type: FilterType = request.kind === 'filter' ? request.type : 'gaussian-blur'
  const start = useRef(context.store.getState()).current
  const layer = activeLayerOf(start)
  const maskTarget = start.editTarget === 'mask' && Boolean(layer?.mask)
  const canvasPreviewPossible = Boolean(layer && !maskTarget && pixelSourceOf(layer))
  const [spec, setSpecState] = useState<FilterSpec>(() => {
    const last = context.editor.getState().lastFilter
    return last && last.type === type ? last : defaultFilter(type)
  })
  const specRef = useRef(spec)
  const [canvasPreview, setCanvasPreview] = useState(true)
  const [working, setWorking] = useState(false)
  const [center, setCenter] = useState<Point>(() => initialCenter(context, start, layer))
  const [boxBusy, setBoxBusy] = useState(false)
  const [canvasBusy, setCanvasBusy] = useState(false)
  /** Bumped to compute the canvas preview again (after an OK that could not apply). */
  const [previewRound, setPreviewRound] = useState(0)
  const [comparing, setComparingState] = useState(false)
  const comparingRef = useRef(false)
  const setComparing = (value: boolean) => {
    comparingRef.current = value
    setComparingState(value)
  }
  const canvasRef = useRef<HTMLCanvasElement>(null)
  const boxPixels = useRef<{ before: PixelBuffer; after: PixelBuffer | null } | null>(null)
  const scale = deviceScale()
  const boxWidth = Math.round(BOX_WIDTH * scale)
  const boxHeight = Math.round(BOX_HEIGHT * scale)
  const fields = FILTER_FIELDS[type]

  const setSpec = (next: FilterSpec) => {
    specRef.current = next
    setSpecState(next)
  }

  // ----- 100% preview window (worker, debounced, stale jobs aborted) -------------------------------
  useEffect(() => {
    if (!layer) return
    const controller = new AbortController()
    const timer = setTimeout(() => {
      const rect: IntRect = { x: Math.round(center.x - boxWidth / 2), y: Math.round(center.y - boxHeight / 2), width: boxWidth, height: boxHeight }
      const margin = Math.max(0, Math.ceil(filterMargin(spec)))
      const outer: IntRect = { x: rect.x - margin, y: rect.y - margin, width: rect.width + 2 * margin, height: rect.height + 2 * margin }
      let source: PixelBuffer
      try {
        source = readTarget(start, layer, outer)
      } catch (error) {
        console.error(error)
        return
      }
      const before = cropBuffer(source, margin, margin, rect.width, rect.height)
      boxPixels.current = { before, after: boxPixels.current?.after ?? null }
      const coverage = start.selection ? selectionCoverage(start.selection, rect) : null
      const mask = marginMask(outer.width, outer.height, margin, coverage, rect.width, rect.height)
      setBoxBusy(true)
      context.imaging.run('filter', { src: source, spec, mask }, { signal: controller.signal })
        .then((filtered) => {
          if (controller.signal.aborted) return
          const after = cropBuffer(filtered, margin, margin, rect.width, rect.height)
          boxPixels.current = { before, after }
          putPixels(canvasRef.current, comparingRef.current ? before : after)
          setBoxBusy(false)
        })
        .catch((error: unknown) => {
          if (!isAbortError(error)) {
            console.error(error)
            setBoxBusy(false)
          }
        })
    }, 90)
    return () => {
      clearTimeout(timer)
      controller.abort()
    }
  }, [center.x, center.y, spec, boxWidth, boxHeight])

  useEffect(() => {
    const pixels = boxPixels.current
    if (pixels) putPixels(canvasRef.current, comparing || !pixels.after ? pixels.before : pixels.after)
  }, [comparing])

  // ----- canvas preview at the displayed proxy level ---------------------------------------------
  useEffect(() => {
    if (!layer || !canvasPreviewPossible) return
    if (!canvasPreview) {
      context.compositor.setPreview(null)
      setCanvasBusy(false)
      return
    }
    const source = pixelSourceOf(layer)
    if (!source) return
    const controller = new AbortController()
    const timer = setTimeout(() => {
      const level = Math.max(0, context.compositor.level | 0)
      const visible = visibleLevelRect(context.commands, start, level)
      if (!visible) return
      const scaled = scaleFilterForLevel(spec, level)
      const margin = Math.max(0, Math.ceil(filterMargin(scaled)))
      let src: PixelBuffer
      try {
        src = readSurfaceLevel(source.surface, level, {
          x: visible.x - margin - levelOffset(source.offsetX, level),
          y: visible.y - margin - levelOffset(source.offsetY, level),
          width: visible.width + 2 * margin,
          height: visible.height + 2 * margin,
        })
      } catch (error) {
        console.error(error)
        return
      }
      const coverage = start.selection ? sampleSelection(start.selection, visible, level) : null
      const mask = marginMask(src.width, src.height, margin, coverage, visible.width, visible.height)
      setCanvasBusy(true)
      context.imaging.run('filter', { src, spec: scaled, mask }, { signal: controller.signal })
        .then((filtered) => {
          if (controller.signal.aborted) return
          const pixels = margin > 0 ? cropBuffer(filtered, margin, margin, visible.width, visible.height) : filtered
          context.compositor.setPreview({ kind: 'layer-pixels', layerId: layer.id, level, rect: visible, pixels })
          setCanvasBusy(false)
        })
        .catch((error: unknown) => {
          if (!isAbortError(error)) {
            console.error(error)
            setCanvasBusy(false)
          }
        })
    }, 160)
    return () => {
      clearTimeout(timer)
      controller.abort()
    }
  }, [canvasPreview, spec, previewRound])

  useEffect(() => () => {
    if (canvasPreviewPossible) context.compositor.setPreview(null)
    const canvas = canvasRef.current
    if (canvas) {
      canvas.width = 1
      canvas.height = 1
    }
  }, [])

  // ----- interaction -------------------------------------------------------------------------------
  const panPreview = (event: ReactPointerEvent<HTMLCanvasElement>) => {
    if (event.button !== 0 || working) return
    event.preventDefault()
    const element = event.currentTarget
    const startX = event.clientX
    const startY = event.clientY
    const origin = center
    let moved = false
    try {
      element.setPointerCapture(event.pointerId)
    } catch {
      // ignore
    }
    setComparing(true)
    const move = (moveEvent: PointerEvent) => {
      const dx = (moveEvent.clientX - startX) * scale
      const dy = (moveEvent.clientY - startY) * scale
      if (!moved && Math.hypot(dx, dy) < 3) return
      moved = true
      setCenter({ x: Math.round(origin.x - dx), y: Math.round(origin.y - dy) })
    }
    const end = () => {
      element.removeEventListener('pointermove', move)
      element.removeEventListener('pointerup', end)
      element.removeEventListener('pointercancel', end)
      setComparing(false)
    }
    element.addEventListener('pointermove', move)
    element.addEventListener('pointerup', end)
    element.addEventListener('pointercancel', end)
  }

  const apply = async () => {
    if (working) return
    setWorking(true)
    let ok = false
    try {
      ok = await applyFilterToLayer(context.commands, specRef.current)
    } finally {
      // The filtered layer replaces the proxy preview before the next frame.
      if (canvasPreviewPossible) context.compositor.setPreview(null)
      setWorking(false)
    }
    if (ok) {
      close()
      context.focusCanvas()
    } else {
      setPreviewRound((round) => round + 1)
    }
  }

  const update = (key: string, value: unknown) => setSpec({ ...specRef.current, [key]: value } as FilterSpec)
  const values = spec as unknown as Record<string, unknown>
  const footer = (
    <CheckField
      label="Preview"
      title={canvasPreviewPossible ? 'Show the result on the canvas (at the current zoom)' : 'There is no canvas preview while a layer mask is the target'}
      checked={canvasPreviewPossible && canvasPreview}
      disabled={!canvasPreviewPossible || working}
      onChange={setCanvasPreview}
    />
  )
  const busy = boxBusy || canvasBusy
  return (
    <DialogFrame
      id={`filter-${type}`}
      title={filterLabel(type)}
      width={300}
      busy={working}
      footerStart={footer}
      onOk={() => void apply()}
      onCancel={() => { close(); context.focusCanvas() }}
      onReset={() => setSpec({ ...defaultFilter(type), ...(type === 'add-noise' ? { seed: (spec as Extract<FilterSpec, { type: 'add-noise' }>).seed } : {}) } as FilterSpec)}
    >
      <div className="ae-filter-preview">
        <canvas
          ref={canvasRef}
          width={boxWidth}
          height={boxHeight}
          style={{ width: BOX_WIDTH, height: BOX_HEIGHT }}
          role="img"
          aria-label={`${filterLabel(type)} preview at 100%. Drag to look elsewhere; hold to see the original.`}
          title="100% preview: drag to look at another spot, hold the mouse button to compare with the original"
          onPointerDown={panPreview}
        />
        <div className="ae-filter-preview-info">
          <span>100%{comparing ? ' · original' : ''}</span>
          <span>{busy ? <><LoaderCircle className="ae-spin" aria-hidden="true" /> Updating…</> : `x ${Math.round(center.x)}, y ${Math.round(center.y)}`}</span>
        </div>
      </div>
      <div className="ae-adjust">
        {fields.length === 0 && <p className="ae-note">This filter has no settings.</p>}
        {fields.map((field) => {
          if (field.kind === 'number') {
            return (
              <SliderField
                key={field.key}
                label={field.label}
                value={Number(values[field.key])}
                min={field.min}
                max={field.max}
                step={field.step ?? 1}
                digits={field.digits ?? 0}
                unit={field.unit}
                disabled={working}
                onChange={(value) => update(field.key, value)}
              />
            )
          }
          if (field.kind === 'select') {
            return <SelectField key={field.key} label={field.label} value={String(values[field.key])} options={field.options} disabled={working} onChange={(value) => update(field.key, value)} />
          }
          return <CheckField key={field.key} label={field.label} checked={Boolean(values[field.key])} disabled={working} onChange={(value) => update(field.key, value)} />
        })}
      </div>
      {maskTarget && <p className="ae-dialog-note">This filters the layer mask of “{layer?.name}”.</p>}
      {start.selection && <p className="ae-dialog-note">Only the selected area changes.</p>}
    </DialogFrame>
  )
}

export const dialog: DialogDefinition = {
  handles: (request) => request.kind === 'filter',
  component: FilterDialog,
}
