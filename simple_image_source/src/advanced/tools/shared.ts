// src/advanced/tools/shared.ts (WP5)
// Infrastructure shared by the Advanced tool controllers (design 5.9):
//   - job tracking for asynchronous commits (worker warps, healing, flood fills), so the editor's settle()
//     can wait for them and block input meanwhile (ToolControllerExtras);
//   - document helpers: the active layer, its pixel source, the paint target (layer pixels or its mask)
//     with Photoshop's refusals (locked, hidden, adjustment, vector layers);
//   - paint compositing on straight RGBA and on masks (every blend mode, transparency lock, eraser,
//     selection coverage), and the stroke engine (dab placement on a sparse per-stroke coverage grid,
//     recomputed from the pre-stroke pixels so opacity caps a stroke exactly as in Photoshop);
//   - floating pixels (move / transform with a selection): lift, hole, place and merge in one undoable
//     step, keeping the original pixels so repeated moves never eat what lies underneath;
//   - live previews at the display level (compositor 'layer-pixels') aligned to the pyramid grid;
//   - overlay drawing primitives and cursors (screen CSS px).
// No DOM access at module level; Node tests load the tools through this module.
import type { BlendMode, BrushTip, IntRect, MaskBuffer, PixelBuffer, Point, Rect, Rgb8, SelectionOp, StrokeSample } from '../../imaging/types.ts'
import type {
  CompositePreview,
  DocumentState,
  SampleSource,
  EditTarget,
  HistoryIcon,
  Layer,
  LayerId,
  LayerMask,
  PixelEditor,
  RasterLayer,
  Selection,
  StrokeSession,
  TiledSurface,
  ToolContext,
  ToolController,
  ToolOptions,
  ToolPointerEvent,
  ViewTransform,
} from '../types.ts'
import { TILE_SIZE } from '../types.ts'
import { pixelEditBlocker } from '../document.ts'
import { compositeRect } from '../composite.ts'
import { applySelectionOp, selectionCoverage } from '../selection.ts'
import { accumulateDab, compositeStroke, placeDabs } from '../../imaging/brush.ts'
import type { BrushDynamics, DabCarry } from '../../imaging/brush.ts'
import { blendInto } from '../../imaging/blend.ts'
import { cropMask } from '../../imaging/mask.ts'
import { downsampleBuffer2x2, levelOffset, readSurfaceLevel } from '../pyramid.ts'
import { tileKey } from '../tiles.ts'
import { visibleDocRect } from '../viewport.ts'

// ---------------------------------------------------------------------------------------------
// Optional controller members (src/advanced/types.ts stays frozen; the editor feature-checks them)
// ---------------------------------------------------------------------------------------------

/** Optional members a tool controller may implement on top of ToolController. */
export interface ToolControllerExtras {
  /** An asynchronous commit (worker job) is running: block document commands and tool switches until whenIdle(). */
  isBusy?(): boolean
  /** Resolves once no asynchronous work of this tool is pending (the editor's settle() awaits it). */
  whenIdle?(): Promise<void>
}

export type AdvancedToolController = ToolController & ToolControllerExtras

export interface JobTracker {
  readonly busy: boolean
  /** Runs `work` as a tracked job; errors propagate to the caller. */
  run<T>(work: () => Promise<T>): Promise<T>
  whenIdle(): Promise<void>
}

export function createJobTracker(onChange?: () => void): JobTracker {
  let running = 0
  let waiters: (() => void)[] = []
  const done = () => {
    running -= 1
    if (running === 0) {
      const list = waiters
      waiters = []
      for (const resolve of list) resolve()
    }
    onChange?.()
  }
  return {
    get busy() {
      return running > 0
    },
    async run<T>(work: () => Promise<T>): Promise<T> {
      running += 1
      onChange?.()
      try {
        return await work()
      } finally {
        done()
      }
    },
    whenIdle(): Promise<void> {
      if (running === 0) return Promise.resolve()
      return new Promise((resolve) => { waiters.push(resolve) })
    },
  }
}

// ---------------------------------------------------------------------------------------------
// Messages
// ---------------------------------------------------------------------------------------------

export function isAbortError(error: unknown): boolean {
  return Boolean(error) && typeof error === 'object' && (error as { name?: unknown }).name === 'AbortError'
}

export function messageOf(error: unknown, fallback: string): string {
  if (error instanceof Error && error.message) return error.message
  if (typeof error === 'string' && error) return error
  return fallback
}

/** Shows an error in the host toast (never for cancellations). */
export function reportError(ctx: ToolContext, error: unknown, fallback: string): void {
  if (isAbortError(error)) return
  ctx.host.notify(messageOf(error, fallback), 'error')
}

// ---------------------------------------------------------------------------------------------
// Rectangles
// ---------------------------------------------------------------------------------------------

export function intersect(a: IntRect, b: IntRect): IntRect | null {
  const x0 = Math.max(a.x, b.x)
  const y0 = Math.max(a.y, b.y)
  const x1 = Math.min(a.x + a.width, b.x + b.width)
  const y1 = Math.min(a.y + a.height, b.y + b.height)
  return x1 > x0 && y1 > y0 ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null
}

/** Bounding box; null and empty rectangles do not contribute. */
export function union(a: IntRect | null | undefined, b: IntRect | null | undefined): IntRect | null {
  const aOk = Boolean(a && a.width > 0 && a.height > 0)
  const bOk = Boolean(b && b.width > 0 && b.height > 0)
  if (!aOk && !bOk) return null
  if (!aOk) return { ...(b as IntRect) }
  if (!bOk) return { ...(a as IntRect) }
  const r = a as IntRect
  const s = b as IntRect
  const x = Math.min(r.x, s.x)
  const y = Math.min(r.y, s.y)
  return { x, y, width: Math.max(r.x + r.width, s.x + s.width) - x, height: Math.max(r.y + r.height, s.y + s.height) - y }
}

export function contains(outer: IntRect, inner: IntRect): boolean {
  return inner.x >= outer.x && inner.y >= outer.y && inner.x + inner.width <= outer.x + outer.width && inner.y + inner.height <= outer.y + outer.height
}

/** Smallest integer rectangle containing a float rectangle. */
export function roundOut(rect: Rect): IntRect {
  const x0 = Math.floor(rect.x)
  const y0 = Math.floor(rect.y)
  return { x: x0, y: y0, width: Math.max(0, Math.ceil(rect.x + rect.width) - x0), height: Math.max(0, Math.ceil(rect.y + rect.height) - y0) }
}

export function translateRect(rect: IntRect, dx: number, dy: number): IntRect {
  return { x: rect.x + dx, y: rect.y + dy, width: rect.width, height: rect.height }
}

export function canvasRect(state: Pick<DocumentState, 'width' | 'height'>): IntRect {
  return { x: 0, y: 0, width: state.width, height: state.height }
}

// ---------------------------------------------------------------------------------------------
// Buffers
// ---------------------------------------------------------------------------------------------

export function emptyPixels(width: number, height: number): PixelBuffer {
  return { width, height, data: new Uint8ClampedArray(Math.max(0, width * height * 4)) }
}

/** Copies `src` (positioned at srcRect) into `dst` (positioned at dstRect) where they overlap. */
export function blitCopy(dst: PixelBuffer, dstRect: IntRect, src: PixelBuffer, srcRect: IntRect): void {
  const overlap = intersect(dstRect, srcRect)
  if (!overlap) return
  const span = overlap.width * 4
  for (let y = overlap.y; y < overlap.y + overlap.height; y += 1) {
    const from = ((y - srcRect.y) * src.width + (overlap.x - srcRect.x)) * 4
    dst.data.set(src.data.subarray(from, from + span), ((y - dstRect.y) * dst.width + (overlap.x - dstRect.x)) * 4)
  }
}

/** Source-over of straight-alpha `src` (at srcRect) onto `dst` (at dstRect), in place. */
export function blitOver(dst: PixelBuffer, dstRect: IntRect, src: PixelBuffer, srcRect: IntRect): void {
  const overlap = intersect(dstRect, srcRect)
  if (!overlap) return
  const d = dst.data
  const s = src.data
  for (let y = overlap.y; y < overlap.y + overlap.height; y += 1) {
    let si = ((y - srcRect.y) * src.width + (overlap.x - srcRect.x)) * 4
    let di = ((y - dstRect.y) * dst.width + (overlap.x - dstRect.x)) * 4
    for (let x = 0; x < overlap.width; x += 1, si += 4, di += 4) {
      const sa = s[si + 3]
      if (sa === 0) continue
      const da = d[di + 3]
      if (sa === 255 || da === 0) {
        d[di] = s[si]
        d[di + 1] = s[si + 1]
        d[di + 2] = s[si + 2]
        d[di + 3] = sa
        continue
      }
      const as = sa / 255
      const ab = da / 255
      const ao = as + ab * (1 - as)
      const ws = as / ao
      const wb = (ab * (1 - as)) / ao
      d[di] = s[si] * ws + d[di] * wb
      d[di + 1] = s[si + 1] * ws + d[di + 1] * wb
      d[di + 2] = s[si + 2] * ws + d[di + 2] * wb
      d[di + 3] = ao * 255
    }
  }
}

export function hasVisiblePixels(buffer: PixelBuffer): boolean {
  const data = buffer.data
  for (let p = 3; p < data.length; p += 4) if (data[p] !== 0) return true
  return false
}

// ---------------------------------------------------------------------------------------------
// Document helpers
// ---------------------------------------------------------------------------------------------

export function findLayer(state: Pick<DocumentState, 'layers'>, id: LayerId | null | undefined): Layer | null {
  if (id === null || id === undefined) return null
  return state.layers.find((layer) => layer.id === id) ?? null
}

export function layerIndexOf(state: Pick<DocumentState, 'layers'>, id: LayerId | null | undefined): number {
  if (id === null || id === undefined) return -1
  return state.layers.findIndex((layer) => layer.id === id)
}

export function activeLayerOf(state: DocumentState): Layer | null {
  return findLayer(state, state.activeLayerId)
}

export function isBackgroundLayer(layer: Layer | null | undefined): boolean {
  return Boolean(layer && layer.kind === 'raster' && layer.isBackground)
}

export interface PixelSource {
  readonly surface: TiledSurface
  readonly offsetX: number
  readonly offsetY: number
}

/** Where a layer's pixels live: the raster surface or the text/shape raster cache (null for adjustments). */
export function pixelSourceOf(layer: Layer): PixelSource | null {
  if (layer.kind === 'raster') return { surface: layer.surface, offsetX: layer.offsetX, offsetY: layer.offsetY }
  if (layer.kind === 'text' || layer.kind === 'shape') return { surface: layer.raster.surface, offsetX: layer.raster.offsetX, offsetY: layer.raster.offsetY }
  return null
}

/** Layer pixels over a document rectangle (transparent where the layer has none). */
export function readLayerRect(layer: Layer, rect: IntRect): PixelBuffer {
  const source = pixelSourceOf(layer)
  if (!source) return emptyPixels(rect.width, rect.height)
  return source.surface.read({ x: rect.x - source.offsetX, y: rect.y - source.offsetY, width: rect.width, height: rect.height })
}

/** Mask values over a document rectangle (the mask's default value where nothing is painted). */
export function readMaskRect(mask: LayerMask, rect: IntRect): MaskBuffer {
  return mask.surface.read({ x: rect.x - mask.offsetX, y: rect.y - mask.offsetY, width: rect.width, height: rect.height })
}

/** Document-space bounds of the pixels a layer shows (null when empty or an adjustment layer). */
export function layerPixelBounds(layer: Layer): IntRect | null {
  const source = pixelSourceOf(layer)
  if (!source) return null
  const bounds = source.surface.contentBounds()
  return bounds ? translateRect(bounds, source.offsetX, source.offsetY) : null
}

/** Alpha of a layer at a document point (its mask included when enabled). */
export function layerAlphaAt(layer: Layer, x: number, y: number): number {
  const px = Math.floor(x)
  const py = Math.floor(y)
  const source = pixelSourceOf(layer)
  if (!source) return 0
  const pixel = source.surface.read({ x: px - source.offsetX, y: py - source.offsetY, width: 1, height: 1 })
  let alpha = pixel.data[3]
  if (alpha && layer.mask && layer.mask.enabled) {
    const value = layer.mask.surface.read({ x: px - layer.mask.offsetX, y: py - layer.mask.offsetY, width: 1, height: 1 }).data[0]
    alpha = Math.round((alpha * value) / 255)
  }
  return alpha
}

/** Topmost visible layer with a visible pixel at the point (Move tool auto-select), or null. */
export function topmostLayerAt(state: DocumentState, point: Point): Layer | null {
  if (point.x < 0 || point.y < 0 || point.x >= state.width || point.y >= state.height) return null
  for (let index = state.layers.length - 1; index >= 0; index -= 1) {
    const layer = state.layers[index]
    if (!layer.visible || layer.opacity <= 0 || layer.kind === 'adjustment') continue
    if (layerAlphaAt(layer, point.x, point.y) > 0) return layer
  }
  return null
}

/**
 * Document-size pixels a sampling tool reads (Magic Wand, Paint Bucket): the active layer's own pixels
 * ('current'), the composite up to and including it ('current-below'), or the whole composite ('all').
 * Adjustment layers have no pixels of their own, so 'current' falls back to the composite below them.
 */
export async function sampleDocumentPixels(ctx: ToolContext, state: DocumentState, sample: SampleSource): Promise<PixelBuffer> {
  const rect = canvasRect(state)
  const active = activeLayerOf(state)
  if (sample === 'current' && active && active.kind !== 'adjustment') return readLayerRect(active, rect)
  if (sample !== 'all' && active) return compositeRect(state, rect, { belowLayerId: active.id, includeBelowLayer: true })
  return ctx.compositor.flatten()
}

// ---------------------------------------------------------------------------------------------
// Selection helpers
// ---------------------------------------------------------------------------------------------

const SELECTION_STRIDE = 1 << 16
let selectionCounter = 1 << 30

/**
 * A selection version above every version the store may have handed out: the store raises its own counter
 * to the versions it is given and increments it for undo/redo restores, so a wide stride keeps versions of
 * different selections distinct (outline caches key on them).
 */
export function nextSelectionVersion(state: Pick<DocumentState, 'selection'>): number {
  selectionCounter = Math.max(selectionCounter, state.selection?.version ?? 0) + SELECTION_STRIDE
  return selectionCounter
}

/** Selection coverage (0..255) at a document point; 255 everywhere without a selection. */
export function selectionValueAt(selection: Selection | null, x: number, y: number): number {
  if (!selection) return 255
  const px = Math.floor(x)
  const py = Math.floor(y)
  const mask = selection.mask
  if (px < 0 || py < 0 || px >= mask.width || py >= mask.height) return 0
  return mask.data[py * mask.width + px]
}

/**
 * Combines a document-size shape mask with the current selection and records it (one history step that
 * does not change the document's content revision). Returns the new selection.
 */
export function commitSelectionShape(ctx: ToolContext, label: string, icon: HistoryIcon, shape: MaskBuffer, op: SelectionOp): Selection | null {
  const state = ctx.store.getState()
  if (shape.width !== state.width || shape.height !== state.height) return state.selection
  const next = applySelectionOp(state.selection, shape, op, nextSelectionVersion(state))
  if (next === state.selection) return next
  ctx.store.transact(label, icon, (tx) => tx.setSelection(next), { affectsOutput: false })
  return next
}

/** Deselect (one history step, no content revision). */
export function deselect(ctx: ToolContext): void {
  if (!ctx.store.getState().selection) return
  ctx.store.transact('Deselect', 'selection', (tx) => tx.setSelection(null), { affectsOutput: false })
}

// ---------------------------------------------------------------------------------------------
// Paint targets
// ---------------------------------------------------------------------------------------------

export interface PaintTarget {
  readonly layer: Layer
  readonly target: EditTarget
  /** Document position of the edited surface's (0, 0). */
  readonly origin: Point
  readonly channels: 1 | 4
  /** Transparency lock (alpha is kept). */
  readonly preserveAlpha: boolean
  /** Writes are clipped to this document rectangle: the canvas (Photoshop paints only what is visible). */
  readonly clip: IntRect | null
}

/**
 * The layer pixels (or the active layer mask) painting tools edit, or null after telling the user why
 * nothing can be painted (no layer, adjustment / vector layer, locked or hidden layer).
 */
export function resolvePaintTarget(ctx: ToolContext, options: { readonly allowMask?: boolean } = {}): PaintTarget | null {
  const state = ctx.store.getState()
  const layer = activeLayerOf(state)
  let target: EditTarget = state.editTarget === 'mask' && layer?.mask ? 'mask' : 'pixels'
  if (target === 'mask' && options.allowMask === false) target = 'pixels'
  const blocker = pixelEditBlocker(layer, target)
  if (blocker || !layer) {
    ctx.host.notify(blocker ?? 'Select a layer first.', 'error')
    return null
  }
  if (target === 'mask' && layer.mask) {
    return { layer, target, origin: { x: layer.mask.offsetX, y: layer.mask.offsetY }, channels: 1, preserveAlpha: false, clip: canvasRect(state) }
  }
  if (layer.kind !== 'raster') {
    ctx.host.notify(`"${layer.name}" cannot be painted on.`, 'error')
    return null
  }
  return {
    layer,
    target: 'pixels',
    origin: { x: layer.offsetX, y: layer.offsetY },
    channels: 4,
    preserveAlpha: layer.locks.transparency || layer.isBackground,
    clip: canvasRect(state),
  }
}

export function toLocalRect(target: Pick<PaintTarget, 'origin'>, rect: IntRect): IntRect {
  return { x: rect.x - target.origin.x, y: rect.y - target.origin.y, width: rect.width, height: rect.height }
}

/** Writes a document-rect region through an editor (pixels or mask). */
export function writeRegion(editor: PixelEditor, target: Pick<PaintTarget, 'origin' | 'channels'>, rect: IntRect, data: PixelBuffer | MaskBuffer): void {
  const local = toLocalRect(target, rect)
  if (target.channels === 4) editor.writePixels(local.x, local.y, data as PixelBuffer)
  else editor.writeMask(local.x, local.y, data as MaskBuffer)
}

// ---------------------------------------------------------------------------------------------
// Colours and paint compositing
// ---------------------------------------------------------------------------------------------

/** Grey value a colour paints into a layer mask (Rec. 601 luma, as Photoshop converts to Gray). */
export function grayOf(color: Rgb8): number {
  return Math.max(0, Math.min(255, Math.round(0.299 * color.r + 0.587 * color.g + 0.114 * color.b)))
}

function clamp01(value: number): number {
  return value <= 0 ? 0 : value >= 1 ? 1 : value
}

export interface PaintOptions {
  readonly color: Rgb8
  /** 0..1, the most one stroke can cover. */
  readonly opacity: number
  readonly mode: BlendMode
  readonly erase?: boolean
  readonly preserveAlpha?: boolean
  /** Selection coverage of the region (0..255), or null. */
  readonly selection: MaskBuffer | null
  /** Dissolve pattern seed. */
  readonly seed?: number
}

/**
 * Paints a solid colour onto the pre-stroke pixels through the stroke coverage (0..1 per pixel):
 * source alpha = coverage * opacity * selection, in the given blend mode; the eraser lowers alpha; the
 * transparency lock keeps alpha (colour mixes only where pixels exist). Returns a new buffer.
 */
export function paintPixels(before: PixelBuffer, coverage: Float32Array, options: PaintOptions): PixelBuffer {
  if (options.erase || options.mode === 'normal') {
    return compositeStroke(before, coverage, {
      color: options.color,
      opacity: options.opacity,
      erase: Boolean(options.erase),
      preserveAlpha: Boolean(options.preserveAlpha),
    }, options.selection)
  }
  const n = before.width * before.height
  const src = new Uint8ClampedArray(n * 4)
  const { r, g, b } = options.color
  for (let i = 0, p = 0; i < n; i += 1, p += 4) {
    src[p] = r
    src[p + 1] = g
    src[p + 2] = b
    src[p + 3] = Math.round(clamp01(coverage[i]) * 255)
  }
  return blendPixels(before, { width: before.width, height: before.height, data: src }, {
    opacity: options.opacity,
    mode: options.mode,
    coverage: options.selection ? options.selection.data : null,
    preserveAlpha: Boolean(options.preserveAlpha),
    seed: options.seed,
  })
}

export interface BlendOptions {
  readonly opacity: number
  readonly mode: BlendMode
  /** Extra per-pixel coverage (0..255), or null. */
  readonly coverage: Uint8Array | null
  readonly preserveAlpha: boolean
  readonly seed?: number
}

/**
 * Composites `source` (same size, straight RGBA) onto `before` with a blend mode, opacity and coverage.
 * With preserveAlpha the colour blends as over an opaque backdrop and the original alpha stays
 * (Photoshop's transparency lock). Returns a new buffer.
 */
export function blendPixels(before: PixelBuffer, source: PixelBuffer, options: BlendOptions): PixelBuffer {
  const n = before.width * before.height
  const out = new Uint8ClampedArray(before.data)
  if (!(options.opacity > 0) || n === 0) return { width: before.width, height: before.height, data: out }
  if (!options.preserveAlpha) {
    blendInto(out, 0, source.data, 0, n, options.mode, clamp01(options.opacity), options.coverage, 0, options.seed ?? 0)
    return { width: before.width, height: before.height, data: out }
  }
  for (let p = 3; p < out.length; p += 4) out[p] = 255
  blendInto(out, 0, source.data, 0, n, options.mode, clamp01(options.opacity), options.coverage, 0, options.seed ?? 0)
  const original = before.data
  for (let p = 0; p < out.length; p += 4) {
    const alpha = original[p + 3]
    if (alpha === 0) {
      out[p] = original[p]
      out[p + 1] = original[p + 1]
      out[p + 2] = original[p + 2]
    }
    out[p + 3] = alpha
  }
  return { width: before.width, height: before.height, data: out }
}

/** Paints a grey value into mask values: m' = m + (value - m) * coverage * opacity * selection. */
export function paintMask(before: MaskBuffer, coverage: Float32Array, value: number, opacity: number, selection: MaskBuffer | null): MaskBuffer {
  const out = new Uint8Array(before.data)
  const amount = clamp01(opacity)
  const sel = selection ? selection.data : null
  const v = Math.max(0, Math.min(255, value))
  for (let i = 0; i < out.length; i += 1) {
    let a = clamp01(coverage[i]) * amount
    if (sel) a *= sel[i] / 255
    if (a <= 0) continue
    out[i] = Math.round(out[i] + (v - out[i]) * a)
  }
  return { width: before.width, height: before.height, data: out }
}

/**
 * Blends per-pixel grey `values` into mask values with per-pixel `alpha` (0..255, e.g. a gradient's
 * opacity or a flood fill), times opacity and optional coverage.
 */
export function blendMaskValues(before: MaskBuffer, values: Uint8Array, alpha: Uint8Array | null, opacity: number, coverage: Uint8Array | null): MaskBuffer {
  const out = new Uint8Array(before.data)
  const amount = clamp01(opacity)
  for (let i = 0; i < out.length; i += 1) {
    let a = amount
    if (alpha) a *= alpha[i] / 255
    if (coverage) a *= coverage[i] / 255
    if (a <= 0) continue
    out[i] = Math.round(out[i] + (values[i] - out[i]) * a)
  }
  return { width: before.width, height: before.height, data: out }
}

/** Grey values (luma) and alpha of an RGBA buffer, for painting a mask with colour content. */
export function grayAndAlpha(source: PixelBuffer): { readonly values: Uint8Array; readonly alpha: Uint8Array } {
  const n = source.width * source.height
  const values = new Uint8Array(n)
  const alpha = new Uint8Array(n)
  const d = source.data
  for (let i = 0, p = 0; i < n; i += 1, p += 4) {
    values[i] = Math.round(0.299 * d[p] + 0.587 * d[p + 1] + 0.114 * d[p + 2])
    alpha[i] = d[p + 3]
  }
  return { values, alpha }
}

// ---------------------------------------------------------------------------------------------
// Stroke engine
// ---------------------------------------------------------------------------------------------

export interface StrokeRegion {
  /** Document rectangle being recomputed. */
  readonly docRect: IntRect
  /** The same rectangle in the edited surface's own coordinates. */
  readonly localRect: IntRect
  /** Pre-stroke content of the region (PixelBuffer for pixels, MaskBuffer for a mask). */
  readonly before: PixelBuffer | MaskBuffer
  /** Accumulated stroke coverage 0..1, region-sized. */
  readonly coverage: Float32Array
  /** Selection coverage of the region (0..255), or null without a selection. */
  readonly selection: MaskBuffer | null
  readonly editor: PixelEditor
}

export interface StrokeEngineOptions {
  readonly label: string
  readonly icon: HistoryIcon
  readonly tip: BrushTip
  /** Fraction of the diameter between dabs. */
  readonly spacing: number
  readonly smoothing: number
  readonly dynamics: BrushDynamics
  /** New content of a region from its pre-stroke content and coverage. */
  readonly paint: (region: StrokeRegion) => PixelBuffer | MaskBuffer
}

export interface StrokeEngine {
  readonly target: PaintTarget
  readonly session: StrokeSession
  /** Feeds pointer samples (document coordinates); `final` catches the smoothed path up with the pointer. */
  add(samples: readonly StrokeSample[], final?: boolean): void
  /** Repaints every touched region (after the paint function's inputs changed). */
  repaintAll(): void
  /** Commits (one history step when anything changed) or cancels the stroke. */
  finish(commit: boolean): void
  /** Last smoothed sample (Shift+click straight lines continue from it). */
  readonly lastPoint: Point | null
  readonly open: boolean
  /** Document bounds of everything the stroke covered so far. */
  readonly bounds: IntRect | null
  /** The coverage over a document rectangle (0..1). */
  coverageIn(rect: IntRect): Float32Array
}

const TILE_AREA = TILE_SIZE * TILE_SIZE

interface CoverageTile {
  readonly tx: number
  readonly ty: number
  readonly data: Float32Array
  dirty: IntRect | null
  touched: IntRect | null
}

/** Starts a stroke session on `target` (one history step on commit). Null after telling the user why not. */
export function beginStrokeEngine(ctx: ToolContext, target: PaintTarget, options: StrokeEngineOptions): StrokeEngine | null {
  let session: StrokeSession
  try {
    session = ctx.store.beginStroke(target.layer.id, target.target, options.label, options.icon)
  } catch (error) {
    reportError(ctx, error, 'This layer cannot be painted on right now.')
    return null
  }
  const state = ctx.store.getState()
  const selection = state.selection
  const tiles = new Map<number, CoverageTile>()
  const carry: DabCarry = { distance: 0, last: null }
  let open = true
  let bounds: IntRect | null = null

  const tileFor = (tx: number, ty: number): CoverageTile => {
    const key = tileKey(tx, ty)
    let tile = tiles.get(key)
    if (!tile) {
      tile = { tx, ty, data: new Float32Array(TILE_AREA), dirty: null, touched: null }
      tiles.set(key, tile)
    }
    return tile
  }

  const extract = (tile: CoverageTile, rect: IntRect): Float32Array => {
    const out = new Float32Array(rect.width * rect.height)
    const left = rect.x - tile.tx * TILE_SIZE
    const top = rect.y - tile.ty * TILE_SIZE
    for (let y = 0; y < rect.height; y += 1) {
      const from = (top + y) * TILE_SIZE + left
      out.set(tile.data.subarray(from, from + rect.width), y * rect.width)
    }
    return out
  }

  const repaint = (tile: CoverageTile, rect: IntRect) => {
    let docRect: IntRect | null = rect
    if (target.clip) docRect = intersect(docRect, target.clip)
    if (!docRect) return
    const localRect = toLocalRect(target, docRect)
    const before = session.editor.readBefore(localRect)
    const coverage = extract(tile, docRect)
    const sel = selection ? selectionCoverage(selection, docRect) : null
    const next = options.paint({ docRect, localRect, before, coverage, selection: sel, editor: session.editor })
    writeRegion(session.editor, target, docRect, next)
  }

  const flush = () => {
    for (const tile of tiles.values()) {
      if (!tile.dirty) continue
      const rect = tile.dirty
      tile.dirty = null
      repaint(tile, rect)
    }
  }

  const engine: StrokeEngine = {
    target,
    session,
    get lastPoint() {
      return carry.last ? { x: carry.last.x, y: carry.last.y } : null
    },
    get open() {
      return open
    },
    get bounds() {
      return bounds
    },
    add(samples: readonly StrokeSample[], final = false): void {
      if (!open || !samples.length) return
      const dabs = placeDabs(samples, options.tip, options.spacing, final ? 0 : options.smoothing, carry, options.dynamics)
      for (const dab of dabs) {
        const half = dab.diameter / 2 + 1
        let reach: IntRect | null = roundOut({ x: dab.x - half, y: dab.y - half, width: 2 * half, height: 2 * half })
        if (target.clip) reach = intersect(reach, target.clip)
        if (!reach) continue
        const tx0 = Math.floor(reach.x / TILE_SIZE)
        const ty0 = Math.floor(reach.y / TILE_SIZE)
        const tx1 = Math.floor((reach.x + reach.width - 1) / TILE_SIZE)
        const ty1 = Math.floor((reach.y + reach.height - 1) / TILE_SIZE)
        for (let ty = ty0; ty <= ty1; ty += 1) {
          for (let tx = tx0; tx <= tx1; tx += 1) {
            const tile = tileFor(tx, ty)
            const touched = accumulateDab(tile.data, TILE_SIZE, TILE_SIZE, tx * TILE_SIZE, ty * TILE_SIZE, dab, options.tip)
            if (touched.width > 0 && touched.height > 0) {
              tile.dirty = union(tile.dirty, touched)
              tile.touched = union(tile.touched, touched)
              bounds = union(bounds, touched)
            }
          }
        }
      }
      flush()
    },
    repaintAll(): void {
      if (!open) return
      for (const tile of tiles.values()) if (tile.touched) repaint(tile, tile.touched)
    },
    finish(commit: boolean): void {
      if (!open) return
      open = false
      try {
        if (commit) flush()
      } finally {
        if (commit) session.commit()
        else session.cancel()
        tiles.clear()
      }
    },
    coverageIn(rect: IntRect): Float32Array {
      const out = new Float32Array(Math.max(0, rect.width * rect.height))
      for (const tile of tiles.values()) {
        const overlap = intersect(rect, { x: tile.tx * TILE_SIZE, y: tile.ty * TILE_SIZE, width: TILE_SIZE, height: TILE_SIZE })
        if (!overlap) continue
        for (let y = overlap.y; y < overlap.y + overlap.height; y += 1) {
          const from = (y - tile.ty * TILE_SIZE) * TILE_SIZE + (overlap.x - tile.tx * TILE_SIZE)
          out.set(tile.data.subarray(from, from + overlap.width), (y - rect.y) * rect.width + (overlap.x - rect.x))
        }
      }
      return out
    },
  }
  return engine
}

/** Stroke samples of a pointer event: its coalesced samples (oldest first), ending at the event itself. */
export function strokeSamples(event: ToolPointerEvent): StrokeSample[] {
  const pressureOf = (value: number) => (event.pointerType === 'pen' ? clamp01(Number.isFinite(value) ? value : 1) : 1)
  const out: StrokeSample[] = []
  for (const sample of event.coalesced ?? []) {
    if (!sample || !Number.isFinite(sample.doc.x) || !Number.isFinite(sample.doc.y)) continue
    out.push({ x: sample.doc.x, y: sample.doc.y, pressure: pressureOf(sample.pressure), time: sample.time })
  }
  const last = out[out.length - 1]
  if (!last || last.x !== event.doc.x || last.y !== event.doc.y) {
    out.push({ x: event.doc.x, y: event.doc.y, pressure: pressureOf(event.pressure), time: event.time })
  }
  return out
}

// ---------------------------------------------------------------------------------------------
// Floating pixels (Move / Free Transform with a selection)
// ---------------------------------------------------------------------------------------------

export interface FloatingPixels {
  readonly layerId: LayerId
  readonly surface: TiledSurface
  /** Document position of the surface's (0, 0) at lift time. */
  readonly origin: Point
  readonly selection: Selection
  /** Document rectangle the pixels were lifted from (the selection bounds). */
  readonly rect: IntRect
  /** Layer pixels times selection coverage, rect-sized. */
  readonly lifted: PixelBuffer
  /** Selection coverage (0..255), rect-sized. */
  readonly coverage: MaskBuffer
  /** Colour left behind (the Background), or null for transparency. */
  readonly fill: Rgb8 | null
  /** False for Alt-drag copies: the source stays. */
  readonly cut: boolean
  /** Writes are clipped to this rectangle (the Background stays canvas-sized), or null. */
  readonly clip: IntRect | null
  /** Layer pixels before the session started, over a growing rectangle. */
  original: { rect: IntRect; pixels: PixelBuffer }
  /** Where the floating pixels were last merged (document rect), or null before the first merge. */
  lastPlaced: IntRect | null
  /** Surface version and selection the session expects (anything else ends it). */
  expectVersion: number
  expectSelection: Selection | null
}

/** Lifts the selected pixels of a raster layer. Null when the selection covers no visible pixel. */
export function liftFloating(state: DocumentState, layer: RasterLayer, selection: Selection, options: { readonly cut: boolean; readonly fill: Rgb8 | null }): FloatingPixels | null {
  const clip = layer.isBackground ? canvasRect(state) : null
  const rect = clip ? intersect(selection.bounds, clip) : { ...selection.bounds }
  if (!rect) return null
  const pixels = readLayerRect(layer, rect)
  const coverage = cropMask(selection.mask, rect)
  const lifted = new Uint8ClampedArray(pixels.data)
  let any = false
  for (let i = 0, p = 3; i < coverage.data.length; i += 1, p += 4) {
    const c = coverage.data[i]
    const value = c === 255 ? lifted[p] : Math.round((lifted[p] * c) / 255)
    lifted[p] = value
    if (value) any = true
    else {
      lifted[p - 3] = 0
      lifted[p - 2] = 0
      lifted[p - 1] = 0
    }
  }
  if (!any) return null
  return {
    layerId: layer.id,
    surface: layer.surface,
    origin: { x: layer.offsetX, y: layer.offsetY },
    selection,
    rect,
    lifted: { width: rect.width, height: rect.height, data: lifted },
    coverage,
    fill: options.fill,
    cut: options.cut,
    clip,
    original: { rect, pixels },
    lastPlaced: null,
    expectVersion: layer.surface.version,
    expectSelection: selection,
  }
}

/** True while the document still shows what the floating session last wrote. */
export function floatingIsCurrent(floating: FloatingPixels, state: DocumentState): boolean {
  const layer = findLayer(state, floating.layerId)
  return Boolean(layer && layer.kind === 'raster' && layer.surface === floating.surface
    && layer.offsetX === floating.origin.x && layer.offsetY === floating.origin.y
    && floating.surface.version === floating.expectVersion && state.selection === floating.expectSelection)
}

function growOriginal(floating: FloatingPixels, rect: IntRect): void {
  if (contains(floating.original.rect, rect)) return
  const grown = union(floating.original.rect, rect) as IntRect
  const fresh = floating.surface.read({ x: grown.x - floating.origin.x, y: grown.y - floating.origin.y, width: grown.width, height: grown.height })
  blitCopy(fresh, grown, floating.original.pixels, floating.original.rect)
  floating.original = { rect: grown, pixels: fresh }
}

/** Original pixels with the hole applied over `rect` (document space). */
function holedOriginal(floating: FloatingPixels, rect: IntRect): PixelBuffer {
  growOriginal(floating, rect)
  const out = emptyPixels(rect.width, rect.height)
  blitCopy(out, rect, floating.original.pixels, floating.original.rect)
  if (!floating.cut) return out
  const hole = intersect(rect, floating.rect)
  if (!hole) return out
  const d = out.data
  const c = floating.coverage.data
  const fill = floating.fill
  for (let y = hole.y; y < hole.y + hole.height; y += 1) {
    let ci = (y - floating.rect.y) * floating.rect.width + (hole.x - floating.rect.x)
    let p = ((y - rect.y) * rect.width + (hole.x - rect.x)) * 4
    for (let x = 0; x < hole.width; x += 1, ci += 1, p += 4) {
      const cover = c[ci]
      if (!cover) continue
      if (fill) {
        const k = cover / 255
        d[p] = Math.round(d[p] + (fill.r - d[p]) * k)
        d[p + 1] = Math.round(d[p + 1] + (fill.g - d[p + 1]) * k)
        d[p + 2] = Math.round(d[p + 2] + (fill.b - d[p + 2]) * k)
      } else {
        d[p + 3] = cover === 255 ? 0 : Math.round((d[p + 3] * (255 - cover)) / 255)
      }
    }
  }
  return out
}

export interface PlacedPixels {
  /** Document rectangle of `pixels`. */
  readonly rect: IntRect
  readonly pixels: PixelBuffer
}

/**
 * The layer content to write for the floating pixels placed at `placed` (null = only the hole): original
 * pixels, the hole (unless copying) and the placed pixels on top, over every area the session has touched
 * or now touches. Null when nothing needs writing.
 */
export function composeFloating(floating: FloatingPixels, placed: PlacedPixels | null): PlacedPixels | null {
  let affected = union(floating.cut ? floating.rect : null, floating.lastPlaced)
  affected = union(affected, placed ? placed.rect : null)
  if (affected && floating.clip) affected = intersect(affected, floating.clip)
  if (!affected) return null
  const out = holedOriginal(floating, affected)
  if (placed) blitOver(out, affected, placed.pixels, placed.rect)
  return { rect: affected, pixels: out }
}

/** Records what a floating merge wrote (call after the transaction). */
export function floatingMerged(floating: FloatingPixels, placed: IntRect | null, selection: Selection | null): void {
  floating.lastPlaced = placed && floating.clip ? intersect(placed, floating.clip) : placed
  floating.expectVersion = floating.surface.version
  floating.expectSelection = selection
}

/** The lifted pixels shifted by whole pixels. */
export function placedAt(floating: FloatingPixels, dx: number, dy: number): PlacedPixels {
  return { rect: translateRect(floating.rect, Math.round(dx), Math.round(dy)), pixels: floating.lifted }
}

// ---------------------------------------------------------------------------------------------
// Display-level previews (compositor 'layer-pixels')
// ---------------------------------------------------------------------------------------------

/** Floor division by 2^level that works for negative numbers. */
export function levelFloor(value: number, level: number): number {
  return Math.floor(value / 2 ** level)
}

/** The visible part of the canvas in level-`level` document pixels (one pixel of margin), or null. */
export function visibleLevelRect(ctx: ToolContext, state: Pick<DocumentState, 'width' | 'height'>, level: number): IntRect | null {
  const view = ctx.view.getView()
  const size = ctx.view.getViewportSize()
  const visible = size.width > 0 && size.height > 0 ? visibleDocRect(view, size) : { x: 0, y: 0, width: state.width, height: state.height }
  const scale = 2 ** level
  const x0 = Math.max(0, Math.floor(visible.x / scale) - 1)
  const y0 = Math.max(0, Math.floor(visible.y / scale) - 1)
  const x1 = Math.min(Math.ceil(state.width / scale), Math.ceil((visible.x + visible.width) / scale) + 1)
  const y1 = Math.min(Math.ceil(state.height / scale), Math.ceil((visible.y + visible.height) / scale) + 1)
  return x1 > x0 && y1 > y0 ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null
}

/** A document rectangle in level pixels (rounded outwards). */
export function toLevelRect(rect: IntRect, level: number): IntRect {
  if (level <= 0) return { ...rect }
  const scale = 2 ** level
  const x0 = Math.floor(rect.x / scale)
  const y0 = Math.floor(rect.y / scale)
  return { x: x0, y: y0, width: Math.ceil((rect.x + rect.width) / scale) - x0 + 1, height: Math.ceil((rect.y + rect.height) / scale) - y0 + 1 }
}

/**
 * Pixels of a layer-local region at pyramid level `level`, computed exactly like the pyramid: the region is
 * widened to multiples of 2^level, read at level 0 and halved `level` times with the alpha-weighted 2x2
 * filter. Returns the buffer and its rectangle in layer-level coordinates.
 */
export function levelPixels(read: (localRect: IntRect) => PixelBuffer, localRect: IntRect, level: number): { readonly pixels: PixelBuffer; readonly rect: IntRect } {
  if (level <= 0) return { pixels: read(localRect), rect: { ...localRect } }
  const scale = 2 ** level
  const x0 = Math.floor(localRect.x / scale) * scale
  const y0 = Math.floor(localRect.y / scale) * scale
  const x1 = Math.ceil((localRect.x + localRect.width) / scale) * scale
  const y1 = Math.ceil((localRect.y + localRect.height) / scale) * scale
  let pixels = read({ x: x0, y: y0, width: x1 - x0, height: y1 - y0 })
  for (let k = 0; k < level; k += 1) pixels = downsampleBuffer2x2(pixels)
  return { pixels, rect: { x: x0 / scale, y: y0 / scale, width: pixels.width, height: pixels.height } }
}

/** Display-level preview of a layer whose pixels moved by (dx, dy), optionally with the unmoved copy kept underneath. */
export function movedLayerPreview(layerId: LayerId, source: PixelSource, dx: number, dy: number, level: number, area: IntRect, keepOriginal: boolean): CompositePreview {
  const pixels = emptyPixels(area.width, area.height)
  const read = (ox: number, oy: number) => readSurfaceLevel(source.surface, level, {
    x: area.x - levelOffset(ox, level),
    y: area.y - levelOffset(oy, level),
    width: area.width,
    height: area.height,
  })
  if (keepOriginal) {
    blitCopy(pixels, area, read(source.offsetX, source.offsetY), area)
    blitOver(pixels, area, read(source.offsetX + dx, source.offsetY + dy), area)
  } else {
    blitCopy(pixels, area, read(source.offsetX + dx, source.offsetY + dy), area)
  }
  return { kind: 'layer-pixels', layerId, level, rect: area, pixels }
}

/**
 * Display-level previews of a floating session: the layer with the hole and the floating pixels placed at
 * a document rectangle (or only the hole). Level buffers are cached per level.
 */
export interface FloatingPreviewer {
  /**
   * Preview with `placed` (null = hole only), or null when nothing is visible. `clipToView` limits the work
   * to the visible part of the canvas (drags); without it the whole affected area is prepared (long sessions
   * where the view may pan).
   */
  preview(ctx: ToolContext, state: DocumentState, level: number, placed: { readonly rect: IntRect; readonly pixels: PixelBuffer; readonly key: unknown } | null,
    clipToView?: boolean): CompositePreview | null
}

export function createFloatingPreviewer(floating: FloatingPixels): FloatingPreviewer {
  const bases = new Map<number, { pixels: PixelBuffer; rect: IntRect } | null>()
  const placedLevels = new Map<string, { key: unknown; pixels: PixelBuffer; rect: IntRect }>()
  const baseRegion = union(floating.cut ? floating.rect : null, floating.lastPlaced)

  const baseAt = (level: number) => {
    if (bases.has(level)) return bases.get(level) ?? null
    let value: { pixels: PixelBuffer; rect: IntRect } | null = null
    if (baseRegion) {
      const local = translateRect(baseRegion, -floating.origin.x, -floating.origin.y)
      const read = (rect: IntRect) => {
        // Live pixels, with the session's region replaced by original + hole.
        const out = floating.surface.read(rect)
        const doc = translateRect(rect, floating.origin.x, floating.origin.y)
        const region = intersect(doc, baseRegion)
        if (region) blitCopy(out, doc, holedOriginal(floating, region), region)
        return out
      }
      value = levelPixels(read, local, level)
    }
    bases.set(level, value)
    return value
  }

  const placedAtLevel = (level: number, placed: { readonly rect: IntRect; readonly pixels: PixelBuffer; readonly key: unknown }) => {
    const cacheKey = String(level)
    const hit = placedLevels.get(cacheKey)
    // The key identifies the pixels; only their position changes while dragging.
    if (hit && hit.key === placed.key) return hit
    const local = { x: 0, y: 0, width: placed.rect.width, height: placed.rect.height }
    const value = levelPixels((rect) => {
      const out = emptyPixels(rect.width, rect.height)
      blitCopy(out, rect, placed.pixels, local)
      return out
    }, local, level)
    const entry = { key: placed.key, pixels: value.pixels, rect: value.rect }
    placedLevels.set(cacheKey, entry)
    return entry
  }

  return {
    preview(ctx, state, level, placed, clipToView = true) {
      const visible = clipToView
        ? visibleLevelRect(ctx, state, level)
        : { x: 0, y: 0, width: Math.ceil(state.width / 2 ** level), height: Math.ceil(state.height / 2 ** level) }
      if (!visible) return null
      const ox = levelOffset(floating.origin.x, level)
      const oy = levelOffset(floating.origin.y, level)
      let area = union(baseRegion ? toLevelRect(baseRegion, level) : null, placed ? toLevelRect(placed.rect, level) : null)
      if (area) area = intersect(area, visible)
      if (!area) return null
      // Live layer pixels at this level, in document-level coordinates.
      const pixels = readSurfaceLevel(floating.surface, level, { x: area.x - ox, y: area.y - oy, width: area.width, height: area.height })
      const base = baseAt(level)
      if (base) blitCopy(pixels, area, base.pixels, translateRect(base.rect, ox, oy))
      if (placed) {
        const entry = placedAtLevel(level, placed)
        const px = levelFloor(placed.rect.x, level)
        const py = levelFloor(placed.rect.y, level)
        blitOver(pixels, area, entry.pixels, { x: px, y: py, width: entry.pixels.width, height: entry.pixels.height })
      }
      return { kind: 'layer-pixels', layerId: floating.layerId, level, rect: area, pixels }
    },
  }
}

/** The compositor's interactive mode (coarser proxies while dragging), when it supports it. */
export function setInteractive(ctx: ToolContext, active: boolean): void {
  const compositor = ctx.compositor as { setInteractive?: (value: boolean) => void }
  if (typeof compositor.setInteractive === 'function') compositor.setInteractive(active)
}

// ---------------------------------------------------------------------------------------------
// Options
// ---------------------------------------------------------------------------------------------

export function optionsOf(ctx: ToolContext): ToolOptions {
  return ctx.editor.getState().options
}

/** Document pixels per `px` CSS px at the current zoom (hit-test tolerances). */
export function screenToDocDistance(ctx: ToolContext, px: number): number {
  const zoom = ctx.view.getView().zoom
  return px / (zoom > 0 ? zoom : 1)
}

/**
 * True for a usable editor context. A controller activated with anything less (an editor still starting,
 * a test probe) stays inert: every handler returns early until a real context arrives.
 */
export function isToolContext(value: unknown): value is ToolContext {
  const ctx = value as Partial<ToolContext> | null | undefined
  return Boolean(ctx && typeof ctx === 'object'
    && ctx.store && typeof ctx.store.getState === 'function' && typeof ctx.store.transact === 'function'
    && ctx.editor && typeof ctx.editor.getState === 'function'
    && ctx.view && typeof ctx.view.getView === 'function'
    && ctx.compositor && ctx.imaging && ctx.host && typeof ctx.host.notify === 'function'
    && typeof ctx.setCursor === 'function' && typeof ctx.setHint === 'function' && typeof ctx.setOverlayElement === 'function')
}

export function isPrimaryButton(event: ToolPointerEvent): boolean {
  return event.button === 0 || (event.button === -1 && (event.buttons & 1) === 1)
}

/**
 * Space-bar panning inside a tool session (crop, free transform, polygonal lasso): holding Space turns
 * drags into view pans without leaving the tool, so the session is never closed just to scroll.
 */
export interface SpacePan {
  readonly held: boolean
  /** Handles Space (returns true when consumed). */
  keyDown(event: KeyboardEvent): boolean
  keyUp(event: KeyboardEvent): boolean
  /** True when the pointer event was used for panning (the tool must ignore it). */
  pointerDown(event: ToolPointerEvent): boolean
  pointerMove(event: ToolPointerEvent): boolean
  pointerUp(event: ToolPointerEvent): boolean
  reset(): void
}

export function createSpacePan(getContext: () => ToolContext | null, restoreCursor: () => void): SpacePan {
  let held = false
  let last: Point | null = null
  return {
    get held() {
      return held
    },
    keyDown(event: KeyboardEvent): boolean {
      if (event.key !== ' ' && event.code !== 'Space') return false
      if (!held) {
        held = true
        getContext()?.setCursor('grab')
      }
      return true
    },
    keyUp(event: KeyboardEvent): boolean {
      if (event.key !== ' ' && event.code !== 'Space') return false
      held = false
      if (!last) restoreCursor()
      return true
    },
    pointerDown(event: ToolPointerEvent): boolean {
      if (!held && event.button !== 1) return false
      last = { x: event.screen.x, y: event.screen.y }
      getContext()?.setCursor('grabbing')
      return true
    },
    pointerMove(event: ToolPointerEvent): boolean {
      if (!last) return false
      const ctx = getContext()
      if (ctx && event.buttons !== 0) ctx.view.panBy(event.screen.x - last.x, event.screen.y - last.y)
      last = { x: event.screen.x, y: event.screen.y }
      return true
    },
    pointerUp(): boolean {
      if (!last) return false
      last = null
      if (held) getContext()?.setCursor('grab')
      else restoreCursor()
      return true
    },
    reset(): void {
      held = false
      last = null
    },
  }
}

// ---------------------------------------------------------------------------------------------
// Snapshot images (overlay previews of transformed pixels)
// ---------------------------------------------------------------------------------------------

export interface SnapshotImage {
  readonly canvas: OffscreenCanvas | HTMLCanvasElement
  /** Source pixels per snapshot pixel (large content is previewed from a reduced copy). */
  readonly scale: number
}

/** A canvas holding `pixels` (halved until its long edge is at most maxEdge), or null without canvas support. */
export function createSnapshot(pixels: PixelBuffer, maxEdge = 2048): SnapshotImage | null {
  if (!(pixels.width > 0 && pixels.height > 0)) return null
  if (typeof OffscreenCanvas === 'undefined' && typeof document === 'undefined') return null
  let buffer = pixels
  let scale = 1
  while (Math.max(buffer.width, buffer.height) > maxEdge) {
    buffer = downsampleBuffer2x2(buffer)
    scale *= 2
  }
  let canvas: OffscreenCanvas | HTMLCanvasElement | null = null
  try {
    if (typeof OffscreenCanvas !== 'undefined') {
      canvas = new OffscreenCanvas(buffer.width, buffer.height)
    } else {
      canvas = document.createElement('canvas')
      canvas.width = buffer.width
      canvas.height = buffer.height
    }
    const context = (canvas as OffscreenCanvas).getContext('2d') as OffscreenCanvasRenderingContext2D | null
    if (!context) {
      releaseSnapshot({ canvas, scale })
      return null
    }
    context.putImageData(new ImageData(buffer.data as Uint8ClampedArray<ArrayBuffer>, buffer.width, buffer.height), 0, 0)
    return { canvas, scale }
  } catch {
    if (canvas) releaseSnapshot({ canvas, scale })
    return null
  }
}

/** Frees a snapshot's canvas memory. */
export function releaseSnapshot(snapshot: SnapshotImage | null): void {
  if (!snapshot) return
  snapshot.canvas.width = 1
  snapshot.canvas.height = 1
}

// ---------------------------------------------------------------------------------------------
// Overlay drawing (screen CSS px; the editor sets the device-pixel-ratio transform)
// ---------------------------------------------------------------------------------------------

/**
 * Makes a controller's drawOverlay work in CSS px even when the editor hands over an unscaled context of a
 * device-pixel-sized overlay canvas: then (identity transform, canvas wider than its CSS box) the ratio is
 * applied around the call. A context the editor already scaled is used as is.
 */
export function withCssPixelOverlay<T extends ToolController>(tool: T): T {
  const draw = tool.drawOverlay.bind(tool)
  tool.drawOverlay = (context: CanvasRenderingContext2D, view: ViewTransform) => {
    const canvas = context.canvas as HTMLCanvasElement | undefined
    const cssWidth = canvas && typeof canvas.clientWidth === 'number' ? canvas.clientWidth : 0
    if (cssWidth > 0 && typeof context.getTransform === 'function') {
      const m = context.getTransform()
      const ratio = canvas ? canvas.width / cssWidth : 1
      if (ratio > 1.01 && Math.abs(m.a - 1) < 1e-6 && Math.abs(m.d - 1) < 1e-6 && Math.abs(m.b) < 1e-9 && Math.abs(m.c) < 1e-9) {
        context.save()
        try {
          context.scale(ratio, ratio)
          draw(context, view)
        } finally {
          context.restore()
        }
        return
      }
    }
    draw(context, view)
  }
  return tool
}

type Overlay = CanvasRenderingContext2D

export function toScreen(view: ViewTransform, p: Point): Point {
  return { x: p.x * view.zoom + view.offsetX, y: p.y * view.zoom + view.offsetY }
}

/** A path drawn twice (white, then black dashes) so it shows on any background. */
export function strokeContrastPath(context: Overlay, points: readonly Point[], closed: boolean, dash = 4): void {
  if (points.length < 2) return
  context.save()
  try {
    context.beginPath()
    context.moveTo(points[0].x, points[0].y)
    for (let i = 1; i < points.length; i += 1) context.lineTo(points[i].x, points[i].y)
    if (closed) context.closePath()
    context.lineWidth = 1
    context.setLineDash([])
    context.strokeStyle = 'rgba(255, 255, 255, 0.95)'
    context.stroke()
    context.setLineDash([dash, dash])
    context.strokeStyle = 'rgba(0, 0, 0, 0.9)'
    context.stroke()
  } finally {
    context.restore()
  }
}

/** A solid outline with a light halo (crop and transform frames). */
export function strokeFramePath(context: Overlay, points: readonly Point[], closed: boolean): void {
  if (points.length < 2) return
  context.save()
  try {
    context.beginPath()
    context.moveTo(points[0].x, points[0].y)
    for (let i = 1; i < points.length; i += 1) context.lineTo(points[i].x, points[i].y)
    if (closed) context.closePath()
    context.setLineDash([])
    context.lineWidth = 3
    context.strokeStyle = 'rgba(255, 255, 255, 0.6)'
    context.stroke()
    context.lineWidth = 1
    context.strokeStyle = 'rgba(17, 17, 17, 0.95)'
    context.stroke()
  } finally {
    context.restore()
  }
}

/** A square handle centred on `p`. */
export function drawHandle(context: Overlay, p: Point, size = 7): void {
  const half = size / 2
  context.save()
  try {
    context.setLineDash([])
    context.fillStyle = '#ffffff'
    context.strokeStyle = '#111111'
    context.lineWidth = 1
    const x = Math.round(p.x - half) + 0.5
    const y = Math.round(p.y - half) + 0.5
    context.fillRect(x, y, size - 1, size - 1)
    context.strokeRect(x, y, size - 1, size - 1)
  } finally {
    context.restore()
  }
}

/** A reference-point marker (circle with a cross). */
export function drawPivot(context: Overlay, p: Point): void {
  context.save()
  try {
    context.setLineDash([])
    context.lineWidth = 1
    context.strokeStyle = '#111111'
    context.fillStyle = 'rgba(255, 255, 255, 0.85)'
    context.beginPath()
    context.arc(p.x, p.y, 4, 0, Math.PI * 2)
    context.fill()
    context.stroke()
    context.beginPath()
    context.moveTo(p.x - 7, p.y)
    context.lineTo(p.x + 7, p.y)
    context.moveTo(p.x, p.y - 7)
    context.lineTo(p.x, p.y + 7)
    context.stroke()
  } finally {
    context.restore()
  }
}

/** Clone-source / sample crosshair. */
export function drawCrosshair(context: Overlay, p: Point, radius = 8): void {
  context.save()
  try {
    context.setLineDash([])
    for (const [width, color] of [[3, 'rgba(255, 255, 255, 0.85)'], [1, 'rgba(17, 17, 17, 0.95)']] as const) {
      context.lineWidth = width
      context.strokeStyle = color
      context.beginPath()
      context.arc(p.x, p.y, radius * 0.6, 0, Math.PI * 2)
      context.moveTo(p.x - radius, p.y)
      context.lineTo(p.x + radius, p.y)
      context.moveTo(p.x, p.y - radius)
      context.lineTo(p.x, p.y + radius)
      context.stroke()
    }
  } finally {
    context.restore()
  }
}

/** Brush-size ring (dark and light circles so it shows on any background). */
export function drawBrushRing(context: Overlay, center: Point, radius: number): void {
  if (!(radius > 0)) return
  context.save()
  try {
    context.setLineDash([])
    context.lineWidth = 1
    context.strokeStyle = 'rgba(255, 255, 255, 0.9)'
    context.beginPath()
    context.arc(center.x, center.y, radius + 0.5, 0, Math.PI * 2)
    context.stroke()
    context.strokeStyle = 'rgba(0, 0, 0, 0.85)'
    context.beginPath()
    context.arc(center.x, center.y, Math.max(0.5, radius - 0.5), 0, Math.PI * 2)
    context.stroke()
  } finally {
    context.restore()
  }
}

/** A small text label with a dark pill (dimensions while dragging). */
export function drawLabel(context: Overlay, text: string, p: Point): void {
  context.save()
  try {
    context.font = '600 11px "Segoe UI", sans-serif'
    const width = context.measureText(text).width + 12
    const x = Math.round(p.x)
    const y = Math.round(p.y)
    context.fillStyle = 'rgba(17, 17, 17, 0.82)'
    context.beginPath()
    if (typeof context.roundRect === 'function') context.roundRect(x, y, width, 20, 4)
    else context.rect(x, y, width, 20)
    context.fill()
    context.fillStyle = '#ffffff'
    context.textBaseline = 'middle'
    context.fillText(text, x + 6, y + 10.5)
  } finally {
    context.restore()
  }
}

function svgCursor(svg: string, x: number, y: number, fallback: string): string {
  return `url("data:image/svg+xml,${encodeURIComponent(svg)}") ${x} ${y}, ${fallback}`
}

/** Photoshop-like curved double arrow for rotating. */
export const ROTATE_CURSOR = svgCursor(
  '<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><g fill="none" stroke-linecap="round" stroke-linejoin="round">'
  + '<path d="M6 15a7 7 0 0 1 9-9" stroke="#fff" stroke-width="4"/><path d="M6 15a7 7 0 0 1 9-9" stroke="#111" stroke-width="1.6"/>'
  + '<path d="M13 3.5 15.5 6 13 8.5M3.5 13 6 15.5 8.5 13" stroke="#fff" stroke-width="4"/><path d="M13 3.5 15.5 6 13 8.5M3.5 13 6 15.5 8.5 13" stroke="#111" stroke-width="1.6"/></g></svg>',
  12, 12, 'crosshair',
)

/** Resize cursor for a handle direction (degrees, 0 = east). */
export function resizeCursor(angle: number): string {
  const a = ((Math.round(angle / 45) % 4) + 4) % 4
  return ['ew-resize', 'nwse-resize', 'ns-resize', 'nesw-resize'][a]
}

/** Cursor for a pointer over a handle of a quad / box drawn at `center`. */
export function handleCursor(handlePoint: Point, center: Point): string {
  return resizeCursor((Math.atan2(handlePoint.y - center.y, handlePoint.x - center.x) * 180) / Math.PI)
}
