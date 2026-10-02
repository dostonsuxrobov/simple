// src/advanced/tools/transform.ts (WP5)
// Free Transform (Ctrl+T; also the Move tool's transform controls). Photoshop semantics (design 5.9):
//   - corner handle = scale, proportional by default (Shift toggles); edge handle = one axis (Shift scales
//     both); Alt scales about the reference point; dragging outside rotates (Shift snaps to 15 degrees);
//     dragging inside moves (Shift keeps it horizontal, vertical or diagonal); Ctrl+corner distorts,
//     Ctrl+Alt+Shift+corner applies perspective, Ctrl+edge skews (Alt: symmetric); the reference point can
//     be dragged; arrow keys nudge 1 px (10 px with Shift); numeric fields via info()/setInfo();
//     Enter or a double-click inside applies, Esc cancels; Space pans without leaving the session.
//   - Targets: the active pixel layer (its linked mask follows), its selected pixels (they float: the
//     source is cleared, or filled with the background colour on the Background, and the selection follows),
//     or a text / shape layer, which stays editable (its transform changes; distort and perspective need
//     a pixel layer).
//   - Preview: the layer is hidden (or shown with the hole) through the compositor and a snapshot is drawn
//     warped on the overlay (affine directly, perspective as an 8 x 8 triangle mesh). Applying runs the
//     bicubic warp in the imaging worker and records exactly one history step ("Free Transform"); whole-
//     pixel moves only shift the layer.
// It is not a ToolId: the editor creates it with createFreeTransform() and routes input to it while
// hasSession() is true; onEnd tells the editor when the session closed.
import type { Affine, Homography, IntRect, PixelBuffer, Point } from '../../imaging/types.ts'
import type { Layer, LayerId, LayerMask, Selection, TiledMask, TiledSurface, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { LIMITS } from '../types.ts'
import { applyHomography } from '../../imaging/transform.ts'
import { createMaskBuffer } from '../../imaging/mask.ts'
import { createSurface, maskFromBuffer, createMaskSurface, surfaceFromBuffer } from '../tiles.ts'
import { pixelEditBlocker } from '../document.ts'
import { selectionFromMask, translateSelection } from '../selection.ts'
import type { ClickStamp, FrameInfo, TransformFrame, TransformHandle } from '../toolGeometry.ts'
import {
  TRANSFORM_HANDLES,
  composeAffine,
  createTransformFrame,
  distortFrame,
  frameAffine,
  frameFromInfo,
  frameHomography,
  frameInfo,
  frameIntegerTranslation,
  hitTestTransform,
  isCornerHandle,
  isDoubleClick,
  isIdentityFrame,
  isValidQuad,
  nudgeDelta,
  perspectiveFrame,
  quadBounds,
  quadCenter,
  rotateFrame,
  scaleFrame,
  skewFrame,
  snapToAngle,
  transformHandlePoint,
  translateQuad,
  warpInverse,
} from '../toolGeometry.ts'
import type { AdvancedToolController, FloatingPixels, FloatingPreviewer, PlacedPixels, SnapshotImage } from './shared.ts'
import {
  activeLayerOf,
  canvasRect,
  composeFloating,
  createFloatingPreviewer,
  createJobTracker,
  createSnapshot,
  createSpacePan,
  drawHandle,
  drawLabel,
  drawPivot,
  findLayer,
  floatingIsCurrent,
  floatingMerged,
  handleCursor,
  intersect,
  isBackgroundLayer,
  isToolContext,
  layerPixelBounds,
  liftFloating,
  nextSelectionVersion,
  placedAt,
  readLayerRect,
  readMaskRect,
  releaseSnapshot,
  reportError,
  ROTATE_CURSOR,
  roundOut,
  screenToDocDistance,
  strokeFramePath,
  toScreen,
} from './shared.ts'

const HANDLE_TOLERANCE_PX = 7
const MESH = 8

// ---------------------------------------------------------------------------------------------
// Warps (shared with the crop tool)
// ---------------------------------------------------------------------------------------------

/** Document rectangle covered by `rect` mapped through a homography (rounded out, one pixel of margin). */
export function warpedBounds(forward: Homography, rect: IntRect): IntRect {
  const corners = [
    { x: rect.x, y: rect.y },
    { x: rect.x + rect.width, y: rect.y },
    { x: rect.x + rect.width, y: rect.y + rect.height },
    { x: rect.x, y: rect.y + rect.height },
  ].map((p) => applyHomography(forward, p))
  const xs = corners.map((p) => p.x)
  const ys = corners.map((p) => p.y)
  const bounds = roundOut({ x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) })
  return { x: bounds.x - 1, y: bounds.y - 1, width: bounds.width + 2, height: bounds.height + 2 }
}

/**
 * Warps pixels whose top-left sits at document (x, y) through a document-space homography (bicubic, in the
 * imaging worker). Output larger than the document limits is clipped to the canvas plus a margin; null when
 * nothing remains.
 */
export async function warpPixels(ctx: ToolContext, src: PixelBuffer, x: number, y: number, forward: Homography, clip?: IntRect | null): Promise<PlacedPixels | null> {
  if (!(src.width > 0 && src.height > 0)) return null
  let out: IntRect | null = warpedBounds(forward, { x, y, width: src.width, height: src.height })
  if (clip) out = intersect(out, clip)
  if (out && out.width * out.height > LIMITS.maxPixels) {
    const state = ctx.store.getState()
    const marginX = Math.round(state.width / 4)
    const marginY = Math.round(state.height / 4)
    out = intersect(out, { x: -marginX, y: -marginY, width: state.width + 2 * marginX, height: state.height + 2 * marginY })
  }
  if (!out) return null
  if (out.width > LIMITS.maxDimension || out.height > LIMITS.maxDimension || out.width * out.height > LIMITS.maxPixels) {
    throw new RangeError('The transformed pixels would be too large. Make the transformation smaller and try again.')
  }
  const inverse = warpInverse(forward, x, y)
  const pixels = await ctx.imaging.run('warp', { src, inverse, out, interpolation: 'bicubic' }, { transfer: false })
  return { rect: out, pixels }
}

/** Warps a layer mask (its non-default part); null when the mask is uniform (nothing to move). */
export async function warpMask(ctx: ToolContext, mask: LayerMask, forward: Homography): Promise<{ readonly surface: TiledMask; readonly offset: Point } | null> {
  const bounds = mask.surface.contentBounds()
  if (!bounds) return null
  const fill = mask.surface.defaultValue
  const values = mask.surface.read(bounds)
  // Coverage as alpha: the part that differs from the default value, so outside the warp stays default.
  const pixels: PixelBuffer = { width: bounds.width, height: bounds.height, data: new Uint8ClampedArray(bounds.width * bounds.height * 4) }
  for (let i = 0, p = 3; i < values.data.length; i += 1, p += 4) pixels.data[p] = fill === 255 ? 255 - values.data[i] : values.data[i]
  const warped = await warpPixels(ctx, pixels, bounds.x + mask.offsetX, bounds.y + mask.offsetY, forward)
  if (!warped) return { surface: createMaskSurface(fill), offset: { x: 0, y: 0 } }
  const out = new Uint8Array(warped.pixels.width * warped.pixels.height)
  for (let i = 0, p = 3; i < out.length; i += 1, p += 4) out[i] = fill === 255 ? 255 - warped.pixels.data[p] : warped.pixels.data[p]
  return {
    surface: maskFromBuffer({ width: warped.pixels.width, height: warped.pixels.height, data: out }, fill, warped.rect.x, warped.rect.y),
    offset: { x: 0, y: 0 },
  }
}

// ---------------------------------------------------------------------------------------------
// Overlay preview
// ---------------------------------------------------------------------------------------------

/** Affine (canvas order) that maps triangle u0..u2 (image coordinates) onto x0..x2. */
function triangleAffine(u: readonly Point[], x: readonly Point[]): Affine | null {
  const [u0, u1, u2] = u
  const [x0, x1, x2] = x
  const det = u0.x * (u1.y - u2.y) + u1.x * (u2.y - u0.y) + u2.x * (u0.y - u1.y)
  if (!(Math.abs(det) > 1e-12)) return null
  const a = (x0.x * (u1.y - u2.y) + x1.x * (u2.y - u0.y) + x2.x * (u0.y - u1.y)) / det
  const c = (x0.x * (u2.x - u1.x) + x1.x * (u0.x - u2.x) + x2.x * (u1.x - u0.x)) / det
  const e = (x0.x * (u1.x * u2.y - u2.x * u1.y) + x1.x * (u2.x * u0.y - u0.x * u2.y) + x2.x * (u0.x * u1.y - u1.x * u0.y)) / det
  const b = (x0.y * (u1.y - u2.y) + x1.y * (u2.y - u0.y) + x2.y * (u0.y - u1.y)) / det
  const d = (x0.y * (u2.x - u1.x) + x1.y * (u0.x - u2.x) + x2.y * (u1.x - u0.x)) / det
  const f = (x0.y * (u1.x * u2.y - u2.x * u1.y) + x1.y * (u2.x * u0.y - u0.x * u2.y) + x2.y * (u0.x * u1.y - u1.x * u0.y)) / det
  return [a, b, c, d, e, f]
}

/** Draws the snapshot of the frame's source warped onto its quad (in document units after the view transform). */
export function drawWarpedSnapshot(context: CanvasRenderingContext2D, view: ViewTransform, snapshot: SnapshotImage, frame: TransformFrame, opacity: number): void {
  const source = frame.source
  const image = snapshot.canvas as CanvasImageSource
  const iw = snapshot.canvas.width
  const ih = snapshot.canvas.height
  context.save()
  try {
    context.globalAlpha = Math.max(0, Math.min(1, opacity))
    context.imageSmoothingEnabled = true
    context.imageSmoothingQuality = 'medium'
    context.translate(view.offsetX, view.offsetY)
    context.scale(view.zoom, view.zoom)
    const affine = frameAffine(frame, 1e-7)
    if (affine) {
      context.transform(affine[0], affine[1], affine[2], affine[3], affine[4], affine[5])
      context.drawImage(image, 0, 0, iw, ih, source.x, source.y, source.width, source.height)
      return
    }
    let h: Homography
    try {
      h = frameHomography(frame)
    } catch {
      return
    }
    const seam = 0.6 / Math.max(1e-6, view.zoom)
    const toImage = (p: Point): Point => ({ x: ((p.x - source.x) / source.width) * iw, y: ((p.y - source.y) / source.height) * ih })
    for (let j = 0; j < MESH; j += 1) {
      for (let i = 0; i < MESH; i += 1) {
        const s = [
          { x: source.x + (i / MESH) * source.width, y: source.y + (j / MESH) * source.height },
          { x: source.x + ((i + 1) / MESH) * source.width, y: source.y + (j / MESH) * source.height },
          { x: source.x + ((i + 1) / MESH) * source.width, y: source.y + ((j + 1) / MESH) * source.height },
          { x: source.x + (i / MESH) * source.width, y: source.y + ((j + 1) / MESH) * source.height },
        ]
        const d = s.map((p) => applyHomography(h, p))
        if (d.some((p) => !Number.isFinite(p.x) || !Number.isFinite(p.y))) continue
        const u = s.map(toImage)
        for (const [a, b, c] of [[0, 1, 2], [0, 2, 3]] as const) {
          const m = triangleAffine([u[a], u[b], u[c]], [d[a], d[b], d[c]])
          if (!m) continue
          const cx = (d[a].x + d[b].x + d[c].x) / 3
          const cy = (d[a].y + d[b].y + d[c].y) / 3
          const grow = (p: Point) => {
            const length = Math.hypot(p.x - cx, p.y - cy) || 1
            return { x: p.x + ((p.x - cx) / length) * seam, y: p.y + ((p.y - cy) / length) * seam }
          }
          const ga = grow(d[a])
          const gb = grow(d[b])
          const gc = grow(d[c])
          context.save()
          context.beginPath()
          context.moveTo(ga.x, ga.y)
          context.lineTo(gb.x, gb.y)
          context.lineTo(gc.x, gc.y)
          context.closePath()
          context.clip()
          context.transform(m[0], m[1], m[2], m[3], m[4], m[5])
          const ux0 = Math.max(0, Math.floor(Math.min(u[a].x, u[b].x, u[c].x)) - 1)
          const uy0 = Math.max(0, Math.floor(Math.min(u[a].y, u[b].y, u[c].y)) - 1)
          const ux1 = Math.min(iw, Math.ceil(Math.max(u[a].x, u[b].x, u[c].x)) + 1)
          const uy1 = Math.min(ih, Math.ceil(Math.max(u[a].y, u[b].y, u[c].y)) + 1)
          if (ux1 > ux0 && uy1 > uy0) context.drawImage(image, ux0, uy0, ux1 - ux0, uy1 - uy0, ux0, uy0, ux1 - ux0, uy1 - uy0)
          context.restore()
        }
      }
    }
  } finally {
    context.restore()
  }
}

/** The frame outline, its eight handles and the reference point. */
export function drawFrame(context: CanvasRenderingContext2D, view: ViewTransform, frame: TransformFrame): void {
  strokeFramePath(context, frame.quad.map((p) => toScreen(view, p)), true)
  for (const handle of TRANSFORM_HANDLES) drawHandle(context, toScreen(view, transformHandlePoint(frame, handle)))
  drawPivot(context, toScreen(view, frame.pivot))
}

/** The pixels as they look on screen: alpha times an enabled layer mask (display only). */
function maskedForDisplay(pixels: PixelBuffer, rect: IntRect, mask: LayerMask | null): PixelBuffer {
  if (!mask || !mask.enabled) return pixels
  const values = readMaskRect(mask, rect).data
  const data = new Uint8ClampedArray(pixels.data)
  for (let i = 0, p = 3; i < values.length; i += 1, p += 4) if (values[i] !== 255) data[p] = Math.round((data[p] * values[i]) / 255)
  return { width: pixels.width, height: pixels.height, data }
}

// ---------------------------------------------------------------------------------------------
// Controller
// ---------------------------------------------------------------------------------------------

export interface FreeTransformOptions {
  /** Called once when a started session closes; `committed` is true when the document changed. */
  readonly onEnd?: (committed: boolean) => void
}

export interface FreeTransformController extends AdvancedToolController {
  /** Numeric fields of the open session (reference point, size, angle), or null. */
  info(): FrameInfo | null
  /** Applies numeric fields (skew and distortion reset). */
  setInfo(info: Partial<FrameInfo>): void
  /** False for text and shape layers (they stay vector: no distort or perspective). */
  canDistort(): boolean
}

type DragMode = 'move' | 'scale' | 'rotate' | 'skew' | 'distort' | 'perspective' | 'pivot'

interface Session {
  readonly kind: 'raster' | 'floating' | 'vector'
  readonly layerId: LayerId
  readonly layer: Layer
  frame: TransformFrame
  readonly start: TransformFrame
  readonly sourcePixels: PixelBuffer | null
  readonly floating: FloatingPixels | null
  readonly previewer: FloatingPreviewer | null
  readonly surface: TiledSurface | null
  readonly surfaceVersion: number
  readonly selection: Selection | null
  snapshot: SnapshotImage | null
  drag: { mode: DragMode; handle: TransformHandle | null; start: Point; frame: TransformFrame; symmetric: boolean } | null
  committing: boolean
  previewLevel: number
}

function sameLayerContent(session: Session, layer: Layer | null): boolean {
  if (!layer) return false
  if (session.kind === 'vector') return layer === session.layer
  if (layer.kind !== 'raster') return false
  return layer.surface === session.surface && layer.surface.version === session.surfaceVersion
}

export function createFreeTransform(options: FreeTransformOptions = {}): FreeTransformController {
  let ctx: ToolContext | null = null
  let session: Session | null = null
  let unsubscribe: (() => void) | null = null
  let lastClick: ClickStamp | null = null
  let warnedVector = false
  const jobs = createJobTracker()
  const pan = createSpacePan(() => ctx, () => ctx?.setCursor('default'))

  const end = (committed: boolean) => {
    const current = session
    if (!current) return
    session = null
    unsubscribe?.()
    unsubscribe = null
    pan.reset()
    releaseSnapshot(current.snapshot)
    if (ctx) {
      ctx.compositor.setPreview(null)
      ctx.setHint(null)
      ctx.setCursor('default')
      ctx.view.requestOverlay()
    }
    options.onEnd?.(committed)
  }

  const refreshPreview = (current: Session) => {
    if (!ctx) return
    if (current.kind === 'floating' && current.previewer) {
      const state = ctx.store.getState()
      current.previewLevel = ctx.compositor.level
      ctx.compositor.setPreview(current.previewer.preview(ctx, state, current.previewLevel, null, false))
    } else {
      ctx.compositor.setPreview({ kind: 'layer-props', layerId: current.layerId, visible: false })
    }
  }

  const begin = (context: ToolContext): boolean => {
    const state = context.store.getState()
    const layer = activeLayerOf(state)
    if (!layer) {
      context.host.notify('Select a layer to transform.', 'error')
      return false
    }
    if (layer.kind === 'adjustment') {
      context.host.notify(`"${layer.name}" is an adjustment layer. Select a pixel, type or shape layer to transform.`, 'error')
      return false
    }
    const selection = state.selection
    let kind: Session['kind']
    let floating: FloatingPixels | null = null
    let sourcePixels: PixelBuffer | null = null
    let bounds: IntRect | null = null
    if (selection && layer.kind === 'raster') {
      const blocker = pixelEditBlocker(layer, 'pixels')
      if (blocker) {
        context.host.notify(blocker, 'error')
        return false
      }
      floating = liftFloating(state, layer, selection, { cut: true, fill: layer.isBackground ? context.editor.getState().background : null })
      if (!floating) {
        context.host.notify('The selected area is empty.', 'error')
        return false
      }
      kind = 'floating'
      bounds = floating.rect
      sourcePixels = floating.lifted
    } else {
      if (isBackgroundLayer(layer) || layer.locks.position) {
        context.host.notify(`"${layer.name}" is locked in place. ${isBackgroundLayer(layer) ? 'Convert it to a normal layer (Layer > Layer from Background) or select an area' : 'Unlock its position'} to transform it.`, 'error')
        return false
      }
      if (layer.kind === 'raster') {
        if (layer.locks.pixels) {
          context.host.notify(`"${layer.name}" is locked. Unlock its pixels in the Layers panel first.`, 'error')
          return false
        }
        kind = 'raster'
      } else {
        kind = 'vector'
      }
      bounds = layerPixelBounds(layer)
      if (!bounds) {
        context.host.notify(`"${layer.name}" is empty; there is nothing to transform.`, 'error')
        return false
      }
      sourcePixels = readLayerRect(layer, bounds)
    }
    const frame = createTransformFrame(bounds)
    session = {
      kind,
      layerId: layer.id,
      layer,
      frame,
      start: frame,
      sourcePixels,
      floating,
      previewer: floating ? createFloatingPreviewer(floating) : null,
      surface: layer.kind === 'raster' ? layer.surface : null,
      surfaceVersion: layer.kind === 'raster' ? layer.surface.version : 0,
      selection,
      snapshot: sourcePixels ? createSnapshot(maskedForDisplay(sourcePixels, bounds, layer.mask)) : null,
      drag: null,
      committing: false,
      previewLevel: context.compositor.level,
    }
    refreshPreview(session)
    // Anything else changing the layer (undo, another command) ends the session without applying it.
    unsubscribe = context.store.subscribe(() => {
      const current = session
      if (!current || current.committing) return
      const now = context.store.getState()
      const live = findLayer(now, current.layerId)
      const valid = current.floating ? floatingIsCurrent(current.floating, now) : sameLayerContent(current, live)
      if (!valid) end(false)
    })
    context.setCursor('default')
    context.setHint('Drag a handle to scale (Shift: free), outside to rotate; Ctrl distorts. Enter applies, Esc cancels.')
    context.view.requestOverlay()
    return true
  }

  const modified = () => Boolean(session && !isIdentityFrame(session.frame, 1e-6))

  const apply = async (context: ToolContext, current: Session) => {
    const frame = current.frame
    const state = context.store.getState()
    const live = findLayer(state, current.layerId)
    if (current.kind === 'vector') {
      if (!live || (live.kind !== 'text' && live.kind !== 'shape')) return false
      const affine = frameAffine(frame, 1e-6)
      if (!affine) throw new RangeError('Distort and Perspective need a pixel layer. Rasterize the layer first.')
      const mask = live.mask && live.mask.linked ? await warpMask(context, live.mask, frameHomography(frame)) : null
      context.store.transact('Free Transform', 'transform', (tx) => {
        if (live.kind === 'text') tx.updateLayer(live.id, { text: { ...live.text, transform: composeAffine(affine, live.text.transform) } })
        else if (live.kind === 'shape') tx.updateLayer(live.id, { shape: { ...live.shape, transform: composeAffine(affine, live.shape.transform) } })
        if (mask) tx.replaceSurface(live.id, 'mask', mask.surface, mask.offset)
      })
      return true
    }
    if (current.kind === 'floating' && current.floating) {
      const floating = current.floating
      const shift = frameIntegerTranslation(frame)
      let placed: PlacedPixels | null
      let selection: Selection | null
      if (shift) {
        placed = placedAt(floating, shift.x, shift.y)
        selection = translateSelection(floating.selection, shift.x, shift.y, nextSelectionVersion(state))
      } else {
        const forward = frameHomography(frame)
        const clip = floating.clip
        placed = await warpPixels(context, floating.lifted, floating.rect.x, floating.rect.y, forward, clip)
        const coverage: PixelBuffer = { width: floating.rect.width, height: floating.rect.height, data: new Uint8ClampedArray(floating.rect.width * floating.rect.height * 4) }
        for (let i = 0, p = 3; i < floating.coverage.data.length; i += 1, p += 4) coverage.data[p] = floating.coverage.data[i]
        const warpedSelection = await warpPixels(context, coverage, floating.rect.x, floating.rect.y, forward, canvasRect(state))
        const now = context.store.getState()
        const mask = createMaskBuffer(now.width, now.height)
        if (warpedSelection) {
          const w = warpedSelection.pixels
          const r = warpedSelection.rect
          for (let y = Math.max(0, r.y); y < Math.min(now.height, r.y + r.height); y += 1) {
            for (let x = Math.max(0, r.x); x < Math.min(now.width, r.x + r.width); x += 1) {
              mask.data[y * now.width + x] = w.data[((y - r.y) * w.width + (x - r.x)) * 4 + 3]
            }
          }
        }
        selection = selectionFromMask(mask, nextSelectionVersion(now))
      }
      const now = context.store.getState()
      if (!floatingIsCurrent(floating, now)) throw new Error('The layer changed while the transformation was being applied, so it was not applied.')
      const result = composeFloating(floating, placed)
      context.store.transact('Free Transform', 'transform', (tx) => {
        if (result) {
          const editor = tx.editPixels(floating.layerId, 'pixels')
          editor.writePixels(result.rect.x - floating.origin.x, result.rect.y - floating.origin.y, result.pixels)
        }
        tx.setSelection(selection)
      })
      floatingMerged(floating, placed ? placed.rect : null, context.store.getState().selection)
      return true
    }
    // Whole pixel layer.
    if (!live || live.kind !== 'raster' || !sameLayerContent(current, live)) {
      throw new Error('The layer changed while the transformation was being applied, so it was not applied.')
    }
    const shift = frameIntegerTranslation(frame)
    if (shift) {
      context.store.transact('Free Transform', 'transform', (tx) => tx.updateLayer(live.id, { offset: { x: live.offsetX + shift.x, y: live.offsetY + shift.y } }))
      return true
    }
    const forward = frameHomography(frame)
    const source = current.start.source
    const warped = current.sourcePixels ? await warpPixels(context, current.sourcePixels, source.x, source.y, forward) : null
    const mask = live.mask && live.mask.linked ? await warpMask(context, live.mask, forward) : null
    const latest = findLayer(context.store.getState(), live.id)
    if (!sameLayerContent(current, latest)) throw new Error('The layer changed while the transformation was being applied, so it was not applied.')
    const surface = warped ? surfaceFromBuffer(warped.pixels, warped.rect.x - live.offsetX, warped.rect.y - live.offsetY) : createSurface()
    context.store.transact('Free Transform', 'transform', (tx) => {
      tx.replaceSurface(live.id, 'pixels', surface)
      if (mask) tx.replaceSurface(live.id, 'mask', mask.surface, mask.offset)
    })
    return true
  }

  const commit = () => {
    const current = session
    const context = ctx
    if (!current || !context || current.committing) return
    if (!modified()) {
      end(false)
      return
    }
    current.committing = true
    current.drag = null
    context.setHint('Applying the transformation…')
    context.setCursor('progress')
    context.view.requestOverlay()
    jobs.run(async () => {
      let committed = false
      try {
        committed = await apply(context, current)
      } catch (error) {
        reportError(context, error, 'The transformation could not be applied.')
      } finally {
        if (session === current) {
          current.committing = false
          end(committed)
        }
      }
    }).catch(() => {})
  }

  const cancel = () => {
    if (!session || session.committing) return
    end(false)
  }

  const tolerance = () => (ctx ? screenToDocDistance(ctx, HANDLE_TOLERANCE_PX) : 4)

  const hoverCursor = (point: Point) => {
    if (!ctx || !session) return
    const hit = hitTestTransform(session.frame, point, tolerance())
    if (hit.kind === 'pivot') ctx.setCursor('move')
    else if (hit.kind === 'handle') ctx.setCursor(handleCursor(transformHandlePoint(session.frame, hit.handle), quadCenter(session.frame.quad)))
    else if (hit.kind === 'inside') ctx.setCursor('move')
    else ctx.setCursor(ROTATE_CURSOR)
  }

  const controller: FreeTransformController = {
    id: 'move',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      lastClick = null
      warnedVector = false
      if (session) return
      begin(context)
    },
    deactivate(): void {
      if (session && !session.committing) {
        // Never drop a transformation silently: a changed frame is applied (one undoable step).
        if (modified()) commit()
        else end(false)
      }
    },
    pointerDown(event: ToolPointerEvent): void {
      const current = session
      if (!ctx || !current || current.committing) return
      if (pan.pointerDown(event)) return
      if (event.button !== 0) return
      const stamp = { time: event.time, x: event.screen.x, y: event.screen.y }
      const hit = hitTestTransform(current.frame, event.doc, tolerance())
      if (hit.kind === 'inside' && isDoubleClick(lastClick, stamp)) {
        lastClick = null
        commit()
        return
      }
      lastClick = stamp
      let mode: DragMode
      let handle: TransformHandle | null = null
      if (hit.kind === 'pivot') mode = 'pivot'
      else if (hit.kind === 'inside') mode = 'move'
      else if (hit.kind === 'outside') mode = 'rotate'
      else {
        handle = hit.handle
        if (event.ctrl && current.kind !== 'vector') {
          if (isCornerHandle(handle)) mode = event.alt && event.shift ? 'perspective' : 'distort'
          else mode = 'skew'
        } else if (event.ctrl && current.kind === 'vector' && !isCornerHandle(handle)) {
          mode = 'skew'
        } else {
          if (event.ctrl && current.kind === 'vector' && !warnedVector) {
            warnedVector = true
            ctx.host.notify('Distort and Perspective need a pixel layer. Rasterize the layer to use them.')
          }
          mode = 'scale'
        }
      }
      current.drag = { mode, handle, start: { x: event.doc.x, y: event.doc.y }, frame: current.frame, symmetric: event.alt }
      ctx.view.requestOverlay()
    },
    pointerMove(event: ToolPointerEvent): void {
      const current = session
      if (!ctx || !current) return
      if (pan.pointerMove(event)) return
      if (current.committing) return
      const drag = current.drag
      if (!drag) {
        hoverCursor(event.doc)
        return
      }
      if (event.buttons === 0) {
        current.drag = null
        return
      }
      // A press that turned into a drag is not the first click of a double-click.
      lastClick = null
      const start = drag.frame
      let quad = start.quad
      let pivot = start.pivot
      switch (drag.mode) {
        case 'pivot': {
          // Snap to the centre and the eight handles within the tolerance.
          let best: Point = { x: event.doc.x, y: event.doc.y }
          let bestDistance = tolerance()
          for (const candidate of [quadCenter(start.quad), ...TRANSFORM_HANDLES.map((h) => transformHandlePoint(start, h))]) {
            const d = Math.hypot(candidate.x - event.doc.x, candidate.y - event.doc.y)
            if (d <= bestDistance) {
              best = candidate
              bestDistance = d
            }
          }
          current.frame = { ...start, pivot: best }
          ctx.view.requestOverlay()
          return
        }
        case 'move': {
          let delta = { x: event.doc.x - drag.start.x, y: event.doc.y - drag.start.y }
          if (event.shift) delta = snapToAngle({ x: 0, y: 0 }, delta, 45)
          quad = translateQuad(start.quad, delta.x, delta.y)
          pivot = { x: start.pivot.x + delta.x, y: start.pivot.y + delta.y }
          break
        }
        case 'rotate':
          quad = rotateFrame(start, drag.start, event.doc, event.shift ? 15 : undefined)
          break
        case 'scale': {
          const handle = drag.handle as TransformHandle
          const proportional = isCornerHandle(handle) ? !event.shift : event.shift
          quad = scaleFrame(start, handle, event.doc, { proportional, fromPivot: event.alt })
          break
        }
        case 'skew':
          quad = skewFrame(start, drag.handle as TransformHandle, drag.start, event.doc, event.alt)
          break
        case 'distort':
          quad = distortFrame(start, drag.handle as TransformHandle, drag.start, event.doc)
          break
        case 'perspective':
          quad = perspectiveFrame(start, drag.handle as TransformHandle, drag.start, event.doc)
          break
        default:
          break
      }
      if (!isValidQuad(quad)) return
      if (drag.mode !== 'move' && drag.mode !== 'rotate') {
        // The reference point keeps its place relative to the frame.
        try {
          const before = frameHomography(start)
          const after = frameHomography({ source: start.source, quad })
          const inverse = warpInverse(before, 0, 0)
          pivot = applyHomography(after, applyHomography(inverse, start.pivot))
          if (!Number.isFinite(pivot.x) || !Number.isFinite(pivot.y)) pivot = start.pivot
        } catch {
          pivot = start.pivot
        }
      }
      current.frame = { source: start.source, quad, pivot }
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (pan.pointerUp(event)) return
      const current = session
      if (!current) return
      current.drag = null
      hoverCursor(event.doc)
    },
    pointerCancel(): void {
      if (session) session.drag = null
      pan.reset()
    },
    keyDown(event: KeyboardEvent): boolean {
      const current = session
      if (!ctx || !current) return false
      if (current.committing) return event.key === 'Enter' || event.key === 'Escape'
      if (event.isComposing) return false
      if (pan.keyDown(event)) return true
      if (event.key === 'Enter') {
        commit()
        return true
      }
      if (event.key === 'Escape') {
        cancel()
        return true
      }
      const delta = nudgeDelta(event.key, event.shiftKey)
      if (delta && !event.ctrlKey && !event.altKey && !event.metaKey) {
        current.frame = {
          source: current.frame.source,
          quad: translateQuad(current.frame.quad, delta.x, delta.y),
          pivot: { x: current.frame.pivot.x + delta.x, y: current.frame.pivot.y + delta.y },
        }
        ctx.view.requestOverlay()
        return true
      }
      return false
    },
    keyUp(event: KeyboardEvent): boolean {
      return pan.held ? pan.keyUp(event) : false
    },
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      const current = session
      if (!ctx || !current) return
      if (current.kind === 'floating' && current.previewLevel !== ctx.compositor.level && !current.committing) refreshPreview(current)
      if (current.snapshot) drawWarpedSnapshot(context, view, current.snapshot, current.frame, current.layer.opacity)
      drawFrame(context, view, current.frame)
      if (current.committing) {
        const bounds = quadBounds(current.frame.quad)
        drawLabel(context, 'Applying…', toScreen(view, { x: bounds.x, y: bounds.y + bounds.height }))
      } else if (current.drag && current.drag.mode !== 'pivot') {
        const info = frameInfo(current.frame)
        const label = current.drag.mode === 'rotate'
          ? `${Math.round(info.angle * 10) / 10}°`
          : `W ${Math.round(info.width)} px  H ${Math.round(info.height)} px`
        const bounds = quadBounds(current.frame.quad)
        drawLabel(context, label, toScreen(view, { x: bounds.x + bounds.width, y: bounds.y + bounds.height }))
      }
    },
    hasSession: () => session !== null,
    commitSession(): void {
      commit()
    },
    cancelSession(): void {
      cancel()
    },
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
    info(): FrameInfo | null {
      return session ? frameInfo(session.frame) : null
    },
    setInfo(info: Partial<FrameInfo>): void {
      if (!session || session.committing) return
      session.frame = frameFromInfo(session.frame, info)
      ctx?.view.requestOverlay()
    },
    canDistort: () => Boolean(session && session.kind !== 'vector'),
  }
  return controller
}
