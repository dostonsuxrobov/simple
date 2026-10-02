// src/advanced/tools/crop.ts (WP5)
// Crop tool (C), Photoshop behaviour (design 5.9, 5.12):
//   - the crop box starts around the whole canvas; drag inside it to draw a new box (until it was changed),
//     drag inside a changed box to move it, drag a handle to resize (aspect preset from the options bar;
//     Shift keeps the current ratio of a free box; Alt resizes about the centre), drag outside to rotate
//     (Shift snaps to 15 degrees), Ctrl+drag draws a straighten line; arrow keys nudge the box;
//   - overlay: the outside is dimmed, rule-of-thirds or grid guides, size readout while dragging;
//   - Enter, a double-click inside the box or commitSession() crops; Esc resets the box;
//   - "Delete cropped pixels" (default on) trims pixel layers to the new canvas, otherwise they keep their
//     pixels outside it; the Background is always trimmed and extended with the background colour; text and
//     shape layers stay editable; a turned box rotates every layer first (bicubic, in the imaging worker).
// Cropping is one history step ("Crop"). Switching tools with a changed box applies it rather than losing
// it (the editor can ask first: hasSession() is true while the box differs from the canvas).
import type { IntRect, PixelBuffer, Point, Rgb8 } from '../../imaging/types.ts'
import type { Layer, TiledMask, TiledSurface, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { affineToHomography } from '../../imaging/transform.ts'
import { surfaceFromBuffer } from '../tiles.ts'
import type { ClickStamp, CropBox, CropHandle } from '../toolGeometry.ts'
import {
  CROP_HANDLES,
  composeAffine,
  cropAspectRatio,
  cropBoxCorners,
  cropBoxFromRect,
  cropCommitRect,
  cropFromLocal,
  cropHandlePoint,
  cropTransform,
  dragCrop,
  fitCropRatio,
  hitTestCrop,
  isDoubleClick,
  moveCrop,
  normalizeAngle,
  nudgeDelta,
  resizeCrop,
  rotateCrop,
  straightenAngle,
} from '../toolGeometry.ts'
import type { AdvancedToolController } from './shared.ts'
import {
  canvasRect,
  createJobTracker,
  createSpacePan,
  drawHandle,
  drawLabel,
  handleCursor,
  isToolContext,
  layerPixelBounds,
  optionsOf,
  readLayerRect,
  reportError,
  ROTATE_CURSOR,
  screenToDocDistance,
  strokeContrastPath,
  strokeFramePath,
  toScreen,
} from './shared.ts'
import { warpMask, warpPixels } from './transform.ts'

const HANDLE_TOLERANCE_PX = 8
const ANGLE_EPSILON = 1e-6

type CropDragMode = 'resize' | 'move' | 'rotate' | 'new' | 'straighten'

interface CropDrag {
  readonly mode: CropDragMode
  readonly handle: CropHandle | null
  readonly start: Point
  readonly box: CropBox
  current: Point
}

/** Pixels over `rect` made opaque over a background colour (the Background never has transparency). */
function flattenOnto(pixels: PixelBuffer, color: Rgb8): PixelBuffer {
  const data = pixels.data
  for (let p = 0; p < data.length; p += 4) {
    const a = data[p + 3]
    if (a === 255) continue
    const k = a / 255
    data[p] = Math.round(data[p] * k + color.r * (1 - k))
    data[p + 1] = Math.round(data[p + 1] * k + color.g * (1 - k))
    data[p + 2] = Math.round(data[p + 2] * k + color.b * (1 - k))
    data[p + 3] = 255
  }
  return pixels
}

function sameBox(a: CropBox, b: CropBox): boolean {
  return Math.abs(a.cx - b.cx) < 1e-6 && Math.abs(a.cy - b.cy) < 1e-6 && Math.abs(a.width - b.width) < 1e-6
    && Math.abs(a.height - b.height) < 1e-6 && Math.abs(normalizeAngle(a.angle - b.angle)) < ANGLE_EPSILON
}

export function createCropTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let box: CropBox | null = null
  let drag: CropDrag | null = null
  let lastClick: ClickStamp | null = null
  let lastRatio: number | null = null
  let lastPreset = ''
  let unsubscribeStore: (() => void) | null = null
  let unsubscribeEditor: (() => void) | null = null
  let canvas = { width: 0, height: 0 }
  /** The user changed the box on the canvas (drag, nudge, straighten); only then does leaving apply it. */
  let touched = false
  const jobs = createJobTracker(() => ctx?.view.requestOverlay())
  const pan = createSpacePan(() => ctx, () => ctx?.setCursor('crosshair'))

  const fullBox = (): CropBox => cropBoxFromRect(canvasRect(canvas))

  const ratioNow = (): number | null => {
    if (!ctx) return null
    const options = optionsOf(ctx).crop
    return cropAspectRatio(options.aspect, options.portrait, canvas)
  }

  /** The box differs from the whole canvas (Enter would crop). */
  const changed = () => Boolean(box && !sameBox(box, fullBox()))
  const modified = () => touched && changed()

  const reset = () => {
    if (!ctx) return
    const state = ctx.store.getState()
    canvas = { width: state.width, height: state.height }
    lastRatio = ratioNow()
    const options = optionsOf(ctx).crop
    lastPreset = `${options.aspect}:${options.portrait}`
    box = fitCropRatio(fullBox(), lastRatio)
    drag = null
    touched = false
    ctx.view.requestOverlay()
  }

  /** Axis-aligned crop: offsets shift, pixel layers are trimmed when asked; one transaction. */
  const cropStraight = (context: ToolContext, rect: IntRect) => {
    const deleteCropped = optionsOf(context).crop.deleteCroppedPixels
    const background = context.editor.getState().background
    context.store.transact('Crop', 'crop', (tx) => {
      for (const layer of [...tx.state.layers]) {
        const maskTarget = layer.mask ? { ...layer.mask, offsetX: layer.mask.offsetX - rect.x, offsetY: layer.mask.offsetY - rect.y } : null
        if (layer.kind === 'raster') {
          if (layer.isBackground || deleteCropped) {
            const bounds = layerPixelBounds(layer)
            let pixels = readLayerRect(layer, rect)
            if (layer.isBackground) pixels = flattenOnto(pixels, background)
            if (bounds || layer.isBackground) tx.replaceSurface(layer.id, 'pixels', surfaceFromBuffer(pixels), { x: 0, y: 0 })
          } else {
            tx.updateLayer(layer.id, { offset: { x: layer.offsetX - rect.x, y: layer.offsetY - rect.y } })
          }
        } else if (layer.kind === 'text' || layer.kind === 'shape') {
          tx.updateLayer(layer.id, { offset: { x: layer.raster.offsetX - rect.x, y: layer.raster.offsetY - rect.y } })
        }
        // Every document position moves with the canvas origin, masks included (linked or not).
        if (maskTarget) {
          const current = tx.state.layers.find((item) => item.id === layer.id)
          const mask = current?.mask
          if (mask && (mask.offsetX !== maskTarget.offsetX || mask.offsetY !== maskTarget.offsetY)) {
            tx.setMask(layer.id, { ...mask, offsetX: maskTarget.offsetX, offsetY: maskTarget.offsetY })
          }
        }
      }
      tx.setCanvas(rect.width, rect.height)
      tx.setSelection(null)
    })
  }

  /** Turned crop: every layer rotates about the box first (worker warps), then one transaction. */
  const cropTurned = async (context: ToolContext, cropBox: CropBox) => {
    const { matrix, width, height } = cropTransform(cropBox)
    const forward = affineToHomography(matrix)
    const deleteCropped = optionsOf(context).crop.deleteCroppedPixels
    const background = context.editor.getState().background
    const state = context.store.getState()
    const target: IntRect = { x: 0, y: 0, width, height }
    interface Plan {
      readonly layer: Layer
      readonly surface: TiledSurface | null
      readonly version: number
      pixels?: { readonly surface: TiledSurface } | null
      mask?: { readonly surface: TiledMask; readonly offset: Point } | null
    }
    const plans: Plan[] = []
    for (const layer of state.layers) {
      const plan: Plan = { layer, surface: layer.kind === 'raster' ? layer.surface : null, version: layer.kind === 'raster' ? layer.surface.version : 0 }
      if (layer.kind === 'raster') {
        const bounds = layerPixelBounds(layer)
        if (bounds) {
          const clip = deleteCropped || layer.isBackground ? target : null
          const warped = await warpPixels(context, readLayerRect(layer, bounds), bounds.x, bounds.y, forward, clip)
          if (layer.isBackground) {
            const pixels = warped ? placeOnCanvas(warped, target) : { width, height, data: new Uint8ClampedArray(width * height * 4) }
            plan.pixels = { surface: surfaceFromBuffer(flattenOnto(pixels, background)) }
          } else {
            plan.pixels = { surface: warped ? surfaceFromBuffer(warped.pixels, warped.rect.x, warped.rect.y) : surfaceFromBuffer({ width: 0, height: 0, data: new Uint8ClampedArray(0) }) }
          }
        } else if (layer.isBackground) {
          plan.pixels = { surface: surfaceFromBuffer(flattenOnto({ width, height, data: new Uint8ClampedArray(width * height * 4) }, background)) }
        }
      }
      if (layer.mask) plan.mask = await warpMask(context, layer.mask, forward)
      plans.push(plan)
    }
    const now = context.store.getState()
    const unchanged = now.layers.length === state.layers.length && now.layers.every((layer, index) => {
      const plan = plans[index]
      if (layer.id !== plan.layer.id) return false
      if (plan.surface) return layer.kind === 'raster' && layer.surface === plan.surface && layer.surface.version === plan.version
      return layer === plan.layer
    })
    if (!unchanged) throw new Error('The image changed while it was being cropped, so the crop was not applied.')
    context.store.transact('Crop', 'crop', (tx) => {
      for (const plan of plans) {
        const layer = plan.layer
        if (plan.pixels) tx.replaceSurface(layer.id, 'pixels', plan.pixels.surface, { x: 0, y: 0 })
        if (layer.kind === 'text') tx.updateLayer(layer.id, { text: { ...layer.text, transform: composeAffine(matrix, layer.text.transform) } })
        else if (layer.kind === 'shape') tx.updateLayer(layer.id, { shape: { ...layer.shape, transform: composeAffine(matrix, layer.shape.transform) } })
        if (plan.mask) tx.replaceSurface(layer.id, 'mask', plan.mask.surface, plan.mask.offset)
      }
      tx.setCanvas(width, height)
      tx.setSelection(null)
    })
  }

  const commit = () => {
    const context = ctx
    const current = box
    if (!context || !current || jobs.busy) return
    if (!changed()) return
    drag = null
    if (Math.abs(normalizeAngle(current.angle)) < ANGLE_EPSILON) {
      const rect = cropCommitRect(current)
      if (rect.x === 0 && rect.y === 0 && rect.width === canvas.width && rect.height === canvas.height) {
        reset()
        return
      }
      try {
        cropStraight(context, rect)
      } catch (error) {
        reportError(context, error, 'The image could not be cropped.')
      }
      reset()
      return
    }
    context.setHint('Cropping…')
    jobs.run(() => cropTurned(context, current))
      .catch((error) => reportError(context, error, 'The image could not be cropped.'))
      .finally(() => {
        if (ctx === context) {
          context.setHint(hint)
          reset()
        }
      })
  }

  const hint = 'Drag to set the crop; drag outside to rotate, Ctrl-drag to straighten. Enter crops, Esc resets.'

  const tolerance = () => (ctx ? screenToDocDistance(ctx, HANDLE_TOLERANCE_PX) : 4)

  const hoverCursor = (point: Point) => {
    if (!ctx || !box) return
    const hit = hitTestCrop(box, point, tolerance())
    if (hit === 'inside') ctx.setCursor(modified() ? 'move' : 'crosshair')
    else if (hit === 'outside') ctx.setCursor(ROTATE_CURSOR)
    else ctx.setCursor(handleCursor(cropHandlePoint(box, hit), { x: box.cx, y: box.cy }))
  }

  return {
    id: 'crop',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      lastClick = null
      reset()
      context.setCursor('crosshair')
      context.setHint(hint)
      unsubscribeStore = context.store.subscribe(() => {
        const state = context.store.getState()
        if (state.width !== canvas.width || state.height !== canvas.height) reset()
      })
      unsubscribeEditor = context.editor.subscribe(() => {
        if (!box) return
        // Guides and other crop options are read while drawing.
        context.view.requestOverlay()
        const options = optionsOf(context).crop
        const preset = `${options.aspect}:${options.portrait}`
        if (preset === lastPreset) return
        lastPreset = preset
        const ratio = ratioNow()
        lastRatio = ratio
        box = ratio ? fitCropRatio(changed() ? box : fullBox(), ratio) : box
        context.view.requestOverlay()
      })
    },
    deactivate(): void {
      // A changed crop box is the user's work: apply it (one undoable step) instead of dropping it.
      if (modified() && !jobs.busy) commit()
      unsubscribeStore?.()
      unsubscribeEditor?.()
      unsubscribeStore = null
      unsubscribeEditor = null
      drag = null
      pan.reset()
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
      box = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || !box || jobs.busy) return
      if (pan.pointerDown(event)) return
      if (event.button !== 0) return
      const stamp = { time: event.time, x: event.screen.x, y: event.screen.y }
      const hit = hitTestCrop(box, event.doc, tolerance())
      if (hit === 'inside' && isDoubleClick(lastClick, stamp)) {
        lastClick = null
        commit()
        return
      }
      lastClick = stamp
      let mode: CropDragMode
      let handle: CropHandle | null = null
      if (event.ctrl) mode = 'straighten'
      else if (hit === 'inside') mode = modified() ? 'move' : 'new'
      else if (hit === 'outside') mode = 'rotate'
      else {
        mode = 'resize'
        handle = hit
      }
      drag = { mode, handle, start: { x: event.doc.x, y: event.doc.y }, box, current: { x: event.doc.x, y: event.doc.y } }
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx || !box) return
      if (pan.pointerMove(event)) return
      if (!drag) {
        hoverCursor(event.doc)
        return
      }
      if (event.buttons === 0) {
        drag = null
        return
      }
      drag.current = { x: event.doc.x, y: event.doc.y }
      // A press that turned into a drag is not the first click of a double-click.
      lastClick = null
      const ratio = lastRatio
      switch (drag.mode) {
        case 'resize':
          box = resizeCrop(drag.box, drag.handle as CropHandle, event.doc, { ratio, keepRatio: event.shift && !ratio, fromCenter: event.alt })
          break
        case 'move':
          box = moveCrop(drag.box, event.doc.x - drag.start.x, event.doc.y - drag.start.y)
          break
        case 'rotate':
          box = rotateCrop(drag.box, drag.start, event.doc, event.shift ? 15 : undefined)
          break
        case 'new':
          if (Math.hypot(event.doc.x - drag.start.x, event.doc.y - drag.start.y) >= screenToDocDistance(ctx, 3)) {
            box = dragCrop(drag.start, event.doc, { ratio, square: event.shift && !ratio, fromCenter: event.alt })
          }
          break
        default:
          break
      }
      if (drag.mode !== 'straighten' && box !== drag.box) touched = true
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (pan.pointerUp(event)) return
      if (!ctx || !drag) return
      const current = drag
      drag = null
      if (current.mode === 'straighten') {
        const angle = straightenAngle(current.start, event.doc)
        if (box && Math.hypot(event.doc.x - current.start.x, event.doc.y - current.start.y) >= screenToDocDistance(ctx, 4)) {
          box = { ...box, angle }
          touched = true
        }
      }
      hoverCursor(event.doc)
      ctx.view.requestOverlay()
    },
    pointerCancel(): void {
      if (drag) {
        box = drag.box
        drag = null
      }
      pan.reset()
      ctx?.view.requestOverlay()
    },
    keyDown(event: KeyboardEvent): boolean {
      if (!ctx || !box) return false
      if (jobs.busy) return event.key === 'Enter' || event.key === 'Escape'
      if (pan.keyDown(event)) return true
      if (event.key === 'Enter') {
        commit()
        return true
      }
      if (event.key === 'Escape') {
        reset()
        return true
      }
      const delta = nudgeDelta(event.key, event.shiftKey)
      if (delta && !event.ctrlKey && !event.altKey && !event.metaKey) {
        box = moveCrop(box, delta.x, delta.y)
        touched = true
        ctx.view.requestOverlay()
        return true
      }
      return false
    },
    keyUp(event: KeyboardEvent): boolean {
      return pan.held ? pan.keyUp(event) : false
    },
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      if (!ctx || !box) return
      const corners = cropBoxCorners(box).map((p) => toScreen(view, p))
      const size = ctx.view.getViewportSize()
      context.save()
      try {
        // Dim everything outside the box.
        context.beginPath()
        context.rect(-1, -1, Math.max(1, size.width) + 2, Math.max(1, size.height) + 2)
        context.moveTo(corners[0].x, corners[0].y)
        for (let i = 1; i < 4; i += 1) context.lineTo(corners[i].x, corners[i].y)
        context.closePath()
        context.fillStyle = 'rgba(0, 0, 0, 0.45)'
        context.fill('evenodd')
        // Guides.
        const overlay = optionsOf(ctx).crop.overlay
        const divisions = overlay === 'thirds' ? 3 : overlay === 'grid' ? 8 : 0
        if (divisions) {
          context.lineWidth = 1
          context.strokeStyle = 'rgba(255, 255, 255, 0.55)'
          context.beginPath()
          for (let i = 1; i < divisions; i += 1) {
            const t = i / divisions - 0.5
            const a = toScreen(view, cropFromLocal(box, { x: t * box.width, y: -box.height / 2 }))
            const b = toScreen(view, cropFromLocal(box, { x: t * box.width, y: box.height / 2 }))
            const c = toScreen(view, cropFromLocal(box, { x: -box.width / 2, y: t * box.height }))
            const d = toScreen(view, cropFromLocal(box, { x: box.width / 2, y: t * box.height }))
            context.moveTo(a.x, a.y)
            context.lineTo(b.x, b.y)
            context.moveTo(c.x, c.y)
            context.lineTo(d.x, d.y)
          }
          context.stroke()
        }
      } finally {
        context.restore()
      }
      strokeFramePath(context, corners, true)
      for (const handle of CROP_HANDLES) drawHandle(context, toScreen(view, cropHandlePoint(box, handle)), 8)
      if (drag && drag.mode === 'straighten') {
        strokeContrastPath(context, [toScreen(view, drag.start), toScreen(view, drag.current)], false)
        const angle = straightenAngle(drag.start, drag.current)
        drawLabel(context, `${Math.round(angle * 10) / 10}°`, { x: toScreen(view, drag.current).x + 12, y: toScreen(view, drag.current).y + 12 })
      } else if (drag || jobs.busy) {
        const text = jobs.busy
          ? 'Cropping…'
          : drag && drag.mode === 'rotate' ? `${Math.round(normalizeAngle(box.angle) * 10) / 10}°` : `${Math.round(box.width)} × ${Math.round(box.height)} px`
        const anchor = corners.reduce((best, p) => (p.y > best.y || (p.y === best.y && p.x > best.x) ? p : best), corners[0])
        drawLabel(context, text, { x: anchor.x + 10, y: anchor.y + 8 })
      }
    },
    hasSession: () => modified(),
    commitSession(): void {
      commit()
    },
    cancelSession(): void {
      if (!jobs.busy) reset()
    },
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}

/** A warped buffer placed on a target canvas rectangle (transparent where it does not reach). */
function placeOnCanvas(placed: { readonly rect: IntRect; readonly pixels: PixelBuffer }, target: IntRect): PixelBuffer {
  const out: PixelBuffer = { width: target.width, height: target.height, data: new Uint8ClampedArray(target.width * target.height * 4) }
  const x0 = Math.max(target.x, placed.rect.x)
  const y0 = Math.max(target.y, placed.rect.y)
  const x1 = Math.min(target.x + target.width, placed.rect.x + placed.rect.width)
  const y1 = Math.min(target.y + target.height, placed.rect.y + placed.rect.height)
  for (let y = y0; y < y1; y += 1) {
    const from = ((y - placed.rect.y) * placed.pixels.width + (x0 - placed.rect.x)) * 4
    out.data.set(placed.pixels.data.subarray(from, from + (x1 - x0) * 4), ((y - target.y) * target.width + (x0 - target.x)) * 4)
  }
  return out
}
