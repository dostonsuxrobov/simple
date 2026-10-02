// src/advanced/tools/move.ts (WP5)
// Move tool (V), Photoshop behaviour (design 5.9):
//   - drag moves the active layer (whole pixels, O(1): only its offset changes; a linked mask follows; an
//     unlinked, targeted mask moves alone); text and shape layers stay editable;
//   - with a selection on a pixel layer the selected pixels float: the source is cleared (filled with the
//     background colour on the Background), the pixels move with the selection, and repeated moves or nudges
//     keep the original pixels underneath intact until the selection changes;
//   - Alt-drag copies (the layer, or the selected pixels); Shift keeps the move horizontal, vertical or
//     diagonal; Ctrl-click (or Auto-Select) picks the topmost layer with a pixel under the pointer;
//   - arrow keys nudge by 1 px (10 px with Shift); a burst of nudges is one history step;
//   - "Show Transform Controls" draws the bounding box; dragging a handle starts a free transform.
// The drag is previewed through the compositor at the display level, so the layer keeps its place in the
// stack, its blend mode and opacity while it moves; releasing records exactly one history step.
import type { IntRect, Point } from '../../imaging/types.ts'
import type { Layer, LayerId, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { duplicateLayer, nextLayerName, pixelEditBlocker } from '../document.ts'
import { ensureMemory, layerBytes } from '../memory.ts'
import { translateSelection } from '../selection.ts'
import { createTransformFrame, hitTestTransform, nudgeDelta, snapToAngle, transformHandlePoint } from '../toolGeometry.ts'
import type { AdvancedToolController, FloatingPixels, FloatingPreviewer, PixelSource } from './shared.ts'
import {
  activeLayerOf,
  composeFloating,
  createFloatingPreviewer,
  findLayer,
  floatingIsCurrent,
  floatingMerged,
  handleCursor,
  intersect,
  isBackgroundLayer,
  isToolContext,
  layerIndexOf,
  layerPixelBounds,
  liftFloating,
  movedLayerPreview,
  nextSelectionVersion,
  optionsOf,
  pixelSourceOf,
  placedAt,
  reportError,
  screenToDocDistance,
  setInteractive,
  toLevelRect,
  topmostLayerAt,
  translateRect,
  union,
  visibleLevelRect,
} from './shared.ts'
import type { FreeTransformController } from './transform.ts'
import { createFreeTransform, drawFrame } from './transform.ts'

interface LayerDrag {
  readonly kind: 'layer'
  readonly layerId: LayerId
  readonly start: Point
  dx: number
  dy: number
  readonly duplicate: boolean
  readonly maskOnly: boolean
  readonly source: PixelSource | null
  readonly bounds: IntRect | null
}

interface FloatDrag {
  readonly kind: 'floating'
  readonly start: Point
  dx: number
  dy: number
  readonly base: Point
  readonly previewer: FloatingPreviewer
}

/** Document position of a layer (its pixels' origin; an adjustment layer's mask). */
function layerOrigin(layer: Layer): Point {
  if (layer.kind === 'raster') return { x: layer.offsetX, y: layer.offsetY }
  if (layer.kind === 'text' || layer.kind === 'shape') return { x: layer.raster.offsetX, y: layer.raster.offsetY }
  return layer.mask ? { x: layer.mask.offsetX, y: layer.mask.offsetY } : { x: 0, y: 0 }
}

function lockedMessage(layer: Layer): string {
  return isBackgroundLayer(layer)
    ? `"${layer.name}" is locked in place. Convert it to a normal layer (Layer > Layer from Background), or select an area to move pixels.`
    : `"${layer.name}" is locked in place. Unlock its position in the Layers panel first.`
}

export function createMoveTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let drag: LayerDrag | FloatDrag | null = null
  let floating: { pixels: FloatingPixels; offset: Point } | null = null
  let transform: FreeTransformController | null = null

  const clearPreview = () => {
    if (!ctx) return
    ctx.compositor.setPreview(null)
    setInteractive(ctx, false)
  }

  const controlsFrame = () => {
    if (!ctx || !optionsOf(ctx).move.showTransformControls) return null
    const state = ctx.store.getState()
    const layer = activeLayerOf(state)
    if (!layer || layer.kind === 'adjustment') return null
    const bounds = state.selection && layer.kind === 'raster' ? state.selection.bounds : layerPixelBounds(layer)
    return bounds ? createTransformFrame(bounds) : null
  }

  /** Writes the floating pixels at `offset` (one history step, or merged into a burst of nudges). */
  const mergeFloating = (offset: Point, label: string, coalesceKey?: string): boolean => {
    if (!ctx || !floating) return false
    const context = ctx
    const f = floating.pixels
    const state = context.store.getState()
    if (!floatingIsCurrent(f, state)) {
      floating = null
      return false
    }
    const placed = placedAt(f, offset.x, offset.y)
    const result = composeFloating(f, placed)
    const selection = translateSelection(f.selection, offset.x, offset.y, nextSelectionVersion(state))
    context.store.transact(label, 'move', (tx) => {
      if (result) tx.editPixels(f.layerId, 'pixels').writePixels(result.rect.x - f.origin.x, result.rect.y - f.origin.y, result.pixels)
      tx.setSelection(selection)
    }, coalesceKey ? { coalesceKey } : undefined)
    floatingMerged(f, placed.rect, context.store.getState().selection)
    floating.offset = offset
    return true
  }

  /** The floating session for the active layer's selection (lifting it when needed), or null after a message. */
  const ensureFloating = (layer: Layer, copy: boolean): boolean => {
    if (!ctx) return false
    const state = ctx.store.getState()
    if (layer.kind !== 'raster' || !state.selection) return false
    if (floating && (copy || floating.pixels.layerId !== layer.id || !floatingIsCurrent(floating.pixels, state))) floating = null
    if (floating) return true
    const lifted = liftFloating(state, layer, state.selection, { cut: !copy, fill: layer.isBackground ? ctx.editor.getState().background : null })
    if (!lifted) {
      ctx.host.notify('The selected area is empty.', 'error')
      return false
    }
    floating = { pixels: lifted, offset: { x: 0, y: 0 } }
    return true
  }

  const updateLayerPreview = (d: LayerDrag) => {
    if (!ctx || !d.source || !d.bounds || d.maskOnly) {
      ctx?.view.requestOverlay()
      return
    }
    const state = ctx.store.getState()
    const level = ctx.compositor.level
    const visible = visibleLevelRect(ctx, state, level)
    let area = union(toLevelRect(d.bounds, level), toLevelRect(translateRect(d.bounds, d.dx, d.dy), level))
    if (area && visible) area = intersect(area, visible)
    if (!area || !visible) {
      ctx.compositor.setPreview({ kind: 'layer-props', layerId: d.layerId, visible: d.duplicate })
      return
    }
    ctx.compositor.setPreview(movedLayerPreview(d.layerId, d.source, d.dx, d.dy, level, area, d.duplicate))
  }

  const updateFloatPreview = (d: FloatDrag) => {
    if (!ctx || !floating) return
    const f = floating.pixels
    const total = { x: d.base.x + d.dx, y: d.base.y + d.dy }
    const placed = { rect: translateRect(f.rect, total.x, total.y), pixels: f.lifted, key: f.lifted }
    ctx.compositor.setPreview(d.previewer.preview(ctx, ctx.store.getState(), ctx.compositor.level, placed))
  }

  const commitLayerMove = (d: LayerDrag) => {
    if (!ctx || (!d.dx && !d.dy)) return
    const context = ctx
    const state = context.store.getState()
    const layer = findLayer(state, d.layerId)
    if (!layer) return
    if (d.duplicate) {
      const bytes = layerBytes(layer)
      const room = ensureMemory(context.store, bytes.pixels + bytes.mask, 'copy this layer')
      if (!room.ok) {
        context.host.notify(room.message, 'error')
        return
      }
      const copy = duplicateLayer(layer, nextLayerName(state, `${layer.name} copy`))
      const origin = layerOrigin(copy)
      context.store.transact('Move', 'move', (tx) => {
        tx.insertLayer(copy, layerIndexOf(tx.state, layer.id) + 1)
        tx.updateLayer(copy.id, { offset: { x: origin.x + d.dx, y: origin.y + d.dy } })
        tx.setActiveLayer(copy.id)
      })
      return
    }
    if (d.maskOnly && layer.mask) {
      const mask = layer.mask
      context.store.transact('Move', 'move', (tx) => tx.setMask(layer.id, { ...mask, offsetX: mask.offsetX + d.dx, offsetY: mask.offsetY + d.dy }))
      return
    }
    const origin = layerOrigin(layer)
    context.store.transact('Move', 'move', (tx) => tx.updateLayer(layer.id, { offset: { x: origin.x + d.dx, y: origin.y + d.dy } }))
  }

  const startTransform = (event: ToolPointerEvent): boolean => {
    if (!ctx) return false
    const frame = controlsFrame()
    if (!frame) return false
    const hit = hitTestTransform(frame, event.doc, screenToDocDistance(ctx, 7))
    if (hit.kind !== 'handle') return false
    const controller = createFreeTransform({
      onEnd: () => {
        if (transform === controller) transform = null
        ctx?.setCursor('move')
        ctx?.view.requestOverlay()
      },
    })
    controller.activate(ctx)
    if (!controller.hasSession()) return true
    transform = controller
    controller.pointerDown(event)
    return true
  }

  const cancelDrag = () => {
    drag = null
    clearPreview()
    ctx?.view.requestOverlay()
  }

  return {
    id: 'move',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      drag = null
      floating = null
      context.setCursor('move')
      context.setHint('Drag to move the layer or the selected pixels. Alt-drag copies, Ctrl-click picks the layer under the pointer, arrows nudge.')
    },
    deactivate(): void {
      if (drag) cancelDrag()
      transform?.deactivate()
      transform = null
      floating = null
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx) return
      if (transform) {
        transform.pointerDown(event)
        return
      }
      if (event.button !== 0) return
      try {
        if (startTransform(event)) return
        const context = ctx
        let state = context.store.getState()
        if (optionsOf(context).move.autoSelect !== event.ctrl) {
          const hit = topmostLayerAt(state, event.doc)
          if (hit && hit.id !== state.activeLayerId) {
            context.store.transact('Select Layer', 'layer', (tx) => tx.setActiveLayer(hit.id))
            state = context.store.getState()
          }
        }
        const layer = activeLayerOf(state)
        if (!layer) {
          context.host.notify('Select a layer to move.', 'error')
          return
        }
        if (!layer.visible) {
          context.host.notify(`"${layer.name}" is hidden. Show it before moving it.`, 'error')
          return
        }
        if (state.selection && layer.kind === 'raster') {
          if (state.editTarget === 'mask') {
            context.host.notify('Moving selected pixels of a layer mask is not supported. Select the layer thumbnail to move its pixels.', 'error')
            return
          }
          const blocker = pixelEditBlocker(layer, 'pixels')
          if (blocker) {
            context.host.notify(blocker, 'error')
            return
          }
          if (!ensureFloating(layer, event.alt) || !floating) return
          drag = { kind: 'floating', start: { x: event.doc.x, y: event.doc.y }, dx: 0, dy: 0, base: floating.offset, previewer: createFloatingPreviewer(floating.pixels) }
          setInteractive(context, true)
          return
        }
        floating = null
        const maskOnly = state.editTarget === 'mask' && Boolean(layer.mask) && !layer.mask?.linked && layer.kind !== 'adjustment'
        if (!maskOnly && !event.alt && (layer.locks.position || isBackgroundLayer(layer))) {
          context.host.notify(lockedMessage(layer), 'error')
          return
        }
        if (layer.kind === 'adjustment' && !layer.mask && !event.alt) return
        drag = {
          kind: 'layer',
          layerId: layer.id,
          start: { x: event.doc.x, y: event.doc.y },
          dx: 0,
          dy: 0,
          duplicate: event.alt,
          maskOnly,
          source: pixelSourceOf(layer),
          bounds: layerPixelBounds(layer),
        }
        setInteractive(context, true)
      } catch (error) {
        cancelDrag()
        reportError(ctx, error, 'The layer could not be moved.')
      }
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx) return
      if (transform) {
        transform.pointerMove(event)
        return
      }
      if (!drag) {
        const frame = controlsFrame()
        if (frame) {
          const hit = hitTestTransform(frame, event.doc, screenToDocDistance(ctx, 7))
          ctx.setCursor(hit.kind === 'handle'
            ? handleCursor(transformHandlePoint(frame, hit.handle), { x: frame.source.x + frame.source.width / 2, y: frame.source.y + frame.source.height / 2 })
            : 'move')
        }
        return
      }
      if (event.buttons === 0) {
        cancelDrag()
        return
      }
      let delta = { x: event.doc.x - drag.start.x, y: event.doc.y - drag.start.y }
      if (event.shift) delta = snapToAngle({ x: 0, y: 0 }, delta, 45)
      const dx = Math.round(delta.x)
      const dy = Math.round(delta.y)
      if (dx === drag.dx && dy === drag.dy) return
      drag.dx = dx
      drag.dy = dy
      try {
        if (drag.kind === 'layer') updateLayerPreview(drag)
        else updateFloatPreview(drag)
      } catch (error) {
        reportError(ctx, error, 'The move could not be previewed.')
      }
    },
    pointerUp(event: ToolPointerEvent): void {
      if (!ctx) return
      if (transform) {
        transform.pointerUp(event)
        return
      }
      const current = drag
      if (!current) return
      drag = null
      clearPreview()
      const context = ctx
      try {
        if (current.kind === 'layer') commitLayerMove(current)
        else if (current.dx || current.dy) mergeFloating({ x: current.base.x + current.dx, y: current.base.y + current.dy }, 'Move')
      } catch (error) {
        reportError(context, error, 'The move could not be applied.')
      }
      context.view.requestOverlay()
    },
    pointerCancel(): void {
      if (transform) {
        transform.pointerCancel()
        return
      }
      cancelDrag()
    },
    keyDown(event: KeyboardEvent): boolean {
      if (!ctx) return false
      if (transform) return transform.keyDown(event)
      if (event.key === 'Escape' && drag) {
        cancelDrag()
        return true
      }
      const delta = nudgeDelta(event.key, event.shiftKey)
      if (!delta || event.ctrlKey || event.altKey || event.metaKey || drag) return false
      const context = ctx
      try {
        const state = context.store.getState()
        const layer = activeLayerOf(state)
        if (!layer) return false
        if (state.selection && layer.kind === 'raster') {
          const blocker = state.editTarget === 'mask' ? 'Nudging selected pixels of a layer mask is not supported.' : pixelEditBlocker(layer, 'pixels')
          if (blocker) {
            context.host.notify(blocker, 'error')
            return true
          }
          if (!ensureFloating(layer, false) || !floating) return true
          mergeFloating({ x: floating.offset.x + delta.x, y: floating.offset.y + delta.y }, 'Nudge', 'nudge-move')
          return true
        }
        floating = null
        if (layer.locks.position || isBackgroundLayer(layer)) {
          context.host.notify(lockedMessage(layer), 'error')
          return true
        }
        const origin = layerOrigin(layer)
        context.store.transact('Nudge', 'move', (tx) => tx.updateLayer(layer.id, { offset: { x: origin.x + delta.x, y: origin.y + delta.y } }), { coalesceKey: 'nudge-move' })
      } catch (error) {
        reportError(context, error, 'The layer could not be moved.')
      }
      return true
    },
    keyUp(event: KeyboardEvent): boolean {
      return transform ? transform.keyUp(event) : false
    },
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      if (transform) {
        transform.drawOverlay(context, view)
        return
      }
      if (drag) return
      const frame = controlsFrame()
      if (frame) drawFrame(context, view, frame)
    },
    hasSession: () => Boolean(transform && transform.hasSession()),
    commitSession(): void {
      transform?.commitSession()
    },
    cancelSession(): void {
      if (transform) transform.cancelSession()
      else if (drag) cancelDrag()
    },
    isBusy: () => Boolean(transform && transform.isBusy?.()),
    whenIdle: () => (transform && transform.whenIdle ? transform.whenIdle() : Promise.resolve()),
  }
}
