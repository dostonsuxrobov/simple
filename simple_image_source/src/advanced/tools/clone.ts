// src/advanced/tools/clone.ts (WP5)
// Clone Stamp (S): Alt+click sets the source; painting copies pixels from source + offset through the brush
// dabs. Aligned keeps the offset from the first stroke for every later stroke; otherwise each stroke starts
// again at the source point. The source is read from a pre-stroke snapshot (the layer's own pixels, the
// composite up to it, or all layers), so a stroke never clones what it just painted. A crosshair marks the
// source. Works on a targeted layer mask too. One history step per stroke.
import type { IntRect, MaskBuffer, PixelBuffer, Point } from '../../imaging/types.ts'
import type { CompositeOptions, PixelEditor, SampleSource, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { compositeRect } from '../composite.ts'
import type { AdvancedToolController, PaintTarget, StrokeEngine, StrokeRegion } from './shared.ts'
import {
  beginStrokeEngine,
  blendMaskValues,
  blendPixels,
  drawCrosshair,
  isToolContext,
  optionsOf,
  reportError,
  resolvePaintTarget,
  strokeSamples,
  toLocalRect,
  toScreen,
  translateRect,
} from './shared.ts'
import { brushCursor, drawBrushCursor } from './brush.ts'

function coverageBytes(region: StrokeRegion): Uint8Array {
  const out = new Uint8Array(region.coverage.length)
  const sel = region.selection ? region.selection.data : null
  for (let i = 0; i < out.length; i += 1) {
    const c = region.coverage[i]
    if (!(c > 0)) continue
    out[i] = Math.round(Math.min(1, c) * (sel ? sel[i] : 255))
  }
  return out
}

/** Source pixels for a document rectangle, read before the stroke changed anything. */
function cloneSource(ctx: ToolContext, target: PaintTarget, sample: SampleSource, editor: PixelEditor, rect: IntRect): PixelBuffer {
  const before = editor.readBefore(toLocalRect(target, rect)) as PixelBuffer
  if (sample === 'current') return before
  const state = ctx.store.getState()
  const options: CompositeOptions = sample === 'all' ? {} : { belowLayerId: target.layer.id, includeBelowLayer: true }
  return compositeRect(state, rect, { ...options, preview: { kind: 'layer-pixels', layerId: target.layer.id, level: 0, rect, pixels: before } })
}

export function createCloneTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let engine: StrokeEngine | null = null
  let hover: Point | null = null
  let source: Point | null = null
  let offset: Point | null = null
  let cursor = ''

  const finish = () => {
    const current = engine
    engine = null
    if (!current || !current.open) return
    try {
      current.finish(true)
    } catch (error) {
      if (ctx) reportError(ctx, error, 'The stroke could not be finished.')
    }
  }

  return {
    id: 'clone-stamp',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      engine = null
      hover = null
      cursor = brushCursor(context.view.getView(), optionsOf(context).clone.size)
      context.setCursor(cursor)
      context.setHint(source ? 'Paint to clone. Alt-click sets a new source.' : 'Alt-click to set the clone source, then paint.')
    },
    deactivate(): void {
      finish()
      hover = null
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0) return
      if (engine) finish()
      if (event.alt) {
        source = { x: Math.round(event.doc.x), y: Math.round(event.doc.y) }
        offset = null
        ctx.setHint('Paint to clone. Alt-click sets a new source.')
        ctx.view.requestOverlay()
        return
      }
      if (!source) {
        ctx.host.notify('Alt-click to set the clone source first.')
        return
      }
      const target = resolvePaintTarget(ctx)
      if (!target) return
      const context = ctx
      const opts = optionsOf(context).clone
      if (!opts.aligned || !offset) offset = { x: Math.round(source.x - event.doc.x), y: Math.round(source.y - event.doc.y) }
      const shift = offset
      const pen = event.pointerType === 'pen'
      const next = beginStrokeEngine(context, target, {
        label: 'Clone Stamp',
        icon: 'clone-stamp',
        tip: { diameter: Math.max(1, opts.size), hardness: opts.hardness, roundness: 1, angle: 0 },
        spacing: opts.spacing,
        smoothing: opts.smoothing,
        dynamics: { pressureSize: pen && opts.pressureSize, pressureFlow: pen && opts.pressureOpacity, flow: opts.flow },
        paint: (region: StrokeRegion): PixelBuffer | MaskBuffer => {
          const sourceRect = translateRect(region.docRect, shift.x, shift.y)
          const coverage = coverageBytes(region)
          if (target.channels === 1) {
            const values = region.editor.readBefore(toLocalRect(target, sourceRect)) as MaskBuffer
            return blendMaskValues(region.before as MaskBuffer, values.data, null, opts.opacity, coverage)
          }
          const pixels = cloneSource(context, target, opts.sample, region.editor, sourceRect)
          return blendPixels(region.before as PixelBuffer, pixels, {
            opacity: opts.opacity,
            mode: opts.blendMode,
            coverage,
            preserveAlpha: target.preserveAlpha,
          })
        },
      })
      if (!next) return
      engine = next
      hover = { x: event.doc.x, y: event.doc.y }
      try {
        engine.add(strokeSamples(event))
      } catch (error) {
        reportError(context, error, 'The stroke could not be painted.')
        finish()
      }
      context.view.requestOverlay()
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx) return
      hover = { x: event.doc.x, y: event.doc.y }
      if (engine) {
        if (event.buttons === 0) finish()
        else {
          try {
            engine.add(strokeSamples(event))
          } catch (error) {
            reportError(ctx, error, 'The stroke could not be painted.')
            finish()
          }
        }
      }
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (!ctx || !engine) return
      try {
        engine.add(strokeSamples(event), true)
      } catch (error) {
        reportError(ctx, error, 'The stroke could not be painted.')
      }
      finish()
      ctx.view.requestOverlay()
    },
    pointerCancel(): void {
      finish()
      ctx?.view.requestOverlay()
    },
    keyDown: () => false,
    keyUp: () => false,
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      if (!ctx) return
      const size = optionsOf(ctx).clone.size
      const next = brushCursor(view, size)
      if (next !== cursor) {
        cursor = next
        ctx.setCursor(next)
      }
      const marker = hover && offset ? { x: hover.x + offset.x, y: hover.y + offset.y } : source
      if (marker) drawCrosshair(context, toScreen(view, marker), 9)
      drawBrushCursor(context, view, hover, size)
    },
    hasSession: () => false,
    commitSession(): void {
      finish()
    },
    cancelSession(): void {
      finish()
    },
  }
}
