// src/advanced/tools/brush.ts (WP5)
// Brush (B) and Eraser (E). Dab stamping from imaging/brush.ts on a per-stroke coverage grid, applied to
// the pre-stroke pixels so opacity caps a stroke exactly (design 5.9):
//   - options: size, hardness, opacity, flow, spacing, smoothing, pen pressure for size / opacity, blend
//     mode (brush); painting stays inside the selection;
//   - Shift+click paints a straight line from the end of the previous stroke; Alt+click picks the
//     foreground colour (brush);
//   - the eraser lowers alpha; on the Background or a transparency-locked layer it paints the background
//     colour; on a targeted layer mask the brush paints the foreground grey and the eraser the background
//     grey (Photoshop);
//   - the stroke is drawn live and becomes exactly one history step when the pointer is released (an
//     interrupted stroke is committed, never dropped).
import type { BrushTip, Point } from '../../imaging/types.ts'
import type { EditTarget, LayerId, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import type { AdvancedToolController, PaintTarget, StrokeEngine, StrokeRegion } from './shared.ts'
import {
  beginStrokeEngine,
  drawBrushRing,
  drawCrosshair,
  grayOf,
  isToolContext,
  optionsOf,
  paintMask,
  paintPixels,
  reportError,
  resolvePaintTarget,
  strokeSamples,
  toScreen,
} from './shared.ts'
import { sampleColorAt } from './eyedropper.ts'
import type { MaskBuffer, PixelBuffer } from '../../imaging/types.ts'

/** Ring radius (CSS px) below which a crosshair shows instead of the brush outline. */
const MIN_RING = 3

export interface BrushCursorState {
  hover: Point | null
}

/** Shared brush-ring overlay for the painting tools. */
export function drawBrushCursor(context: CanvasRenderingContext2D, view: ViewTransform, hover: Point | null, size: number): void {
  if (!hover) return
  const center = toScreen(view, hover)
  const radius = (size / 2) * view.zoom
  if (radius >= MIN_RING) drawBrushRing(context, center, radius)
  else drawCrosshair(context, center, 6)
}

export function brushCursor(view: ViewTransform, size: number): string {
  return (size / 2) * view.zoom >= MIN_RING ? 'none' : 'crosshair'
}

export function createBrushTool(kind: 'brush' | 'eraser'): AdvancedToolController {
  let ctx: ToolContext | null = null
  let engine: StrokeEngine | null = null
  let hover: Point | null = null
  let lastStroke: { layerId: LayerId; target: EditTarget; point: Point } | null = null
  let cursor = ''

  const options = () => (ctx ? (kind === 'brush' ? optionsOf(ctx).brush : optionsOf(ctx).eraser) : null)

  const updateCursor = () => {
    const current = options()
    if (!ctx || !current) return
    cursor = brushCursor(ctx.view.getView(), current.size)
    ctx.setCursor(cursor)
  }

  const finish = () => {
    const current = engine
    engine = null
    if (!current || !current.open) return
    try {
      current.finish(true)
      const point = current.lastPoint
      if (point) lastStroke = { layerId: current.target.layer.id, target: current.target.target, point }
    } catch (error) {
      if (ctx) reportError(ctx, error, 'The stroke could not be finished.')
    }
  }

  const painter = (target: PaintTarget) => {
    if (!ctx) return null
    const editor = ctx.editor.getState()
    const opts = kind === 'brush' ? optionsOf(ctx).brush : optionsOf(ctx).eraser
    const foreground = editor.foreground
    const background = editor.background
    return (region: StrokeRegion): PixelBuffer | MaskBuffer => {
      if (target.channels === 1) {
        const value = grayOf(kind === 'brush' ? foreground : background)
        return paintMask(region.before as MaskBuffer, region.coverage, value, opts.opacity, region.selection)
      }
      const before = region.before as PixelBuffer
      if (kind === 'eraser') {
        // Locked transparency (the Background): the eraser paints the background colour instead.
        if (target.preserveAlpha) {
          return paintPixels(before, region.coverage, { color: background, opacity: opts.opacity, mode: 'normal', preserveAlpha: true, selection: region.selection })
        }
        return paintPixels(before, region.coverage, { color: background, opacity: opts.opacity, mode: 'normal', erase: true, selection: region.selection })
      }
      return paintPixels(before, region.coverage, {
        color: foreground,
        opacity: opts.opacity,
        mode: opts.blendMode,
        preserveAlpha: target.preserveAlpha,
        selection: region.selection,
        seed: region.docRect.x * 73856093 ^ region.docRect.y * 19349663,
      })
    }
  }

  return {
    id: kind,
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      engine = null
      hover = null
      updateCursor()
      context.setHint(kind === 'brush'
        ? 'Drag to paint. Shift-click draws a straight line; Alt-click picks a colour; [ and ] change the size.'
        : 'Drag to erase. On the Background the eraser paints the background colour.')
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
      if (event.alt && kind === 'brush') {
        sampleColorAt(ctx, event, false)
        return
      }
      const opts = options()
      if (!opts) return
      const target = resolvePaintTarget(ctx)
      if (!target) return
      const paint = painter(target)
      if (!paint) return
      const pen = event.pointerType === 'pen'
      const tip: BrushTip = { diameter: Math.max(1, opts.size), hardness: opts.hardness, roundness: 1, angle: 0 }
      const next = beginStrokeEngine(ctx, target, {
        label: kind === 'brush' ? 'Brush Tool' : 'Eraser',
        icon: kind,
        tip,
        spacing: opts.spacing,
        smoothing: opts.smoothing,
        dynamics: { pressureSize: pen && opts.pressureSize, pressureFlow: pen && opts.pressureOpacity, flow: opts.flow },
        paint,
      })
      if (!next) return
      engine = next
      hover = { x: event.doc.x, y: event.doc.y }
      try {
        const straight = event.shift && lastStroke && lastStroke.layerId === target.layer.id && lastStroke.target === target.target
        const samples = strokeSamples(event)
        if (straight && lastStroke) {
          const from = lastStroke.point
          engine.add([{ x: from.x, y: from.y, pressure: samples[0].pressure, time: event.time }, ...samples.slice(-1)], true)
        } else {
          engine.add(samples)
        }
      } catch (error) {
        reportError(ctx, error, 'The stroke could not be painted.')
        finish()
      }
      ctx.view.requestOverlay()
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx) return
      hover = { x: event.doc.x, y: event.doc.y }
      if (engine) {
        if (event.buttons === 0) {
          finish()
        } else {
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
      // An interrupted stroke keeps what was painted (one history step).
      finish()
      ctx?.view.requestOverlay()
    },
    keyDown: () => false,
    keyUp: () => false,
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      const opts = options()
      if (!opts) return
      // The ring replaces the cursor once it is big enough to see (zoom and size change it).
      const next = brushCursor(view, opts.size)
      if (ctx && next !== cursor) {
        cursor = next
        ctx.setCursor(next)
      }
      drawBrushCursor(context, view, hover, opts.size)
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
