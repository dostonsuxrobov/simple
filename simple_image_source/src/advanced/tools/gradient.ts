// src/advanced/tools/gradient.ts (WP5)
// Gradient tool (G, Shift+G cycles with Paint Bucket): drag a line; on release the active layer (or its
// mask) is filled inside the selection (or the whole canvas) with a linear, radial, angle, reflected or
// diamond gradient. Presets: foreground to background, foreground to transparent, black to white; reverse
// and dither; Shift snaps the line to 45 degrees; the tool's opacity and blend mode apply; transparency-
// locked layers keep their alpha. Alt+click picks the foreground colour. One history step.
import type { GradientSpec, GradientStop, IntRect, OpacityStop, PixelBuffer, Point, Rgb8 } from '../../imaging/types.ts'
import type { ToolContext, ToolOptions, ToolPointerEvent, ViewTransform } from '../types.ts'
import { renderGradient } from '../../imaging/gradient.ts'
import { cropMask } from '../../imaging/mask.ts'
import { pixelEditBlocker } from '../document.ts'
import { distance, snapToAngle } from '../toolGeometry.ts'
import type { AdvancedToolController, PaintTarget } from './shared.ts'
import {
  blendMaskValues,
  blendPixels,
  canvasRect,
  createJobTracker,
  drawHandle,
  findLayer,
  grayAndAlpha,
  intersect,
  isToolContext,
  optionsOf,
  readLayerRect,
  readMaskRect,
  reportError,
  resolvePaintTarget,
  strokeContrastPath,
  toScreen,
  writeRegion,
} from './shared.ts'
import { sampleColorAt } from './eyedropper.ts'

const BLACK: Rgb8 = { r: 0, g: 0, b: 0 }
const WHITE: Rgb8 = { r: 255, g: 255, b: 255 }
/** Rows rendered between yields to the event loop. */
const STRIPE_ROWS = 64

/** The gradient a preset describes (foreground / background colours come from the editor). */
export function gradientSpecFor(options: ToolOptions['gradient'], from: Point, to: Point, foreground: Rgb8, background: Rgb8): GradientSpec {
  let stops: GradientStop[]
  let opacityStops: OpacityStop[] = [{ position: 0, opacity: 1 }, { position: 1, opacity: 1 }]
  switch (options.preset) {
    case 'foreground-transparent':
      stops = [{ position: 0, color: foreground }, { position: 1, color: foreground }]
      opacityStops = [{ position: 0, opacity: 1 }, { position: 1, opacity: 0 }]
      break
    case 'black-white':
      stops = [{ position: 0, color: BLACK }, { position: 1, color: WHITE }]
      break
    default:
      stops = [{ position: 0, color: foreground }, { position: 1, color: background }]
  }
  return { kind: options.kind, from, to, stops, opacityStops, reverse: options.reverse, dither: options.dither }
}

function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0))
}

export function createGradientTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let drag: { start: Point; end: Point; target: PaintTarget } | null = null
  const jobs = createJobTracker(() => ctx?.setCursor(jobs.busy ? 'progress' : 'crosshair'))

  const apply = async (context: ToolContext, target: PaintTarget, start: Point, end: Point) => {
    const state = context.store.getState()
    const options = optionsOf(context).gradient
    const colors = context.editor.getState()
    let region: IntRect | null = state.selection ? state.selection.bounds : canvasRect(state)
    if (region && target.clip) region = intersect(region, target.clip)
    if (!region) return
    const rect = region
    const spec = gradientSpecFor(options, start, end, colors.foreground, colors.background)
    // Render in stripes so a large canvas never blocks the window for long.
    const pixels: PixelBuffer = { width: rect.width, height: rect.height, data: new Uint8ClampedArray(rect.width * rect.height * 4) }
    let sliceStart = Date.now()
    for (let top = 0; top < rect.height; top += STRIPE_ROWS) {
      const rows = Math.min(STRIPE_ROWS, rect.height - top)
      const stripe = renderGradient(rect.width, rows, { x: rect.x, y: rect.y + top }, spec)
      pixels.data.set(stripe.data, top * rect.width * 4)
      if (Date.now() - sliceStart > 12) {
        await yieldToEventLoop()
        sliceStart = Date.now()
      }
    }
    const now = context.store.getState()
    if (now.width !== state.width || now.height !== state.height) return
    const layer = findLayer(now, target.layer.id)
    if (!layer || pixelEditBlocker(layer, target.target)) return
    const coverage = now.selection ? cropMask(now.selection.mask, rect).data : null
    context.store.transact('Gradient', 'gradient', (tx) => {
      const editor = tx.editPixels(layer.id, target.target)
      if (target.channels === 1 && layer.mask) {
        const before = readMaskRect(layer.mask, rect)
        const { values, alpha } = grayAndAlpha(pixels)
        writeRegion(editor, target, rect, blendMaskValues(before, values, alpha, options.opacity, coverage))
        return
      }
      const before = readLayerRect(layer, rect)
      writeRegion(editor, target, rect, blendPixels(before, pixels, {
        opacity: options.opacity,
        mode: options.blendMode,
        coverage,
        preserveAlpha: target.preserveAlpha,
      }))
    })
  }

  return {
    id: 'gradient',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      drag = null
      context.setCursor('crosshair')
      context.setHint('Drag to draw a gradient. Shift keeps the line at 45° steps; Alt-click picks a colour.')
    },
    deactivate(): void {
      drag = null
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0 || jobs.busy) return
      if (event.alt) {
        sampleColorAt(ctx, event, false)
        return
      }
      const target = resolvePaintTarget(ctx)
      if (!target) return
      drag = { start: { x: event.doc.x, y: event.doc.y }, end: { x: event.doc.x, y: event.doc.y }, target }
      ctx.view.requestOverlay()
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx || !drag) return
      drag.end = event.shift ? snapToAngle(drag.start, event.doc, 45) : { x: event.doc.x, y: event.doc.y }
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (!ctx || !drag) return
      const current = drag
      drag = null
      ctx.view.requestOverlay()
      current.end = event.shift ? snapToAngle(current.start, event.doc, 45) : { x: event.doc.x, y: event.doc.y }
      if (!(distance(current.start, current.end) >= 0.5)) return
      const context = ctx
      jobs.run(() => apply(context, current.target, current.start, current.end))
        .catch((error) => reportError(context, error, 'The gradient could not be drawn.'))
    },
    pointerCancel(): void {
      drag = null
      ctx?.view.requestOverlay()
    },
    keyDown(event: KeyboardEvent): boolean {
      if (event.key === 'Escape' && drag) {
        drag = null
        ctx?.view.requestOverlay()
        return true
      }
      return false
    },
    keyUp: () => false,
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      if (!drag) return
      const a = toScreen(view, drag.start)
      const b = toScreen(view, drag.end)
      strokeContrastPath(context, [a, b], false, 6)
      drawHandle(context, a, 6)
      drawHandle(context, b, 6)
    },
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {
      drag = null
    },
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}
