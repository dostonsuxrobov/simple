// src/advanced/tools/shape.ts (WP5)
// Shape tool (U, Shift+U cycles rectangle / ellipse / line / arrow): drag draws a new shape layer above the
// active layer with the options bar's fill, stroke, stroke width and corner radius. Shift draws a square or
// circle (lines snap to 45 degrees), Alt draws from the centre. Rectangles and ellipses snap to whole pixels
// so their edges stay crisp. The layer stays editable (a vector spec with a transform). One history step.
import type { Affine, Point, Rect } from '../../imaging/types.ts'
import type { ShapeKind, ShapeSpec, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { createShapeLayer, nextLayerName } from '../document.ts'
import { drawShape } from '../../shared/vector.ts'
import { constrainedRect, lineFromDrag } from '../toolGeometry.ts'
import type { AdvancedToolController } from './shared.ts'
import {
  drawLabel,
  isToolContext,
  layerIndexOf,
  optionsOf,
  reportError,
  toScreen,
} from './shared.ts'

const LABELS: Readonly<Record<ShapeKind, { readonly history: string; readonly layer: string }>> = Object.freeze({
  rectangle: { history: 'Rectangle Tool', layer: 'Rectangle' },
  ellipse: { history: 'Ellipse Tool', layer: 'Ellipse' },
  line: { history: 'Line Tool', layer: 'Line' },
  arrow: { history: 'Arrow Tool', layer: 'Arrow' },
})

interface ShapeDrag {
  readonly anchor: Point
  current: Point
  readonly startScreen: Point
  moved: boolean
  shift: boolean
  alt: boolean
}

/** The shape spec a drag describes (shape-local geometry, translation in the transform). */
export function shapeSpecFromDrag(options: { readonly kind: ShapeKind; readonly fill: ShapeSpec['fill']; readonly stroke: ShapeSpec['stroke']; readonly strokeWidth: number; readonly cornerRadius: number },
  anchor: Point, point: Point, modifiers: { readonly shift: boolean; readonly alt: boolean }, lineColor: ShapeSpec['stroke']): ShapeSpec | null {
  const base = { fill: options.fill, stroke: options.stroke, strokeWidth: Math.max(0, options.strokeWidth), arrowHeads: 'none' as const }
  if (options.kind === 'line' || options.kind === 'arrow') {
    const { start, end } = lineFromDrag(anchor, point, { snap: modifiers.shift, fromCenter: modifiers.alt })
    if (Math.hypot(end.x - start.x, end.y - start.y) < 1) return null
    const transform: Affine = [1, 0, 0, 1, start.x, start.y]
    return {
      ...base,
      kind: options.kind,
      x1: 0,
      y1: 0,
      x2: end.x - start.x,
      y2: end.y - start.y,
      // Lines draw with the stroke colour (or the fill); fall back to the foreground so they never vanish.
      stroke: options.stroke ?? options.fill ?? lineColor,
      fill: null,
      strokeWidth: Math.max(1, options.strokeWidth),
      cornerRadius: 0,
      arrowHeads: options.kind === 'arrow' ? 'end' : 'none',
      transform,
    }
  }
  const raw: Rect = constrainedRect(anchor, point, { square: modifiers.shift, fromCenter: modifiers.alt })
  const x0 = Math.round(raw.x)
  const y0 = Math.round(raw.y)
  const width = Math.round(raw.x + raw.width) - x0
  const height = Math.round(raw.y + raw.height) - y0
  if (width < 1 || height < 1) return null
  return {
    ...base,
    kind: options.kind,
    x1: 0,
    y1: 0,
    x2: width,
    y2: height,
    cornerRadius: options.kind === 'rectangle' ? Math.max(0, options.cornerRadius) : 0,
    fill: options.fill,
    stroke: options.stroke,
    transform: [1, 0, 0, 1, x0, y0],
  }
}

export function createShapeTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let drag: ShapeDrag | null = null

  const specNow = (): ShapeSpec | null => {
    if (!ctx || !drag) return null
    const foreground = ctx.editor.getState().foreground
    return shapeSpecFromDrag(optionsOf(ctx).shape, drag.anchor, drag.current, { shift: drag.shift, alt: drag.alt }, { ...foreground, a: 255 })
  }

  const create = (spec: ShapeSpec) => {
    if (!ctx) return
    const context = ctx
    const labels = LABELS[spec.kind]
    const state = context.store.getState()
    const layer = createShapeLayer(nextLayerName(state, labels.layer), spec)
    context.store.transact(labels.history, 'shape', (tx) => {
      const active = layerIndexOf(tx.state, tx.state.activeLayerId)
      tx.insertLayer(layer, active >= 0 ? active + 1 : tx.state.layers.length)
      tx.setActiveLayer(layer.id)
    })
  }

  return {
    id: 'shape',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      drag = null
      context.setCursor('crosshair')
      context.setHint('Drag to draw a shape. Shift keeps proportions (45° for lines), Alt draws from the centre.')
    },
    deactivate(): void {
      drag = null
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0) return
      drag = {
        anchor: { x: event.doc.x, y: event.doc.y },
        current: { x: event.doc.x, y: event.doc.y },
        startScreen: { x: event.screen.x, y: event.screen.y },
        moved: false,
        shift: event.shift,
        alt: event.alt,
      }
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx || !drag) return
      if (event.buttons === 0) {
        drag = null
        ctx.view.requestOverlay()
        return
      }
      drag.shift = event.shift
      drag.alt = event.alt
      drag.current = { x: event.doc.x, y: event.doc.y }
      if (!drag.moved && Math.hypot(event.screen.x - drag.startScreen.x, event.screen.y - drag.startScreen.y) >= 3) drag.moved = true
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (!ctx || !drag) return
      drag.current = { x: event.doc.x, y: event.doc.y }
      drag.shift = event.shift
      drag.alt = event.alt
      const moved = drag.moved
      const spec = moved ? specNow() : null
      drag = null
      ctx.view.requestOverlay()
      if (!spec) return
      try {
        create(spec)
      } catch (error) {
        reportError(ctx, error, 'The shape could not be drawn.')
      }
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
      if (!drag || !drag.moved) return
      const spec = specNow()
      if (!spec) return
      context.save()
      try {
        context.translate(view.offsetX, view.offsetY)
        context.scale(view.zoom, view.zoom)
        drawShape(context, spec)
      } catch {
        // Preview only.
      } finally {
        context.restore()
      }
      const end = toScreen(view, drag.current)
      const w = Math.abs(spec.x2 - spec.x1)
      const h = Math.abs(spec.y2 - spec.y1)
      const text = spec.kind === 'line' || spec.kind === 'arrow' ? `${Math.round(Math.hypot(w, h))} px` : `${Math.round(w)} × ${Math.round(h)}`
      drawLabel(context, text, { x: end.x + 14, y: end.y + 14 })
    },
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {
      drag = null
    },
  }
}
