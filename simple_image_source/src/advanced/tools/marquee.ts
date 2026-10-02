// src/advanced/tools/marquee.ts (WP5)
// Rectangular and Elliptical Marquee (M, Shift+M cycles). Photoshop behaviour:
//   - with a selection, Shift / Alt / Shift+Alt at pointer-down add / subtract / intersect; otherwise
//     Shift draws a square or circle and Alt draws from the centre (keys used to pick the operation
//     constrain only after being released and pressed again);
//   - options: feather, anti-alias (ellipse; rectangles snap to whole pixels), normal / fixed ratio /
//     fixed size;
//   - a click without a drag deselects; dragging inside the selection (New mode) moves the outline;
//     arrow keys nudge the outline by 1 px (10 px with Shift).
// One history step per selection change; selection changes never mark the document modified.
import type { Point, Rect, SelectionOp } from '../../imaging/types.ts'
import type { HistoryIcon, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { createMaskBuffer, rasterizeEllipse, rasterizeRect } from '../../imaging/mask.ts'
import { translateSelection } from '../selection.ts'
import { marqueeRect, modifierState, nudgeDelta, selectionOpFromModifiers, snapRectToPixels } from '../toolGeometry.ts'
import type { AdvancedToolController } from './shared.ts'
import {
  commitSelectionShape,
  createJobTracker,
  deselect,
  drawLabel,
  isToolContext,
  nextSelectionVersion,
  optionsOf,
  reportError,
  selectionValueAt,
  strokeContrastPath,
  toScreen,
} from './shared.ts'

interface DrawDrag {
  readonly mode: 'draw'
  readonly anchor: Point
  current: Point
  readonly op: SelectionOp
  shiftLatched: boolean
  altLatched: boolean
  square: boolean
  fromCenter: boolean
  readonly startScreen: Point
  moved: boolean
}

interface MoveDrag {
  readonly mode: 'move'
  readonly anchor: Point
  dx: number
  dy: number
}

/** Ellipse outline as a polygon (screen or document points). */
export function ellipsePoints(rect: Rect, segments = 96): Point[] {
  const cx = rect.x + rect.width / 2
  const cy = rect.y + rect.height / 2
  const points: Point[] = []
  for (let i = 0; i < segments; i += 1) {
    const t = (i / segments) * Math.PI * 2
    points.push({ x: cx + (rect.width / 2) * Math.cos(t), y: cy + (rect.height / 2) * Math.sin(t) })
  }
  return points
}

/** Nudges (or moves) the selection outline by whole pixels; coalesced into one step per burst of key presses. */
export function moveSelectionOutline(ctx: ToolContext, dx: number, dy: number, coalesce: boolean): boolean {
  const state = ctx.store.getState()
  const selection = state.selection
  if (!selection || (!dx && !dy)) return false
  const next = translateSelection(selection, dx, dy, nextSelectionVersion(state))
  ctx.store.transact(coalesce ? 'Nudge Selection' : 'Move Selection', 'selection', (tx) => tx.setSelection(next), {
    affectsOutput: false,
    coalesceKey: coalesce ? 'nudge-selection' : undefined,
  })
  return true
}

export function createMarqueeTool(shape: 'rect' | 'ellipse'): AdvancedToolController {
  const id = shape === 'rect' ? 'marquee-rect' : 'marquee-ellipse'
  const label = shape === 'rect' ? 'Rectangular Marquee' : 'Elliptical Marquee'
  const icon: HistoryIcon = id
  let ctx: ToolContext | null = null
  let drag: DrawDrag | MoveDrag | null = null
  const jobs = createJobTracker()

  const currentRect = (d: DrawDrag): Rect => {
    if (!ctx) return { x: 0, y: 0, width: 0, height: 0 }
    const options = optionsOf(ctx).marquee
    const rect = marqueeRect(d.anchor, d.current, { square: d.square, fromCenter: d.fromCenter }, options)
    return shape === 'rect' ? snapRectToPixels(rect) : rect
  }

  const finishDraw = async (d: DrawDrag) => {
    if (!ctx) return
    const context = ctx
    const options = optionsOf(context).marquee
    const state = context.store.getState()
    if (!d.moved && options.style !== 'fixed-size') {
      if (d.op === 'replace') deselect(context)
      return
    }
    const rect = currentRect(d)
    if (!(rect.width > 0 && rect.height > 0)) {
      if (d.op === 'replace') deselect(context)
      return
    }
    let mask = createMaskBuffer(state.width, state.height)
    const drawn = shape === 'rect'
      ? rasterizeRect(mask, rect, false)
      : rasterizeEllipse(mask, rect, options.antiAlias)
    if (!drawn) {
      if (d.op === 'replace') deselect(context)
      return
    }
    if (options.feather > 0) {
      await jobs.run(async () => {
        mask = await context.imaging.run('feather', { mask, radius: options.feather })
      })
    }
    commitSelectionShape(context, label, icon, mask, d.op)
  }

  const updateHoverCursor = (event: ToolPointerEvent) => {
    if (!ctx) return
    const state = ctx.store.getState()
    const options = optionsOf(ctx).marquee
    const inside = Boolean(state.selection) && options.op === 'replace' && !event.shift && !event.alt
      && selectionValueAt(state.selection, event.doc.x, event.doc.y) > 0
    ctx.setCursor(inside ? 'move' : 'crosshair')
  }

  return {
    id,
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      drag = null
      context.setCursor('crosshair')
      context.setHint('Drag to select. Shift adds, Alt subtracts; Shift draws a square, Alt draws from the centre.')
    },
    deactivate(): void {
      drag = null
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0 || jobs.busy) return
      const state = ctx.store.getState()
      const options = optionsOf(ctx).marquee
      const mods = selectionOpFromModifiers(event.shift, event.alt, Boolean(state.selection), options.op)
      if (state.selection && mods.op === 'replace' && options.op === 'replace' && !event.shift && !event.alt
        && selectionValueAt(state.selection, event.doc.x, event.doc.y) > 0) {
        drag = { mode: 'move', anchor: { x: event.doc.x, y: event.doc.y }, dx: 0, dy: 0 }
        return
      }
      drag = {
        mode: 'draw',
        anchor: { x: event.doc.x, y: event.doc.y },
        current: { x: event.doc.x, y: event.doc.y },
        op: mods.op,
        shiftLatched: mods.shiftLatched,
        altLatched: mods.altLatched,
        square: event.shift && !mods.shiftLatched,
        fromCenter: event.alt && !mods.altLatched,
        startScreen: { x: event.screen.x, y: event.screen.y },
        moved: false,
      }
      ctx.view.requestOverlay()
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx) return
      if (!drag) {
        updateHoverCursor(event)
        return
      }
      if (drag.mode === 'move') {
        drag.dx = Math.round(event.doc.x - drag.anchor.x)
        drag.dy = Math.round(event.doc.y - drag.anchor.y)
        ctx.view.requestOverlay()
        return
      }
      const shift = modifierState(drag.shiftLatched, event.shift)
      const alt = modifierState(drag.altLatched, event.alt)
      drag.shiftLatched = shift.latched
      drag.altLatched = alt.latched
      drag.square = shift.active
      drag.fromCenter = alt.active
      drag.current = { x: event.doc.x, y: event.doc.y }
      if (!drag.moved && Math.hypot(event.screen.x - drag.startScreen.x, event.screen.y - drag.startScreen.y) >= 2) drag.moved = true
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (!ctx || !drag) return
      const current = drag
      drag = null
      ctx.view.requestOverlay()
      const context = ctx
      try {
        if (current.mode === 'move') {
          moveSelectionOutline(context, current.dx, current.dy, false)
          return
        }
        current.current = { x: event.doc.x, y: event.doc.y }
        finishDraw(current).catch((error) => reportError(context, error, 'The selection could not be made.'))
      } catch (error) {
        reportError(context, error, 'The selection could not be made.')
      }
    },
    pointerCancel(): void {
      drag = null
      ctx?.view.requestOverlay()
    },
    keyDown(event: KeyboardEvent): boolean {
      if (!ctx) return false
      if (event.key === 'Escape' && drag) {
        drag = null
        ctx.view.requestOverlay()
        return true
      }
      const delta = nudgeDelta(event.key, event.shiftKey)
      if (delta && !drag && !event.ctrlKey && !event.altKey && !event.metaKey) {
        try {
          return moveSelectionOutline(ctx, delta.x, delta.y, true)
        } catch (error) {
          reportError(ctx, error, 'The selection could not be moved.')
          return true
        }
      }
      return false
    },
    keyUp: () => false,
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      if (!ctx || !drag) return
      if (drag.mode === 'move') {
        const selection = ctx.store.getState().selection
        if (!selection) return
        const b = selection.bounds
        const corners = [
          { x: b.x + drag.dx, y: b.y + drag.dy },
          { x: b.x + b.width + drag.dx, y: b.y + drag.dy },
          { x: b.x + b.width + drag.dx, y: b.y + b.height + drag.dy },
          { x: b.x + drag.dx, y: b.y + b.height + drag.dy },
        ].map((p) => toScreen(view, p))
        strokeContrastPath(context, corners, true)
        return
      }
      if (!drag.moved) return
      const rect = currentRect(drag)
      const points = shape === 'rect'
        ? [{ x: rect.x, y: rect.y }, { x: rect.x + rect.width, y: rect.y }, { x: rect.x + rect.width, y: rect.y + rect.height }, { x: rect.x, y: rect.y + rect.height }]
        : ellipsePoints(rect)
      strokeContrastPath(context, points.map((p) => toScreen(view, p)), true)
      const corner = toScreen(view, drag.current)
      drawLabel(context, `${Math.round(rect.width)} × ${Math.round(rect.height)}`, { x: corner.x + 14, y: corner.y + 14 })
    },
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {
      drag = null
      ctx?.view.requestOverlay()
    },
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}
