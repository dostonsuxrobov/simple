// src/advanced/tools/lasso.ts (WP5)
// Lasso and Polygonal Lasso (L, Shift+L cycles).
//   - Lasso: drag a freehand outline; releasing closes it.
//   - Polygonal Lasso: click vertices (Shift snaps the segment to 45 degrees); double-click, Enter or a
//     click on the first vertex closes the polygon; Backspace / Delete removes the last vertex; Esc cancels.
//     An open polygon is a tool session (switching tools closes it rather than losing it).
//   - Selection modifiers as for the marquees (picked at the first press); options: feather, anti-alias.
// One history step per selection change; never marks the document modified.
import type { Point, SelectionOp } from '../../imaging/types.ts'
import type { HistoryIcon, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { createMaskBuffer, rasterizePolygon } from '../../imaging/mask.ts'
import type { ClickStamp } from '../toolGeometry.ts'
import {
  appendPathPoint,
  closesPolygon,
  isDegeneratePolygon,
  isDoubleClick,
  nudgeDelta,
  selectionOpFromModifiers,
  simplifyPath,
  snapToAngle,
} from '../toolGeometry.ts'
import type { AdvancedToolController } from './shared.ts'
import {
  commitSelectionShape,
  createJobTracker,
  createSpacePan,
  deselect,
  isToolContext,
  optionsOf,
  reportError,
  screenToDocDistance,
  strokeContrastPath,
  toScreen,
} from './shared.ts'
import { moveSelectionOutline } from './marquee.ts'

/** Close the polygon when clicking within this many CSS px of its first vertex. */
const CLOSE_TOLERANCE_PX = 7

interface LassoPath {
  readonly op: SelectionOp
  readonly points: Point[]
  hover: Point | null
}

export function createLassoTool(kind: 'free' | 'polygon'): AdvancedToolController {
  const id = kind === 'free' ? 'lasso' : 'lasso-polygon'
  const label = kind === 'free' ? 'Lasso' : 'Polygonal Lasso'
  const icon: HistoryIcon = id
  let ctx: ToolContext | null = null
  let path: LassoPath | null = null
  let lastClick: ClickStamp | null = null
  const jobs = createJobTracker()
  const pan = createSpacePan(() => ctx, () => ctx?.setCursor('crosshair'))

  const commit = async (current: LassoPath) => {
    if (!ctx) return
    const context = ctx
    const options = optionsOf(context).lasso
    const points = simplifyPath(current.points, 0.2)
    if (isDegeneratePolygon(points)) {
      if (current.op === 'replace') deselect(context)
      return
    }
    const state = context.store.getState()
    let mask = createMaskBuffer(state.width, state.height)
    const drawn = rasterizePolygon(mask, points, options.antiAlias, 'nonzero')
    if (!drawn) {
      if (current.op === 'replace') deselect(context)
      return
    }
    if (options.feather > 0) {
      await jobs.run(async () => {
        mask = await context.imaging.run('feather', { mask, radius: options.feather })
      })
    }
    commitSelectionShape(context, label, icon, mask, current.op)
  }

  const close = () => {
    if (!ctx || !path) return
    const current = path
    path = null
    ctx.view.requestOverlay()
    const context = ctx
    commit(current).catch((error) => reportError(context, error, 'The selection could not be made.'))
  }

  const cancel = () => {
    path = null
    pan.reset()
    ctx?.view.requestOverlay()
  }

  const tolerance = () => (ctx ? screenToDocDistance(ctx, CLOSE_TOLERANCE_PX) : 1)

  const vertexFor = (event: ToolPointerEvent): Point => {
    const last = path?.points[path.points.length - 1]
    return event.shift && last ? snapToAngle(last, event.doc, 45) : { x: event.doc.x, y: event.doc.y }
  }

  return {
    id,
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      path = null
      lastClick = null
      context.setCursor('crosshair')
      context.setHint(kind === 'free'
        ? 'Drag around the area to select. Shift adds, Alt subtracts.'
        : 'Click to add points; double-click, press Enter or click the first point to close. Backspace removes a point.')
    },
    deactivate(): void {
      // An open polygon is the user's work: close it instead of dropping it.
      if (kind === 'polygon' && path && path.points.length >= 3) close()
      else cancel()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || jobs.busy) return
      if (kind === 'polygon' && path && pan.pointerDown(event)) return
      if (event.button !== 0) return
      const state = ctx.store.getState()
      const stamp = { time: event.time, x: event.screen.x, y: event.screen.y }
      if (kind === 'free') {
        const mods = selectionOpFromModifiers(event.shift, event.alt, Boolean(state.selection), optionsOf(ctx).lasso.op)
        path = { op: mods.op, points: [{ x: event.doc.x, y: event.doc.y }], hover: null }
        ctx.view.requestOverlay()
        return
      }
      if (!path) {
        const mods = selectionOpFromModifiers(event.shift, event.alt, Boolean(state.selection), optionsOf(ctx).lasso.op)
        path = { op: mods.op, points: [{ x: event.doc.x, y: event.doc.y }], hover: null }
        lastClick = stamp
        ctx.view.requestOverlay()
        return
      }
      const double = isDoubleClick(lastClick, stamp)
      lastClick = stamp
      if (double || closesPolygon(path.points, event.doc, tolerance())) {
        close()
        return
      }
      appendPathPoint(path.points, vertexFor(event), 0.01)
      ctx.view.requestOverlay()
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx || !path) return
      if (pan.pointerMove(event)) return
      if (kind === 'free') {
        if (event.buttons === 0) return
        const step = Math.max(0.25, screenToDocDistance(ctx, 1))
        for (const sample of event.coalesced ?? []) appendPathPoint(path.points, sample.doc, step)
        appendPathPoint(path.points, event.doc, step)
      } else {
        path.hover = vertexFor(event)
      }
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (pan.pointerUp(event)) return
      if (!ctx || !path || kind !== 'free') return
      appendPathPoint(path.points, event.doc, 0.01)
      if (path.points.length < 3) {
        const op = path.op
        path = null
        ctx.view.requestOverlay()
        if (op === 'replace') {
          try {
            deselect(ctx)
          } catch (error) {
            reportError(ctx, error, 'The selection could not be changed.')
          }
        }
        return
      }
      close()
    },
    pointerCancel(): void {
      if (kind === 'free') cancel()
    },
    keyDown(event: KeyboardEvent): boolean {
      if (!ctx) return false
      if (path) {
        if (event.key === 'Escape') {
          cancel()
          return true
        }
        if (kind === 'polygon' && event.key === 'Enter') {
          if (path.points.length >= 3) close()
          else cancel()
          return true
        }
        if (kind === 'polygon' && (event.key === 'Backspace' || event.key === 'Delete')) {
          path.points.pop()
          if (!path.points.length) cancel()
          else ctx.view.requestOverlay()
          return true
        }
        if (kind === 'polygon' && pan.keyDown(event)) return true
        return false
      }
      const delta = nudgeDelta(event.key, event.shiftKey)
      if (delta && !event.ctrlKey && !event.altKey && !event.metaKey) {
        try {
          return moveSelectionOutline(ctx, delta.x, delta.y, true)
        } catch (error) {
          reportError(ctx, error, 'The selection could not be moved.')
          return true
        }
      }
      return false
    },
    keyUp(event: KeyboardEvent): boolean {
      return pan.held ? pan.keyUp(event) : false
    },
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      if (!path || !path.points.length) return
      const points = path.points.map((p) => toScreen(view, p))
      if (kind === 'polygon' && path.hover) points.push(toScreen(view, path.hover))
      if (points.length >= 2) strokeContrastPath(context, points, kind === 'free')
      if (kind === 'polygon') {
        const first = points[0]
        const nearFirst = path.hover && path.points.length >= 3 && closesPolygon(path.points, path.hover, tolerance())
        context.save()
        try {
          context.setLineDash([])
          context.lineWidth = 1
          context.fillStyle = nearFirst ? '#111111' : '#ffffff'
          context.strokeStyle = '#111111'
          context.beginPath()
          context.arc(first.x, first.y, nearFirst ? 5 : 3.5, 0, Math.PI * 2)
          context.fill()
          context.stroke()
        } finally {
          context.restore()
        }
      }
    },
    hasSession: () => kind === 'polygon' && path !== null,
    commitSession(): void {
      if (path && path.points.length >= 3) close()
      else cancel()
    },
    cancelSession(): void {
      cancel()
    },
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}
