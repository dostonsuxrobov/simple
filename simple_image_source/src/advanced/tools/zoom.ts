// src/advanced/tools/zoom.ts (WP5)
// Zoom tool (Z): click zooms in one step of Photoshop's ladder at the pointer, Alt+click zooms out, dragging
// right / left zooms continuously about the press point (scrubby zoom), double-click shows 100%. Records no
// history.
import type { Point } from '../../imaging/types.ts'
import type { ToolContext, ToolPointerEvent } from '../types.ts'
import type { ClickStamp } from '../toolGeometry.ts'
import { isDoubleClick, scrubbyZoom } from '../toolGeometry.ts'
import { stepZoom } from '../viewport.ts'
import type { AdvancedToolController } from './shared.ts'
import { isToolContext } from './shared.ts'

const SCRUB_THRESHOLD = 4

export function createZoomTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let drag: { start: Point; zoom: number; scrubbing: boolean; alt: boolean } | null = null
  let lastClick: ClickStamp | null = null
  let altDown = false

  const cursor = () => ctx?.setCursor(altDown ? 'zoom-out' : 'zoom-in')

  return {
    id: 'zoom',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      drag = null
      lastClick = null
      altDown = false
      cursor()
      context.setHint('Click to zoom in, Alt-click to zoom out, drag left or right to zoom smoothly. Double-click for 100%.')
    },
    deactivate(): void {
      drag = null
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0) return
      altDown = event.alt
      const stamp = { time: event.time, x: event.screen.x, y: event.screen.y }
      if (isDoubleClick(lastClick, stamp)) {
        lastClick = null
        drag = null
        ctx.view.actualPixels()
        return
      }
      lastClick = stamp
      drag = { start: { x: event.screen.x, y: event.screen.y }, zoom: ctx.view.getView().zoom, scrubbing: false, alt: event.alt }
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx) return
      if (!drag) {
        if (altDown !== event.alt) {
          altDown = event.alt
          cursor()
        }
        return
      }
      const dx = event.screen.x - drag.start.x
      if (!drag.scrubbing && Math.abs(dx) < SCRUB_THRESHOLD) return
      drag.scrubbing = true
      lastClick = null
      ctx.view.zoomAt(scrubbyZoom(drag.zoom, dx, ctx.view.getViewportSize().dpr), drag.start)
    },
    pointerUp(event: ToolPointerEvent): void {
      if (!ctx || !drag) return
      const current = drag
      drag = null
      if (current.scrubbing) return
      const direction = event.alt || current.alt ? -1 : 1
      const view = ctx.view.getView()
      ctx.view.zoomAt(stepZoom(view.zoom, ctx.view.getViewportSize().dpr, direction), { x: event.screen.x, y: event.screen.y })
    },
    pointerCancel(): void {
      drag = null
    },
    keyDown(event: KeyboardEvent): boolean {
      if (event.key === 'Alt' && !altDown) {
        altDown = true
        cursor()
      }
      return false
    },
    keyUp(event: KeyboardEvent): boolean {
      if (event.key === 'Alt' && altDown) {
        altDown = false
        cursor()
      }
      return false
    },
    drawOverlay(): void {},
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {},
  }
}
