// src/advanced/tools/hand.ts (WP5)
// Hand tool (H, or hold Space): drag pans the view; double-click fits the image on screen. Never edits the
// document, so it records no history.
import type { Point } from '../../imaging/types.ts'
import type { ToolContext, ToolPointerEvent } from '../types.ts'
import type { ClickStamp } from '../toolGeometry.ts'
import { isDoubleClick } from '../toolGeometry.ts'
import type { AdvancedToolController } from './shared.ts'
import { isToolContext } from './shared.ts'

export function createHandTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let last: Point | null = null
  let lastClick: ClickStamp | null = null

  const stop = () => {
    last = null
    ctx?.setCursor('grab')
  }

  return {
    id: 'hand',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      last = null
      lastClick = null
      context.setCursor('grab')
      context.setHint('Drag to move around the image. Double-click to fit it on screen.')
    },
    deactivate(): void {
      last = null
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button > 1) return
      const stamp = { time: event.time, x: event.screen.x, y: event.screen.y }
      if (event.button === 0 && isDoubleClick(lastClick, stamp)) {
        lastClick = null
        last = null
        ctx.view.fit()
        return
      }
      lastClick = stamp
      last = { x: event.screen.x, y: event.screen.y }
      ctx.setCursor('grabbing')
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx || !last) return
      if (event.buttons === 0) {
        stop()
        return
      }
      const dx = event.screen.x - last.x
      const dy = event.screen.y - last.y
      if (dx || dy) ctx.view.panBy(dx, dy)
      last = { x: event.screen.x, y: event.screen.y }
    },
    pointerUp(): void {
      stop()
    },
    pointerCancel(): void {
      stop()
    },
    keyDown: () => false,
    keyUp: () => false,
    drawOverlay(): void {},
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {},
  }
}
