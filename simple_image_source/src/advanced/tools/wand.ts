// src/advanced/tools/wand.ts (WP5)
// Magic Wand (W): click selects pixels similar to the clicked one (tolerance, contiguous, anti-alias;
// sample the current layer or all layers). Selection modifiers as for the marquees. The flood fill runs in
// the imaging worker. One history step; never marks the document modified.
import type { ToolContext, ToolPointerEvent } from '../types.ts'
import { nudgeDelta, selectionOpFromModifiers } from '../toolGeometry.ts'
import type { AdvancedToolController } from './shared.ts'
import {
  commitSelectionShape,
  createJobTracker,
  isToolContext,
  optionsOf,
  reportError,
  sampleDocumentPixels,
} from './shared.ts'
import { moveSelectionOutline } from './marquee.ts'

export function createWandTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  const jobs = createJobTracker(() => ctx?.setCursor(jobs.busy ? 'progress' : 'crosshair'))

  return {
    id: 'magic-wand',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      context.setCursor('crosshair')
      context.setHint('Click to select similar colours. Shift adds, Alt subtracts. Tolerance is in the options bar.')
    },
    deactivate(): void {
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0 || jobs.busy) return
      const context = ctx
      const state = context.store.getState()
      const x = Math.floor(event.doc.x)
      const y = Math.floor(event.doc.y)
      if (x < 0 || y < 0 || x >= state.width || y >= state.height) return
      const options = optionsOf(context).wand
      const { op } = selectionOpFromModifiers(event.shift, event.alt, Boolean(state.selection), options.op)
      jobs.run(async () => {
        const src = await sampleDocumentPixels(context, state, options.sample)
        const mask = await context.imaging.run('flood', {
          src,
          seed: { x, y },
          options: { tolerance: options.tolerance, contiguous: options.contiguous, antiAlias: options.antiAlias, compareAlpha: true },
        })
        const now = context.store.getState()
        if (now.width !== state.width || now.height !== state.height) return
        commitSelectionShape(context, 'Magic Wand', 'magic-wand', mask, op)
      }).catch((error) => reportError(context, error, 'The Magic Wand could not make a selection.'))
    },
    pointerMove(): void {},
    pointerUp(): void {},
    pointerCancel(): void {},
    keyDown(event: KeyboardEvent): boolean {
      if (!ctx) return false
      const delta = nudgeDelta(event.key, event.shiftKey)
      if (!delta || event.ctrlKey || event.altKey || event.metaKey) return false
      try {
        return moveSelectionOutline(ctx, delta.x, delta.y, true)
      } catch (error) {
        reportError(ctx, error, 'The selection could not be moved.')
        return true
      }
    },
    keyUp: () => false,
    drawOverlay(): void {},
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {},
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}
