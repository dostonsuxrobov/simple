// src/advanced/tools/eyedropper.ts (WP5)
// Eyedropper (I): click (or drag) sets the foreground colour, Alt sets the background colour. Sample size
// 1 / 3 / 5 / 11 px and source (current layer, current and below, all layers) come from the options bar.
// While sampling, a ring shows the new colour (top) against the current one (bottom). No history.
import type { Point, Rgb8 } from '../../imaging/types.ts'
import type { ToolContext, ToolPointerEvent } from '../types.ts'
import type { AdvancedToolController } from './shared.ts'
import { isToolContext, optionsOf } from './shared.ts'

function css(color: Rgb8): string {
  return `rgb(${color.r}, ${color.g}, ${color.b})`
}

/** Samples the document at the event and stores it as the foreground (or background) colour. */
export function sampleColorAt(ctx: ToolContext, event: Pick<ToolPointerEvent, 'doc'>, background: boolean): Rgb8 | null {
  const state = ctx.store.getState()
  const { x, y } = event.doc
  if (!(x >= 0 && y >= 0 && x < state.width && y < state.height)) return null
  const { size, sample } = optionsOf(ctx).eyedropper
  const color = ctx.compositor.sample(x, y, size, sample)
  if (!color || color.a === 0) return null
  const rgb = { r: color.r, g: color.g, b: color.b }
  ctx.editor.update(background ? { background: rgb } : { foreground: rgb })
  return rgb
}

export function createEyedropperTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let sampling: { background: boolean; previous: Rgb8; current: Rgb8 | null; screen: Point } | null = null

  const pick = (event: ToolPointerEvent) => {
    if (!ctx || !sampling) return
    sampling.screen = { x: event.screen.x, y: event.screen.y }
    const color = sampleColorAt(ctx, event, sampling.background)
    if (color) sampling.current = color
    ctx.view.requestOverlay()
  }

  return {
    id: 'eyedropper',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      sampling = null
      context.setCursor('crosshair')
      context.setHint('Click to pick the foreground colour. Alt-click picks the background colour.')
    },
    deactivate(): void {
      sampling = null
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0) return
      const editor = ctx.editor.getState()
      sampling = { background: event.alt, previous: event.alt ? editor.background : editor.foreground, current: null, screen: event.screen }
      pick(event)
    },
    pointerMove(event: ToolPointerEvent): void {
      if (sampling && event.buttons !== 0) pick(event)
    },
    pointerUp(): void {
      if (!sampling) return
      sampling = null
      ctx?.view.requestOverlay()
    },
    pointerCancel(): void {
      sampling = null
      ctx?.view.requestOverlay()
    },
    keyDown: () => false,
    keyUp: () => false,
    drawOverlay(context: CanvasRenderingContext2D): void {
      if (!sampling) return
      const { x, y } = sampling.screen
      const outer = 34
      const inner = 22
      context.save()
      try {
        context.setLineDash([])
        const halves: [number, number, Rgb8][] = [[Math.PI, 2 * Math.PI, sampling.current ?? sampling.previous], [0, Math.PI, sampling.previous]]
        for (const [from, to, color] of halves) {
          context.beginPath()
          context.arc(x, y, outer, from, to)
          context.arc(x, y, inner, to, from, true)
          context.closePath()
          context.fillStyle = css(color)
          context.fill()
        }
        context.lineWidth = 1
        context.strokeStyle = 'rgba(17, 17, 17, 0.6)'
        context.beginPath()
        context.arc(x, y, outer + 0.5, 0, Math.PI * 2)
        context.stroke()
        context.strokeStyle = 'rgba(255, 255, 255, 0.85)'
        context.beginPath()
        context.arc(x, y, inner - 0.5, 0, Math.PI * 2)
        context.stroke()
      } finally {
        context.restore()
      }
    },
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {},
  }
}
