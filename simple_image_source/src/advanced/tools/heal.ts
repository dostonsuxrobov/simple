// src/advanced/tools/heal.ts (WP5)
// Spot Healing Brush (J): paint over a blemish; on release the painted area is healed from a matching
// nearby source with a seamless clone (imaging/inpaint.ts healRegion, run in the imaging worker). Sample
// the current layer, the composite below, or all layers (healing onto an empty layer). Healing stays inside
// the selection and keeps alpha on transparency-locked layers. One history step per stroke.
import type { BrushTip, Dab, IntRect, MaskBuffer, PixelBuffer, Point } from '../../imaging/types.ts'
import type { ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { accumulateDab, placeDabs } from '../../imaging/brush.ts'
import type { DabCarry } from '../../imaging/brush.ts'
import { healMargin } from '../../imaging/inpaint.ts'
import { cropMask } from '../../imaging/mask.ts'
import { compositeRect } from '../composite.ts'
import { pixelEditBlocker } from '../document.ts'
import type { AdvancedToolController, PaintTarget } from './shared.ts'
import {
  canvasRect,
  createJobTracker,
  findLayer,
  intersect,
  isToolContext,
  optionsOf,
  readLayerRect,
  reportError,
  resolvePaintTarget,
  roundOut,
  strokeSamples,
  toScreen,
  union,
  writeRegion,
} from './shared.ts'
import { brushCursor, drawBrushCursor } from './brush.ts'

interface HealStroke {
  readonly target: PaintTarget
  readonly tip: BrushTip
  readonly dabs: Dab[]
  readonly carry: DabCarry
  bounds: IntRect | null
}

let healSeed = 1

/** Pixels that differ between two same-size buffers, as a bounding rectangle in buffer coordinates. */
function changedBounds(a: PixelBuffer, b: PixelBuffer): IntRect | null {
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -1
  let y1 = -1
  const width = a.width
  for (let i = 0, p = 0; p < a.data.length; i += 1, p += 4) {
    if (a.data[p] === b.data[p] && a.data[p + 1] === b.data[p + 1] && a.data[p + 2] === b.data[p + 2] && a.data[p + 3] === b.data[p + 3]) continue
    const x = i % width
    const y = (i - x) / width
    if (x < x0) x0 = x
    if (y < y0) y0 = y
    if (x > x1) x1 = x
    if (y > y1) y1 = y
  }
  return x1 < 0 ? null : { x: x0, y: y0, width: x1 - x0 + 1, height: y1 - y0 + 1 }
}

export function createHealTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let stroke: HealStroke | null = null
  let hover: Point | null = null
  let cursor = ''
  const jobs = createJobTracker(() => {
    if (!ctx) return
    if (jobs.busy) ctx.setCursor('progress')
    else ctx.setCursor(cursor)
    ctx.setHint(jobs.busy ? 'Healing…' : 'Paint over a blemish to heal it.')
  })

  const addSamples = (current: HealStroke, event: ToolPointerEvent, final: boolean) => {
    const dabs = placeDabs(strokeSamples(event), current.tip, 0.15, final ? 0 : 0.1, current.carry, {})
    for (const dab of dabs) {
      current.dabs.push(dab)
      const half = dab.diameter / 2 + 1
      current.bounds = union(current.bounds, roundOut({ x: dab.x - half, y: dab.y - half, width: 2 * half, height: 2 * half }))
    }
  }

  const heal = async (context: ToolContext, current: HealStroke) => {
    const state = context.store.getState()
    const canvas = canvasRect(state)
    const holeBounds = current.bounds ? intersect(current.bounds, canvas) : null
    if (!holeBounds) return
    const margin = healMargin(holeBounds)
    const region = intersect({ x: holeBounds.x - margin, y: holeBounds.y - margin, width: holeBounds.width + 2 * margin, height: holeBounds.height + 2 * margin }, canvas)
    if (!region) return
    // The painted area as a hole mask over the region.
    const coverage = new Float32Array(region.width * region.height)
    for (const dab of current.dabs) accumulateDab(coverage, region.width, region.height, region.x, region.y, dab, current.tip)
    const hole: MaskBuffer = { width: region.width, height: region.height, data: new Uint8Array(coverage.length) }
    const selection = state.selection ? cropMask(state.selection.mask, region).data : null
    let any = false
    for (let i = 0; i < coverage.length; i += 1) {
      if (coverage[i] > 0.02 && (!selection || selection[i] > 0)) {
        hole.data[i] = 255
        any = true
      }
    }
    if (!any) return
    const layer = current.target.layer
    const sample = optionsOf(context).heal.sample
    const src = sample === 'current'
      ? readLayerRect(layer, region)
      : compositeRect(state, region, sample === 'all' ? {} : { belowLayerId: layer.id, includeBelowLayer: true })
    healSeed += 1
    const healed = await context.imaging.run('heal', { src, hole, options: { seed: healSeed } }, { transfer: false })
    const changed = changedBounds(src, healed)
    if (!changed) return
    const now = context.store.getState()
    if (now.width !== state.width || now.height !== state.height) return
    const live = findLayer(now, layer.id)
    if (!live || live.kind !== 'raster' || pixelEditBlocker(live, 'pixels')) return
    const rect = { x: region.x + changed.x, y: region.y + changed.y, width: changed.width, height: changed.height }
    context.store.transact('Spot Healing Brush', 'spot-healing', (tx) => {
      const editor = tx.editPixels(live.id, 'pixels')
      const before = readLayerRect(live, rect)
      const out = new Uint8ClampedArray(before.data)
      const preserveAlpha = current.target.preserveAlpha
      for (let y = 0; y < rect.height; y += 1) {
        for (let x = 0; x < rect.width; x += 1) {
          const ri = (changed.y + y) * region.width + (changed.x + x)
          const rp = ri * 4
          const p = (y * rect.width + x) * 4
          if (src.data[rp] === healed.data[rp] && src.data[rp + 1] === healed.data[rp + 1] && src.data[rp + 2] === healed.data[rp + 2]
            && src.data[rp + 3] === healed.data[rp + 3]) continue
          const k = selection ? selection[ri] / 255 : 1
          if (k <= 0) continue
          out[p] = Math.round(before.data[p] + (healed.data[rp] - before.data[p]) * k)
          out[p + 1] = Math.round(before.data[p + 1] + (healed.data[rp + 1] - before.data[p + 1]) * k)
          out[p + 2] = Math.round(before.data[p + 2] + (healed.data[rp + 2] - before.data[p + 2]) * k)
          if (!preserveAlpha) out[p + 3] = Math.round(before.data[p + 3] + (healed.data[rp + 3] - before.data[p + 3]) * k)
        }
      }
      writeRegion(editor, current.target, rect, { width: rect.width, height: rect.height, data: out })
    })
  }

  return {
    id: 'spot-healing',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      stroke = null
      hover = null
      cursor = brushCursor(context.view.getView(), optionsOf(context).heal.size)
      context.setCursor(cursor)
      context.setHint('Paint over a blemish to heal it.')
    },
    deactivate(): void {
      // A stroke in progress is healed rather than dropped.
      const current = stroke
      stroke = null
      if (current && ctx) {
        const context = ctx
        jobs.run(() => heal(context, current)).catch((error) => reportError(context, error, 'The area could not be healed.'))
      }
      hover = null
      ctx?.view.requestOverlay()
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0 || jobs.busy) return
      const state = ctx.store.getState()
      if (state.editTarget === 'mask') {
        ctx.host.notify('Spot healing works on layer pixels. Select the layer thumbnail instead of its mask.', 'error')
        return
      }
      const target = resolvePaintTarget(ctx, { allowMask: false })
      if (!target) return
      const options = optionsOf(ctx).heal
      stroke = {
        target,
        tip: { diameter: Math.max(1, options.size), hardness: options.hardness, roundness: 1, angle: 0 },
        dabs: [],
        carry: { distance: 0, last: null },
        bounds: null,
      }
      hover = { x: event.doc.x, y: event.doc.y }
      addSamples(stroke, event, false)
      ctx.view.requestOverlay()
    },
    pointerMove(event: ToolPointerEvent): void {
      if (!ctx) return
      hover = { x: event.doc.x, y: event.doc.y }
      if (stroke && event.buttons !== 0) addSamples(stroke, event, false)
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (!ctx || !stroke) return
      const current = stroke
      stroke = null
      addSamples(current, event, true)
      ctx.view.requestOverlay()
      const context = ctx
      jobs.run(() => heal(context, current)).catch((error) => reportError(context, error, 'The area could not be healed.'))
    },
    pointerCancel(): void {
      const current = stroke
      stroke = null
      if (current && ctx) {
        const context = ctx
        jobs.run(() => heal(context, current)).catch((error) => reportError(context, error, 'The area could not be healed.'))
      }
      ctx?.view.requestOverlay()
    },
    keyDown: () => false,
    keyUp: () => false,
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      if (!ctx) return
      const size = optionsOf(ctx).heal.size
      if (stroke && stroke.dabs.length) {
        context.save()
        try {
          context.beginPath()
          for (const dab of stroke.dabs) {
            const center = toScreen(view, dab)
            const radius = Math.max(0.5, (dab.diameter / 2) * view.zoom)
            context.moveTo(center.x + radius, center.y)
            context.arc(center.x, center.y, radius, 0, Math.PI * 2)
          }
          context.fillStyle = 'rgba(30, 30, 30, 0.38)'
          context.fill('nonzero')
        } finally {
          context.restore()
        }
      }
      if (!jobs.busy) {
        const next = brushCursor(view, size)
        if (next !== cursor) {
          cursor = next
          ctx.setCursor(next)
        }
      }
      drawBrushCursor(context, view, hover, size)
    },
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {},
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}
