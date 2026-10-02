// src/advanced/tools/bucket.ts (WP5)
// Paint Bucket (G, Shift+G cycles with Gradient): click fills the region similar to the clicked pixel
// (tolerance, contiguous, anti-alias; sample the current layer or all layers) with the foreground colour,
// using the tool's opacity and blend mode, inside the selection, keeping alpha on transparency-locked
// layers. Fills a layer mask with the foreground grey when the mask is targeted. Alt+click picks the
// foreground colour. One history step.
import type { PixelBuffer } from '../../imaging/types.ts'
import type { ToolContext, ToolPointerEvent } from '../types.ts'
import { cropMask, maskBounds } from '../../imaging/mask.ts'
import { pixelEditBlocker } from '../document.ts'
import type { AdvancedToolController, PaintTarget } from './shared.ts'
import {
  blendMaskValues,
  blendPixels,
  createJobTracker,
  findLayer,
  grayOf,
  intersect,
  isToolContext,
  optionsOf,
  readLayerRect,
  readMaskRect,
  reportError,
  resolvePaintTarget,
  sampleDocumentPixels,
  writeRegion,
} from './shared.ts'
import { sampleColorAt } from './eyedropper.ts'

/** The pixels a flood fill compares on a mask target: the mask values as grey. */
function maskAsPixels(target: PaintTarget, width: number, height: number): PixelBuffer {
  const mask = target.layer.mask
  const data = new Uint8ClampedArray(width * height * 4)
  if (!mask) return { width, height, data }
  const values = readMaskRect(mask, { x: 0, y: 0, width, height }).data
  for (let i = 0, p = 0; i < values.length; i += 1, p += 4) {
    data[p] = values[i]
    data[p + 1] = values[i]
    data[p + 2] = values[i]
    data[p + 3] = 255
  }
  return { width, height, data }
}

export function createBucketTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  const jobs = createJobTracker(() => ctx?.setCursor(jobs.busy ? 'progress' : 'crosshair'))

  const fill = async (context: ToolContext, target: PaintTarget, seed: { x: number; y: number }) => {
    const state = context.store.getState()
    const options = optionsOf(context).bucket
    const color = context.editor.getState().foreground
    const src = target.channels === 1 ? maskAsPixels(target, state.width, state.height) : await sampleDocumentPixels(context, state, options.sample)
    const flood = await context.imaging.run('flood', {
      src,
      seed,
      options: { tolerance: options.tolerance, contiguous: options.contiguous, antiAlias: options.antiAlias, compareAlpha: true },
    })
    const now = context.store.getState()
    if (now.width !== state.width || now.height !== state.height) return
    const layer = findLayer(now, target.layer.id)
    if (!layer || pixelEditBlocker(layer, target.target)) return
    let region = maskBounds(flood)
    if (region && target.clip) region = intersect(region, target.clip)
    if (region && now.selection) region = intersect(region, now.selection.bounds)
    if (!region) return
    const rect = region
    const coverage = cropMask(flood, rect).data
    if (now.selection) {
      const selection = cropMask(now.selection.mask, rect).data
      for (let i = 0; i < coverage.length; i += 1) coverage[i] = Math.round((coverage[i] * selection[i]) / 255)
    }
    context.store.transact('Paint Bucket', 'paint-bucket', (tx) => {
      const editor = tx.editPixels(layer.id, target.target)
      if (target.channels === 1 && layer.mask) {
        const before = readMaskRect(layer.mask, rect)
        const values = new Uint8Array(before.data.length).fill(grayOf(color))
        writeRegion(editor, target, rect, blendMaskValues(before, values, null, options.opacity, coverage))
        return
      }
      const before = readLayerRect(layer, rect)
      const solid = new Uint8ClampedArray(rect.width * rect.height * 4)
      for (let p = 0; p < solid.length; p += 4) {
        solid[p] = color.r
        solid[p + 1] = color.g
        solid[p + 2] = color.b
        solid[p + 3] = 255
      }
      const result = blendPixels(before, { width: rect.width, height: rect.height, data: solid }, {
        opacity: options.opacity,
        mode: options.blendMode,
        coverage,
        preserveAlpha: target.preserveAlpha,
      })
      writeRegion(editor, target, rect, result)
    })
  }

  return {
    id: 'paint-bucket',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      context.setCursor('crosshair')
      context.setHint('Click to fill similar colours with the foreground colour. Alt-click picks a colour.')
    },
    deactivate(): void {
      ctx?.setHint(null)
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || event.button !== 0 || jobs.busy) return
      const context = ctx
      if (event.alt) {
        sampleColorAt(context, event, false)
        return
      }
      const state = context.store.getState()
      const x = Math.floor(event.doc.x)
      const y = Math.floor(event.doc.y)
      if (x < 0 || y < 0 || x >= state.width || y >= state.height) return
      const target = resolvePaintTarget(context)
      if (!target) return
      jobs.run(() => fill(context, target, { x, y })).catch((error) => reportError(context, error, 'The area could not be filled.'))
    },
    pointerMove(): void {},
    pointerUp(): void {},
    pointerCancel(): void {},
    keyDown: () => false,
    keyUp: () => false,
    drawOverlay(): void {},
    hasSession: () => false,
    commitSession(): void {},
    cancelSession(): void {},
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}
