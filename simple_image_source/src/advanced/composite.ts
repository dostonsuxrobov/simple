// src/advanced/composite.ts (WP3)
// The pure, per-tile layer compositor (design 5.5). Synchronous, deterministic and DOM-free: the display
// compositor, flatten/export, thumbnails, the eyedropper and Node tests all use it.
//
// Per block (a level tile clipped to the request), on straight RGBA, one rounding per layer:
//   out = transparent
//   for each unit bottom to top (a layer, or a clipping group = base + the clipped layers above it):
//     skip when the base is hidden (its clipped layers are hidden too)
//     single layer: draw(out, layer)
//     group:        G = transparent; draw(G, base, opacity 1, normal)        (base mask applied here)
//                   for each clipped C: draw(G, C) with G's alpha locked      (source-atop)
//                   blend(out, G, base.blendMode, base.opacity)
//   draw(adjustment): adjusted = kernel(dst); dst.rgb = mix(dst.rgb, B(dst.rgb, adjusted), mask * opacity);
//                     alpha unchanged
//   draw(raster/text/shape): src = level pixels of the layer (offset floor(o / 2^level)); coverage = layer mask;
//                     previews applied; blendInto(dst, src, mode, opacity, coverage)
// Dissolve keeps a source pixel when hash2(x, y, hash(layer id)) < its effective alpha (document
// coordinates, so the pattern stays put across tiles), then composites it opaque. Dithering adjustments
// (gradient maps) draw their noise per document row on tile-anchored indices for the same reason.
// The first layer drawn onto a still transparent block is copied rather than blended when that is exact
// (every mode but Dissolve reduces to a copy over transparency).
// Previews: 'layer-pixels' replaces the layer's pixels inside `rect`, given in document coordinates of
// pyramid level `level` (ignored when compositing another level).
// Requests are split into tile-aligned blocks of at most 256 x 256, so scratch memory stays bounded and any
// two requests covering the same pixels give the same bytes.
import type { AdjustmentKernel, AdjustmentSpec, BlendMode, IntRect, OpOptions, PixelBuffer, Rgba8 } from '../imaging/types.ts'
import type {
  AdjustmentLayer,
  CompositeOptions,
  CompositePreview,
  CompositeRectFn,
  DocumentState,
  Layer,
  LayerId,
  SampleSource,
  Selection,
  TiledSurface,
} from './types.ts'
import { TILE_SIZE } from './types.ts'
import { blendInto } from '../imaging/blend.ts'
import { compileAdjustment } from '../imaging/adjustments.ts'
import { progressReporter, throwIfAborted } from '../imaging/buffer.ts'
import { hash2, hashString, mixSeed } from '../imaging/random.ts'
import { MAX_LEVEL, levelOffset, levelSize, readMaskLevel, readSurfaceLevel } from './pyramid.ts'
import { tileIndex, toIntRect } from './tiles.ts'

export type CompositeDocument = Pick<DocumentState, 'width' | 'height' | 'layers'>

// ---------------------------------------------------------------------------------------------
// Scratch buffers (compositing is synchronous and never re-entrant, so module-level pools are safe)
// ---------------------------------------------------------------------------------------------

const BLOCK_PIXELS = TILE_SIZE * TILE_SIZE
const ACC = 0
const SRC = 1
const T1 = 2
const T0 = 3
const GROUP = 4
const COV = 0
const SEL = 1

const rgbaPool: Uint8ClampedArray[] = []
const maskPool: Uint8Array[] = []
/** One tile-wide row for dithering kernels (see runKernel). */
const ditherRow = new Uint8ClampedArray(TILE_SIZE * 4)

function rgbaScratch(slot: number, pixels: number): Uint8ClampedArray {
  let buffer = rgbaPool[slot]
  if (!buffer) {
    buffer = new Uint8ClampedArray(BLOCK_PIXELS * 4)
    rgbaPool[slot] = buffer
  }
  return buffer.subarray(0, pixels * 4)
}

function maskScratch(slot: number, pixels: number): Uint8Array {
  let buffer = maskPool[slot]
  if (!buffer) {
    buffer = new Uint8Array(BLOCK_PIXELS)
    maskPool[slot] = buffer
  }
  return buffer.subarray(0, pixels)
}

// ---------------------------------------------------------------------------------------------
// Layer helpers
// ---------------------------------------------------------------------------------------------

interface PixelSource {
  readonly surface: TiledSurface
  readonly offsetX: number
  readonly offsetY: number
}

/** Where a layer's pixels live: the raster surface, or the text/shape raster cache. Null for adjustments. */
export function layerPixelSource(layer: Layer): PixelSource | null {
  if (layer.kind === 'raster') return { surface: layer.surface, offsetX: layer.offsetX, offsetY: layer.offsetY }
  if (layer.kind === 'text' || layer.kind === 'shape') {
    return { surface: layer.raster.surface, offsetX: layer.raster.offsetX, offsetY: layer.raster.offsetY }
  }
  return null
}

interface LayerProps {
  readonly visible: boolean
  readonly opacity: number
  readonly blendMode: BlendMode
}

function clamp01(value: number): number {
  return value >= 1 ? 1 : value > 0 ? value : 0
}

function propsOf(layer: Layer, preview: CompositePreview | null): LayerProps {
  if (preview && preview.kind === 'layer-props' && preview.layerId === layer.id) {
    return {
      visible: preview.visible ?? layer.visible,
      opacity: clamp01(preview.opacity ?? layer.opacity),
      blendMode: preview.blendMode ?? layer.blendMode,
    }
  }
  return { visible: layer.visible, opacity: clamp01(layer.opacity), blendMode: layer.blendMode }
}

const kernels = new WeakMap<AdjustmentSpec, AdjustmentKernel>()

function kernelFor(spec: AdjustmentSpec): AdjustmentKernel {
  let kernel = kernels.get(spec)
  if (!kernel) {
    kernel = compileAdjustment(spec)
    kernels.set(spec, kernel)
  }
  return kernel
}

function seedOf(layer: Layer, level: number): number {
  return mixSeed(hashString(layer.id), level)
}

// ---------------------------------------------------------------------------------------------
// Plan (resolved once per request)
// ---------------------------------------------------------------------------------------------

interface PlanLayer {
  readonly layer: Layer
  readonly props: LayerProps
}

interface PlanUnit {
  readonly base: PlanLayer
  readonly clipped: readonly PlanLayer[]
}

interface Plan {
  readonly units: readonly PlanUnit[]
  /** onlyLayerId: the layer composited in isolation (raw pixels, no mask or opacity). */
  readonly only: Layer | null
  readonly onlyMode: boolean
}

function planComposite(doc: CompositeDocument, options: CompositeOptions): Plan {
  const preview = options.preview ?? null
  const all = doc.layers ?? []
  if (options.onlyLayerId !== undefined && options.onlyLayerId !== null) {
    return { units: [], only: all.find((layer) => layer.id === options.onlyLayerId) ?? null, onlyMode: true }
  }
  let layers: readonly Layer[] = all
  if (options.belowLayerId !== undefined && options.belowLayerId !== null) {
    const index = all.findIndex((layer) => layer.id === options.belowLayerId)
    if (index >= 0) layers = all.slice(0, options.includeBelowLayer ? index + 1 : index)
  }
  // Clipped layers join the nearest non-clipped layer below; a clipped layer with nothing below is a base.
  const groups: { base: Layer; clipped: Layer[] }[] = []
  for (const layer of layers) {
    const last = groups[groups.length - 1]
    if (layer.clipped && last) last.clipped.push(layer)
    else groups.push({ base: layer, clipped: [] })
  }
  const units: PlanUnit[] = []
  for (const group of groups) {
    const props = propsOf(group.base, preview)
    if (!props.visible || props.opacity <= 0) continue
    const clipped: PlanLayer[] = []
    for (const layer of group.clipped) {
      const clippedProps = propsOf(layer, preview)
      if (clippedProps.visible && clippedProps.opacity > 0) clipped.push({ layer, props: clippedProps })
    }
    units.push({ base: { layer: group.base, props }, clipped })
  }
  return { units, only: null, onlyMode: false }
}

// ---------------------------------------------------------------------------------------------
// Block context
// ---------------------------------------------------------------------------------------------

interface BlockContext {
  /** Block in level pixels (document space at `level`). */
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly n: number
  readonly level: number
  readonly preview: CompositePreview | null
}

function sourceTouches(source: PixelSource, ctx: BlockContext): boolean {
  const bounds = source.surface.contentBounds()
  if (!bounds) return false
  const scale = 2 ** ctx.level
  const ox = levelOffset(source.offsetX, ctx.level)
  const oy = levelOffset(source.offsetY, ctx.level)
  // Level-0, layer-local footprint of the block, with one proxy pixel of slack for the offset rounding.
  const x0 = (ctx.x - ox - 1) * scale
  const y0 = (ctx.y - oy - 1) * scale
  const x1 = (ctx.x + ctx.width - ox + 1) * scale
  const y1 = (ctx.y + ctx.height - oy + 1) * scale
  return bounds.x < x1 && bounds.x + bounds.width > x0 && bounds.y < y1 && bounds.y + bounds.height > y0
}

/** Layer-mask coverage of the block, or null when the layer has no enabled mask. */
function maskCoverage(layer: Layer, ctx: BlockContext): Uint8Array | null {
  const mask = layer.mask
  if (!mask || !mask.enabled) return null
  const coverage = maskScratch(COV, ctx.n)
  readMaskLevel(mask.surface, ctx.level, {
    x: ctx.x - levelOffset(mask.offsetX, ctx.level),
    y: ctx.y - levelOffset(mask.offsetY, ctx.level),
    width: ctx.width,
    height: ctx.height,
  }, { width: ctx.width, height: ctx.height, data: coverage })
  return coverage
}

/** Selection coverage of the block at the block's level (box average of the full-resolution mask). */
function selectionCoverage(selection: Selection | null, ctx: BlockContext): Uint8Array | null {
  if (!selection) return null
  const coverage = maskScratch(SEL, ctx.n)
  coverage.fill(0)
  const mask = selection.mask
  const scale = 2 ** ctx.level
  if (scale === 1) {
    const x0 = Math.max(0, ctx.x)
    const x1 = Math.min(mask.width, ctx.x + ctx.width)
    if (x1 <= x0) return coverage
    for (let y = Math.max(0, ctx.y); y < Math.min(mask.height, ctx.y + ctx.height); y += 1) {
      const from = y * mask.width + x0
      coverage.set(mask.data.subarray(from, from + (x1 - x0)), (y - ctx.y) * ctx.width + (x0 - ctx.x))
    }
    return coverage
  }
  const area = scale * scale
  for (let row = 0; row < ctx.height; row += 1) {
    const sy0 = (ctx.y + row) * scale
    for (let col = 0; col < ctx.width; col += 1) {
      const sx0 = (ctx.x + col) * scale
      let sum = 0
      for (let sy = Math.max(0, sy0); sy < Math.min(mask.height, sy0 + scale); sy += 1) {
        const base = sy * mask.width
        for (let sx = Math.max(0, sx0); sx < Math.min(mask.width, sx0 + scale); sx += 1) sum += mask.data[base + sx]
      }
      coverage[row * ctx.width + col] = Math.round(sum / area)
    }
  }
  return coverage
}

/** Dissolve: keep a pixel opaque when hash2(x, y, seed) < its effective alpha, otherwise drop it. */
function dissolveInPlace(src: Uint8ClampedArray, coverage: Uint8Array | null, opacity: number, ctx: BlockContext, seed: number): void {
  for (let i = 0, p = 3; i < ctx.n; i += 1, p += 4) {
    const alpha = src[p]
    if (alpha === 0) continue
    let effective = (alpha / 255) * opacity
    if (coverage) effective *= coverage[i] / 255
    const x = ctx.x + (i % ctx.width)
    const y = ctx.y + Math.floor(i / ctx.width)
    src[p] = hash2(x, y, seed) < effective ? 255 : 0
  }
}

function forceOpaque(data: Uint8ClampedArray): void {
  for (let p = 3; p < data.length; p += 4) data[p] = 255
}

function dithers(kernel: AdjustmentKernel): boolean {
  const spec = kernel.spec
  return spec.type === 'gradient-map' && Boolean(spec.dither)
}

/**
 * Runs `kernel` over the block's pixels in place. A dithering kernel (gradient map with dither) draws its
 * noise from the pixel index and the seed, so it runs row by row on a tile-anchored row buffer with a seed
 * per document row: every request that covers a pixel then gives it the same value.
 */
function runKernel(kernel: AdjustmentKernel, data: Uint8ClampedArray, ctx: BlockContext, seed: number): void {
  if (!dithers(kernel)) {
    kernel.apply(data, 0, ctx.n, seed)
    return
  }
  const tileLeft = tileIndex(ctx.x) * TILE_SIZE
  const lead = ctx.x - tileLeft
  const span = ctx.width * 4
  for (let row = 0; row < ctx.height; row += 1) {
    const from = row * span
    ditherRow.set(data.subarray(from, from + span), lead * 4)
    kernel.apply(ditherRow, lead, lead + ctx.width, mixSeed(seed, ctx.y + row, tileLeft))
    data.set(ditherRow.subarray(lead * 4, lead * 4 + span), from)
  }
}

/**
 * Applies an adjustment kernel to `dst` (alpha unchanged) mixed by opacity * coverage, with the layer's
 * blend mode: dst.rgb = mix(dst.rgb, B(dst.rgb, kernel(dst.rgb)), opacity * coverage).
 */
function applyKernel(dst: Uint8ClampedArray, ctx: BlockContext, kernel: AdjustmentKernel, opacity: number, mode: BlendMode,
  coverage: Uint8Array | null, seed: number): void {
  const n = ctx.n
  const plain = mode === 'normal' || mode === 'dissolve'
  if (plain && !coverage && opacity >= 1) {
    runKernel(kernel, dst, ctx, seed)
    return
  }
  const adjusted = rgbaScratch(T1, n)
  adjusted.set(dst)
  runKernel(kernel, adjusted, ctx, seed)
  if (plain) {
    for (let i = 0, p = 0; i < n; i += 1, p += 4) {
      if (dst[p + 3] === 0) continue
      let weight = opacity
      if (coverage) {
        const c = coverage[i]
        if (c === 0) continue
        if (c !== 255) weight *= c / 255
      }
      if (weight >= 1) {
        dst[p] = adjusted[p]
        dst[p + 1] = adjusted[p + 1]
        dst[p + 2] = adjusted[p + 2]
        continue
      }
      dst[p] = dst[p] + (adjusted[p] - dst[p]) * weight
      dst[p + 1] = dst[p + 1] + (adjusted[p + 1] - dst[p + 1]) * weight
      dst[p + 2] = dst[p + 2] + (adjusted[p + 2] - dst[p + 2]) * weight
    }
    return
  }
  // Other modes: blend the adjusted colour over an opaque copy of the backdrop, keep the backdrop alpha.
  const backdrop = rgbaScratch(T0, n)
  backdrop.set(dst)
  forceOpaque(backdrop)
  forceOpaque(adjusted)
  blendInto(backdrop, 0, adjusted, 0, n, mode, opacity, coverage, 0, 0)
  for (let p = 0; p < n * 4; p += 4) {
    if (dst[p + 3] === 0) continue
    dst[p] = backdrop[p]
    dst[p + 1] = backdrop[p + 1]
    dst[p + 2] = backdrop[p + 2]
  }
}

/**
 * Source-atop composite of `src` onto `group` (the group's alpha is locked), W3C-exact:
 * Co = as * ((1 - ab) Cs + ab B(Cb, Cs)) + (1 - as) Cb, computed as ab * T1 + (1 - ab) * T0 where T1 / T0
 * are the mode / normal composites over an opaque copy of the backdrop.
 */
function clipInto(group: Uint8ClampedArray, src: Uint8ClampedArray, n: number, mode: BlendMode, opacity: number, coverage: Uint8Array | null): void {
  const t1 = rgbaScratch(T1, n)
  t1.set(group)
  forceOpaque(t1)
  blendInto(t1, 0, src, 0, n, mode, opacity, coverage, 0, 0)
  let t0: Uint8ClampedArray | null = null
  if (mode !== 'normal') {
    t0 = rgbaScratch(T0, n)
    t0.set(group)
    forceOpaque(t0)
    blendInto(t0, 0, src, 0, n, 'normal', opacity, coverage, 0, 0)
  }
  for (let p = 0; p < n * 4; p += 4) {
    const ab = group[p + 3]
    if (ab === 0) continue
    if (ab === 255 || !t0) {
      group[p] = t1[p]
      group[p + 1] = t1[p + 1]
      group[p + 2] = t1[p + 2]
      continue
    }
    const k = ab / 255
    group[p] = t1[p] * k + t0[p] * (1 - k)
    group[p + 1] = t1[p + 1] * k + t0[p + 1] * (1 - k)
    group[p + 2] = t1[p + 2] * k + t0[p + 2] * (1 - k)
  }
}

/** Reads a layer's pixels for the block and applies the pixel previews that target it. */
function readLayerPixels(layer: Layer, source: PixelSource, ctx: BlockContext, target: Uint8ClampedArray): void {
  readSurfaceLevel(source.surface, ctx.level, {
    x: ctx.x - levelOffset(source.offsetX, ctx.level),
    y: ctx.y - levelOffset(source.offsetY, ctx.level),
    width: ctx.width,
    height: ctx.height,
  }, { width: ctx.width, height: ctx.height, data: target })
  const preview = ctx.preview
  if (!preview || preview.layerId !== layer.id) return
  if (preview.kind === 'layer-pixels') {
    if ((preview.level | 0) !== ctx.level) return
    const pixels = preview.pixels
    const rect = preview.rect
    const x0 = Math.max(ctx.x, rect.x)
    const y0 = Math.max(ctx.y, rect.y)
    const x1 = Math.min(ctx.x + ctx.width, rect.x + Math.min(rect.width, pixels.width))
    const y1 = Math.min(ctx.y + ctx.height, rect.y + Math.min(rect.height, pixels.height))
    if (x1 <= x0 || y1 <= y0) return
    for (let y = y0; y < y1; y += 1) {
      const from = ((y - rect.y) * pixels.width + (x0 - rect.x)) * 4
      target.set(pixels.data.subarray(from, from + (x1 - x0) * 4), ((y - ctx.y) * ctx.width + (x0 - ctx.x)) * 4)
    }
  } else if (preview.kind === 'adjustment') {
    // Destructive-dialog preview on a pixel layer: the spec applied inside the selection.
    const kernel = kernelFor(preview.spec)
    if (kernel.isIdentity) return
    applyKernel(target, ctx, kernel, 1, 'normal', selectionCoverage(preview.selection, ctx), seedOf(layer, ctx.level))
  }
}

function previewTargets(layer: Layer, preview: CompositePreview | null): boolean {
  return Boolean(preview && preview.layerId === layer.id && (preview.kind === 'layer-pixels' || preview.kind === 'adjustment'))
}

/** Draws one layer onto `dst`. alphaLocked = inside a clipping group (source-atop). */
function drawLayer(dst: Uint8ClampedArray, layer: Layer, opacity: number, mode: BlendMode, ctx: BlockContext, alphaLocked: boolean): void {
  if (opacity <= 0) return
  if (layer.kind === 'adjustment') {
    drawAdjustment(dst, layer, opacity, mode, ctx)
    return
  }
  const source = layerPixelSource(layer)
  if (!source) return
  if (!sourceTouches(source, ctx) && !previewTargets(layer, ctx.preview)) return
  const src = rgbaScratch(SRC, ctx.n)
  readLayerPixels(layer, source, ctx, src)
  let coverage = maskCoverage(layer, ctx)
  let blendMode = mode
  let amount = opacity
  if (mode === 'dissolve') {
    dissolveInPlace(src, coverage, opacity, ctx, seedOf(layer, ctx.level))
    blendMode = 'normal'
    amount = 1
    coverage = null
  }
  if (alphaLocked) clipInto(dst, src, ctx.n, blendMode, amount, coverage)
  else blendInto(dst, 0, src, 0, ctx.n, blendMode, amount, coverage, 0, 0)
}

function drawAdjustment(dst: Uint8ClampedArray, layer: AdjustmentLayer, opacity: number, mode: BlendMode, ctx: BlockContext): void {
  const preview = ctx.preview
  const spec = preview && preview.kind === 'adjustment' && preview.layerId === layer.id ? preview.spec : layer.adjustment
  if (!spec) return
  const kernel = kernelFor(spec)
  if (kernel.isIdentity) return
  applyKernel(dst, ctx, kernel, opacity, mode, maskCoverage(layer, ctx), seedOf(layer, ctx.level))
}

function anyAlpha(data: Uint8ClampedArray): boolean {
  for (let p = 3; p < data.length; p += 4) if (data[p] !== 0) return true
  return false
}

/** Composites a finished clipping group onto `acc` with the base layer's blend mode and opacity. */
function mergeGroup(acc: Uint8ClampedArray, group: Uint8ClampedArray, base: PlanLayer, ctx: BlockContext): void {
  const { blendMode, opacity } = base.props
  if (blendMode === 'dissolve') {
    dissolveInPlace(group, null, opacity, ctx, seedOf(base.layer, ctx.level))
    blendInto(acc, 0, group, 0, ctx.n, 'normal', 1, null, 0, 0)
    return
  }
  blendInto(acc, 0, group, 0, ctx.n, blendMode, opacity, null, 0, 0)
}

/**
 * The first layer drawn onto a still transparent block: every blend mode except Dissolve reduces to a copy
 * there (W3C formula with ab = 0), so an unmasked layer at full opacity is copied instead of blended.
 * Returns 'skipped' when the layer has nothing in the block, 'no' when it must be blended normally.
 */
function copyOntoEmpty(acc: Uint8ClampedArray, unit: PlanUnit, ctx: BlockContext): 'copied' | 'skipped' | 'no' {
  const { layer, props } = unit.base
  if (unit.clipped.length || layer.kind === 'adjustment' || props.opacity < 1 || props.blendMode === 'dissolve') return 'no'
  if (layer.mask && layer.mask.enabled) return 'no'
  const source = layerPixelSource(layer)
  if (!source) return 'no'
  if (!sourceTouches(source, ctx) && !previewTargets(layer, ctx.preview)) return 'skipped'
  readLayerPixels(layer, source, ctx, acc)
  // Same bytes as blending over transparency: invisible pixels carry no colour.
  for (let p = 3; p < ctx.n * 4; p += 4) {
    if (acc[p] === 0) {
      acc[p - 3] = 0
      acc[p - 2] = 0
      acc[p - 1] = 0
    }
  }
  return 'copied'
}

function compositeBlock(plan: Plan, acc: Uint8ClampedArray, ctx: BlockContext): void {
  if (plan.onlyMode) {
    const layer = plan.only
    if (!layer) return
    const source = layerPixelSource(layer)
    if (!source) return
    if (!sourceTouches(source, ctx) && !previewTargets(layer, ctx.preview)) return
    readLayerPixels(layer, source, ctx, acc)
    return
  }
  let empty = true
  for (const unit of plan.units) {
    const base = unit.base
    if (empty) {
      const shortcut = copyOntoEmpty(acc, unit, ctx)
      if (shortcut === 'skipped') continue
      if (shortcut === 'copied') {
        empty = false
        continue
      }
    }
    empty = false
    if (!unit.clipped.length) {
      drawLayer(acc, base.layer, base.props.opacity, base.props.blendMode, ctx, false)
      continue
    }
    const group = rgbaScratch(GROUP, ctx.n)
    group.fill(0)
    if (base.layer.kind === 'adjustment') {
      // The adjustment affects what is below; the layers clipped to it show through its mask.
      drawLayer(acc, base.layer, base.props.opacity, base.props.blendMode, ctx, false)
      for (const clipped of unit.clipped) drawLayer(group, clipped.layer, clipped.props.opacity, clipped.props.blendMode, ctx, false)
      const coverage = maskCoverage(base.layer, ctx)
      if (coverage) {
        for (let i = 0, p = 3; i < ctx.n; i += 1, p += 4) {
          const c = coverage[i]
          if (c !== 255) group[p] = (group[p] * c) / 255
        }
      }
    } else {
      drawLayer(group, base.layer, 1, 'normal', ctx, false)
      if (!anyAlpha(group)) continue
      for (const clipped of unit.clipped) drawLayer(group, clipped.layer, clipped.props.opacity, clipped.props.blendMode, ctx, true)
    }
    mergeGroup(acc, group, base, ctx)
  }
}

// ---------------------------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------------------------

function checkLevel(level: number | undefined): number {
  const value = Math.floor(Number(level ?? 0))
  if (!Number.isFinite(value) || value < 0) return 0
  return Math.min(MAX_LEVEL, value)
}

/**
 * Composites `rect` (level pixels) into `target` with the rect's top-left at (targetX, targetY), replacing
 * what the target held there. Pixels outside the document read as transparent.
 */
export function compositeInto(target: PixelBuffer, targetX: number, targetY: number, doc: CompositeDocument, rect: IntRect,
  options: CompositeOptions = {}): void {
  const r = toIntRect(rect)
  if (r.width <= 0 || r.height <= 0) return
  const level = checkLevel(options.level)
  const size = levelSize(doc.width, doc.height, level)
  const x0 = Math.max(0, r.x)
  const y0 = Math.max(0, r.y)
  const x1 = Math.min(size.width, r.x + r.width)
  const y1 = Math.min(size.height, r.y + r.height)
  const plan = planComposite(doc, options)
  const preview = options.preview ?? null
  const out = target.data
  const stride = target.width * 4
  // Clear the requested region first (pixels outside the document stay transparent).
  for (let y = 0; y < r.height; y += 1) {
    const ty = targetY + y
    if (ty < 0 || ty >= target.height) continue
    const from = Math.max(0, targetX)
    const to = Math.min(target.width, targetX + r.width)
    if (to > from) out.fill(0, ty * stride + from * 4, ty * stride + to * 4)
  }
  if (x1 <= x0 || y1 <= y0) return
  if (!plan.onlyMode && !plan.units.length) return
  for (let ty = tileIndex(y0); ty * TILE_SIZE < y1; ty += 1) {
    for (let tx = tileIndex(x0); tx * TILE_SIZE < x1; tx += 1) {
      const bx0 = Math.max(x0, tx * TILE_SIZE)
      const by0 = Math.max(y0, ty * TILE_SIZE)
      const bx1 = Math.min(x1, (tx + 1) * TILE_SIZE)
      const by1 = Math.min(y1, (ty + 1) * TILE_SIZE)
      const width = bx1 - bx0
      const height = by1 - by0
      const n = width * height
      const ox = targetX + (bx0 - r.x)
      const oy = targetY + (by0 - r.y)
      const direct = ox === 0 && oy === 0 && target.width === width && target.height === height
      const acc = direct ? out : rgbaScratch(ACC, n)
      if (!direct) acc.fill(0)
      compositeBlock(plan, acc, { x: bx0, y: by0, width, height, n, level, preview })
      if (direct) continue
      for (let row = 0; row < height; row += 1) {
        const yy = oy + row
        if (yy < 0 || yy >= target.height) continue
        const cx0 = Math.max(0, ox)
        const cx1 = Math.min(target.width, ox + width)
        if (cx1 <= cx0) continue
        const from = (row * width + (cx0 - ox)) * 4
        out.set(acc.subarray(from, from + (cx1 - cx0) * 4), yy * stride + cx0 * 4)
      }
    }
  }
}

/** Pure, synchronous, deterministic composite of `rect` (in level pixels; see CompositeOptions). */
export const compositeRect: CompositeRectFn = (doc, rect, options) => {
  const r = toIntRect(rect)
  const out: PixelBuffer = { width: r.width, height: r.height, data: new Uint8ClampedArray(r.width * r.height * 4) }
  compositeInto(out, 0, 0, doc, r, options ?? {})
  return out
}

/** Full-resolution composite of the whole document, built in tile rows (progress and abort per row). */
export function flattenDocument(doc: CompositeDocument, options?: OpOptions): PixelBuffer {
  const width = Math.max(0, doc.width | 0)
  const height = Math.max(0, doc.height | 0)
  const out: PixelBuffer = { width, height, data: new Uint8ClampedArray(width * height * 4) }
  const report = progressReporter(options)
  for (let top = 0; top < height; top += TILE_SIZE) {
    throwIfAborted(options?.signal)
    const rows = Math.min(TILE_SIZE, height - top)
    compositeInto(out, 0, top, doc, { x: 0, y: top, width, height: rows })
    report((top + rows) / height)
  }
  throwIfAborted(options?.signal)
  report(1)
  return out
}

/**
 * Eyedropper / Info readout: the average colour of a size x size square centred on (x, y) (alpha-weighted,
 * so transparent pixels do not darken it). 'current' samples the active layer's own pixels, 'current-below'
 * the composite up to and including it, 'all' the whole composite. Outside the canvas reads transparent.
 */
export function sampleDocument(doc: DocumentState, x: number, y: number, size: number, source: SampleSource,
  activeLayerId: LayerId | null = doc.activeLayerId): Rgba8 {
  const side = Math.max(1, Math.round(Number(size) || 1))
  const left = Math.floor(x) - Math.floor((side - 1) / 2)
  const top = Math.floor(y) - Math.floor((side - 1) / 2)
  const x0 = Math.max(0, left)
  const y0 = Math.max(0, top)
  const x1 = Math.min(doc.width, left + side)
  const y1 = Math.min(doc.height, top + side)
  if (!(x1 > x0 && y1 > y0) || !Number.isFinite(x) || !Number.isFinite(y)) return { r: 0, g: 0, b: 0, a: 0 }
  let options: CompositeOptions = {}
  const active = activeLayerId ? doc.layers.find((layer) => layer.id === activeLayerId) ?? null : null
  if (active && source === 'current' && active.kind !== 'adjustment') options = { onlyLayerId: active.id }
  else if (active && source !== 'all') options = { belowLayerId: active.id, includeBelowLayer: true }
  const pixels = compositeRect(doc, { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }, options)
  let sumA = 0
  let r = 0
  let g = 0
  let b = 0
  const data = pixels.data
  for (let p = 0; p < data.length; p += 4) {
    const a = data[p + 3]
    sumA += a
    r += data[p] * a
    g += data[p + 1] * a
    b += data[p + 2] * a
  }
  const count = data.length / 4
  if (sumA === 0) return { r: 0, g: 0, b: 0, a: 0 }
  return { r: Math.round(r / sumA), g: Math.round(g / sumA), b: Math.round(b / sumA), a: Math.round(sumA / count) }
}
