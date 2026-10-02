// src/imaging/mask.ts (WP4)
// 8-bit coverage masks: selection shapes, combine operations, Select > Modify, bounds and the
// marching-ants outline. Pure and DOM-free (worker and Node tests load it).
//
// Conventions
//   - A pixel (i, j) is the unit square [i, i + 1] x [j, j + 1]; its centre is (i + 0.5, j + 0.5).
//   - Rasterizers combine into `target` with max (union), so several shapes can share one scratch mask,
//     and return the tight bounds of the coverage they produced (null when nothing was drawn).
//   - Anti-aliased shapes use exact horizontal coverage on 16 sub-scanlines per pixel row (rectangles
//     use exact area coverage); aliased shapes include a pixel when its centre is inside.
//   - Select > Modify follows Photoshop: Feather is a Gaussian with sigma = radius / 2, the canvas edge
//     is not treated as a selection edge (clamp to edge), Expand / Contract are exact Euclidean with a
//     one-pixel anti-aliased edge measured on the 50% threshold.
import type { FillRule, IntRect, MaskBuffer, OpOptions, PixelBuffer, Point, Rect, SelectionOp } from './types.ts'
import { regionDistanceTransform } from './distance.ts'

/** Sub-scanlines per pixel row for anti-aliased ellipses and polygons. */
const SUBSAMPLES = 16

function abortError(): Error {
  const error = new Error('The operation was cancelled.')
  error.name = 'AbortError'
  return error
}

function assertSize(width: number, height: number): void {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
    throw new RangeError(`Invalid mask size ${width} x ${height}.`)
  }
}

function checkMask(mask: MaskBuffer, label = 'mask'): void {
  if (!mask || !mask.data) throw new TypeError(`The ${label} is missing.`)
  assertSize(mask.width, mask.height)
  if (mask.data.length !== mask.width * mask.height) {
    throw new RangeError(`The ${label} data must hold ${mask.width * mask.height} values, not ${mask.data.length}.`)
  }
}

function clampByte(value: number): number {
  if (!(value > 0)) return 0
  if (value >= 255) return 255
  return Math.round(value)
}

/** round(a * b / 255) for 0..255 integers, exact (Blinn). */
function mul255(a: number, b: number): number {
  const t = a * b + 128
  return (t + (t >> 8)) >> 8
}

export function createMaskBuffer(width: number, height: number, fill = 0): MaskBuffer {
  assertSize(width, height)
  const data = new Uint8Array(width * height)
  const value = clampByte(fill)
  if (value !== 0) data.fill(value)
  return { width, height, data }
}

export function cloneMask(mask: MaskBuffer): MaskBuffer {
  checkMask(mask)
  return { width: mask.width, height: mask.height, data: new Uint8Array(mask.data) }
}

/** Integer intersection of a rect with [0, width) x [0, height), or null when empty. */
export function clipRect(rect: IntRect, width: number, height: number): IntRect | null {
  const x0 = Math.max(0, rect.x)
  const y0 = Math.max(0, rect.y)
  const x1 = Math.min(width, rect.x + rect.width)
  const y1 = Math.min(height, rect.y + rect.height)
  if (!(x1 > x0 && y1 > y0)) return null
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
}

/** Copy of `rect` out of `mask`; areas outside the mask read as 0. */
export function cropMask(mask: MaskBuffer, rect: IntRect): MaskBuffer {
  checkMask(mask)
  const width = Math.max(0, Math.floor(rect.width))
  const height = Math.max(0, Math.floor(rect.height))
  const out = createMaskBuffer(width, height)
  const x = Math.floor(rect.x)
  const y = Math.floor(rect.y)
  const clip = clipRect({ x, y, width, height }, mask.width, mask.height)
  if (!clip) return out
  for (let row = clip.y; row < clip.y + clip.height; row += 1) {
    const from = row * mask.width + clip.x
    out.data.set(mask.data.subarray(from, from + clip.width), (row - y) * width + (clip.x - x))
  }
  return out
}

/** Copies `src` into `dst` with its top-left at (x, y), clipped to `dst`. */
export function pasteMask(dst: MaskBuffer, src: MaskBuffer, x: number, y: number): void {
  checkMask(dst, 'target mask')
  checkMask(src, 'source mask')
  const ox = Math.round(x)
  const oy = Math.round(y)
  const clip = clipRect({ x: ox, y: oy, width: src.width, height: src.height }, dst.width, dst.height)
  if (!clip) return
  for (let row = clip.y; row < clip.y + clip.height; row += 1) {
    const from = (row - oy) * src.width + (clip.x - ox)
    dst.data.set(src.data.subarray(from, from + clip.width), row * dst.width + clip.x)
  }
}

// ---------------------------------------------------------------------------------------------
// Rasterization
// ---------------------------------------------------------------------------------------------

interface Box {
  readonly x0: number
  readonly y0: number
  readonly x1: number
  readonly y1: number
}

function normalizeRect(rect: Rect): Box | null {
  let { x, y, width, height } = rect
  if (![x, y, width, height].every(Number.isFinite)) return null
  if (width < 0) { x += width; width = -width }
  if (height < 0) { y += height; height = -height }
  if (!(width > 0 && height > 0)) return null
  return { x0: x, y0: y, x1: x + width, y1: y + height }
}

class BoundsTracker {
  minX = Infinity
  minY = Infinity
  maxX = -Infinity
  maxY = -Infinity

  add(x0: number, x1: number, y: number): void {
    if (x0 < this.minX) this.minX = x0
    if (x1 > this.maxX) this.maxX = x1
    if (y < this.minY) this.minY = y
    if (y > this.maxY) this.maxY = y
  }

  rect(): IntRect | null {
    if (this.maxX < this.minX) return null
    return { x: this.minX, y: this.minY, width: this.maxX - this.minX + 1, height: this.maxY - this.minY + 1 }
  }
}

/** Pushes the sorted, disjoint inside intervals [a0, b0, a1, b1, ...] of a shape on the line y = ys. */
type SpanSource = (ys: number, out: number[]) => void

/**
 * Coverage from horizontal spans. Anti-aliased: SUBSAMPLES sub-scanlines per row, each contributing the
 * exact horizontal overlap of its spans with every pixel. Aliased: one scanline through the pixel
 * centres; pixel i is inside when a <= i + 0.5 < b.
 */
function fillSpans(target: MaskBuffer, top: number, bottom: number, antiAlias: boolean, spans: SpanSource): IntRect | null {
  const { width, data } = target
  const bounds = new BoundsTracker()
  const list: number[] = []
  if (!antiAlias) {
    for (let y = top; y < bottom; y += 1) {
      list.length = 0
      spans(y + 0.5, list)
      const row = y * width
      for (let k = 0; k + 1 < list.length; k += 2) {
        const i0 = Math.max(0, Math.ceil(list[k] - 0.5))
        const i1 = Math.min(width, Math.ceil(list[k + 1] - 0.5))
        if (i1 <= i0) continue
        data.fill(255, row + i0, row + i1)
        bounds.add(i0, i1 - 1, y)
      }
    }
    return bounds.rect()
  }
  const acc = new Float64Array(width + 2)
  const delta = new Float64Array(width + 2)
  const weight = 1 / SUBSAMPLES
  for (let y = top; y < bottom; y += 1) {
    let lo = width
    let hi = -1
    for (let s = 0; s < SUBSAMPLES; s += 1) {
      list.length = 0
      spans(y + (s + 0.5) * weight, list)
      for (let k = 0; k + 1 < list.length; k += 2) {
        const a = list[k] < 0 ? 0 : list[k]
        const b = list[k + 1] > width ? width : list[k + 1]
        if (!(b > a)) continue
        const ia = Math.floor(a)
        const ib = Math.floor(b)
        if (ia === ib) {
          acc[ia] += (b - a) * weight
        } else {
          acc[ia] += (ia + 1 - a) * weight
          if (ib > ia + 1) {
            delta[ia + 1] += weight
            delta[ib] -= weight
          }
          if (ib < width) acc[ib] += (b - ib) * weight
        }
        if (ia < lo) lo = ia
        const end = ib < width ? ib : width - 1
        if (end > hi) hi = end
      }
    }
    if (hi < lo) continue
    const row = y * width
    let running = 0
    let first = -1
    let last = -1
    for (let x = lo; x <= hi; x += 1) {
      running += delta[x]
      const coverage = acc[x] + running
      acc[x] = 0
      delta[x] = 0
      const value = coverage <= 0 ? 0 : coverage >= 1 ? 255 : Math.round(coverage * 255)
      if (value === 0) continue
      if (value > data[row + x]) data[row + x] = value
      if (first < 0) first = x
      last = x
    }
    delta[width] = 0
    delta[width + 1] = 0
    if (first >= 0) bounds.add(first, last, y)
  }
  return bounds.rect()
}

/**
 * Rectangle (Rectangular Marquee). Anti-aliased: exact area coverage of fractional edges.
 * Combined into `target` with max. Returns the bounds of the drawn coverage, or null.
 */
export function rasterizeRect(target: MaskBuffer, rect: Rect, antiAlias: boolean): IntRect | null {
  checkMask(target, 'target mask')
  const box = normalizeRect(rect)
  if (!box) return null
  const { width, height, data } = target
  if (!antiAlias) {
    const i0 = Math.max(0, Math.ceil(box.x0 - 0.5))
    const i1 = Math.min(width, Math.ceil(box.x1 - 0.5))
    const j0 = Math.max(0, Math.ceil(box.y0 - 0.5))
    const j1 = Math.min(height, Math.ceil(box.y1 - 0.5))
    if (i0 >= i1 || j0 >= j1) return null
    for (let y = j0; y < j1; y += 1) data.fill(255, y * width + i0, y * width + i1)
    return { x: i0, y: j0, width: i1 - i0, height: j1 - j0 }
  }
  const x0 = Math.max(0, box.x0)
  const x1 = Math.min(width, box.x1)
  const y0 = Math.max(0, box.y0)
  const y1 = Math.min(height, box.y1)
  if (!(x1 > x0 && y1 > y0)) return null
  const i0 = Math.floor(x0)
  const i1 = Math.ceil(x1)
  const columns = new Float64Array(i1 - i0)
  for (let i = i0; i < i1; i += 1) columns[i - i0] = Math.min(i + 1, x1) - Math.max(i, x0)
  const bounds = new BoundsTracker()
  for (let j = Math.floor(y0); j < Math.ceil(y1); j += 1) {
    const cy = Math.min(j + 1, y1) - Math.max(j, y0)
    const row = j * width
    let first = -1
    let last = -1
    for (let i = i0; i < i1; i += 1) {
      const value = Math.round(columns[i - i0] * cy * 255)
      if (value <= 0) continue
      if (value > data[row + i]) data[row + i] = value > 255 ? 255 : value
      if (first < 0) first = i
      last = i
    }
    if (first >= 0) bounds.add(first, last, j)
  }
  return bounds.rect()
}

/** Ellipse inscribed in `rect` (Elliptical Marquee). Combined into `target` with max. */
export function rasterizeEllipse(target: MaskBuffer, rect: Rect, antiAlias: boolean): IntRect | null {
  checkMask(target, 'target mask')
  const box = normalizeRect(rect)
  if (!box) return null
  const cx = (box.x0 + box.x1) / 2
  const cy = (box.y0 + box.y1) / 2
  const rx = (box.x1 - box.x0) / 2
  const ry = (box.y1 - box.y0) / 2
  const top = Math.max(0, Math.floor(box.y0))
  const bottom = Math.min(target.height, Math.ceil(box.y1))
  if (top >= bottom || box.x1 <= 0 || box.x0 >= target.width) return null
  return fillSpans(target, top, bottom, antiAlias, (ys, out) => {
    const t = (ys - cy) / ry
    if (!(t > -1 && t < 1)) return
    const half = rx * Math.sqrt(1 - t * t)
    out.push(cx - half, cx + half)
  })
}

/**
 * Closed polygon (Lasso, Polygonal Lasso); the last point connects back to the first. Non-finite points
 * are ignored. 'nonzero' (default, Photoshop lasso) or 'evenodd'. Combined into `target` with max.
 */
export function rasterizePolygon(target: MaskBuffer, points: readonly Point[], antiAlias: boolean, rule: FillRule = 'nonzero'): IntRect | null {
  checkMask(target, 'target mask')
  const xs: number[] = []
  const ys: number[] = []
  for (const point of points ?? []) {
    if (point && Number.isFinite(point.x) && Number.isFinite(point.y)) {
      xs.push(point.x)
      ys.push(point.y)
    }
  }
  const count = xs.length
  if (count < 3) return null
  // Non-horizontal edges as (ymin, ymax, x at ymin, dx/dy, direction), sorted by ymin.
  const order: number[] = []
  const eTop: number[] = []
  const eBottom: number[] = []
  const eX: number[] = []
  const eSlope: number[] = []
  const eDir: number[] = []
  let minY = Infinity
  let maxY = -Infinity
  for (let i = 0; i < count; i += 1) {
    const j = i + 1 === count ? 0 : i + 1
    const ax = xs[i]
    const ay = ys[i]
    const bx = xs[j]
    const by = ys[j]
    if (ay < minY) minY = ay
    if (ay > maxY) maxY = ay
    if (ay === by) continue
    const down = by > ay
    const top = down ? ay : by
    const bottom = down ? by : ay
    eTop.push(top)
    eBottom.push(bottom)
    eX.push(down ? ax : bx)
    eSlope.push((bx - ax) / (by - ay))
    eDir.push(down ? 1 : -1)
    order.push(order.length)
  }
  if (!order.length) return null
  order.sort((a, b) => eTop[a] - eTop[b])
  const top = Math.max(0, Math.floor(minY))
  const bottom = Math.min(target.height, Math.ceil(maxY))
  if (top >= bottom) return null
  const evenOdd = rule === 'evenodd'
  const active = new Int32Array(order.length)
  let activeCount = 0
  let next = 0
  const crossX = new Float64Array(order.length)
  const crossDir = new Int8Array(order.length)
  return fillSpans(target, top, bottom, antiAlias, (ysample, out) => {
    while (next < order.length && eTop[order[next]] <= ysample) active[activeCount++] = order[next++]
    let crossings = 0
    let keep = 0
    for (let k = 0; k < activeCount; k += 1) {
      const e = active[k]
      if (eBottom[e] <= ysample) continue
      active[keep++] = e
      // Half-open [top, bottom): the edge is active from its top row onwards.
      const x = eX[e] + (ysample - eTop[e]) * eSlope[e]
      const dir = eDir[e]
      // Insertion sort by x.
      let at = crossings
      while (at > 0 && crossX[at - 1] > x) {
        crossX[at] = crossX[at - 1]
        crossDir[at] = crossDir[at - 1]
        at -= 1
      }
      crossX[at] = x
      crossDir[at] = dir
      crossings += 1
    }
    activeCount = keep
    let winding = 0
    let start = 0
    for (let k = 0; k < crossings; k += 1) {
      const before = evenOdd ? (winding & 1) !== 0 : winding !== 0
      winding += crossDir[k]
      const after = evenOdd ? (winding & 1) !== 0 : winding !== 0
      if (!before && after) start = crossX[k]
      else if (before && !after && crossX[k] > start) out.push(start, crossX[k])
    }
  })
}

// ---------------------------------------------------------------------------------------------
// Combine, invert, bounds
// ---------------------------------------------------------------------------------------------

/**
 * dst = op(dst, src), in place. replace: src; add: max; subtract: dst * (255 - src) / 255;
 * intersect: dst * src / 255 (rounded exactly). `rect`, when given, promises that `src` is 0 outside it:
 * the result is identical to combining the whole masks, only faster.
 */
export function combineMasks(dst: MaskBuffer, src: MaskBuffer, op: SelectionOp, rect?: IntRect): void {
  checkMask(dst, 'target mask')
  checkMask(src, 'source mask')
  if (dst.width !== src.width || dst.height !== src.height) {
    throw new RangeError(`Cannot combine a ${src.width} x ${src.height} mask into a ${dst.width} x ${dst.height} mask.`)
  }
  const { width, height } = dst
  const a = dst.data
  const b = src.data
  const full = { x: 0, y: 0, width, height }
  const region = rect ? clipRect(rect, width, height) : full
  if (op === 'replace' || op === 'intersect') {
    // Everything outside the region becomes 0 (src is 0 there).
    if (!region) {
      a.fill(0)
      return
    }
    a.fill(0, 0, region.y * width)
    a.fill(0, (region.y + region.height) * width)
    for (let y = region.y; y < region.y + region.height; y += 1) {
      const row = y * width
      a.fill(0, row, row + region.x)
      a.fill(0, row + region.x + region.width, row + width)
    }
  }
  if (!region) return
  for (let y = region.y; y < region.y + region.height; y += 1) {
    const start = y * width + region.x
    const end = start + region.width
    switch (op) {
      case 'replace':
        a.set(b.subarray(start, end), start)
        break
      case 'add':
        for (let i = start; i < end; i += 1) if (b[i] > a[i]) a[i] = b[i]
        break
      case 'subtract':
        for (let i = start; i < end; i += 1) if (b[i] !== 0 && a[i] !== 0) a[i] = mul255(a[i], 255 - b[i])
        break
      case 'intersect':
        for (let i = start; i < end; i += 1) a[i] = a[i] === 0 || b[i] === 0 ? 0 : mul255(a[i], b[i])
        break
      default:
        throw new RangeError(`Unknown selection operation "${String(op)}".`)
    }
  }
}

/** In place: 255 - value. */
export function invertMask(mask: MaskBuffer): void {
  checkMask(mask)
  const data = mask.data
  for (let i = 0; i < data.length; i += 1) data[i] = 255 - data[i]
}

/**
 * Tight bounds of the pixels whose value is >= threshold (default 1, any coverage), or null.
 * `within` restricts the search to a region (pixels outside it are ignored).
 */
export function maskBounds(mask: MaskBuffer, threshold = 1, within?: IntRect): IntRect | null {
  checkMask(mask)
  const { width, data } = mask
  const region = clipRect(within ?? { x: 0, y: 0, width, height: mask.height }, width, mask.height)
  if (!region) return null
  const t = Math.max(0, Math.min(255, Math.ceil(Number.isFinite(threshold) ? threshold : 1)))
  const x0 = region.x
  const x1 = region.x + region.width
  const y0 = region.y
  const y1 = region.y + region.height
  if (t === 0) return region
  const rowHas = (y: number): boolean => {
    const row = y * width
    for (let x = x0; x < x1; x += 1) if (data[row + x] >= t) return true
    return false
  }
  let top = -1
  for (let y = y0; y < y1; y += 1) {
    if (rowHas(y)) { top = y; break }
  }
  if (top < 0) return null
  let bottom = top
  for (let y = y1 - 1; y > top; y -= 1) {
    if (rowHas(y)) { bottom = y; break }
  }
  let left = x1
  let right = x0 - 1
  for (let y = top; y <= bottom; y += 1) {
    const row = y * width
    for (let x = x0; x < left; x += 1) {
      if (data[row + x] >= t) { left = x; break }
    }
    for (let x = x1 - 1; x > right; x -= 1) {
      if (data[row + x] >= t) { right = x; break }
    }
  }
  return { x: left, y: top, width: right - left + 1, height: bottom - top + 1 }
}

/** The alpha channel as a mask ("Load selection from layer", Ctrl+click on a layer thumbnail). */
export function maskFromAlpha(src: PixelBuffer): MaskBuffer {
  if (!src || !src.data || src.data.length !== src.width * src.height * 4) {
    throw new RangeError('The pixel buffer does not match its size.')
  }
  const out = createMaskBuffer(src.width, src.height)
  const data = src.data
  for (let i = 0, p = 3; i < out.data.length; i += 1, p += 4) out.data[i] = data[p]
  return out
}

/** Box-averaged copy reduced by an integer factor (display-level outlines). Size rounds up. */
export function downsampleMask(mask: MaskBuffer, factor: number): MaskBuffer {
  checkMask(mask)
  const f = Math.max(1, Math.floor(factor))
  if (f === 1) return cloneMask(mask)
  const width = Math.ceil(mask.width / f)
  const height = Math.ceil(mask.height / f)
  const out = createMaskBuffer(width, height)
  const sums = new Float64Array(width)
  const counts = new Float64Array(width)
  const column = new Int32Array(mask.width)
  for (let x = 0; x < mask.width; x += 1) column[x] = (x / f) | 0
  for (let oy = 0; oy < height; oy += 1) {
    sums.fill(0)
    counts.fill(0)
    const yEnd = Math.min(mask.height, (oy + 1) * f)
    for (let y = oy * f; y < yEnd; y += 1) {
      const row = y * mask.width
      for (let x = 0; x < mask.width; x += 1) {
        const ox = column[x]
        sums[ox] += mask.data[row + x]
        counts[ox] += 1
      }
    }
    for (let ox = 0; ox < width; ox += 1) out.data[oy * width + ox] = Math.round(sums[ox] / counts[ox])
  }
  return out
}

// ---------------------------------------------------------------------------------------------
// Select > Modify: feather, expand, contract
// ---------------------------------------------------------------------------------------------

interface BlurPlan {
  /** Box radii applied in sequence, or a normalised Gaussian kernel. */
  readonly boxes: readonly number[] | null
  readonly kernel: Float64Array | null
  /** Total reach of the plan in pixels. */
  readonly support: number
}

function blurPlan(sigma: number): BlurPlan {
  if (sigma < 2) {
    const reach = Math.max(1, Math.ceil(3 * sigma))
    const kernel = new Float64Array(2 * reach + 1)
    let sum = 0
    for (let k = -reach; k <= reach; k += 1) {
      const w = Math.exp(-(k * k) / (2 * sigma * sigma))
      kernel[k + reach] = w
      sum += w
    }
    for (let k = 0; k < kernel.length; k += 1) kernel[k] /= sum
    return { boxes: null, kernel, support: reach }
  }
  // Three box passes approximating the Gaussian (box widths from Kovesi / Kutskir).
  const passes = 3
  const ideal = Math.sqrt((12 * sigma * sigma) / passes + 1)
  let lower = Math.floor(ideal)
  if (lower % 2 === 0) lower -= 1
  const upper = lower + 2
  const m = Math.round((12 * sigma * sigma - passes * lower * lower - 4 * passes * lower - 3 * passes) / (-4 * lower - 4))
  const boxes: number[] = []
  for (let i = 0; i < passes; i += 1) boxes.push(((i < m ? lower : upper) - 1) / 2)
  return { boxes, kernel: null, support: boxes.reduce((a, b) => a + b, 0) }
}

function boxLine(src: Float32Array, dst: Float32Array, n: number, radius: number): void {
  if (radius <= 0) {
    dst.set(src.subarray(0, n))
    return
  }
  const last = n - 1
  const scale = 1 / (2 * radius + 1)
  let sum = (radius + 1) * src[0]
  for (let k = 1; k <= radius; k += 1) sum += src[k < last ? k : last]
  for (let i = 0; i < n; i += 1) {
    dst[i] = sum * scale
    const add = i + radius + 1
    const sub = i - radius
    sum += src[add < last ? add : last] - src[sub > 0 ? sub : 0]
  }
}

function kernelLine(src: Float32Array, dst: Float32Array, n: number, kernel: Float64Array): void {
  const reach = (kernel.length - 1) / 2
  const last = n - 1
  for (let i = 0; i < n; i += 1) {
    let acc = 0
    for (let k = -reach; k <= reach; k += 1) {
      let j = i + k
      if (j < 0) j = 0
      else if (j > last) j = last
      acc += kernel[k + reach] * src[j]
    }
    dst[i] = acc
  }
}

/** Runs the plan on `line` (length n) using `tmp`; returns the array holding the result. */
function blurLine(line: Float32Array, tmp: Float32Array, n: number, plan: BlurPlan): Float32Array {
  if (plan.kernel) {
    kernelLine(line, tmp, n, plan.kernel)
    return tmp
  }
  let from = line
  let to = tmp
  for (const radius of plan.boxes as readonly number[]) {
    boxLine(from, to, n, radius)
    const swap = from
    from = to
    to = swap
  }
  return from
}

function expandRect(rect: IntRect, margin: number, width: number, height: number): IntRect {
  return clipRect({ x: rect.x - margin, y: rect.y - margin, width: rect.width + 2 * margin, height: rect.height + 2 * margin }, width, height) as IntRect
}

/**
 * Select > Modify > Feather: Gaussian blur of the coverage with sigma = radius / 2 (exact kernel below
 * sigma 2, three box passes above), clamping at the canvas edge so a selection touching the edge keeps
 * full coverage there. Total coverage is preserved. Returns a new mask.
 */
export function featherMask(mask: MaskBuffer, radius: number, options?: OpOptions): MaskBuffer {
  checkMask(mask)
  if (!(radius > 0) || !Number.isFinite(radius)) return cloneMask(mask)
  const { width, height } = mask
  const bounds = maskBounds(mask)
  if (!bounds) return cloneMask(mask)
  const plan = blurPlan(radius / 2)
  const region = expandRect(bounds, Math.ceil(plan.support) + 1, width, height)
  const rw = region.width
  const rh = region.height
  const signal = options?.signal
  const onProgress = options?.onProgress
  // Horizontal passes, row by row, into a float copy of the region.
  const buffer = new Float32Array(rw * rh)
  const line = new Float32Array(Math.max(rw, rh))
  const tmp = new Float32Array(Math.max(rw, rh))
  for (let y = 0; y < rh; y += 1) {
    if ((y & 63) === 0 && signal?.aborted) throw abortError()
    const from = (region.y + y) * width + region.x
    for (let x = 0; x < rw; x += 1) line[x] = mask.data[from + x]
    const result = blurLine(line, tmp, rw, plan)
    buffer.set(result.subarray(0, rw), y * rw)
  }
  onProgress?.(0.5)
  // Vertical passes in column strips (cache friendly), written straight into the output.
  const out = createMaskBuffer(width, height)
  const STRIP = 16
  const column = new Float32Array(rh * STRIP)
  const columnTmp = new Float32Array(rh)
  let reported = 0.5
  for (let sx = 0; sx < rw; sx += STRIP) {
    if (signal?.aborted) throw abortError()
    const strip = Math.min(STRIP, rw - sx)
    for (let y = 0; y < rh; y += 1) {
      const row = y * rw + sx
      for (let c = 0; c < strip; c += 1) column[c * rh + y] = buffer[row + c]
    }
    for (let c = 0; c < strip; c += 1) {
      const view = column.subarray(c * rh, (c + 1) * rh)
      const result = blurLine(view, columnTmp, rh, plan)
      const x = region.x + sx + c
      for (let y = 0; y < rh; y += 1) out.data[(region.y + y) * width + x] = clampByte(result[y])
    }
    const fraction = 0.5 + 0.5 * ((sx + strip) / rw)
    if (onProgress && fraction - reported >= 0.05) {
      reported = fraction
      onProgress(fraction)
    }
  }
  onProgress?.(1)
  return out
}

/**
 * Select > Modify > Expand: the 50% selection grown by `pixels` with exact Euclidean (round) corners and a
 * one-pixel anti-aliased edge; straight edges move by exactly `pixels`. Never shrinks coverage.
 */
export function expandMask(mask: MaskBuffer, pixels: number, options?: OpOptions): MaskBuffer {
  checkMask(mask)
  const out = cloneMask(mask)
  if (!(pixels > 0) || !Number.isFinite(pixels)) return out
  const { width, height } = mask
  const bounds = maskBounds(mask, 128)
  if (!bounds) return out
  const region = expandRect(bounds, Math.ceil(pixels) + 2, width, height)
  const rw = region.width
  const sites = new Uint8Array(rw * region.height)
  for (let y = 0; y < region.height; y += 1) {
    const from = (region.y + y) * width + region.x
    for (let x = 0; x < rw; x += 1) sites[y * rw + x] = mask.data[from + x] >= 128 ? 1 : 0
  }
  const distance = regionDistanceTransform(sites, rw, region.height, options)
  for (let y = 0; y < region.height; y += 1) {
    const row = (region.y + y) * width + region.x
    for (let x = 0; x < rw; x += 1) {
      const coverage = pixels + 0.5 - distance[y * rw + x]
      if (coverage <= 0) continue
      const value = coverage >= 1 ? 255 : Math.round(coverage * 255)
      if (value > out.data[row + x]) out.data[row + x] = value
    }
  }
  return out
}

/**
 * Select > Modify > Contract: the 50% selection shrunk by `pixels` (exact Euclidean distance to the
 * unselected area, one-pixel anti-aliased edge; the canvas edge is not a selection edge, as in
 * Photoshop). Never grows coverage.
 */
export function contractMask(mask: MaskBuffer, pixels: number, options?: OpOptions): MaskBuffer {
  checkMask(mask)
  const out = cloneMask(mask)
  if (!(pixels > 0) || !Number.isFinite(pixels)) return out
  const { width, height } = mask
  const bounds = maskBounds(mask)
  if (!bounds) return out
  const region = expandRect(bounds, Math.ceil(pixels) + 2, width, height)
  const rw = region.width
  const sites = new Uint8Array(rw * region.height)
  for (let y = 0; y < region.height; y += 1) {
    const from = (region.y + y) * width + region.x
    for (let x = 0; x < rw; x += 1) sites[y * rw + x] = mask.data[from + x] < 128 ? 1 : 0
  }
  const distance = regionDistanceTransform(sites, rw, region.height, options)
  for (let y = 0; y < region.height; y += 1) {
    const row = (region.y + y) * width + region.x
    for (let x = 0; x < rw; x += 1) {
      const coverage = distance[y * rw + x] - pixels + 0.5
      if (coverage >= 1) continue
      const value = coverage <= 0 ? 0 : Math.round(coverage * 255)
      if (value < out.data[row + x]) out.data[row + x] = value
    }
  }
  return out
}

// ---------------------------------------------------------------------------------------------
// Outline (marching ants)
// ---------------------------------------------------------------------------------------------

/**
 * Boundary of the pixels whose value is >= threshold (default 128, Photoshop's 50% rule), traced along
 * pixel edges and merged into maximal straight runs, as segments [x0, y0, x1, y1, ...] in mask pixel
 * units times `scale`. A rectangle yields exactly four segments. Pass a downsampled mask and its factor as
 * `scale` to outline at a display level.
 */
export function traceOutline(mask: MaskBuffer, threshold = 128, scale = 1): Float32Array {
  checkMask(mask)
  const { width, height, data } = mask
  const t = Math.max(1, Math.min(255, Math.ceil(Number.isFinite(threshold) ? threshold : 128)))
  const s = Number.isFinite(scale) && scale > 0 ? scale : 1
  let segments = new Float32Array(256)
  let length = 0
  const push = (x0: number, y0: number, x1: number, y1: number) => {
    if (length + 4 > segments.length) {
      const grown = new Float32Array(segments.length * 2)
      grown.set(segments)
      segments = grown
    }
    segments[length++] = x0 * s
    segments[length++] = y0 * s
    segments[length++] = x1 * s
    segments[length++] = y1 * s
  }
  if (width === 0 || height === 0) return new Float32Array(0)
  const inside = new Uint8Array(width * height)
  for (let i = 0; i < inside.length; i += 1) inside[i] = data[i] >= t ? 1 : 0
  // Horizontal edges on the line y = j separate row j - 1 (above) from row j (below).
  for (let j = 0; j <= height; j += 1) {
    const above = (j - 1) * width
    const below = j * width
    let start = -1
    for (let x = 0; x <= width; x += 1) {
      let differs = false
      if (x < width) {
        const a = j > 0 ? inside[above + x] : 0
        const b = j < height ? inside[below + x] : 0
        differs = a !== b
      }
      if (differs) {
        if (start < 0) start = x
      } else if (start >= 0) {
        push(start, j, x, j)
        start = -1
      }
    }
  }
  // Vertical edges on the line x = i separate column i - 1 from column i; runs grow down the rows.
  const runStart = new Int32Array(width + 1).fill(-1)
  for (let y = 0; y <= height; y += 1) {
    const row = y * width
    for (let i = 0; i <= width; i += 1) {
      let differs = false
      if (y < height) {
        const a = i > 0 ? inside[row + i - 1] : 0
        const b = i < width ? inside[row + i] : 0
        differs = a !== b
      }
      if (differs) {
        if (runStart[i] < 0) runStart[i] = y
      } else if (runStart[i] >= 0) {
        push(i, runStart[i], i, y)
        runStart[i] = -1
      }
    }
  }
  return segments.slice(0, length)
}
