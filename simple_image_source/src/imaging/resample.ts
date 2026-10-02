// src/imaging/resample.ts (WP2)
// High-quality resizing. Separable two-pass filtering on premultiplied float colour (transparent edges never
// darken), with the kernel stretched by the shrink factor so downscaling is anti-aliased (as in Pillow).
// When shrinking, borders use whole-sample mirroring ("reflect 101": -1 -> 1), which keeps periodic patterns
// periodic, so a 2x reduction of a one-pixel checkerboard is uniform grey; when enlarging, the edge sample
// repeats so gradients stay monotone up to the border. Rows stream through the stages with small row
// caches, so peak memory is about the output buffer plus a few rows, whatever the scale. Pure and DOM-free.
import type { OpOptions, PixelBuffer, ResampleMethod } from './types.ts'
import { assertBuffer, chunkRowsFor, cloneBuffer, createBuffer, finishedRun, runRowsSync, throwIfAborted, unpremultiplyInto } from './buffer.ts'
import type { RowRun } from './buffer.ts'

// ---------------------------------------------------------------------------------------------
// Kernels (shared with transform.ts)
// ---------------------------------------------------------------------------------------------

/** Keys cubic; a = -0.5 is Catmull-Rom (the 'bicubic' interpolation). Support 2. */
export function cubicWeight(x: number, a = -0.5): number {
  const t = x < 0 ? -x : x
  if (t < 1) return ((a + 2) * t - (a + 3)) * t * t + 1
  if (t < 2) return ((a * t - 5 * a) * t + 8 * a) * t - 4 * a
  return 0
}

/** Lanczos windowed sinc with `lobes` lobes (3 = 'lanczos3'). Exactly 0 at non-zero integers. */
export function lanczosWeight(x: number, lobes = 3): number {
  const t = x < 0 ? -x : x
  if (t === 0) return 1
  if (t >= lobes) return 0
  if (t === Math.floor(t)) return 0
  const px = Math.PI * t
  return (lobes * Math.sin(px) * Math.sin(px / lobes)) / (px * px)
}

/** Triangle (tent) filter: bilinear interpolation. Support 1. */
export function triangleWeight(x: number): number {
  const t = x < 0 ? -x : x
  return t < 1 ? 1 - t : 0
}

interface Kernel {
  readonly support: number
  readonly weight: (x: number) => number
}

const KERNELS: Readonly<Record<'bilinear' | 'bicubic' | 'lanczos3', Kernel>> = Object.freeze({
  bilinear: { support: 1, weight: triangleWeight },
  bicubic: { support: 2, weight: (x: number) => cubicWeight(x) },
  lanczos3: { support: 3, weight: (x: number) => lanczosWeight(x, 3) },
})

/** Mirror an out-of-range index into [0, length) without repeating the edge sample (-1 -> 1, length -> length - 2). */
export function reflectIndex(index: number, length: number): number {
  if (length <= 1) return 0
  if (index >= 0 && index < length) return index
  const period = 2 * (length - 1)
  let i = index % period
  if (i < 0) i += period
  return i < length ? i : period - i
}

// ---------------------------------------------------------------------------------------------
// Contributions (which input samples, with which weights, make each output sample)
// ---------------------------------------------------------------------------------------------

interface Contributions {
  readonly outLength: number
  readonly taps: number
  /** outLength * taps input indices (padding repeats a valid index with weight 0). */
  readonly index: Int32Array
  readonly weight: Float64Array
  /** Smallest and largest input index each output sample reads. */
  readonly low: Int32Array
  readonly high: Int32Array
}

function finish(outLength: number, taps: number, index: Int32Array, weight: Float64Array): Contributions {
  const low = new Int32Array(outLength)
  const high = new Int32Array(outLength)
  for (let o = 0; o < outLength; o += 1) {
    let lo = Infinity
    let hi = -Infinity
    for (let k = 0; k < taps; k += 1) {
      const i = index[o * taps + k]
      if (i < lo) lo = i
      if (i > hi) hi = i
    }
    low[o] = lo
    high[o] = hi
  }
  return { outLength, taps, index, weight, low, high }
}

function clampIndex(index: number, length: number): number {
  return index < 0 ? 0 : index >= length ? length - 1 : index
}

function kernelContributions(inLength: number, outLength: number, kernel: Kernel): Contributions {
  const scale = inLength / outLength
  const filterScale = Math.max(1, scale)
  const support = kernel.support * filterScale
  const taps = 2 * Math.ceil(support) + 1
  const index = new Int32Array(outLength * taps)
  const weight = new Float64Array(outLength * taps)
  // Shrinking mirrors the border (periodic detail stays periodic); enlarging repeats the edge sample, which
  // keeps gradients monotone up to the border instead of folding them back.
  const mapIndex = scale > 1 ? reflectIndex : clampIndex
  for (let o = 0; o < outLength; o += 1) {
    const centre = (o + 0.5) * scale
    const base = o * taps
    let count = 0
    let sum = 0
    const first = Math.floor(centre - support - 0.5)
    const last = Math.ceil(centre + support - 0.5)
    for (let k = first; k <= last && count < taps; k += 1) {
      const w = kernel.weight((k + 0.5 - centre) / filterScale)
      if (w === 0) continue
      index[base + count] = mapIndex(k, inLength)
      weight[base + count] = w
      sum += w
      count += 1
    }
    if (count === 0) {
      index[base] = reflectIndex(Math.min(inLength - 1, Math.max(0, Math.floor(centre))), inLength)
      weight[base] = 1
      sum = 1
      count = 1
    }
    for (let k = 0; k < count; k += 1) weight[base + k] /= sum
    for (let k = count; k < taps; k += 1) index[base + k] = index[base + count - 1]
  }
  return finish(outLength, taps, index, weight)
}

/** Exact area weights: each output sample averages the input interval it covers (fractional overlaps). */
function areaContributions(inLength: number, outLength: number): Contributions {
  const scale = inLength / outLength
  const taps = Math.ceil(scale) + 1
  const index = new Int32Array(outLength * taps)
  const weight = new Float64Array(outLength * taps)
  for (let o = 0; o < outLength; o += 1) {
    const start = o * scale
    const end = Math.min(inLength, (o + 1) * scale)
    const base = o * taps
    let count = 0
    let sum = 0
    for (let k = Math.floor(start); k < end && count < taps; k += 1) {
      const overlap = Math.min(end, k + 1) - Math.max(start, k)
      if (overlap <= 0) continue
      index[base + count] = Math.min(inLength - 1, k)
      weight[base + count] = overlap
      sum += overlap
      count += 1
    }
    if (count === 0) {
      index[base] = Math.min(inLength - 1, Math.floor(start))
      weight[base] = 1
      sum = 1
      count = 1
    }
    for (let k = 0; k < count; k += 1) weight[base + k] /= sum
    for (let k = count; k < taps; k += 1) index[base + k] = index[base + count - 1]
  }
  return finish(outLength, taps, index, weight)
}

// ---------------------------------------------------------------------------------------------
// Streaming row pipeline (premultiplied RGBA floats, 0..255 units)
// ---------------------------------------------------------------------------------------------

interface RowSource {
  readonly width: number
  readonly height: number
  /** Writes row y into `target` (width * 4 floats). Consumers request rows in non-decreasing windows. */
  read(y: number, target: Float32Array): void
}

function pixelRows(src: PixelBuffer): RowSource {
  const { width, height, data } = src
  return {
    width,
    height,
    read(y, target) {
      const base = y * width * 4
      for (let j = 0; j < width * 4; j += 4) {
        const a = data[base + j + 3]
        if (a === 255) {
          target[j] = data[base + j]
          target[j + 1] = data[base + j + 1]
          target[j + 2] = data[base + j + 2]
          target[j + 3] = 255
        } else if (a === 0) {
          target[j] = 0
          target[j + 1] = 0
          target[j + 2] = 0
          target[j + 3] = 0
        } else {
          const k = a / 255
          target[j] = data[base + j] * k
          target[j + 1] = data[base + j + 1] * k
          target[j + 2] = data[base + j + 2] * k
          target[j + 3] = a
        }
      }
    },
  }
}

function horizontalStage(source: RowSource, contributions: Contributions): RowSource {
  const scratch = new Float32Array(source.width * 4)
  const { taps, index, weight, outLength } = contributions
  return {
    width: outLength,
    height: source.height,
    read(y, target) {
      source.read(y, scratch)
      for (let o = 0; o < outLength; o += 1) {
        let r = 0
        let g = 0
        let b = 0
        let a = 0
        const base = o * taps
        for (let k = 0; k < taps; k += 1) {
          const w = weight[base + k]
          if (w === 0) continue
          const s = index[base + k] * 4
          r += w * scratch[s]
          g += w * scratch[s + 1]
          b += w * scratch[s + 2]
          a += w * scratch[s + 3]
        }
        const d = o * 4
        target[d] = r
        target[d + 1] = g
        target[d + 2] = b
        target[d + 3] = a
      }
    },
  }
}

function verticalStage(source: RowSource, contributions: Contributions): RowSource {
  const rowLength = source.width * 4
  const { taps, index, weight, low, outLength } = contributions
  const cache = new Map<number, Float32Array>()
  const pool: Float32Array[] = []
  const accumulator = new Float64Array(rowLength)
  const rowOf = (y: number): Float32Array => {
    let row = cache.get(y)
    if (!row) {
      row = pool.pop() ?? new Float32Array(rowLength)
      source.read(y, row)
      cache.set(y, row)
    }
    return row
  }
  return {
    width: source.width,
    height: outLength,
    read(y, target) {
      const lowest = low[y]
      for (const [key, row] of cache) {
        if (key < lowest) {
          cache.delete(key)
          pool.push(row)
        }
      }
      accumulator.fill(0)
      const base = y * taps
      for (let k = 0; k < taps; k += 1) {
        const w = weight[base + k]
        if (w === 0) continue
        const row = rowOf(index[base + k])
        for (let j = 0; j < rowLength; j += 1) accumulator[j] += w * row[j]
      }
      for (let j = 0; j < rowLength; j += 1) target[j] = accumulator[j]
    },
  }
}

type AxisFilter = 'nearest' | 'bilinear' | 'bicubic' | 'lanczos3' | 'area'

/** The filter stages for one axis, as [method, intermediate length] steps. */
function axisPlan(method: ResampleMethod, inLength: number, outLength: number): { filter: AxisFilter; length: number }[] {
  if (inLength === outLength) return []
  if (method !== 'auto') return [{ filter: method, length: outLength }]
  const shrink = inLength / outLength
  if (shrink > 2) {
    const intermediate = outLength * 2
    return intermediate < inLength
      ? [{ filter: 'area', length: intermediate }, { filter: 'lanczos3', length: outLength }]
      : [{ filter: 'lanczos3', length: outLength }]
  }
  if (shrink > 1) return [{ filter: 'lanczos3', length: outLength }]
  return [{ filter: 'bicubic', length: outLength }]
}

function contributionsFor(filter: AxisFilter, inLength: number, outLength: number): Contributions {
  if (filter === 'area') return areaContributions(inLength, outLength)
  if (filter === 'nearest') {
    const scale = inLength / outLength
    const index = new Int32Array(outLength)
    const weight = new Float64Array(outLength).fill(1)
    for (let o = 0; o < outLength; o += 1) index[o] = Math.min(inLength - 1, Math.floor((o + 0.5) * scale))
    return finish(outLength, 1, index, weight)
  }
  return kernelContributions(inLength, outLength, KERNELS[filter])
}

const MAX_SIDE = 1 << 17
const MAX_PIXELS = 1 << 28

/**
 * The resumable form of resample (worker handlers run it in chunks). Rows stream through row caches, so
 * process() must be called with consecutive, increasing row ranges.
 */
export function startResample(src: PixelBuffer, width: number, height: number, method: ResampleMethod = 'auto'): RowRun {
  assertBuffer(src)
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
    throw new RangeError(`Invalid target size ${width} x ${height}.`)
  }
  if (width > MAX_SIDE || height > MAX_SIDE || width * height > MAX_PIXELS) {
    throw new RangeError(`The target size ${width} x ${height} is too large.`)
  }
  if (src.width < 1 || src.height < 1) throw new RangeError('Cannot resize an empty image.')
  if (width === src.width && height === src.height) return finishedRun(cloneBuffer(src))
  const chosen: ResampleMethod = method === 'nearest' || method === 'bilinear' || method === 'bicubic' || method === 'lanczos3' || method === 'area' ? method : 'auto'
  if (chosen === 'nearest') return finishedRun(nearestResize(src, width, height))

  let source = pixelRows(src)
  let currentWidth = src.width
  for (const step of axisPlan(chosen, src.width, width)) {
    source = horizontalStage(source, contributionsFor(step.filter, currentWidth, step.length))
    currentWidth = step.length
  }
  let currentHeight = src.height
  for (const step of axisPlan(chosen, src.height, height)) {
    source = verticalStage(source, contributionsFor(step.filter, currentHeight, step.length))
    currentHeight = step.length
  }
  const output = createBuffer(width, height)
  const row = new Float32Array(width * 4)
  const pipeline = source
  let next = 0
  return {
    output,
    rows: height,
    chunkRows: chunkRowsFor(width, 1 << 18),
    process(startRow, endRow) {
      if (startRow !== next) throw new RangeError('Resample rows must be processed in order.')
      for (let y = startRow; y < endRow; y += 1) {
        pipeline.read(y, row)
        unpremultiplyInto(row, 0, output.data, y * width * 4, width)
      }
      next = endRow
    },
  }
}

/**
 * Resizes to width x height (positive integers). Methods: 'nearest' (exact copies), 'bilinear', 'bicubic'
 * (Catmull-Rom), 'lanczos3', 'area' (exact box average) and 'auto' (area pre-filter then Lanczos-3 when
 * shrinking by more than 2x, Lanczos-3 when shrinking, bicubic when enlarging; chosen per axis).
 * Same size returns an identical copy. Throws AbortError when cancelled.
 */
export function resample(src: PixelBuffer, width: number, height: number, method: ResampleMethod = 'auto', options?: OpOptions): PixelBuffer {
  throwIfAborted(options?.signal)
  return runRowsSync(startResample(src, width, height, method), options)
}

function nearestResize(src: PixelBuffer, width: number, height: number): PixelBuffer {
  const out = createBuffer(width, height)
  const xs = new Int32Array(width)
  for (let x = 0; x < width; x += 1) xs[x] = Math.min(src.width - 1, Math.floor(((x + 0.5) * src.width) / width))
  const s = src.data
  const d = out.data
  for (let y = 0; y < height; y += 1) {
    const sy = Math.min(src.height - 1, Math.floor(((y + 0.5) * src.height) / height))
    const sourceRow = sy * src.width
    const targetRow = y * width
    for (let x = 0; x < width; x += 1) {
      const from = (sourceRow + xs[x]) * 4
      const to = (targetRow + x) * 4
      d[to] = s[from]
      d[to + 1] = s[from + 1]
      d[to + 2] = s[from + 2]
      d[to + 3] = s[from + 3]
    }
  }
  return out
}
