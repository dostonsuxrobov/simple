// src/imaging/filters.ts (WP2)
// Neighbourhood filters (design section 5.11). Every filter reads straight RGBA8 and returns a new buffer.
// Blurs, sharpening, mosaic and edge filters work on premultiplied float colour, so transparent pixels never
// bleed dark fringes into visible ones; image edges are clamped (edge pixels repeat). Selection masking is the
// caller's job (buffer.ts mixByMask) or startFilter's. filterMargin() gives the rows of context a horizontal
// stripe needs; with the stripe's absolute first row (applyFilterStripe, startFilter) every filter is
// stripe-exact, so chunked and striped runs match a whole-image run byte for byte. Pure and DOM-free.
import type { FilterSpec, FilterType, MaskBuffer, OpOptions, PixelBuffer } from './types.ts'
import {
  abortError,
  assertBuffer,
  assertMask,
  chunkRowsFor,
  cloneBuffer,
  createBuffer,
  finishedRun,
  fromPremultiplied,
  mixPixelsByMask,
  progressReporter,
  throwIfAborted,
  toPremultiplied,
} from './buffer.ts'
import type { RowRun } from './buffer.ts'
import { hash2, mixSeed } from './random.ts'

// ---------------------------------------------------------------------------------------------
// Catalogue, defaults and margins
// ---------------------------------------------------------------------------------------------

/** Filter menu order: Blur, Sharpen, Noise, Pixelate, Stylize. */
export const FILTER_TYPES: readonly FilterType[] = Object.freeze([
  'gaussian-blur', 'motion-blur', 'unsharp-mask', 'sharpen', 'add-noise', 'median', 'reduce-noise', 'pixelate', 'emboss', 'find-edges',
] as const)

const LABELS: Readonly<Record<FilterType, string>> = Object.freeze({
  'gaussian-blur': 'Gaussian Blur',
  'motion-blur': 'Motion Blur',
  'unsharp-mask': 'Unsharp Mask',
  sharpen: 'Sharpen',
  'add-noise': 'Add Noise',
  median: 'Median',
  'reduce-noise': 'Reduce Noise',
  pixelate: 'Mosaic',
  emboss: 'Emboss',
  'find-edges': 'Find Edges',
})

/** Menu and history label ("Sharpen More" for the strong sharpen spec). */
export function filterLabel(type: FilterType, spec?: FilterSpec): string {
  if (spec && spec.type === 'sharpen' && spec.strength === 'more') return 'Sharpen More'
  return LABELS[type] ?? 'Filter'
}

function createDefault(type: FilterType): FilterSpec {
  switch (type) {
    case 'gaussian-blur': return { type, radius: 2 }
    case 'motion-blur': return { type, angle: 0, distance: 10 }
    case 'unsharp-mask': return { type, amount: 50, radius: 1, threshold: 0 }
    case 'sharpen': return { type, strength: 'normal' }
    case 'add-noise': return { type, amount: 10, distribution: 'uniform', monochromatic: false, seed: 1 }
    case 'median': return { type, radius: 1 }
    case 'reduce-noise': return { type, strength: 6, preserveDetails: 60 }
    case 'pixelate': return { type, cellSize: 8 }
    case 'emboss': return { type, angle: 135, height: 3, amount: 100 }
    case 'find-edges': return { type }
    default: throw new RangeError(`Unknown filter type "${String(type)}".`)
  }
}

/** Default dialog settings for a filter (a fresh object). */
export function defaultFilter<T extends FilterType>(type: T): Extract<FilterSpec, { type: T }> {
  return createDefault(type) as Extract<FilterSpec, { type: T }>
}

function num(value: unknown, fallback: number, min: number, max: number): number {
  const n = typeof value === 'number' && Number.isFinite(value) ? value : fallback
  return n < min ? min : n > max ? max : n
}

/** Box radii of the three-pass Gaussian approximation (Kutskir's boxesForGauss). */
function boxRadii(sigma: number): number[] {
  const passes = 3
  const ideal = Math.sqrt((12 * sigma * sigma) / passes + 1)
  let lower = Math.floor(ideal)
  if (lower % 2 === 0) lower -= 1
  const upper = lower + 2
  const m = Math.round((12 * sigma * sigma - passes * lower * lower - 4 * passes * lower - 3 * passes) / (-4 * lower - 4))
  const radii: number[] = []
  for (let pass = 0; pass < passes; pass += 1) radii.push(((pass < m ? lower : upper) - 1) / 2)
  return radii
}

const EXACT_SIGMA_LIMIT = 2

function blurSupport(sigma: number): number {
  if (!(sigma > 0)) return 0
  if (sigma < EXACT_SIGMA_LIMIT) return Math.max(1, Math.ceil(3 * sigma))
  return Math.max(Math.ceil(3 * sigma), boxRadii(sigma).reduce((sum, radius) => sum + radius, 0))
}

function reduceNoiseRadius(strength: number): number {
  return Math.min(8, Math.max(1, Math.ceil(2 * (1 + strength))))
}

/** Rows of context above and below a stripe that make a striped run identical to a whole-image run. */
export function filterMargin(spec: FilterSpec): number {
  switch (spec.type) {
    case 'gaussian-blur': return blurSupport(num(spec.radius, 2, 0.1, 250))
    case 'motion-blur': return Math.ceil(num(spec.distance, 10, 1, 999) / 2) + 1
    case 'unsharp-mask': return blurSupport(num(spec.radius, 1, 0.1, 250))
    case 'sharpen': return 1
    case 'add-noise': return 0
    case 'median': return Math.round(num(spec.radius, 1, 1, 100))
    case 'reduce-noise': return reduceNoiseRadius(num(spec.strength, 6, 0, 10))
    case 'pixelate': return Math.round(num(spec.cellSize, 8, 2, 200))
    case 'emboss': return Math.ceil(num(spec.height, 3, 1, 10)) + 1
    case 'find-edges': return 1
    default: return 0
  }
}

/**
 * Margin for the worker client's stripe splitting, which cannot tell a stripe its absolute first row; null
 * when the filter depends on absolute rows: Add Noise (noise pattern), Mosaic (cell grid) and steep Motion
 * Blur (line digitisation). Use as `client.setStripeMargin('filter', (input) => stripeMargin(input.spec))`.
 */
export function stripeMargin(spec: FilterSpec): number | null {
  if (spec.type === 'add-noise' || spec.type === 'pixelate') return null
  if (spec.type === 'motion-blur') {
    const theta = (num(spec.angle, 0, -360, 360) * Math.PI) / 180
    if (Math.abs(Math.sin(theta)) > Math.abs(Math.cos(theta))) return null
  }
  return filterMargin(spec)
}

// ---------------------------------------------------------------------------------------------
// Progress and cancellation across passes
// ---------------------------------------------------------------------------------------------

interface Job {
  readonly signal: AbortSignal | undefined
  readonly report: (fraction: number) => void
  readonly units: number
  done: number
}

function createJob(options: OpOptions | undefined, units: number): Job {
  return { signal: options?.signal, report: progressReporter(options), units: Math.max(1, units), done: 0 }
}

function tick(job: Job, units: number): void {
  if (job.signal && job.signal.aborted) throw abortError()
  job.done += units
  job.report(Math.min(0.99, job.done / job.units))
}

// ---------------------------------------------------------------------------------------------
// Gaussian machinery on premultiplied floats
// ---------------------------------------------------------------------------------------------

/** Normalised Gaussian weights for k in [-R, R], R = ceil(3 sigma) (at least 1). */
export function gaussianKernel(sigma: number): Float64Array {
  const radius = Math.max(1, Math.ceil(3 * sigma))
  const kernel = new Float64Array(2 * radius + 1)
  let sum = 0
  for (let k = -radius; k <= radius; k += 1) {
    const w = Math.exp(-(k * k) / (2 * sigma * sigma))
    kernel[k + radius] = w
    sum += w
  }
  for (let index = 0; index < kernel.length; index += 1) kernel[index] /= sum
  return kernel
}

// Passes run on Float64 lines (one row, or one column of a strip): V8 keeps them monomorphic and fast.
// Lines are interleaved RGBA, premultiplied, 0..255 units; ends are clamped (edge samples repeat).

/** Box average of radius r along a line of `count` RGBA samples. */
function boxPass(src: Float64Array, dst: Float64Array, count: number, r: number): void {
  const inv = 1 / (2 * r + 1)
  const last = (count - 1) * 4
  let a0 = (r + 1) * src[0]
  let a1 = (r + 1) * src[1]
  let a2 = (r + 1) * src[2]
  let a3 = (r + 1) * src[3]
  for (let k = 1; k <= r; k += 1) {
    const s = (k < count ? k : count - 1) * 4
    a0 += src[s]
    a1 += src[s + 1]
    a2 += src[s + 2]
    a3 += src[s + 3]
  }
  let x = 0
  // Left edge: the sample leaving the window is clamped to index 0.
  const endLeft = Math.min(count, r + 1)
  for (; x < endLeft; x += 1) {
    const d = x * 4
    dst[d] = a0 * inv
    dst[d + 1] = a1 * inv
    dst[d + 2] = a2 * inv
    dst[d + 3] = a3 * inv
    const add = x + r + 1
    const s = (add < count ? add : count - 1) * 4
    a0 += src[s] - src[0]
    a1 += src[s + 1] - src[1]
    a2 += src[s + 2] - src[2]
    a3 += src[s + 3] - src[3]
  }
  // Middle: no clamping.
  const endMiddle = count - r - 1
  for (; x < endMiddle; x += 1) {
    const d = x * 4
    dst[d] = a0 * inv
    dst[d + 1] = a1 * inv
    dst[d + 2] = a2 * inv
    dst[d + 3] = a3 * inv
    const s = d + (r + 1) * 4
    const t = d - r * 4
    a0 += src[s] - src[t]
    a1 += src[s + 1] - src[t + 1]
    a2 += src[s + 2] - src[t + 2]
    a3 += src[s + 3] - src[t + 3]
  }
  // Right edge: the sample entering the window is clamped to the last index.
  for (; x < count; x += 1) {
    const d = x * 4
    dst[d] = a0 * inv
    dst[d + 1] = a1 * inv
    dst[d + 2] = a2 * inv
    dst[d + 3] = a3 * inv
    const sub = x - r
    const t = (sub > 0 ? sub : 0) * 4
    a0 += src[last] - src[t]
    a1 += src[last + 1] - src[t + 1]
    a2 += src[last + 2] - src[t + 2]
    a3 += src[last + 3] - src[t + 3]
  }
}

/** Exact convolution with a symmetric kernel along a line of `count` RGBA samples. */
function kernelPass(src: Float64Array, dst: Float64Array, count: number, kernel: Float64Array): void {
  const radius = (kernel.length - 1) / 2
  const taps = kernel.length
  const last = count - 1
  const leftEnd = Math.min(count, radius)
  const rightStart = Math.max(leftEnd, count - radius)
  for (let x = 0; x < count; x += 1) {
    if (x === leftEnd && x < rightStart) {
      // Middle: every tap is inside the line.
      for (; x < rightStart; x += 1) {
        let r = 0
        let g = 0
        let b = 0
        let a = 0
        let s = (x - radius) * 4
        for (let k = 0; k < taps; k += 1, s += 4) {
          const w = kernel[k]
          r += w * src[s]
          g += w * src[s + 1]
          b += w * src[s + 2]
          a += w * src[s + 3]
        }
        const d = x * 4
        dst[d] = r
        dst[d + 1] = g
        dst[d + 2] = b
        dst[d + 3] = a
      }
      if (x >= count) break
    }
    let r = 0
    let g = 0
    let b = 0
    let a = 0
    for (let k = -radius; k <= radius; k += 1) {
      const sx = x + k
      const s = (sx < 0 ? 0 : sx > last ? last : sx) * 4
      const w = kernel[k + radius]
      r += w * src[s]
      g += w * src[s + 1]
      b += w * src[s + 2]
      a += w * src[s + 3]
    }
    const d = x * 4
    dst[d] = r
    dst[d + 1] = g
    dst[d + 2] = b
    dst[d + 3] = a
  }
}

/** Exact separable kernel below sigma 2, three box passes (sizes from sigma) otherwise. */
interface BlurPlan {
  readonly kernel: Float64Array | null
  readonly radii: readonly number[]
  /** Samples of edge padding per side (sum of the box radii; 0 for the exact kernel, which clamps itself). */
  readonly pad: number
}

function blurPlan(sigma: number): BlurPlan {
  if (sigma < EXACT_SIGMA_LIMIT) return { kernel: gaussianKernel(sigma), radii: [], pad: 0 }
  const radii = boxRadii(sigma).filter((radius) => radius > 0)
  return { kernel: null, radii, pad: radii.reduce((sum, radius) => sum + radius, 0) }
}

/**
 * Blurs a line of `count` samples stored at sample offset plan.pad of `a` (a and b hold count + 2 * pad
 * samples) and returns whichever array holds the result at the same offset. The padding is filled with the
 * edge samples first, so the box cascade equals a convolution of the edge-extended line (each pass clamping
 * already blurred values would over-weight the edge pixel).
 */
function blurLine(a: Float64Array, b: Float64Array, count: number, plan: BlurPlan): Float64Array {
  if (plan.kernel) {
    kernelPass(a, b, count, plan.kernel)
    return b
  }
  const pad = plan.pad
  const first = pad * 4
  const last = (pad + count - 1) * 4
  for (let k = 0; k < pad; k += 1) {
    const before = k * 4
    const after = (pad + count + k) * 4
    a[before] = a[first]
    a[before + 1] = a[first + 1]
    a[before + 2] = a[first + 2]
    a[before + 3] = a[first + 3]
    a[after] = a[last]
    a[after + 1] = a[last + 1]
    a[after + 2] = a[last + 2]
    a[after + 3] = a[last + 3]
  }
  const total = count + 2 * pad
  let from = a
  let to = b
  for (const radius of plan.radii) {
    boxPass(from, to, total, radius)
    const swap = from
    from = to
    to = swap
  }
  return from
}

/** Premultiplies row `y` of `src` into `line` (Float64) starting at sample `offset`. */
function premultiplyRow(src: PixelBuffer, y: number, line: Float64Array, offset: number): void {
  const data = src.data
  const base = y * src.width * 4
  const shift = offset * 4
  for (let j = 0; j < src.width * 4; j += 4) {
    const a = data[base + j + 3]
    const d = shift + j
    if (a === 255) {
      line[d] = data[base + j]
      line[d + 1] = data[base + j + 1]
      line[d + 2] = data[base + j + 2]
      line[d + 3] = 255
    } else if (a === 0) {
      line[d] = 0
      line[d + 1] = 0
      line[d + 2] = 0
      line[d + 3] = 0
    } else {
      const k = a / 255
      line[d] = data[base + j] * k
      line[d + 1] = data[base + j + 1] * k
      line[d + 2] = data[base + j + 2] * k
      line[d + 3] = a
    }
  }
}

/** Horizontal half of the blur: every row premultiplied and blurred into a float image. */
function blurRows(src: PixelBuffer, plan: BlurPlan, job: Job): Float32Array {
  const { width, height } = src
  const rows = new Float32Array(width * height * 4)
  const padded = (width + 2 * plan.pad) * 4
  const a = new Float64Array(padded)
  const b = new Float64Array(padded)
  const start = plan.pad * 4
  for (let y = 0; y < height; y += 1) {
    premultiplyRow(src, y, a, plan.pad)
    rows.set(blurLine(a, b, width, plan).subarray(start, start + width * 4), y * width * 4)
    if ((y & 63) === 63) tick(job, 64)
  }
  tick(job, height & 63)
  return rows
}

const STRIP_COLUMNS = 16

/**
 * Vertical half of the blur, in cache-friendly strips of columns. For each strip, `emit(x0, columns, strip,
 * columnLength)` receives the blurred columns: column c, row y at strip[c * columnLength + y * 4].
 */
function blurColumns(
  rows: Float32Array,
  width: number,
  height: number,
  plan: BlurPlan,
  job: Job,
  emit: (x0: number, columns: number, strip: Float64Array, columnLength: number) => void,
): void {
  const columnLength = height * 4
  const strip = new Float64Array(STRIP_COLUMNS * columnLength)
  const padded = (height + 2 * plan.pad) * 4
  const a = new Float64Array(padded)
  const b = new Float64Array(padded)
  const start = plan.pad * 4
  for (let x0 = 0; x0 < width; x0 += STRIP_COLUMNS) {
    const columns = Math.min(STRIP_COLUMNS, width - x0)
    for (let y = 0; y < height; y += 1) {
      const rowBase = (y * width + x0) * 4
      for (let c = 0; c < columns; c += 1) {
        const s = rowBase + c * 4
        const d = c * columnLength + y * 4
        strip[d] = rows[s]
        strip[d + 1] = rows[s + 1]
        strip[d + 2] = rows[s + 2]
        strip[d + 3] = rows[s + 3]
      }
    }
    for (let c = 0; c < columns; c += 1) {
      const column = strip.subarray(c * columnLength, (c + 1) * columnLength)
      a.set(column, start)
      column.set(blurLine(a, b, height, plan).subarray(start, start + columnLength))
    }
    emit(x0, columns, strip, columnLength)
    tick(job, (height * columns) / width)
  }
}

// ---------------------------------------------------------------------------------------------
// Individual filters
// ---------------------------------------------------------------------------------------------

function gaussianBlur(src: PixelBuffer, sigma: number, options: OpOptions | undefined): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height * 2)
  const plan = blurPlan(sigma)
  const rows = blurRows(src, plan, job)
  const out = createBuffer(width, height)
  const o = out.data
  blurColumns(rows, width, height, plan, job, (x0, columns, strip, columnLength) => {
    for (let y = 0; y < height; y += 1) {
      const rowBase = (y * width + x0) * 4
      for (let c = 0; c < columns; c += 1) {
        const s = c * columnLength + y * 4
        const d = rowBase + c * 4
        let a = strip[s + 3]
        if (!(a > 0.5)) {
          o[d] = 0; o[d + 1] = 0; o[d + 2] = 0; o[d + 3] = 0
          continue
        }
        if (a > 255) a = 255
        o[d + 3] = a
        if (a === 255) {
          o[d] = strip[s]; o[d + 1] = strip[s + 1]; o[d + 2] = strip[s + 2]
        } else {
          const k = 255 / a
          o[d] = strip[s] * k; o[d + 1] = strip[s + 1] * k; o[d + 2] = strip[s + 2] * k
        }
      }
    }
  })
  return out
}

/** v + amount * (v - gauss(v)) on premultiplied colour where |v - gauss| >= threshold; alpha is kept. */
function unsharpMask(src: PixelBuffer, amount: number, sigma: number, threshold: number, options: OpOptions | undefined): PixelBuffer {
  const { width, height, data } = src
  const job = createJob(options, height * 2)
  const plan = blurPlan(sigma)
  const rows = blurRows(src, plan, job)
  const out = createBuffer(width, height)
  const o = out.data
  const k = amount / 100
  blurColumns(rows, width, height, plan, job, (x0, columns, strip, columnLength) => {
    for (let y = 0; y < height; y += 1) {
      const rowBase = (y * width + x0) * 4
      for (let c = 0; c < columns; c += 1) {
        const s = c * columnLength + y * 4
        const d = rowBase + c * 4
        const a = data[d + 3]
        o[d + 3] = a
        if (a === 0) continue
        const scale = a / 255
        for (let channel = 0; channel < 3; channel += 1) {
          const v = data[d + channel] * scale
          const diff = v - strip[s + channel]
          let sharpened = diff >= threshold || -diff >= threshold ? v + k * diff : v
          if (sharpened > a) sharpened = a
          o[d + channel] = sharpened / scale
        }
      }
    }
  })
  return out
}

function sharpen(src: PixelBuffer, strength: 'normal' | 'more', options: OpOptions | undefined): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height)
  const k = strength === 'more' ? 0.6 : 0.25
  const p = toPremultiplied(src)
  const out = new Float32Array(p.length)
  const rowLength = width * 4
  for (let y = 0; y < height; y += 1) {
    const up = (y > 0 ? y - 1 : 0) * rowLength
    const row = y * rowLength
    const down = (y < height - 1 ? y + 1 : height - 1) * rowLength
    for (let x = 0; x < width; x += 1) {
      const left = (x > 0 ? x - 1 : 0) * 4
      const centre = x * 4
      const right = (x < width - 1 ? x + 1 : width - 1) * 4
      for (let c = 0; c < 3; c += 1) {
        out[row + centre + c] = (1 + 4 * k) * p[row + centre + c]
          - k * (p[up + centre + c] + p[down + centre + c] + p[row + left + c] + p[row + right + c])
      }
      out[row + centre + 3] = p[row + centre + 3]
    }
    if ((y & 63) === 63) tick(job, 64)
  }
  return fromPremultiplied(out, width, height)
}

function addNoise(src: PixelBuffer, spec: Extract<FilterSpec, { type: 'add-noise' }>, options: OpOptions | undefined, rowOffset: number): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height)
  const amplitude = (num(spec.amount, 10, 0.1, 400) / 100) * 127.5
  const gaussian = spec.distribution === 'gaussian'
  const sigma = amplitude / 2.5
  const mono = Boolean(spec.monochromatic)
  const seed = num(spec.seed, 1, -2147483648, 4294967295) >>> 0
  const seedB = mixSeed(seed, 0x2545f491)
  const out = cloneBuffer(src)
  const data = out.data
  const sample = (x: number, y: number): number => {
    const u = hash2(x, y, seed)
    if (!gaussian) return (2 * u - 1) * amplitude
    const v = hash2(x, y, seedB)
    return sigma * Math.sqrt(-2 * Math.log(1 - u)) * Math.cos(2 * Math.PI * v)
  }
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = (y * width + x) * 4
      if (data[i + 3] === 0) continue
      if (mono) {
        const n = sample(x * 4, y + rowOffset)
        data[i] += n
        data[i + 1] += n
        data[i + 2] += n
      } else {
        data[i] += sample(x * 4, y + rowOffset)
        data[i + 1] += sample(x * 4 + 1, y + rowOffset)
        data[i + 2] += sample(x * 4 + 2, y + rowOffset)
      }
    }
    if ((y & 63) === 63) tick(job, 64)
  }
  return out
}

/**
 * Median per channel over a (2r+1)^2 square with clamped edges (Huang's sliding histogram, O(r) per pixel).
 * Colour medians only count visible pixels (alpha > 0), so transparent neighbours never darken an edge;
 * alpha takes the median of all pixels.
 */
function median(src: PixelBuffer, radius: number, options: OpOptions | undefined): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height)
  const data = src.data
  const out = createBuffer(width, height)
  const o = out.data
  const hist = new Int32Array(4 * 256)
  const lastX = width - 1
  const lastY = height - 1
  const windowRows = new Int32Array(2 * radius + 1)
  // Median trackers per channel: current median value and the count of samples below it.
  const med = new Int32Array(4)
  const below = new Int32Array(4)
  let visible = 0
  let all = 0

  const addColumn = (x: number, sign: number) => {
    const cx = (x < 0 ? 0 : x > lastX ? lastX : x) * 4
    for (let k = 0; k < windowRows.length; k += 1) {
      const i = windowRows[k] + cx
      const a = data[i + 3]
      hist[768 + a] += sign
      if (a < med[3]) below[3] += sign
      all += sign
      if (a === 0) continue
      visible += sign
      const r = data[i]
      const g = data[i + 1]
      const b = data[i + 2]
      hist[r] += sign
      hist[256 + g] += sign
      hist[512 + b] += sign
      if (r < med[0]) below[0] += sign
      if (g < med[1]) below[1] += sign
      if (b < med[2]) below[2] += sign
    }
  }

  const find = (channel: number, count: number): number => {
    const target = (count - 1) >> 1
    const base = channel * 256
    let m = med[channel]
    let lt = below[channel]
    while (lt > target && m > 0) {
      m -= 1
      lt -= hist[base + m]
    }
    while (lt + hist[base + m] <= target && m < 255) {
      lt += hist[base + m]
      m += 1
    }
    med[channel] = m
    below[channel] = lt
    return m
  }

  for (let y = 0; y < height; y += 1) {
    hist.fill(0)
    med.fill(0)
    below.fill(0)
    visible = 0
    all = 0
    for (let k = -radius; k <= radius; k += 1) {
      const yy = y + k < 0 ? 0 : y + k > lastY ? lastY : y + k
      windowRows[k + radius] = yy * width * 4
    }
    for (let x = -radius; x <= radius; x += 1) addColumn(x, 1)
    for (let x = 0; x < width; x += 1) {
      const d = (y * width + x) * 4
      o[d + 3] = find(3, all)
      if (visible > 0) {
        o[d] = find(0, visible)
        o[d + 1] = find(1, visible)
        o[d + 2] = find(2, visible)
      }
      if (x < lastX) {
        addColumn(x - radius, -1)
        addColumn(x + radius + 1, 1)
      }
    }
    if ((y & 15) === 15) tick(job, 16)
  }
  return out
}

/** Edge-preserving smoothing: a separable bilateral approximation (horizontal then vertical). */
function reduceNoise(src: PixelBuffer, strength: number, preserveDetails: number, options: OpOptions | undefined): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height * 2)
  const sigmaS = 1 + strength
  const sigmaR = 0.02 + 0.2 * (1 - preserveDetails / 100)
  const radius = reduceNoiseRadius(strength)
  const spatial = new Float64Array(radius + 1)
  for (let k = 0; k <= radius; k += 1) spatial[k] = Math.exp(-(k * k) / (2 * sigmaS * sigmaS))
  // Range weight exp(-q) tabulated for q in [0, 20).
  const RANGE_STEPS = 2048
  const rangeScale = RANGE_STEPS / 20
  const range = new Float64Array(RANGE_STEPS + 1)
  for (let index = 0; index <= RANGE_STEPS; index += 1) range[index] = index === RANGE_STEPS ? 0 : Math.exp(-index / rangeScale)
  const rangeFactor = (1 / (2 * sigmaR * sigmaR * 3 * 255 * 255)) * rangeScale
  const p = toPremultiplied(src)
  const q = new Float32Array(p.length)

  /**
   * One 1-D bilateral pass over rows (horizontal: taps along x) or columns (taps along y), always visiting
   * pixels in row-major order so memory access stays sequential.
   */
  const pass = (input: Float32Array, output: Float32Array, horizontal: boolean) => {
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const centre = (y * width + x) * 4
        const r0 = input[centre]
        const g0 = input[centre + 1]
        const b0 = input[centre + 2]
        const a0 = input[centre + 3]
        let sr = 0
        let sg = 0
        let sb = 0
        let sa = 0
        let sw = 0
        for (let k = -radius; k <= radius; k += 1) {
          let s: number
          if (horizontal) {
            const xx = x + k < 0 ? 0 : x + k >= width ? width - 1 : x + k
            s = (y * width + xx) * 4
          } else {
            const yy = y + k < 0 ? 0 : y + k >= height ? height - 1 : y + k
            s = (yy * width + x) * 4
          }
          const dr = input[s] - r0
          const dg = input[s + 1] - g0
          const db = input[s + 2] - b0
          const da = input[s + 3] - a0
          const qv = (dr * dr + dg * dg + db * db + da * da) * rangeFactor
          const w = spatial[k < 0 ? -k : k] * (qv >= RANGE_STEPS ? 0 : range[qv | 0])
          sr += w * input[s]
          sg += w * input[s + 1]
          sb += w * input[s + 2]
          sa += w * input[s + 3]
          sw += w
        }
        output[centre] = sr / sw
        output[centre + 1] = sg / sw
        output[centre + 2] = sb / sw
        output[centre + 3] = sa / sw
      }
      if ((y & 63) === 63) tick(job, 64)
    }
  }
  pass(p, q, true)
  pass(q, p, false)
  return fromPremultiplied(p, width, height)
}

/** Mosaic: alpha-weighted average per cell, cells aligned to the buffer origin. */
function pixelate(src: PixelBuffer, cellSize: number, options: OpOptions | undefined, rowOffset: number): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height)
  const data = src.data
  const out = createBuffer(width, height)
  const o = out.data
  const phase = ((rowOffset % cellSize) + cellSize) % cellSize
  for (let cellTop = -phase; cellTop < height; cellTop += cellSize) {
    const top = Math.max(0, cellTop)
    const bottom = Math.min(height, cellTop + cellSize)
    for (let left = 0; left < width; left += cellSize) {
      const right = Math.min(width, left + cellSize)
      let sr = 0
      let sg = 0
      let sb = 0
      let sa = 0
      for (let y = top; y < bottom; y += 1) {
        for (let x = left; x < right; x += 1) {
          const i = (y * width + x) * 4
          const a = data[i + 3]
          sr += data[i] * a
          sg += data[i + 1] * a
          sb += data[i + 2] * a
          sa += a
        }
      }
      const count = (bottom - top) * (right - left)
      const alpha = sa / count
      const r = sa > 0 ? sr / sa : 0
      const g = sa > 0 ? sg / sa : 0
      const b = sa > 0 ? sb / sa : 0
      for (let y = top; y < bottom; y += 1) {
        for (let x = left; x < right; x += 1) {
          const i = (y * width + x) * 4
          o[i] = r
          o[i + 1] = g
          o[i + 2] = b
          o[i + 3] = alpha
        }
      }
    }
    tick(job, bottom - top)
  }
  return out
}

function bilinearPlane(plane: Float32Array, width: number, height: number, x: number, y: number): number {
  const fx = x < 0 ? 0 : x > width - 1 ? width - 1 : x
  const fy = y < 0 ? 0 : y > height - 1 ? height - 1 : y
  const x0 = Math.floor(fx)
  const y0 = Math.floor(fy)
  const x1 = x0 < width - 1 ? x0 + 1 : x0
  const y1 = y0 < height - 1 ? y0 + 1 : y0
  const tx = fx - x0
  const ty = fy - y0
  const top = plane[y0 * width + x0] + (plane[y0 * width + x1] - plane[y0 * width + x0]) * tx
  const bottom = plane[y1 * width + x0] + (plane[y1 * width + x1] - plane[y1 * width + x0]) * tx
  return top + (bottom - top) * ty
}

/** Emboss: 128 + amount * (Y(p + d * height) - Y(p - d * height)) along the angle; grey output, alpha kept. */
function emboss(src: PixelBuffer, angle: number, heightPx: number, amount: number, options: OpOptions | undefined): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height)
  const data = src.data
  const plane = new Float32Array(width * height)
  for (let pixel = 0; pixel < plane.length; pixel += 1) {
    const i = pixel * 4
    plane[pixel] = ((0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]) * data[i + 3]) / 255
  }
  const theta = (angle * Math.PI) / 180
  const dx = Math.cos(theta) * heightPx
  const dy = -Math.sin(theta) * heightPx
  const k = amount / 100
  const out = createBuffer(width, height)
  const o = out.data
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const v = 128 + k * (bilinearPlane(plane, width, height, x + dx, y + dy) - bilinearPlane(plane, width, height, x - dx, y - dy))
      const i = (y * width + x) * 4
      o[i] = v
      o[i + 1] = v
      o[i + 2] = v
      o[i + 3] = data[i + 3]
    }
    if ((y & 63) === 63) tick(job, 64)
  }
  return out
}

/** Find Edges: per-channel Sobel magnitude on premultiplied colour, 255 - magnitude (dark edges on white). */
function findEdges(src: PixelBuffer, options: OpOptions | undefined): PixelBuffer {
  const { width, height } = src
  const job = createJob(options, height)
  const p = toPremultiplied(src)
  const out = createBuffer(width, height)
  const o = out.data
  const rowLength = width * 4
  for (let y = 0; y < height; y += 1) {
    const up = (y > 0 ? y - 1 : 0) * rowLength
    const row = y * rowLength
    const down = (y < height - 1 ? y + 1 : height - 1) * rowLength
    for (let x = 0; x < width; x += 1) {
      const left = (x > 0 ? x - 1 : 0) * 4
      const centre = x * 4
      const right = (x < width - 1 ? x + 1 : width - 1) * 4
      for (let c = 0; c < 3; c += 1) {
        const gx = (p[up + right + c] + 2 * p[row + right + c] + p[down + right + c]) - (p[up + left + c] + 2 * p[row + left + c] + p[down + left + c])
        const gy = (p[down + left + c] + 2 * p[down + centre + c] + p[down + right + c]) - (p[up + left + c] + 2 * p[up + centre + c] + p[up + right + c])
        o[row + centre + c] = 255 - Math.sqrt(gx * gx + gy * gy)
      }
      o[row + centre + 3] = src.data[row + centre + 3]
    }
    if ((y & 63) === 63) tick(job, 64)
  }
  return out
}

// ---------------------------------------------------------------------------------------------
// Motion blur: box average along digital lines, O(N) for any distance
// ---------------------------------------------------------------------------------------------

/** First index in [0, n) where `test` holds for a monotone predicate (false...false true...true); n if none. */
function firstWhere(n: number, test: (index: number) => boolean): number {
  let low = 0
  let high = n
  while (low < high) {
    const mid = (low + high) >> 1
    if (test(mid)) high = mid
    else low = mid + 1
  }
  return low
}

function motionBlur(src: PixelBuffer, angle: number, distance: number, options: OpOptions | undefined, rowOffset: number): PixelBuffer {
  const { width, height } = src
  const theta = (angle * Math.PI) / 180
  const ux = Math.cos(theta)
  const uy = -Math.sin(theta)
  const xMajor = Math.abs(ux) >= Math.abs(uy)
  // Consecutive samples of a digital line are one major-axis pixel apart, i.e. 1 / |cos| (or 1 / |sin|)
  // pixels along the motion, so `distance` pixels hold distance * |cos| (or |sin|) samples.
  const n = distance * (xMajor ? Math.abs(ux) : Math.abs(uy))
  if (!(n > 1)) return cloneBuffer(src)
  const major = xMajor ? width : height
  const minor = xMajor ? height : width
  const slope = xMajor ? uy / ux : ux / uy
  const offset = new Int32Array(major)
  // Steep lines run along y: digitise them in absolute rows so stripes match a whole-image run.
  for (let t = 0; t < major; t += 1) offset[t] = Math.round(slope * (xMajor ? t : t + rowOffset))
  let minOffset = offset[0]
  let maxOffset = offset[0]
  for (let t = 1; t < major; t += 1) {
    if (offset[t] < minOffset) minOffset = offset[t]
    if (offset[t] > maxOffset) maxOffset = offset[t]
  }
  const increasing = slope >= 0
  const lines = minor + maxOffset - minOffset
  const job = createJob(options, lines)
  const half = (n - 1) / 2
  const a = Math.floor(half)
  const f = half - a
  const inv = 1 / n
  const p = toPremultiplied(src)
  const out = new Float32Array(p.length)
  const line = new Float64Array(major * 4)
  const prefix = new Float64Array((major + 1) * 4)
  const index = new Int32Array(major)
  let processed = 0
  for (let start = -maxOffset; start <= minor - 1 - minOffset; start += 1) {
    // Line pixels: t with 0 <= start + offset[t] <= minor - 1 (a contiguous range because offset is monotone).
    const low = -start
    const high = minor - 1 - start
    const t0 = increasing ? firstWhere(major, (t) => offset[t] >= low) : firstWhere(major, (t) => offset[t] <= high)
    const t1 = (increasing ? firstWhere(major, (t) => offset[t] > high) : firstWhere(major, (t) => offset[t] < low)) - 1
    processed += 1
    if ((processed & 127) === 0) tick(job, 128)
    if (t1 < t0) continue
    const length = t1 - t0 + 1
    for (let k = 0; k < length; k += 1) {
      const t = t0 + k
      const m = start + offset[t]
      const pixel = xMajor ? m * width + t : t * width + m
      index[k] = pixel
      const s = pixel * 4
      const d = k * 4
      line[d] = p[s]
      line[d + 1] = p[s + 1]
      line[d + 2] = p[s + 2]
      line[d + 3] = p[s + 3]
      prefix[d + 4] = prefix[d] + p[s]
      prefix[d + 5] = prefix[d + 1] + p[s + 1]
      prefix[d + 6] = prefix[d + 2] + p[s + 2]
      prefix[d + 7] = prefix[d + 3] + p[s + 3]
    }
    const lastIndex = length - 1
    const lastBase = lastIndex * 4
    for (let k = 0; k < length; k += 1) {
      const lo = k - a
      const hi = k + a
      const loClamped = lo < 0 ? 0 : lo
      const hiClamped = hi > lastIndex ? lastIndex : hi
      const before = lo < 0 ? -lo : 0
      const after = hi > lastIndex ? hi - lastIndex : 0
      const fracLow = (k - a - 1 < 0 ? 0 : k - a - 1) * 4
      const fracHigh = (k + a + 1 > lastIndex ? lastIndex : k + a + 1) * 4
      const d = index[k] * 4
      for (let c = 0; c < 4; c += 1) {
        let sum = prefix[(hiClamped + 1) * 4 + c] - prefix[loClamped * 4 + c] + before * line[c] + after * line[lastBase + c]
        if (f > 0) sum += f * (line[fracLow + c] + line[fracHigh + c])
        out[d + c] = sum * inv
      }
    }
  }
  return fromPremultiplied(out, width, height)
}

// ---------------------------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------------------------

function runFilter(src: PixelBuffer, spec: FilterSpec, options: OpOptions | undefined, rowOffset: number): PixelBuffer {
  switch (spec.type) {
    case 'gaussian-blur':
      return gaussianBlur(src, num(spec.radius, 2, 0.1, 250), options)
    case 'motion-blur':
      return motionBlur(src, num(spec.angle, 0, -360, 360), num(spec.distance, 10, 1, 999), options, rowOffset)
    case 'unsharp-mask':
      return unsharpMask(src, num(spec.amount, 50, 1, 500), num(spec.radius, 1, 0.1, 250), num(spec.threshold, 0, 0, 255), options)
    case 'sharpen':
      return sharpen(src, spec.strength === 'more' ? 'more' : 'normal', options)
    case 'add-noise':
      return addNoise(src, spec, options, rowOffset)
    case 'median':
      return median(src, Math.round(num(spec.radius, 1, 1, 100)), options)
    case 'reduce-noise':
      return reduceNoise(src, num(spec.strength, 6, 0, 10), num(spec.preserveDetails, 60, 0, 100), options)
    case 'pixelate':
      return pixelate(src, Math.round(num(spec.cellSize, 8, 2, 200)), options, rowOffset)
    case 'emboss':
      return emboss(src, num(spec.angle, 135, -360, 360), num(spec.height, 3, 1, 10), num(spec.amount, 100, 1, 500), options)
    case 'find-edges':
      return findEdges(src, options)
    default:
      throw new RangeError(`Unknown filter type "${String((spec as { type?: unknown }).type)}".`)
  }
}

function checkSpec(spec: FilterSpec): void {
  if (!spec || typeof spec !== 'object') throw new RangeError('The filter is missing.')
  if (!Object.prototype.hasOwnProperty.call(LABELS, spec.type)) throw new RangeError(`Unknown filter type "${String((spec as { type?: unknown }).type)}".`)
}

/** Runs one filter on the whole buffer and returns a new buffer. Throws AbortError when cancelled. */
export function applyFilter(src: PixelBuffer, spec: FilterSpec, options?: OpOptions): PixelBuffer {
  return applyFilterStripe(src, spec, 0, options)
}

/**
 * applyFilter for a horizontal stripe of a larger image whose first row is absolute row `rowOffset`.
 * Position-dependent filters (Add Noise, Mosaic cells, steep Motion Blur lines) use absolute rows, so a
 * stripe padded by filterMargin(spec) rows reproduces the whole-image result exactly.
 */
export function applyFilterStripe(stripe: PixelBuffer, spec: FilterSpec, rowOffset: number, options?: OpOptions): PixelBuffer {
  assertBuffer(stripe)
  checkSpec(spec)
  throwIfAborted(options?.signal)
  if (stripe.width === 0 || stripe.height === 0) return cloneBuffer(stripe)
  const result = runFilter(stripe, spec, options, Math.round(Number(rowOffset) || 0))
  throwIfAborted(options?.signal)
  options?.onProgress?.(1)
  return result
}

/**
 * The resumable form of applyFilter, optionally mixed back through a selection mask (0..255). Rows are
 * filtered in stripes padded by filterMargin(spec), so any chunking gives the same bytes as one run.
 */
export function startFilter(src: PixelBuffer, spec: FilterSpec, mask: MaskBuffer | null = null): RowRun {
  assertBuffer(src)
  checkSpec(spec)
  if (mask) assertMask(mask, src.width, src.height)
  const { width, height } = src
  if (width === 0 || height === 0) return finishedRun(cloneBuffer(src))
  const margin = filterMargin(spec)
  const rowBytes = width * 4
  let chunkRows = chunkRowsFor(width, 1 << 20, Math.max(64, 4 * margin))
  if (chunkRows + 2 * margin >= height) chunkRows = height
  const output = createBuffer(width, height)
  return {
    output,
    rows: height,
    chunkRows,
    process(startRow, endRow) {
      const from = Math.max(0, startRow - margin)
      const to = Math.min(height, endRow + margin)
      const stripe = from === 0 && to === height ? src : { width, height: to - from, data: src.data.subarray(from * rowBytes, to * rowBytes) }
      const result = runFilter(stripe, spec, undefined, from)
      output.data.set(result.data.subarray((startRow - from) * rowBytes, (endRow - from) * rowBytes), startRow * rowBytes)
      if (mask) mixPixelsByMask(src, output, mask, 1, startRow * width, endRow * width, output.data)
    },
  }
}
