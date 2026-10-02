// src/imaging/buffer.ts (WP2)
// Pixel-buffer basics: allocation, copies, exact geometric permutations, masked mixing, rectangles,
// plus the small helpers every imaging operation shares (validation, abort, progress, premultiplied
// float conversion). Pure and DOM-free; buffers are straight-alpha RGBA8 (types.ts PixelBuffer).
// Every function returns new buffers and leaves its inputs untouched unless its name says otherwise
// (pasteBuffer writes into dst).
import type { IntRect, MaskBuffer, OpOptions, PixelBuffer, Rgba8 } from './types.ts'

// ---------------------------------------------------------------------------------------------
// Validation, abort and progress helpers
// ---------------------------------------------------------------------------------------------

function isDimension(value: number): boolean {
  return Number.isInteger(value) && value >= 0
}

/** Throws a RangeError unless `buffer` is a consistent PixelBuffer. */
export function assertBuffer(buffer: PixelBuffer, name = 'image'): void {
  if (!buffer || typeof buffer !== 'object') throw new RangeError(`The ${name} is missing.`)
  const { width, height, data } = buffer
  if (!isDimension(width) || !isDimension(height)) throw new RangeError(`The ${name} has an invalid size (${width} x ${height}).`)
  if (!data || data.length !== width * height * 4) {
    throw new RangeError(`The ${name} pixel data does not match its size (${data ? data.length : 0} bytes for ${width} x ${height}).`)
  }
}

/** Throws a RangeError unless `mask` is a MaskBuffer of exactly width x height. */
export function assertMask(mask: MaskBuffer, width: number, height: number, name = 'mask'): void {
  if (!mask || typeof mask !== 'object') throw new RangeError(`The ${name} is missing.`)
  if (mask.width !== width || mask.height !== height) {
    throw new RangeError(`The ${name} is ${mask.width} x ${mask.height} but the image is ${width} x ${height}.`)
  }
  if (!mask.data || mask.data.length !== width * height) {
    throw new RangeError(`The ${name} data does not match its size (${mask.data ? mask.data.length : 0} values for ${width} x ${height}).`)
  }
}

/** An Error named 'AbortError' (the cancellation contract of OpOptions.signal). */
export function abortError(message = 'The operation was cancelled.'): Error {
  const error = new Error(message)
  error.name = 'AbortError'
  return error
}

export function throwIfAborted(signal: AbortSignal | undefined | null): void {
  if (signal && signal.aborted) throw abortError()
}

/**
 * Wraps OpOptions.onProgress so an operation can report freely: forwards at most about 20 increasing
 * values (steps of 5%) plus the final 1.
 */
export function progressReporter(options: OpOptions | undefined | null): (fraction: number) => void {
  const onProgress = options && options.onProgress
  if (typeof onProgress !== 'function') return () => {}
  let last = 0
  return (fraction: number) => {
    const value = fraction >= 1 ? 1 : fraction > 0 ? fraction : 0
    if (value <= last) return
    if (value < 1 && value - last < 0.05) return
    last = value
    onProgress(value)
  }
}

// ---------------------------------------------------------------------------------------------
// Resumable row runs (lets worker handlers yield between chunks so an 'abort' message can land)
// ---------------------------------------------------------------------------------------------

/**
 * An operation split by output rows. process(start, end) computes rows [start, end) of `output`; callers
 * process rows in increasing order (some runs stream state from one chunk to the next).
 */
export interface RowRun {
  readonly output: PixelBuffer
  readonly rows: number
  /** Rows per chunk that keep one chunk around 10..50 ms of work. */
  readonly chunkRows: number
  process(startRow: number, endRow: number): void
}

/** A run whose result is already complete (identity fast paths). */
export function finishedRun(output: PixelBuffer): RowRun {
  return { output, rows: output.height, chunkRows: Math.max(1, output.height), process() {} }
}

/** Rows per chunk for about `pixelsPerChunk` output pixels, at least `minimum` rows. */
export function chunkRowsFor(width: number, pixelsPerChunk: number, minimum = 1): number {
  return Math.max(minimum, Math.ceil(pixelsPerChunk / Math.max(1, width)))
}

/** Runs every chunk synchronously, checking the signal between chunks and reporting progress. */
export function runRowsSync(run: RowRun, options?: OpOptions): PixelBuffer {
  const signal = options?.signal
  const report = progressReporter(options)
  const step = Math.max(1, run.chunkRows)
  for (let start = 0; start < run.rows; start += step) {
    throwIfAborted(signal)
    const end = Math.min(run.rows, start + step)
    run.process(start, end)
    report(end / run.rows)
  }
  throwIfAborted(signal)
  report(1)
  return run.output
}

/**
 * Like runRowsSync but yields to the event loop between chunks. A module worker only handles its next
 * message (such as 'abort') between tasks, so synchronous handlers could never observe cancellation.
 */
export async function runRowsAsync(run: RowRun, options?: OpOptions): Promise<PixelBuffer> {
  const signal = options?.signal
  const report = progressReporter(options)
  const step = Math.max(1, run.chunkRows)
  for (let start = 0; start < run.rows; start += step) {
    throwIfAborted(signal)
    const end = Math.min(run.rows, start + step)
    run.process(start, end)
    report(end / run.rows)
    if (end < run.rows) await yieldToEventLoop()
  }
  throwIfAborted(signal)
  report(1)
  return run.output
}

let yieldChannel: MessageChannel | null = null
const yieldWaiters: (() => void)[] = []

/** Resolves on a fresh macrotask (setImmediate in Node, a MessageChannel in browsers and workers; no 4 ms clamp). */
export function yieldToEventLoop(): Promise<void> {
  const immediate = (globalThis as { setImmediate?: (callback: () => void) => unknown }).setImmediate
  if (typeof immediate === 'function') return new Promise((resolve) => { immediate(resolve) })
  if (typeof MessageChannel === 'function') {
    if (!yieldChannel) {
      yieldChannel = new MessageChannel()
      yieldChannel.port1.onmessage = () => {
        const next = yieldWaiters.shift()
        if (next) next()
      }
    }
    const channel = yieldChannel
    return new Promise((resolve) => {
      yieldWaiters.push(resolve)
      channel.port2.postMessage(null)
    })
  }
  return new Promise((resolve) => { setTimeout(resolve, 0) })
}

// ---------------------------------------------------------------------------------------------
// Allocation and copies
// ---------------------------------------------------------------------------------------------

/** A new width x height buffer, transparent black unless `fill` is given. */
export function createBuffer(width: number, height: number, fill?: Rgba8): PixelBuffer {
  if (!isDimension(width) || !isDimension(height)) throw new RangeError(`Invalid image size ${width} x ${height}.`)
  const data = new Uint8ClampedArray(width * height * 4)
  if (fill && data.length > 0 && (fill.r || fill.g || fill.b || fill.a)) {
    data[0] = fill.r
    data[1] = fill.g
    data[2] = fill.b
    data[3] = fill.a
    // Doubling copy: endian-agnostic and fast.
    for (let filled = 4; filled < data.length; filled *= 2) data.copyWithin(filled, 0, Math.min(filled, data.length - filled))
  }
  return { width, height, data }
}

export function cloneBuffer(src: PixelBuffer): PixelBuffer {
  return { width: src.width, height: src.height, data: new Uint8ClampedArray(src.data) }
}

/** Pixels of `rect` (may extend past the image or be negative); transparent where it lies outside. */
export function cropBuffer(src: PixelBuffer, rect: IntRect): PixelBuffer {
  const width = Math.max(0, Math.round(rect.width))
  const height = Math.max(0, Math.round(rect.height))
  const left = Math.round(rect.x)
  const top = Math.round(rect.y)
  const out = createBuffer(width, height)
  const x0 = Math.max(0, left)
  const x1 = Math.min(src.width, left + width)
  if (x1 <= x0) return out
  const y0 = Math.max(0, top)
  const y1 = Math.min(src.height, top + height)
  for (let y = y0; y < y1; y += 1) {
    const from = (y * src.width + x0) * 4
    out.data.set(src.data.subarray(from, from + (x1 - x0) * 4), ((y - top) * width + (x0 - left)) * 4)
  }
  return out
}

/** Writes `src` into `dst` at (x, y), replacing pixels (no blending); clipped to `dst`. Mutates dst. */
export function pasteBuffer(dst: PixelBuffer, src: PixelBuffer, x: number, y: number): void {
  const left = Math.round(x)
  const top = Math.round(y)
  const x0 = Math.max(0, left)
  const x1 = Math.min(dst.width, left + src.width)
  if (x1 <= x0) return
  const y0 = Math.max(0, top)
  const y1 = Math.min(dst.height, top + src.height)
  for (let row = y0; row < y1; row += 1) {
    const from = ((row - top) * src.width + (x0 - left)) * 4
    dst.data.set(src.data.subarray(from, from + (x1 - x0) * 4), (row * dst.width + x0) * 4)
  }
}

/** A Uint32 view of the pixels when the backing memory is aligned (one element per pixel), else null. */
function pixelWords(data: Uint8ClampedArray): Uint32Array | null {
  if (data.byteOffset % 4 !== 0 || data.length % 4 !== 0) return null
  return new Uint32Array(data.buffer, data.byteOffset, data.length / 4)
}

// ---------------------------------------------------------------------------------------------
// Exact geometric permutations
// ---------------------------------------------------------------------------------------------

export function flipHorizontal(src: PixelBuffer): PixelBuffer {
  const { width, height } = src
  const out = createBuffer(width, height)
  const from = pixelWords(src.data)
  const to = pixelWords(out.data)
  for (let y = 0; y < height; y += 1) {
    const row = y * width
    if (from && to) {
      for (let x = 0; x < width; x += 1) to[row + x] = from[row + width - 1 - x]
    } else {
      for (let x = 0; x < width; x += 1) {
        const s = (row + width - 1 - x) * 4
        const d = (row + x) * 4
        out.data[d] = src.data[s]
        out.data[d + 1] = src.data[s + 1]
        out.data[d + 2] = src.data[s + 2]
        out.data[d + 3] = src.data[s + 3]
      }
    }
  }
  return out
}

export function flipVertical(src: PixelBuffer): PixelBuffer {
  const { width, height } = src
  const out = createBuffer(width, height)
  const rowBytes = width * 4
  for (let y = 0; y < height; y += 1) {
    const from = (height - 1 - y) * rowBytes
    out.data.set(src.data.subarray(from, from + rowBytes), y * rowBytes)
  }
  return out
}

/** Quarter turn; the result is height x width. `clockwise` as seen on screen (y points down). */
export function rotate90(src: PixelBuffer, clockwise: boolean): PixelBuffer {
  const { width, height } = src
  const out = createBuffer(height, width)
  const from = pixelWords(src.data)
  const to = pixelWords(out.data)
  // Output pixel (u, v) with u in [0, height), v in [0, width).
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const u = clockwise ? height - 1 - y : y
      const v = clockwise ? x : width - 1 - x
      const s = y * width + x
      const d = v * height + u
      if (from && to) {
        to[d] = from[s]
      } else {
        out.data[d * 4] = src.data[s * 4]
        out.data[d * 4 + 1] = src.data[s * 4 + 1]
        out.data[d * 4 + 2] = src.data[s * 4 + 2]
        out.data[d * 4 + 3] = src.data[s * 4 + 3]
      }
    }
  }
  return out
}

export function rotate180(src: PixelBuffer): PixelBuffer {
  const { width, height } = src
  const out = createBuffer(width, height)
  const count = width * height
  const from = pixelWords(src.data)
  const to = pixelWords(out.data)
  if (from && to) {
    for (let index = 0; index < count; index += 1) to[index] = from[count - 1 - index]
    return out
  }
  for (let index = 0; index < count; index += 1) {
    const s = (count - 1 - index) * 4
    const d = index * 4
    out.data[d] = src.data[s]
    out.data[d + 1] = src.data[s + 1]
    out.data[d + 2] = src.data[s + 2]
    out.data[d + 3] = src.data[s + 3]
  }
  return out
}

/** True when any pixel has alpha below 255. */
export function hasTransparency(src: PixelBuffer): boolean {
  const data = src.data
  for (let index = 3; index < data.length; index += 4) {
    if (data[index] !== 255) return true
  }
  return false
}

// ---------------------------------------------------------------------------------------------
// Masked mixing
// ---------------------------------------------------------------------------------------------

/**
 * original + (processed - original) * mask/255 * opacity, as a new buffer. Where the two alphas differ the
 * mix is done on premultiplied colour so transparent edges never pick up the colour of hidden pixels.
 * mask null = everywhere; opacity is clamped to [0, 1].
 */
export function mixByMask(original: PixelBuffer, processed: PixelBuffer, mask: MaskBuffer | null, opacity: number): PixelBuffer {
  if (original.width !== processed.width || original.height !== processed.height) {
    throw new RangeError(`Cannot mix a ${processed.width} x ${processed.height} result into a ${original.width} x ${original.height} image.`)
  }
  const { width, height } = original
  if (mask) assertMask(mask, width, height)
  const amount = opacity >= 1 ? 1 : opacity > 0 ? opacity : 0
  if (amount === 0) return cloneBuffer(original)
  if (!mask && amount === 1) return cloneBuffer(processed)
  const out = new Uint8ClampedArray(original.data.length)
  mixPixelsByMask(original, processed, mask, amount, 0, width * height, out)
  return { width, height, data: out }
}

/**
 * mixByMask for pixels [start, end) only, written into `out` (which may be processed.data: each pixel is
 * read before it is written). Sizes are not re-validated.
 */
export function mixPixelsByMask(
  original: PixelBuffer,
  processed: PixelBuffer,
  mask: MaskBuffer | null,
  opacity: number,
  start: number,
  end: number,
  out: Uint8ClampedArray,
): void {
  const amount = opacity >= 1 ? 1 : opacity > 0 ? opacity : 0
  const a = original.data
  const b = processed.data
  for (let pixel = start; pixel < end; pixel += 1) {
    const i = pixel * 4
    const coverage = mask ? mask.data[pixel] : 255
    if (coverage === 0) {
      out[i] = a[i]; out[i + 1] = a[i + 1]; out[i + 2] = a[i + 2]; out[i + 3] = a[i + 3]
      continue
    }
    const w = coverage === 255 ? amount : (coverage / 255) * amount
    if (w >= 1) {
      out[i] = b[i]; out[i + 1] = b[i + 1]; out[i + 2] = b[i + 2]; out[i + 3] = b[i + 3]
      continue
    }
    const a0 = a[i + 3]
    const a1 = b[i + 3]
    if (a0 === a1) {
      out[i] = a[i] + (b[i] - a[i]) * w
      out[i + 1] = a[i + 1] + (b[i + 1] - a[i + 1]) * w
      out[i + 2] = a[i + 2] + (b[i + 2] - a[i + 2]) * w
      out[i + 3] = a0
      continue
    }
    const alpha = a0 + (a1 - a0) * w
    out[i + 3] = alpha
    if (alpha <= 0) {
      out[i] = 0; out[i + 1] = 0; out[i + 2] = 0
      continue
    }
    const k0 = (a0 * (1 - w)) / alpha
    const k1 = (a1 * w) / alpha
    out[i] = a[i] * k0 + b[i] * k1
    out[i + 1] = a[i + 1] * k0 + b[i + 1] * k1
    out[i + 2] = a[i + 2] * k0 + b[i + 2] * k1
  }
}

// ---------------------------------------------------------------------------------------------
// Rectangles
// ---------------------------------------------------------------------------------------------

/** Overlap of two rectangles, or null when they do not overlap (touching edges do not overlap). */
export function intersectRect(a: IntRect, b: IntRect): IntRect | null {
  const x = Math.max(a.x, b.x)
  const y = Math.max(a.y, b.y)
  const right = Math.min(a.x + a.width, b.x + b.width)
  const bottom = Math.min(a.y + a.height, b.y + b.height)
  if (right <= x || bottom <= y) return null
  return { x, y, width: right - x, height: bottom - y }
}

/** Bounding box of two rectangles; an empty rectangle (zero width or height) does not contribute. */
export function unionRect(a: IntRect, b: IntRect): IntRect {
  const aEmpty = !(a.width > 0 && a.height > 0)
  const bEmpty = !(b.width > 0 && b.height > 0)
  if (aEmpty && bEmpty) return { x: a.x, y: a.y, width: 0, height: 0 }
  if (aEmpty) return { x: b.x, y: b.y, width: b.width, height: b.height }
  if (bEmpty) return { x: a.x, y: a.y, width: a.width, height: a.height }
  const x = Math.min(a.x, b.x)
  const y = Math.min(a.y, b.y)
  return { x, y, width: Math.max(a.x + a.width, b.x + b.width) - x, height: Math.max(a.y + a.height, b.y + b.height) - y }
}

// ---------------------------------------------------------------------------------------------
// Premultiplied float form (filters, resampling and warps work on it so transparent edges never darken)
// ---------------------------------------------------------------------------------------------

/** RGBA floats in 0..255 units with colour premultiplied: P = c * a / 255 (exact for opaque pixels). */
export function toPremultiplied(src: PixelBuffer): Float32Array {
  const data = src.data
  const out = new Float32Array(data.length)
  for (let i = 0; i < data.length; i += 4) {
    const a = data[i + 3]
    if (a === 255) {
      out[i] = data[i]; out[i + 1] = data[i + 1]; out[i + 2] = data[i + 2]; out[i + 3] = 255
    } else if (a !== 0) {
      const k = a / 255
      out[i] = data[i] * k; out[i + 1] = data[i + 1] * k; out[i + 2] = data[i + 2] * k; out[i + 3] = a
    }
  }
  return out
}

/** Writes premultiplied floats `p` (from index pStart) back as straight RGBA8 into `out` (from byte outStart). */
export function unpremultiplyInto(p: Float32Array | Float64Array, pStart: number, out: Uint8ClampedArray, outStart: number, pixels: number): void {
  for (let n = 0; n < pixels; n += 1) {
    const s = pStart + n * 4
    const d = outStart + n * 4
    let a = p[s + 3]
    // Uint8ClampedArray rounds half to even, so alpha 0.5 is stored as 0: treat it as transparent.
    if (!(a > 0.5)) {
      out[d] = 0; out[d + 1] = 0; out[d + 2] = 0; out[d + 3] = 0
      continue
    }
    if (a > 255) a = 255
    out[d + 3] = a
    if (a === 255) {
      out[d] = p[s]; out[d + 1] = p[s + 1]; out[d + 2] = p[s + 2]
    } else {
      const k = 255 / a
      // Colour cannot exceed alpha in premultiplied form; clamp overshoot from sharpening kernels.
      const r = p[s] > a ? a : p[s]
      const g = p[s + 1] > a ? a : p[s + 1]
      const b = p[s + 2] > a ? a : p[s + 2]
      out[d] = r * k; out[d + 1] = g * k; out[d + 2] = b * k
    }
  }
}

/** Premultiplied floats (width * height * 4) -> a new straight RGBA8 buffer. */
export function fromPremultiplied(p: Float32Array | Float64Array, width: number, height: number): PixelBuffer {
  const out = createBuffer(width, height)
  unpremultiplyInto(p, 0, out.data, 0, width * height)
  return out
}
