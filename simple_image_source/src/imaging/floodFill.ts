// src/imaging/floodFill.ts (WP4)
// Magic Wand and Paint Bucket region growing. Pure and DOM-free.
//
// A pixel matches the seed when max(|dr|, |dg|, |db|[, |da|]) <= tolerance. With `compareAlpha` the
// colour channels are compared premultiplied, so fully transparent pixels match each other whatever
// RGB they hide (straight-alpha layers keep arbitrary colour under alpha 0). Contiguous mode grows
// 4-connected spans from the seed (a one-pixel diagonal line is a barrier, as in Photoshop); global mode
// takes every matching pixel. Anti-aliasing softens the pixels just outside the region (4-neighbours of
// matched pixels) to clamp(1 - (diff - tolerance) / 8).
import type { FloodOptions, MaskBuffer, OpOptions, PixelBuffer, Point } from './types.ts'

/** Width of the anti-aliasing ramp beyond the tolerance, in colour levels. */
const SOFT_RANGE = 8

function abortError(): Error {
  const error = new Error('The operation was cancelled.')
  error.name = 'AbortError'
  return error
}

/** Colour distance of every pixel to the seed colour is computed on demand by this closure. */
function differenceFunction(data: Uint8ClampedArray, seedIndex: number, compareAlpha: boolean): (pixel: number) => number {
  const p = seedIndex * 4
  if (!compareAlpha) {
    const r = data[p]
    const g = data[p + 1]
    const b = data[p + 2]
    return (pixel) => {
      const q = pixel * 4
      let d = Math.abs(data[q] - r)
      const dg = Math.abs(data[q + 1] - g)
      if (dg > d) d = dg
      const db = Math.abs(data[q + 2] - b)
      return db > d ? db : d
    }
  }
  const a = data[p + 3]
  const r = (data[p] * a) / 255
  const g = (data[p + 1] * a) / 255
  const b = (data[p + 2] * a) / 255
  return (pixel) => {
    const q = pixel * 4
    const qa = data[q + 3]
    let d = Math.abs(qa - a)
    const dr = Math.abs((data[q] * qa) / 255 - r)
    if (dr > d) d = dr
    const dg = Math.abs((data[q + 1] * qa) / 255 - g)
    if (dg > d) d = dg
    const db = Math.abs((data[q + 2] * qa) / 255 - b)
    return db > d ? db : d
  }
}

/**
 * Selection coverage of the region grown from `seed` (document pixel coordinates, floored). Matched pixels
 * are 255; with anti-aliasing the ring just outside gets partial coverage. A seed outside the image
 * yields an empty mask.
 */
export function floodFill(src: PixelBuffer, seed: Point, options: FloodOptions, op?: OpOptions): MaskBuffer {
  if (!src || !src.data || src.data.length !== src.width * src.height * 4) {
    throw new RangeError('The pixel buffer does not match its size.')
  }
  const { width, height, data } = src
  const out = new Uint8Array(width * height)
  const result: MaskBuffer = { width, height, data: out }
  const sx = Math.floor(Number(seed?.x))
  const sy = Math.floor(Number(seed?.y))
  if (!(sx >= 0 && sy >= 0 && sx < width && sy < height)) return result
  const tolerance = Math.max(0, Math.min(255, Number.isFinite(options?.tolerance) ? options.tolerance : 32))
  const diff = differenceFunction(data, sy * width + sx, Boolean(options?.compareAlpha))
  const signal = op?.signal
  const onProgress = op?.onProgress

  let minY = sy
  let maxY = sy
  if (options?.contiguous === false) {
    minY = height
    maxY = -1
    for (let y = 0; y < height; y += 1) {
      if ((y & 255) === 0 && signal?.aborted) throw abortError()
      const row = y * width
      let any = false
      for (let x = 0; x < width; x += 1) {
        if (diff(row + x) <= tolerance) {
          out[row + x] = 255
          any = true
        }
      }
      if (any) {
        if (y < minY) minY = y
        maxY = y
      }
    }
  } else {
    // Scanline fill: each stack entry is a pixel to grow a horizontal span from.
    const stack: number[] = [sy * width + sx]
    let spans = 0
    while (stack.length) {
      const start = stack.pop() as number
      if (out[start] !== 0) continue
      if (diff(start) > tolerance) continue
      if ((++spans & 4095) === 0) {
        if (signal?.aborted) throw abortError()
      }
      const y = (start / width) | 0
      const row = y * width
      let left = start - row
      let right = left
      while (left > 0 && out[row + left - 1] === 0 && diff(row + left - 1) <= tolerance) left -= 1
      while (right + 1 < width && out[row + right + 1] === 0 && diff(row + right + 1) <= tolerance) right += 1
      out.fill(255, row + left, row + right + 1)
      if (y < minY) minY = y
      if (y > maxY) maxY = y
      for (const ny of [y - 1, y + 1]) {
        if (ny < 0 || ny >= height) continue
        const nrow = ny * width
        let inRun = false
        for (let x = left; x <= right; x += 1) {
          const index = nrow + x
          const candidate = out[index] === 0 && diff(index) <= tolerance
          if (candidate && !inRun) {
            stack.push(index)
            inRun = true
          } else if (!candidate) {
            inRun = false
          }
        }
      }
    }
  }
  onProgress?.(0.8)

  if (options?.antiAlias && maxY >= minY) {
    // Partial coverage for unmatched 4-neighbours of matched pixels. Their difference is always above
    // the tolerance (contiguous: otherwise they would have been filled; global: by definition), so the
    // value stays below 255 and cannot be mistaken for a matched pixel.
    const y0 = Math.max(0, minY - 1)
    const y1 = Math.min(height - 1, maxY + 1)
    for (let y = y0; y <= y1; y += 1) {
      const row = y * width
      for (let x = 0; x < width; x += 1) {
        const index = row + x
        if (out[index] !== 0) continue
        const touches = (x > 0 && out[index - 1] === 255)
          || (x + 1 < width && out[index + 1] === 255)
          || (y > 0 && out[index - width] === 255)
          || (y + 1 < height && out[index + width] === 255)
        if (!touches) continue
        const coverage = 1 - (diff(index) - tolerance) / SOFT_RANGE
        if (coverage > 0) out[index] = Math.min(254, Math.round(coverage * 255))
      }
    }
  }
  onProgress?.(1)
  return result
}
