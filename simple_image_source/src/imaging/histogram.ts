// src/imaging/histogram.ts (WP2)
// 256-bin histograms for the Levels dialog, the Curves backdrop and the auto adjustments. Pure, DOM-free.
import type { Histogram, MaskBuffer, PixelBuffer } from './types.ts'
import { assertBuffer, assertMask } from './buffer.ts'

/**
 * Counts every pixel the mask selects (mask > 0; no mask = all). `alpha` bins count all of those pixels by
 * alpha (bin 0 = fully transparent); red/green/blue/luma and `count` only include pixels with alpha > 0.
 * Luma is Rec. 709 of the gamma-encoded values, rounded to the nearest bin.
 */
export function computeHistogram(src: PixelBuffer, mask: MaskBuffer | null = null): Histogram {
  assertBuffer(src)
  if (mask) assertMask(mask, src.width, src.height)
  const red = new Uint32Array(256)
  const green = new Uint32Array(256)
  const blue = new Uint32Array(256)
  const luma = new Uint32Array(256)
  const alpha = new Uint32Array(256)
  const data = src.data
  const total = src.width * src.height
  let count = 0
  for (let pixel = 0; pixel < total; pixel += 1) {
    if (mask && mask.data[pixel] === 0) continue
    const i = pixel * 4
    const a = data[i + 3]
    alpha[a] += 1
    if (a === 0) continue
    const r = data[i]
    const g = data[i + 1]
    const b = data[i + 2]
    red[r] += 1
    green[g] += 1
    blue[b] += 1
    luma[((2126 * r + 7152 * g + 722 * b + 5000) / 10000) | 0] += 1
    count += 1
  }
  return { red, green, blue, luma, alpha, count }
}

/** Sum of all bins. */
export function histogramTotal(bins: ArrayLike<number>): number {
  let total = 0
  for (let index = 0; index < bins.length; index += 1) total += bins[index]
  return total
}

/**
 * The smallest bin value v whose cumulative count reaches `fraction` of the total (0..1). Returns -1 for an
 * empty histogram. percentile(bins, 0.5) is the median.
 */
export function histogramPercentile(bins: ArrayLike<number>, fraction: number): number {
  const total = histogramTotal(bins)
  if (total <= 0) return -1
  const target = Math.max(0, Math.min(1, fraction)) * total
  let cumulative = 0
  for (let index = 0; index < bins.length; index += 1) {
    cumulative += bins[index]
    if (cumulative >= target && cumulative > 0) return index
  }
  return bins.length - 1
}

/** Mean bin value (weighted by counts), or -1 when empty. */
export function histogramMean(bins: ArrayLike<number>): number {
  let total = 0
  let sum = 0
  for (let index = 0; index < bins.length; index += 1) {
    total += bins[index]
    sum += bins[index] * index
  }
  return total > 0 ? sum / total : -1
}
