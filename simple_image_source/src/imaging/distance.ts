// src/imaging/distance.ts (WP4)
// Exact Euclidean distance transforms in O(N) (Felzenszwalb & Huttenlocher, "Distance Transforms of
// Sampled Functions", 2012). Used by Select > Modify > Expand / Contract, spot healing (hole dilation,
// border band, inscribed radius) and Telea inpainting. Pure and DOM-free.
//
// Both transforms are separable: a column pass finds, for every pixel, the row distance to the nearest
// site in its own column (two linear scans); a row pass takes the lower envelope of the per-column
// parabolas. Sites are the pixels whose `binary` value is non-zero.
import type { OpOptions } from './types.ts'

type Metric = 'centre' | 'region'

function abortError(): Error {
  const error = new Error('The operation was cancelled.')
  error.name = 'AbortError'
  return error
}

function checkInput(binary: Uint8Array, width: number, height: number): void {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
    throw new RangeError(`Invalid distance transform size ${width} x ${height}.`)
  }
  if (!binary || binary.length !== width * height) {
    throw new RangeError(`The distance transform input must hold ${width * height} values.`)
  }
}

/**
 * Lower envelope of the parabolas f(q) + (x - (q + shift))^2 sampled at integer x (FH algorithm 1).
 * Non-finite f(q) are skipped; a line without sites yields Infinity.
 */
function lowerEnvelope(f: Float64Array, n: number, shift: number, out: Float64Array, v: Int32Array, z: Float64Array): void {
  let k = -1
  for (let q = 0; q < n; q += 1) {
    const fq = f[q]
    if (fq === Infinity) continue
    const pq = q + shift
    if (k < 0) {
      k = 0
      v[0] = q
      z[0] = -Infinity
      z[1] = Infinity
      continue
    }
    let s = 0
    for (;;) {
      const r = v[k]
      const pr = r + shift
      s = ((fq + pq * pq) - (f[r] + pr * pr)) / (2 * (pq - pr))
      if (s > z[k] || k === 0) break
      k -= 1
    }
    // z[0] is -Infinity, so the loop above always stops with s > z[k] (k may be 0).
    if (s <= z[k]) {
      // Only reachable at k === 0 with s === -Infinity (cannot happen for finite inputs); replace.
      v[0] = q
      z[1] = Infinity
      continue
    }
    k += 1
    v[k] = q
    z[k] = s
    z[k + 1] = Infinity
  }
  if (k < 0) {
    out.fill(Infinity, 0, n)
    return
  }
  let j = 0
  for (let x = 0; x < n; x += 1) {
    while (z[j + 1] < x) j += 1
    const d = x - (v[j] + shift)
    out[x] = d * d + f[v[j]]
  }
}

function transform(binary: Uint8Array, width: number, height: number, metric: Metric, options?: OpOptions): Float32Array {
  checkInput(binary, width, height)
  const total = width * height
  const out = new Float32Array(total)
  if (total === 0) return out
  const signal = options?.signal
  const onProgress = options?.onProgress

  // Column pass: row distance to the nearest site in the same column (Infinity when none).
  const last = new Int32Array(width).fill(-1)
  for (let y = 0; y < height; y += 1) {
    const row = y * width
    for (let x = 0; x < width; x += 1) {
      if (binary[row + x] !== 0) last[x] = y
      out[row + x] = last[x] >= 0 ? y - last[x] : Infinity
    }
  }
  if (signal?.aborted) throw abortError()
  last.fill(-1)
  for (let y = height - 1; y >= 0; y -= 1) {
    const row = y * width
    for (let x = 0; x < width; x += 1) {
      if (binary[row + x] !== 0) last[x] = y
      if (last[x] >= 0) {
        const d = last[x] - y
        if (d < out[row + x]) out[row + x] = d
      }
    }
  }
  onProgress?.(0.3)

  // Row pass: lower envelope across columns.
  const f = new Float64Array(width)
  const result = new Float64Array(width)
  const other = metric === 'region' ? new Float64Array(width) : null
  const v = new Int32Array(width)
  const z = new Float64Array(width + 1)
  const progressEvery = Math.max(64, Math.ceil(height / 8))
  for (let y = 0; y < height; y += 1) {
    if ((y & 63) === 0 && signal?.aborted) throw abortError()
    const row = y * width
    if (metric === 'centre') {
      for (let x = 0; x < width; x += 1) {
        const k = out[row + x]
        f[x] = k * k
      }
      lowerEnvelope(f, width, 0, result, v, z)
    } else {
      // Distance to the closed unit square of a site: h(t) = max(|t| - 0.5, 0)^2 per axis. The column
      // pass already found the nearest site row; h is monotone in |t|, so f = h(k). Across columns,
      // h(x - q) is the parabola centred at q + 0.5 for sites left of x and at q - 0.5 for sites right
      // of x; each envelope over-estimates the other side, so the minimum of both envelopes and the
      // pixel's own column (h(0) = 0) is exact.
      for (let x = 0; x < width; x += 1) {
        const k = out[row + x]
        f[x] = k > 0 ? (k - 0.5) * (k - 0.5) : 0
      }
      lowerEnvelope(f, width, 0.5, result, v, z)
      lowerEnvelope(f, width, -0.5, other as Float64Array, v, z)
      for (let x = 0; x < width; x += 1) {
        let best = f[x]
        const a = result[x]
        const b = (other as Float64Array)[x]
        if (a < best) best = a
        if (b < best) best = b
        result[x] = best
      }
    }
    for (let x = 0; x < width; x += 1) out[row + x] = Math.sqrt(result[x])
    if (onProgress && y > 0 && y % progressEvery === 0) onProgress(0.3 + 0.7 * (y / height))
  }
  onProgress?.(1)
  return out
}

/**
 * Exact Euclidean distance from every pixel centre to the centre of the nearest pixel whose `binary`
 * value is non-zero (0 on those pixels). Infinity everywhere when there is no such pixel.
 */
export function distanceTransform(binary: Uint8Array, width: number, height: number, options?: OpOptions): Float32Array {
  return transform(binary, width, height, 'centre', options)
}

/**
 * Exact Euclidean distance from every pixel centre to the region covered by the non-zero pixels, each
 * taken as the closed unit square around its centre (0 on those pixels; Infinity when there are none).
 * Along a row or column this is the centre distance minus 0.5; near corners it measures to the true
 * corner, so a one-pixel coverage ramp at a fixed distance reproduces the Minkowski sum (or difference)
 * of the pixel region with a disk.
 */
export function regionDistanceTransform(binary: Uint8Array, width: number, height: number, options?: OpOptions): Float32Array {
  return transform(binary, width, height, 'region', options)
}
