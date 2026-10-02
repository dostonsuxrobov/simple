// src/imaging/transform.ts (WP2)
// Arbitrary rotation, affine and perspective warps (inverse mapping with nearest / bilinear / bicubic
// (Catmull-Rom) / Lanczos-3 sampling of premultiplied colour, so edges blend into transparency without dark
// fringes) and homography helpers for free transform. Coordinates: pixel (i, j) covers [i, i + 1) x [j, j + 1),
// so its centre is (i + 0.5, j + 0.5); y points down; positive angles turn clockwise on screen. Pure, DOM-free.
import type { Affine, Homography, IntRect, Interpolation, OpOptions, PixelBuffer, Point, RotateFit, Rgba8, Size } from './types.ts'
import { assertBuffer, chunkRowsFor, cloneBuffer, createBuffer, finishedRun, rotate180, rotate90, runRowsSync, throwIfAborted } from './buffer.ts'
import type { RowRun } from './buffer.ts'

const TRANSPARENT: Rgba8 = Object.freeze({ r: 0, g: 0, b: 0, a: 0 })
const MAX_SIDE = 1 << 17
const MAX_PIXELS = 1 << 28

// ---------------------------------------------------------------------------------------------
// Sampling
// ---------------------------------------------------------------------------------------------

/** Writes the sample at index-space position (x, y) (pixel centres on integers) into out[o..o+3]. */
type Sampler = (x: number, y: number, out: Uint8ClampedArray, o: number) => void

/**
 * clampEdges: kernel taps past the image repeat the edge pixel instead of reading the background (used when
 * the output lies entirely inside the source, so no background may leak into it).
 */
function createSampler(src: PixelBuffer, interpolation: Interpolation, background: Rgba8, clampEdges = false): Sampler {
  const { width, height, data } = src
  const bgA = clampByte(background.a)
  // Background in premultiplied form, 0..255*255 units like the accumulators below.
  const bgR = clampByte(background.r) * bgA
  const bgG = clampByte(background.g) * bgA
  const bgB = clampByte(background.b) * bgA
  const writeBackground = (out: Uint8ClampedArray, o: number) => {
    out[o] = background.r
    out[o + 1] = background.g
    out[o + 2] = background.b
    out[o + 3] = bgA
  }
  const write = (out: Uint8ClampedArray, o: number, r: number, g: number, b: number, a: number) => {
    // r, g, b are premultiplied (colour * alpha); a in 0..255 (may overshoot with negative lobes).
    if (!(a > 0.5)) {
      out[o] = 0; out[o + 1] = 0; out[o + 2] = 0; out[o + 3] = 0
      return
    }
    // Colour divides by the unclamped alpha, so an alpha overshoot never brightens the colour.
    const limit = a * 255
    const k = 1 / a
    out[o] = (r > limit ? limit : r) * k
    out[o + 1] = (g > limit ? limit : g) * k
    out[o + 2] = (b > limit ? limit : b) * k
    out[o + 3] = a > 255 ? 255 : a
  }
  const copyPixel = (x: number, y: number, out: Uint8ClampedArray, o: number) => {
    if (x < 0 || y < 0 || x >= width || y >= height) {
      writeBackground(out, o)
      return
    }
    const s = (y * width + x) * 4
    out[o] = data[s]; out[o + 1] = data[s + 1]; out[o + 2] = data[s + 2]; out[o + 3] = data[s + 3]
  }

  if (interpolation === 'nearest') return (x, y, out, o) => copyPixel(Math.round(x), Math.round(y), out, o)

  const support = interpolation === 'bilinear' ? 1 : interpolation === 'lanczos3' ? 3 : 2
  const taps = support * 2
  const fillWeights = interpolation === 'bilinear' ? bilinearWeights : interpolation === 'lanczos3' ? lanczos3Weights : bicubicWeights
  const wx = new Float64Array(taps)
  const wy = new Float64Array(taps)
  const rowStride = width * 4
  const generic: Sampler = (x, y, out, o) => {
    if (!(x > -support && y > -support && x < width - 1 + support && y < height - 1 + support)) {
      writeBackground(out, o)
      return
    }
    const x0 = Math.floor(x)
    const y0 = Math.floor(y)
    const fx = x - x0
    const fy = y - y0
    // Every kernel interpolates (weight 1 at 0, 0 at other integers): an exact position copies the pixel,
    // hidden colour of transparent pixels included, so identity and integer moves are byte-exact.
    if (fx === 0 && fy === 0) {
      copyPixel(x0, y0, out, o)
      return
    }
    const firstX = x0 - support + 1
    const firstY = y0 - support + 1
    fillWeights(fx, wx)
    fillWeights(fy, wy)
    let r = 0
    let g = 0
    let b = 0
    let a = 0
    if (firstX >= 0 && firstY >= 0 && firstX + taps <= width && firstY + taps <= height) {
      // Every tap inside: no bounds checks.
      let rowStart = (firstY * width + firstX) * 4
      for (let j = 0; j < taps; j += 1) {
        const w2 = wy[j]
        let s = rowStart
        for (let i = 0; i < taps; i += 1) {
          const k = data[s + 3] * wx[i] * w2
          r += k * data[s]
          g += k * data[s + 1]
          b += k * data[s + 2]
          a += k
          s += 4
        }
        rowStart += rowStride
      }
      write(out, o, r, g, b, a)
      return
    }
    for (let j = 0; j < taps; j += 1) {
      const w2 = wy[j]
      if (w2 === 0) continue
      const sy = firstY + j
      for (let i = 0; i < taps; i += 1) {
        const w = wx[i] * w2
        if (w === 0) continue
        const sx = firstX + i
        if (sx >= 0 && sy >= 0 && sx < width && sy < height) {
          const s = (sy * width + sx) * 4
          const wa = w * data[s + 3]
          r += wa * data[s]
          g += wa * data[s + 1]
          b += wa * data[s + 2]
          a += wa
        } else if (clampEdges) {
          const s = ((sy < 0 ? 0 : sy >= height ? height - 1 : sy) * width + (sx < 0 ? 0 : sx >= width ? width - 1 : sx)) * 4
          const wa = w * data[s + 3]
          r += wa * data[s]
          g += wa * data[s + 1]
          b += wa * data[s + 2]
          a += wa
        } else if (bgA !== 0) {
          r += w * bgR
          g += w * bgG
          b += w * bgB
          a += w * bgA
        }
      }
    }
    write(out, o, r, g, b, a)
  }
  if (interpolation !== 'bicubic') return generic

  // Catmull-Rom fast path: inline weights and an unrolled 4 x 4 interior (the default rotate/warp kernel).
  return (x, y, out, o) => {
    const x0 = Math.floor(x)
    const y0 = Math.floor(y)
    if (!(x0 >= 1 && y0 >= 1 && x0 + 2 < width && y0 + 2 < height)) {
      generic(x, y, out, o)
      return
    }
    const fx = x - x0
    const fy = y - y0
    if (fx === 0 && fy === 0) {
      copyPixel(x0, y0, out, o)
      return
    }
    const fx2 = fx * fx
    const fx3 = fx2 * fx
    const fy2 = fy * fy
    const fy3 = fy2 * fy
    const wx0 = -0.5 * fx3 + fx2 - 0.5 * fx
    const wx1 = 1.5 * fx3 - 2.5 * fx2 + 1
    const wx2 = -1.5 * fx3 + 2 * fx2 + 0.5 * fx
    const wx3 = 0.5 * fx3 - 0.5 * fx2
    let r = 0
    let g = 0
    let b = 0
    let a = 0
    let s = ((y0 - 1) * width + (x0 - 1)) * 4
    for (let j = 0; j < 4; j += 1) {
      const w = j === 0 ? -0.5 * fy3 + fy2 - 0.5 * fy : j === 1 ? 1.5 * fy3 - 2.5 * fy2 + 1 : j === 2 ? -1.5 * fy3 + 2 * fy2 + 0.5 * fy : 0.5 * fy3 - 0.5 * fy2
      let k = data[s + 3] * wx0 * w
      r += k * data[s]
      g += k * data[s + 1]
      b += k * data[s + 2]
      a += k
      k = data[s + 7] * wx1 * w
      r += k * data[s + 4]
      g += k * data[s + 5]
      b += k * data[s + 6]
      a += k
      k = data[s + 11] * wx2 * w
      r += k * data[s + 8]
      g += k * data[s + 9]
      b += k * data[s + 10]
      a += k
      k = data[s + 15] * wx3 * w
      r += k * data[s + 12]
      g += k * data[s + 13]
      b += k * data[s + 14]
      a += k
      s += rowStride
    }
    write(out, o, r, g, b, a)
  }
}

// Tap weights for a sample at fraction f (0 <= f < 1) past tap index `support - 1`; one call per axis.

/** Taps x0, x0 + 1. */
function bilinearWeights(f: number, out: Float64Array): void {
  out[0] = 1 - f
  out[1] = f
}

/** Catmull-Rom taps x0 - 1 .. x0 + 2. */
function bicubicWeights(f: number, out: Float64Array): void {
  const f2 = f * f
  const f3 = f2 * f
  out[0] = -0.5 * f3 + f2 - 0.5 * f
  out[1] = 1.5 * f3 - 2.5 * f2 + 1
  out[2] = -1.5 * f3 + 2 * f2 + 0.5 * f
  out[3] = 0.5 * f3 - 0.5 * f2
}

const SIN_THIRDS = [-3, -2, -1, 0, 1, 2].map((m) => Math.sin((Math.PI * m) / 3))
const COS_THIRDS = [-3, -2, -1, 0, 1, 2].map((m) => Math.cos((Math.PI * m) / 3))

const LANCZOS_STEPS = 16384
let lanczosTable: Float64Array | null = null

/**
 * Lanczos-3 taps from a table at 1/16384 px resolution (the trig per sample otherwise dominates a Lanczos
 * warp). Each row is normalised to sum 1: raw Lanczos weights sum to 1 +- 1%, which would leave opaque
 * pixels slightly transparent (alpha 252..254) after a warp.
 */
function lanczos3Weights(f: number, out: Float64Array): void {
  if (!lanczosTable) {
    const table = new Float64Array((LANCZOS_STEPS + 1) * 6)
    const row = new Float64Array(6)
    for (let step = 0; step <= LANCZOS_STEPS; step += 1) {
      lanczos3WeightsExact(step / LANCZOS_STEPS, row)
      const sum = row[0] + row[1] + row[2] + row[3] + row[4] + row[5]
      for (let k = 0; k < 6; k += 1) table[step * 6 + k] = row[k] / sum
    }
    lanczosTable = table
  }
  const base = Math.round(f * LANCZOS_STEPS) * 6
  out[0] = lanczosTable[base]
  out[1] = lanczosTable[base + 1]
  out[2] = lanczosTable[base + 2]
  out[3] = lanczosTable[base + 3]
  out[4] = lanczosTable[base + 4]
  out[5] = lanczosTable[base + 5]
}

/**
 * Lanczos-3 taps x0 - 2 .. x0 + 3 at distances d = f + 2 - k. sin(pi d) = (-1)^k sin(pi f) and
 * sin(pi d / 3) follows from the angle-addition formula, so three trig calls cover all six taps.
 */
function lanczos3WeightsExact(f: number, out: Float64Array): void {
  const sinF = Math.sin(Math.PI * f)
  const sinThird = Math.sin((Math.PI * f) / 3)
  const cosThird = Math.cos((Math.PI * f) / 3)
  for (let k = 0; k < 6; k += 1) {
    const d = f + 2 - k
    if (d === 0) {
      out[k] = 1
      continue
    }
    const m = 2 - k + 3
    const sinD = k % 2 === 0 ? sinF : -sinF
    const sinD3 = sinThird * COS_THIRDS[m] + cosThird * SIN_THIRDS[m]
    out[k] = d <= -3 || d >= 3 ? 0 : (3 * sinD * sinD3) / (Math.PI * Math.PI * d * d)
  }
}

function clampByte(value: number): number {
  const n = Number(value)
  return n > 0 ? (n < 255 ? n : 255) : 0
}

function checkOutput(width: number, height: number): void {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
    throw new RangeError(`Invalid output size ${width} x ${height}.`)
  }
  if (width > MAX_SIDE || height > MAX_SIDE || width * height > MAX_PIXELS) {
    throw new RangeError(`The output size ${width} x ${height} is too large.`)
  }
}

function readInterpolation(value: Interpolation | undefined): Interpolation {
  return value === 'nearest' || value === 'bilinear' || value === 'lanczos3' ? value : 'bicubic'
}

// ---------------------------------------------------------------------------------------------
// Rotation
// ---------------------------------------------------------------------------------------------

/** Size of the bounding box of a width x height rectangle turned by `degrees` (the 'expand' canvas). */
export function rotatedBounds(width: number, height: number, degrees: number): Size {
  const t = (degrees * Math.PI) / 180
  const c = Math.abs(Math.cos(t))
  const s = Math.abs(Math.sin(t))
  return {
    width: Math.max(width > 0 ? 1 : 0, Math.ceil(width * c + height * s - 1e-6)),
    height: Math.max(height > 0 ? 1 : 0, Math.ceil(width * s + height * c - 1e-6)),
  }
}

/**
 * Largest centred rectangle with the original aspect ratio inside the image turned by `degrees`:
 * scale = min(W / (W|cos| + H|sin|), H / (W|sin| + H|cos|)); the result is floor(W * scale) x floor(H * scale).
 */
export function inscribedSize(width: number, height: number, degrees: number): Size {
  if (!(width > 0 && height > 0)) return { width: 0, height: 0 }
  const t = (degrees * Math.PI) / 180
  const c = Math.abs(Math.cos(t))
  const s = Math.abs(Math.sin(t))
  const scale = Math.min(width / (width * c + height * s), height / (width * s + height * c))
  return { width: Math.max(1, Math.floor(width * scale + 1e-6)), height: Math.max(1, Math.floor(height * scale + 1e-6)) }
}

function normalizeDegrees(degrees: number): number {
  let d = degrees % 360
  if (d < 0) d += 360
  // Snap values a hair away from a quarter turn (floating error from UI arithmetic).
  const quarter = Math.round(d / 90) * 90
  if (Math.abs(d - quarter) < 1e-9) d = quarter % 360
  return d
}

/** The resumable form of rotateArbitrary (worker handlers run it in chunks of output rows). */
export function startRotate(
  src: PixelBuffer,
  degrees: number,
  fit: RotateFit,
  interpolation: Interpolation = 'bicubic',
  background: Rgba8 = TRANSPARENT,
): RowRun {
  assertBuffer(src)
  if (!Number.isFinite(degrees)) throw new RangeError('The rotation angle must be a finite number.')
  const d = normalizeDegrees(degrees)
  const { width, height } = src
  const square = width === height
  if (d === 0) return finishedRun(cloneBuffer(src))
  if (d === 180) return finishedRun(rotate180(src))
  if ((d === 90 || d === 270) && (fit === 'expand' || square)) return finishedRun(rotate90(src, d === 90))
  const size = fit === 'expand' ? rotatedBounds(width, height, d) : fit === 'same-size' ? { width, height } : inscribedSize(width, height, d)
  checkOutput(size.width, size.height)
  const t = (d * Math.PI) / 180
  const cos = Math.cos(t)
  const sin = Math.sin(t)
  // Output point P maps back to source point c + R(-theta) (P - c').
  const inCx = width / 2
  const inCy = height / 2
  const outCx = size.width / 2
  const outCy = size.height / 2
  const inverse: Affine = [cos, -sin, sin, cos, inCx - cos * outCx - sin * outCy, inCy + sin * outCx - cos * outCy]
  return affineRun(src, inverse, { x: 0, y: 0, width: size.width, height: size.height }, readInterpolation(interpolation), background ?? TRANSPARENT, fit === 'crop-inscribed')
}

/**
 * Rotates clockwise by `degrees` about the centre. fit 'expand' grows the canvas to the rotated bounds,
 * 'same-size' keeps width x height, 'crop-inscribed' keeps the largest centred rectangle of the original
 * aspect ratio with no empty corners. Areas outside the source become `background` (default transparent).
 * Multiples of 90 degrees are exact pixel permutations whenever the output geometry allows it.
 */
export function rotateArbitrary(
  src: PixelBuffer,
  degrees: number,
  fit: RotateFit,
  interpolation: Interpolation = 'bicubic',
  background: Rgba8 = TRANSPARENT,
  options?: OpOptions,
): PixelBuffer {
  throwIfAborted(options?.signal)
  return runRowsSync(startRotate(src, degrees, fit, interpolation, background), options)
}

// ---------------------------------------------------------------------------------------------
// Warps
// ---------------------------------------------------------------------------------------------

const WARP_CHUNK_PIXELS = 1 << 17

function affineRun(src: PixelBuffer, inverse: Affine, out: IntRect, interpolation: Interpolation, background: Rgba8, clampEdges = false): RowRun {
  const [a, b, c, d, e, f] = inverse
  const output = createBuffer(out.width, out.height)
  const sample = createSampler(src, interpolation, background, clampEdges)
  const o = output.data
  return {
    output,
    rows: out.height,
    chunkRows: chunkRowsFor(out.width, WARP_CHUNK_PIXELS),
    process(startRow, endRow) {
      for (let j = startRow; j < endRow; j += 1) {
        const Y = out.y + j + 0.5
        const X0 = out.x + 0.5
        // Source position (continuous) minus 0.5 = index space.
        let sx = a * X0 + c * Y + e - 0.5
        let sy = b * X0 + d * Y + f - 0.5
        let index = j * out.width * 4
        for (let i = 0; i < out.width; i += 1) {
          sample(sx, sy, o, index)
          sx += a
          sy += b
          index += 4
        }
      }
    },
  }
}

/** The resumable form of warpAffine. */
export function startWarpAffine(src: PixelBuffer, inverse: Affine, out: IntRect, interpolation: Interpolation = 'bicubic'): RowRun {
  assertBuffer(src)
  checkOutput(out.width, out.height)
  if (!inverse || inverse.length !== 6 || !inverse.every(Number.isFinite)) throw new RangeError('The transform is not a valid affine matrix.')
  return affineRun(src, inverse, out, readInterpolation(interpolation), TRANSPARENT)
}

/**
 * Affine warp: `inverse` maps destination coordinates to source coordinates (canvas order
 * x' = a x + c y + e, y' = b x + d y + f). The result covers `out` in destination space; outside the
 * source it is transparent.
 */
export function warpAffine(src: PixelBuffer, inverse: Affine, out: IntRect, interpolation: Interpolation = 'bicubic', options?: OpOptions): PixelBuffer {
  throwIfAborted(options?.signal)
  return runRowsSync(startWarpAffine(src, inverse, out, interpolation), options)
}

/** The resumable form of warpPerspective. */
export function startWarpPerspective(src: PixelBuffer, inverse: Homography, out: IntRect, interpolation: Interpolation = 'bicubic'): RowRun {
  assertBuffer(src)
  checkOutput(out.width, out.height)
  if (!inverse || inverse.length !== 9 || !inverse.every(Number.isFinite)) throw new RangeError('The transform is not a valid 3 x 3 matrix.')
  const [h0, h1, h2, h3, h4, h5, h6, h7, h8] = inverse
  if (Math.abs(h6) < 1e-15 && Math.abs(h7) < 1e-15 && Math.abs(h8 - 1) < 1e-15) {
    return affineRun(src, [h0, h3, h1, h4, h2, h5], out, readInterpolation(interpolation), TRANSPARENT)
  }
  const output = createBuffer(out.width, out.height)
  const sample = createSampler(src, readInterpolation(interpolation), TRANSPARENT)
  const o = output.data
  return {
    output,
    rows: out.height,
    chunkRows: chunkRowsFor(out.width, WARP_CHUNK_PIXELS),
    process(startRow, endRow) {
      for (let j = startRow; j < endRow; j += 1) {
        const Y = out.y + j + 0.5
        for (let i = 0; i < out.width; i += 1) {
          const X = out.x + i + 0.5
          const w = h6 * X + h7 * Y + h8
          if (!(w > 1e-12)) continue
          const sx = (h0 * X + h1 * Y + h2) / w - 0.5
          const sy = (h3 * X + h4 * Y + h5) / w - 0.5
          sample(sx, sy, o, (j * out.width + i) * 4)
        }
      }
    },
  }
}

/**
 * Perspective warp: `inverse` maps destination coordinates to source coordinates ([x, y, 1] -> [X, Y, W],
 * point (X / W, Y / W)). Points with W <= 0 lie behind the projection and stay transparent.
 */
export function warpPerspective(src: PixelBuffer, inverse: Homography, out: IntRect, interpolation: Interpolation = 'bicubic', options?: OpOptions): PixelBuffer {
  throwIfAborted(options?.signal)
  return runRowsSync(startWarpPerspective(src, inverse, out, interpolation), options)
}

// ---------------------------------------------------------------------------------------------
// Homographies
// ---------------------------------------------------------------------------------------------

export function applyHomography(h: Homography, p: Point): Point {
  const w = h[6] * p.x + h[7] * p.y + h[8]
  return { x: (h[0] * p.x + h[1] * p.y + h[2]) / w, y: (h[3] * p.x + h[4] * p.y + h[5]) / w }
}

function multiply3(a: readonly number[], b: readonly number[]): number[] {
  const out = new Array<number>(9)
  for (let row = 0; row < 3; row += 1) {
    for (let column = 0; column < 3; column += 1) {
      out[row * 3 + column] = a[row * 3] * b[column] + a[row * 3 + 1] * b[3 + column] + a[row * 3 + 2] * b[6 + column]
    }
  }
  return out
}

function normalizeHomography(values: readonly number[]): Homography {
  let scale = values[8]
  if (Math.abs(scale) < 1e-12) {
    scale = 0
    for (const value of values) if (Math.abs(value) > Math.abs(scale)) scale = value
  }
  if (!scale || !Number.isFinite(scale)) throw new RangeError('The transform cannot be represented.')
  return values.map((value) => value / scale) as unknown as Homography
}

/** Inverse of a homography (normalised so the last entry is 1 when possible). Throws for a singular matrix. */
export function invertHomography(h: Homography): Homography {
  const [a, b, c, d, e, f, g, k, i] = h
  const A = e * i - f * k
  const B = -(d * i - f * g)
  const C = d * k - e * g
  const det = a * A + b * B + c * C
  const scale = Math.max(...h.map((value) => Math.abs(value)))
  if (!Number.isFinite(det) || Math.abs(det) <= 1e-14 * Math.pow(scale || 1, 3)) throw new RangeError('The transform is singular and cannot be inverted.')
  const adjugate = [
    A, -(b * i - c * k), b * f - c * e,
    B, a * i - c * g, -(a * f - c * d),
    C, -(a * k - b * g), a * e - b * d,
  ]
  return normalizeHomography(adjugate.map((value) => value / det))
}

/** Hartley normalisation: centroid to the origin, mean distance sqrt(2). Returns the 3x3 matrix. */
function normalizer(points: readonly Point[]): number[] {
  let mx = 0
  let my = 0
  for (const p of points) {
    mx += p.x
    my += p.y
  }
  mx /= points.length
  my /= points.length
  let distance = 0
  for (const p of points) distance += Math.hypot(p.x - mx, p.y - my)
  distance /= points.length
  if (!(distance > 1e-12)) throw new RangeError('The corner points are degenerate (they coincide).')
  const s = Math.SQRT2 / distance
  return [s, 0, -s * mx, 0, s, -s * my, 0, 0, 1]
}

/**
 * The homography that maps from[k] to to[k] for the four corner pairs (direct linear transform with
 * Hartley normalisation and partial pivoting). Throws a RangeError for degenerate quads.
 */
export function homographyFromQuads(from: readonly Point[], to: readonly Point[]): Homography {
  if (!from || !to || from.length !== 4 || to.length !== 4) throw new RangeError('A perspective transform needs exactly four corner pairs.')
  for (const p of [...from, ...to]) {
    if (!p || !Number.isFinite(p.x) || !Number.isFinite(p.y)) throw new RangeError('The corner points must be finite numbers.')
  }
  const tFrom = normalizer(from)
  const tTo = normalizer(to)
  const nf = from.map((p) => ({ x: tFrom[0] * p.x + tFrom[2], y: tFrom[4] * p.y + tFrom[5] }))
  const nt = to.map((p) => ({ x: tTo[0] * p.x + tTo[2], y: tTo[4] * p.y + tTo[5] }))
  // 8 x 9 augmented system for h0..h7 with h8 = 1.
  const m: number[][] = []
  for (let k = 0; k < 4; k += 1) {
    const { x, y } = nf[k]
    const { x: u, y: v } = nt[k]
    m.push([x, y, 1, 0, 0, 0, -x * u, -y * u, u])
    m.push([0, 0, 0, x, y, 1, -x * v, -y * v, v])
  }
  for (let column = 0; column < 8; column += 1) {
    let pivot = column
    for (let row = column + 1; row < 8; row += 1) if (Math.abs(m[row][column]) > Math.abs(m[pivot][column])) pivot = row
    if (Math.abs(m[pivot][column]) < 1e-10) throw new RangeError('The corner points are degenerate (three of them lie on one line).')
    if (pivot !== column) {
      const swap = m[pivot]
      m[pivot] = m[column]
      m[column] = swap
    }
    for (let row = 0; row < 8; row += 1) {
      if (row === column) continue
      const factor = m[row][column] / m[column][column]
      if (factor === 0) continue
      for (let k = column; k < 9; k += 1) m[row][k] -= factor * m[column][k]
    }
  }
  const hn = [...Array.from({ length: 8 }, (_, row) => m[row][8] / m[row][row]), 1]
  // H = T_to^-1 * Hn * T_from
  const sTo = tTo[0]
  const tToInverse = [1 / sTo, 0, -tTo[2] / sTo, 0, 1 / sTo, -tTo[5] / sTo, 0, 0, 1]
  return normalizeHomography(multiply3(multiply3(tToInverse, hn), tFrom))
}

/** The homography form of an affine matrix (canvas order). */
export function affineToHomography(m: Affine): Homography {
  return [m[0], m[2], m[4], m[1], m[3], m[5], 0, 0, 1]
}

/** Inverse of an affine matrix (canvas order). Throws for a singular matrix. */
export function invertAffine(m: Affine): Affine {
  const [a, b, c, d, e, f] = m
  const det = a * d - b * c
  if (!Number.isFinite(det) || Math.abs(det) < 1e-14) throw new RangeError('The transform is singular and cannot be inverted.')
  return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det]
}
