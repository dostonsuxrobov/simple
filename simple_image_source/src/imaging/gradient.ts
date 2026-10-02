// src/imaging/gradient.ts (WP4)
// Gradient tool rendering: linear, radial, angle, reflected and diamond gradients with colour stops
// (midpoint bias u^(ln 0.5 / ln m), as in Photoshop's gradient editor), opacity stops, reverse and
// dither. Pure and DOM-free; tiles can be rendered independently (dither noise is keyed on document
// coordinates), so a document-size gradient never needs one huge buffer.
//
// Geometry, with d = to - from and L = |d| (pixel centres sit at origin + (i + 0.5, j + 0.5)):
//   linear:    t = (p - from) . d / L^2
//   reflected: t = |(p - from) . d| / L^2
//   radial:    t = |p - from| / L
//   diamond:   t = (|a| + |b|) / L, a and b the coordinates of p - from along d and across it
//   angle:     t = counter-clockwise (on screen) sweep from d around `from`, divided by 360 degrees
// t is clamped to [0, 1] and reversed when spec.reverse.
import type { GradientSpec, GradientStop, OpacityStop, PixelBuffer, Point } from './types.ts'

/** Colour lookup resolution (linear interpolation between entries). */
const LUT_SIZE = 4096

export interface GradientSample {
  /** 0..255, unrounded. */
  readonly r: number
  readonly g: number
  readonly b: number
  readonly a: number
}

function clamp01(value: number): number {
  return value <= 0 ? 0 : value >= 1 ? 1 : value
}

function sortedStops(stops: readonly GradientStop[]): GradientStop[] {
  const list = (stops ?? []).filter((stop) => stop && stop.color && Number.isFinite(stop.position))
    .map((stop) => ({ ...stop, position: clamp01(stop.position) }))
  list.sort((a, b) => a.position - b.position)
  return list.length ? list : [{ position: 0, color: { r: 0, g: 0, b: 0 } }]
}

function sortedOpacity(stops: readonly OpacityStop[]): OpacityStop[] {
  const list = (stops ?? []).filter((stop) => stop && Number.isFinite(stop.position) && Number.isFinite(stop.opacity))
    .map((stop) => ({ position: clamp01(stop.position), opacity: clamp01(stop.opacity) }))
  list.sort((a, b) => a.position - b.position)
  return list
}

/** Index i of the segment [stops[i], stops[i + 1]] containing t, or -1 / length - 1 outside. */
function segmentOf(positions: readonly number[], t: number): number {
  if (t <= positions[0]) return -1
  for (let i = 0; i < positions.length - 1; i += 1) if (t <= positions[i + 1]) return i
  return positions.length - 1
}

function biased(u: number, midpoint: number | undefined): number {
  const m = Math.min(0.99, Math.max(0.01, Number.isFinite(midpoint) ? (midpoint as number) : 0.5))
  if (m === 0.5 || u <= 0 || u >= 1) return u
  return Math.pow(u, Math.log(0.5) / Math.log(m))
}

function colorAt(stops: readonly GradientStop[], t: number): [number, number, number] {
  const positions = stops.map((stop) => stop.position)
  const i = segmentOf(positions, t)
  if (i < 0) return [stops[0].color.r, stops[0].color.g, stops[0].color.b]
  if (i >= stops.length - 1) {
    const last = stops[stops.length - 1].color
    return [last.r, last.g, last.b]
  }
  const a = stops[i]
  const b = stops[i + 1]
  const span = b.position - a.position
  const u = span > 0 ? biased((t - a.position) / span, a.midpoint) : 1
  return [
    a.color.r + (b.color.r - a.color.r) * u,
    a.color.g + (b.color.g - a.color.g) * u,
    a.color.b + (b.color.b - a.color.b) * u,
  ]
}

function opacityAt(stops: readonly OpacityStop[], t: number): number {
  if (!stops.length) return 1
  const positions = stops.map((stop) => stop.position)
  const i = segmentOf(positions, t)
  if (i < 0) return stops[0].opacity
  if (i >= stops.length - 1) return stops[stops.length - 1].opacity
  const a = stops[i]
  const b = stops[i + 1]
  const span = b.position - a.position
  const u = span > 0 ? (t - a.position) / span : 1
  return a.opacity + (b.opacity - a.opacity) * u
}

/** Exact (unrounded) colour of the gradient at parameter t in [0, 1], before `reverse`. */
export function gradientColorAt(spec: Pick<GradientSpec, 'stops' | 'opacityStops'>, t: number): GradientSample {
  const tt = clamp01(Number.isFinite(t) ? t : 0)
  const [r, g, b] = colorAt(sortedStops(spec.stops), tt)
  return { r, g, b, a: opacityAt(sortedOpacity(spec.opacityStops), tt) * 255 }
}

/** Gradient parameter t in [0, 1] at document point (x, y), after `reverse`. */
export function gradientParameter(spec: Pick<GradientSpec, 'kind' | 'from' | 'to' | 'reverse'>, x: number, y: number): number {
  const evaluate = parameterFunction(spec.kind, spec.from, spec.to)
  const t = clamp01(evaluate(x, y))
  return spec.reverse ? 1 - t : t
}

function parameterFunction(kind: GradientSpec['kind'], from: Point, to: Point): (x: number, y: number) => number {
  const fx = from.x
  const fy = from.y
  const dx = to.x - from.x
  const dy = to.y - from.y
  const length2 = dx * dx + dy * dy
  const length = Math.sqrt(length2)
  if (!(length > 1e-9) && kind !== 'angle') {
    // A click without a drag: every pixel lies beyond the end point.
    return () => 1
  }
  switch (kind) {
    case 'radial':
      return (x, y) => Math.sqrt((x - fx) * (x - fx) + (y - fy) * (y - fy)) / length
    case 'reflected':
      return (x, y) => Math.abs(((x - fx) * dx + (y - fy) * dy) / length2)
    case 'diamond':
      return (x, y) => {
        const a = ((x - fx) * dx + (y - fy) * dy) / length
        const b = (-(x - fx) * dy + (y - fy) * dx) / length
        return (Math.abs(a) + Math.abs(b)) / length
      }
    case 'angle': {
      const base = length > 1e-9 ? Math.atan2(dy, dx) : 0
      const turn = 2 * Math.PI
      return (x, y) => {
        // Screen y points down, so a counter-clockwise sweep on screen decreases atan2.
        let t = (base - Math.atan2(y - fy, x - fx)) / turn
        t -= Math.floor(t)
        return t
      }
    }
    case 'linear':
    default:
      return (x, y) => ((x - fx) * dx + (y - fy) * dy) / length2
  }
}

/** Integer hash of a document pixel, channel and seed to [0, 1) (stable across tiles). */
function noise(x: number, y: number, channel: number, seed: number): number {
  let h = Math.imul(x | 0, 0x27d4eb2d) ^ Math.imul(y | 0, 0x165667b1) ^ Math.imul(seed | 0, 0x9e3779b9) ^ Math.imul(channel + 1, 0x85ebca6b)
  h = Math.imul(h ^ (h >>> 15), 0x2c1b3c6d)
  h = Math.imul(h ^ (h >>> 12), 0x297a2d39)
  h ^= h >>> 15
  return (h >>> 0) / 4294967296
}

/**
 * Renders `spec` into a new width x height straight-RGBA buffer whose top-left pixel sits at document
 * position `origin`. Dither adds per-pixel noise in [-0.5, 0.5) before rounding, so every channel stays
 * within 1 level of the undithered value; `seed` varies the noise.
 */
export function renderGradient(width: number, height: number, origin: Point, spec: GradientSpec, seed = 0): PixelBuffer {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 0 || height < 0) {
    throw new RangeError(`Invalid gradient size ${width} x ${height}.`)
  }
  const data = new Uint8ClampedArray(width * height * 4)
  const stops = sortedStops(spec.stops)
  const opacity = sortedOpacity(spec.opacityStops)
  // Lookup table of unrounded RGBA values over t in [0, 1], plus the exact stop values at their positions.
  const lut = new Float32Array((LUT_SIZE + 1) * 4)
  for (let i = 0; i <= LUT_SIZE; i += 1) {
    const t = i / LUT_SIZE
    const [r, g, b] = colorAt(stops, t)
    lut[i * 4] = r
    lut[i * 4 + 1] = g
    lut[i * 4 + 2] = b
    lut[i * 4 + 3] = opacityAt(opacity, t) * 255
  }
  const parameter = parameterFunction(spec.kind, spec.from, spec.to)
  const reverse = Boolean(spec.reverse)
  const dither = Boolean(spec.dither)
  const ox = Number.isFinite(origin?.x) ? origin.x : 0
  const oy = Number.isFinite(origin?.y) ? origin.y : 0
  const baseX = Math.floor(ox)
  const baseY = Math.floor(oy)
  const exact = Float64Array.from(new Set([...stops.map((stop) => stop.position), ...opacity.map((stop) => stop.position)]))
  const isExact = (t: number): boolean => {
    for (let k = 0; k < exact.length; k += 1) if (exact[k] === t) return true
    return false
  }
  const value = [0, 0, 0, 0]
  for (let j = 0; j < height; j += 1) {
    const py = oy + j + 0.5
    for (let i = 0; i < width; i += 1) {
      let t = parameter(ox + i + 0.5, py)
      t = t <= 0 ? 0 : t >= 1 ? 1 : t
      if (reverse) t = 1 - t
      if (isExact(t)) {
        // Exactly on a stop: evaluate without the table so stop colours are hit exactly.
        const [r, g, b] = colorAt(stops, t)
        value[0] = r
        value[1] = g
        value[2] = b
        value[3] = opacityAt(opacity, t) * 255
      } else {
        const f = t * LUT_SIZE
        let k = Math.floor(f)
        if (k >= LUT_SIZE) k = LUT_SIZE - 1
        const w = f - k
        const p = k * 4
        value[0] = lut[p] + (lut[p + 4] - lut[p]) * w
        value[1] = lut[p + 1] + (lut[p + 5] - lut[p + 1]) * w
        value[2] = lut[p + 2] + (lut[p + 6] - lut[p + 2]) * w
        value[3] = lut[p + 3] + (lut[p + 7] - lut[p + 3]) * w
      }
      const q = (j * width + i) * 4
      if (dither) {
        const gx = baseX + i
        const gy = baseY + j
        for (let c = 0; c < 4; c += 1) data[q + c] = Math.round(value[c] + noise(gx, gy, c, seed) - 0.5)
      } else {
        data[q] = Math.round(value[0])
        data[q + 1] = Math.round(value[1])
        data[q + 2] = Math.round(value[2])
        data[q + 3] = Math.round(value[3])
      }
    }
  }
  return { width, height, data }
}
