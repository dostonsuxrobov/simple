// src/imaging/brush.ts (WP4)
// Brush engine core shared by the Advanced brush, eraser, clone stamp and spot-healing tools. Pure and
// DOM-free.
//
// Dab coverage at distance r from the centre of a dab of radius R and hardness h:
//   hard (h >= 0.999): clamp(R + 0.5 - r, 0, 1)
//   soft:              (1 - smoothstep(clamp((r - R*h) / (R - R*h)))) * clamp(R + 0.5 - r, 0, 1)
// Elliptical tips (roundness < 1, rotated by `angle` degrees counter-clockwise on screen) use the
// normalised elliptical radius for the profile and the first-order distance to the ellipse for the edge.
// Dabs with R < 2 are supersampled (8 x 8) so tiny brushes keep their true area.
//
// Dabs are placed at equal arc length (spacing * diameter) along the input path smoothed by an
// exponential moving average (alpha = 1 - 0.8 * smoothing). A per-stroke coverage buffer accumulates
// S = S + dab * flow * (1 - S); the painted pixel is blend(before, colour, S * opacity * selection),
// computed from the pre-stroke pixels, so opacity caps a stroke exactly as in Photoshop.
import type { BrushTip, Dab, IntRect, MaskBuffer, PixelBuffer, Rgb8, StrokeSample } from './types.ts'

const HARD = 0.999
const SUPERSAMPLE_BELOW_RADIUS = 2
const SUPERSAMPLE = 8
/** Smallest distance between two dabs, px (keeps 1 px brushes at tiny spacing bounded). */
const MIN_STEP = 0.1

/** Optional pen dynamics for placeDabs. */
export interface BrushDynamics {
  /** Pen pressure scales the diameter by lerp(0.1, 1, pressure). Default false. */
  readonly pressureSize?: boolean
  /** Pen pressure scales the flow. Default false. */
  readonly pressureFlow?: boolean
  /** Base flow 0..1. Default 1. */
  readonly flow?: number
}

/** State carried between placeDabs calls of one stroke (start with { distance: 0, last: null }). */
export interface DabCarry {
  /** Arc length travelled since the last dab. */
  distance: number
  /** Last smoothed sample. */
  last: StrokeSample | null
}

function clamp01(value: number): number {
  return value <= 0 ? 0 : value >= 1 ? 1 : value
}

function smoothstep(t: number): number {
  return t * t * (3 - 2 * t)
}

interface TipShape {
  readonly radius: number
  readonly inner: number
  readonly hard: boolean
  readonly round: boolean
  readonly roundness: number
  readonly cos: number
  readonly sin: number
}

function tipShape(tip: BrushTip, diameter: number): TipShape {
  const radius = Math.max(0.5, Number.isFinite(diameter) ? diameter / 2 : 0.5)
  const hardness = clamp01(Number.isFinite(tip.hardness) ? tip.hardness : 1)
  const roundness = Math.max(0.01, Math.min(1, Number.isFinite(tip.roundness) ? tip.roundness : 1))
  const radians = ((Number.isFinite(tip.angle) ? tip.angle : 0) * Math.PI) / 180
  return {
    radius,
    inner: radius * hardness,
    hard: hardness >= HARD,
    round: roundness >= 0.999,
    roundness,
    // Counter-clockwise on screen (y down) is a negative mathematical angle.
    cos: Math.cos(-radians),
    sin: Math.sin(-radians),
  }
}

/** Profile value at a point (dx, dy) from the dab centre; edge = signed distance inside the outline. */
function profileAt(shape: TipShape, dx: number, dy: number): number {
  let r: number
  let edge: number
  if (shape.round) {
    r = Math.sqrt(dx * dx + dy * dy)
    edge = shape.radius - r
  } else {
    const u = dx * shape.cos + dy * shape.sin
    const v = (-dx * shape.sin + dy * shape.cos) / shape.roundness
    r = Math.sqrt(u * u + v * v)
    if (r > 0) {
      // |grad r| for r = sqrt(u^2 + (v'/rho)^2) expressed with v already divided by rho.
      const gu = u / r
      const gv = v / (shape.roundness * r)
      edge = (shape.radius - r) / Math.sqrt(gu * gu + gv * gv)
    } else {
      edge = shape.radius * shape.roundness
    }
  }
  const aa = clamp01(edge + 0.5)
  if (aa === 0 || shape.hard) return aa
  if (r <= shape.inner) return aa
  const span = shape.radius - shape.inner
  const t = clamp01((r - shape.inner) / span)
  return (1 - smoothstep(t)) * aa
}

/** Supersampled profile for tiny dabs: point-sampled membership (and soft profile) on an 8 x 8 grid. */
function supersampledAt(shape: TipShape, dx: number, dy: number): number {
  let sum = 0
  for (let j = 0; j < SUPERSAMPLE; j += 1) {
    const sy = dy + (j + 0.5) / SUPERSAMPLE - 0.5
    for (let i = 0; i < SUPERSAMPLE; i += 1) {
      const sx = dx + (i + 0.5) / SUPERSAMPLE - 0.5
      let r: number
      if (shape.round) {
        r = Math.sqrt(sx * sx + sy * sy)
      } else {
        const u = sx * shape.cos + sy * shape.sin
        const v = (-sx * shape.sin + sy * shape.cos) / shape.roundness
        r = Math.sqrt(u * u + v * v)
      }
      if (r > shape.radius) continue
      if (shape.hard || r <= shape.inner) sum += 1
      else sum += 1 - smoothstep(clamp01((r - shape.inner) / (shape.radius - shape.inner)))
    }
  }
  return sum / (SUPERSAMPLE * SUPERSAMPLE)
}

function coverageFunction(shape: TipShape): (dx: number, dy: number) => number {
  return shape.radius < SUPERSAMPLE_BELOW_RADIUS ? (dx, dy) => supersampledAt(shape, dx, dy) : (dx, dy) => profileAt(shape, dx, dy)
}

/** Half extent of the dab footprint in px (covers the anti-aliased edge and rotated ellipses). */
function reach(shape: TipShape): number {
  return shape.radius + 1
}

/**
 * The tip's coverage on a size x size grid (odd size), centred on the middle pixel's centre. Used for
 * cursors, previews and tip thumbnails; strokes call accumulateDab, which evaluates the same profile at
 * the dab's exact sub-pixel position.
 */
export function dabMask(tip: BrushTip): { readonly size: number; readonly data: Float32Array } {
  const shape = tipShape(tip, tip.diameter)
  const half = Math.ceil(reach(shape))
  const size = 2 * half + 1
  const data = new Float32Array(size * size)
  const coverage = coverageFunction(shape)
  const centre = size / 2
  for (let y = 0; y < size; y += 1) {
    for (let x = 0; x < size; x += 1) data[y * size + x] = coverage(x + 0.5 - centre, y + 0.5 - centre)
  }
  return { size, data }
}

function lerp(a: number, b: number, t: number): number {
  return a + (b - a) * t
}

/**
 * Dabs for newly arrived samples of a stroke. `carry` is updated in place and must start as
 * { distance: 0, last: null }; the first sample of a stroke always produces a dab. Positions are smoothed
 * with an exponential moving average (alpha = 1 - 0.8 * smoothing); pass smoothing 0 with the final
 * sample to catch up with the pointer on release. `spacing` is the fraction of the (pressure-scaled)
 * diameter between dabs; Photoshop's default is 0.25.
 */
export function placeDabs(samples: readonly StrokeSample[], tip: BrushTip, spacing: number, smoothing: number,
  carry: { distance: number; last: StrokeSample | null }, dynamics: BrushDynamics = {}): Dab[] {
  const dabs: Dab[] = []
  const baseDiameter = Math.max(0.1, Number.isFinite(tip.diameter) ? tip.diameter : 1)
  const baseFlow = clamp01(dynamics.flow ?? 1)
  const step = Math.max(0, Number.isFinite(spacing) ? spacing : 0.25)
  const alpha = 1 - 0.8 * clamp01(Number.isFinite(smoothing) ? smoothing : 0)
  const diameterAt = (pressure: number) => (dynamics.pressureSize ? baseDiameter * lerp(0.1, 1, clamp01(pressure)) : baseDiameter)
  const flowAt = (pressure: number) => (dynamics.pressureFlow ? baseFlow * clamp01(pressure) : baseFlow)
  const emit = (x: number, y: number, pressure: number) => {
    dabs.push({ x, y, diameter: diameterAt(pressure), flow: flowAt(pressure) })
  }
  for (const raw of samples) {
    if (!raw || !Number.isFinite(raw.x) || !Number.isFinite(raw.y)) continue
    const pressure = Number.isFinite(raw.pressure) ? clamp01(raw.pressure) : 1
    const last = carry.last
    if (!last) {
      const first = { x: raw.x, y: raw.y, pressure, time: raw.time }
      carry.last = first
      carry.distance = 0
      emit(first.x, first.y, pressure)
      continue
    }
    const next = {
      x: last.x + alpha * (raw.x - last.x),
      y: last.y + alpha * (raw.y - last.y),
      pressure: last.pressure + alpha * (pressure - last.pressure),
      time: raw.time,
    }
    const dx = next.x - last.x
    const dy = next.y - last.y
    const length = Math.sqrt(dx * dx + dy * dy)
    if (length > 0) {
      let position = 0
      for (;;) {
        const t = position / length
        const gap = Math.max(MIN_STEP, step * diameterAt(lerp(last.pressure, next.pressure, t)))
        // A pressure drop can shrink the gap below the distance already travelled: dab right here.
        const need = Math.max(0, gap - carry.distance)
        if (position + need > length) {
          carry.distance += length - position
          break
        }
        position += need
        carry.distance = 0
        const u = position / length
        emit(last.x + dx * u, last.y + dy * u, lerp(last.pressure, next.pressure, u))
      }
    }
    carry.last = next
  }
  return dabs
}

/**
 * Accumulates one dab into a per-stroke coverage buffer (width x height floats whose (0, 0) sits at
 * (originX, originY) in the dab's coordinate space): S = S + dab * flow * (1 - S), so S never exceeds 1.
 * Returns the touched rectangle in the dab's coordinate space (width 0 when the dab misses the buffer).
 */
export function accumulateDab(coverage: Float32Array, width: number, height: number, originX: number, originY: number,
  dab: Dab, tip: BrushTip): IntRect {
  if (coverage.length < width * height) throw new RangeError('The coverage buffer is smaller than width x height.')
  const shape = tipShape(tip, dab.diameter)
  const flow = clamp01(dab.flow)
  const cx = dab.x - originX
  const cy = dab.y - originY
  const extent = reach(shape)
  const x0 = Math.max(0, Math.floor(cx - extent))
  const y0 = Math.max(0, Math.floor(cy - extent))
  const x1 = Math.min(width, Math.ceil(cx + extent))
  const y1 = Math.min(height, Math.ceil(cy + extent))
  if (!(x1 > x0 && y1 > y0) || flow === 0 || !Number.isFinite(cx) || !Number.isFinite(cy)) {
    return { x: originX + Math.min(Math.max(0, Math.round(cx)), width), y: originY + Math.min(Math.max(0, Math.round(cy)), height), width: 0, height: 0 }
  }
  const at = coverageFunction(shape)
  for (let y = y0; y < y1; y += 1) {
    const dy = y + 0.5 - cy
    const row = y * width
    for (let x = x0; x < x1; x += 1) {
      const value = at(x + 0.5 - cx, dy)
      if (value <= 0) continue
      const a = value * flow
      const s = coverage[row + x]
      coverage[row + x] = s + a * (1 - s)
    }
  }
  return { x: originX + x0, y: originY + y0, width: x1 - x0, height: y1 - y0 }
}

/** How a stroke's accumulated coverage is applied to the pre-stroke pixels (normal blend mode). */
export interface StrokePaint {
  readonly color: Rgb8
  /** 0..1; the most a single stroke can cover (Photoshop opacity). */
  readonly opacity: number
  /** Eraser: alpha' = alpha * (1 - S * opacity * selection); colour untouched; no effect with preserveAlpha. */
  readonly erase?: boolean
  /** Transparency lock: alpha is kept, colour mixes only where the pixel already has alpha. */
  readonly preserveAlpha?: boolean
}

/**
 * Applies a stroke's coverage (same size as `before`) to the pre-stroke pixels in normal mode:
 * source-over with source alpha S * opacity * selection, computed in straight alpha with one rounding.
 * Returns a new buffer; `before` is not modified. Blend modes other than normal use blend.ts.
 */
export function compositeStroke(before: PixelBuffer, coverage: Float32Array, paint: StrokePaint, selection?: MaskBuffer | null): PixelBuffer {
  const { width, height } = before
  const count = width * height
  if (before.data.length !== count * 4) throw new RangeError('The pixel buffer does not match its size.')
  if (coverage.length < count) throw new RangeError('The coverage buffer does not match the pixels.')
  if (selection && (selection.width !== width || selection.height !== height)) throw new RangeError('The selection does not match the pixels.')
  const src = before.data
  const data = new Uint8ClampedArray(src)
  const opacity = clamp01(paint.opacity)
  const cr = paint.color.r
  const cg = paint.color.g
  const cb = paint.color.b
  const sel = selection?.data ?? null
  for (let i = 0, p = 0; i < count; i += 1, p += 4) {
    let a = clamp01(coverage[i]) * opacity
    if (sel) a *= sel[i] / 255
    if (a <= 0) continue
    const ab = src[p + 3] / 255
    if (paint.erase) {
      // A transparency-locked layer keeps its alpha; the eraser tool paints the background colour there.
      if (!paint.preserveAlpha) data[p + 3] = Math.round(src[p + 3] * (1 - a))
      continue
    }
    if (paint.preserveAlpha) {
      if (ab === 0) continue
      data[p] = Math.round(src[p] + (cr - src[p]) * a)
      data[p + 1] = Math.round(src[p + 1] + (cg - src[p + 1]) * a)
      data[p + 2] = Math.round(src[p + 2] + (cb - src[p + 2]) * a)
      continue
    }
    const ao = a + ab * (1 - a)
    const kb = (ab * (1 - a)) / ao
    const ks = a / ao
    data[p] = Math.round(cr * ks + src[p] * kb)
    data[p + 1] = Math.round(cg * ks + src[p + 1] * kb)
    data[p + 2] = Math.round(cb * ks + src[p + 2] * kb)
    data[p + 3] = Math.round(ao * 255)
  }
  return { width, height, data }
}
