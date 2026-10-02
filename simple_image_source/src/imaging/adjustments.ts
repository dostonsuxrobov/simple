// src/imaging/adjustments.ts (WP2)
// The 13 Photoshop-style adjustments plus Simple mode's quick adjust (design section 5.10).
// Values are gamma-encoded sRGB normalised to [0, 1]. Per-channel operations compile to three 256-entry
// lookup tables (exact, several hundred Mpx/s); colour-mixing operations run per pixel after the tables.
// Lum = 0.3 R + 0.59 G + 0.11 B (W3C) for colour mixing; Y601 = 0.299 R + 0.587 G + 0.114 B for Threshold and
// Gradient Map. Photoshop's exact formulas are proprietary; Brightness/Contrast, Vibrance, Black & White,
// Photo Filter and the Hue/Saturation range windows are documented approximations. Levels and Curves follow
// their published definitions. Pure and DOM-free.
import type {
  AdjustmentKernel,
  AdjustmentSpec,
  AdjustmentType,
  ColorBalanceShift,
  CurvePoint,
  GradientStop,
  HslShift,
  HueRange,
  LevelsChannel,
  MaskBuffer,
  OpOptions,
  PixelBuffer,
  QuickAdjust,
  Rgb8,
} from './types.ts'
import { assertBuffer, assertMask, chunkRowsFor, finishedRun, runRowsSync, throwIfAborted } from './buffer.ts'
import type { RowRun } from './buffer.ts'
import { SRGB_TO_LINEAR, linearToSrgb } from './color.ts'
import { hash2 } from './random.ts'

// ---------------------------------------------------------------------------------------------
// Catalogue
// ---------------------------------------------------------------------------------------------

/** The 13 adjustment types in Photoshop's Image > Adjustments menu order ('quick' is Simple-only). */
export const ADJUSTMENT_TYPES: readonly AdjustmentType[] = Object.freeze([
  'brightness-contrast', 'levels', 'curves', 'exposure', 'vibrance', 'hue-saturation', 'color-balance',
  'black-white', 'photo-filter', 'invert', 'posterize', 'threshold', 'gradient-map',
] as const)

const LABELS: Readonly<Record<AdjustmentType, string>> = Object.freeze({
  'brightness-contrast': 'Brightness/Contrast',
  levels: 'Levels',
  curves: 'Curves',
  exposure: 'Exposure',
  vibrance: 'Vibrance',
  'hue-saturation': 'Hue/Saturation',
  'color-balance': 'Color Balance',
  'black-white': 'Black & White',
  'photo-filter': 'Photo Filter',
  invert: 'Invert',
  posterize: 'Posterize',
  threshold: 'Threshold',
  'gradient-map': 'Gradient Map',
  quick: 'Adjust',
})

/** Menu and history label, e.g. "Hue/Saturation". */
export function adjustmentLabel(type: AdjustmentType): string {
  return LABELS[type] ?? 'Adjustment'
}

function identityLevels(): LevelsChannel {
  return { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 }
}

function identityCurve(): CurvePoint[] {
  return [{ x: 0, y: 0 }, { x: 255, y: 255 }]
}

function zeroBalance(): ColorBalanceShift {
  return { cyanRed: 0, magentaGreen: 0, yellowBlue: 0 }
}

function createDefault(type: AdjustmentType): AdjustmentSpec {
  switch (type) {
    case 'brightness-contrast': return { type, brightness: 0, contrast: 0, legacy: false }
    case 'levels': return { type, rgb: identityLevels(), red: identityLevels(), green: identityLevels(), blue: identityLevels() }
    case 'curves': return { type, rgb: identityCurve(), red: identityCurve(), green: identityCurve(), blue: identityCurve() }
    case 'exposure': return { type, exposure: 0, offset: 0, gamma: 1 }
    case 'vibrance': return { type, vibrance: 0, saturation: 0 }
    case 'hue-saturation': return { type, colorize: false, master: { hue: 0, saturation: 0, lightness: 0 }, ranges: {} }
    case 'color-balance': return { type, shadows: zeroBalance(), midtones: zeroBalance(), highlights: zeroBalance(), preserveLuminosity: true }
    case 'black-white': return { type, reds: 40, yellows: 60, greens: 40, cyans: 60, blues: 20, magentas: 80, tint: null }
    // Photoshop's default Photo Filter: Warming Filter (85), density 25%.
    case 'photo-filter': return { type, color: { r: 236, g: 138, b: 0 }, density: 25, preserveLuminosity: true }
    case 'invert': return { type }
    case 'posterize': return { type, levels: 4 }
    case 'threshold': return { type, level: 128 }
    case 'gradient-map': return {
      type,
      stops: [{ position: 0, color: { r: 0, g: 0, b: 0 }, midpoint: 0.5 }, { position: 1, color: { r: 255, g: 255, b: 255 } }],
      reverse: false,
      dither: false,
    }
    case 'quick': return { type, exposure: 0, brightness: 0, contrast: 0, highlights: 0, shadows: 0, saturation: 0, warmth: 0, auto: null }
    default: throw new RangeError(`Unknown adjustment type "${String(type)}".`)
  }
}

/** Photoshop's default settings for an adjustment (a fresh object). Most defaults are the identity. */
export function defaultAdjustment<T extends AdjustmentType>(type: T): Extract<AdjustmentSpec, { type: T }> {
  return createDefault(type) as Extract<AdjustmentSpec, { type: T }>
}

// ---------------------------------------------------------------------------------------------
// Small numeric helpers
// ---------------------------------------------------------------------------------------------

function clamp01(value: number): number {
  return value <= 0 ? 0 : value >= 1 ? 1 : value
}

/** A finite number clamped to [min, max], or `fallback` when the input is not a finite number. */
function num(value: unknown, fallback: number, min: number, max: number): number {
  const n = typeof value === 'number' && Number.isFinite(value) ? value : fallback
  return n < min ? min : n > max ? max : n
}

function smoothstep01(value: number): number {
  return value * value * (3 - 2 * value)
}

// ---------------------------------------------------------------------------------------------
// Lookup tables
// ---------------------------------------------------------------------------------------------

/** Three 256-entry tables back to back: R at 0, G at 256, B at 512. */
type ChannelLut = Uint8Array

function buildLut(red: (value: number) => number, green: (value: number) => number, blue: (value: number) => number): ChannelLut {
  const lut = new Uint8Array(768)
  for (let value = 0; value < 256; value += 1) {
    lut[value] = clampByte(red(value))
    lut[256 + value] = clampByte(green(value))
    lut[512 + value] = clampByte(blue(value))
  }
  return lut
}

function clampByte(value: number): number {
  if (!(value > 0)) return 0
  if (value >= 255) return 255
  return Math.round(value)
}

function isIdentityLut(lut: ChannelLut): boolean {
  for (let index = 0; index < 768; index += 1) {
    if (lut[index] !== (index & 255)) return false
  }
  return true
}

/** lut2 after lut1. */
function composeLuts(first: ChannelLut, second: ChannelLut): ChannelLut {
  const out = new Uint8Array(768)
  for (let channel = 0; channel < 768; channel += 256) {
    for (let value = 0; value < 256; value += 1) out[channel + value] = second[channel + first[channel + value]]
  }
  return out
}

function applyLut(data: Uint8ClampedArray, lut: ChannelLut, start: number, end: number): void {
  for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
    data[i] = lut[data[i]]
    data[i + 1] = lut[256 + data[i + 1]]
    data[i + 2] = lut[512 + data[i + 2]]
  }
}

// ---------------------------------------------------------------------------------------------
// Levels and Curves
// ---------------------------------------------------------------------------------------------

/** Photoshop Levels on a 0..255 value (float in, float out): ob + (ow - ob) * clamp((v - ib) / (iw - ib))^(1 / gamma). */
function levelsValue(channel: LevelsChannel, value: number): number {
  const inBlack = num(channel.inBlack, 0, 0, 255)
  const inWhite = num(channel.inWhite, 255, 0, 255)
  const gamma = num(channel.gamma, 1, 0.01, 9.99)
  const outBlack = num(channel.outBlack, 0, 0, 255)
  const outWhite = num(channel.outWhite, 255, 0, 255)
  let t = inWhite > inBlack ? (value - inBlack) / (inWhite - inBlack) : value >= inBlack ? 1 : 0
  t = clamp01(t)
  if (gamma !== 1) t = Math.pow(t, 1 / gamma)
  return outBlack + (outWhite - outBlack) * t
}

function isIdentityLevels(channel: LevelsChannel): boolean {
  return num(channel.inBlack, 0, 0, 255) === 0 && num(channel.inWhite, 255, 0, 255) === 255 && num(channel.gamma, 1, 0.01, 9.99) === 1
    && num(channel.outBlack, 0, 0, 255) === 0 && num(channel.outWhite, 255, 0, 255) === 255
}

/** One Levels channel as a 256-entry table. */
export function levelsLut(channel: LevelsChannel): Uint8Array {
  const lut = new Uint8Array(256)
  for (let value = 0; value < 256; value += 1) lut[value] = clampByte(levelsValue(channel, value))
  return lut
}

/**
 * A curve as a function on 0..255 (float in, float out): natural cubic spline through the points, flat
 * before the first and after the last point, clamped to [0, 255]. Points are sorted; for duplicate x the
 * last one wins; no points = identity; one point = constant.
 */
export function curveFunction(points: readonly CurvePoint[]): (x: number) => number {
  const byX = new Map<number, number>()
  for (const point of points ?? []) {
    if (!point || !Number.isFinite(point.x) || !Number.isFinite(point.y)) continue
    byX.set(num(point.x, 0, 0, 255), num(point.y, 0, 0, 255))
  }
  const xs = [...byX.keys()].sort((a, b) => a - b)
  const ys = xs.map((x) => byX.get(x) as number)
  const n = xs.length
  if (n === 0) return (x) => (x <= 0 ? 0 : x >= 255 ? 255 : x)
  if (n === 1) return () => ys[0]
  // Second derivatives with natural boundary conditions (Thomas algorithm).
  const second = new Float64Array(n)
  if (n > 2) {
    const c = new Float64Array(n)
    const d = new Float64Array(n)
    for (let i = 1; i < n - 1; i += 1) {
      const h0 = xs[i] - xs[i - 1]
      const h1 = xs[i + 1] - xs[i]
      const a = h0
      const b = 2 * (h0 + h1)
      const rhs = 6 * ((ys[i + 1] - ys[i]) / h1 - (ys[i] - ys[i - 1]) / h0)
      const denominator = b - a * c[i - 1]
      c[i] = h1 / denominator
      d[i] = (rhs - a * d[i - 1]) / denominator
    }
    for (let i = n - 2; i >= 1; i -= 1) second[i] = d[i] - c[i] * second[i + 1]
  }
  return (x: number) => {
    if (x <= xs[0]) return ys[0]
    if (x >= xs[n - 1]) return ys[n - 1]
    let segment = 0
    while (segment < n - 2 && x > xs[segment + 1]) segment += 1
    const x0 = xs[segment]
    const x1 = xs[segment + 1]
    const h = x1 - x0
    const a = (x1 - x) / h
    const b = (x - x0) / h
    const y = a * ys[segment] + b * ys[segment + 1] + ((a * a * a - a) * second[segment] + (b * b * b - b) * second[segment + 1]) * (h * h) / 6
    return y <= 0 ? 0 : y >= 255 ? 255 : y
  }
}

/** A curve as a 256-entry table (see curveFunction). */
export function curveLut(points: readonly CurvePoint[]): Uint8Array {
  const curve = curveFunction(points)
  const lut = new Uint8Array(256)
  for (let value = 0; value < 256; value += 1) lut[value] = clampByte(curve(value))
  return lut
}

// ---------------------------------------------------------------------------------------------
// Tone helpers shared by Brightness/Contrast and quick adjust
// ---------------------------------------------------------------------------------------------

/** Modern contrast on [0, 1]: c > 0 blends towards smoothstep, c < 0 scales around 0.5. c in -100..100. */
function modernContrast(value: number, contrast: number): number {
  if (contrast > 0) return value + (smoothstep01(value) - value) * (contrast / 100)
  if (contrast < 0) return 0.5 + (value - 0.5) * (1 + contrast / 100)
  return value
}

/** Fritsch-Carlson monotone cubic through (xs, ys); xs strictly increasing. */
function monotoneCubic(xs: readonly number[], ys: readonly number[]): (x: number) => number {
  const n = xs.length
  const delta = new Float64Array(n - 1)
  for (let k = 0; k < n - 1; k += 1) delta[k] = (ys[k + 1] - ys[k]) / (xs[k + 1] - xs[k])
  const m = new Float64Array(n)
  m[0] = delta[0]
  m[n - 1] = delta[n - 2]
  for (let k = 1; k < n - 1; k += 1) m[k] = delta[k - 1] * delta[k] <= 0 ? 0 : (delta[k - 1] + delta[k]) / 2
  for (let k = 0; k < n - 1; k += 1) {
    if (delta[k] === 0) {
      m[k] = 0
      m[k + 1] = 0
      continue
    }
    const a = m[k] / delta[k]
    const b = m[k + 1] / delta[k]
    const norm = a * a + b * b
    if (norm > 9) {
      const tau = 3 / Math.sqrt(norm)
      m[k] = tau * a * delta[k]
      m[k + 1] = tau * b * delta[k]
    }
  }
  return (x: number) => {
    if (x <= xs[0]) return ys[0]
    if (x >= xs[n - 1]) return ys[n - 1]
    let k = 0
    while (k < n - 2 && x > xs[k + 1]) k += 1
    const h = xs[k + 1] - xs[k]
    const t = (x - xs[k]) / h
    const t2 = t * t
    const t3 = t2 * t
    return (2 * t3 - 3 * t2 + 1) * ys[k] + (t3 - 2 * t2 + t) * h * m[k] + (-2 * t3 + 3 * t2) * ys[k + 1] + (t3 - t2) * h * m[k + 1]
  }
}

// ---------------------------------------------------------------------------------------------
// Per-pixel stages
// ---------------------------------------------------------------------------------------------

/** In place on straight RGBA8, pixels [start, end); `seed` drives dithering. Alpha is never touched. */
type PixelStage = (data: Uint8ClampedArray, start: number, end: number, seed: number) => void

/** Writes an RGB triple in [0, 1] (W3C ClipColor applied around luminosity `l`). */
function writeClipped(data: Uint8ClampedArray, i: number, r: number, g: number, b: number): void {
  const l = 0.3 * r + 0.59 * g + 0.11 * b
  const n = r < g ? (r < b ? r : b) : (g < b ? g : b)
  const x = r > g ? (r > b ? r : b) : (g > b ? g : b)
  if (n < 0 && l - n > 1e-12) {
    const f = l / (l - n)
    r = l + (r - l) * f
    g = l + (g - l) * f
    b = l + (b - l) * f
  }
  if (x > 1 && x - l > 1e-12) {
    const f = (1 - l) / (x - l)
    r = l + (r - l) * f
    g = l + (g - l) * f
    b = l + (b - l) * f
  }
  data[i] = r * 255
  data[i + 1] = g * 255
  data[i + 2] = b * 255
}

/** c' = L + (c - L) * factor with Lum and ClipColor: the shared saturation step of vibrance and quick adjust. */
function saturationStage(factor: number): PixelStage {
  return (data, start, end) => {
    for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
      const r = data[i] / 255
      const g = data[i + 1] / 255
      const b = data[i + 2] / 255
      if (r === g && g === b) continue
      const l = 0.3 * r + 0.59 * g + 0.11 * b
      writeClipped(data, i, l + (r - l) * factor, l + (g - l) * factor, l + (b - l) * factor)
    }
  }
}

function vibranceStage(vibrance: number, saturation: number): PixelStage {
  const kVibrance = vibrance / 100
  const kSaturation = 1 + saturation / 100
  return (data, start, end) => {
    for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
      const r = data[i] / 255
      const g = data[i + 1] / 255
      const b = data[i + 2] / 255
      const max = r > g ? (r > b ? r : b) : (g > b ? g : b)
      const min = r < g ? (r < b ? r : b) : (g < b ? g : b)
      const chroma = max - min
      if (chroma <= 0) continue
      let kv = kVibrance * (1 - chroma) * (1 - chroma)
      // Skin protection: hues 0..50 degrees (red is the maximum, green >= blue, (g - b) / chroma <= 5/6).
      if (kv !== 0 && max === r && g >= b && (g - b) <= (5 / 6) * chroma) kv *= 0.5
      const factor = (1 + kv) * kSaturation
      if (factor === 1) continue
      const l = 0.3 * r + 0.59 * g + 0.11 * b
      writeClipped(data, i, l + (r - l) * factor, l + (g - l) * factor, l + (b - l) * factor)
    }
  }
}

const RANGE_CENTRES: Readonly<Record<HueRange, number>> = Object.freeze({
  reds: 0, yellows: 60, greens: 120, cyans: 180, blues: 240, magentas: 300,
})

/** Photoshop's default range windows: full weight within 15 degrees of the centre, linear fall-off to 45. */
export function hueRangeWeight(range: HueRange, hue: number): number {
  let distance = Math.abs((((hue - RANGE_CENTRES[range]) % 360) + 540) % 360 - 180)
  if (!Number.isFinite(distance)) distance = 180
  if (distance <= 15) return 1
  if (distance >= 45) return 0
  return (45 - distance) / 30
}

function hueToChannel(p: number, q: number, t: number): number {
  if (t < 0) t += 1
  else if (t > 1) t -= 1
  if (t < 1 / 6) return p + (q - p) * 6 * t
  if (t < 0.5) return q
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
  return p
}

function lighten(value: number, lightness: number): number {
  return lightness > 0 ? value + (1 - value) * (lightness / 100) : value * (1 + lightness / 100)
}

function hueSaturationStage(spec: Extract<AdjustmentSpec, { type: 'hue-saturation' }>): PixelStage {
  const master = readShift(spec.master)
  const ranges: { readonly centre: HueRange; readonly shift: HslShift }[] = []
  if (!spec.colorize && spec.ranges) {
    for (const key of Object.keys(RANGE_CENTRES) as HueRange[]) {
      const shift = spec.ranges[key]
      if (!shift) continue
      const clean = readShift(shift)
      if (clean.hue || clean.saturation || clean.lightness) ranges.push({ centre: key, shift: clean })
    }
  }
  const colorizeHue = ((master.hue % 360) + 360) % 360
  const colorizeSat = Math.max(0, master.saturation) / 100
  return (data, start, end) => {
    for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
      let r = data[i] / 255
      let g = data[i + 1] / 255
      let b = data[i + 2] / 255
      const max = r > g ? (r > b ? r : b) : (g > b ? g : b)
      const min = r < g ? (r < b ? r : b) : (g < b ? g : b)
      const l = (max + min) / 2
      const d = max - min
      let hue = 0
      let s = 0
      if (d > 0) {
        s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
        if (max === r) hue = (g - b) / d + (g < b ? 6 : 0)
        else if (max === g) hue = (b - r) / d + 2
        else hue = (r - g) / d + 4
        hue *= 60
      }
      let lightRange = 0
      let changed = false
      if (spec.colorize) {
        hue = colorizeHue
        s = colorizeSat
        changed = true
      } else {
        let hueShift = master.hue
        let satRange = 0
        if (d > 0) {
          for (const range of ranges) {
            const w = hueRangeWeight(range.centre, hue)
            if (w === 0) continue
            hueShift += w * range.shift.hue
            satRange += w * range.shift.saturation
            lightRange += w * range.shift.lightness
          }
        }
        if (d > 0 && (hueShift !== 0 || satRange !== 0 || master.saturation !== 0)) {
          hue += hueShift
          if (satRange !== 0) s = clamp01(s * (1 + satRange / 100))
          if (master.saturation !== 0) s = clamp01(s * (1 + master.saturation / 100))
          changed = true
        }
      }
      if (changed) {
        if (s <= 0) {
          r = l
          g = l
          b = l
        } else {
          let h = (hue % 360) / 360
          if (h < 0) h += 1
          const q = l < 0.5 ? l * (1 + s) : l + s - l * s
          const p = 2 * l - q
          r = hueToChannel(p, q, h + 1 / 3)
          g = hueToChannel(p, q, h)
          b = hueToChannel(p, q, h - 1 / 3)
        }
      }
      if (lightRange !== 0) {
        r = lighten(r, lightRange)
        g = lighten(g, lightRange)
        b = lighten(b, lightRange)
        changed = true
      }
      if (master.lightness !== 0) {
        r = lighten(r, master.lightness)
        g = lighten(g, master.lightness)
        b = lighten(b, master.lightness)
        changed = true
      }
      if (!changed) continue
      data[i] = r * 255
      data[i + 1] = g * 255
      data[i + 2] = b * 255
    }
  }
}

function readShift(shift: HslShift | undefined): HslShift {
  return {
    hue: num(shift?.hue, 0, -360, 360),
    saturation: num(shift?.saturation, 0, -100, 100),
    lightness: num(shift?.lightness, 0, -100, 100),
  }
}

function readBalance(shift: ColorBalanceShift | undefined): ColorBalanceShift {
  return {
    cyanRed: num(shift?.cyanRed, 0, -100, 100),
    magentaGreen: num(shift?.magentaGreen, 0, -100, 100),
    yellowBlue: num(shift?.yellowBlue, 0, -100, 100),
  }
}

/** GIMP's colour balance: tonal ranges weighted by HSL lightness, shifts added per channel. */
function colorBalanceStage(spec: Extract<AdjustmentSpec, { type: 'color-balance' }>): PixelStage {
  const shadows = readBalance(spec.shadows)
  const midtones = readBalance(spec.midtones)
  const highlights = readBalance(spec.highlights)
  const preserve = spec.preserveLuminosity !== false
  // Offsets per channel indexed by max + min (0..510, lightness = index / 510).
  const offR = new Float64Array(511)
  const offG = new Float64Array(511)
  const offB = new Float64Array(511)
  const a = 0.25
  const bEdge = 0.333
  const scale = 0.7
  for (let index = 0; index <= 510; index += 1) {
    const l = index / 510
    const ws = clamp01((l - bEdge) / -a + 0.5) * scale
    const wm = clamp01((l - bEdge) / a + 0.5) * clamp01((l + bEdge - 1) / -a + 0.5) * scale
    const wh = clamp01((l + bEdge - 1) / a + 0.5) * scale
    offR[index] = (shadows.cyanRed * ws + midtones.cyanRed * wm + highlights.cyanRed * wh) / 100
    offG[index] = (shadows.magentaGreen * ws + midtones.magentaGreen * wm + highlights.magentaGreen * wh) / 100
    offB[index] = (shadows.yellowBlue * ws + midtones.yellowBlue * wm + highlights.yellowBlue * wh) / 100
  }
  return (data, start, end) => {
    for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
      const r8 = data[i]
      const g8 = data[i + 1]
      const b8 = data[i + 2]
      const max8 = r8 > g8 ? (r8 > b8 ? r8 : b8) : (g8 > b8 ? g8 : b8)
      const min8 = r8 < g8 ? (r8 < b8 ? r8 : b8) : (g8 < b8 ? g8 : b8)
      const index = max8 + min8
      let r = clamp01(r8 / 255 + offR[index])
      let g = clamp01(g8 / 255 + offG[index])
      let b = clamp01(b8 / 255 + offB[index])
      if (preserve) {
        // Restore the original HSL lightness, keeping the new hue and saturation.
        const lOriginal = index / 510
        const max = r > g ? (r > b ? r : b) : (g > b ? g : b)
        const min = r < g ? (r < b ? r : b) : (g < b ? g : b)
        const d = max - min
        if (d <= 0) {
          r = lOriginal
          g = lOriginal
          b = lOriginal
        } else {
          const l = (max + min) / 2
          const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
          let h: number
          if (max === r) h = (g - b) / d + (g < b ? 6 : 0)
          else if (max === g) h = (b - r) / d + 2
          else h = (r - g) / d + 4
          h /= 6
          const q = lOriginal < 0.5 ? lOriginal * (1 + s) : lOriginal + s - lOriginal * s
          const p = 2 * lOriginal - q
          r = hueToChannel(p, q, h + 1 / 3)
          g = hueToChannel(p, q, h)
          b = hueToChannel(p, q, h - 1 / 3)
        }
      }
      data[i] = r * 255
      data[i + 1] = g * 255
      data[i + 2] = b * 255
    }
  }
}

function blackWhiteStage(spec: Extract<AdjustmentSpec, { type: 'black-white' }>): PixelStage {
  const wReds = num(spec.reds, 40, -200, 300) / 100
  const wYellows = num(spec.yellows, 60, -200, 300) / 100
  const wGreens = num(spec.greens, 40, -200, 300) / 100
  const wCyans = num(spec.cyans, 60, -200, 300) / 100
  const wBlues = num(spec.blues, 20, -200, 300) / 100
  const wMagentas = num(spec.magentas, 80, -200, 300) / 100
  const tint = spec.tint ? readRgb(spec.tint) : null
  const tintR = tint ? tint.r / 255 : 0
  const tintG = tint ? tint.g / 255 : 0
  const tintB = tint ? tint.b / 255 : 0
  const tintLum = 0.3 * tintR + 0.59 * tintG + 0.11 * tintB
  return (data, start, end) => {
    for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
      const r = data[i] / 255
      const g = data[i + 1] / 255
      const b = data[i + 2] / 255
      let max: number
      let mid: number
      let min: number
      let primary: number
      let secondary: number
      if (r >= g) {
        if (g >= b) { max = r; mid = g; min = b; primary = wReds; secondary = wYellows }
        else if (r >= b) { max = r; mid = b; min = g; primary = wReds; secondary = wMagentas }
        else { max = b; mid = r; min = g; primary = wBlues; secondary = wMagentas }
      } else if (r >= b) {
        max = g; mid = r; min = b; primary = wGreens; secondary = wYellows
      } else if (g >= b) {
        max = g; mid = b; min = r; primary = wGreens; secondary = wCyans
      } else {
        max = b; mid = g; min = r; primary = wBlues; secondary = wCyans
      }
      const gray = clamp01(min + (mid - min) * secondary + (max - mid) * primary)
      if (!tint) {
        const v = gray * 255
        data[i] = v
        data[i + 1] = v
        data[i + 2] = v
        continue
      }
      const shift = gray - tintLum
      writeClipped(data, i, tintR + shift, tintG + shift, tintB + shift)
    }
  }
}

function readRgb(color: Rgb8): Rgb8 {
  return { r: num(color?.r, 0, 0, 255), g: num(color?.g, 0, 0, 255), b: num(color?.b, 0, 0, 255) }
}

/** W3C Overlay with the image as backdrop and the filter colour as source, per channel in [0, 1]. */
function overlay(backdrop: number, source: number): number {
  return backdrop <= 0.5 ? 2 * backdrop * source : 1 - 2 * (1 - backdrop) * (1 - source)
}

function photoFilterLut(spec: Extract<AdjustmentSpec, { type: 'photo-filter' }>): ChannelLut {
  const color = readRgb(spec.color)
  const density = num(spec.density, 25, 0, 100) / 100
  const channel = (filter: number) => (value: number) => {
    const c = value / 255
    return (c + (overlay(c, filter / 255) - c) * density) * 255
  }
  return buildLut(channel(color.r), channel(color.g), channel(color.b))
}

/** Photo Filter with "Preserve Luminosity": the filtered colour moved back to the original luminosity. */
function preserveLuminosityStage(lut: ChannelLut): PixelStage {
  return (data, start, end) => {
    for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
      const r0 = data[i] / 255
      const g0 = data[i + 1] / 255
      const b0 = data[i + 2] / 255
      const r = lut[data[i]] / 255
      const g = lut[256 + data[i + 1]] / 255
      const b = lut[512 + data[i + 2]] / 255
      const shift = (0.3 * r0 + 0.59 * g0 + 0.11 * b0) - (0.3 * r + 0.59 * g + 0.11 * b)
      writeClipped(data, i, r + shift, g + shift, b + shift)
    }
  }
}

function thresholdStage(level: number): PixelStage {
  const limit = level * 1000
  return (data, start, end) => {
    for (let i = start * 4, stop = end * 4; i < stop; i += 4) {
      const v = 299 * data[i] + 587 * data[i + 1] + 114 * data[i + 2] >= limit ? 255 : 0
      data[i] = v
      data[i + 1] = v
      data[i + 2] = v
    }
  }
}

/** Colour of a gradient at t in [0, 1] (stops sorted), channels in 0..255 floats. */
function gradientColorAt(stops: readonly GradientStop[], t: number, out: Float64Array, offset: number): void {
  const first = stops[0]
  const last = stops[stops.length - 1]
  if (t <= first.position) {
    out[offset] = first.color.r; out[offset + 1] = first.color.g; out[offset + 2] = first.color.b
    return
  }
  if (t >= last.position) {
    out[offset] = last.color.r; out[offset + 1] = last.color.g; out[offset + 2] = last.color.b
    return
  }
  let k = 0
  while (k < stops.length - 2 && t > stops[k + 1].position) k += 1
  const a = stops[k]
  const b = stops[k + 1]
  const span = b.position - a.position
  let u = span > 0 ? (t - a.position) / span : 1
  const midpoint = num(a.midpoint, 0.5, 0.01, 0.99)
  if (midpoint !== 0.5 && u > 0 && u < 1) u = Math.pow(u, Math.log(0.5) / Math.log(midpoint))
  out[offset] = a.color.r + (b.color.r - a.color.r) * u
  out[offset + 1] = a.color.g + (b.color.g - a.color.g) * u
  out[offset + 2] = a.color.b + (b.color.b - a.color.b) * u
}

/** Sorted, sanitised gradient stops (black to white when none are usable). */
export function normalizeStops(stops: readonly GradientStop[] | null | undefined): GradientStop[] {
  const clean = (stops ?? [])
    .filter((stop) => stop && stop.color && Number.isFinite(stop.position))
    .map((stop) => ({ position: clamp01(stop.position), color: readRgb(stop.color), midpoint: num(stop.midpoint, 0.5, 0.01, 0.99) }))
    .sort((a, b) => a.position - b.position)
  if (clean.length) return clean
  return [{ position: 0, color: { r: 0, g: 0, b: 0 }, midpoint: 0.5 }, { position: 1, color: { r: 255, g: 255, b: 255 }, midpoint: 0.5 }]
}

const GRADIENT_STEPS = 1020

function gradientMapStage(spec: Extract<AdjustmentSpec, { type: 'gradient-map' }>): PixelStage {
  const stops = normalizeStops(spec.stops)
  const table = new Float64Array((GRADIENT_STEPS + 1) * 3)
  for (let index = 0; index <= GRADIENT_STEPS; index += 1) {
    const t = index / GRADIENT_STEPS
    gradientColorAt(stops, spec.reverse ? 1 - t : t, table, index * 3)
  }
  const dither = Boolean(spec.dither)
  return (data, start, end, seed) => {
    for (let pixel = start; pixel < end; pixel += 1) {
      const i = pixel * 4
      // Y601 in quarter levels: 0..1020.
      const index = ((299 * data[i] + 587 * data[i + 1] + 114 * data[i + 2]) / 250 + 0.5) | 0
      const t = index * 3
      const offset = dither ? hash2(pixel, 0, seed) - 0.5 : 0
      data[i] = table[t] + offset
      data[i + 1] = table[t + 1] + offset
      data[i + 2] = table[t + 2] + offset
    }
  }
}

// ---------------------------------------------------------------------------------------------
// Quick adjust (Simple mode)
// ---------------------------------------------------------------------------------------------

function readQuick(spec: QuickAdjust): QuickAdjust {
  return {
    exposure: num(spec.exposure, 0, -100, 100),
    brightness: num(spec.brightness, 0, -100, 100),
    contrast: num(spec.contrast, 0, -100, 100),
    highlights: num(spec.highlights, 0, -100, 100),
    shadows: num(spec.shadows, 0, -100, 100),
    saturation: num(spec.saturation, 0, -100, 100),
    warmth: num(spec.warmth, 0, -100, 100),
    auto: spec.auto && typeof spec.auto === 'object' ? spec.auto : null,
  }
}

function quickLut(quick: QuickAdjust): ChannelLut {
  const exposureGain = Math.pow(2, (1.5 * quick.exposure) / 100)
  const warm = Math.pow(2, (0.25 * quick.warmth) / 100)
  const brightnessExponent = Math.pow(2, -quick.brightness / 100)
  const tone = quick.shadows !== 0 || quick.highlights !== 0
    ? monotoneCubic([0, 0.25, 0.5, 0.75, 1], [0, 0.25 + (0.12 * quick.shadows) / 100, 0.5, 0.75 + (0.12 * quick.highlights) / 100, 1])
    : null
  const auto = quick.auto
  const channel = (gain: number) => (value: number) => {
    let v = value
    if (gain !== 1) v = linearToSrgb(SRGB_TO_LINEAR[value] * gain) * 255
    if (auto) v = levelsValue(auto, v)
    let u = v / 255
    if (brightnessExponent !== 1) u = Math.pow(clamp01(u), brightnessExponent)
    if (tone) u = tone(u)
    if (quick.contrast !== 0) u = modernContrast(clamp01(u), quick.contrast)
    return u * 255
  }
  return buildLut(channel(exposureGain * warm), channel(exposureGain), channel(exposureGain / warm))
}

// ---------------------------------------------------------------------------------------------
// Compilation
// ---------------------------------------------------------------------------------------------

interface CompiledParts {
  /** Runs first; null when the per-channel part is the identity. */
  readonly lut: ChannelLut | null
  /** Runs after the tables; null when absent. */
  readonly pixel: PixelStage | null
}

const PARTS = new WeakMap<AdjustmentKernel, CompiledParts>()

function lutParts(lut: ChannelLut, pixel: PixelStage | null = null): CompiledParts {
  return { lut: isIdentityLut(lut) ? null : lut, pixel }
}

function compileParts(spec: AdjustmentSpec): CompiledParts {
  switch (spec.type) {
    case 'brightness-contrast': {
      const brightness = num(spec.brightness, 0, -150, 150)
      const contrast = num(spec.contrast, 0, -100, 100)
      if (brightness === 0 && contrast === 0) return { lut: null, pixel: null }
      const exponent = Math.pow(2, -brightness / 100)
      const f = spec.legacy
        ? (value: number) => clamp01(0.5 + (clamp01(value / 255 + brightness / 255) - 0.5) * (1 + contrast / 100)) * 255
        : (value: number) => clamp01(modernContrast(Math.pow(value / 255, exponent), contrast)) * 255
      return lutParts(buildLut(f, f, f))
    }
    case 'levels': {
      const rgb = spec.rgb ?? identityLevels()
      const channel = (levels: LevelsChannel | undefined) => {
        const own = levels ?? identityLevels()
        const ownIdentity = isIdentityLevels(own)
        return (value: number) => levelsValue(rgb, ownIdentity ? value : levelsValue(own, value))
      }
      return lutParts(buildLut(channel(spec.red), channel(spec.green), channel(spec.blue)))
    }
    case 'curves': {
      const master = curveFunction(spec.rgb ?? identityCurve())
      const channel = (points: readonly CurvePoint[] | undefined) => {
        const own = curveFunction(points ?? identityCurve())
        return (value: number) => master(own(value))
      }
      return lutParts(buildLut(channel(spec.red), channel(spec.green), channel(spec.blue)))
    }
    case 'exposure': {
      const gain = Math.pow(2, num(spec.exposure, 0, -20, 20))
      const offset = num(spec.offset, 0, -0.5, 0.5)
      const gamma = num(spec.gamma, 1, 0.01, 9.99)
      if (gain === 1 && offset === 0 && gamma === 1) return { lut: null, pixel: null }
      const f = (value: number) => {
        const x = SRGB_TO_LINEAR[value] * gain + offset
        if (!(x > 0)) return 0
        return linearToSrgb(gamma === 1 ? x : Math.pow(x, 1 / gamma)) * 255
      }
      return lutParts(buildLut(f, f, f))
    }
    case 'vibrance': {
      const vibrance = num(spec.vibrance, 0, -100, 100)
      const saturation = num(spec.saturation, 0, -100, 100)
      if (vibrance === 0 && saturation === 0) return { lut: null, pixel: null }
      return { lut: null, pixel: vibranceStage(vibrance, saturation) }
    }
    case 'hue-saturation': {
      const master = readShift(spec.master)
      const anyRange = Object.values(spec.ranges ?? {}).some((shift) => {
        const clean = readShift(shift)
        return clean.hue !== 0 || clean.saturation !== 0 || clean.lightness !== 0
      })
      if (!spec.colorize && master.hue % 360 === 0 && master.saturation === 0 && master.lightness === 0 && !anyRange) {
        return { lut: null, pixel: null }
      }
      return { lut: null, pixel: hueSaturationStage(spec) }
    }
    case 'color-balance': {
      const all = [spec.shadows, spec.midtones, spec.highlights].map(readBalance)
      if (all.every((shift) => shift.cyanRed === 0 && shift.magentaGreen === 0 && shift.yellowBlue === 0)) return { lut: null, pixel: null }
      return { lut: null, pixel: colorBalanceStage(spec) }
    }
    case 'black-white':
      return { lut: null, pixel: blackWhiteStage(spec) }
    case 'photo-filter': {
      const lut = photoFilterLut(spec)
      if (isIdentityLut(lut)) return { lut: null, pixel: null }
      if (spec.preserveLuminosity === false) return { lut, pixel: null }
      return { lut: null, pixel: preserveLuminosityStage(lut) }
    }
    case 'invert': {
      const f = (value: number) => 255 - value
      return lutParts(buildLut(f, f, f))
    }
    case 'posterize': {
      const levels = Math.round(num(spec.levels, 4, 2, 255))
      const f = (value: number) => (Math.round((value / 255) * (levels - 1)) / (levels - 1)) * 255
      return lutParts(buildLut(f, f, f))
    }
    case 'threshold':
      return { lut: null, pixel: thresholdStage(num(spec.level, 128, 1, 255)) }
    case 'gradient-map':
      return { lut: null, pixel: gradientMapStage(spec) }
    case 'quick': {
      const quick = readQuick(spec)
      const pixel = quick.saturation !== 0 ? saturationStage(1 + quick.saturation / 100) : null
      return lutParts(quickLut(quick), pixel)
    }
    default:
      throw new RangeError(`Unknown adjustment type "${String((spec as { type?: unknown })?.type)}".`)
  }
}

/** Compiles a spec once; apply() then runs on any number of pixel ranges (tiles, stripes, previews). */
export function compileAdjustment(spec: AdjustmentSpec): AdjustmentKernel {
  if (!spec || typeof spec !== 'object') throw new RangeError('The adjustment is missing.')
  const parts = compileParts(spec)
  const isIdentity = !parts.lut && !parts.pixel
  const kernel: AdjustmentKernel = {
    spec,
    isIdentity,
    apply(data: Uint8ClampedArray, startPixel = 0, endPixel = data.length >> 2, ditherSeed = 0) {
      if (isIdentity) return
      const total = data.length >> 2
      const start = Math.max(0, Math.min(total, Math.floor(startPixel)))
      const end = Math.max(start, Math.min(total, Math.floor(endPixel)))
      if (parts.lut) applyLut(data, parts.lut, start, end)
      if (parts.pixel) parts.pixel(data, start, end, ditherSeed | 0)
    },
  }
  PARTS.set(kernel, parts)
  return Object.freeze(kernel)
}

/** Kernels in order with consecutive table-only kernels folded into one table. */
function planStages(specs: readonly AdjustmentSpec[]): CompiledParts[] {
  const stages: CompiledParts[] = []
  let pending: ChannelLut | null = null
  for (const spec of specs) {
    const parts = PARTS.get(compileAdjustment(spec)) as CompiledParts
    if (!parts.lut && !parts.pixel) continue
    if (parts.lut) pending = pending ? composeLuts(pending, parts.lut) : parts.lut
    if (parts.pixel) {
      stages.push({ lut: pending, pixel: parts.pixel })
      pending = null
    }
  }
  if (pending && !isIdentityLut(pending)) stages.push({ lut: pending, pixel: null })
  return stages
}

/**
 * The resumable form of applyAdjustments (worker handlers run it in chunks). Pixel indices are global, so
 * any chunking gives the same bytes as one whole-image run (gradient-map dither included).
 */
export function startAdjustments(src: PixelBuffer, specs: readonly AdjustmentSpec[], mask: MaskBuffer | null = null, opacity = 1): RowRun {
  assertBuffer(src)
  if (mask) assertMask(mask, src.width, src.height)
  const amount = opacity >= 1 ? 1 : opacity > 0 ? opacity : 0
  const output: PixelBuffer = { width: src.width, height: src.height, data: new Uint8ClampedArray(src.data) }
  const stages = amount > 0 ? planStages(specs ?? []) : []
  if (!stages.length) return finishedRun(output)
  const out = output.data
  const original = src.data
  const width = src.width
  return {
    output,
    rows: src.height,
    chunkRows: chunkRowsFor(width, 1 << 20),
    process(startRow, endRow) {
      const start = startRow * width
      const end = endRow * width
      for (const stage of stages) {
        if (stage.lut) applyLut(out, stage.lut, start, end)
        if (stage.pixel) stage.pixel(out, start, end, 0)
      }
      if (!mask && amount >= 1) return
      // Alpha is unchanged by every kernel, so a straight per-channel mix is exact.
      for (let pixel = start; pixel < end; pixel += 1) {
        const coverage = mask ? mask.data[pixel] : 255
        const w = coverage === 255 ? amount : (coverage / 255) * amount
        if (w >= 1) continue
        const i = pixel * 4
        if (w <= 0) {
          out[i] = original[i]; out[i + 1] = original[i + 1]; out[i + 2] = original[i + 2]
          continue
        }
        out[i] = original[i] + (out[i] - original[i]) * w
        out[i + 1] = original[i + 1] + (out[i + 1] - original[i + 1]) * w
        out[i + 2] = original[i + 2] + (out[i + 2] - original[i + 2]) * w
      }
    },
  }
}

/**
 * Applies `specs` in order to a copy of `src`, then mixes the result back by `mask` (0..255) and `opacity`
 * (0..1). Alpha is never changed. Returns a new buffer (identity specs give identical bytes).
 */
export function applyAdjustments(
  src: PixelBuffer,
  specs: readonly AdjustmentSpec[],
  mask: MaskBuffer | null = null,
  opacity = 1,
  options?: OpOptions,
): PixelBuffer {
  throwIfAborted(options?.signal)
  return runRowsSync(startAdjustments(src, specs, mask, opacity), options)
}
