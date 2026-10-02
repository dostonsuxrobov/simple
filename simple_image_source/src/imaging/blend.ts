// src/imaging/blend.ts (WP2)
// Photoshop's 27 layer blend modes in software (design section 5.5 and Appendix B). Compositing uses the
// W3C general formula on straight alpha:
//   as = srcA * opacity * coverage, ao = as + ab (1 - as),
//   Co = (as (1 - ab) Cs + as ab B(Cb, Cs) + (1 - as) ab Cb) / ao.
// Separable modes go through a lazily built 256 x 256 table of B per mode (exact at 8-bit inputs); the
// non-separable modes (Hue, Saturation, Color, Luminosity) use W3C Lum/SetLum/SetSat/ClipColor; Darker and
// Lighter Color compare R + G + B; Soft Light is Photoshop's variant (not the W3C one); Dissolve keeps a pixel
// when hash2(...) < as. Pure and DOM-free.
import type { BlendMode } from './types.ts'
import { hash2 } from './random.ts'

/** All 27 modes in Photoshop menu order. */
export const BLEND_MODES: readonly BlendMode[] = Object.freeze([
  'normal', 'dissolve',
  'darken', 'multiply', 'color-burn', 'linear-burn', 'darker-color',
  'lighten', 'screen', 'color-dodge', 'linear-dodge', 'lighter-color',
  'overlay', 'soft-light', 'hard-light', 'vivid-light', 'linear-light', 'pin-light', 'hard-mix',
  'difference', 'exclusion', 'subtract', 'divide',
  'hue', 'saturation', 'color', 'luminosity',
] as const)

const LABELS: Readonly<Record<BlendMode, string>> = Object.freeze({
  normal: 'Normal',
  dissolve: 'Dissolve',
  darken: 'Darken',
  multiply: 'Multiply',
  'color-burn': 'Color Burn',
  'linear-burn': 'Linear Burn',
  'darker-color': 'Darker Color',
  lighten: 'Lighten',
  screen: 'Screen',
  'color-dodge': 'Color Dodge',
  'linear-dodge': 'Linear Dodge (Add)',
  'lighter-color': 'Lighter Color',
  overlay: 'Overlay',
  'soft-light': 'Soft Light',
  'hard-light': 'Hard Light',
  'vivid-light': 'Vivid Light',
  'linear-light': 'Linear Light',
  'pin-light': 'Pin Light',
  'hard-mix': 'Hard Mix',
  difference: 'Difference',
  exclusion: 'Exclusion',
  subtract: 'Subtract',
  divide: 'Divide',
  hue: 'Hue',
  saturation: 'Saturation',
  color: 'Color',
  luminosity: 'Luminosity',
})

/** Photoshop's menu label, e.g. "Linear Dodge (Add)". */
export function blendModeLabel(mode: BlendMode): string {
  return isBlendMode(mode) ? LABELS[mode] : 'Normal'
}

/** True for the 27 supported mode names. */
export function isBlendMode(mode: unknown): mode is BlendMode {
  return typeof mode === 'string' && Object.prototype.hasOwnProperty.call(LABELS, mode)
}

/**
 * Canvas globalCompositeOperation for the 16 W3C modes, for cheap previews only (canvas Soft Light is the W3C
 * formula, and canvas storage is premultiplied 8-bit). The other 11 modes have no canvas equivalent.
 */
export const CANVAS_COMPOSITE: Readonly<Partial<Record<BlendMode, string>>> = Object.freeze({
  normal: 'source-over',
  multiply: 'multiply',
  screen: 'screen',
  overlay: 'overlay',
  darken: 'darken',
  lighten: 'lighten',
  'color-dodge': 'color-dodge',
  'color-burn': 'color-burn',
  'hard-light': 'hard-light',
  'soft-light': 'soft-light',
  difference: 'difference',
  exclusion: 'exclusion',
  hue: 'hue',
  saturation: 'saturation',
  color: 'color',
  luminosity: 'luminosity',
})

const NON_SEPARABLE: ReadonlySet<BlendMode> = new Set<BlendMode>(['darker-color', 'lighter-color', 'hue', 'saturation', 'color', 'luminosity'])

/** True for modes that work per channel (all but Darker/Lighter Color, Hue, Saturation, Color, Luminosity). */
export function isSeparableBlendMode(mode: BlendMode): boolean {
  return isBlendMode(mode) && !NON_SEPARABLE.has(mode)
}

// ---------------------------------------------------------------------------------------------
// Blend functions B(Cb, Cs) on [0, 1]
// ---------------------------------------------------------------------------------------------

function screen(cb: number, cs: number): number {
  return cb + cs - cb * cs
}

function colorBurn(cb: number, cs: number): number {
  if (cb >= 1) return 1
  if (cs <= 0) return 0
  return 1 - Math.min(1, (1 - cb) / cs)
}

function colorDodge(cb: number, cs: number): number {
  if (cb <= 0) return 0
  if (cs >= 1) return 1
  return Math.min(1, cb / (1 - cs))
}

function hardLight(cb: number, cs: number): number {
  return cs <= 0.5 ? cb * 2 * cs : screen(cb, 2 * cs - 1)
}

/**
 * Separable blend function B(Cb, Cs) for channel values in [0, 1]. Throws a RangeError for the non-separable
 * modes (use blendColor for those).
 */
export function blendChannel(mode: BlendMode, backdrop: number, source: number): number {
  const cb = backdrop
  const cs = source
  switch (mode) {
    case 'normal':
    case 'dissolve':
      return cs
    case 'darken': return Math.min(cb, cs)
    case 'multiply': return cb * cs
    case 'color-burn': return colorBurn(cb, cs)
    case 'linear-burn': return Math.max(0, cb + cs - 1)
    case 'lighten': return Math.max(cb, cs)
    case 'screen': return screen(cb, cs)
    case 'color-dodge': return colorDodge(cb, cs)
    case 'linear-dodge': return Math.min(1, cb + cs)
    case 'overlay': return hardLight(cs, cb)
    case 'soft-light':
      return cs <= 0.5
        ? 2 * cb * cs + cb * cb * (1 - 2 * cs)
        : 2 * cb * (1 - cs) + Math.sqrt(cb) * (2 * cs - 1)
    case 'hard-light': return hardLight(cb, cs)
    case 'vivid-light': return cs <= 0.5 ? colorBurn(cb, 2 * cs) : colorDodge(cb, 2 * (cs - 0.5))
    case 'linear-light': return Math.min(1, Math.max(0, cb + 2 * cs - 1))
    case 'pin-light': return cs <= 0.5 ? Math.min(cb, 2 * cs) : Math.max(cb, 2 * (cs - 0.5))
    // Tolerance keeps 8-bit pairs that sum to exactly 255 on the "1" side despite floating error.
    case 'hard-mix': return cb + cs >= 1 - 1e-9 ? 1 : 0
    case 'difference': return Math.abs(cb - cs)
    case 'exclusion': return cb + cs - 2 * cb * cs
    case 'subtract': return Math.max(0, cb - cs)
    case 'divide': return cs <= 0 ? 1 : Math.min(1, cb / cs)
    default:
      throw new RangeError(`"${String(mode)}" is not a separable blend mode; use blendColor.`)
  }
}

function lumOf(r: number, g: number, b: number): number {
  return 0.3 * r + 0.59 * g + 0.11 * b
}

/** W3C ClipColor + SetLum into `out` (avoids allocation in the pixel loop). */
function setLumInto(r: number, g: number, b: number, l: number, out: Float64Array): void {
  const d = l - lumOf(r, g, b)
  r += d
  g += d
  b += d
  const lum = lumOf(r, g, b)
  const n = Math.min(r, g, b)
  const x = Math.max(r, g, b)
  if (n < 0 && lum - n > 1e-12) {
    const f = lum / (lum - n)
    r = lum + (r - lum) * f
    g = lum + (g - lum) * f
    b = lum + (b - lum) * f
  }
  if (x > 1 && x - lum > 1e-12) {
    const f = (1 - lum) / (x - lum)
    r = lum + (r - lum) * f
    g = lum + (g - lum) * f
    b = lum + (b - lum) * f
  }
  out[0] = r
  out[1] = g
  out[2] = b
}

/** W3C SetSat into `out`. */
function setSatInto(r: number, g: number, b: number, s: number, out: Float64Array): void {
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  if (max <= min) {
    out[0] = 0
    out[1] = 0
    out[2] = 0
    return
  }
  const k = s / (max - min)
  out[0] = (r - min) * k
  out[1] = (g - min) * k
  out[2] = (b - min) * k
}

const SCRATCH = new Float64Array(3)

/** Non-separable B(Cb, Cs) into `out`, channels in [0, 1]. */
function nonSeparableInto(mode: BlendMode, br: number, bg: number, bb: number, sr: number, sg: number, sb: number, out: Float64Array): void {
  switch (mode) {
    case 'darker-color':
      if (sr + sg + sb < br + bg + bb) { out[0] = sr; out[1] = sg; out[2] = sb } else { out[0] = br; out[1] = bg; out[2] = bb }
      return
    case 'lighter-color':
      if (sr + sg + sb > br + bg + bb) { out[0] = sr; out[1] = sg; out[2] = sb } else { out[0] = br; out[1] = bg; out[2] = bb }
      return
    case 'hue': {
      const backdropSat = Math.max(br, bg, bb) - Math.min(br, bg, bb)
      setSatInto(sr, sg, sb, backdropSat, SCRATCH)
      setLumInto(SCRATCH[0], SCRATCH[1], SCRATCH[2], lumOf(br, bg, bb), out)
      return
    }
    case 'saturation': {
      const sourceSat = Math.max(sr, sg, sb) - Math.min(sr, sg, sb)
      setSatInto(br, bg, bb, sourceSat, SCRATCH)
      setLumInto(SCRATCH[0], SCRATCH[1], SCRATCH[2], lumOf(br, bg, bb), out)
      return
    }
    case 'color':
      setLumInto(sr, sg, sb, lumOf(br, bg, bb), out)
      return
    case 'luminosity':
      setLumInto(br, bg, bb, lumOf(sr, sg, sb), out)
      return
    default:
      out[0] = blendChannel(mode, br, sr)
      out[1] = blendChannel(mode, bg, sg)
      out[2] = blendChannel(mode, bb, sb)
  }
}

/** B(Cb, Cs) for whole colours (every mode; separable modes apply per channel). Channels in [0, 1]. */
export function blendColor(mode: BlendMode, backdrop: readonly [number, number, number], source: readonly [number, number, number]): [number, number, number] {
  const out = new Float64Array(3)
  nonSeparableInto(mode, backdrop[0], backdrop[1], backdrop[2], source[0], source[1], source[2], out)
  return [out[0], out[1], out[2]]
}

// ---------------------------------------------------------------------------------------------
// Pixel compositing
// ---------------------------------------------------------------------------------------------

const TABLES = new Map<BlendMode, Float32Array>()

/** B * 255 for every 8-bit (Cb, Cs) pair, index (Cb << 8) | Cs. */
function separableTable(mode: BlendMode): Float32Array {
  let table = TABLES.get(mode)
  if (!table) {
    table = new Float32Array(65536)
    for (let cb = 0; cb < 256; cb += 1) {
      for (let cs = 0; cs < 256; cs += 1) table[(cb << 8) | cs] = blendChannel(mode, cb / 255, cs / 255) * 255
    }
    TABLES.set(mode, table)
  }
  return table
}

function checkRange(name: string, index: number, count: number, length: number, stride: number): void {
  if (!Number.isInteger(index) || index < 0 || index % stride !== 0 || index + count * stride > length) {
    throw new RangeError(`blendInto: ${name} ${index} for ${count} pixels is outside the buffer (length ${length}${stride > 1 ? '; indices are byte offsets, multiples of 4' : ''}).`)
  }
}

/**
 * Composites `pixels` consecutive straight-RGBA8 pixels of `src` (from byte offset srcIndex) onto `dst` (from
 * byte offset dstIndex) in place. Offsets are array indices (multiples of 4); `coverage` (0..255, one byte per
 * pixel, from coverageIndex) multiplies the source alpha, as does `opacity` (0..1). Dissolve keeps pixel n
 * when hash2(dstIndex / 4 + n, 0, seed) < as, so callers fold the tile or row position into `seed` for a
 * pattern that stays put while scrolling.
 */
export function blendInto(
  dst: Uint8ClampedArray,
  dstIndex: number,
  src: Uint8ClampedArray,
  srcIndex: number,
  pixels: number,
  mode: BlendMode,
  opacity: number,
  coverage: Uint8Array | null = null,
  coverageIndex = 0,
  seed = 0,
): void {
  if (!(pixels > 0)) return
  const amount = opacity >= 1 ? 1 : opacity > 0 ? opacity : 0
  if (amount === 0) return
  checkRange('dstIndex', dstIndex, pixels, dst.length, 4)
  checkRange('srcIndex', srcIndex, pixels, src.length, 4)
  if (coverage) checkRange('coverageIndex', coverageIndex, pixels, coverage.length, 1)
  if (!isBlendMode(mode)) throw new RangeError(`Unknown blend mode "${String(mode)}".`)
  const kOpacity = amount / 255
  if (mode === 'normal' || mode === 'dissolve') {
    blendNormal(dst, dstIndex, src, srcIndex, pixels, kOpacity, coverage, coverageIndex, mode === 'dissolve', seed)
  } else if (NON_SEPARABLE.has(mode)) {
    blendNonSeparable(dst, dstIndex, src, srcIndex, pixels, mode, kOpacity, coverage, coverageIndex)
  } else {
    blendSeparable(dst, dstIndex, src, srcIndex, pixels, separableTable(mode), kOpacity, coverage, coverageIndex)
  }
}

// The three loops below share one shape: skip invisible source pixels, copy over an empty backdrop, use the
// simplified formula over an opaque backdrop (ao = 1), and the full W3C formula otherwise.

function blendNormal(
  dst: Uint8ClampedArray, dstIndex: number, src: Uint8ClampedArray, srcIndex: number, pixels: number,
  kOpacity: number, coverage: Uint8Array | null, coverageIndex: number, dissolve: boolean, seed: number,
): void {
  const pixelBase = dstIndex >> 2
  for (let n = 0; n < pixels; n += 1) {
    const s = srcIndex + n * 4
    const sourceAlpha = src[s + 3]
    if (sourceAlpha === 0) continue
    let as = sourceAlpha * kOpacity
    if (coverage) {
      const c = coverage[coverageIndex + n]
      if (c === 0) continue
      if (c !== 255) as *= c / 255
    }
    if (dissolve) {
      if (!(hash2(pixelBase + n, 0, seed) < as)) continue
      as = 1
    }
    const d = dstIndex + n * 4
    const backdropAlpha = dst[d + 3]
    if (as >= 1 || backdropAlpha === 0) {
      dst[d] = src[s]
      dst[d + 1] = src[s + 1]
      dst[d + 2] = src[s + 2]
      dst[d + 3] = as >= 1 ? 255 : as * 255
      continue
    }
    if (backdropAlpha === 255) {
      const k = 1 - as
      dst[d] = src[s] * as + dst[d] * k
      dst[d + 1] = src[s + 1] * as + dst[d + 1] * k
      dst[d + 2] = src[s + 2] * as + dst[d + 2] * k
      continue
    }
    const ab = backdropAlpha / 255
    const ao = as + ab - as * ab
    const wSource = as / ao
    const wBackdrop = ((1 - as) * ab) / ao
    dst[d] = src[s] * wSource + dst[d] * wBackdrop
    dst[d + 1] = src[s + 1] * wSource + dst[d + 1] * wBackdrop
    dst[d + 2] = src[s + 2] * wSource + dst[d + 2] * wBackdrop
    dst[d + 3] = ao * 255
  }
}

function blendSeparable(
  dst: Uint8ClampedArray, dstIndex: number, src: Uint8ClampedArray, srcIndex: number, pixels: number,
  table: Float32Array, kOpacity: number, coverage: Uint8Array | null, coverageIndex: number,
): void {
  for (let n = 0; n < pixels; n += 1) {
    const s = srcIndex + n * 4
    const sourceAlpha = src[s + 3]
    if (sourceAlpha === 0) continue
    let as = sourceAlpha * kOpacity
    if (coverage) {
      const c = coverage[coverageIndex + n]
      if (c === 0) continue
      if (c !== 255) as *= c / 255
    }
    const d = dstIndex + n * 4
    const backdropAlpha = dst[d + 3]
    const cs0 = src[s]
    const cs1 = src[s + 1]
    const cs2 = src[s + 2]
    if (backdropAlpha === 0) {
      dst[d] = cs0
      dst[d + 1] = cs1
      dst[d + 2] = cs2
      dst[d + 3] = as * 255
      continue
    }
    const cb0 = dst[d]
    const cb1 = dst[d + 1]
    const cb2 = dst[d + 2]
    const b0 = table[(cb0 << 8) | cs0]
    const b1 = table[(cb1 << 8) | cs1]
    const b2 = table[(cb2 << 8) | cs2]
    if (backdropAlpha === 255) {
      const k = 1 - as
      dst[d] = b0 * as + cb0 * k
      dst[d + 1] = b1 * as + cb1 * k
      dst[d + 2] = b2 * as + cb2 * k
      continue
    }
    const ab = backdropAlpha / 255
    const ao = as + ab - as * ab
    const wSource = (as * (1 - ab)) / ao
    const wBlend = (as * ab) / ao
    const wBackdrop = ((1 - as) * ab) / ao
    dst[d] = wSource * cs0 + wBlend * b0 + wBackdrop * cb0
    dst[d + 1] = wSource * cs1 + wBlend * b1 + wBackdrop * cb1
    dst[d + 2] = wSource * cs2 + wBlend * b2 + wBackdrop * cb2
    dst[d + 3] = ao * 255
  }
}

function blendNonSeparable(
  dst: Uint8ClampedArray, dstIndex: number, src: Uint8ClampedArray, srcIndex: number, pixels: number,
  mode: BlendMode, kOpacity: number, coverage: Uint8Array | null, coverageIndex: number,
): void {
  const mixed = new Float64Array(3)
  for (let n = 0; n < pixels; n += 1) {
    const s = srcIndex + n * 4
    const sourceAlpha = src[s + 3]
    if (sourceAlpha === 0) continue
    let as = sourceAlpha * kOpacity
    if (coverage) {
      const c = coverage[coverageIndex + n]
      if (c === 0) continue
      if (c !== 255) as *= c / 255
    }
    const d = dstIndex + n * 4
    const backdropAlpha = dst[d + 3]
    const cs0 = src[s]
    const cs1 = src[s + 1]
    const cs2 = src[s + 2]
    if (backdropAlpha === 0) {
      dst[d] = cs0
      dst[d + 1] = cs1
      dst[d + 2] = cs2
      dst[d + 3] = as * 255
      continue
    }
    const cb0 = dst[d]
    const cb1 = dst[d + 1]
    const cb2 = dst[d + 2]
    nonSeparableInto(mode, cb0 / 255, cb1 / 255, cb2 / 255, cs0 / 255, cs1 / 255, cs2 / 255, mixed)
    const b0 = mixed[0] * 255
    const b1 = mixed[1] * 255
    const b2 = mixed[2] * 255
    if (backdropAlpha === 255) {
      const k = 1 - as
      dst[d] = b0 * as + cb0 * k
      dst[d + 1] = b1 * as + cb1 * k
      dst[d + 2] = b2 * as + cb2 * k
      continue
    }
    const ab = backdropAlpha / 255
    const ao = as + ab - as * ab
    const wSource = (as * (1 - ab)) / ao
    const wBlend = (as * ab) / ao
    const wBackdrop = ((1 - as) * ab) / ao
    dst[d] = wSource * cs0 + wBlend * b0 + wBackdrop * cb0
    dst[d + 1] = wSource * cs1 + wBlend * b1 + wBackdrop * cb1
    dst[d + 2] = wSource * cs2 + wBlend * b2 + wBackdrop * cb2
    dst[d + 3] = ao * 255
  }
}
