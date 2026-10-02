// src/advanced/psdMapping.ts (WP7)
// Pure, DOM-free mapping between ag-psd's structures and the Advanced document model (design 6.2):
//   - file structure: sniffPsd() (header), readPsdStructure() (merged-alpha flag and ICC profile, which
//     ag-psd does not expose), psdHeaderProblem() (friendly refusals: PSB, CMYK/Lab, bit depth, size);
//   - layer properties: blend modes (all 27 Photoshop modes), opacity quantisation, locks, advanced blending;
//   - adjustment layers: the 13 Photoshop adjustments both ways (channel mixer, selective color and color
//     lookup have no equivalent and map to null);
//   - text: Photoshop point/paragraph text <-> TextSpec (font, size, colour, leading, tracking, alignment,
//     transform). Point text is anchored at the first baseline (left, centre or right by justification),
//     TextSpec at the top-left of the first line box; the conversion uses the same layout as rendering;
//   - pixels: 16/32-bit to 8-bit, mask channels, the merged image's white matte, vector-mask rasterization.
// ag-psd is only referenced for types here, so Node tests and the editor can load this module cheaply.
import type {
  AdjustmentSpec,
  Affine,
  BlendMode,
  ColorBalanceShift,
  CurvePoint,
  GradientStop,
  HslShift,
  HueRange,
  LevelsChannel,
  MaskBuffer,
  Point,
  Rgb8,
} from '../imaging/types.ts'
import type { LayerLocks, TextSpec, TextStyle } from './types.ts'
import type {
  AdjustmentLayer as PsdAdjustment,
  BezierPath,
  BlendMode as PsdBlendMode,
  Color as PsdColor,
  CurvesAdjustmentChannel,
  HueSaturationAdjustmentChannel,
  Justification,
  Layer as PsdLayer,
  LayerTextData,
  LayerVectorMask,
  LevelsAdjustmentChannel,
  ParagraphStyle as PsdParagraphStyle,
  TextStyle as PsdTextStyle,
} from 'ag-psd'
import { combineMasks, createMaskBuffer, invertMask, rasterizePolygon } from '../imaging/mask.ts'
import { linearToSrgb } from '../imaging/color.ts'
import { layoutText } from '../shared/vector.ts'
import type { TextMeasure } from '../shared/vector.ts'

// ---------------------------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------------------------

function finite(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value) ? value : fallback
}

function clamp(value: number, min: number, max: number): number {
  return value < min ? min : value > max ? max : value
}

function clamp01(value: number): number {
  return value <= 0 ? 0 : value >= 1 ? 1 : value
}

function byte(value: unknown, fallback = 0): number {
  return clamp(Math.round(finite(value, fallback)), 0, 255)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

/** Photoshop stores layer opacity as one byte: round(opacity * 255) / 255 (0.5 -> 0.50196). */
export function quantizeOpacity(opacity: number): number {
  return Math.round(clamp01(finite(opacity, 1)) * 255) / 255
}

/** a ∘ b (apply b first, then a), both in canvas order [a, b, c, d, e, f]. */
export function multiplyAffine(a: Affine, b: Affine): Affine {
  return [
    a[0] * b[0] + a[2] * b[1],
    a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3],
    a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4],
    a[1] * b[4] + a[3] * b[5] + a[5],
  ]
}

function translation(x: number, y: number): Affine {
  return [1, 0, 0, 1, x, y]
}

function affineFrom(value: unknown): Affine {
  if (Array.isArray(value) && value.length >= 6 && value.slice(0, 6).every((n) => typeof n === 'number' && Number.isFinite(n))) {
    return [value[0], value[1], value[2], value[3], value[4], value[5]]
  }
  return [1, 0, 0, 1, 0, 0]
}

// ---------------------------------------------------------------------------------------------
// File header and structure (the main process checked these too; the renderer re-checks every input)
// ---------------------------------------------------------------------------------------------

export interface PsdHeader {
  /** 1 = PSD, 2 = PSB (large document). */
  readonly version: number
  readonly channels: number
  readonly width: number
  readonly height: number
  /** Bits per channel: 1, 8, 16 or 32. */
  readonly depth: number
  /** 0 Bitmap, 1 Grayscale, 2 Indexed, 3 RGB, 4 CMYK, 7 Multichannel, 8 Duotone, 9 Lab. */
  readonly colorMode: number
}

export const PSD_COLOR_MODES: Readonly<Record<number, string>> = Object.freeze({
  0: 'Bitmap', 1: 'Grayscale', 2: 'Indexed', 3: 'RGB', 4: 'CMYK', 7: 'Multichannel', 8: 'Duotone', 9: 'Lab',
})

/** The fixed 26-byte header of a PSD/PSB, or null when `bytes` do not start with '8BPS'. */
export function sniffPsd(bytes: Uint8Array): PsdHeader | null {
  if (!bytes || bytes.length < 26) return null
  if (bytes[0] !== 0x38 || bytes[1] !== 0x42 || bytes[2] !== 0x50 || bytes[3] !== 0x53) return null
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return {
    version: view.getUint16(4),
    channels: view.getUint16(12),
    height: view.getUint32(14),
    width: view.getUint32(18),
    depth: view.getUint16(22),
    colorMode: view.getUint16(24),
  }
}

export const PSD_TOO_LARGE_MESSAGE = 'This image is too large to edit safely. The limit is 50 megapixels and 20,000 pixels per side.'
export const PSD_DAMAGED_MESSAGE = 'This Photoshop document is damaged or uses a layout Simple can\'t read.'

/** A friendly reason Simple cannot open a document with this header, or null when it can. */
export function psdHeaderProblem(header: PsdHeader, limits: { readonly maxPixels: number; readonly maxDimension: number }): string | null {
  if (header.version === 2) return 'Large Photoshop documents (.psb) are not supported. Save the document as .psd in Photoshop first.'
  if (header.version !== 1) return PSD_DAMAGED_MESSAGE
  if (![0, 1, 2, 3].includes(header.colorMode)) {
    return `This PSD uses ${PSD_COLOR_MODES[header.colorMode] || 'an unsupported'} color. In Photoshop choose Image > Mode > RGB Color, then save a copy.`
  }
  if (![1, 8, 16, 32].includes(header.depth)) return 'This PSD uses an unsupported bit depth.'
  if ((header.colorMode === 0) !== (header.depth === 1)) return PSD_DAMAGED_MESSAGE
  if (header.colorMode === 2 && header.depth !== 8) return 'This PSD uses an unsupported bit depth.'
  if (header.channels < 1 || header.channels > 56) return PSD_DAMAGED_MESSAGE
  if (header.channels > 16) {
    return 'This Photoshop document has more than 16 channels. Delete extra alpha channels in Photoshop (Channels panel), then save a copy.'
  }
  if (header.width < 1 || header.height < 1) return PSD_DAMAGED_MESSAGE
  if (header.width > limits.maxDimension || header.height > limits.maxDimension || header.width * header.height > limits.maxPixels) {
    return PSD_TOO_LARGE_MESSAGE
  }
  return null
}

export interface PsdStructure {
  /**
   * True when the layer count is negative: the merged image's first extra channel is its transparency
   * (otherwise extra channels are saved selections and the merged image is opaque).
   */
  readonly mergedAlpha: boolean
  /** Embedded ICC profile (image resource 1039), or null. */
  readonly iccProfile: Uint8Array | null
}

/** Reads what ag-psd does not expose: the merged-alpha flag and the ICC profile. Tolerates damaged data. */
export function readPsdStructure(bytes: Uint8Array): PsdStructure {
  const result = { mergedAlpha: false, iccProfile: null as Uint8Array | null }
  if (!sniffPsd(bytes)) return result
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  const length = bytes.length
  const u32 = (offset: number) => (offset + 4 <= length ? view.getUint32(offset) : -1)
  let offset = 26
  const colorModeLength = u32(offset)
  if (colorModeLength < 0) return result
  offset += 4 + colorModeLength
  const resourcesLength = u32(offset)
  if (resourcesLength < 0) return result
  const resourcesEnd = Math.min(length, offset + 4 + resourcesLength)
  let at = offset + 4
  while (at + 12 <= resourcesEnd) {
    if (bytes[at] !== 0x38 || bytes[at + 1] !== 0x42 || bytes[at + 2] !== 0x49 || bytes[at + 3] !== 0x4d) break // '8BIM'
    const id = view.getUint16(at + 4)
    const nameLength = bytes[at + 6]
    // Pascal name: length byte + text, padded to an even size.
    const cursor = at + 6 + ((nameLength + 2) & ~1)
    const size = u32(cursor)
    if (size < 0) break
    const dataStart = cursor + 4
    const dataEnd = dataStart + size
    if (dataEnd > resourcesEnd) break
    if (id === 1039 && size > 132) result.iccProfile = bytes.subarray(dataStart, dataEnd)
    at = dataEnd + (size % 2)
  }
  offset += 4 + resourcesLength
  const layerAndMaskLength = u32(offset)
  if (layerAndMaskLength > 0) {
    const layerInfoLength = u32(offset + 4)
    if (layerInfoLength >= 2 && offset + 10 <= length) result.mergedAlpha = view.getInt16(offset + 8) < 0
  }
  return result
}

/** The 'desc' text of an ICC profile (v2 textDescriptionType or v4 multiLocalizedUnicodeType), or null. */
export function iccProfileDescription(profile: Uint8Array | null): string | null {
  if (!profile || profile.length < 132) return null
  const view = new DataView(profile.buffer, profile.byteOffset, profile.byteLength)
  const count = view.getUint32(128)
  for (let index = 0; index < Math.min(count, 256); index += 1) {
    const entry = 132 + index * 12
    if (entry + 12 > profile.length) break
    const signature = String.fromCharCode(profile[entry], profile[entry + 1], profile[entry + 2], profile[entry + 3])
    if (signature !== 'desc') continue
    const start = view.getUint32(entry + 4)
    const size = view.getUint32(entry + 8)
    if (start + Math.min(size, 12) > profile.length || size < 12) return null
    const type = String.fromCharCode(profile[start], profile[start + 1], profile[start + 2], profile[start + 3])
    if (type === 'desc') {
      const textLength = view.getUint32(start + 8)
      const end = Math.min(profile.length, start + 12 + textLength)
      let text = ''
      for (let i = start + 12; i < end && profile[i] !== 0; i += 1) text += String.fromCharCode(profile[i])
      return text.trim() || null
    }
    if (type === 'mluc' && start + 28 <= profile.length) {
      const records = view.getUint32(start + 8)
      if (!records) return null
      const textLength = view.getUint32(start + 20)
      const textOffset = view.getUint32(start + 24)
      let text = ''
      for (let i = start + textOffset; i + 1 < Math.min(profile.length, start + textOffset + textLength); i += 2) {
        const code = view.getUint16(i)
        if (!code) break
        text += String.fromCharCode(code)
      }
      return text.trim() || null
    }
    return null
  }
  return null
}

/** True for sRGB profiles ("sRGB IEC61966-2.1", "sRGB built-in", ...): Simple works in sRGB. */
export function isSrgbProfile(profile: Uint8Array | null): boolean {
  if (!profile) return true
  const description = iccProfileDescription(profile)
  return description ? /\bs\s*rgb\b/i.test(description) : false
}

/** Pixels per inch from resolutionInfo (Photoshop stores pixels per inch whatever the display unit is). */
export function ppiFromResolution(info: { readonly horizontalResolution?: number } | undefined | null): number {
  const value = finite(info?.horizontalResolution, 72)
  return value > 0 && value < 100_000 ? Math.round(value * 1000) / 1000 : 72
}

export function resolutionFromPpi(ppi: number): {
  horizontalResolution: number; horizontalResolutionUnit: 'PPI'; widthUnit: 'Inches'
  verticalResolution: number; verticalResolutionUnit: 'PPI'; heightUnit: 'Inches'
} {
  const value = finite(ppi, 72) > 0 ? finite(ppi, 72) : 72
  return {
    horizontalResolution: value, horizontalResolutionUnit: 'PPI', widthUnit: 'Inches',
    verticalResolution: value, verticalResolutionUnit: 'PPI', heightUnit: 'Inches',
  }
}

// ---------------------------------------------------------------------------------------------
// Blend modes
// ---------------------------------------------------------------------------------------------

const BLEND_FROM_PSD: Readonly<Record<string, BlendMode>> = Object.freeze({
  normal: 'normal', dissolve: 'dissolve',
  darken: 'darken', multiply: 'multiply', 'color burn': 'color-burn', 'linear burn': 'linear-burn', 'darker color': 'darker-color',
  lighten: 'lighten', screen: 'screen', 'color dodge': 'color-dodge', 'linear dodge': 'linear-dodge', 'lighter color': 'lighter-color',
  overlay: 'overlay', 'soft light': 'soft-light', 'hard light': 'hard-light', 'vivid light': 'vivid-light',
  'linear light': 'linear-light', 'pin light': 'pin-light', 'hard mix': 'hard-mix',
  difference: 'difference', exclusion: 'exclusion', subtract: 'subtract', divide: 'divide',
  hue: 'hue', saturation: 'saturation', color: 'color', luminosity: 'luminosity',
})

const BLEND_TO_PSD: Readonly<Record<BlendMode, PsdBlendMode>> = Object.freeze(Object.fromEntries(
  Object.entries(BLEND_FROM_PSD).map(([psd, mode]) => [mode, psd]),
) as Record<BlendMode, PsdBlendMode>)

/**
 * A layer's blend mode. Unknown modes become Normal with supported = false. 'pass through' (a group mode)
 * maps to Normal and is supported: groups are dissolved or flattened by the importer.
 */
export function blendFromPsd(mode: string | undefined): { mode: BlendMode; supported: boolean } {
  if (mode === undefined || mode === 'pass through') return { mode: 'normal', supported: true }
  const mapped = BLEND_FROM_PSD[mode]
  return mapped ? { mode: mapped, supported: true } : { mode: 'normal', supported: false }
}

/** ag-psd's name for a blend mode ('color-burn' -> 'color burn'). */
export function blendToPsd(mode: BlendMode): string {
  return BLEND_TO_PSD[mode] ?? 'normal'
}

/** Photoshop's "special eight": fill opacity and layer opacity differ for these modes. */
export const FILL_OPACITY_SENSITIVE_MODES: ReadonlySet<BlendMode> = new Set<BlendMode>([
  'color-burn', 'linear-burn', 'color-dodge', 'linear-dodge', 'vivid-light', 'linear-light', 'hard-mix', 'difference',
])

// ---------------------------------------------------------------------------------------------
// Locks and advanced blending
// ---------------------------------------------------------------------------------------------

export function locksFromPsd(layer: Pick<PsdLayer, 'transparencyProtected' | 'protected'>): LayerLocks {
  const locks = layer.protected ?? {}
  return {
    pixels: Boolean(locks.composite),
    position: Boolean(locks.position),
    transparency: Boolean(locks.transparency || layer.transparencyProtected),
  }
}

export function locksToPsd(locks: LayerLocks): Pick<PsdLayer, 'transparencyProtected' | 'protected'> {
  return {
    transparencyProtected: Boolean(locks.transparency),
    protected: { transparency: Boolean(locks.transparency), composite: Boolean(locks.pixels), position: Boolean(locks.position) },
  }
}

function isDefaultRange(range: readonly number[] | undefined): boolean {
  return !range || (range.length >= 4 && range[0] === 0 && range[1] === 0 && range[2] === 255 && range[3] === 255)
}

/** Photoshop "Blending Options" Simple does not reproduce (Blend If, knockout, channel restrictions, ...). */
export function advancedBlendingNotes(layer: PsdLayer): string[] {
  const notes: string[] = []
  const ranges = layer.blendingRanges
  if (ranges && (!isDefaultRange(ranges.compositeGrayBlendSource) || !isDefaultRange(ranges.compositeGraphBlendDestinationRange)
    || ranges.ranges.some((range) => !isDefaultRange(range.sourceRange) || !isDefaultRange(range.destRange)))) notes.push('Blend If')
  if (layer.knockout) notes.push('knockout')
  if (layer.blendClippendElements === false) notes.push('clipped layers blended separately')
  if (layer.blendInteriorElements) notes.push('interior effects blended as a group')
  if (layer.channelBlendingRestrictions && layer.channelBlendingRestrictions.length) notes.push('channel restrictions')
  if (layer.transparencyShapesLayer === false) notes.push('transparency does not shape the layer')
  if (layer.layerMaskAsGlobalMask) notes.push('mask hides effects')
  return notes
}

/** True when the layer has at least one enabled layer style (effects that Simple imports without). */
export function hasEnabledEffects(layer: PsdLayer): boolean {
  const effects = layer.effects
  if (!effects || effects.disabled) return false
  for (const [key, value] of Object.entries(effects)) {
    if (key === 'disabled' || key === 'scale') continue
    const list = Array.isArray(value) ? value : [value]
    for (const effect of list) if (isRecord(effect) && effect.enabled !== false && effect.present !== false) return true
  }
  return false
}

// ---------------------------------------------------------------------------------------------
// Colours
// ---------------------------------------------------------------------------------------------

/**
 * Where a colour came from: ag-psd uses different scales for binary colour records (RGB/CMYK 0..255,
 * HSB 0..1, Lab normalised), descriptors (RGB 0..255, percentages, Lab in L*a*b* units) and text engine
 * data (RGB 0..255).
 */
export type PsdColorSource = 'binary' | 'descriptor' | 'engine'

function labToRgb(l: number, a: number, b: number): Rgb8 {
  // CIE L*a*b* (D50, Photoshop's PCS) -> XYZ -> Bradford D50->D65 -> linear sRGB -> sRGB.
  const fy = (l + 16) / 116
  const fx = fy + a / 500
  const fz = fy - b / 200
  const inv = (t: number) => (t > 6 / 29 ? t * t * t : 3 * (6 / 29) * (6 / 29) * (t - 4 / 29))
  const x50 = 0.96422 * inv(fx)
  const y50 = inv(fy)
  const z50 = 0.82521 * inv(fz)
  const x = 0.9555766 * x50 - 0.0230393 * y50 + 0.0631636 * z50
  const y = -0.0282895 * x50 + 1.0099416 * y50 + 0.0210077 * z50
  const z = 0.0122982 * x50 - 0.020483 * y50 + 1.3299098 * z50
  const r = 3.2404542 * x - 1.5371385 * y - 0.4985314 * z
  const g = -0.969266 * x + 1.8760108 * y + 0.041556 * z
  const bl = 0.0556434 * x - 0.2040259 * y + 1.0572252 * z
  return { r: Math.round(linearToSrgb(r) * 255), g: Math.round(linearToSrgb(g) * 255), b: Math.round(linearToSrgb(bl) * 255) }
}

function hsbToRgb(h: number, s: number, v: number): Rgb8 {
  const hue = ((h % 360) + 360) % 360 / 60
  const c = v * s
  const x = c * (1 - Math.abs((hue % 2) - 1))
  const [r, g, b] = hue < 1 ? [c, x, 0] : hue < 2 ? [x, c, 0] : hue < 3 ? [0, c, x] : hue < 4 ? [0, x, c] : hue < 5 ? [x, 0, c] : [c, 0, x]
  const m = v - c
  return { r: byte((r + m) * 255), g: byte((g + m) * 255), b: byte((b + m) * 255) }
}

/** Any ag-psd colour as 8-bit sRGB (Lab, HSB, CMYK and gray are converted), or null when absent. */
export function psdColorToRgb(color: PsdColor | undefined | null, source: PsdColorSource): Rgb8 | null {
  if (!isRecord(color)) return null
  const c = color as Record<string, number>
  if ('r' in c && 'g' in c && 'b' in c) return { r: byte(c.r), g: byte(c.g), b: byte(c.b) }
  if ('fr' in c) return { r: byte(finite(c.fr, 0) * 255), g: byte(finite(c.fg, 0) * 255), b: byte(finite(c.fb, 0) * 255) }
  if ('l' in c && 'a' in c && 'b' in c) {
    let l = finite(c.l, 0)
    let a = finite(c.a, 0)
    let b = finite(c.b, 0)
    if (source === 'binary' && Math.abs(l) <= 1 && Math.abs(a) <= 1 && Math.abs(b) <= 1) {
      l *= 100
      a *= 128
      b *= 128
    }
    return labToRgb(l, a, b)
  }
  if ('h' in c && 's' in c && 'b' in c) {
    const h = source === 'binary' ? finite(c.h, 0) * 360 : finite(c.h, 0) <= 1 ? finite(c.h, 0) * 360 : finite(c.h, 0)
    const s = finite(c.s, 0) > 1 ? finite(c.s, 0) / 100 : finite(c.s, 0)
    const v = finite(c.b, 0) > 1 ? finite(c.b, 0) / 100 : finite(c.b, 0)
    return hsbToRgb(h, clamp01(s), clamp01(v))
  }
  if ('c' in c && 'm' in c && 'y' in c && 'k' in c) {
    const scale = source === 'descriptor' ? 100 : 255
    const cc = clamp01(finite(c.c, 0) / scale)
    const mm = clamp01(finite(c.m, 0) / scale)
    const yy = clamp01(finite(c.y, 0) / scale)
    const kk = clamp01(finite(c.k, 0) / scale)
    return { r: byte(255 * (1 - cc) * (1 - kk)), g: byte(255 * (1 - mm) * (1 - kk)), b: byte(255 * (1 - yy) * (1 - kk)) }
  }
  if ('k' in c) {
    const k = source === 'descriptor' ? 255 - finite(c.k, 0) * 2.55 : finite(c.k, 0)
    const value = byte(k)
    return { r: value, g: value, b: value }
  }
  return null
}

function rgbOut(color: Rgb8): { r: number; g: number; b: number } {
  return { r: byte(color.r), g: byte(color.g), b: byte(color.b) }
}

// ---------------------------------------------------------------------------------------------
// Adjustment layers
// ---------------------------------------------------------------------------------------------

const HUE_RANGES: readonly HueRange[] = ['reds', 'yellows', 'greens', 'cyans', 'blues', 'magentas']

/** Photoshop's default Hue/Saturation range windows (degrees: ramp start, range start, range end, ramp end). */
const HUE_RANGE_WINDOWS: Readonly<Record<HueRange, readonly [number, number, number, number]>> = Object.freeze({
  reds: [315, 345, 15, 45],
  yellows: [15, 45, 75, 105],
  greens: [75, 105, 135, 165],
  cyans: [135, 165, 195, 225],
  blues: [195, 225, 255, 285],
  magentas: [255, 285, 315, 345],
})

function levelsFromPsd(channel: LevelsAdjustmentChannel | undefined): LevelsChannel {
  if (!channel) return { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 }
  return {
    inBlack: byte(channel.shadowInput, 0),
    inWhite: byte(channel.highlightInput, 255),
    gamma: clamp(finite(channel.midtoneInput, 1), 0.01, 9.99),
    outBlack: byte(channel.shadowOutput, 0),
    outWhite: byte(channel.highlightOutput, 255),
  }
}

function levelsToPsd(channel: LevelsChannel): LevelsAdjustmentChannel {
  return {
    shadowInput: byte(channel.inBlack, 0),
    highlightInput: byte(channel.inWhite, 255),
    shadowOutput: byte(channel.outBlack, 0),
    highlightOutput: byte(channel.outWhite, 255),
    midtoneInput: Math.round(clamp(finite(channel.gamma, 1), 0.01, 9.99) * 100) / 100,
  }
}

function curveFromPsd(channel: CurvesAdjustmentChannel | undefined): CurvePoint[] {
  const points = (channel ?? [])
    .map((point) => ({ x: byte(point?.input), y: byte(point?.output) }))
    .sort((a, b) => a.x - b.x)
  const unique: CurvePoint[] = []
  for (const point of points) {
    if (unique.length && unique[unique.length - 1].x === point.x) unique[unique.length - 1] = point
    else unique.push(point)
  }
  if (unique.length < 2) return [{ x: 0, y: 0 }, { x: 255, y: 255 }]
  return unique.slice(0, 16)
}

function curveToPsd(points: readonly CurvePoint[]): CurvesAdjustmentChannel {
  const clean = curveFromPsd((points ?? []).map((point) => ({ input: point.x, output: point.y })))
  return clean.map((point) => ({ input: point.x, output: point.y }))
}

function shiftFromPsd(channel: HueSaturationAdjustmentChannel | undefined): HslShift {
  return {
    hue: clamp(Math.round(finite(channel?.hue, 0)), -180, 180),
    saturation: clamp(Math.round(finite(channel?.saturation, 0)), -100, 100),
    lightness: clamp(Math.round(finite(channel?.lightness, 0)), -100, 100),
  }
}

function balanceFromPsd(value: { cyanRed?: number; magentaGreen?: number; yellowBlue?: number } | undefined): ColorBalanceShift {
  return {
    cyanRed: clamp(Math.round(finite(value?.cyanRed, 0)), -100, 100),
    magentaGreen: clamp(Math.round(finite(value?.magentaGreen, 0)), -100, 100),
    yellowBlue: clamp(Math.round(finite(value?.yellowBlue, 0)), -100, 100),
  }
}

function balanceToPsd(value: ColorBalanceShift): { cyanRed: number; magentaGreen: number; yellowBlue: number } {
  return balanceFromPsd(value)
}

/**
 * Gradient stops. Photoshop keeps each midpoint on the stop that ends the segment (the midpoint between
 * the previous stop and this one); GradientStop keeps it on the stop that starts the segment.
 */
function gradientStopsFromPsd(stops: readonly { color: PsdColor; location: number; midpoint: number }[]): GradientStop[] {
  const sorted = stops
    .map((stop) => ({ position: clamp01(finite(stop.location, 0)), color: psdColorToRgb(stop.color, 'binary') ?? { r: 0, g: 0, b: 0 }, midpoint: finite(stop.midpoint, 0.5) }))
    .sort((a, b) => a.position - b.position)
  // The last stop starts no segment, so it carries no midpoint.
  return sorted.map((stop, index) => (index + 1 < sorted.length
    ? { position: stop.position, color: stop.color, midpoint: clamp(sorted[index + 1].midpoint, 0.05, 0.95) }
    : { position: stop.position, color: stop.color }))
}

function gradientStopsToPsd(stops: readonly GradientStop[]): { color: PsdColor; location: number; midpoint: number }[] {
  const sorted = [...stops].sort((a, b) => a.position - b.position)
  return sorted.map((stop, index) => ({
    color: rgbOut(stop.color),
    location: clamp01(finite(stop.position, 0)),
    midpoint: index === 0 ? 0.5 : clamp(finite(sorted[index - 1].midpoint, 0.5), 0.05, 0.95),
  }))
}

/** An ag-psd adjustment as an AdjustmentSpec; null for adjustments Simple cannot apply. */
export function adjustmentFromPsd(input: unknown): AdjustmentSpec | null {
  if (!isRecord(input) || typeof input.type !== 'string') return null
  const a = input as unknown as PsdAdjustment
  switch (a.type) {
    case 'brightness/contrast':
      return {
        type: 'brightness-contrast',
        brightness: clamp(Math.round(finite(a.brightness, 0)), -150, 150),
        contrast: clamp(Math.round(finite(a.contrast, 0)), -50, 100),
        legacy: Boolean(a.useLegacy),
      }
    case 'levels':
      return { type: 'levels', rgb: levelsFromPsd(a.rgb), red: levelsFromPsd(a.red), green: levelsFromPsd(a.green), blue: levelsFromPsd(a.blue) }
    case 'curves':
      return { type: 'curves', rgb: curveFromPsd(a.rgb), red: curveFromPsd(a.red), green: curveFromPsd(a.green), blue: curveFromPsd(a.blue) }
    case 'exposure':
      return {
        type: 'exposure',
        exposure: clamp(finite(a.exposure, 0), -20, 20),
        offset: clamp(finite(a.offset, 0), -0.5, 0.5),
        gamma: clamp(finite(a.gamma, 1), 0.01, 9.99),
      }
    case 'vibrance':
      return { type: 'vibrance', vibrance: clamp(Math.round(finite(a.vibrance, 0)), -100, 100), saturation: clamp(Math.round(finite(a.saturation, 0)), -100, 100) }
    case 'hue/saturation': {
      // ag-psd reads the colorize flag (high byte of `a`) and the colorization values (b, c, d) into `master`.
      const master = a.master
      const colorize = Boolean(master && (finite(master.a, 0) & 0xff00))
      if (colorize && master) {
        return {
          type: 'hue-saturation',
          colorize: true,
          master: {
            hue: clamp(Math.round(finite(master.b, 0)), 0, 360),
            saturation: clamp(Math.round(finite(master.c, 25)), 0, 100),
            lightness: clamp(Math.round(finite(master.d, 0)), -100, 100),
          },
          ranges: {},
        }
      }
      const ranges: Partial<Record<HueRange, HslShift>> = {}
      for (const range of HUE_RANGES) {
        const shift = shiftFromPsd(a[range])
        if (shift.hue || shift.saturation || shift.lightness) ranges[range] = shift
      }
      return { type: 'hue-saturation', colorize: false, master: shiftFromPsd(master), ranges }
    }
    case 'color balance':
      return {
        type: 'color-balance',
        shadows: balanceFromPsd(a.shadows),
        midtones: balanceFromPsd(a.midtones),
        highlights: balanceFromPsd(a.highlights),
        preserveLuminosity: a.preserveLuminosity !== false,
      }
    case 'black & white': {
      const tint = a.useTint ? psdColorToRgb(a.tintColor, 'descriptor') : null
      const value = (v: number | undefined, fallback: number) => clamp(Math.round(finite(v, fallback)), -200, 300)
      return {
        type: 'black-white',
        reds: value(a.reds, 40),
        yellows: value(a.yellows, 60),
        greens: value(a.greens, 40),
        cyans: value(a.cyans, 60),
        blues: value(a.blues, 20),
        magentas: value(a.magentas, 80),
        tint,
      }
    }
    case 'photo filter':
      return {
        type: 'photo-filter',
        color: psdColorToRgb(a.color, 'binary') ?? { r: 236, g: 138, b: 0 },
        density: clamp(Math.round(finite(a.density, 25)), 0, 100),
        preserveLuminosity: a.preserveLuminosity !== false,
      }
    case 'invert':
      return { type: 'invert' }
    case 'posterize':
      return { type: 'posterize', levels: clamp(Math.round(finite(a.levels, 4)), 2, 255) }
    case 'threshold':
      return { type: 'threshold', level: clamp(Math.round(finite(a.level, 128)), 1, 255) }
    case 'gradient map': {
      if (a.gradientType === 'noise') return null
      const stops = gradientStopsFromPsd(a.colorStops ?? [])
      return {
        type: 'gradient-map',
        stops: stops.length >= 2 ? stops : [
          { position: 0, color: { r: 0, g: 0, b: 0 }, midpoint: 0.5 },
          { position: 1, color: { r: 255, g: 255, b: 255 }, midpoint: 0.5 },
        ],
        reverse: Boolean(a.reverse),
        dither: Boolean(a.dither),
      }
    }
    default:
      return null
  }
}

/** Human name of an ag-psd adjustment type for messages ("Channel Mixer"). */
export function psdAdjustmentLabel(input: unknown): string {
  const type = isRecord(input) && typeof input.type === 'string' ? input.type : 'adjustment'
  return type.split(/[ /&]+/).filter(Boolean).map((word) => word[0].toUpperCase() + word.slice(1)).join(' ')
}

/** An AdjustmentSpec as an ag-psd adjustment; null when Photoshop has no equivalent (Simple's quick adjust). */
export function adjustmentToPsd(spec: AdjustmentSpec): unknown {
  switch (spec.type) {
    case 'brightness-contrast':
      return {
        type: 'brightness/contrast',
        brightness: clamp(Math.round(finite(spec.brightness, 0)), -150, 150),
        contrast: clamp(Math.round(finite(spec.contrast, 0)), -50, 100),
        meanValue: 127,
        useLegacy: Boolean(spec.legacy),
        labColorOnly: false,
      } satisfies PsdAdjustment
    case 'levels':
      return { type: 'levels', rgb: levelsToPsd(spec.rgb), red: levelsToPsd(spec.red), green: levelsToPsd(spec.green), blue: levelsToPsd(spec.blue) } satisfies PsdAdjustment
    case 'curves':
      return { type: 'curves', rgb: curveToPsd(spec.rgb), red: curveToPsd(spec.red), green: curveToPsd(spec.green), blue: curveToPsd(spec.blue) } satisfies PsdAdjustment
    case 'exposure':
      return {
        type: 'exposure',
        exposure: clamp(finite(spec.exposure, 0), -20, 20),
        offset: clamp(finite(spec.offset, 0), -0.5, 0.5),
        gamma: clamp(finite(spec.gamma, 1), 0.01, 9.99),
      } satisfies PsdAdjustment
    case 'vibrance':
      return { type: 'vibrance', vibrance: clamp(Math.round(finite(spec.vibrance, 0)), -100, 100), saturation: clamp(Math.round(finite(spec.saturation, 0)), -100, 100) } satisfies PsdAdjustment
    case 'hue-saturation': {
      const channel = (range: HueRange): HueSaturationAdjustmentChannel => {
        const shift = spec.colorize ? undefined : spec.ranges?.[range]
        const [a, b, c, d] = HUE_RANGE_WINDOWS[range]
        return {
          a, b, c, d,
          hue: clamp(Math.round(finite(shift?.hue, 0)), -180, 180),
          saturation: clamp(Math.round(finite(shift?.saturation, 0)), -100, 100),
          lightness: clamp(Math.round(finite(shift?.lightness, 0)), -100, 100),
        }
      }
      const master: HueSaturationAdjustmentChannel = spec.colorize
        ? {
          a: 0x0100, // colorize flag in the high byte, then the colorization hue/saturation/lightness
          b: clamp(Math.round(finite(spec.master.hue, 0)), 0, 360),
          c: clamp(Math.round(finite(spec.master.saturation, 25)), 0, 100),
          d: clamp(Math.round(finite(spec.master.lightness, 0)), -100, 100),
          hue: 0,
          saturation: 0,
          lightness: 0,
        }
        : {
          a: 0, b: 0, c: 25, d: 0,
          hue: clamp(Math.round(finite(spec.master.hue, 0)), -180, 180),
          saturation: clamp(Math.round(finite(spec.master.saturation, 0)), -100, 100),
          lightness: clamp(Math.round(finite(spec.master.lightness, 0)), -100, 100),
        }
      return {
        type: 'hue/saturation',
        master,
        reds: channel('reds'),
        yellows: channel('yellows'),
        greens: channel('greens'),
        cyans: channel('cyans'),
        blues: channel('blues'),
        magentas: channel('magentas'),
      } satisfies PsdAdjustment
    }
    case 'color-balance':
      return {
        type: 'color balance',
        shadows: balanceToPsd(spec.shadows),
        midtones: balanceToPsd(spec.midtones),
        highlights: balanceToPsd(spec.highlights),
        preserveLuminosity: Boolean(spec.preserveLuminosity),
      } satisfies PsdAdjustment
    case 'black-white':
      return {
        type: 'black & white',
        reds: Math.round(finite(spec.reds, 40)),
        yellows: Math.round(finite(spec.yellows, 60)),
        greens: Math.round(finite(spec.greens, 40)),
        cyans: Math.round(finite(spec.cyans, 60)),
        blues: Math.round(finite(spec.blues, 20)),
        magentas: Math.round(finite(spec.magentas, 80)),
        useTint: Boolean(spec.tint),
        tintColor: rgbOut(spec.tint ?? { r: 225, g: 211, b: 179 }),
      } satisfies PsdAdjustment
    case 'photo-filter':
      return {
        type: 'photo filter',
        color: rgbOut(spec.color),
        density: clamp(Math.round(finite(spec.density, 25)), 0, 100),
        preserveLuminosity: Boolean(spec.preserveLuminosity),
      } satisfies PsdAdjustment
    case 'invert':
      return { type: 'invert' } satisfies PsdAdjustment
    case 'posterize':
      return { type: 'posterize', levels: clamp(Math.round(finite(spec.levels, 4)), 2, 255) } satisfies PsdAdjustment
    case 'threshold':
      return { type: 'threshold', level: clamp(Math.round(finite(spec.level, 128)), 1, 255) } satisfies PsdAdjustment
    case 'gradient-map':
      return {
        type: 'gradient map',
        name: 'Custom',
        gradientType: 'solid',
        smoothness: 1,
        colorStops: gradientStopsToPsd(spec.stops),
        opacityStops: [{ opacity: 1, location: 0, midpoint: 0.5 }, { opacity: 1, location: 1, midpoint: 0.5 }],
        reverse: Boolean(spec.reverse),
        dither: Boolean(spec.dither),
      } satisfies PsdAdjustment
    default:
      return null
  }
}

// ---------------------------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------------------------

export interface ResolvedFont {
  /** CSS family, e.g. 'Arial'. */
  readonly family: string
  readonly weight: 400 | 700
  readonly italic: boolean
}

export type FontResolver = (postScriptName: string) => ResolvedFont

export interface PsdFontChoice {
  /** PostScript name Photoshop looks the font up by, e.g. 'Arial-BoldMT'. */
  readonly postScriptName: string
  /** Ask Photoshop to embolden / slant a regular face (used when the exact face is unknown). */
  readonly fauxBold: boolean
  readonly fauxItalic: boolean
}

const DEFAULT_PSD_FONT = 'ArialMT'

function alignFromJustification(justification: Justification | undefined): TextStyle['align'] {
  if (justification === 'center' || justification === 'justify-center') return 'center'
  if (justification === 'right' || justification === 'justify-right') return 'right'
  return 'left'
}

function styleKey(style: PsdTextStyle | undefined): string {
  if (!style) return ''
  const color = psdColorToRgb(style.fillColor, 'engine')
  return JSON.stringify([style.font?.name ?? '', finite(style.fontSize, 0), Boolean(style.fauxBold), Boolean(style.fauxItalic),
    color ? [color.r, color.g, color.b] : null, Boolean(style.underline), finite(style.tracking, 0)])
}

/** Photoshop's insertion point of point text (text-local x) for the line layout's alignment. */
function pointAnchor(align: TextStyle['align'], width: number): number {
  return align === 'center' ? width / 2 : align === 'right' ? width : 0
}

/**
 * Photoshop text -> TextSpec. `notes` lists what Simple draws differently once the text is edited (until
 * then the layer shows Photoshop's own pixels). `measure` must be the one used for rendering.
 */
export function textFromPsd(text: LayerTextData, resolveFont: FontResolver, measure?: TextMeasure): { spec: TextSpec; notes: string[] } {
  const notes: string[] = []
  const runs = text.styleRuns ?? []
  const style: PsdTextStyle = { ...(text.style ?? {}), ...(runs[0]?.style ?? {}) }
  if (runs.length > 1 && new Set(runs.map((run) => styleKey({ ...(text.style ?? {}), ...run.style }))).size > 1) {
    notes.push('mixed character styles use the first style')
  }
  const paragraph: PsdParagraphStyle = { ...(text.paragraphStyle ?? {}), ...(text.paragraphStyleRuns?.[0]?.style ?? {}) }
  if (text.warp && text.warp.style && text.warp.style !== 'none') notes.push('warped text is drawn straight')
  if (text.orientation === 'vertical') notes.push('vertical text is drawn horizontally')
  if (text.textPath) notes.push('text on a path is drawn straight')
  const font = resolveFont(String(style.font?.name || DEFAULT_PSD_FONT))
  const fontSize = finite(style.fontSize, 12) > 0 ? finite(style.fontSize, 12) : 12
  const leading = finite(style.leading, 0)
  const lineHeight = style.autoLeading === false && leading > 0
    ? leading / fontSize
    : clamp(finite(paragraph.autoLeading, 1.2), 0.5, 10)
  const textStyle: TextStyle = {
    fontFamily: font.family,
    fontSize,
    fontWeight: style.fauxBold ? 700 : font.weight,
    italic: Boolean(style.fauxItalic) || font.italic,
    underline: Boolean(style.underline),
    color: psdColorToRgb(style.fillColor, 'engine') ?? { r: 0, g: 0, b: 0 },
    align: alignFromJustification(paragraph.justification),
    lineHeight,
    letterSpacing: (finite(style.tracking, 0) * fontSize) / 1000,
  }
  const content = String(text.text ?? '').replace(/\r\n?/g, '\n')
  let matrix = affineFrom(text.transform)
  const hScale = finite(style.horizontalScale, 1)
  const vScale = finite(style.verticalScale, 1)
  if ((hScale !== 1 || vScale !== 1) && hScale > 0 && vScale > 0) matrix = multiplyAffine(matrix, [hScale, 0, 0, vScale, 0, 0])
  const box = text.boxBounds
  if (text.shapeType === 'box' && Array.isArray(box) && box.length >= 4 && finite(box[2], 0) > finite(box[0], 0)) {
    const left = finite(box[0], 0)
    const top = finite(box[1], 0)
    return {
      spec: { text: content, style: textStyle, boxWidth: finite(box[2], 0) - left, transform: multiplyAffine(matrix, translation(left, top)) },
      notes,
    }
  }
  const layout = layoutText({ text: content, style: textStyle, boxWidth: null }, measure)
  const baseline = layout.lines[0]?.baseline ?? layout.ascent
  const anchorX = pointAnchor(textStyle.align, layout.width)
  const base = Array.isArray(text.pointBase) && text.pointBase.length >= 2 ? text.pointBase : [0, 0]
  const transform = multiplyAffine(matrix, translation(finite(base[0], 0) - anchorX, finite(base[1], 0) - baseline))
  return { spec: { text: content, style: textStyle, boxWidth: null, transform }, notes }
}

/** TextSpec -> Photoshop text data (point text anchored at its first baseline, or paragraph text). */
export function textToPsd(spec: TextSpec, font: PsdFontChoice, measure?: TextMeasure): LayerTextData {
  const style = spec.style
  const fontSize = finite(style.fontSize, 12) > 0 ? finite(style.fontSize, 12) : 12
  const layout = layoutText(spec, measure)
  const matrix = affineFrom(spec.transform)
  const lineHeight = finite(style.lineHeight, 1.2)
  const autoLeading = Math.abs(lineHeight - 1.2) < 1e-9
  const psdStyle: PsdTextStyle = {
    font: { name: font.postScriptName || DEFAULT_PSD_FONT },
    fontSize,
    fauxBold: font.fauxBold,
    fauxItalic: font.fauxItalic,
    autoLeading,
    ...(autoLeading ? {} : { leading: lineHeight * fontSize }),
    tracking: Math.round((finite(style.letterSpacing, 0) / fontSize) * 1000),
    underline: Boolean(style.underline),
    fillColor: rgbOut(style.color),
  }
  const common = {
    text: spec.text,
    antiAlias: 'smooth' as const,
    orientation: 'horizontal' as const,
    style: psdStyle,
    paragraphStyle: { justification: style.align, autoLeading: 1.2 },
  }
  const boxWidth = spec.boxWidth
  if (boxWidth !== null && boxWidth !== undefined && Number.isFinite(boxWidth) && boxWidth > 0) {
    return {
      ...common,
      transform: [...matrix],
      shapeType: 'box',
      boxBounds: [0, 0, boxWidth, layout.height],
      left: 0,
      top: 0,
      right: boxWidth,
      bottom: layout.height,
    }
  }
  const baseline = layout.lines[0]?.baseline ?? layout.ascent
  const anchorX = pointAnchor(style.align, layout.width)
  return {
    ...common,
    transform: [...multiplyAffine(matrix, translation(anchorX, baseline))],
    shapeType: 'point',
    pointBase: [0, 0],
    left: -anchorX,
    top: -baseline,
    right: layout.width - anchorX,
    bottom: layout.height - baseline,
  }
}

// ---------------------------------------------------------------------------------------------
// Pixels
// ---------------------------------------------------------------------------------------------

let u16ToU8: Uint8Array | null = null

/** Rounded 16-bit -> 8-bit table (exact inverse of v * 257). */
function u16Table(): Uint8Array {
  if (!u16ToU8) {
    u16ToU8 = new Uint8Array(65536)
    for (let value = 0; value < 65536; value += 1) u16ToU8[value] = Math.floor((value * 255 + 32767) / 65535)
  }
  return u16ToU8
}

export type PsdSampleArray = Uint8ClampedArray | Uint8Array | Uint16Array | Float32Array

/**
 * ag-psd RGBA samples of any depth as straight 8-bit RGBA. 8-bit data is returned as is (no copy);
 * 16-bit is rounded; 32-bit (linear light) is sRGB-encoded and clamped, alpha linearly.
 */
export function toRgba8(data: PsdSampleArray, pixels: number): Uint8ClampedArray {
  const length = pixels * 4
  if (data instanceof Uint8ClampedArray) return data.length === length ? data : data.subarray(0, length)
  if (data instanceof Uint8Array) return new Uint8ClampedArray(data.buffer, data.byteOffset, Math.min(length, data.length))
  const out = new Uint8ClampedArray(length)
  if (data instanceof Uint16Array) {
    const table = u16Table()
    for (let i = 0; i < length; i += 1) out[i] = table[data[i]]
    return out
  }
  for (let i = 0; i < length; i += 4) {
    out[i] = Math.round(linearToSrgb(data[i]) * 255)
    out[i + 1] = Math.round(linearToSrgb(data[i + 1]) * 255)
    out[i + 2] = Math.round(linearToSrgb(data[i + 2]) * 255)
    out[i + 3] = Math.round(clamp01(data[i + 3]) * 255)
  }
  return out
}

/** The gray (red) channel of ag-psd's RGBA mask data of any depth, as 8-bit coverage. */
export function maskChannel(data: PsdSampleArray, width: number, height: number): MaskBuffer {
  const count = width * height
  const out = new Uint8Array(count)
  if (data instanceof Uint16Array) {
    const table = u16Table()
    for (let p = 0; p < count; p += 1) out[p] = table[data[p * 4]]
  } else if (data instanceof Float32Array) {
    for (let p = 0; p < count; p += 1) out[p] = Math.round(clamp01(data[p * 4]) * 255)
  } else {
    for (let p = 0; p < count; p += 1) out[p] = data[p * 4]
  }
  return { width, height, data: out }
}

/** Photoshop mattes the merged image of a transparent document with white; undo that in place. */
export function removeWhiteMatte(data: Uint8ClampedArray): void {
  for (let i = 0; i < data.length; i += 4) {
    const alpha = data[i + 3]
    if (alpha === 0 || alpha === 255) continue
    const a = alpha / 255
    const matte = 255 * (1 - a)
    data[i] = (data[i] - matte) / a
    data[i + 1] = (data[i + 1] - matte) / a
    data[i + 2] = (data[i + 2] - matte) / a
  }
}

/** Sets every alpha to 255 in place (a merged image whose extra channels are saved selections). */
export function forceOpaque(data: Uint8ClampedArray): void {
  for (let i = 3; i < data.length; i += 4) data[i] = 255
}

// ---------------------------------------------------------------------------------------------
// Vector masks
// ---------------------------------------------------------------------------------------------

function cubicPoints(p0: Point, p1: Point, p2: Point, p3: Point, out: Point[]): void {
  const straight = p1.x === p0.x && p1.y === p0.y && p2.x === p3.x && p2.y === p3.y
  if (straight) {
    out.push(p3)
    return
  }
  const span = Math.hypot(p1.x - p0.x, p1.y - p0.y) + Math.hypot(p2.x - p1.x, p2.y - p1.y) + Math.hypot(p3.x - p2.x, p3.y - p2.y)
  const steps = clamp(Math.ceil(span / 3), 4, 96)
  for (let step = 1; step <= steps; step += 1) {
    const t = step / steps
    const u = 1 - t
    const a = u * u * u
    const b = 3 * u * u * t
    const c = 3 * u * t * t
    const d = t * t * t
    out.push({ x: a * p0.x + b * p1.x + c * p2.x + d * p3.x, y: a * p0.y + b * p1.y + c * p2.y + d * p3.y })
  }
}

/** A closed Bézier path (knot points: [in.x, in.y, anchor.x, anchor.y, out.x, out.y]) as a polygon. */
export function flattenBezierPath(path: Pick<BezierPath, 'knots'>): Point[] {
  const knots = (path.knots ?? []).filter((knot) => Array.isArray(knot?.points) && knot.points.length >= 6 && knot.points.every(Number.isFinite))
  if (knots.length < 2) return []
  const anchor = (index: number): Point => ({ x: knots[index].points[2], y: knots[index].points[3] })
  const points: Point[] = [anchor(0)]
  for (let index = 0; index < knots.length; index += 1) {
    const next = (index + 1) % knots.length
    const current = knots[index].points
    const following = knots[next].points
    cubicPoints(anchor(index), { x: current[4], y: current[5] }, { x: following[0], y: following[1] }, anchor(next), points)
  }
  return points
}

/**
 * Rasterizes a vector mask over the document (width x height): path components combine in order
 * (combine = add, subtract, intersect, exclude = xor) starting from empty, or full when
 * fillStartsWithAllPixels; then inverted when asked. Null when the mask is disabled or draws nothing.
 */
export function rasterizeVectorMask(vector: LayerVectorMask | undefined, width: number, height: number): MaskBuffer | null {
  if (!vector || vector.disable || !(width > 0 && height > 0)) return null
  const result = createMaskBuffer(width, height, vector.fillStartsWithAllPixels ? 255 : 0)
  let drew = Boolean(vector.fillStartsWithAllPixels)
  for (const path of vector.paths ?? []) {
    if (path.open) continue
    const polygon = flattenBezierPath(path)
    if (polygon.length < 3) continue
    const shape = createMaskBuffer(width, height, 0)
    const bounds = rasterizePolygon(shape, polygon, true, path.fillRule === 'non-zero' ? 'nonzero' : 'evenodd')
    const operation = path.operation ?? 'combine'
    if (operation === 'exclude') {
      const a = result.data
      const b = shape.data
      for (let i = 0; i < a.length; i += 1) if (b[i]) a[i] = Math.abs(a[i] - b[i])
    } else {
      combineMasks(result, shape, operation === 'subtract' ? 'subtract' : operation === 'intersect' ? 'intersect' : 'add', bounds ?? { x: 0, y: 0, width: 0, height: 0 })
    }
    drew = true
  }
  if (!drew) return null
  if (vector.invert) invertMask(result)
  return result
}

/** a = a * b / 255 for two coverage masks of the same size (both masks hide what either hides). */
export function multiplyMasks(a: MaskBuffer, b: MaskBuffer): void {
  combineMasks(a, b, 'intersect')
}
