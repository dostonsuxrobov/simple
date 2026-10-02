// src/imaging/color.ts (WP2)
// Colour math shared by adjustments, blending, looks and the colour UI. Pure and DOM-free.
// Conventions: RGB triples passed as numbers are in [0, 1] unless a name says 8 (0..255);
// hue is in degrees [0, 360); HSL saturation and lightness are in [0, 1].
import type { Rgb8 } from './types.ts'

function clamp01(value: number): number {
  return value <= 0 ? 0 : value >= 1 ? 1 : value
}

/** sRGB transfer: encoded [0, 1] -> linear light [0, 1]. */
export function srgbToLinear(encoded: number): number {
  const c = clamp01(encoded)
  return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4)
}

/** sRGB transfer: linear light [0, 1] -> encoded [0, 1] (not rounded). */
export function linearToSrgb(linear: number): number {
  const v = clamp01(linear)
  return v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055
}

/** Linear light of every 8-bit sRGB code value (256 entries). */
export const SRGB_TO_LINEAR: Float32Array = (() => {
  const table = new Float32Array(256)
  for (let value = 0; value < 256; value += 1) table[value] = srgbToLinear(value / 255)
  return table
})()

/** Linear light [0, 1] (clamped) -> rounded 8-bit sRGB code value. Exact inverse of SRGB_TO_LINEAR. */
export function linearToSrgb8(value: number): number {
  return Math.round(linearToSrgb(value) * 255)
}

/** RGB in [0, 1] -> [hue degrees [0, 360), saturation [0, 1], lightness [0, 1]] (HSL). */
export function rgbToHsl(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const l = (max + min) / 2
  const d = max - min
  if (d <= 0) return [0, 0, l]
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min)
  let h: number
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0)
  else if (max === g) h = (b - r) / d + 2
  else h = (r - g) / d + 4
  h *= 60
  if (h >= 360) h -= 360
  return [h, s, l]
}

function hueChannel(p: number, q: number, t: number): number {
  if (t < 0) t += 1
  else if (t > 1) t -= 1
  if (t < 1 / 6) return p + (q - p) * 6 * t
  if (t < 1 / 2) return q
  if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
  return p
}

/** HSL (hue in degrees, any value; s and l in [0, 1]) -> RGB in [0, 1]. */
export function hslToRgb(h: number, s: number, l: number): [number, number, number] {
  const sat = clamp01(s)
  const light = clamp01(l)
  if (sat <= 0) return [light, light, light]
  let hue = (h % 360) / 360
  if (hue < 0) hue += 1
  const q = light < 0.5 ? light * (1 + sat) : light + sat - light * sat
  const p = 2 * light - q
  return [hueChannel(p, q, hue + 1 / 3), hueChannel(p, q, hue), hueChannel(p, q, hue - 1 / 3)]
}

/** RGB in [0, 1] -> [hue degrees [0, 360), saturation [0, 1], value [0, 1]] (HSB/HSV, for colour pickers). */
export function rgbToHsv(r: number, g: number, b: number): [number, number, number] {
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  const d = max - min
  if (d <= 0) return [0, 0, max]
  let h: number
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0)
  else if (max === g) h = (b - r) / d + 2
  else h = (r - g) / d + 4
  h *= 60
  if (h >= 360) h -= 360
  return [h, max > 0 ? d / max : 0, max]
}

/** HSB/HSV (hue in degrees; s and v in [0, 1]) -> RGB in [0, 1]. */
export function hsvToRgb(h: number, s: number, v: number): [number, number, number] {
  const sat = clamp01(s)
  const value = clamp01(v)
  let hue = h % 360
  if (hue < 0) hue += 360
  const sector = hue / 60
  const i = Math.floor(sector) % 6
  const f = sector - Math.floor(sector)
  const p = value * (1 - sat)
  const q = value * (1 - sat * f)
  const t = value * (1 - sat * (1 - f))
  switch (i) {
    case 0: return [value, t, p]
    case 1: return [q, value, p]
    case 2: return [p, value, t]
    case 3: return [p, q, value]
    case 4: return [t, p, value]
    default: return [value, p, q]
  }
}

/** W3C Compositing Lum(C) = 0.3 R + 0.59 G + 0.11 B (any consistent channel scale). */
export function lum(r: number, g: number, b: number): number {
  return 0.3 * r + 0.59 * g + 0.11 * b
}

/** W3C Sat(C) = max - min. */
export function sat(r: number, g: number, b: number): number {
  return Math.max(r, g, b) - Math.min(r, g, b)
}

/** W3C ClipColor: pulls an out-of-gamut colour towards its luminosity until it fits [0, 1]. */
export function clipColor(rgb: readonly [number, number, number]): [number, number, number] {
  let [r, g, b] = rgb
  const l = lum(r, g, b)
  const n = Math.min(r, g, b)
  const x = Math.max(r, g, b)
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
  return [clamp01(r), clamp01(g), clamp01(b)]
}

/** W3C SetLum: shifts the colour to luminosity `l` (in [0, 1]) and clips it into gamut. */
export function setLum(rgb: [number, number, number], l: number): [number, number, number] {
  const d = l - lum(rgb[0], rgb[1], rgb[2])
  return clipColor([rgb[0] + d, rgb[1] + d, rgb[2] + d])
}

/** W3C SetSat: rescales the colour so max - min equals `s`, keeping the order of the channels. */
export function setSat(rgb: [number, number, number], s: number): [number, number, number] {
  const [r, g, b] = rgb
  const max = Math.max(r, g, b)
  const min = Math.min(r, g, b)
  if (max <= min) return [0, 0, 0]
  const scale = s / (max - min)
  return [(r - min) * scale, (g - min) * scale, (b - min) * scale]
}

/** '#rgb', '#rrggbb' (with or without '#', any case, surrounding spaces ignored) -> Rgb8, else null. */
export function parseHex(hex: string): Rgb8 | null {
  if (typeof hex !== 'string') return null
  let text = hex.trim()
  if (text.startsWith('#')) text = text.slice(1)
  if (!/^[0-9a-fA-F]+$/.test(text)) return null
  if (text.length === 3) text = text[0] + text[0] + text[1] + text[1] + text[2] + text[2]
  if (text.length !== 6) return null
  const value = parseInt(text, 16)
  return { r: (value >> 16) & 255, g: (value >> 8) & 255, b: value & 255 }
}

/** Rgb8 -> '#rrggbb' (lowercase, the form input[type=color] reports). Channels are rounded and clamped. */
export function toHex(color: Rgb8): string {
  const part = (value: number) => {
    const v = Math.max(0, Math.min(255, Math.round(Number(value) || 0)))
    return v.toString(16).padStart(2, '0')
  }
  return `#${part(color.r)}${part(color.g)}${part(color.b)}`
}
