// src/simple/encodeHints.ts (WP8)
// How an edited image should be re-encoded so Save does not quietly lower its quality (design 4.9,
// IMAGE-SIE-14). Read from the source file's bytes in the renderer (no pixel decode):
//   - WebP: a lossless source (VP8L bitstream) must be saved losslessly. Chromium's encoder writes VP8L only
//     at quality exactly 1 (0.92 and 0.999 give lossy VP8 + ALPH, about 10x larger for graphics).
//   - JPEG: the source quality estimated from its luminance quantization table (IJG scaling), so a 95%
//     photo is not saved at 92% and a 75% one is not inflated. Clamped to 0.75..0.98 for saving.
// Pure and DOM-free (Node-tested).

/** IJG standard luminance quantization table (natural order). */
const STD_LUMINANCE: readonly number[] = [
  16, 11, 10, 16, 24, 40, 51, 61,
  12, 12, 14, 19, 26, 58, 60, 55,
  14, 13, 16, 24, 40, 57, 69, 56,
  14, 17, 22, 29, 51, 87, 80, 62,
  18, 22, 37, 56, 68, 109, 103, 77,
  24, 35, 55, 64, 81, 104, 113, 92,
  49, 64, 78, 87, 103, 121, 120, 101,
  72, 92, 95, 98, 112, 100, 103, 99,
]

/** Zigzag position -> natural index (DQT stores tables in zigzag order). */
export const JPEG_NATURAL_ORDER: readonly number[] = [
  0, 1, 8, 16, 9, 2, 3, 10, 17, 24, 32, 25, 18, 11, 4, 5,
  12, 19, 26, 33, 40, 48, 41, 34, 27, 20, 13, 6, 7, 14, 21, 28,
  35, 42, 49, 56, 57, 50, 43, 36, 29, 22, 15, 23, 30, 37, 44, 51,
  58, 59, 52, 45, 38, 31, 39, 46, 53, 60, 61, 54, 47, 55, 62, 63,
]

export const DEFAULT_LOSSY_QUALITY = 0.92
export const MIN_SAVE_QUALITY = 0.75
export const MAX_SAVE_QUALITY = 0.98

export interface EncodeHints {
  /** The source WebP was lossless: re-encode at quality 1 (VP8L). */
  readonly webpLossless: boolean
  /** Estimated source JPEG quality 1..100, or null when unknown. */
  readonly jpegQuality: number | null
}

export const NO_ENCODE_HINTS: EncodeHints = Object.freeze({ webpLossless: false, jpegQuality: null })

function ascii(bytes: Uint8Array, offset: number, length: number): string {
  let text = ''
  for (let index = 0; index < length && offset + index < bytes.length; index += 1) text += String.fromCharCode(bytes[offset + index])
  return text
}

function u32le(bytes: Uint8Array, offset: number): number {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0
}

/** True when the WebP's image bitstream is VP8L (lossless): first chunk VP8L, or VP8X followed by VP8L before any VP8. */
export function isLosslessWebp(bytes: Uint8Array): boolean {
  if (bytes.length < 20 || ascii(bytes, 0, 4) !== 'RIFF' || ascii(bytes, 8, 4) !== 'WEBP') return false
  let offset = 12
  for (let guard = 0; guard < 4096 && offset + 8 <= bytes.length; guard += 1) {
    const type = ascii(bytes, offset, 4)
    const size = u32le(bytes, offset + 4)
    if (type === 'VP8L') return true
    if (type === 'VP8 ') return false
    if (type === 'ANMF') {
      // Animated: judge the first frame's bitstream (frame header is 16 bytes, then its chunks).
      const frameEnd = Math.min(bytes.length, offset + 8 + size)
      let inner = offset + 8 + 16
      while (inner + 8 <= frameEnd) {
        const innerType = ascii(bytes, inner, 4)
        if (innerType === 'VP8L') return true
        if (innerType === 'VP8 ') return false
        inner += 8 + u32le(bytes, inner + 4) + (u32le(bytes, inner + 4) & 1)
      }
      return false
    }
    if (size > bytes.length) return false
    offset += 8 + size + (size & 1)
  }
  return false
}

/** The first 8-bit luminance (id 0) quantization table in natural order, or null. */
export function jpegLuminanceTable(bytes: Uint8Array): number[] | null {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return null
  let offset = 2
  while (offset + 4 <= bytes.length) {
    if (bytes[offset] !== 0xff) return null
    let marker = bytes[offset + 1]
    // Fill bytes.
    while (marker === 0xff && offset + 2 < bytes.length) {
      offset += 1
      marker = bytes[offset + 1]
    }
    if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
      offset += 2
      continue
    }
    if (marker === 0xd9 || marker === 0xda) return null
    const length = (bytes[offset + 2] << 8) | bytes[offset + 3]
    if (length < 2) return null
    const end = offset + 2 + length
    if (marker === 0xdb) {
      let p = offset + 4
      while (p < end && p < bytes.length) {
        const precision = bytes[p] >> 4
        const id = bytes[p] & 0x0f
        const entryBytes = precision ? 2 : 1
        if (p + 1 + 64 * entryBytes > bytes.length) return null
        if (id === 0) {
          const table = new Array<number>(64)
          for (let k = 0; k < 64; k += 1) {
            const value = precision ? (bytes[p + 1 + 2 * k] << 8) | bytes[p + 2 + 2 * k] : bytes[p + 1 + k]
            table[JPEG_NATURAL_ORDER[k]] = value
          }
          return table
        }
        p += 1 + 64 * entryBytes
      }
    }
    offset = end
  }
  return null
}

/** The table libjpeg builds for `quality` (1..100) from the standard one (baseline: entries 1..255). */
export function scaledLuminanceTable(quality: number): number[] {
  const q = Math.max(1, Math.min(100, Math.round(quality)))
  const scale = q < 50 ? Math.floor(5000 / q) : 200 - q * 2
  return STD_LUMINANCE.map((value) => Math.max(1, Math.min(255, Math.floor((value * scale + 50) / 100))))
}

/** Estimated JPEG quality (1..100) from the luminance table, or null when the file has none. */
export function estimateJpegQuality(bytes: Uint8Array): number | null {
  const table = jpegLuminanceTable(bytes)
  if (!table) return null
  let best = 0
  let bestError = Infinity
  for (let quality = 1; quality <= 100; quality += 1) {
    const expected = scaledLuminanceTable(quality)
    let error = 0
    for (let index = 0; index < 64; index += 1) error += Math.abs(expected[index] - table[index])
    if (error < bestError || (error === bestError && quality > best)) {
      bestError = error
      best = quality
    }
  }
  return best || null
}

/** Encoder quality for `format` given the source hints (undefined = the encoder's default, e.g. PNG). */
export function saveQuality(format: string, hints: EncodeHints): number | undefined {
  if (format === 'png') return undefined
  if (format === 'webp') return hints.webpLossless ? 1 : DEFAULT_LOSSY_QUALITY
  if (format === 'jpeg') {
    if (hints.jpegQuality === null) return DEFAULT_LOSSY_QUALITY
    return Math.max(MIN_SAVE_QUALITY, Math.min(MAX_SAVE_QUALITY, hints.jpegQuality / 100))
  }
  return DEFAULT_LOSSY_QUALITY
}

/** Hints for a freshly opened source. */
export function encodeHintsFor(format: string, bytes: Uint8Array): EncodeHints {
  const normalized = format.toLowerCase()
  if (normalized === 'webp') return { webpLossless: isLosslessWebp(bytes), jpegQuality: null }
  if (normalized === 'jpg' || normalized === 'jpeg') return { webpLossless: false, jpegQuality: estimateJpegQuality(bytes) }
  return NO_ENCODE_HINTS
}
