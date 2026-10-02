'use strict'

// Metadata carry-over for edited saves. The renderer re-encodes pixels with canvas.toBlob, which writes no
// EXIF, XMP, IPTC, DPI or PNG text. These pure functions read those blocks from the source file's bytes and
// splice them into the newly encoded JPEG, PNG or WebP without decoding or re-encoding any pixels.
//
// Rules:
// - EXIF Orientation is reset to 1 (the renderer's pixels are already upright), PixelXDimension/
//   PixelYDimension follow the new size, and the IFD1 thumbnail (a stale picture of the old pixels) is dropped.
// - ICC: Chromium converts decoded pixels to sRGB before they reach the canvas. A source sRGB profile is kept;
//   a wide-gamut profile (Adobe RGB, Display P3, ...) is NOT re-attached by default, because tagging sRGB pixels
//   with it would shift every color. `colorProfile: 'keep'` forces it (for pixels that were never converted).
// - PNG: pHYs, tEXt, zTXt and iTXt are kept; EXIF travels as eXIf, XMP as iTXt "XML:com.adobe.xmp".
// - WebP: the file is wrapped in VP8X with ICCP, EXIF and "XMP " chunks.
// Nothing here throws for damaged metadata: a block that cannot be parsed safely is dropped and reported.

const zlib = require('node:zlib')
const { imageDimensions, sniffFormat, toBytes } = require('./image-files.cjs')

const MAX_TEXT_BYTES = 16 * 1024 * 1024
const MAX_ICC_BYTES = 16 * 1024 * 1024
const JPEG_SEGMENT_PAYLOAD = 65533
const EXIF_HEADER = Buffer.from('Exif\0\0', 'latin1')
const XMP_NAMESPACE = Buffer.from('http://ns.adobe.com/xap/1.0/\0', 'latin1')
const XMP_EXTENSION_NAMESPACE = Buffer.from('http://ns.adobe.com/xmp/extension/\0', 'latin1')
const ICC_HEADER = Buffer.from('ICC_PROFILE\0', 'latin1')
const PHOTOSHOP_HEADER = Buffer.from('Photoshop 3.0\0', 'latin1')
const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
const PNG_XMP_KEYWORD = 'XML:com.adobe.xmp'
/** Photoshop image resources worth keeping in APP13: IPTC-NAA, its digest and the resolution info. */
const KEPT_PHOTOSHOP_RESOURCES = new Set([0x0404, 0x0425, 0x03ed])

// #region primitives

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n += 1) {
    let c = n
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()

function crc32(...buffers) {
  let crc = 0xffffffff
  for (const buffer of buffers) {
    for (let index = 0; index < buffer.length; index += 1) crc = CRC_TABLE[(crc ^ buffer[index]) & 0xff] ^ (crc >>> 8)
  }
  return (crc ^ 0xffffffff) >>> 0
}

function startsWith(buffer, prefix) {
  return buffer.length >= prefix.length && buffer.subarray(0, prefix.length).equals(prefix)
}

function pngChunk(type, data) {
  const header = Buffer.alloc(8)
  header.writeUInt32BE(data.length, 0)
  header.write(type, 4, 'latin1')
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(header.subarray(4, 8), data), 0)
  return Buffer.concat([header, data, crc])
}

function readPngChunks(bytes) {
  if (!startsWith(bytes, PNG_SIGNATURE)) throw new Error('Not a PNG file.')
  const chunks = []
  let offset = 8
  while (offset + 12 <= bytes.length) {
    const length = bytes.readUInt32BE(offset)
    const type = bytes.subarray(offset + 4, offset + 8).toString('latin1')
    const end = offset + 12 + length
    if (end > bytes.length) throw new Error('The PNG chunk list is damaged.')
    chunks.push({ type, data: bytes.subarray(offset + 8, offset + 8 + length), raw: bytes.subarray(offset, end) })
    offset = end
    if (type === 'IEND') break
  }
  return chunks
}

function jpegSegment(marker, payload) {
  if (payload.length > JPEG_SEGMENT_PAYLOAD) throw new Error('A JPEG segment is too large.')
  const header = Buffer.from([0xff, marker, 0, 0])
  header.writeUInt16BE(payload.length + 2, 2)
  return Buffer.concat([header, payload])
}

/** Marker segments up to (not including) the first SOS, and where the entropy-coded remainder starts. */
function readJpegSegments(bytes) {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) throw new Error('Not a JPEG file.')
  const segments = []
  let offset = 2
  while (offset < bytes.length) {
    if (bytes[offset] !== 0xff) throw new Error('The JPEG marker list is damaged.')
    const markerStart = offset
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1
    const marker = bytes[offset]
    offset += 1
    if (marker === 0xda || marker === 0xd9) return { segments, rest: bytes.subarray(markerStart) }
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue
    if (offset + 2 > bytes.length) break
    const length = bytes.readUInt16BE(offset)
    if (length < 2 || offset + length > bytes.length) throw new Error('The JPEG marker list is damaged.')
    segments.push({ marker, data: bytes.subarray(offset + 2, offset + length) })
    offset += length
  }
  throw new Error('The JPEG has no image data.')
}

function readRiffChunks(bytes) {
  if (bytes.length < 12 || bytes.subarray(0, 4).toString('latin1') !== 'RIFF' || bytes.subarray(8, 12).toString('latin1') !== 'WEBP') {
    throw new Error('Not a WebP file.')
  }
  const chunks = []
  let offset = 12
  while (offset + 8 <= bytes.length) {
    const type = bytes.subarray(offset, offset + 4).toString('latin1')
    const size = bytes.readUInt32LE(offset + 4)
    if (offset + 8 + size > bytes.length) throw new Error('The WebP chunk list is damaged.')
    chunks.push({ type, data: bytes.subarray(offset + 8, offset + 8 + size) })
    offset += 8 + size + (size & 1)
  }
  return chunks
}

function riffChunk(type, data) {
  const header = Buffer.alloc(8)
  header.write(type, 0, 'latin1')
  header.writeUInt32LE(data.length, 4)
  return Buffer.concat(data.length & 1 ? [header, data, Buffer.alloc(1)] : [header, data])
}

function inflateCapped(data, limit) {
  return zlib.inflateSync(data, { maxOutputLength: limit })
}

// #endregion primitives

// #region EXIF (TIFF structure)

const TIFF_TYPE_SIZES = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8, 13: 4, 16: 8, 17: 8, 18: 8 }
const RESOLUTION_UNITS = { 2: 'in', 3: 'cm' }

/**
 * Reads, and optionally patches, a TIFF/EXIF block. Returns a copy where Orientation is 1, the pixel
 * dimensions are `width`/`height` (when given), and IFD1 (the embedded thumbnail) is unlinked and, when it
 * sits at the end of the block as cameras write it, cut off. Throws when the structure is unsafe to edit.
 */
function patchExif(input, { width = null, height = null } = {}) {
  const tiff = Buffer.from(toBytes(input))
  if (tiff.length < 8) throw new Error('The EXIF block is too short.')
  const little = tiff[0] === 0x49 && tiff[1] === 0x49
  if (!little && !(tiff[0] === 0x4d && tiff[1] === 0x4d)) throw new Error('The EXIF byte order is unknown.')
  const u16 = (offset) => (little ? tiff.readUInt16LE(offset) : tiff.readUInt16BE(offset))
  const u32 = (offset) => (little ? tiff.readUInt32LE(offset) : tiff.readUInt32BE(offset))
  const w16 = (offset, value) => (little ? tiff.writeUInt16LE(value, offset) : tiff.writeUInt16BE(value, offset))
  const w32 = (offset, value) => (little ? tiff.writeUInt32LE(value >>> 0, offset) : tiff.writeUInt32BE(value >>> 0, offset))
  if (u16(2) !== 42) throw new Error('The EXIF block is not TIFF data.')
  const info = { orientation: null, resolution: null, hasThumbnail: false, hasGps: false }
  const resolution = {}
  const visited = new Set()
  let maxEnd = 8

  const writeSize = (entry, type, count, value) => {
    if (!value || count !== 1) return
    if (type === 3 && value <= 0xffff) w16(entry + 8, value)
    else if (type === 4) w32(entry + 8, value)
  }
  const rational = (offset) => {
    const denominator = u32(offset + 4)
    return denominator ? u32(offset) / denominator : null
  }

  function walk(offset, kind, depth = 0) {
    if (!offset || depth > 4 || visited.has(offset)) return null
    if (offset + 2 > tiff.length) throw new Error('An EXIF directory is out of range.')
    visited.add(offset)
    const count = u16(offset)
    const end = offset + 2 + count * 12 + 4
    if (count > 1000 || end > tiff.length) throw new Error('An EXIF directory is out of range.')
    maxEnd = Math.max(maxEnd, end)
    for (let index = 0; index < count; index += 1) {
      const entry = offset + 2 + index * 12
      const tag = u16(entry)
      const type = u16(entry + 2)
      const valueCount = u32(entry + 4)
      const size = (TIFF_TYPE_SIZES[type] || 1) * valueCount
      if (size > 4) {
        const valueOffset = u32(entry + 8)
        if (valueOffset + size <= tiff.length) maxEnd = Math.max(maxEnd, valueOffset + size)
        if (kind === 'ifd0' && type === 5 && (tag === 0x011a || tag === 0x011b) && valueOffset + 8 <= tiff.length) {
          resolution[tag === 0x011a ? 'x' : 'y'] = rational(valueOffset)
        }
      }
      if (kind === 'ifd0') {
        if (tag === 0x0112 && type === 3) {
          info.orientation = u16(entry + 8)
          w16(entry + 8, 1)
        } else if (tag === 0x0128 && type === 3) {
          resolution.unit = u16(entry + 8)
        } else if (tag === 0x0100) writeSize(entry, type, valueCount, width)
        else if (tag === 0x0101) writeSize(entry, type, valueCount, height)
        else if (tag === 0x8769) walk(u32(entry + 8), 'exif', depth + 1)
        else if (tag === 0x8825) { info.hasGps = true; walk(u32(entry + 8), 'gps', depth + 1) }
      } else if (kind === 'exif') {
        if (tag === 0xa002) writeSize(entry, type, valueCount, width)
        else if (tag === 0xa003) writeSize(entry, type, valueCount, height)
        else if (tag === 0xa005) walk(u32(entry + 8), 'interop', depth + 1)
      }
    }
    return end - 4
  }

  const nextPointer = walk(u32(4), 'ifd0')
  if (nextPointer == null) throw new Error('The EXIF block has no main directory.')
  const ifd1 = u32(nextPointer)
  let result = tiff
  if (ifd1) {
    info.hasThumbnail = true
    w32(nextPointer, 0)
    if (ifd1 >= maxEnd && ifd1 <= tiff.length) result = tiff.subarray(0, ifd1)
  }
  const unit = RESOLUTION_UNITS[resolution.unit ?? 2]
  if (unit && resolution.x > 0 && resolution.y > 0) info.resolution = { x: resolution.x, y: resolution.y, unit }
  return { tiff: Buffer.from(result), info }
}

// #endregion EXIF

// #region XMP and ICC helpers

function patchXmp(input, { width = null, height = null } = {}) {
  let text = toBytes(input).toString('utf8')
  const setValue = (name, value) => {
    if (value == null) return
    const attribute = new RegExp(`(${name}\\s*=\\s*["'])[^"']*(["'])`, 'g')
    const element = new RegExp(`(<${name}>)[^<]*(</${name}>)`, 'g')
    text = text.replace(attribute, (_whole, open, close) => `${open}${value}${close}`)
      .replace(element, (_whole, open, close) => `${open}${value}${close}`)
  }
  setValue('tiff:Orientation', 1)
  setValue('exif:PixelXDimension', width)
  setValue('exif:PixelYDimension', height)
  setValue('tiff:ImageWidth', width)
  setValue('tiff:ImageLength', height)
  return Buffer.from(text, 'utf8')
}

/** Color space and description of an ICC profile, and whether it is an sRGB profile. */
function iccInfo(icc) {
  const info = { colorSpace: null, description: '', srgb: false }
  try {
    if (icc.length < 132 || icc.subarray(36, 40).toString('latin1') !== 'acsp') return info
    info.colorSpace = icc.subarray(16, 20).toString('latin1').trim()
    const count = Math.min(icc.readUInt32BE(128), 200)
    for (let index = 0; index < count; index += 1) {
      const entry = 132 + index * 12
      if (entry + 12 > icc.length) break
      if (icc.subarray(entry, entry + 4).toString('latin1') !== 'desc') continue
      const offset = icc.readUInt32BE(entry + 4)
      const size = icc.readUInt32BE(entry + 8)
      if (offset + Math.max(size, 16) > icc.length) break
      const type = icc.subarray(offset, offset + 4).toString('latin1')
      if (type === 'desc') {
        const length = icc.readUInt32BE(offset + 8)
        info.description = icc.subarray(offset + 12, Math.min(icc.length, offset + 12 + length)).toString('latin1').replace(/\0+$/, '')
      } else if (type === 'mluc' && offset + 28 <= icc.length) {
        const length = icc.readUInt32BE(offset + 20)
        const start = offset + icc.readUInt32BE(offset + 24)
        const text = icc.subarray(start, Math.min(icc.length, start + length))
        info.description = new TextDecoder('utf-16be').decode(text).replace(/\0+$/, '')
      }
      break
    }
  } catch { /* an unreadable profile is treated as not sRGB */ }
  info.srgb = info.colorSpace === 'RGB' && /sRGB|IEC\s*61966-2[.-]1/i.test(info.description)
  return info
}

function filterPhotoshopResources(irb) {
  const kept = []
  let offset = 0
  while (offset + 12 <= irb.length && irb.subarray(offset, offset + 4).toString('latin1') === '8BIM') {
    const id = irb.readUInt16BE(offset + 4)
    const nameLength = irb[offset + 6]
    let cursor = offset + 6 + 1 + nameLength
    if ((1 + nameLength) % 2) cursor += 1
    if (cursor + 4 > irb.length) break
    const size = irb.readUInt32BE(cursor)
    const end = cursor + 4 + size + (size % 2)
    if (end > irb.length + 1) break
    if (KEPT_PHOTOSHOP_RESOURCES.has(id)) kept.push(irb.subarray(offset, Math.min(end, irb.length)))
    offset = end
  }
  return kept.length && kept.some((block) => block.readUInt16BE(4) === 0x0404) ? Buffer.concat(kept) : null
}

function pngTextKeyword(data) {
  const end = data.indexOf(0)
  return end > 0 ? data.subarray(0, end).toString('latin1') : ''
}

function pngXmpText(data) {
  // iTXt: keyword\0 compressionFlag compressionMethod languageTag\0 translatedKeyword\0 text
  let offset = data.indexOf(0) + 1
  const compressed = data[offset] === 1
  offset += 2
  offset = data.indexOf(0, offset) + 1
  offset = data.indexOf(0, offset) + 1
  if (offset <= 0) return null
  const text = data.subarray(offset)
  return compressed ? inflateCapped(text, MAX_TEXT_BYTES) : Buffer.from(text)
}

function avifBitDepth(bytes) {
  const index = bytes.indexOf(Buffer.from('pixi'))
  if (index < 4 || index + 10 > bytes.length) return 8
  const channels = bytes[index + 8]
  return channels > 0 && index + 9 + channels <= bytes.length ? Math.max(...bytes.subarray(index + 9, index + 9 + channels)) : 8
}

// #endregion helpers

// #region extraction

/**
 * Metadata blocks of a source image. Unknown or damaged parts are skipped; the result is always an object
 * (with `format: null` when the bytes are not an image this module reads).
 */
function extractMetadata(value) {
  const bytes = toBytes(value)
  const format = sniffFormat(bytes)
  const source = {
    format,
    exif: null,
    xmp: null,
    iptc: null,
    icc: null,
    iccName: null,
    density: null,
    pngText: [],
    pngSrgbIntent: null,
    pngColorChunks: false,
    bitDepth: 8,
    colorModel: 'rgb',
    orientation: null,
    problems: [],
  }
  try {
    if (format === 'jpeg') extractJpeg(bytes, source)
    else if (format === 'png') extractPng(bytes, source)
    else if (format === 'webp') extractWebp(bytes, source)
    else if (format === 'avif') source.bitDepth = avifBitDepth(bytes)
    else if (format === 'psd' && bytes.length >= 26) source.bitDepth = bytes.readUInt16BE(22)
  } catch (error) {
    source.problems.push(error.message)
  }
  if (source.exif) {
    try { source.orientation = patchExif(source.exif).info.orientation } catch { /* reported when splicing */ }
  }
  return source
}

function extractJpeg(bytes, source) {
  const { segments } = readJpegSegments(bytes)
  const iccChunks = []
  let adobeTransform = null
  for (const { marker, data } of segments) {
    if (marker === 0xe0 && startsWith(data, Buffer.from('JFIF\0', 'latin1')) && data.length >= 12) {
      const units = data[7]
      const x = data.readUInt16BE(8)
      const y = data.readUInt16BE(10)
      if (x > 0 && y > 0 && !(units === 0 && x === 1 && y === 1)) source.density = { x, y, unit: units === 1 ? 'in' : units === 2 ? 'cm' : null }
    } else if (marker === 0xe1 && startsWith(data, Buffer.from('Exif\0', 'latin1')) && !source.exif) {
      source.exif = Buffer.from(data.subarray(6))
    } else if (marker === 0xe1 && startsWith(data, XMP_NAMESPACE) && !source.xmp) {
      source.xmp = Buffer.from(data.subarray(XMP_NAMESPACE.length))
    } else if (marker === 0xe1 && startsWith(data, XMP_EXTENSION_NAMESPACE)) {
      source.problems.push('Extended XMP')
    } else if (marker === 0xe2 && startsWith(data, ICC_HEADER) && data.length > 14) {
      iccChunks.push({ sequence: data[12], data: data.subarray(14) })
    } else if (marker === 0xed && startsWith(data, PHOTOSHOP_HEADER) && !source.iptc) {
      source.iptc = filterPhotoshopResources(data.subarray(PHOTOSHOP_HEADER.length))
    } else if (marker === 0xee && startsWith(data, Buffer.from('Adobe', 'latin1')) && data.length >= 12) {
      adobeTransform = data[11]
    } else if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker) && data.length >= 6) {
      source.bitDepth = data[0]
      const components = data[5]
      source.colorModel = components === 1 ? 'gray' : components === 4 ? (adobeTransform === 2 ? 'ycck' : 'cmyk') : 'rgb'
    }
  }
  if (iccChunks.length) {
    iccChunks.sort((a, b) => a.sequence - b.sequence)
    const icc = Buffer.concat(iccChunks.map((chunk) => chunk.data))
    if (icc.length <= MAX_ICC_BYTES) source.icc = icc
  }
}

function extractPng(bytes, source) {
  let textBytes = 0
  for (const { type, data } of readPngChunks(bytes)) {
    if (type === 'IHDR' && data.length >= 13) {
      source.bitDepth = data[8]
      source.colorModel = data[9] === 0 || data[9] === 4 ? 'gray' : data[9] === 3 ? 'indexed' : 'rgb'
    } else if (type === 'pHYs' && data.length >= 9) {
      const x = data.readUInt32BE(0)
      const y = data.readUInt32BE(4)
      if (x > 0 && y > 0) source.density = { x, y, unit: data[8] === 1 ? 'm' : null }
    } else if (type === 'iCCP') {
      const nameEnd = data.indexOf(0)
      try {
        if (nameEnd > 0 && data[nameEnd + 1] === 0) {
          source.icc = inflateCapped(data.subarray(nameEnd + 2), MAX_ICC_BYTES)
          source.iccName = data.subarray(0, nameEnd).toString('latin1')
        }
      } catch { source.problems.push('ICC profile') }
    } else if (type === 'sRGB' && data.length >= 1) {
      source.pngSrgbIntent = data[0]
    } else if (type === 'gAMA' || type === 'cHRM') {
      source.pngColorChunks = true
    } else if (type === 'eXIf') {
      source.exif = Buffer.from(data)
    } else if (type === 'iTXt' && pngTextKeyword(data) === PNG_XMP_KEYWORD && !source.xmp) {
      try { source.xmp = pngXmpText(data) } catch { source.problems.push('XMP') }
    } else if (type === 'tEXt' || type === 'zTXt' || type === 'iTXt') {
      textBytes += data.length
      if (textBytes <= MAX_TEXT_BYTES) source.pngText.push({ type, data: Buffer.from(data) })
    }
  }
}

function extractWebp(bytes, source) {
  for (const { type, data } of readRiffChunks(bytes)) {
    if (type === 'ICCP' && data.length <= MAX_ICC_BYTES) source.icc = Buffer.from(data)
    else if (type === 'EXIF') source.exif = Buffer.from(startsWith(data, EXIF_HEADER) ? data.subarray(6) : data)
    else if (type === 'XMP ') source.xmp = Buffer.from(data)
  }
}

/** True when the source has anything spliceMetadata could carry over. */
function hasMetadata(source) {
  return Boolean(source && (source.exif || source.xmp || source.iptc || source.icc || source.density || source.pngText.length || source.pngSrgbIntent != null))
}

/** Approximate memory held by an extracted metadata object (for caches). */
function metadataSize(source) {
  if (!source) return 0
  return [source.exif, source.xmp, source.iptc, source.icc, ...source.pngText.map((chunk) => chunk.data)]
    .reduce((total, block) => total + (block ? block.length : 0), 256)
}

/** Fidelity losses a re-encode from canvas pixels always causes (8-bit RGB output), as user-facing text. */
function fidelityWarnings(source, outputFormat = null) {
  const warnings = []
  if (!source || !source.format) return warnings
  if (source.bitDepth > 8 && outputFormat !== 'psd') {
    warnings.push(`The original image had ${source.bitDepth} bits per channel; the saved image has 8 bits per channel.`)
  }
  if (source.colorModel === 'cmyk' || source.colorModel === 'ycck') {
    warnings.push('The original CMYK colors were converted to RGB.')
  }
  return warnings
}

// #endregion extraction

// #region splicing

function densityToJfif(density) {
  if (!density) return null
  let { x, y } = density
  let units = 0
  if (density.unit === 'in') units = 1
  else if (density.unit === 'cm') units = 2
  else if (density.unit === 'm') { units = 1; x *= 0.0254; y *= 0.0254 }
  x = Math.round(x)
  y = Math.round(y)
  if (x < 1 || y < 1 || x > 0xffff || y > 0xffff) return null
  return { units, x, y }
}

function densityToPhys(density) {
  if (!density) return null
  let { x, y } = density
  let unit = 1
  if (density.unit === 'in') { x /= 0.0254; y /= 0.0254 } else if (density.unit === 'cm') { x *= 100; y *= 100 } else if (density.unit !== 'm') unit = 0
  x = Math.round(x)
  y = Math.round(y)
  if (x < 1 || y < 1 || x > 0x7fffffff || y > 0x7fffffff) return null
  const data = Buffer.alloc(9)
  data.writeUInt32BE(x, 0)
  data.writeUInt32BE(y, 4)
  data[8] = unit
  return data
}

function iccLabel(info) {
  return info.description ? `ICC profile (${info.description})` : 'ICC profile'
}

/**
 * Decides which blocks of `source` go into an output of `outputFormat`, patching EXIF/XMP for the new size.
 * Returns the parts plus `kept`/`dropped` labels.
 */
function planParts(source, outputFormat, dimensions, colorProfile) {
  const kept = []
  const dropped = []
  const parts = { exif: null, xmp: null, iptc: null, icc: null, iccName: source.iccName, density: source.density, pngText: [], srgbIntent: null }
  if (source.exif) {
    try {
      const patched = patchExif(source.exif, dimensions)
      parts.exif = patched.tiff
      if (!parts.density && patched.info.resolution) parts.density = patched.info.resolution
      if (outputFormat === 'jpeg' && parts.exif.length > JPEG_SEGMENT_PAYLOAD - EXIF_HEADER.length) {
        parts.exif = null
        dropped.push('EXIF (too large for a JPEG segment)')
      } else kept.push('EXIF')
    } catch {
      dropped.push('EXIF (damaged)')
    }
  }
  if (source.xmp) {
    parts.xmp = patchXmp(source.xmp, dimensions)
    if (outputFormat === 'jpeg' && parts.xmp.length > JPEG_SEGMENT_PAYLOAD - XMP_NAMESPACE.length) {
      parts.xmp = null
      dropped.push('XMP (too large for a JPEG segment)')
    } else kept.push('XMP')
  }
  if (source.iptc) {
    if (outputFormat === 'jpeg' && source.iptc.length <= JPEG_SEGMENT_PAYLOAD - PHOTOSHOP_HEADER.length) {
      parts.iptc = source.iptc
      kept.push('IPTC')
    } else dropped.push('IPTC')
  }
  if (source.icc) {
    const info = iccInfo(source.icc)
    if (colorProfile === 'keep' || (colorProfile !== 'drop' && info.srgb)) {
      parts.icc = source.icc
      kept.push(iccLabel(info))
    } else {
      dropped.push(colorProfile === 'drop' ? iccLabel(info) : `${iccLabel(info)}: colors were converted to sRGB`)
    }
  }
  if (outputFormat === 'png') {
    parts.pngText = source.pngText
    if (parts.pngText.length) kept.push('PNG text')
    if (!parts.icc) {
      if (source.pngSrgbIntent != null) parts.srgbIntent = source.pngSrgbIntent
      else if (source.icc || source.pngColorChunks) parts.srgbIntent = 0
    }
  } else if (source.pngText.length) {
    dropped.push('PNG text')
  }
  if (parts.density && outputFormat !== 'webp') kept.push('Resolution (DPI)')
  return { parts, kept, dropped }
}

function spliceJpeg(output, parts) {
  const { segments, rest } = readJpegSegments(output)
  const jfif = densityToJfif(parts.density)
  const header = []
  const others = []
  let sawJfif = false
  for (const segment of segments) {
    const { marker, data } = segment
    if (marker === 0xe0 && startsWith(data, Buffer.from('JFIF\0', 'latin1')) && data.length >= 12 && !sawJfif) {
      sawJfif = true
      const patched = Buffer.from(data)
      if (jfif) {
        patched[7] = jfif.units
        patched.writeUInt16BE(jfif.x, 8)
        patched.writeUInt16BE(jfif.y, 10)
      }
      header.unshift(jpegSegment(0xe0, patched))
      continue
    }
    if (marker === 0xe1 && (startsWith(data, Buffer.from('Exif\0', 'latin1')) ? parts.exif : startsWith(data, XMP_NAMESPACE) ? parts.xmp : false)) continue
    if (marker === 0xe2 && startsWith(data, ICC_HEADER) && parts.icc) continue
    if (marker === 0xed && startsWith(data, PHOTOSHOP_HEADER) && parts.iptc) continue
    others.push(jpegSegment(marker, data))
  }
  if (!sawJfif && jfif && !parts.exif) {
    const data = Buffer.from([0x4a, 0x46, 0x49, 0x46, 0, 1, 1, jfif.units, 0, 0, 0, 0, 0, 0])
    data.writeUInt16BE(jfif.x, 8)
    data.writeUInt16BE(jfif.y, 10)
    header.unshift(jpegSegment(0xe0, data))
  }
  if (parts.exif) header.push(jpegSegment(0xe1, Buffer.concat([EXIF_HEADER, parts.exif])))
  if (parts.xmp) header.push(jpegSegment(0xe1, Buffer.concat([XMP_NAMESPACE, parts.xmp])))
  if (parts.icc) {
    const chunkSize = JPEG_SEGMENT_PAYLOAD - ICC_HEADER.length - 2
    const count = Math.ceil(parts.icc.length / chunkSize)
    if (count <= 255) {
      for (let index = 0; index < count; index += 1) {
        const piece = parts.icc.subarray(index * chunkSize, (index + 1) * chunkSize)
        header.push(jpegSegment(0xe2, Buffer.concat([ICC_HEADER, Buffer.from([index + 1, count]), piece])))
      }
    }
  }
  if (parts.iptc) header.push(jpegSegment(0xed, Buffer.concat([PHOTOSHOP_HEADER, parts.iptc])))
  return Buffer.concat([Buffer.from([0xff, 0xd8]), ...header, ...others, rest])
}

function splicePng(output, parts) {
  const chunks = readPngChunks(output)
  if (!chunks.length || chunks[0].type !== 'IHDR') throw new Error('The PNG has no header.')
  const inserted = []
  const replaced = new Set()
  if (parts.icc) {
    const name = /^[\x20-\x7e]{1,79}$/.test(parts.iccName || '') ? parts.iccName : 'ICC Profile'
    inserted.push(pngChunk('iCCP', Buffer.concat([Buffer.from(`${name}\0\0`, 'latin1'), zlib.deflateSync(parts.icc)])))
    for (const type of ['iCCP', 'sRGB', 'gAMA', 'cHRM']) replaced.add(type)
  } else if (parts.srgbIntent != null && !chunks.some((chunk) => chunk.type === 'iCCP' || chunk.type === 'sRGB')) {
    inserted.push(pngChunk('sRGB', Buffer.from([parts.srgbIntent & 3])))
  }
  const phys = densityToPhys(parts.density)
  if (phys) { inserted.push(pngChunk('pHYs', phys)); replaced.add('pHYs') }
  if (parts.exif) { inserted.push(pngChunk('eXIf', parts.exif)); replaced.add('eXIf') }
  if (parts.xmp) {
    inserted.push(pngChunk('iTXt', Buffer.concat([Buffer.from(`${PNG_XMP_KEYWORD}\0\0\0\0\0`, 'latin1'), parts.xmp])))
  }
  for (const chunk of parts.pngText) inserted.push(pngChunk(chunk.type, chunk.data))
  const body = chunks.slice(1).filter((chunk) => {
    if (replaced.has(chunk.type)) return false
    if (parts.xmp && chunk.type === 'iTXt' && pngTextKeyword(chunk.data) === PNG_XMP_KEYWORD) return false
    return true
  })
  return Buffer.concat([PNG_SIGNATURE, chunks[0].raw, ...inserted, ...body.map((chunk) => chunk.raw)])
}

function spliceWebp(output, parts) {
  const chunks = readRiffChunks(output)
  const existing = chunks.find((chunk) => chunk.type === 'VP8X')
  let flags = existing ? existing.data[0] : 0
  let dimensions
  if (existing && existing.data.length >= 10) {
    dimensions = { width: existing.data.readUIntLE(4, 3) + 1, height: existing.data.readUIntLE(7, 3) + 1 }
  } else {
    dimensions = imageDimensions(output, '.webp')
    if (!dimensions) throw new Error('The WebP size could not be read.')
    const lossless = chunks.find((chunk) => chunk.type === 'VP8L')
    if (lossless && lossless.data.length >= 5 && (lossless.data.readUInt32LE(1) >>> 28) & 1) flags |= 0x10
    if (chunks.some((chunk) => chunk.type === 'ALPH')) flags |= 0x10
  }
  const icc = parts.icc || chunks.find((chunk) => chunk.type === 'ICCP')?.data || null
  const exif = parts.exif || chunks.find((chunk) => chunk.type === 'EXIF')?.data || null
  const xmp = parts.xmp || chunks.find((chunk) => chunk.type === 'XMP ')?.data || null
  flags &= ~(0x20 | 0x08 | 0x04)
  if (icc) flags |= 0x20
  if (exif) flags |= 0x08
  if (xmp) flags |= 0x04
  const vp8x = Buffer.alloc(10)
  vp8x[0] = flags
  vp8x.writeUIntLE(dimensions.width - 1, 4, 3)
  vp8x.writeUIntLE(dimensions.height - 1, 7, 3)
  const body = chunks.filter((chunk) => !['VP8X', 'ICCP', 'EXIF', 'XMP '].includes(chunk.type))
  const payload = Buffer.concat([
    Buffer.from('WEBP', 'latin1'),
    riffChunk('VP8X', vp8x),
    ...(icc ? [riffChunk('ICCP', icc)] : []),
    ...body.map((chunk) => riffChunk(chunk.type, chunk.data)),
    ...(exif ? [riffChunk('EXIF', exif)] : []),
    ...(xmp ? [riffChunk('XMP ', xmp)] : []),
  ])
  const header = Buffer.alloc(8)
  header.write('RIFF', 0, 'latin1')
  header.writeUInt32LE(payload.length, 4)
  return Buffer.concat([header, payload])
}

/**
 * Copies `source` metadata (from extractMetadata) into freshly encoded `output` bytes (JPEG, PNG or WebP).
 * Options: `colorProfile` 'auto' (default: keep sRGB profiles only), 'keep' or 'drop'.
 * Never throws for metadata problems: on failure the output comes back unchanged with a `dropped` note.
 */
function spliceMetadata(outputValue, source, options = {}) {
  const output = toBytes(outputValue)
  const result = { bytes: output, kept: [], dropped: [] }
  if (!source || !source.format || !hasMetadata(source)) return result
  const outputFormat = sniffFormat(output)
  if (!['jpeg', 'png', 'webp'].includes(outputFormat)) return result
  const colorProfile = ['keep', 'drop'].includes(options.colorProfile) ? options.colorProfile : 'auto'
  const dimensions = imageDimensions(output, outputFormat === 'jpeg' ? '.jpg' : `.${outputFormat}`) || {}
  const plan = planParts(source, outputFormat, { width: dimensions.width ?? null, height: dimensions.height ?? null }, colorProfile)
  if (outputFormat === 'webp' && source.density && !plan.parts.exif) plan.dropped.push('Resolution (DPI) outside EXIF')
  for (const problem of source.problems) plan.dropped.push(`${problem} (unreadable)`)
  const nothingToWrite = !plan.parts.exif && !plan.parts.xmp && !plan.parts.iptc && !plan.parts.icc && !plan.parts.density
    && !plan.parts.pngText.length && plan.parts.srgbIntent == null
  if (nothingToWrite) return { bytes: output, kept: [], dropped: plan.dropped }
  try {
    const bytes = outputFormat === 'jpeg' ? spliceJpeg(output, plan.parts)
      : outputFormat === 'png' ? splicePng(output, plan.parts)
        : spliceWebp(output, plan.parts)
    return { bytes, kept: plan.kept, dropped: plan.dropped }
  } catch (error) {
    return { bytes: output, kept: [], dropped: [`All metadata (${error.message})`] }
  }
}

// #endregion splicing

module.exports = {
  crc32,
  extractMetadata,
  fidelityWarnings,
  hasMetadata,
  iccInfo,
  metadataSize,
  patchExif,
  patchXmp,
  readJpegSegments,
  readPngChunks,
  readRiffChunks,
  spliceMetadata,
}
