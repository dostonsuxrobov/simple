const fs = require('node:fs/promises')
const path = require('node:path')
const crypto = require('node:crypto')
const zlib = require('node:zlib')

const MAX_IMAGE_BYTES = 512 * 1024 * 1024
const MAX_IMAGE_DIMENSION = 20_000
const MAX_IMAGE_PIXELS = 50_000_000
/** Decompressed size cap for .svgz (and gzip-compressed .svg) files. A gzip bomb stops here. */
const MAX_SVG_BYTES = 32 * 1024 * 1024
/** Vector images have no intrinsic pixels: small ones rasterize at this long edge, huge ones are capped. */
const SVG_MIN_RASTER_EDGE = 2048
const SVG_MAX_RASTER_EDGE = 8192
const SUPPORTED_EXTENSIONS = Object.freeze([
  '.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg', '.avif',
  '.jfif', '.jpe', '.jif', '.apng', '.ico', '.svgz', '.psd',
])
/** Extensions whose files can be overwritten in place (when the content really is that format and still). */
const EDITABLE_EXTENSIONS = Object.freeze(['.png', '.jpg', '.jpeg', '.webp', '.jfif', '.jpe', '.jif', '.apng', '.psd'])

/** Canonical content formats. `extensions[0]` is the extension a new file of this format gets. */
const FORMATS = Object.freeze({
  png: Object.freeze({ label: 'PNG', mime: 'image/png', extensions: Object.freeze(['.png', '.apng']) }),
  jpeg: Object.freeze({ label: 'JPEG', mime: 'image/jpeg', extensions: Object.freeze(['.jpg', '.jpeg', '.jfif', '.jpe', '.jif']) }),
  webp: Object.freeze({ label: 'WebP', mime: 'image/webp', extensions: Object.freeze(['.webp']) }),
  gif: Object.freeze({ label: 'GIF', mime: 'image/gif', extensions: Object.freeze(['.gif']) }),
  bmp: Object.freeze({ label: 'BMP', mime: 'image/bmp', extensions: Object.freeze(['.bmp']) }),
  svg: Object.freeze({ label: 'SVG', mime: 'image/svg+xml', extensions: Object.freeze(['.svg', '.svgz']) }),
  avif: Object.freeze({ label: 'AVIF', mime: 'image/avif', extensions: Object.freeze(['.avif']) }),
  ico: Object.freeze({ label: 'icon', mime: 'image/x-icon', extensions: Object.freeze(['.ico']) }),
  psd: Object.freeze({ label: 'Photoshop', mime: 'image/vnd.adobe.photoshop', extensions: Object.freeze(['.psd']) }),
})
const FORMAT_BY_EXTENSION = Object.freeze(Object.fromEntries(
  Object.entries(FORMATS).flatMap(([format, info]) => info.extensions.map((extension) => [extension, format])),
))
const MIME_BY_EXTENSION = Object.freeze(Object.fromEntries(
  SUPPORTED_EXTENSIONS.map((extension) => [extension, FORMATS[FORMAT_BY_EXTENSION[extension]].mime]),
))
/** Recognised but not decodable by Chromium: a specific message beats "not supported". */
const UNSUPPORTED_FORMAT_MESSAGES = Object.freeze({
  tiff: 'TIFF images are not supported yet. Save the image as PNG or JPEG in another program first.',
  heic: 'HEIC photos are not supported yet. Convert the photo to JPEG first (Windows Photos can do this).',
  jxl: 'JPEG XL images are not supported yet.',
  psb: 'Large Photoshop documents (.psb) are not supported. Save the document as .psd in Photoshop first.',
})
const PSD_COLOR_MODES = Object.freeze({ 0: 'Bitmap', 1: 'Grayscale', 2: 'Indexed', 3: 'RGB', 4: 'CMYK', 7: 'Multichannel', 8: 'Duotone', 9: 'Lab' })

const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])

function checkedDimensions(width, height) {
  const normalizedWidth = Math.round(Math.abs(Number(width)))
  const normalizedHeight = Math.round(Math.abs(Number(height)))
  if (!Number.isFinite(normalizedWidth) || !Number.isFinite(normalizedHeight) || normalizedWidth < 1 || normalizedHeight < 1) return null
  return { width: normalizedWidth, height: normalizedHeight }
}

function jpegDimensions(bytes) {
  let offset = 2
  const startOfFrame = new Set([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf])
  while (offset + 8 < bytes.length) {
    if (bytes[offset] !== 0xff) { offset += 1; continue }
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1
    const marker = bytes[offset]
    offset += 1
    if (marker === 0xd8 || marker === 0xd9 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue
    if (offset + 2 > bytes.length) break
    const length = bytes.readUInt16BE(offset)
    if (length < 2 || offset + length > bytes.length) break
    if (startOfFrame.has(marker) && length >= 7) return checkedDimensions(bytes.readUInt16BE(offset + 5), bytes.readUInt16BE(offset + 3))
    offset += length
  }
  return null
}

function webpDimensions(bytes) {
  const kind = bytes.subarray(12, 16).toString('ascii')
  if (kind === 'VP8X' && bytes.length >= 30) {
    return checkedDimensions(bytes.readUIntLE(24, 3) + 1, bytes.readUIntLE(27, 3) + 1)
  }
  if (kind === 'VP8L' && bytes.length >= 25 && bytes[20] === 0x2f) {
    const bits = bytes.readUInt32LE(21)
    return checkedDimensions((bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1)
  }
  if (kind === 'VP8 ' && bytes.length >= 30 && bytes[23] === 0x9d && bytes[24] === 0x01 && bytes[25] === 0x2a) {
    return checkedDimensions(bytes.readUInt16LE(26) & 0x3fff, bytes.readUInt16LE(28) & 0x3fff)
  }
  return null
}

// #region SVG (text decoding, safe DOCTYPE handling, intrinsic and raster size)

function isGzip(bytes) {
  return bytes.length >= 3 && bytes[0] === 0x1f && bytes[1] === 0x8b && bytes[2] === 0x08
}

/** Gunzips with a hard output cap so a decompression bomb fails fast instead of exhausting memory. */
function gunzipSvg(bytes) {
  try {
    return zlib.gunzipSync(bytes, { maxOutputLength: MAX_SVG_BYTES })
  } catch (error) {
    if (error && (error.code === 'ERR_BUFFER_TOO_LARGE' || error instanceof RangeError)) {
      throw new Error('This compressed SVG expands beyond the 32 MB safety limit.')
    }
    throw new Error('This compressed SVG is damaged and could not be decompressed.')
  }
}

/** Decodes SVG/XML text by BOM (UTF-8, UTF-16 LE/BE), BOM-less UTF-16, or a Latin-1 encoding declaration. */
function decodeSvgText(bytes, limit = bytes.length) {
  const slice = bytes.subarray(0, Math.min(bytes.length, limit))
  if (slice.length >= 2 && slice[0] === 0xff && slice[1] === 0xfe) return new TextDecoder('utf-16le').decode(slice.subarray(2))
  if (slice.length >= 2 && slice[0] === 0xfe && slice[1] === 0xff) return new TextDecoder('utf-16be').decode(slice.subarray(2))
  if (slice.length >= 3 && slice[0] === 0xef && slice[1] === 0xbb && slice[2] === 0xbf) return new TextDecoder('utf-8').decode(slice.subarray(3))
  if (slice.length >= 4 && slice[1] === 0 && slice[3] === 0 && slice[0] !== 0 && slice[2] !== 0) return new TextDecoder('utf-16le').decode(slice)
  if (slice.length >= 4 && slice[0] === 0 && slice[2] === 0 && slice[1] !== 0 && slice[3] !== 0) return new TextDecoder('utf-16be').decode(slice)
  const head = slice.subarray(0, 256).toString('latin1')
  const declared = /^\s*<\?xml[^>]*\bencoding\s*=\s*["']([A-Za-z0-9._-]+)["']/i.exec(head)?.[1]?.toLowerCase()
  if (declared && /^(?:iso-8859-1|latin-?1|windows-1252|cp1252|us-ascii|ascii)$/.test(declared)) return new TextDecoder('windows-1252').decode(slice)
  return new TextDecoder('utf-8').decode(slice)
}

/**
 * Finds the end of a `<!DOCTYPE ...>` declaration starting at `start`, honouring quoted strings,
 * comments and an internal subset `[ ... ]` (which may itself contain `>` characters).
 */
function doctypeEnd(text, start) {
  let index = start + 9
  let inSubset = false
  while (index < text.length) {
    const char = text[index]
    if (char === '"' || char === "'") {
      const close = text.indexOf(char, index + 1)
      if (close < 0) return -1
      index = close + 1
      continue
    }
    if (inSubset && text.startsWith('<!--', index)) {
      const close = text.indexOf('-->', index + 4)
      if (close < 0) return -1
      index = close + 3
      continue
    }
    if (char === '[' && !inSubset) inSubset = true
    else if (char === ']' && inSubset) inSubset = false
    else if (char === '>' && !inSubset) return index + 1
    index += 1
  }
  return -1
}

/**
 * Walks the XML prolog (whitespace, XML declaration, processing instructions, comments, one DOCTYPE)
 * and returns where the root element starts plus the DOCTYPE span, or null when the prolog is not XML.
 */
function svgProlog(text) {
  let index = text.charCodeAt(0) === 0xfeff ? 1 : 0
  let doctype = null
  while (index < text.length) {
    const char = text[index]
    if (char === ' ' || char === '\t' || char === '\r' || char === '\n') { index += 1; continue }
    if (text.startsWith('<?', index)) {
      const close = text.indexOf('?>', index + 2)
      if (close < 0) return null
      index = close + 2
      continue
    }
    if (text.startsWith('<!--', index)) {
      const close = text.indexOf('-->', index + 4)
      if (close < 0) return null
      index = close + 3
      continue
    }
    if (/^<!doctype\s/i.test(text.slice(index, index + 10))) {
      if (doctype) return null
      const end = doctypeEnd(text, index)
      if (end < 0) return null
      doctype = { start: index, end }
      index = end
      continue
    }
    break
  }
  if (!/^<svg(?:[\s/>]|$)/i.test(text.slice(index, index + 5))) return null
  return { rootStart: index, doctype }
}

/** End index (exclusive) of the start tag beginning at `start`, honouring quoted attribute values. */
function tagEnd(text, start) {
  let index = start + 1
  while (index < text.length) {
    const char = text[index]
    if (char === '"' || char === "'") {
      const close = text.indexOf(char, index + 1)
      if (close < 0) return -1
      index = close + 1
      continue
    }
    if (char === '>') return index + 1
    index += 1
  }
  return -1
}

function looksLikeSvg(bytes) {
  const text = decodeSvgText(bytes, 64 * 1024)
  return Boolean(svgProlog(text))
}

function svgLength(value) {
  const match = /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)\s*(px|in|cm|mm|q|pt|pc)?\s*$/i.exec(String(value || ''))
  if (!match) return null
  const factors = { px: 1, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, q: 96 / 101.6, pt: 96 / 72, pc: 16 }
  const length = Number(match[1]) * factors[(match[2] || 'px').toLowerCase()]
  return Number.isFinite(length) && length > 0 ? length : null
}

function svgAttribute(tag, name) {
  const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, 'i').exec(tag)
  return match?.[1] ?? match?.[2] ?? match?.[3] ?? null
}

function svgViewBox(tag) {
  const values = String(svgAttribute(tag, 'viewBox') || '').trim().split(/[\s,]+/).map(Number)
  if (values.length !== 4 || !values.every(Number.isFinite) || values[2] <= 0 || values[3] <= 0) return null
  return { x: values[0], y: values[1], width: values[2], height: values[3] }
}

/** Intrinsic CSS-pixel size of a root `<svg>` start tag, following the CSS replaced-element rules. */
function svgTagSize(tag) {
  const viewBox = svgViewBox(tag)
  let width = svgLength(svgAttribute(tag, 'width'))
  let height = svgLength(svgAttribute(tag, 'height'))
  if (viewBox) {
    if (width == null && height == null) { width = viewBox.width; height = viewBox.height }
    else if (width == null) width = height * (viewBox.width / viewBox.height)
    else if (height == null) height = width * (viewBox.height / viewBox.width)
  }
  return { width: width ?? 300, height: height ?? 150, viewBox }
}

/**
 * Raster size for a vector image: small or unit-less drawings (icons, viewBox-only files) render at a
 * 2048 px long edge instead of Chromium's 300x150 default, and very large ones are capped so the editor
 * stays within the canvas limits. The aspect ratio is kept.
 */
function svgRasterSize(width, height) {
  const longEdge = Math.max(width, height)
  const target = Math.min(Math.max(longEdge, SVG_MIN_RASTER_EDGE), SVG_MAX_RASTER_EDGE)
  let scale = target / longEdge
  const pixels = width * scale * height * scale
  if (pixels > MAX_IMAGE_PIXELS) scale *= Math.sqrt(MAX_IMAGE_PIXELS / pixels)
  return {
    width: Math.max(1, Math.min(MAX_IMAGE_DIMENSION, Math.floor(width * scale))),
    height: Math.max(1, Math.min(MAX_IMAGE_DIMENSION, Math.floor(height * scale))),
  }
}

function svgDimensions(bytes) {
  const text = decodeSvgText(isGzip(bytes) ? gunzipSvg(bytes) : bytes, 256 * 1024)
  const prolog = svgProlog(text)
  const start = prolog ? prolog.rootStart : text.search(/<svg\b/i)
  if (start < 0) return null
  const end = tagEnd(text, start)
  if (end < 0) return null
  const size = svgTagSize(text.slice(start, end))
  return checkedDimensions(size.width, size.height)
}

const XML_PREDEFINED_ENTITIES = new Set(['amp', 'lt', 'gt', 'quot', 'apos'])
const MAX_ENTITY_EXPANSION = 8 * 1024 * 1024

/**
 * Internal general entities from a DOCTYPE subset, e.g. Illustrator's `<!ENTITY ns_svg "http://...">`.
 * External (SYSTEM/PUBLIC) and parameter entities are never resolved.
 */
function internalEntities(doctypeText) {
  const entities = new Map()
  const pattern = /<!ENTITY\s+([A-Za-z_:][\w.:-]*)\s+("([^"]*)"|'([^']*)')\s*>/g
  let match
  while ((match = pattern.exec(doctypeText))) {
    const [, name, , doubleQuoted, singleQuoted] = match
    if (XML_PREDEFINED_ENTITIES.has(name) || entities.has(name)) continue
    entities.set(name, doubleQuoted ?? singleQuoted ?? '')
  }
  // Resolve references between entities (bounded depth, so recursive definitions cannot loop).
  for (let pass = 0; pass < 4; pass += 1) {
    let changed = false
    for (const [name, value] of entities) {
      const expanded = value.replace(/&([A-Za-z_:][\w.:-]*);/g, (whole, reference) => (
        reference !== name && entities.has(reference) && !/&[A-Za-z_:][\w.:-]*;/.test(entities.get(reference)) ? entities.get(reference) : whole
      ))
      if (expanded.length > MAX_ENTITY_EXPANSION) throw new Error('This SVG declares entities that expand beyond the safety limit.')
      if (expanded !== value) { entities.set(name, expanded); changed = true }
    }
    if (!changed) break
  }
  return entities
}

function expandEntities(text, entities) {
  if (!entities.size) return text
  let total = 0
  const expanded = text.replace(/&([A-Za-z_:][\w.:-]*);/g, (whole, name) => {
    if (!entities.has(name)) return whole
    const value = entities.get(name)
    total += value.length
    if (total > MAX_ENTITY_EXPANSION) throw new Error('This SVG declares entities that expand beyond the safety limit.')
    return value
  })
  return expanded
}

function formatSvgNumber(value) {
  return String(Math.round(value * 1000) / 1000)
}

/**
 * Turns SVG/SVGZ bytes into a self-contained UTF-8 SVG the renderer can rasterize at a sensible size.
 * Nothing is executed or fetched: the text is only decoded, the DOCTYPE is removed after its internal
 * entities are expanded (bounded), the XML encoding declaration is dropped (the output is UTF-8), and the
 * root `width`/`height` are rewritten to the raster size (adding a viewBox when the drawing had none), so
 * `<img>.naturalWidth/Height` equals the size reported here.
 */
function prepareSvg(value) {
  let bytes = toBytes(value)
  const compressed = isGzip(bytes)
  if (compressed) bytes = gunzipSvg(bytes)
  if (bytes.length > MAX_SVG_BYTES) throw new Error('This SVG is larger than the 32 MB safety limit.')
  let text = decodeSvgText(bytes)
  const prolog = svgProlog(text)
  if (!prolog) throw new Error('The file content is not a valid SVG image.')
  let rootStart = prolog.rootStart
  if (prolog.doctype) {
    const entities = internalEntities(text.slice(prolog.doctype.start, prolog.doctype.end))
    const before = text.slice(0, prolog.doctype.start)
    const after = expandEntities(text.slice(prolog.doctype.end), entities)
    text = before + after
    rootStart -= prolog.doctype.end - prolog.doctype.start
  }
  // The text is re-encoded as UTF-8, so a UTF-16 or Latin-1 declaration would now be wrong.
  const declaration = /^(﻿?\s*<\?xml\b[^>]*?)\s+encoding\s*=\s*(?:"[^"]*"|'[^']*')/i.exec(text)
  if (declaration) {
    const removed = declaration[0].length - declaration[1].length
    text = declaration[1] + text.slice(declaration[0].length)
    rootStart -= removed
  }
  if (text.charCodeAt(0) === 0xfeff) { text = text.slice(1); rootStart -= 1 }
  const rootEnd = tagEnd(text, rootStart)
  if (rootEnd < 0) throw new Error('The file content is not a valid SVG image.')
  const tag = text.slice(rootStart, rootEnd)
  const intrinsic = svgTagSize(tag)
  const intrinsicSize = checkedDimensions(intrinsic.width, intrinsic.height)
  if (!intrinsicSize) throw new Error('The image dimensions could not be read safely.')
  const raster = svgRasterSize(intrinsic.width, intrinsic.height)
  const selfClosing = /\/\s*>$/.test(tag)
  let attributes = tag.slice(4, tag.length - (selfClosing ? 2 : 1))
    .replace(/(^|\s)(?:width|height)\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)/gi, '$1')
  if (!intrinsic.viewBox) attributes += ` viewBox="0 0 ${formatSvgNumber(intrinsic.width)} ${formatSvgNumber(intrinsic.height)}"`
  if (!/(?:^|\s)preserveAspectRatio\s*=/.test(attributes) && !intrinsic.viewBox) attributes += ' preserveAspectRatio="none"'
  const rewritten = `<svg${attributes.replace(/\s+$/, '')} width="${raster.width}" height="${raster.height}"${selfClosing ? '/>' : '>'}`
  text = text.slice(0, rootStart) + rewritten + text.slice(rootEnd)
  return {
    bytes: Buffer.from(text, 'utf8'),
    compressed,
    intrinsic: intrinsicSize,
    raster,
  }
}

// #endregion SVG

function avifDimensions(bytes) {
  const marker = Buffer.from('ispe')
  let offset = 0
  let largest = null
  while ((offset = bytes.indexOf(marker, offset)) >= 0) {
    if (offset >= 4 && offset + 16 <= bytes.length) {
      const boxSize = bytes.readUInt32BE(offset - 4)
      if (boxSize >= 20 && offset - 4 + boxSize <= bytes.length) {
        const candidate = checkedDimensions(bytes.readUInt32BE(offset + 8), bytes.readUInt32BE(offset + 12))
        if (candidate && (!largest || candidate.width * candidate.height > largest.width * largest.height)) largest = candidate
      }
    }
    offset += 4
  }
  return largest
}

/** Largest entry of an ICO/CUR directory (a stored width or height of 0 means 256). */
function icoDimensions(bytes) {
  if (bytes.length < 6) return null
  const count = bytes.readUInt16LE(4)
  let largest = null
  for (let index = 0; index < count && 6 + (index + 1) * 16 <= bytes.length; index += 1) {
    const entry = 6 + index * 16
    let width = bytes[entry] || 256
    let height = bytes[entry + 1] || 256
    const dataSize = bytes.readUInt32LE(entry + 8)
    const dataOffset = bytes.readUInt32LE(entry + 12)
    // A PNG-compressed entry knows its real size (it can exceed 256 px).
    if (dataOffset + 24 <= bytes.length && dataSize >= 24 && bytes.subarray(dataOffset, dataOffset + 8).equals(PNG_SIGNATURE)) {
      width = bytes.readUInt32BE(dataOffset + 16)
      height = bytes.readUInt32BE(dataOffset + 20)
    }
    const candidate = checkedDimensions(width, height)
    if (candidate && (!largest || candidate.width * candidate.height > largest.width * largest.height)) largest = candidate
  }
  return largest
}

function psdHeader(bytes) {
  if (bytes.length < 26 || bytes.subarray(0, 4).toString('ascii') !== '8BPS') return null
  return {
    version: bytes.readUInt16BE(4),
    channels: bytes.readUInt16BE(12),
    height: bytes.readUInt32BE(14),
    width: bytes.readUInt32BE(18),
    depth: bytes.readUInt16BE(22),
    colorMode: bytes.readUInt16BE(24),
  }
}

/** Format name ('png', 'jpeg', ...) for an extension (with or without the dot), or null. */
function formatOfExtension(extension) {
  const value = String(extension || '').toLowerCase()
  return FORMAT_BY_EXTENSION[value.startsWith('.') ? value : `.${value}`] || null
}

/**
 * Identifies image content by its magic bytes alone. Returns a canonical format from FORMATS,
 * 'gzip' for gzip data (possibly .svgz), a key of UNSUPPORTED_FORMAT_MESSAGES, or null.
 */
function sniffFormat(value) {
  const bytes = toBytes(value)
  if (bytes.length >= 8 && bytes.subarray(0, 8).equals(PNG_SIGNATURE)) return 'png'
  if (bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpeg'
  if (bytes.length >= 16 && bytes.subarray(0, 4).toString('ascii') === 'RIFF' && bytes.subarray(8, 12).toString('ascii') === 'WEBP') return 'webp'
  if (bytes.length >= 10 && ['GIF87a', 'GIF89a'].includes(bytes.subarray(0, 6).toString('ascii'))) return 'gif'
  if (bytes.length >= 26 && bytes[0] === 0x42 && bytes[1] === 0x4d && [12, 16, 40, 52, 56, 64, 108, 124].includes(bytes.readUInt32LE(14))) return 'bmp'
  if (bytes.length >= 16 && bytes.subarray(4, 8).toString('ascii') === 'ftyp') {
    const boxSize = bytes.readUInt32BE(0)
    const brands = bytes.subarray(8, Math.min(bytes.length, boxSize >= 16 && boxSize <= 4096 ? boxSize : 64)).toString('latin1')
    if (/avif|avis/.test(brands)) return 'avif'
    if (/hei[cmsx]|hev[cmsx]|mif1|msf1/.test(brands)) return 'heic'
    return null
  }
  if (bytes.length >= 22 && bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 1 && bytes[3] === 0 && bytes.readUInt16LE(4) > 0) return 'ico'
  if (bytes.length >= 26 && bytes.subarray(0, 4).toString('ascii') === '8BPS') {
    const version = bytes.readUInt16BE(4)
    return version === 1 ? 'psd' : version === 2 ? 'psb' : null
  }
  if (bytes.length >= 8 && ((bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0) || (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && bytes[3] === 0x2a))) return 'tiff'
  if ((bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0x0a) || (bytes.length >= 12 && bytes.readUInt32BE(0) === 12 && bytes.subarray(4, 8).toString('ascii') === 'JXL ')) return 'jxl'
  if (isGzip(bytes)) return 'gzip'
  if (looksLikeSvg(bytes)) return 'svg'
  return null
}

/**
 * Animation in the source. APNG: an `acTL` chunk before the first IDAT (frame count from acTL).
 * WebP: VP8X animation flag / `ANMF` frame chunks. GIF: more than one image descriptor.
 * AVIF image sequences ('avis' brand) are flagged with an unknown frame count.
 */
function detectAnimation(value, format = sniffFormat(value)) {
  const bytes = toBytes(value)
  const still = { animated: false, frameCount: 1 }
  if (format === 'png') {
    let offset = 8
    while (offset + 12 <= bytes.length) {
      const length = bytes.readUInt32BE(offset)
      const type = bytes.subarray(offset + 4, offset + 8).toString('latin1')
      if (type === 'IDAT' || type === 'IEND') break
      if (type === 'acTL' && length >= 8 && offset + 16 <= bytes.length) {
        const frames = bytes.readUInt32BE(offset + 8)
        return { animated: frames > 1, frameCount: Math.max(1, frames) }
      }
      if (length > bytes.length) break
      offset += 12 + length
    }
    return still
  }
  if (format === 'webp') {
    let flagged = false
    let frames = 0
    let offset = 12
    while (offset + 8 <= bytes.length) {
      const type = bytes.subarray(offset, offset + 4).toString('latin1')
      const length = bytes.readUInt32LE(offset + 4)
      if (type === 'VP8X' && offset + 9 <= bytes.length) flagged = (bytes[offset + 8] & 0x02) !== 0
      else if (type === 'ANMF') frames += 1
      if (length > bytes.length) break
      offset += 8 + length + (length & 1)
    }
    if (frames > 1) return { animated: true, frameCount: frames }
    if (flagged && frames === 0) return { animated: true, frameCount: null }
    return still
  }
  if (format === 'gif') {
    const frames = gifFrameCount(bytes)
    return { animated: frames > 1, frameCount: Math.max(1, frames) }
  }
  if (format === 'avif' && bytes.length >= 16) {
    const boxSize = bytes.readUInt32BE(0)
    const brands = bytes.subarray(8, Math.min(bytes.length, boxSize >= 16 && boxSize <= 4096 ? boxSize : 64)).toString('latin1')
    if (bytes.subarray(8, 12).toString('latin1') === 'avis' || /avis/.test(brands)) return { animated: true, frameCount: null }
  }
  return still
}

function skipGifSubBlocks(bytes, offset) {
  while (offset < bytes.length) {
    const size = bytes[offset]
    offset += 1
    if (size === 0) return offset
    offset += size
  }
  return -1
}

function gifFrameCount(bytes) {
  if (bytes.length < 13) return 0
  let offset = 13
  const flags = bytes[10]
  if (flags & 0x80) offset += 3 * (1 << ((flags & 0x07) + 1))
  let frames = 0
  while (offset < bytes.length) {
    const block = bytes[offset]
    if (block === 0x3b) break
    if (block === 0x21) {
      offset = skipGifSubBlocks(bytes, offset + 2)
      if (offset < 0) break
      continue
    }
    if (block === 0x2c) {
      if (offset + 10 > bytes.length) break
      frames += 1
      const local = bytes[offset + 9]
      offset += 10
      if (local & 0x80) offset += 3 * (1 << ((local & 0x07) + 1))
      offset = skipGifSubBlocks(bytes, offset + 1)
      if (offset < 0) break
      continue
    }
    break
  }
  return frames
}

function dimensionsForFormat(bytes, format) {
  if (format === 'png' && bytes.length >= 24) return checkedDimensions(bytes.readUInt32BE(16), bytes.readUInt32BE(20))
  if (format === 'jpeg') return jpegDimensions(bytes)
  if (format === 'webp') return webpDimensions(bytes)
  if (format === 'gif' && bytes.length >= 10) return checkedDimensions(bytes.readUInt16LE(6), bytes.readUInt16LE(8))
  if (format === 'bmp' && bytes.length >= 26) {
    const dibSize = bytes.readUInt32LE(14)
    return dibSize === 12
      ? checkedDimensions(bytes.readUInt16LE(18), bytes.readUInt16LE(20))
      : checkedDimensions(bytes.readInt32LE(18), bytes.readInt32LE(22))
  }
  if (format === 'svg') return svgDimensions(bytes)
  if (format === 'avif') return avifDimensions(bytes)
  if (format === 'ico') return icoDimensions(bytes)
  if (format === 'psd') {
    const header = psdHeader(bytes)
    return header ? checkedDimensions(header.width, header.height) : null
  }
  return null
}

/** Intrinsic dimensions of `value` read as the format of `extension` (SVG: CSS pixels before rasterizing). */
function imageDimensions(value, extension) {
  const bytes = toBytes(value)
  return dimensionsForFormat(bytes, formatOfExtension(extension))
}

function assertDimensionsWithinLimits(dimensions) {
  if (!dimensions) throw new Error('The image dimensions could not be read safely.')
  if (
    dimensions.width > MAX_IMAGE_DIMENSION
    || dimensions.height > MAX_IMAGE_DIMENSION
    || dimensions.width * dimensions.height > MAX_IMAGE_PIXELS
  ) {
    throw new Error('This image is too large to edit safely. The limit is 50 megapixels and 20,000 pixels per side.')
  }
  return dimensions
}

/**
 * Validates the pixel size before the renderer decodes anything. A vector SVG is checked at the size it
 * will be rasterized at (see svgRasterSize), so a huge drawing opens scaled down instead of being refused.
 */
function validateImageDimensions(value, extension) {
  const bytes = toBytes(value)
  const format = formatOfExtension(extension)
  if (format === 'svg') {
    const intrinsic = dimensionsForFormat(bytes, 'svg')
    if (!intrinsic) throw new Error('The image dimensions could not be read safely.')
    return assertDimensionsWithinLimits(svgRasterSize(intrinsic.width, intrinsic.height))
  }
  return assertDimensionsWithinLimits(dimensionsForFormat(bytes, format))
}

function toBytes(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return Buffer.from(value || [])
}

function extensionOf(value) {
  return path.extname(String(value || '')).toLowerCase()
}

function isSupportedExtension(value) {
  return SUPPORTED_EXTENSIONS.includes(extensionOf(value))
}

function safeStem(value) {
  return path.basename(String(value || 'Untitled image'), path.extname(String(value || '')))
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim() || 'Untitled image'
}

/** True when `bytes` really are the format the extension names (aliases such as .jfif and .apng included). */
function matchesSignature(bytes, extension) {
  const value = toBytes(bytes)
  const ext = String(extension || '').toLowerCase()
  const expected = formatOfExtension(ext)
  if (!expected) return false
  const actual = sniffFormat(value)
  if (expected === 'svg') return actual === 'svg' || (actual === 'gzip' && ext === '.svgz')
  return actual === expected
}

/**
 * Photoshop documents Simple can open: PSD (not PSB) in Bitmap, Grayscale, Indexed or RGB color at 1, 8, 16
 * or 32 bits, with at most 16 channels (the PSD reader's limit). CMYK, Lab, Duotone and Multichannel get
 * the Photoshop steps to convert a copy.
 */
function assertPsdSupported(bytes) {
  const header = psdHeader(bytes)
  if (!header) throw new Error('This Photoshop document is damaged.')
  if (header.version === 2) throw new Error(UNSUPPORTED_FORMAT_MESSAGES.psb)
  if (header.version !== 1) throw new Error('This Photoshop document is damaged or uses an unknown version.')
  if (![0, 1, 2, 3].includes(header.colorMode)) {
    throw new Error(`This PSD uses ${PSD_COLOR_MODES[header.colorMode] || 'an unsupported'} color. In Photoshop choose Image > Mode > RGB Color, then save a copy.`)
  }
  if (![1, 8, 16, 32].includes(header.depth)) throw new Error('This PSD uses an unsupported bit depth.')
  if ((header.colorMode === 0) !== (header.depth === 1) || (header.colorMode === 2 && header.depth !== 8)) {
    throw new Error('This PSD uses an unsupported bit depth.')
  }
  if (header.channels < 1 || header.channels > 56) throw new Error('This Photoshop document is damaged.')
  if (header.channels > 16) {
    throw new Error('This Photoshop document has more than 16 channels. Delete extra alpha channels in Photoshop (Channels panel), then save a copy.')
  }
}

/**
 * Strict check that `value` is a supported image of exactly the format `extension` names. Used for bytes
 * the app produced (save, export, clipboard, print). Opening a file uses inspectImageBytes instead, which
 * trusts the content over the extension.
 */
function validateImageBytes(value, extension) {
  const bytes = toBytes(value)
  const ext = extensionOf(`file${extension.startsWith('.') ? extension : `.${extension}`}`)
  if (!SUPPORTED_EXTENSIONS.includes(ext)) throw new Error('This image format is not supported.')
  if (!bytes.length) throw new Error('This image is empty.')
  if (bytes.length > MAX_IMAGE_BYTES) throw new Error('This image is larger than the 512 MB safety limit.')
  if (!matchesSignature(bytes, ext)) throw new Error(`The file content does not match the ${ext.slice(1).toUpperCase()} format.`)
  if (formatOfExtension(ext) === 'psd') assertPsdSupported(bytes)
  return bytes
}

/** Name for a copy of `name` in `format`, e.g. 'photo.jpg' (really WebP) -> 'photo.webp'. */
function nameWithFormat(name, format, suffix = '') {
  return `${safeStem(name)}${suffix}${FORMATS[format]?.extensions[0] || '.png'}`
}

/**
 * Everything the open handlers need to know about a file's bytes, judged by content (magic bytes), not
 * by the name: the real format, whether that disagrees with the extension, the size to decode at, and
 * whether the source is animated. SVG/SVGZ bytes come back as a normalized, self-contained UTF-8 SVG.
 */
function inspectImageBytes(value, name = '') {
  let bytes = toBytes(value)
  if (!bytes.length) throw new Error('This image is empty.')
  if (bytes.length > MAX_IMAGE_BYTES) throw new Error('This image is larger than the 512 MB safety limit.')
  const extension = extensionOf(name)
  const extensionFormat = formatOfExtension(extension)
  let format = sniffFormat(bytes)
  if (format && UNSUPPORTED_FORMAT_MESSAGES[format]) throw new Error(UNSUPPORTED_FORMAT_MESSAGES[format])
  if (format === 'gzip') {
    if (extensionFormat && extensionFormat !== 'svg') throw new Error('This file is compressed and is not an image.')
    format = 'svg'
  }
  if (!format) {
    throw new Error(extensionFormat
      ? `The file content does not match the ${extension.slice(1).toUpperCase()} format, and is not another supported image.`
      : 'This file is not a supported image.')
  }
  if (format === 'psd') assertPsdSupported(bytes)
  let dimensions
  let intrinsic = null
  if (format === 'svg') {
    const svg = prepareSvg(bytes)
    bytes = svg.bytes
    intrinsic = svg.intrinsic
    dimensions = assertDimensionsWithinLimits(svg.raster)
  } else {
    dimensions = assertDimensionsWithinLimits(dimensionsForFormat(bytes, format))
  }
  const animation = detectAnimation(bytes, format)
  const mismatch = Boolean(extensionFormat) && extensionFormat !== format
  // The format string keeps the familiar extension spelling ('jpg' / 'jpeg') when it is accurate.
  const payloadFormat = !mismatch && ['.jpg', '.jpeg'].includes(extension)
    ? extension.slice(1)
    : format === 'jpeg' ? 'jpg' : format
  const canSaveInPlace = !mismatch && !animation.animated && EDITABLE_EXTENSIONS.includes(extension)
  const notices = []
  if (mismatch) {
    notices.push(`This file is actually a${/^[aeiou]/i.test(FORMATS[format].label) ? 'n' : ''} ${FORMATS[format].label} image. Save will create ${nameWithFormat(name, saveFormatFor(format))}.`)
  }
  if (animation.animated) {
    notices.push(`Animated image${animation.frameCount ? ` (${animation.frameCount} frames)` : ''}: only the first frame is editable. Save creates a still copy.`)
  }
  if (format === 'svg' && intrinsic && (intrinsic.width !== dimensions.width || intrinsic.height !== dimensions.height)) {
    notices.push(`SVG rasterized at ${dimensions.width} × ${dimensions.height}.`)
  }
  let suggestedName = null
  if (!canSaveInPlace) {
    const saveFormat = saveFormatFor(format)
    suggestedName = animation.animated
      ? nameWithFormat(name, 'png', ' (frame 1)')
      : nameWithFormat(name, saveFormat)
  }
  return {
    bytes,
    format,
    payloadFormat,
    extension,
    extensionFormat,
    mime: FORMATS[format].mime,
    mismatch,
    canSaveInPlace,
    width: dimensions.width,
    height: dimensions.height,
    intrinsic,
    animated: animation.animated,
    frameCount: animation.frameCount,
    notices,
    suggestedName,
  }
}

/** The format an edited copy of a `format` source is saved in by default. */
function saveFormatFor(format) {
  if (format === 'jpeg' || format === 'webp' || format === 'psd') return format
  return 'png'
}

function outputExtension(format) {
  const normalized = String(format || '').toLowerCase()
  if (normalized === 'jpeg' || normalized === 'jpg') return '.jpg'
  if (normalized === 'webp') return '.webp'
  if (normalized === 'png') return '.png'
  if (normalized === 'psd') return '.psd'
  throw new Error('Choose PNG, JPEG, WebP, or PSD as the save format.')
}

/** True when a file with `extension` may hold `format` output without renaming (e.g. JPEG into .jfif). */
function isCompatibleOutputExtension(extension, format) {
  const expected = formatOfExtension(outputExtension(format))
  return formatOfExtension(extension) === expected && expected !== 'svg'
}

function ensureOutputExtension(filePath, format) {
  const expected = outputExtension(format)
  const current = extensionOf(filePath)
  if (current === expected || (current && isCompatibleOutputExtension(current, format))) return filePath
  return `${filePath}${expected}`
}

async function atomicWrite(targetPath, value) {
  const bytes = toBytes(value)
  const directory = path.dirname(targetPath)
  const extension = path.extname(targetPath) || '.tmp'
  const token = crypto.randomUUID()
  const temporary = path.join(directory, `.${path.basename(targetPath, extension)}-${token}${extension}.tmp`)
  const backup = path.join(directory, `.${path.basename(targetPath)}-${token}.bak`)
  let backedUp = false
  await fs.writeFile(temporary, bytes)
  try {
    try {
      await fs.rename(targetPath, backup)
      backedUp = true
    } catch (error) {
      if (error.code !== 'ENOENT') throw error
    }
    await fs.rename(temporary, targetPath)
    if (backedUp) await fs.rm(backup, { force: true })
  } catch (error) {
    await fs.rm(temporary, { force: true }).catch(() => {})
    if (backedUp) await fs.rename(backup, targetPath).catch(() => {})
    throw error
  }
}

module.exports = {
  EDITABLE_EXTENSIONS,
  FORMATS,
  FORMAT_BY_EXTENSION,
  MAX_IMAGE_BYTES,
  MAX_IMAGE_DIMENSION,
  MAX_IMAGE_PIXELS,
  MAX_SVG_BYTES,
  MIME_BY_EXTENSION,
  SUPPORTED_EXTENSIONS,
  atomicWrite,
  decodeSvgText,
  detectAnimation,
  ensureOutputExtension,
  extensionOf,
  formatOfExtension,
  inspectImageBytes,
  isCompatibleOutputExtension,
  isSupportedExtension,
  imageDimensions,
  matchesSignature,
  nameWithFormat,
  outputExtension,
  prepareSvg,
  psdHeader,
  safeStem,
  saveFormatFor,
  sniffFormat,
  svgRasterSize,
  toBytes,
  validateImageDimensions,
  validateImageBytes,
}
