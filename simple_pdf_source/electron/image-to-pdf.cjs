'use strict'

// Images to PDF pages. PNG and JPEG are embedded as they are (JPEG bytes are
// never recompressed); TIFF is read here, page by page, without a decoder
// dependency: fax (CCITT) and JPEG strips are passed through to the PDF, and
// uncompressed, PackBits, LZW and Deflate pixels are re-packed losslessly.
// Formats only a browser engine decodes (WebP, GIF, BMP, AVIF, ICO, SVG) are
// converted by office-import.cjs before they reach this module.

const zlib = require('node:zlib')

/** Read the bounded EXIF IFD0 Orientation tag; malformed metadata is optional. */
function jpegOrientation(value) {
  const bytes = Buffer.isBuffer(value) ? value : Buffer.from(value)
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return 1
  let offset = 2
  while (offset + 4 <= bytes.length) {
    if (bytes[offset] !== 0xff) return 1
    while (bytes[offset + 1] === 0xff) offset++
    const marker = bytes[offset + 1]
    if (marker === 0xda || marker === 0xd9) return 1
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) { offset += 2; continue }
    if (offset + 4 > bytes.length) return 1
    const size = bytes.readUInt16BE(offset + 2)
    const end = offset + 2 + size
    if (size < 2 || end > bytes.length) return 1
    const start = offset + 4
    if (marker === 0xe1 && end - start >= 14 && bytes.toString('ascii',start,start+6) === 'Exif\0\0') {
      const tiff = start + 6
      const endian = bytes.toString('ascii',tiff,tiff+2)
      if (endian !== 'II' && endian !== 'MM') return 1
      const read16 = (position) => endian === 'II' ? bytes.readUInt16LE(position) : bytes.readUInt16BE(position)
      const read32 = (position) => endian === 'II' ? bytes.readUInt32LE(position) : bytes.readUInt32BE(position)
      if (read16(tiff+2) !== 42) return 1
      const directory = tiff + read32(tiff+4)
      if (directory < tiff+8 || directory+2 > end) return 1
      const count = read16(directory)
      if (count > Math.floor((end-directory-2)/12)) return 1
      for (let index=0;index<count;index++) {
        const entry=directory+2+index*12
        if (read16(entry) !== 0x0112) continue
        if (read16(entry+2) !== 3 || read32(entry+4) !== 1) return 1
        const orientation=read16(entry+8)
        return orientation>=1&&orientation<=8 ? orientation : 1
      }
    }
    offset=end
  }
  return 1
}

/** PDF uses a bottom-left origin; EXIF describes top-left display coordinates. */
function orientationPlacement(orientation,width,height) {
  const matrices = {
    1:[1,0,0,1,0,0], 2:[-1,0,0,1,width,0],
    3:[-1,0,0,-1,width,height], 4:[1,0,0,-1,0,height],
    5:[0,-1,-1,0,height,width], 6:[0,-1,1,0,0,width],
    7:[0,1,1,0,0,0], 8:[0,1,-1,0,height,0],
  }
  const swapped=orientation>=5&&orientation<=8
  return { width:swapped?height:width, height:swapped?width:height, matrix:matrices[orientation]||matrices[1] }
}

function toBuffer(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return Buffer.from(value || [])
}

// ---------------------------------------------------------------------------
// Sniffing
// ---------------------------------------------------------------------------

const BMP_HEADER_SIZES = new Set([12, 16, 40, 52, 56, 64, 108, 124])
const HEIC_BRANDS = new Set(['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'mif1', 'msf1'])

/** Text at the start of a file, past a BOM, XML declaration, comments and DOCTYPE. */
function markupStart(bytes) {
  let text
  if (bytes[0] === 0xff && bytes[1] === 0xfe) text = bytes.subarray(2, 4098).toString('utf16le')
  else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const swapped = Buffer.from(bytes.subarray(2, 4098))
    if (swapped.length % 2) return ''
    text = swapped.swap16().toString('utf16le')
  } else text = bytes.subarray(bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf ? 3 : 0, 4096).toString('latin1')
  let rest = text.replace(/^\s+/, '')
  for (let guard = 0; guard < 20; guard += 1) {
    const before = rest
    rest = rest.replace(/^<\?[\s\S]*?\?>\s*/, '').replace(/^<!--[\s\S]*?-->\s*/, '').replace(/^<!DOCTYPE[^>[]*(\[[\s\S]*?\])?\s*>\s*/i, '')
    if (rest === before) break
  }
  return rest
}

/**
 * The image format of `value` from its first bytes.
 * @returns {'png'|'jpeg'|'gif'|'webp'|'bmp'|'tiff'|'ico'|'avif'|'heic'|'svg'|null}
 */
function sniffImageType(value) {
  const bytes = toBuffer(value)
  if (bytes.length < 4) return null
  if (bytes.length >= 8 && bytes.readUInt32BE(0) === 0x89504e47 && bytes.readUInt32BE(4) === 0x0d0a1a0a) return 'png'
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'jpeg'
  if (bytes.length >= 6 && /^GIF8[79]a$/.test(bytes.toString('latin1', 0, 6))) return 'gif'
  if (bytes.length >= 12 && bytes.toString('latin1', 0, 4) === 'RIFF' && bytes.toString('latin1', 8, 12) === 'WEBP') return 'webp'
  if ((bytes[0] === 0x49 && bytes[1] === 0x49 && (bytes[2] === 0x2a || bytes[2] === 0x2b) && bytes[3] === 0)
    || (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && (bytes[3] === 0x2a || bytes[3] === 0x2b))) return 'tiff'
  if (bytes.length >= 26 && bytes[0] === 0x42 && bytes[1] === 0x4d && BMP_HEADER_SIZES.has(bytes.readUInt32LE(14))
    && bytes.readUInt32LE(10) >= 26 && bytes.readUInt32LE(10) < bytes.length + 1024) return 'bmp'
  if (bytes.length >= 12 && bytes.toString('latin1', 4, 8) === 'ftyp') {
    const brands = [bytes.toString('latin1', 8, 12)]
    const boxSize = bytes.readUInt32BE(0)
    for (let offset = 16; offset + 4 <= Math.min(boxSize, bytes.length, 64); offset += 4) brands.push(bytes.toString('latin1', offset, offset + 4))
    if (brands.includes('avif') || brands.includes('avis')) return 'avif'
    if (brands.some((brand) => HEIC_BRANDS.has(brand))) return 'heic'
  }
  if (bytes.length >= 22 && bytes.readUInt16LE(0) === 0 && bytes.readUInt16LE(2) === 1) {
    const count = bytes.readUInt16LE(4)
    const size = bytes.readUInt32LE(14)
    const offset = bytes.readUInt32LE(18)
    if (count > 0 && count <= 256 && size > 0 && offset >= 6 + count * 16 && offset < bytes.length) return 'ico'
  }
  if (/^<svg[\s>:]/i.test(markupStart(bytes))) return 'svg'
  return null
}

const SVG_UNITS = { px: 0.75, pt: 1, pc: 12, mm: 72 / 25.4, cm: 72 / 2.54, in: 72, q: 72 / 101.6, em: 12, ex: 6 }

function svgLength(value) {
  const match = /^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?)\s*(px|pt|pc|mm|cm|in|q|em|ex)?\s*$/i.exec(String(value ?? ''))
  if (!match) return null
  const number = Number(match[1]) * SVG_UNITS[(match[2] || 'px').toLowerCase()]
  return Number.isFinite(number) && number > 0 ? number : null
}

/**
 * The size, in points, at which an SVG draws: its width and height (CSS
 * units), else its viewBox in CSS pixels, else the 300 x 150 pixel default.
 * @param {string} text the SVG source
 * @returns {{width: number, height: number}}
 */
function svgDimensions(text) {
  const tag = /<svg\b[^>]*>/i.exec(String(text || ''))?.[0] || ''
  const attribute = (name) => new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, 'i').exec(tag)
  const read = (name) => { const match = attribute(name); return match ? (match[1] ?? match[2]) : null }
  let width = svgLength(read('width'))
  let height = svgLength(read('height'))
  const box = String(read('viewBox') || '').trim().split(/[\s,]+/).map(Number)
  const boxWidth = box.length === 4 && box[2] > 0 && Number.isFinite(box[2]) ? box[2] : null
  const boxHeight = box.length === 4 && box[3] > 0 && Number.isFinite(box[3]) ? box[3] : null
  if (boxWidth && boxHeight) {
    if (width && !height) height = width * boxHeight / boxWidth
    else if (height && !width) width = height * boxWidth / boxHeight
    else if (!width && !height) { width = boxWidth * 0.75; height = boxHeight * 0.75 }
  }
  return { width: width || 225, height: height || 112.5 }
}

// ---------------------------------------------------------------------------
// PNG and JPEG
// ---------------------------------------------------------------------------

async function addImagePage(pdfDoc, value, extension, { maxDimension = 841.89 } = {}) {
  const { pushGraphicsState, popGraphicsState, concatTransformationMatrix } = require('pdf-lib')
  // The bytes decide the decoder, so a PNG named .jpg still opens.
  const sniffed = sniffImageType(value)
  const lower = sniffed === 'png' ? '.png' : sniffed === 'jpeg' ? '.jpg' : String(extension).toLowerCase()
  if (!['.png','.jpg','.jpeg'].includes(lower)) throw new Error('Choose a PNG or JPEG image to add to the PDF.')
  // pdf-lib's JPEG parser reads from buffer offset zero, so pooled Node buffers
  // and sliced typed arrays must be copied into an exact independent byte view.
  const bytes=new Uint8Array(value)
  const image=lower==='.png' ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes)
  if (!(image.width>0&&image.height>0)) throw new Error('The image dimensions are invalid.')
  const limit=Number.isFinite(maxDimension)&&maxDimension>0 ? maxDimension : 841.89
  const scale=Math.min(1,limit/image.width,limit/image.height)
  const width=image.width*scale
  const height=image.height*scale
  const placement=orientationPlacement(lower==='.png'?1:jpegOrientation(bytes),width,height)
  const page=pdfDoc.addPage([placement.width,placement.height])
  page.pushOperators(pushGraphicsState(),concatTransformationMatrix(...placement.matrix))
  page.drawImage(image,{x:0,y:0,width,height})
  page.pushOperators(popGraphicsState())
  return page
}

// ---------------------------------------------------------------------------
// TIFF
// ---------------------------------------------------------------------------

const TIFF_TYPE_SIZES = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8, 13: 4 }
const MAX_TIFF_PAGES = 2000
const MAX_DECODED_BYTES = 768 * 1024 * 1024

function tiffError(message) {
  const error = new Error(message)
  error.name = 'IMAGE_DAMAGED'
  error.code = 'IMAGE_DAMAGED'
  return error
}

function unsupportedTiff(detail) {
  const error = new Error(`This TIFF image uses ${detail}, which Simple can't read. Save it as PNG, JPEG or a standard TIFF, then try again.`)
  error.name = 'UNSUPPORTED_IMAGE'
  error.code = 'UNSUPPORTED_IMAGE'
  return error
}

/** Reads every image directory of a TIFF file (thumbnails are skipped). */
function readTiffDirectories(bytes) {
  if (bytes.length < 8) throw tiffError('This TIFF image is damaged.')
  const little = bytes[0] === 0x49
  if (bytes.readUInt16LE(2) === 43 || bytes.readUInt16BE(2) === 43) throw unsupportedTiff('the BigTIFF layout')
  const u16 = (offset) => little ? bytes.readUInt16LE(offset) : bytes.readUInt16BE(offset)
  const u32 = (offset) => little ? bytes.readUInt32LE(offset) : bytes.readUInt32BE(offset)
  const valueAt = (type, offset) => {
    switch (type) {
      case 1: case 2: case 7: return bytes[offset]
      case 6: return bytes.readInt8(offset)
      case 3: return u16(offset)
      case 8: return little ? bytes.readInt16LE(offset) : bytes.readInt16BE(offset)
      case 4: case 13: return u32(offset)
      case 9: return little ? bytes.readInt32LE(offset) : bytes.readInt32BE(offset)
      case 5: { const denominator = u32(offset + 4); return denominator ? u32(offset) / denominator : 0 }
      case 10: { const denominator = little ? bytes.readInt32LE(offset + 4) : bytes.readInt32BE(offset + 4); return denominator ? (little ? bytes.readInt32LE(offset) : bytes.readInt32BE(offset)) / denominator : 0 }
      case 11: return little ? bytes.readFloatLE(offset) : bytes.readFloatBE(offset)
      case 12: return little ? bytes.readDoubleLE(offset) : bytes.readDoubleBE(offset)
      default: return 0
    }
  }
  const directories = []
  const seen = new Set()
  let offset = u32(4)
  while (offset && !seen.has(offset) && directories.length < MAX_TIFF_PAGES) {
    seen.add(offset)
    if (offset + 2 > bytes.length) break
    const count = u16(offset)
    if (offset + 2 + count * 12 + 4 > bytes.length) throw tiffError('This TIFF image is damaged.')
    const tags = new Map()
    for (let index = 0; index < count; index += 1) {
      const entry = offset + 2 + index * 12
      const tag = u16(entry)
      const type = u16(entry + 2)
      const total = u32(entry + 4)
      const size = TIFF_TYPE_SIZES[type]
      if (!size || total > 64 * 1024 * 1024) continue
      const inline = size * total <= 4
      const start = inline ? entry + 8 : u32(entry + 8)
      if (start + size * total > bytes.length) continue
      if (tag === 347) { tags.set(tag, bytes.subarray(start, start + total)); continue }
      const values = new Array(Math.min(total, 1 << 22))
      for (let item = 0; item < values.length; item += 1) values[item] = valueAt(type, start + item * size)
      tags.set(tag, values)
    }
    const next = u32(offset + 2 + count * 12)
    const subfile = tags.get(254)?.[0] || 0
    if (!(subfile & 1)) directories.push(tags)
    offset = next
  }
  if (!directories.length) throw tiffError('This TIFF image has no pages.')
  return directories
}

function packBits(input, expected) {
  const output = Buffer.alloc(expected)
  let read = 0
  let written = 0
  while (read < input.length && written < expected) {
    const header = input.readInt8(read++)
    if (header >= 0) {
      const length = Math.min(header + 1, expected - written, input.length - read)
      input.copy(output, written, read, read + length)
      read += header + 1
      written += length
    } else if (header !== -128) {
      const length = Math.min(1 - header, expected - written)
      output.fill(input[read++] ?? 0, written, written + length)
      written += length
    }
  }
  return output
}

function lzw(input, expected) {
  const output = Buffer.alloc(expected)
  const prefix = new Int32Array(4096)
  const suffix = new Uint8Array(4096)
  const lengths = new Uint16Array(4096)
  for (let code = 0; code < 256; code += 1) { prefix[code] = -1; suffix[code] = code; lengths[code] = 1 }
  let written = 0
  let width = 9
  let next = 258
  let previous = -1
  let bitBuffer = 0
  let bitCount = 0
  let read = 0
  const stack = new Uint8Array(4096)
  const emit = (code) => {
    let length = lengths[code]
    let cursor = code
    for (let index = length - 1; index >= 0; index -= 1) { stack[index] = suffix[cursor]; cursor = prefix[cursor] }
    const room = Math.min(length, expected - written)
    for (let index = 0; index < room; index += 1) output[written++] = stack[index]
    return stack[0]
  }
  while (written < expected) {
    while (bitCount < width && read < input.length) { bitBuffer = ((bitBuffer << 8) | input[read++]) >>> 0; bitCount += 8 }
    if (bitCount < width) break
    const code = (bitBuffer >>> (bitCount - width)) & ((1 << width) - 1)
    bitCount -= width
    bitBuffer &= (1 << bitCount) - 1
    if (code === 257) break
    if (code === 256) { width = 9; next = 258; previous = -1; continue }
    if (previous === -1) {
      if (code > 255) break
      emit(code)
      previous = code
      continue
    }
    let first
    if (code < next) {
      first = emit(code)
      if (next < 4096) { prefix[next] = previous; suffix[next] = first; lengths[next] = lengths[previous] + 1; next += 1 }
    } else if (code === next && next < 4096) {
      let cursor = previous
      while (prefix[cursor] !== -1) cursor = prefix[cursor]
      first = suffix[cursor]
      prefix[next] = previous; suffix[next] = first; lengths[next] = lengths[previous] + 1; next += 1
      emit(code)
    } else break
    previous = code
    // TIFF LZW widens the code one entry early.
    if (next + 1 >= (1 << width) && width < 12) width += 1
  }
  return output
}

function reverseBits(input) {
  const output = Buffer.allocUnsafe(input.length)
  for (let index = 0; index < input.length; index += 1) {
    let value = input[index]
    value = ((value & 0xf0) >> 4) | ((value & 0x0f) << 4)
    value = ((value & 0xcc) >> 2) | ((value & 0x33) << 2)
    value = ((value & 0xaa) >> 1) | ((value & 0x55) << 1)
    output[index] = value
  }
  return output
}

/** Joins a JPEGTables segment and a strip that relies on it into one JPEG. */
function completeJpeg(strip, tables) {
  if (!tables || tables.length < 4 || strip[0] !== 0xff || strip[1] !== 0xd8) return strip
  const end = tables[tables.length - 2] === 0xff && tables[tables.length - 1] === 0xd9 ? tables.length - 2 : tables.length
  return Buffer.concat([tables.subarray(0, end), strip.subarray(2)])
}

function tiffResolution(tags) {
  const unit = tags.get(296)?.[0] ?? 2
  const perInch = unit === 3 ? 2.54 : unit === 2 ? 1 : 0
  const x = Number(tags.get(282)?.[0]) * perInch
  const y = Number(tags.get(283)?.[0]) * perInch
  if (!(x >= 20 && x <= 10_000)) return null
  return { x, y: y >= 20 && y <= 10_000 ? y : x }
}

/**
 * One TIFF page as PDF image pieces: either decoded pixels for a single image
 * XObject, or pass-through CCITT/JPEG strips and tiles placed side by side.
 */
function tiffPage(bytes, tags) {
  const get = (tag, fallback) => tags.get(tag)?.[0] ?? fallback
  const width = get(256, 0)
  const height = get(257, 0)
  if (!(width > 0 && height > 0 && width <= 1 << 20 && height <= 1 << 20)) throw tiffError('This TIFF image is damaged.')
  const compression = get(259, 1)
  const samples = get(277, 1)
  const bits = tags.get(258) || [1]
  const bitsPerSample = bits[0] || 1
  const photometric = get(262, samples >= 3 ? 2 : 1)
  const planar = get(284, 1)
  const predictor = get(317, 1)
  const fillOrder = get(266, 1)
  const sampleFormat = get(339, 1)
  const extra = tags.get(338) || []
  const tiled = tags.has(322) && tags.has(324)
  const pieceWidth = tiled ? get(322, width) : width
  const pieceHeight = tiled ? get(323, height) : Math.min(get(278, height) || height, height)
  const offsets = tags.get(tiled ? 324 : 273) || []
  const counts = tags.get(tiled ? 325 : 279) || []
  if (!offsets.length) throw tiffError('This TIFF image is damaged.')
  const across = tiled ? Math.ceil(width / pieceWidth) : 1
  const down = Math.ceil(height / pieceHeight)
  const planes = planar === 2 ? samples : 1
  if (offsets.length < across * down * planes) throw tiffError('This TIFF image is damaged.')
  const piece = (index) => {
    const start = offsets[index]
    const length = counts[index] ?? (bytes.length - start)
    if (!(start >= 0 && start < bytes.length)) return Buffer.alloc(0)
    return bytes.subarray(start, Math.min(bytes.length, start + length))
  }
  const page = { width, height, resolution: tiffResolution(tags), orientation: get(274, 1) }

  if (sampleFormat === 3) throw unsupportedTiff('floating-point samples')
  if (compression === 2 || compression === 3 || compression === 4) {
    if (samples !== 1 || bitsPerSample !== 1) throw tiffError('This TIFF image is damaged.')
    const options = get(compression === 4 ? 293 : 292, 0)
    if (compression === 3 && (options & 2)) throw unsupportedTiff('uncompressed fax mode')
    const parameters = {
      K: compression === 4 ? -1 : compression === 3 && (options & 1) ? 1 : 0,
      Columns: pieceWidth,
      BlackIs1: false,
      EncodedByteAlign: compression === 2 || (compression === 3 && Boolean(options & 4)),
      EndOfBlock: false,
    }
    page.pieces = []
    for (let row = 0; row < down; row += 1) {
      for (let column = 0; column < across; column += 1) {
        const index = row * across + column
        const rows = tiled ? pieceHeight : Math.min(pieceHeight, height - row * pieceHeight)
        let data = piece(index)
        if (fillOrder === 2) data = reverseBits(data)
        page.pieces.push({
          x: column * pieceWidth, y: row * pieceHeight, width: pieceWidth, height: rows,
          data,
          dict: { ColorSpace: 'DeviceGray', BitsPerComponent: 1, Filter: 'CCITTFaxDecode', DecodeParms: { ...parameters, Rows: rows }, ...(photometric === 1 ? { Decode: [1, 0] } : {}) },
        })
      }
    }
    return page
  }
  if (compression === 7) {
    if (planar !== 1 || bitsPerSample !== 8 || ![1, 3, 4].includes(samples)) throw unsupportedTiff('an unusual JPEG layout')
    const tables = tags.get(347)
    const colorSpace = samples === 1 ? 'DeviceGray' : samples === 4 ? 'DeviceCMYK' : 'DeviceRGB'
    page.pieces = []
    for (let row = 0; row < down; row += 1) {
      for (let column = 0; column < across; column += 1) {
        const index = row * across + column
        const rows = tiled ? pieceHeight : Math.min(pieceHeight, height - row * pieceHeight)
        const data = completeJpeg(piece(index), Buffer.isBuffer(tables) ? tables : null)
        const dict = { ColorSpace: colorSpace, BitsPerComponent: 8, Filter: 'DCTDecode' }
        if (samples === 3 && photometric === 2) dict.DecodeParms = { ColorTransform: 0 }
        if (samples === 1 && photometric === 0) dict.Decode = [1, 0]
        page.pieces.push({ x: column * pieceWidth, y: row * pieceHeight, width: pieceWidth, height: rows, data, dict })
      }
    }
    return page
  }
  if (![1, 5, 8, 32946, 32773].includes(compression)) throw unsupportedTiff(compression === 6 ? 'old-style JPEG compression' : `compression type ${compression}`)
  if (![1, 2, 4, 8, 16].includes(bitsPerSample) || bits.some((value) => value !== bitsPerSample)) throw unsupportedTiff(`${bitsPerSample}-bit samples`)
  if (photometric === 6) throw unsupportedTiff('uncompressed YCbCr colour')
  if (![0, 1, 2, 3, 5].includes(photometric)) throw unsupportedTiff('this colour model')
  const colorSamples = photometric === 2 ? 3 : photometric === 5 ? 4 : 1
  if (samples < colorSamples) throw tiffError('This TIFF image is damaged.')
  const alphaIndex = samples > colorSamples && (extra[0] === 1 || extra[0] === 2) && photometric !== 3 ? colorSamples : -1
  const unpremultiply = extra[0] === 1
  if (bitsPerSample < 8 && samples !== 1) throw unsupportedTiff('packed colour samples')

  // Rows of the whole image, one plane per sample when planar.
  const sampleBytes = bitsPerSample === 16 ? 2 : 1
  const rowBytes = (pixels, perPixel) => Math.ceil(pixels * perPixel * bitsPerSample / 8)
  const perPixel = planar === 2 ? 1 : samples
  const fullRow = rowBytes(width, perPixel)
  if (fullRow * height * planes > MAX_DECODED_BYTES) throw unsupportedTiff('an image too large to convert')
  const planesData = Array.from({ length: planes }, () => Buffer.alloc(fullRow * height))
  const little = bytes[0] === 0x49
  for (let plane = 0; plane < planes; plane += 1) {
    for (let row = 0; row < down; row += 1) {
      for (let column = 0; column < across; column += 1) {
        const index = plane * across * down + row * across + column
        const pieceRow = rowBytes(pieceWidth, perPixel)
        const expected = pieceRow * pieceHeight
        const raw = piece(index)
        let data
        // Undoing the predictor writes in place, so never into the caller's bytes.
        if (compression === 1) data = predictor === 2 ? Buffer.from(raw) : raw
        else if (compression === 32773) data = packBits(raw, expected)
        else if (compression === 5) data = lzw(raw, expected)
        else {
          try { data = zlib.inflateSync(raw, { finishFlush: zlib.constants.Z_SYNC_FLUSH }) } catch { data = Buffer.alloc(expected) }
        }
        if (predictor === 2 && bitsPerSample >= 8) {
          for (let line = 0; line < pieceHeight; line += 1) {
            const base = line * pieceRow
            if (sampleBytes === 1) {
              for (let index2 = perPixel; index2 < pieceRow && base + index2 < data.length; index2 += 1) data[base + index2] = (data[base + index2] + data[base + index2 - perPixel]) & 0xff
            } else {
              for (let index2 = perPixel * 2; index2 + 1 < pieceRow && base + index2 + 1 < data.length; index2 += 2) {
                const read = (at) => little ? data.readUInt16LE(at) : data.readUInt16BE(at)
                const value = (read(base + index2) + read(base + index2 - perPixel * 2)) & 0xffff
                if (little) data.writeUInt16LE(value, base + index2)
                else data.writeUInt16BE(value, base + index2)
              }
            }
          }
        } else if (predictor === 3) throw unsupportedTiff('a floating-point predictor')
        const rows = Math.min(pieceHeight, height - row * pieceHeight)
        const xBytes = rowBytes(column * pieceWidth, perPixel)
        const copyBytes = Math.min(pieceRow, fullRow - xBytes)
        if (bitsPerSample < 8 && tiled && xBytes * 8 !== column * pieceWidth * perPixel * bitsPerSample) throw unsupportedTiff('unaligned bit-packed tiles')
        for (let line = 0; line < rows; line += 1) {
          const from = line * pieceRow
          if (from >= data.length) break
          data.copy(planesData[plane], (row * pieceHeight + line) * fullRow + xBytes, from, Math.min(data.length, from + copyBytes))
        }
      }
    }
  }

  // 16-bit samples become 8-bit; planes are interleaved into pixels.
  const toByte = (buffer, at) => sampleBytes === 2 ? buffer[little ? at + 1 : at] : buffer[at]
  if (bitsPerSample < 8) {
    if (photometric === 3) {
      const map = tags.get(320) || []
      const entries = 1 << bitsPerSample
      const lookup = Buffer.alloc(entries * 3)
      for (let index = 0; index < entries; index += 1) for (let channel = 0; channel < 3; channel += 1) lookup[index * 3 + channel] = (map[channel * entries + index] || 0) >> 8
      page.pixels = { data: planesData[0], dict: { ColorSpace: ['Indexed', 'DeviceRGB', entries - 1, lookup], BitsPerComponent: bitsPerSample } }
      return page
    }
    page.pixels = { data: planesData[0], dict: { ColorSpace: 'DeviceGray', BitsPerComponent: bitsPerSample, ...(photometric === 0 ? { Decode: [1, 0] } : {}) } }
    return page
  }
  const pixelCount = width * height
  const color = Buffer.alloc(pixelCount * colorSamples)
  const alpha = alphaIndex >= 0 ? Buffer.alloc(pixelCount) : null
  for (let row = 0; row < height; row += 1) {
    for (let x = 0; x < width; x += 1) {
      const pixel = row * width + x
      const sampleAt = (sample) => planar === 2
        ? toByte(planesData[sample], row * fullRow + x * sampleBytes)
        : toByte(planesData[0], row * fullRow + (x * samples + sample) * sampleBytes)
      for (let sample = 0; sample < colorSamples; sample += 1) color[pixel * colorSamples + sample] = sampleAt(sample)
      if (alpha) {
        const value = sampleAt(alphaIndex)
        alpha[pixel] = value
        if (unpremultiply && value > 0 && value < 255) {
          for (let sample = 0; sample < colorSamples; sample += 1) color[pixel * colorSamples + sample] = Math.min(255, Math.round(color[pixel * colorSamples + sample] * 255 / value))
        }
      }
    }
  }
  if (photometric === 3) {
    const map = tags.get(320) || []
    const entries = 256
    const lookup = Buffer.alloc(entries * 3)
    for (let index = 0; index < entries; index += 1) for (let channel = 0; channel < 3; channel += 1) lookup[index * 3 + channel] = (map[channel * entries + index] || 0) >> 8
    page.pixels = { data: color, dict: { ColorSpace: ['Indexed', 'DeviceRGB', 255, lookup], BitsPerComponent: 8 } }
  } else {
    const colorSpace = colorSamples === 3 ? 'DeviceRGB' : colorSamples === 4 ? 'DeviceCMYK' : 'DeviceGray'
    page.pixels = { data: color, dict: { ColorSpace: colorSpace, BitsPerComponent: 8, ...(photometric === 0 ? { Decode: [1, 0] } : {}) } }
  }
  if (alpha && alpha.some((value) => value !== 255)) page.alpha = alpha
  return page
}

function colorSpaceObject(context, value) {
  const { PDFHexString } = require('pdf-lib')
  if (!Array.isArray(value)) return value
  const [kind, base, maximum, lookup] = value
  return context.obj([kind, base, maximum, PDFHexString.of(lookup.toString('hex'))])
}

/**
 * Adds one PDF page per TIFF page. Pages keep their physical size when the
 * file records a resolution; otherwise they are sized like other images.
 * @returns {Promise<object[]>} the added pdf-lib pages
 */
async function addTiffPages(pdfDoc, value, { maxDimension = 841.89 } = {}) {
  const {
    pushGraphicsState, popGraphicsState, concatTransformationMatrix, drawObject, rectangle, clip, endPath,
  } = require('pdf-lib')
  const bytes = toBuffer(value)
  const context = pdfDoc.context
  const added = []
  for (const tags of readTiffDirectories(bytes)) {
    const page = tiffPage(bytes, tags)
    let width
    let height
    if (page.resolution) {
      width = page.width / page.resolution.x * 72
      height = page.height / page.resolution.y * 72
      const scale = Math.min(1, 14_400 / width, 14_400 / height)
      width *= scale
      height *= scale
    } else {
      const limit = Number.isFinite(maxDimension) && maxDimension > 0 ? maxDimension : 841.89
      const scale = Math.min(1, limit / page.width, limit / page.height)
      width = page.width * scale
      height = page.height * scale
    }
    const placement = orientationPlacement(page.orientation >= 1 && page.orientation <= 8 ? page.orientation : 1, width, height)
    const pdfPage = pdfDoc.addPage([placement.width, placement.height])
    const operators = [pushGraphicsState(), concatTransformationMatrix(...placement.matrix)]
    if (page.pixels) {
      const dict = { Type: 'XObject', Subtype: 'Image', Width: page.width, Height: page.height, ...page.pixels.dict }
      dict.ColorSpace = colorSpaceObject(context, dict.ColorSpace)
      if (page.alpha) {
        dict.SMask = context.register(context.flateStream(page.alpha, { Type: 'XObject', Subtype: 'Image', Width: page.width, Height: page.height, ColorSpace: 'DeviceGray', BitsPerComponent: 8 }))
      }
      const name = pdfPage.node.newXObject('Image', context.register(context.flateStream(page.pixels.data, dict)))
      operators.push(concatTransformationMatrix(width, 0, 0, height, 0, 0), drawObject(name))
    } else {
      // Strips and tiles are placed in image pixels; edge tiles are clipped.
      const scaleX = width / page.width
      const scaleY = height / page.height
      operators.push(rectangle(0, 0, width, height), clip(), endPath())
      for (const piece of page.pieces) {
        const dict = { Type: 'XObject', Subtype: 'Image', Width: piece.width, Height: piece.height, ...piece.dict }
        const name = pdfPage.node.newXObject('Image', context.register(context.stream(piece.data, dict)))
        // A hair of overlap hides seams between strips in anti-aliased viewers.
        const overlap = page.pieces.length > 1 ? Math.min(0.05, scaleY / 4) : 0
        const bottom = height - (piece.y + piece.height) * scaleY
        operators.push(
          pushGraphicsState(),
          concatTransformationMatrix(piece.width * scaleX, 0, 0, piece.height * scaleY + overlap, piece.x * scaleX, bottom - overlap / 2),
          drawObject(name),
          popGraphicsState(),
        )
      }
    }
    operators.push(popGraphicsState())
    pdfPage.pushOperators(...operators)
    added.push(pdfPage)
  }
  return added
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let index = 0; index < 256; index += 1) {
    let value = index
    for (let bit = 0; bit < 8; bit += 1) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1
    table[index] = value >>> 0
  }
  return table
})()

function crc32(...parts) {
  let value = 0xffffffff
  for (const part of parts) for (const byte of part) value = CRC_TABLE[(value ^ byte) & 0xff] ^ (value >>> 8)
  return (value ^ 0xffffffff) >>> 0
}

/** An 8-bit gray (1), gray+alpha (2), RGB (3) or RGBA (4) PNG. */
function encodePng(width, height, channels, pixels) {
  const stride = width * channels
  const raw = Buffer.alloc((stride + 1) * height)
  for (let row = 0; row < height; row += 1) pixels.copy(raw, row * (stride + 1) + 1, row * stride, (row + 1) * stride)
  const chunk = (type, body) => {
    const head = Buffer.alloc(8)
    head.writeUInt32BE(body.length, 0)
    head.write(type, 4, 'latin1')
    const tail = Buffer.alloc(4)
    tail.writeUInt32BE(crc32(head.subarray(4), body), 0)
    return Buffer.concat([head, body, tail])
  }
  const header = Buffer.alloc(13)
  header.writeUInt32BE(width, 0)
  header.writeUInt32BE(height, 4)
  header[8] = 8
  header[9] = { 1: 0, 2: 4, 3: 2, 4: 6 }[channels]
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', zlib.deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

/**
 * The first page of a TIFF as a PNG (or its JPEG when it is one JPEG strip),
 * for placing it as a picture on a page.
 * @returns {{data: Buffer, mime: 'image/png'|'image/jpeg', width: number, height: number}}
 */
function tiffFirstImage(value) {
  const bytes = toBuffer(value)
  const page = tiffPage(bytes, readTiffDirectories(bytes)[0])
  const { width, height } = page
  if (!page.pixels) {
    const piece = page.pieces.length === 1 ? page.pieces[0] : null
    if (piece?.dict.Filter === 'DCTDecode' && piece.dict.ColorSpace !== 'DeviceCMYK' && !piece.dict.DecodeParms && !piece.dict.Decode) {
      return { data: piece.data, mime: 'image/jpeg', width, height }
    }
    throw unsupportedTiff('fax or striped JPEG compression for a picture (add it as pages instead)')
  }
  const { data, dict } = page.pixels
  const bits = dict.BitsPerComponent
  const invert = Array.isArray(dict.Decode) && dict.Decode[0] === 1
  const indexed = Array.isArray(dict.ColorSpace) ? dict.ColorSpace[3] : null
  const sourceChannels = indexed ? 1 : dict.ColorSpace === 'DeviceRGB' ? 3 : dict.ColorSpace === 'DeviceCMYK' ? 4 : 1
  const outputChannels = (indexed || sourceChannels >= 3 ? 3 : 1) + (page.alpha ? 1 : 0)
  const output = Buffer.alloc(width * height * outputChannels)
  const rowBytes = Math.ceil(width * sourceChannels * bits / 8)
  const maximum = (1 << bits) - 1
  for (let row = 0; row < height; row += 1) {
    for (let x = 0; x < width; x += 1) {
      const at = (row * width + x) * outputChannels
      let channels
      if (bits < 8) {
        const bit = x * bits
        const raw = (data[row * rowBytes + (bit >> 3)] >> (8 - bits - (bit & 7))) & maximum
        channels = indexed ? [...indexed.subarray(raw * 3, raw * 3 + 3)] : [Math.round(raw * 255 / maximum)]
      } else {
        const start = row * rowBytes + x * sourceChannels
        if (indexed) channels = [...indexed.subarray(data[start] * 3, data[start] * 3 + 3)]
        else if (sourceChannels === 4) {
          const black = data[start + 3]
          channels = [0, 1, 2].map((channel) => Math.round((255 - data[start + channel]) * (255 - black) / 255))
        } else channels = [...data.subarray(start, start + sourceChannels)]
      }
      if (invert && !indexed) channels = channels.map((channel) => 255 - channel)
      for (let channel = 0; channel < channels.length; channel += 1) output[at + channel] = channels[channel]
      if (page.alpha) output[at + outputChannels - 1] = page.alpha[row * width + x]
    }
  }
  return { data: encodePng(width, height, outputChannels, output), mime: 'image/png', width, height }
}

/**
 * Adds the pages of a PNG, JPEG or TIFF image (one per TIFF page).
 * @returns {Promise<object[]>} the added pdf-lib pages
 */
async function addImagePages(pdfDoc, value, extension = '', options = {}) {
  const type = sniffImageType(value)
  if (type === 'tiff' || (!type && /^\.?tiff?$/i.test(String(extension)))) return addTiffPages(pdfDoc, value, options)
  return [await addImagePage(pdfDoc, value, extension, options)]
}

async function imageToPdfBytes(buffer,extension,title='Converted image') {
  const {PDFDocument}=require('pdf-lib')
  const pdfDoc=await PDFDocument.create()
  await addImagePages(pdfDoc,buffer,extension)
  pdfDoc.setTitle(title)
  pdfDoc.setCreator('simple')
  return pdfDoc.save()
}

module.exports={addImagePage,addImagePages,addTiffPages,encodePng,imageToPdfBytes,jpegOrientation,orientationPlacement,sniffImageType,svgDimensions,tiffFirstImage}
