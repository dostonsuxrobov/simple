const fs = require('node:fs/promises')
const path = require('node:path')
const crypto = require('node:crypto')

const MAX_IMAGE_BYTES = 512 * 1024 * 1024
const MAX_IMAGE_DIMENSION = 20_000
const MAX_IMAGE_PIXELS = 50_000_000
const SUPPORTED_EXTENSIONS = Object.freeze(['.png', '.jpg', '.jpeg', '.webp', '.gif', '.bmp', '.svg', '.avif'])
const EDITABLE_EXTENSIONS = Object.freeze(['.png', '.jpg', '.jpeg', '.webp'])
const MIME_BY_EXTENSION = Object.freeze({
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
  '.gif': 'image/gif',
  '.bmp': 'image/bmp',
  '.svg': 'image/svg+xml',
  '.avif': 'image/avif',
})

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

function svgLength(value) {
  const match = /^\s*([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)\s*(px|in|cm|mm|q|pt|pc)?\s*$/i.exec(String(value || ''))
  if (!match) return null
  const factors = { px: 1, in: 96, cm: 96 / 2.54, mm: 96 / 25.4, q: 96 / 101.6, pt: 96 / 72, pc: 16 }
  return Number(match[1]) * factors[(match[2] || 'px').toLowerCase()]
}

function svgAttribute(tag, name) {
  const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)'|([^\\s>]+))`, 'i').exec(tag)
  return match?.[1] ?? match?.[2] ?? match?.[3] ?? null
}

function svgDimensions(bytes) {
  const text = bytes.subarray(0, Math.min(bytes.length, 64 * 1024)).toString('utf8')
  const tag = /<svg\b[^>]*>/i.exec(text)?.[0]
  if (!tag) return null
  const viewBox = String(svgAttribute(tag, 'viewBox') || '').trim().split(/[\s,]+/).map(Number)
  const viewBoxWidth = viewBox.length === 4 && Number.isFinite(viewBox[2]) ? Math.abs(viewBox[2]) : null
  const viewBoxHeight = viewBox.length === 4 && Number.isFinite(viewBox[3]) ? Math.abs(viewBox[3]) : null
  const width = svgLength(svgAttribute(tag, 'width')) ?? viewBoxWidth ?? 300
  const height = svgLength(svgAttribute(tag, 'height')) ?? viewBoxHeight ?? 150
  return checkedDimensions(width, height)
}

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

function imageDimensions(value, extension) {
  const bytes = toBytes(value)
  const ext = extensionOf(`file${extension.startsWith('.') ? extension : `.${extension}`}`)
  if (ext === '.png' && bytes.length >= 24) return checkedDimensions(bytes.readUInt32BE(16), bytes.readUInt32BE(20))
  if (ext === '.jpg' || ext === '.jpeg') return jpegDimensions(bytes)
  if (ext === '.webp') return webpDimensions(bytes)
  if (ext === '.gif' && bytes.length >= 10) return checkedDimensions(bytes.readUInt16LE(6), bytes.readUInt16LE(8))
  if (ext === '.bmp' && bytes.length >= 26) {
    const dibSize = bytes.readUInt32LE(14)
    return dibSize === 12
      ? checkedDimensions(bytes.readUInt16LE(18), bytes.readUInt16LE(20))
      : checkedDimensions(bytes.readInt32LE(18), bytes.readInt32LE(22))
  }
  if (ext === '.svg') return svgDimensions(bytes)
  if (ext === '.avif') return avifDimensions(bytes)
  return null
}

function validateImageDimensions(value, extension) {
  const dimensions = imageDimensions(value, extension)
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

function matchesSignature(bytes, extension) {
  switch (extension) {
    case '.png':
      return bytes.length >= 8 && bytes.subarray(0, 8).equals(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]))
    case '.jpg':
    case '.jpeg':
      // JPEG decoders permit trailing application bytes after the EOI marker,
      // so the required SOI marker is the compatible content check here.
      return bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff
    case '.webp':
      return bytes.length >= 12 && bytes.subarray(0, 4).toString('ascii') === 'RIFF' && bytes.subarray(8, 12).toString('ascii') === 'WEBP'
    case '.gif':
      return bytes.length >= 10 && ['GIF87a', 'GIF89a'].includes(bytes.subarray(0, 6).toString('ascii'))
    case '.bmp':
      return bytes.length >= 26 && bytes[0] === 0x42 && bytes[1] === 0x4d
    case '.svg': {
      const prefix = bytes.subarray(0, Math.min(bytes.length, 16 * 1024)).toString('utf8').replace(/^\uFEFF/, '')
      return /^(?:(?:\s+)|(?:<\?xml[^>]*>\s*)|(?:<!--[^]*?-->\s*)|(?:<!doctype\s+svg[^>]*>\s*))*<svg(?:\s|>)/i.test(prefix)
    }
    case '.avif': {
      if (bytes.length < 16 || bytes.subarray(4, 8).toString('ascii') !== 'ftyp') return false
      const brands = bytes.subarray(8, Math.min(bytes.length, 64)).toString('ascii')
      return /(?:avif|avis)/.test(brands)
    }
    default:
      return false
  }
}

function validateImageBytes(value, extension) {
  const bytes = toBytes(value)
  const ext = extensionOf(`file${extension.startsWith('.') ? extension : `.${extension}`}`)
  if (!SUPPORTED_EXTENSIONS.includes(ext)) throw new Error('This image format is not supported.')
  if (!bytes.length) throw new Error('This image is empty.')
  if (bytes.length > MAX_IMAGE_BYTES) throw new Error('This image is larger than the 512 MB safety limit.')
  if (!matchesSignature(bytes, ext)) throw new Error(`The file content does not match the ${ext.slice(1).toUpperCase()} format.`)
  return bytes
}

function outputExtension(format) {
  const normalized = String(format || '').toLowerCase()
  if (normalized === 'jpeg' || normalized === 'jpg') return '.jpg'
  if (normalized === 'webp') return '.webp'
  if (normalized === 'png') return '.png'
  throw new Error('Choose PNG, JPEG, or WebP as the save format.')
}

function ensureOutputExtension(filePath, format) {
  const expected = outputExtension(format)
  const current = extensionOf(filePath)
  if (expected === '.jpg' && ['.jpg', '.jpeg'].includes(current)) return filePath
  if (current === expected) return filePath
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
  MAX_IMAGE_BYTES,
  MAX_IMAGE_DIMENSION,
  MAX_IMAGE_PIXELS,
  MIME_BY_EXTENSION,
  SUPPORTED_EXTENSIONS,
  atomicWrite,
  ensureOutputExtension,
  extensionOf,
  isSupportedExtension,
  imageDimensions,
  matchesSignature,
  outputExtension,
  safeStem,
  toBytes,
  validateImageDimensions,
  validateImageBytes,
}
