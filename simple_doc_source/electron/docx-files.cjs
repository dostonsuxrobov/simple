const { inflateRawSync } = require('node:zlib')

const MAX_FILE_BYTES = 256 * 1024 * 1024
const MAX_ZIP_ENTRIES = 20_000
const MAX_ZIP_ENTRY_BYTES = 512 * 1024 * 1024
const MAX_ZIP_TOTAL_BYTES = 1024 * 1024 * 1024
const MAX_CONTENT_TYPES_BYTES = 4 * 1024 * 1024
const MAIN_DOCUMENT_CONTENT_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
  'application/vnd.ms-word.document.macroEnabled.main+xml',
])

function toBytes(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  if (value && value.type === 'Buffer' && Array.isArray(value.data)) return Buffer.from(value.data)
  return Buffer.from(value)
}

function normalizedEntryName(value) {
  const name = value.replace(/\\/g, '/')
  if (!name || name.startsWith('/') || name.split('/').includes('..')) {
    throw new Error('This DOCX contains an unsafe file path.')
  }
  return name
}

function decodeXml(bytes) {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return bytes.subarray(3).toString('utf8')
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return bytes.subarray(2).toString('utf16le')
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const littleEndian = Buffer.from(bytes.subarray(2))
    littleEndian.swap16()
    return littleEndian.toString('utf16le')
  }
  if (bytes.length >= 4 && bytes[1] === 0 && bytes[3] === 0) return bytes.toString('utf16le')
  if (bytes.length >= 4 && bytes[0] === 0 && bytes[2] === 0) {
    const littleEndian = Buffer.from(bytes)
    littleEndian.swap16()
    return littleEndian.toString('utf16le')
  }
  return bytes.toString('utf8')
}

function decodeXmlEntities(value) {
  return value.replace(/&(#x[0-9a-f]+|#\d+|amp|apos|gt|lt|quot);/gi, (entity, name) => {
    const normalized = name.toLowerCase()
    if (normalized === 'amp') return '&'
    if (normalized === 'apos') return "'"
    if (normalized === 'gt') return '>'
    if (normalized === 'lt') return '<'
    if (normalized === 'quot') return '"'
    const codePoint = normalized.startsWith('#x')
      ? Number.parseInt(normalized.slice(2), 16)
      : Number.parseInt(normalized.slice(1), 10)
    return Number.isSafeInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff
      ? String.fromCodePoint(codePoint)
      : entity
  })
}

function attributesFromXmlTag(source) {
  const attributes = new Map()
  const pattern = /([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/gs
  for (const match of source.matchAll(pattern)) attributes.set(match[1], decodeXmlEntities(match[3]))
  return attributes
}

function declaredMainDocumentPart(contentTypesXml) {
  const overridePattern = /<(?:[A-Za-z_][\w.-]*:)?Override\b([^>]*)>/gs
  for (const match of contentTypesXml.matchAll(overridePattern)) {
    const attributes = attributesFromXmlTag(match[1])
    if (!MAIN_DOCUMENT_CONTENT_TYPES.has(attributes.get('ContentType'))) continue
    const partName = attributes.get('PartName')
    if (!partName) continue
    return normalizedEntryName(partName.startsWith('/') ? partName.slice(1) : partName)
  }
  return null
}

function readEntry(bytes, entry) {
  if (entry.uncompressedSize > MAX_CONTENT_TYPES_BYTES) {
    throw new Error('This DOCX has an unusually large content-types manifest.')
  }
  const offset = entry.localHeaderOffset
  if (offset + 30 > bytes.length || bytes.readUInt32LE(offset) !== 0x04034b50) {
    throw new Error('This DOCX archive has a damaged file entry.')
  }
  const nameLength = bytes.readUInt16LE(offset + 26)
  const extraLength = bytes.readUInt16LE(offset + 28)
  const dataOffset = offset + 30 + nameLength + extraLength
  const dataEnd = dataOffset + entry.compressedSize
  if (dataEnd > bytes.length) throw new Error('This DOCX archive has a truncated file entry.')
  const localName = normalizedEntryName(bytes.subarray(offset + 30, offset + 30 + nameLength).toString('utf8'))
  if (localName !== entry.name) throw new Error('This DOCX archive has a mismatched file entry.')
  const compressed = bytes.subarray(dataOffset, dataEnd)
  let result
  if (entry.method === 0) result = Buffer.from(compressed)
  else if (entry.method === 8) {
    try {
      result = inflateRawSync(compressed, { maxOutputLength: MAX_CONTENT_TYPES_BYTES })
    } catch {
      throw new Error('This DOCX has a damaged content-types manifest.')
    }
  } else {
    throw new Error('This DOCX uses unsupported compression for its content-types manifest.')
  }
  if (result.length !== entry.uncompressedSize) {
    throw new Error('This DOCX has a damaged content-types manifest.')
  }
  return result
}

function validateDocxPackage(data) {
  const bytes = toBytes(data)
  const minimumEocdOffset = Math.max(0, bytes.length - 65_557)
  let eocdOffset = -1
  for (let offset = bytes.length - 22; offset >= minimumEocdOffset; offset -= 1) {
    if (bytes.readUInt32LE(offset) === 0x06054b50) {
      eocdOffset = offset
      break
    }
  }
  if (eocdOffset < 0) throw new Error('This DOCX archive is incomplete.')

  const entryCount = bytes.readUInt16LE(eocdOffset + 10)
  const centralSize = bytes.readUInt32LE(eocdOffset + 12)
  const centralOffset = bytes.readUInt32LE(eocdOffset + 16)
  if (entryCount === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) {
    throw new Error('This DOCX uses an unsupported ZIP64 layout.')
  }
  if (!entryCount || entryCount > MAX_ZIP_ENTRIES || centralOffset + centralSize > bytes.length) {
    throw new Error('This DOCX archive has an unsafe directory structure.')
  }

  const entries = new Map()
  let cursor = centralOffset
  let totalUncompressed = 0
  for (let index = 0; index < entryCount; index += 1) {
    if (cursor + 46 > bytes.length || bytes.readUInt32LE(cursor) !== 0x02014b50) {
      throw new Error('This DOCX archive has a damaged file directory.')
    }
    const flags = bytes.readUInt16LE(cursor + 8)
    const method = bytes.readUInt16LE(cursor + 10)
    const compressedSize = bytes.readUInt32LE(cursor + 20)
    const uncompressedSize = bytes.readUInt32LE(cursor + 24)
    const nameLength = bytes.readUInt16LE(cursor + 28)
    const extraLength = bytes.readUInt16LE(cursor + 30)
    const commentLength = bytes.readUInt16LE(cursor + 32)
    const localHeaderOffset = bytes.readUInt32LE(cursor + 42)
    const nextCursor = cursor + 46 + nameLength + extraLength + commentLength
    if (
      nextCursor > bytes.length
      || compressedSize === 0xffffffff
      || uncompressedSize === 0xffffffff
      || localHeaderOffset === 0xffffffff
    ) {
      throw new Error('This DOCX archive contains an unsupported entry.')
    }
    if ((flags & 0x1) !== 0) throw new Error('Password-protected DOCX files are not supported.')
    if (uncompressedSize > MAX_ZIP_ENTRY_BYTES) throw new Error('This DOCX contains an entry that is too large.')
    totalUncompressed += uncompressedSize
    if (totalUncompressed > MAX_ZIP_TOTAL_BYTES) throw new Error('This DOCX expands beyond the safe size limit.')

    const entryName = normalizedEntryName(bytes.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8'))
    if (entries.has(entryName)) throw new Error('This DOCX archive contains duplicate file entries.')
    entries.set(entryName, { name: entryName, compressedSize, uncompressedSize, method, localHeaderOffset })
    cursor = nextCursor
  }
  if (cursor > centralOffset + centralSize || !entries.has('[Content_Types].xml')) {
    throw new Error('This archive does not contain a complete Word document.')
  }

  const contentTypes = decodeXml(readEntry(bytes, entries.get('[Content_Types].xml')))
  const declaredPart = declaredMainDocumentPart(contentTypes)
  const mainDocumentPart = declaredPart || (entries.has('word/document.xml') ? 'word/document.xml' : null)
  if (!mainDocumentPart || !entries.has(mainDocumentPart)) {
    throw new Error('This archive does not contain its declared main Word document part.')
  }
  return { mainDocumentPart }
}

function validateDocxBytes(data) {
  const bytes = toBytes(data)
  if (bytes.byteLength < 4 || bytes.byteLength > MAX_FILE_BYTES) throw new Error('The DOCX file is empty or too large.')
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) {
    throw new Error('Password-protected Word documents are not supported. Remove the password in Word, then try the .docx file again.')
  }
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('This is not a valid DOCX file.')
  validateDocxPackage(bytes)
  return bytes
}

module.exports = {
  MAIN_DOCUMENT_CONTENT_TYPES,
  MAX_FILE_BYTES,
  declaredMainDocumentPart,
  validateDocxBytes,
  validateDocxPackage,
}
