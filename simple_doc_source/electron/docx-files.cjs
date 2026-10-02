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

/** Central-directory entries of a validated, non-ZIP64 package, keyed by name. */
function listZipEntries(data) {
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
  if (cursor > centralOffset + centralSize) throw new Error('This archive does not contain a complete Word document.')
  return entries
}

function validateDocxPackage(data) {
  const bytes = toBytes(data)
  const entries = listZipEntries(bytes)
  if (!entries.has('[Content_Types].xml')) {
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

// ---- Files Simple Docs opens (DOC-017, DOC-SIE-15) --------------------------------
// Word packages open in the editor directly: a template (.dotx/.dotm) as a new untitled
// document, a macro-enabled file without its macros. RTF, OpenDocument text, web pages,
// Markdown and plain text are read by the renderer's native importers (src/importers),
// without LibreOffice or any network access. Everything but a plain .docx saves as a
// separate .docx, so a source file is never overwritten in another format. The
// renderer's drop check (OPENABLE_DOCUMENT in src/main.ts) mirrors this table.

/** Largest file the renderer's importers read (src/importers MAX_IMPORT_BYTES). */
const MAX_IMPORT_BYTES = 100 * 1024 * 1024

const OPEN_FORMATS = Object.freeze({
  '.docx': 'docx',
  '.docm': 'docm',
  '.dotx': 'dotx',
  '.dotm': 'dotm',
  '.doc': 'doc',
  '.rtf': 'rtf',
  '.odt': 'odt',
  '.html': 'html',
  '.htm': 'html',
  '.md': 'md',
  '.markdown': 'md',
  '.txt': 'txt',
})
const WORD_PACKAGE_FORMATS = new Set(['docx', 'docm', 'dotx', 'dotm'])
const IMPORT_FORMATS = new Set(['rtf', 'odt', 'html', 'md', 'txt'])
const TEMPLATE_CONTENT_TYPES = new Set([
  'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml',
  'application/vnd.ms-word.template.macroEnabledTemplate.main+xml',
])

/** Open dialog filters: everything Simple Docs reads first, then each family. */
const OPEN_DIALOG_FILTERS = Object.freeze([
  { name: 'All supported documents', extensions: Object.keys(OPEN_FORMATS).map((extension) => extension.slice(1)) },
  { name: 'Word documents', extensions: ['docx', 'docm', 'doc'] },
  { name: 'Word templates', extensions: ['dotx', 'dotm'] },
  { name: 'OpenDocument text', extensions: ['odt'] },
  { name: 'Rich Text', extensions: ['rtf'] },
  { name: 'Web pages', extensions: ['html', 'htm'] },
  { name: 'Markdown', extensions: ['md', 'markdown'] },
  { name: 'Plain text', extensions: ['txt'] },
  { name: 'All files', extensions: ['*'] },
].map((filter) => Object.freeze({ ...filter, extensions: Object.freeze(filter.extensions) })))

const UNSUPPORTED_FILE_MESSAGE = 'Simple Docs can’t open this kind of file. It opens Word documents and templates, OpenDocument text, Rich Text, web pages, Markdown and plain text.'

/** The format a file name declares (see OPEN_FORMATS), or null. */
function openFormatForName(name) {
  const match = /\.[^.\\/]+$/.exec(String(name ?? ''))
  return match ? OPEN_FORMATS[match[0].toLowerCase()] ?? null : null
}

function isOpenableDocumentPath(filePath) {
  return typeof filePath === 'string' && openFormatForName(filePath) !== null
}

const OLE_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]
const ODT_MIMETYPE = 'application/vnd.oasis.opendocument.text'

/** An OpenDocument package's type: the first ZIP entry is a stored "mimetype" (null when it is not). */
function openDocumentMimetype(bytes) {
  if (bytes.length < 38 || bytes.readUInt32LE(0) !== 0x04034b50 || bytes.readUInt16LE(8) !== 0) return null
  const nameLength = bytes.readUInt16LE(26)
  const extraLength = bytes.readUInt16LE(28)
  const size = bytes.readUInt32LE(18)
  if (nameLength !== 8 || bytes.toString('latin1', 30, 38) !== 'mimetype' || size > 200) return null
  const start = 30 + nameLength + extraLength
  return bytes.toString('latin1', start, Math.min(bytes.length, start + size)).trim() || null
}

/**
 * What a file's bytes are, whatever its name says: 'word' (an OOXML package), 'odt',
 * 'odf' (an OpenDocument package of unknown type), 'zip' (another archive), 'ole'
 * (Word 97-2003 or an encrypted package), 'rtf', 'html', 'text', 'binary' or 'empty'.
 */
function sniffDocumentBytes(data) {
  const bytes = toBytes(data)
  if (!bytes.length) return 'empty'
  if (bytes.length >= 4 && bytes.readUInt32LE(0) === 0x04034b50) {
    const mimetype = openDocumentMimetype(bytes)
    // Text documents and their templates; a spreadsheet or a presentation is not one.
    if (mimetype) return mimetype.startsWith(ODT_MIMETYPE) ? 'odt' : 'zip'
    try {
      const entries = listZipEntries(bytes)
      if (entries.has('[Content_Types].xml')) return 'word'
      // Some writers do not put "mimetype" first; the importer checks the type itself.
      if (entries.has('content.xml') && entries.has('mimetype')) return 'odf'
    } catch {
      // A damaged or partial archive; the name decides how it is reported.
    }
    return 'zip'
  }
  if (OLE_SIGNATURE.every((value, index) => bytes[index] === value)) return 'ole'
  if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0xfe && bytes[1] === 0xff)) {
    const head = bytes.subarray(2, Math.min(bytes.length, 2050))
    const littleEndian = bytes[0] === 0xff
    let text = ''
    for (let index = 0; index + 1 < head.length; index += 2) text += String.fromCharCode(littleEndian ? head[index] | (head[index + 1] << 8) : (head[index] << 8) | head[index + 1])
    return /^\s*<(?:!doctype\s+html|html[\s>])/i.test(text) ? 'html' : /^\s*\{\\rtf/.test(text) ? 'rtf' : 'text'
  }
  const head = bytes.subarray(0, Math.min(bytes.length, 8192))
  const start = head[0] === 0xef && head[1] === 0xbb && head[2] === 0xbf ? 3 : 0
  const text = head.toString('latin1', start, Math.min(head.length, start + 2048))
  if (/^\s*\{\\rtf/.test(text)) return 'rtf'
  // A PDF can be plain ASCII at its start; it is never a text document.
  if (text.startsWith('%PDF-') || head.subarray(start).includes(0)) return 'binary'
  if (/^\s*(?:<\?xml[^>]*>\s*)?(?:<!--[\s\S]*?-->\s*)*<(?:!doctype\s+html|html[\s>]|head[\s>]|body[\s>])/i.test(text)) return 'html'
  return 'text'
}

/** Whether an OOXML package carries macros (vbaProject.bin anywhere in it). */
function packageHasMacros(entries) {
  for (const name of entries.keys()) if (/(?:^|\/)vbaProject\.bin$/i.test(name)) return true
  return false
}

/**
 * The Word package details the open path needs: whether it is a template, whether it
 * has macros. Call after validateDocxBytes.
 */
function wordPackageInfo(data) {
  const bytes = toBytes(data)
  const entries = listZipEntries(bytes)
  let template = false
  const contentTypes = entries.get('[Content_Types].xml')
  if (contentTypes) {
    const xml = decodeXml(readEntry(bytes, contentTypes))
    const overridePattern = /<(?:[A-Za-z_][\w.-]*:)?Override\b([^>]*)>/gs
    for (const match of xml.matchAll(overridePattern)) {
      if (TEMPLATE_CONTENT_TYPES.has(attributesFromXmlTag(match[1]).get('ContentType'))) template = true
    }
  }
  return { template, macros: packageHasMacros(entries) }
}

/**
 * How to open a file: its declared format, corrected by its content. Returns one of
 * OPEN_FORMATS' values ('docx', 'docm', 'dotx', 'dotm', 'doc', 'rtf', 'odt', 'html',
 * 'md', 'txt') and whether the name disagreed with the content ('renamed'). Throws a
 * plain-language error for files Simple Docs cannot read.
 */
function documentOpenKind(fileName, data) {
  const bytes = toBytes(data)
  const declared = openFormatForName(fileName)
  const content = sniffDocumentBytes(bytes)
  if (content === 'empty') throw new Error('This file is empty.')
  // Word 97-2003 files keep their own reader, which also recognizes RTF, web pages,
  // web archives and DOCX packages saved with a .doc name.
  if (declared === 'doc') return { kind: 'doc', renamed: false }
  if (content === 'rtf') return { kind: 'rtf', renamed: declared !== 'rtf' }
  if (content === 'odt') return { kind: 'odt', renamed: declared !== 'odt' }
  if (content === 'odf' && declared === 'odt') return { kind: 'odt', renamed: false }
  if (content === 'word') {
    if (WORD_PACKAGE_FORMATS.has(declared)) return { kind: declared, renamed: false }
    return { kind: 'docx', renamed: true }
  }
  if (content === 'ole') {
    // An encrypted (password-protected) Word package is an OLE file too; its reader explains.
    if (WORD_PACKAGE_FORMATS.has(declared)) return { kind: declared, renamed: false }
    return { kind: 'doc', renamed: true }
  }
  if (content === 'zip' || content === 'odf' || content === 'binary') {
    if (WORD_PACKAGE_FORMATS.has(declared)) return { kind: declared, renamed: false }
    if (IMPORT_FORMATS.has(declared)) throw new Error(`This file is not a readable ${declared === 'odt' ? 'OpenDocument' : 'text'} document.`)
    throw new Error(UNSUPPORTED_FILE_MESSAGE)
  }
  // Text: a web page, Markdown or plain text (or flat OpenDocument XML, or a damaged
  // RTF file, whose importers explain what is wrong).
  if (WORD_PACKAGE_FORMATS.has(declared)) {
    if (content === 'html') return { kind: 'html', renamed: true }
    return { kind: declared, renamed: false }
  }
  if (declared === 'md' || declared === 'txt' || declared === 'odt' || declared === 'rtf') return { kind: declared, renamed: false }
  // A web page, or a file without a known extension (chosen through "All files").
  return { kind: content === 'html' || declared === 'html' ? 'html' : 'txt', renamed: false }
}

module.exports = {
  IMPORT_FORMATS,
  MAIN_DOCUMENT_CONTENT_TYPES,
  MAX_FILE_BYTES,
  MAX_IMPORT_BYTES,
  OPEN_DIALOG_FILTERS,
  OPEN_FORMATS,
  UNSUPPORTED_FILE_MESSAGE,
  WORD_PACKAGE_FORMATS,
  declaredMainDocumentPart,
  documentOpenKind,
  isOpenableDocumentPath,
  listZipEntries,
  openFormatForName,
  sniffDocumentBytes,
  validateDocxBytes,
  validateDocxPackage,
  wordPackageInfo,
}
