// Vendored from simple/shared/electron/validators.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// Dependency-free structural validators. safeWriteFile() runs one on every
// temp file before it replaces a user's file, so a serializer bug or a
// truncated write can never overwrite good work with a broken file.
//
// They check structure, not meaning: a ZIP package must have an intact
// central directory, local headers and CRCs; a PDF must start with %PDF- and
// end with a startxref/%%EOF trailer; images must reach their end marker.
// Deeper checks that need a library (pdf-lib load, SheetJS reopen,
// validateDocxBytes) are injected by the workspace through `validate`.
//
// Validators accept a validator id ('ooxml-word'), a registry format id
// ('docx') or an extension ('.docx'); formats.json maps the last two.

const zlib = require('node:zlib')

/** Validator ids, as used by the "validator" field of formats.json. */
const VALIDATOR_IDS = Object.freeze([
  'none', 'pdf', 'zip', 'ooxml-word', 'ooxml-excel', 'ooxml-powerpoint', 'odf', 'cfb',
  'png', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'ico', 'text', 'json',
])

const MAX_ZIP_ENTRIES = 20_000
// Inflating is synchronous, so optional entries stop being CRC-checked after
// this many bytes; required parts are always checked.
const DEFAULT_CRC_BUDGET = 256 * 1024 * 1024
// Text up to this size is decoded whole (for source-text and markup checks);
// larger text is checked in streamed chunks so it never hits string limits.
const MAX_WHOLE_TEXT_BYTES = 64 * 1024 * 1024
const TEXT_CHUNK_BYTES = 8 * 1024 * 1024
const TEXT_TAIL_BYTES = 64 * 1024
const ODF_MEDIA_TYPES = Object.freeze({
  odt: 'application/vnd.oasis.opendocument.text',
  ott: 'application/vnd.oasis.opendocument.text-template',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  ots: 'application/vnd.oasis.opendocument.spreadsheet-template',
  odp: 'application/vnd.oasis.opendocument.presentation',
  otp: 'application/vnd.oasis.opendocument.presentation-template',
  odg: 'application/vnd.oasis.opendocument.graphics',
})
const OOXML_KINDS = Object.freeze({
  'ooxml-word': { label: 'a Word document', parts: ['word/document.xml'], folder: 'word/', contentType: /wordprocessingml|ms-word/i },
  'ooxml-excel': { label: 'an Excel workbook', parts: ['xl/workbook.xml', 'xl/workbook.bin'], folder: 'xl/', contentType: /spreadsheetml|ms-excel/i },
  'ooxml-powerpoint': { label: 'a PowerPoint presentation', parts: ['ppt/presentation.xml'], folder: 'ppt/', contentType: /presentationml|ms-powerpoint/i },
})
const TEXT_ENCODING_ALIASES = Object.freeze({
  utf8: 'utf-8',
  'utf-8-bom': 'utf-8',
  'utf-16': 'utf-16le',
  utf16le: 'utf-16le',
  'utf-16le': 'utf-16le',
  utf16be: 'utf-16be',
  'utf-16be': 'utf-16be',
  ucs2: 'utf-16le',
  'ucs-2': 'utf-16le',
  latin1: 'windows-1252',
  'iso-8859-1': 'windows-1252',
  ascii: 'windows-1252',
  'us-ascii': 'windows-1252',
})
const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
const CFB_SIGNATURE = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
const CFB_END_OF_CHAIN = 0xfffffffe
const CFB_FREE = 0xffffffff

/** Thrown inside a validator; validateBytes() turns it into {ok:false, reason}. */
class StructureError extends Error {
  constructor(reason) {
    super(reason)
    this.name = 'StructureError'
  }
}

function fail(reason) {
  throw new StructureError(reason)
}

function toBuffer(bytes) {
  if (Buffer.isBuffer(bytes)) return bytes
  if (bytes instanceof Uint8Array) return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes instanceof ArrayBuffer) return Buffer.from(bytes)
  if (typeof bytes === 'string') return Buffer.from(bytes, 'utf8')
  throw new TypeError('Validators need a Buffer, Uint8Array, ArrayBuffer or string.')
}

let crcTable = null
/**
 * CRC-32 (ZIP/PNG polynomial). Uses zlib.crc32 when the runtime has it.
 * @param {Buffer} data
 * @returns {number}
 */
function crc32(data) {
  if (typeof zlib.crc32 === 'function') return zlib.crc32(data) >>> 0
  if (!crcTable) {
    crcTable = new Uint32Array(256)
    for (let n = 0; n < 256; n += 1) {
      let c = n
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      crcTable[n] = c >>> 0
    }
  }
  let crc = 0xffffffff
  for (let index = 0; index < data.length; index += 1) crc = crcTable[(crc ^ data[index]) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

// ---------------------------------------------------------------------------
// Format id resolution
// ---------------------------------------------------------------------------

let registryMaps = null
function registry() {
  if (registryMaps) return registryMaps
  const byId = new Map()
  const byExtension = new Map()
  try {
    const formats = require('./formats.json').formats || []
    for (const format of formats) {
      if (!format || typeof format.id !== 'string' || typeof format.validator !== 'string') continue
      byId.set(format.id.toLowerCase(), format.validator)
      for (const extension of format.extensions || []) {
        const key = String(extension).toLowerCase().replace(/^\./, '')
        if (!byExtension.has(key)) byExtension.set(key, { validator: format.validator, id: format.id })
      }
    }
  } catch {
    // Without the registry only validator ids resolve; unknown formats skip the structural check.
  }
  registryMaps = { byId, byExtension }
  return registryMaps
}

/**
 * Resolves a validator id, registry format id or extension to a validator.
 * @param {string} formatOrValidator for example 'ooxml-word', 'docx' or '.docx'
 * @returns {{validator: string, format: string|null}|null} null when unknown
 */
function resolveValidator(formatOrValidator) {
  if (typeof formatOrValidator !== 'string' || !formatOrValidator.trim()) return null
  const key = formatOrValidator.trim().toLowerCase()
  const { byId, byExtension } = registry()
  if (!key.startsWith('.') && byId.has(key)) return { validator: byId.get(key), format: key }
  if (!key.startsWith('.') && VALIDATOR_IDS.includes(key)) return { validator: key, format: null }
  const extension = byExtension.get(key.replace(/^\./, ''))
  if (extension) return { validator: extension.validator, format: extension.id.toLowerCase() }
  return null
}

// ---------------------------------------------------------------------------
// PDF
// ---------------------------------------------------------------------------

/**
 * %PDF- within the first 1 KB, %%EOF within the last 2 KB, and a startxref
 * that points inside the file at a cross-reference table or stream.
 * @param {Buffer} bytes
 */
function validatePdf(bytes) {
  const head = bytes.subarray(0, 1024).toString('latin1')
  const headerAt = head.indexOf('%PDF-')
  if (headerAt < 0) fail('has no %PDF- header')
  const tailStart = Math.max(0, bytes.length - 2048)
  const tail = bytes.subarray(tailStart).toString('latin1')
  const eofAt = tail.lastIndexOf('%%EOF')
  if (eofAt < 0) fail('has no %%EOF end marker, so the file is incomplete')
  const startxrefAt = tail.lastIndexOf('startxref', eofAt)
  if (startxrefAt < 0) fail('has no startxref before its %%EOF marker')
  const match = /^startxref\s+(\d+)\s*$/.exec(tail.slice(startxrefAt, eofAt))
  if (!match) fail('has an unreadable startxref value')
  const offset = Number(match[1])
  if (!(offset < bytes.length)) fail('points its startxref past the end of the file')
  const looksLikeXref = (position) => {
    if (position < 0 || position >= bytes.length) return false
    const text = bytes.subarray(position, Math.min(bytes.length, position + 64)).toString('latin1').replace(/^[\0\t\n\f\r ]+/, '')
    return text.startsWith('xref') || /^\d+\s+\d+\s+obj\b/.test(text)
  }
  if (!looksLikeXref(offset) && !looksLikeXref(offset + headerAt)) fail('has a startxref that does not point at a cross-reference table')
  return { version: head.slice(headerAt + 5, headerAt + 8) }
}

// ---------------------------------------------------------------------------
// ZIP, OOXML, ODF
// ---------------------------------------------------------------------------

function readZip64Extra(extra, entry) {
  let cursor = 0
  while (cursor + 4 <= extra.length) {
    const id = extra.readUInt16LE(cursor)
    const size = extra.readUInt16LE(cursor + 2)
    const body = extra.subarray(cursor + 4, cursor + 4 + size)
    if (id === 0x0001) {
      let position = 0
      const next = () => {
        if (position + 8 > body.length) fail(`entry "${entry.name}" has a damaged ZIP64 field`)
        const value = Number(body.readBigUInt64LE(position))
        position += 8
        return value
      }
      if (entry.uncompressedSize === 0xffffffff) entry.uncompressedSize = next()
      if (entry.compressedSize === 0xffffffff) entry.compressedSize = next()
      if (entry.localOffset === 0xffffffff) entry.localOffset = next()
      return
    }
    cursor += 4 + size
  }
  if (entry.uncompressedSize === 0xffffffff || entry.compressedSize === 0xffffffff || entry.localOffset === 0xffffffff) {
    fail(`entry "${entry.name}" needs a ZIP64 field it does not have`)
  }
}

/**
 * Reads and bounds-checks a ZIP central directory.
 * @param {Buffer} bytes
 * @param {{maxEntries?: number}} [options]
 * @returns {{entries: Map<string, object>, ordered: object[], cdOffset: number, cdSize: number}}
 */
function readZipDirectory(bytes, options = {}) {
  const maxEntries = options.maxEntries || MAX_ZIP_ENTRIES
  const size = bytes.length
  if (size < 22) fail('is too short to be a ZIP package')
  let eocd = -1
  const lowest = Math.max(0, size - 22 - 0xffff)
  for (let offset = size - 22; offset >= lowest; offset -= 1) {
    if (bytes[offset] === 0x50 && bytes[offset + 1] === 0x4b && bytes[offset + 2] === 0x05 && bytes[offset + 3] === 0x06
      && offset + 22 + bytes.readUInt16LE(offset + 20) === size) {
      eocd = offset
      break
    }
  }
  if (eocd < 0) fail('has no end-of-archive record, so the file is incomplete')
  let disk = bytes.readUInt16LE(eocd + 4)
  let directoryDisk = bytes.readUInt16LE(eocd + 6)
  let total = bytes.readUInt16LE(eocd + 10)
  let cdSize = bytes.readUInt32LE(eocd + 12)
  let cdOffset = bytes.readUInt32LE(eocd + 16)
  let directoryLimit = eocd
  if (total === 0xffff || cdSize === 0xffffffff || cdOffset === 0xffffffff || bytes.readUInt16LE(eocd + 8) === 0xffff) {
    const locator = eocd - 20
    if (locator < 0 || bytes.readUInt32LE(locator) !== 0x07064b50) fail('declares ZIP64 sizes but has no ZIP64 locator')
    const record = Number(bytes.readBigUInt64LE(locator + 8))
    if (record + 56 > locator || bytes.readUInt32LE(record) !== 0x06064b50) fail('has a damaged ZIP64 end record')
    disk = bytes.readUInt32LE(record + 16)
    directoryDisk = bytes.readUInt32LE(record + 20)
    total = Number(bytes.readBigUInt64LE(record + 32))
    cdSize = Number(bytes.readBigUInt64LE(record + 40))
    cdOffset = Number(bytes.readBigUInt64LE(record + 48))
    directoryLimit = record
  }
  if (disk !== 0 || directoryDisk !== 0) fail('is split across several disks')
  if (total > maxEntries) fail(`has ${total} entries, more than the ${maxEntries} Simple accepts`)
  if (cdOffset + cdSize > directoryLimit) fail('has a central directory that runs past the end of the archive')
  const entries = new Map()
  const ordered = []
  const directoryEnd = cdOffset + cdSize
  let cursor = cdOffset
  for (let index = 0; index < total; index += 1) {
    if (cursor + 46 > directoryEnd || bytes.readUInt32LE(cursor) !== 0x02014b50) fail('has a damaged central directory')
    const nameLength = bytes.readUInt16LE(cursor + 28)
    const extraLength = bytes.readUInt16LE(cursor + 30)
    const commentLength = bytes.readUInt16LE(cursor + 32)
    const next = cursor + 46 + nameLength + extraLength + commentLength
    if (next > directoryEnd) fail('has a central directory record that runs past its end')
    const nameBytes = bytes.subarray(cursor + 46, cursor + 46 + nameLength)
    const entry = {
      name: nameBytes.toString('utf8'),
      nameBytes,
      flags: bytes.readUInt16LE(cursor + 8),
      method: bytes.readUInt16LE(cursor + 10),
      crc32: bytes.readUInt32LE(cursor + 16),
      compressedSize: bytes.readUInt32LE(cursor + 20),
      uncompressedSize: bytes.readUInt32LE(cursor + 24),
      localOffset: bytes.readUInt32LE(cursor + 42),
      dataStart: null,
    }
    readZip64Extra(bytes.subarray(cursor + 46 + nameLength, cursor + 46 + nameLength + extraLength), entry)
    if (entries.has(entry.name)) fail(`contains "${entry.name}" twice`)
    entries.set(entry.name, entry)
    ordered.push(entry)
    cursor = next
  }
  return { entries, ordered, cdOffset, cdSize }
}

function locateEntry(bytes, zip, entry) {
  if (entry.dataStart !== null) return entry.dataStart
  const header = entry.localOffset
  if (header + 30 > zip.cdOffset || bytes.readUInt32LE(header) !== 0x04034b50) fail(`has no local header for "${entry.name}"`)
  const nameLength = bytes.readUInt16LE(header + 26)
  const extraLength = bytes.readUInt16LE(header + 28)
  if (!bytes.subarray(header + 30, header + 30 + nameLength).equals(entry.nameBytes)) fail(`has a local header for "${entry.name}" that does not match its directory record`)
  const dataStart = header + 30 + nameLength + extraLength
  if (dataStart + entry.compressedSize > zip.cdOffset) fail(`has data for "${entry.name}" that runs past the end of the archive`)
  entry.dataStart = dataStart
  return dataStart
}

function inflateEntry(bytes, zip, entry) {
  const start = locateEntry(bytes, zip, entry)
  const raw = bytes.subarray(start, start + entry.compressedSize)
  if (entry.flags & 0x1) fail(`has an encrypted entry "${entry.name}"`)
  if (entry.method === 0) {
    if (raw.length !== entry.uncompressedSize) fail(`has a stored entry "${entry.name}" with the wrong size`)
    return raw
  }
  if (entry.method !== 8) fail(`uses an unsupported compression method for "${entry.name}"`)
  let data
  try {
    data = zlib.inflateRawSync(raw, { maxOutputLength: Math.max(1, entry.uncompressedSize) })
  } catch {
    fail(`has an entry "${entry.name}" that can't be decompressed`)
  }
  if (data.length !== entry.uncompressedSize) fail(`has an entry "${entry.name}" with the wrong size`)
  return data
}

/**
 * Checks every local header and, within a budget, every entry's CRC.
 * @param {Buffer} bytes
 * @param {object} zip result of readZipDirectory
 * @param {{crc?: 'all'|'required'|'none', required?: Iterable<string>, crcBudget?: number}} [options]
 */
function checkZipEntries(bytes, zip, options = {}) {
  const mode = options.crc || 'all'
  const required = new Set(options.required || [])
  const budget = options.crcBudget ?? DEFAULT_CRC_BUDGET
  let inflated = 0
  for (const entry of zip.ordered) {
    locateEntry(bytes, zip, entry)
    const isRequired = required.has(entry.name)
    if (mode === 'none' || (mode === 'required' && !isRequired)) continue
    if ((entry.flags & 0x1) || (entry.method !== 0 && entry.method !== 8)) {
      if (isRequired) inflateEntry(bytes, zip, entry)
      continue
    }
    if (!isRequired && inflated + entry.uncompressedSize > budget) continue
    const data = inflateEntry(bytes, zip, entry)
    inflated += data.length
    if (crc32(data) !== entry.crc32) fail(`has an entry "${entry.name}" that fails its CRC check`)
  }
}

/**
 * Lists the entries of a ZIP package (bounds-checked).
 * @param {Buffer|Uint8Array} bytes
 * @returns {Array<{name: string, method: number, compressedSize: number, uncompressedSize: number, crc32: number}>}
 */
function readZipEntries(bytes) {
  const buffer = toBuffer(bytes)
  return readZipDirectory(buffer).ordered.map(({ name, method, compressedSize, uncompressedSize, crc32: crc }) => ({ name, method, compressedSize, uncompressedSize, crc32: crc }))
}

/**
 * Reads one entry of a ZIP package, or null when it is not there.
 * @param {Buffer|Uint8Array} bytes
 * @param {string} name
 * @returns {Buffer|null}
 */
function readZipEntry(bytes, name) {
  const buffer = toBuffer(bytes)
  const zip = readZipDirectory(buffer)
  const entry = zip.entries.get(name)
  return entry ? inflateEntry(buffer, zip, entry) : null
}

function validateZip(bytes, options = {}) {
  const zip = readZipDirectory(bytes, options)
  checkZipEntries(bytes, zip, options)
  return { entries: zip.ordered.length }
}

function attributeOf(tag, name) {
  const match = new RegExp(`\\b${name}\\s*=\\s*(["'])(.*?)\\1`, 's').exec(tag)
  return match ? match[2] : null
}

function normalizePart(target, base = '') {
  let decoded = target
  try { decoded = decodeURIComponent(target) } catch {}
  const parts = []
  for (const part of `${decoded.startsWith('/') ? '' : base}${decoded}`.replace(/\\/g, '/').split('/')) {
    if (!part || part === '.') continue
    if (part === '..') parts.pop()
    else parts.push(part)
  }
  return parts.join('/')
}

function validateOoxml(bytes, validator, options = {}) {
  const kind = OOXML_KINDS[validator]
  const zip = readZipDirectory(bytes, options)
  if (!zip.entries.has('[Content_Types].xml')) fail('is missing [Content_Types].xml')
  let mainPart = null
  const relsEntry = zip.entries.get('_rels/.rels')
  if (relsEntry) {
    const rels = inflateEntry(bytes, zip, relsEntry).toString('utf8')
    for (const tag of rels.match(/<(?:\w+:)?Relationship\b[^>]*>/g) || []) {
      const type = attributeOf(tag, 'Type') || ''
      const target = attributeOf(tag, 'Target')
      if (target && /\/officeDocument$/.test(type) && !/^External$/i.test(attributeOf(tag, 'TargetMode') || '')) {
        mainPart = normalizePart(target)
        break
      }
    }
  }
  if (!mainPart) mainPart = kind.parts.find((part) => zip.entries.has(part)) || null
  if (!mainPart || !zip.entries.has(mainPart)) fail(`is missing the main part of ${kind.label}`)
  const contentTypes = inflateEntry(bytes, zip, zip.entries.get('[Content_Types].xml')).toString('utf8')
  if (!/<(?:\w+:)?Types\b/.test(contentTypes)) fail('has an unreadable [Content_Types].xml')
  let declared = null
  for (const tag of contentTypes.match(/<(?:\w+:)?Override\b[^>]*>/g) || []) {
    const partName = attributeOf(tag, 'PartName')
    if (partName && normalizePart(partName).toLowerCase() === mainPart.toLowerCase()) {
      declared = attributeOf(tag, 'ContentType')
      break
    }
  }
  const matchesKind = declared ? kind.contentType.test(declared) : mainPart.startsWith(kind.folder)
  if (!matchesKind) fail(`is not ${kind.label}`)
  checkZipEntries(bytes, zip, { ...options, required: ['[Content_Types].xml', '_rels/.rels', mainPart].filter((name) => zip.entries.has(name)) })
  return { mainPart, entries: zip.ordered.length }
}

// The ODF package rule says "mimetype" must be the first entry and stored.
// SheetJS (Calc's ODS writer) puts it third and may compress it; LibreOffice
// and SheetJS still open such files, so by default that is reported in
// `notes` instead of failing. `strictOdf: true` enforces the rule.
function validateOdf(bytes, options = {}) {
  const zip = readZipDirectory(bytes, options)
  const mimeEntry = zip.entries.get('mimetype')
  if (!mimeEntry) fail('is missing its mimetype entry')
  const mime = inflateEntry(bytes, zip, mimeEntry).toString('latin1').trim()
  const expected = ODF_MEDIA_TYPES[options.format]
  if (expected ? (mime !== expected && mime !== `${expected}-template`) : !mime.startsWith('application/vnd.oasis.opendocument.')) {
    fail(`declares the media type "${mime.slice(0, 80)}"${expected ? ` instead of ${expected}` : ''}`)
  }
  if (!zip.entries.has('content.xml')) fail('is missing content.xml')
  checkZipEntries(bytes, zip, { ...options, required: ['mimetype', 'content.xml'] })
  const notes = []
  if (mimeEntry.localOffset !== 0) notes.push('mimetype-not-first')
  if (mimeEntry.method !== 0) notes.push('mimetype-compressed')
  if (options.strictOdf && notes.length) fail('does not start with a stored mimetype entry')
  return { mediaType: mime, entries: zip.ordered.length, conformant: !notes.length, notes }
}

// ---------------------------------------------------------------------------
// Compound File Binary (.doc, .xls, .ppt)
// ---------------------------------------------------------------------------

function validateCfb(bytes) {
  if (bytes.length < 512) fail('is too short to be an Office 97–2003 file')
  if (!bytes.subarray(0, 8).equals(CFB_SIGNATURE)) fail('has no compound-file signature')
  if (bytes.readUInt16LE(28) !== 0xfffe) fail('has an invalid byte-order mark')
  const sectorShift = bytes.readUInt16LE(30)
  if (sectorShift !== 9 && sectorShift !== 12) fail(`has an unsupported sector size (shift ${sectorShift})`)
  const sectorSize = 1 << sectorShift
  if (bytes.length % sectorSize !== 0) fail('is not a whole number of sectors long, so the file is incomplete')
  const sectorCount = bytes.length / sectorSize - 1
  const fatCount = bytes.readUInt32LE(44)
  const firstDirectory = bytes.readUInt32LE(48)
  const miniCutoff = bytes.readUInt32LE(56) || 4096
  const firstMiniFat = bytes.readUInt32LE(60)
  const miniFatCount = bytes.readUInt32LE(64)
  let difatSector = bytes.readUInt32LE(68)
  const difatCount = bytes.readUInt32LE(72)
  if (fatCount === 0 || fatCount > sectorCount) fail('has an impossible allocation table size')
  const sectorOffset = (id) => (id + 1) * sectorSize
  const fatSectors = []
  for (let index = 0; index < 109 && fatSectors.length < fatCount; index += 1) fatSectors.push(bytes.readUInt32LE(76 + index * 4))
  const perSector = sectorSize / 4
  for (let walked = 0; fatSectors.length < fatCount; walked += 1) {
    if (walked >= Math.max(difatCount, 1) + 1 || difatSector >= sectorCount) fail('has a damaged allocation table index')
    const base = sectorOffset(difatSector)
    for (let index = 0; index < perSector - 1 && fatSectors.length < fatCount; index += 1) fatSectors.push(bytes.readUInt32LE(base + index * 4))
    difatSector = bytes.readUInt32LE(base + (perSector - 1) * 4)
  }
  const fat = new Uint32Array(fatSectors.length * perSector)
  fatSectors.forEach((id, index) => {
    if (id >= sectorCount) fail('refers to allocation table sectors past the end of the file')
    const base = sectorOffset(id)
    for (let entry = 0; entry < perSector; entry += 1) fat[index * perSector + entry] = bytes.readUInt32LE(base + entry * 4)
  })
  const chain = (start, what) => {
    const sectors = []
    const seen = new Set()
    for (let id = start; id !== CFB_END_OF_CHAIN; id = fat[id]) {
      if (id >= sectorCount || id >= fat.length || seen.has(id)) fail(`has a damaged ${what} chain`)
      seen.add(id)
      sectors.push(id)
    }
    return sectors
  }
  const directory = chain(firstDirectory, 'directory')
  if (!directory.length) fail('has no directory')
  if (miniFatCount > 0) chain(firstMiniFat, 'mini allocation table')
  const rootOffset = sectorOffset(directory[0])
  if (bytes[rootOffset + 66] !== 5) fail('has no root storage')
  const miniStreamSize = Number(bytes.readBigUInt64LE(rootOffset + 120)) % 2 ** 32
  const miniStreamStart = bytes.readUInt32LE(rootOffset + 116)
  if (miniStreamSize > 0 && miniStreamStart !== CFB_END_OF_CHAIN && chain(miniStreamStart, 'mini stream').length * sectorSize < miniStreamSize) {
    fail('has a mini stream shorter than its declared size')
  }
  for (const id of directory) {
    for (let base = sectorOffset(id); base < sectorOffset(id) + sectorSize; base += 128) {
      const type = bytes[base + 66]
      if (type !== 2) continue
      const streamSize = Number(bytes.readBigUInt64LE(base + 120)) % 2 ** 32
      if (streamSize < miniCutoff || streamSize === 0) continue
      const start = bytes.readUInt32LE(base + 116)
      if (start === CFB_END_OF_CHAIN || start === CFB_FREE) fail('has a stream with no data')
      if (chain(start, 'stream').length * sectorSize < streamSize) fail('has a stream shorter than its declared size')
    }
  }
  return { sectorSize, sectors: sectorCount }
}

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

function validatePng(bytes) {
  if (bytes.length < 8 || !bytes.subarray(0, 8).equals(PNG_SIGNATURE)) fail('has no PNG signature')
  let cursor = 8
  let first = true
  let sawData = false
  for (;;) {
    if (cursor + 12 > bytes.length) fail('ends before its IEND chunk, so the file is incomplete')
    const length = bytes.readUInt32BE(cursor)
    const type = bytes.subarray(cursor + 4, cursor + 8).toString('latin1')
    const end = cursor + 12 + length
    if (length > 0x7fffffff || end > bytes.length) fail(`has a ${type} chunk that runs past the end of the file`)
    if (!/^[A-Za-z]{4}$/.test(type)) fail('has a damaged chunk header')
    if (first && (type !== 'IHDR' || length !== 13)) fail('does not start with an IHDR chunk')
    first = false
    if (crc32(bytes.subarray(cursor + 4, cursor + 8 + length)) !== bytes.readUInt32BE(cursor + 8 + length)) fail(`has a ${type} chunk that fails its CRC check`)
    if (type === 'IDAT') sawData = true
    cursor = end
    if (type === 'IEND') {
      if (length !== 0) fail('has a damaged IEND chunk')
      if (!sawData) fail('has no image data')
      if (cursor !== bytes.length) fail('has data after its IEND chunk')
      return { width: bytes.readUInt32BE(16), height: bytes.readUInt32BE(20) }
    }
  }
}

function validateJpeg(bytes) {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) fail('has no JPEG start marker')
  let cursor = 2
  let sawFrame = false
  let sawScan = false
  let progressive = false
  while (cursor < bytes.length) {
    if (bytes[cursor] !== 0xff) fail('has a damaged segment')
    while (cursor < bytes.length && bytes[cursor] === 0xff) cursor += 1
    if (cursor >= bytes.length) break
    const marker = bytes[cursor]
    cursor += 1
    if (marker === 0xd9) {
      if (!sawFrame || !sawScan) fail('ends before its image data')
      for (let index = cursor; index < bytes.length; index += 1) if (bytes[index] !== 0) fail('has data after its end marker')
      return { progressive }
    }
    if (marker === 0xd8 || marker === 0x00) fail('has a misplaced marker')
    if ((marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) continue
    if (cursor + 2 > bytes.length) break
    const length = bytes.readUInt16BE(cursor)
    if (length < 2 || cursor + length > bytes.length) fail('has a segment that runs past the end of the file')
    if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
      sawFrame = true
      if (marker === 0xc2 || marker === 0xc6 || marker === 0xca || marker === 0xce) progressive = true
    }
    cursor += length
    if (marker === 0xda) {
      if (!sawFrame) fail('has image data before its frame header')
      sawScan = true
      // Entropy-coded data runs until a marker that is not a stuffed byte or a restart marker.
      while (cursor + 1 < bytes.length && !(bytes[cursor] === 0xff && bytes[cursor + 1] !== 0x00 && !(bytes[cursor + 1] >= 0xd0 && bytes[cursor + 1] <= 0xd7))) cursor += 1
      if (cursor + 1 >= bytes.length) break
    }
  }
  fail('has no end marker, so the file is incomplete')
}

function skipGifSubBlocks(bytes, cursor) {
  for (;;) {
    if (cursor >= bytes.length) fail('ends inside a data block, so the file is incomplete')
    const size = bytes[cursor]
    cursor += 1
    if (size === 0) return cursor
    cursor += size
  }
}

function validateGif(bytes) {
  const signature = bytes.subarray(0, 6).toString('latin1')
  if (signature !== 'GIF87a' && signature !== 'GIF89a') fail('has no GIF signature')
  if (bytes.length < 14) fail('is too short to be a GIF image')
  const packed = bytes[10]
  let cursor = 13 + (packed & 0x80 ? 3 * 2 ** ((packed & 0x07) + 1) : 0)
  let frames = 0
  while (cursor < bytes.length) {
    const block = bytes[cursor]
    cursor += 1
    if (block === 0x3b) {
      if (!frames) fail('has no image')
      for (let index = cursor; index < bytes.length; index += 1) if (bytes[index] !== 0) fail('has data after its trailer')
      return { frames }
    }
    if (block === 0x21) {
      cursor = skipGifSubBlocks(bytes, cursor + 1)
    } else if (block === 0x2c) {
      if (cursor + 9 > bytes.length) fail('ends inside an image header, so the file is incomplete')
      const localPacked = bytes[cursor + 8]
      cursor += 9 + (localPacked & 0x80 ? 3 * 2 ** ((localPacked & 0x07) + 1) : 0)
      cursor = skipGifSubBlocks(bytes, cursor + 1)
      frames += 1
    } else {
      fail('has a damaged block')
    }
  }
  fail('has no trailer, so the file is incomplete')
}

function validateWebp(bytes) {
  if (bytes.length < 20 || bytes.subarray(0, 4).toString('latin1') !== 'RIFF' || bytes.subarray(8, 12).toString('latin1') !== 'WEBP') fail('has no WebP signature')
  if (bytes.readUInt32LE(4) + 8 !== bytes.length) fail('is not the length its header declares, so the file is incomplete')
  let cursor = 12
  let first = null
  while (cursor < bytes.length) {
    if (cursor + 8 > bytes.length) fail('has a damaged chunk header')
    const type = bytes.subarray(cursor, cursor + 4).toString('latin1')
    const size = bytes.readUInt32LE(cursor + 4)
    const end = cursor + 8 + size + (size % 2)
    if (end > bytes.length) fail(`has a ${type.trim()} chunk that runs past the end of the file`)
    if (first === null) first = type
    cursor = end
  }
  if (!['VP8 ', 'VP8L', 'VP8X'].includes(first)) fail('does not start with an image chunk')
  return { kind: first.trim() }
}

function validateBmp(bytes) {
  if (bytes.length < 26 || bytes[0] !== 0x42 || bytes[1] !== 0x4d) fail('has no BMP signature')
  if (bytes.readUInt32LE(2) !== bytes.length) fail('is not the length its header declares, so the file is incomplete')
  const dataOffset = bytes.readUInt32LE(10)
  const headerSize = bytes.readUInt32LE(14)
  if (![12, 16, 40, 52, 56, 64, 108, 124].includes(headerSize) || 14 + headerSize > bytes.length) fail('has an unknown header')
  if (dataOffset < 14 + headerSize || dataOffset >= bytes.length) fail('has a pixel offset outside the file')
  if (headerSize === 12) return { width: bytes.readUInt16LE(18), height: bytes.readUInt16LE(20) }
  const width = bytes.readInt32LE(18)
  const height = bytes.readInt32LE(22)
  const bitCount = bytes.readUInt16LE(28)
  const compression = headerSize >= 40 ? bytes.readUInt32LE(30) : 0
  const imageSize = headerSize >= 40 ? bytes.readUInt32LE(34) : 0
  if (width <= 0 || height === 0) fail('has an impossible size')
  const needed = compression === 0 || compression === 3 || compression === 6
    ? Math.floor((bitCount * width + 31) / 32) * 4 * Math.abs(height)
    : imageSize
  if (dataOffset + needed > bytes.length) fail('ends before its pixel data, so the file is incomplete')
  return { width, height: Math.abs(height) }
}

function validateTiff(bytes) {
  if (bytes.length < 8) fail('is too short to be a TIFF image')
  const order = bytes.subarray(0, 2).toString('latin1')
  if (order !== 'II' && order !== 'MM') fail('has no TIFF byte-order mark')
  const little = order === 'II'
  const u16 = (offset) => (little ? bytes.readUInt16LE(offset) : bytes.readUInt16BE(offset))
  const u32 = (offset) => (little ? bytes.readUInt32LE(offset) : bytes.readUInt32BE(offset))
  const magic = u16(2)
  if (magic === 43) return { bigTiff: true }
  if (magic !== 42) fail('has no TIFF signature')
  const typeSizes = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8, 13: 4 }
  const values = (entry) => {
    const type = u16(entry + 2)
    const count = u32(entry + 4)
    const size = (typeSizes[type] || 0) * count
    if (!size) return []
    const at = size <= 4 ? entry + 8 : u32(entry + 8)
    if (at + size > bytes.length) fail('has a tag that points past the end of the file')
    const list = []
    for (let index = 0; index < count && index < 100_000; index += 1) list.push(type === 3 ? u16(at + index * 2) : type === 4 || type === 13 ? u32(at + index * 4) : 0)
    return list
  }
  let ifd = u32(4)
  let pages = 0
  const seen = new Set()
  while (ifd !== 0) {
    if (seen.has(ifd) || ifd + 2 > bytes.length) fail('has a damaged directory')
    seen.add(ifd)
    const count = u16(ifd)
    if (ifd + 2 + count * 12 + 4 > bytes.length) fail('has a directory that runs past the end of the file')
    const tags = new Map()
    for (let index = 0; index < count; index += 1) {
      const entry = ifd + 2 + index * 12
      tags.set(u16(entry), entry)
    }
    for (const [offsetsTag, countsTag] of [[273, 279], [324, 325]]) {
      if (!tags.has(offsetsTag)) continue
      const offsets = values(tags.get(offsetsTag))
      const counts = tags.has(countsTag) ? values(tags.get(countsTag)) : []
      offsets.forEach((offset, index) => {
        if (offset + (counts[index] || 0) > bytes.length) fail('has image data past the end of the file, so the file is incomplete')
      })
    }
    pages += 1
    ifd = u32(ifd + 2 + count * 12)
    if (pages > 10_000) fail('has too many pages')
  }
  if (!pages) fail('has no image directory')
  return { pages }
}

function validateIco(bytes) {
  if (bytes.length < 6 || bytes.readUInt16LE(0) !== 0) fail('has no icon header')
  const type = bytes.readUInt16LE(2)
  const count = bytes.readUInt16LE(4)
  if ((type !== 1 && type !== 2) || count === 0) fail('has no icon header')
  const tableEnd = 6 + count * 16
  if (tableEnd > bytes.length) fail('ends inside its image table, so the file is incomplete')
  for (let index = 0; index < count; index += 1) {
    const entry = 6 + index * 16
    const size = bytes.readUInt32LE(entry + 8)
    const offset = bytes.readUInt32LE(entry + 12)
    if (offset < tableEnd || offset + size > bytes.length || size < 8) fail('has an image that runs past the end of the file')
    const image = bytes.subarray(offset, offset + size)
    if (image.subarray(0, 8).equals(PNG_SIGNATURE)) validatePng(image)
    else if (image.readUInt32LE(0) !== 40) fail('has an image in an unknown format')
  }
  return { images: count }
}

// ---------------------------------------------------------------------------
// Text and JSON
// ---------------------------------------------------------------------------

function normalizeEncoding(value) {
  if (!value) return null
  const key = String(value).trim().toLowerCase()
  return TEXT_ENCODING_ALIASES[key] || key
}

function detectBom(bytes) {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { encoding: 'utf-8', length: 3 }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return { encoding: 'utf-16le', length: 2 }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return { encoding: 'utf-16be', length: 2 }
  return null
}

function countReplacement(text) {
  let count = 0
  for (let index = text.indexOf('\uFFFD'); index >= 0; index = text.indexOf('\uFFFD', index + 1)) count += 1
  return count
}

function stripTrailingNoise(text) {
  return text.replace(/(?:\s|<!--[\s\S]*?-->)+$/, '')
}

function checkXmlRoot(head, tail = head) {
  const body = head.replace(/^\s*(?:<\?[\s\S]*?\?>\s*|<!--[\s\S]*?-->\s*|<!DOCTYPE[^>[]*(?:\[[\s\S]*?\])?\s*>\s*)*/i, '')
  const root = /^<([A-Za-z_][\w.:-]*)/.exec(body)
  if (!root) fail('has no XML root element')
  const end = stripTrailingNoise(head === tail ? body : tail)
  const selfClosed = head === tail && /\/>$/.test(end) && !end.slice(0, -2).includes('<', 1)
  if (!end.endsWith(`</${root[1]}>`) && !selfClosed) fail(`ends before its closing </${root[1]}> tag, so the file is incomplete`)
}

/**
 * Strictly decodes text. Up to `limit` bytes the whole text comes back;
 * larger input is checked in streamed chunks and only its head and tail are
 * kept, so a huge export never hits string-length limits.
 * @returns {{valid: true, whole: boolean, text: string|null, head: string, tail: string}|{valid: false}|null}
 *   null when TextDecoder does not know the encoding
 */
function decodeChecked(body, encoding, limit) {
  let decoder
  try {
    decoder = new TextDecoder(encoding, { fatal: true, ignoreBOM: true })
  } catch {
    return null
  }
  let head = ''
  try {
    if (body.length <= limit) {
      const text = decoder.decode(body)
      return { valid: true, whole: true, text, head: text, tail: text }
    }
    for (let offset = 0; offset < body.length; offset += TEXT_CHUNK_BYTES) {
      const chunk = decoder.decode(body.subarray(offset, offset + TEXT_CHUNK_BYTES), { stream: true })
      if (!offset) head = chunk.slice(0, TEXT_TAIL_BYTES)
    }
    decoder.decode()
  } catch {
    return { valid: false }
  }
  let tailStart = Math.max(0, body.length - TEXT_TAIL_BYTES)
  if (encoding.startsWith('utf-16')) tailStart -= tailStart % 2
  const tail = new TextDecoder(encoding, { ignoreBOM: true }).decode(body.subarray(tailStart))
  return { valid: true, whole: false, text: null, head, tail }
}

function checkRtfGroups(text) {
  if (!text.startsWith('{\\rtf')) fail('has no {\\rtf header')
  let depth = 0
  let index = 0
  for (; index < text.length; index += 1) {
    const character = text[index]
    if (character === '\\') {
      const binary = /^\\bin(\d+) ?/.exec(text.slice(index, index + 20))
      if (binary) index += binary[0].length - 1 + Number(binary[1])
      else index += 1
    } else if (character === '{') depth += 1
    else if (character === '}') {
      depth -= 1
      if (depth === 0) break
      if (depth < 0) fail('has unbalanced braces')
    }
  }
  if (depth !== 0) fail('ends before its last closing brace, so the file is incomplete')
  if (text.slice(index + 1).replace(/[\s\0]+/g, '')) fail('has text after its last closing brace')
}

/**
 * Text check: the bytes decode with the declared (or byte-order-mark) encoding
 * without invalid sequences. When the caller passes the source `text`, the
 * decoded text must equal it (UTF-8 and UTF-16) or add no U+FFFD (legacy code
 * pages). Markup formats must reach their closing tag; JSON must parse.
 * Undeclared bytes that are not UTF-8 are a legacy code page and pass.
 * @param {Buffer} bytes
 * @param {{format?: string, encoding?: string, text?: string, maxWholeTextBytes?: number}} [options]
 */
function validateText(bytes, options = {}) {
  const declared = normalizeEncoding(options.encoding)
  const bom = detectBom(bytes)
  if (declared && bom && bom.encoding !== declared) fail(`starts with a ${bom.encoding} byte-order mark but should be ${declared}`)
  const encoding = declared || (bom && bom.encoding) || null
  const body = bom ? bytes.subarray(bom.length) : bytes
  if ((encoding === 'utf-16le' || encoding === 'utf-16be') && body.length % 2) {
    fail(`has an odd number of bytes for ${encoding} text, so the file is incomplete`)
  }
  let decoded = decodeChecked(body, encoding || 'utf-8', options.maxWholeTextBytes ?? MAX_WHOLE_TEXT_BYTES)
  if (decoded && !decoded.valid) {
    if (encoding) fail(`isn't valid ${encoding === 'utf-8' ? 'UTF-8' : encoding} text`)
    decoded = null
  }
  const notes = []
  if (typeof options.text === 'string' && decoded) {
    const expected = options.text.replace(/^\uFEFF/, '')
    const exactCodec = encoding === null || ['utf-8', 'utf-16le', 'utf-16be'].includes(encoding)
    if (!decoded.whole) notes.push('the source text was not compared because the file is very large')
    else if (exactCodec) {
      if (decoded.text !== expected) fail(decoded.text.length < expected.length ? 'is shorter than the document text, so the file is incomplete' : 'does not match the document text')
    } else if (countReplacement(decoded.text) > countReplacement(expected)) {
      fail(`has characters that ${encoding} can't hold`)
    }
  }
  const format = String(options.format || '').toLowerCase()
  if (decoded && format) {
    const head = decoded.head.replace(/^\s+/, '')
    const tail = decoded.whole ? head : decoded.tail
    if (format === 'json') {
      if (decoded.whole) {
        try { JSON.parse(decoded.text) } catch { fail("isn't valid JSON") }
      } else if (!/^[[{]/.test(head) || !/[\]}]$/.test(tail.trimEnd())) {
        fail("isn't complete JSON")
      }
    } else if (format === 'html' || format === 'htm') {
      if (/^(?:<!doctype html[^>]*>\s*)?(?:<!--[\s\S]*?-->\s*)*<html[\s>]/i.test(head) && !/<\/html\s*>$/i.test(stripTrailingNoise(tail))) {
        fail('ends before its closing </html> tag, so the file is incomplete')
      }
    } else if (format === 'xml' || format === 'svg' || format === 'fods') {
      checkXmlRoot(head, tail)
    } else if (format === 'rtf') {
      if (decoded.whole) checkRtfGroups(head)
      else if (!head.startsWith('{\\rtf') || !tail.replace(/[\s\0]+$/, '').endsWith('}')) fail('ends before its last closing brace, so the file is incomplete')
    }
  }
  const result = { encoding: encoding || (decoded ? 'utf-8' : 'unknown') }
  if (notes.length) result.notes = notes
  return result
}

function validateJson(bytes, options = {}) {
  const result = validateText(bytes, { ...options, format: 'json' })
  if (result.encoding === 'unknown') fail("isn't valid UTF-8 JSON")
  return result
}

// ---------------------------------------------------------------------------
// Entry points
// ---------------------------------------------------------------------------

const VALIDATORS = Object.freeze({
  none: () => ({}),
  pdf: (bytes) => validatePdf(bytes),
  zip: (bytes, options) => validateZip(bytes, options),
  'ooxml-word': (bytes, options) => validateOoxml(bytes, 'ooxml-word', options),
  'ooxml-excel': (bytes, options) => validateOoxml(bytes, 'ooxml-excel', options),
  'ooxml-powerpoint': (bytes, options) => validateOoxml(bytes, 'ooxml-powerpoint', options),
  odf: (bytes, options) => validateOdf(bytes, options),
  cfb: (bytes) => validateCfb(bytes),
  png: (bytes) => validatePng(bytes),
  jpeg: (bytes) => validateJpeg(bytes),
  gif: (bytes) => validateGif(bytes),
  webp: (bytes) => validateWebp(bytes),
  bmp: (bytes) => validateBmp(bytes),
  tiff: (bytes) => validateTiff(bytes),
  ico: (bytes) => validateIco(bytes),
  text: (bytes, options) => validateText(bytes, options),
  json: (bytes, options) => validateJson(bytes, options),
})

/**
 * Checks that bytes are a structurally complete file of the given format.
 *
 * @param {string} formatOrValidator validator id ('ooxml-excel'), registry format id ('xlsx') or extension ('.xlsx')
 * @param {Buffer|Uint8Array|ArrayBuffer|string} bytes
 * @param {object} [options]
 * @param {string} [options.format] registry format id when `formatOrValidator` is a validator id (ODF media type, HTML/XML/RTF checks)
 * @param {string} [options.encoding] declared text encoding ('utf-8', 'utf-16le', 'windows-1252', …)
 * @param {string} [options.text] the source text the bytes were encoded from
 * @param {number} [options.maxWholeTextBytes] text above this size (default 64 MB) is checked in streamed chunks
 * @param {'all'|'required'|'none'} [options.crc='all'] which ZIP entries get a CRC check
 * @param {number} [options.crcBudget] stop CRC-checking optional entries after this many inflated bytes
 * @param {boolean} [options.strictOdf=false] fail ODF packages whose mimetype entry is not first and stored
 * @returns {{ok: true, validator: string|null, format: string|null, details?: object, skipped?: boolean}
 *   | {ok: false, validator: string, format: string|null, reason: string}}
 *   `reason` completes the sentence "The file …"; an unknown format is skipped, never failed.
 */
function validateBytes(formatOrValidator, bytes, options = {}) {
  const resolved = resolveValidator(formatOrValidator)
  if (!resolved) return { ok: true, validator: null, format: null, skipped: true }
  const format = resolved.format || (options.format ? String(options.format).toLowerCase() : null)
  const run = VALIDATORS[resolved.validator]
  if (!run) return { ok: true, validator: resolved.validator, format, skipped: true }
  try {
    const details = run(toBuffer(bytes), { ...options, format: format || options.format })
    return { ok: true, validator: resolved.validator, format, details }
  } catch (error) {
    if (error instanceof StructureError) return { ok: false, validator: resolved.validator, format, reason: error.message }
    if (error instanceof RangeError) return { ok: false, validator: resolved.validator, format, reason: 'ends unexpectedly, so the file is incomplete' }
    throw error
  }
}

/**
 * Reads a file and validates it. Convenience for tests and for checking a
 * file another tool produced.
 * @param {string} formatOrValidator
 * @param {string} filePath
 * @param {object} [options] see validateBytes
 * @returns {Promise<ReturnType<typeof validateBytes>>}
 */
async function validateFile(formatOrValidator, filePath, options = {}) {
  const bytes = await require('node:fs/promises').readFile(filePath)
  return validateBytes(formatOrValidator, bytes, options)
}

module.exports = {
  ODF_MEDIA_TYPES,
  VALIDATOR_IDS,
  checkZipEntries,
  crc32,
  readZipEntries,
  readZipEntry,
  resolveValidator,
  validateBytes,
  validateFile,
}
