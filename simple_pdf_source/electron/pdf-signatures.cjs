'use strict'

const fs = require('node:fs/promises')
const zlib = require('node:zlib')

// A signature dictionary records the signed byte ranges of the file in
// /ByteRange [offset length offset length]. Signers write it uncompressed
// (they patch the numbers after writing the file), and names may touch their
// neighbours (`/Type/Sig/ByteRange[0 840 960 240]`), so the scan accepts any
// whitespace instead of matching one spelling such as '/Type /Sig'.
const BYTE_RANGE = Buffer.from('/ByteRange')
const SIG_FLAGS = Buffer.from('/SigFlags')
const CERTIFICATION_NAMES = [Buffer.from('/DocMDP'), Buffer.from('/UR3')]
const OBJECT_STREAM = Buffer.from('/ObjStm')
const STREAM_KEYWORD = Buffer.from('stream')
const END_STREAM = Buffer.from('endstream')
const PDF_WHITESPACE = new Set([0x00, 0x09, 0x0a, 0x0c, 0x0d, 0x20])
const PDF_DELIMITERS = new Set([0x28, 0x29, 0x3c, 0x3e, 0x5b, 0x5d, 0x7b, 0x7d, 0x2f, 0x25])
const MAX_RANGE_TEXT = 512
const MAX_SIGNATURES = 1_000
const MAX_OBJECT_STREAMS = 20_000
const MAX_INFLATED_BYTES = 64 * 1024 * 1024

function asBuffer(value) {
  if (Buffer.isBuffer(value)) return value
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  if (value instanceof ArrayBuffer) return Buffer.from(value)
  return Buffer.from(value || [])
}

function isDigit(byte) {
  return byte >= 0x30 && byte <= 0x39
}

function endsName(bytes, index) {
  return index >= bytes.length || PDF_WHITESPACE.has(bytes[index]) || PDF_DELIMITERS.has(bytes[index])
}

/** Parse `[a b c d …]` following a /ByteRange key. Placeholders (`/****`) fail. */
function parseByteRange(bytes, start) {
  const end = Math.min(bytes.length, start + MAX_RANGE_TEXT)
  let index = start
  // The name must end here; `/ByteRangeX` is a different key.
  if (index < end && !PDF_WHITESPACE.has(bytes[index]) && bytes[index] !== 0x5b) return null
  while (index < end && PDF_WHITESPACE.has(bytes[index])) index += 1
  if (bytes[index] !== 0x5b) return null
  index += 1
  const values = []
  while (index < end) {
    while (index < end && PDF_WHITESPACE.has(bytes[index])) index += 1
    if (bytes[index] === 0x5d) break
    if (!isDigit(bytes[index])) return null
    let value = 0
    while (index < end && isDigit(bytes[index])) {
      value = value * 10 + bytes[index] - 0x30
      index += 1
      if (value > Number.MAX_SAFE_INTEGER / 10) return null
    }
    values.push(value)
    if (values.length > 64) return null
  }
  if (bytes[index] !== 0x5d || values.length < 4 || values.length % 2) return null
  return values
}

/**
 * True when the ranges still describe this file: the excluded gap holds the
 * hex /Contents string. A full rewrite (any pdf-lib save) moves the objects, so
 * this tells a valid signed revision from leftover signature objects.
 */
function rangeIsIntact(range, byteAt, length) {
  if (range[0] !== 0) return false
  for (let index = 0; index < range.length; index += 2) {
    const offset = range[index]
    const size = range[index + 1]
    if (offset + size > length) return false
    if (index + 2 < range.length && offset + size >= range[index + 2]) return false
  }
  const gapStart = range[0] + range[1]
  const gapEnd = range[2]
  return byteAt(gapStart) === 0x3c && byteAt(gapEnd - 1) === 0x3e
}

function scanByteRanges(bytes) {
  const signatures = []
  let index = bytes.indexOf(BYTE_RANGE)
  while (index >= 0 && signatures.length < MAX_SIGNATURES) {
    const range = parseByteRange(bytes, index + BYTE_RANGE.length)
    if (range) signatures.push({ range, intact: rangeIsIntact(range, (position) => bytes[position], bytes.length) })
    index = bytes.indexOf(BYTE_RANGE, index + BYTE_RANGE.length)
  }
  return signatures
}

const FIELDS = Buffer.from('/Fields')
const ACROFORM_WINDOW = 4096

/**
 * AcroForm /SigFlags with AppendOnly (2): the document holds signatures.
 * Only counted next to the AcroForm's /Fields, so a stray mention elsewhere
 * (or SignaturesExist alone, which unsigned fields also set) does not count.
 */
function hasAppendOnlySigFlags(bytes) {
  let index = bytes.indexOf(SIG_FLAGS)
  while (index >= 0) {
    let cursor = index + SIG_FLAGS.length
    if (endsName(bytes, cursor)) {
      while (cursor < bytes.length && PDF_WHITESPACE.has(bytes[cursor])) cursor += 1
      let value = 0
      let digits = 0
      while (cursor < bytes.length && isDigit(bytes[cursor]) && digits < 10) {
        value = value * 10 + bytes[cursor] - 0x30
        cursor += 1
        digits += 1
      }
      const window = bytes.subarray(Math.max(0, index - ACROFORM_WINDOW), Math.min(bytes.length, cursor + ACROFORM_WINDOW))
      if (digits && (value & 2) && window.includes(FIELDS)) return true
    }
    index = bytes.indexOf(SIG_FLAGS, index + SIG_FLAGS.length)
  }
  return false
}

/** Certification (DocMDP) and usage-rights (UR3) signatures. */
function hasCertificationMarker(bytes) {
  return CERTIFICATION_NAMES.some((name) => {
    let index = bytes.indexOf(name)
    while (index >= 0) {
      if (endsName(bytes, index + name.length)) return true
      index = bytes.indexOf(name, index + name.length)
    }
    return false
  })
}

/**
 * Fields and catalogs may be compressed into object streams. Inflate them
 * (bounded) and look for the same markers plus a signed /FT /Sig field.
 */
function scanObjectStreams(bytes) {
  const found = { byteRanges: 0, appendOnly: false, certified: false, signedFields: 0 }
  let inflatedTotal = 0
  let streams = 0
  let index = bytes.indexOf(OBJECT_STREAM)
  while (index >= 0 && streams < MAX_OBJECT_STREAMS && inflatedTotal < MAX_INFLATED_BYTES) {
    streams += 1
    const keyword = bytes.indexOf(STREAM_KEYWORD, index)
    if (keyword < 0) break
    if (keyword - index > 4096) {
      index = bytes.indexOf(OBJECT_STREAM, index + OBJECT_STREAM.length)
      continue
    }
    const dictionaryStart = bytes.lastIndexOf('<<', index)
    const dictionary = bytes.subarray(Math.max(0, dictionaryStart, index - 4096), keyword).toString('latin1')
    let dataStart = keyword + STREAM_KEYWORD.length
    if (bytes[dataStart] === 0x0d) dataStart += 1
    if (bytes[dataStart] === 0x0a) dataStart += 1
    const dataEnd = bytes.indexOf(END_STREAM, dataStart)
    index = bytes.indexOf(OBJECT_STREAM, Math.max(dataEnd, index + OBJECT_STREAM.length))
    if (dataEnd < 0 || !/\/FlateDecode/.test(dictionary) || /\/DecodeParms/.test(dictionary)) continue
    let inflated
    try {
      inflated = zlib.inflateSync(bytes.subarray(dataStart, dataEnd), {
        finishFlush: zlib.constants.Z_SYNC_FLUSH,
        maxOutputLength: MAX_INFLATED_BYTES - inflatedTotal,
      })
    } catch {
      continue
    }
    inflatedTotal += inflated.length
    found.byteRanges += scanByteRanges(inflated).length
    found.appendOnly ||= hasAppendOnlySigFlags(inflated)
    found.certified ||= hasCertificationMarker(inflated)
    const text = inflated.toString('latin1')
    const signedField = /\/FT\s*\/Sig(?![\w#.-])(?:(?!>>)[\s\S]){0,2000}?\/V\s*\d+\s+\d+\s+R|\/V\s*\d+\s+\d+\s+R(?:(?!<<|>>)[\s\S]){0,2000}?\/FT\s*\/Sig(?![\w#.-])/g
    found.signedFields += (text.match(signedField) || []).length
  }
  return found
}

function summarize(signatures, extra = {}) {
  const signed = signatures.length > 0 || Boolean(extra.appendOnly || extra.certified
    || extra.compressedSignatures || extra.signedFields || extra.structural)
  return {
    signed,
    count: signatures.length + (extra.compressedSignatures || 0),
    intact: signatures.some((signature) => signature.intact),
  }
}

/**
 * Detect signed PDFs from their bytes, and optionally from a parsed pdf-lib
 * document as well. `intact` is true when at least one signature still covers
 * this exact file, so overwriting it would destroy a valid signed revision.
 */
function detectSignatures(value, pdfDoc) {
  const bytes = asBuffer(value)
  const signatures = scanByteRanges(bytes)
  const compressed = scanObjectStreams(bytes)
  return summarize(signatures, {
    appendOnly: compressed.appendOnly || hasAppendOnlySigFlags(bytes),
    certified: compressed.certified || hasCertificationMarker(bytes),
    compressedSignatures: compressed.byteRanges,
    signedFields: compressed.signedFields,
    structural: pdfDoc ? documentSignatureStatus(pdfDoc).signed : false,
  })
}

/** Plain /ByteRange signatures of a file on disk, read in chunks. */
async function fileSignatureStatus(filePath, chunkSize = 4 * 1024 * 1024) {
  const handle = await fs.open(filePath, 'r')
  try {
    const { size } = await handle.stat()
    const overlap = BYTE_RANGE.length + MAX_RANGE_TEXT
    const signatures = []
    const readByte = async (position) => {
      if (position < 0 || position >= size) return -1
      const one = Buffer.alloc(1)
      await handle.read(one, 0, 1, position)
      return one[0]
    }
    const seen = new Set()
    for (let start = 0; start < size && signatures.length < MAX_SIGNATURES; start += chunkSize) {
      const length = Math.min(size - start, chunkSize + overlap)
      const chunk = Buffer.alloc(length)
      await handle.read(chunk, 0, length, start)
      let index = chunk.indexOf(BYTE_RANGE)
      while (index >= 0) {
        const absolute = start + index
        if (!seen.has(absolute)) {
          seen.add(absolute)
          const range = parseByteRange(chunk, index + BYTE_RANGE.length)
          if (range) {
            const gapStart = range[0] + range[1]
            const opening = await readByte(gapStart)
            const closing = await readByte(range[2] - 1)
            const byteAt = (position) => (position === gapStart ? opening : position === range[2] - 1 ? closing : -1)
            signatures.push({ range, intact: rangeIsIntact(range, byteAt, size) })
          }
        }
        index = chunk.indexOf(BYTE_RANGE, index + BYTE_RANGE.length)
      }
    }
    return summarize(signatures)
  } finally {
    await handle.close()
  }
}

/**
 * Structural detection on a parsed pdf-lib document: signed signature fields
 * (/FT /Sig, possibly inherited, with a /V dictionary), AcroForm /SigFlags with
 * AppendOnly, or catalog /Perms (DocMDP certification or UR usage rights).
 * Unsigned signature fields alone do not make a document signed.
 */
function documentSignatureStatus(pdfDoc) {
  const { PDFArray, PDFDict, PDFName, PDFNumber, PDFRef } = require('pdf-lib')
  const { context, catalog } = pdfDoc
  const perms = catalog.lookupMaybe(PDFName.of('Perms'), PDFDict)
  const certified = Boolean(perms && ['DocMDP', 'UR', 'UR3'].some((key) => perms.has(PDFName.of(key))))
  const acroForm = catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const sigFlagsValue = acroForm?.lookupMaybe(PDFName.of('SigFlags'), PDFNumber)
  const sigFlags = sigFlagsValue ? sigFlagsValue.asNumber() : 0
  let signedFields = 0
  const visited = new Set()
  const pending = []
  const fields = acroForm?.lookupMaybe(PDFName.of('Fields'), PDFArray)
  if (fields) for (let index = 0; index < fields.size(); index += 1) pending.push({ value: fields.get(index), type: null })
  while (pending.length && visited.size < 200_000) {
    const { value, type } = pending.pop()
    if (value instanceof PDFRef) {
      if (visited.has(value.tag)) continue
      visited.add(value.tag)
    }
    const field = context.lookup(value)
    if (!(field instanceof PDFDict)) continue
    const ownType = field.get(PDFName.of('FT'))
    const fieldType = ownType instanceof PDFName ? ownType.decodeText() : type
    if (fieldType === 'Sig' && context.lookup(field.get(PDFName.of('V'))) instanceof PDFDict) signedFields += 1
    const kids = field.lookupMaybe(PDFName.of('Kids'), PDFArray)
    if (kids) for (let index = 0; index < kids.size(); index += 1) pending.push({ value: kids.get(index), type: fieldType })
  }
  return {
    signed: certified || (sigFlags & 2) !== 0 || signedFields > 0,
    signedFields,
    certified,
    sigFlags,
  }
}

module.exports = { detectSignatures, documentSignatureStatus, fileSignatureStatus, parseByteRange }
