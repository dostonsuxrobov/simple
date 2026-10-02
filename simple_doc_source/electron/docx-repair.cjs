'use strict'

// Repairs applied to DOCX packages written by the editor before they reach the
// user's disk. The editor's writer labels every picture it does not recognize
// as "imageN.png" with an image/png content type, and (WordCanvas 0.12.0)
// writes mid-document section breaks with empty header/footer relationship ids
// (r:id=""), which Word reports as unreadable content. Unchanged packages are
// returned byte-for-byte; only a package that needs a repair is rewritten.

const { inflateRawSync, constants: zlibConstants } = require('node:zlib')
const JSZip = require('jszip')
const { listZipEntries, validateDocxBytes } = require('./docx-files.cjs')

const IMAGE_SIGNATURES = Object.freeze([
  { extension: 'png', contentType: 'image/png', test: (b) => b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47 },
  { extension: 'jpeg', contentType: 'image/jpeg', test: (b) => b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff },
  { extension: 'gif', contentType: 'image/gif', test: (b) => b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38 },
  { extension: 'bmp', contentType: 'image/bmp', test: (b) => b[0] === 0x42 && b[1] === 0x4d && b.length >= 14 },
  { extension: 'tiff', contentType: 'image/tiff', test: (b) => (b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 0x2a) },
  { extension: 'webp', contentType: 'image/webp', test: (b) => b.subarray(0, 4).toString('latin1') === 'RIFF' && b.subarray(8, 12).toString('latin1') === 'WEBP' },
  { extension: 'emf', contentType: 'image/x-emf', test: (b) => b.length >= 44 && b.readUInt32LE(0) === 1 && b.subarray(40, 44).toString('latin1') === ' EMF' },
  { extension: 'wmf', contentType: 'image/x-wmf', test: (b) => (b.length >= 4 && b.readUInt32LE(0) === 0x9ac6cdd7) || ((b[0] === 1 || b[0] === 2) && b[1] === 0 && b[2] === 9 && b[3] === 0) },
  { extension: 'ico', contentType: 'image/x-icon', test: (b) => b[0] === 0 && b[1] === 0 && b[2] === 1 && b[3] === 0 },
  { extension: 'avif', contentType: 'image/avif', test: (b) => b.subarray(4, 8).toString('latin1') === 'ftyp' && /^avi[fs]/.test(b.subarray(8, 12).toString('latin1')) },
  { extension: 'svg', contentType: 'image/svg+xml', test: (b) => /^(?:﻿)?\s*(?:<\?xml[^>]*>\s*)?(?:<!--[\s\S]*?-->\s*)*(?:<!DOCTYPE svg[^>]*>\s*)?<svg[\s>]/i.test(b.subarray(0, 1024).toString('utf8')) },
])
const EXTENSION_ALIASES = Object.freeze({ jpg: 'jpeg', jpe: 'jpeg', tif: 'tiff' })

function detectImage(bytes) {
  if (!bytes || bytes.length < 4) return null
  return IMAGE_SIGNATURES.find((signature) => signature.test(bytes)) || null
}

/** The first decompressed bytes of an entry without inflating all of it. */
function peekEntry(bytes, entry, count = 1024) {
  const offset = entry.localHeaderOffset
  if (offset + 30 > bytes.length || bytes.readUInt32LE(offset) !== 0x04034b50) return null
  const dataOffset = offset + 30 + bytes.readUInt16LE(offset + 26) + bytes.readUInt16LE(offset + 28)
  const compressed = bytes.subarray(dataOffset, Math.min(bytes.length, dataOffset + entry.compressedSize))
  if (entry.method === 0) return compressed.subarray(0, count)
  if (entry.method !== 8) return null
  try {
    return inflateRawSync(compressed.subarray(0, Math.min(compressed.length, Math.max(4096, count * 4))), { finishFlush: zlibConstants.Z_SYNC_FLUSH }).subarray(0, count)
  } catch { return null }
}

function inflateEntry(bytes, entry) {
  const offset = entry.localHeaderOffset
  const dataOffset = offset + 30 + bytes.readUInt16LE(offset + 26) + bytes.readUInt16LE(offset + 28)
  const compressed = bytes.subarray(dataOffset, dataOffset + entry.compressedSize)
  if (entry.method === 0) return Buffer.from(compressed)
  return inflateRawSync(compressed, { maxOutputLength: Math.max(1, entry.uncompressedSize) })
}

const EMPTY_BAND_REFERENCE = /<w:(header|footer)Reference\b[^>]*\br:id=""[^>]*\/>/g

function describeEmptyBands(xml) {
  const sections = [...xml.matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)]
  const affected = new Set()
  sections.forEach((section, index) => {
    if (/<w:(?:header|footer)Reference\b[^>]*\br:id=""/.test(section[0])) affected.add(index + 1)
  })
  return [...affected]
}

/**
 * Inspect an exported DOCX and repair mislabeled media and empty header/footer
 * references. Returns { bytes, changed, repairs, warnings }.
 */
async function repairDocxPackage(data) {
  const bytes = validateDocxBytes(data)
  const entries = listZipEntries(bytes)
  const relabel = []
  for (const entry of entries.values()) {
    const match = /^word\/media\/([^/]+)\.([a-z0-9]+)$/i.exec(entry.name)
    if (!match) continue
    const head = peekEntry(bytes, entry)
    const detected = detectImage(head)
    if (!detected) continue
    const current = EXTENSION_ALIASES[match[2].toLowerCase()] || match[2].toLowerCase()
    if (current !== detected.extension) relabel.push({ name: entry.name, base: match[1], detected })
  }
  const documentEntry = entries.get('word/document.xml')
  let documentXml = null
  let emptySections = []
  if (documentEntry) {
    try {
      documentXml = inflateEntry(bytes, documentEntry).toString('utf8')
      emptySections = describeEmptyBands(documentXml)
    } catch { documentXml = null }
  }
  if (!relabel.length && !emptySections.length) return { bytes, changed: false, repairs: [], warnings: [] }

  const zip = await JSZip.loadAsync(bytes)
  const repairs = []
  const warnings = []
  if (relabel.length) {
    const renamed = new Map()
    const taken = new Set(Object.keys(zip.files))
    for (const item of relabel) {
      let target = `word/media/${item.base}.${item.detected.extension}`
      for (let suffix = 2; taken.has(target); suffix += 1) target = `word/media/${item.base}-${suffix}.${item.detected.extension}`
      taken.add(target)
      const file = zip.file(item.name)
      const content = await file.async('nodebuffer')
      zip.remove(item.name)
      zip.file(target, content, { date: file.date, compression: 'DEFLATE', createFolders: false })
      renamed.set(item.name.slice('word/'.length), target.slice('word/'.length))
    }
    // Every relationship part under word/ addresses media relative to word/.
    for (const name of Object.keys(zip.files).filter((entry) => /^word\/_rels\/[^/]+\.rels$/.test(entry))) {
      const file = zip.file(name)
      let xml = await file.async('string')
      let touched = false
      for (const [from, to] of renamed) {
        const pattern = new RegExp(`(Target=")(?:\\./)?${from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(")`, 'g')
        const next = xml.replace(pattern, `$1${to}$2`)
        if (next !== xml) { xml = next; touched = true }
      }
      if (touched) zip.file(name, xml, { date: file.date, createFolders: false })
    }
    const typesFile = zip.file('[Content_Types].xml')
    let types = await typesFile.async('string')
    for (const extension of [...new Set(relabel.map((item) => item.detected.extension))]) {
      if (new RegExp(`<Default\\b[^>]*Extension="${extension}"`, 'i').test(types)) continue
      const contentType = relabel.find((item) => item.detected.extension === extension).detected.contentType
      types = types.replace(/<Types\b[^>]*>/, (open) => `${open}<Default Extension="${extension}" ContentType="${contentType}"/>`)
    }
    zip.file('[Content_Types].xml', types, { date: typesFile.date, createFolders: false })
    repairs.push(`${relabel.length} ${relabel.length === 1 ? 'picture was' : 'pictures were'} relabeled with ${relabel.length === 1 ? 'its' : 'their'} real image format.`)
  }
  if (emptySections.length && documentXml !== null) {
    const repairedXml = documentXml.replace(EMPTY_BAND_REFERENCE, '')
    const file = zip.file('word/document.xml')
    zip.file('word/document.xml', repairedXml, { date: file.date, createFolders: false })
    const list = emptySections.join(', ')
    repairs.push(`Removed empty header/footer references from section ${list}.`)
    warnings.push(`The header or footer of section ${list} could not be saved, so Word will show that section without it. Check that section’s header and footer before sharing the file.`)
  }
  const output = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 }, platform: 'DOS' })
  return { bytes: validateDocxBytes(output), changed: true, repairs, warnings }
}

module.exports = { detectImage, peekEntry, repairDocxPackage }
