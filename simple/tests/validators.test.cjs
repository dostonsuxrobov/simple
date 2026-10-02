'use strict'

// Structural validators (simple/shared/electron/validators.cjs). They run on
// scratch copies of the samples in "Simple test examples", on tiny real images
// made once with Pillow (embedded below), and on deliberately damaged copies.
// A validator must pass every intact file and fail every truncated one.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')

const SHARED = path.resolve(__dirname, '..', 'shared', 'electron')
const validators = require(path.join(SHARED, 'validators.cjs'))
const { validateBytes, resolveValidator } = validators

const SAMPLES = path.resolve(__dirname, '..', '..', 'Simple test examples')
const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-validators-'))
test.after(() => fs.rmSync(ROOT, { recursive: true, force: true }))

const EXPECTED_VALIDATOR = Object.freeze({
  '.docx': 'ooxml-word',
  '.xlsx': 'ooxml-excel',
  '.ods': 'odf',
  '.xls': 'cfb',
  '.pdf': 'pdf',
  '.html': 'text',
  '.md': 'text',
})

/** Copies of the samples in a scratch folder; the originals are only read. */
function sampleCopies() {
  if (!fs.existsSync(SAMPLES)) return []
  const folder = path.join(ROOT, 'samples')
  fs.mkdirSync(folder, { recursive: true })
  return fs.readdirSync(SAMPLES)
    .filter((name) => EXPECTED_VALIDATOR[path.extname(name).toLowerCase()])
    .sort()
    .map((name) => {
      const copy = path.join(folder, name)
      if (!fs.existsSync(copy)) fs.copyFileSync(path.join(SAMPLES, name), copy)
      return { name, extension: path.extname(name).toLowerCase(), path: copy, bytes: fs.readFileSync(copy) }
    })
}

// Tiny real images written by Pillow 12.3 (6x4 pixels unless noted).
const IMAGES = Object.freeze({
  png: ['png', 'iVBORw0KGgoAAAANSUhEUgAAAAYAAAAECAYAAACtBE5DAAAAIUlEQVR4nGNUTX59ggELYAIRt+aImGOVUEt5cxKrBDYAALQcBkMtKXZcAAAAAElFTkSuQmCC'],
  apng: ['png', 'iVBORw0KGgoAAAANSUhEUgAAAAYAAAAECAYAAACtBE5DAAAACGFjVEwAAAACAAAAAPONk3AAAAAaZmNUTAAAAAAAAAAGAAAABAAAAAAAAAAAAAEACgAAI8YFEAAAAB9JREFUeJxjVE1+/Z8BC2ACEbfmiGCXUEt5g10CGwAAmV0FegPhMLQAAAAaZmNUTAAAAAEAAAAGAAAABAAAAAAAAAAAAAEACgAAuLXvxAAAACNmZEFUAAAAAnicY7yjpvafAQtgAhHKN29il7irro5dAhsAAIV2BS8wdvBxAAAAAElFTkSuQmCC'],
  jpeg: ['jpeg', '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAoHBwgHBgoICAgLCgoLDhgQDg0NDh0VFhEYIx8lJCIfIiEmKzcvJik0KSEiMEExNDk7Pj4+JS5ESUM8SDc9Pjv/2wBDAQoLCw4NDhwQEBw7KCIoOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozv/wAARCAAEAAYDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAcEAACAgIDAAAAAAAAAAAAAAABAgAxA0ERInH/xAAUAQEAAAAAAAAAAAAAAAAAAAAG/8QAFBEBAAAAAAAAAAAAAAAAAAAAAP/aAAwDAQACEQMRAD8Aju5yMGYKCAB1ULQ415e7iIj4ef/Z'],
  jpegProgressive: ['jpeg', '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAoHBwgHBgoICAgLCgoLDhgQDg0NDh0VFhEYIx8lJCIfIiEmKzcvJik0KSEiMEExNDk7Pj4+JS5ESUM8SDc9Pjv/2wBDAQoLCw4NDhwQEBw7KCIoOzs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozs7Ozv/wgARCAAEAAYDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAP/xAAVAQEBAAAAAAAAAAAAAAAAAAAEBf/aAAwDAQACEAMQAAABiDRv/8QAFRABAQAAAAAAAAAAAAAAAAAAADH/2gAIAQEAAQUCr//EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQMBAT8Bf//EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQIBAT8Bf//EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQEABj8Cf//EABgQAAIDAAAAAAAAAAAAAAAAAAABEUGB/9oACAEBAAE/IW5ViP/aAAwDAQACAAMAAAAQC//EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQMBAT8Qf//EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQIBAT8Qf//EABoQAQACAwEAAAAAAAAAAAAAAAExQQARIXH/2gAIAQEAAT8QcFAgHAg1Xk3n/9k='],
  gif: ['gif', 'R0lGODdhBgAEAIIAAABmzDNmzABm/zNm/////wAAAAAAAAAAACwAAAAABgAEAAAIFQAHBBAoYMAAAgUDCABAcKBAggMCAgA7'],
  gifAnimated: ['gif', 'R0lGODlhBgAEAIIAAABmzDNmzABm/zNm/////wAAAAAAAAAAACH/C05FVFNDQVBFMi4wAwEAAAAh+QQACgAAACwAAAAABgAEAAAIFQAHBBAoYMAAAgUDCABAcKBAggMCAgAh+QQBCgAFACwAAAAABgAEAILMAAD/AADMMzP/MzP///8AAAAAAAAAAAAIFQAFCBwooEAAAQcFABAwYGDDhQICAgA7'],
  webpLossy: ['webp', 'UklGRkYAAABXRUJQVlA4IDoAAADwAQCdASoGAAQAAoBCJagCdLoAArdntQAA/qYMA+amwHi5/p7bsYU3/pNn/E2f8TZ8kj/sWctgZOgA'],
  webpLossless: ['webp', 'UklGRiQAAABXRUJQVlA4TBgAAAAvBcAAEA9wsv+b/J/9/4H/8YD3/ET0Pxo='],
  webpAnimated: ['webp', 'UklGRpAAAABXRUJQVlA4WAoAAAASAAAABQAAAwAAQU5JTQYAAAAAAAAAAABBTk1GLgAAAAAAAAAAAAUAAAMAAGQAAAJWUDhMFgAAAC8FwAAAD/Axzpdon2J9/oMHIxDR/xBBTk1GLgAAAAAAAAAAAAUAAAMAAGQAAABWUDhMFQAAAC8FwAAQDzAmwz0m8x/weM9PRP+jAQA='],
  bmp: ['bmp', 'Qk2GAAAAAAAAADYAAAAoAAAABgAAAAQAAAABABgAAAAAAFAAAADEDgAAxA4AAAAAAAAAAAAA62Ml62Ml62Ml62Ml62Ml62MlAADrYyXrYyXrYyXrYyXrYyXrYyUAAOtjJf///+tjJetjJetjJetjJQAA62Ml62Ml62Ml62Ml62Ml62MlAAA='],
  tiff: ['tiff', 'SUkqAAgAAAAKAAABBAABAAAABgAAAAEBBAABAAAABAAAAAIBAwADAAAAhgAAAAMBAwABAAAAAQAAAAYBAwABAAAAAgAAABEBBAABAAAAjAAAABUBAwABAAAAAwAAABYBBAABAAAABAAAABcBBAABAAAASAAAABwBAwABAAAAAQAAAAAAAAAIAAgACAAlY+slY+slY+slY+slY+slY+slY+v///8lY+slY+slY+slY+slY+slY+slY+slY+slY+slY+slY+slY+slY+slY+slY+slY+s='],
  tiffLzw: ['tiff', 'SUkqACQAAACACUxuuBQSBwWEQeBv+GQmDQ+HRGFRCJxJ1wEACgAAAQMAAQAAAAYAAAABAQMAAQAAAAQAAAACAQMAAwAAAKIAAAADAQMAAQAAAAUAAAAGAQMAAQAAAAIAAAARAQQAAQAAAAgAAAAVAQMAAQAAAAMAAAAWAQMAAQAAAAQAAAAXAQQAAQAAABsAAAAcAQMAAQAAAAEAAAAAAAAACAAIAAgA'],
  ico16: ['ico', 'AAABAAEAEBAAAAAAIABrAAAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAADJJREFUeJxjVE1+/Z+BAsAEIm7NEaHMALWUN5QZQAlgokj3qAFgwMRAIWAaNYBh4MMAACnNBZJVnK9zAAAAAElFTkSuQmCC'],
})

function image(name) {
  return Buffer.from(IMAGES[name][1], 'base64')
}

function truncated(bytes, fraction = 0.9) {
  return bytes.subarray(0, Math.floor(bytes.length * fraction))
}

function assertFails(result, pattern, label) {
  assert.equal(result.ok, false, `${label} should fail`)
  if (pattern) assert.match(result.reason, pattern, label)
}

/**
 * Minimal ZIP writer for fixtures: entries are { name, data, method } with
 * method 0 (stored) or 8 (deflate). `zip64` writes ZIP64 directory records.
 */
function buildZip(entries, { zip64 = false } = {}) {
  const locals = []
  const centrals = []
  let offset = 0
  for (const entry of entries) {
    const name = Buffer.from(entry.name)
    const data = Buffer.from(entry.data)
    const method = entry.method ?? 8
    const stored = method === 8 ? zlib.deflateRawSync(data) : data
    const crc = zlib.crc32(data) >>> 0
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt16LE(method, 8)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(stored.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(name.length, 26)
    locals.push(local, name, stored)
    const extra = zip64 ? Buffer.alloc(28) : Buffer.alloc(0)
    if (zip64) {
      extra.writeUInt16LE(0x0001, 0)
      extra.writeUInt16LE(24, 2)
      extra.writeBigUInt64LE(BigInt(data.length), 4)
      extra.writeBigUInt64LE(BigInt(stored.length), 12)
      extra.writeBigUInt64LE(BigInt(offset), 20)
    }
    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(zip64 ? 45 : 20, 4)
    central.writeUInt16LE(zip64 ? 45 : 20, 6)
    central.writeUInt16LE(method, 10)
    central.writeUInt32LE(crc, 16)
    central.writeUInt32LE(zip64 ? 0xffffffff : stored.length, 20)
    central.writeUInt32LE(zip64 ? 0xffffffff : data.length, 24)
    central.writeUInt16LE(name.length, 28)
    central.writeUInt16LE(extra.length, 30)
    central.writeUInt32LE(zip64 ? 0xffffffff : offset, 42)
    centrals.push(central, name, extra)
    offset += 30 + name.length + stored.length
  }
  const directory = Buffer.concat(centrals)
  const parts = [...locals, directory]
  if (zip64) {
    const record = Buffer.alloc(56)
    record.writeUInt32LE(0x06064b50, 0)
    record.writeBigUInt64LE(44n, 4)
    record.writeUInt16LE(45, 12)
    record.writeUInt16LE(45, 14)
    record.writeBigUInt64LE(BigInt(entries.length), 24)
    record.writeBigUInt64LE(BigInt(entries.length), 32)
    record.writeBigUInt64LE(BigInt(directory.length), 40)
    record.writeBigUInt64LE(BigInt(offset), 48)
    const locator = Buffer.alloc(20)
    locator.writeUInt32LE(0x07064b50, 0)
    locator.writeBigUInt64LE(BigInt(offset + directory.length), 8)
    locator.writeUInt32LE(1, 16)
    parts.push(record, locator)
  }
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(zip64 ? 0xffff : entries.length, 8)
  end.writeUInt16LE(zip64 ? 0xffff : entries.length, 10)
  end.writeUInt32LE(zip64 ? 0xffffffff : directory.length, 12)
  end.writeUInt32LE(zip64 ? 0xffffffff : offset, 16)
  parts.push(end)
  return Buffer.concat(parts)
}

const ODS_MANIFEST = '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
const ODS_CONTENT = '<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>'

// ---------------------------------------------------------------------------
// The samples
// ---------------------------------------------------------------------------

test('every sample copy passes the validator its format names', (t) => {
  const samples = sampleCopies()
  if (!samples.length) return t.skip(`no samples in ${SAMPLES}`)
  assert.ok(samples.length >= 8, 'docx, html, md, pdf (2), ods, xls and xlsx samples are expected')
  for (const sample of samples) {
    const result = validateBytes(sample.extension, sample.bytes)
    assert.equal(result.ok, true, `${sample.name}: ${result.reason}`)
    assert.equal(result.validator, EXPECTED_VALIDATOR[sample.extension], sample.name)
    assert.ok(!result.skipped, `${sample.name} must really be checked`)
  }
})

test('every sample copy truncated to 90 % fails', async (t) => {
  const samples = sampleCopies()
  if (!samples.length) return t.skip(`no samples in ${SAMPLES}`)
  for (const sample of samples) {
    const cutPath = path.join(ROOT, 'samples', `truncated-${sample.name}`)
    fs.writeFileSync(cutPath, truncated(sample.bytes))
    if (sample.extension === '.md') {
      // Plain text has no end marker; the source text is what proves it complete.
      assert.equal((await validators.validateFile('.md', cutPath)).ok, true, 'Markdown alone cannot show truncation')
      assertFails(await validators.validateFile('.md', cutPath, { text: sample.bytes.toString('utf8') }), /shorter than the document text/, sample.name)
      continue
    }
    assertFails(await validators.validateFile(sample.extension, cutPath), /incomplete|past the end|end-of-archive|%%EOF/, sample.name)
  }
})

test('a compound file cut on a sector boundary still fails', (t) => {
  const sample = sampleCopies().find((item) => item.extension === '.xls')
  if (!sample) return t.skip('no .xls sample')
  const sectors = Math.floor((sample.bytes.length * 0.9) / 512)
  assertFails(validateBytes('xls', sample.bytes.subarray(0, sectors * 512)), /chain|past the end|shorter/, 'xls on a sector boundary')
  const damaged = Buffer.from(sample.bytes)
  damaged.writeUInt32LE(0x7fffffff, 48)
  assertFails(validateBytes('cfb', damaged), /directory chain/, 'a directory pointer past the end')
  assertFails(validateBytes('cfb', Buffer.alloc(1024)), /signature/, 'zeros')
})

test('a ZIP package with a corrupted CRC fails', (t) => {
  const sample = sampleCopies().find((item) => item.extension === '.docx')
  if (!sample) return t.skip('no .docx sample')
  const damaged = Buffer.from(sample.bytes)
  const name = Buffer.from('word/document.xml')
  let record = damaged.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]))
  while (record >= 0 && !damaged.subarray(record + 46, record + 46 + name.length).equals(name)) {
    record = damaged.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]), record + 4)
  }
  assert.ok(record > 0, 'the central record of word/document.xml exists')
  damaged.writeUInt32LE(damaged.readUInt32LE(record + 16) ^ 0x00ff00ff, record + 16)
  assertFails(validateBytes('docx', damaged), /"word\/document\.xml" that fails its CRC check/, 'docx CRC')

  const stored = buildZip([{ name: 'a.txt', data: 'hello world', method: 0 }])
  assert.equal(validateBytes('zip', stored).ok, true)
  const flipped = Buffer.from(stored)
  flipped[30 + 'a.txt'.length + 2] ^= 0x20
  assertFails(validateBytes('zip', flipped), /fails its CRC check/, 'a flipped stored byte')
  assert.equal(validateBytes('zip', flipped, { crc: 'none' }).ok, true, 'crc:none checks headers only')
})

test('packages must be the declared kind', (t) => {
  const samples = sampleCopies()
  const docx = samples.find((item) => item.extension === '.docx')
  const xlsx = samples.find((item) => item.extension === '.xlsx')
  const ods = samples.find((item) => item.extension === '.ods')
  if (!docx || !xlsx || !ods) return t.skip('docx, xlsx and ods samples are needed')
  assertFails(validateBytes('xlsx', docx.bytes), /is not an Excel workbook/, 'docx as xlsx')
  assertFails(validateBytes('pptx', xlsx.bytes), /is not a PowerPoint presentation/, 'xlsx as pptx')
  assertFails(validateBytes('odt', ods.bytes), /instead of application\/vnd\.oasis\.opendocument\.text/, 'ods as odt')
  assertFails(validateBytes('docx', ods.bytes), /\[Content_Types\]\.xml/, 'ods as docx')
  assert.equal(validateBytes('ods', ods.bytes).details.conformant, true)
  assert.deepEqual(validators.readZipEntries(docx.bytes).map((entry) => entry.name).includes('word/document.xml'), true)
  assert.match(validators.readZipEntry(docx.bytes, '[Content_Types].xml').toString('utf8'), /<Types/)
  assert.equal(validators.readZipEntry(docx.bytes, 'missing.xml'), null)
})

test('ODF written with mimetype out of place (as SheetJS does) passes with a note; strictOdf rejects it', () => {
  const bytes = buildZip([
    { name: 'META-INF/manifest.xml', data: ODS_MANIFEST },
    { name: 'mimetype', data: 'application/vnd.oasis.opendocument.spreadsheet', method: 8 },
    { name: 'content.xml', data: ODS_CONTENT },
  ])
  const result = validateBytes('ods', bytes)
  assert.equal(result.ok, true, result.reason)
  assert.deepEqual(result.details.notes, ['mimetype-not-first', 'mimetype-compressed'])
  assertFails(validateBytes('ods', bytes, { strictOdf: true }), /stored mimetype/, 'strict ODF')
  const conformant = buildZip([
    { name: 'mimetype', data: 'application/vnd.oasis.opendocument.spreadsheet', method: 0 },
    { name: 'content.xml', data: ODS_CONTENT },
  ])
  assert.equal(validateBytes('ods', conformant, { strictOdf: true }).details.conformant, true)
  assertFails(validateBytes('ods', buildZip([{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.spreadsheet', method: 0 }])), /content\.xml/, 'no content')
})

test('ZIP64 directory records are understood', () => {
  const bytes = buildZip([
    { name: '[Content_Types].xml', data: '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>' },
    { name: 'word/document.xml', data: '<w:document/>' },
  ], { zip64: true })
  const result = validateBytes('docx', bytes)
  assert.equal(result.ok, true, result.reason)
  assert.equal(result.details.mainPart, 'word/document.xml')
  assertFails(validateBytes('docx', truncated(bytes, 0.97)), /end-of-archive|ZIP64/, 'cut ZIP64')
})

// ---------------------------------------------------------------------------
// Images
// ---------------------------------------------------------------------------

test('every generated image passes, and every truncated copy fails', () => {
  for (const [name, [validator]] of Object.entries(IMAGES)) {
    const bytes = image(name)
    const result = validateBytes(validator, bytes)
    assert.equal(result.ok, true, `${name}: ${result.reason}`)
    assertFails(validateBytes(validator, truncated(bytes)), null, `${name} truncated`)
    assertFails(validateBytes(validator, bytes.subarray(0, 3)), null, `${name} cut to 3 bytes`)
  }
  assert.equal(validateBytes('jpeg', image('jpegProgressive')).details.progressive, true)
  assert.equal(validateBytes('gif', image('gifAnimated')).details.frames, 2)
  assert.equal(validateBytes('webp', image('webpAnimated')).details.kind, 'VP8X')
})

test('a PNG without IEND, or with a damaged chunk, fails', () => {
  const png = image('png')
  assertFails(validateBytes('png', png.subarray(0, png.length - 12)), /IEND/, 'no IEND')
  const damaged = Buffer.from(png)
  damaged[40] ^= 0x01
  assertFails(validateBytes('png', damaged), /CRC/, 'flipped IDAT byte')
  assertFails(validateBytes('png', Buffer.concat([png, Buffer.from('trailing')])), /after its IEND/, 'data after IEND')
  assertFails(validateBytes('png', image('jpeg')), /PNG signature/, 'a JPEG named .png')
})

test('damaged JPEG, GIF, WebP, BMP, TIFF and ICO files fail', () => {
  const jpeg = image('jpeg')
  assertFails(validateBytes('jpeg', jpeg.subarray(0, jpeg.length - 2)), /end marker/, 'JPEG without EOI')
  assertFails(validateBytes('jpeg', Buffer.concat([jpeg, Buffer.from('MP4 data')])), /after its end marker/, 'JPEG with appended data')
  assert.equal(validateBytes('jpeg', Buffer.concat([jpeg, Buffer.alloc(16)])).ok, true, 'zero padding after EOI is tolerated')
  const gif = image('gif')
  assertFails(validateBytes('gif', gif.subarray(0, gif.length - 1)), /trailer/, 'GIF without trailer')
  const webp = Buffer.from(image('webpLossy'))
  webp.writeUInt32LE(webp.readUInt32LE(4) + 2, 4)
  assertFails(validateBytes('webp', webp), /length its header declares/, 'WebP size')
  const bmp = Buffer.from(image('bmp'))
  bmp.writeUInt32LE(bmp.length + 100, 2)
  assertFails(validateBytes('bmp', bmp), /length its header declares/, 'BMP size')
  const tiff = Buffer.from(image('tiff'))
  tiff.writeUInt32LE(tiff.length + 50, 4)
  assertFails(validateBytes('tiff', tiff), /directory/, 'TIFF directory offset past the end')
  const ico = Buffer.from(image('ico16'))
  ico.writeUInt32LE(ico.length, 6 + 12)
  assertFails(validateBytes('ico', ico), /past the end/, 'ICO image offset past the end')
})

// ---------------------------------------------------------------------------
// PDF, text and JSON
// ---------------------------------------------------------------------------

test('PDF structure: header, end marker and a startxref that points at the xref table', (t) => {
  const sample = sampleCopies().find((item) => item.extension === '.pdf')
  if (!sample) return t.skip('no PDF sample')
  assertFails(validateBytes('pdf', Buffer.from(`junk${sample.bytes.toString('latin1').slice(5)}`, 'latin1')), /%PDF-/, 'no header')
  const text = sample.bytes.toString('latin1')
  const at = text.lastIndexOf('startxref')
  const offset = Number(/startxref\s+(\d+)/.exec(text.slice(at))[1])
  const wrongPointer = Buffer.from(text.slice(0, at) + text.slice(at).replace(String(offset), String(offset + 7)), 'latin1')
  assertFails(validateBytes('pdf', wrongPointer), /does not point at a cross-reference table/, 'startxref off by 7')
  const pastEnd = Buffer.from(text.slice(0, at) + text.slice(at).replace(String(offset), '99999999'), 'latin1')
  assertFails(validateBytes('pdf', pastEnd), /past the end/, 'startxref past the end')
  const minimal = '%PDF-1.4\n1 0 obj\n<< /Type /Catalog >>\nendobj\nxref\n0 2\n0000000000 65535 f \n0000000009 00000 n \ntrailer\n<< /Size 2 /Root 1 0 R >>\nstartxref\n44\n%%EOF\n'
  assert.equal(validateBytes('pdf', Buffer.from(minimal)).ok, true)
})

test('text: encodings, byte-order marks and the source text', () => {
  assert.equal(validateBytes('txt', Buffer.from('plain words\n')).details.encoding, 'utf-8')
  assertFails(validateBytes('txt', Buffer.from([0x61, 0xc3, 0x28]), { encoding: 'utf-8' }), /valid UTF-8/, 'invalid UTF-8')
  assert.equal(validateBytes('csv', Buffer.from([0x63, 0x61, 0x66, 0xe9]), {}).details.encoding, 'unknown', 'undeclared legacy bytes are not judged')
  assert.equal(validateBytes('csv', Buffer.from([0x63, 0x61, 0x66, 0xe9]), { encoding: 'windows-1252', text: 'café' }).ok, true)
  const utf16 = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('a;b\r\n1;2\r\n', 'utf16le')])
  assert.equal(validateBytes('csv', utf16, { text: 'a;b\r\n1;2\r\n' }).details.encoding, 'utf-16le')
  assertFails(validateBytes('csv', utf16.subarray(0, utf16.length - 1)), /odd number of bytes/, 'odd UTF-16')
  assertFails(validateBytes('csv', utf16, { encoding: 'utf-8' }), /byte-order mark/, 'BOM disagrees with the declared encoding')
  assertFails(validateBytes('txt', Buffer.from('Hello wor'), { text: 'Hello world' }), /shorter/, 'truncated text')
  assertFails(validateBytes('txt', Buffer.from('Hello there'), { text: 'Hello world' }), /does not match/, 'wrong text')
  assert.equal(validateBytes('txt', Buffer.from('\uFEFFHi', 'utf8'), { text: '\uFEFFHi' }).ok, true)
  const sjis = Buffer.from([0x82, 0xa0, 0x82])
  assertFails(validateBytes('txt', sjis, { encoding: 'shift_jis' }), /valid shift_jis/, 'cut Shift-JIS')
})

test('markup and JSON must reach their closing structure', () => {
  const html = '<!DOCTYPE html>\n<html lang="en"><head><title>x</title></head><body><p>Hello</p></body></html>\n'
  assert.equal(validateBytes('html', Buffer.from(html)).ok, true)
  assertFails(validateBytes('.htm', Buffer.from(html.slice(0, -20))), /<\/html>/, 'cut HTML')
  assert.equal(validateBytes('html', Buffer.from('<table><tr><td>1</td></tr></table>')).ok, true, 'a fragment has no closing html tag to check')
  const xml = '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n<!-- made by Simple -->\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet/></Workbook>\n'
  assert.equal(validateBytes('xml', Buffer.from(xml)).ok, true)
  assertFails(validateBytes('xml', Buffer.from(xml.slice(0, -14))), /<\/Workbook>/, 'cut XML')
  assert.equal(validateBytes('svg', Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"/>')).ok, true)
  const rtf = '{\\rtf1\\ansi{\\fonttbl{\\f0 Arial;}}\\f0 Braces \\{ and \\} stay text.\\par}'
  assert.equal(validateBytes('rtf', Buffer.from(rtf)).ok, true)
  assertFails(validateBytes('rtf', Buffer.from(rtf.slice(0, -1))), /closing brace/, 'cut RTF')
  assert.equal(validateBytes('json', Buffer.from('\uFEFF{"a":[1,2,3]}', 'utf8')).ok, true)
  assertFails(validateBytes('json', Buffer.from('{"a":[1,2')), /valid JSON/, 'cut JSON')
  assertFails(validateBytes('json', Buffer.from('{ // comment\n"a": 1 }')), /valid JSON/, 'JSON with comments')
})

test('large text is checked in streamed chunks instead of being decoded whole', () => {
  const line = 'Größe,Value,naïve café ✓\n'
  const text = line.repeat(4000)
  const big = Buffer.from(text)
  const streamed = { maxWholeTextBytes: 1024 }
  assert.equal(validateBytes('csv', big, { ...streamed, encoding: 'utf-8' }).ok, true)
  const broken = Buffer.from(big)
  broken[50_000] = 0xff
  assertFails(validateBytes('csv', broken, { ...streamed, encoding: 'utf-8' }), /valid UTF-8/, 'an invalid byte in the middle')
  assertFails(validateBytes('csv', big.subarray(0, big.length - 2), { ...streamed, encoding: 'utf-8' }), /valid UTF-8/, 'cut inside a character')
  assert.deepEqual(validateBytes('csv', big, { ...streamed, text }).details.notes, ['the source text was not compared because the file is very large'])
  const utf16 = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')])
  assert.equal(validateBytes('csv', utf16, streamed).details.encoding, 'utf-16le')
  const html = Buffer.from(`<!DOCTYPE html><html><body>${'<p>row</p>'.repeat(5000)}</body></html>\n`)
  assert.equal(validateBytes('html', html, streamed).ok, true)
  assertFails(validateBytes('html', html.subarray(0, html.length - 10), streamed), /<\/html>/, 'large cut HTML')
  const json = Buffer.from(JSON.stringify({ rows: Array.from({ length: 5000 }, (_, index) => index) }))
  assert.equal(validateBytes('json', json, streamed).ok, true)
  assertFails(validateBytes('json', json.subarray(0, json.length - 3), streamed), /complete JSON/, 'large cut JSON')
  const xml = Buffer.from(`<?xml version="1.0"?><Workbook>${'<Row/>'.repeat(5000)}</Workbook>`)
  assert.equal(validateBytes('xml', xml, streamed).ok, true)
  assertFails(validateBytes('xml', xml.subarray(0, xml.length - 4), streamed), /<\/Workbook>/, 'large cut XML')
})

// ---------------------------------------------------------------------------
// Resolution
// ---------------------------------------------------------------------------

test('format ids, extensions and validator ids resolve; unknown formats are skipped, never failed', () => {
  assert.deepEqual(resolveValidator('docx'), { validator: 'ooxml-word', format: 'docx' })
  assert.deepEqual(resolveValidator('.XLSX'), { validator: 'ooxml-excel', format: 'xlsx' })
  assert.deepEqual(resolveValidator('ooxml-powerpoint'), { validator: 'ooxml-powerpoint', format: null })
  assert.equal(resolveValidator('png').validator, 'png')
  assert.equal(resolveValidator('.jpg').validator, 'jpeg')
  assert.equal(resolveValidator('.htm').validator, 'text')
  assert.equal(resolveValidator('nope'), null)
  assert.deepEqual(validateBytes('nope', Buffer.from('x')), { ok: true, validator: null, format: null, skipped: true })
  assert.equal(validateBytes('none', Buffer.from('anything')).ok, true)

  const registry = JSON.parse(fs.readFileSync(path.join(SHARED, 'formats.json'), 'utf8'))
  const named = new Set(registry.formats.map((format) => format.validator).filter(Boolean))
  for (const validator of named) assert.ok(validators.VALIDATOR_IDS.includes(validator), `formats.json names an unknown validator "${validator}"`)
})

test('validators accept Uint8Array and ArrayBuffer input and do not throw on garbage', () => {
  const png = image('png')
  const view = new Uint8Array(png.buffer.slice(png.byteOffset, png.byteOffset + png.length))
  assert.equal(validateBytes('png', view).ok, true)
  assert.equal(validateBytes('png', view.buffer).ok, true)
  for (const validator of validators.VALIDATOR_IDS) {
    for (const garbage of [Buffer.alloc(0), Buffer.from([0xff]), Buffer.alloc(600, 0xff), Buffer.from('PK\u0005\u0006')]) {
      assert.doesNotThrow(() => validateBytes(validator, garbage), `${validator} on ${garbage.length} bytes`)
    }
  }
})
