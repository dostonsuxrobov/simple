const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const Module = require('node:module')
const { pathToFileURL } = require('node:url')
const { imageDimensions, validateImageBytes } = require('../electron/image-files.cjs')
const {
  crc32,
  extractMetadata,
  fidelityWarnings,
  iccInfo,
  patchExif,
  readJpegSegments,
  readPngChunks,
  readRiffChunks,
  spliceMetadata,
} = require('../electron/image-metadata.cjs')

// #region builders

const TYPE_SIZES = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 7: 1 }

/** Serializes IFDs ({ ifd0, exif, gps, ifd1 }) into a TIFF block; '@exif'/'@gps'/'@thumb' are resolved offsets. */
function buildTiff({ ifd0, exif = null, gps = null, ifd1 = null, thumbnail = null }, little = true) {
  const order = [['ifd0', ifd0], ['exif', exif], ['gps', gps], ['ifd1', ifd1]].filter(([, entries]) => entries)
  const valueBytes = (entry) => {
    if (Buffer.isBuffer(entry.values)) return entry.values
    const size = TYPE_SIZES[entry.type]
    const buffer = Buffer.alloc(entry.type === 5 ? entry.values.length * 4 : entry.values.length * size)
    entry.values.forEach((value, index) => {
      if (entry.type === 3) little ? buffer.writeUInt16LE(value, index * 2) : buffer.writeUInt16BE(value, index * 2)
      else if (entry.type === 4 || entry.type === 5) little ? buffer.writeUInt32LE(value, index * 4) : buffer.writeUInt32BE(value, index * 4)
      else buffer[index] = value
    })
    return buffer
  }
  const layout = {}
  let offset = 8
  for (const [name, entries] of order) {
    const data = entries.map((entry) => (typeof entry.values?.[0] === 'string' ? Buffer.alloc(4) : valueBytes(entry)))
    const extra = data.reduce((total, bytes) => total + (bytes.length > 4 ? bytes.length + (bytes.length & 1) : 0), 0)
    layout[name] = { offset, size: 2 + entries.length * 12 + 4 + extra }
    offset += layout[name].size
  }
  layout.thumb = { offset }
  const tiff = Buffer.alloc(offset + (thumbnail ? thumbnail.length : 0))
  const w16 = (o, v) => (little ? tiff.writeUInt16LE(v, o) : tiff.writeUInt16BE(v, o))
  const w32 = (o, v) => (little ? tiff.writeUInt32LE(v, o) : tiff.writeUInt32BE(v, o))
  tiff.write(little ? 'II' : 'MM', 0, 'latin1')
  w16(2, 42)
  w32(4, 8)
  for (const [name, entries] of order) {
    const start = layout[name].offset
    let dataOffset = start + 2 + entries.length * 12 + 4
    w16(start, entries.length)
    entries.forEach((entry, index) => {
      const at = start + 2 + index * 12
      w16(at, entry.tag)
      w16(at + 2, entry.type)
      if (typeof entry.values?.[0] === 'string') {
        w32(at + 4, 1)
        w32(at + 8, layout[entry.values[0].slice(1)].offset)
        return
      }
      const bytes = valueBytes(entry)
      w32(at + 4, Buffer.isBuffer(entry.values) ? entry.values.length : entry.type === 5 ? entry.values.length / 2 : entry.values.length)
      if (bytes.length <= 4) bytes.copy(tiff, at + 8)
      else {
        w32(at + 8, dataOffset)
        bytes.copy(tiff, dataOffset)
        dataOffset += bytes.length + (bytes.length & 1)
      }
    })
    w32(start + 2 + entries.length * 12, name === 'ifd0' && ifd1 ? layout.ifd1.offset : 0)
  }
  if (thumbnail) thumbnail.copy(tiff, layout.thumb.offset)
  return tiff
}

function cameraExif(little = true) {
  const thumbnail = Buffer.concat([Buffer.from([0xff, 0xd8]), Buffer.alloc(300, 0xab), Buffer.from([0xff, 0xd9])])
  return buildTiff({
    ifd0: [
      { tag: 0x010f, type: 2, values: Buffer.from('Canon\0') },
      { tag: 0x0112, type: 3, values: [6] },
      { tag: 0x011a, type: 5, values: [300, 1] },
      { tag: 0x011b, type: 5, values: [300, 1] },
      { tag: 0x0128, type: 3, values: [2] },
      { tag: 0x8769, type: 4, values: ['@exif'] },
      { tag: 0x8825, type: 4, values: ['@gps'] },
    ],
    exif: [
      { tag: 0x9003, type: 2, values: Buffer.from('2024:05:06 07:08:09\0') },
      { tag: 0xa002, type: 4, values: [4000] },
      { tag: 0xa003, type: 3, values: [3000] },
    ],
    gps: [{ tag: 0x0000, type: 1, values: [2, 3, 0, 0] }],
    ifd1: [
      { tag: 0x0103, type: 3, values: [6] },
      { tag: 0x0201, type: 4, values: ['@thumb'] },
      { tag: 0x0202, type: 4, values: [thumbnail.length] },
    ],
    thumbnail,
  }, little)
}

/** Independent TIFF reader for assertions: { ifd0, exif, nextIfd } with tag -> value (numbers, strings). */
function readTiff(tiff) {
  const little = tiff.toString('latin1', 0, 2) === 'II'
  const u16 = (o) => (little ? tiff.readUInt16LE(o) : tiff.readUInt16BE(o))
  const u32 = (o) => (little ? tiff.readUInt32LE(o) : tiff.readUInt32BE(o))
  const readIfd = (offset) => {
    const tags = {}
    const count = u16(offset)
    for (let index = 0; index < count; index += 1) {
      const at = offset + 2 + index * 12
      const tag = u16(at)
      const type = u16(at + 2)
      const n = u32(at + 4)
      const size = (TYPE_SIZES[type] || 1) * n
      const valueAt = size > 4 ? u32(at + 8) : at + 8
      if (type === 2) tags[tag] = tiff.toString('latin1', valueAt, valueAt + n).replace(/\0+$/, '')
      else if (type === 3) tags[tag] = u16(valueAt)
      else if (type === 4) tags[tag] = u32(valueAt)
      else if (type === 5) tags[tag] = u32(valueAt) / u32(valueAt + 4)
    }
    return { tags, next: u32(offset + 2 + count * 12) }
  }
  const ifd0 = readIfd(u32(4))
  return { ifd0: ifd0.tags, exif: ifd0.tags[0x8769] ? readIfd(ifd0.tags[0x8769]).tags : {}, nextIfd: ifd0.next }
}

function iccProfile(description, colorSpace = 'RGB ') {
  const text = Buffer.from(`${description}\0`, 'latin1')
  const desc = Buffer.alloc(12 + text.length + 67)
  desc.write('desc', 0, 'latin1')
  desc.writeUInt32BE(text.length, 8)
  text.copy(desc, 12)
  const profile = Buffer.alloc(144 + desc.length)
  profile.writeUInt32BE(profile.length, 0)
  profile.write('mntr', 12, 'latin1')
  profile.write(colorSpace, 16, 'latin1')
  profile.write('XYZ ', 20, 'latin1')
  profile.write('acsp', 36, 'latin1')
  profile.writeUInt32BE(1, 128)
  profile.write('desc', 132, 'latin1')
  profile.writeUInt32BE(144, 136)
  profile.writeUInt32BE(desc.length, 140)
  desc.copy(profile, 144)
  return profile
}

const SRGB_ICC = iccProfile('sRGB IEC61966-2.1')
const ADOBE_ICC = iccProfile('Adobe RGB (1998)')

function jpegSegment(marker, payload) {
  const header = Buffer.from([0xff, marker, 0, 0])
  header.writeUInt16BE(payload.length + 2, 2)
  return Buffer.concat([header, payload])
}

function jfifPayload(units, x, y) {
  const data = Buffer.from([0x4a, 0x46, 0x49, 0x46, 0, 1, 1, units, 0, 0, 0, 0, 0, 0])
  data.writeUInt16BE(x, 8)
  data.writeUInt16BE(y, 10)
  return data
}

function makeJpeg({ width = 64, height = 48, jfif = [0, 1, 1], app = [], precision = 8, components = 3 } = {}) {
  const sof = Buffer.alloc(6 + components * 3)
  sof[0] = precision
  sof.writeUInt16BE(height, 1)
  sof.writeUInt16BE(width, 3)
  sof[5] = components
  return Buffer.concat([
    Buffer.from([0xff, 0xd8]),
    ...(jfif ? [jpegSegment(0xe0, jfifPayload(...jfif))] : []),
    ...app,
    jpegSegment(0xdb, Buffer.alloc(65, 1)),
    jpegSegment(0xc0, sof),
    jpegSegment(0xc4, Buffer.alloc(20, 2)),
    jpegSegment(0xda, Buffer.from([1, 1, 0, 0, 0x3f, 0])),
    Buffer.from([0x12, 0x34, 0xff, 0x00, 0x56, 0xff, 0xd9]),
  ])
}

const XMP = '<x:xmpmeta xmlns:x="adobe:ns:meta/"><rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#"><rdf:Description xmlns:tiff="http://ns.adobe.com/tiff/1.0/" xmlns:exif="http://ns.adobe.com/exif/1.0/" xmlns:dc="http://purl.org/dc/elements/1.1/" tiff:Orientation="6" exif:PixelXDimension="4000"><dc:creator>Ada</dc:creator></rdf:Description></rdf:RDF></x:xmpmeta>'

function photoshopIrb() {
  const resource = (id, data) => {
    const header = Buffer.alloc(12)
    header.write('8BIM', 0, 'latin1')
    header.writeUInt16BE(id, 4)
    header.writeUInt32BE(data.length, 8)
    return Buffer.concat([header, data, Buffer.alloc(data.length & 1)])
  }
  return Buffer.concat([resource(0x0404, Buffer.from([0x1c, 0x02, 0x78, 0, 5, 0x48, 0x65, 0x6c, 0x6c, 0x6f])), resource(0x0409, Buffer.alloc(40, 9))])
}

function cameraJpeg({ icc = SRGB_ICC, little = true } = {}) {
  const half = Math.ceil(icc.length / 2)
  const iccSegment = (index, piece) => jpegSegment(0xe2, Buffer.concat([Buffer.from('ICC_PROFILE\0', 'latin1'), Buffer.from([index, 2]), piece]))
  return makeJpeg({
    width: 4000,
    height: 3000,
    jfif: [1, 300, 300],
    app: [
      jpegSegment(0xe1, Buffer.concat([Buffer.from('Exif\0\0', 'latin1'), cameraExif(little)])),
      jpegSegment(0xe1, Buffer.concat([Buffer.from('http://ns.adobe.com/xap/1.0/\0', 'latin1'), Buffer.from(XMP)])),
      iccSegment(2, icc.subarray(half)),
      iccSegment(1, icc.subarray(0, half)),
      jpegSegment(0xed, Buffer.concat([Buffer.from('Photoshop 3.0\0', 'latin1'), photoshopIrb()])),
    ],
  })
}

function pngChunk(type, data = Buffer.alloc(0)) {
  const header = Buffer.alloc(8)
  header.writeUInt32BE(data.length, 0)
  header.write(type, 4, 'latin1')
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(header.subarray(4), data), 0)
  return Buffer.concat([header, data, crc])
}

function makePng({ width = 4, height = 3, bitDepth = 8, before = [], after = [] } = {}) {
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(width, 0)
  ihdr.writeUInt32BE(height, 4)
  ihdr[8] = bitDepth
  ihdr[9] = 6
  const rowBytes = width * 4 * (bitDepth / 8)
  const raw = Buffer.alloc((rowBytes + 1) * height, 0x40)
  for (let row = 0; row < height; row += 1) raw[row * (rowBytes + 1)] = 0
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', ihdr),
    ...before,
    pngChunk('IDAT', zlib.deflateSync(raw)),
    ...after,
    pngChunk('IEND'),
  ])
}

function phys(ppm) {
  const data = Buffer.alloc(9)
  data.writeUInt32BE(ppm, 0)
  data.writeUInt32BE(ppm, 4)
  data[8] = 1
  return data
}

function richPng() {
  const exif = buildTiff({ ifd0: [{ tag: 0x010f, type: 2, values: Buffer.from('Nikon\0') }, { tag: 0x0112, type: 3, values: [8] }] })
  return makePng({
    bitDepth: 16,
    before: [
      pngChunk('iCCP', Buffer.concat([Buffer.from('sRGB profile\0\0', 'latin1'), zlib.deflateSync(SRGB_ICC)])),
      pngChunk('pHYs', phys(11811)),
      pngChunk('eXIf', exif),
      pngChunk('tEXt', Buffer.from('Title\0Harbour at dusk', 'latin1')),
      pngChunk('iTXt', Buffer.concat([Buffer.from('XML:com.adobe.xmp\0\0\0\0\0', 'latin1'), Buffer.from(XMP)])),
    ],
    after: [
      pngChunk('zTXt', Buffer.concat([Buffer.from('Comment\0\0', 'latin1'), zlib.deflateSync(Buffer.from('compressed note'))])),
      pngChunk('iTXt', Buffer.concat([Buffer.from('Description\0\0\0en\0\0', 'latin1'), Buffer.from('Ünïcode text')])),
    ],
  })
}

function riffChunk(type, data) {
  const header = Buffer.alloc(8)
  header.write(type, 0, 'latin1')
  header.writeUInt32LE(data.length, 4)
  return Buffer.concat(data.length & 1 ? [header, data, Buffer.alloc(1)] : [header, data])
}

function riff(chunks) {
  const payload = Buffer.concat([Buffer.from('WEBP'), ...chunks])
  const header = Buffer.alloc(8)
  header.write('RIFF', 0, 'latin1')
  header.writeUInt32LE(payload.length, 4)
  return Buffer.concat([header, payload])
}

function vp8l(width, height, alpha) {
  const data = Buffer.alloc(17)
  data[0] = 0x2f
  data.writeUInt32LE(((width - 1) | ((height - 1) << 14) | (alpha ? 1 << 28 : 0)) >>> 0, 1)
  return data
}

function pngChunkTypes(bytes) {
  return readPngChunks(bytes).map((chunk) => chunk.type)
}

function assertPngCrcs(bytes) {
  let offset = 8
  while (offset < bytes.length) {
    const length = bytes.readUInt32BE(offset)
    const expected = crc32(bytes.subarray(offset + 4, offset + 8 + length))
    assert.equal(bytes.readUInt32BE(offset + 8 + length), expected, `CRC of ${bytes.toString('latin1', offset + 4, offset + 8)}`)
    offset += 12 + length
  }
  assert.equal(offset, bytes.length)
}

// #endregion builders

test('crc32 matches the standard check value', () => {
  assert.equal(crc32(Buffer.from('123456789')), 0xcbf43926)
})

test('EXIF patching resets orientation, updates the size and drops the IFD1 thumbnail (both byte orders)', () => {
  for (const little of [true, false]) {
    const source = cameraExif(little)
    const { tiff, info } = patchExif(source, { width: 48, height: 64 })
    assert.equal(info.orientation, 6)
    assert.equal(info.hasThumbnail, true)
    assert.deepEqual(info.resolution, { x: 300, y: 300, unit: 'in' })
    assert.ok(tiff.length < source.length - 300, 'thumbnail bytes are cut off')
    const parsed = readTiff(tiff)
    assert.equal(parsed.ifd0[0x0112], 1)
    assert.equal(parsed.ifd0[0x010f], 'Canon')
    assert.equal(parsed.ifd0[0x011a], 300)
    assert.equal(parsed.exif[0x9003], '2024:05:06 07:08:09')
    assert.equal(parsed.exif[0xa002], 48)
    assert.equal(parsed.exif[0xa003], 64)
    assert.equal(parsed.nextIfd, 0)
    assert.equal(readTiff(source).ifd0[0x0112], 6, 'the input is not modified')
  }
  assert.throws(() => patchExif(Buffer.from('not tiff data')), /byte order/)
})

test('JPEG -> JPEG keeps EXIF (orientation 1), XMP, sRGB ICC, IPTC and density', () => {
  const source = extractMetadata(cameraJpeg())
  assert.equal(source.format, 'jpeg')
  assert.equal(source.orientation, 6)
  assert.deepEqual(source.icc, SRGB_ICC, 'ICC chunks are reassembled in sequence order')
  assert.deepEqual(source.density, { x: 300, y: 300, unit: 'in' })
  const output = makeJpeg({ width: 3000, height: 4000 })
  const result = spliceMetadata(output, source)
  assert.deepEqual(result.kept, ['EXIF', 'XMP', 'IPTC', 'ICC profile (sRGB IEC61966-2.1)', 'Resolution (DPI)'])
  assert.deepEqual(result.dropped, [])
  assert.doesNotThrow(() => validateImageBytes(result.bytes, '.jpg'))
  assert.deepEqual(imageDimensions(result.bytes, '.jpg'), { width: 3000, height: 4000 })
  const { segments, rest } = readJpegSegments(result.bytes)
  assert.equal(segments[0].marker, 0xe0, 'JFIF stays first')
  assert.ok(rest.equals(readJpegSegments(output).rest), 'scan data is untouched')
  const back = extractMetadata(result.bytes)
  const exif = readTiff(back.exif)
  assert.equal(exif.ifd0[0x0112], 1)
  assert.equal(exif.ifd0[0x010f], 'Canon')
  assert.equal(exif.exif[0x9003], '2024:05:06 07:08:09')
  assert.deepEqual([exif.exif[0xa002], exif.exif[0xa003], exif.nextIfd], [3000, 4000, 0])
  assert.match(back.xmp.toString(), /tiff:Orientation="1"/)
  assert.match(back.xmp.toString(), /exif:PixelXDimension="3000"/)
  assert.match(back.xmp.toString(), /<dc:creator>Ada<\/dc:creator>/)
  assert.deepEqual(back.icc, SRGB_ICC)
  assert.deepEqual(back.density, { x: 300, y: 300, unit: 'in' })
  assert.ok(back.iptc.includes(Buffer.from([0x1c, 0x02, 0x78])), 'IPTC record kept')
  assert.equal(back.iptc.indexOf(Buffer.from([0x38, 0x42, 0x49, 0x4d, 0x04, 0x09])), -1, 'Photoshop thumbnail resource dropped')
  // Splicing again (e.g. a second save) does not duplicate blocks.
  const twice = spliceMetadata(result.bytes, back)
  assert.equal(readJpegSegments(twice.bytes).segments.filter((segment) => segment.marker === 0xe1).length, 2)
})

test('wide-gamut ICC profiles are not re-attached to sRGB pixels unless asked', () => {
  assert.equal(iccInfo(SRGB_ICC).srgb, true)
  assert.equal(iccInfo(ADOBE_ICC).srgb, false)
  const source = extractMetadata(cameraJpeg({ icc: ADOBE_ICC, little: false }))
  const auto = spliceMetadata(makeJpeg(), source)
  assert.equal(extractMetadata(auto.bytes).icc, null)
  assert.ok(auto.dropped.includes('ICC profile (Adobe RGB (1998)): colors were converted to sRGB'))
  assert.equal(readTiff(extractMetadata(auto.bytes).exif).ifd0[0x0112], 1, 'big-endian EXIF patched too')
  const kept = spliceMetadata(makeJpeg(), source, { colorProfile: 'keep' })
  assert.deepEqual(extractMetadata(kept.bytes).icc, ADOBE_ICC)
  const dropped = spliceMetadata(makeJpeg(), extractMetadata(cameraJpeg()), { colorProfile: 'drop' })
  assert.equal(extractMetadata(dropped.bytes).icc, null)
})

test('PNG -> PNG keeps pHYs, tEXt/zTXt/iTXt, eXIf (orientation 1), XMP and the sRGB profile', () => {
  const source = extractMetadata(richPng())
  assert.equal(source.bitDepth, 16)
  assert.equal(source.pngText.length, 3)
  assert.match(fidelityWarnings(source).join(' '), /16 bits per channel; the saved image has 8 bits per channel/)
  const output = makePng({ width: 3, height: 4 })
  const result = spliceMetadata(output, source)
  assertPngCrcs(result.bytes)
  assert.doesNotThrow(() => validateImageBytes(result.bytes, '.png'))
  const types = pngChunkTypes(result.bytes)
  assert.equal(types[0], 'IHDR')
  const firstIdat = types.indexOf('IDAT')
  for (const type of ['iCCP', 'pHYs', 'eXIf']) assert.ok(types.indexOf(type) > 0 && types.indexOf(type) < firstIdat, `${type} before IDAT`)
  assert.equal(types.filter((type) => type === 'iTXt').length, 2)
  const back = extractMetadata(result.bytes)
  assert.deepEqual(back.density, { x: 11811, y: 11811, unit: 'm' })
  assert.deepEqual(back.icc, SRGB_ICC)
  assert.equal(readTiff(back.exif).ifd0[0x0112], 1)
  assert.equal(readTiff(back.exif).ifd0[0x010f], 'Nikon')
  assert.match(back.xmp.toString(), /tiff:Orientation="1"/)
  assert.deepEqual(back.pngText.map((chunk) => chunk.data.toString('latin1').split('\0')[0]), ['Title', 'Comment', 'Description'])
  assert.equal(back.bitDepth, 8, 'pixels stay as encoded')
  const idat = readPngChunks(result.bytes).find((chunk) => chunk.type === 'IDAT').data
  assert.equal(zlib.inflateSync(idat).length, (3 * 4 + 1) * 4)
})

test('JPEG -> PNG and PNG -> JPEG carry EXIF, XMP and DPI across formats', () => {
  const toPng = spliceMetadata(makePng({ width: 48, height: 64 }), extractMetadata(cameraJpeg()))
  assertPngCrcs(toPng.bytes)
  const pngBack = extractMetadata(toPng.bytes)
  assert.equal(readTiff(pngBack.exif).ifd0[0x0112], 1)
  assert.equal(readTiff(pngBack.exif).exif[0xa002], 48)
  assert.deepEqual(pngBack.density, { x: 11811, y: 11811, unit: 'm' })
  assert.ok(toPng.dropped.includes('IPTC'))
  const toJpeg = spliceMetadata(makeJpeg(), extractMetadata(richPng()))
  const jpegBack = extractMetadata(toJpeg.bytes)
  assert.equal(readTiff(jpegBack.exif).ifd0[0x010f], 'Nikon')
  assert.deepEqual(jpegBack.density, { x: 300, y: 300, unit: 'in' })
  assert.ok(toJpeg.dropped.includes('PNG text'))
})

test('WebP output is wrapped in VP8X with ICCP, EXIF and XMP chunks', () => {
  const output = riff([riffChunk('VP8L', vp8l(30, 20, true))])
  const result = spliceMetadata(output, extractMetadata(cameraJpeg()))
  assert.doesNotThrow(() => validateImageBytes(result.bytes, '.webp'))
  assert.equal(result.bytes.readUInt32LE(4), result.bytes.length - 8, 'RIFF size')
  const chunks = readRiffChunks(result.bytes)
  assert.deepEqual(chunks.map((chunk) => chunk.type), ['VP8X', 'ICCP', 'VP8L', 'EXIF', 'XMP '])
  const vp8x = chunks[0].data
  assert.equal(vp8x[0], 0x20 | 0x10 | 0x08 | 0x04, 'ICC, alpha, EXIF and XMP flags')
  assert.deepEqual([vp8x.readUIntLE(4, 3) + 1, vp8x.readUIntLE(7, 3) + 1], [30, 20])
  assert.deepEqual(imageDimensions(result.bytes, '.webp'), { width: 30, height: 20 })
  const back = extractMetadata(result.bytes)
  assert.equal(readTiff(back.exif).ifd0[0x0112], 1)
  assert.equal(readTiff(back.exif).exif[0xa002], 30)
  assert.deepEqual(back.icc, SRGB_ICC)
  assert.match(back.xmp.toString(), /dc:creator/)

  // WebP source whose EXIF chunk carries the "Exif\0\0" prefix some writers add.
  const sourceWebp = riff([
    riffChunk('VP8X', Buffer.from([0x08, 0, 0, 0, 29, 0, 0, 19, 0, 0])),
    riffChunk('VP8L', vp8l(30, 20, false)),
    riffChunk('EXIF', Buffer.concat([Buffer.from('Exif\0\0', 'latin1'), cameraExif()])),
  ])
  const again = spliceMetadata(riff([riffChunk('VP8L', vp8l(20, 30, false))]), extractMetadata(sourceWebp))
  const exif = readTiff(extractMetadata(again.bytes).exif)
  assert.deepEqual([exif.ifd0[0x0112], exif.exif[0xa002], exif.exif[0xa003]], [1, 20, 30])
  assert.equal(readRiffChunks(again.bytes)[0].data[0], 0x08)
})

test('damaged or missing metadata never breaks the encoded output', () => {
  const brokenExif = makeJpeg({ app: [jpegSegment(0xe1, Buffer.concat([Buffer.from('Exif\0\0', 'latin1'), Buffer.from('II*\0\xff\xff\xff\x00', 'latin1')]))] })
  const output = makeJpeg()
  const result = spliceMetadata(output, extractMetadata(brokenExif))
  assert.ok(result.dropped.includes('EXIF (damaged)'))
  assert.equal(extractMetadata(result.bytes).exif, null)
  assert.equal(spliceMetadata(output, extractMetadata(makeJpeg())).bytes, output, 'nothing to add: output unchanged')
  assert.equal(spliceMetadata(Buffer.from('GIF89a'), extractMetadata(cameraJpeg())).bytes.toString(), 'GIF89a')
  const garbage = extractMetadata(Buffer.from([0xff, 0xd8, 0xff, 0xe1, 0xff, 0xff]))
  assert.equal(garbage.format, 'jpeg')
  assert.ok(garbage.problems.length > 0)
  const cmyk = extractMetadata(makeJpeg({ components: 4 }))
  assert.equal(cmyk.colorModel, 'cmyk')
  assert.match(fidelityWarnings(cmyk).join(' '), /CMYK colors were converted to RGB/)
  assert.deepEqual(fidelityWarnings(extractMetadata(makePng())), [])
})

// #region electron/main.cjs file:save with a stubbed electron module

async function loadMain(temporaryRoot, dialogResults) {
  const handlers = new Map()
  const dialogCalls = []
  const stub = () => new Proxy(function stubbed() {}, {
    get: (_target, key) => (key === 'then' ? undefined : key === 'isDestroyed' ? () => false : stub()),
    apply: () => stub(),
    construct: () => stub(),
  })
  const electron = {
    app: { isPackaged: true, requestSingleInstanceLock: () => true, on: () => {}, quit: () => {}, getVersion: () => '0', whenReady: () => Promise.resolve() },
    BrowserWindow: Object.assign(function BrowserWindow() { return stub() }, {
      fromWebContents: () => ({ isDestroyed: () => false }),
      getAllWindows: () => [],
      getFocusedWindow: () => null,
    }),
    clipboard: {},
    dialog: {
      showOpenDialog: async () => ({ canceled: true }),
      showSaveDialog: async (_window, options) => { dialogCalls.push(options); return dialogResults.shift() || { canceled: true } },
    },
    ipcMain: { handle: (channel, handler) => handlers.set(channel, handler), on: () => {} },
    nativeImage: {},
    shell: {},
  }
  const previous = [process.env.TEMP, process.env.TMP]
  process.env.TEMP = temporaryRoot
  process.env.TMP = temporaryRoot
  const originalLoad = Module._load
  Module._load = function load(request, parent, isMain) {
    return request === 'electron' ? electron : originalLoad.call(this, request, parent, isMain)
  }
  try {
    const mainPath = require.resolve('../electron/main.cjs')
    delete require.cache[mainPath]
    require(mainPath)
    await new Promise((resolve) => setTimeout(resolve, 50))
  } finally {
    Module._load = originalLoad
    process.env.TEMP = previous[0]
    process.env.TMP = previous[1]
  }
  const rendererUrl = pathToFileURL(path.join(__dirname, '..', 'dist', 'index.html')).href
  const sender = { id: 7, getURL: () => rendererUrl, once: () => {} }
  return { dialogCalls, invoke: (channel, ...args) => handlers.get(channel)({ sender, senderFrame: { url: rendererUrl } }, ...args) }
}

test('file:save splices source metadata, reports fidelity warnings and protects animations', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-save-'))
  try {
    const dialogResults = []
    const { invoke, dialogCalls } = await loadMain(directory, dialogResults)
    const photo = path.join(directory, 'photo.jpg')
    await fs.writeFile(photo, cameraJpeg())
    const opened = await invoke('file:open-path', photo)
    assert.deepEqual([opened.orientation, opened.width, opened.height], [6, 3000, 4000], 'upright size for EXIF orientation 6')

    // Edited in-place save with the token: metadata kept, orientation reset.
    const edited = makeJpeg({ width: 3000, height: 4000 })
    const saved = await invoke('file:save', { data: new Uint8Array(edited), path: photo, name: 'photo.jpg', format: 'jpeg', forceDialog: false, sourceToken: opened.sourceToken })
    assert.equal(saved.path, photo)
    assert.deepEqual(saved.warnings, [])
    assert.ok(saved.metadata.kept.includes('EXIF'))
    const written = extractMetadata(await fs.readFile(photo))
    assert.equal(readTiff(written.exif).ifd0[0x0112], 1)
    assert.equal(readTiff(written.exif).ifd0[0x010f], 'Canon')

    // Without a token the in-place file itself is the metadata source (the current renderer sends none).
    await fs.writeFile(photo, cameraJpeg())
    await invoke('file:save', { data: new Uint8Array(edited), path: photo, name: 'photo.jpg', format: 'jpeg', forceDialog: false })
    assert.equal(readTiff(extractMetadata(await fs.readFile(photo)).exif).exif[0x9003], '2024:05:06 07:08:09')

    // preserveMetadata: false strips it.
    await fs.writeFile(photo, cameraJpeg())
    await invoke('file:save', { data: new Uint8Array(edited), path: photo, name: 'photo.jpg', format: 'jpeg', forceDialog: false, preserveMetadata: false })
    assert.deepEqual(await fs.readFile(photo), edited)

    // An unchanged save keeps the original bytes exactly (no double splice, orientation untouched).
    const original = cameraJpeg()
    await fs.writeFile(photo, original)
    await invoke('file:save', { data: new Uint8Array(original), path: photo, name: 'photo.jpg', format: 'jpeg', forceDialog: false })
    assert.deepEqual(await fs.readFile(photo), original)

    // A 16-bit PNG source: Save As (dialog) with the token reports the bit-depth reduction.
    const deep = path.join(directory, 'deep.png')
    await fs.writeFile(deep, richPng())
    const deepOpened = await invoke('file:open-path', deep)
    assert.equal(deepOpened.bitDepth, 16)
    const copy = path.join(directory, 'deep copy.png')
    dialogResults.push({ canceled: false, filePath: copy })
    const deepSaved = await invoke('file:save', { data: new Uint8Array(makePng()), path: null, name: 'deep.png', format: 'png', forceDialog: true, sourceToken: deepOpened.sourceToken })
    assert.equal(deepSaved.path, copy)
    assert.match(deepSaved.warnings.join(' '), /16 bits per channel/)
    assert.equal(dialogCalls.at(-1).defaultPath, path.join(directory, 'deep.png'), 'Save As starts in the source folder')
    assert.equal(extractMetadata(await fs.readFile(copy)).pngText.length, 3)

    // An animated PNG is never overwritten in place by a still frame.
    const animation = Buffer.concat([
      makePng().subarray(0, 33),
      pngChunk('acTL', Buffer.from([0, 0, 0, 3, 0, 0, 0, 0])),
      makePng().subarray(33),
    ])
    const animated = path.join(directory, 'anim.png')
    await fs.writeFile(animated, animation)
    const animatedOpened = await invoke('file:open-path', animated)
    assert.deepEqual([animatedOpened.animated, animatedOpened.frameCount, animatedOpened.directSave], [true, 3, false])
    const still = await invoke('file:save', { data: new Uint8Array(makePng()), path: animated, name: 'anim.png', format: 'png', forceDialog: false })
    assert.equal(still, null, 'the dialog was cancelled')
    assert.equal(dialogCalls.at(-1).defaultPath, path.join(directory, 'anim (frame 1).png'))
    assert.equal(dialogCalls.at(-1).title, 'Save a still copy')
    assert.deepEqual(await fs.readFile(animated), animation, 'the animation is untouched')
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
})

// #endregion
