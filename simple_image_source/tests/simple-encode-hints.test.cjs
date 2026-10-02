'use strict'
// WP8 / IMAGE-SIE-14: re-encoding hints read from the source bytes. A lossless WebP (VP8L) must be saved
// at quality 1 (Chromium's only lossless setting); a JPEG's quality is estimated from its luminance
// quantization table so Save keeps roughly the source quality instead of a fixed 0.92.
const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const E = load('simple/encodeHints.ts')

function chunk(type, payload) {
  const size = Buffer.alloc(4)
  size.writeUInt32LE(payload.length)
  return Buffer.concat([Buffer.from(type, 'latin1'), size, payload, payload.length & 1 ? Buffer.alloc(1) : Buffer.alloc(0)])
}

function riff(...chunks) {
  const body = Buffer.concat([Buffer.from('WEBP', 'latin1'), ...chunks])
  const size = Buffer.alloc(4)
  size.writeUInt32LE(body.length)
  return new Uint8Array(Buffer.concat([Buffer.from('RIFF', 'latin1'), size, body]))
}

/** A minimal JPEG header with the luminance table libjpeg writes for `quality`, in zigzag order. */
function jpegWithQuality(quality, extraTables = false) {
  const natural = E.scaledLuminanceTable(quality)
  const zigzag = E.JPEG_NATURAL_ORDER.map((index) => natural[index])
  const tables = [0x00, ...zigzag]
  if (extraTables) tables.unshift(0x01, ...new Array(64).fill(99))
  const dqt = Buffer.from([0xff, 0xdb, ((tables.length + 2) >> 8) & 0xff, (tables.length + 2) & 0xff, ...tables])
  const app0 = Buffer.from([0xff, 0xe0, 0x00, 0x10, 0x4a, 0x46, 0x49, 0x46, 0x00, 0x01, 0x01, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00])
  return new Uint8Array(Buffer.concat([Buffer.from([0xff, 0xd8]), app0, dqt, Buffer.from([0xff, 0xda, 0x00, 0x02, 0xff, 0xd9])]))
}

test('lossless WebP is recognised from its first image chunk', () => {
  assert.equal(E.isLosslessWebp(riff(chunk('VP8L', Buffer.from([0x2f, 0, 0, 0, 0])))), true)
  assert.equal(E.isLosslessWebp(riff(chunk('VP8 ', Buffer.alloc(10)))), false)
  assert.equal(E.isLosslessWebp(riff(chunk('VP8X', Buffer.alloc(10)), chunk('ICCP', Buffer.alloc(7)), chunk('VP8L', Buffer.alloc(5)))), true)
  assert.equal(E.isLosslessWebp(riff(chunk('VP8X', Buffer.alloc(10)), chunk('ALPH', Buffer.alloc(9)), chunk('VP8 ', Buffer.alloc(10)))), false)
  const frame = Buffer.concat([Buffer.alloc(16), chunk('VP8L', Buffer.alloc(5))])
  assert.equal(E.isLosslessWebp(riff(chunk('VP8X', Buffer.alloc(10)), chunk('ANIM', Buffer.alloc(6)), chunk('ANMF', frame))), true)
  assert.equal(E.isLosslessWebp(new Uint8Array(Buffer.from('not a webp at all, really'))), false)
})

test('JPEG quality is estimated from the IJG-scaled luminance table', () => {
  for (const quality of [25, 50, 75, 85, 92, 95, 100]) {
    assert.equal(E.estimateJpegQuality(jpegWithQuality(quality)), quality, `quality ${quality}`)
  }
  assert.equal(E.estimateJpegQuality(jpegWithQuality(85, true)), 85, 'the luminance table is found among several')
  assert.equal(E.estimateJpegQuality(new Uint8Array([0x89, 0x50, 0x4e, 0x47])), null)
  assert.deepEqual(E.scaledLuminanceTable(50).slice(0, 4), [16, 11, 10, 16])
})

test('save quality: lossless WebP at 1, JPEG at its estimate clamped to 0.75..0.98', () => {
  assert.equal(E.saveQuality('png', E.NO_ENCODE_HINTS), undefined)
  assert.equal(E.saveQuality('webp', { webpLossless: true, jpegQuality: null }), 1)
  assert.equal(E.saveQuality('webp', E.NO_ENCODE_HINTS), E.DEFAULT_LOSSY_QUALITY)
  assert.equal(E.saveQuality('jpeg', { webpLossless: false, jpegQuality: 95 }), 0.95)
  assert.equal(E.saveQuality('jpeg', { webpLossless: false, jpegQuality: 40 }), 0.75)
  assert.equal(E.saveQuality('jpeg', { webpLossless: false, jpegQuality: 100 }), 0.98)
  assert.equal(E.saveQuality('jpeg', E.NO_ENCODE_HINTS), E.DEFAULT_LOSSY_QUALITY)
  assert.deepEqual(E.encodeHintsFor('jpg', jpegWithQuality(80)), { webpLossless: false, jpegQuality: 80 })
  assert.deepEqual(E.encodeHintsFor('webp', riff(chunk('VP8L', Buffer.alloc(5)))), { webpLossless: true, jpegQuality: null })
  assert.deepEqual(E.encodeHintsFor('png', new Uint8Array(4)), E.NO_ENCODE_HINTS)
})
