'use strict'
// WP2 imaging core: pixel-buffer basics (allocation, crops, exact flips and quarter turns, masked mixing,
// rectangles, premultiplied conversion), colour math and the seeded random helpers.
const test = require('node:test')
const assert = require('node:assert/strict')
const path = require('node:path')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping (not disabled by --no-experimental-strip-types), '
  + 'or run node with --experimental-strip-types.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const B = load('imaging/buffer.ts')
const C = load('imaging/color.ts')
const R = load('imaging/random.ts')

function randomBuffer(width, height, seed = 1, opaque = false) {
  const next = R.mulberry32(seed)
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i += 1) data[i] = Math.floor(next() * 256)
  if (opaque) for (let i = 3; i < data.length; i += 4) data[i] = 255
  return { width, height, data }
}

function pixel(buffer, x, y) {
  const i = (y * buffer.width + x) * 4
  return [...buffer.data.subarray(i, i + 4)]
}

function same(a, b) {
  return a.width === b.width && a.height === b.height && Buffer.compare(Buffer.from(a.data), Buffer.from(b.data)) === 0
}

test('createBuffer allocates transparent or filled pixels and rejects invalid sizes', () => {
  const empty = B.createBuffer(3, 2)
  assert.equal(empty.data.length, 24)
  assert.ok(empty.data.every((value) => value === 0))
  const filled = B.createBuffer(5, 3, { r: 10, g: 20, b: 30, a: 40 })
  for (let y = 0; y < 3; y += 1) for (let x = 0; x < 5; x += 1) assert.deepEqual(pixel(filled, x, y), [10, 20, 30, 40])
  assert.equal(B.createBuffer(0, 7).data.length, 0)
  assert.throws(() => B.createBuffer(-1, 2), RangeError)
  assert.throws(() => B.createBuffer(1.5, 2), RangeError)
  assert.throws(() => B.assertBuffer({ width: 2, height: 2, data: new Uint8ClampedArray(15) }), /does not match/)
})

test('cloneBuffer copies, cropBuffer is transparent outside, pasteBuffer overwrites with clipping', () => {
  const src = randomBuffer(6, 4, 3)
  const copy = B.cloneBuffer(src)
  assert.ok(same(copy, src))
  copy.data[0] ^= 255
  assert.notEqual(copy.data[0], src.data[0], 'the clone is independent')

  const inside = B.cropBuffer(src, { x: 1, y: 1, width: 3, height: 2 })
  assert.deepEqual(pixel(inside, 0, 0), pixel(src, 1, 1))
  assert.deepEqual(pixel(inside, 2, 1), pixel(src, 3, 2))
  const partial = B.cropBuffer(src, { x: -2, y: 3, width: 4, height: 3 })
  assert.equal(partial.width, 4)
  assert.equal(partial.height, 3)
  assert.deepEqual(pixel(partial, 0, 0), [0, 0, 0, 0])
  assert.deepEqual(pixel(partial, 2, 0), pixel(src, 0, 3))
  assert.deepEqual(pixel(partial, 3, 1), [0, 0, 0, 0], 'rows below the image are transparent')
  assert.ok(B.cropBuffer(src, { x: 10, y: 10, width: 2, height: 2 }).data.every((value) => value === 0))

  const dst = B.createBuffer(4, 4, { r: 1, g: 2, b: 3, a: 4 })
  const patch = B.createBuffer(3, 3, { r: 9, g: 9, b: 9, a: 9 })
  B.pasteBuffer(dst, patch, 2, -1)
  assert.deepEqual(pixel(dst, 2, 0), [9, 9, 9, 9])
  assert.deepEqual(pixel(dst, 3, 1), [9, 9, 9, 9])
  assert.deepEqual(pixel(dst, 2, 2), [1, 2, 3, 4], 'outside the patch is untouched')
  assert.deepEqual(pixel(dst, 1, 0), [1, 2, 3, 4])
  B.pasteBuffer(dst, patch, 10, 10)
})

test('flips twice and four quarter turns are the identity; turns move pixels the right way', () => {
  const src = randomBuffer(7, 5, 11)
  assert.ok(same(B.flipHorizontal(B.flipHorizontal(src)), src))
  assert.ok(same(B.flipVertical(B.flipVertical(src)), src))
  assert.deepEqual(pixel(B.flipHorizontal(src), 0, 2), pixel(src, 6, 2))
  assert.deepEqual(pixel(B.flipVertical(src), 3, 0), pixel(src, 3, 4))

  let turned = src
  for (let turn = 0; turn < 4; turn += 1) turned = B.rotate90(turned, true)
  assert.ok(same(turned, src), 'four clockwise quarter turns')
  turned = src
  for (let turn = 0; turn < 4; turn += 1) turned = B.rotate90(turned, false)
  assert.ok(same(turned, src), 'four counter-clockwise quarter turns')
  assert.ok(same(B.rotate90(B.rotate90(src, true), false), src))

  const cw = B.rotate90(src, true)
  assert.equal(cw.width, 5)
  assert.equal(cw.height, 7)
  assert.deepEqual(pixel(cw, 4, 0), pixel(src, 0, 0), 'clockwise: the top-left corner goes to the top-right')
  assert.deepEqual(pixel(cw, 0, 6), pixel(src, 6, 4), 'clockwise: the bottom-right corner goes to the bottom-left')
  const ccw = B.rotate90(src, false)
  assert.deepEqual(pixel(ccw, 0, 6), pixel(src, 0, 0), 'counter-clockwise: the top-left corner goes to the bottom-left')

  assert.ok(same(B.rotate180(B.rotate180(src)), src))
  assert.ok(same(B.rotate180(src), B.rotate90(B.rotate90(src, true), true)))
})

test('geometry helpers work on views that are not 4-byte aligned', () => {
  const src = randomBuffer(5, 3, 21)
  const backing = new Uint8ClampedArray(src.data.length + 1)
  backing.set(src.data, 1)
  const view = { width: 5, height: 3, data: new Uint8ClampedArray(backing.buffer, 1, src.data.length) }
  assert.ok(same(B.flipHorizontal(view), B.flipHorizontal(src)))
  assert.ok(same(B.rotate90(view, true), B.rotate90(src, true)))
  assert.ok(same(B.rotate180(view), B.rotate180(src)))
})

test('hasTransparency', () => {
  const opaque = B.createBuffer(4, 4, { r: 0, g: 0, b: 0, a: 255 })
  assert.equal(B.hasTransparency(opaque), false)
  opaque.data[4 * 9 + 3] = 254
  assert.equal(B.hasTransparency(opaque), true)
})

test('mixByMask follows the mask and opacity and mixes differing alpha without dark fringes', () => {
  const original = B.createBuffer(4, 1, { r: 0, g: 0, b: 0, a: 255 })
  const processed = B.createBuffer(4, 1, { r: 200, g: 100, b: 50, a: 255 })
  const mask = { width: 4, height: 1, data: Uint8Array.from([0, 255, 128, 64]) }
  const mixed = B.mixByMask(original, processed, mask, 1)
  assert.deepEqual(pixel(mixed, 0, 0), [0, 0, 0, 255])
  assert.deepEqual(pixel(mixed, 1, 0), [200, 100, 50, 255])
  assert.deepEqual(pixel(mixed, 2, 0), [100, 50, 25, 255])
  const half = B.mixByMask(original, processed, null, 0.5)
  assert.deepEqual(pixel(half, 3, 0), [100, 50, 25, 255])
  assert.ok(same(B.mixByMask(original, processed, null, 0), original))
  assert.ok(same(B.mixByMask(original, processed, null, 1), processed))
  assert.notEqual(B.mixByMask(original, processed, null, 1).data, processed.data, 'always a new buffer')

  const clear = B.createBuffer(1, 1)
  const white = B.createBuffer(1, 1, { r: 255, g: 255, b: 255, a: 255 })
  const edge = B.mixByMask(clear, white, null, 0.5)
  assert.deepEqual(pixel(edge, 0, 0), [255, 255, 255, 128], 'colour of the visible pixel, not grey')
  assert.throws(() => B.mixByMask(original, B.createBuffer(3, 1), null, 1), RangeError)
  assert.throws(() => B.mixByMask(original, processed, { width: 2, height: 2, data: new Uint8Array(4) }, 1), RangeError)
})

test('intersectRect and unionRect', () => {
  assert.deepEqual(B.intersectRect({ x: 0, y: 0, width: 10, height: 10 }, { x: 5, y: -5, width: 10, height: 8 }), { x: 5, y: 0, width: 5, height: 3 })
  assert.equal(B.intersectRect({ x: 0, y: 0, width: 5, height: 5 }, { x: 5, y: 0, width: 5, height: 5 }), null, 'touching edges do not overlap')
  assert.deepEqual(B.unionRect({ x: 0, y: 0, width: 2, height: 2 }, { x: -3, y: 4, width: 1, height: 1 }), { x: -3, y: 0, width: 5, height: 5 })
  assert.deepEqual(B.unionRect({ x: 50, y: 50, width: 0, height: 0 }, { x: 1, y: 2, width: 3, height: 4 }), { x: 1, y: 2, width: 3, height: 4 }, 'empty rectangles do not contribute')
})

test('premultiplied float conversion round-trips every colour at every visible alpha', () => {
  const width = 256
  const height = 255
  const data = new Uint8ClampedArray(width * height * 4)
  for (let a = 1; a <= 255; a += 1) {
    for (let c = 0; c < 256; c += 1) {
      const i = ((a - 1) * width + c) * 4
      data[i] = c
      data[i + 1] = 255 - c
      data[i + 2] = (c * 7) & 255
      data[i + 3] = a
    }
  }
  const src = { width, height, data }
  const back = B.fromPremultiplied(B.toPremultiplied(src), width, height)
  assert.ok(same(back, src), 'low-alpha colours survive (unlike a canvas round trip)')
  const transparent = { width: 1, height: 1, data: Uint8ClampedArray.from([9, 9, 9, 0]) }
  assert.deepEqual([...B.fromPremultiplied(B.toPremultiplied(transparent), 1, 1).data], [0, 0, 0, 0])
})

test('progressReporter forwards about 20 increasing values and the final 1; throwIfAborted names AbortError', () => {
  const seen = []
  const report = B.progressReporter({ onProgress: (value) => seen.push(value) })
  for (let step = 0; step <= 1000; step += 1) report(step / 1000)
  report(1)
  assert.ok(seen.length <= 21, `${seen.length} calls`)
  assert.equal(seen.at(-1), 1)
  assert.ok(seen.every((value, index) => index === 0 || value > seen[index - 1]))
  B.progressReporter(undefined)(0.5)
  const controller = new AbortController()
  B.throwIfAborted(controller.signal)
  controller.abort()
  assert.throws(() => B.throwIfAborted(controller.signal), { name: 'AbortError' })
})

test('sRGB transfer tables round-trip all 256 code values', () => {
  assert.equal(C.SRGB_TO_LINEAR.length, 256)
  assert.equal(C.SRGB_TO_LINEAR[0], 0)
  assert.equal(C.SRGB_TO_LINEAR[255], 1)
  for (let value = 0; value < 256; value += 1) assert.equal(C.linearToSrgb8(C.SRGB_TO_LINEAR[value]), value)
  assert.ok(Math.abs(C.SRGB_TO_LINEAR[128] - 0.2158605) < 1e-6)
  assert.equal(C.linearToSrgb8(2), 255)
  assert.equal(C.linearToSrgb8(-1), 0)
})

test('HSL and HSV conversions round-trip and use degrees', () => {
  assert.deepEqual(C.rgbToHsl(1, 0, 0), [0, 1, 0.5])
  assert.deepEqual(C.rgbToHsl(0, 0, 1).map((v) => +v.toFixed(6)), [240, 1, 0.5])
  assert.deepEqual(C.rgbToHsl(0.5, 0.5, 0.5), [0, 0, 0.5])
  assert.deepEqual(C.rgbToHsv(1, 0.5, 0).map((v) => +v.toFixed(6)), [30, 1, 1])
  for (let r = 0; r <= 255; r += 51) {
    for (let g = 0; g <= 255; g += 17) {
      for (let b = 0; b <= 255; b += 85) {
        const [h, s, l] = C.rgbToHsl(r / 255, g / 255, b / 255)
        const back = C.hslToRgb(h, s, l).map((v) => Math.round(v * 255))
        assert.deepEqual(back, [r, g, b])
        const [hh, ss, vv] = C.rgbToHsv(r / 255, g / 255, b / 255)
        assert.deepEqual(C.hsvToRgb(hh, ss, vv).map((v) => Math.round(v * 255)), [r, g, b])
      }
    }
  }
  assert.deepEqual(C.hslToRgb(360 + 120, 1, 0.5).map((v) => Math.round(v * 255)), [0, 255, 0], 'hue wraps')
})

test('W3C lum, setLum and setSat', () => {
  assert.ok(Math.abs(C.lum(1, 1, 1) - 1) < 1e-12)
  assert.ok(Math.abs(C.lum(1, 0, 0) - 0.3) < 1e-12)
  const shifted = C.setLum([0.2, 0.4, 0.6], 0.7)
  assert.ok(Math.abs(C.lum(...shifted) - 0.7) < 1e-9)
  const clipped = C.setLum([1, 0, 0], 0.9)
  assert.ok(clipped.every((v) => v >= 0 && v <= 1), 'ClipColor keeps the result in gamut')
  assert.ok(Math.abs(C.lum(...clipped) - 0.9) < 1e-9)
  assert.deepEqual(C.setSat([0.2, 0.5, 0.8], 0.3).map((v) => +v.toFixed(9)), [0, 0.15, 0.3])
  assert.deepEqual(C.setSat([0.4, 0.4, 0.4], 0.5), [0, 0, 0])
  assert.equal(C.sat(0.1, 0.9, 0.3), 0.8)
})

test('parseHex and toHex', () => {
  assert.deepEqual(C.parseHex('#A0B1C2'), { r: 160, g: 177, b: 194 })
  assert.deepEqual(C.parseHex(' 0f8 '), { r: 0, g: 255, b: 136 })
  assert.equal(C.parseHex('#12345'), null)
  assert.equal(C.parseHex('zzzzzz'), null)
  assert.equal(C.parseHex(''), null)
  assert.equal(C.toHex({ r: 160, g: 177, b: 194 }), '#a0b1c2')
  assert.equal(C.toHex({ r: -4, g: 300, b: 7.6 }), '#00ff08')
})

test('seeded random helpers are deterministic and spread evenly', () => {
  const a = R.mulberry32(42)
  const b = R.mulberry32(42)
  const values = Array.from({ length: 1000 }, () => a())
  assert.deepEqual(values, Array.from({ length: 1000 }, () => b()))
  assert.ok(values.every((v) => v >= 0 && v < 1))
  assert.notDeepEqual(values.slice(0, 10), Array.from({ length: 10 }, R.mulberry32(43)))

  assert.equal(R.hash2(12, 34, 5), R.hash2(12, 34, 5))
  assert.notEqual(R.hash2(12, 34, 5), R.hash2(13, 34, 5))
  assert.notEqual(R.hash2(12, 34, 5), R.hash2(12, 34, 6))
  assert.notEqual(R.hash2(1, 2, 0), R.hash2(2, 1, 0), 'x and y are not interchangeable')
  const bins = new Array(10).fill(0)
  let sum = 0
  for (let y = 0; y < 100; y += 1) {
    for (let x = 0; x < 100; x += 1) {
      const v = R.hash2(x, y, 7)
      assert.ok(v >= 0 && v < 1)
      bins[Math.floor(v * 10)] += 1
      sum += v
    }
  }
  assert.ok(Math.abs(sum / 10000 - 0.5) < 0.02)
  assert.ok(bins.every((count) => count > 850 && count < 1150), `bins ${bins}`)
  assert.equal(R.mixSeed(1, 2, 3), R.mixSeed(1, 2, 3))
  assert.notEqual(R.mixSeed(1, 2, 3), R.mixSeed(1, 3, 2))
  assert.equal(R.hashString('layer-1'), R.hashString('layer-1'))
  assert.notEqual(R.hashString('layer-1'), R.hashString('layer-2'))
})
