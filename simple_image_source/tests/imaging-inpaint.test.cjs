'use strict'
// WP4 spot healing and inpainting (src/imaging/inpaint.ts): exact results on flat colour, linear ramps,
// harmonic fields and periodic texture; Telea on thin scratches; fallbacks; the 'heal' worker handler.
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

const { healMargin, healRegion, inpaintTelea } = load('imaging/inpaint.ts')
const { handlers } = load('shared/worker-ops/mask.ts')

function image(width, height, colour) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) data.set(colour(x, y), (y * width + x) * 4)
  }
  return { width, height, data }
}

/** Paints a blemish over the clean image wherever `inside` holds and returns the stroke mask. */
function blemish(picture, inside, colour = [255, 0, 255, 255]) {
  const { width, height } = picture
  const hole = { width, height, data: new Uint8Array(width * height) }
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      if (!inside(x, y)) continue
      hole.data[y * width + x] = 255
      picture.data.set(colour, (y * width + x) * 4)
    }
  }
  return hole
}

/** Mean and max absolute RGB error over the pixels where `where` holds. */
function errors(actual, expected, where) {
  let total = 0
  let count = 0
  let worst = 0
  for (let y = 0; y < actual.height; y += 1) {
    for (let x = 0; x < actual.width; x += 1) {
      if (!where(x, y)) continue
      for (let c = 0; c < 3; c += 1) {
        const e = Math.abs(actual.data[(y * actual.width + x) * 4 + c] - expected.data[(y * actual.width + x) * 4 + c])
        total += e
        worst = Math.max(worst, e)
        count += 1
      }
    }
  }
  return { mean: total / count, worst }
}

const everywhere = () => true

test('healing a square hole in a flat colour is exact', () => {
  const flat = (x, y) => [90, 140, 200, 255]
  const clean = image(160, 160, flat)
  const damaged = image(160, 160, flat)
  const hole = blemish(damaged, (x, y) => x >= 70 && x < 90 && y >= 70 && y < 90)
  const healed = healRegion(damaged, hole)
  assert.deepEqual(healed.data, clean.data)
  assert.notEqual(healed.data, damaged.data, 'returns a new buffer')
  assert.equal(damaged.data[(80 * 160 + 80) * 4], 255, 'the input is not modified')
})

test('healing a square hole in a linear gradient has a mean error of at most 6 levels', () => {
  const ramp = (x, y) => [Math.round((x * 255) / 159), Math.round((y * 200) / 159), 128, 255]
  const clean = image(160, 160, ramp)
  const damaged = image(160, 160, ramp)
  const inside = (x, y) => x >= 70 && x < 90 && y >= 70 && y < 90
  const hole = blemish(damaged, inside)
  const healed = healRegion(damaged, hole)
  const result = errors(healed, clean, (x, y) => x >= 66 && x < 94 && y >= 66 && y < 94)
  assert.ok(result.mean <= 6, `mean ${result.mean}`)
  assert.ok(result.mean < 1 && result.worst <= 2, `a ramp is reproduced almost exactly (mean ${result.mean}, worst ${result.worst})`)
  // Pixels away from the dilated hole are untouched.
  const far = errors(healed, clean, (x, y) => x < 60 || x >= 100 || y < 60 || y >= 100)
  assert.equal(far.worst, 0)
})

/** u = x^2 - y^2 is harmonic, so source - target differences are harmonic too and the seamless clone
 *  must recover the field up to rounding. */
function harmonic(centre) {
  return (x, y) => {
    const u = ((x - centre) ** 2 - (y - centre) ** 2) / 900
    return [Math.round(128 + u), Math.round(60 + x * 0.3), Math.round(200 - y * 0.25), 255]
  }
}

test('large holes converge: a harmonic field is reproduced across a 120 px hole', () => {
  // Exercises the multigrid Laplace solver (the ring candidates fit in this image).
  const size = 640
  const field = harmonic(320)
  const clean = image(size, size, field)
  const damaged = image(size, size, field)
  const inside = (x, y) => (x - 320) ** 2 + (y - 320) ** 2 <= 60 * 60
  const hole = blemish(damaged, inside)
  const started = performance.now()
  const healed = healRegion(damaged, hole)
  const elapsed = performance.now() - started
  const result = errors(healed, clean, inside)
  assert.ok(result.mean < 1 && result.worst <= 2, `mean ${result.mean}, worst ${result.worst}`)
  assert.ok(elapsed < 5000, `took ${elapsed} ms`)
})

test('when no source ring fits the image, the in-image offsets are scanned instead', () => {
  // 420 px: ring 1 overlaps the hole and its band, rings 1.5+ leave the image; other offsets fit.
  const size = 420
  const field = harmonic(210)
  const clean = image(size, size, field)
  const damaged = image(size, size, field)
  const inside = (x, y) => (x - 210) ** 2 + (y - 210) ** 2 <= 60 * 60
  const hole = blemish(damaged, inside)
  const result = errors(healRegion(damaged, hole), clean, inside)
  assert.ok(result.mean < 1 && result.worst <= 2, `mean ${result.mean}, worst ${result.worst}`)
})

test('healing in a periodic texture copies a matching source exactly', () => {
  const stripes = (x, y) => [x % 8 < 4 ? 30 : 220, 100 + (y % 5) * 10, 50, 255]
  const clean = image(200, 200, stripes)
  const damaged = image(200, 200, stripes)
  const hole = blemish(damaged, (x, y) => (x - 100) ** 2 + (y - 100) ** 2 <= 144)
  assert.deepEqual(healRegion(damaged, hole).data, clean.data)
})

test('Telea fills a 2 px scratch', () => {
  const smooth = (x, y) => [40 + x, 60 + y, Math.round(100 + 40 * Math.sin(x / 15)), 255]
  const clean = image(120, 80, smooth)
  const damaged = image(120, 80, smooth)
  const scratch = (x, y) => y >= 39 && y < 41 && x >= 10 && x < 110
  const hole = blemish(damaged, scratch, [255, 255, 255, 255])
  const direct = inpaintTelea(damaged, hole)
  const telea = errors(direct, clean, scratch)
  assert.ok(telea.worst <= 2, `Telea worst ${telea.worst}`)
  assert.equal(errors(direct, clean, (x, y) => !scratch(x, y)).worst, 0, 'only the hole changes')
  // The spot-healing entry point routes thin strokes to Telea as well.
  const viaHeal = healRegion(damaged, hole)
  const healed = errors(viaHeal, clean, (x, y) => y >= 36 && y < 44 && x >= 6 && x < 114)
  assert.ok(healed.worst <= 2, `heal worst ${healed.worst}`)
  // A diagonal one-pixel wire on a flat background disappears completely.
  const flat = image(60, 60, () => [10, 120, 30, 255])
  const wire = blemish(image(60, 60, () => [10, 120, 30, 255]), (x, y) => x === y && x > 5 && x < 55, [0, 0, 0, 255])
  const wired = image(60, 60, (x, y) => (x === y && x > 5 && x < 55 ? [0, 0, 0, 255] : [10, 120, 30, 255]))
  assert.deepEqual(healRegion(wired, wire).data, flat.data)
})

test('falls back to inpainting when no source fits, and handles trivial masks', () => {
  const ramp = (x, y) => [x * 6, y * 6, 100, 255]
  const clean = image(40, 40, ramp)
  const damaged = image(40, 40, ramp)
  const inside = (x, y) => x >= 10 && x < 30 && y >= 10 && y < 30
  const hole = blemish(damaged, inside, [0, 0, 0, 255])
  const healed = healRegion(damaged, hole)
  const result = errors(healed, clean, inside)
  assert.ok(result.mean < 2, `fallback mean ${result.mean}`)
  // Empty and full masks return a copy.
  const empty = healRegion(clean, { width: 40, height: 40, data: new Uint8Array(1600) })
  assert.deepEqual(empty.data, clean.data)
  assert.notEqual(empty.data, clean.data)
  assert.deepEqual(healRegion(clean, { width: 40, height: 40, data: new Uint8Array(1600).fill(255) }).data, clean.data)
  assert.deepEqual(inpaintTelea(clean, { width: 40, height: 40, data: new Uint8Array(1600) }).data, clean.data)
  assert.throws(() => healRegion(clean, { width: 39, height: 40, data: new Uint8Array(1560) }), RangeError)
  assert.throws(() => inpaintTelea({ width: 2, height: 2, data: new Uint8ClampedArray(4) }, { width: 2, height: 2, data: new Uint8Array(4) }), RangeError)
})

test('results are deterministic, healMargin covers the search, progress is bounded', () => {
  const noise = (x, y) => {
    const h = Math.imul(x * 374761393 + y * 668265263, 1274126177) >>> 0
    return [h & 255, (h >>> 8) & 255, (h >>> 16) & 255, 255]
  }
  const picture = image(220, 220, noise)
  const hole = blemish(image(220, 220, noise), (x, y) => (x - 110) ** 2 + (y - 110) ** 2 <= 100)
  const fractions = []
  const first = healRegion(picture, hole, { seed: 3 }, { onProgress: (value) => fractions.push(value) })
  assert.deepEqual(healRegion(picture, hole, { seed: 3 }).data, first.data)
  assert.ok(fractions.length <= 20, `${fractions.length} progress calls`)
  assert.equal(fractions.at(-1), 1)
  const margin = healMargin({ x: 100, y: 100, width: 20, height: 20 })
  assert.ok(margin >= 3 * 24 && margin < 200, `margin ${margin}`)
})

test('the heal worker handler matches the direct call and honours cancellation', async () => {
  const ramp = (x, y) => [x * 2, y, 77, 255]
  const damaged = image(100, 90, ramp)
  const hole = blemish(damaged, (x, y) => x >= 40 && x < 52 && y >= 40 && y < 50)
  const direct = healRegion(damaged, hole, {})
  const viaWorker = await handlers.heal({
    src: { width: 100, height: 90, data: Uint8ClampedArray.from(damaged.data) },
    hole: { width: 100, height: 90, data: Uint8Array.from(hole.data) },
    options: {},
  }, {})
  assert.deepEqual(viaWorker.data, direct.data)
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(handlers.heal({ src: damaged, hole, options: {} }, { signal: controller.signal }), { name: 'AbortError' })
  // Cancelling while running stops with an AbortError too.
  const big = image(400, 400, (x, y) => [x % 256, y % 256, 9, 255])
  const bigHole = blemish(image(400, 400, () => [0, 0, 0, 255]), (x, y) => (x - 200) ** 2 + (y - 200) ** 2 <= 70 * 70)
  const running = new AbortController()
  assert.throws(() => healRegion(big, bigHole, {}, { signal: running.signal, onProgress: () => running.abort() }), { name: 'AbortError' })
})
