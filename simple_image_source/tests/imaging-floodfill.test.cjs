'use strict'
// WP4 flood fill (src/imaging/floodFill.ts): Magic Wand / Paint Bucket tolerance, contiguity
// (4-connected), alpha comparison, anti-aliased edges and the 'flood' worker handler.
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

const { floodFill } = load('imaging/floodFill.ts')
const { handlers } = load('shared/worker-ops/mask.ts')

const WAND = { tolerance: 32, contiguous: true, antiAlias: false, compareAlpha: false }

function image(width, height, colour) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) data.set(colour(x, y), (y * width + x) * 4)
  }
  return { width, height, data }
}

const selected = (m) => {
  const out = []
  for (let i = 0; i < m.data.length; i += 1) if (m.data[i] === 255) out.push(i)
  return out
}

test('tolerance selects exactly the pixels within max(|dr|, |dg|, |db|) of the seed', () => {
  // A horizontal ramp: value = x.
  const ramp = image(256, 3, (x) => [x, 128, 255 - x, 255])
  const m = floodFill(ramp, { x: 100.7, y: 1.2 }, { ...WAND, tolerance: 10 })
  for (let y = 0; y < 3; y += 1) {
    for (let x = 0; x < 256; x += 1) assert.equal(m.data[y * 256 + x], x >= 90 && x <= 110 ? 255 : 0, `pixel ${x},${y}`)
  }
  const exact = floodFill(ramp, { x: 100, y: 1 }, { ...WAND, tolerance: 0 })
  assert.equal(selected(exact).length, 3)
  const everything = floodFill(ramp, { x: 0, y: 0 }, { ...WAND, tolerance: 255 })
  assert.equal(selected(everything).length, 256 * 3)
})

test('contiguous fills one region; global selects every matching pixel', () => {
  const twoSquares = image(20, 10, (x, y) => ((x >= 2 && x < 6) || (x >= 12 && x < 16)) && y >= 2 && y < 6 ? [200, 30, 30, 255] : [20, 20, 20, 255])
  const contiguous = floodFill(twoSquares, { x: 3, y: 3 }, WAND)
  assert.equal(selected(contiguous).length, 16)
  assert.equal(contiguous.data[3 * 20 + 13], 0)
  const global = floodFill(twoSquares, { x: 3, y: 3 }, { ...WAND, contiguous: false })
  assert.equal(selected(global).length, 32)
  assert.equal(global.data[3 * 20 + 13], 255)
})

test('regions are 4-connected: a one-pixel diagonal line is a barrier', () => {
  const diagonal = image(12, 12, (x, y) => (x === y ? [0, 0, 0, 255] : [255, 255, 255, 255]))
  const m = floodFill(diagonal, { x: 8, y: 2 }, WAND)
  // Upper-right triangle only: (12 * 11) / 2 pixels with x > y.
  assert.equal(selected(m).length, 66)
  for (const index of selected(m)) assert.ok(index % 12 > ((index / 12) | 0))
})

test('fills large irregular regions completely (scanline stack correctness)', () => {
  // A spiral-ish maze of walls: every free pixel is reachable from the seed.
  const width = 64
  const height = 48
  const wall = (x, y) => (y % 8 === 4 && x % 31 !== 3) || (x % 16 === 8 && y % 23 === 7)
  const maze = image(width, height, (x, y) => (wall(x, y) ? [0, 0, 0, 255] : [240, 240, 240, 255]))
  const m = floodFill(maze, { x: 0, y: 0 }, WAND)
  let free = 0
  for (let y = 0; y < height; y += 1) for (let x = 0; x < width; x += 1) if (!wall(x, y)) free += 1
  assert.equal(selected(m).length, free)
})

test('alpha: transparent pixels match each other whatever colour they hide', () => {
  const layer = image(10, 4, (x) => (x < 5 ? [x * 50, 255 - x * 40, x * 20, 0] : [10, 200, 30, 255]))
  const withAlpha = floodFill(layer, { x: 0, y: 0 }, { ...WAND, tolerance: 0, compareAlpha: true })
  assert.equal(selected(withAlpha).length, 20)
  assert.ok(selected(withAlpha).every((index) => index % 10 < 5))
  // Without alpha comparison the hidden colours differ and only the seed column matches.
  const colourOnly = floodFill(layer, { x: 0, y: 0 }, { ...WAND, tolerance: 0, compareAlpha: false })
  assert.equal(selected(colourOnly).length, 4)
  // Alpha differences count: 50% transparent red is 128 levels away from opaque red.
  const halves = image(4, 1, (x) => [255, 0, 0, x < 2 ? 255 : 127])
  assert.equal(selected(floodFill(halves, { x: 0, y: 0 }, { ...WAND, compareAlpha: true })).length, 2)
  assert.equal(selected(floodFill(halves, { x: 0, y: 0 }, { ...WAND, compareAlpha: false })).length, 4)
})

test('anti-aliasing softens the ring just outside the region to clamp(1 - (diff - tolerance) / 8)', () => {
  // Background 100; a 3-pixel bar at 100 + 36 next to one at 100 + 50; tolerance 32.
  const strip = image(9, 1, (x) => [x < 4 ? 100 : x < 6 ? 136 : 150, 100, 100, 255])
  const soft = floodFill(strip, { x: 0, y: 0 }, { ...WAND, antiAlias: true })
  assert.deepEqual([...soft.data.slice(0, 4)], [255, 255, 255, 255])
  assert.equal(soft.data[4], Math.round((1 - (36 - 32) / 8) * 255), 'diff 36 -> 50% coverage')
  assert.equal(soft.data[5], 0, 'only direct neighbours of the region are softened')
  const hard = floodFill(strip, { x: 0, y: 0 }, WAND)
  assert.equal(hard.data[4], 0)
  // A hard edge far beyond the tolerance stays crisp, and partial values never reach 255.
  const edge = image(6, 6, (x) => (x < 3 ? [0, 0, 0, 255] : [255, 255, 255, 255]))
  const crisp = floodFill(edge, { x: 0, y: 0 }, { ...WAND, antiAlias: true })
  assert.ok(crisp.data.every((value) => value === 0 || value === 255))
  // Premultiplied difference 154 * 53 / 255 = 32.008 is just above the tolerance: coverage rounds to 255
  // but is held at 254 so the ring is never mistaken for matched pixels.
  const near = { width: 2, height: 1, data: Uint8ClampedArray.from([0, 0, 0, 53, 154, 0, 0, 53]) }
  assert.deepEqual([...floodFill(near, { x: 0, y: 0 }, { ...WAND, antiAlias: true, compareAlpha: true }).data], [255, 254])
})

test('seeds outside the image select nothing; bad buffers are rejected', () => {
  const small = image(4, 4, () => [1, 2, 3, 255])
  assert.equal(selected(floodFill(small, { x: -1, y: 0 }, WAND)).length, 0)
  assert.equal(selected(floodFill(small, { x: 4, y: 0 }, WAND)).length, 0)
  assert.equal(selected(floodFill(small, { x: NaN, y: 0 }, WAND)).length, 0)
  assert.equal(selected(floodFill(small, { x: 3.99, y: 3.99 }, WAND)).length, 16)
  assert.throws(() => floodFill({ width: 2, height: 2, data: new Uint8ClampedArray(3) }, { x: 0, y: 0 }, WAND), RangeError)
})

test('the flood worker handler matches the direct call and honours cancellation', async () => {
  const picture = image(30, 20, (x, y) => [(x * 7) % 256, (y * 11) % 256, 90, 255])
  const options = { tolerance: 40, contiguous: true, antiAlias: true, compareAlpha: true }
  const direct = floodFill(picture, { x: 5, y: 5 }, options)
  const viaWorker = await handlers.flood({ src: { width: 30, height: 20, data: Uint8ClampedArray.from(picture.data) }, seed: { x: 5, y: 5 }, options }, {})
  assert.deepEqual(viaWorker.data, direct.data)
  // A plain Uint8Array (as some clipboard paths produce) is accepted too.
  const bytes = await handlers.flood({ src: { width: 30, height: 20, data: Uint8Array.from(picture.data) }, seed: { x: 5, y: 5 }, options }, {})
  assert.deepEqual(bytes.data, direct.data)
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(handlers.flood({ src: picture, seed: { x: 0, y: 0 }, options }, { signal: controller.signal }), { name: 'AbortError' })
  await assert.rejects(handlers.flood({ src: { width: 3, height: 3, data: new Uint8ClampedArray(5) }, seed: { x: 0, y: 0 }, options }, {}), TypeError)
})
