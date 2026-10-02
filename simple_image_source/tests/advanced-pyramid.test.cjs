'use strict'
// WP3 mip pyramid (src/advanced/pyramid.ts) and proxy compositing: alpha-weighted 2x2 levels that never
// darken transparent edges, exact agreement with a reference downsample (including negative tiles),
// incremental rebuilds after edits, masks, the cache budget, and composite level 1 of a normal-mode
// stack against a downsample of level 0 (design acceptance: within +-2).
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

const pyramid = load('advanced/pyramid.ts')
const tiles = load('advanced/tiles.ts')
const composite = load('advanced/composite.ts')
const documentModule = load('advanced/document.ts')
const memory = load('advanced/memory.ts')

function noise(width, height, seed = 1, alpha = null) {
  const data = new Uint8ClampedArray(width * height * 4)
  let s = seed >>> 0
  for (let i = 0; i < data.length; i += 4) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0
    data[i] = s & 255
    data[i + 1] = (s >>> 8) & 255
    data[i + 2] = (s >>> 16) & 255
    data[i + 3] = alpha === null ? s >>> 24 : alpha
  }
  return { width, height, data }
}

function maxDifference(a, b) {
  assert.equal(a.length, b.length)
  let max = 0
  for (let i = 0; i < a.length; i += 1) max = Math.max(max, Math.abs(a[i] - b[i]))
  return max
}

/**
 * Max difference over the proxy pixels that are fully inside the document. With an odd width or height
 * the last proxy column/row covers half a pixel outside the canvas: the reference counts it transparent,
 * the proxy keeps whatever a layer has there (that half pixel is clipped away when it is drawn).
 */
function interiorDifference(proxy, reference, docWidth, docHeight, scale) {
  const fullColumns = Math.floor(docWidth / scale)
  const fullRows = Math.floor(docHeight / scale)
  let max = 0
  for (let y = 0; y < fullRows; y += 1) {
    for (let x = 0; x < fullColumns; x += 1) {
      const i = (y * reference.width + x) * 4
      for (let c = 0; c < 4; c += 1) max = Math.max(max, Math.abs(proxy.data[i + c] - reference.data[i + c]))
    }
  }
  return max
}

test('level geometry', () => {
  assert.equal(pyramid.maxPyramidLevel(256, 256), 0)
  assert.equal(pyramid.maxPyramidLevel(257, 10), 1)
  assert.equal(pyramid.maxPyramidLevel(6000, 4000), 5)
  assert.equal(pyramid.maxPyramidLevel(20000, 2500), 7)
  assert.deepEqual(pyramid.levelSize(1001, 7, 1), { width: 501, height: 4 })
  assert.deepEqual(pyramid.levelSize(1001, 7, 3), { width: 126, height: 1 })
  assert.equal(pyramid.levelOffset(-1, 1), -1)
  assert.equal(pyramid.levelOffset(-2, 1), -1)
  assert.equal(pyramid.levelOffset(-3, 1), -2)
  assert.equal(pyramid.levelOffset(5, 2), 1)
})

test('the 2x2 reduction is alpha-weighted: transparent neighbours never darken an edge', () => {
  const edge = { width: 2, height: 2, data: new Uint8ClampedArray([
    200, 100, 50, 255, 0, 0, 0, 0,
    0, 0, 0, 0, 0, 0, 0, 0,
  ]) }
  const out = pyramid.downsampleBuffer2x2(edge)
  assert.deepEqual([...out.data], [200, 100, 50, 64])
  const mixed = { width: 2, height: 1, data: new Uint8ClampedArray([255, 0, 0, 255, 0, 0, 255, 85]) }
  // Weighted by alpha 255 : 85 -> red 191, blue 64; alpha (255 + 85 + 0 + 0) / 4 = 85.
  assert.deepEqual([...pyramid.downsampleBuffer2x2(mixed).data], [191, 0, 64, 85])
})

test('surface level tiles match a reference downsample exactly, also across negative tiles', () => {
  const src = noise(700, 530, 21)
  const surface = tiles.surfaceFromBuffer(src, -300, -260)
  // Level 0 reads the surface itself.
  assert.deepEqual(pyramid.readSurfaceLevel(surface, 0, { x: -300, y: -260, width: 700, height: 530 }).data, src.data)
  // Build the level-0 reference covering the aligned area that holds the content.
  const aligned = surface.read({ x: -512, y: -512, width: 1024, height: 1024 })
  const level1 = pyramid.downsampleBuffer2x2(aligned)
  const level2 = pyramid.downsampleBuffer2x2(level1)
  assert.deepEqual(pyramid.readSurfaceLevel(surface, 1, { x: -256, y: -256, width: 512, height: 512 }).data, level1.data)
  assert.deepEqual(pyramid.readSurfaceLevel(surface, 2, { x: -128, y: -128, width: 256, height: 256 }).data, level2.data)
  // A sub-rectangle and an area outside everything.
  const part = pyramid.readSurfaceLevel(surface, 1, { x: -100, y: 20, width: 90, height: 33 })
  for (let row = 0; row < 33; row += 1) {
    const from = ((20 + 256 + row) * 512 + (-100 + 256)) * 4
    assert.deepEqual(part.data.subarray(row * 90 * 4, (row + 1) * 90 * 4), level1.data.subarray(from, from + 90 * 4))
  }
  assert.ok(pyramid.readSurfaceLevel(surface, 3, { x: 500, y: 500, width: 20, height: 20 }).data.every((v) => v === 0))
  assert.equal(pyramid.surfaceLevelTile(surface, 4, 10, 10), null)
})

test('an edit rebuilds exactly the proxy tiles above it', () => {
  const src = noise(600, 300, 5, 255)
  const surface = tiles.surfaceFromBuffer(src)
  const before = pyramid.readSurfaceLevel(surface, 1, { x: 0, y: 0, width: 300, height: 150 })
  const untouchedTile = pyramid.surfaceLevelTile(surface, 1, 1, 0)
  assert.ok(untouchedTile)
  // Paint one pixel in level-0 tile (0, 0): level-1 tile (0, 0) changes, level-1 tile (1, 0) does not.
  surface.write(10, 10, { width: 2, height: 2, data: new Uint8ClampedArray(16).fill(255) })
  const after = pyramid.readSurfaceLevel(surface, 1, { x: 0, y: 0, width: 300, height: 150 })
  const changed = []
  for (let i = 0; i < after.data.length; i += 4) if (after.data[i] !== before.data[i] || after.data[i + 3] !== before.data[i + 3]) changed.push(i / 4)
  assert.deepEqual(changed, [5 * 300 + 5], 'only the proxy pixel over the edit changes')
  assert.deepEqual([...after.data.subarray(changed[0] * 4, changed[0] * 4 + 4)], [255, 255, 255, 255])
  assert.equal(pyramid.surfaceLevelTile(surface, 1, 1, 0), untouchedTile, 'the neighbouring proxy tile is reused')
  // Deleting the tile makes its proxy transparent.
  surface.setTile(0, 0, undefined)
  assert.ok(pyramid.readSurfaceLevel(surface, 1, { x: 0, y: 0, width: 128, height: 128 }).data.every((v) => v === 0))
})

test('mask levels average plainly and keep the default value where nothing is stored', () => {
  const mask = tiles.createMaskSurface(255)
  mask.write(0, 0, { width: 2, height: 2, data: new Uint8Array([0, 0, 0, 100]) })
  const level1 = pyramid.readMaskLevel(mask, 1, { x: -1, y: -1, width: 3, height: 3 })
  // Pixel (0, 0) averages 0, 0, 0, 100 -> 25; everything else is the default 255.
  assert.deepEqual([...level1.data], [255, 255, 255, 255, 25, 255, 255, 255, 255])
  assert.equal(pyramid.maskLevelTile(tiles.createMaskSurface(0), 2, 0, 0), null)
  const hidden = tiles.createMaskSurface(0)
  hidden.write(300, 0, { width: 2, height: 2, data: new Uint8Array([255, 255, 255, 255]) })
  assert.deepEqual([...pyramid.readMaskLevel(hidden, 1, { x: 149, y: 0, width: 3, height: 1 }).data], [0, 255, 0])
})

test('the proxy cache honours its budget and reports its bytes', () => {
  const surface = tiles.surfaceFromBuffer(noise(1024, 512, 9))
  pyramid.clearPyramidCache()
  const expected = pyramid.readSurfaceLevel(surface, 2, { x: 0, y: 0, width: 256, height: 128 })
  assert.ok(pyramid.pyramidCacheBytes() > 0)
  assert.ok(memory.cacheBytes() >= pyramid.pyramidCacheBytes(), 'registered with memory accounting')
  pyramid.setPyramidCacheBudget(0)
  assert.equal(pyramid.pyramidCacheBytes(), 0)
  assert.deepEqual(pyramid.readSurfaceLevel(surface, 2, { x: 0, y: 0, width: 256, height: 128 }).data, expected.data, 'rebuilt on demand')
  pyramid.setPyramidCacheBudget(pyramid.DEFAULT_PYRAMID_BUDGET)
  memory.trimCaches()
  assert.equal(pyramid.pyramidCacheBytes(), 0)
})

/** Photo-like content: smooth colour and (optionally) smooth alpha. */
function smooth(width, height, seed, alpha = null) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const i = (y * width + x) * 4
      data[i] = 128 + 90 * Math.sin(x / 37 + seed)
      data[i + 1] = 128 + 90 * Math.cos(y / 41 + seed * 2)
      data[i + 2] = 128 + 60 * Math.sin((x + y) / 53 + seed * 3)
      data[i + 3] = alpha === null ? 140 + 100 * Math.sin(x / 45 + y / 61 + seed) : alpha
    }
  }
  return { width, height, data }
}

test('composite level 1 of a normal-mode stack matches a 2x2 downsample of level 0 within 2', () => {
  // Each layer is downsampled on its own, so a proxy equals the downsampled composite when blending is
  // linear over every 2x2 block: opaque pixels with layer opacity (even noise), or photo-like content
  // (smooth colour and alpha). Where a semi-transparent layer edge splits a block over a busy backdrop,
  // proxies may differ more (display only; flatten and export always use level 0). Proxies place layers
  // at floor(offset / 2), so offsets here are even (an odd offset may shift a layer by one proxy pixel).
  const W = 517
  const H = 389
  const layers = [
    documentModule.createRasterLayer({ name: 'Background', isBackground: true, surface: tiles.surfaceFromBuffer(smooth(W, H, 5, 255)) }),
    documentModule.createRasterLayer({ name: 'A', surface: tiles.surfaceFromBuffer(smooth(301, 199, 1), 40, 60), opacity: 0.6 }),
    documentModule.createRasterLayer({ name: 'C', surface: tiles.surfaceFromBuffer(smooth(W, H, 2)), offsetX: -100, offsetY: 32 }),
    // Busy opaque content on top, aligned to 4 px so it stays block-aligned at levels 1 and 2.
    documentModule.createRasterLayer({ name: 'B', surface: tiles.surfaceFromBuffer(noise(252, 300, 33, 255)), offsetX: 200, offsetY: 48, opacity: 0.8 }),
  ]
  const doc = { width: W, height: H, layers }
  const full = composite.flattenDocument(doc)
  const reference = pyramid.downsampleBuffer2x2(full)
  const proxy = composite.compositeRect(doc, { x: 0, y: 0, width: reference.width, height: reference.height }, { level: 1 })
  const level1 = interiorDifference(proxy, reference, W, H, 2)
  assert.ok(level1 <= 2, `level 1 max difference ${level1}`)
  // Level 2 against a 4x4 downsample of level 0.
  const reference2 = pyramid.downsampleBuffer2x2(reference)
  const proxy2 = composite.compositeRect(doc, { x: 0, y: 0, width: reference2.width, height: reference2.height }, { level: 2 })
  const level2 = interiorDifference(proxy2, reference2, W, H, 4)
  assert.ok(level2 <= 2, `level 2 max difference ${level2}`)
  // An even-sized document has no partial edge pixels at level 1: the whole proxy matches.
  const even = { width: 516, height: 388, layers }
  const evenReference = pyramid.downsampleBuffer2x2(composite.flattenDocument(even))
  const evenProxy = composite.compositeRect(even, { x: 0, y: 0, width: 258, height: 194 }, { level: 1 })
  assert.ok(maxDifference(evenProxy.data, evenReference.data) <= 2, `even document max difference ${maxDifference(evenProxy.data, evenReference.data)}`)
})
