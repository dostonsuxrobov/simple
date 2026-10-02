'use strict'
// WP3 sparse tiled storage (src/advanced/tiles.ts): addressing with negative tiles, exact read/write
// round trips across tile boundaries, compaction of empty tiles, masks with a default value, content
// bounds, per-tile and footprint versions, ownership rules of setTile/clone.
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

const tiles = load('advanced/tiles.ts')
const { TILE_SIZE } = load('advanced/types.ts')

/** Deterministic RGBA noise with a mix of transparent, partial and opaque pixels. */
function noise(width, height, seed = 1, alpha = null) {
  const data = new Uint8ClampedArray(width * height * 4)
  let s = seed >>> 0
  for (let i = 0; i < data.length; i += 4) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0
    data[i] = s & 255
    data[i + 1] = (s >>> 8) & 255
    data[i + 2] = (s >>> 16) & 255
    data[i + 3] = alpha === null ? Math.max(1, s >>> 24) : alpha
  }
  return { width, height, data }
}

function maskNoise(width, height, seed = 3) {
  const data = new Uint8Array(width * height)
  let s = seed >>> 0
  for (let i = 0; i < data.length; i += 1) {
    s = (Math.imul(s, 1664525) + 1013904223) >>> 0
    data[i] = s >>> 24
  }
  return { width, height, data }
}

/** Reference crop with transparent (or `fill`) outside the source. */
function crop(src, x, y, width, height, channels = 4, fill = 0) {
  const data = channels === 4 ? new Uint8ClampedArray(width * height * 4) : new Uint8Array(width * height).fill(fill)
  for (let row = 0; row < height; row += 1) {
    for (let col = 0; col < width; col += 1) {
      const sx = x + col
      const sy = y + row
      if (sx < 0 || sy < 0 || sx >= src.width || sy >= src.height) continue
      for (let c = 0; c < channels; c += 1) data[(row * width + col) * channels + c] = src.data[(sy * src.width + sx) * channels + c]
    }
  }
  return { width, height, data }
}

test('tile keys round-trip for negative and positive tile coordinates', () => {
  for (const [tx, ty] of [[0, 0], [-1, -1], [5, -7], [-0x8000, 0x7fff], [0x7fff, -0x8000], [123, 456]]) {
    const key = tiles.tileKey(tx, ty)
    assert.deepEqual(tiles.tileCoords(key), { tx, ty })
  }
  assert.equal(tiles.tileKey(0, 0), 0x8000 * 0x10000 + 0x8000)
  assert.equal(tiles.tileIndex(-1), -1)
  assert.equal(tiles.tileIndex(-256), -1)
  assert.equal(tiles.tileIndex(-257), -2)
  assert.equal(tiles.tileIndex(255), 0)
  assert.equal(tiles.tileIndex(256), 1)
  assert.deepEqual(tiles.tileRect(-1, 2), { x: -256, y: 512, width: 256, height: 256 })
  assert.deepEqual(tiles.tilesInRect({ x: -10, y: 250, width: 20, height: 10 }), [
    { tx: -1, ty: 0 }, { tx: 0, ty: 0 }, { tx: -1, ty: 1 }, { tx: 0, ty: 1 },
  ])
  assert.deepEqual(tiles.tilesInRect({ x: 0, y: 0, width: 0, height: 5 }), [])
  assert.deepEqual(tiles.tileAlignedRect({ x: -10, y: 5, width: 300, height: 10 }), { x: -256, y: 0, width: 768, height: 256 })
})

test('surfaces round-trip reads and writes across tile boundaries and negative coordinates', () => {
  const surface = tiles.createSurface()
  assert.equal(surface.channels, 4)
  assert.equal(surface.tileCount, 0)
  assert.equal(surface.byteSize, 0)
  assert.equal(surface.contentBounds(), null)
  const src = noise(600, 420, 7)
  const touched = surface.write(-300, -200, src)
  assert.deepEqual(touched, { x: -512, y: -256, width: 1024, height: 512 })
  // Exact read of the written region.
  assert.deepEqual(surface.read({ x: -300, y: -200, width: 600, height: 420 }).data, src.data)
  // Sub-rectangles spanning tiles, partially outside the written area.
  for (const [x, y, w, h] of [[-301, -201, 50, 40], [-10, -10, 300, 300], [250, 200, 80, 30], [-1000, -1000, 10, 10], [-44, 0, 1, 257]]) {
    const expected = crop(src, x + 300, y + 200, w, h)
    assert.deepEqual(surface.read({ x, y, width: w, height: h }).data, expected.data, `read ${x},${y} ${w}x${h}`)
  }
  // Reading into a provided target clears it first.
  const target = { width: 30, height: 20, data: new Uint8ClampedArray(30 * 20 * 4).fill(9) }
  surface.read({ x: 290, y: 210, width: 30, height: 20 }, target)
  assert.deepEqual(target.data, crop(src, 590, 410, 30, 20).data)
  assert.throws(() => surface.read({ x: 0, y: 0, width: 31, height: 20 }, target), RangeError)
  // An overlapping write replaces only its own pixels.
  const patch = noise(40, 300, 9)
  surface.write(240, -60, patch)
  const reference = crop(src, 0, 0, 600, 420)
  for (let row = 0; row < 300; row += 1) {
    for (let col = 0; col < 40; col += 1) {
      const x = 240 + col + 300
      const y = -60 + row + 200
      if (x >= 600 || y >= 420) continue
      reference.data.set(patch.data.subarray((row * 40 + col) * 4, (row * 40 + col) * 4 + 4), (y * 600 + x) * 4)
    }
  }
  assert.deepEqual(surface.read({ x: -300, y: -200, width: 600, height: 420 }).data, reference.data)
})

test('write honours a source rectangle and clips it to the source buffer', () => {
  const surface = tiles.createSurface()
  const src = noise(100, 80, 11)
  const touched = surface.write(10, 20, src, { x: -5, y: 70, width: 30, height: 30 })
  // The source rect is clipped to x 0..25, y 70..80; the destination shifts with the clip.
  assert.deepEqual(touched, { x: 0, y: 0, width: 256, height: 256 })
  assert.deepEqual(surface.contentBounds(), { x: 15, y: 20, width: 25, height: 10 })
  assert.deepEqual(surface.read({ x: 15, y: 20, width: 25, height: 10 }).data, crop(src, 0, 70, 25, 10).data)
  assert.deepEqual(surface.write(0, 0, src, { x: 200, y: 0, width: 10, height: 10 }), { x: 0, y: 0, width: 0, height: 0 })
  assert.throws(() => surface.write(0, 0, { width: 2, height: 2, data: new Uint8ClampedArray(15) }), RangeError)
  assert.throws(() => surface.write(0, 0, { width: 1, height: 1, data: new Float32Array(4) }), TypeError)
  // Plain byte arrays (decoders) are copied like clamped ones.
  surface.write(-7, -7, { width: 1, height: 1, data: new Uint8Array([1, 2, 3, 4]) })
  assert.deepEqual([...surface.read({ x: -7, y: -7, width: 1, height: 1 }).data], [1, 2, 3, 4])
  assert.ok(surface.read({ x: 0, y: 0, width: 1, height: 1 }).data instanceof Uint8ClampedArray)
})

test('empty tiles are never allocated and tiles that become empty are freed', () => {
  const surface = tiles.createSurface()
  surface.write(-100, -100, noise(300, 300, 1, 0))
  assert.equal(surface.tileCount, 0, 'writing transparent pixels into empty space allocates nothing')
  surface.write(-10, -10, noise(20, 20, 2))
  assert.equal(surface.tileCount, 4)
  assert.equal(surface.byteSize, 4 * TILE_SIZE * TILE_SIZE * 4)
  // Clearing part of the content keeps the tiles that still hold pixels.
  surface.write(-10, -10, noise(10, 20, 3, 0))
  assert.equal(surface.tileCount, 2)
  assert.deepEqual(surface.contentBounds(), { x: 0, y: -10, width: 10, height: 20 })
  surface.write(0, -10, noise(10, 20, 4, 0))
  assert.equal(surface.tileCount, 0)
  assert.equal(surface.contentBounds(), null)
  // setTile with a cleared tile keeps it until compaction.
  surface.write(5, 5, noise(3, 3, 5))
  surface.setTile(0, 0, new Uint8ClampedArray(TILE_SIZE * TILE_SIZE * 4))
  assert.equal(surface.tileCount, 1)
  assert.equal(surface.contentBounds(), null)
  assert.equal(tiles.compactGrid(surface), 1)
  assert.equal(surface.tileCount, 0)
  assert.equal(tiles.isEmptyTile(new Uint8ClampedArray(16), 4), true)
  assert.equal(tiles.isEmptyTile(new Uint8ClampedArray([0, 0, 0, 1]), 4), false)
  assert.equal(tiles.isEmptyTile(new Uint8ClampedArray([9, 9, 9, 0, 0, 0, 0, 0]), 4), true, 'colour under zero alpha is still empty')
})

test('content bounds are tight and follow every change', () => {
  const surface = tiles.createSurface()
  const pixel = { width: 1, height: 1, data: new Uint8ClampedArray([1, 2, 3, 4]) }
  surface.write(-257, 1000, pixel)
  assert.deepEqual(surface.contentBounds(), { x: -257, y: 1000, width: 1, height: 1 })
  surface.write(511, -3, pixel)
  assert.deepEqual(surface.contentBounds(), { x: -257, y: -3, width: 769, height: 1004 })
  const before = surface.contentBounds()
  assert.equal(surface.contentBounds(), before, 'cached per version')
  surface.setTile(tiles.tileIndex(511), tiles.tileIndex(-3), undefined)
  assert.deepEqual(surface.contentBounds(), { x: -257, y: 1000, width: 1, height: 1 })
})

test('versions are monotonic, per tile, and visible through pyramid footprints', () => {
  const a = tiles.createSurface()
  const b = tiles.createSurface()
  assert.equal(a.version, 0)
  a.write(0, 0, noise(10, 10, 1))
  const v1 = a.version
  assert.ok(v1 > 0)
  b.write(0, 0, noise(10, 10, 2))
  assert.ok(b.version > v1, 'versions are global across surfaces')
  a.write(300, 0, noise(10, 10, 3))
  assert.ok(a.version > b.version)
  assert.equal(a.tileVersion(0, 0), v1)
  assert.equal(a.tileVersion(1, 0), a.version)
  assert.equal(a.tileVersion(5, 5), 0)
  // Footprints: tile (1, 0) lies in level-1 tile (0, 0) and level-3 tile (0, 0).
  assert.equal(tiles.regionVersion(a, 1, 0, 0), a.version)
  assert.equal(tiles.regionVersion(a, 3, 0, 0), a.version)
  assert.equal(tiles.regionVersion(a, 1, 1, 0), 0)
  // Negative tiles map to negative footprints (floor division).
  a.write(-1, -1, noise(1, 1, 4))
  assert.equal(a.tileVersion(-1, -1), a.version)
  assert.equal(tiles.regionVersion(a, 1, -1, -1), a.version)
  assert.equal(tiles.regionVersion(a, 4, -1, -1), a.version)
  // A write of transparent pixels into nothing changes nothing.
  const quiet = a.version
  a.write(5000, 5000, noise(4, 4, 1, 0))
  assert.equal(a.version, quiet)
})

test('masks read their default value where nothing is stored and free tiles equal to it', () => {
  const reveal = tiles.createMaskSurface(255)
  assert.equal(reveal.channels, 1)
  assert.equal(reveal.defaultValue, 255)
  assert.ok(reveal.read({ x: -5, y: -5, width: 10, height: 10 }).data.every((v) => v === 255))
  const data = maskNoise(300, 300, 5)
  reveal.write(-100, 50, data)
  assert.deepEqual(reveal.read({ x: -100, y: 50, width: 300, height: 300 }).data, data.data)
  assert.deepEqual(reveal.read({ x: -150, y: 0, width: 100, height: 100 }).data, crop(data, -50, -50, 100, 100, 1, 255).data)
  // Writing the default everywhere frees every tile.
  reveal.write(-100, 50, { width: 300, height: 300, data: new Uint8Array(300 * 300).fill(255) })
  assert.equal(reveal.tileCount, 0)
  assert.equal(reveal.contentBounds(), null)
  const hide = tiles.createMaskSurface(0)
  hide.write(10, 10, { width: 2, height: 1, data: new Uint8Array([0, 200]) })
  assert.deepEqual(hide.contentBounds(), { x: 11, y: 10, width: 1, height: 1 })
  assert.equal(hide.byteSize, TILE_SIZE * TILE_SIZE)
  assert.throws(() => tiles.createMaskSurface(128), RangeError)
  assert.throws(() => hide.setTile(0, 0, new Uint8ClampedArray(TILE_SIZE * TILE_SIZE)), RangeError, 'mask tiles are Uint8Array')
  const fromBuffer = tiles.maskFromBuffer({ width: 3, height: 1, data: new Uint8Array([255, 7, 255]) }, 255, -2, 0)
  assert.equal(fromBuffer.tileCount, 1)
  assert.deepEqual(fromBuffer.read({ x: -3, y: 0, width: 5, height: 1 }).data, new Uint8Array([255, 255, 7, 255, 255]))
})

test('setTile validates, replaces and deletes whole tiles; forEachTile visits allocated tiles', () => {
  const surface = tiles.createSurface()
  assert.throws(() => surface.setTile(0, 0, new Uint8ClampedArray(10)), RangeError)
  assert.throws(() => surface.setTile(0, 0, new Uint8Array(TILE_SIZE * TILE_SIZE * 4)), RangeError)
  assert.throws(() => surface.setTile(0x8000, 0, new Uint8ClampedArray(TILE_SIZE * TILE_SIZE * 4)), RangeError)
  const tile = new Uint8ClampedArray(TILE_SIZE * TILE_SIZE * 4)
  tile.set([10, 20, 30, 40], (5 * TILE_SIZE + 6) * 4)
  surface.setTile(-2, 3, tile)
  assert.equal(surface.tile(-2, 3), tile, 'the surface takes ownership of the array')
  assert.deepEqual([...surface.read({ x: -512 + 6, y: 768 + 5, width: 1, height: 1 }).data], [10, 20, 30, 40])
  const seen = []
  surface.forEachTile((tx, ty, data) => seen.push([tx, ty, data === tile]))
  assert.deepEqual(seen, [[-2, 3, true]])
  surface.setTile(-2, 3, undefined)
  assert.equal(surface.tile(-2, 3), undefined)
  assert.equal(surface.tileCount, 0)
  assert.throws(() => surface.write(0x7fff * TILE_SIZE, 0, noise(TILE_SIZE + 1, 1, 1)), RangeError, 'too far from the canvas')
})

test('clone is a deep copy and factories place buffers at an offset', () => {
  const src = noise(70, 50, 13)
  const surface = tiles.surfaceFromBuffer(src, 250, -20)
  assert.deepEqual(surface.contentBounds(), { x: 250, y: -20, width: 70, height: 50 })
  const copy = surface.clone()
  assert.notEqual(copy.tile(1, 0), surface.tile(1, 0))
  assert.deepEqual(copy.read({ x: 250, y: -20, width: 70, height: 50 }).data, src.data)
  copy.write(250, -20, noise(70, 50, 1, 0))
  assert.equal(copy.contentBounds(), null)
  assert.deepEqual(surface.read({ x: 250, y: -20, width: 70, height: 50 }).data, src.data, 'the original is untouched')
  assert.ok(tiles.isTiledSurface(surface))
  assert.ok(!tiles.isTiledMask(surface))
  assert.ok(tiles.isTiledMask(tiles.createMaskSurface(0)))
  assert.ok(!tiles.isTiledSurface({ channels: 4 }))
})
