'use strict'
// WP3 pure compositor (src/advanced/composite.ts): exact single-layer reproduction, opacity, blend modes,
// layer masks, adjustment layers under coverage, hidden layers and clipping groups, Dissolve and dither
// stability across requests, previews, onlyLayerId / belowLayerId, flatten == compositeRect, sampling,
// and the 1 MP x 5 layer budget.
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

const composite = load('advanced/composite.ts')
const tiles = load('advanced/tiles.ts')
const doc = load('advanced/document.ts')
const blend = load('imaging/blend.ts')
const adjustments = load('imaging/adjustments.ts')
const selection = load('advanced/selection.ts')
const mask = load('imaging/mask.ts')

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

function solid(width, height, rgba) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i += 4) data.set(rgba, i)
  return { width, height, data }
}

function raster(name, buffer, init = {}) {
  return doc.createRasterLayer({ name, surface: tiles.surfaceFromBuffer(buffer), ...init })
}

function px(buffer, x, y) {
  const i = (y * buffer.width + x) * 4
  return [...buffer.data.subarray(i, i + 4)]
}

function near(actual, expected, tolerance = 1, message = '') {
  assert.equal(actual.length, expected.length)
  for (let i = 0; i < actual.length; i += 1) {
    assert.ok(Math.abs(actual[i] - expected[i]) <= tolerance, `${message} channel ${i}: ${actual} vs ${expected}`)
  }
}

function maskLayer(width, height, values, offsetX = 0, offsetY = 0, enabled = true) {
  const surface = tiles.maskFromBuffer({ width, height, data: Uint8Array.from(values) }, 255, 0, 0)
  return doc.createLayerMask({ surface, offsetX, offsetY, enabled })
}

const W = 64
const H = 48
const WHITE = [255, 255, 255, 255]

test('a single normal layer reproduces its pixels exactly; outside the document is transparent', () => {
  const src = noise(50, 40, 3)
  const layer = raster('Layer 1', src, { offsetX: 10, offsetY: -5 })
  const out = composite.compositeRect({ width: W, height: H, layers: [layer] }, { x: -4, y: -4, width: W + 8, height: H + 8 })
  for (let y = -4; y < H + 4; y += 1) {
    for (let x = -4; x < W + 4; x += 1) {
      const got = px(out, x + 4, y + 4)
      const lx = x - 10
      const ly = y + 5
      const inside = x >= 0 && y >= 0 && x < W && y < H && lx >= 0 && ly >= 0 && lx < 50 && ly < 40
      let expected = [0, 0, 0, 0]
      if (inside) {
        expected = px(src, lx, ly)
        if (expected[3] === 0) expected = [0, 0, 0, 0]
      }
      assert.deepEqual(got, expected, `pixel ${x},${y}`)
    }
  }
  // Every blend mode except Dissolve composites over transparency as a copy.
  for (const mode of ['multiply', 'screen', 'difference', 'luminosity']) {
    const moded = raster('m', src, { offsetX: 10, offsetY: -5, blendMode: mode })
    assert.deepEqual(composite.compositeRect({ width: W, height: H, layers: [moded] }, { x: 0, y: 0, width: W, height: H }).data,
      composite.compositeRect({ width: W, height: H, layers: [layer] }, { x: 0, y: 0, width: W, height: H }).data, mode)
  }
})

test('opacity over white gives the W3C values', () => {
  const background = raster('Background', solid(W, H, WHITE), { isBackground: true })
  const layer = raster('Layer 1', solid(W, H, [201, 101, 51, 255]), { opacity: 0.5 })
  const out = composite.compositeRect({ width: W, height: H, layers: [background, layer] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(out, 5, 5), [228, 178, 153, 255])
  // A half-transparent pixel at 50% opacity: as = 128 / 255 * 0.5.
  const half = raster('Layer 2', solid(W, H, [201, 101, 51, 128]), { opacity: 0.5 })
  const out2 = composite.compositeRect({ width: W, height: H, layers: [background, half] }, { x: 0, y: 0, width: W, height: H })
  const as = (128 / 255) * 0.5
  near(px(out2, 1, 1), [201 * as + 255 * (1 - as), 101 * as + 255 * (1 - as), 51 * as + 255 * (1 - as), 255])
  // Over transparency the alpha scales instead.
  const alone = composite.compositeRect({ width: W, height: H, layers: [layer] }, { x: 0, y: 0, width: 2, height: 2 })
  near(px(alone, 0, 0), [201, 101, 51, 127.5])
})

test('blend modes use the shared blend functions', () => {
  const backdrop = [90, 160, 220, 255]
  const source = [200, 40, 120, 255]
  const background = raster('Background', solid(4, 4, backdrop), { isBackground: true })
  for (const mode of ['multiply', 'screen', 'overlay', 'soft-light', 'color-dodge', 'hard-mix', 'hue', 'luminosity', 'darker-color']) {
    const layer = raster(mode, solid(4, 4, source), { blendMode: mode })
    const out = composite.compositeRect({ width: 4, height: 4, layers: [background, layer] }, { x: 0, y: 0, width: 4, height: 4 })
    const expected = blend.blendColor(mode, backdrop.slice(0, 3).map((v) => v / 255), source.slice(0, 3).map((v) => v / 255)).map((v) => v * 255)
    near(px(out, 2, 2), [...expected, 255], 1, mode)
  }
})

test('layer masks apply only under their coverage and can be disabled', () => {
  const background = raster('Background', solid(W, H, WHITE), { isBackground: true })
  // Mask 8 x 1 at document (4, 10): 0, 0, 128, 255, 255, ... ; elsewhere the default reveals.
  const values = [0, 0, 128, 255, 255, 255, 255, 0]
  const red = [255, 0, 0, 255]
  const layer = raster('Red', solid(W, H, red), { mask: maskLayer(8, 1, values, 4, 10) })
  const out = composite.compositeRect({ width: W, height: H, layers: [background, layer] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(out, 4, 10), WHITE)
  assert.deepEqual(px(out, 5, 10), WHITE)
  near(px(out, 6, 10), [255, 127, 127, 255])
  assert.deepEqual(px(out, 7, 10), red)
  assert.deepEqual(px(out, 11, 10), WHITE)
  assert.deepEqual(px(out, 30, 30), red, 'the default value reveals')
  const disabled = { ...layer, mask: { ...layer.mask, enabled: false } }
  const out2 = composite.compositeRect({ width: W, height: H, layers: [background, disabled] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(out2, 4, 10), red)
  // A hide-all mask with one revealed pixel.
  const hidden = tiles.createMaskSurface(0)
  hidden.write(20, 20, { width: 1, height: 1, data: new Uint8Array([255]) })
  const masked = raster('Masked', solid(W, H, red), { mask: doc.createLayerMask({ surface: hidden }) })
  const out3 = composite.compositeRect({ width: W, height: H, layers: [background, masked] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(out3, 20, 20), red)
  assert.deepEqual(px(out3, 21, 20), WHITE)
})

test('adjustment layers change only what is below them, only under their mask, mixed by opacity', () => {
  const below = raster('Background', noise(W, H, 9, 255), { isBackground: true })
  const above = raster('Top', solid(8, 8, [10, 20, 30, 255]), { offsetX: 40, offsetY: 0 })
  // Invert the top half of the document only.
  const values = new Array(W * H).fill(0)
  for (let i = 0; i < W * (H / 2); i += 1) values[i] = 255
  const invert = doc.createAdjustmentLayer('Invert 1', { type: 'invert' }, maskLayer(W, H, values))
  const layers = [below, invert, above]
  const out = composite.compositeRect({ width: W, height: H, layers }, { x: 0, y: 0, width: W, height: H })
  const base = composite.compositeRect({ width: W, height: H, layers: [below] }, { x: 0, y: 0, width: W, height: H })
  for (const [x, y] of [[0, 0], [17, 5], [63, 23]]) {
    const b = px(base, x, y)
    assert.deepEqual(px(out, x, y), [255 - b[0], 255 - b[1], 255 - b[2], 255], `inverted ${x},${y}`)
  }
  assert.deepEqual(px(out, 5, 30), px(base, 5, 30), 'outside the mask')
  assert.deepEqual(px(out, 41, 1), [10, 20, 30, 255], 'the layer above is not adjusted')
  // Opacity 50% mixes halfway.
  const halfInvert = { ...invert, opacity: 0.5 }
  const mixed = composite.compositeRect({ width: W, height: H, layers: [below, halfInvert] }, { x: 0, y: 0, width: W, height: H })
  const b = px(base, 3, 3)
  near(px(mixed, 3, 3), [b[0] + (255 - 2 * b[0]) * 0.5, b[1] + (255 - 2 * b[1]) * 0.5, b[2] + (255 - 2 * b[2]) * 0.5, 255])
  // Transparent pixels stay transparent.
  const sparse = raster('Dot', solid(1, 1, [100, 100, 100, 255]), { offsetX: 2, offsetY: 2 })
  const out2 = composite.compositeRect({ width: W, height: H, layers: [sparse, doc.createAdjustmentLayer('Invert 2', { type: 'invert' })] },
    { x: 0, y: 0, width: 5, height: 5 })
  assert.deepEqual(px(out2, 2, 2), [155, 155, 155, 255])
  assert.equal(px(out2, 0, 0)[3], 0)
})

test('hidden layers and hidden clipping bases are skipped; clipped layers stay inside the base alpha', () => {
  const background = raster('Background', solid(W, H, WHITE), { isBackground: true })
  const squareBuffer = solid(10, 10, [0, 0, 255, 255])
  for (let i = 0; i < 10; i += 1) squareBuffer.data[(5 * 10 + i) * 4 + 3] = 128
  const base = raster('Base', squareBuffer, { offsetX: 20, offsetY: 20 })
  const clipped = raster('Clipped', solid(W, H, [255, 0, 0, 255]), { clipped: true })
  const layers = [background, base, clipped]
  const out = composite.compositeRect({ width: W, height: H, layers }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(out, 22, 22), [255, 0, 0, 255], 'clipped colour inside the base')
  assert.deepEqual(px(out, 5, 5), WHITE, 'nothing outside the base')
  near(px(out, 22, 25), [255, 127, 127, 255], 1, 'half-transparent base row keeps its alpha')
  // The base's opacity applies to the whole group.
  const faded = composite.compositeRect({ width: W, height: H, layers: [background, { ...base, opacity: 0.5 }, clipped] }, { x: 0, y: 0, width: W, height: H })
  near(px(faded, 22, 22), [255, 127, 127, 255])
  // A multiply clipped layer blends with the base, not with the white below.
  const multiply = raster('Multiply', solid(W, H, [128, 128, 128, 255]), { clipped: true, blendMode: 'multiply' })
  const out2 = composite.compositeRect({ width: W, height: H, layers: [background, base, multiply] }, { x: 0, y: 0, width: W, height: H })
  near(px(out2, 22, 22), [0, 0, 128, 255])
  // Hidden layers and hidden bases.
  const hiddenClipped = composite.compositeRect({ width: W, height: H, layers: [background, base, { ...clipped, visible: false }] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(hiddenClipped, 22, 22), [0, 0, 255, 255])
  const hiddenBase = composite.compositeRect({ width: W, height: H, layers: [background, { ...base, visible: false }, clipped] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(hiddenBase, 22, 22), WHITE, 'a hidden base hides its clipped layers')
  assert.deepEqual(px(hiddenBase, 5, 5), WHITE)
  // A clipped adjustment adjusts only the base.
  const clippedInvert = { ...doc.createAdjustmentLayer('Invert', { type: 'invert' }), clipped: true }
  const out3 = composite.compositeRect({ width: W, height: H, layers: [background, base, clippedInvert] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(out3, 22, 22), [255, 255, 0, 255])
  assert.deepEqual(px(out3, 5, 5), WHITE, 'the white below the group is not inverted')
})

test('Dissolve and dithered gradient maps are stable across different requests', () => {
  const background = raster('Background', noise(300, 300, 4, 255), { isBackground: true })
  const dissolve = raster('Dissolve', noise(300, 300, 5), { blendMode: 'dissolve', opacity: 0.7 })
  const map = doc.createAdjustmentLayer('Gradient Map 1', {
    type: 'gradient-map',
    stops: [{ position: 0, color: { r: 20, g: 0, b: 90 } }, { position: 1, color: { r: 250, g: 200, b: 40 } }],
    reverse: false,
    dither: true,
  })
  const d = { width: 300, height: 300, layers: [background, dissolve, map] }
  const full = composite.flattenDocument(d)
  assert.deepEqual(composite.compositeRect(d, { x: 0, y: 0, width: 300, height: 300 }).data, full.data)
  for (const [x, y, w, h] of [[13, 7, 100, 61], [255, 250, 3, 40], [0, 299, 300, 1], [128, 128, 172, 150]]) {
    const part = composite.compositeRect(d, { x, y, width: w, height: h })
    for (let row = 0; row < h; row += 1) {
      const from = ((y + row) * 300 + x) * 4
      assert.deepEqual(part.data.subarray(row * w * 4, (row + 1) * w * 4), full.data.subarray(from, from + w * 4), `rect ${x},${y} row ${row}`)
    }
  }
  // Past the document edge everything is transparent.
  const beyond = composite.compositeRect(d, { x: 290, y: 0, width: 20, height: 2 })
  assert.deepEqual(beyond.data.subarray(0, 40), full.data.subarray(290 * 4, 300 * 4))
  assert.ok(beyond.data.subarray(40, 80).every((v) => v === 0))
  // Dissolve pixels are either fully kept or dropped.
  const alone = composite.compositeRect({ width: 300, height: 300, layers: [dissolve] }, { x: 0, y: 0, width: 300, height: 300 })
  let kept = 0
  for (let p = 3; p < alone.data.length; p += 4) {
    assert.ok(alone.data[p] === 0 || alone.data[p] === 255)
    if (alone.data[p]) kept += 1
  }
  assert.ok(kept > 300 * 300 * 0.2 && kept < 300 * 300 * 0.5, `kept ${kept}`)
})

test('previews change opacity, blend, visibility, adjustments and pixels without touching the layers', () => {
  const background = raster('Background', solid(W, H, WHITE), { isBackground: true })
  const layer = raster('Layer 1', solid(W, H, [201, 101, 51, 255]))
  const d = { width: W, height: H, layers: [background, layer] }
  const rect = { x: 0, y: 0, width: W, height: H }
  const faded = composite.compositeRect(d, rect, { preview: { kind: 'layer-props', layerId: layer.id, opacity: 0.5 } })
  assert.deepEqual(px(faded, 3, 3), [228, 178, 153, 255])
  const hidden = composite.compositeRect(d, rect, { preview: { kind: 'layer-props', layerId: layer.id, visible: false } })
  assert.deepEqual(px(hidden, 3, 3), WHITE)
  const multiply = composite.compositeRect(d, rect, { preview: { kind: 'layer-props', layerId: layer.id, blendMode: 'multiply' } })
  assert.deepEqual(px(multiply, 3, 3), [201, 101, 51, 255])
  // Destructive-dialog preview on a pixel layer, inside a selection only.
  const shape = mask.createMaskBuffer(W, H)
  mask.rasterizeRect(shape, { x: 0, y: 0, width: 10, height: H }, false)
  const sel = selection.selectionFromMask(shape, 1)
  const inverted = composite.compositeRect(d, rect, { preview: { kind: 'adjustment', layerId: layer.id, spec: { type: 'invert' }, selection: sel } })
  assert.deepEqual(px(inverted, 3, 3), [54, 154, 204, 255])
  assert.deepEqual(px(inverted, 30, 3), [201, 101, 51, 255])
  // Live spec of an adjustment layer.
  const levels = doc.createAdjustmentLayer('Levels 1', adjustments.defaultAdjustment('levels'))
  const d2 = { width: W, height: H, layers: [background, layer, levels] }
  assert.deepEqual(px(composite.compositeRect(d2, rect), 3, 3), [201, 101, 51, 255], 'default Levels is identity')
  const live = composite.compositeRect(d2, rect, { preview: { kind: 'adjustment', layerId: levels.id, spec: { type: 'invert' }, selection: null } })
  assert.deepEqual(px(live, 3, 3), [54, 154, 204, 255])
  // Filter preview pixels at a level (document level coordinates).
  const pixels = solid(2, 2, [1, 2, 3, 255])
  const previewed = composite.compositeRect(d, rect, { preview: { kind: 'layer-pixels', layerId: layer.id, level: 0, rect: { x: 4, y: 4, width: 2, height: 2 }, pixels } })
  assert.deepEqual(px(previewed, 5, 5), [1, 2, 3, 255])
  assert.deepEqual(px(previewed, 6, 6), [201, 101, 51, 255])
  const otherLevel = composite.compositeRect(d, rect, { preview: { kind: 'layer-pixels', layerId: layer.id, level: 1, rect: { x: 4, y: 4, width: 2, height: 2 }, pixels } })
  assert.deepEqual(px(otherLevel, 5, 5), [201, 101, 51, 255], 'a preview for another level is ignored')
  assert.deepEqual(px(composite.compositeRect(d, rect), 3, 3), [201, 101, 51, 255], 'layers are untouched')
})

test('onlyLayerId, belowLayerId and includeBelowLayer select layers', () => {
  const background = raster('Background', solid(W, H, WHITE), { isBackground: true })
  const middle = raster('Middle', solid(W, H, [10, 200, 10, 255]), { opacity: 0.5, mask: maskLayer(1, 1, [0]) })
  const top = raster('Top', solid(4, 4, [0, 0, 255, 255]))
  const d = { width: W, height: H, layers: [background, middle, top] }
  const rect = { x: 0, y: 0, width: 4, height: 4 }
  assert.deepEqual(px(composite.compositeRect(d, rect, { onlyLayerId: middle.id }), 0, 0), [10, 200, 10, 255], 'raw pixels: no mask, no opacity')
  near(px(composite.compositeRect(d, rect, { belowLayerId: top.id }), 1, 1), [132.5, 227.5, 132.5, 255])
  assert.deepEqual(px(composite.compositeRect(d, rect, { belowLayerId: middle.id }), 1, 1), WHITE)
  assert.deepEqual(px(composite.compositeRect(d, rect, { belowLayerId: top.id, includeBelowLayer: true }), 1, 1), [0, 0, 255, 255])
  assert.equal(px(composite.compositeRect(d, rect, { onlyLayerId: 'missing' }), 0, 0)[3], 0)
  assert.deepEqual(px(composite.compositeRect(d, rect, { onlyLayerId: middle.id }), 0, 0), [10, 200, 10, 255])
  // The mask hides the middle layer only at (0, 0).
  assert.deepEqual(px(composite.compositeRect(d, { x: 0, y: 0, width: 4, height: 4 }, { belowLayerId: top.id }), 0, 0), WHITE)
})

test('flattenDocument equals compositeRect over the full document and supports progress and abort', () => {
  const layers = [
    raster('Background', noise(530, 270, 1, 255), { isBackground: true }),
    raster('A', noise(400, 300, 2), { offsetX: -50, offsetY: 20, blendMode: 'overlay', opacity: 0.8 }),
    doc.createAdjustmentLayer('Curves', { type: 'posterize', levels: 5 }, maskLayer(3, 1, [0, 255, 0], 100, 100)),
    raster('B', noise(100, 100, 3), { offsetX: 300, offsetY: 150, blendMode: 'hue' }),
  ]
  const d = { width: 530, height: 270, layers }
  const progress = []
  const flat = composite.flattenDocument(d, { onProgress: (value) => progress.push(value) })
  assert.deepEqual(flat.data, composite.compositeRect(d, { x: 0, y: 0, width: 530, height: 270 }).data)
  assert.equal(progress[progress.length - 1], 1)
  const controller = new AbortController()
  controller.abort()
  assert.throws(() => composite.flattenDocument(d, { signal: controller.signal }), (error) => error.name === 'AbortError')
  // compositeInto writes a rectangle at an offset of a larger target and leaves the rest alone.
  const target = { width: 20, height: 20, data: new Uint8ClampedArray(20 * 20 * 4).fill(7) }
  composite.compositeInto(target, 5, 6, d, { x: 100, y: 100, width: 10, height: 10 })
  assert.deepEqual(px(target, 0, 0), [7, 7, 7, 7])
  assert.deepEqual(px(target, 5, 6), px(flat, 100, 100))
  assert.deepEqual(px(target, 14, 15), px(flat, 109, 109))
})

test('text and shape layers composite their raster caches', () => {
  const shape = doc.createShapeLayer('Rectangle 1', {
    kind: 'rectangle', x1: 10, y1: 10, x2: 30, y2: 20, fill: { r: 0, g: 128, b: 255, a: 255 }, stroke: null,
    strokeWidth: 0, cornerRadius: 0, arrowHeads: 'none', transform: [1, 0, 0, 1, 0, 0],
  })
  const out = composite.compositeRect({ width: W, height: H, layers: [shape] }, { x: 0, y: 0, width: W, height: H })
  assert.deepEqual(px(out, 20, 15), [0, 128, 255, 255])
  assert.equal(px(out, 5, 5)[3], 0)
  const raster = doc.rasterCacheFrom(solid(2, 2, [9, 8, 7, 255]), 3, 4, 'key')
  const text = { ...shape, kind: 'text', text: { text: 'x', style: {}, boxWidth: null, transform: [1, 0, 0, 1, 0, 0] }, raster }
  delete text.shape
  assert.deepEqual(px(composite.compositeRect({ width: W, height: H, layers: [text] }, { x: 0, y: 0, width: W, height: H }), 4, 5), [9, 8, 7, 255])
})

test('sampleDocument averages with alpha weighting from the chosen source', () => {
  const background = raster('Background', solid(W, H, [200, 0, 0, 255]), { isBackground: true })
  const top = raster('Top', solid(1, 1, [0, 0, 200, 255]), { offsetX: 10, offsetY: 10 })
  const state = { width: W, height: H, layers: [background, top], activeLayerId: top.id }
  assert.deepEqual(composite.sampleDocument(state, 10.7, 10.2, 1, 'all'), { r: 0, g: 0, b: 200, a: 255 })
  assert.deepEqual(composite.sampleDocument(state, 10, 10, 3, 'all'), { r: 178, g: 0, b: 22, a: 255 })
  // Current layer only: one opaque pixel among transparent ones does not darken.
  assert.deepEqual(composite.sampleDocument(state, 10, 10, 3, 'current'), { r: 0, g: 0, b: 200, a: 28 })
  assert.deepEqual(composite.sampleDocument({ ...state, activeLayerId: background.id }, 10, 10, 1, 'current-below'), { r: 200, g: 0, b: 0, a: 255 })
  assert.deepEqual(composite.sampleDocument(state, -5, 3, 1, 'all'), { r: 0, g: 0, b: 0, a: 0 })
  assert.deepEqual(composite.sampleDocument(state, 0, 0, 3, 'all'), { r: 200, g: 0, b: 0, a: 255 }, 'clipped to the canvas')
})

test('1 MP x 5 layers composites in under 150 ms', () => {
  const size = 1000
  const layers = [
    raster('Background', noise(size, size, 1, 255), { isBackground: true }),
    raster('A', noise(size, size, 2), { opacity: 0.7 }),
    raster('B', noise(size, size, 3), { blendMode: 'multiply' }),
    raster('C', noise(size, size, 4), { blendMode: 'screen', opacity: 0.5 }),
    raster('D', noise(size, size, 5), { blendMode: 'overlay' }),
  ]
  const d = { width: size, height: size, layers }
  // One warm-up pass (lazy blend tables, JIT), then the best of four: other processes share the CPU.
  composite.compositeRect(d, { x: 0, y: 0, width: size, height: size })
  let best = Infinity
  for (let run = 0; run < 4; run += 1) {
    const started = performance.now()
    composite.compositeRect(d, { x: 0, y: 0, width: size, height: size })
    best = Math.min(best, performance.now() - started)
  }
  assert.ok(best < 150, `best of 4: ${best.toFixed(1)} ms`)
})

// ---------------------------------------------------------------------------------------------
// Display compositor (src/advanced/compositor.ts) against a minimal fake canvas: the level window must
// hold exactly the composited tiles once settle() resolves, follow edits, previews, zoom levels and
// re-centring; flatten() / renderToCanvas() give one consistent state even when edits land mid-pass.
// The real-GPU pixel match runs in Electron (scratch harness), not here.
// ---------------------------------------------------------------------------------------------

const viewport = load('advanced/viewport.ts')

class FakeContext {
  constructor(canvas) {
    this.canvas = canvas
    this.imageSmoothingEnabled = true
    this.imageSmoothingQuality = 'low'
    this.globalAlpha = 1
    this.globalCompositeOperation = 'source-over'
    this.fillStyle = '#000'
    this.draws = []
  }

  pixels() {
    const c = this.canvas
    if (!c.data || c.data.length !== c.width * c.height * 4) c.data = new Uint8ClampedArray(c.width * c.height * 4)
    return c.data
  }

  putImageData(image, x, y, dirtyX = 0, dirtyY = 0, dirtyWidth = image.width, dirtyHeight = image.height) {
    const out = this.pixels()
    if (arguments.length > 3) this.partialPuts = (this.partialPuts ?? 0) + 1
    for (let row = Math.max(0, dirtyY); row < Math.min(image.height, dirtyY + dirtyHeight); row += 1) {
      const ty = y + row
      if (ty < 0 || ty >= this.canvas.height) continue
      for (let col = Math.max(0, dirtyX); col < Math.min(image.width, dirtyX + dirtyWidth); col += 1) {
        const tx = x + col
        if (tx < 0 || tx >= this.canvas.width) continue
        const from = (row * image.width + col) * 4
        out.set(image.data.subarray(from, from + 4), (ty * this.canvas.width + tx) * 4)
      }
    }
  }

  clearRect(x, y, width, height) {
    const out = this.pixels()
    for (let ty = Math.max(0, y); ty < Math.min(this.canvas.height, y + height); ty += 1) {
      out.fill(0, (ty * this.canvas.width + Math.max(0, x)) * 4, (ty * this.canvas.width + Math.min(this.canvas.width, x + width)) * 4)
    }
  }

  drawImage(source, ...args) {
    this.draws.push({ source, args, smoothing: this.imageSmoothingEnabled })
    if (args.length === 2 && source.data) {
      // 1:1 copy (window re-centring).
      const [dx, dy] = args
      this.putImageData({ width: source.width, height: source.height, data: source.data }, dx, dy)
    }
  }

  createPattern() { return { pattern: true } }
  setTransform() {}
  save() {}
  restore() {}
  beginPath() {}
  rect() {}
  clip() {}
  translate() {}
  fillRect() {}
}

class FakeCanvas {
  constructor(width = 300, height = 150) {
    this._width = width
    this._height = height
    this.data = null
    this.context = null
    FakeCanvas.created.push(this)
  }

  get width() { return this._width }
  set width(value) { this._width = value; this.data = null }
  get height() { return this._height }
  set height(value) { this._height = value; this.data = null }

  getContext() {
    if (!this.context) this.context = new FakeContext(this)
    return this.context
  }
}
FakeCanvas.created = []

class FakeImageData {
  constructor(data, width, height) {
    assert.equal(data.length, width * height * 4, 'ImageData size')
    this.data = data
    this.width = width
    this.height = height
  }
}

async function withFakeDom(run) {
  const saved = { OffscreenCanvas: globalThis.OffscreenCanvas, ImageData: globalThis.ImageData, document: globalThis.document }
  FakeCanvas.created = []
  globalThis.OffscreenCanvas = FakeCanvas
  globalThis.ImageData = FakeImageData
  globalThis.document = { visibilityState: 'visible', createElement: () => new FakeCanvas() }
  try {
    await run()
  } finally {
    for (const [key, value] of Object.entries(saved)) {
      if (value === undefined) delete globalThis[key]
      else globalThis[key] = value
    }
  }
}

const compositorModule = load('advanced/compositor.ts')

function makeDisplayStore() {
  let counter = 0
  const host = { nextRevision: () => (counter += 1), setRevision: () => {} }
  const layers = [
    raster('Background', noise(700, 520, 41, 255), { isBackground: true }),
    raster('A', noise(400, 300, 42), { offsetX: 100, offsetY: 60, blendMode: 'multiply' }),
    doc.createAdjustmentLayer('Posterize 1', { type: 'posterize', levels: 6 }),
    raster('B', noise(200, 200, 43), { offsetX: -50, offsetY: 350, opacity: 0.6 }),
  ]
  const store = doc.createDocumentStore({ width: 700, height: 520, ppi: 72, layers, host, baseLabel: 'Open' })
  return { store, layers }
}

/** The live window canvas (the newest canvas of at least one tile that holds pixels). */
function liveWindow() {
  const live = FakeCanvas.created.filter((canvas) => canvas.width >= 256 && canvas.height >= 256 && canvas.data)
  return live[live.length - 1]
}

function assertWindowMatches(store, view, size, level, message) {
  const state = store.getState()
  const win = liveWindow()
  assert.ok(win, `${message}: a window exists`)
  const counts = viewport.levelTileCounts(state, level)
  const visible = viewport.visibleTileRange(view, size, state, level, 0)
  const range = {
    tx0: Math.max(0, visible.tx0 - 1),
    ty0: Math.max(0, visible.ty0 - 1),
    tx1: Math.min(counts.columns, visible.tx1 + 1),
    ty1: Math.min(counts.rows, visible.ty1 + 1),
  }
  assert.equal(win.width, (range.tx1 - range.tx0) * 256, `${message}: window width`)
  for (let ty = visible.ty0; ty < visible.ty1; ty += 1) {
    for (let tx = visible.tx0; tx < visible.tx1; tx += 1) {
      const rect = viewport.levelTileRect(state, level, tx, ty)
      const expected = composite.compositeRect(state, rect, { level })
      for (let row = 0; row < rect.height; row += 1) {
        const from = (((ty - range.ty0) * 256 + row) * win.width + (tx - range.tx0) * 256) * 4
        const got = Buffer.from(win.data.buffer, win.data.byteOffset + from, rect.width * 4)
        const want = Buffer.from(expected.data.buffer, expected.data.byteOffset + row * rect.width * 4, rect.width * 4)
        if (Buffer.compare(got, want) !== 0) assert.fail(`${message}: tile ${tx},${ty} row ${row} differs`)
      }
    }
  }
}

test('the display compositor settles to exactly the composited tiles and follows edits', async () => {
  await withFakeDom(async () => {
    const { store, layers } = makeDisplayStore()
    const compositor = compositorModule.createCompositor(store)
    const screen = new FakeCanvas()
    const size = { width: 500, height: 300, dpr: 1 }
    let view = { zoom: 1, offsetX: -20, offsetY: -10 }
    assert.equal(await compositor.settle(), undefined, 'nothing attached: settles at once')
    compositor.attach(screen)
    compositor.setView(view, size)
    assert.deepEqual([screen.width, screen.height], [500, 300], 'the screen canvas gets device pixels')
    await compositor.settle()
    assert.equal(compositor.level, 0)
    assert.equal(compositor.stats().pendingVisible, 0)
    assertWindowMatches(store, view, size, 0, '100%')
    const draws = screen.getContext().draws
    assert.ok(draws.length > 0, 'the window was drawn to the screen')
    assert.equal(draws[draws.length - 1].smoothing, false, '1:1 blits without smoothing')
    // An edit invalidates only what it touches.
    const composited = compositor.stats().tilesComposited
    store.transact('Paint', 'brush', (tx) => tx.editPixels(layers[1].id, 'pixels').writePixels(10, 10, noise(20, 20, 7, 255)))
    await compositor.settle()
    assertWindowMatches(store, view, size, 0, 'after an edit')
    assert.ok(compositor.stats().tilesComposited - composited <= 4, 'only the touched tiles were recomposited')
    // An adjustment change affects everything.
    store.transact('Posterize', 'adjustment', (tx) => tx.updateLayer(layers[2].id, { adjustment: { type: 'posterize', levels: 3 } }))
    await compositor.settle()
    assertWindowMatches(store, view, size, 0, 'after an adjustment change')
    // A small pan whose tiles are all current (the ring) still waits for a frame that draws the new view.
    const drawnBefore = draws.length
    view = { zoom: 1, offsetX: -30, offsetY: -12 }
    compositor.setView(view, size)
    await compositor.settle()
    assert.ok(draws.length > drawnBefore, 'settle() waits for the new view to be drawn')
    const last = draws[draws.length - 1].args
    assert.equal(last[4] - last[0], -30, 'drawn at the new offset (destination x - source x = origin)')
    // Pan far enough to re-centre the window.
    view = { zoom: 1, offsetX: -300, offsetY: -250 }
    compositor.setView(view, size)
    await compositor.settle()
    assertWindowMatches(store, view, size, 0, 'after panning')
    // Zoom out to a proxy level.
    view = { zoom: 0.4, offsetX: 0, offsetY: 0 }
    compositor.setView(view, size)
    await compositor.settle()
    assert.equal(compositor.level, 1)
    assertWindowMatches(store, view, size, 1, 'at 40%')
    // Undo through the store invalidates too.
    store.history.undo()
    await compositor.settle()
    assertWindowMatches(store, view, size, 1, 'after undo')
    // Interactive mode uses one level coarser, and refines when it ends.
    compositor.setInteractive(true)
    assert.equal(compositor.level, 2)
    await compositor.settle()
    assertWindowMatches(store, view, size, 2, 'interactive')
    compositor.setInteractive(false)
    await compositor.settle()
    assertWindowMatches(store, view, size, 1, 'refined')
    // Sampling matches the pure sampler.
    assert.deepEqual(compositor.sample(30, 40, 3, 'all'), composite.sampleDocument(store.getState(), 30, 40, 3, 'all'))
    compositor.dispose()
    assert.equal(await compositor.settle(), undefined)
  })
})

test('compositor previews show on screen but never in flatten(); flatten and renderToCanvas are exact', async () => {
  await withFakeDom(async () => {
    const { store, layers } = makeDisplayStore()
    const compositor = compositorModule.createCompositor(store)
    const screen = new FakeCanvas()
    const size = { width: 800, height: 600, dpr: 1 }
    const view = { zoom: 1, offsetX: 0, offsetY: 0 }
    compositor.attach(screen)
    compositor.setView(view, size)
    await compositor.settle()
    const preview = { kind: 'layer-props', layerId: layers[1].id, opacity: 0.25, blendMode: 'screen' }
    compositor.setPreview(preview)
    await compositor.settle()
    const expected = composite.compositeRect(store.getState(), { x: 0, y: 0, width: 256, height: 256 }, { preview })
    assert.deepEqual(liveWindow().data.subarray(0, 256 * 4), expected.data.subarray(0, 256 * 4), 'the preview is on screen')
    const reference = composite.flattenDocument(store.getState())
    const progress = []
    const flat = await compositor.flatten({ onProgress: (value) => progress.push(value) })
    assert.deepEqual(flat.data, reference.data, 'flatten ignores the preview and reuses no previewed tile')
    assert.equal(progress[progress.length - 1], 1)
    compositor.setPreview(null)
    await compositor.settle()
    assert.deepEqual(liveWindow().data.subarray(0, 256 * 4), reference.data.subarray(0, 256 * 4))
    const canvas = await compositor.renderToCanvas()
    assert.deepEqual([canvas.width, canvas.height], [700, 520])
    assert.deepEqual(canvas.data, reference.data)
    const controller = new AbortController()
    controller.abort()
    await assert.rejects(compositor.flatten({ signal: controller.signal }), (error) => error.name === 'AbortError')
    compositor.dispose()
    await assert.rejects(compositor.flatten(), /closed/)
  })
})

test('flatten() restarts when the document changes during the pass', async () => {
  await withFakeDom(async () => {
    let counter = 0
    const host = { nextRevision: () => (counter += 1), setRevision: () => {} }
    const layers = [
      raster('Background', noise(1600, 1600, 51, 255), { isBackground: true }),
      raster('A', noise(1600, 1600, 52), { blendMode: 'overlay' }),
      raster('B', noise(1600, 1600, 53), { blendMode: 'luminosity', opacity: 0.5 }),
    ]
    const store = doc.createDocumentStore({ width: 1600, height: 1600, ppi: 72, layers, host, baseLabel: 'Open' })
    const compositor = compositorModule.createCompositor(store)
    let progress = 0
    const pending = compositor.flatten({ onProgress: (value) => { progress = value } })
    // Edit while flatten is yielding between tile rows (rows below the edit are already done).
    await new Promise((resolve) => setTimeout(resolve, 0))
    const progressAtEdit = progress
    store.transact('Paint', 'brush', (tx) => tx.editPixels(layers[0].id, 'pixels').writePixels(0, 0, noise(1600, 100, 54, 255)))
    const flat = await pending
    assert.ok(progressAtEdit > 0 && progressAtEdit < 1, `the edit landed mid-pass (progress ${progressAtEdit})`)
    assert.deepEqual(flat.data, composite.flattenDocument(store.getState()).data, 'one consistent, final state')
    compositor.dispose()
  })
})

// ---------------------------------------------------------------------------------------------
// Shared vector rasterizer (src/shared/vector.ts), used by Advanced text/shape layers and Simple markup:
// text layout, bounds, cache keys and the canvas-free shape rasterizer (Node has no canvas).
// ---------------------------------------------------------------------------------------------

const vector = load('shared/vector.ts')

const STYLE = { ...vector.DEFAULT_TEXT_STYLE, fontSize: 20, lineHeight: 1.5, letterSpacing: 0 }
/** Every character is 10 px wide; ascent 16, descent 4. */
const measure10 = (text) => ({ width: Array.from(text).length * 10, ascent: 16, descent: 4 })

function shapeSpec(extra) {
  return { kind: 'rectangle', x1: 0, y1: 0, x2: 10, y2: 10, fill: { r: 255, g: 0, b: 0, a: 255 }, stroke: null, strokeWidth: 0, cornerRadius: 0, arrowHeads: 'none', transform: [1, 0, 0, 1, 0, 0], ...extra }
}

function coverage(result) {
  let sum = 0
  for (let p = 3; p < result.pixels.data.length; p += 4) sum += result.pixels.data[p]
  return sum / 255
}

test('text layout: lines, word wrap, long words, alignment and baselines', () => {
  const point = vector.layoutText({ text: 'Hello\nab', style: STYLE, boxWidth: null }, measure10)
  assert.deepEqual(point.lines.map((line) => [line.text, line.x, line.width]), [['Hello', 0, 50], ['ab', 0, 20]])
  assert.equal(point.width, 50)
  assert.equal(point.lineHeight, 30)
  assert.equal(point.height, 60)
  // Half leading (30 - 20) / 2 = 5 above the ascent.
  assert.deepEqual(point.lines.map((line) => [line.top, line.baseline]), [[0, 21], [30, 51]])
  const wrapped = vector.layoutText({ text: 'one two three four', style: STYLE, boxWidth: 75 }, measure10)
  assert.deepEqual(wrapped.lines.map((line) => line.text), ['one two', 'three', 'four'])
  assert.equal(wrapped.width, 75)
  const broken = vector.layoutText({ text: 'abcdefghij', style: STYLE, boxWidth: 35 }, measure10)
  assert.deepEqual(broken.lines.map((line) => line.text), ['abc', 'def', 'ghi', 'j'])
  const centred = vector.layoutText({ text: 'ab\nabcd', style: { ...STYLE, align: 'center' }, boxWidth: null }, measure10)
  assert.deepEqual(centred.lines.map((line) => line.x), [10, 0])
  const right = vector.layoutText({ text: 'ab', style: { ...STYLE, align: 'right' }, boxWidth: 100 }, measure10)
  assert.equal(right.lines[0].x, 80)
  assert.deepEqual(vector.layoutText({ text: '', style: STYLE, boxWidth: null }, measure10).lines.map((line) => line.text), [''])
  assert.equal(vector.cssFont({ ...STYLE, fontFamily: 'Segoe "UI"; x', fontWeight: 700, italic: true }), 'italic 700 20px "Segoe UI x", sans-serif')
  // textBounds maps the layout box through the transform.
  const spec = { text: 'Hi', style: STYLE, boxWidth: null, transform: [2, 0, 0, 2, 100, 50] }
  const layout = vector.layoutText(spec)
  assert.deepEqual(vector.textBounds(spec), { x: 100, y: 50, width: layout.width * 2, height: layout.height * 2 })
})

test('shape bounds include strokes and arrow heads and follow the transform', () => {
  assert.deepEqual(vector.shapeBounds(shapeSpec({ stroke: { r: 0, g: 0, b: 0, a: 255 }, strokeWidth: 4 })), { x: -2, y: -2, width: 14, height: 14 })
  assert.deepEqual(vector.shapeBounds(shapeSpec({ x1: 10, x2: 0, y1: 10, y2: 0 })), { x: 0, y: 0, width: 10, height: 10 }, 'endpoints in any order')
  assert.equal(vector.arrowHeadLength(1), 10)
  assert.equal(vector.arrowHeadLength(5), 20)
  const arrow = vector.shapeBounds(shapeSpec({ kind: 'arrow', x1: 0, y1: 0, x2: 100, y2: 0, stroke: { r: 0, g: 0, b: 0, a: 255 }, strokeWidth: 5, arrowHeads: 'end' }))
  assert.deepEqual(arrow, { x: -20, y: -20, width: 140, height: 40 })
  const rotated = vector.shapeBounds(shapeSpec({ transform: [0, 1, -1, 0, 50, 0] }))
  assert.deepEqual(rotated, { x: 40, y: 0, width: 10, height: 10 })
})

test('the canvas-free rasterizer covers exactly the shape area', () => {
  const rect = vector.rasterizeShapePure(shapeSpec({ x1: 2.5, y1: 3, x2: 22.5, y2: 13 }))
  assert.ok(Math.abs(coverage(rect) - 200) < 1, `rectangle area ${coverage(rect)}`)
  const inside = rect.pixels.data
  const at = (x, y) => [...inside.subarray(((y - rect.offsetY) * rect.pixels.width + (x - rect.offsetX)) * 4, ((y - rect.offsetY) * rect.pixels.width + (x - rect.offsetX)) * 4 + 4)]
  assert.deepEqual(at(10, 8), [255, 0, 0, 255])
  assert.equal(at(2, 8)[3], 128, 'half-covered edge pixel')
  const ellipse = vector.rasterizeShapePure(shapeSpec({ kind: 'ellipse', x1: 0, y1: 0, x2: 60, y2: 40 }))
  const ellipseArea = Math.PI * 30 * 20
  assert.ok(Math.abs(coverage(ellipse) - ellipseArea) / ellipseArea < 0.01, `ellipse area ${coverage(ellipse)} vs ${ellipseArea}`)
  // A stroke-only rectangle is a ring of the stroke width centred on the outline.
  const ring = vector.rasterizeShapePure(shapeSpec({ x1: 0, y1: 0, x2: 40, y2: 40, fill: null, stroke: { r: 0, g: 0, b: 255, a: 255 }, strokeWidth: 4 }))
  assert.ok(Math.abs(coverage(ring) - (44 * 44 - 36 * 36)) < 2, `ring area ${coverage(ring)}`)
  // Lines are capsules (round caps): length * width + pi * r^2.
  const line = vector.rasterizeShapePure(shapeSpec({ kind: 'line', x1: 10, y1: 10, x2: 60, y2: 10, fill: null, stroke: { r: 0, g: 0, b: 0, a: 255 }, strokeWidth: 6 }))
  const capsule = 50 * 6 + Math.PI * 9
  assert.ok(Math.abs(coverage(line) - capsule) / capsule < 0.02, `line area ${coverage(line)} vs ${capsule}`)
  // Translucent fill keeps its alpha; arrows add their heads.
  const translucent = vector.rasterizeShapePure(shapeSpec({ fill: { r: 0, g: 0, b: 0, a: 128 } }))
  assert.equal(Math.max(...translucent.pixels.data.filter((v, i) => i % 4 === 3)), 128)
  const shaft = coverage(vector.rasterizeShapePure(shapeSpec({ kind: 'line', x1: 0, y1: 0, x2: 100, y2: 0, stroke: { r: 0, g: 0, b: 0, a: 255 }, strokeWidth: 4 })))
  const headed = coverage(vector.rasterizeShapePure(shapeSpec({ kind: 'arrow', x1: 0, y1: 0, x2: 100, y2: 0, stroke: { r: 0, g: 0, b: 0, a: 255 }, strokeWidth: 4, arrowHeads: 'both' })))
  // Two heads of length max(10, 4 * 4) = 16 and half-width 0.45 * 16, plus the shaft between them.
  const arrowArea = 2 * (0.5 * 16 * 2 * 0.45 * 16) + (100 - 2 * 16) * 4
  assert.ok(Math.abs(headed - arrowArea) / arrowArea < 0.03, `arrow area ${headed} vs ${arrowArea}`)
  assert.ok(headed > shaft, 'heads add area')
  // Without a canvas, rasterizeShape uses the same pure rasterizer.
  assert.deepEqual(vector.rasterizeShape(shapeSpec({ kind: 'ellipse', x2: 30, y2: 17 })).pixels.data, vector.rasterizeShapePure(shapeSpec({ kind: 'ellipse', x2: 30, y2: 17 })).pixels.data)
  assert.throws(() => vector.rasterizeShapePure(shapeSpec({ x2: 30000, y2: 10 })), RangeError)
})

test('spec keys ignore whole-pixel moves only', () => {
  const base = shapeSpec({ transform: [1, 0, 0, 1, 10.25, 5] })
  const moved = shapeSpec({ transform: [1, 0, 0, 1, 13.25, -2] })
  assert.equal(vector.specKey(base), vector.specKey(moved))
  assert.deepEqual(vector.wholePixelShift(base, moved), { x: 3, y: -7 })
  assert.notEqual(vector.specKey(base), vector.specKey(shapeSpec({ transform: [1, 0, 0, 1, 10.5, 5] })))
  assert.equal(vector.wholePixelShift(base, shapeSpec({ transform: [1, 0, 0, 1, 10.5, 5] })), null)
  assert.notEqual(vector.specKey(base), vector.specKey({ ...base, fill: { r: 0, g: 0, b: 0, a: 255 } }))
  const text = { text: 'A', style: STYLE, boxWidth: null, transform: [1, 0, 0, 1, 0, 0] }
  assert.ok(vector.isTextSpec(text) && !vector.isTextSpec(base))
  assert.notEqual(vector.specKey(text), vector.specKey({ ...text, text: 'B' }))
  assert.equal(vector.specKey(text), vector.specKey({ ...text, transform: [1, 0, 0, 1, 5, 9] }))
  // Text needs a canvas; empty text needs nothing.
  assert.throws(() => vector.rasterizeTextSync(text), /canvas/)
  assert.equal(vector.rasterizeTextSync({ ...text, text: '  ' }).pixels.width, 0)
})

test('small edits recomposite only the changed part of a tile and stay exact', async () => {
  await withFakeDom(async () => {
    const { store, layers } = makeDisplayStore()
    const compositor = compositorModule.createCompositor(store)
    const screen = new FakeCanvas()
    const size = { width: 700, height: 520, dpr: 1 }
    let view = { zoom: 1, offsetX: 0, offsetY: 0 }
    compositor.attach(screen)
    compositor.setView(view, size)
    await compositor.settle()
    const full = compositor.stats().tilesComposited
    // A stroke of small dabs on layer B (some start left of the canvas).
    const stroke = store.beginStroke(layers[3].id, 'pixels', 'Brush Tool', 'brush')
    for (let i = 0; i < 12; i += 1) {
      stroke.editor.writePixels(10 + i * 17, 20 + i * 5, noise(9, 9, 60 + i, 255))
      await compositor.settle()
      assertWindowMatches(store, view, size, 0, `after dab ${i}`)
    }
    stroke.commit()
    const stats = compositor.stats()
    assert.equal(stats.tilesComposited, full, 'no whole tile was recomposited for the dabs')
    // The first two dabs lie entirely left of the canvas; the other ten each refresh part of one tile.
    assert.equal(stats.partialRefreshes, 10)
    assert.equal(liveWindow().context.partialPuts, 10, 'only the changed part was uploaded')
    // Undo restores exactly, and a proxy level refreshes partially too.
    store.history.undo()
    await compositor.settle()
    assertWindowMatches(store, view, size, 0, 'after undo')
    view = { zoom: 0.5, offsetX: 0, offsetY: 0 }
    compositor.setView(view, size)
    await compositor.settle()
    const proxyFull = compositor.stats().tilesComposited
    store.transact('Dab', 'brush', (tx) => tx.editPixels(layers[1].id, 'pixels').writePixels(40, 40, noise(6, 6, 77, 255)))
    await compositor.settle()
    assertWindowMatches(store, view, size, 1, 'proxy after a dab')
    assert.equal(compositor.stats().tilesComposited, proxyFull)
    // A large change still recomposites whole tiles.
    store.transact('Fill', 'fill', (tx) => tx.editPixels(layers[0].id, 'pixels').writePixels(0, 0, noise(700, 520, 78, 255)))
    await compositor.settle()
    assertWindowMatches(store, view, size, 1, 'proxy after a fill')
    assert.ok(compositor.stats().tilesComposited > proxyFull)
    // flatten() only reuses fully current tiles.
    assert.deepEqual((await compositor.flatten()).data, composite.flattenDocument(store.getState()).data)
    compositor.dispose()
    // A dab into a tile that composited to nothing (no pixel buffer yet).
    let counter = 0
    const sparseLayer = raster('Dot', noise(4, 4, 90, 255), { offsetX: 600, offsetY: 400 })
    const sparse = doc.createDocumentStore({ width: 700, height: 520, ppi: 72, layers: [sparseLayer], host: { nextRevision: () => (counter += 1), setRevision: () => {} }, baseLabel: 'Open' })
    const second = compositorModule.createCompositor(sparse)
    second.attach(new FakeCanvas())
    view = { zoom: 1, offsetX: 0, offsetY: 0 }
    second.setView(view, size)
    await second.settle()
    sparse.transact('Dab', 'brush', (tx) => tx.editPixels(sparseLayer.id, 'pixels').writePixels(-590, -390, noise(5, 5, 91, 255)))
    await second.settle()
    assertWindowMatches(sparse, view, size, 0, 'dab into an empty tile')
    assert.ok(second.stats().partialRefreshes >= 1)
    second.dispose()
  })
})
