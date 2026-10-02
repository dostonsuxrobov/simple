'use strict'
// WP4 gradients (src/imaging/gradient.ts): stop colours at stop positions, midpoint bias, opacity stops,
// the five gradient geometries, reverse, dither bounds and tile-independent rendering.
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

const { gradientColorAt, gradientParameter, renderGradient } = load('imaging/gradient.ts')

const STOPS = [
  { position: 0, color: { r: 0, g: 0, b: 0 } },
  { position: 0.3, color: { r: 255, g: 0, b: 128 } },
  { position: 0.7, color: { r: 17, g: 201, b: 99 } },
  { position: 1, color: { r: 10, g: 250, b: 30 } },
]

function spec(patch = {}) {
  return { kind: 'linear', from: { x: 0, y: 0 }, to: { x: 100, y: 0 }, stops: STOPS, opacityStops: [], reverse: false, dither: false, ...patch }
}

const pixel = (buffer, x, y) => Array.from(buffer.data.slice((y * buffer.width + x) * 4, (y * buffer.width + x) * 4 + 4))

test('linear gradients hit their stop colours at stop positions', () => {
  // origin.x = -0.5 puts pixel centre i at document x = i, so t = i / 100 exactly.
  const g = renderGradient(101, 2, { x: -0.5, y: 0 }, spec())
  assert.deepEqual(pixel(g, 0, 0), [0, 0, 0, 255])
  assert.deepEqual(pixel(g, 30, 1), [255, 0, 128, 255])
  assert.deepEqual(pixel(g, 70, 0), [17, 201, 99, 255])
  assert.deepEqual(pixel(g, 100, 1), [10, 250, 30, 255])
  // Linear interpolation between stops (midpoint 0.5).
  assert.deepEqual(pixel(g, 15, 0), [128, 0, 64, 255])
  // Beyond the end points the end colours extend.
  const wide = renderGradient(140, 1, { x: -20.5, y: 0 }, spec())
  assert.deepEqual(pixel(wide, 0, 0), [0, 0, 0, 255])
  assert.deepEqual(pixel(wide, 139, 0), [10, 250, 30, 255])
})

test('midpoints bias the interpolation and opacity stops drive alpha', () => {
  const biased = { stops: [{ position: 0, color: { r: 0, g: 0, b: 0 }, midpoint: 0.25 }, { position: 1, color: { r: 200, g: 100, b: 0 } }], opacityStops: [] }
  const atMidpoint = gradientColorAt(biased, 0.25)
  assert.ok(Math.abs(atMidpoint.r - 100) < 1e-9 && Math.abs(atMidpoint.g - 50) < 1e-9, 'the midpoint is the 50% mix')
  assert.ok(gradientColorAt(biased, 0.5).r > 100)
  const faded = spec({ stops: [{ position: 0, color: { r: 9, g: 9, b: 9 } }], opacityStops: [{ position: 0, opacity: 1 }, { position: 1, opacity: 0 }] })
  const g = renderGradient(101, 1, { x: -0.5, y: 0 }, faded)
  assert.equal(pixel(g, 0, 0)[3], 255)
  assert.equal(pixel(g, 50, 0)[3], 128)
  assert.equal(pixel(g, 100, 0)[3], 0)
  assert.deepEqual(pixel(g, 100, 0).slice(0, 3), [9, 9, 9], 'straight alpha keeps the colour')
  // Unsorted stops are sorted; missing stops fall back to opaque black.
  const unsorted = renderGradient(3, 1, { x: -0.5, y: 0 }, spec({ to: { x: 2, y: 0 }, stops: [{ position: 1, color: { r: 200, g: 0, b: 0 } }, { position: 0, color: { r: 0, g: 0, b: 200 } }] }))
  assert.deepEqual(pixel(unsorted, 0, 0), [0, 0, 200, 255])
  assert.deepEqual(pixel(unsorted, 2, 0), [200, 0, 0, 255])
  assert.deepEqual(pixel(renderGradient(1, 1, { x: 0, y: 0 }, spec({ stops: [] })), 0, 0), [0, 0, 0, 255])
})

test('geometry of the five gradient kinds, and reverse', () => {
  const from = { x: 50, y: 50 }
  const to = { x: 90, y: 50 }
  const t = (kind, x, y, reverse = false) => gradientParameter({ kind, from, to, reverse }, x, y)
  assert.equal(t('linear', 70, 10), 0.5)
  assert.equal(t('linear', 20, 50), 0, 'clamped before the start')
  assert.equal(t('reflected', 30, 50), 0.5, 'mirrored about the start')
  assert.equal(t('radial', 50, 80), 0.75)
  assert.equal(t('diamond', 70, 50), 0.5)
  assert.equal(t('diamond', 60, 60), 0.5, '|a| + |b| in the drag frame')
  assert.equal(t('diamond', 50, 90), 1)
  assert.equal(t('angle', 90, 50), 0)
  assert.ok(Math.abs(t('angle', 50, 10) - 0.25) < 1e-12, 'counter-clockwise on screen (up is a quarter turn)')
  assert.ok(Math.abs(t('angle', 10, 50) - 0.5) < 1e-12)
  assert.ok(Math.abs(t('angle', 50, 90) - 0.75) < 1e-12)
  assert.equal(t('linear', 70, 10, true), 0.5)
  assert.equal(t('linear', 80, 10, true), 0.25)
  // Rendered radial gradient matches the parameter.
  const radial = renderGradient(101, 101, { x: 0, y: 0 }, spec({ kind: 'radial', from: { x: 50.5, y: 50.5 }, to: { x: 100.5, y: 50.5 }, stops: [{ position: 0, color: { r: 0, g: 0, b: 0 } }, { position: 1, color: { r: 200, g: 200, b: 200 } }] }))
  assert.deepEqual(pixel(radial, 50, 50), [0, 0, 0, 255])
  assert.deepEqual(pixel(radial, 75, 50), [100, 100, 100, 255])
  assert.deepEqual(pixel(radial, 50, 25), [100, 100, 100, 255])
  // A click without a drag paints the end colour everywhere (no NaN).
  const click = renderGradient(4, 4, { x: 0, y: 0 }, spec({ to: { x: 0, y: 0 } }))
  assert.deepEqual(pixel(click, 2, 2), [10, 250, 30, 255])
})

test('dither stays within +-1 of the exact value, is seeded and tile independent', () => {
  const base = spec({ stops: [{ position: 0, color: { r: 20, g: 40, b: 60 } }, { position: 1, color: { r: 27, g: 41, b: 70 } }], opacityStops: [{ position: 0, opacity: 1 }, { position: 1, opacity: 0.9 }] })
  const plain = renderGradient(101, 64, { x: -0.5, y: 0 }, base)
  const dithered = renderGradient(101, 64, { x: -0.5, y: 0 }, { ...base, dither: true }, 7)
  let changed = 0
  for (let i = 0; i < plain.data.length; i += 1) {
    const difference = Math.abs(dithered.data[i] - plain.data[i])
    assert.ok(difference <= 1, `channel ${i} differs by ${difference}`)
    if (difference) changed += 1
  }
  assert.ok(changed > plain.data.length * 0.1, 'dither actually breaks up the bands')
  // Mean is preserved: dithering does not shift the gradient.
  const mean = (buffer, channel) => {
    let total = 0
    for (let i = channel; i < buffer.data.length; i += 4) total += buffer.data[i]
    return total / (buffer.data.length / 4)
  }
  const exact = (channel) => {
    let total = 0
    for (let x = 0; x < 101; x += 1) total += [20, 40, 60][channel] + ([27, 41, 70][channel] - [20, 40, 60][channel]) * (x / 100)
    return total / 101
  }
  for (let channel = 0; channel < 3; channel += 1) assert.ok(Math.abs(mean(dithered, channel) - exact(channel)) < 0.1)
  // Same seed, same noise; another seed, other noise.
  assert.deepEqual(renderGradient(101, 64, { x: -0.5, y: 0 }, { ...base, dither: true }, 7).data, dithered.data)
  assert.notDeepEqual(renderGradient(101, 64, { x: -0.5, y: 0 }, { ...base, dither: true }, 8).data, dithered.data)
  // Rendering in two tiles equals rendering the whole area (noise keyed on document pixels).
  const left = renderGradient(50, 64, { x: -0.5, y: 0 }, { ...base, dither: true }, 7)
  const right = renderGradient(51, 64, { x: 49.5, y: 0 }, { ...base, dither: true }, 7)
  for (let y = 0; y < 64; y += 1) {
    for (let x = 0; x < 101; x += 1) {
      const tile = x < 50 ? pixel(left, x, y) : pixel(right, x - 50, y)
      assert.deepEqual(tile, pixel(dithered, x, y))
    }
  }
})

test('input validation', () => {
  assert.throws(() => renderGradient(-1, 2, { x: 0, y: 0 }, spec()), RangeError)
  assert.throws(() => renderGradient(2.5, 2, { x: 0, y: 0 }, spec()), RangeError)
  assert.equal(renderGradient(0, 3, { x: 0, y: 0 }, spec()).data.length, 0)
})
