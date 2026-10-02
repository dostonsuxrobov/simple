'use strict'
// WP4 distance transforms (src/imaging/distance.ts): exact Euclidean distances in O(N), checked against
// brute force for both metrics (pixel centres, and pixel squares used by Expand / Contract).
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

const { distanceTransform, regionDistanceTransform } = load('imaging/distance.ts')

function random(seed) {
  let state = seed >>> 0
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    return state / 4294967296
  }
}

function bruteForce(binary, width, height, region) {
  const out = new Float64Array(width * height)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let best = Infinity
      for (let qy = 0; qy < height; qy += 1) {
        for (let qx = 0; qx < width; qx += 1) {
          if (!binary[qy * width + qx]) continue
          let dx = Math.abs(x - qx)
          let dy = Math.abs(y - qy)
          if (region) {
            dx = Math.max(dx - 0.5, 0)
            dy = Math.max(dy - 0.5, 0)
          }
          best = Math.min(best, dx * dx + dy * dy)
        }
      }
      out[y * width + x] = Math.sqrt(best)
    }
  }
  return out
}

test('both transforms match brute force on random sparse and dense site sets', () => {
  const rnd = random(20261002)
  for (let trial = 0; trial < 24; trial += 1) {
    const width = 1 + Math.floor(rnd() * 33)
    const height = 1 + Math.floor(rnd() * 33)
    const density = trial % 3 === 0 ? 0.6 : rnd() * 0.15
    const binary = new Uint8Array(width * height)
    for (let i = 0; i < binary.length; i += 1) binary[i] = rnd() < density ? 1 + Math.floor(rnd() * 254) : 0
    for (const region of [false, true]) {
      const fast = region ? regionDistanceTransform(binary, width, height) : distanceTransform(binary, width, height)
      const slow = bruteForce(binary, width, height, region)
      assert.ok(fast instanceof Float32Array)
      for (let i = 0; i < fast.length; i += 1) {
        if (slow[i] === Infinity) assert.equal(fast[i], Infinity)
        else assert.ok(Math.abs(fast[i] - slow[i]) < 1e-4, `${region ? 'region' : 'centre'} ${width}x${height} pixel ${i}: ${fast[i]} vs ${slow[i]}`)
      }
    }
  }
})

test('a single site gives exact Euclidean distances for both metrics', () => {
  const width = 21
  const height = 17
  const binary = new Uint8Array(width * height)
  binary[8 * width + 10] = 1
  const centre = distanceTransform(binary, width, height)
  const region = regionDistanceTransform(binary, width, height)
  assert.equal(centre[8 * width + 10], 0)
  assert.equal(region[8 * width + 10], 0)
  // (dx, dy) = (3, 4): 5 between centres; to the site's square sqrt(2.5^2 + 3.5^2).
  assert.ok(Math.abs(centre[12 * width + 13] - 5) < 1e-6)
  assert.ok(Math.abs(region[12 * width + 13] - Math.hypot(2.5, 3.5)) < 1e-6)
  // Along an axis the square metric is the centre distance minus half a pixel.
  assert.ok(Math.abs(region[8 * width + 16] - 5.5) < 1e-6)
  assert.ok(Math.abs(centre[8 * width + 16] - 6) < 1e-6)
})

test('no sites yields Infinity everywhere; empty images are fine', () => {
  const none = distanceTransform(new Uint8Array(12), 4, 3)
  assert.ok(Array.from(none).every((value) => value === Infinity))
  assert.ok(Array.from(regionDistanceTransform(new Uint8Array(12), 3, 4)).every((value) => value === Infinity))
  assert.equal(distanceTransform(new Uint8Array(0), 0, 5).length, 0)
  const all = regionDistanceTransform(new Uint8Array(6).fill(9), 3, 2)
  assert.ok(Array.from(all).every((value) => value === 0))
})

test('input validation, progress and cancellation', () => {
  assert.throws(() => distanceTransform(new Uint8Array(5), 2, 3), RangeError)
  assert.throws(() => distanceTransform(new Uint8Array(6), 2.5, 3), RangeError)
  const fractions = []
  const binary = new Uint8Array(200 * 200)
  binary[0] = 1
  distanceTransform(binary, 200, 200, { onProgress: (value) => fractions.push(value) })
  assert.ok(fractions.length > 0 && fractions.length <= 20, `${fractions.length} progress calls`)
  assert.equal(fractions.at(-1), 1)
  const controller = new AbortController()
  controller.abort()
  assert.throws(() => distanceTransform(binary, 200, 200, { signal: controller.signal }), { name: 'AbortError' })
})
