'use strict'
// WP4 brush engine (src/imaging/brush.ts): dab coverage and area, tip shape, dab placement at equal arc
// length with smoothing and pressure, per-stroke coverage accumulation and the opacity cap.
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

const { accumulateDab, compositeStroke, dabMask, placeDabs } = load('imaging/brush.ts')

const tip = (diameter, hardness = 1, roundness = 1, angle = 0) => ({ diameter, hardness, roundness, angle })
const sum = (array) => array.reduce((total, value) => total + value, 0)
const sample = (x, y, pressure = 1, time = 0) => ({ x, y, pressure, time })

test('a hard dab covers pi r^2 within 2%', () => {
  for (const diameter of [3, 6, 10, 25, 60, 151]) {
    const { size, data } = dabMask(tip(diameter))
    assert.equal(size % 2, 1, 'odd grid centred on a pixel')
    assert.equal(data.length, size * size)
    const area = (Math.PI * diameter * diameter) / 4
    assert.ok(Math.abs(sum(data) - area) <= 0.02 * area, `diameter ${diameter}: ${sum(data)} vs ${area}`)
    const centre = (size - 1) / 2
    assert.equal(data[centre * size + centre], 1)
    assert.equal(data[0], 0)
    // Anti-aliased edge: some pixels are partially covered.
    assert.ok(data.some((value) => value > 0 && value < 1))
  }
  // Accumulating one full-flow dab at a sub-pixel position covers the same area.
  const width = 80
  const coverage = new Float32Array(width * width)
  accumulateDab(coverage, width, width, 0, 0, { x: 40.3, y: 39.7, diameter: 30, flow: 1 }, tip(30))
  assert.ok(Math.abs(sum(coverage) - Math.PI * 225) <= 0.02 * Math.PI * 225)
})

test('soft dabs fall off smoothly; elliptical tips keep their area and rotate', () => {
  const soft = dabMask(tip(41, 0))
  const hard = dabMask(tip(41, 1))
  const c = (soft.size - 1) / 2
  assert.equal(soft.data[c * soft.size + c], 1)
  for (let x = c; x < soft.size - 1; x += 1) {
    assert.ok(soft.data[c * soft.size + x + 1] <= soft.data[c * soft.size + x] + 1e-9, 'monotone towards the edge')
  }
  assert.ok(sum(soft.data) < sum(hard.data) * 0.7)
  // Half hardness: fully opaque out to R * h.
  const half = dabMask(tip(41, 0.5))
  assert.equal(half.data[c * half.size + c + 9], 1)
  assert.ok(half.data[c * half.size + c + 15] < 1)
  // Roundness 0.5: an ellipse with semi-axes R and R / 2.
  const flat = dabMask(tip(60, 1, 0.5, 0))
  const ellipse = Math.PI * 30 * 15
  assert.ok(Math.abs(sum(flat.data) - ellipse) <= 0.02 * ellipse)
  const m = (flat.size - 1) / 2
  assert.equal(flat.data[m * flat.size + m + 25], 1, 'long axis horizontal at 0 degrees')
  assert.equal(flat.data[(m + 25) * flat.size + m], 0)
  // 90 degrees transposes the tip.
  const turned = dabMask(tip(60, 1, 0.5, 90))
  let worst = 0
  for (let y = 0; y < flat.size; y += 1) {
    for (let x = 0; x < flat.size; x += 1) worst = Math.max(worst, Math.abs(turned.data[x * flat.size + y] - flat.data[y * flat.size + x]))
  }
  assert.ok(worst < 1e-6, `rotation mismatch ${worst}`)
})

test('dab spacing is even at 25% along a straight stroke, across calls', () => {
  const brush = tip(20)
  const carry = { distance: 0, last: null }
  const dabs = []
  // Irregular sample spacing, several calls.
  const xs = [0, 3, 3.5, 11, 30, 31, 57.25, 90, 100]
  for (const x of xs) dabs.push(...placeDabs([sample(x, 10)], brush, 0.25, 0, carry))
  assert.deepEqual({ x: dabs[0].x, y: dabs[0].y }, { x: 0, y: 10 }, 'the first sample always gets a dab')
  for (let i = 1; i < dabs.length; i += 1) {
    assert.ok(Math.abs(dabs[i].x - dabs[i - 1].x - 5) < 1e-9, `gap ${i}: ${dabs[i].x - dabs[i - 1].x}`)
    assert.equal(dabs[i].y, 10)
    assert.equal(dabs[i].diameter, 20)
    assert.equal(dabs[i].flow, 1)
  }
  assert.equal(dabs.length, 21)
  assert.ok(Math.abs(carry.distance - 0) < 1e-9)
  // Spacing follows the arc length of a polyline, including around corners.
  const corner = { distance: 0, last: null }
  const path2 = placeDabs([sample(0, 0), sample(7, 0), sample(7, 9)], tip(8), 0.25, 0, corner)
  assert.equal(path2.length, 9)
  assert.deepEqual({ x: path2[3].x, y: path2[3].y }, { x: 6, y: 0 })
  assert.deepEqual({ x: path2[4].x, y: path2[4].y }, { x: 7, y: 1 })
})

test('pressure scales size and flow when enabled; smoothing lags and catches up', () => {
  const brush = tip(40)
  const carry = { distance: 0, last: null }
  const dabs = placeDabs([sample(0, 0, 0.5), sample(100, 0, 0.5)], brush, 0.25, 0, carry, { pressureSize: true, pressureFlow: true, flow: 0.8 })
  const expectedDiameter = 40 * (0.1 + 0.9 * 0.5)
  for (const dab of dabs) {
    assert.ok(Math.abs(dab.diameter - expectedDiameter) < 1e-9)
    assert.ok(Math.abs(dab.flow - 0.4) < 1e-9)
  }
  for (let i = 1; i < dabs.length; i += 1) assert.ok(Math.abs(dabs[i].x - dabs[i - 1].x - 0.25 * expectedDiameter) < 1e-9)
  // Without dynamics the pen pressure is ignored.
  const plain = placeDabs([sample(0, 0, 0.2)], brush, 0.25, 0, { distance: 0, last: null })
  assert.equal(plain[0].diameter, 40)
  assert.equal(plain[0].flow, 1)
  // Smoothing 1 (alpha 0.2): the stroke trails the pointer; smoothing 0 on release catches up.
  const smooth = { distance: 0, last: null }
  placeDabs([sample(0, 0), sample(100, 0)], tip(4), 0.25, 1, smooth)
  assert.ok(Math.abs(smooth.last.x - 20) < 1e-9)
  const tail = placeDabs([sample(100, 0)], tip(4), 0.25, 0, smooth)
  assert.ok(Math.abs(smooth.last.x - 100) < 1e-9)
  assert.ok(Math.abs(tail.at(-1).x - 100) <= 1)
})

test('accumulated stroke coverage never exceeds the opacity cap', () => {
  const width = 120
  const height = 60
  const coverage = new Float32Array(width * height)
  const brush = tip(24, 0.6)
  const carry = { distance: 0, last: null }
  // A scribble that crosses itself many times.
  const samples = []
  for (let i = 0; i <= 200; i += 1) samples.push(sample(60 + 40 * Math.sin(i / 7), 30 + 20 * Math.sin(i / 3), 1, i))
  const dabs = placeDabs(samples, brush, 0.1, 0, carry, { flow: 0.5 })
  assert.ok(dabs.length > 100)
  for (const dab of dabs) accumulateDab(coverage, width, height, 0, 0, dab, brush)
  const peak = Math.max(...coverage)
  assert.ok(peak <= 1 && peak > 0.99, `peak coverage ${peak}`)
  // Painting black at 60% opacity over white can never get darker than 40% grey.
  const white = { width, height, data: new Uint8ClampedArray(width * height * 4).fill(255) }
  const painted = compositeStroke(white, coverage, { color: { r: 0, g: 0, b: 0 }, opacity: 0.6 })
  let darkest = 255
  for (let i = 0; i < width * height; i += 1) {
    darkest = Math.min(darkest, painted.data[i * 4])
    assert.equal(painted.data[i * 4 + 3], 255)
  }
  assert.ok(darkest >= Math.round(255 * 0.4) - 1, `darkest ${darkest}`)
  assert.notEqual(painted.data, white.data, 'the pre-stroke pixels are not modified')
  assert.equal(white.data[0], 255)
})

test('compositeStroke: source-over on transparency, eraser, transparency lock and selection', () => {
  const before = { width: 3, height: 1, data: Uint8ClampedArray.from([255, 0, 0, 255, 0, 0, 0, 0, 0, 0, 255, 128]) }
  const full = Float32Array.from([1, 1, 1])
  // Painting at 50% onto transparent keeps the paint colour (straight alpha) with half alpha.
  const paint = compositeStroke(before, full, { color: { r: 0, g: 255, b: 0 }, opacity: 0.5 })
  assert.deepEqual([...paint.data.slice(4, 8)], [0, 255, 0, 128])
  assert.deepEqual([...paint.data.slice(0, 4)], [128, 128, 0, 255])
  // Eraser.
  const erased = compositeStroke(before, full, { color: { r: 0, g: 0, b: 0 }, opacity: 0.5, erase: true })
  assert.deepEqual([...erased.data], [255, 0, 0, 128, 0, 0, 0, 0, 0, 0, 255, 64])
  // Transparency lock keeps alpha and leaves empty pixels empty.
  const locked = compositeStroke(before, full, { color: { r: 0, g: 255, b: 0 }, opacity: 1, preserveAlpha: true })
  assert.deepEqual([...locked.data], [0, 255, 0, 255, 0, 0, 0, 0, 0, 255, 0, 128])
  assert.deepEqual([...compositeStroke(before, full, { color: { r: 0, g: 0, b: 0 }, opacity: 1, erase: true, preserveAlpha: true }).data], [...before.data])
  // The selection scales the coverage.
  const selection = { width: 3, height: 1, data: Uint8Array.from([0, 255, 255]) }
  const masked = compositeStroke(before, full, { color: { r: 0, g: 0, b: 255 }, opacity: 1 }, selection)
  assert.deepEqual([...masked.data.slice(0, 4)], [255, 0, 0, 255])
  assert.deepEqual([...masked.data.slice(4, 8)], [0, 0, 255, 255])
})

test('accumulateDab works in tiles and reports the touched rectangle', () => {
  const brush = tip(10)
  const dab = { x: 260, y: 5, diameter: 10, flow: 1 }
  // The dab straddles two 256-px tiles.
  const left = new Float32Array(256 * 256)
  const right = new Float32Array(256 * 256)
  const a = accumulateDab(left, 256, 256, 0, 0, dab, brush)
  const b = accumulateDab(right, 256, 256, 256, 0, dab, brush)
  assert.deepEqual(a, { x: 254, y: 0, width: 2, height: 11 })
  assert.deepEqual(b, { x: 256, y: 0, width: 10, height: 11 })
  const whole = new Float32Array(512 * 16)
  accumulateDab(whole, 512, 16, 0, 0, dab, brush)
  for (let y = 0; y < 11; y += 1) {
    for (let x = 250; x < 270; x += 1) {
      const tiled = x < 256 ? left[y * 256 + x] : right[y * 256 + x - 256]
      assert.ok(Math.abs(tiled - whole[y * 512 + x]) < 1e-7)
    }
  }
  const miss = accumulateDab(left, 256, 256, 0, 0, { x: 900, y: 900, diameter: 10, flow: 1 }, brush)
  assert.equal(miss.width, 0)
  assert.throws(() => accumulateDab(new Float32Array(3), 2, 2, 0, 0, dab, brush), RangeError)
})
