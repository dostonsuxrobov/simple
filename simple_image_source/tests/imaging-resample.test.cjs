'use strict'
// WP2 resampling: constant images stay constant, a 2x reduction of a checkerboard is uniform grey, transparent
// edges never darken, nearest/area are exact, kernels and border mirroring behave, the 'resample' worker op.
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

const S = load('imaging/resample.ts')
const R = load('imaging/random.ts')
const { handlers } = load('shared/worker-ops/color.ts')

const METHODS = ['nearest', 'bilinear', 'bicubic', 'lanczos3', 'area', 'auto']

function image(width, height, f) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const [r, g, b, a = 255] = f(x, y)
      data.set([r, g, b, a], (y * width + x) * 4)
    }
  }
  return { width, height, data }
}

function px(buffer, x, y) {
  const i = (y * buffer.width + x) * 4
  return [...buffer.data.subarray(i, i + 4)]
}

test('a constant image stays constant for every method, up and down, odd sizes included', () => {
  const src = image(37, 23, () => [200, 100, 50, 255])
  const translucent = image(16, 9, () => [10, 220, 130, 90])
  for (const method of METHODS) {
    for (const [width, height] of [[18, 11], [5, 3], [80, 47], [37, 70], [1, 1], [111, 2]]) {
      for (const source of [src, translucent]) {
        const out = S.resample(source, width, height, method)
        assert.equal(out.width, width)
        assert.equal(out.height, height)
        const expected = px(source, 0, 0)
        for (let i = 0; i < out.data.length; i += 4) {
          assert.deepEqual([...out.data.subarray(i, i + 4)], expected, `${method} ${width}x${height}`)
        }
      }
    }
  }
})

test('2x reduction of a one-pixel checkerboard is uniform mid grey', () => {
  const board = image(64, 48, (x, y) => ((x + y) % 2 ? [255, 255, 255] : [0, 0, 0]))
  for (const method of ['bilinear', 'bicubic', 'lanczos3', 'area', 'auto']) {
    const out = S.resample(board, 32, 24, method)
    const first = px(out, 0, 0)
    assert.ok(first[0] >= 127 && first[0] <= 128, `${method}: ${first}`)
    for (let y = 0; y < 24; y += 1) {
      for (let x = 0; x < 32; x += 1) assert.deepEqual(px(out, x, y), first, `${method} uniform at ${x},${y}`)
    }
  }
  const odd = image(63, 47, (x, y) => ((x + y) % 2 ? [255, 255, 255] : [0, 0, 0]))
  const shrunk = S.resample(odd, 21, 15, 'auto')
  for (let i = 0; i < shrunk.data.length; i += 4) assert.ok(Math.abs(shrunk.data[i] - 127.5) <= 20, 'a 3x reduction is close to grey')
})

test('transparent edges never darken when shrinking or enlarging', () => {
  const square = image(40, 40, (x, y) => (x >= 10 && x < 30 && y >= 10 && y < 30 ? [255, 255, 255, 255] : [0, 0, 0, 0]))
  for (const method of METHODS) {
    for (const [width, height] of [[20, 20], [13, 13], [97, 97]]) {
      const out = S.resample(square, width, height, method)
      let edge = 0
      for (let i = 0; i < out.data.length; i += 4) {
        const a = out.data[i + 3]
        if (a === 0) continue
        if (a < 255) edge += 1
        assert.ok(out.data[i] >= 254 && out.data[i + 1] >= 254 && out.data[i + 2] >= 254, `${method} ${width}: ${[...out.data.subarray(i, i + 4)]}`)
      }
      // An exact 2x area reduction of a square on even pixel boundaries has no partial pixels.
      if (method !== 'nearest' && !(method === 'area' && width === 20)) assert.ok(edge > 0, `${method} ${width}: soft edges exist`)
    }
  }
})

test('nearest replicates and picks exact pixels; area averages exactly; same size is an identical copy', () => {
  const next = R.mulberry32(3)
  const src = image(5, 4, () => [next() * 256, next() * 256, next() * 256, next() * 256])
  const doubled = S.resample(src, 10, 8, 'nearest')
  for (let y = 0; y < 8; y += 1) for (let x = 0; x < 10; x += 1) assert.deepEqual(px(doubled, x, y), px(src, x >> 1, y >> 1))
  assert.deepEqual([...S.resample(doubled, 5, 4, 'nearest').data], [...src.data])

  const quad = image(4, 4, (x, y) => [x * 40 + y * 4, 10 * (x + y), 255 - x * 50, 255])
  const halved = S.resample(quad, 2, 2, 'area')
  for (let y = 0; y < 2; y += 1) {
    for (let x = 0; x < 2; x += 1) {
      for (let c = 0; c < 3; c += 1) {
        let sum = 0
        for (let dy = 0; dy < 2; dy += 1) for (let dx = 0; dx < 2; dx += 1) sum += px(quad, x * 2 + dx, y * 2 + dy)[c]
        assert.ok(Math.abs(px(halved, x, y)[c] - sum / 4) <= 0.5)
      }
    }
  }
  const copy = S.resample(src, 5, 4, 'lanczos3')
  assert.notEqual(copy.data, src.data)
  assert.deepEqual([...copy.data], [...src.data])
})

test('enlarging a linear gradient stays monotone and hits the end values', () => {
  const ramp = image(16, 1, (x) => [x * 17, x * 17, x * 17])
  for (const method of ['bilinear', 'bicubic', 'lanczos3', 'auto']) {
    const out = S.resample(ramp, 64, 4, method)
    for (let y = 0; y < 4; y += 1) {
      for (let x = 1; x < 64; x += 1) assert.ok(px(out, x, y)[0] >= px(out, x - 1, y)[0] - 1, `${method} monotone at ${x}`)
    }
    assert.ok(px(out, 0, 0)[0] <= 2 && px(out, 63, 0)[0] >= 253)
  }
})

test('kernels, border mirroring and validation', () => {
  assert.equal(S.cubicWeight(0), 1)
  assert.equal(S.cubicWeight(1), 0)
  assert.equal(S.cubicWeight(2), 0)
  assert.equal(S.lanczosWeight(0), 1)
  assert.equal(S.lanczosWeight(1), 0)
  assert.equal(S.lanczosWeight(-2), 0)
  assert.equal(S.lanczosWeight(3), 0)
  assert.equal(S.triangleWeight(0.25), 0.75)
  for (const t of [0, 0.2, 0.5, 0.9]) {
    const cubic = [-1, 0, 1, 2].reduce((sum, k) => sum + S.cubicWeight(t - k), 0)
    assert.ok(Math.abs(cubic - 1) < 1e-12, 'Catmull-Rom is a partition of unity')
  }
  assert.deepEqual([-3, -2, -1, 0, 4, 5, 6, 7, 9].map((i) => S.reflectIndex(i, 5)), [3, 2, 1, 0, 4, 3, 2, 1, 1])
  assert.equal(S.reflectIndex(-4, 1), 0)
  const src = image(4, 4, () => [1, 2, 3, 4])
  assert.throws(() => S.resample(src, 0, 4), RangeError)
  assert.throws(() => S.resample(src, 2.5, 4), RangeError)
  assert.throws(() => S.resample(src, 200000, 4), RangeError)
  assert.throws(() => S.resample({ width: 0, height: 0, data: new Uint8ClampedArray(0) }, 2, 2), RangeError)
})

test('resample reports progress, honours abort and the worker op wraps it', async () => {
  const next = R.mulberry32(5)
  const src = image(300, 200, () => [next() * 256, next() * 256, next() * 256, 255])
  const seen = []
  S.resample(src, 1200, 800, 'auto', { onProgress: (value) => seen.push(value) })
  assert.ok(seen.length >= 2 && seen.length <= 21, `${seen.length} progress calls`)
  assert.equal(seen.at(-1), 1)
  const small = []
  S.resample(src, 120, 80, 'auto', { onProgress: (value) => small.push(value) })
  assert.deepEqual(small, [1], 'a single chunk reports completion once')
  const controller = new AbortController()
  assert.throws(() => S.resample(src, 700, 500, 'bicubic', {
    signal: controller.signal,
    onProgress: (value) => { if (value > 0.1) controller.abort() },
  }), { name: 'AbortError' })
  const before = new Uint8ClampedArray(src.data)
  const out = await handlers.resample({ src, width: 150, height: 100, method: 'lanczos3' }, {})
  assert.deepEqual([...out.data], [...S.resample(src, 150, 100, 'lanczos3').data])
  assert.deepEqual(src.data, before, 'the input is left intact')
})

test('startResample streams rows: any chunking gives the same bytes, rows must come in order', () => {
  const next = R.mulberry32(12)
  const src = image(53, 41, () => [next() * 256, next() * 256, next() * 256, next() * 256])
  for (const method of ['bilinear', 'bicubic', 'lanczos3', 'area', 'auto', 'nearest']) {
    for (const [width, height] of [[17, 13], [120, 97], [53, 200]]) {
      const whole = S.resample(src, width, height, method)
      for (const rows of [1, 5, height]) {
        const run = S.startResample(src, width, height, method)
        for (let start = 0; start < run.rows; start += rows) run.process(start, Math.min(run.rows, start + rows))
        assert.deepEqual([...run.output.data], [...whole.data], `${method} ${width}x${height} in chunks of ${rows}`)
      }
    }
  }
  const run = S.startResample(src, 20, 20, 'bicubic')
  run.process(0, 5)
  assert.throws(() => run.process(10, 12), RangeError)
})

test('a 6x reduction with auto is close to an exact area average', () => {
  const next = R.mulberry32(9)
  const src = image(240, 180, (x, y) => {
    const v = 128 + 100 * Math.sin(x / 13) * Math.cos(y / 17)
    return [v, 255 - v, (v + next() * 8) | 0, 255]
  })
  const auto = S.resample(src, 40, 30, 'auto')
  const area = S.resample(src, 40, 30, 'area')
  let total = 0
  for (let i = 0; i < auto.data.length; i += 4) total += Math.abs(auto.data[i] - area.data[i])
  assert.ok(total / (auto.data.length / 4) < 3, `mean difference ${(total / (auto.data.length / 4)).toFixed(2)}`)
})
