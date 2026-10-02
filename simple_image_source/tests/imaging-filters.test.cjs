'use strict'
// WP2 filters and looks: blur/sharpen/noise/median/stylize behaviour, transparent-edge handling, stripe
// margins that make striped worker runs match whole-image runs, abort/progress, the 'filter' and 'look'
// worker ops, and the radius-10 Gaussian budget on 4 MP.
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

const F = load('imaging/filters.ts')
const L = load('imaging/looks.ts')
const R = load('imaging/random.ts')
const { handlers } = load('shared/worker-ops/color.ts')

function image(width, height, f) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const [r, g, b, a = 255] = f(x, y)
      const i = (y * width + x) * 4
      data[i] = r
      data[i + 1] = g
      data[i + 2] = b
      data[i + 3] = a
    }
  }
  return { width, height, data }
}

function randomImage(width, height, seed = 1, opaque = true) {
  const next = R.mulberry32(seed)
  return image(width, height, () => [next() * 256, next() * 256, next() * 256, opaque ? 255 : next() * 256])
}

function px(buffer, x, y) {
  const i = (y * buffer.width + x) * 4
  return [...buffer.data.subarray(i, i + 4)]
}

function maxDiff(a, b) {
  let max = 0
  for (let i = 0; i < a.data.length; i += 1) max = Math.max(max, Math.abs(a.data[i] - b.data[i]))
  return max
}

function sameBytes(a, b) {
  return a.width === b.width && a.height === b.height && Buffer.compare(Buffer.from(a.data), Buffer.from(b.data)) === 0
}

function channelMean(buffer, channel) {
  let sum = 0
  for (let i = channel; i < buffer.data.length; i += 4) sum += buffer.data[i]
  return sum / (buffer.data.length / 4)
}

test('catalogue: ten filters with labels, fresh defaults and margins', () => {
  assert.equal(F.FILTER_TYPES.length, 10)
  for (const type of F.FILTER_TYPES) {
    const spec = F.defaultFilter(type)
    assert.equal(spec.type, type)
    assert.ok(F.filterLabel(type).length > 0)
    assert.ok(Number.isInteger(F.filterMargin(spec)) && F.filterMargin(spec) >= 0)
  }
  assert.equal(F.filterLabel('pixelate'), 'Mosaic')
  assert.equal(F.filterLabel('sharpen', { type: 'sharpen', strength: 'more' }), 'Sharpen More')
  assert.equal(F.filterMargin({ type: 'gaussian-blur', radius: 4 }), 12)
  assert.equal(F.filterMargin({ type: 'add-noise', amount: 10, distribution: 'uniform', monochromatic: false, seed: 1 }), 0)
  assert.equal(F.stripeMargin({ type: 'add-noise', amount: 10, distribution: 'uniform', monochromatic: false, seed: 1 }), null)
  assert.equal(F.stripeMargin({ type: 'pixelate', cellSize: 8 }), null)
  assert.equal(F.stripeMargin({ type: 'median', radius: 3 }), 3)
  assert.throws(() => F.defaultFilter('nope'), RangeError)
  assert.throws(() => F.applyFilter(randomImage(2, 2), { type: 'nope' }), RangeError)
})

test('Gaussian blur preserves constant images and the mean, and radius 0.1 is near identity', () => {
  const constant = image(37, 23, () => [200, 100, 50, 255])
  for (const radius of [0.1, 0.7, 1.5, 2, 5.5, 30, 250]) {
    assert.ok(sameBytes(F.applyFilter(constant, { type: 'gaussian-blur', radius }), constant), `radius ${radius} keeps a constant image`)
  }
  const noisy = randomImage(64, 48, 5)
  const blurred = F.applyFilter(noisy, { type: 'gaussian-blur', radius: 4 })
  for (let channel = 0; channel < 3; channel += 1) {
    assert.ok(Math.abs(channelMean(blurred, channel) - channelMean(noisy, channel)) < 1, 'the mean survives')
  }
  assert.ok(maxDiff(F.applyFilter(noisy, { type: 'gaussian-blur', radius: 0.1 }), noisy) <= 1)
  const impulse = image(41, 41, (x, y) => (x === 20 && y === 20 ? [255, 255, 255] : [0, 0, 0]))
  const spread = F.applyFilter(impulse, { type: 'gaussian-blur', radius: 3 })
  assert.deepEqual(px(spread, 17, 20), px(spread, 23, 20), 'symmetric response')
  assert.deepEqual(px(spread, 20, 17), px(spread, 20, 23))
  assert.deepEqual(px(spread, 20, 23), px(spread, 23, 20))
})

test('blurs leave no dark fringe beside a transparent edge', () => {
  const half = image(40, 20, (x) => (x < 20 ? [255, 255, 255, 255] : [0, 0, 0, 0]))
  const specs = [
    { type: 'gaussian-blur', radius: 1.2 },
    { type: 'gaussian-blur', radius: 6 },
    { type: 'motion-blur', angle: 0, distance: 15 },
    { type: 'motion-blur', angle: 35, distance: 15 },
    { type: 'unsharp-mask', amount: 150, radius: 3, threshold: 0 },
    { type: 'reduce-noise', strength: 5, preserveDetails: 30 },
    { type: 'median', radius: 2 },
    { type: 'pixelate', cellSize: 6 },
  ]
  for (const spec of specs) {
    const out = F.applyFilter(half, spec)
    let visible = 0
    for (let i = 0; i < out.data.length; i += 4) {
      if (out.data[i + 3] === 0) continue
      visible += 1
      assert.ok(out.data[i] >= 254 && out.data[i + 1] >= 254 && out.data[i + 2] >= 254, `${spec.type}: visible pixels stay white, got ${[...out.data.subarray(i, i + 4)]}`)
    }
    assert.ok(visible > 0)
  }
  const soft = F.applyFilter(half, { type: 'gaussian-blur', radius: 4 })
  assert.ok(px(soft, 20, 10)[3] > 0 && px(soft, 20, 10)[3] < 255, 'alpha is blurred across the edge')
})

test('Unsharp Mask on a flat image is the identity; it raises edge contrast and honours the threshold', () => {
  const flat = image(30, 20, () => [120, 60, 200, 255])
  assert.ok(sameBytes(F.applyFilter(flat, { type: 'unsharp-mask', amount: 500, radius: 5, threshold: 0 }), flat))
  assert.ok(sameBytes(F.applyFilter(flat, { type: 'unsharp-mask', amount: 80, radius: 0.5, threshold: 0 }), flat))
  const edge = image(30, 4, (x) => (x < 15 ? [80, 80, 80] : [160, 160, 160]))
  const sharp = F.applyFilter(edge, { type: 'unsharp-mask', amount: 100, radius: 2, threshold: 0 })
  assert.ok(px(sharp, 14, 1)[0] < 80 && px(sharp, 15, 1)[0] > 160, 'overshoot on both sides of the edge')
  assert.deepEqual(px(sharp, 2, 1), [80, 80, 80, 255], 'far from the edge nothing changes')
  const subtle = image(30, 4, (x) => (x < 15 ? [100, 100, 100] : [104, 104, 104]))
  assert.ok(sameBytes(F.applyFilter(subtle, { type: 'unsharp-mask', amount: 300, radius: 2, threshold: 10 }), subtle), 'differences below the threshold are left alone')
  const sharpened = F.applyFilter(edge, { type: 'sharpen', strength: 'normal' })
  const more = F.applyFilter(edge, { type: 'sharpen', strength: 'more' })
  assert.ok(px(more, 14, 1)[0] < px(sharpened, 14, 1)[0], 'Sharpen More is stronger')
  assert.ok(sameBytes(F.applyFilter(flat, { type: 'sharpen', strength: 'more' }), flat))
})

test('Median removes isolated salt noise and keeps straight edges', () => {
  const salt = image(40, 30, (x, y) => ((x * 7 + y * 13) % 29 === 0 ? [255, 255, 255] : [10, 20, 30]))
  const cleaned = F.applyFilter(salt, { type: 'median', radius: 1 })
  for (let y = 0; y < 30; y += 1) for (let x = 0; x < 40; x += 1) assert.deepEqual(px(cleaned, x, y), [10, 20, 30, 255])
  const edge = image(20, 20, (x) => (x < 10 ? [0, 0, 0] : [255, 255, 255]))
  assert.ok(sameBytes(F.applyFilter(edge, { type: 'median', radius: 3 }), edge))
  const wide = F.applyFilter(randomImage(30, 30, 9), { type: 'median', radius: 5 })
  assert.equal(wide.width, 30)
})

test('Add Noise is deterministic per seed, keeps alpha and the average', () => {
  const gray = image(64, 64, () => [128, 128, 128, 200])
  const spec = { type: 'add-noise', amount: 20, distribution: 'gaussian', monochromatic: false, seed: 42 }
  const a = F.applyFilter(gray, spec)
  assert.ok(sameBytes(a, F.applyFilter(gray, spec)), 'same seed, same pixels')
  assert.ok(!sameBytes(a, F.applyFilter(gray, { ...spec, seed: 43 })), 'another seed differs')
  for (let i = 3; i < a.data.length; i += 4) assert.equal(a.data[i], 200)
  assert.ok(Math.abs(channelMean(a, 0) - 128) < 2)
  const mono = F.applyFilter(gray, { ...spec, monochromatic: true, distribution: 'uniform' })
  for (let i = 0; i < mono.data.length; i += 4) {
    assert.equal(mono.data[i], mono.data[i + 1])
    assert.equal(mono.data[i], mono.data[i + 2])
  }
  const transparent = image(4, 4, () => [10, 10, 10, 0])
  assert.ok(sameBytes(F.applyFilter(transparent, spec), transparent), 'hidden pixels are untouched')
})

test('Motion blur averages along its angle only', () => {
  const columns = image(40, 30, (x) => (x % 2 ? [255, 255, 255] : [0, 0, 0]))
  const horizontal = F.applyFilter(columns, { type: 'motion-blur', angle: 0, distance: 10 })
  for (let y = 0; y < 30; y += 1) for (let x = 6; x < 34; x += 1) assert.ok(Math.abs(px(horizontal, x, y)[0] - 127.5) <= 13, `stripes across the motion blur out (${px(horizontal, x, y)[0]})`)
  const vertical = F.applyFilter(columns, { type: 'motion-blur', angle: 90, distance: 10 })
  assert.ok(sameBytes(vertical, columns), 'stripes along the motion stay sharp')
  const constant = image(25, 25, () => [30, 60, 90, 255])
  assert.ok(sameBytes(F.applyFilter(constant, { type: 'motion-blur', angle: 33, distance: 40 }), constant))
  const noisy = randomImage(50, 50, 4)
  const diagonal = F.applyFilter(noisy, { type: 'motion-blur', angle: -60, distance: 25 })
  assert.ok(Math.abs(channelMean(diagonal, 1) - channelMean(noisy, 1)) < 3)
})

test('Reduce Noise smooths flat noise and keeps strong edges', () => {
  const next = R.mulberry32(8)
  const noisy = image(60, 40, (x) => {
    const base = x < 30 ? 60 : 200
    const n = Math.round((next() - 0.5) * 16)
    return [base + n, base + n, base + n]
  })
  const out = F.applyFilter(noisy, { type: 'reduce-noise', strength: 6, preserveDetails: 40 })
  const variance = (buffer, x0, x1) => {
    let sum = 0
    let squares = 0
    let count = 0
    for (let y = 5; y < 35; y += 1) {
      for (let x = x0; x < x1; x += 1) {
        const v = px(buffer, x, y)[0]
        sum += v
        squares += v * v
        count += 1
      }
    }
    return squares / count - (sum / count) ** 2
  }
  assert.ok(variance(out, 5, 22) < variance(noisy, 5, 22) / 3, 'flat noise drops')
  assert.ok(px(out, 26, 20)[0] < 90 && px(out, 33, 20)[0] > 170, 'the edge survives')
})

test('Mosaic, Emboss and Find Edges', () => {
  const src = randomImage(20, 12, 3)
  const mosaic = F.applyFilter(src, { type: 'pixelate', cellSize: 8 })
  assert.deepEqual(px(mosaic, 0, 0), px(mosaic, 7, 7))
  assert.deepEqual(px(mosaic, 8, 0), px(mosaic, 15, 7))
  assert.deepEqual(px(mosaic, 16, 8), px(mosaic, 19, 11), 'partial cells at the edge')
  let sum = 0
  for (let y = 0; y < 8; y += 1) for (let x = 0; x < 8; x += 1) sum += px(src, x, y)[1]
  assert.ok(Math.abs(px(mosaic, 3, 3)[1] - sum / 64) <= 1)
  const weighted = F.applyFilter(image(2, 1, (x) => (x === 0 ? [255, 0, 0, 255] : [0, 0, 255, 0])), { type: 'pixelate', cellSize: 2 })
  assert.deepEqual(px(weighted, 1, 0), [255, 0, 0, 128], 'alpha-weighted colour, averaged alpha')

  const flat = image(16, 16, () => [90, 140, 200, 77])
  const embossed = F.applyFilter(flat, { type: 'emboss', angle: 135, height: 3, amount: 200 })
  for (let i = 0; i < embossed.data.length; i += 4) assert.deepEqual([...embossed.data.subarray(i, i + 4)], [128, 128, 128, 77])
  const step = image(16, 16, (x) => (x < 8 ? [0, 0, 0] : [255, 255, 255]))
  const relief = F.applyFilter(step, { type: 'emboss', angle: 0, height: 2, amount: 100 })
  assert.notEqual(px(relief, 7, 5)[0], 128)
  assert.equal(px(relief, 2, 5)[0], 128)

  const plain = image(10, 10, () => [100, 100, 100])
  const edges = F.applyFilter(plain, { type: 'find-edges' })
  assert.deepEqual(px(edges, 5, 5), [255, 255, 255, 255], 'no edges: white')
  const boundary = F.applyFilter(step, { type: 'find-edges' })
  assert.ok(px(boundary, 7, 5)[0] < 50 && px(boundary, 2, 5)[0] === 255)
})

test('stripe margins make striped runs identical to whole-image runs', () => {
  const src = randomImage(37, 61, 12, false)
  const specs = [
    { type: 'gaussian-blur', radius: 1.3 },
    { type: 'gaussian-blur', radius: 3.5 },
    { type: 'unsharp-mask', amount: 120, radius: 2.5, threshold: 3 },
    { type: 'sharpen', strength: 'more' },
    { type: 'median', radius: 2 },
    { type: 'reduce-noise', strength: 3, preserveDetails: 50 },
    { type: 'emboss', angle: 45, height: 3, amount: 150 },
    { type: 'find-edges' },
    { type: 'motion-blur', angle: 0, distance: 9 },
    { type: 'motion-blur', angle: 25, distance: 12 },
  ]
  const rowBytes = src.width * 4
  for (const spec of specs) {
    const whole = F.applyFilter(src, spec)
    const margin = F.stripeMargin(spec)
    assert.equal(typeof margin, 'number')
    for (const [top, bottom] of [[0, 20], [20, 41], [41, 61]]) {
      const from = Math.max(0, top - margin)
      const to = Math.min(src.height, bottom + margin)
      const stripe = { width: src.width, height: to - from, data: src.data.slice(from * rowBytes, to * rowBytes) }
      const part = F.applyFilter(stripe, spec)
      const expected = whole.data.subarray(top * rowBytes, bottom * rowBytes)
      const actual = part.data.subarray((top - from) * rowBytes, (bottom - from) * rowBytes)
      assert.equal(Buffer.compare(Buffer.from(actual), Buffer.from(expected)), 0, `${spec.type} rows ${top}..${bottom}`)
    }
  }
})

test('filters honour abort and report bounded progress', () => {
  const src = randomImage(120, 90, 2)
  for (const type of F.FILTER_TYPES) {
    const seen = []
    F.applyFilter(src, F.defaultFilter(type), { onProgress: (value) => seen.push(value) })
    assert.ok(seen.length >= 1 && seen.length <= 22, `${type}: ${seen.length} progress calls`)
    assert.equal(seen.at(-1), 1)
    assert.ok(seen.every((value, index) => index === 0 || value >= seen[index - 1]))
    const early = new AbortController()
    early.abort()
    assert.throws(() => F.applyFilter(src, F.defaultFilter(type), { signal: early.signal }), { name: 'AbortError' })
  }
  const controller = new AbortController()
  assert.throws(() => F.applyFilter(src, { type: 'gaussian-blur', radius: 8 }, {
    signal: controller.signal,
    onProgress: (value) => { if (value > 0.2) controller.abort() },
  }), { name: 'AbortError' })
})

test('filter and look worker ops: masks, inputs left intact', async () => {
  const src = randomImage(30, 20, 6)
  const before = new Uint8ClampedArray(src.data)
  const mask = { width: 30, height: 20, data: Uint8Array.from({ length: 600 }, (_, i) => (i % 30 < 15 ? 255 : 0)) }
  const spec = { type: 'gaussian-blur', radius: 3 }
  const out = await handlers.filter({ src, spec, mask }, {})
  const full = F.applyFilter(src, spec)
  for (let y = 0; y < 20; y += 1) {
    assert.deepEqual(px(out, 3, y), px(full, 3, y), 'selected pixels are filtered')
    assert.deepEqual(px(out, 25, y), px(src, 25, y), 'unselected pixels are untouched')
  }
  assert.deepEqual(src.data, before)
  const look = await handlers.look({ src, look: 'mono', intensity: 100 }, {})
  assert.ok(sameBytes(look, L.applyLook(src, 'mono', 100)))
  assert.deepEqual(src.data, before)
  await assert.rejects(async () => handlers.filter({ src, spec, mask: { width: 2, height: 2, data: new Uint8Array(4) } }, {}), RangeError)
})

test('looks: nine looks, intensity blends with the original, vignette darkens corners only', () => {
  assert.deepEqual([...L.LOOK_IDS], ['none', 'vivid', 'warm', 'cool', 'mono', 'sepia', 'vintage', 'dramatic', 'fade'])
  assert.equal(Object.keys(L.LOOKS).length, 9)
  assert.ok(Object.isFrozen(L.LOOKS) && Object.isFrozen(L.LOOKS.sepia.specs))
  for (const id of L.LOOK_IDS) assert.ok(L.LOOKS[id].label.length > 0)
  const src = randomImage(40, 30, 10, false)
  assert.ok(sameBytes(L.applyLook(src, 'none', 100), src))
  assert.ok(sameBytes(L.applyLook(src, 'vivid', 0), src))
  for (const id of L.LOOK_IDS) {
    const out = L.applyLook(src, id, 100)
    for (let i = 3; i < out.data.length; i += 4) assert.equal(out.data[i], src.data[i], `${id} keeps alpha`)
  }
  const mono = L.applyLook(src, 'mono', 100)
  for (let i = 0; i < mono.data.length; i += 4) {
    assert.equal(mono.data[i], mono.data[i + 1])
    assert.equal(mono.data[i], mono.data[i + 2])
  }
  const full = L.applyLook(src, 'warm', 100)
  const half = L.applyLook(src, 'warm', 50)
  for (let i = 0; i < src.data.length; i += 1) {
    if (i % 4 === 3) continue
    assert.ok(Math.abs(half.data[i] - (src.data[i] + full.data[i]) / 2) <= 1)
  }
  const gray = image(41, 41, () => [200, 200, 200, 255])
  const vignette = L.applyLook(gray, 'dramatic', 100)
  assert.ok(px(vignette, 0, 0)[0] < px(vignette, 20, 20)[0] - 30, 'corners darken')
  const centre = L.applyLook(image(41, 41, () => [200, 200, 200, 255]), 'vintage', 100)
  assert.equal(px(centre, 20, 20)[0], px(centre, 20, 19)[0], 'the centre is untouched by the vignette')
  assert.throws(() => L.applyLook(src, 'nope', 100), RangeError)
})

function runInChunks(run, rows) {
  for (let start = 0; start < run.rows; start += rows) run.process(start, Math.min(run.rows, start + rows))
  return run.output
}

test('with absolute row offsets every filter is stripe-exact, noise, mosaic and steep motion blur included', () => {
  const src = randomImage(29, 70, 15, false)
  const rowBytes = src.width * 4
  const specs = [
    { type: 'add-noise', amount: 40, distribution: 'gaussian', monochromatic: false, seed: 9 },
    { type: 'add-noise', amount: 25, distribution: 'uniform', monochromatic: true, seed: 2 },
    { type: 'pixelate', cellSize: 6 },
    { type: 'pixelate', cellSize: 11 },
    { type: 'motion-blur', angle: 70, distance: 15 },
    { type: 'motion-blur', angle: -90, distance: 8 },
    { type: 'gaussian-blur', radius: 2.5 },
  ]
  for (const spec of specs) {
    const whole = F.applyFilter(src, spec)
    const margin = F.filterMargin(spec)
    for (const [top, bottom] of [[0, 17], [17, 40], [40, 70]]) {
      const from = Math.max(0, top - margin)
      const to = Math.min(src.height, bottom + margin)
      const stripe = { width: src.width, height: to - from, data: src.data.slice(from * rowBytes, to * rowBytes) }
      const part = F.applyFilterStripe(stripe, spec, from)
      const actual = part.data.subarray((top - from) * rowBytes, (bottom - from) * rowBytes)
      const expected = whole.data.subarray(top * rowBytes, bottom * rowBytes)
      assert.equal(Buffer.compare(Buffer.from(actual), Buffer.from(expected)), 0, `${JSON.stringify(spec)} rows ${top}..${bottom}`)
    }
  }
  assert.equal(F.stripeMargin({ type: 'motion-blur', angle: 70, distance: 15 }), null, 'the worker client cannot pass row offsets')
  assert.equal(typeof F.stripeMargin({ type: 'motion-blur', angle: 20, distance: 15 }), 'number')
})

test('startFilter and startLook give the same bytes in any chunking, masks included', () => {
  const src = randomImage(33, 300, 21, false)
  const mask = { width: 33, height: 300, data: Uint8Array.from({ length: 33 * 300 }, (_, i) => (i * 37) % 256) }
  for (const type of F.FILTER_TYPES) {
    const spec = F.defaultFilter(type)
    const whole = F.applyFilter(src, spec)
    for (const rows of [1, 7, 64, 300]) {
      const chunked = runInChunks(F.startFilter(src, spec), rows)
      assert.ok(sameBytes(chunked, whole), `${type} in chunks of ${rows}`)
    }
    const run = F.startFilter(src, spec)
    assert.ok(run.chunkRows >= 1)
    const masked = runInChunks(F.startFilter(src, spec, mask), 50)
    const B = load('imaging/buffer.ts')
    assert.ok(sameBytes(masked, B.mixByMask(src, whole, mask, 1)), `${type} through a mask`)
  }
  for (const id of L.LOOK_IDS) {
    for (const intensity of [100, 35]) {
      const whole = L.applyLook(src, id, intensity)
      assert.ok(sameBytes(runInChunks(L.startLook(src, id, intensity), 13), whole), `${id} at ${intensity}%`)
    }
  }
})

test('a long filter job in the worker runtime stops soon after an abort message', async () => {
  const { createImagingWorkerRuntime, HANDLERS } = load('shared/worker-ops/index.ts')
  const width = 1000
  const height = 3000
  const next = R.mulberry32(31)
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i += 1) data[i] = next() * 256
  const src = { width, height, data }
  const spec = { type: 'find-edges' }
  const chunks = Math.ceil(height / F.startFilter(src, spec).chunkRows)
  assert.ok(chunks >= 3, 'the job spans several chunks')
  const messages = []
  let abortSent = false
  const finished = new Promise((resolve) => {
    const runtime = createImagingWorkerRuntime(HANDLERS, (message) => {
      messages.push(message)
      if (message.type === 'progress' && !abortSent) {
        abortSent = true
        // Like a real worker: the abort arrives as a separate task, so the handler must yield to see it.
        setImmediate(() => runtime.handle({ type: 'abort', id: 1 }))
      }
      if (message.type === 'result' || message.type === 'error') resolve(message)
    })
    runtime.handle({ type: 'run', id: 1, op: 'filter', input: { src, spec, mask: null } })
  })
  const final = await finished
  assert.equal(final.type, 'error')
  assert.equal(final.name, 'AbortError')
  const progress = messages.filter((message) => message.type === 'progress')
  assert.ok(progress.length < chunks, `stopped after ${progress.length} of ${chunks} chunks`)
})

test('a radius-10 Gaussian on 4 MP finishes under 600 ms', () => {
  const width = 2000
  const height = 2000
  const next = R.mulberry32(77)
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i += 1) data[i] = (i & 3) === 3 ? 255 : next() * 256
  const src = { width, height, data }
  let best = Infinity
  for (let run = 0; run < 3; run += 1) {
    const started = process.hrtime.bigint()
    F.applyFilter(src, { type: 'gaussian-blur', radius: 10 })
    best = Math.min(best, Number(process.hrtime.bigint() - started) / 1e6)
    if (best < 600) break
  }
  assert.ok(best < 600, `best of up to 3 runs: ${best.toFixed(0)} ms`)
})
