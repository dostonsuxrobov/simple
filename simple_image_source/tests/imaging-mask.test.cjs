'use strict'
// WP4 selection masks (src/imaging/mask.ts): anti-aliased shape coverage against analytic areas, combine
// truth tables, invert, feather mass preservation, exact Euclidean expand / contract, bounds, the
// pixel-edge outline, and the matching worker-op handlers.
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

const mask = load('imaging/mask.ts')
const { handlers } = load('shared/worker-ops/mask.ts')

const coverage = (m) => m.data.reduce((sum, value) => sum + value, 0) / 255
const partials = (m) => m.data.reduce((count, value) => count + (value > 0 && value < 255 ? 1 : 0), 0)
const isBinary = (m) => m.data.every((value) => value === 0 || value === 255)

function random(seed) {
  let state = seed >>> 0
  return () => {
    state = (Math.imul(state, 1664525) + 1013904223) >>> 0
    return state / 4294967296
  }
}

function randomMask(width, height, seed) {
  const rnd = random(seed)
  const out = mask.createMaskBuffer(width, height)
  for (let i = 0; i < out.data.length; i += 1) {
    const r = rnd()
    out.data[i] = r < 0.3 ? 0 : r < 0.6 ? 255 : Math.floor(rnd() * 256)
  }
  return out
}

function polygonArea(points) {
  let sum = 0
  for (let i = 0; i < points.length; i += 1) {
    const a = points[i]
    const b = points[(i + 1) % points.length]
    sum += a.x * b.y - b.x * a.y
  }
  return Math.abs(sum) / 2
}

test('createMaskBuffer validates its size and fills', () => {
  const m = mask.createMaskBuffer(3, 2, 200)
  assert.equal(m.width, 3)
  assert.equal(m.height, 2)
  assert.ok(m.data instanceof Uint8Array)
  assert.deepEqual([...m.data], [200, 200, 200, 200, 200, 200])
  assert.deepEqual([...mask.createMaskBuffer(2, 1, 999).data], [255, 255])
  assert.throws(() => mask.createMaskBuffer(-1, 2), RangeError)
  assert.throws(() => mask.createMaskBuffer(2, 1.5), RangeError)
})

test('anti-aliased rectangles have exact area coverage with partial edges', () => {
  for (const rect of [{ x: 10.3, y: 20.7, width: 30.25, height: 15.5 }, { x: 5, y: 6, width: 40, height: 9 }, { x: 0.5, y: 0.5, width: 0.5, height: 0.5 }]) {
    const m = mask.createMaskBuffer(100, 60)
    const bounds = mask.rasterizeRect(m, rect, true)
    const area = rect.width * rect.height
    assert.ok(Math.abs(coverage(m) - area) <= Math.max(0.005 * area, 0.01), `rect area ${coverage(m)} vs ${area}`)
    assert.deepEqual(bounds, mask.maskBounds(m))
  }
  const m = mask.createMaskBuffer(100, 60)
  mask.rasterizeRect(m, { x: 10.3, y: 20.7, width: 30.25, height: 15.5 }, true)
  assert.ok(partials(m) > 0, 'fractional edges are partially covered')
  // Interior pixels are fully covered; the left column [10, 11) is covered 0.7 horizontally.
  assert.equal(m.data[30 * 100 + 20], 255)
  assert.ok(Math.abs(m.data[30 * 100 + 10] - 0.7 * 255) <= 0.5)
  // Integer rectangles are crisp either way.
  const crisp = mask.createMaskBuffer(50, 50)
  mask.rasterizeRect(crisp, { x: 5, y: 6, width: 10, height: 7 }, true)
  assert.ok(isBinary(crisp))
  assert.equal(coverage(crisp), 70)
})

test('aliased rectangles include pixels whose centre is inside; shapes clip and normalise', () => {
  const m = mask.createMaskBuffer(40, 40)
  const bounds = mask.rasterizeRect(m, { x: 2.4, y: 3.6, width: 5.2, height: 4 }, false)
  // Centres 2.5 .. 7.5 are inside [2.4, 7.6); rows: centres 4.5 .. 7.5 inside [3.6, 7.6).
  assert.deepEqual(bounds, { x: 2, y: 4, width: 6, height: 4 })
  assert.ok(isBinary(m))
  assert.equal(coverage(m), 24)
  // Negative sizes are normalised; outside or empty shapes return null.
  const n = mask.createMaskBuffer(40, 40)
  assert.deepEqual(mask.rasterizeRect(n, { x: 12, y: 12, width: -4, height: -4 }, false), { x: 8, y: 8, width: 4, height: 4 })
  assert.equal(mask.rasterizeRect(n, { x: 50, y: 0, width: 4, height: 4 }, true), null)
  assert.equal(mask.rasterizeRect(n, { x: 5, y: 5, width: 0, height: 4 }, true), null)
  // Clipped to the target.
  const clipped = mask.createMaskBuffer(10, 10)
  assert.deepEqual(mask.rasterizeRect(clipped, { x: -5, y: 8, width: 8, height: 10 }, true), { x: 0, y: 8, width: 3, height: 2 })
})

test('rasterizers combine with max so several shapes can share one scratch mask', () => {
  const m = mask.createMaskBuffer(30, 30)
  m.data.fill(100)
  mask.rasterizeRect(m, { x: 10.5, y: 10, width: 5, height: 5 }, true)
  assert.equal(m.data[12 * 30 + 12], 255)
  assert.equal(m.data[12 * 30 + 10], Math.max(100, Math.round(0.5 * 255)))
  assert.equal(m.data[0], 100, 'pixels outside the shape keep their value')
})

test('anti-aliased ellipses are within 0.5% of the analytic area and have soft edges', () => {
  for (const rect of [{ x: 20.3, y: 30.1, width: 120.5, height: 80.25 }, { x: 3.2, y: 4.7, width: 12, height: 7.5 }, { x: 50, y: 10, width: 30, height: 30 }]) {
    const m = mask.createMaskBuffer(180, 140)
    const bounds = mask.rasterizeEllipse(m, rect, true)
    const area = Math.PI * (rect.width / 2) * (rect.height / 2)
    assert.ok(Math.abs(coverage(m) - area) <= 0.005 * area, `ellipse ${coverage(m)} vs ${area}`)
    assert.ok(partials(m) > 0)
    assert.deepEqual(bounds, mask.maskBounds(m))
  }
  const aliased = mask.createMaskBuffer(180, 140)
  mask.rasterizeEllipse(aliased, { x: 20.3, y: 30.1, width: 120.5, height: 80.25 }, false)
  assert.ok(isBinary(aliased))
  const area = Math.PI * 60.25 * 40.125
  assert.ok(Math.abs(coverage(aliased) - area) <= 0.01 * area)
  // Centre-inclusion: the centre of the box is selected, the corners of the box are not.
  assert.equal(aliased.data[70 * 180 + 80], 255)
  assert.equal(aliased.data[31 * 180 + 21], 0)
})

test('anti-aliased polygons are within 0.5% of the analytic area (convex and concave)', () => {
  const shapes = [
    [{ x: 10.5, y: 10.2 }, { x: 180.3, y: 40.7 }, { x: 60.1, y: 170.9 }],
    [{ x: 20, y: 20 }, { x: 170, y: 25.5 }, { x: 90.25, y: 80 }, { x: 160, y: 170 }, { x: 30.5, y: 150 }],
    // A lasso-like star outline with many points.
    Array.from({ length: 200 }, (_, i) => {
      const angle = (i / 200) * Math.PI * 2
      const r = 60 + 20 * Math.sin(angle * 7)
      return { x: 100 + r * Math.cos(angle), y: 95 + r * Math.sin(angle) }
    }),
  ]
  for (const points of shapes) {
    const m = mask.createMaskBuffer(200, 200)
    const bounds = mask.rasterizePolygon(m, points, true)
    const area = polygonArea(points)
    assert.ok(Math.abs(coverage(m) - area) <= 0.005 * area, `polygon ${coverage(m)} vs ${area}`)
    assert.ok(partials(m) > 0)
    assert.deepEqual(bounds, mask.maskBounds(m))
  }
  const aliased = mask.createMaskBuffer(200, 200)
  mask.rasterizePolygon(aliased, shapes[0], false)
  assert.ok(isBinary(aliased))
  assert.ok(Math.abs(coverage(aliased) - polygonArea(shapes[0])) <= 0.01 * polygonArea(shapes[0]))
})

test('polygon fill rules: a pentagram is solid with nonzero and hollow with evenodd', () => {
  const star = Array.from({ length: 5 }, (_, i) => {
    const angle = -Math.PI / 2 + (i * 4 * Math.PI) / 5
    return { x: 50 + 40 * Math.cos(angle), y: 50 + 40 * Math.sin(angle) }
  })
  const nonzero = mask.createMaskBuffer(100, 100)
  const evenodd = mask.createMaskBuffer(100, 100)
  mask.rasterizePolygon(nonzero, star, true, 'nonzero')
  mask.rasterizePolygon(evenodd, star, true, 'evenodd')
  assert.equal(nonzero.data[52 * 100 + 50], 255, 'the centre pentagon is inside with nonzero')
  assert.equal(evenodd.data[52 * 100 + 50], 0, 'and outside with evenodd')
  // Inner pentagon area: the star's points are spaced so the hole is substantial.
  assert.ok(coverage(nonzero) - coverage(evenodd) > 500)
  // Degenerate input.
  const m = mask.createMaskBuffer(10, 10)
  assert.equal(mask.rasterizePolygon(m, [{ x: 1, y: 1 }, { x: 5, y: 5 }], true), null)
  assert.equal(mask.rasterizePolygon(m, [{ x: 1, y: 1 }, { x: 8, y: 1 }, { x: 4, y: 1 }], true), null)
  assert.notEqual(mask.rasterizePolygon(m, [{ x: 1, y: 1 }, { x: NaN, y: 3 }, { x: 8, y: 1 }, { x: 4, y: 8 }], true), null)
})

test('combine operations follow their truth tables', () => {
  const a = { width: 4, height: 1, data: Uint8Array.from([0, 0, 255, 255]) }
  const b = { width: 4, height: 1, data: Uint8Array.from([0, 255, 0, 255]) }
  const run = (op) => {
    const dst = { width: 4, height: 1, data: Uint8Array.from(a.data) }
    mask.combineMasks(dst, b, op)
    return [...dst.data]
  }
  assert.deepEqual(run('replace'), [0, 255, 0, 255])
  assert.deepEqual(run('add'), [0, 255, 255, 255])
  assert.deepEqual(run('subtract'), [0, 0, 255, 0])
  assert.deepEqual(run('intersect'), [0, 0, 0, 255])
  assert.throws(() => mask.combineMasks(mask.createMaskBuffer(2, 2), mask.createMaskBuffer(2, 3), 'add'), RangeError)
  assert.throws(() => mask.combineMasks(mask.createMaskBuffer(2, 2), mask.createMaskBuffer(2, 2), 'xor'), RangeError)
})

test('partial-coverage combinations round exactly (every 8-bit pair)', () => {
  const dst = mask.createMaskBuffer(256, 256)
  const src = mask.createMaskBuffer(256, 256)
  for (let i = 0; i < 65536; i += 1) {
    dst.data[i] = i >> 8
    src.data[i] = i & 255
  }
  const subtract = { width: 256, height: 256, data: Uint8Array.from(dst.data) }
  const intersect = { width: 256, height: 256, data: Uint8Array.from(dst.data) }
  const add = { width: 256, height: 256, data: Uint8Array.from(dst.data) }
  mask.combineMasks(subtract, src, 'subtract')
  mask.combineMasks(intersect, src, 'intersect')
  mask.combineMasks(add, src, 'add')
  for (let i = 0; i < 65536; i += 1) {
    const x = i >> 8
    const y = i & 255
    if (subtract.data[i] !== Math.round((x * (255 - y)) / 255)) assert.fail(`subtract ${x} ${y}`)
    if (intersect.data[i] !== Math.round((x * y) / 255)) assert.fail(`intersect ${x} ${y}`)
    if (add.data[i] !== Math.max(x, y)) assert.fail(`add ${x} ${y}`)
  }
})

test('the rect hint never changes the result of a combine', () => {
  for (const op of ['replace', 'add', 'subtract', 'intersect']) {
    const base = randomMask(37, 23, 5)
    const shape = mask.createMaskBuffer(37, 23)
    mask.rasterizeEllipse(shape, { x: 8.5, y: 4.2, width: 17, height: 11 }, true)
    const hint = mask.maskBounds(shape)
    const full = { width: 37, height: 23, data: Uint8Array.from(base.data) }
    const fast = { width: 37, height: 23, data: Uint8Array.from(base.data) }
    mask.combineMasks(full, shape, op)
    mask.combineMasks(fast, shape, op, hint)
    assert.deepEqual(fast.data, full.data, op)
  }
})

test('invert twice is identity', () => {
  const m = randomMask(31, 17, 9)
  const before = Uint8Array.from(m.data)
  mask.invertMask(m)
  assert.equal(m.data[0], 255 - before[0])
  mask.invertMask(m)
  assert.deepEqual(m.data, before)
})

test('feathering preserves total coverage within 1% and spreads the edge', () => {
  const m = mask.createMaskBuffer(300, 260)
  mask.rasterizeEllipse(m, { x: 90, y: 80, width: 120, height: 90 }, true)
  const total = coverage(m)
  for (const radius of [0.6, 2, 3.9, 4, 12, 30]) {
    const feathered = mask.featherMask(m, radius)
    assert.notEqual(feathered.data, m.data, 'returns a new mask')
    assert.ok(Math.abs(coverage(feathered) - total) <= 0.01 * total, `radius ${radius}: ${coverage(feathered)} vs ${total}`)
    assert.ok(partials(feathered) > partials(m), `radius ${radius} softens the edge`)
  }
  assert.ok(Math.abs(coverage(mask.featherMask(m, 30)) - total) / total < 0.002)
  // The canvas edge is not a selection edge: Select All stays fully selected.
  const all = mask.featherMask(mask.createMaskBuffer(64, 48, 255), 25)
  assert.ok(all.data.every((value) => value === 255))
  // Radius 0 and an empty mask are copies.
  assert.deepEqual(mask.featherMask(m, 0).data, m.data)
  assert.ok(mask.featherMask(mask.createMaskBuffer(20, 20), 5).data.every((value) => value === 0))
})

test('expand and contract of a square are exact with Euclidean corners', () => {
  const size = 120
  const side = 40
  const n = 10
  const square = mask.createMaskBuffer(size, size)
  mask.rasterizeRect(square, { x: 40, y: 40, width: side, height: side }, false)
  const grown = mask.expandMask(square, n)
  // Straight edges move by exactly n pixels.
  assert.deepEqual(mask.maskBounds(grown), { x: 30, y: 30, width: 60, height: 60 })
  for (let y = 40; y < 80; y += 1) {
    assert.equal(grown.data[y * size + 30], 255)
    assert.equal(grown.data[y * size + 29], 0)
    assert.equal(grown.data[y * size + 89], 255)
    assert.equal(grown.data[y * size + 90], 0)
  }
  // Area = Minkowski sum with a disk of radius n (Steiner): L^2 + 4 L n + pi n^2.
  const steiner = side * side + 4 * side * n + Math.PI * n * n
  assert.ok(Math.abs(coverage(grown) - steiner) <= 0.005 * steiner, `${coverage(grown)} vs ${steiner}`)
  // Corner pixels follow the circle of radius n around the true corner (40, 40).
  let worst = 0
  for (let y = 26; y < 40; y += 1) {
    for (let x = 26; x < 40; x += 1) {
      let inside = 0
      for (let j = 0; j < 16; j += 1) {
        for (let i = 0; i < 16; i += 1) {
          const px = x + (i + 0.5) / 16
          const py = y + (j + 0.5) / 16
          if ((px - 40) ** 2 + (py - 40) ** 2 <= n * n) inside += 1
        }
      }
      worst = Math.max(worst, Math.abs(grown.data[y * size + x] / 255 - inside / 256))
    }
  }
  assert.ok(worst < 0.1, `corner coverage error ${worst}`)
  // Contracting a convex square gives the inner square with sharp corners: exact and crisp.
  const shrunk = mask.contractMask(square, n)
  assert.deepEqual(mask.maskBounds(shrunk), { x: 50, y: 50, width: side - 2 * n, height: side - 2 * n })
  assert.ok(isBinary(shrunk))
  assert.equal(coverage(shrunk), (side - 2 * n) ** 2)
  // Contracting the expanded square by the same amount (a closing) restores the square; only the four
  // corner pixels soften, and they stay more than 50% selected.
  const closed = mask.contractMask(grown, n)
  assert.deepEqual(mask.maskBounds(closed, 128), { x: 40, y: 40, width: side, height: side })
  const changed = []
  for (let i = 0; i < closed.data.length; i += 1) if (closed.data[i] !== square.data[i]) changed.push(i)
  assert.ok(changed.length <= 12, `${changed.length} pixels changed`)
  for (const i of changed) {
    const x = i % size
    const y = (i / size) | 0
    assert.ok(Math.min(x - 40, 79 - x) <= 1 && Math.min(y - 40, 79 - y) <= 1, `pixel (${x}, ${y}) is at a corner`)
    assert.ok(closed.data[i] >= 128)
  }
})

test('expand never shrinks, contract never grows; canvas edges and empty masks', () => {
  const m = randomMask(40, 30, 3)
  const grown = mask.expandMask(m, 2)
  const shrunk = mask.contractMask(m, 2)
  for (let i = 0; i < m.data.length; i += 1) {
    assert.ok(grown.data[i] >= m.data[i])
    assert.ok(shrunk.data[i] <= m.data[i])
  }
  assert.ok(mask.contractMask(mask.createMaskBuffer(30, 20, 255), 5).data.every((value) => value === 255), 'Select All does not contract from the canvas edge')
  assert.ok(mask.expandMask(mask.createMaskBuffer(30, 20), 5).data.every((value) => value === 0))
  assert.deepEqual(mask.expandMask(m, 0).data, m.data)
})

test('maskBounds honours the threshold and an optional region; maskFromAlpha copies alpha', () => {
  const m = mask.createMaskBuffer(10, 8)
  m.data[2 * 10 + 3] = 40
  m.data[5 * 10 + 7] = 200
  assert.deepEqual(mask.maskBounds(m), { x: 3, y: 2, width: 5, height: 4 })
  assert.deepEqual(mask.maskBounds(m, 128), { x: 7, y: 5, width: 1, height: 1 })
  assert.equal(mask.maskBounds(m, 201), null)
  assert.deepEqual(mask.maskBounds(m, 1, { x: 0, y: 0, width: 5, height: 5 }), { x: 3, y: 2, width: 1, height: 1 })
  assert.equal(mask.maskBounds(mask.createMaskBuffer(5, 5)), null)
  const pixels = { width: 2, height: 2, data: Uint8ClampedArray.from([1, 2, 3, 0, 4, 5, 6, 128, 7, 8, 9, 255, 0, 0, 0, 7]) }
  assert.deepEqual([...mask.maskFromAlpha(pixels).data], [0, 128, 255, 7])
})

test('the outline of a rectangle is four edges along pixel boundaries', () => {
  const m = mask.createMaskBuffer(120, 120)
  mask.rasterizeRect(m, { x: 40, y: 30, width: 40, height: 25 }, false)
  const outline = mask.traceOutline(m)
  assert.ok(outline instanceof Float32Array)
  assert.equal(outline.length / 4, 4)
  const edges = []
  for (let i = 0; i < outline.length; i += 4) edges.push(Array.from(outline.slice(i, i + 4)).join(','))
  assert.deepEqual(edges.sort(), ['40,30,40,55', '40,30,80,30', '40,55,80,55', '80,30,80,55'].sort())
  // Scale maps a downsampled mask back to document units.
  const scaled = mask.traceOutline(m, 128, 2)
  assert.equal(scaled.length, 16)
  assert.ok(Array.from(scaled).every((value, i) => value === outline[i] * 2))
  // Two rectangles -> eight edges; empty -> none; full -> the canvas border.
  mask.rasterizeRect(m, { x: 5, y: 80, width: 10, height: 10 }, false)
  assert.equal(mask.traceOutline(m).length / 4, 8)
  assert.equal(mask.traceOutline(mask.createMaskBuffer(10, 10)).length, 0)
  assert.equal(mask.traceOutline(mask.createMaskBuffer(10, 6, 255)).length / 4, 4)
  // The 50% rule: 127 is outside, 128 inside.
  const half = mask.createMaskBuffer(4, 1)
  half.data.set([127, 128, 128, 0])
  assert.deepEqual(Array.from(mask.traceOutline(half)).join(','), '1,0,3,0,1,1,3,1,1,0,1,1,3,0,3,1')
})

test('crop, paste and downsample helpers', () => {
  const m = mask.createMaskBuffer(6, 4)
  for (let i = 0; i < m.data.length; i += 1) m.data[i] = i * 10
  const crop = mask.cropMask(m, { x: 4, y: 2, width: 4, height: 3 })
  assert.deepEqual([...crop.data], [160, 170, 0, 0, 220, 230, 0, 0, 0, 0, 0, 0])
  const target = mask.createMaskBuffer(3, 3)
  mask.pasteMask(target, crop, -1, 1)
  assert.deepEqual([...target.data], [0, 0, 0, 170, 0, 0, 230, 0, 0])
  const down = mask.downsampleMask({ width: 3, height: 2, data: Uint8Array.from([0, 255, 100, 255, 255, 50]) }, 2)
  assert.deepEqual({ width: down.width, height: down.height, data: [...down.data] }, { width: 2, height: 1, data: [191, 75] })
})

test('worker handlers feather, expand and contract like the direct calls', async () => {
  const m = mask.createMaskBuffer(80, 60)
  mask.rasterizeEllipse(m, { x: 20, y: 15, width: 40, height: 30 }, true)
  const copy = () => ({ width: 80, height: 60, data: Uint8Array.from(m.data) })
  const options = {}
  assert.deepEqual((await handlers.feather({ mask: copy(), radius: 6 }, options)).data, mask.featherMask(m, 6).data)
  assert.deepEqual((await handlers.expand({ mask: copy(), pixels: 4 }, options)).data, mask.expandMask(m, 4).data)
  assert.deepEqual((await handlers.contract({ mask: copy(), pixels: 4 }, options)).data, mask.contractMask(m, 4).data)
  const controller = new AbortController()
  controller.abort()
  await assert.rejects(handlers.feather({ mask: copy(), radius: 6 }, { signal: controller.signal }), { name: 'AbortError' })
  // An abort that arrives while the job waits for its first turn (a queued worker message) stops it.
  const queued = new AbortController()
  const pending = handlers.expand({ mask: copy(), pixels: 4 }, { signal: queued.signal })
  queued.abort()
  await assert.rejects(pending, { name: 'AbortError' })
  await assert.rejects(handlers.expand({ mask: copy(), pixels: -1 }, options), RangeError)
  await assert.rejects(handlers.contract({ mask: { width: 4, height: 4, data: new Uint8Array(3) }, pixels: 1 }, options), TypeError)
})
