'use strict'
// WP2 transforms: arbitrary rotation (identity at 0/360, exact quarter turns, expand/same-size/inscribed
// fits, background fill), inscribed size closed form, affine/perspective warps (identity is exact),
// homography solving and inversion, and the 'rotate'/'warp' worker ops.
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

const T = load('imaging/transform.ts')
const B = load('imaging/buffer.ts')
const R = load('imaging/random.ts')
const { handlers } = load('shared/worker-ops/color.ts')

const INTERPOLATIONS = ['nearest', 'bilinear', 'bicubic', 'lanczos3']

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

function randomImage(width, height, seed = 1, opaque = false) {
  const next = R.mulberry32(seed)
  return image(width, height, () => [next() * 256, next() * 256, next() * 256, opaque ? 255 : next() * 256])
}

function px(buffer, x, y) {
  const i = (y * buffer.width + x) * 4
  return [...buffer.data.subarray(i, i + 4)]
}

function same(a, b) {
  return a.width === b.width && a.height === b.height && Buffer.compare(Buffer.from(a.data), Buffer.from(b.data)) === 0
}

function maxDiff(a, b) {
  let max = 0
  for (let i = 0; i < a.data.length; i += 1) max = Math.max(max, Math.abs(a.data[i] - b.data[i]))
  return max
}

test('rotateArbitrary at 0 and 360 degrees is the identity for every fit and interpolation', () => {
  const src = randomImage(23, 17, 4)
  for (const degrees of [0, 360, -360, 720, 1e-12]) {
    for (const fit of ['expand', 'same-size', 'crop-inscribed']) {
      for (const interpolation of INTERPOLATIONS) {
        const out = T.rotateArbitrary(src, degrees, fit, interpolation)
        assert.equal(out.width, src.width)
        assert.equal(out.height, src.height)
        assert.ok(maxDiff(out, src) <= 1, `${degrees} ${fit} ${interpolation}`)
        assert.notEqual(out.data, src.data)
      }
    }
  }
})

test('quarter turns are exact permutations and turn clockwise', () => {
  const src = randomImage(9, 5, 2)
  assert.ok(same(T.rotateArbitrary(src, 90, 'expand'), B.rotate90(src, true)))
  assert.ok(same(T.rotateArbitrary(src, -90, 'expand'), B.rotate90(src, false)))
  assert.ok(same(T.rotateArbitrary(src, 270, 'expand'), B.rotate90(src, false)))
  assert.ok(same(T.rotateArbitrary(src, 180, 'same-size'), B.rotate180(src)))
  assert.ok(same(T.rotateArbitrary(src, 90.0000000000001, 'expand'), B.rotate90(src, true)), 'snaps float noise')
  const square = randomImage(8, 8, 5)
  assert.ok(same(T.rotateArbitrary(square, 90, 'same-size'), B.rotate90(square, true)))
  // A non-square same-size quarter turn samples: the centre pixel still lands exactly.
  const odd = image(7, 5, (x, y) => [x * 30, y * 50, 0, 255])
  const turned = T.rotateArbitrary(odd, 90, 'same-size', 'nearest')
  assert.equal(turned.width, 7)
  assert.deepEqual(px(turned, 3, 2), px(odd, 3, 2))
  // Clockwise on screen: a marker right of centre moves below centre.
  const marker = image(21, 21, (x, y) => (x === 18 && y === 10 ? [255, 0, 0, 255] : [0, 0, 0, 255]))
  const cw = T.rotateArbitrary(marker, 90, 'same-size', 'nearest')
  assert.deepEqual(px(cw, 10, 18), [255, 0, 0, 255])
  const small = T.rotateArbitrary(marker, 30, 'same-size', 'bilinear')
  const angle = (30 * Math.PI) / 180
  const tx = 10.5 + 8 * Math.cos(angle) - 0.5
  const ty = 10.5 + 8 * Math.sin(angle) - 0.5
  let best = [0, 0, -1]
  for (let y = 0; y < 21; y += 1) for (let x = 0; x < 21; x += 1) if (px(small, x, y)[0] > best[2]) best = [x, y, px(small, x, y)[0]]
  assert.ok(Math.hypot(best[0] - tx, best[1] - ty) <= 1, `marker at ${best} (expected near ${tx.toFixed(1)}, ${ty.toFixed(1)})`)
})

test('expand fits the rotated bounds, corners take the background, the centre survives', () => {
  const src = image(40, 20, () => [200, 120, 40, 255])
  const out = T.rotateArbitrary(src, 30, 'expand')
  const bounds = T.rotatedBounds(40, 20, 30)
  assert.equal(out.width, bounds.width)
  assert.equal(out.height, bounds.height)
  assert.equal(bounds.width, Math.ceil(40 * Math.cos(Math.PI / 6) + 20 * Math.sin(Math.PI / 6) - 1e-6))
  assert.deepEqual(px(out, 0, 0), [0, 0, 0, 0], 'transparent corner')
  assert.deepEqual(px(out, out.width >> 1, out.height >> 1), [200, 120, 40, 255])
  for (let i = 0; i < out.data.length; i += 4) {
    if (out.data[i + 3] > 0) assert.ok(Math.abs(out.data[i] - 200) <= 1 && Math.abs(out.data[i + 2] - 40) <= 1, 'no dark fringe at the rotated edge')
  }
  const filled = T.rotateArbitrary(src, 30, 'expand', 'bicubic', { r: 255, g: 255, b: 255, a: 255 })
  assert.deepEqual(px(filled, 0, 0), [255, 255, 255, 255], 'background colour fills the corners')
})

test('inscribedSize matches the closed form and crop-inscribed leaves no empty corners', () => {
  for (const [width, height] of [[400, 300], [300, 400], [1000, 1000], [4032, 3024], [17, 91]]) {
    for (const degrees of [-44.9, -10, -0.1, 0, 0.5, 3, 10, 27.5, 45, 90, 135]) {
      const t = (degrees * Math.PI) / 180
      const c = Math.abs(Math.cos(t))
      const s = Math.abs(Math.sin(t))
      const scale = Math.min(width / (width * c + height * s), height / (width * s + height * c))
      const size = T.inscribedSize(width, height, degrees)
      assert.equal(size.width, Math.max(1, Math.floor(width * scale + 1e-6)), `${width}x${height} at ${degrees}`)
      assert.equal(size.height, Math.max(1, Math.floor(height * scale + 1e-6)))
    }
  }
  assert.deepEqual(T.inscribedSize(400, 300, 0), { width: 400, height: 300 })
  const src = randomImage(120, 80, 6, true)
  const cropped = T.rotateArbitrary(src, 10, 'crop-inscribed')
  assert.deepEqual({ width: cropped.width, height: cropped.height }, T.inscribedSize(120, 80, 10))
  for (let i = 3; i < cropped.data.length; i += 4) assert.ok(cropped.data[i] >= 254, 'no transparent corners')
})

test('opaque images stay opaque: inscribed crops everywhere, other fits inside the source', () => {
  const src = randomImage(160, 120, 19, true)
  for (const interpolation of INTERPOLATIONS) {
    for (const degrees of [10, -4.5, 33]) {
      const cropped = T.rotateArbitrary(src, degrees, 'crop-inscribed', interpolation)
      for (let i = 3; i < cropped.data.length; i += 4) assert.equal(cropped.data[i], 255, `${interpolation} ${degrees}: inscribed crop pixel ${(i - 3) / 4}`)
      const same = T.rotateArbitrary(src, degrees, 'same-size', interpolation)
      for (let y = 40; y < 80; y += 1) {
        for (let x = 50; x < 110; x += 1) assert.equal(px(same, x, y)[3], 255, `${interpolation} ${degrees}: interior (${x}, ${y})`)
      }
    }
  }
  const flat = image(64, 64, () => [200, 100, 50, 255])
  for (const interpolation of INTERPOLATIONS) {
    assert.deepEqual(px(T.rotateArbitrary(flat, 17, 'same-size', interpolation), 32, 32), [200, 100, 50, 255], `${interpolation}: weights sum to one`)
  }
})

test('rotating there and back restores the interior closely', () => {
  const src = image(60, 60, (x, y) => [128 + 100 * Math.sin(x / 6), 128 + 100 * Math.cos(y / 7), (x * y) % 256, 255])
  const there = T.rotateArbitrary(src, 17, 'same-size', 'lanczos3')
  const back = T.rotateArbitrary(there, -17, 'same-size', 'lanczos3')
  let total = 0
  let count = 0
  for (let y = 20; y < 40; y += 1) {
    for (let x = 20; x < 40; x += 1) {
      for (let c = 0; c < 2; c += 1) {
        total += Math.abs(px(back, x, y)[c] - px(src, x, y)[c])
        count += 1
      }
    }
  }
  assert.ok(total / count < 2, `mean error ${(total / count).toFixed(2)}`)
})

test('warpPerspective and warpAffine with the identity are exact for every interpolation', () => {
  const src = randomImage(31, 19, 8)
  const out = { x: 0, y: 0, width: 31, height: 19 }
  for (const interpolation of INTERPOLATIONS) {
    assert.ok(same(T.warpPerspective(src, [1, 0, 0, 0, 1, 0, 0, 0, 1], out, interpolation), src), `perspective ${interpolation}`)
    assert.ok(same(T.warpPerspective(src, [2, 0, 0, 0, 2, 0, 0, 0, 2], out, interpolation), src), 'a scaled identity is the identity')
    assert.ok(same(T.warpAffine(src, [1, 0, 0, 1, 0, 0], out, interpolation), src), `affine ${interpolation}`)
  }
  // Integer translation: destination (x, y) reads source (x + 3, y - 2).
  const moved = T.warpAffine(src, [1, 0, 0, 1, 3, -2], { x: 0, y: 0, width: 31, height: 19 }, 'bicubic')
  assert.deepEqual(px(moved, 4, 5), px(src, 7, 3))
  assert.deepEqual(px(moved, 30, 0), [0, 0, 0, 0], 'outside the source is transparent')
  // The output rectangle is in destination space.
  const window = T.warpPerspective(src, [1, 0, 0, 0, 1, 0, 0, 0, 1], { x: 5, y: 4, width: 6, height: 3 })
  assert.deepEqual(px(window, 0, 0), px(src, 5, 4))
  assert.deepEqual(px(window, 5, 2), px(src, 10, 6))
})

test('homographyFromQuads maps the corners; inversion and application agree', () => {
  const from = [{ x: 0, y: 0 }, { x: 100, y: 0 }, { x: 100, y: 80 }, { x: 0, y: 80 }]
  const to = [{ x: 12, y: 7 }, { x: 140, y: 20 }, { x: 120, y: 130 }, { x: -5, y: 90 }]
  const h = T.homographyFromQuads(from, to)
  assert.equal(h[8], 1)
  for (let k = 0; k < 4; k += 1) {
    const p = T.applyHomography(h, from[k])
    assert.ok(Math.abs(p.x - to[k].x) < 1e-6 && Math.abs(p.y - to[k].y) < 1e-6, `corner ${k}`)
  }
  const inverse = T.invertHomography(h)
  for (const p of [{ x: 50, y: 40 }, { x: 3, y: 77 }, { x: 99, y: 1 }]) {
    const q = T.applyHomography(inverse, T.applyHomography(h, p))
    assert.ok(Math.abs(q.x - p.x) < 1e-6 && Math.abs(q.y - p.y) < 1e-6)
  }
  const big = T.homographyFromQuads(
    [{ x: 0, y: 0 }, { x: 20000, y: 0 }, { x: 20000, y: 15000 }, { x: 0, y: 15000 }],
    [{ x: 100, y: 50 }, { x: 19000, y: 300 }, { x: 20500, y: 14000 }, { x: -400, y: 15500 }],
  )
  const corner = T.applyHomography(big, { x: 20000, y: 15000 })
  assert.ok(Math.abs(corner.x - 20500) < 1e-4 && Math.abs(corner.y - 14000) < 1e-4, 'large coordinates stay accurate')
  const affine = [1.5, 0.2, -0.3, 0.9, 10, -4]
  const asHomography = T.affineToHomography(affine)
  const viaH = T.applyHomography(asHomography, { x: 3, y: 5 })
  assert.ok(Math.abs(viaH.x - (1.5 * 3 - 0.3 * 5 + 10)) < 1e-12 && Math.abs(viaH.y - (0.2 * 3 + 0.9 * 5 - 4)) < 1e-12)
  const roundTrip = T.invertAffine(T.invertAffine(affine))
  roundTrip.forEach((value, index) => assert.ok(Math.abs(value - affine[index]) < 1e-12))
  assert.throws(() => T.homographyFromQuads(from, [{ x: 0, y: 0 }, { x: 1, y: 1 }, { x: 2, y: 2 }, { x: 3, y: 3 }]), RangeError)
  assert.throws(() => T.homographyFromQuads(from, to.slice(0, 3)), RangeError)
  assert.throws(() => T.invertHomography([1, 2, 3, 2, 4, 6, 0, 0, 0]), RangeError)
  assert.throws(() => T.invertAffine([1, 2, 2, 4, 0, 0]), RangeError)
})

test('a perspective warp built from quads moves a marker where the quads say', () => {
  const src = image(40, 40, (x, y) => (x >= 18 && x <= 21 && y >= 18 && y <= 21 ? [255, 255, 255, 255] : [0, 0, 0, 255]))
  const from = [{ x: 0, y: 0 }, { x: 40, y: 0 }, { x: 40, y: 40 }, { x: 0, y: 40 }]
  const to = [{ x: 10, y: 5 }, { x: 70, y: 0 }, { x: 60, y: 60 }, { x: 0, y: 50 }]
  const forward = T.homographyFromQuads(from, to)
  const out = T.warpPerspective(src, T.invertHomography(forward), { x: 0, y: 0, width: 72, height: 62 }, 'bilinear')
  const centre = T.applyHomography(forward, { x: 20, y: 20 })
  assert.ok(px(out, Math.floor(centre.x), Math.floor(centre.y))[0] > 200, 'the marker lands at the mapped centre')
  assert.equal(px(out, 71, 61)[3], 0, 'outside the warped quad is transparent')
})

test('rotate and warp runs give the same bytes in any chunking', () => {
  const src = randomImage(45, 31, 13)
  const runs = [
    () => T.startRotate(src, 23, 'expand', 'bicubic'),
    () => T.startRotate(src, -8, 'crop-inscribed', 'lanczos3'),
    () => T.startRotate(src, 90, 'same-size', 'bilinear'),
    () => T.startRotate(src, 180, 'expand'),
    () => T.startWarpAffine(src, [0.8, 0.1, -0.2, 1.1, 3, -2], { x: -4, y: 2, width: 50, height: 40 }, 'bicubic'),
    () => T.startWarpPerspective(src, [1, 0.05, 0, 0.02, 1, 0, 0.001, 0.0005, 1], { x: 0, y: 0, width: 47, height: 33 }, 'bilinear'),
  ]
  for (const make of runs) {
    const whole = make()
    for (let start = 0; start < whole.rows; start += whole.rows) whole.process(start, whole.rows)
    for (const rows of [1, 6]) {
      const run = make()
      for (let start = 0; start < run.rows; start += rows) run.process(start, Math.min(run.rows, start + rows))
      assert.ok(same(run.output, whole.output), `chunks of ${rows}`)
    }
  }
  assert.ok(same(T.rotateArbitrary(src, 23, 'expand', 'bicubic'), (() => { const r = runs[0](); r.process(0, r.rows); return r.output })()))
})

test('rotate and warp worker ops wrap the transforms and leave inputs intact', async () => {
  const src = randomImage(24, 16, 3)
  const before = new Uint8ClampedArray(src.data)
  const rotated = await handlers.rotate({ src, degrees: 12, fit: 'expand', interpolation: 'bicubic', background: { r: 0, g: 0, b: 0, a: 0 } }, {})
  assert.ok(same(rotated, T.rotateArbitrary(src, 12, 'expand', 'bicubic')))
  const warped = await handlers.warp({ src, inverse: [1, 0, 0, 0, 1, 0, 0, 0, 1], out: { x: 0, y: 0, width: 24, height: 16 }, interpolation: 'bilinear' }, {})
  assert.ok(same(warped, src))
  assert.deepEqual(src.data, before)
  const controller = new AbortController()
  controller.abort()
  assert.throws(() => handlers.rotate({ src, degrees: 12, fit: 'expand', interpolation: 'bicubic', background: { r: 0, g: 0, b: 0, a: 0 } }, { signal: controller.signal }), { name: 'AbortError' })
  assert.throws(() => T.rotateArbitrary(src, Number.NaN, 'expand'), RangeError)
  assert.throws(() => T.warpPerspective(src, [1, 0, 0], { x: 0, y: 0, width: 2, height: 2 }), RangeError)
})
