'use strict'
// WP2 blending: the 16 W3C modes match the Compositing and Blending Level 1 formulas and the 11 extra
// Photoshop modes match the design's Appendix B, on a 9 x 9 grid of sample values within +-1, through the
// W3C general compositing formula with partial alphas. Opacity 0 is a no-op, coverage scales linearly,
// Dissolve is deterministic. The reference below is written independently of src/imaging/blend.ts.
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

const Bl = load('imaging/blend.ts')
const R = load('imaging/random.ts')
const { BLEND_MODE_MENU } = load('advanced/types.ts')

// ---------------------------------------------------------------------------------------------
// Reference formulas (W3C Compositing and Blending Level 1, section 10; Appendix B of the design)
// ---------------------------------------------------------------------------------------------

const W3C_SEPARABLE = {
  normal: (b, s) => s,
  multiply: (b, s) => b * s,
  screen: (b, s) => b + s - b * s,
  overlay: (b, s) => W3C_SEPARABLE['hard-light'](s, b),
  darken: (b, s) => Math.min(b, s),
  lighten: (b, s) => Math.max(b, s),
  'color-dodge': (b, s) => (b === 0 ? 0 : s === 1 ? 1 : Math.min(1, b / (1 - s))),
  'color-burn': (b, s) => (b === 1 ? 1 : s === 0 ? 0 : 1 - Math.min(1, (1 - b) / s)),
  'hard-light': (b, s) => (s <= 0.5 ? b * 2 * s : W3C_SEPARABLE.screen(b, 2 * s - 1)),
  // Appendix B: Soft Light follows Photoshop, not the W3C D(Cb) curve.
  'soft-light': (b, s) => (s <= 0.5 ? 2 * b * s + b * b * (1 - 2 * s) : 2 * b * (1 - s) + Math.sqrt(b) * (2 * s - 1)),
  difference: (b, s) => Math.abs(b - s),
  exclusion: (b, s) => b + s - 2 * b * s,
}

const EXTRA_SEPARABLE = {
  'linear-burn': (b, s) => Math.max(0, b + s - 1),
  'linear-dodge': (b, s) => Math.min(1, b + s),
  'vivid-light': (b, s) => (s <= 0.5 ? W3C_SEPARABLE['color-burn'](b, 2 * s) : W3C_SEPARABLE['color-dodge'](b, 2 * (s - 0.5))),
  'linear-light': (b, s) => Math.min(1, Math.max(0, b + 2 * s - 1)),
  'pin-light': (b, s) => (s <= 0.5 ? Math.min(b, 2 * s) : Math.max(b, 2 * (s - 0.5))),
  'hard-mix': (b, s) => (Math.round((b + s) * 255) >= 255 ? 1 : 0),
  subtract: (b, s) => Math.max(0, b - s),
  divide: (b, s) => (s === 0 ? 1 : Math.min(1, b / s)),
}

const Lum = ([r, g, b]) => 0.3 * r + 0.59 * g + 0.11 * b
function ClipColor(c) {
  const l = Lum(c)
  const n = Math.min(...c)
  const x = Math.max(...c)
  let out = c
  if (n < 0) out = out.map((v) => l + ((v - l) * l) / (l - n))
  if (x > 1) out = out.map((v) => l + ((v - l) * (1 - l)) / (x - l))
  return out
}
const SetLum = (c, l) => ClipColor(c.map((v) => v + (l - Lum(c))))
const Sat = (c) => Math.max(...c) - Math.min(...c)
function SetSat(c, s) {
  const order = [0, 1, 2].sort((i, j) => c[i] - c[j])
  const [min, mid, max] = order
  const out = [0, 0, 0]
  if (c[max] > c[min]) {
    out[mid] = ((c[mid] - c[min]) * s) / (c[max] - c[min])
    out[max] = s
  }
  return out
}

const W3C_NON_SEPARABLE = {
  hue: (b, s) => SetLum(SetSat(s, Sat(b)), Lum(b)),
  saturation: (b, s) => SetLum(SetSat(b, Sat(s)), Lum(b)),
  color: (b, s) => SetLum(s, Lum(b)),
  luminosity: (b, s) => SetLum(b, Lum(s)),
}

const sum3 = (c) => c[0] + c[1] + c[2]
const EXTRA_NON_SEPARABLE = {
  'darker-color': (b, s) => (sum3(s) < sum3(b) ? s : b),
  'lighter-color': (b, s) => (sum3(s) > sum3(b) ? s : b),
}

function referenceBlend(mode, backdrop, source) {
  if (W3C_NON_SEPARABLE[mode]) return W3C_NON_SEPARABLE[mode](backdrop, source)
  if (EXTRA_NON_SEPARABLE[mode]) return EXTRA_NON_SEPARABLE[mode](backdrop, source)
  const f = W3C_SEPARABLE[mode] ?? EXTRA_SEPARABLE[mode]
  return backdrop.map((b, i) => f(b, source[i]))
}

/** W3C general formula; colours in [0, 1], alphas in [0, 1]. Returns RGBA in 0..255 floats. */
function referenceComposite(mode, cb, ab, cs, as) {
  const ao = as + ab * (1 - as)
  if (ao === 0) return [0, 0, 0, 0]
  const blended = referenceBlend(mode, cb, cs)
  const co = cb.map((b, i) => (as * (1 - ab) * cs[i] + as * ab * blended[i] + (1 - as) * ab * b) / ao)
  return [...co.map((v) => v * 255), ao * 255]
}

const GRID = [0, 32, 64, 96, 128, 160, 192, 224, 255]
const ALPHAS = [[255, 255, 1], [128, 255, 1], [255, 100, 1], [77, 180, 1], [200, 0, 1], [255, 255, 0.6], [180, 90, 0.35]]

function checkMode(mode) {
  let worst = 0
  for (let i = 0; i < 9; i += 1) {
    for (let j = 0; j < 9; j += 1) {
      const backdrop = [GRID[i], GRID[(i + 3) % 9], GRID[(i + 6) % 9]]
      const source = [GRID[j], GRID[(j + 5) % 9], GRID[(j + 2) % 9]]
      for (const [sourceAlpha, backdropAlpha, opacity] of ALPHAS) {
        const dst = Uint8ClampedArray.from([...backdrop, backdropAlpha])
        const src = Uint8ClampedArray.from([...source, sourceAlpha])
        Bl.blendInto(dst, 0, src, 0, 1, mode, opacity)
        const expected = referenceComposite(mode, backdrop.map((v) => v / 255), backdropAlpha / 255, source.map((v) => v / 255), (sourceAlpha / 255) * opacity)
        for (let c = 0; c < 4; c += 1) {
          const diff = Math.abs(dst[c] - expected[c])
          worst = Math.max(worst, diff)
          assert.ok(diff <= 1, `${mode}: backdrop ${backdrop}/${backdropAlpha} source ${source}/${sourceAlpha} opacity ${opacity}: got ${[...dst]} expected ${expected.map((v) => v.toFixed(2))}`)
        }
      }
    }
  }
  return worst
}

test('the 16 W3C modes match the spec formulas on a 9 x 9 grid within +-1', () => {
  const modes = [...Object.keys(W3C_SEPARABLE), ...Object.keys(W3C_NON_SEPARABLE)]
  assert.equal(modes.length, 16)
  for (const mode of modes) checkMode(mode)
})

test('the 11 extra Photoshop modes match Appendix B within +-1', () => {
  const modes = [...Object.keys(EXTRA_SEPARABLE), ...Object.keys(EXTRA_NON_SEPARABLE)]
  assert.equal(modes.length, 10, 'Dissolve, the eleventh, is checked separately')
  for (const mode of modes) checkMode(mode)
})

test('blendChannel and blendColor agree with the reference', () => {
  for (const mode of Bl.BLEND_MODES) {
    for (let i = 0; i < 9; i += 1) {
      for (let j = 0; j < 9; j += 1) {
        const b = [GRID[i], GRID[(i + 4) % 9], GRID[(i + 7) % 9]].map((v) => v / 255)
        const s = [GRID[j], GRID[(j + 2) % 9], GRID[(j + 6) % 9]].map((v) => v / 255)
        const expected = mode === 'dissolve' ? s : referenceBlend(mode, b, s)
        const actual = Bl.blendColor(mode, b, s)
        actual.forEach((value, c) => assert.ok(Math.abs(value - expected[c]) < 1e-9, `${mode} blendColor`))
        if (Bl.isSeparableBlendMode(mode)) {
          assert.ok(Math.abs(Bl.blendChannel(mode, b[0], s[0]) - expected[0]) < 1e-9, `${mode} blendChannel`)
        }
      }
    }
  }
  for (const mode of ['hue', 'saturation', 'color', 'luminosity', 'darker-color', 'lighter-color']) {
    assert.equal(Bl.isSeparableBlendMode(mode), false)
    assert.throws(() => Bl.blendChannel(mode, 0.2, 0.4), RangeError)
  }
})

test('mode lists: 27 modes in menu order, canvas equivalents only for the 16 W3C modes', () => {
  assert.equal(Bl.BLEND_MODES.length, 27)
  assert.deepEqual([...Bl.BLEND_MODES], BLEND_MODE_MENU.filter((entry) => entry !== '-'))
  assert.deepEqual(Object.keys(Bl.CANVAS_COMPOSITE).sort(), [...Object.keys(W3C_SEPARABLE), ...Object.keys(W3C_NON_SEPARABLE)].sort())
  assert.equal(Bl.CANVAS_COMPOSITE.normal, 'source-over')
  assert.equal(Bl.blendModeLabel('linear-dodge'), 'Linear Dodge (Add)')
  assert.equal(Bl.isBlendMode('multiply'), true)
  assert.equal(Bl.isBlendMode('toString'), false)
  assert.throws(() => Bl.blendInto(new Uint8ClampedArray(4), 0, new Uint8ClampedArray(4).fill(255), 0, 1, 'toString', 1), RangeError)
})

test('opacity 0 and coverage 0 are no-ops; coverage scales the result linearly', () => {
  const next = R.mulberry32(11)
  const count = 64
  const backdrop = Uint8ClampedArray.from({ length: count * 4 }, (_, i) => (i % 4 === 3 ? 60 + (i % 7) * 25 : next() * 256))
  const source = Uint8ClampedArray.from({ length: count * 4 }, (_, i) => (i % 4 === 3 ? 100 + (i % 5) * 30 : next() * 256))
  for (const mode of Bl.BLEND_MODES) {
    const dst = new Uint8ClampedArray(backdrop)
    Bl.blendInto(dst, 0, source, 0, count, mode, 0)
    assert.deepEqual(dst, backdrop, `${mode}: opacity 0`)
    Bl.blendInto(dst, 0, source, 0, count, mode, 1, new Uint8Array(count), 0)
    assert.deepEqual(dst, backdrop, `${mode}: coverage 0`)
  }
  for (const mode of Bl.BLEND_MODES.filter((m) => m !== 'dissolve')) {
    const full = new Uint8ClampedArray(backdrop)
    Bl.blendInto(full, 0, source, 0, count, mode, 1)
    for (const coverageValue of [64, 128, 191]) {
      const partial = new Uint8ClampedArray(backdrop)
      Bl.blendInto(partial, 0, source, 0, count, mode, 1, new Uint8Array(count).fill(coverageValue), 0)
      const t = coverageValue / 255
      for (let p = 0; p < count; p += 1) {
        const i = p * 4
        // The W3C formula is linear in the source alpha for premultiplied results: compare those.
        const exactFull = referenceComposite(mode, [...backdrop.subarray(i, i + 3)].map((v) => v / 255), backdrop[i + 3] / 255, [...source.subarray(i, i + 3)].map((v) => v / 255), source[i + 3] / 255)
        const exactPart = referenceComposite(mode, [...backdrop.subarray(i, i + 3)].map((v) => v / 255), backdrop[i + 3] / 255, [...source.subarray(i, i + 3)].map((v) => v / 255), (source[i + 3] / 255) * t)
        for (let c = 0; c < 3; c += 1) {
          const premultipliedFull = (exactFull[c] * exactFull[3]) / 255
          const premultipliedBackdrop = (backdrop[i + c] * backdrop[i + 3]) / 255
          const lerp = premultipliedBackdrop + (premultipliedFull - premultipliedBackdrop) * t
          assert.ok(Math.abs((exactPart[c] * exactPart[3]) / 255 - lerp) < 1e-6, `${mode}: linear in coverage`)
          assert.ok(Math.abs((partial[i + c] * partial[i + 3]) / 255 - lerp) <= 1.5, `${mode}: implementation follows`)
        }
        assert.ok(Math.abs(partial[i + 3] - (backdrop[i + 3] + (full[i + 3] - backdrop[i + 3]) * t)) <= 1)
      }
    }
  }
})

test('Dissolve keeps a pixel when hash2 < alpha: deterministic, all-or-nothing', () => {
  const count = 10000
  const backdrop = new Uint8ClampedArray(count * 4).fill(0)
  for (let i = 3; i < backdrop.length; i += 4) backdrop[i] = 255
  const source = new Uint8ClampedArray(count * 4).fill(255)
  const a = new Uint8ClampedArray(backdrop)
  const b = new Uint8ClampedArray(backdrop)
  Bl.blendInto(a, 0, source, 0, count, 'dissolve', 0.3, null, 0, 77)
  Bl.blendInto(b, 0, source, 0, count, 'dissolve', 0.3, null, 0, 77)
  assert.deepEqual(a, b, 'same seed, same pattern')
  let kept = 0
  for (let p = 0; p < count; p += 1) {
    const value = a[p * 4]
    assert.ok(value === 0 || value === 255, 'pixels are either replaced or untouched')
    if (value === 255) {
      kept += 1
      assert.equal(R.hash2(p, 0, 77) < 0.3, true)
    }
  }
  assert.ok(Math.abs(kept / count - 0.3) < 0.03, `kept ${kept}`)
  const c = new Uint8ClampedArray(backdrop)
  Bl.blendInto(c, 0, source, 0, count, 'dissolve', 0.3, null, 0, 78)
  assert.notDeepEqual(c, a, 'another seed, another pattern')
  const opaque = new Uint8ClampedArray(backdrop)
  Bl.blendInto(opaque, 0, source, 0, count, 'dissolve', 1)
  assert.ok(opaque.every((value) => value === 255), 'fully opaque dissolve is normal')
})

test('blendInto works on runs inside larger buffers and validates its indices', () => {
  const dst = new Uint8ClampedArray(16 * 4).fill(10)
  const src = new Uint8ClampedArray(16 * 4).fill(200)
  for (let i = 3; i < src.length; i += 4) src[i] = 255
  const coverage = new Uint8Array(16).fill(255)
  Bl.blendInto(dst, 4 * 4, src, 8 * 4, 3, 'normal', 1, coverage, 5)
  for (let p = 0; p < 16; p += 1) {
    const expected = p >= 4 && p < 7 ? 200 : 10
    assert.equal(dst[p * 4], expected, `pixel ${p}`)
  }
  const transparent = new Uint8ClampedArray(4)
  Bl.blendInto(transparent, 0, Uint8ClampedArray.from([9, 8, 7, 100]), 0, 1, 'multiply', 0.5)
  assert.deepEqual([...transparent], [9, 8, 7, 50], 'over a transparent backdrop the source colour shows as is')
  assert.throws(() => Bl.blendInto(dst, 2, src, 0, 1, 'normal', 1), /multiples of 4/)
  assert.throws(() => Bl.blendInto(dst, 0, src, 60, 2, 'normal', 1), RangeError)
  assert.throws(() => Bl.blendInto(dst, 0, src, 0, 2, 'normal', 1, new Uint8Array(1), 0), RangeError)
  Bl.blendInto(dst, 0, src, 0, 0, 'normal', 1)
})

test('blending a 1 MP layer is fast enough for interactive compositing', () => {
  const count = 1 << 20
  const next = R.mulberry32(5)
  const backdrop = Uint8ClampedArray.from({ length: count * 4 }, () => next() * 256)
  const source = Uint8ClampedArray.from({ length: count * 4 }, () => next() * 256)
  for (const mode of ['normal', 'multiply', 'luminosity']) {
    const dst = new Uint8ClampedArray(backdrop)
    Bl.blendInto(dst, 0, source, 0, 4096, mode, 0.8)
    let best = Infinity
    for (let run = 0; run < 3; run += 1) {
      dst.set(backdrop)
      const started = process.hrtime.bigint()
      Bl.blendInto(dst, 0, source, 0, count, mode, 0.8)
      best = Math.min(best, Number(process.hrtime.bigint() - started) / 1e6)
    }
    assert.ok(best < 250, `${mode}: ${best.toFixed(0)} ms for 1 MP`)
  }
})
