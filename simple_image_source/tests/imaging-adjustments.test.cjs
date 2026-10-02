'use strict'
// WP2 adjustments: identity specs are byte-exact, Levels/Curves follow their definitions, the colour
// adjustments behave like Photoshop's, quick adjust folds into tables, masks and opacity mix correctly, the
// 'adjust' worker op wraps it all, and table-based kernels are fast.
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

const A = load('imaging/adjustments.ts')
const C = load('imaging/color.ts')
const R = load('imaging/random.ts')
const { handlers } = load('shared/worker-ops/color.ts')

function randomBuffer(width, height, seed = 1, opaque = false) {
  const next = R.mulberry32(seed)
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i += 1) data[i] = Math.floor(next() * 256)
  if (opaque) for (let i = 3; i < data.length; i += 4) data[i] = 255
  return { width, height, data }
}

/** One row of the given colours (alpha 255 unless given). */
function row(...colors) {
  const data = new Uint8ClampedArray(colors.length * 4)
  colors.forEach((color, index) => data.set([color[0], color[1], color[2], color.length > 3 ? color[3] : 255], index * 4))
  return { width: colors.length, height: 1, data }
}

function grayRamp() {
  return row(...Array.from({ length: 256 }, (_, v) => [v, v, v]))
}

function px(buffer, index) {
  return [...buffer.data.subarray(index * 4, index * 4 + 4)]
}

function apply(src, ...specs) {
  return A.applyAdjustments(src, specs)
}

function maxDiff(a, b) {
  let max = 0
  for (let i = 0; i < a.data.length; i += 1) max = Math.max(max, Math.abs(a.data[i] - b.data[i]))
  return max
}

function alphaUnchanged(src, out) {
  for (let i = 3; i < src.data.length; i += 4) if (src.data[i] !== out.data[i]) return false
  return true
}

const IDENTITY_SPECS = () => [
  A.defaultAdjustment('brightness-contrast'),
  { type: 'brightness-contrast', brightness: 0, contrast: 0, legacy: true },
  A.defaultAdjustment('levels'),
  A.defaultAdjustment('curves'),
  A.defaultAdjustment('exposure'),
  A.defaultAdjustment('vibrance'),
  A.defaultAdjustment('hue-saturation'),
  { ...A.defaultAdjustment('hue-saturation'), master: { hue: 360, saturation: 0, lightness: 0 } },
  A.defaultAdjustment('color-balance'),
  { ...A.defaultAdjustment('color-balance'), preserveLuminosity: false },
  { ...A.defaultAdjustment('photo-filter'), density: 0 },
  A.defaultAdjustment('quick'),
]

test('identity specs compile to identity kernels and return identical bytes', () => {
  const src = randomBuffer(97, 31, 5)
  for (const spec of IDENTITY_SPECS()) {
    const kernel = A.compileAdjustment(spec)
    assert.equal(kernel.isIdentity, true, `${spec.type} is the identity`)
    const out = A.applyAdjustments(src, [spec])
    assert.notEqual(out.data, src.data, 'a new buffer')
    assert.equal(Buffer.compare(Buffer.from(out.data), Buffer.from(src.data)), 0, `${spec.type} keeps every byte`)
  }
  const out = A.applyAdjustments(src, IDENTITY_SPECS())
  assert.equal(Buffer.compare(Buffer.from(out.data), Buffer.from(src.data)), 0, 'a chain of identities too')
})

test('every adjustment type leaves alpha untouched and has a label and a default', () => {
  const src = randomBuffer(64, 16, 9)
  const types = [...A.ADJUSTMENT_TYPES, 'quick']
  assert.equal(A.ADJUSTMENT_TYPES.length, 13)
  for (const type of types) {
    const spec = A.defaultAdjustment(type)
    assert.equal(spec.type, type)
    assert.ok(A.adjustmentLabel(type).length > 0)
    const changed = type === 'quick' ? { ...spec, exposure: 30, saturation: 40 } : spec
    assert.ok(alphaUnchanged(src, apply(src, changed)), `${type} keeps alpha`)
  }
  assert.equal(A.adjustmentLabel('hue-saturation'), 'Hue/Saturation')
  assert.equal(A.adjustmentLabel('black-white'), 'Black & White')
  const first = A.defaultAdjustment('levels')
  first.rgb.gamma = 3
  assert.equal(A.defaultAdjustment('levels').rgb.gamma, 1, 'defaults are fresh objects')
  assert.throws(() => A.defaultAdjustment('nope'), RangeError)
  assert.throws(() => A.compileAdjustment({ type: 'nope' }), RangeError)
})

test('Levels maps known values exactly and the gamma midpoint lands on 128 +- 1', () => {
  const ramp = grayRamp()
  const levels = (rgb) => ({ ...A.defaultAdjustment('levels'), rgb: { ...A.defaultAdjustment('levels').rgb, ...rgb } })
  const out = apply(ramp, levels({ inBlack: 50, inWhite: 200 }))
  for (let v = 0; v <= 50; v += 1) assert.equal(px(out, v)[0], 0)
  for (let v = 200; v <= 255; v += 1) assert.equal(px(out, v)[0], 255)
  for (let v = 51; v < 200; v += 1) assert.equal(px(out, v)[0], Math.round(((v - 50) / 150) * 255))
  const output = apply(ramp, levels({ outBlack: 20, outWhite: 220 }))
  assert.equal(px(output, 0)[0], 20)
  assert.equal(px(output, 255)[0], 220)
  assert.equal(px(output, 128)[0], Math.round(20 + (200 * 128) / 255))
  const inverted = apply(ramp, levels({ outBlack: 255, outWhite: 0 }))
  assert.equal(px(inverted, 0)[0], 255)
  assert.equal(px(inverted, 255)[0], 0)

  for (const gamma of [0.3, 0.5, 0.8, 1.6, 2, 3.5]) {
    for (const [inBlack, inWhite] of [[0, 255], [20, 230]]) {
      const midpoint = Math.round(inBlack + (inWhite - inBlack) * Math.pow(0.5, gamma))
      const lut = A.levelsLut({ inBlack, inWhite, gamma, outBlack: 0, outWhite: 255 })
      assert.ok(Math.abs(lut[midpoint] - 128) <= 1, `gamma ${gamma}: input ${midpoint} -> ${lut[midpoint]}`)
    }
  }
  assert.ok(A.levelsLut({ inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 }).every((value, index) => value === index))
})

test('Levels applies the individual channel before the RGB master', () => {
  const spec = {
    ...A.defaultAdjustment('levels'),
    red: { inBlack: 0, inWhite: 128, gamma: 1, outBlack: 0, outWhite: 255 },
    rgb: { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 128 },
  }
  const out = apply(row([64, 64, 64]), spec)
  // red: 64 -> 127.5 by the channel, then * 128/255 by the master.
  assert.equal(px(out, 0)[0], Math.round((((64 / 128) * 255) * 128) / 255))
  assert.equal(px(out, 0)[1], Math.round((64 * 128) / 255))
})

test('Curves pass through their points, stay flat outside them and chain channel then master', () => {
  const points = [{ x: 0, y: 0 }, { x: 64, y: 90 }, { x: 128, y: 128 }, { x: 192, y: 170 }, { x: 255, y: 255 }]
  const lut = A.curveLut(points)
  for (const point of points) assert.equal(lut[point.x], point.y)
  for (let v = 1; v < 256; v += 1) assert.ok(lut[v] >= lut[v - 1], 'this S-curve is monotone')
  const flat = A.curveLut([{ x: 200, y: 180 }, { x: 30, y: 60 }])
  for (let v = 0; v <= 30; v += 1) assert.equal(flat[v], 60)
  for (let v = 200; v <= 255; v += 1) assert.equal(flat[v], 180)
  assert.equal(flat[115], Math.round(60 + ((115 - 30) * 120) / 170), 'two points give a straight line')
  assert.ok(A.curveLut([{ x: 0, y: 0 }, { x: 255, y: 255 }]).every((value, index) => value === index))
  assert.ok(A.curveLut([{ x: 0, y: 300 }, { x: 255, y: -20 }]).every((value) => value >= 0 && value <= 255), 'clamped to 0..255')

  const linear = [{ x: 0, y: 0 }, { x: 255, y: 255 }]
  const spec = { type: 'curves', rgb: [{ x: 0, y: 0 }, { x: 255, y: 128 }], red: [{ x: 0, y: 255 }, { x: 255, y: 0 }], green: linear, blue: linear }
  const out = apply(row([0, 0, 0], [255, 255, 255]), spec)
  assert.deepEqual(px(out, 0), [128, 0, 0, 255], 'red: inverted by its channel curve, then halved by the master')
  assert.deepEqual(px(out, 1), [0, 128, 128, 255])
})

test('Invert twice is the identity and Posterize yields exactly n levels', () => {
  const src = randomBuffer(50, 20, 3)
  const once = apply(src, { type: 'invert' })
  for (let i = 0; i < src.data.length; i += 4) {
    assert.equal(once.data[i], 255 - src.data[i])
    assert.equal(once.data[i + 3], src.data[i + 3])
  }
  assert.equal(Buffer.compare(Buffer.from(apply(src, { type: 'invert' }, { type: 'invert' }).data), Buffer.from(src.data)), 0)
  const ramp = grayRamp()
  for (const levels of [2, 3, 4, 5, 7, 16, 64, 128, 255]) {
    const out = apply(ramp, { type: 'posterize', levels })
    const distinct = new Set()
    for (let v = 0; v < 256; v += 1) distinct.add(px(out, v)[0])
    assert.equal(distinct.size, levels, `posterize ${levels}`)
    assert.ok(distinct.has(0) && distinct.has(255))
  }
})

test('Hue +360 is the identity within +-1, also through the HSL path', () => {
  const src = randomBuffer(64, 64, 17, true)
  const shortcut = apply(src, { ...A.defaultAdjustment('hue-saturation'), master: { hue: 360, saturation: 0, lightness: 0 } })
  assert.equal(maxDiff(shortcut, src), 0)
  const throughHsl = apply(src, { ...A.defaultAdjustment('hue-saturation'), master: { hue: 360, saturation: 1e-9, lightness: 0 } })
  assert.ok(maxDiff(throughHsl, src) <= 1, `max difference ${maxDiff(throughHsl, src)}`)
  const halfTurn = { ...A.defaultAdjustment('hue-saturation'), master: { hue: 180, saturation: 0, lightness: 0 } }
  assert.ok(maxDiff(apply(src, halfTurn, halfTurn), src) <= 1)
  const red = apply(row([255, 0, 0]), { ...A.defaultAdjustment('hue-saturation'), master: { hue: 120, saturation: 0, lightness: 0 } })
  assert.deepEqual(px(red, 0), [0, 255, 0, 255], '+120 degrees turns red into green')
})

test('Hue/Saturation saturation, lightness, colorize and range windows', () => {
  const hs = (master, extra = {}) => ({ ...A.defaultAdjustment('hue-saturation'), master: { hue: 0, saturation: 0, lightness: 0, ...master }, ...extra })
  const color = row([200, 80, 40], [128, 128, 128])
  const gray = apply(color, hs({ saturation: -100 }))
  const l = Math.round((200 + 40) / 2)
  assert.deepEqual(px(gray, 0), [l, l, l, 255], '-100 saturation gives HSL lightness')
  assert.deepEqual(px(apply(color, hs({ lightness: 100 })), 0), [255, 255, 255, 255])
  assert.deepEqual(px(apply(color, hs({ lightness: -100 })), 0), [0, 0, 0, 255])
  assert.deepEqual(px(apply(color, hs({ lightness: 50 })), 1), [192, 192, 192, 255])
  const colorized = apply(color, hs({ hue: 240, saturation: 100 }, { colorize: true }))
  for (const index of [0, 1]) {
    const [hue, s, light] = C.rgbToHsl(...px(colorized, index).slice(0, 3).map((v) => v / 255))
    const original = px(color, index)
    const originalLight = (Math.max(...original.slice(0, 3)) + Math.min(...original.slice(0, 3))) / 2 / 255
    assert.ok(Math.abs(hue - 240) < 1.5, `colorize sets the hue (${hue})`)
    assert.ok(s > 0.98, 'colorize sets the saturation, greys included')
    assert.ok(Math.abs(light - originalLight) < 0.005, 'colorize keeps HSL lightness')
  }
  // Ranges: shifting only the blues leaves reds and greys alone.
  const ranges = hs({}, { ranges: { blues: { hue: 0, saturation: -100, lightness: 0 } } })
  const mixed = apply(row([220, 30, 30], [30, 30, 220], [100, 100, 100]), ranges)
  assert.deepEqual(px(mixed, 0), [220, 30, 30, 255])
  const blue = px(mixed, 1)
  assert.equal(blue[0], blue[2], 'blue is desaturated')
  assert.deepEqual(px(mixed, 2), [100, 100, 100, 255])
  assert.equal(A.hueRangeWeight('reds', 0), 1)
  assert.equal(A.hueRangeWeight('reds', 345), 1)
  assert.equal(A.hueRangeWeight('reds', 30), 0.5)
  assert.equal(A.hueRangeWeight('yellows', 30), 0.5)
  assert.equal(A.hueRangeWeight('reds', 60), 0)
})

test('Brightness/Contrast, Exposure and Vibrance', () => {
  const ramp = grayRamp()
  const legacy = apply(ramp, { type: 'brightness-contrast', brightness: 20, contrast: 0, legacy: true })
  for (let v = 0; v < 256; v += 1) assert.equal(px(legacy, v)[0], Math.min(255, v + 20))
  const brighter = apply(ramp, { type: 'brightness-contrast', brightness: 60, contrast: 0, legacy: false })
  for (let v = 1; v < 255; v += 1) assert.ok(px(brighter, v)[0] >= v)
  assert.equal(px(brighter, 0)[0], 0, 'modern brightness keeps black')
  assert.equal(px(brighter, 255)[0], 255, 'modern brightness keeps white')
  const contrast = apply(ramp, { type: 'brightness-contrast', brightness: 0, contrast: 80, legacy: false })
  assert.ok(px(contrast, 64)[0] < 64 && px(contrast, 192)[0] > 192)
  const flat = apply(ramp, { type: 'brightness-contrast', brightness: 0, contrast: -50, legacy: false })
  assert.ok(px(flat, 0)[0] > 0 && px(flat, 255)[0] < 255)

  const stop = apply(row([50, 100, 200]), { type: 'exposure', exposure: 1, offset: 0, gamma: 1 })
  assert.deepEqual(px(stop, 0).slice(0, 3), [50, 100, 200].map((v) => C.linearToSrgb8(C.SRGB_TO_LINEAR[v] * 2)), '+1 stop doubles linear light')
  const darker = apply(row([200, 200, 200]), { type: 'exposure', exposure: -2, offset: 0, gamma: 1 })
  assert.equal(px(darker, 0)[0], C.linearToSrgb8(C.SRGB_TO_LINEAR[200] / 4))

  const vib = { type: 'vibrance', vibrance: 100, saturation: 0 }
  const colors = row([140, 120, 100], [250, 20, 20], [120, 120, 120], [100, 110, 160])
  const out = apply(colors, vib)
  const chroma = (p) => Math.max(...p.slice(0, 3)) - Math.min(...p.slice(0, 3))
  assert.deepEqual(px(out, 2), [120, 120, 120, 255], 'greys stay grey')
  const mutedGain = chroma(px(out, 3)) / chroma(px(colors, 3))
  const vividGain = chroma(px(out, 1)) / chroma(px(colors, 1))
  assert.ok(mutedGain > vividGain, 'vibrance boosts muted colours more than saturated ones')
  const desaturated = apply(colors, { type: 'vibrance', vibrance: 0, saturation: -100 })
  for (let index = 0; index < 4; index += 1) {
    const p = px(desaturated, index)
    assert.ok(Math.max(...p.slice(0, 3)) - Math.min(...p.slice(0, 3)) <= 1, 'saturation -100 removes colour')
  }
})

test('Color Balance, Black & White, Photo Filter, Threshold and Gradient Map', () => {
  const mid = row([128, 128, 128], [40, 40, 40], [230, 230, 230])
  const warmMid = apply(mid, { ...A.defaultAdjustment('color-balance'), midtones: { cyanRed: 100, magentaGreen: 0, yellowBlue: 0 }, preserveLuminosity: false })
  assert.ok(px(warmMid, 0)[0] > 128 + 40, 'midtones +100 red moves mid grey towards red')
  assert.ok(px(warmMid, 1)[0] - 40 < px(warmMid, 0)[0] - 128, 'shadows move less than midtones')
  const preserved = apply(mid, { ...A.defaultAdjustment('color-balance'), midtones: { cyanRed: 60, magentaGreen: -20, yellowBlue: 30 }, preserveLuminosity: true })
  for (let index = 0; index < 3; index += 1) {
    const p = px(preserved, index)
    const lightness = (Math.max(...p.slice(0, 3)) + Math.min(...p.slice(0, 3))) / 2
    const original = px(mid, index)[0]
    assert.ok(Math.abs(lightness - original) <= 1, `preserve luminosity keeps HSL lightness (${lightness} vs ${original})`)
  }

  const bw = apply(row([255, 0, 0], [255, 255, 0], [0, 0, 255], [90, 90, 90], [255, 0, 255]), A.defaultAdjustment('black-white'))
  assert.deepEqual(px(bw, 0), [102, 102, 102, 255], 'reds 40%')
  assert.deepEqual(px(bw, 1), [153, 153, 153, 255], 'yellows 60%')
  assert.deepEqual(px(bw, 2), [51, 51, 51, 255], 'blues 20%')
  assert.deepEqual(px(bw, 3), [90, 90, 90, 255], 'greys are unchanged')
  assert.deepEqual(px(bw, 4), [204, 204, 204, 255], 'magentas 80%')
  const tinted = apply(row([90, 90, 90]), { ...A.defaultAdjustment('black-white'), tint: { r: 225, g: 211, b: 179 } })
  const t = px(tinted, 0)
  assert.ok(t[0] > t[2], 'the tint colours the grey')
  assert.ok(Math.abs(C.lum(t[0], t[1], t[2]) - 90) <= 1, 'the tinted grey keeps its luminosity')

  const filterSpec = { type: 'photo-filter', color: { r: 236, g: 138, b: 0 }, density: 50, preserveLuminosity: false }
  const warmed = px(apply(row([128, 128, 128]), filterSpec), 0)
  assert.ok(warmed[0] > warmed[1] && warmed[1] > warmed[2], 'warming filter')
  const keep = px(apply(row([128, 128, 128], [60, 120, 200]), { ...filterSpec, preserveLuminosity: true }), 1)
  assert.ok(Math.abs(C.lum(keep[0], keep[1], keep[2]) - C.lum(60, 120, 200)) <= 1, 'preserve luminosity')

  const threshold = apply(row([128, 128, 128], [127, 127, 127], [255, 0, 0], [0, 255, 0]), { type: 'threshold', level: 128 })
  assert.deepEqual([0, 1, 2, 3].map((i) => px(threshold, i)[0]), [255, 0, 0, 255], 'Y601 >= level is white')

  const ramp = grayRamp()
  const map = apply(ramp, A.defaultAdjustment('gradient-map'))
  for (let v = 0; v < 256; v += 1) assert.deepEqual(px(map, v), [v, v, v, 255], 'black-to-white map of a grey ramp')
  const reversed = apply(ramp, { ...A.defaultAdjustment('gradient-map'), reverse: true })
  for (let v = 0; v < 256; v += 1) assert.equal(px(reversed, v)[0], 255 - v)
  const duo = apply(row([0, 0, 0], [255, 255, 255], [128, 128, 128]), {
    type: 'gradient-map',
    stops: [{ position: 0, color: { r: 20, g: 40, b: 200 } }, { position: 1, color: { r: 250, g: 200, b: 10 } }],
    reverse: false,
    dither: false,
  })
  assert.deepEqual(px(duo, 0), [20, 40, 200, 255])
  assert.deepEqual(px(duo, 1), [250, 200, 10, 255])
  const biased = apply(row([128, 128, 128]), {
    type: 'gradient-map',
    stops: [{ position: 0, color: { r: 0, g: 0, b: 0 }, midpoint: 0.25 }, { position: 1, color: { r: 255, g: 255, b: 255 } }],
    reverse: false,
    dither: false,
  })
  assert.ok(px(biased, 0)[0] > 160, 'a midpoint of 25% reaches the half colour early')
  const dithered = A.applyAdjustments(ramp, [{ ...A.defaultAdjustment('gradient-map'), dither: true }])
  assert.ok(maxDiff(dithered, map) <= 1, 'dither stays within +-1')
})

test('Quick adjust: exposure, warmth, tone curve, saturation and hidden auto levels', () => {
  const quick = (values) => ({ ...A.defaultAdjustment('quick'), ...values })
  const ramp = grayRamp()
  const exposed = apply(ramp, quick({ exposure: 50 }))
  for (let v = 1; v < 255; v += 1) assert.ok(px(exposed, v)[0] >= v)
  const warm = px(apply(row([128, 128, 128]), quick({ warmth: 60 })), 0)
  assert.ok(warm[0] > 128 && warm[2] < 128 && warm[1] === 128)
  const shadows = apply(ramp, quick({ shadows: 100 }))
  assert.ok(px(shadows, 64)[0] > 64 + 20 && px(shadows, 192)[0] - 192 < 5, 'shadows lift the dark quarter')
  const highlights = apply(ramp, quick({ highlights: -100 }))
  assert.ok(px(highlights, 191)[0] < 191 - 20 && px(highlights, 255)[0] === 255, 'highlights recover the light quarter')
  const gray = apply(row([200, 80, 40]), quick({ saturation: -100 }))
  const g = px(gray, 0)
  assert.ok(Math.max(g[0], g[1], g[2]) - Math.min(g[0], g[1], g[2]) <= 1)
  const autoRamp = row(...Array.from({ length: 128 }, (_, i) => [64 + i, 64 + i, 64 + i]))
  const stretched = apply(autoRamp, quick({ auto: { inBlack: 64, inWhite: 191, gamma: 1, outBlack: 0, outWhite: 255 } }))
  assert.equal(px(stretched, 0)[0], 0)
  assert.equal(px(stretched, 127)[0], 255)
  const kernel = A.compileAdjustment(quick({ contrast: 30, brightness: -20, saturation: 10 }))
  assert.equal(kernel.isIdentity, false)
})

test('masks and opacity mix the result back; kernels honour pixel ranges', () => {
  const src = randomBuffer(8, 2, 13, true)
  const mask = { width: 8, height: 2, data: Uint8Array.from({ length: 16 }, (_, i) => (i % 2 ? 255 : 0)) }
  const masked = A.applyAdjustments(src, [{ type: 'invert' }], mask)
  for (let pixel = 0; pixel < 16; pixel += 1) {
    const expected = pixel % 2 ? 255 - src.data[pixel * 4] : src.data[pixel * 4]
    assert.equal(masked.data[pixel * 4], expected)
  }
  const half = A.applyAdjustments(src, [{ type: 'invert' }], null, 0.5)
  for (let i = 0; i < src.data.length; i += 4) assert.ok(Math.abs(half.data[i] - 127.5) <= 1)
  assert.equal(maxDiff(A.applyAdjustments(src, [{ type: 'invert' }], null, 0), src), 0)
  assert.throws(() => A.applyAdjustments(src, [], { width: 2, height: 2, data: new Uint8Array(4) }), RangeError)

  const data = new Uint8ClampedArray(src.data)
  A.compileAdjustment({ type: 'invert' }).apply(data, 3, 5)
  for (let pixel = 0; pixel < 16; pixel += 1) {
    const inside = pixel >= 3 && pixel < 5
    assert.equal(data[pixel * 4], inside ? 255 - src.data[pixel * 4] : src.data[pixel * 4])
  }
})

test('the adjust worker op matches applyAdjustments, never mutates its input and honours abort', async () => {
  const src = randomBuffer(40, 30, 21)
  const before = new Uint8ClampedArray(src.data)
  const specs = [{ type: 'invert' }, { ...A.defaultAdjustment('quick'), contrast: 20, saturation: -30 }]
  const viaOp = await handlers.adjust({ src, specs, mask: null, opacity: 1 }, {})
  assert.equal(maxDiff(viaOp, A.applyAdjustments(src, specs)), 0)
  assert.deepEqual(src.data, before, 'inline runs share objects with the caller; the input must stay intact')
  const controller = new AbortController()
  controller.abort()
  assert.throws(() => handlers.adjust({ src, specs, mask: null, opacity: 1 }, { signal: controller.signal }), { name: 'AbortError' })
  assert.throws(() => handlers.adjust(null, {}), RangeError)
})

test('startAdjustments gives the same bytes in any chunking (dithered gradient map, mask and opacity included)', () => {
  const src = randomBuffer(37, 90, 23)
  const mask = { width: 37, height: 90, data: Uint8Array.from({ length: 37 * 90 }, (_, i) => (i * 53) % 256) }
  const specs = [
    { ...A.defaultAdjustment('quick'), exposure: 20, saturation: -30 },
    { type: 'gradient-map', stops: [{ position: 0, color: { r: 10, g: 20, b: 90 } }, { position: 1, color: { r: 250, g: 230, b: 120 } }], reverse: false, dither: true },
    { ...A.defaultAdjustment('hue-saturation'), master: { hue: 30, saturation: 20, lightness: -10 } },
  ]
  const whole = A.applyAdjustments(src, specs, mask, 0.7)
  for (const rows of [1, 4, 33, 90]) {
    const run = A.startAdjustments(src, specs, mask, 0.7)
    for (let start = 0; start < run.rows; start += rows) run.process(start, Math.min(run.rows, start + rows))
    assert.equal(Buffer.compare(Buffer.from(run.output.data), Buffer.from(whole.data)), 0, `chunks of ${rows} rows`)
  }
  const identity = A.startAdjustments(src, [A.defaultAdjustment('levels')])
  assert.equal(Buffer.compare(Buffer.from(identity.output.data), Buffer.from(src.data)), 0, 'an identity run is complete at once')
})

test('table-based adjustments run at 150 Mpx/s or better', () => {
  const width = 2000
  const height = 2000
  const src = randomBuffer(width, height, 2, true)
  const kernels = [
    A.compileAdjustment({ ...A.defaultAdjustment('levels'), rgb: { inBlack: 10, inWhite: 240, gamma: 1.2, outBlack: 0, outWhite: 255 } }),
    A.compileAdjustment({ type: 'curves', rgb: [{ x: 0, y: 0 }, { x: 128, y: 150 }, { x: 255, y: 255 }], red: [{ x: 0, y: 0 }, { x: 255, y: 255 }], green: [{ x: 0, y: 0 }, { x: 255, y: 255 }], blue: [{ x: 0, y: 0 }, { x: 255, y: 255 }] }),
    A.compileAdjustment({ ...A.defaultAdjustment('quick'), exposure: 20, contrast: 15, warmth: 10 }),
  ]
  for (const kernel of kernels) {
    const data = new Uint8ClampedArray(src.data)
    kernel.apply(data)
    let best = Infinity
    for (let run = 0; run < 7; run += 1) {
      const started = process.hrtime.bigint()
      kernel.apply(data)
      best = Math.min(best, Number(process.hrtime.bigint() - started) / 1e6)
    }
    const mpxPerSecond = (width * height) / 1e6 / (best / 1000)
    assert.ok(mpxPerSecond >= 150, `${kernel.spec.type}: ${mpxPerSecond.toFixed(0)} Mpx/s (best of 7: ${best.toFixed(1)} ms for 4 MP)`)
  }
})
