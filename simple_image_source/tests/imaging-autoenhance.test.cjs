'use strict'
// WP2 histogram and automatic corrections: histogram counting rules, Simple mode's Auto (hidden Levels,
// gamma, saturation, warmth, highlight/shadow nudges) and Photoshop's Auto Tone / Contrast / Color.
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

const H = load('imaging/histogram.ts')
const Auto = load('imaging/autoEnhance.ts')
const A = load('imaging/adjustments.ts')
const R = load('imaging/random.ts')
const { handlers } = load('shared/worker-ops/color.ts')

/** width x height image whose pixel colour is f(x, y) -> [r, g, b, a?]. */
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

function channelRange(buffer, channel) {
  let min = 255
  let max = 0
  for (let i = channel; i < buffer.data.length; i += 4) {
    min = Math.min(min, buffer.data[i])
    max = Math.max(max, buffer.data[i])
  }
  return [min, max]
}

function mean(buffer, channel) {
  let sum = 0
  for (let i = channel; i < buffer.data.length; i += 4) sum += buffer.data[i]
  return sum / (buffer.data.length / 4)
}

/** A 64..191 grey gradient: the Simple-mode Auto fixture. */
const gradient = () => image(256, 64, (x) => {
  const v = 64 + Math.floor(x / 2)
  return [v, v, v]
})

test('computeHistogram counts visible pixels per channel, alpha for everything the mask selects', () => {
  const src = image(4, 2, (x, y) => (y === 0 ? [x * 10, 100, 200, 255] : [50, 60, 70, x === 0 ? 0 : 128]))
  const histogram = H.computeHistogram(src)
  assert.equal(histogram.count, 7, 'the fully transparent pixel is not counted')
  assert.equal(histogram.alpha[0], 1)
  assert.equal(histogram.alpha[255], 4)
  assert.equal(histogram.alpha[128], 3)
  assert.equal(histogram.red[50], 3)
  assert.equal(histogram.green[100], 4)
  assert.equal(histogram.blue[200], 4)
  const luma = Math.round(0.2126 * 50 + 0.7152 * 60 + 0.0722 * 70)
  assert.equal(histogram.luma[luma], 3, 'Rec. 709 luma')
  const total = (bins) => bins.reduce((sum, value) => sum + value, 0)
  for (const key of ['red', 'green', 'blue', 'luma']) assert.equal(total(histogram[key]), 7)
  const mask = { width: 4, height: 2, data: Uint8Array.from([255, 255, 0, 0, 0, 0, 0, 255]) }
  const masked = H.computeHistogram(src, mask)
  assert.equal(masked.count, 3)
  assert.equal(total(masked.alpha), 3)
  assert.throws(() => H.computeHistogram(src, { width: 1, height: 1, data: new Uint8Array(1) }), RangeError)
})

test('histogram percentile and mean helpers', () => {
  const bins = new Uint32Array(256)
  bins[10] = 1
  bins[20] = 2
  bins[30] = 1
  assert.equal(H.histogramPercentile(bins, 0), 10)
  assert.equal(H.histogramPercentile(bins, 0.25), 10)
  assert.equal(H.histogramPercentile(bins, 0.5), 20)
  assert.equal(H.histogramPercentile(bins, 1), 30)
  assert.equal(H.histogramMean(bins), 20)
  assert.equal(H.histogramPercentile(new Uint32Array(256), 0.5), -1)
})

test('Auto stretches a 64..191 gradient to at least 95% of the range', () => {
  const src = gradient()
  const result = Auto.autoEnhance(src)
  const auto = result.quick.auto
  assert.ok(auto, 'hidden levels are set')
  assert.ok(Math.abs(auto.inBlack - 64) <= 1 && Math.abs(auto.inWhite - 191) <= 1, JSON.stringify(auto))
  assert.ok(auto.gamma >= 0.95 && auto.gamma <= 1.05, 'a symmetric gradient needs almost no gamma')
  assert.equal(result.quick.warmth, 0, 'neutral grey has no cast')
  assert.ok(result.strength > 0.5)
  const out = A.applyAdjustments(src, [{ type: 'quick', ...result.quick }])
  const [min, max] = channelRange(out, 0)
  assert.ok(max - min >= 0.95 * 255, `output spans ${min}..${max}`)
})

test('Auto leaves a well exposed full-range image almost alone', () => {
  const next = R.mulberry32(4)
  const src = image(300, 200, (x) => {
    const v = Math.round((x / 299) * 255) + Math.floor(next() * 21) - 10
    return [v, v, v, 255]
  })
  const result = Auto.autoEnhance(src)
  const auto = result.quick.auto
  assert.ok(!auto || (auto.inBlack <= 3 && auto.inWhite >= 252), JSON.stringify(auto))
  assert.ok(result.strength < 0.35, `strength ${result.strength}`)
})

test('Auto warms a blue cast, cools a warm cast and boosts muted colour', () => {
  const blueCast = image(200, 100, (x) => {
    const v = 60 + Math.floor(x * 0.7)
    return [v - 12, v, v + 14]
  })
  const cool = Auto.autoEnhance(blueCast).quick
  assert.ok(cool.warmth > 0 && cool.warmth <= 30, `warmth ${cool.warmth}`)
  assert.equal(cool.saturation, 12, 'muted image gets +12 saturation')
  const warmCast = image(200, 100, (x) => {
    const v = 60 + Math.floor(x * 0.7)
    return [v + 14, v, v - 12]
  })
  const warm = Auto.autoEnhance(warmCast).quick
  assert.ok(warm.warmth < 0 && warm.warmth >= -30, `warmth ${warm.warmth}`)
  const vivid = image(100, 100, (x) => (x % 2 ? [250, 20, 30] : [20, 40, 240]))
  assert.equal(Auto.autoEnhance(vivid).quick.saturation, 0, 'already colourful')
})

test('Auto recovers clipped highlights and lifts crushed shadows', () => {
  const clipped = image(100, 100, (x, y) => (y < 10 ? [255, 255, 255] : [100 + (x % 50), 100, 100]))
  assert.equal(Auto.autoEnhance(clipped).quick.highlights, -10)
  const crushed = image(100, 100, (x, y) => (y < 10 ? [0, 0, 0] : [100 + (x % 50), 100, 100]))
  assert.equal(Auto.autoEnhance(crushed).quick.shadows, 10)
  const transparent = image(10, 10, () => [200, 10, 10, 0])
  assert.deepEqual(Auto.autoEnhance(transparent), {
    quick: { exposure: 0, brightness: 0, contrast: 0, highlights: 0, shadows: 0, saturation: 0, warmth: 0, auto: null },
    strength: 0,
  })
  const flat = image(20, 20, () => [120, 120, 120])
  assert.equal(Auto.autoEnhance(flat).quick.auto, null, 'a flat image is never stretched')
})

test('Auto samples a bounded proxy, so large images stay fast', () => {
  const big = image(2400, 1800, (x, y) => [(x + y) & 255, x & 255, y & 255])
  const started = Date.now()
  const result = Auto.autoEnhance(big)
  assert.ok(Date.now() - started < 1500, 'about 1 MP of samples')
  assert.ok(result.strength >= 0 && result.strength <= 1)
})

test('Auto Tone stretches each channel, Auto Contrast keeps colour balance, Auto Color neutralises a cast', () => {
  const src = image(256, 16, (x) => [40 + Math.floor(x * 0.5), 80 + Math.floor(x * 0.4), 20 + Math.floor(x * 0.6)])
  const tone = Auto.autoTone(src)
  assert.equal(tone.type, 'levels')
  const toned = A.applyAdjustments(src, [tone])
  for (let channel = 0; channel < 3; channel += 1) {
    const [min, max] = channelRange(toned, channel)
    assert.ok(min <= 2 && max >= 253, `channel ${channel}: ${min}..${max}`)
  }
  const contrast = Auto.autoContrast(src)
  assert.deepEqual(contrast.red, { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 })
  assert.ok(contrast.rgb.inBlack >= 19 && contrast.rgb.inWhite <= 237, JSON.stringify(contrast.rgb))
  const contrasted = A.applyAdjustments(src, [contrast])
  assert.ok(mean(contrasted, 1) > mean(contrasted, 0), 'green stays greener than red')

  const cast = image(200, 100, (x, y) => {
    const v = 30 + Math.floor((x / 199) * 200)
    return y < 5 ? [5, 5, 15] : y > 94 ? [235, 240, 255] : [v, v + 6, Math.min(255, v + 25)]
  })
  const color = Auto.autoColor(cast)
  const neutral = A.applyAdjustments(cast, [color])
  const index = (50 * 200 + 100) * 4
  const [r, g, b] = neutral.data.subarray(index, index + 3)
  assert.ok(Math.abs(r - b) <= Math.abs(cast.data[index] - cast.data[index + 2]) / 3, `cast reduced: ${r},${g},${b}`)
  const none = Auto.autoColor(image(4, 4, () => [0, 0, 0, 0]))
  assert.deepEqual(none.rgb, { inBlack: 0, inWhite: 255, gamma: 1, outBlack: 0, outWhite: 255 })
})

test('histogram and autoEnhance worker ops', async () => {
  const src = gradient()
  const histogram = await handlers.histogram({ src, mask: null }, {})
  assert.equal(histogram.count, src.width * src.height)
  const result = await handlers.autoEnhance({ src }, {})
  assert.deepEqual(result, Auto.autoEnhance(src))
  const controller = new AbortController()
  controller.abort()
  assert.throws(() => handlers.autoEnhance({ src }, { signal: controller.signal }), { name: 'AbortError' })
})
