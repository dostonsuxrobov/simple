const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const fixture = require('./fixtures/scan-fixture.cjs')

// Retouching scanned glyphs (design 4.8.3 / 6.2(6)): the line "Invoice
// number 48213 ..." of the noisy and the hard (stained, blurred, low
// contrast) fixture scans, retouched the way the preparation worker does it.
const retouchModule = () => import('../electron/ocr-retouch.mjs')
const styleModule = () => import('../electron/scan-style.mjs')
const preprocessModule = () => import('../electron/ocr-preprocess.mjs')

const DPI = 300
const SCALE = DPI / 72
const PAGE_HEIGHT = 792
const LINE_START = 'Invoice'

const rasters = new Map()
async function scanOf(name) {
  if (!rasters.has(name)) {
    rasters.set(name, (async () => {
      const variant = await fixture.buildScanVariant(name)
      return { variant, raster: await fixture.rasterize(variant.pdf, DPI) }
    })())
  }
  return rasters.get(name)
}

/**
 * The region scanEdit renders for a line (the line's box grown by 0.35 line
 * heights and a quarter em), its pixels, and the line geometry in them.
 */
function lineRegion({ variant, raster }, startsWith = LINE_START, transform = (value) => value) {
  const line = variant.truth.lines.find((item) => item.text.startsWith(startsWith))
  const size = line.size
  const first = line.words[0]
  const last = line.words.at(-1)
  const dir = first.baseline.dir
  const up = { x: -dir.y, y: dir.x }
  const origin = first.baseline.origin
  const length = Math.hypot(last.baseline.origin.x + dir.x * last.inkWidth - origin.x, last.baseline.origin.y + dir.y * last.inkWidth - origin.y)
  const corner = (u, v) => ({ x: origin.x + dir.x * u + up.x * v, y: origin.y + dir.y * u + up.y * v })
  const corners = [corner(0, -0.3 * size), corner(length, -0.3 * size), corner(0, 0.95 * size), corner(length, 0.95 * size)]
  const x0 = Math.floor((Math.min(...corners.map((c) => c.x)) - 0.25 * size) * SCALE)
  const x1 = Math.ceil((Math.max(...corners.map((c) => c.x)) + 0.25 * size) * SCALE)
  const y0 = Math.floor((PAGE_HEIGHT - Math.max(...corners.map((c) => c.y)) - 0.42 * size) * SCALE)
  const y1 = Math.ceil((PAGE_HEIGHT - Math.min(...corners.map((c) => c.y)) + 0.42 * size) * SCALE)
  const width = x1 - x0
  const height = y1 - y0
  const pixels = new Uint8Array(width * height)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) pixels[y * width + x] = transform(raster.data[(y + y0) * raster.width + x + x0])
  }
  const toPixels = (point) => ({ x: point.x * SCALE - x0, y: (PAGE_HEIGHT - point.y) * SCALE - y0 })
  const base = toPixels(origin)
  return {
    line, size, pixels, width, height,
    baseline: { x: base.x, y: base.y, dx: dir.x, dy: -dir.y },
    length: length * SCALE,
    fontSize: size * SCALE,
    quads: line.words.map((word) => {
      const o = word.baseline.origin
      const at = (u, v) => toPixels({ x: o.x + dir.x * u + up.x * v, y: o.y + dir.y * u + up.y * v })
      return { quad: [at(0, 0.75 * size), at(word.inkWidth, 0.75 * size), at(word.inkWidth, -0.24 * size), at(0, -0.24 * size)] }
    }),
  }
}

async function retouchRegion(region, options = {}) {
  const { retouchLine } = await retouchModule()
  const { estimateScanStyle } = await styleModule()
  const common = { channels: 1, dpi: DPI, baseline: region.baseline, length: region.length, fontSize: region.fontSize, text: region.line.text }
  const { style, features } = estimateScanStyle(region.pixels, region.width, region.height, common)
  const result = retouchLine(region.pixels, region.width, region.height, {
    ...common, xHeight: features.xHeight, fontClass: style.fontClass, segment: !options.targets, targets: options.targets, seed: options.seed ?? 'page-1-line-invoice',
  })
  return { result, style }
}

const median = (values) => {
  const sorted = Float64Array.from(values).sort()
  return sorted[sorted.length >> 1]
}

for (const name of ['noisy300', 'hard300']) {
  test(`${name}: the retouched area matches the paper around it`, { timeout: 120_000 }, async () => {
    const { composePatch, compositePatch, rgbaOf } = await retouchModule()
    const region = lineRegion(await scanOf(name))
    const { result } = await retouchRegion(region)
    assert.equal(result.segmented, true, 'the line\'s words were found in the pixels')
    assert.deepEqual(result.words.map((word) => word.text), region.line.text.split(' '))
    const patch = composePatch(result, null)
    const composite = compositePatch(rgbaOf(region.pixels, region.width, region.height, 1), region.width, region.height, patch)
    // Per word: the filled pixels against the paper in a box around the word.
    let maskedSquares = 0
    let maskedCount = 0
    let paperSquares = 0
    let paperCount = 0
    const margin = Math.round(0.3 * region.fontSize)
    for (const word of result.words) {
      const labels = new Set(word.labels)
      const box = word.box
      const masked = []
      const paper = []
      for (let y = Math.max(0, box.y0 - margin); y < Math.min(region.height, box.y1 + margin); y += 1) {
        for (let x = Math.max(0, box.x0 - margin); x < Math.min(region.width, box.x1 + margin); x += 1) {
          const index = y * region.width + x
          if (labels.has(result.labels[index])) masked.push(composite[index * 4])
          else if (result.labels[index] < 0 && !result.blocked[index]) paper.push(region.pixels[index])
        }
      }
      // Like for like: paper close to white is clipped at 255, so its noise is
      // lopsided and its mean sits below its median; the fill reproduces both.
      const level = median(paper)
      const mean = (values) => values.reduce((sum, value) => sum + value, 0) / values.length
      assert.ok(Math.abs(median(masked) - level) <= 4, `${name} "${word.text}": filled median ${median(masked)} vs paper ${level}`)
      assert.ok(Math.abs(mean(masked) - mean(paper)) <= 4, `${name} "${word.text}": filled mean ${mean(masked).toFixed(1)} vs paper ${mean(paper).toFixed(1)}`)
      for (const value of masked) { maskedSquares += (value - level) ** 2; maskedCount += 1 }
      for (const value of paper) { paperSquares += (value - level) ** 2; paperCount += 1 }
    }
    const ratio = Math.sqrt(maskedSquares / maskedCount) / Math.sqrt(paperSquares / paperCount)
    assert.ok(ratio >= 0.6 && ratio <= 1.6, `${name}: grain ${ratio.toFixed(2)} x the paper's`)
  })

  test(`${name}: nothing outside the grown glyph mask changes, and the neighbouring lines keep every ink pixel`, { timeout: 120_000 }, async () => {
    const { composePatch, compositePatch, rgbaOf, lineFrame, estimateLevels } = await retouchModule()
    const region = lineRegion(await scanOf(name))
    const { result } = await retouchRegion(region)
    const patch = composePatch(result, null)
    // The mask is the target glyphs grown by the dilation radius; the feather adds one pixel (8 neighbours).
    const grown = new Uint8Array(region.width * region.height)
    for (let y = 0; y < region.height; y += 1) {
      for (let x = 0; x < region.width; x += 1) {
        if (result.labels[y * region.width + x] < 0) continue
        for (let dy = -1; dy <= 1; dy += 1) {
          for (let dx = -1; dx <= 1; dx += 1) {
            const xx = x + dx
            const yy = y + dy
            if (xx >= 0 && yy >= 0 && xx < region.width && yy < region.height) grown[yy * region.width + xx] = 1
          }
        }
      }
    }
    for (let y = 0; y < patch.height; y += 1) {
      for (let x = 0; x < patch.width; x += 1) {
        const alpha = patch.rgba[(y * patch.width + x) * 4 + 3]
        if (alpha) assert.equal(grown[(y + patch.y) * region.width + x + patch.x], 1, `alpha ${alpha} outside the mask at ${x + patch.x},${y + patch.y}`)
      }
    }
    const composite = compositePatch(rgbaOf(region.pixels, region.width, region.height, 1), region.width, region.height, patch)
    const frame = lineFrame(region.baseline)
    const levels = estimateLevels(region.pixels)
    const threshold = levels.paper - 0.25 * levels.contrast
    let before = 0
    let after = 0
    for (let y = 0; y < region.height; y += 1) {
      for (let x = 0; x < region.width; x += 1) {
        const v = frame.v(x, y)
        // Above the line's ascenders or below its descenders: the neighbours' band.
        if (v <= region.fontSize * 1.02 && v >= -region.fontSize * 0.36) continue
        const index = y * region.width + x
        if (region.pixels[index] < threshold) before += 1
        if (composite[index * 4] < threshold) after += 1
      }
    }
    assert.ok(before > 50, `${name}: the region reaches into the neighbouring lines (${before} ink pixels)`)
    assert.equal(after, before, `${name}: ink pixels in the neighbouring lines' band`)
  })
}

test('a bilevel scan gets a hard mask and a flat paper fill', { timeout: 60_000 }, async () => {
  const { composePatch } = await retouchModule()
  const region = lineRegion(await scanOf('clean300'), LINE_START, (value) => (value < 128 ? 0 : 255))
  const { result } = await retouchRegion(region)
  assert.equal(result.bilevel, true)
  const patch = composePatch(result, null)
  const alphas = new Set()
  const fills = new Set()
  for (let index = 0; index < patch.rgba.length; index += 4) {
    alphas.add(patch.rgba[index + 3])
    if (patch.rgba[index + 3]) fills.add(`${patch.rgba[index]},${patch.rgba[index + 1]},${patch.rgba[index + 2]}`)
  }
  assert.deepEqual([...alphas].sort((a, b) => a - b), [0, 255])
  assert.deepEqual([...fills], ['255,255,255'])
  assert.deepEqual(result.noiseSigma, [0, 0, 0])
})

test('a patch for some words leaves the other words untouched; the same key gives the same patch', { timeout: 60_000 }, async () => {
  const { composePatch } = await retouchModule()
  const region = lineRegion(await scanOf('noisy300'))
  const { result } = await retouchRegion(region)
  const index = result.words.findIndex((word) => word.text === '48213')
  const word = result.words[index]
  const patch = composePatch(result, new Set(word.labels))
  const others = new Set(result.words.filter((_, position) => position !== index).flatMap((item) => item.labels))
  let covered = 0
  for (let y = 0; y < patch.height; y += 1) {
    for (let x = 0; x < patch.width; x += 1) {
      if (!patch.rgba[(y * patch.width + x) * 4 + 3]) continue
      const label = result.labels[(y + patch.y) * region.width + x + patch.x]
      assert.ok(!others.has(label), 'the patch reaches another word')
      if (label >= 0) covered += 1
    }
  }
  assert.ok(covered > 200, `the word's glyphs are covered (${covered} px)`)
  // Its box: the word's ink grown by the dilation radius and the feather.
  assert.ok(patch.x >= word.box.x0 - result.radius - 1 && patch.x + patch.width <= word.box.x1 + result.radius + 1)

  const again = await retouchRegion(region)
  assert.deepEqual(composePatch(again.result, new Set(word.labels)).rgba, patch.rgba)
  const other = await retouchRegion(region, { seed: 'another line' })
  assert.notDeepEqual(composePatch(other.result, new Set(word.labels)).rgba, patch.rgba)
})

test('word boxes from recognition restrict the mask to those words', { timeout: 60_000 }, async () => {
  const { composePatch } = await retouchModule()
  const region = lineRegion(await scanOf('noisy300'))
  const { result } = await retouchRegion(region, { targets: region.quads })
  assert.equal(result.words.length, region.line.words.length)
  const patch = composePatch(result, new Set([2]))
  assert.ok(patch.width < region.length / 8, 'one word, not the line')
})

async function createTesseract(directory) {
  const Tesseract = require('tesseract.js')
  // As scripts/test-ocr-accuracy.mjs: pin the SIMD core the app ships.
  const workerPath = path.join(directory, 'tesseract-simd-worker.cjs')
  fs.writeFileSync(workerPath, `const Module = require('node:module')
const resolve = Module._resolveFilename
Module._resolveFilename = function (request, ...rest) {
  return resolve.call(this, request === 'tesseract.js-core/tesseract-core-relaxedsimd-lstm' ? 'tesseract.js-core/tesseract-core-simd-lstm' : request, ...rest)
}
require(${JSON.stringify(require.resolve('tesseract.js/src/worker-script/node/index.js'))})
`)
  return Tesseract.createWorker('eng', Tesseract.OEM.LSTM_ONLY, {
    workerPath,
    langPath: path.join(__dirname, '..', 'node_modules', '@tesseract.js-data', 'eng', '4.0.0_best_int'),
    cacheMethod: 'none',
    gzip: true,
  })
}

test('ghost test: Tesseract finds no word where the line was retouched', { timeout: 240_000 }, async () => {
  const { composePatch, compositePatch, rgbaOf, lineFrame } = await retouchModule()
  const { encodePgm } = await preprocessModule()
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-retouch-ghost-'))
  let worker
  try {
    worker = await createTesseract(directory)
    await worker.setParameters({ user_defined_dpi: String(DPI), tessedit_pageseg_mode: '6' })
    for (const name of ['noisy300', 'hard300']) {
      const region = lineRegion(await scanOf(name))
      const { result } = await retouchRegion(region)
      const composite = compositePatch(rgbaOf(region.pixels, region.width, region.height, 1), region.width, region.height, composePatch(result, null))
      const gray = Uint8Array.from({ length: region.width * region.height }, (_, index) => composite[index * 4])
      const frame = lineFrame(region.baseline)
      const inBand = (bbox) => {
        const x = (bbox.x0 + bbox.x1) / 2
        const y = (bbox.y0 + bbox.y1) / 2
        const v = frame.v(x, y)
        const u = frame.u(x, y)
        return v > -0.3 * region.fontSize && v < 0.9 * region.fontSize && u > 0 && u < region.length
      }
      const wordsOf = async (data) => {
        const { data: page } = await worker.recognize(encodePgm({ width: region.width, height: region.height, data }), {}, { text: false, blocks: true })
        return (page.blocks || []).flatMap((block) => block.paragraphs.flatMap((paragraph) => paragraph.lines.flatMap((line) => line.words)))
      }
      // The untouched line reads well, so the test can see a ghost.
      const original = (await wordsOf(region.pixels)).filter((word) => inBand(word.bbox))
      assert.ok(original.filter((word) => word.confidence > 80).length >= 10, `${name}: the original line reads (${original.map((word) => word.text).join(' ')})`)
      const ghosts = (await wordsOf(gray)).filter((word) => inBand(word.bbox) && word.confidence > 50)
      assert.deepEqual(ghosts.map((word) => `${word.text} (${Math.round(word.confidence)})`), [], `${name}: words left where the line was`)
    }
  } finally {
    await worker?.terminate()
    fs.rmSync(directory, { recursive: true, force: true })
  }
})
