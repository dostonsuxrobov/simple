const test = require('node:test')
const assert = require('node:assert/strict')
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib')
const fixture = require('./fixtures/scan-fixture.cjs')

// Font style of scanned lines measured from pixels (design 4.8.3 / 6.2(5)):
// lines rendered at 300 DPI in Times, Helvetica and Courier, regular and
// bold, at 9, 11 and 14 pt, each between two neighbouring lines.
const styleModule = () => import('../electron/scan-style.mjs')

const DPI = 300
const SCALE = DPI / 72
const PAGE = [612, 220]
const LINES = [
  'The committee reviewed the proposed budget for the third quarter',
  'Invoice number 48213 was paid in full on July 14. The remaining',
  'Several members noted that the warehouse lease expires on',
]
const FILLER = 'Ms. Alvarez will prepare a summary of vendor quotes'
const FONTS = [
  { font: 'TimesRoman', fontClass: 'serif', weight: 400 },
  { font: 'TimesRomanBold', fontClass: 'serif', weight: 700 },
  { font: 'Helvetica', fontClass: 'sans', weight: 400 },
  { font: 'HelveticaBold', fontClass: 'sans', weight: 700 },
  { font: 'Courier', fontClass: 'mono', weight: 400 },
  { font: 'CourierBold', fontClass: 'mono', weight: 700 },
]
const SIZES = [9, 11, 14]

/** One page per font and size: a filler line, the three sample lines and another filler, 1.25 em apart. */
async function renderPage(fontName, size, color = [0.08, 0.08, 0.1]) {
  const doc = await PDFDocument.create()
  const page = doc.addPage(PAGE)
  const font = await doc.embedFont(StandardFonts[fontName])
  const lines = []
  let baseline = 175
  for (const text of [FILLER, ...LINES, FILLER]) {
    page.drawText(text, { x: 60, y: baseline, size, font, color: rgb(...color) })
    lines.push({ text, baseline, width: font.widthOfTextAtSize(text, size) })
    baseline -= size * 1.25
  }
  return { bytes: await doc.save(), lines: lines.slice(1, -1) }
}

/** The pixels around one line and its baseline in them (what scanEdit hands the worker). */
function lineRegion(raster, line, size) {
  const x0 = Math.floor((60 - size * 0.3) * SCALE)
  const x1 = Math.ceil((60 + line.width + size * 0.3) * SCALE)
  const y0 = Math.floor((PAGE[1] - (line.baseline + size * 1.2)) * SCALE)
  const y1 = Math.ceil((PAGE[1] - (line.baseline - size * 0.6)) * SCALE)
  const width = x1 - x0
  const height = y1 - y0
  const channels = raster.channels
  const pixels = new Uint8Array(width * height * channels)
  for (let y = 0; y < height; y += 1) {
    pixels.set(raster.data.subarray(((y + y0) * raster.width + x0) * channels, ((y + y0) * raster.width + x1) * channels), y * width * channels)
  }
  return {
    pixels, width, height, channels,
    baseline: { x: 60 * SCALE - x0, y: (PAGE[1] - line.baseline) * SCALE - y0, dx: 1, dy: 0 },
    length: line.width * SCALE,
  }
}

async function measureAll({ degrade } = {}) {
  const { estimateScanStyle } = await styleModule()
  const results = []
  for (const { font, fontClass, weight } of FONTS) {
    for (const size of SIZES) {
      const page = await renderPage(font, size)
      let raster = await fixture.rasterize(page.bytes, DPI)
      if (degrade) raster = fixture.degrade(raster, { ...degrade, seed: size * 7 + font.length })
      page.lines.forEach((line, index) => {
        const region = lineRegion(raster, line, size)
        // The OCR layer's font size is Tesseract's row height: close, not exact.
        const hint = size * (0.88 + 0.12 * index)
        const { style } = estimateScanStyle(region.pixels, region.width, region.height, {
          channels: region.channels, dpi: DPI, baseline: region.baseline, length: region.length, fontSize: hint * SCALE, text: line.text,
        })
        results.push({ font, size, line: index, fontClass, weight, style })
      })
    }
  }
  return results
}

function summarize(results) {
  const share = (predicate) => results.filter(predicate).length / results.length
  return {
    lines: results.length,
    fontClass: share((r) => r.style.fontClass === r.fontClass),
    weight: share((r) => r.style.fontWeight === r.weight),
    size: share((r) => Math.abs(r.style.fontSize - r.size) / r.size <= 0.08),
    worstSize: Math.max(...results.map((r) => Math.abs(r.style.fontSize - r.size) / r.size)),
    misses: results.filter((r) => r.style.fontClass !== r.fontClass || r.style.fontWeight !== r.weight || Math.abs(r.style.fontSize - r.size) / r.size > 0.08)
      .map((r) => `${r.font} ${r.size}pt line ${r.line}: ${r.style.fontClass} ${r.style.fontWeight} ${r.style.fontSize}pt`),
  }
}

test('class, weight and size of clean 300 DPI lines', { timeout: 180_000 }, async () => {
  const summary = summarize(await measureAll())
  assert.ok(summary.fontClass >= 0.95, `class ${summary.fontClass}: ${summary.misses.join('; ')}`)
  assert.ok(summary.weight >= 0.95, `weight ${summary.weight}: ${summary.misses.join('; ')}`)
  // Every size within 8 %.
  assert.equal(summary.size, 1, `size: ${summary.misses.join('; ')}`)
})

test('class, weight and size survive scanner grain, specks and uneven light', { timeout: 180_000 }, async () => {
  const summary = summarize(await measureAll({ degrade: { noise: 14, specks: 0.0005, gradient: true } }))
  assert.ok(summary.fontClass >= 0.95, `class ${summary.fontClass}: ${summary.misses.join('; ')}`)
  assert.ok(summary.weight >= 0.95, `weight ${summary.weight}: ${summary.misses.join('; ')}`)
  assert.ok(summary.size >= 0.95, `size ${summary.size}: ${summary.misses.join('; ')}`)
})

test('the stand-in family follows the class, and the ink and paper colours are measured', { timeout: 60_000 }, async () => {
  const { estimateScanStyle, SCAN_FONT_FAMILIES } = await styleModule()
  const blue = [0.1, 0.2, 0.6]
  for (const { font, fontClass } of [FONTS[0], FONTS[2], FONTS[4]]) {
    const page = await renderPage(font, 11, blue)
    const raster = await fixture.rasterize(page.bytes, DPI, { color: true })
    const region = lineRegion(raster, page.lines[1], 11)
    const { style } = estimateScanStyle(region.pixels, region.width, region.height, {
      channels: 3, dpi: DPI, baseline: region.baseline, length: region.length, fontSize: 11 * SCALE, text: page.lines[1].text,
    })
    assert.equal(style.fontClass, fontClass, font)
    assert.equal(style.fontFamily, SCAN_FONT_FAMILIES[fontClass])
    style.color.forEach((channel, index) => assert.ok(Math.abs(channel - blue[index]) <= 0.06, `${font} ink ${style.color} vs ${blue}`))
    style.background.forEach((channel) => assert.ok(channel >= 0.97, `${font} paper ${style.background}`))
  }
})

test('upright text is not taken for italic, slanted text is', { timeout: 60_000 }, async () => {
  const { estimateScanStyle } = await styleModule()
  for (const [font, italic] of [['Helvetica', false], ['HelveticaOblique', true], ['TimesRoman', false], ['TimesRomanItalic', true]]) {
    const page = await renderPage(font, 11)
    const raster = await fixture.rasterize(page.bytes, DPI)
    const region = lineRegion(raster, page.lines[0], 11)
    const { style } = estimateScanStyle(region.pixels, region.width, region.height, {
      channels: 1, dpi: DPI, baseline: region.baseline, length: region.length, fontSize: 11 * SCALE, text: page.lines[0].text,
    })
    assert.equal(style.italic, italic, font)
  }
})
