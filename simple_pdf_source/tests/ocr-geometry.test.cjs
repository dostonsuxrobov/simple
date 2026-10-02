const test = require('node:test')
const assert = require('node:assert/strict')
const { PDFDocument, degrees } = require('pdf-lib')

const load = () => import('../electron/ocr-geometry.mjs')
const DPI = 300
const S = DPI / 72

/** A pdf.js viewport (PageViewport) for a page with a CropBox offset, as renderForOcr builds it. */
async function viewportFor({ rotate = 0, viewRotation = 0, crop = { x: 50, y: 40, width: 612, height: 792 } } = {}) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const doc = await PDFDocument.create()
  const page = doc.addPage([700, 900])
  page.setCropBox(crop.x, crop.y, crop.width, crop.height)
  if (rotate) page.setRotation(degrees(rotate))
  const pdf = await pdfjs.getDocument({ data: await doc.save(), disableWorker: true, isEvalSupported: false, verbosity: 0 }).promise
  try {
    const pdfPage = await pdf.getPage(1)
    const viewport = pdfPage.getViewport({ scale: S, rotation: (pdfPage.rotate + viewRotation) % 360 })
    return (x, y) => {
      const [px, py] = viewport.convertToPdfPoint(x, y)
      return { x: px, y: py }
    }
  } finally {
    await pdf.destroy()
  }
}

const close = (actual, expected, tolerance = 0.01, label = '') => assert.ok(Math.abs(actual - expected) <= tolerance, `${label} ${actual} vs ${expected}`)
const closePoint = (actual, expected, tolerance = 0.01, label = '') => {
  close(actual.x, expected.x, tolerance, `${label} x`)
  close(actual.y, expected.y, tolerance, `${label} y`)
}

const word = (text, x0, y0, x1, y1, confidence = 95) => ({ text, confidence, bbox: { x0, y0, x1, y1 } })
const line = (words, baseline, rowHeight = 60) => ({ words, baseline, rowAttributes: { rowHeight }, bbox: { x0: words[0].bbox.x0, y0: 0, x1: words.at(-1).bbox.x1, y1: 0 } })
const page = (lines, isLtr = 1) => [{ paragraphs: [{ is_ltr: isLtr, lines }] }]
const context = (toPdf, extra = {}) => ({ toPdf, dpi: DPI, pageIndex: 2, contentKey: 'key', rotation: 0, ...extra })

test('words map through /Rotate and CropBox offsets to hand-computed PDF geometry', async () => {
  const { tesseractToOcrPage } = await load()
  // Hand-computed mapping of an OCR pixel (X, Y) for CropBox [50 40 662 832].
  const expected = {
    0: (X, Y) => ({ x: 50 + X / S, y: 832 - Y / S }),
    90: (X, Y) => ({ x: 50 + Y / S, y: 40 + X / S }),
    180: (X, Y) => ({ x: 662 - X / S, y: 40 + Y / S }),
    270: (X, Y) => ({ x: 662 - Y / S, y: 832 - X / S }),
  }
  const directions = { 0: { x: 1, y: 0 }, 90: { x: 0, y: 1 }, 180: { x: -1, y: 0 }, 270: { x: 0, y: -1 } }
  const blocks = page([line([word('Alpha', 300, 560, 500, 610), word('Beta', 540, 555, 700, 600)], { x0: 300, y0: 600, x1: 700, y1: 600 })])
  for (const rotation of [0, 90, 180, 270]) {
    for (const [label, toPdf] of [
      ['/Rotate', await viewportFor({ rotate: rotation })],
      ['view rotation', await viewportFor({ viewRotation: rotation })],
    ]) {
      const result = tesseractToOcrPage(blocks, context(toPdf, { rotation }))
      const [alpha, beta] = result.paragraphs[0].lines[0].words
      const map = expected[rotation]
      closePoint(alpha.origin, map(300, 600), 0.01, `${label} ${rotation} alpha origin`)
      closePoint(beta.origin, map(540, 600), 0.01, `${label} ${rotation} beta origin`)
      assert.deepEqual(alpha.dir, directions[rotation], `${label} ${rotation} direction`)
      close(alpha.width, 200 / S, 0.01, `${label} ${rotation} width`)
      close(alpha.gap, 40 / S, 0.01, `${label} ${rotation} gap`)
      close(beta.width, 160 / S, 0.01)
      assert.equal(beta.gap, 0)
      close(result.paragraphs[0].lines[0].fontSize, 60 / S, 0.01, 'font size')
      closePoint(alpha.quad[0], map(300, 560), 0.01, 'quad top-left')
      closePoint(alpha.quad[2], map(500, 610), 0.01, 'quad bottom-right')
      assert.equal(result.rotation, rotation)
    }
  }
})

test('deskewed OCR coordinates map back to the skewed scan', async () => {
  const { tesseractToOcrPage } = await load()
  const { deskewPoint } = await import('../electron/ocr-preprocess.mjs')
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const deskew = { degrees: 3, cx: 1275, cy: 1650 }
  const blocks = page([line([word('one', 400, 950, 600, 1000), word('two', 650, 950, 900, 1000)], { x0: 400, y0: 1000, x1: 900, y1: 1000 })])
  const result = tesseractToOcrPage(blocks, context(toPdf, { deskew }))
  const [one, two] = result.paragraphs[0].lines[0].words
  const pdfOf = (x, y) => {
    const point = deskewPoint(x, y, deskew)
    return { x: point.x / S, y: 792 - point.y / S }
  }
  closePoint(one.origin, pdfOf(400, 1000), 0.01, 'origin')
  closePoint(two.origin, pdfOf(650, 1000), 0.01, 'second origin')
  const angle = (3 * Math.PI) / 180
  closePoint(one.dir, { x: Math.cos(angle), y: -Math.sin(angle) }, 1e-9, 'direction follows the scan')
  close(one.width, 200 / S, 0.01)
  close(one.gap, 50 / S, 0.01)
  assert.equal(result.deskewDegrees, 3)
  closePoint(one.quad[3], pdfOf(400, 1000), 0.01, 'quad bottom-left')
})

test('baselines within 0.5 degree of the page axis snap; each word keeps its own baseline height', async () => {
  const { tesseractToOcrPage } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const sloped = (degreesOfSlope) => {
    const slope = Math.tan((degreesOfSlope * Math.PI) / 180)
    const at = (x) => 1000 + slope * (x - 300)
    return {
      at,
      blocks: page([line([word('first', 300, 950, 500, at(500)), word('second', 550, 950, 800, at(800)), word('third', 850, 950, 1300, at(1300))], { x0: 300, y0: at(300), x1: 1300, y1: at(1300) })]),
    }
  }
  const snapped = sloped(0.4)
  const result = tesseractToOcrPage(snapped.blocks, context(toPdf))
  const words = result.paragraphs[0].lines[0].words
  for (const item of words) assert.deepEqual(item.dir, { x: 1, y: 0 })
  close(words[0].origin.y, 792 - snapped.at(300) / S, 0.01)
  close(words[2].origin.y, 792 - snapped.at(850) / S, 0.01)
  // 0.4 degrees over 550 px is 3.84 px, 0.92 pt.
  close(words[0].origin.y - words[2].origin.y, (550 * Math.tan((0.4 * Math.PI) / 180)) / S, 0.01, 'words step down with the baseline')
  close(words[1].width, 250 / S, 0.01, 'width along the snapped axis')

  const tilted = tesseractToOcrPage(sloped(0.6).blocks, context(toPdf))
  const dir = tilted.paragraphs[0].lines[0].words[0].dir
  close(Math.atan2(-dir.y, dir.x) * 180 / Math.PI, 0.6, 1e-6, 'not snapped')
})

test('glyph bottoms of baseline letters correct a low or tilted Tesseract baseline', async () => {
  const { tesseractToOcrPage } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  // Letters sit on y = 600; Q's tail and the descenders reach lower. Tesseract reported 604 -> 602.
  const glyphs = (text, x0, bottoms) => text.split('').map((character, index) => ({
    text: character,
    bbox: { x0: x0 + index * 40, y0: 540, x1: x0 + index * 40 + 34, y1: bottoms[character] ?? 600 },
    is_superscript: 0,
    is_subscript: 0,
    is_dropcap: 0,
  }))
  const tails = { Q: 604, y: 616, p: 616 }
  const words = [
    { ...word('Quarterly', 300, 540, 660, 616), symbols: glyphs('Quarterly', 300, tails) },
    { ...word('Report', 700, 540, 940, 616), symbols: glyphs('Report', 700, tails) },
  ]
  const result = tesseractToOcrPage(page([line(words, { x0: 300, y0: 604, x1: 940, y1: 602 })]), context(toPdf))
  const [first, second] = result.paragraphs[0].lines[0].words
  close(first.origin.y, 792 - 600 / S, 0.01, 'first word on the true baseline')
  close(second.origin.y, 792 - 600 / S, 0.01, 'second word on the true baseline')
  assert.deepEqual(first.dir, { x: 1, y: 0 })

  // Too few baseline letters: Tesseract's baseline is kept.
  const short = [{ ...word('yp', 300, 540, 380, 616), symbols: glyphs('yp', 300, tails) }]
  close(tesseractToOcrPage(page([line(short, { x0: 300, y0: 604, x1: 380, y1: 604 })]), context(toPdf)).paragraphs[0].lines[0].words[0].origin.y, 792 - 604 / S, 0.01)
})

test('lines Tesseract read vertically map along their own baseline', async () => {
  const { tesseractToOcrPage, tesseractReadingStats } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  // Text turned 90 degrees clockwise: it runs downwards, its top faces right,
  // the baseline is a vertical line on the left and glyph boxes are empty.
  const empty = { text: 'x', bbox: { x0: 0, y0: 0, x1: 0, y1: 0 }, is_superscript: 0, is_subscript: 0, is_dropcap: 0 }
  const blocks = page([line(
    [{ ...word('Alpha', 1470, 300, 1520, 500), symbols: [empty] }, { ...word('Beta', 1475, 540, 1520, 700), symbols: [empty] }],
    { x0: 1480, y0: 300, x1: 1480, y1: 1300 },
    50,
  )])
  const [alpha, beta] = tesseractToOcrPage(blocks, context(toPdf)).paragraphs[0].lines[0].words
  const pdfOf = (x, y) => ({ x: x / S, y: 792 - y / S })
  assert.deepEqual(alpha.dir, { x: 0, y: -1 })
  closePoint(alpha.origin, pdfOf(1480, 300), 0.01, 'origin on the baseline at the word start')
  closePoint(beta.origin, pdfOf(1480, 540), 0.01)
  close(alpha.width, 200 / S, 0.01)
  close(alpha.gap, 40 / S, 0.01)
  closePoint(alpha.quad[0], pdfOf(1520, 300), 0.01, 'top-left in reading orientation')
  closePoint(alpha.quad[2], pdfOf(1470, 500), 0.01, 'bottom-right in reading orientation')
  close(tesseractToOcrPage(blocks, context(toPdf)).paragraphs[0].lines[0].fontSize, 50 / S, 0.01)

  const horizontal = line([word('flat', 100, 550, 200, 600)], { x0: 100, y0: 600, x1: 200, y1: 600 })
  assert.deepEqual(tesseractReadingStats([...blocks, ...page([horizontal])]), { lines: 2, verticalLines: 1, words: 3 })
  assert.deepEqual(tesseractReadingStats(null), { lines: 0, verticalLines: 0, words: 0 })
})

test('the orientation policy probes weak or vertical readings and prefers horizontal, clearly better ones', async () => {
  const { needsOrientationCheck, pickOrientation, readingScore } = await load()
  // Figures measured on the fixtures: upright 94-95, wrong orientations 30-35.
  const upright = { tesseractConfidence: 95, meanConfidence: 95, wordCount: 127, lines: 13, verticalLines: 0 }
  const junk = { tesseractConfidence: 35, meanConfidence: 65, wordCount: 53, lines: 20, verticalLines: 0 }
  const vertical = { tesseractConfidence: 95, meanConfidence: 95, wordCount: 127, lines: 13, verticalLines: 13 }
  assert.equal(needsOrientationCheck(upright), false)
  assert.equal(needsOrientationCheck(junk), true, 'the kept-word mean (65) alone would miss this')
  assert.equal(needsOrientationCheck(vertical), true)
  assert.equal(needsOrientationCheck({ ...upright, lines: 0, wordCount: 0 }), false, 'nothing recognised: nothing to turn')
  assert.equal(needsOrientationCheck({ ...upright, meanConfidence: 40, wordCount: 12 }), true, 'design rule: few weak words')
  assert.equal(readingScore(vertical), readingScore(upright) / 2)

  const probe = (correction, reading) => ({ correction, reading })
  assert.equal(pickOrientation(junk, [probe(90, upright), probe(180, vertical), probe(270, junk)]).correction, 90)
  assert.equal(pickOrientation(vertical, [probe(90, junk), probe(180, junk), probe(270, { ...upright, tesseractConfidence: 88 })]).correction, 270, 'horizontal within 10 points of a vertical reading')
  assert.equal(pickOrientation(vertical, [probe(90, junk), probe(180, junk), probe(270, { ...upright, tesseractConfidence: 80 })]), null)
  assert.equal(pickOrientation({ ...upright, tesseractConfidence: 58 }, [probe(90, junk), probe(180, junk), probe(270, junk)]), null, 'a weak page that is upright stays')
  assert.equal(pickOrientation(junk, [probe(90, { ...junk, tesseractConfidence: 55 })]), null, 'not 25 points better')
  assert.equal(pickOrientation(junk, [probe(90, { ...upright, wordCount: 0 })]), null)
})

test('a degenerate baseline falls back to the median bottom of words without descenders', async () => {
  const { tesseractToOcrPage } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const blocks = page([{ words: [word('ab', 100, 550, 200, 600), word('gy', 220, 560, 300, 625), word('cd', 320, 552, 420, 600)], baseline: { x0: 0, y0: 0, x1: 0, y1: 0 } }])
  const [first] = tesseractToOcrPage(blocks, context(toPdf)).paragraphs[0].lines[0].words
  close(first.origin.y, 792 - 600 / S, 0.01)
  assert.deepEqual(first.dir, { x: 1, y: 0 })
})

test('confidence filters drop junk words and lines', async () => {
  const { tesseractToOcrPage } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const base = { x0: 100, y0: 600, x1: 1000, y1: 600 }
  const blocks = page([
    line([word('Good', 100, 550, 200, 600, 95), word('ab', 220, 550, 300, 600, 30), word('-', 320, 550, 340, 600, 30), word('Fine', 360, 550, 460, 600, 95), word('x', 480, 550, 500, 600, 10)], base),
    line([word('Okay', 100, 650, 200, 700, 90), word('cd', 220, 650, 300, 700, 30), word('zz', 320, 650, 400, 700, 20)], { ...base, y0: 700, y1: 700 }),
    line([word('noise', 100, 750, 200, 800, 20), word('junk', 220, 750, 300, 800, 22)], { ...base, y0: 800, y1: 800 }),
    line([word('  ', 100, 850, 200, 900, 99), word('Kept', 220, 850, 300, 900, 99)], { ...base, y0: 900, y1: 900 }),
  ])
  const result = tesseractToOcrPage(blocks, context(toPdf))
  const lines = result.paragraphs[0].lines
  // Line 1: mean 53 >= 50 keeps the weak 'ab'; '-' has no letters; 'x' is under 15.
  assert.equal(lines[0].text, 'Good ab Fine')
  // Line 2: mean 46.7 < 50 drops the weak 'cd' and 'zz'.
  assert.equal(lines[1].text, 'Okay')
  // Line 3: mean 21 < 25 is dropped entirely; line 4 ignores the blank word.
  assert.equal(lines[2].text, 'Kept')
  assert.equal(lines.length, 3)
  assert.equal(result.wordCount, 5)
  close(result.meanConfidence, (95 + 30 + 95 + 90 + 99) / 5, 1e-9)
  close(lines[0].words[0].gap, 20 / S, 0.01, 'gap to the next kept word')
  close(lines[0].words[2].gap, 0, 1e-9, 'last kept word')
})

test('mixed pages keep only image words that do not duplicate native text', async () => {
  const { tesseractToOcrPage } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const blocks = page([line([word('native', 300, 550, 500, 600), word('scan', 600, 550, 800, 600), word('outside', 1500, 550, 1800, 600)], { x0: 300, y0: 600, x1: 1800, y1: 600 })])
  const pdfRect = (x0, y0, x1, y1) => ({ x: x0 / S, y: 792 - y1 / S, width: (x1 - x0) / S, height: (y1 - y0) / S })
  const result = tesseractToOcrPage(blocks, context(toPdf, {
    nativeText: [pdfRect(290, 545, 510, 605)],
    imageRects: [pdfRect(0, 0, 1200, 3300)],
  }))
  assert.deepEqual(result.paragraphs[0].lines[0].words.map((item) => item.text), ['scan'])
  const overlapping = tesseractToOcrPage(blocks, context(toPdf, { nativeText: [pdfRect(560, 560, 700, 640)] }))
  assert.deepEqual(overlapping.paragraphs[0].lines[0].words.map((item) => item.text), ['native', 'outside'], 'IoU above 0.3 is a duplicate')
  const none = tesseractToOcrPage(blocks, context(toPdf, { imageRects: [] }))
  assert.equal(none.wordCount, 0, 'an empty image list keeps nothing')
  assert.deepEqual(none.paragraphs, [])
})

test('text is normalised for the GlyphLessFont layer', async () => {
  const { normalizeOcrText, tesseractToOcrPage } = await load()
  assert.equal(normalizeOcrText('ﬁnancial ﬃce ﬀ ﬂ ﬄ ﬅ ﬆ'), 'financial ffice ff fl ffl st st')
  assert.equal(normalizeOcrText('Café'), 'Café')
  assert.equal(normalizeOcrText('a\u{1F600}b'), 'a�b')
  assert.equal(normalizeOcrText('x\uD800y\uDC00z'), 'x�y�z')
  assert.equal(normalizeOcrText('a\u0001b\u007f\u0085c\td'), 'abcd')
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const result = tesseractToOcrPage(page([line([word('ﬁle', 100, 550, 200, 600), word('\u0007', 220, 550, 260, 600)], { x0: 100, y0: 600, x1: 260, y1: 600 })]), context(toPdf))
  assert.deepEqual(result.paragraphs[0].lines[0].words.map((item) => item.text), ['file'])
})

test('font size falls back to the median word height when the row height is implausible', async () => {
  const { tesseractToOcrPage } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const words = [word('tall', 100, 540, 200, 600), word('row', 220, 550, 300, 600)]
  const base = { x0: 100, y0: 600, x1: 300, y1: 600 }
  close(tesseractToOcrPage(page([line(words, base, 500)]), context(toPdf)).paragraphs[0].lines[0].fontSize, 55 / S, 0.01)
  close(tesseractToOcrPage(page([line(words, base, 0)]), context(toPdf)).paragraphs[0].lines[0].fontSize, 55 / S, 0.01)
  close(tesseractToOcrPage(page([line(words, base, 75)]), context(toPdf)).paragraphs[0].lines[0].fontSize, 75 / S, 0.01)
})

test('results carry identity and engine data, and convert to the layer payload', async () => {
  const { tesseractToOcrPage, ocrPageToLayerPayload, buildOcrLayerOperation } = await load()
  const toPdf = await viewportFor({ crop: { x: 0, y: 0, width: 612, height: 792 } })
  const blocks = [
    { paragraphs: [{ is_ltr: 0, lines: [line([word('Alpha', 300, 560, 500, 610, 80), word('Beta', 540, 555, 700, 600, 90)], { x0: 300, y0: 600, x1: 700, y1: 600 })] }] },
    { paragraphs: [{ is_ltr: true, lines: [] }] },
  ]
  const result = tesseractToOcrPage(blocks, context(toPdf, { language: 'eng', orientationCorrectedBy: 90, rotation: 90 }))
  assert.equal(result.schema, 1)
  assert.equal(result.pageIndex, 2)
  assert.equal(result.contentKey, 'key')
  assert.deepEqual(result.engine, { name: 'tesseract.js', version: '7.0.0', core: 'simd-lstm', model: 'best_int', language: 'eng' })
  assert.equal(result.orientationCorrectedBy, 90)
  assert.equal(result.paragraphs.length, 1)
  assert.equal(result.paragraphs[0].ltr, false)
  assert.equal(result.paragraphs[0].lines[0].id, 'ocr-3-l1')
  assert.equal(result.paragraphs[0].id, 'ocr-3-p1')
  const payload = ocrPageToLayerPayload(result)
  assert.equal(payload.pageIndex, 2)
  assert.deepEqual(Object.keys(payload.lines[0].words[0]).sort(), ['dx', 'dy', 'gap', 'text', 'width', 'x', 'y'])
  assert.equal(payload.lines[0].words[0].x, Math.round(result.paragraphs[0].lines[0].words[0].origin.x * 1000) / 1000)
  assert.equal(payload.lines[0].fontSize, Math.round(result.paragraphs[0].lines[0].fontSize * 1000) / 1000)
  const empty = tesseractToOcrPage([], context(toPdf, { pageIndex: 4 }))
  const operation = buildOcrLayerOperation([result, empty], { replaceExisting: true })
  assert.equal(operation.type, 'ocr-text-layer')
  assert.equal(operation.replaceExisting, true)
  assert.deepEqual(operation.meta, { engine: 'tesseract.js 7.0.0', language: 'eng' })
  assert.deepEqual(operation.pages.map((item) => [item.pageIndex, item.lines.length]), [[2, 1], [4, 0]], 'pages without words stay so their old layer is replaced')
  assert.throws(() => tesseractToOcrPage([], { dpi: 300, pageIndex: 0 }), /toPdf/)
  assert.throws(() => tesseractToOcrPage([], context(() => ({ x: 0, y: 0 }))), /degenerate/)
})
