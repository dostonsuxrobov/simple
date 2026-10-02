'use strict'

// Recognize text (OCR, design WP3): page classification (rules, geometry and
// real pdf.js classification of scan, recognised-scan, mixed, born-digital,
// outlined and blank pages), run planning, the completion messages and the
// recognition cache. The renderer modules are transpiled from src/ as is.
const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const ts = require('typescript')
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib')
const fixture = require('./fixtures/scan-fixture.cjs')
const layer = require('../electron/ocr-text-layer.cjs')

const OCR_DIR = path.join(__dirname, '../src/lib/ocr')

function loadModule(file, stubs) {
  const source = fs.readFileSync(file, 'utf8')
  const { outputText } = ts.transpileModule(source, {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022, esModuleInterop: true },
    fileName: file,
  })
  const module = { exports: {} }
  const localRequire = (id) => {
    if (Object.prototype.hasOwnProperty.call(stubs, id)) return stubs[id]
    throw new Error(`${path.basename(file)}: unexpected import ${id}`)
  }
  // eslint-disable-next-line no-new-func
  new Function('require', 'module', 'exports', outputText)(localRequire, module, module.exports)
  return module.exports
}

let modulesPromise
function modules() {
  modulesPromise ??= (async () => {
    const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
    const textAppearance = await import('../electron/text-appearance.mjs')
    const types = loadModule(path.join(OCR_DIR, 'types.ts'), {})
    const classifier = loadModule(path.join(OCR_DIR, 'pageClassifier.ts'), {
      '../../../electron/text-appearance.mjs': textAppearance,
      '../pdf': { getPageTextContent: (page) => page.getTextContent(), pdfjs },
      './types': types,
    })
    const cache = loadModule(path.join(OCR_DIR, 'ocrCache.ts'), {})
    const runner = loadModule(path.join(OCR_DIR, 'ocrRunner.ts'), {
      './ocrCache': cache,
      './pageClassifier': classifier,
      './types': types,
    })
    return { pdfjs, classifier, cache, runner }
  })()
  return modulesPromise
}

async function classifyBytes(bytes) {
  const { pdfjs, classifier } = await modules()
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, disableFontFace: true }).promise
  try {
    return await classifier.classifyPage(await pdf.getPage(1))
  } finally {
    await pdf.destroy()
  }
}

/** A perfect recognition of the fixture written as Simple's invisible layer. */
async function recognisedScan(scan) {
  const lines = scan.truth.lines.map((line) => ({
    fontSize: line.size,
    words: line.words.map((word, index) => {
      const next = line.words[index + 1]
      const gap = next ? Math.max(0, Math.hypot(next.baseline.origin.x - word.baseline.origin.x, next.baseline.origin.y - word.baseline.origin.y) - word.inkWidth) : 0
      return { text: word.text, x: word.baseline.origin.x, y: word.baseline.origin.y, dx: word.baseline.dir.x, dy: word.baseline.dir.y, width: word.inkWidth, gap }
    }),
  }))
  const doc = await PDFDocument.load(scan.pdf)
  const op = layer.validateOcrLayerOperation({ type: 'ocr-text-layer', meta: { engine: 'tesseract.js 7.0.0', language: 'eng' }, pages: [{ pageIndex: 0, lines }] }, 1)
  layer.addOcrTextLayer(doc, 0, op.pages[0].lines, { fontRef: layer.addGlyphlessFont(doc), meta: op.meta })
  return doc.save()
}

const stats = (overrides) => ({ visibleChars: 0, invisibleChars: 0, imageCoverage: 0, hasImages: false, vectorOps: 0, ...overrides })

test('page kinds follow the design rules, at their thresholds', async () => {
  const { classifier } = await modules()
  const kind = (overrides) => classifier.pageKindFromStats(stats(overrides))
  assert.equal(kind({}), 'blank')
  assert.equal(kind({ vectorOps: 4 }), 'blank')
  assert.equal(kind({ vectorOps: 5 }), 'native', 'a page with a few drawn lines is not blank')
  assert.equal(kind({ hasImages: true, imageCoverage: 1 }), 'scan')
  assert.equal(kind({ hasImages: true, imageCoverage: 0.25, visibleChars: 19, invisibleChars: 19 }), 'scan', 'a stamp or page number keeps a scan a scan')
  assert.equal(kind({ hasImages: true, imageCoverage: 0.24 }), 'native', 'a small picture does not make a scan')
  assert.equal(kind({ hasImages: true, imageCoverage: 1, invisibleChars: 20 }), 'searchable-scan')
  assert.equal(kind({ invisibleChars: 400 }), 'searchable-scan', 'recognised text over vector art')
  assert.equal(kind({ hasImages: true, imageCoverage: 1, visibleChars: 20 }), 'mixed')
  assert.equal(kind({ visibleChars: 3, vectorOps: 200 }), 'vector-only')
  assert.equal(kind({ visibleChars: 3, vectorOps: 199 }), 'native')
  assert.equal(kind({ visibleChars: 900 }), 'native')
  const scan = classifier.scanStateFromStats(stats({ hasImages: true, imageCoverage: 1 }))
  assert.equal(scan.needsOcr, true)
  assert.equal(scan.hasOcrText, false)
  const searchable = classifier.scanStateFromStats(stats({ hasImages: true, imageCoverage: 1, invisibleChars: 300 }), { ocrBySimple: true })
  assert.deepEqual([searchable.needsOcr, searchable.hasOcrText, searchable.ocrBySimple], [false, true, true])
})

test('image coverage is the union of image areas inside the page', async () => {
  const { classifier } = await modules()
  const view = { x: 0, y: 0, width: 612, height: 792 }
  assert.equal(classifier.imageCoverage([{ x: 0, y: 0, width: 612, height: 792 }], view), 1)
  assert.equal(classifier.imageCoverage([], view), 0)
  assert.ok(Math.abs(classifier.imageCoverage([{ x: 0, y: 0, width: 306, height: 792 }], view) - 0.5) < 0.02)
  const overlapping = classifier.imageCoverage([{ x: 0, y: 0, width: 306, height: 792 }, { x: 150, y: 0, width: 306, height: 792 }], view)
  assert.ok(Math.abs(overlapping - 456 / 612) < 0.03, `overlapping images count once (${overlapping})`)
  assert.equal(classifier.imageCoverage([{ x: 700, y: 0, width: 300, height: 300 }], view), 0, 'images off the crop box do not count')
  const cropped = { x: 50, y: 40, width: 792, height: 612 }
  assert.equal(classifier.imageCoverage([{ x: 0, y: 0, width: 900, height: 700 }], cropped), 1)
  assert.equal(classifier.isPageCoveringImage({ x: 0, y: 0, width: 612, height: 700 }, [0, 0, 612, 792]), true)
  assert.equal(classifier.isPageCoveringImage({ x: 0, y: 0, width: 612, height: 600 }, [0, 0, 612, 792]), false)
})

test('image placement, drawn resolution and path counts come from the operator list', async () => {
  const { classifier } = await modules()
  const OPS = { save: 1, restore: 2, transform: 3, paintImageXObject: 4, paintFormXObjectBegin: 5, paintFormXObjectEnd: 6, constructPath: 7, fill: 8, stroke: 9, paintImageMaskXObject: 10 }
  const list = {
    fnArray: [1, 3, 4, 2, 5, 3, 4, 6, 7, 8, 7, 9, 1, 3, 10, 2],
    argsArray: [
      [], [612, 0, 0, 792, 0, 0], ['img_1', 2550, 3300], [],
      [[0.5, 0, 0, 0.5, 0, 0], null], [200, 0, 0, 100, 10, 20], ['img_2', 400, 200], [],
      [], [], [], [],
      [], [0, 100, -50, 0, 300, 300], [{ width: 120, height: 60 }], [],
    ],
  }
  const geometry = classifier.pageGeometry(list, OPS)
  assert.equal(geometry.vectorOps, 4)
  assert.equal(geometry.images.length, 3)
  assert.deepEqual(geometry.images[0].rect, { x: 0, y: 0, width: 612, height: 792 })
  assert.equal(geometry.images[0].dpi, 300)
  assert.deepEqual(geometry.images[1].rect, { x: 5, y: 10, width: 100, height: 50 }, 'nested form transforms apply')
  assert.equal(geometry.images[1].dpi, 288)
  // A rotated placement: 120 px across 100 pt and 60 px across 50 pt.
  assert.deepEqual(geometry.images[2].rect, { x: 250, y: 300, width: 50, height: 100 })
  assert.equal(geometry.images[2].dpi, 86)
})

test('text item bounds follow the baseline, size and direction', async () => {
  const { classifier } = await modules()
  const near = (rect, expected) => Object.entries(expected).every(([key, value]) => Math.abs(rect[key] - value) < 1e-9)
  const flat = classifier.textItemRect({ str: 'Hello', transform: [11, 0, 0, 11, 72, 700], width: 30, height: 11 }, { ascent: 0.8, descent: -0.2 })
  assert.ok(near(flat, { x: 72, y: 697.8, width: 30, height: 11 }), JSON.stringify(flat))
  const turned = classifier.textItemRect({ str: 'Up', transform: [0, 10, -10, 0, 100, 100], width: 20, height: 10 }, { ascent: 0.8, descent: -0.2 })
  assert.ok(Math.abs(turned.x - 92) < 1e-9 && Math.abs(turned.width - 10) < 1e-9 && Math.abs(turned.y - 100) < 1e-9 && Math.abs(turned.height - 20) < 1e-9, JSON.stringify(turned))
  assert.equal(classifier.textItemRect({ str: 'x' }), null)
})

test('pdf.js pages are classified as scan, recognised scan, mixed, born-digital, outlined and blank', { timeout: 120_000 }, async () => {
  const scan = await fixture.buildScanVariant('clean300')
  const scanState = await classifyBytes(scan.pdf)
  assert.equal(scanState.kind, 'scan')
  assert.equal(scanState.needsOcr, true)
  assert.equal(scanState.nativeDpi, 300)
  assert.ok(scanState.imageCoverage > 0.99)

  const recognised = await classifyBytes(await recognisedScan(scan))
  assert.equal(recognised.kind, 'searchable-scan')
  assert.equal(recognised.hasOcrText, true)
  assert.equal(recognised.ocrBySimple, true)
  assert.equal(recognised.visibleChars, 0)
  assert.ok(recognised.invisibleChars > 500, `invisible characters: ${recognised.invisibleChars}`)

  const vector = await fixture.vectorPage()
  const native = await classifyBytes(vector.bytes)
  assert.equal(native.kind, 'native')
  assert.ok(native.visibleChars > 500)
  assert.equal(native.visibleTextRects.length > 10, true)

  const mixedDoc = await PDFDocument.load(scan.pdf)
  const font = await mixedDoc.embedFont(StandardFonts.TimesRoman)
  mixedDoc.getPage(0).drawText('A typed caption over the scanned page', { x: 72, y: 60, size: 11, font, color: rgb(0, 0, 0) })
  const mixed = await classifyBytes(await mixedDoc.save())
  assert.equal(mixed.kind, 'mixed')
  assert.equal(mixed.imageRects.length, 1)
  assert.equal(mixed.visibleTextRects.length, 1)
  const caption = mixed.visibleTextRects[0]
  assert.ok(caption.x >= 71 && caption.x <= 73 && caption.y < 60 && caption.y + caption.height > 60, JSON.stringify(caption))

  const outlinedDoc = await PDFDocument.create()
  const outlined = outlinedDoc.addPage([612, 792])
  for (let index = 0; index < 150; index += 1) outlined.drawRectangle({ x: 40 + (index % 30) * 18, y: 600 - Math.floor(index / 30) * 20, width: 9, height: 12, color: rgb(0, 0, 0) })
  assert.equal((await classifyBytes(await outlinedDoc.save())).kind, 'vector-only')

  const blankDoc = await PDFDocument.create()
  blankDoc.addPage([612, 792])
  const blank = await classifyBytes(await blankDoc.save())
  assert.equal(blank.kind, 'blank')
  assert.equal(blank.needsOcr, false)
})

test('a run reads scans, never doubles recognised text, and reads only pictures on pages with text', async () => {
  const { classifier, runner } = await modules()
  const state = (overrides, extra) => classifier.scanStateFromStats(stats(overrides), extra)
  const scan = state({ hasImages: true, imageCoverage: 1 }, { nativeDpi: 400 })
  const recognised = state({ hasImages: true, imageCoverage: 1, invisibleChars: 300 })
  const image = { x: 0, y: 300, width: 612, height: 492 }
  const caption = { x: 72, y: 60, width: 200, height: 12 }
  const mixed = state({ hasImages: true, imageCoverage: 0.6, visibleChars: 300 }, { imageRects: [image], visibleTextRects: [caption] })
  const native = state({ visibleChars: 3000 }, { visibleTextRects: [caption] })
  const nativeWithPhoto = state({ visibleChars: 3000, hasImages: true, imageCoverage: 0.1 }, { imageRects: [image], visibleTextRects: [caption] })
  const outlined = state({ visibleChars: 3, vectorOps: 2000 }, { visibleTextRects: [caption] })
  const blank = state({})
  const plan = (pageState, explicit, replaceExisting = false) => runner.planOcrPage(7, pageState, { explicit, replaceExisting })

  assert.deepEqual(plan(scan, false), { pageIndex: 7, action: 'recognize', nativeDpi: 400 })
  assert.deepEqual(plan(recognised, false), { pageIndex: 7, action: 'skip', reason: 'has-ocr-text' })
  assert.deepEqual(plan(recognised, true), { pageIndex: 7, action: 'skip', reason: 'has-ocr-text' }, 'even when picked: replacing is a choice')
  assert.equal(plan(recognised, false, true).action, 'recognize')
  assert.deepEqual(plan(mixed, false), { pageIndex: 7, action: 'skip', reason: 'has-text' }, 'mixed pages only on request')
  assert.deepEqual(plan(mixed, true), { pageIndex: 7, action: 'recognize', nativeText: [caption], imageRects: [image] })
  assert.deepEqual(plan(native, true), { pageIndex: 7, action: 'skip', reason: 'has-text' })
  assert.deepEqual(plan(nativeWithPhoto, true), { pageIndex: 7, action: 'recognize', nativeText: [caption], imageRects: [image] })
  assert.deepEqual(plan(outlined, false), { pageIndex: 7, action: 'recognize', nativeText: [caption] }, 'outlined text has no searchable text')
  assert.deepEqual(plan(blank, true), { pageIndex: 7, action: 'skip', reason: 'blank' })
  const stamped = state({ hasImages: true, imageCoverage: 1, visibleChars: 9 }, { visibleTextRects: [caption] })
  assert.deepEqual(plan(stamped, false), { pageIndex: 7, action: 'recognize', nativeText: [caption] }, 'a stamp on a scan is not read twice')
  assert.equal(runner.pageNeedsRecognition(scan, false), true)
  assert.equal(runner.pageNeedsRecognition(recognised, false), false)
  assert.equal(runner.pageNeedsRecognition(recognised, true), true)
  assert.equal(runner.hasRecognizedText([scan, recognised]), true)
  assert.equal(runner.hasRecognizedText([scan, native]), false)

  // The cache key names the filter regions: other regions keep other words.
  assert.equal(runner.regionSignature({}), '')
  const signature = runner.regionSignature({ nativeText: [caption], imageRects: [image] })
  assert.match(signature, /^\|regions:[0-9a-f]{8}$/)
  assert.equal(runner.regionSignature({ nativeText: [{ ...caption }], imageRects: [{ ...image }] }), signature)
  assert.notEqual(runner.regionSignature({ nativeText: [{ ...caption, x: 90 }], imageRects: [image] }), signature)
})

function result(pageIndex, words, overrides = {}) {
  return { schema: 1, pageIndex, contentKey: `key-${pageIndex}`, engine: { name: 'tesseract.js', version: '7.0.0', core: 'simd-lstm', language: 'eng', model: 'best_int' }, rotation: 0, dpi: 300, deskewDegrees: 0, orientationCorrectedBy: 0, meanConfidence: 91, wordCount: words, paragraphs: [], ...overrides }
}

function outcome(overrides = {}) {
  return { results: [], contentKeys: new Map(), skipped: [], failed: [], planned: 0, stopped: false, operation: null, ...overrides }
}

test('completion messages follow the design copy', async () => {
  const { runner } = await modules()
  const summary = (value, applied = true) => runner.summarizeOcrRun(outcome(value), applied)
  assert.equal(summary({ results: [result(0, 120)], planned: 1 }).message, 'Text recognized on 1 page. You can now search, select, and edit it.')
  assert.equal(
    summary({ results: [result(0, 120), result(3, 80, { meanConfidence: 52 }), result(5, 10)], planned: 3 }).message,
    'Text recognized on 3 pages. You can now search, select, and edit it. · some words on page 4 may be inaccurate',
  )
  const sideways = summary({ results: [result(1, 90, { orientationCorrectedBy: 90 })], planned: 1 })
  assert.equal(sideways.message, 'Text recognized on 1 page. You can now search, select, and edit it. · Page 2 looks sideways')
  assert.deepEqual(sideways.sideways, [{ pageIndex: 1, degrees: 90 }])
  assert.equal(summary({ results: [result(0, 50), result(1, 40), result(2, 30), result(3, 20), result(4, 10)], planned: 12, stopped: true }).message, 'Stopped — text added to 5 of 12 pages.')
  assert.equal(summary({ planned: 12, stopped: true }, false).message, 'Stopped — no text was added.')
  assert.equal(summary({ results: [result(0, 0)], planned: 1 }, false).message, 'No text was found on this page.')
  assert.equal(
    summary({ results: [result(0, 40)], failed: [{ pageIndex: 2, message: 'Text recognition failed on this page.' }], planned: 2 }).message,
    'Text recognized on 1 page. You can now search, select, and edit it. · page 3 could not be read',
  )
  assert.equal(summary({ skipped: [{ pageIndex: 0, reason: 'has-ocr-text' }] }, false).message, 'This text was already recognized. Choose “Replace text recognized earlier” to read it again.')
  assert.equal(summary({ skipped: [{ pageIndex: 0, reason: 'has-text' }] }, false).message, 'These pages already have text.')
  assert.equal(runner.describePages([6, 3]), 'pages 4 and 7')
  assert.equal(runner.describePages([0, 1, 2, 3, 4]), 'pages 1, 2, 3 and 2 more')
  assert.equal(runner.formatOcrDuration(runner.estimateOcrSeconds(44, 2)), 'About 40 seconds')
  assert.equal(runner.formatOcrDuration(5), 'A few seconds')
  assert.equal(runner.formatOcrDuration(200), 'About 3 minutes')
  assert.equal(runner.pagesToWrite([result(0, 0), result(1, 4)], false).length, 1, 'pages without words are written only to replace text')
  assert.equal(runner.pagesToWrite([result(0, 0), result(1, 4)], true).length, 2)
})

test('the recognition cache is a bounded LRU keyed by content, indexed per document', async () => {
  const { cache } = await modules()
  cache.clearOcrCache()
  for (let index = 0; index < cache.OCR_CACHE_LIMITS.pages + 5; index += 1) cache.storeOcrResult(`content-${index}`, result(index, 3))
  assert.equal(cache.ocrCacheStats().pages, cache.OCR_CACHE_LIMITS.pages)
  assert.equal(cache.getCachedOcrResult('content-0'), undefined, 'the oldest results go first')
  assert.equal(cache.getCachedOcrResult('content-6')?.pageIndex, 6)
  cache.storeOcrResult('content-new', result(9, 3))
  assert.ok(cache.getCachedOcrResult('content-6'), 'reading a result keeps it')
  assert.equal(cache.getCachedOcrResult('content-5'), undefined, 'the least recently used result made room')
  assert.ok(cache.getCachedOcrResult('content-7'))

  const documentA = {}
  const documentB = {}
  cache.bindOcrPage(documentA, 2, 'content-new')
  assert.equal(cache.ocrResultForPage(documentA, 2)?.pageIndex, 2, 'a result is rebound to the page it now sits on')
  assert.equal(cache.ocrResultForPage(documentA, 9), undefined)
  assert.equal(cache.ocrResultForPage(documentB, 2), undefined, 'each document has its own index')
  cache.clearOcrCache()
  assert.equal(cache.ocrResultForPage(documentA, 2), undefined, 'an evicted result is simply missing')
})

test('a text export of a recognised scan carries its text, not the scan or the OCR font', { timeout: 120_000 }, async () => {
  const { pdfjs } = await modules()
  const sourcePath = path.join(__dirname, '../src/lib/pdfExport.ts')
  const { outputText } = ts.transpileModule(fs.readFileSync(sourcePath, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
  })
  const exporter = { exports: {} }
  // eslint-disable-next-line no-new-func
  new Function('module', 'exports', 'require', outputText)(exporter, exporter.exports, require)
  const { extractPdfText, fontInfoFromName } = exporter.exports
  assert.deepEqual(fontInfoFromName('GlyphLessFont'), { bold: false, italic: false, family: '' })
  assert.equal(fontInfoFromName('ABCDEF+Georgia-BoldItalic').family, 'Georgia', 'real fonts keep their family')

  const scan = await fixture.buildScanVariant('clean300')
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await recognisedScan(scan)), isEvalSupported: false, disableFontFace: true }).promise
  try {
    const [page] = await extractPdfText(pdf, [0])
    const collapse = (value) => value.replace(/\s+/g, ' ').trim()
    assert.equal(collapse(page.text), collapse(scan.truth.lines.map((line) => line.text).join(' ')))
    const runs = page.blocks.flatMap((block) => block.type === 'table'
      ? block.rows.flat(2).flatMap((line) => line.runs)
      : block.lines ? block.lines.flatMap((line) => line.runs) : [])
    assert.ok(runs.length > 0)
    assert.equal(runs.filter((run) => run.font).length, 0, `OCR text must not name a font: ${JSON.stringify([...new Set(runs.map((run) => run.font))])}`)
    assert.equal(page.blocks.filter((block) => block.type === 'image').length, 0, 'the scan picture under recognised text is not exported again')
  } finally {
    await pdf.destroy()
  }
})
