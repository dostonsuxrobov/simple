// OCR accuracy and position check (WP1): synthetic scans -> ocr-preprocess ->
// tesseract.js (Node, same pinned core and model as the app) -> ocr-geometry ->
// comparison with the laid-out ground truth.
//
//   node scripts/test-ocr-accuracy.mjs [variant,variant,...]
//
// When electron/ocr-text-layer.cjs (WP2) is present, the result is also written
// as an invisible text layer and searched with pdf.js.
import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import { existsSync } from 'node:fs'
import { createRequire } from 'node:module'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const require = createRequire(import.meta.url)
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const fixture = require('../tests/fixtures/scan-fixture.cjs')
const Tesseract = require('tesseract.js')
const preprocess = await import('../electron/ocr-preprocess.mjs')
const geometry = await import('../electron/ocr-geometry.mjs')

// Design 6.3: >= 98% words, >= 97% on the stained low-contrast page. The two
// sideways scans go through the orientation fallback (the same policy as
// recognizePage) and must end up corrected by the expected rotation.
const TARGETS = { clean300: 0.98, clean200: 0.98, low150: 0.98, noisy300: 0.98, skew4: 0.98, hard300: 0.97, rotated90Crop: 0.98, color300: 0.98, sideways300: 0.98, sidewaysCw300: 0.98 }
const EXPECTED_CORRECTION = { sideways300: 90, sidewaysCw300: 270 }
const MAX_CENTRE_ERROR_PT = 1.5
const MAX_BASELINE_ERROR_PT = 1.0
const MAX_LINE_EXTENT_ERROR = 0.02
const SEARCH_TERMS = ['negotiations', '7741-0093', '$1,250.75']

const requested = (process.argv[2] ?? Object.keys(TARGETS).join(',')).split(',').map((name) => name.trim()).filter(Boolean)
for (const name of requested) assert.ok(TARGETS[name], `Unknown variant ${name}`)

/** renderForOcr's chooseOcrDpi(): scans between 280 and 400 DPI keep their resolution. */
const chooseOcrDpi = (nativeDpi) => (nativeDpi >= 280 && nativeDpi <= 400 ? nativeDpi : 300)
const comparable = (text) => text.replace(/[“”]/g, '"').replace(/[‘’]/g, "'").replace(/[–—]/g, '-')

function alignWords(truth, recognised) {
  const rows = truth.length + 1
  const columns = recognised.length + 1
  const cost = new Uint32Array(rows * columns)
  for (let i = 0; i < rows; i += 1) cost[i * columns] = i
  for (let j = 0; j < columns; j += 1) cost[j] = j
  for (let i = 1; i < rows; i += 1) {
    for (let j = 1; j < columns; j += 1) {
      const substitute = cost[(i - 1) * columns + j - 1] + (truth[i - 1] === recognised[j - 1] ? 0 : 1)
      cost[i * columns + j] = Math.min(cost[(i - 1) * columns + j] + 1, cost[i * columns + j - 1] + 1, substitute)
    }
  }
  const matches = []
  const edits = []
  let i = truth.length
  let j = recognised.length
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && cost[i * columns + j] === cost[(i - 1) * columns + j - 1] + (truth[i - 1] === recognised[j - 1] ? 0 : 1)) {
      if (truth[i - 1] === recognised[j - 1]) matches.push([i - 1, j - 1])
      else edits.push(`${truth[i - 1]} -> ${recognised[j - 1]}`)
      i -= 1
      j -= 1
    } else if (i > 0 && cost[i * columns + j] === cost[(i - 1) * columns + j] + 1) {
      edits.push(`missing ${truth[i - 1]}`)
      i -= 1
    } else {
      edits.push(`extra ${recognised[j - 1]}`)
      j -= 1
    }
  }
  return { errors: cost[truth.length * columns + recognised.length], matches: matches.reverse(), edits: edits.reverse() }
}

async function createEngine(temporaryDirectory) {
  // tesseract.js' Node worker picks the relaxed-SIMD core when V8 supports it;
  // the app pins the plain SIMD core, so redirect that one require.
  const workerPath = path.join(temporaryDirectory, 'tesseract-simd-worker.cjs')
  await fs.writeFile(workerPath, `const Module = require('node:module')
const resolve = Module._resolveFilename
Module._resolveFilename = function (request, ...rest) {
  return resolve.call(this, request === 'tesseract.js-core/tesseract-core-relaxedsimd-lstm' ? 'tesseract.js-core/tesseract-core-simd-lstm' : request, ...rest)
}
require(${JSON.stringify(require.resolve('tesseract.js/src/worker-script/node/index.js'))})
`)
  return Tesseract.createWorker('eng', Tesseract.OEM.LSTM_ONLY, {
    workerPath,
    langPath: path.join(root, 'node_modules', '@tesseract.js-data', 'eng', '4.0.0_best_int'),
    cacheMethod: 'none',
    gzip: true,
    errorHandler: (error) => console.error('tesseract:', error),
  })
}

async function loadLayerWriter() {
  const file = path.join(root, 'electron', 'ocr-text-layer.cjs')
  if (!existsSync(file)) return { skipped: 'electron/ocr-text-layer.cjs (WP2) is not present' }
  const writer = require(file)
  if (typeof writer.addGlyphlessFont !== 'function' || typeof writer.addOcrTextLayer !== 'function') {
    return { skipped: 'electron/ocr-text-layer.cjs does not export addGlyphlessFont/addOcrTextLayer yet' }
  }
  const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-ocr-search-'))
  try {
    for (const name of ['search', 'documentSearch']) {
      const ts = (await import('typescript')).default
      const source = await fs.readFile(path.join(root, 'src', 'lib', `${name}.ts`), 'utf8')
      const output = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 } }).outputText
      await fs.writeFile(path.join(temporaryDirectory, `${name}.mjs`), output.replace("from './search'", "from './search.mjs'"))
    }
    const { createDocumentSearch } = await import(pathToFileURL(path.join(temporaryDirectory, 'documentSearch.mjs')).href)
    return { writer, createDocumentSearch, cleanup: () => fs.rm(temporaryDirectory, { recursive: true, force: true }) }
  } catch (error) {
    await fs.rm(temporaryDirectory, { recursive: true, force: true })
    throw error
  }
}

async function searchableCheck(layer, variant, result) {
  const { PDFDocument } = require('pdf-lib')
  const operation = geometry.buildOcrLayerOperation([result], { replaceExisting: false })
  const document = await PDFDocument.load(variant.pdf)
  const op = typeof layer.writer.validateOcrLayerOperation === 'function' ? layer.writer.validateOcrLayerOperation(operation, document.getPageCount()) : operation
  const fontRef = layer.writer.addGlyphlessFont(document)
  for (const page of op.pages) layer.writer.addOcrTextLayer(document, page.pageIndex, page.lines, { fontRef, meta: op.meta })
  const bytes = await document.save({ useObjectStreams: true })
  const pdf = await pdfjs.getDocument({ data: bytes, disableWorker: true, isEvalSupported: false, verbosity: 0 }).promise
  try {
    const search = layer.createDocumentSearch(async (proxy, index) => {
      const page = await proxy.getPage(index + 1)
      const content = await page.getTextContent()
      return content.items.map((item) => ('str' in item ? item.str + (item.hasEOL ? '\n' : ' ') : '')).join('')
    })
    const found = {}
    for (const term of SEARCH_TERMS) found[term] = (await search(pdf, term)).reduce((sum, match) => sum + match.count, 0)
    return found
  } finally {
    await pdf.destroy()
  }
}

const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-ocr-accuracy-'))
let worker
let layer
const report = []
try {
  layer = await loadLayerWriter()
  const started = performance.now()
  worker = await createEngine(temporaryDirectory)
  const initMs = Math.round(performance.now() - started)
  console.log(`tesseract.js ${require('tesseract.js/package.json').version}, eng best_int, SIMD core; worker ready in ${initMs} ms`)
  if (layer.skipped) console.log(`searchable-layer stage skipped: ${layer.skipped}`)

  // One recognition pass the way recognizePage() does it, at an extra
  // orientation correction (the render turned clockwise, as pdf.js rotation).
  const recognizeAt = async (name, variant, dpi, correction) => {
    const pdf = await pdfjs.getDocument({ data: new Uint8Array(variant.pdf), disableWorker: true, isEvalSupported: false, verbosity: 0 }).promise
    let toPdf
    let rotation
    let pdfFromPixel
    let raster
    try {
      const page = await pdf.getPage(1)
      rotation = (page.rotate + correction) % 360
      const viewport = page.getViewport({ scale: dpi / 72, rotation })
      raster = await fixture.rasterize(variant.pdf, dpi, { color: true })
      for (let turn = 0; turn < correction / 90; turn += 1) raster = fixture.rotateRaster90cw(raster)
      assert.ok(Math.abs(viewport.width - raster.width) <= 1 && Math.abs(viewport.height - raster.height) <= 1, `${name}: pdf.js viewport ${viewport.width}x${viewport.height} vs raster ${raster.width}x${raster.height}`)
      const sx = viewport.width / raster.width
      const sy = viewport.height / raster.height
      toPdf = (x, y) => {
        const [px, py] = viewport.convertToPdfPoint(x * sx, y * sy)
        return { x: px, y: py }
      }
      const o = toPdf(0, 0)
      const ex = toPdf(1, 0)
      const ey = toPdf(0, 1)
      pdfFromPixel = [ex.x - o.x, ex.y - o.y, ey.x - o.x, ey.y - o.y, o.x, o.y]
    } finally {
      await pdf.destroy()
    }
    let lap = performance.now()
    const prepared = preprocess.preparePage(raster.data, raster.width, raster.height, { channels: raster.channels, dpi })
    const prepareMs = Math.round(performance.now() - lap)
    const pgm = preprocess.encodePgm(prepared.image)
    const contentKey = `${await preprocess.sha256Hex(pgm)}|${preprocess.engineSignature()}|${pdfFromPixel.map((value) => Math.round(value * 1e6) / 1e6).join(',')}`
    assert.equal(prepared.blank, false, `${name}: page judged blank`)
    lap = performance.now()
    await worker.setParameters(preprocess.tesseractParameters({ dpi, psm: '3' }))
    const { data } = await worker.recognize(Buffer.from(pgm), {}, { text: false, blocks: true })
    const recognizeMs = Math.round(performance.now() - lap)
    const result = geometry.tesseractToOcrPage(data.blocks, { toPdf, dpi, pageIndex: 0, contentKey, rotation, deskew: prepared.deskew, language: 'eng', orientationCorrectedBy: correction })
    const stats = geometry.tesseractReadingStats(data.blocks)
    const reading = { tesseractConfidence: data.confidence, meanConfidence: result.meanConfidence, wordCount: result.wordCount, lines: stats.lines, verticalLines: stats.verticalLines }
    return { result, reading, prepared, prepareMs, recognizeMs }
  }

  for (const name of requested) {
    const buildStarted = performance.now()
    const variant = await fixture.buildScanVariant(name)
    const buildMs = Math.round(performance.now() - buildStarted)
    const dpi = chooseOcrDpi(variant.nativeDpi)

    let pass = await recognizeAt(name, variant, dpi, 0)
    const { prepareMs, recognizeMs } = pass
    let probes = null
    if (geometry.needsOrientationCheck(pass.reading)) {
      probes = []
      for (const correction of [90, 180, 270]) probes.push({ correction, reading: (await recognizeAt(name, variant, 150, correction)).reading })
      const best = geometry.pickOrientation(pass.reading, probes)
      if (best) {
        const corrected = await recognizeAt(name, variant, dpi, best.correction)
        if (geometry.readingScore(corrected.reading) >= geometry.readingScore(pass.reading)) pass = corrected
      }
    }
    const { result, prepared } = pass
    const ocrWords = result.paragraphs.flatMap((paragraph) => paragraph.lines.flatMap((line) => line.words.map((word) => ({ ...word, line }))))
    const truthWords = variant.truth.words
    const alignment = alignWords(truthWords.map((word) => word.text), ocrWords.map((word) => comparable(word.text)))
    const accuracy = 1 - alignment.errors / truthWords.length

    let maxCentre = 0
    let maxBaseline = 0
    let centreSum = 0
    for (const [truthIndex, ocrIndex] of alignment.matches) {
      const truth = truthWords[truthIndex]
      const word = ocrWords[ocrIndex]
      const centre = { x: (word.quad[0].x + word.quad[2].x) / 2, y: (word.quad[0].y + word.quad[2].y) / 2 }
      const centreError = Math.hypot(centre.x - truth.centre.x, centre.y - truth.centre.y)
      const { origin, dir } = truth.baseline
      const baselineError = Math.abs((word.origin.x - origin.x) * -dir.y + (word.origin.y - origin.y) * dir.x)
      maxCentre = Math.max(maxCentre, centreError)
      maxBaseline = Math.max(maxBaseline, baselineError)
      centreSum += centreError
    }

    // Line extents: OCR lines that hold exactly one truth line's words.
    let maxExtentError = 0
    let linesCompared = 0
    const truthLineOf = new Map(alignment.matches.map(([truthIndex, ocrIndex]) => [ocrIndex, truthWords[truthIndex].lineIndex]))
    const ocrIndexOf = new Map(ocrWords.map((word, index) => [word.quad, index]))
    for (const line of result.paragraphs.flatMap((paragraph) => paragraph.lines)) {
      const truthLines = new Set(line.words.map((word) => truthLineOf.get(ocrIndexOf.get(word.quad))))
      if (truthLines.size !== 1 || truthLines.has(undefined)) continue
      const truthLine = variant.truth.lines[[...truthLines][0]]
      if (truthLine.words.length !== line.words.length) continue
      const first = line.words[0]
      const last = line.words[line.words.length - 1]
      const extent = (last.origin.x - first.origin.x) * first.dir.x + (last.origin.y - first.origin.y) * first.dir.y + last.width
      const inkWidth = (() => {
        const a = truthLine.words[0]
        const b = truthLine.words[truthLine.words.length - 1]
        const along = (point) => point.x * a.baseline.dir.x + point.y * a.baseline.dir.y
        return along(b.centre) + b.inkWidth / 2 - (along(a.centre) - a.inkWidth / 2)
      })()
      maxExtentError = Math.max(maxExtentError, Math.abs(extent - inkWidth) / inkWidth)
      linesCompared += 1
    }

    let search = null
    if (!layer.skipped) search = await searchableCheck(layer, variant, result)

    const row = {
      variant: name,
      dpi,
      words: `${truthWords.length - alignment.errors}/${truthWords.length}`,
      accuracy: Math.round(accuracy * 10000) / 100,
      meanConfidence: Math.round(result.meanConfidence),
      maxCentrePt: Math.round(maxCentre * 100) / 100,
      meanCentrePt: Math.round((centreSum / Math.max(1, alignment.matches.length)) * 100) / 100,
      maxBaselinePt: Math.round(maxBaseline * 100) / 100,
      maxLineExtentPct: Math.round(maxExtentError * 1000) / 10,
      linesCompared,
      deskew: prepared.deskew ? prepared.deskew.degrees : 0,
      denoised: prepared.stats.denoised,
      binarized: prepared.stats.binarized,
      correctedBy: result.orientationCorrectedBy,
      ...(probes ? { probes: probes.map(({ correction, reading }) => `${correction}:${reading.tesseractConfidence}/${reading.wordCount}${reading.verticalLines ? 'v' : ''}`) } : {}),
      buildMs,
      prepareMs,
      recognizeMs,
      edits: alignment.edits.slice(0, 6),
      ...(search ? { search } : {}),
    }
    report.push(row)
    const orientation = probes ? `  orientation: corrected by ${row.correctedBy} (probes ${row.probes.join(', ')})` : ''
    console.log(`${name.padEnd(14)} ${row.words.padStart(7)} ${String(row.accuracy).padStart(6)}%  centre<=${row.maxCentrePt.toFixed(2)}pt baseline<=${row.maxBaselinePt.toFixed(2)}pt extent<=${row.maxLineExtentPct}% (${linesCompared} lines)  prep ${prepareMs} ms, ocr ${recognizeMs} ms${orientation}${row.edits.length ? `  [${row.edits.join(' | ')}]` : ''}${search ? `  search ${JSON.stringify(search)}` : ''}`)
  }

  const failures = []
  for (const row of report) {
    if (row.correctedBy !== (EXPECTED_CORRECTION[row.variant] ?? 0)) failures.push(`${row.variant}: orientation corrected by ${row.correctedBy}, expected ${EXPECTED_CORRECTION[row.variant] ?? 0}`)
    if (row.accuracy / 100 < TARGETS[row.variant]) failures.push(`${row.variant}: ${row.accuracy}% words < ${TARGETS[row.variant] * 100}%`)
    if (row.maxCentrePt > MAX_CENTRE_ERROR_PT) failures.push(`${row.variant}: word centre off by ${row.maxCentrePt} pt`)
    if (row.maxBaselinePt > MAX_BASELINE_ERROR_PT) failures.push(`${row.variant}: baseline off by ${row.maxBaselinePt} pt`)
    if (row.maxLineExtentPct / 100 > MAX_LINE_EXTENT_ERROR) failures.push(`${row.variant}: line extent off by ${row.maxLineExtentPct}%`)
    if (row.search) for (const term of SEARCH_TERMS) if (row.search[term] < 1) failures.push(`${row.variant}: search did not find ${term}`)
  }
  console.log(JSON.stringify({ ok: failures.length === 0, failures, report }))
  if (failures.length) {
    console.error(failures.join('\n'))
    process.exitCode = 1
  }
} finally {
  await worker?.terminate()
  await layer?.cleanup?.()
  await fs.rm(temporaryDirectory, { recursive: true, force: true })
}
