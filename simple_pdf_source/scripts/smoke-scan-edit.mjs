// Edit scanned text (design WP4, 6.4(3)) end to end in the built app:
//   main    a scan recognised by Simple: click "48213", the line opens with its
//           retouch patch and matched font; type 48214 and apply; correct
//           another line's recognised text only; delete a third line; save.
//           The saved file has 48214 once and 48213 nowhere, unchanged lines
//           and words are pixel-identical, the retouched line reads as the new
//           word, the deleted line leaves no ghost, and the reopened file
//           finds 48214.
//   foreign a scan with another tool's invisible GlyphLessFont layer (D2):
//           the editor shows real glyphs and the save draws no double text.
//
//   npm run build:web && node scripts/smoke-scan-edit.mjs [--keep] [--only=main,foreign]
//
// SIMPLE_SCAN_EDIT_DIR sets where fixtures and Electron profiles go (default:
// the system temp folder). Every app instance gets its own profile and
// remote debugging port; every process started is killed (taskkill /T /F).
import assert from 'node:assert/strict'
import { spawn, spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import fs from 'node:fs/promises'
import { createRequire } from 'node:module'
import net from 'node:net'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const require = createRequire(import.meta.url)
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const fixture = require('../tests/fixtures/scan-fixture.cjs')
const layer = require('../electron/ocr-text-layer.cjs')
const { PDFDocument, PDFHexString, beginText, endText, popGraphicsState, pushGraphicsState, setCharacterSqueeze, setFontAndSize, setTextMatrix, setTextRenderingMode, showText, TextRenderingMode } = require('pdf-lib')
const args = process.argv.slice(2)
const keep = args.includes('--keep')
const only = new Set((args.find((arg) => arg.startsWith('--only='))?.slice(7) || 'main,foreign').split(',').filter(Boolean))
const relaxedTiming = process.env.SIMPLE_SCAN_EDIT_RELAXED_TIMING === '1'
const STEP_TIMEOUT_MS = 30_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))
const log = (...parts) => console.error('[smoke-scan-edit]', ...parts.map((part) => (typeof part === 'string' ? part : JSON.stringify(part))))
const collapse = (text) => String(text || '').replace(/\s+/g, ' ').trim()

// ---------------------------------------------------------------------------
// Fixtures

/** Layer words from the fixture's ground truth (what a perfect recognition writes). */
function truthLines(variant) {
  return variant.truth.lines.map((line) => ({
    fontSize: line.size,
    words: line.words.map((word, index) => {
      const next = line.words[index + 1]
      const end = { x: word.baseline.origin.x + word.baseline.dir.x * word.inkWidth, y: word.baseline.origin.y + word.baseline.dir.y * word.inkWidth }
      return {
        text: word.text, x: word.baseline.origin.x, y: word.baseline.origin.y, dx: word.baseline.dir.x, dy: word.baseline.dir.y,
        width: word.inkWidth, gap: next ? Math.hypot(next.baseline.origin.x - end.x, next.baseline.origin.y - end.y) : 0,
      }
    }),
  }))
}

/** A scan with Simple's own searchable layer (as Recognize text writes it). */
async function simpleScan(name) {
  const variant = await fixture.buildScanVariant(name)
  const doc = await PDFDocument.load(variant.pdf)
  const fontRef = layer.addGlyphlessFont(doc)
  layer.addOcrTextLayer(doc, 0, truthLines(variant), { fontRef, meta: { engine: 'tesseract.js 7.0.0', language: 'eng' } })
  return { variant, bytes: await doc.save({ useObjectStreams: true }) }
}

/**
 * A scan with another tool's layer, the way Tesseract's own PDF renderer and
 * OCRmyPDF write it: GlyphLessFont, one text object per word, each word
 * stretched with Tz, no marked content, another producer.
 */
async function foreignScan(name) {
  const variant = await fixture.buildScanVariant(name)
  const doc = await PDFDocument.load(variant.pdf)
  const page = doc.getPage(0)
  const fontRef = layer.addGlyphlessFont(doc)
  const key = page.node.newFontDictionary('f-0-0', fontRef)
  for (const line of truthLines(variant)) {
    for (const word of line.words) {
      const text = layer.normalizeLayerText(word.text)
      const scale = (100 * word.width) / (0.5 * line.fontSize * text.length)
      page.pushOperators(
        pushGraphicsState(), beginText(), setTextRenderingMode(TextRenderingMode.Invisible), setFontAndSize(key, line.fontSize),
        setTextMatrix(word.dx, word.dy, -word.dy, word.dx, word.x, word.y), setCharacterSqueeze(scale),
        showText(PDFHexString.of(layer.utf16Hex(text))), endText(), popGraphicsState(),
      )
    }
  }
  doc.setProducer('Tesseract 5 / OCRmyPDF (test fixture)')
  doc.setCreator('scanner')
  return { variant, bytes: await doc.save({ useObjectStreams: false }) }
}

// ---------------------------------------------------------------------------
// Verification helpers (Node)

async function pageText(bytes) {
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, disableFontFace: true, verbosity: 0 }).promise
  try {
    const content = await (await pdf.getPage(1)).getTextContent()
    return content.items.map((item) => ('str' in item ? item.str : '')).join(' ')
  } finally {
    await pdf.destroy()
  }
}

const occurrences = (text, word) => text.split(word).length - 1

/** Pixels (PDF points, page origin bottom-left) that differ between two 150 DPI renders. */
async function differences(before, after, dpi = 150) {
  const a = await fixture.rasterize(before, dpi)
  const b = await fixture.rasterize(after, dpi)
  const scale = 72 / dpi
  const points = []
  for (let y = 0; y < a.height; y += 1) {
    for (let x = 0; x < a.width; x += 1) {
      if (a.data[y * a.width + x] !== b.data[y * b.width + x]) points.push({ x: (x + 0.5) * scale, y: 792 - (y + 0.5) * scale })
    }
  }
  return points
}

const within = (point, rect, slack = 0) => point.x >= rect.x - slack && point.x <= rect.x + rect.width + slack
  && point.y >= rect.y - slack && point.y <= rect.y + rect.height + slack

async function createTesseract(directory) {
  const Tesseract = require('tesseract.js')
  const workerPath = path.join(directory, 'tesseract-simd-worker.cjs')
  await fs.writeFile(workerPath, `const Module = require('node:module')
const resolve = Module._resolveFilename
Module._resolveFilename = function (request, ...rest) {
  return resolve.call(this, request === 'tesseract.js-core/tesseract-core-relaxedsimd-lstm' ? 'tesseract.js-core/tesseract-core-simd-lstm' : request, ...rest)
}
require(${JSON.stringify(require.resolve('tesseract.js/src/worker-script/node/index.js'))})
`)
  const worker = await Tesseract.createWorker('eng', Tesseract.OEM.LSTM_ONLY, {
    workerPath,
    langPath: path.join(root, 'node_modules', '@tesseract.js-data', 'eng', '4.0.0_best_int'),
    cacheMethod: 'none',
    gzip: true,
  })
  await worker.setParameters({ user_defined_dpi: '300', tessedit_pageseg_mode: '6' })
  return worker
}

/** Words Tesseract reads in a PDF-space rect of the saved page (rendered at 300 DPI). */
async function readRegion(worker, bytes, rect) {
  const { encodePgm } = await import(pathToFileURL(path.join(root, 'electron', 'ocr-preprocess.mjs')).href)
  const raster = await fixture.rasterize(bytes, 300)
  const s = 300 / 72
  const x0 = Math.max(0, Math.floor(rect.x * s))
  const x1 = Math.min(raster.width, Math.ceil((rect.x + rect.width) * s))
  const y0 = Math.max(0, Math.floor((792 - rect.y - rect.height) * s))
  const y1 = Math.min(raster.height, Math.ceil((792 - rect.y) * s))
  const width = x1 - x0
  const height = y1 - y0
  const data = new Uint8Array(width * height)
  for (let y = 0; y < height; y += 1) data.set(raster.data.subarray((y + y0) * raster.width + x0, (y + y0) * raster.width + x1), y * width)
  const { data: page } = await worker.recognize(encodePgm({ width, height, data }), {}, { text: false, blocks: true })
  return (page.blocks || []).flatMap((block) => block.paragraphs.flatMap((paragraph) => paragraph.lines.flatMap((line) => line.words)))
    .map((word) => ({ text: word.text, confidence: word.confidence, centreY: 792 - (y0 + (word.bbox.y0 + word.bbox.y1) / 2) / s }))
}

/** The ink box of a fixture line (PDF points), from its words. */
function lineBox(variant, startsWith) {
  const line = variant.truth.lines.find((item) => item.text.startsWith(startsWith))
  const first = line.words[0]
  const last = line.words.at(-1)
  const size = line.size
  const x0 = first.baseline.origin.x
  const x1 = last.baseline.origin.x + last.inkWidth
  const ys = line.words.map((word) => word.baseline.origin.y)
  return { line, rect: { x: x0 - 0.4 * size, y: Math.min(...ys) - 0.4 * size, width: x1 - x0 + 0.8 * size, height: Math.max(...ys) - Math.min(...ys) + 1.4 * size } }
}

// ---------------------------------------------------------------------------
// The app (Electron with remote debugging), driven over CDP

function electronExecutable() {
  const packaged = process.env.SIMPLE_TEST_EXECUTABLE ? path.resolve(process.env.SIMPLE_TEST_EXECUTABLE) : ''
  if (packaged) return { executable: packaged, appArgs: [] }
  const found = [process.env.ELECTRON_PATH, path.join(root, 'node_modules', 'electron', 'dist', 'electron.exe')].filter(Boolean).find((candidate) => existsSync(candidate))
  if (!found) throw new Error('Electron is not installed (node_modules/electron/dist/electron.exe).')
  return { executable: found, appArgs: ['.'] }
}

function freePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer()
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address()
      server.close(() => resolve(port))
    })
  })
}

async function launch(file, profile) {
  const { executable, appArgs } = electronExecutable()
  const port = await freePort()
  const app = spawn(executable, [...appArgs, `--remote-debugging-port=${port}`, `--user-data-dir=${profile}`, file], { cwd: root, windowsHide: true, stdio: 'ignore' })
  const kill = () => {
    if (app.exitCode === null) spawnSync('taskkill.exe', ['/pid', String(app.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
  }
  process.once('exit', kill)
  const deadline = Date.now() + 90_000
  let target
  while (!target && Date.now() < deadline) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl && /index\.html/.test(item.url))
    } catch { /* the app is opening its debugging endpoint */ }
    if (!target) await pause(60)
  }
  if (!target) { kill(); throw new Error('The renderer did not start') }
  const socket = new WebSocket(target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  let requestId = 0
  const pending = new Map()
  const runtimeErrors = []
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data))
    if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params.exceptionDetails?.exception?.description || message.params.exceptionDetails?.text)
    const request = pending.get(message.id)
    if (!request) return
    pending.delete(message.id)
    clearTimeout(request.timeout)
    if (message.error) request.reject(new Error(message.error.message))
    else if (message.result?.exceptionDetails) request.reject(new Error(message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text))
    else request.resolve(message.result?.result?.value ?? message.result)
  })
  const call = (method, params = {}, timeoutMs = STEP_TIMEOUT_MS) => new Promise((resolve, reject) => {
    const id = ++requestId
    const timeout = setTimeout(() => { pending.delete(id); reject(new Error(`Timed out: ${method}`)) }, timeoutMs)
    pending.set(id, { resolve, reject, timeout })
    socket.send(JSON.stringify({ id, method, params }))
  })
  await call('Runtime.enable')
  const evaluate = (expression) => call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  const close = async () => {
    try { socket.close() } catch { /* already closed */ }
    app.kill()
    await Promise.race([new Promise((resolve) => app.once('exit', resolve)), pause(2_000)])
    kill()
    process.removeListener('exit', kill)
  }
  return { call, evaluate, close, runtimeErrors }
}

function driver(session) {
  const { call, evaluate } = session
  const surface = `document.querySelector('.continuous-page-slot.is-current .page-surface')`
  const state = () => evaluate(`({
    toast: document.querySelector('.toast')?.textContent || '',
    editor: document.querySelector('.inline-pdf-text-editor')?.value ?? null,
    frame: document.querySelector('.inline-text-frame')?.dataset.scanEdit ?? null,
    font: document.querySelector('[data-scan-matched-font]')?.textContent || '',
    busy: Boolean(document.querySelector('.busy-overlay')),
  })`)
  const waitFor = async (expression, label, milliseconds = STEP_TIMEOUT_MS) => {
    const deadline = Date.now() + milliseconds
    while (Date.now() < deadline) {
      if (await evaluate(`Boolean(${expression})`)) return
      await pause(25)
    }
    throw new Error(`Timed out: ${label}; ${JSON.stringify(await state())}`)
  }
  const key = async (keyName, code, modifiers = 0) => {
    const virtual = ({ Enter: 13, Escape: 27, Delete: 46, Tab: 9 })[keyName] || keyName.toUpperCase().charCodeAt(0)
    const parameters = { key: keyName, code, modifiers, windowsVirtualKeyCode: virtual, nativeVirtualKeyCode: virtual }
    await call('Input.dispatchKeyEvent', { type: 'keyDown', ...parameters })
    await call('Input.dispatchKeyEvent', { type: 'keyUp', ...parameters })
  }
  const clickAt = async (point) => {
    await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: point.x, y: point.y })
    await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'left', buttons: 1, clickCount: 1 })
    await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'left', buttons: 0, clickCount: 1 })
  }
  const clickExpression = async (expression, label) => {
    await waitFor(expression, `click target ${label}`)
    const point = await evaluate(`(() => { const element = ${expression}; element.scrollIntoView({ block: 'nearest', inline: 'nearest' }); const box = element.getBoundingClientRect(); return { x: box.left + box.width / 2, y: box.top + box.height / 2 } })()`)
    await clickAt(point)
  }
  /** PDF point -> window point on the current (unrotated) page. */
  const pagePoint = async (x, y) => {
    const geometry = await evaluate(`(() => { const box = ${surface}.getBoundingClientRect(); return { left: box.left, top: box.top, width: box.width, height: box.height } })()`)
    const scale = geometry.width / 612
    return { x: geometry.left + x * scale, y: geometry.top + (792 - y) * scale, scale, geometry }
  }
  /** A window rect (DOMRect-like) in PDF points. */
  const toPdfRect = async (box) => {
    const { geometry, scale } = await pagePoint(0, 792)
    return { x: (box.left - geometry.left) / scale, y: 792 - (box.top + box.height - geometry.top) / scale, width: box.width / scale, height: box.height / scale }
  }
  const waitForPage = (label) => waitFor(`Boolean(${surface}) && !${surface}.classList.contains('is-page-transitioning') && !document.querySelector('.document-refreshing') && !document.querySelector('.busy-overlay')`, label, 60_000)
  /** Click a fixture word (its ink centre) and wait for the scan edit of its line. */
  const editWord = async (word, lineStart) => {
    const point = await pagePoint(word.centre.x, word.centre.y)
    const started = Date.now()
    await clickAt(point)
    await waitFor(`document.querySelector('.inline-pdf-text-editor')?.value?.startsWith(${JSON.stringify(lineStart)})`, `editor for "${word.text}"`)
    return started
  }
  /** Select `from` inside the editor and type `to` over it. */
  const replaceInEditor = async (from, to) => {
    const ok = await evaluate(`(() => {
      const editor = document.querySelector('.inline-pdf-text-editor')
      const start = editor.value.indexOf(${JSON.stringify(from)})
      if (start < 0) return false
      editor.focus()
      editor.setSelectionRange(start, start + ${from.length})
      return true
    })()`)
    assert.ok(ok, `"${from}" is in the editor`)
    await call('Input.insertText', { text: to })
    await waitFor(`document.querySelector('.inline-pdf-text-editor')?.value?.includes(${JSON.stringify(to)})`, `typed "${to}"`)
  }
  return { state, waitFor, key, clickAt, clickExpression, pagePoint, toPdfRect, waitForPage, editWord, replaceInEditor, surface }
}

// ---------------------------------------------------------------------------
// Scenarios

async function scenarioMain(workDirectory, tesseract) {
  const report = {}
  const scan = await simpleScan('clean300')
  const file = path.join(workDirectory, 'scan-recognised.pdf')
  await fs.writeFile(file, scan.bytes)
  const original = new Uint8Array(scan.bytes)
  const word = scan.variant.truth.words.find((item) => item.text === '48213')
  const correctWord = scan.variant.truth.words.find((item) => item.text === '7741-0093')
  const titleWord = scan.variant.truth.words.find((item) => item.text === 'Operations')
  const session = await launch(file, path.join(workDirectory, 'profile-main'))
  let rects
  try {
    const ui = driver(session)
    await ui.waitForPage('scanned page')
    await ui.clickExpression(`document.querySelector('button[aria-label="Edit text (E)"]')`, 'Edit PDF')
    await ui.waitFor(`document.querySelector('.edit-inspector-empty > strong')?.textContent === 'Scanned page with recognized text'`, 'recognised-scan inspector state')
    await ui.waitFor(`document.querySelectorAll('.continuous-page-slot.is-current [data-pdf-text-invisible="true"]').length > 5`, 'invisible text layer')
    report.ocrSpans = await session.evaluate(`document.querySelectorAll('.continuous-page-slot.is-current [data-pdf-text-ocr="simple"]').length`)

    // 1. Click 48213: the line opens; the patch arrives; the font is matched.
    const clicked = await ui.editWord(word, 'Invoice number 48213')
    await ui.waitFor(`document.querySelector('.inline-text-frame')?.dataset.scanEdit === 'ready' && document.querySelector('[data-scan-patch]')`, 'retouch patch', 15_000)
    report.patchMs = Date.now() - clicked
    report.editorLine = await session.evaluate(`document.querySelector('.inline-pdf-text-editor').value`)
    assert.equal(collapse(report.editorLine), collapse(word && scan.variant.truth.lines[word.lineIndex].text))
    if (!relaxedTiming) assert.ok(report.patchMs <= 1_500, `the patch took ${report.patchMs} ms (target 500 ms)`)
    report.matchedFont = await session.evaluate(`document.querySelector('[data-scan-matched-font]')?.textContent || ''`)
    const match = /^Matched font: (.+) · (Regular|Bold|Italic|Bold Italic) · ([\d.]+) pt$/.exec(report.matchedFont)
    assert.ok(match, `matched font: ${report.matchedFont}`)
    assert.equal(match[1], 'Arial')
    assert.equal(match[2], 'Regular')
    assert.ok(Math.abs(Number(match[3]) - 10.5) / 10.5 <= 0.08, `matched size ${match[3]} pt vs 10.5`)
    report.editorFont = await session.evaluate(`getComputedStyle(document.querySelector('.inline-pdf-text-editor')).fontFamily`)
    assert.ok(!/GlyphLess|g_d\d/i.test(report.editorFont), `editor font ${report.editorFont}`)
    report.badge = await session.evaluate(`document.querySelector('[data-scan-badge]')?.textContent || ''`)
    assert.match(report.badge, /Scanned text/)

    // 2. Type 48214 over 48213 and apply.
    await ui.replaceInEditor('48213', '48214')
    await ui.key('Enter', 'Enter', 2)
    await ui.waitFor(`!document.querySelector('.inline-pdf-text-editor')`, 'edit applied')
    // Where the committed edit draws: its patch and its new word.
    rects = await session.evaluate(`(() => {
      const box = (element) => { const r = element.getBoundingClientRect(); return { left: r.left, top: r.top, width: r.width, height: r.height } }
      const patches = [...document.querySelectorAll('.continuous-page-slot.is-current [data-scan-patch]')].map(box)
      const texts = [...document.querySelectorAll('.continuous-page-slot.is-current .text-overlay.is-replacement')].map((element) => ({ text: element.textContent, ...box(element) }))
      return { patches, texts }
    })()`)
    assert.equal(rects.patches.length, 1, 'one patch: the changed word')
    assert.equal(rects.texts.length, 1)
    assert.equal(collapse(rects.texts[0].text), '48214', 'only the changed word is redrawn')
    report.patchRect = await ui.toPdfRect(rects.patches[0])
    report.textRect = await ui.toPdfRect(rects.texts[0])

    // 3. Correct another line's recognised text only.
    await ui.editWord(correctWord, 'transferred to account 7741-0093')
    await ui.clickExpression(`[...document.querySelectorAll('.edit-inspector button')].find((button) => button.textContent.trim() === 'Recognized text only')`, 'Recognized text only')
    await ui.waitFor(`[...document.querySelectorAll('.edit-inspector button')].find((button) => button.textContent.trim() === 'Recognized text only')?.getAttribute('aria-pressed') === 'true'`, 'mode switched')
    await ui.replaceInEditor('7741-0093', '7741-0094')
    await ui.key('Enter', 'Enter', 2)
    await ui.waitFor(`!document.querySelector('.inline-pdf-text-editor')`, 'correction applied')

    // 4. Delete the title line.
    await ui.editWord(titleWord, 'Quarterly Operations Report')
    await ui.waitFor(`document.querySelector('.inline-text-frame')?.dataset.scanEdit === 'ready'`, 'title prepared', 15_000)
    await ui.clickExpression(`[...document.querySelectorAll('.edit-inspector button')].find((button) => button.textContent.trim() === 'Delete')`, 'Delete')
    await ui.waitFor(`!document.querySelector('.inline-pdf-text-editor')`, 'line deleted')

    // 5. Save in place.
    await session.evaluate(`document.activeElement?.blur?.()`)
    await ui.key('s', 'KeyS', 2)
    await ui.waitFor(`/Saved/.test(document.querySelector('.toast')?.textContent || '')`, 'save toast', 60_000)
    report.saveToast = (await ui.state()).toast
    const errors = session.runtimeErrors.filter((message) => message && !/ResizeObserver loop/.test(message))
    assert.deepEqual(errors, [], 'uncaught renderer errors')
  } finally {
    await session.close()
  }

  // Verify the saved file.
  const saved = new Uint8Array(await fs.readFile(file))
  const text = await pageText(saved)
  report.text = { '48214': occurrences(text, '48214'), '48213': occurrences(text, '48213'), '7741-0094': occurrences(text, '7741-0094'), '7741-0093': occurrences(text, '7741-0093'), Operations: occurrences(text, 'Operations'), Invoice: occurrences(text, 'Invoice') }
  assert.equal(report.text['48214'], 1, 'the new word once')
  assert.equal(report.text['48213'], 0, 'the old word nowhere')
  assert.equal(report.text['7741-0094'], 1, 'the corrected recognised text')
  assert.equal(report.text['7741-0093'], 0)
  assert.equal(report.text.Operations, 0, 'the deleted line')
  assert.equal(report.text.Invoice, 1, 'unchanged words keep their recognised text')

  // Pixel identity: nothing changes outside the changed word and the deleted line.
  const changed = await differences(original, saved)
  const title = lineBox(scan.variant, 'Quarterly')
  const correctedLine = lineBox(scan.variant, 'transferred')
  const slack = 72 / 150
  const outsidePatch = changed.filter((point) => !within(point, report.patchRect, slack) && !within(point, title.rect, slack))
  const outsideWord = outsidePatch.filter((point) => !within(point, report.textRect, slack))
  report.pixels = { changed: changed.length, outsidePatch: outsidePatch.length, outsidePatchAndNewWord: outsideWord.length, inCorrectedLine: changed.filter((point) => within(point, correctedLine.rect, 0)).length }
  log('pixels', report.pixels, { patch: report.patchRect, text: report.textRect })
  assert.equal(report.pixels.outsidePatchAndNewWord, 0, 'pixels changed outside the edited word')
  assert.equal(report.pixels.inCorrectedLine, 0, '"Recognized text only" changed what the line looks like')

  // Tesseract reads the new word where the old one was, and nothing where the title was.
  const invoice = lineBox(scan.variant, 'Invoice')
  const read = await readRegion(tesseract, saved, invoice.rect)
  const found = read.find((item) => item.text.replace(/[^\d]/g, '') === '48214')
  report.ocrNewWord = found ? { text: found.text, confidence: Math.round(found.confidence) } : read.map((item) => item.text).join(' ')
  assert.ok(found && found.confidence >= 80, `Tesseract on the saved line: ${JSON.stringify(report.ocrNewWord)}`)
  const titleRead = await readRegion(tesseract, saved, title.rect)
  const titleBand = { min: title.line.words[0].baseline.origin.y - 0.3 * title.line.size, max: title.line.words[0].baseline.origin.y + title.line.size }
  report.titleGhosts = titleRead.filter((item) => item.confidence > 50 && item.centreY >= titleBand.min && item.centreY <= titleBand.max).map((item) => item.text)
  assert.deepEqual(report.titleGhosts, [], 'words left where the deleted line was')

  // Reopen in Simple: search finds the new word.
  const reopened = await launch(file, path.join(workDirectory, 'profile-reopen'))
  try {
    const ui = driver(reopened)
    await ui.waitForPage('saved page')
    await reopened.evaluate(`document.activeElement?.blur?.()`)
    await ui.key('f', 'KeyF', 2)
    await ui.waitFor(`document.activeElement?.getAttribute('aria-label') === 'Find in document'`, 'search box')
    await reopened.call('Input.insertText', { text: '48214' })
    await ui.key('Enter', 'Enter')
    await ui.waitFor(`/match/.test(document.querySelector('.search-status')?.textContent || '') && !/Searching/.test(document.querySelector('.search-status')?.textContent || '')`, 'search result')
    report.reopenSearch = await reopened.evaluate(`document.querySelector('.search-status')?.textContent || ''`)
    assert.match(report.reopenSearch, /\b1 match/)
  } finally {
    await reopened.close()
  }
  return report
}

async function scenarioForeign(workDirectory) {
  const report = {}
  const scan = await foreignScan('noisy300')
  const file = path.join(workDirectory, 'scan-foreign-ocr.pdf')
  await fs.writeFile(file, scan.bytes)
  const original = new Uint8Array(scan.bytes)
  const word = scan.variant.truth.words.find((item) => item.text === 'warehouse')
  const lineText = scan.variant.truth.lines[word.lineIndex].text
  const session = await launch(file, path.join(workDirectory, 'profile-foreign'))
  let rects
  try {
    const ui = driver(session)
    await ui.waitForPage('scanned page')
    await ui.clickExpression(`document.querySelector('button[aria-label="Edit text (E)"]')`, 'Edit PDF')
    await ui.waitFor(`document.querySelectorAll('.continuous-page-slot.is-current [data-pdf-text-invisible="true"]').length > 5`, 'invisible text layer')
    await ui.editWord(word, lineText.split(' ').slice(0, 3).join(' '))
    // The editor shows the line in a real font, never GlyphLessFont's boxes.
    report.editorFont = await session.evaluate(`getComputedStyle(document.querySelector('.inline-pdf-text-editor')).fontFamily`)
    assert.ok(!/GlyphLess|g_d\d|monospace/i.test(report.editorFont), `editor font ${report.editorFont}`)
    await ui.waitFor(`document.querySelector('.inline-text-frame')?.dataset.scanEdit === 'ready'`, 'retouch patch', 15_000)
    report.matchedFont = await session.evaluate(`document.querySelector('[data-scan-matched-font]')?.textContent || ''`)
    assert.match(report.matchedFont, /^Matched font: Times New Roman/)
    await ui.replaceInEditor('warehouse', 'storehouse')
    await ui.key('Enter', 'Enter', 2)
    await ui.waitFor(`!document.querySelector('.inline-pdf-text-editor')`, 'edit applied')
    rects = await session.evaluate(`(() => {
      const box = (element) => { const r = element.getBoundingClientRect(); return { left: r.left, top: r.top, width: r.width, height: r.height } }
      return {
        patches: [...document.querySelectorAll('.continuous-page-slot.is-current [data-scan-patch]')].map(box),
        texts: [...document.querySelectorAll('.continuous-page-slot.is-current .text-overlay.is-replacement')].map(box),
      }
    })()`)
    report.patchRects = await Promise.all(rects.patches.map((box) => ui.toPdfRect(box)))
    report.textRects = await Promise.all(rects.texts.map((box) => ui.toPdfRect(box)))
    await session.evaluate(`document.activeElement?.blur?.()`)
    await ui.key('s', 'KeyS', 2)
    await ui.waitFor(`/Saved/.test(document.querySelector('.toast')?.textContent || '')`, 'save toast', 60_000)
    const errors = session.runtimeErrors.filter((message) => message && !/ResizeObserver loop/.test(message))
    assert.deepEqual(errors, [], 'uncaught renderer errors')
  } finally {
    await session.close()
  }
  const saved = new Uint8Array(await fs.readFile(file))
  const text = await pageText(saved)
  report.text = { storehouse: occurrences(text, 'storehouse'), warehouse: occurrences(text, 'warehouse') }
  assert.equal(report.text.storehouse, 1)
  assert.equal(report.text.warehouse, 0, 'no double text: the old recognised word is gone')
  const changed = await differences(original, saved)
  const slack = 72 / 150
  const allowed = [...report.patchRects, ...report.textRects]
  report.pixels = { changed: changed.length, outside: changed.filter((point) => !allowed.some((rect) => within(point, rect, slack))).length }
  assert.ok(report.pixels.changed > 0, 'the word was replaced on the page')
  assert.equal(report.pixels.outside, 0, 'pixels changed outside the edited word')
  return report
}

// ---------------------------------------------------------------------------

const workRoot = process.env.SIMPLE_SCAN_EDIT_DIR ? path.resolve(process.env.SIMPLE_SCAN_EDIT_DIR) : os.tmpdir()
await fs.mkdir(workRoot, { recursive: true })
const workDirectory = await fs.mkdtemp(path.join(workRoot, 'simple-scan-edit-'))
const reports = {}
const failures = []
let tesseract
try {
  assert.ok(existsSync(path.join(root, 'dist', 'index.html')), 'dist/index.html is missing; run npm run build:web first')
  tesseract = await createTesseract(workDirectory)
  const scenarios = [['main', () => scenarioMain(workDirectory, tesseract)], ['foreign', () => scenarioForeign(workDirectory)]].filter(([name]) => only.has(name))
  for (const [name, run] of scenarios) {
    let report
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      try {
        report = { ...(await run()), ok: true }
        break
      } catch (error) {
        report = { ok: false, error: String(error?.stack || error).slice(0, 4000) }
        // One retry: the machine may be busy (timeouts).
        if (attempt === 1) log(`${name}: attempt 1 failed, retrying once: ${report.error.split('\n')[0]}`)
      }
    }
    reports[name] = report
    if (!report.ok) failures.push(`${name}: ${report.error}`)
    console.log(`${name}: ${report.ok ? 'ok' : 'FAIL'} ${JSON.stringify(report)}`)
  }
} finally {
  await tesseract?.terminate().catch(() => {})
  if (!keep) await fs.rm(workDirectory, { recursive: true, force: true, maxRetries: 8, retryDelay: 250 })
  else console.log(`kept ${workDirectory}`)
}
console.log(JSON.stringify({ ok: failures.length === 0, failures, reports }))
if (failures.length) process.exitCode = 1
