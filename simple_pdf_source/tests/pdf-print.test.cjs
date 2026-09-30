'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const { EventEmitter } = require('node:events')
const fs = require('node:fs')
const path = require('node:path')
const {
  PDFArray,
  PDFContentStream,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRawStream,
  PDFString,
  decodePDFRawStream,
  degrees,
  rgb,
} = require('pdf-lib')
const {
  calculatePrintPlacement,
  loadPdfViewerForPrint,
  nativePrintOptions,
  normalizePrintLayout,
  preparePrintPdf,
  printWebContentsSilently,
  resolvePrinter,
  visiblePageBox,
} = require('../electron/pdf-print.cjs')

function closeTo(actual, expected, epsilon = 0.02) {
  assert.ok(Math.abs(actual - expected) <= epsilon, `Expected ${actual} to be within ${epsilon} of ${expected}`)
}

function readNumberArray(dictionary, name) {
  const array = dictionary.lookup(PDFName.of(name))
  assert.ok(array instanceof PDFArray)
  return Array.from({ length: array.size() }, (_, index) => array.lookup(index, PDFNumber).asNumber())
}

function decodedContent(stream) {
  if (stream instanceof PDFRawStream) return Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1')
  if (stream instanceof PDFContentStream) return stream.getContentsString()
  throw new Error('Unexpected PDF content stream type.')
}

function operatorNumbers(content, operator, count) {
  const number = '(-?(?:\\d+(?:\\.\\d*)?|\\.\\d+))'
  const match = content.match(new RegExp(`${Array.from({ length: count }, () => number).join('\\s+')}\\s+${operator}(?:\\s|$)`))
  assert.ok(match, `Expected ${operator} in content stream:\n${content}`)
  return match.slice(1).map(Number)
}

function addAnnotation(document, page, name, rect, extra = {}) {
  const annotation = document.context.obj({
    Type: 'Annot',
    Subtype: 'Square',
    Rect: rect,
    NM: PDFString.of(name),
    F: 4,
    ...extra,
  })
  page.node.addAnnot(document.context.register(annotation))
}

async function fixture() {
  const document = await PDFDocument.create()
  const first = document.addPage([300, 600])
  first.drawRectangle({ x: 0, y: 0, width: 300, height: 600, color: rgb(0.9, 0.2, 0.2) })
  const second = document.addPage([400, 250])
  second.setRotation(degrees(90))
  second.drawRectangle({ x: 10, y: 10, width: 120, height: 70, color: rgb(0.2, 0.4, 0.9) })
  return document.save()
}

test('print layout clamps untrusted values and uses stable defaults', () => {
  assert.deepEqual(normalizePrintLayout({ paperSize: 'Poster', marginMode: 'huge', scaleMode: 'magic', customScale: 99 }), {
    paperSize: 'Letter',
    landscape: false,
    marginMode: 'normal',
    scaleMode: 'fit',
    customScale: 4,
  })
})

test('native print options are always silent and omit deviceName for the Windows default', () => {
  const defaults = nativePrintOptions({ paperSize: 'A4', landscape: true, copies: 2, color: false, collate: true })
  assert.equal(defaults.silent, true)
  assert.equal(defaults.pageSize, 'A4')
  assert.equal(defaults.landscape, true)
  assert.equal(defaults.copies, 2)
  assert.equal(defaults.color, false)
  assert.equal(Object.hasOwn(defaults, 'deviceName'), false)

  const named = nativePrintOptions({ paperSize: 'Letter', deviceName: 'Office Printer', duplexMode: 'shortEdge' })
  assert.equal(named.silent, true)
  assert.equal(named.deviceName, 'Office Printer')
  assert.equal(named.duplexMode, 'shortEdge')
})

test('printer resolution keeps Windows default implicit and validates explicit devices', () => {
  const printers = [{ name: 'Office', displayName: 'Office' }, { name: 'Archive', displayName: 'Archive' }]
  assert.deepEqual(resolvePrinter(printers), { name: '', displayName: 'Default Windows printer' })
  assert.equal(resolvePrinter(printers, 'Archive').name, 'Archive')
  assert.throws(() => resolvePrinter([], ''), /No printer is available/)
  assert.throws(() => resolvePrinter(printers, 'Missing'), /no longer available/)
})

test('final print handoff sends silent options and returns driver failure without a native dialog route', async () => {
  let received = null
  const webContents = {
    isDestroyed: () => false,
    print(options, callback) {
      received = options
      callback(false, 'printer offline')
    },
  }
  const result = await printWebContentsSilently(webContents, nativePrintOptions({ paperSize: 'Legal' }))
  assert.equal(received.silent, true)
  assert.equal(Object.hasOwn(received, 'deviceName'), false)
  assert.deepEqual(result, { success: false, failureReason: 'printer offline' })

  const dialogSource = fs.readFileSync(path.join(__dirname, '..', 'src', 'components', 'PrintDialog.tsx'), 'utf8')
  const preloadSource = fs.readFileSync(path.join(__dirname, '..', 'electron', 'preload.cjs'), 'utf8')
  assert.doesNotMatch(dialogSource, /System preview/i)
  assert.doesNotMatch(preloadSource, /['"]pdf:print['"]/) 
})

test('fit, actual size, shrink, and custom scale produce honest placement', () => {
  const fit = calculatePrintPlacement(1000, 500, 0, { paperSize: 'Letter', landscape: false, marginMode: 'normal', scaleMode: 'fit' })
  assert.equal(fit.scale, 0.54)
  assert.equal(fit.cropped, false)

  const actual = calculatePrintPlacement(1000, 500, 0, { paperSize: 'Letter', landscape: false, marginMode: 'normal', scaleMode: 'actual' })
  assert.equal(actual.scale, 1)
  assert.equal(actual.cropped, true)

  const shrinkSmall = calculatePrintPlacement(100, 100, 0, { paperSize: 'Letter', landscape: false, marginMode: 'normal', scaleMode: 'shrink' })
  assert.equal(shrinkSmall.scale, 1)

  const custom = calculatePrintPlacement(100, 100, 0, { paperSize: 'A4', landscape: true, marginMode: 'minimum', scaleMode: 'custom', customScale: 2 })
  assert.equal(custom.scale, 2)
  assert.equal(custom.paperWidth > custom.paperHeight, true)
})

test('rotated pages retain rotation while their physical paper orientation stays exact', () => {
  const placement = calculatePrintPlacement(400, 250, 90, { paperSize: 'Letter', landscape: true, marginMode: 'normal', scaleMode: 'fit' })
  assert.equal(placement.paperWidth, 792)
  assert.equal(placement.paperHeight, 612)
  assert.equal(placement.rawPaperWidth, 612)
  assert.equal(placement.rawPaperHeight, 792)
})

test('visible page geometry matches the renderer CropBox and MediaBox intersection', async () => {
  const document = await PDFDocument.create()
  const page = document.addPage([100, 100])
  page.setMediaBox(10, 20, 100, 100)
  page.setCropBox(0, 50, 80, 100)
  assert.deepEqual(visiblePageBox(page), { x: 10, y: 50, width: 70, height: 70 })

  page.setCropBox(500, 500, 20, 20)
  assert.deepEqual(visiblePageBox(page), { x: 10, y: 20, width: 100, height: 100 })
})

test('print preparation creates selected paper pages without rasterizing source content', async () => {
  const bytes = await fixture()
  const output = await preparePrintPdf(bytes, {
    pageIndices: [1, 0],
    paperSize: 'A4',
    landscape: true,
    marginMode: 'minimum',
    scaleMode: 'fit',
    customScale: 1,
  })
  const reopened = await PDFDocument.load(output)
  assert.equal(reopened.getPageCount(), 2)

  const rotated = reopened.getPage(0)
  assert.equal(rotated.getRotation().angle, 90)
  assert.ok(Math.abs(rotated.getWidth() - 595.28) < 0.02)
  assert.ok(Math.abs(rotated.getHeight() - 841.89) < 0.02)

  const normal = reopened.getPage(1)
  assert.equal(normal.getRotation().angle, 0)
  assert.ok(Math.abs(normal.getWidth() - 841.89) < 0.02)
  assert.ok(Math.abs(normal.getHeight() - 595.28) < 0.02)
  const normalPlacement = calculatePrintPlacement(300, 600, 0, {
    paperSize: 'A4', landscape: true, marginMode: 'minimum', scaleMode: 'fit', customScale: 1,
  })
  const normalMediaBox = normal.getMediaBox()
  assert.equal(normalMediaBox.x, 0)
  assert.equal(normalMediaBox.y, 0)
  closeTo(normalMediaBox.width, normalPlacement.rawPaperWidth)
  closeTo(normalMediaBox.height, normalPlacement.rawPaperHeight)
})

test('print preparation sizes from CropBox, clips vector content, and clips annotations', async () => {
  const document = await PDFDocument.create()
  const page = document.addPage([500, 400])
  page.setMediaBox(20, 30, 500, 400)
  page.setCropBox(120, 110, 200, 100)
  page.drawRectangle({ x: 25, y: 35, width: 80, height: 60, color: rgb(1, 0, 0) })
  page.drawRectangle({ x: 120, y: 110, width: 200, height: 100, color: rgb(0, 0.7, 0.2) })
  addAnnotation(document, page, 'inside', [140, 130, 180, 160], { QuadPoints: [140, 160, 180, 160, 140, 130, 180, 130] })
  addAnnotation(document, page, 'outside', [40, 40, 80, 80])
  addAnnotation(document, page, 'partial-without-appearance', [115, 120, 140, 145])
  const crossingAppearance = document.context.flateStream(Buffer.from('q 0.9 0.1 0.1 rg 0 0 60 60 re f Q'), {
    Type: 'XObject', Subtype: 'Form', FormType: 1, BBox: [0, 0, 60, 60], Matrix: [1, 0, 0, 1, 0, 0], Resources: {},
  })
  const crossingAppearances = document.context.obj({ N: document.context.register(crossingAppearance) })
  addAnnotation(document, page, 'crossing', [300, 180, 360, 240], {
    L: [300, 180, 360, 240],
    AP: crossingAppearances,
  })

  const placement = calculatePrintPlacement(200, 100, 0, {
    paperSize: 'Letter', landscape: false, marginMode: 'normal', scaleMode: 'fit', customScale: 1,
  })
  const output = await preparePrintPdf(await document.save(), {
    paperSize: 'Letter', landscape: false, marginMode: 'normal', scaleMode: 'fit', customScale: 1,
  })
  const reopened = await PDFDocument.load(output)
  const printedPage = reopened.getPage(0)

  assert.deepEqual(printedPage.getMediaBox(), { x: 0, y: 0, width: 612, height: 792 })
  assert.deepEqual(printedPage.getCropBox(), { x: 0, y: 0, width: 612, height: 792 })
  closeTo(placement.scale, 2.7)
  closeTo(placement.contentWidth, 540)
  closeTo(placement.contentHeight, 270)

  const contents = printedPage.node.Contents()
  assert.ok(contents instanceof PDFArray)
  const prefix = decodedContent(contents.lookup(0))
  const clipRectangle = operatorNumbers(prefix, 're', 4)
  const transform = operatorNumbers(prefix, 'cm', 6)
  ;[clipRectangle[0], clipRectangle[1], clipRectangle[2], clipRectangle[3]].forEach((value, index) => {
    closeTo(value, [placement.x, placement.y, placement.contentWidth, placement.contentHeight][index])
  })
  assert.match(prefix, /\bW\s+n\b/)
  ;[transform[0], transform[1], transform[2], transform[3], transform[4], transform[5]].forEach((value, index) => {
    closeTo(value, [placement.scale, 0, 0, placement.scale, placement.x - 120 * placement.scale, placement.y - 110 * placement.scale][index])
  })

  const annotations = printedPage.node.Annots()
  assert.ok(annotations instanceof PDFArray)
  const byName = new Map()
  for (let index = 0; index < annotations.size(); index += 1) {
    const annotation = annotations.lookup(index)
    assert.ok(annotation instanceof PDFDict)
    const name = annotation.lookup(PDFName.of('NM'), PDFString).decodeText()
    byName.set(name, annotation)
  }
  assert.deepEqual([...byName.keys()].sort(), ['crossing', 'inside'])

  const clipRight = placement.x + placement.contentWidth
  const clipTop = placement.y + placement.contentHeight
  const crossingRect = readNumberArray(byName.get('crossing'), 'Rect')
  closeTo(crossingRect[2], clipRight)
  closeTo(crossingRect[3], clipTop)
  const crossingLine = readNumberArray(byName.get('crossing'), 'L')
  assert.ok(crossingLine.every((value, index) => index % 2 === 0 ? value <= clipRight : value <= clipTop))

  const clippedAppearances = byName.get('crossing').lookup(PDFName.of('AP'), PDFDict)
  const clippedNormalAppearance = clippedAppearances.lookup(PDFName.of('N'))
  assert.ok(clippedNormalAppearance instanceof PDFRawStream || clippedNormalAppearance instanceof PDFContentStream)
  const appearanceBounds = readNumberArray(clippedNormalAppearance.dict, 'BBox')
  appearanceBounds.forEach((value, index) => closeTo(value, crossingRect[index]))
  const appearanceContent = decodedContent(clippedNormalAppearance)
  assert.match(appearanceContent, /\bW\s+n\b/)
  assert.match(appearanceContent, /\/OriginalAppearance\s+Do\b/)
  const appearanceTransform = operatorNumbers(appearanceContent, 'cm', 6)
  appearanceTransform.forEach((value, index) => closeTo(value, [2.7, 0, 0, 2.7, 522, 450][index]))
})

test('a rotated cropped page keeps its rotation and uses the matching raw paper axes', async () => {
  const document = await PDFDocument.create()
  const page = document.addPage([500, 400])
  page.setCropBox(50, 60, 200, 100)
  page.setRotation(degrees(90))
  page.drawRectangle({ x: 50, y: 60, width: 200, height: 100, color: rgb(0.1, 0.3, 0.8) })

  const output = await preparePrintPdf(await document.save(), {
    paperSize: 'Letter', landscape: false, marginMode: 'normal', scaleMode: 'fit', customScale: 1,
  })
  const reopened = await PDFDocument.load(output)
  const printedPage = reopened.getPage(0)
  const placement = calculatePrintPlacement(200, 100, 90, {
    paperSize: 'Letter', landscape: false, marginMode: 'normal', scaleMode: 'fit', customScale: 1,
  })

  assert.equal(printedPage.getRotation().angle, 90)
  closeTo(printedPage.getMediaBox().width, placement.rawPaperWidth)
  closeTo(printedPage.getMediaBox().height, placement.rawPaperHeight)
  const prefix = decodedContent(printedPage.node.Contents().lookup(0))
  const clipRectangle = operatorNumbers(prefix, 're', 4)
  clipRectangle.forEach((value, index) => closeTo(value, [placement.x, placement.y, placement.contentWidth, placement.contentHeight][index]))
})

test('print preparation rejects an out-of-range page before writing output', async () => {
  await assert.rejects(
    preparePrintPdf(await fixture(), { pageIndices: [2], paperSize: 'Letter' }),
    /page range/i,
  )
})

test('direct-print viewer waits for its explicit decoded-document signal, not shell load events', async () => {
  const webContents = new EventEmitter()
  const browserWindow = new EventEmitter()
  browserWindow.webContents = webContents
  let loadCalls = 0
  browserWindow.loadURL = async (url) => {
    loadCalls += 1
    assert.equal(url, 'file:///prepared.pdf')
  }

  let resolved = false
  const ready = loadPdfViewerForPrint(browserWindow, 'file:///prepared.pdf', { timeoutMs: 1_000 })
    .then(() => { resolved = true })
  await new Promise((resolve) => setImmediate(resolve))
  assert.equal(loadCalls, 1)

  webContents.emit('did-stop-loading')
  webContents.emit('page-title-updated', {}, 'prepared.pdf', false)
  await new Promise((resolve) => setImmediate(resolve))
  assert.equal(resolved, false)

  webContents.emit('page-title-updated', {}, 'prepared.pdf', true)
  await ready
  assert.equal(resolved, true)
  assert.equal(webContents.listenerCount('page-title-updated'), 0)
  assert.equal(webContents.listenerCount('did-fail-load'), 0)
  assert.equal(browserWindow.listenerCount('closed'), 0)
})

test('direct-print viewer readiness listener is installed before a fast load starts', async () => {
  const webContents = new EventEmitter()
  const browserWindow = new EventEmitter()
  browserWindow.webContents = webContents
  browserWindow.loadURL = async () => {
    webContents.emit('page-title-updated', {}, 'prepared.pdf', true)
  }
  await loadPdfViewerForPrint(browserWindow, 'file:///prepared.pdf', { timeoutMs: 1_000 })
})

test('direct-print viewer readiness rejects a main-frame load failure and cleans up', async () => {
  const webContents = new EventEmitter()
  const browserWindow = new EventEmitter()
  browserWindow.webContents = webContents
  browserWindow.loadURL = async () => {
    webContents.emit('did-fail-load', {}, -3, 'aborted', 'file:///prepared.pdf', true)
  }
  await assert.rejects(
    loadPdfViewerForPrint(browserWindow, 'file:///prepared.pdf', { timeoutMs: 1_000 }),
    /could not load \(-3\): aborted/i,
  )
  assert.equal(webContents.listenerCount('page-title-updated'), 0)
  assert.equal(browserWindow.listenerCount('closed'), 0)
})

test('print submission times out without retrying and ignores a late success', async () => {
  const renderer = new EventEmitter()
  let submissions = 0
  let callback
  renderer.print = (_options, cb) => { submissions += 1; callback = cb }
  const result = await printWebContentsSilently(renderer, { silent: true }, { timeoutMs: 5 })
  assert.equal(result.success, false)
  assert.match(result.failureReason, /Check the print queue.*duplicate/)
  callback(true)
  assert.equal(submissions, 1)
  assert.equal(renderer.listenerCount('destroyed'), 0)
  assert.equal(renderer.listenerCount('render-process-gone'), 0)
})

test('print submission observes a renderer crash and releases listeners', async () => {
  const renderer = new EventEmitter()
  renderer.print = () => queueMicrotask(() => renderer.emit('render-process-gone', {}, { reason: 'crashed' }))
  const result = await printWebContentsSilently(renderer, { silent: true })
  assert.equal(result.success, false)
  assert.match(result.failureReason, /renderer stopped/)
  assert.equal(renderer.listenerCount('destroyed'), 0)
})
