const test = require('node:test')
const assert = require('node:assert/strict')
const { PDFDocument, StandardFonts } = require('pdf-lib')
const {
  MAX_PDF_BYTES,
  composePrintPdf,
  nativePrintOptions,
  normalizePrintSettings,
  parsePageRange,
  printPageGeometry,
  resolvePrinter,
  safeTextForFont,
  validatePdfBytes,
} = require('../electron/print-layout.cjs')

async function fixturePdf() {
  const document = await PDFDocument.create()
  document.addPage([400, 600]).drawText('First page', { x: 30, y: 550, size: 24 })
  document.addPage([500, 300]).drawText('Second page', { x: 30, y: 250, size: 24 })
  return new Uint8Array(await document.save())
}

test('page-range parser accepts normalized unions and rejects unsafe ranges', () => {
  assert.deepEqual(parsePageRange('3, 1-2, 2', 4), [0, 1, 2])
  assert.throws(() => parsePageRange('', 4), /Enter a page range/)
  assert.throws(() => parsePageRange('3-1', 4), /ascending/)
  assert.throws(() => parsePageRange('1, 5', 4), /outside this 4-page document/)
  assert.throws(() => parsePageRange('1;2', 4), /not a valid page or range/)
})

test('print settings normalize paper, margins, and scale to bounded values', () => {
  const settings = normalizePrintSettings({
    paper: 'A4',
    orientation: 'landscape',
    margins: { preset: 'custom', custom: { top: -4, right: 99, bottom: 0.5, left: '1.25' } },
    scaling: 'custom',
    scalePercent: 900,
    pages: 'custom',
    pageRange: '1-2',
  })
  assert.equal(settings.paper, 'A4')
  assert.equal(settings.orientation, 'landscape')
  assert.ok(settings.paperWidth > settings.paperHeight)
  assert.equal(settings.margins.top, 0)
  assert.ok(settings.margins.right < settings.paperWidth / 144)
  assert.equal(settings.margins.left, 1.25)
  assert.equal(settings.scalePercent, 200)
  assert.equal(settings.pages, 'custom')
})

test('document-size print preserves source paper and all pages by default', async () => {
  const composition = await composePrintPdf({
    data: await fixturePdf(),
    name: 'Quarterly plan.docx',
    settings: { paper: 'Document', orientation: 'portrait', margins: { preset: 'none' }, scaling: 'fit' },
  })
  assert.equal(composition.sourcePageCount, 2)
  assert.equal(composition.outputPageCount, 2)
  assert.deepEqual(composition.placements.map((placement) => placement.sourcePage), [1, 2])
  assert.equal(composition.placements[0].scale, 1)
  assert.equal(composition.placements[0].clipped, false)

  const result = await PDFDocument.load(composition.data)
  assert.equal(result.getPageCount(), 2)
  assert.deepEqual(result.getPage(0).getSize(), { width: 400, height: 600 })
  assert.equal(result.getTitle(), 'Quarterly plan.docx')
})

test('document-size print preserves every selected page size and orientation', async () => {
  const composition = await composePrintPdf({
    data: await fixturePdf(),
    name: 'Mixed paper.docx',
    settings: { paper: 'Document', orientation: 'portrait', margins: { preset: 'none' }, scaling: 'fit' },
  })
  const result = await PDFDocument.load(composition.data)
  assert.deepEqual(result.getPage(0).getSize(), { width: 400, height: 600 })
  assert.deepEqual(result.getPage(1).getSize(), { width: 500, height: 300 })
  assert.deepEqual(composition.placements.map(({ scale, clipped }) => ({ scale, clipped })), [
    { scale: 1, clipped: false },
    { scale: 1, clipped: false },
  ])
  assert.equal(composition.mixedPaperSizes, true)
})

test('composition uses one model for range, paper, scale, and page marks', async () => {
  const composition = await composePrintPdf({
    data: await fixturePdf(),
    name: 'Status report',
    settings: {
      paper: 'A4',
      orientation: 'landscape',
      margins: { preset: 'normal' },
      scaling: 'custom',
      scalePercent: 150,
      pages: 'custom',
      pageRange: '2',
      printTitle: true,
      printPageNumbers: true,
      centerContent: true,
    },
  })
  assert.equal(composition.sourcePageCount, 2)
  assert.equal(composition.outputPageCount, 1)
  assert.deepEqual(composition.placements, [{ sourcePage: 2, scale: 1.5, clipped: true }])
  const result = await PDFDocument.load(composition.data)
  const page = result.getPage(0)
  assert.ok(page.getWidth() > page.getHeight())
  assert.ok(Math.abs(page.getWidth() - 841.89) < 0.02)
  assert.ok(Math.abs(page.getHeight() - 595.28) < 0.02)
  assert.equal(composition.mixedPaperSizes, false)
})

// Text positions (x, y) of every `Tm` text matrix in a page's content streams.
function textPositions(document, pageIndex) {
  const { PDFArray, PDFRawStream, decodePDFRawStream } = require('pdf-lib')
  const contents = document.getPage(pageIndex).node.Contents()
  const streams = contents instanceof PDFArray ? contents.asArray().map((ref) => document.context.lookup(ref)) : [contents]
  const text = streams.map((stream) => Buffer.from(stream instanceof PDFRawStream ? decodePDFRawStream(stream).decode() : stream.getContents()).toString('latin1')).join('\n')
  return [...text.matchAll(/(-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) (-?[\d.]+) Tm/g)].map((match) => ({ x: Number(match[5]), y: Number(match[6]) }))
}

test('page marks stay inside a printable inset and their bands with the None margin preset', async () => {
  const composition = await composePrintPdf({
    data: await fixturePdf(),
    name: 'Margins none',
    settings: { paper: 'Letter', orientation: 'portrait', margins: { preset: 'none' }, scaling: 'fit', printTitle: true, printPageNumbers: true },
  })
  const result = await PDFDocument.load(composition.data)
  const height = result.getPage(0).getHeight()
  const positions = textPositions(result, 0)
  assert.equal(positions.length, 2, 'one title and one page number')
  const [titleMark, numberMark] = [...positions].sort((left, right) => right.y - left.y)
  const capHeight = 0.718 * 8
  const descent = 0.207 * 8
  // Typical printers cannot print within about 4–6 mm (12–17 pt) of the edge.
  assert.ok(height - (titleMark.y + capHeight) >= 17, `title top ${height - (titleMark.y + capHeight)} pt from the top edge`)
  assert.ok(titleMark.x >= 17, 'title starts inside the left printable edge')
  assert.ok(numberMark.y - descent >= 17, `page-number descenders ${numberMark.y - descent} pt from the bottom edge`)

  const geometry = printPageGeometry({ width: 612, height: 792 }, { top: 0, right: 0, bottom: 0, left: 0 }, { printTitle: true, printPageNumbers: true })
  // The content box starts above the footer band and ends below the header band.
  assert.ok(geometry.contentBottom >= geometry.pageNumberBaseline + 8)
  assert.ok(geometry.contentBottom + geometry.availableHeight <= geometry.titleBaseline - descent)
  assert.ok(geometry.titleBaseline + capHeight <= 792 - 17)
  // Without marks the None preset still uses the full sheet.
  const plain = printPageGeometry({ width: 612, height: 792 }, { top: 0, right: 0, bottom: 0, left: 0 }, {})
  assert.deepEqual([plain.contentBottom, plain.availableHeight, plain.availableWidth], [0, 792, 612])
  // Generous margins keep the marks inside the margin edge, as before.
  const normal = printPageGeometry({ width: 612, height: 792 }, { top: 54, right: 54, bottom: 54, left: 54 }, { printTitle: true, printPageNumbers: true })
  assert.ok(normal.titleBaseline + capHeight <= 792 - 54)
  assert.ok(normal.pageNumberBaseline - descent >= 54)
  assert.equal(normal.availableHeight, 792 - 108 - 36)
})

test('Unicode document titles remain metadata-exact and use a safe printable fallback', async () => {
  const title = 'Résumé — 東京 · Привет 🧾.docx'
  const composition = await composePrintPdf({
    data: await fixturePdf(),
    name: title,
    settings: {
      paper: 'Letter',
      orientation: 'portrait',
      margins: { preset: 'normal' },
      scaling: 'fit',
      printTitle: true,
    },
  })
  const result = await PDFDocument.load(composition.data)
  assert.equal(result.getTitle(), title)
  assert.equal(result.getPageCount(), 2)

  const fontDocument = await PDFDocument.create()
  const font = await fontDocument.embedFont(StandardFonts.Helvetica)
  const printable = safeTextForFont(title, font)
  assert.match(printable, /^Résumé/)
  assert.notEqual(printable, title)
  assert.doesNotThrow(() => font.encodeText(printable))
})

test('native print settings carry fixed and composed document paper without invalid default-size combinations', () => {
  const a4 = nativePrintOptions({
    paper: 'A4', orientation: 'landscape', paperWidth: 841.89, paperHeight: 595.28, mixedPaperSizes: false,
    deviceName: 'Office Printer', copies: 3, color: false, collate: true, duplexMode: 'longEdge',
  })
  assert.equal(a4.silent, true)
  assert.equal(a4.deviceName, 'Office Printer')
  assert.equal(a4.copies, 3)
  assert.equal(a4.color, false)
  assert.equal(a4.collate, true)
  assert.equal(a4.duplexMode, 'longEdge')
  assert.equal(a4.pageSize, 'A4')
  assert.equal(a4.landscape, true)
  assert.equal(Object.hasOwn(a4, 'usePrinterDefaultPageSize'), false)

  const document = nativePrintOptions({
    paper: 'Document', orientation: 'landscape', paperWidth: 500, paperHeight: 300, mixedPaperSizes: false, deviceName: 'Office Printer',
  })
  assert.deepEqual(document.pageSize, { width: 105833, height: 176389 })
  assert.equal(document.landscape, true)
  assert.equal(Object.hasOwn(document, 'usePrinterDefaultPageSize'), false)

  const mixed = nativePrintOptions({
    paper: 'Document', orientation: 'portrait', paperWidth: 400, paperHeight: 600, mixedPaperSizes: true, deviceName: 'Office Printer',
  })
  assert.equal(Object.hasOwn(mixed, 'pageSize'), false)
  assert.equal(mixed.usePrinterDefaultPageSize, true)
  assert.throws(() => nativePrintOptions({ paper: 'Document', orientation: 'portrait', deviceName: 'Office Printer' }), /Paper width/)
  const defaults = nativePrintOptions({ paper: 'A4', orientation: 'portrait' })
  assert.equal(defaults.silent, true)
  assert.equal(Object.hasOwn(defaults, 'deviceName'), false)
})

test('printer resolution honors an explicit choice, then the Windows default, and fails clearly', () => {
  const printers = [
    { name: 'Backup', displayName: 'Backup' },
    { name: 'Office', displayName: 'Office printer' },
  ]
  assert.equal(resolvePrinter(printers, 'Backup').name, 'Backup')
  assert.equal(resolvePrinter(printers).name, '')
  assert.equal(resolvePrinter(printers).displayName, 'Default Windows printer')
  assert.throws(() => resolvePrinter([], ''), /No printer is available/)
  assert.throws(() => resolvePrinter(printers, 'Missing'), /no longer available/)
})

test('print byte validation rejects spoofed, empty, and oversized payloads', async () => {
  assert.throws(() => validatePdfBytes(Buffer.from('hello')), /not a PDF/)
  assert.throws(() => validatePdfBytes(Buffer.alloc(0)), /invalid size/)
  assert.throws(() => validatePdfBytes({ byteLength: MAX_PDF_BYTES + 1 }), /byte array/)
  const bytes = await fixturePdf()
  assert.equal(validatePdfBytes(bytes).subarray(0, 5).toString('ascii'), '%PDF-')
})
