const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const path = require('node:path')
const { PDFDocument } = require('pdf-lib')
const { imageDimensions } = require('../electron/image-files.cjs')
const {
  MAX_PDF_PAGE_POINTS,
  ensurePdfExtension,
  imageToPdfBytes,
  pdfPageDimensions,
} = require('../electron/pdf-export.cjs')

test('PDF output names always end in .pdf', () => {
  assert.equal(ensurePdfExtension('photo'), 'photo.pdf')
  assert.equal(ensurePdfExtension('photo.PDF'), 'photo.PDF')
  assert.equal(ensurePdfExtension('photo.png'), 'photo.pdf')
  assert.throws(() => ensurePdfExtension(''), /invalid/)
})

test('PDF pages preserve aspect ratio and cap the longest edge', () => {
  assert.deepEqual(pdfPageDimensions(640, 480), { width: 640, height: 480, scale: 1 })
  const large = pdfPageDimensions(4000, 3000)
  assert.equal(large.width, MAX_PDF_PAGE_POINTS)
  assert.ok(Math.abs(large.height - MAX_PDF_PAGE_POINTS * 0.75) < 0.001)
  assert.ok(Math.abs(large.width / large.height - 4 / 3) < 0.000001)
  assert.throws(() => pdfPageDimensions(0, 100), /invalid/)
})

test('a PNG becomes a valid one-page PDF with image metadata', async () => {
  const png = await fs.readFile(path.join(__dirname, '..', 'public', 'brand-icon.png'))
  const source = imageDimensions(png, '.png')
  assert.ok(source)
  const bytes = await imageToPdfBytes(png, 'Brand icon')
  assert.equal(bytes.subarray(0, 5).toString('ascii'), '%PDF-')

  const pdf = await PDFDocument.load(bytes)
  assert.equal(pdf.getPageCount(), 1)
  assert.equal(pdf.getTitle(), 'Brand icon')
  assert.equal(pdf.getCreator(), 'simple')
  const page = pdf.getPage(0)
  const expected = pdfPageDimensions(source.width, source.height)
  assert.ok(Math.abs(page.getWidth() - expected.width) < 0.001)
  assert.ok(Math.abs(page.getHeight() - expected.height) < 0.001)
})
