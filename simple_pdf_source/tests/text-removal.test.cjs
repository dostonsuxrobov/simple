const test = require('node:test')
const assert = require('node:assert/strict')
const { PDFDocument, StandardFonts, rgb, degrees } = require('pdf-lib')
const { removeNativeText } = require('../electron/text-removal.cjs')

async function fixture(withText, rotation = 0, nested = false) {
  const doc = await PDFDocument.create()
  const page = doc.addPage([400, 500])
  page.setRotation(degrees(rotation))
  page.setCropBox(10, 15, 380, 470)
  for (let x = 0; x < 400; x += 10) {
    page.drawRectangle({ x, y: 0, width: 10, height: 500, color: rgb(x / 400, .4, 1 - x / 400) })
  }
  const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')
  page.drawImage(await doc.embedPng(png), { x: 40, y: 300, width: 100, height: 140 })
  page.drawLine({ start: { x: 0, y: 405 }, end: { x: 400, y: 405 }, thickness: 2, color: rgb(1, 1, 0) })
  const font = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Keep this neighboring line', { x: 40, y: 418, size: 12, font })
  if (nested) {
    const inner = await PDFDocument.create()
    const p = inner.addPage([400, 500])
    if (withText) p.drawText('REMOVE ME', { x: 40, y: 400, size: 12 })
    // Keep the form nonempty even in the expected fixture.
    p.drawText('KEEP FORM', { x: 40, y: 250, size: 12 })
    const [embedded] = await doc.embedPdf(await inner.save())
    page.drawPage(embedded)
    const second = doc.addPage([400, 500])
    second.drawPage(embedded)
  } else if (withText) page.drawText('REMOVE ME', { x: 40, y: 400, size: 12, font })
  return doc.save()
}

const edit = { type: 'text', pageIndex: 0, cover: true, originalText: 'REMOVE ME',
  originalRect: { x: 40, y: 398, width: 75, height: 12 } }

async function rendered(bytes, pageIndex = 0) {
  const mupdf = await import('mupdf')
  const doc = mupdf.Document.openDocument(bytes, 'application/pdf')
  const page = doc.loadPage(pageIndex)
  const pix = page.toPixmap(mupdf.Matrix.scale(2, 2), mupdf.ColorSpace.DeviceRGB, false)
  const structured = page.toStructuredText()
  try { return { pixels: Buffer.from(pix.getPixels()), text: structured.asText() } }
  finally { structured.destroy(); pix.destroy(); page.destroy(); doc.destroy() }
}

for (const rotation of [0, 90, 180, 270]) {
  test(`text-only removal preserves every background pixel, crop, neighbors, and rotation ${rotation}`, async () => {
    const output = await removeNativeText(await fixture(true, rotation), [edit])
    const actual = await rendered(output)
    const expected = await rendered(await fixture(false, rotation))
    assert.equal(actual.text.includes('REMOVE'), false)
    assert.match(actual.text, /Keep this neighboring line/)
    assert.deepEqual(actual.pixels, expected.pixels)
  })
}

test('nested text is removed without changing a shared form on another page', async () => {
  const input = await fixture(true, 0, true)
  const output = await removeNativeText(input, [edit])
  assert.deepEqual((await rendered(output)).pixels, (await rendered(await fixture(false, 0, true))).pixels)
  assert.deepEqual((await rendered(output, 1)).pixels, (await rendered(input, 1)).pixels)
})

test('existing pending redactions are not applied by a text edit', async () => {
  const mupdf = await import('mupdf')
  const doc = mupdf.Document.openDocument(await fixture(true), 'application/pdf')
  const page = doc.loadPage(0)
  const annot = page.createAnnotation('Redact')
  annot.setRect([0, 0, 390, 90])
  const buffer = doc.saveToBuffer()
  const input = Uint8Array.from(buffer.asUint8Array())
  buffer.destroy(); annot.destroy(); page.destroy(); doc.destroy()
  const output = await removeNativeText(input, [edit])
  const reopened = mupdf.Document.openDocument(output, 'application/pdf')
  const resultPage = reopened.loadPage(0)
  const annotations = resultPage.getAnnotations()
  try {
    assert.equal(annotations.filter(a => a.getType() === 'Redact').length, 1)
    assert.match((await rendered(output)).text, /Keep this neighboring line/)
  } finally { annotations.forEach(a => a.destroy()); resultPage.destroy(); reopened.destroy() }
})

test('editing linked text preserves its hyperlink', async () => {
  const mupdf = await import('mupdf')
  const doc = mupdf.Document.openDocument(await fixture(true), 'application/pdf')
  const page = doc.loadPage(0)
  const rect = mupdf.Rect.transform([40, 398, 115, 410], page.getTransform())
  const link = page.createLink(rect, 'https://example.com/keep')
  const buffer = doc.saveToBuffer()
  const input = Uint8Array.from(buffer.asUint8Array())
  buffer.destroy(); link.destroy(); page.destroy(); doc.destroy()
  const output = await removeNativeText(input, [edit])
  const reopened = mupdf.Document.openDocument(output, 'application/pdf')
  const resultPage = reopened.loadPage(0), links = resultPage.getLinks()
  try { assert.deepEqual(links.map(link => link.getURI()), ['https://example.com/keep']) }
  finally { links.forEach(link => link.destroy()); resultPage.destroy(); reopened.destroy() }
})
