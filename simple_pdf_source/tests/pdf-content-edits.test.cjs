'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { PDFDocument, PDFName, decodePDFRawStream } = require('pdf-lib')
const { listPageImageDraws, readPageContentBytes, removePageImageDraws } = require('../electron/pdf-content-edits.cjs')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { ONE_PIXEL_PNG, decodedStreams, pageTexts, pngDataUrl } = require('./helpers/pdf-test-utils.cjs')

// WinAnsi bytes 0x80-0x9F (’ “ ” • – €) in a literal string, an inline image
// whose samples are in the same range, and one image XObject to remove.
const TEXT = Buffer.from([...Buffer.from('BT /F1 12 Tf 40 700 Td (Don'), 0x92, ...Buffer.from('t '), 0x93,
  ...Buffer.from('Q'), 0x94, 0x20, 0x95, 0x20, 0x96, 0x20, 0x80, ...Buffer.from('5) Tj ET\n')])
const INLINE_SAMPLES = Buffer.from([0x81, 0x8a, 0x92, 0x9e])
const INLINE = Buffer.concat([
  Buffer.from('q 20 0 0 20 300 600 cm BI /W 2 /H 2 /BPC 8 /CS /G ID '),
  INLINE_SAMPLES,
  Buffer.from(' EI Q\n'),
])
const IMAGE_DRAW = Buffer.from('q 100 0 0 50 40 500 cm /Im1 Do Q\n')
const CONTENT = Buffer.concat([TEXT, INLINE, IMAGE_DRAW])
const IMAGE_RECT = { x: 40, y: 500, width: 100, height: 50 }

async function fixture() {
  const doc = await PDFDocument.create()
  const page = doc.addPage([400, 800])
  const image = await doc.embedPng(ONE_PIXEL_PNG)
  const font = doc.context.register(doc.context.obj({
    Type: 'Font', Subtype: 'Type1', BaseFont: 'Helvetica', Encoding: 'WinAnsiEncoding',
  }))
  page.node.set(PDFName.of('Resources'), doc.context.obj({ Font: { F1: font }, XObject: { Im1: image.ref } }))
  page.node.set(PDFName.of('Contents'), doc.context.register(doc.context.flateStream(CONTENT)))
  return doc.save()
}

test('removing an image keeps every other content byte, including 0x80-0x9F text and inline image data', async () => {
  const doc = await PDFDocument.load(await fixture())
  const page = doc.getPage(0)
  assert.deepEqual(Buffer.from(readPageContentBytes(page)), CONTENT)
  assert.equal(listPageImageDraws(page).length, 1)

  assert.equal(removePageImageDraws(page, [IMAGE_RECT]), 1)

  const contents = doc.context.lookup(page.node.get(PDFName.of('Contents')))
  const rewritten = Buffer.from(decodePDFRawStream(contents).decode())
  const start = CONTENT.indexOf('/Im1 Do')
  const expected = Buffer.from(CONTENT)
  expected.fill(0x20, start, start + '/Im1 Do'.length)
  assert.deepEqual(rewritten, expected, 'only the /Im1 Do operator may change')
  assert.ok(rewritten.includes(INLINE_SAMPLES), 'inline image samples survive')
  assert.equal(listPageImageDraws(page).length, 0)
  // The removed image is no longer a page resource, so it is not saved.
  const xobjects = page.node.Resources().lookup(PDFName.of('XObject'))
  assert.equal(xobjects.has(PDFName.of('Im1')), false)
})

test('an image kept by another draw on the page stays a resource', async () => {
  const doc = await PDFDocument.load(await fixture())
  const page = doc.getPage(0)
  const twice = Buffer.concat([CONTENT, Buffer.from('q 50 0 0 50 200 100 cm /Im1 Do Q\n')])
  page.node.set(PDFName.of('Contents'), doc.context.register(doc.context.flateStream(twice)))
  assert.equal(removePageImageDraws(page, [IMAGE_RECT]), 1)
  assert.equal(listPageImageDraws(page).length, 1)
  assert.equal(page.node.Resources().lookup(PDFName.of('XObject')).has(PDFName.of('Im1')), true)
})

test('moving a native image through Save keeps curly quotes, dashes, bullets, euro and inline images intact', async () => {
  const { invoke } = loadMain()
  const overlay = {
    id: 'image-move', type: 'object', kind: 'image', pageIndex: 0,
    rect: { x: 200, y: 300, width: 100, height: 50 },
    originalRect: IMAGE_RECT,
    dataUrl: pngDataUrl(),
    opacity: 1,
    cover: true,
  }
  const output = await invoke('pdf:flatten-overlays', await fixture(), [overlay], {}, {})
  const [text] = await pageTexts(output)
  assert.match(text, /Don’t “Q” • – €5/)
  const streams = await decodedStreams(output)
  assert.ok(streams.some((stream) => stream.data.includes(INLINE_SAMPLES)), 'inline image bytes are unchanged')
  assert.ok(streams.some((stream) => stream.data.includes(TEXT)), 'the text operator bytes are unchanged')
})
