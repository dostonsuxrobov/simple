const test = require('node:test')
const assert = require('node:assert/strict')

const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')
const JPEG = Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0x10, 0x4a, 0x46, 0x49, 0x46, 0, 1])
const GIF = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64')
const WEBP = Buffer.concat([Buffer.from('RIFF'), Buffer.from([30, 0, 0, 0]), Buffer.from('WEBPVP8L'), Buffer.alloc(8)])
const BMP = Buffer.concat([Buffer.from('BM'), Buffer.alloc(30)])
const SVG = Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>')
const TIFF = Buffer.from('II*\u0000abcdefgh')

test('picture formats are recognized from their bytes (DOC-SIE-17)', async () => {
  const { sniffImageType, isPortableImageType } = await import('../src/exporters/images.js')
  assert.equal(sniffImageType(PNG), 'image/png')
  assert.equal(sniffImageType(JPEG), 'image/jpeg')
  assert.equal(sniffImageType(GIF), 'image/gif')
  assert.equal(sniffImageType(WEBP), 'image/webp')
  assert.equal(sniffImageType(BMP), 'image/bmp')
  assert.equal(sniffImageType(SVG), 'image/svg+xml')
  assert.equal(sniffImageType(TIFF), 'image/tiff')
  assert.equal(sniffImageType(Buffer.from('nope')), null)
  assert.equal(isPortableImageType('image/png'), true)
  assert.equal(isPortableImageType('image/gif'), false)
})

test('PNG and JPEG pass through; other pictures become PNG/JPEG for PDF and print (DOC-SIE-17)', async () => {
  const { normalizeImageBytes } = await import('../src/exporters/images.js')
  const calls = []
  const codec = async (bytes, type) => { calls.push(type); return { bytes: type === 'image/svg+xml' ? PNG : JPEG, type: 'ignored' } }
  const png = await normalizeImageBytes(PNG, { codec })
  assert.equal(png.converted, false)
  assert.equal(png.bytes, PNG, 'PNG bytes are kept as they are')
  const gif = await normalizeImageBytes(GIF, { codec })
  assert.deepEqual([gif.converted, gif.type, gif.from], [true, 'image/jpeg', 'image/gif'])
  const svg = await normalizeImageBytes(SVG, { codec })
  assert.deepEqual([svg.type, svg.from], ['image/png', 'image/svg+xml'])
  assert.deepEqual(calls, ['image/gif', 'image/svg+xml'])
  // A codec that does not produce PNG/JPEG is never trusted.
  await assert.rejects(normalizeImageBytes(WEBP, { codec: async () => ({ bytes: GIF }) }), /could not be converted/)
  await assert.rejects(normalizeImageBytes(Buffer.from('nope'), { codec }), /not recognized/)
})

test('document pictures are normalized on a copy, once per source, with failures reported', async () => {
  const { normalizeDocumentImages } = await import('../src/exporters/images.js')
  const sources = { 'blob:png': PNG, 'blob:gif': GIF, 'blob:webp': WEBP, 'blob:tiff': TIFF }
  const image = (id, src) => ({ kind: 'image', id, revision: 1, src, widthPx: 10, heightPx: 10, align: 'left' })
  const document = {
    section: { pageWidthPx: 816, pageHeightPx: 1056, header: [image('header', 'blob:gif')] },
    blocks: [image('a', 'blob:png'), image('b', 'blob:gif'), { kind: 'table', id: 't', revision: 1, rows: [{ cells: [{ id: 'c', blocks: [image('c', 'blob:webp')] }] }] }, image('d', 'blob:tiff')],
  }
  const before = JSON.stringify(document)
  const resolved = []
  const registered = []
  const result = await normalizeDocumentImages(document, {
    resolve: async (source) => { resolved.push(source); return { type: '', arrayBuffer: async () => sources[source].buffer.slice(sources[source].byteOffset, sources[source].byteOffset + sources[source].byteLength) } },
    register: (bytes, type) => { registered.push(type); return `blob:normalized-${registered.length}` },
    codec: async (_bytes, type) => {
      if (type === 'image/tiff') throw new Error('TIFF cannot be decoded in the browser.')
      return { bytes: PNG }
    },
  })
  assert.equal(JSON.stringify(document), before, 'the live document is not changed')
  assert.deepEqual(resolved.sort(), ['blob:gif', 'blob:png', 'blob:tiff', 'blob:webp'])
  assert.equal(result.converted, 3, 'header and body copies of one GIF plus the WebP')
  assert.equal(registered.length, 2)
  assert.equal(result.document.blocks[0].src, 'blob:png')
  assert.equal(result.document.blocks[1].src, result.document.section.header[0].src)
  assert.match(result.document.blocks[2].rows[0].cells[0].blocks[0].src, /^blob:normalized-/)
  assert.equal(result.document.blocks[3].src, 'blob:tiff', 'an undecodable picture keeps its original bytes for DOCX')
  assert.equal(result.failed, 1)
  assert.match(result.warnings[0], /1 picture could not be converted/)
})

test('a failed Office page view never blocks printing or PDF export (doc-io-fidelity-6)', async () => {
  const { createLayoutPdfSource } = await import('../src/exporters/layout-pdf.js')
  const layoutPdf = new Uint8Array([1, 2, 3])
  const editorPdf = new Uint8Array([9])
  let renders = 0
  const render = async () => { renders += 1; return editorPdf }

  const working = createLayoutPdfSource(async () => layoutPdf)
  assert.equal(await working.pdfOr(render), layoutPdf)
  assert.equal(await working.pdfOr(render), layoutPdf)
  assert.equal(renders, 0)
  assert.equal(working.failed, false)

  let loads = 0
  const failing = createLayoutPdfSource(async () => { loads += 1; throw new Error('Document conversion took too long. The original file is unchanged.') })
  // The page view still sees the rejection and can explain it.
  await assert.rejects(failing.promise, /took too long/)
  assert.equal(failing.failed, true)
  assert.equal(await failing.pdfOr(render), editorPdf)
  assert.equal(await failing.pdfOr(render), editorPdf, 'every later print/export also falls back')
  assert.equal(renders, 2)
  assert.equal(loads, 1, 'the failed conversion is not retried on every print')
  assert.match(String(failing.error), /took too long/)

  // A synchronous throw from the loader is handled the same way.
  const throwing = createLayoutPdfSource(() => { throw new Error('IPC unavailable') })
  assert.equal(await throwing.pdfOr(render), editorPdf)
})
