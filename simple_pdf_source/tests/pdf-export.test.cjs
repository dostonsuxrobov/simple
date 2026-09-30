'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const test = require('node:test')
const JSZip = require('jszip')
const mammoth = require('mammoth')
const ts = require('typescript')
const {
  buildTextExport,
  imageExportFileName,
  normalizeTextPages,
  safeExportBaseName,
} = require('../electron/pdf-export.cjs')

function rendererExportHelpers() {
  const sourcePath = path.join(__dirname, '..', 'src', 'lib', 'pdfExport.ts')
  const output = ts.transpileModule(fs.readFileSync(sourcePath, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2022 },
  }).outputText
  const loaded = { exports: {} }
  Function('module', 'exports', 'require', output)(loaded, loaded.exports, require)
  return loaded.exports
}

const pages = [
  { pageNumber: 1, text: 'First page\nA < B & C' },
  { pageNumber: 3, text: 'Third page\nSecond line' },
]

test('export names stay local to the chosen destination', () => {
  assert.equal(safeExportBaseName('Quarterly: Report?.pdf'), 'Quarterly Report')
  assert.equal(safeExportBaseName('Quarterly.v2'), 'Quarterly.v2')
  assert.equal(safeExportBaseName('Quarterly.V2.PDF'), 'Quarterly.V2')
  assert.equal(imageExportFileName('Quarterly Report', 3, 12, 'png'), 'Quarterly Report - page 003.png')
  assert.equal(imageExportFileName('Quarterly Report', 3, 12, 'jpeg'), 'Quarterly Report - page 003.jpg')
  assert.equal(imageExportFileName('Quarterly Report', 3, 12, 'webp'), 'Quarterly Report - page 003.webp')
})

test('text page validation rejects empty and unreasonable inputs', () => {
  assert.throws(() => normalizeTextPages([]), /between 1 and 10,000/i)
  assert.throws(() => normalizeTextPages([{ pageNumber: 0, text: '' }]), /invalid page number/i)
  assert.deepEqual(normalizeTextPages(pages), pages)
})

test('plain-text and HTML exports retain page boundaries and escape active markup', async () => {
  const text = (await buildTextExport('txt', pages, 'Sample')).toString('utf8')
  assert.match(text, /Page 1/)
  assert.match(text, /\f/)
  assert.match(text, /Page 3/)

  const html = (await buildTextExport('html', pages, 'Sample')).toString('utf8')
  assert.match(html, /<section class="page">/)
  assert.match(html, /A &lt; B &amp; C/)
  assert.doesNotMatch(html, /A < B & C/)

  const markdown = (await buildTextExport('md', pages, 'Sample')).toString('utf8')
  assert.match(markdown, /^# Sample/m)
  assert.match(markdown, /^## Page 3/m)
})

test('Word export is a valid DOCX package with page breaks and escaped text', async () => {
  const bytes = await buildTextExport('docx', [...pages, { pageNumber: 4, text: 'Safe\u0001 text\u000b remains' }], 'Sample & Review')
  const zip = await JSZip.loadAsync(bytes)
  assert.ok(zip.file('[Content_Types].xml'))
  assert.ok(zip.file('word/document.xml'))
  const documentXml = await zip.file('word/document.xml').async('string')
  assert.match(documentXml, /w:type="page"/)
  assert.match(documentXml, /Sample &amp; Review/)
  assert.match(documentXml, /A &lt; B &amp; C/)
  assert.match(documentXml, /Safe text remains/)
  assert.doesNotMatch(documentXml, /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/)
  const extracted = await mammoth.extractRawText({ buffer: bytes })
  assert.match(extracted.value, /First page/)
  assert.match(extracted.value, /Safe text remains/)
})

test('renderer text extraction releases a page after a failed text read', async () => {
  const { extractPdfText } = rendererExportHelpers()
  let cleanups = 0
  const pdf = {
    getPage: async () => ({
      getTextContent: async () => { throw new Error('broken text stream') },
      cleanup: () => { cleanups += 1 },
    }),
  }
  await assert.rejects(extractPdfText(pdf, [0]), /broken text stream/)
  assert.equal(cleanups, 1)
})

test('renderer image conversion releases its page and canvas after an encoding setup failure', async () => {
  const { renderPdfPageImage } = rendererExportHelpers()
  let cleanups = 0
  const canvas = { width: 0, height: 0, getContext: () => null }
  const previousDocument = global.document
  global.document = { createElement: () => canvas }
  const pdf = {
    getPage: async () => ({
      rotate: 90,
      getViewport: ({ scale }) => ({ width: 100 * scale, height: 200 * scale }),
      cleanup: () => { cleanups += 1 },
    }),
  }
  try {
    await assert.rejects(renderPdfPageImage(pdf, 0, 'png', 2, 0.9), /canvas could not be created/i)
  } finally {
    global.document = previousDocument
  }
  assert.equal(cleanups, 1)
  assert.equal(canvas.width, 0)
  assert.equal(canvas.height, 0)
})

test('renderer image conversion applies intrinsic page rotation and requests the chosen encoding', async () => {
  const { renderPdfPageImage } = rendererExportHelpers()
  const viewportRequests = []
  const canvas = {
    width: 0,
    height: 0,
    getContext: () => ({}),
    toBlob: (callback, mimeType, quality) => {
      assert.equal(mimeType, 'image/webp')
      assert.equal(quality, 0.8)
      callback(new Blob([Uint8Array.of(1, 2, 3)], { type: mimeType }))
    },
  }
  let renderedViewport = null
  let cleanups = 0
  const previousDocument = global.document
  global.document = { createElement: () => canvas }
  const pdf = {
    getPage: async () => ({
      rotate: 270,
      getViewport: (options) => {
        viewportRequests.push(options)
        return { width: 120 * options.scale, height: 80 * options.scale }
      },
      render: ({ viewport }) => {
        renderedViewport = viewport
        return { promise: Promise.resolve() }
      },
      cleanup: () => { cleanups += 1 },
    }),
  }
  let result
  try {
    result = await renderPdfPageImage(pdf, 0, 'webp', 2, 0.8)
  } finally {
    global.document = previousDocument
  }
  assert.deepEqual([...result], [1, 2, 3])
  assert.deepEqual(viewportRequests.map((request) => request.rotation), [270, 270])
  assert.deepEqual(renderedViewport, { width: 240, height: 160 })
  assert.equal(cleanups, 1)
  assert.equal(canvas.width, 0)
  assert.equal(canvas.height, 0)
})

test('extremely large PDF pages respect image-export limits below 0.1x scale', async () => {
  const { renderPdfPageImage } = rendererExportHelpers()
  let rendered = false
  const canvas = {
    width: 0, height: 0, getContext: () => ({}),
    toBlob(callback) {
      assert.ok(this.width <= 8192 && this.height <= 8192)
      assert.ok(this.width * this.height <= 64_000_000)
      callback(new Blob([Uint8Array.of(1)], { type: 'image/png' }))
    },
  }
  const previousDocument = global.document
  global.document = { createElement: () => canvas }
  try {
    await renderPdfPageImage({ getPage: async () => ({
      rotate: 0,
      getViewport: ({ scale }) => ({ width: 200_000 * scale, height: 200_000 * scale }),
      render: () => { rendered = true; return { promise: Promise.resolve() } },
      cleanup() {},
    }) }, 0, 'png', 3, .9)
    assert.equal(rendered, true)
  } finally { global.document = previousDocument }
})
