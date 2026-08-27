import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { createRequire } from 'node:module'
import { PDFDocument } from 'pdf-lib'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'
import { createServer } from 'vite'

const require = createRequire(import.meta.url)
const { replacePdfOutlines } = require('../electron/pdf-outlines.cjs')
const pdfLib = require('pdf-lib')
const outputDirectory = path.resolve('tmp/pdfs')
const outputPath = path.join(outputDirectory, 'native-outline-interoperability.pdf')

await fs.mkdir(outputDirectory, { recursive: true })
const source = await PDFDocument.create()
source.addPage([400, 600])
source.addPage([400, 600])
source.addPage([400, 600])

const outlineEntries = [
  { title: 'Chapter 1', pageIndex: 0, depth: 0, bold: true },
  { title: 'Résumé 日本語', pageIndex: 1, depth: 1, italic: true, color: [32, 96, 192] },
  { title: 'Chapter 2', pageIndex: 2, depth: 0 },
]
const result = replacePdfOutlines(source, outlineEntries, pdfLib)
assert.deepEqual(result, { written: 3, skipped: 0 })
const objectCount = source.context.enumerateIndirectObjects().length
assert.deepEqual(replacePdfOutlines(source, outlineEntries, pdfLib), { written: 3, skipped: 0 })
assert.equal(source.context.enumerateIndirectObjects().length, objectCount)
await fs.writeFile(outputPath, await source.save({ useObjectStreams: true }))

const bytes = new Uint8Array(await fs.readFile(outputPath))
const loaded = await pdfjs.getDocument({ data: bytes, disableWorker: true }).promise
const outline = await loaded.getOutline()
assert.equal(outline?.length, 2)
assert.equal(outline?.[0].title, 'Chapter 1')
assert.equal(outline?.[0].items.length, 1)
assert.equal(outline?.[0].items[0].title, 'Résumé 日本語')
assert.equal(outline?.[1].title, 'Chapter 2')

const destinations = [outline[0].dest, outline[0].items[0].dest, outline[1].dest]
const pages = await Promise.all(destinations.map(async (destination) => {
  assert.ok(Array.isArray(destination))
  return loaded.getPageIndex(destination[0])
}))
assert.deepEqual(pages, [0, 1, 2])
await loaded.destroy()

const vite = await createServer({ server: { middlewareMode: true }, appType: 'custom', logLevel: 'silent' })
try {
  const { getPdfOutlineBookmarks } = await vite.ssrLoadModule('/src/lib/pdf.ts')
  const pageReference = { num: 12, gen: 0 }
  const mockPdf = {
    numPages: 3,
    getOutline: async () => [
      {
        title: 'Named destination', bold: false, italic: false, color: new Uint8ClampedArray([0, 0, 0]),
        dest: 'chapter-two', url: null, count: 1,
        items: [{ title: 'Direct child', bold: true, italic: false, color: new Uint8ClampedArray([1, 2, 3]), dest: [2, { name: 'Fit' }], url: null, items: [] }],
      },
      { title: 'Heading only', bold: false, italic: true, color: new Uint8ClampedArray([0, 0, 0]), dest: null, url: null, items: [] },
    ],
    getDestination: async (name) => name === 'chapter-two' ? [pageReference, { name: 'Fit' }] : null,
    cachedPageNumber: () => null,
    getPageIndex: async (reference) => reference === pageReference ? 1 : Promise.reject(new Error('bad ref')),
  }
  const flattened = await getPdfOutlineBookmarks(mockPdf)
  assert.deepEqual(flattened.map(({ title, pageIndex, pageNumber, depth }) => ({ title, pageIndex, pageNumber, depth })), [
    { title: 'Named destination', pageIndex: 1, pageNumber: 2, depth: 0 },
    { title: 'Direct child', pageIndex: 2, pageNumber: 3, depth: 1 },
    { title: 'Heading only', pageIndex: null, pageNumber: null, depth: 0 },
  ])
} finally {
  await vite.close()
}

console.log(JSON.stringify({ outputPath, outlineTitles: [outline[0].title, outline[0].items[0].title, outline[1].title], pages }))
await fs.rm(outputPath, { force: true })
