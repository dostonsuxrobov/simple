import assert from 'node:assert/strict'
import { createRequire } from 'node:module'
import fs from 'node:fs/promises'
import path from 'node:path'
import { PDFDocument } from 'pdf-lib'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const require = createRequire(import.meta.url)
const { listPageImageDraws } = require('../electron/pdf-content-edits.cjs')
const inputPath = path.resolve(process.argv[2] || 'tmp/pdfs/simple-editor-saved.pdf')
const bytes = new Uint8Array(await fs.readFile(inputPath))

const structural = await PDFDocument.load(bytes)
assert.equal(listPageImageDraws(structural.getPage(0)).length, 0, 'Deleted native image invocation remained in the saved PDF')

const rendered = await pdfjs.getDocument({ data: bytes.slice(), disableWorker: true, isEvalSupported: false }).promise
const page = await rendered.getPage(1)
const content = (await page.getTextContent()).items
const text = content
  .map((item) => 'str' in item ? item.str : '')
  .join(' ')
  .replace(/\s+/g, ' ')
  .trim()
const finalReplacement = 'Select precisely four words, not the complete line.'
assert.match(text, /Select only these four words, not the complete line\./, 'Original source run was unexpectedly corrupted')
assert.ok(text.includes(finalReplacement), 'Final replacement was not written')
assert.ok(!text.includes('Select exactly four words, not the complete line.'), 'An intermediate replacement layer survived the second save')
assert.equal(text.split(finalReplacement).length - 1, 1, 'Final replacement was written more than once')
const originalRun = content.find((item) => 'str' in item && item.str.startsWith('Select only these'))
const replacementRun = content.find((item) => 'str' in item && item.str.includes('precisely'))
assert.ok(originalRun && replacementRun, 'Source and replacement baselines could not be inspected')
assert.ok(Math.abs(originalRun.transform[5] - replacementRun.transform[5]) < 0.1, 'Editing moved the native text baseline')
assert.ok(Math.abs(replacementRun.transform[1]) < 0.01 && Math.abs(replacementRun.transform[2]) < 0.01, 'Page viewing rotation was incorrectly baked into the replacement text')
await rendered.destroy()

console.log(JSON.stringify({ inputPath, pages: structural.getPageCount(), imageDraws: 0, finalReplacementCount: 1, intermediateReplacementCount: 0 }))
