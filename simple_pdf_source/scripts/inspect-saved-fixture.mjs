import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const inputPath = path.resolve(process.argv[2] || 'tmp/pdfs/simple-editor-saved.pdf')
const data = new Uint8Array(await fs.readFile(inputPath))
const pdf = await pdfjs.getDocument({ data, disableWorker: true, isEvalSupported: false }).promise
const rotations = []
for (let index = 0; index < pdf.numPages; index += 1) rotations.push((await pdf.getPage(index + 1)).rotate)
const editedPage = await pdf.getPage(4)
const editedText = (await editedPage.getTextContent()).items.map((item) => 'str' in item ? item.str : '').join(' ')
const outline = await pdf.getOutline()
const titles = []
const stack = [...(outline || [])].reverse()
while (stack.length) {
  const item = stack.pop()
  titles.push(item.title)
  if (item.items?.length) stack.push(...item.items.slice().reverse())
}

assert.equal(pdf.numPages, 8)
assert.equal(rotations[4], 90)
assert.ok(titles.includes('Page 5'))
assert.match(editedText, /Direct style-preserving edit\s+sample/)
console.log(JSON.stringify({ inputPath, pages: pdf.numPages, rotations, outlineTitles: titles, editedText }))
await pdf.destroy()
