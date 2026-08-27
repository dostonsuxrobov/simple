import fs from 'node:fs/promises'
import path from 'node:path'
import { createRequire } from 'node:module'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'

const require = createRequire(import.meta.url)
const { replacePdfOutlines } = require('../electron/pdf-outlines.cjs')
const pdfLib = require('pdf-lib')
const outputDirectory = path.resolve('tmp/pdfs')
const outputPath = path.join(outputDirectory, 'simple-interaction-fixture.pdf')
const dropPath = path.join(outputDirectory, 'simple-drop-page.pdf')

await fs.mkdir(outputDirectory, { recursive: true })
const pdf = await PDFDocument.create()
const regular = await pdf.embedFont(StandardFonts.Helvetica)
const bold = await pdf.embedFont(StandardFonts.HelveticaBold)
const italic = await pdf.embedFont(StandardFonts.TimesRomanItalic)
const image = await pdf.embedPng(Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
)))

for (let index = 0; index < 8; index += 1) {
  const page = pdf.addPage(index === 3 ? [720, 500] : [500, 700])
  const { width, height } = page.getSize()
  page.drawRectangle({ x: 0, y: 0, width, height, color: rgb(1, 1, 1) })
  page.drawText(`Page ${index + 1}`, { x: 44, y: height - 54, size: 22, font: bold, color: rgb(0.04, 0.04, 0.05) })
  page.drawRectangle({ x: 42, y: height - 128, width: Math.min(410, width - 84), height: 42, color: rgb(0.88, 0.94, 1) })
  page.drawText('Select only these four words, not the complete line.', {
    x: 52,
    y: height - 113,
    size: 16,
    font: index % 2 ? italic : regular,
    color: rgb(0.08, 0.17, 0.29),
  })
  page.drawText('Inline style-preserving edit sample', {
    x: 52,
    y: height - 174,
    size: 19,
    font: index % 2 ? bold : italic,
    color: rgb(0.32, 0.12, 0.46),
  })
  page.drawText('Continuous scrolling verification content', { x: 52, y: 80, size: 13, font: regular })
  if (index === 0) {
    page.drawImage(image, { x: 300, y: height - 286, width: 132, height: 78 })
    page.drawText('Text remains selectable above image', {
      x: 306,
      y: height - 246,
      size: 11,
      font: bold,
      color: rgb(0.1, 0.1, 0.12),
    })
  }
}

replacePdfOutlines(pdf, [
  { title: 'Opening section', pageIndex: 0, depth: 0, bold: true },
  { title: 'Nested detail', pageIndex: 2, depth: 1, italic: true, color: [40, 90, 170] },
  { title: 'Landscape page', pageIndex: 3, depth: 0 },
  { title: 'Final section', pageIndex: 7, depth: 0 },
], pdfLib)

pdf.setTitle('simple interaction fixture')
pdf.setProducer('simple QA')
await fs.writeFile(outputPath, await pdf.save({ useObjectStreams: true }))

const droppedPdf = await PDFDocument.create()
const droppedPage = droppedPdf.addPage([420, 420])
const droppedFont = await droppedPdf.embedFont(StandardFonts.HelveticaBold)
droppedPage.drawText('Dropped page', { x: 105, y: 205, size: 26, font: droppedFont })
await fs.writeFile(dropPath, await droppedPdf.save({ useObjectStreams: true }))
console.log(JSON.stringify({ outputPath, dropPath }))
