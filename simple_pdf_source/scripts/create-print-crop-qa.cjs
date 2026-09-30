'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib')
const { preparePrintPdf } = require('../electron/pdf-print.cjs')

async function main() {
  const outputDirectory = path.resolve(__dirname, '..', 'tmp', 'pdfs')
  await fs.mkdir(outputDirectory, { recursive: true })

  const source = await PDFDocument.create()
  const page = source.addPage([600, 500])
  page.setCropBox(100, 80, 300, 240)
  page.drawRectangle({ x: 0, y: 0, width: 600, height: 500, color: rgb(0.9, 0.08, 0.08) })
  page.drawRectangle({ x: 100, y: 80, width: 300, height: 240, color: rgb(0.08, 0.68, 0.3) })
  page.drawRectangle({ x: 100, y: 80, width: 300, height: 240, borderColor: rgb(0.05, 0.15, 0.55), borderWidth: 8 })
  const font = await source.embedFont(StandardFonts.HelveticaBold)
  page.drawText('ONLY THIS GREEN CROP SHOULD PRINT', { x: 126, y: 194, size: 18, font, color: rgb(1, 1, 1) })
  page.drawText('HIDDEN RED CONTENT', { x: 405, y: 390, size: 20, font, color: rgb(1, 1, 1) })

  const sourceBytes = await source.save()
  const targetBytes = await preparePrintPdf(sourceBytes, {
    paperSize: 'Letter',
    landscape: false,
    marginMode: 'normal',
    scaleMode: 'fit',
    customScale: 1,
  })
  const sourcePath = path.join(outputDirectory, 'print-crop-source.pdf')
  const targetPath = path.join(outputDirectory, 'print-crop-prepared.pdf')
  await Promise.all([fs.writeFile(sourcePath, sourceBytes), fs.writeFile(targetPath, targetBytes)])
  process.stdout.write(`${targetPath}\n`)
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
