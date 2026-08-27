import assert from 'node:assert/strict'
import { createRequire } from 'node:module'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const require = createRequire(import.meta.url)
const { listPageImageDraws, removePageImageDraws } = require('../electron/pdf-content-edits.cjs')

const ONE_PIXEL_PNG = Uint8Array.from(Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
))

const source = await PDFDocument.create()
const page = source.addPage([500, 700])
const font = await source.embedFont(StandardFonts.Helvetica)
const image = await source.embedPng(ONE_PIXEL_PNG)
page.drawRectangle({ x: 0, y: 0, width: 500, height: 700, color: rgb(0.8, 0.9, 1) })
page.drawText('independent text survives', { x: 42, y: 330, size: 18, font })
page.drawImage(image, { x: 40, y: 500, width: 120, height: 60 })
page.drawImage(image, { x: 300, y: 500, width: 120, height: 60 })

const reopened = await PDFDocument.load(await source.save())
const editedPage = reopened.getPage(0)
const before = listPageImageDraws(editedPage)
assert.equal(before.length, 2)
assert.equal(removePageImageDraws(editedPage, [{ x: 40, y: 500, width: 120, height: 60 }]), 1)
const after = listPageImageDraws(editedPage)
assert.equal(after.length, 1)
assert.ok(Math.abs(after[0].rect.x - 300) < 0.1)

const output = await reopened.save()
const verified = await PDFDocument.load(output)
const verifiedPage = verified.getPage(0)
assert.equal(listPageImageDraws(verifiedPage).length, 1)
const rendered = await pdfjs.getDocument({ data: output.slice(), disableWorker: true }).promise
const text = (await (await rendered.getPage(1)).getTextContent()).items
  .map((item) => 'str' in item ? item.str : '')
  .join(' ')
assert.match(text, /independent text survives/)
await rendered.destroy()

console.log('pdf content edit tests passed')
