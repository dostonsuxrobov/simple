'use strict'

const { PDFDocument, PDFRawStream, decodePDFRawStream } = require('pdf-lib')

const ONE_PIXEL_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
)

function pngDataUrl(bytes = ONE_PIXEL_PNG) {
  return `data:image/png;base64,${Buffer.from(bytes).toString('base64')}`
}

async function pdfjs() {
  return import('pdfjs-dist/legacy/build/pdf.mjs')
}

/** Text items of each page, as pdf.js reports them. */
async function textItems(bytes) {
  const { getDocument } = await pdfjs()
  const document = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise
  try {
    const pages = []
    for (let number = 1; number <= document.numPages; number += 1) {
      const content = await (await document.getPage(number)).getTextContent()
      pages.push(content.items.filter((item) => 'str' in item))
    }
    return pages
  } finally {
    await document.destroy()
  }
}

async function pageTexts(bytes) {
  return (await textItems(bytes)).map((items) => items.map((item) => item.str).join(''))
}

/** Every stream of a PDF, decoded where pdf-lib can decode it. */
async function decodedStreams(bytes) {
  const document = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false })
  const streams = []
  for (const [ref, object] of document.context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFRawStream)) continue
    let data
    try { data = decodePDFRawStream(object).decode() } catch { data = object.getContents() }
    streams.push({ ref, dict: object.dict, data: Buffer.from(data) })
  }
  return streams
}

/** True when the text appears in any (decoded) stream or in the raw file. */
async function fileContainsText(bytes, text) {
  const raw = Buffer.from(bytes)
  const needles = [Buffer.from(text, 'latin1'), Buffer.from(Buffer.from(text, 'latin1').toString('hex').toUpperCase()), Buffer.from(Buffer.from(text, 'latin1').toString('hex'))]
  if (needles.some((needle) => raw.includes(needle))) return true
  for (const stream of await decodedStreams(bytes)) {
    if (needles.some((needle) => stream.data.includes(needle))) return true
  }
  return false
}

/**
 * True when `text` survives anywhere in the file: raw bytes, decoded streams,
 * or any string value of any object (object streams included).
 */
async function documentContainsText(bytes, text) {
  if (await fileContainsText(bytes, text)) return true
  const { PDFArray, PDFDict, PDFHexString, PDFString } = require('pdf-lib')
  const document = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false })
  const seen = new Set()
  const visit = (value) => {
    if (!value || seen.has(value)) return false
    if (value instanceof PDFString || value instanceof PDFHexString) {
      try { return value.decodeText().includes(text) } catch { return false }
    }
    if (value instanceof PDFDict) {
      seen.add(value)
      return value.entries().some(([, child]) => visit(child))
    }
    if (value instanceof PDFArray) {
      seen.add(value)
      return value.asArray().some(visit)
    }
    return false
  }
  for (const [, object] of document.context.enumerateIndirectObjects()) {
    if (visit(object instanceof PDFRawStream ? object.dict : object)) return true
  }
  return false
}

/** Page dictionaries present in the file, whether or not they are in the page tree. */
async function pageObjectCount(bytes) {
  const { PDFDict, PDFName } = require('pdf-lib')
  const document = await PDFDocument.load(bytes, { updateMetadata: false, throwOnInvalidObject: false })
  return document.context.enumerateIndirectObjects()
    .filter(([, object]) => object instanceof PDFDict && object.get(PDFName.of('Type'))?.toString() === '/Page').length
}

module.exports = {
  ONE_PIXEL_PNG,
  decodedStreams,
  documentContainsText,
  fileContainsText,
  pageObjectCount,
  pageTexts,
  pdfjs,
  pngDataUrl,
  textItems,
}
