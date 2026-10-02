'use strict'

const { PDFDocument, StandardFonts } = require('pdf-lib')

async function mupdf() {
  return import('mupdf')
}

/** A one-page PDF showing `text`, encrypted by mupdf with `options`. */
async function encryptedPdf(text, options) {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 300])
  page.drawText(text, { x: 20, y: 250, size: 14, font: await doc.embedFont(StandardFonts.Helvetica) })
  const plain = await doc.save({ useObjectStreams: false })
  const { Document } = await mupdf()
  const source = Document.openDocument(plain, 'application/pdf')
  try {
    return Buffer.from(source.saveToBuffer(options).asUint8Array())
  } finally {
    source.destroy()
  }
}

/**
 * Rewrite a classic-xref PDF into the layout of a linearized ("Fast Web
 * View") file: the full trailer, with /Encrypt, sits in a first-page section
 * right after the header; the final trailer holds only /Size; the file is
 * padded so the last 256 KB contain no /Encrypt at all.
 */
function linearizedLayout(input, padding = 300 * 1024) {
  const text = Buffer.from(input).toString('latin1')
  const startxref = text.lastIndexOf('startxref')
  const xrefStart = Number(/startxref\s+(\d+)/.exec(text.slice(startxref))[1])
  if (!text.startsWith('xref', xrefStart)) throw new Error('The fixture needs a classic xref table.')
  const trailerStart = text.indexOf('trailer', xrefStart)
  const trailer = text.slice(trailerStart + 'trailer'.length, startxref).trim()
  const headerEnd = text.indexOf('\n', text.indexOf('\n') + 1) + 1
  const table = text.slice(xrefStart, trailerStart).split(/\r?\n/).filter(Boolean)
  const [first, count] = table[1].trim().split(/\s+/).map(Number)
  const size = /\/Size\s+(\d+)/.exec(trailer)[1]
  const objects = text.slice(headerEnd, xrefStart)
  const pad = `%${'p'.repeat(78)}\n`.repeat(Math.ceil(padding / 80))
  const firstSection = (previous) => `xref\n0 1\n0000000000 65535 f \ntrailer\n${trailer.replace(/^<</, `<< /Prev ${String(previous).padStart(10, '0')}`)}\n`
  const delta = firstSection(0).length
  const mainXref = headerEnd + delta + objects.length + pad.length
  const entries = table.slice(2, 2 + count).map((line) => {
    const [offset, generation, type] = line.trim().split(/\s+/)
    return `${String(type === 'n' ? Number(offset) + delta : Number(offset)).padStart(10, '0')} ${generation} ${type} `
  })
  return Buffer.from(text.slice(0, headerEnd) + firstSection(mainXref) + objects + pad
    + `xref\n${first} ${count}\n${entries.join('\n')}\ntrailer\n<< /Size ${size} >>\nstartxref\n${headerEnd}\n%%EOF\n`, 'latin1')
}

module.exports = { encryptedPdf, linearizedLayout }
