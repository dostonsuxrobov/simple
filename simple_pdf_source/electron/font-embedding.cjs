'use strict'

// fontkit 1.1.1 writes the source CFF byte length into the one-byte OffSize
// field. Values outside 1..4 cause FreeType/Acrobat-compatible readers to
// reject the saved face and substitute a visibly different font. INDEX
// offsets have their own sizes; the header uses the maximum four-byte size.
function normalizeCffSubsetHeader(data) {
  if (data.length < 4 || data[0] !== 1 || data[2] < 4) throw new Error('The embedded CFF font has an invalid header.')
  const result = Uint8Array.from(data)
  result[3] = 4
  return result
}

async function embedPdfFont(pdfDoc, bytes, requiredText) {
  if (requiredText !== undefined) {
    // Validate on a disposable document. A failed embedder registered on the
    // real document would be retried by save(), even after choosing a fallback.
    const { PDFDocument } = require('pdf-lib')
    const trial = await PDFDocument.create()
    trial.registerFontkit(require('@pdf-lib/fontkit'))
    const candidate = await embedPdfFont(trial, bytes)
    candidate.encodeText(String(requiredText))
    await candidate.embed()
  }
  const font = await pdfDoc.embedFont(bytes, { subset: true })
  const embedder = font.embedder
  if (embedder?.isCFF()) {
    const serialize = embedder.serializeFont.bind(embedder)
    embedder.serializeFont = async () => normalizeCffSubsetHeader(await serialize())
  }
  return font
}

module.exports = { embedPdfFont, normalizeCffSubsetHeader }
