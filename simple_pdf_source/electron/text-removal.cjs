'use strict'

// Keep the PDF content behind edited glyphs: no filled rectangles, raster
// patches, or image/vector redactions. Loaded only when native text is edited.
async function removeNativeText(data, edits) {
  const targets = edits.filter(edit => edit.type === 'text' && edit.cover && edit.originalRect)
  if (!targets.length) return data
  const path = require('node:path')
  const fs = require('node:fs')
  const { pathToFileURL } = require('node:url')
  const bundled = path.join(__dirname, '../vendor/mupdf/dist/mupdf.js')
  const mupdf = await import(pathToFileURL(fs.existsSync(bundled) ? bundled : require.resolve('mupdf')).href)
  const doc = mupdf.Document.openDocument(data, 'application/pdf')
  try {
    for (const pageIndex of new Set(targets.map(edit => edit.pageIndex))) {
      const page = doc.loadPage(pageIndex)
      try {
        const inverse = mupdf.Matrix.invert(page.getTransform())
        const pageTargets = targets.filter(edit => edit.pageIndex === pageIndex)
        const hits = new Set()
        const quads = []
        const text = page.toStructuredText('preserve-whitespace')
        try {
          text.walk({ onChar(character, origin, font, size, quad) {
            const x = (quad[0] + quad[2] + quad[4] + quad[6]) / 4
            const y = (quad[1] + quad[3] + quad[5] + quad[7]) / 4
            const px = x * inverse[0] + y * inverse[2] + inverse[4]
            const py = x * inverse[1] + y * inverse[3] + inverse[5]
            let selected = false
            pageTargets.forEach((edit, index) => {
              const r = edit.originalRect
              if (px >= r.x - 0.1 && px <= r.x + r.width + 0.1
                && py >= r.y - 0.1 && py <= r.y + r.height + 0.1) {
                hits.add(index)
                selected = true
              }
            })
            // A point inside the glyph selects that glyph without sweeping up
            // adjacent lines through font ascender/descender bounding boxes.
            if (selected) quads.push([x - .01, y - .01, x + .01, y - .01, x - .01, y + .01, x + .01, y + .01])
          } })
        } finally { text.destroy() }
        if (pageTargets.some((edit, index) => String(edit.originalText || '').trim() && !hits.has(index))) {
          throw new Error(`The original text on page ${pageIndex + 1} could not be located. Reopen the text selection before saving.`)
        }
        if (quads.length) {
          // The redaction engine also drops intersecting links. This is a
          // content edit, so preserve the page's existing annotation objects
          // (links, comments, widgets, and pending user redactions) verbatim.
          const pageObject = page.getObject()
          const originalAnnotations = pageObject.get('Annots')
          const savedAnnotations = doc.newArray()
          for (let i = 0; i < originalAnnotations.length; i++) {
            const value = originalAnnotations.get(i)
            savedAnnotations.push(value)
            value.destroy()
          }
          const annotation = page.createAnnotation('Redact')
          try {
            annotation.setQuadPoints(quads)
            // Apply just our annotation; existing user redaction annotations
            // elsewhere on the page must remain pending and untouched.
            annotation.applyRedaction(0, mupdf.PDFPage.REDACT_IMAGE_NONE,
              mupdf.PDFPage.REDACT_LINE_ART_NONE, mupdf.PDFPage.REDACT_TEXT_REMOVE)
          } finally {
            if (savedAnnotations.length) pageObject.put('Annots', savedAnnotations)
            else pageObject.delete('Annots')
            annotation.destroy()
            savedAnnotations.destroy()
            originalAnnotations.destroy()
            pageObject.destroy()
          }
        }
      } finally { page.destroy() }
    }
    const buffer = doc.saveToBuffer('compress')
    try { return Uint8Array.from(buffer.asUint8Array()) } finally { buffer.destroy() }
  } finally { doc.destroy() }
}

module.exports = { removeNativeText }
