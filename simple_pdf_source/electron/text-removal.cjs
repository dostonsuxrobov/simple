'use strict'

const { loadMupdf } = require('./mupdf-loader.cjs')

const ALIGNED_GLYPH_TOLERANCE = 15 * Math.PI / 180
// Where to probe a glyph for a redaction point, as fractions of its quad
// (along the baseline, up from the bottom). The centre comes first.
const PROBES = [[0.5, 0.5], [0.3, 0.5], [0.7, 0.5], [0.5, 0.3], [0.5, 0.7], [0.3, 0.3], [0.7, 0.7], [0.3, 0.7], [0.7, 0.3]]

function angleDifference(left, right) {
  const difference = Math.abs(left - right) % (2 * Math.PI)
  return Math.min(difference, 2 * Math.PI - difference)
}

// mupdf quads are [ul, ur, ll, lr]; interpolate inside that box.
function quadPoint(quad, along, up) {
  const lowX = quad[4] + (quad[6] - quad[4]) * along
  const lowY = quad[5] + (quad[7] - quad[5]) * along
  const highX = quad[0] + (quad[2] - quad[0]) * along
  const highY = quad[1] + (quad[3] - quad[1]) * along
  return [lowX + (highX - lowX) * up, lowY + (highY - lowY) * up]
}

function insideQuad(quad, x, y) {
  const corners = [[quad[0], quad[1]], [quad[2], quad[3]], [quad[6], quad[7]], [quad[4], quad[5]]]
  let sign = 0
  for (let index = 0; index < 4; index += 1) {
    const [ax, ay] = corners[index]
    const [bx, by] = corners[(index + 1) % 4]
    const cross = (bx - ax) * (y - ay) - (by - ay) * (x - ax)
    if (Math.abs(cross) < 1e-9) continue
    if (!sign) sign = Math.sign(cross)
    else if (Math.sign(cross) !== sign) return false
  }
  return true
}

function quadBounds(quad) {
  const xs = [quad[0], quad[2], quad[4], quad[6]]
  const ys = [quad[1], quad[3], quad[5], quad[7]]
  return { x0: Math.min(...xs), y0: Math.min(...ys), x1: Math.max(...xs), y1: Math.max(...ys) }
}

/**
 * The redaction removes every glyph whose box contains the point, so use a
 * point of the target glyph that no other glyph covers (overlapping
 * watermarks, stamps, touching superscripts). The centre is kept whenever
 * it is free, which is the common case.
 */
function redactionPoint(glyph, others) {
  let fallback = null
  for (const [along, up] of PROBES) {
    const [x, y] = quadPoint(glyph.quad, along, up)
    fallback ||= [x, y]
    const covered = others.some((other) => x >= other.bounds.x0 && x <= other.bounds.x1
      && y >= other.bounds.y0 && y <= other.bounds.y1 && insideQuad(other.quad, x, y))
    if (!covered) return [x, y]
  }
  return fallback
}

/** A tiny quad around a point: it selects the glyph(s) whose box holds the point. */
function pointQuad([x, y]) {
  return [x - .01, y - .01, x + .01, y - .01, x - .01, y + .01, x + .01, y + .01]
}

/**
 * Remove the text under `quads` (mupdf page space) from one page, keeping
 * images, vector art and every existing annotation object.
 */
function redactTextAt(mupdf, doc, page, quads) {
  if (!quads.length) return
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

// Keep the PDF content behind edited glyphs: no filled rectangles, raster
// patches, or image/vector redactions. Loaded only when native text is edited.
//
// `options.onUnlocated(edit)` reports an edit whose original glyphs cannot be
// found and skips it; without it such an edit aborts with an error, as before.
async function removeNativeText(data, edits, options = {}) {
  const targets = edits.filter(edit => edit.type === 'text' && edit.cover && edit.originalRect)
  if (!targets.length) return data
  const mupdf = await loadMupdf()
  const doc = mupdf.Document.openDocument(data, 'application/pdf')
  try {
    for (const pageIndex of new Set(targets.map(edit => edit.pageIndex))) {
      const page = doc.loadPage(pageIndex)
      try {
        const inverse = mupdf.Matrix.invert(page.getTransform())
        const pageTargets = targets.filter(edit => edit.pageIndex === pageIndex)
        // Rotated runs (vertical labels, diagonal stamps) have an axis-aligned
        // box that also covers neighbouring horizontal text. Prefer glyphs that
        // run in the edit's direction; fall back to every glyph in the box.
        const matches = pageTargets.map(() => ({ all: [], aligned: [] }))
        const glyphs = []
        const text = page.toStructuredText('preserve-whitespace')
        try {
          text.walk({ onChar(character, origin, font, size, quad) {
            const x = (quad[0] + quad[2] + quad[4] + quad[6]) / 4
            const y = (quad[1] + quad[3] + quad[5] + quad[7]) / 4
            const px = x * inverse[0] + y * inverse[2] + inverse[4]
            const py = x * inverse[1] + y * inverse[3] + inverse[5]
            const dx = quad[2] - quad[0]
            const dy = quad[3] - quad[1]
            const glyphAngle = Math.atan2(dx * inverse[1] + dy * inverse[3], dx * inverse[0] + dy * inverse[2])
            const index = glyphs.length
            glyphs.push({ quad: Array.from(quad), bounds: quadBounds(quad) })
            pageTargets.forEach((edit, editIndex) => {
              const r = edit.originalRect
              if (px >= r.x - 0.1 && px <= r.x + r.width + 0.1
                && py >= r.y - 0.1 && py <= r.y + r.height + 0.1) {
                matches[editIndex].all.push(index)
                if (angleDifference(glyphAngle, Number(edit.angle) || 0) <= ALIGNED_GLYPH_TOLERANCE) matches[editIndex].aligned.push(index)
              }
            })
          } })
        } finally { text.destroy() }
        const selected = new Set()
        pageTargets.forEach((edit, editIndex) => {
          const { all, aligned } = matches[editIndex]
          if (String(edit.originalText || '').trim() && !all.length) {
            if (typeof options.onUnlocated === 'function') {
              options.onUnlocated(edit)
              return
            }
            throw new Error(`The original text on page ${pageIndex + 1} could not be located. Reopen the text selection before saving.`)
          }
          const rotated = Math.abs(Number(edit.angle) || 0) >= 0.01
          for (const index of rotated && aligned.length ? aligned : all) selected.add(index)
        })
        const others = glyphs.filter((_, index) => !selected.has(index))
        // A point inside the glyph selects that glyph without sweeping up
        // adjacent lines through font ascender/descender bounding boxes.
        const quads = [...selected].map((index) => pointQuad(redactionPoint(glyphs[index], others)))
        redactTextAt(mupdf, doc, page, quads)
      } finally { page.destroy() }
    }
    const buffer = doc.saveToBuffer('compress')
    try { return Uint8Array.from(buffer.asUint8Array()) } finally { buffer.destroy() }
  } finally { doc.destroy() }
}

module.exports = { insideQuad, pointQuad, quadBounds, quadPoint, redactTextAt, redactionPoint, removeNativeText }
