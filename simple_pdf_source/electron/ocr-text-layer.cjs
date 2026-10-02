'use strict'

// Searchable text layer for recognised (OCR) pages, design section 4.7 (WP2).
//
// Each recognised word is written as invisible text (render mode 3) in
// Tesseract's GlyphLessFont, stretched to the word's ink width, with a
// separate stretched space between words. The page image is never touched, so
// the page looks exactly as before while search, selection, copy, mark-up and
// text export read the recognised words. The text sits inside
// `/SimpleOCR <<…>> BDC … EMC` in its own content stream, appended after the
// page's existing content (which pdf-lib wraps in q … Q), so it is drawn with
// the default coordinate system.
//
// Layers written earlier (by Simple or by another OCR tool) are removed with
// mupdf when text is recognised again: every glyph drawn in render mode 3 on
// the page is redacted at a point that no visible glyph covers.

const { codedError } = require('./pdf-problems.cjs')
const { loadMupdf } = require('./mupdf-loader.cjs')
const { insideQuad, pointQuad, quadBounds, quadPoint, redactTextAt } = require('./text-removal.cjs')

/**
 * Tesseract's GlyphLessFont (the 572-byte pdf.ttf distributed with
 * tesseract-ocr under the Apache-2.0 licence): one 0.5 em glyph that every
 * character code maps to. Embedded as base64 so the bundled backend never
 * reads it from disk. Its vertical metrics
 * are patched from 1 em / 0 to 0.8 / -0.2 em (1638 / -410 of 2048 units in
 * head.yMin/yMax, hhea ascender/descender and the OS/2 typo and win values;
 * table checksums and head.checkSumAdjustment recomputed). pdf.js takes ascent
 * and descent from these tables, so selection boxes cover the scanned ink.
 */
const GLYPHLESS_TTF_BASE64 = 'AAEAAAAKAIAAAwAgT1MvMl7czWAAAAEoAAAAYGNtYXAACgA0AAABkAAAAB5nbHlmFSJBJAAAAbgAAAAYaGVhZAt57jEAAACsAAAANmhoZWEKaAJpAAAA5AAAACRobXR4BAAAAAAAAYgAAAAIbG9jYQAMAAAAAAGwAAAABm1heHAABAAFAAABCAAAACBuYW1l8usW2gAAAdAAAABLcG9zdAABAAEAAAIcAAAAIAABAAAAAQAAo8pxEl8PPPUEBwgAAAAAAM+a/G4AAAAA1MOn8gAA/mYEAAZmAAAAEAACAAAAAAAAAAEAAAZm/mYAAAQAAAAAAAQAAAEAAAAAAAAAAAAAAAAAAAACAAEAAAACAAQAAQAAAAAAAQAAAAAAAAAAAAAAAAAAAAAAAwAAAZAABQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAUAAQABAAAAAAAAAAAAAAAAAAAAAAAAAAAAR09PRwBAAAAAAAZm/mYAAAZmAZqAAAAAAAAAAAAAAAAAAAABAAAAAAAABAAAAAAAAAIAAQAAAAAAFAADAAAAAAAUAAYACgAAAAAAAAAAAAAAAAAMAAAAAQAAAAAEAAgAAAMAADEhESEEAPwACAAAAAADACoAAAADAAAABQAWAAAAAQAAAAAABQALABYAAwABBAkABQAWAAAAVgBlAHIAcwBpAG8AbgAgADEALgAwVmVyc2lvbiAxLjAAAAEAAAAAAAAAAAAAAAAAAQAAAAAAAAAAAAAAAAAAAAA='

/** Font name used by Tesseract and OCRmyPDF; other tools recognise the layer by it. */
const GLYPHLESS_FONT_NAME = 'GlyphLessFont'
/** Marked-content tag (and resource key prefix) of Simple's layers. */
const OCR_MARK = 'SimpleOCR'
const OCR_LAYER_VERSION = 1
/** Glyph advance of the font in text space units (DW 500). */
const GLYPH_ADVANCE = 0.5

const LIMITS = Object.freeze({
  minFontSize: 0.5,
  maxFontSize: 500,
  /** Coordinates and lengths beyond this are not page geometry. */
  maxCoordinate: 1e6,
  maxLength: 1e5,
  maxWordsPerPage: 25_000,
  /** Recognised text per mutation, counted as UTF-16 (the layer's encoding). */
  maxTextBytes: 4 * 1024 * 1024,
  /** Horizontal scaling (Tz) is kept inside this range. */
  minScale: 1,
  maxScale: 5000,
  /** |dir| may differ from 1 by this much before it is rejected (it is renormalised). */
  directionTolerance: 0.05,
})

/** Invisible glyph origins in the device census and stext walk match within this (pt). */
const ORIGIN_TOLERANCE = 0.05
// Where to probe a glyph for a redaction point, as fractions of its quad
// (along the baseline, up from the bottom). The centre comes first.
const PROBES = [[0.5, 0.5], [0.3, 0.5], [0.7, 0.5], [0.5, 0.3], [0.5, 0.7], [0.3, 0.3], [0.7, 0.7], [0.3, 0.7], [0.7, 0.3]]

let pdfLibModule
function getPdfLib() {
  pdfLibModule ||= require('pdf-lib')
  return pdfLibModule
}

function invalidLayer(message) {
  return codedError('OCR_LAYER_INVALID', `The recognized text could not be added: ${message}`)
}

/** A fresh copy of the patched GlyphLessFont program. */
function glyphlessFontBytes() {
  return new Uint8Array(Buffer.from(GLYPHLESS_TTF_BASE64, 'base64'))
}

/** CIDToGIDMap stream data: all 65 536 CIDs (UTF-16 code units) draw glyph 1. */
function cidToGidMap() {
  const map = new Uint8Array(65536 * 2)
  for (let index = 1; index < map.length; index += 2) map[index] = 1
  return map
}

/**
 * ToUnicode CMap: code = UTF-16 code unit = Unicode value. One bfrange per
 * high byte (at most 100 per block), skipping the surrogate block D800-DFFF,
 * which cannot be a character on its own.
 */
function toUnicodeCMap() {
  const ranges = []
  for (let high = 0; high < 256; high += 1) {
    if (high >= 0xd8 && high <= 0xdf) continue
    const hex = high.toString(16).padStart(2, '0').toUpperCase()
    ranges.push(`<${hex}00> <${hex}FF> <${hex}00>`)
  }
  const blocks = []
  for (let index = 0; index < ranges.length; index += 100) {
    const block = ranges.slice(index, index + 100)
    blocks.push(`${block.length} beginbfrange\n${block.join('\n')}\nendbfrange`)
  }
  return [
    '/CIDInit /ProcSet findresource begin',
    '12 dict begin',
    'begincmap',
    '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def',
    '/CMapName /Adobe-Identity-UCS def',
    '/CMapType 2 def',
    '1 begincodespacerange',
    '<0000> <FFFF>',
    'endcodespacerange',
    ...blocks,
    'endcmap',
    'CMapName currentdict /CMap defineresource pop',
    'end',
    'end',
    '',
  ].join('\n')
}

/**
 * Add the GlyphLessFont (Type0 / Identity-H over a CIDFontType2 whose
 * CIDToGIDMap sends every code to glyph 1) to the document. Returns the
 * reference of the Type0 font dictionary; one font serves every page.
 */
function addGlyphlessFont(pdfDoc) {
  const { PDFString } = getPdfLib()
  const context = pdfDoc.context
  const program = glyphlessFontBytes()
  const fontFile = context.register(context.flateStream(program, { Length1: program.length }))
  const descriptor = context.register(context.obj({
    Type: 'FontDescriptor',
    FontName: GLYPHLESS_FONT_NAME,
    Flags: 5,
    FontBBox: [0, -200, 500, 800],
    ItalicAngle: 0,
    Ascent: 800,
    Descent: -200,
    CapHeight: 700,
    StemV: 80,
    FontFile2: fontFile,
  }))
  const cidFont = context.register(context.obj({
    Type: 'Font',
    Subtype: 'CIDFontType2',
    BaseFont: GLYPHLESS_FONT_NAME,
    CIDSystemInfo: { Registry: PDFString.of('Adobe'), Ordering: PDFString.of('Identity'), Supplement: 0 },
    FontDescriptor: descriptor,
    DW: 500,
    CIDToGIDMap: context.register(context.flateStream(cidToGidMap())),
  }))
  return context.register(context.obj({
    Type: 'Font',
    Subtype: 'Type0',
    BaseFont: GLYPHLESS_FONT_NAME,
    Encoding: 'Identity-H',
    DescendantFonts: [cidFont],
    ToUnicode: context.register(context.flateStream(toUnicodeCMap())),
  }))
}

/**
 * The page's /Font resource name for `fontRef`, added when missing. Pages that
 * share one resource dictionary reuse a single entry.
 */
function glyphlessFontKey(page, fontRef) {
  const { PDFDict, PDFName } = getPdfLib()
  page.node.normalize()
  const fonts = page.node.Resources()?.lookupMaybe(PDFName.of('Font'), PDFDict)
  for (const [key, value] of fonts ? fonts.entries() : []) {
    if (value === fontRef) return key
  }
  return page.node.newFontDictionary(OCR_MARK, fontRef)
}

/**
 * Text the 2-byte GlyphLessFont encoding can carry (the same rules as
 * ocr-geometry.mjs normalizeOcrText): ligatures decomposed, NFC, characters
 * outside the BMP and lone surrogates replaced with U+FFFD, controls removed.
 */
const LIGATURES = new Map([
  ['\uFB00', 'ff'], ['\uFB01', 'fi'], ['\uFB02', 'fl'], ['\uFB03', 'ffi'], ['\uFB04', 'ffl'], ['\uFB05', 'st'], ['\uFB06', 'st'],
])
function normalizeLayerText(text) {
  const value = String(text ?? '').replace(/[\uFB00-\uFB06]/g, (ligature) => LIGATURES.get(ligature) ?? ligature).normalize('NFC')
  let out = ''
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index)
    if (code >= 0xd800 && code <= 0xdbff) {
      const next = value.charCodeAt(index + 1)
      if (next >= 0xdc00 && next <= 0xdfff) index += 1
      out += '\uFFFD'
    } else if (code >= 0xdc00 && code <= 0xdfff) {
      out += '\uFFFD'
    } else if (code < 0x20 || (code >= 0x7f && code <= 0x9f)) {
      continue
    } else {
      out += value[index]
    }
  }
  return out
}

/** UTF-16BE hex of normalised text: one 2-byte code (= CID) per code unit. */
function utf16Hex(text) {
  let hex = ''
  for (let index = 0; index < text.length; index += 1) hex += text.charCodeAt(index).toString(16).padStart(4, '0')
  return hex.toUpperCase()
}

/** A PDF number: rounded, never "-0" and never in exponent notation (inputs are bounded). */
function formatNumber(value, digits = 3) {
  const factor = 10 ** digits
  const rounded = Math.round(value * factor) / factor
  return rounded === 0 ? '0' : String(rounded)
}

/** Printable ASCII metadata as a PDF literal string. */
function literalString(value) {
  return `(${String(value).replace(/[\\()]/g, '\\$&')})`
}

function sanitizeMeta(meta) {
  const clean = (value, pattern, fallback) => {
    const text = typeof value === 'string' ? value.replace(pattern, '').trim().slice(0, 64) : ''
    return text || fallback
  }
  return {
    engine: clean(meta?.engine, /[^\x20-\x7e]/g, 'tesseract.js'),
    language: clean(meta?.language, /[^A-Za-z0-9_+-]/g, 'eng'),
  }
}

const isNumber = (value) => typeof value === 'number' && Number.isFinite(value)

/**
 * Validate and normalise one page's lines (main never trusts the renderer):
 * finite numbers, |dir| = 1 (renormalised), font sizes in [0.5, 500],
 * non-negative widths and gaps, text normalised. Words that normalise to
 * nothing are dropped, and so are lines left without words. `budget` counts
 * the text of the whole mutation.
 */
function validateLines(lines, where, budget = { bytes: 0 }) {
  if (!Array.isArray(lines)) throw invalidLayer(`${where} has no list of lines.`)
  const result = []
  let words = 0
  lines.forEach((line, lineIndex) => {
    const at = `${where}, line ${lineIndex + 1}`
    if (!line || typeof line !== 'object' || !Array.isArray(line.words)) throw invalidLayer(`${at} has no list of words.`)
    const fontSize = line.fontSize
    if (!isNumber(fontSize) || fontSize < LIMITS.minFontSize || fontSize > LIMITS.maxFontSize) {
      throw invalidLayer(`${at} has an invalid font size.`)
    }
    const lineWords = []
    line.words.forEach((word, wordIndex) => {
      const label = `${at}, word ${wordIndex + 1}`
      if (!word || typeof word !== 'object' || typeof word.text !== 'string') throw invalidLayer(`${label} has no text.`)
      budget.bytes += word.text.length * 2
      if (budget.bytes > LIMITS.maxTextBytes) throw invalidLayer('there is too much text for one step.')
      for (const key of ['x', 'y', 'dx', 'dy', 'width', 'gap']) {
        if (!isNumber(word[key])) throw invalidLayer(`${label} has an invalid ${key}.`)
      }
      if (Math.abs(word.x) > LIMITS.maxCoordinate || Math.abs(word.y) > LIMITS.maxCoordinate) throw invalidLayer(`${label} lies far outside the page.`)
      if (word.width < 0 || word.gap < 0 || word.width > LIMITS.maxLength || word.gap > LIMITS.maxLength) throw invalidLayer(`${label} has an invalid width or gap.`)
      const length = Math.hypot(word.dx, word.dy)
      if (!(Math.abs(length - 1) <= LIMITS.directionTolerance)) throw invalidLayer(`${label} has an invalid direction.`)
      const text = normalizeLayerText(word.text).trim()
      if (!text) return
      lineWords.push({ text, x: word.x, y: word.y, dx: word.dx / length, dy: word.dy / length, width: word.width, gap: word.gap })
    })
    if (!lineWords.length) return
    words += lineWords.length
    if (words > LIMITS.maxWordsPerPage) throw invalidLayer(`${where} has more than ${LIMITS.maxWordsPerPage.toLocaleString('en-US')} words.`)
    result.push({ fontSize, words: lineWords })
  })
  return result
}

/**
 * Validate the 'ocr-text-layer' mutation against the document's page count.
 * Returns the normalised operation; throws OCR_LAYER_INVALID otherwise.
 */
function validateOcrLayerOperation(operation, pageCount) {
  if (!operation || typeof operation !== 'object' || operation.type !== 'ocr-text-layer') throw invalidLayer('the request is not a text layer.')
  if (!Array.isArray(operation.pages)) throw invalidLayer('the request lists no pages.')
  const budget = { bytes: 0 }
  const seen = new Set()
  const pages = operation.pages.map((page, index) => {
    const pageIndex = page?.pageIndex
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= pageCount) {
      throw invalidLayer(`entry ${index + 1} names a page that is not in the document.`)
    }
    if (seen.has(pageIndex)) throw invalidLayer(`page ${pageIndex + 1} is listed twice.`)
    seen.add(pageIndex)
    return { pageIndex, lines: validateLines(page.lines, `page ${pageIndex + 1}`, budget) }
  })
  return { type: 'ocr-text-layer', replaceExisting: operation.replaceExisting === true, meta: sanitizeMeta(operation.meta), pages }
}

const horizontalScale = (advance, glyphs, fontSize) => (
  Math.max(LIMITS.minScale, Math.min(LIMITS.maxScale, (100 * advance) / (GLYPH_ADVANCE * fontSize * glyphs)))
)

/**
 * Spacing limits from pdf.js text extraction, in em. A gap between words
 * below 0.102 em reads as letter tracking (below 0.03 em the words are joined
 * without a space); letters moved apart by more than 0.102 em read as words,
 * and moved back by more than 0.2 em as a new item.
 */
const MIN_WORD_GAP = 0.11
const MAX_LETTER_SPREAD = 0.09
const MAX_LETTER_SQUEEZE = 0.18

const clamp = (value, low, high) => Math.max(low, Math.min(high, value))
const median = (values) => {
  const sorted = [...values].sort((a, b) => a - b)
  const middle = sorted.length >> 1
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2
}

/**
 * One horizontal scale for every word of a slanted line. pdf.js measures
 * slanted text in units of each glyph's horizontal scale, so words stretched
 * differently would read as one line each. The median word fit is used,
 * lowered where needed so every word can still be narrowed (with character
 * spacing) to leave a space before the next one.
 */
function sharedLineScale(fontSize, words, advances) {
  let scale = median(words.map((word) => word.width / (word.text.length * GLYPH_ADVANCE * fontSize)))
  words.forEach((word, index) => {
    const advance = advances[index]
    if (advance === null || advance < word.width * 0.5) return
    const glyphs = word.text.length
    scale = Math.min(scale, (advance - MIN_WORD_GAP * fontSize + (glyphs - 1) * MAX_LETTER_SQUEEZE * fontSize) / (glyphs * GLYPH_ADVANCE * fontSize))
  })
  // Rounded as written, so every word carries exactly the same value.
  return Number(formatNumber(clamp(scale * 100, LIMITS.minScale, LIMITS.maxScale)))
}

/**
 * The text runs of one line: each word with its own text matrix along its
 * direction, fitted to its ink width, and after it a separate stretched space
 * up to the next word. Words on page-axis lines are stretched (Tz) to fit;
 * on slanted lines they share one Tz and are fitted with character spacing
 * (Tc) inside the range pdf.js still reads as one word. Gaps come from the
 * word positions; a word that touches or overlaps the next one is narrowed
 * (to at most half) so a space still separates them in text extraction.
 */
function lineRuns(line) {
  const { fontSize, words } = line
  const advances = words.map((word, index) => {
    const next = words[index + 1]
    return next ? (next.x - word.x) * word.dx + (next.y - word.y) * word.dy : null
  })
  const slanted = words.some((word) => word.dx !== 0 && word.dy !== 0)
  const shared = slanted ? sharedLineScale(fontSize, words, advances) : null
  const minimum = MIN_WORD_GAP * fontSize
  const runs = []
  words.forEach((word, index) => {
    const advance = advances[index]
    const glyphs = word.text.length
    let scale
    let spacing = 0
    // Where the pen stands after the word, measured from its origin.
    let end
    if (shared === null) {
      let width = word.width
      if (advance !== null && advance > 0 && advance - width < minimum) width = Math.max(word.width * 0.5, advance - minimum)
      scale = horizontalScale(width, glyphs, fontSize)
      end = (glyphs * GLYPH_ADVANCE * fontSize * scale) / 100
    } else {
      const h = shared / 100
      const natural = glyphs * GLYPH_ADVANCE * fontSize * h
      if (glyphs > 1) {
        spacing = Number(formatNumber(clamp((word.width - natural) / ((glyphs - 1) * h), (-MAX_LETTER_SQUEEZE * fontSize) / h, (MAX_LETTER_SPREAD * fontSize) / h)))
      }
      scale = shared
      end = natural + glyphs * spacing * h
    }
    runs.push({ matrix: [word.dx, word.dy, -word.dy, word.dx, word.x, word.y], scale, spacing, hex: utf16Hex(word.text) })
    if (advance === null) return
    // A next word behind this one (right-to-left order) keeps the requested gap.
    const space = advance > 0 ? Math.max(0, advance - end) : word.gap
    runs.push({ scale: horizontalScale(space, 1, fontSize), spacing: 0, hex: '0020' })
  })
  return runs
}

const markProperties = ({ engine, language }) => (
  `<</Engine ${literalString(engine)} /Lang ${literalString(language)} /Version ${OCR_LAYER_VERSION}>>`
)

/**
 * Content-stream source of one page's layer, inside `q … Q` so no text state
 * (render mode 3 in particular) reaches content appended after it. Highlight
 * boxes follow the scanned words while pdf.js and mupdf read one item per line.
 */
function ocrLayerContent(fontName, lines, meta = {}) {
  const out = ['q', `/${OCR_MARK} ${markProperties(sanitizeMeta(meta))} BDC`, 'BT', '3 Tr 0 Tc 0 Tw 0 Ts']
  let scale = null
  let spacing = '0'
  for (const line of lines) {
    out.push(`${fontName} ${formatNumber(line.fontSize)} Tf`)
    for (const run of lineRuns(line)) {
      if (run.matrix) {
        const [a, b, c, d, e, f] = run.matrix
        out.push(`${formatNumber(a, 6)} ${formatNumber(b, 6)} ${formatNumber(c, 6)} ${formatNumber(d, 6)} ${formatNumber(e)} ${formatNumber(f)} Tm`)
      }
      const operators = []
      if (formatNumber(run.spacing) !== spacing) operators.push(`${spacing = formatNumber(run.spacing)} Tc`)
      if (formatNumber(run.scale) !== scale) operators.push(`${scale = formatNumber(run.scale)} Tz`)
      out.push([...operators, `<${run.hex}> Tj`].join(' '))
    }
  }
  out.push('ET', 'EMC', 'Q', '')
  return out.join('\n')
}

/**
 * Write one page's recognised lines as an invisible text layer, in a content
 * stream of its own after the page's content. `lines` are validated here too.
 * Returns the number of words written.
 */
function addOcrTextLayer(pdfDoc, pageIndex, lines, { fontRef, meta } = {}) {
  if (!fontRef) throw new Error('addOcrTextLayer needs the GlyphLessFont reference (addGlyphlessFont).')
  const page = pdfDoc.getPage(pageIndex)
  const valid = validateLines(lines, `page ${pageIndex + 1}`)
  if (!valid.length) return 0
  const fontName = glyphlessFontKey(page, fontRef)
  const context = pdfDoc.context
  page.node.addContentStream(context.register(context.flateStream(ocrLayerContent(fontName.toString(), valid, meta))))
  return valid.reduce((sum, line) => sum + line.words.length, 0)
}

/**
 * Words of a run of text spread over `width` from `origin` along `dir`: an
 * even pitch for the letters and spaces of 0.2-0.5 em, so text extraction
 * reads the run as one line with single spaces.
 */
function runWords(text, origin, dir, width, fontSize) {
  const tokens = text.match(/\S+|\s+/g) || []
  const letters = tokens.reduce((sum, token) => sum + (/\S/.test(token) ? token.length : 0), 0)
  const spaces = tokens.reduce((sum, token) => sum + (/\S/.test(token) ? 0 : token.length), 0)
  const even = width / Math.max(1, letters + spaces)
  let space = Math.min(0.5 * fontSize, Math.max(0.2 * fontSize, even))
  let pitch = (width - spaces * space) / Math.max(1, letters)
  if (!(pitch >= 0.05 * fontSize)) space = pitch = even
  const words = []
  let offset = 0
  for (const token of tokens) {
    const advance = token.length * (/\S/.test(token) ? pitch : space)
    if (/\S/.test(token)) {
      words.push({ text: token, x: origin.x + offset * dir.x, y: origin.y + offset * dir.y, dx: dir.x, dy: dir.y, width: advance, gap: 0 })
    } else if (words.length) {
      words[words.length - 1].gap = advance
    }
    offset += advance
  }
  return words
}

/**
 * Draw one invisible run (scan edits in "recognized text only" mode): `text`
 * from `origin` along the unit `dir`, spread over `width`. Operators go to the
 * page's pdf-lib content stream, so a caller can roll them back like any other
 * drawing. `fontKey` comes from glyphlessFontKey(). Returns false when the
 * text has nothing to write.
 */
function writeInvisibleRun(page, fontKey, { text, origin, dir, width, fontSize, meta } = {}) {
  const {
    PDFHexString, PDFName, PDFOperator, PDFString,
    beginText, endText, popGraphicsState, pushGraphicsState, setCharacterSpacing, setCharacterSqueeze,
    setFontAndSize, setTextMatrix, setTextRenderingMode, setTextRise, setWordSpacing, showText, TextRenderingMode,
  } = getPdfLib()
  const [line] = validateLines([{
    fontSize,
    words: [{ text, x: origin?.x, y: origin?.y, dx: dir?.x, dy: dir?.y, width, gap: 0 }],
  }], 'the edited text')
  if (!line) return false
  const [run] = line.words
  const words = runWords(run.text, { x: run.x, y: run.y }, { x: run.dx, y: run.dy }, run.width, line.fontSize)
  const { engine, language } = sanitizeMeta(meta)
  const properties = page.doc.context.obj({ Engine: PDFString.of(engine), Lang: PDFString.of(language), Version: OCR_LAYER_VERSION })
  const name = typeof fontKey === 'string' ? fontKey.replace(/^\//, '') : fontKey
  const operators = [
    pushGraphicsState(),
    PDFOperator.of('BDC', [PDFName.of(OCR_MARK), properties]),
    beginText(),
    setTextRenderingMode(TextRenderingMode.Invisible),
    setCharacterSpacing(0),
    setWordSpacing(0),
    setTextRise(0),
    setFontAndSize(name, line.fontSize),
  ]
  for (const entry of lineRuns({ fontSize: line.fontSize, words })) {
    if (entry.matrix) operators.push(setTextMatrix(...entry.matrix))
    operators.push(setCharacterSpacing(entry.spacing), setCharacterSqueeze(entry.scale), showText(PDFHexString.of(entry.hex)))
  }
  operators.push(endText(), PDFOperator.of('EMC'), popGraphicsState())
  page.pushOperators(...operators)
  return true
}

const yieldToEventLoop = () => new Promise((resolve) => setImmediate(resolve))

/**
 * Glyphs drawn on a page (its content, not its annotations), from mupdf's
 * device interface: render mode 3 arrives as ignoreText, everything else
 * that marks or clips the page as fill/stroke/clip text.
 */
function pageGlyphs(mupdf, page) {
  const invisible = []
  const visible = []
  const collect = (target, text, ctm, withQuad) => {
    try {
      text.walk({
        showGlyph(font, trm, glyph, unicode, wmode) {
          const matrix = mupdf.Matrix.concat(trm, ctm)
          const entry = { x: matrix[4], y: matrix[5] }
          if (withQuad) {
            // The redaction box of a glyph runs from the origin along its
            // advance and from below the baseline to above it; this quad
            // stays inside it for any font ([ul, ur, ll, lr], mupdf order).
            const advance = glyph >= 0 ? font.advanceGlyph(glyph, wmode) : 0
            const point = (u, v) => [u * matrix[0] + v * matrix[2] + matrix[4], u * matrix[1] + v * matrix[3] + matrix[5]]
            if (advance > 0) entry.quad = [...point(0, 0.6), ...point(advance, 0.6), ...point(0, -0.1), ...point(advance, -0.1)]
            entry.unicode = unicode
          }
          target.push(entry)
        },
      })
    } finally {
      text.destroy()
    }
  }
  const device = new mupdf.Device({
    ignoreText: (text, ctm) => collect(invisible, text, ctm, true),
    fillText: (text, ctm) => collect(visible, text, ctm, false),
    strokeText: (text, _stroke, ctm) => collect(visible, text, ctm, false),
    clipText: (text, ctm) => collect(visible, text, ctm, false),
    clipStrokeText: (text, _stroke, ctm) => collect(visible, text, ctm, false),
  })
  try {
    page.runPageContents(device, mupdf.Matrix.identity)
  } finally {
    device.close()
    device.destroy()
  }
  return { invisible, visible }
}

/** Points bucketed on a grid so origins are matched without a quadratic scan. */
class PointGrid {
  constructor(cell = 1) {
    this.cell = cell
    this.cells = new Map()
  }

  key(x, y) {
    return `${Math.floor(x / this.cell)},${Math.floor(y / this.cell)}`
  }

  add(x, y, value) {
    const key = this.key(x, y)
    const list = this.cells.get(key)
    if (list) list.push(value)
    else this.cells.set(key, [value])
  }

  /** Values whose point lies within `tolerance` of (x, y). */
  near(x, y, tolerance, pointOf) {
    const found = []
    const x0 = Math.floor((x - tolerance) / this.cell)
    const x1 = Math.floor((x + tolerance) / this.cell)
    const y0 = Math.floor((y - tolerance) / this.cell)
    const y1 = Math.floor((y + tolerance) / this.cell)
    for (let gx = x0; gx <= x1; gx += 1) {
      for (let gy = y0; gy <= y1; gy += 1) {
        for (const value of this.cells.get(`${gx},${gy}`) || []) {
          const point = pointOf(value)
          if (Math.abs(point.x - x) <= tolerance && Math.abs(point.y - y) <= tolerance) found.push(value)
        }
      }
    }
    return found
  }
}

/** Glyph boxes indexed on a coarse grid for "is this point covered" checks. */
class BoxIndex {
  constructor(glyphs, cell = 24) {
    this.cell = cell
    this.cells = new Map()
    for (const glyph of glyphs) {
      const { x0, y0, x1, y1 } = glyph.bounds
      for (let gx = Math.floor(x0 / cell); gx <= Math.floor(x1 / cell); gx += 1) {
        for (let gy = Math.floor(y0 / cell); gy <= Math.floor(y1 / cell); gy += 1) {
          const key = `${gx},${gy}`
          const list = this.cells.get(key)
          if (list) list.push(glyph)
          else this.cells.set(key, [glyph])
        }
      }
    }
  }

  covers(x, y) {
    const list = this.cells.get(`${Math.floor(x / this.cell)},${Math.floor(y / this.cell)}`)
    return Boolean(list?.some((glyph) => x >= glyph.bounds.x0 && x <= glyph.bounds.x1
      && y >= glyph.bounds.y0 && y <= glyph.bounds.y1 && insideQuad(glyph.quad, x, y)))
  }
}

/** A point of `quad` that no protected glyph covers, or null. */
function freePoint(quad, protectedBoxes) {
  for (const [along, up] of PROBES) {
    const [x, y] = quadPoint(quad, along, up)
    if (!protectedBoxes.covers(x, y)) return [x, y]
  }
  return null
}

/**
 * Redaction quads for every invisible glyph of a page. A glyph is found in
 * mupdf's structured text by its origin, so its real box is known; glyphs the
 * structured text leaves out use a box inside their advance. Characters at an
 * origin where a visible glyph is also drawn are protected, and an invisible
 * glyph whose every probe point lies on a protected character is kept: visible
 * text is never removed with it.
 */
function invisibleRedactionQuads(mupdf, page, census) {
  const invisibleAt = new PointGrid()
  census.invisible.forEach((glyph, index) => invisibleAt.add(glyph.x, glyph.y, index))
  const visibleAt = new PointGrid()
  census.visible.forEach((glyph, index) => visibleAt.add(glyph.x, glyph.y, index))
  const originOf = (list) => (index) => list[index]

  const targets = []
  const protectedGlyphs = []
  const covered = new Set()
  const structured = page.toStructuredText('preserve-whitespace')
  try {
    structured.walk({
      onChar(_character, [x, y], _font, _size, quad) {
        const glyph = { quad: Array.from(quad), bounds: quadBounds(quad) }
        const invisible = invisibleAt.near(x, y, ORIGIN_TOLERANCE, originOf(census.invisible))
        const visible = visibleAt.near(x, y, ORIGIN_TOLERANCE, originOf(census.visible))
        if (invisible.length && !visible.length) {
          targets.push(glyph)
          for (const index of invisible) covered.add(index)
        } else {
          protectedGlyphs.push(glyph)
        }
      },
    })
  } finally {
    structured.destroy()
  }

  let kept = 0
  census.invisible.forEach((glyph, index) => {
    if (covered.has(index)) return
    if (visibleAt.near(glyph.x, glyph.y, ORIGIN_TOLERANCE, originOf(census.visible)).length) kept += 1
    else if (glyph.quad) targets.push({ quad: glyph.quad, bounds: quadBounds(glyph.quad) })
    else kept += 1
  })
  const protectedBoxes = new BoxIndex(protectedGlyphs)
  const quads = []
  for (const target of targets) {
    const point = freePoint(target.quad, protectedBoxes)
    if (point) quads.push(pointQuad(point))
    else kept += 1
  }
  return { quads, kept }
}

/**
 * Remove every invisible (render mode 3) glyph from the listed pages: earlier
 * OCR layers, Simple's or another tool's. Images, vector art, visible text and
 * annotations stay exactly as they are. Pages without invisible text are left
 * untouched, and so are the bytes when no page has any. Afterwards each page
 * is counted again; invisible glyphs that could not be removed without also
 * removing visible text are reported through `options.onKept`, and any other
 * glyph left behind fails with OCR_TEXT_NOT_REPLACED.
 */
async function removeInvisibleText(data, pageIndices, options = {}) {
  const indices = [...new Set(Array.isArray(pageIndices) ? pageIndices : [])].filter(Number.isInteger).sort((a, b) => a - b)
  if (!indices.length) return data
  const mupdf = await loadMupdf()
  const doc = mupdf.Document.openDocument(data, 'application/pdf')
  try {
    const pageCount = doc.countPages()
    let changed = false
    let first = true
    for (const pageIndex of indices) {
      if (pageIndex < 0 || pageIndex >= pageCount) continue
      // mupdf runs synchronously; let IPC and window events through between pages.
      if (!first) await yieldToEventLoop()
      first = false
      const page = doc.loadPage(pageIndex)
      try {
        const census = pageGlyphs(mupdf, page)
        if (!census.invisible.length) continue
        const { quads, kept } = invisibleRedactionQuads(mupdf, page, census)
        if (quads.length) {
          redactTextAt(mupdf, doc, page, quads)
          changed = true
        }
        const left = pageGlyphs(mupdf, page).invisible.length
        if (left > kept) {
          throw codedError('OCR_TEXT_NOT_REPLACED',
            `The text recognized earlier on page ${pageIndex + 1} could not be removed, so it was not replaced. Recognize the page again without replacing its text.`,
            { pageIndex, left, kept })
        }
        if (left > 0 && typeof options.onKept === 'function') options.onKept({ pageIndex, count: left })
      } finally {
        page.destroy()
      }
    }
    if (!changed) return data
    const buffer = doc.saveToBuffer('compress')
    try { return Uint8Array.from(buffer.asUint8Array()) } finally { buffer.destroy() }
  } finally {
    doc.destroy()
  }
}

/**
 * Invisible (render mode 3) and visible glyph counts of each listed page (all
 * pages by default), from the same census removeInvisibleText() uses.
 */
async function invisibleTextCensus(data, pageIndices) {
  const mupdf = await loadMupdf()
  const doc = mupdf.Document.openDocument(data, 'application/pdf')
  try {
    const pageCount = doc.countPages()
    const indices = Array.isArray(pageIndices) ? pageIndices : Array.from({ length: pageCount }, (_, index) => index)
    const result = []
    for (const pageIndex of indices) {
      if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= pageCount) continue
      const page = doc.loadPage(pageIndex)
      try {
        const census = pageGlyphs(mupdf, page)
        result.push({ pageIndex, invisible: census.invisible.length, visible: census.visible.length })
      } finally {
        page.destroy()
      }
    }
    return result
  } finally {
    doc.destroy()
  }
}

module.exports = {
  GLYPHLESS_FONT_NAME,
  GLYPHLESS_TTF_BASE64,
  LIMITS,
  OCR_MARK,
  addGlyphlessFont,
  addOcrTextLayer,
  glyphlessFontBytes,
  glyphlessFontKey,
  invisibleTextCensus,
  normalizeLayerText,
  ocrLayerContent,
  removeInvisibleText,
  toUnicodeCMap,
  utf16Hex,
  validateOcrLayerOperation,
  writeInvisibleRun,
}
