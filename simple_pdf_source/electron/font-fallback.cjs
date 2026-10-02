'use strict'

/**
 * Per-character font fallback from installed Windows fonts.
 *
 * pdf-lib draws every character with the one font it is given; characters
 * that font lacks become glyph 0 (.notdef), which most viewers draw as
 * nothing. The editor (Chromium) falls back per character, so '你好' or '✅'
 * looked right on screen and vanished in the saved file. FontRuns splits text
 * into runs of characters that one font can draw, choosing the first
 * installed face that has each character.
 */
const fs = require('node:fs/promises')
const path = require('node:path')
const { embedPdfFont } = require('./font-embedding.cjs')

let fontkitModule
function getFontkit() {
  fontkitModule ||= require('@pdf-lib/fontkit')
  return fontkitModule
}

function fontsDirectory() {
  return path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts')
}

const CJK_FACES = {
  han: { key: 'microsoft-yahei', files: { regular: 'msyh.ttc', bold: 'msyhbd.ttc' } },
  kana: { key: 'yu-gothic', files: { regular: 'YuGothM.ttc', bold: 'YuGothB.ttc' } },
  hangul: { key: 'malgun-gothic', files: { regular: 'malgun.ttf', bold: 'malgunbd.ttf' } },
}

// Faces tried, in order, for characters the requested font cannot draw.
// Every entry is optional: missing files are skipped.
const FALLBACK_FACES = [
  { key: 'segoe-ui', files: { regular: 'segoeui.ttf', bold: 'segoeuib.ttf', italic: 'segoeuii.ttf', boldItalic: 'segoeuiz.ttf' } },
  { key: 'segoe-ui-symbol', files: { regular: 'seguisym.ttf' } },
  'cjk',
  { key: 'microsoft-jhenghei', files: { regular: 'msjh.ttc', bold: 'msjhbd.ttc' } },
  { key: 'nirmala-ui', files: { regular: ['Nirmala.ttc', 'Nirmala.ttf'], bold: ['NirmalaB.ttc', 'NirmalaB.ttf'] } },
  { key: 'leelawadee-ui', files: { regular: 'LeelawUI.ttf', bold: 'LeelaUIb.ttf' } },
  { key: 'ebrima', files: { regular: 'ebrima.ttf', bold: 'ebrimabd.ttf' } },
  { key: 'gadugi', files: { regular: 'gadugi.ttf', bold: 'gadugib.ttf' } },
  { key: 'sylfaen', files: { regular: 'sylfaen.ttf' } },
  { key: 'myanmar-text', files: { regular: 'mmrtext.ttf', bold: 'mmrtextb.ttf' } },
  { key: 'microsoft-himalaya', files: { regular: 'himalaya.ttf' } },
  { key: 'mongolian-baiti', files: { regular: 'monbaiti.ttf' } },
  { key: 'javanese-text', files: { regular: 'javatext.ttf' } },
  { key: 'segoe-ui-historic', files: { regular: 'seguihis.ttf' } },
  { key: 'simsun', files: { regular: 'simsun.ttc' } },
  { key: 'simsun-extb', files: { regular: 'simsunb.ttf' } },
  { key: 'simsun-extg', files: { regular: 'SimsunExtG.ttf' } },
  { key: 'mingliu-extb', files: { regular: 'mingliub.ttc' } },
  // Monochrome outlines of the emoji; colour layers are not used.
  { key: 'segoe-ui-emoji', files: { regular: 'seguiemj.ttf' } },
  { key: 'arial-unicode', files: { regular: 'ARIALUNI.TTF' } },
]

const KANA = /[\u3040-\u30ff\u31f0-\u31ff\uff66-\uff9f]/u
const HANGUL = /[\u1100-\u11ff\u3130-\u318f\uac00-\ud7af]/u
const COMBINING_MARK = /\p{M}/u

function scriptHints(text) {
  const value = String(text || '')
  return { kana: KANA.test(value), hangul: HANGUL.test(value) }
}

function orderedFaces(hints = {}) {
  const cjk = hints.kana
    ? [CJK_FACES.kana, CJK_FACES.han, CJK_FACES.hangul]
    : hints.hangul
      ? [CJK_FACES.hangul, CJK_FACES.han, CJK_FACES.kana]
      : [CJK_FACES.han, CJK_FACES.kana, CJK_FACES.hangul]
  return FALLBACK_FACES.flatMap((face) => (face === 'cjk' ? cjk : [face]))
}

function styleKey(style = {}) {
  return style.bold && style.italic ? 'boldItalic' : style.bold ? 'bold' : style.italic ? 'italic' : 'regular'
}

function faceFiles(face, style) {
  const files = face.files[styleKey(style)] || (style.bold ? face.files.bold : undefined) || face.files.regular
  return Array.isArray(files) ? files : [files]
}

/** Characters with no visible glyph of their own (joiners, variation selectors…). */
function isDefaultIgnorable(codePoint) {
  return codePoint === 0xad || codePoint === 0x34f || codePoint === 0x61c
    || (codePoint >= 0x115f && codePoint <= 0x1160) || (codePoint >= 0x17b4 && codePoint <= 0x17b5)
    || (codePoint >= 0x180b && codePoint <= 0x180f) || (codePoint >= 0x200b && codePoint <= 0x200f)
    || (codePoint >= 0x202a && codePoint <= 0x202e) || (codePoint >= 0x2060 && codePoint <= 0x206f)
    || codePoint === 0x3164 || (codePoint >= 0xfe00 && codePoint <= 0xfe0f) || codePoint === 0xfeff
    || codePoint === 0xffa0 || (codePoint >= 0xfff0 && codePoint <= 0xfff8)
    || (codePoint >= 0x1bca0 && codePoint <= 0x1bca3) || (codePoint >= 0x1d173 && codePoint <= 0x1d17a)
    || (codePoint >= 0xe0000 && codePoint <= 0xe0fff)
}

function isControl(codePoint) {
  return (codePoint < 0x20 && codePoint !== 0x09) || (codePoint >= 0x7f && codePoint < 0xa0)
}

// Process-wide cache of parsed system fonts. Parsing is lazy (tables are read
// on demand) and only faces that are actually needed are opened. The cache is
// dropped after a minute of inactivity so large CJK fonts do not stay in memory.
const faceCache = new Map()
let faceCacheTimer = null
const FACE_CACHE_IDLE_MS = 60 * 1000

function touchFaceCache() {
  if (faceCacheTimer) clearTimeout(faceCacheTimer)
  faceCacheTimer = setTimeout(() => {
    faceCache.clear()
    faceCacheTimer = null
  }, FACE_CACHE_IDLE_MS)
  faceCacheTimer.unref?.()
}

async function loadFaceFile(fileName) {
  const filePath = path.isAbsolute(fileName) ? fileName : path.join(fontsDirectory(), fileName)
  const key = filePath.toLowerCase()
  touchFaceCache()
  if (!faceCache.has(key)) {
    faceCache.set(key, (async () => {
      try {
        const raw = await fs.readFile(filePath)
        const parsed = getFontkit().create(raw)
        const font = Array.isArray(parsed?.fonts) ? parsed.fonts[0] : parsed
        if (!font || typeof font.hasGlyphForCodePoint !== 'function') return null
        return { filePath, raw, font, collection: Array.isArray(parsed?.fonts), standalone: null }
      } catch {
        return null
      }
    })())
  }
  return faceCache.get(key)
}

async function loadFace(face, style) {
  for (const file of faceFiles(face, style)) {
    const loaded = await loadFaceFile(file)
    if (loaded) return loaded
  }
  return null
}

/**
 * pdf-lib embeds single fonts only. Copy the first font of a TrueType
 * collection (.ttc) into a standalone sfnt; tables are copied verbatim.
 */
function extractCollectionFont(value, fontIndex = 0) {
  const bytes = Buffer.isBuffer(value) ? value : Buffer.from(value)
  if (bytes.toString('latin1', 0, 4) !== 'ttcf') return bytes
  const count = bytes.readUInt32BE(8)
  if (fontIndex >= count) throw new RangeError('The font collection does not contain that font.')
  const offset = bytes.readUInt32BE(12 + 4 * fontIndex)
  const sfntVersion = bytes.readUInt32BE(offset)
  const tableCount = bytes.readUInt16BE(offset + 4)
  const tables = []
  for (let index = 0; index < tableCount; index += 1) {
    const record = offset + 12 + 16 * index
    tables.push({
      tag: bytes.toString('latin1', record, record + 4),
      checksum: bytes.readUInt32BE(record + 4),
      offset: bytes.readUInt32BE(record + 8),
      length: bytes.readUInt32BE(record + 12),
    })
  }
  const headerSize = 12 + 16 * tableCount
  const total = tables.reduce((size, table) => size + ((table.length + 3) & ~3), headerSize)
  const output = Buffer.alloc(total)
  const entrySelector = Math.floor(Math.log2(Math.max(1, tableCount)))
  const searchRange = 2 ** entrySelector * 16
  output.writeUInt32BE(sfntVersion, 0)
  output.writeUInt16BE(tableCount, 4)
  output.writeUInt16BE(searchRange, 6)
  output.writeUInt16BE(entrySelector, 8)
  output.writeUInt16BE(tableCount * 16 - searchRange, 10)
  let cursor = headerSize
  tables.forEach((table, index) => {
    const record = 12 + 16 * index
    output.write(table.tag, record, 'latin1')
    output.writeUInt32BE(table.checksum, record + 4)
    output.writeUInt32BE(cursor, record + 8)
    output.writeUInt32BE(table.length, record + 12)
    bytes.copy(output, cursor, table.offset, table.offset + table.length)
    cursor += (table.length + 3) & ~3
  })
  return output
}

function standaloneBytes(face) {
  if (!face.standalone) face.standalone = face.collection ? extractCollectionFont(face.raw) : face.raw
  return face.standalone
}

/** Whether a pdf-lib font can draw a code point. */
function fontCovers(font, codePoint) {
  const embedder = font?.embedder
  if (typeof embedder?.font?.hasGlyphForCodePoint === 'function') return embedder.font.hasGlyphForCodePoint(codePoint)
  if (typeof embedder?.encoding?.canEncodeUnicodeCodePoint === 'function') return embedder.encoding.canEncodeUnicodeCodePoint(codePoint)
  try {
    font.encodeText(String.fromCodePoint(codePoint))
    return true
  } catch {
    return false
  }
}

/** Embedded fallback fonts for one pdf-lib document. */
class FontLibrary {
  constructor(pdfDoc) {
    this.pdfDoc = pdfDoc
    this.embedded = new Map()
  }

  async embedFace(face) {
    const key = face.filePath.toLowerCase()
    if (!this.embedded.has(key)) {
      this.embedded.set(key, (async () => {
        this.pdfDoc.registerFontkit(getFontkit())
        return embedPdfFont(this.pdfDoc, standaloneBytes(face))
      })().catch(() => null))
    }
    return this.embedded.get(key)
  }

  /** First installed face, in fallback order, that can draw `codePoint`. */
  async fallbackFor(codePoint, style = {}, hints = {}) {
    // Supplementary Private Use planes: no system font defines them.
    if (codePoint >= 0xf0000) return null
    for (const face of orderedFaces(hints)) {
      const loaded = await loadFace(face, style)
      if (!loaded || !loaded.font.hasGlyphForCodePoint(codePoint)) continue
      const font = await this.embedFace(loaded)
      if (font) return font
    }
    return null
  }

  /**
   * One font able to draw all of `text`, for widgets (form field appearances
   * use a single font): `preferred` faces first, then the fallback order.
   * @returns {Promise<{ font: import('pdf-lib').PDFFont, complete: boolean } | null>}
   */
  async singleFontFor(text, { preferred = [], style = {} } = {}) {
    const all = [...new Set(Array.from(String(text || ''), (character) => character.codePointAt(0)))]
      .filter((codePoint) => !isControl(codePoint) && !isDefaultIgnorable(codePoint) && codePoint !== 0x09)
    // Supplementary Private Use characters exist in no system font: choose
    // the font for the rest, and report the result as incomplete.
    const codePoints = all.filter((codePoint) => codePoint < 0xf0000)
    const complete = codePoints.length === all.length
    let best = null
    const candidates = [...preferred, ...orderedFaces(scriptHints(text))]
    for (const candidate of candidates) {
      if (candidate?.pdfFont) {
        // A base-14 font cannot even encode characters outside its set.
        if (!complete && typeof candidate.pdfFont.embedder?.font?.hasGlyphForCodePoint !== 'function') continue
        const covered = codePoints.filter((codePoint) => fontCovers(candidate.pdfFont, codePoint)).length
        if (covered === codePoints.length) return { font: candidate.pdfFont, complete }
        if (!best || covered > best.covered) best = { covered, font: candidate.pdfFont }
        continue
      }
      const loaded = candidate?.filePath ? candidate : await loadFace(candidate, style)
      if (!loaded) continue
      const covered = codePoints.filter((codePoint) => loaded.font.hasGlyphForCodePoint(codePoint)).length
      if (covered === codePoints.length) {
        const font = await this.embedFace(loaded)
        if (font) return { font, complete }
      }
      if (!best || covered > best.covered) best = { covered, face: loaded }
    }
    if (!best) return null
    const font = best.font || await this.embedFace(best.face)
    return font ? { font, complete: false } : null
  }

  /** Prepare runs for `text` drawn with `primary`, falling back per character. */
  async runsFor(primary, text, style = {}) {
    const runs = new FontRuns(primary)
    await runs.cover(this, text, style)
    return runs
  }
}

/**
 * Text measured and drawn as runs of one font each. Measuring and drawing
 * use the same runs, so layout stays consistent with the output.
 */
class FontRuns {
  constructor(primary) {
    this.primary = primary
    this.assigned = new Map()
    this.missing = new Set()
    this.replacement = null
  }

  async cover(library, text, style = {}) {
    const hints = scriptHints(text)
    for (const character of String(text || '')) {
      const codePoint = character.codePointAt(0)
      if (this.assigned.has(codePoint)) continue
      if (codePoint === 0x09 || codePoint === 0x0a || codePoint === 0x0d || isControl(codePoint) || isDefaultIgnorable(codePoint)) {
        this.assigned.set(codePoint, undefined)
        continue
      }
      if (fontCovers(this.primary, codePoint)) {
        this.assigned.set(codePoint, this.primary)
        continue
      }
      const fallback = await library.fallbackFor(codePoint, style, hints)
      this.assigned.set(codePoint, fallback || null)
      if (!fallback) this.missing.add(character)
    }
    if (this.missing.size && !this.replacement) {
      this.replacement = fontCovers(this.primary, 0xfffd) ? this.primary : await library.fallbackFor(0xfffd, style, hints)
    }
    return this
  }

  get missingCharacters() {
    return [...this.missing]
  }

  /** @returns {Array<{ font: import('pdf-lib').PDFFont, text: string }>} */
  runs(text) {
    const runs = []
    let current = null
    const append = (font, value) => {
      if (!current || current.font !== font) {
        current = { font, text: '' }
        runs.push(current)
      }
      current.text += value
    }
    for (const character of String(text || '')) {
      const codePoint = character.codePointAt(0)
      if (codePoint === 0x0a || codePoint === 0x0d || isControl(codePoint)) continue
      if (codePoint === 0x09) {
        const font = current?.font && fontCovers(current.font, 0x20) ? current.font : this.primary
        append(font, ' ')
        continue
      }
      if (isDefaultIgnorable(codePoint)) {
        // Joiners and variation selectors stay with their base character
        // when that font maps them; otherwise they are dropped, never tofu.
        if (current && fontCovers(current.font, codePoint)) current.text += character
        continue
      }
      if (current && COMBINING_MARK.test(character) && fontCovers(current.font, codePoint)) {
        current.text += character
        continue
      }
      let font = this.assigned.has(codePoint) ? this.assigned.get(codePoint) : this.primary
      if (font === undefined) font = this.primary
      if (font === null) {
        if (!this.replacement) continue
        append(this.replacement, '\ufffd')
        continue
      }
      append(font, character)
    }
    return runs
  }

  widthOfTextAtSize(text, size) {
    return this.runs(text).reduce((width, run) => width + run.font.widthOfTextAtSize(run.text, size), 0)
  }

  heightAtSize(size, options) {
    return this.primary.heightAtSize(size, options)
  }
}

module.exports = {
  FontLibrary,
  FontRuns,
  extractCollectionFont,
  fontCovers,
  isDefaultIgnorable,
  loadFaceFile,
  scriptHints,
}
