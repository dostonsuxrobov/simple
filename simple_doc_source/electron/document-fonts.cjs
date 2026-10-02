const fs = require('node:fs/promises')
const path = require('node:path')
const { fontCoverage } = require('./font-coverage.cjs')

// Use fonts already licensed and installed on this computer. Never bundle or
// copy Windows font files into a release. Missing families retain engine fallbacks.
const FONT_FAMILIES = [
  ['Arial', ['arial.ttf', 'arialbd.ttf', 'ariali.ttf', 'arialbi.ttf']],
  ['Calibri', ['calibri.ttf', 'calibrib.ttf', 'calibrii.ttf', 'calibriz.ttf']],
  ['Cambria', ['cambria.ttf', 'cambriab.ttf', 'cambriai.ttf', 'cambriaz.ttf']],
  ['Times New Roman', ['times.ttf', 'timesbd.ttf', 'timesi.ttf', 'timesbi.ttf']],
  ['Courier New', ['cour.ttf', 'courbd.ttf', 'couri.ttf', 'courbi.ttf']],
  ['Georgia', ['georgia.ttf', 'georgiab.ttf', 'georgiai.ttf', 'georgiaz.ttf']],
  ['Verdana', ['verdana.ttf', 'verdanab.ttf', 'verdanai.ttf', 'verdanaz.ttf']],
  ['Tahoma', ['tahoma.ttf', 'tahomabd.ttf']],
  ['Segoe UI', ['segoeui.ttf', 'segoeuib.ttf', 'segoeuii.ttf', 'segoeuiz.ttf']],
  ['Trebuchet MS', ['trebuc.ttf', 'trebucbd.ttf', 'trebucit.ttf', 'trebucbi.ttf']],
  ['Consolas', ['consola.ttf', 'consolab.ttf', 'consolai.ttf', 'consolaz.ttf']],
  ['Garamond', ['GARA.TTF', 'GARABD.TTF', 'GARAIT.TTF']],
  ['Arial Narrow', ['ARIALN.TTF', 'ARIALNB.TTF', 'ARIALNI.TTF', 'ARIALNBI.TTF']],
  ['Aptos', ['aptos.ttf', 'aptos-bold.ttf', 'aptos-italic.ttf', 'aptos-bolditalic.ttf']],
  ['Malgun Gothic', ['malgun.ttf', 'malgunbd.ttf']],
]
const FACE_STYLES = ['regular', 'bold', 'italic', 'boldItalic']
const SFNT_VERSIONS = [0x00010000, 0x4f54544f, 0x74727565]
/** The engine refuses font payloads over 16 MB (PDF export), so larger files are never offered. */
const MAX_FONT_BYTES = 16 * 1024 * 1024
/** Families beyond the curated set that one window may load eagerly (recent and document fonts). */
const MAX_EXTRA_FAMILIES = 32

function defaultFontRoots() {
  return [
    path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts'),
    ...(process.env.LOCALAPPDATA ? [path.join(process.env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts')] : []),
  ]
}

/** Line metrics from the head, hhea and OS/2 tables, or null when the face is unusable or may not be embedded. */
function metricsFromTables(head, hhea, os2) {
  if (!head || head.length < 20 || !hhea || hhea.length < 10) return null
  // Restricted or bitmap-only embedding cannot be used by the vector PDF exporter.
  if (os2?.length >= 10 && (os2.readUInt16BE(8) & 0x0202)) return null
  const units = head.readUInt16BE(18)
  if (!units) return null
  const hheaAscent = hhea.readInt16BE(4)
  const hheaDescent = Math.abs(hhea.readInt16BE(6))
  const lineGap = Math.max(0, hhea.readInt16BE(8))
  // Word sizes a single-spaced line like Windows GDI: usWinAscent + usWinDescent
  // plus the external leading that the hhea line gap adds beyond them. Using the
  // bare hhea ascender/descender made Calibri lines 1.0 em instead of Word's
  // 1.2207 em. The external leading is carried below the baseline.
  let ascentUnits = hheaAscent
  let descentUnits = hheaDescent + lineGap
  if (os2?.length >= 78) {
    const winAscent = os2.readUInt16BE(74)
    const winDescent = os2.readUInt16BE(76)
    if (winAscent > 0 && winAscent + winDescent > 0) {
      const externalLeading = Math.max(0, lineGap - ((winAscent + winDescent) - (hheaAscent + hheaDescent)))
      ascentUnits = winAscent
      descentUnits = winDescent + externalLeading
    }
  }
  const ascent = ascentUnits / units
  const descent = descentUnits / units
  if (ascent <= 0 || ascent > 4 || descent < 0 || descent > 4) return null
  return { ascent, descent }
}

function fontMetrics(value) {
  const bytes = Buffer.from(value)
  if (bytes.length < 12 || !SFNT_VERSIONS.includes(bytes.readUInt32BE(0))) return null
  const count = bytes.readUInt16BE(4)
  if (count > 256 || bytes.length < 12 + count * 16) return null
  const tables = new Map()
  for (let index = 0; index < count; index += 1) {
    const position = 12 + index * 16
    const offset = bytes.readUInt32BE(position + 8)
    const length = bytes.readUInt32BE(position + 12)
    if (offset + length > bytes.length) return null
    tables.set(bytes.toString('ascii', position, position + 4), bytes.subarray(offset, offset + length))
  }
  return metricsFromTables(tables.get('head'), tables.get('hhea'), tables.get('OS/2'))
}

// ---------------------------------------------------------------------------------------
// Installed font families (every TTF/OTF on this computer), for the font list

/** UTF-16BE (Windows and Unicode name records) or Mac Roman ASCII text. */
function decodeName(record) {
  if (record.platform === 3 || record.platform === 0) {
    if (record.bytes.length % 2) return null
    return Buffer.from(record.bytes).swap16().toString('utf16le')
  }
  if (record.platform === 1 && record.encoding === 0) return record.bytes.toString('latin1')
  return null
}

/**
 * The family (name ID 1) and style (name ID 2) a face reports, preferring the Windows
 * English (United States) names that Word and Windows show in their font lists.
 */
function fontNames(name) {
  if (!name || name.length < 6) return { family: null, subfamily: null }
  const count = name.readUInt16BE(2)
  const storage = name.readUInt16BE(4)
  const records = []
  for (let index = 0; index < Math.min(count, 2048); index += 1) {
    const at = 6 + index * 12
    if (at + 12 > name.length) break
    const start = storage + name.readUInt16BE(at + 10)
    const length = name.readUInt16BE(at + 8)
    if (start + length > name.length) continue
    records.push({ platform: name.readUInt16BE(at), encoding: name.readUInt16BE(at + 2), language: name.readUInt16BE(at + 4), id: name.readUInt16BE(at + 6), bytes: name.subarray(start, start + length) })
  }
  const preferences = [
    (record) => record.platform === 3 && record.language === 0x0409,
    (record) => record.platform === 3 && (record.language & 0xff) === 0x09,
    (record) => record.platform === 3,
    (record) => record.platform === 0,
    (record) => record.platform === 1 && record.language === 0,
  ]
  const pick = (id) => {
    const candidates = records.filter((record) => record.id === id)
    for (const preferred of preferences) {
      for (const record of candidates.filter(preferred)) {
        const value = decodeName(record)?.replace(/\u0000/g, '').trim()
        if (value) return value
      }
    }
    return null
  }
  return { family: pick(1), subfamily: pick(2) }
}

/** True when the face maps Unicode characters (symbol-encoded fonts such as Wingdings do not). */
function hasUnicodeCmap(cmap) {
  if (!cmap || cmap.length < 4) return false
  const count = cmap.readUInt16BE(2)
  for (let index = 0; index < Math.min(count, 64); index += 1) {
    const at = 4 + index * 8
    if (at + 8 > cmap.length) break
    const platform = cmap.readUInt16BE(at)
    const encoding = cmap.readUInt16BE(at + 2)
    if (platform === 0 || (platform === 3 && (encoding === 1 || encoding === 10))) return true
  }
  return false
}

/** Which of the four Word styles a face is, from OS/2 fsSelection (or head macStyle). */
function faceStyle(head, os2) {
  let bold = false
  let italic = false
  if (os2?.length >= 64) {
    const selection = os2.readUInt16BE(62)
    bold = Boolean(selection & 0x20)
    italic = Boolean(selection & 0x01)
  } else if (head?.length >= 46) {
    const macStyle = head.readUInt16BE(44)
    bold = Boolean(macStyle & 0x01)
    italic = Boolean(macStyle & 0x02)
  }
  return bold ? (italic ? 'boldItalic' : 'bold') : italic ? 'italic' : 'regular'
}

/** Reads only the table directory and the small tables a face's name, style and metrics need. */
async function readFaceTables(filePath) {
  const handle = await fs.open(filePath, 'r')
  try {
    const stat = await handle.stat()
    if (!stat.isFile() || stat.size < 12 || stat.size > MAX_FONT_BYTES) return null
    const header = Buffer.alloc(12)
    await handle.read(header, 0, 12, 0)
    if (!SFNT_VERSIONS.includes(header.readUInt32BE(0))) return null
    const count = header.readUInt16BE(4)
    if (!count || count > 256) return null
    const directory = Buffer.alloc(count * 16)
    await handle.read(directory, 0, directory.length, 12)
    const tables = new Map()
    for (let index = 0; index < count; index += 1) {
      const at = index * 16
      const tag = directory.toString('latin1', at, at + 4)
      if (!['name', 'head', 'hhea', 'OS/2', 'cmap'].includes(tag)) continue
      const offset = directory.readUInt32BE(at + 8)
      let length = directory.readUInt32BE(at + 12)
      if (offset + length > stat.size) continue
      // Only the cmap encoding records are needed, never its glyph maps.
      if (tag === 'cmap') length = Math.min(length, 4 + 64 * 8)
      if (length > 512 * 1024) continue
      const data = Buffer.alloc(length)
      await handle.read(data, 0, length, offset)
      tables.set(tag, data)
    }
    return { tables, size: stat.size }
  } finally {
    await handle.close()
  }
}

/** One installed face: family, style and metrics, or null for collections, symbol fonts and restricted faces. */
async function scanFontFile(filePath) {
  try {
    const read = await readFaceTables(filePath)
    if (!read) return null
    const { tables } = read
    const { family, subfamily } = fontNames(tables.get('name'))
    if (!family || family.startsWith('.') || family.length > 80) return null
    if (!hasUnicodeCmap(tables.get('cmap'))) return null
    const metrics = metricsFromTables(tables.get('head'), tables.get('hhea'), tables.get('OS/2'))
    if (!metrics) return null
    const os2 = tables.get('OS/2')
    return {
      file: filePath,
      family,
      subfamily: subfamily || '',
      style: faceStyle(tables.get('head'), os2),
      weight: os2?.length >= 6 ? os2.readUInt16BE(4) : 400,
      metrics,
    }
  } catch {
    return null
  }
}

const IDEAL_WEIGHT = { regular: 400, italic: 400, bold: 700, boldItalic: 700 }

/** Groups faces into families with Word's four style slots (the face closest to 400/700 wins a slot). */
function groupFaces(faces) {
  const families = new Map()
  for (const face of faces) {
    if (!face) continue
    const key = face.family.toLowerCase()
    let entry = families.get(key)
    if (!entry) {
      entry = { family: face.family, faces: {} }
      families.set(key, entry)
    }
    const current = entry.faces[face.style]
    const distance = (candidate) => Math.abs((candidate.weight || 400) - IDEAL_WEIGHT[face.style])
    if (!current || distance(face) < distance(current) || (distance(face) === distance(current) && face.file < current.file)) entry.faces[face.style] = face
  }
  for (const entry of families.values()) {
    // A family without a regular face (Arial Black, a single-weight font) uses its first face.
    if (!entry.faces.regular) {
      const first = FACE_STYLES.map((style) => entry.faces[style]).find(Boolean)
      if (first) entry.faces = { regular: first }
    }
  }
  return [...families.values()].sort((a, b) => a.family.localeCompare(b.family, 'en', { sensitivity: 'base' }))
}

/** Every TTF/OTF family installed for all users and for the current user. */
async function scanInstalledFontFamilies(options = {}) {
  const roots = options.roots || defaultFontRoots()
  const files = []
  for (const root of roots) {
    let names = []
    try { names = await fs.readdir(root) } catch { continue }
    for (const name of names) if (/\.(?:ttf|otf)$/i.test(name)) files.push(path.join(root, name))
  }
  const faces = []
  const concurrency = Math.max(1, options.concurrency || 12)
  for (let index = 0; index < files.length; index += concurrency) {
    faces.push(...await Promise.all(files.slice(index, index + concurrency).map(scanFontFile)))
  }
  return groupFaces(faces)
}

async function installedDocumentFonts(options = {}) {
  const roots = options.roots || defaultFontRoots()
  const files = options.files || new Map()
  const fonts = []
  for (const [family, names] of FONT_FAMILIES) {
    const faces = {}
    let sizing = null
    let simpleCoverage
    for (let index = 0; index < names.length; index += 1) {
      for (const root of roots) {
        try {
          const filePath = path.join(root, names[index])
          const stat = await fs.stat(filePath)
          if (!stat.isFile() || stat.size > MAX_FONT_BYTES) continue
          const bytes = await fs.readFile(filePath)
          const metrics = fontMetrics(bytes)
          if (!metrics) continue
          const token = String(files.size)
          files.set(token, filePath)
          faces[FACE_STYLES[index]] = `simple-font://installed/${token}`
          if (index === 0) {
            sizing = metrics
            if (family === 'Malgun Gothic') simpleCoverage = fontCoverage(bytes)
          }
          break
        } catch { /* Uninstalled fonts use the bundled fallback. */ }
      }
    }
    if (faces.regular && sizing) fonts.push({ family, faces, sizing, ...simpleCoverage ? { simpleCoverage } : {} })
  }
  return { fonts, files }
}

const familyKey = (family) => String(family ?? '').trim().replace(/\s+/g, ' ').toLowerCase()

/**
 * The fonts one main process serves: the curated set (loaded by every window), the
 * installed family list (scanned once, on first use) and, on request, any other
 * installed family as an engine font definition. Faces are served by token through
 * the simple-font:// protocol; font files never leave this computer.
 */
function createFontCatalog(options = {}) {
  const roots = options.roots || defaultFontRoots()
  const files = new Map()
  const tokens = new Map()
  let curated = null
  let scanned = null
  const tokenFor = (filePath) => {
    let token = tokens.get(filePath)
    if (token === undefined) {
      token = String(files.size)
      files.set(token, filePath)
      tokens.set(filePath, token)
    }
    return token
  }
  const curatedFonts = () => curated ||= installedDocumentFonts({ roots, files }).then((result) => {
    for (const [token, filePath] of result.files) tokens.set(filePath, token)
    return result.fonts
  })
  const families = () => scanned ||= scanInstalledFontFamilies({ roots }).catch(() => [])
  const definition = (entry) => {
    const regular = entry.faces.regular
    if (!regular) return null
    const faces = {}
    for (const style of FACE_STYLES) if (entry.faces[style]) faces[style] = `simple-font://installed/${tokenFor(entry.faces[style].file)}`
    return { family: entry.family, faces, sizing: regular.metrics }
  }
  return {
    /** Engine font definitions: the curated set plus the requested installed families (unknown ones are skipped). */
    async documentFonts(requested = []) {
      const fonts = [...await curatedFonts()]
      const wanted = [...new Set((Array.isArray(requested) ? requested : []).map(familyKey).filter(Boolean))]
        .filter((key) => !fonts.some((font) => familyKey(font.family) === key))
        .slice(0, MAX_EXTRA_FAMILIES)
      if (!wanted.length) return fonts
      const byKey = new Map((await families()).map((entry) => [familyKey(entry.family), entry]))
      for (const key of wanted) {
        const entry = byKey.get(key)
        const font = entry ? definition(entry) : null
        if (font) fonts.push(font)
      }
      return fonts
    },
    /** Every installed family Simple can show and embed, by name, with how many styles it has. */
    async familyList() {
      return (await families()).map((entry) => ({ family: entry.family, styles: FACE_STYLES.filter((style) => entry.faces[style]).length }))
    },
    /** The font file behind a simple-font:// token, or null. */
    async fileForToken(token) {
      await curatedFonts()
      return files.get(String(token)) ?? null
    },
  }
}

module.exports = { createFontCatalog, faceStyle, fontMetrics, fontNames, groupFaces, hasUnicodeCmap, installedDocumentFonts, scanFontFile, scanInstalledFontFamilies }
