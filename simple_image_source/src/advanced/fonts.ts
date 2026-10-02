// src/advanced/fonts.ts (WP7)
// Font names between Photoshop and the browser (design 5.13). PSD text stores PostScript names
// ('ArialMT', 'SegoeUI-Bold'); canvas text needs a CSS family plus weight and italic.
//   - A built-in table covers the curated Windows families (and common Adobe defaults) both ways.
//   - loadLocalFonts() asks Chromium for the installed fonts (window.queryLocalFonts; Chromium may need a
//     user gesture the first time) and from then on maps every installed PostScript name exactly.
//   - Anything else uses a heuristic: strip 'MT' / 'PS' / 'PSMT', split the '-Style' suffix and CamelCase.
// No DOM access at module load (Node tests import it).
import type { PsdFontChoice, ResolvedFont } from './psdMapping.ts'
import type { TextStyle } from './types.ts'
import { FONT_FAMILIES } from '../shared/vector.ts'

interface Face {
  readonly postScriptName: string
  readonly family: string
  readonly weight: 400 | 700
  readonly italic: boolean
}

type FaceStyle = 'regular' | 'bold' | 'italic' | 'boldItalic'

function face(postScriptName: string, family: string, style: FaceStyle): Face {
  return { postScriptName, family, weight: style === 'bold' || style === 'boldItalic' ? 700 : 400, italic: style === 'italic' || style === 'boldItalic' }
}

/** The usual four faces of a Windows family whose PostScript names are Family, Family-Bold, ... */
function fourFaces(family: string, compact = family.replace(/\s+/g, '')): Face[] {
  return [
    face(compact, family, 'regular'),
    face(`${compact}-Bold`, family, 'bold'),
    face(`${compact}-Italic`, family, 'italic'),
    face(`${compact}-BoldItalic`, family, 'boldItalic'),
  ]
}

const KNOWN_FACES: readonly Face[] = Object.freeze([
  face('ArialMT', 'Arial', 'regular'),
  face('Arial-BoldMT', 'Arial', 'bold'),
  face('Arial-ItalicMT', 'Arial', 'italic'),
  face('Arial-BoldItalicMT', 'Arial', 'boldItalic'),
  ...fourFaces('Segoe UI'),
  ...fourFaces('Calibri'),
  ...fourFaces('Cambria'),
  ...fourFaces('Candara'),
  ...fourFaces('Consolas'),
  ...fourFaces('Constantia'),
  ...fourFaces('Corbel'),
  ...fourFaces('Georgia'),
  ...fourFaces('Verdana'),
  ...fourFaces('Sitka Text'),
  face('CourierNewPSMT', 'Courier New', 'regular'),
  face('CourierNewPS-BoldMT', 'Courier New', 'bold'),
  face('CourierNewPS-ItalicMT', 'Courier New', 'italic'),
  face('CourierNewPS-BoldItalicMT', 'Courier New', 'boldItalic'),
  face('TimesNewRomanPSMT', 'Times New Roman', 'regular'),
  face('TimesNewRomanPS-BoldMT', 'Times New Roman', 'bold'),
  face('TimesNewRomanPS-ItalicMT', 'Times New Roman', 'italic'),
  face('TimesNewRomanPS-BoldItalicMT', 'Times New Roman', 'boldItalic'),
  face('TrebuchetMS', 'Trebuchet MS', 'regular'),
  face('TrebuchetMS-Bold', 'Trebuchet MS', 'bold'),
  face('TrebuchetMS-Italic', 'Trebuchet MS', 'italic'),
  face('Trebuchet-BoldItalic', 'Trebuchet MS', 'boldItalic'),
  face('Tahoma', 'Tahoma', 'regular'),
  face('Tahoma-Bold', 'Tahoma', 'bold'),
  face('Impact', 'Impact', 'regular'),
  face('Bahnschrift', 'Bahnschrift', 'regular'),
  face('SegoePrint', 'Segoe Print', 'regular'),
  face('SegoePrint-Bold', 'Segoe Print', 'bold'),
  face('SegoeScript', 'Segoe Script', 'regular'),
  face('SegoeScript-Bold', 'Segoe Script', 'bold'),
  // Photoshop's default and other common Adobe faces (often not installed; the names still round-trip).
  face('MyriadPro-Regular', 'Myriad Pro', 'regular'),
  face('MyriadPro-Bold', 'Myriad Pro', 'bold'),
  face('MyriadPro-It', 'Myriad Pro', 'italic'),
  face('MyriadPro-BoldIt', 'Myriad Pro', 'boldItalic'),
  face('MinionPro-Regular', 'Minion Pro', 'regular'),
  face('MinionPro-Bold', 'Minion Pro', 'bold'),
  face('MinionPro-It', 'Minion Pro', 'italic'),
  face('MinionPro-BoldIt', 'Minion Pro', 'boldItalic'),
  face('Helvetica', 'Helvetica', 'regular'),
  face('Helvetica-Bold', 'Helvetica', 'bold'),
  face('Helvetica-Oblique', 'Helvetica', 'italic'),
  face('Helvetica-BoldOblique', 'Helvetica', 'boldItalic'),
])

/** Extra PostScript names that only map one way (import): weights Simple folds into 400 / 700. */
const IMPORT_ALIASES: readonly Face[] = Object.freeze([
  face('SegoeUI-Semibold', 'Segoe UI', 'bold'),
  face('SegoeUI-SemiboldItalic', 'Segoe UI', 'boldItalic'),
  face('SegoeUI-Light', 'Segoe UI', 'regular'),
  face('SegoeUI-LightItalic', 'Segoe UI', 'italic'),
  face('SegoeUI-Semilight', 'Segoe UI', 'regular'),
  face('SegoeUI-Black', 'Segoe UI', 'bold'),
  face('ArialNarrow', 'Arial Narrow', 'regular'),
  face('ArialNarrow-Bold', 'Arial Narrow', 'bold'),
  face('Arial-Black', 'Arial Black', 'regular'),
  face('Calibri-Light', 'Calibri Light', 'regular'),
])

const knownByName = new Map<string, Face>()
for (const entry of [...KNOWN_FACES, ...IMPORT_ALIASES]) knownByName.set(entry.postScriptName.toLowerCase(), entry)

// ---------------------------------------------------------------------------------------------
// Installed fonts (queryLocalFonts)
// ---------------------------------------------------------------------------------------------

interface LocalFontData {
  readonly postscriptName: string
  readonly fullName: string
  readonly family: string
  readonly style: string
}

let localByName: Map<string, Face> | null = null
let localFamilies: readonly string[] = []
let localLoad: Promise<boolean> | null = null

const BOLD_STYLE = /bold|black|heavy|semi\s*bold|demi|extra\s*bold|ultra\s*bold/i
const ITALIC_STYLE = /italic|oblique|kursiv/i

function faceFromLocal(font: LocalFontData): Face {
  return {
    postScriptName: font.postscriptName,
    family: font.family,
    weight: BOLD_STYLE.test(font.style) ? 700 : 400,
    italic: ITALIC_STYLE.test(font.style),
  }
}

/** Remembers installed fonts (tests and loadLocalFonts call this). */
export function setLocalFonts(fonts: readonly LocalFontData[]): void {
  const byName = new Map<string, Face>()
  const families = new Set<string>()
  for (const font of fonts) {
    if (!font || typeof font.postscriptName !== 'string' || typeof font.family !== 'string') continue
    byName.set(font.postscriptName.toLowerCase(), faceFromLocal(font))
    families.add(font.family)
  }
  localByName = byName
  localFamilies = Object.freeze([...families].sort((a, b) => a.localeCompare(b)))
}

/** Forgets installed fonts (tests). */
export function clearLocalFonts(): void {
  localByName = null
  localFamilies = []
  localLoad = null
}

export function localFontsLoaded(): boolean {
  return localByName !== null
}

/**
 * Reads the installed fonts once through window.queryLocalFonts(). Resolves false when the API is
 * missing, refused (Chromium may require a user gesture) or slower than `timeoutMs`; never throws.
 */
export function loadLocalFonts(timeoutMs = 2000): Promise<boolean> {
  if (localByName) return Promise.resolve(true)
  if (localLoad) return localLoad
  const query = (globalThis as { queryLocalFonts?: () => Promise<readonly LocalFontData[]> }).queryLocalFonts
  if (typeof query !== 'function') return Promise.resolve(false)
  const attempt = (async () => {
    try {
      const fonts = await query.call(globalThis)
      setLocalFonts(fonts)
      return true
    } catch {
      return false
    }
  })()
  const timeout = new Promise<boolean>((resolve) => { setTimeout(() => resolve(false), Math.max(0, timeoutMs)) })
  localLoad = Promise.race([attempt, timeout]).then((loaded) => {
    if (!loaded) localLoad = null
    return loaded
  })
  return localLoad
}

/** Families for a font menu: the curated Windows list plus installed families once loaded. */
export function availableFontFamilies(): string[] {
  return [...new Set([...FONT_FAMILIES, ...localFamilies])].sort((a, b) => a.localeCompare(b))
}

// ---------------------------------------------------------------------------------------------
// PostScript name -> CSS family
// ---------------------------------------------------------------------------------------------

function splitWords(text: string): string {
  return text
    .replace(/_/g, ' ')
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .replace(/\s+/g, ' ')
    .trim()
}

/** Heuristic for unknown names: 'TimesNewRomanPS-BoldMT' -> Times New Roman, bold. */
export function guessFontFromPostScript(postScriptName: string): ResolvedFont {
  const name = String(postScriptName ?? '').trim()
  if (!name) return { family: 'Arial', weight: 400, italic: false }
  const dash = name.indexOf('-')
  let base = dash > 0 ? name.slice(0, dash) : name
  let style = dash > 0 ? name.slice(dash + 1) : ''
  base = base.replace(/(?:PSMT|PS|MT)$/, '') || base
  style = style.replace(/(?:PSMT|PS|MT)$/, '')
  const weight: 400 | 700 = BOLD_STYLE.test(style) || /(?:^|[a-z])Bd$/.test(style) ? 700 : 400
  const italic = ITALIC_STYLE.test(style) || /(?:^|[a-z])It$/.test(style) || style === 'It'
  return { family: splitWords(base) || name, weight, italic }
}

/** A PostScript name as a CSS family + weight + italic: installed fonts, then the table, then the heuristic. */
export function resolvePostScriptFont(postScriptName: string): ResolvedFont {
  const key = String(postScriptName ?? '').trim().toLowerCase()
  const known = localByName?.get(key) ?? knownByName.get(key)
  if (known) return { family: known.family, weight: known.weight, italic: known.italic }
  const guess = guessFontFromPostScript(postScriptName)
  // Prefer the installed spelling of a guessed family ('Segoe ui' -> 'Segoe UI').
  const installed = [...localFamilies, ...FONT_FAMILIES].find((family) => family.toLowerCase() === guess.family.toLowerCase())
  return installed ? { ...guess, family: installed } : guess
}

// ---------------------------------------------------------------------------------------------
// CSS family -> PostScript name (export)
// ---------------------------------------------------------------------------------------------

function findFace(faces: Iterable<Face>, family: string, weight: 400 | 700, italic: boolean): { exact: Face | null; regular: Face | null } {
  const wanted = family.trim().toLowerCase()
  let exact: Face | null = null
  let regular: Face | null = null
  for (const entry of faces) {
    if (entry.family.toLowerCase() !== wanted) continue
    if (!exact && entry.weight === weight && entry.italic === italic) exact = entry
    if (!regular && entry.weight === 400 && !entry.italic) regular = entry
  }
  return { exact, regular }
}

/**
 * The PostScript name Photoshop should look a text style up by. An exact installed or known face wins;
 * a family known only in its regular face uses that face with faux bold / italic; anything else gets a
 * conventional guess ('Family-Bold').
 */
export function postScriptNameFor(style: Pick<TextStyle, 'fontFamily' | 'fontWeight' | 'italic'>): PsdFontChoice {
  const family = String(style.fontFamily ?? '').trim() || 'Arial'
  const weight: 400 | 700 = style.fontWeight === 700 ? 700 : 400
  const italic = Boolean(style.italic)
  for (const source of [localByName ? localByName.values() : [], KNOWN_FACES]) {
    const { exact, regular } = findFace(source, family, weight, italic)
    if (exact) return { postScriptName: exact.postScriptName, fauxBold: false, fauxItalic: false }
    if (regular) return { postScriptName: regular.postScriptName, fauxBold: weight === 700, fauxItalic: italic }
  }
  const compact = family.replace(/[^A-Za-z0-9]+/g, '') || 'Arial'
  const suffix = weight === 700 ? (italic ? '-BoldItalic' : '-Bold') : italic ? '-Italic' : ''
  return { postScriptName: `${compact}${suffix}`, fauxBold: false, fauxItalic: false }
}
