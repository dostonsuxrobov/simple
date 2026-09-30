/** Curated Windows / Office font families for the font pickers, with CSS fallbacks. */

export type FontCategory = 'sans-serif' | 'serif' | 'monospace' | 'display' | 'handwriting' | 'symbol'

export interface FontFamilyInfo {
  name: string
  category: FontCategory
  /** East Asian or Indic UI font: preview its name in the UI face if it isn't installed. */
  script?: 'cjk' | 'indic' | 'multi'
}

export const FONT_FAMILIES: readonly FontFamilyInfo[] = [
  { name: 'Aptos', category: 'sans-serif' },
  { name: 'Aptos Display', category: 'sans-serif' },
  { name: 'Aptos Narrow', category: 'sans-serif' },
  { name: 'Arial', category: 'sans-serif' },
  { name: 'Arial Black', category: 'display' },
  { name: 'Arial Narrow', category: 'sans-serif' },
  { name: 'Bahnschrift', category: 'sans-serif' },
  { name: 'Book Antiqua', category: 'serif' },
  { name: 'Bookman Old Style', category: 'serif' },
  { name: 'Calibri', category: 'sans-serif' },
  { name: 'Calibri Light', category: 'sans-serif' },
  { name: 'Cambria', category: 'serif' },
  { name: 'Cambria Math', category: 'serif' },
  { name: 'Candara', category: 'sans-serif' },
  { name: 'Cascadia Code', category: 'monospace' },
  { name: 'Cascadia Mono', category: 'monospace' },
  { name: 'Century', category: 'serif' },
  { name: 'Century Gothic', category: 'sans-serif' },
  { name: 'Comic Sans MS', category: 'handwriting' },
  { name: 'Consolas', category: 'monospace' },
  { name: 'Constantia', category: 'serif' },
  { name: 'Corbel', category: 'sans-serif' },
  { name: 'Courier New', category: 'monospace' },
  { name: 'Ebrima', category: 'sans-serif', script: 'multi' },
  { name: 'Franklin Gothic Book', category: 'sans-serif' },
  { name: 'Franklin Gothic Medium', category: 'sans-serif' },
  { name: 'Gabriola', category: 'display' },
  { name: 'Gadugi', category: 'sans-serif', script: 'multi' },
  { name: 'Garamond', category: 'serif' },
  { name: 'Georgia', category: 'serif' },
  { name: 'Gill Sans MT', category: 'sans-serif' },
  { name: 'Impact', category: 'display' },
  { name: 'Ink Free', category: 'handwriting' },
  { name: 'Leelawadee UI', category: 'sans-serif', script: 'multi' },
  { name: 'Lucida Console', category: 'monospace' },
  { name: 'Lucida Sans Unicode', category: 'sans-serif' },
  { name: 'Malgun Gothic', category: 'sans-serif', script: 'cjk' },
  { name: 'Microsoft JhengHei', category: 'sans-serif', script: 'cjk' },
  { name: 'Microsoft Sans Serif', category: 'sans-serif' },
  { name: 'Microsoft YaHei', category: 'sans-serif', script: 'cjk' },
  { name: 'MingLiU', category: 'serif', script: 'cjk' },
  { name: 'MS Gothic', category: 'monospace', script: 'cjk' },
  { name: 'MS Mincho', category: 'serif', script: 'cjk' },
  { name: 'MS PGothic', category: 'sans-serif', script: 'cjk' },
  { name: 'Nirmala UI', category: 'sans-serif', script: 'indic' },
  { name: 'Palatino Linotype', category: 'serif' },
  { name: 'Rockwell', category: 'serif' },
  { name: 'Segoe Print', category: 'handwriting' },
  { name: 'Segoe Script', category: 'handwriting' },
  { name: 'Segoe UI', category: 'sans-serif' },
  { name: 'Segoe UI Light', category: 'sans-serif' },
  { name: 'Segoe UI Semibold', category: 'sans-serif' },
  { name: 'Segoe UI Symbol', category: 'symbol' },
  { name: 'SimHei', category: 'sans-serif', script: 'cjk' },
  { name: 'SimSun', category: 'serif', script: 'cjk' },
  { name: 'Sitka Text', category: 'serif' },
  { name: 'Sylfaen', category: 'serif', script: 'multi' },
  { name: 'Symbol', category: 'symbol' },
  { name: 'Tahoma', category: 'sans-serif' },
  { name: 'Times New Roman', category: 'serif' },
  { name: 'Trebuchet MS', category: 'sans-serif' },
  { name: 'Tw Cen MT', category: 'sans-serif' },
  { name: 'Verdana', category: 'sans-serif' },
  { name: 'Webdings', category: 'symbol' },
  { name: 'Wingdings', category: 'symbol' },
  { name: 'Yu Gothic', category: 'sans-serif', script: 'cjk' },
  { name: 'Yu Mincho', category: 'serif', script: 'cjk' },
]

/** Excel's font size list. Free entry accepts 1-409 in half-point steps. */
export const FONT_SIZES: readonly number[] = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]

/** The workbook body font used when a cell has no explicit font (the app's Normal style). */
export const DEFAULT_FONT = { name: 'Aptos', size: 11 } as const

const CATEGORY_FALLBACK: Record<FontCategory, string> = {
  'sans-serif': '"Segoe UI", Arial, sans-serif',
  serif: '"Times New Roman", Georgia, serif',
  monospace: 'Consolas, "Courier New", monospace',
  display: '"Segoe UI", Arial, sans-serif',
  handwriting: '"Segoe Print", "Comic Sans MS", cursive',
  symbol: '"Segoe UI Symbol", sans-serif',
}

export function fontInfo(name: string | undefined): FontFamilyInfo | undefined {
  const key = String(name || '').trim().toLowerCase()
  return FONT_FAMILIES.find((font) => font.name.toLowerCase() === key)
}

export function isSymbolFont(name: string | undefined) {
  return fontInfo(name)?.category === 'symbol'
}

/** CSS `font-family` for a spreadsheet font name with a category-appropriate fallback. */
export function fontStack(name: string | undefined) {
  const family = String(name || '').replace(/["\\\r\n]/g, '').trim()
  const category = fontInfo(family)?.category || 'sans-serif'
  return family ? `"${family}", ${CATEGORY_FALLBACK[category]}` : `"${DEFAULT_FONT.name}", Calibri, ${CATEGORY_FALLBACK['sans-serif']}`
}

/** Parses a typed font size; Excel accepts 1-409 points rounded to the nearest half point. */
export function parseFontSize(input: string | number | undefined): number | null {
  const value = typeof input === 'number' ? input : Number(String(input ?? '').trim().replace(',', '.'))
  if (!Number.isFinite(value) || value < 1 || value > 409) return null
  return Math.round(value * 2) / 2
}

const installedCache = new Map<string, boolean>()
let probeContext: CanvasRenderingContext2D | null | undefined

/**
 * Whether a font is installed, by comparing rendered widths against generic fallbacks.
 * Returns undefined outside a browser. Results are cached per name.
 */
export function isFontInstalled(name: string): boolean | undefined {
  const family = name.replace(/["\\\r\n]/g, '').trim()
  if (!family) return undefined
  const cached = installedCache.get(family)
  if (cached !== undefined) return cached
  if (typeof document === 'undefined') return undefined
  if (probeContext === undefined) {
    try {
      probeContext = document.createElement('canvas').getContext('2d')
    } catch {
      probeContext = null
    }
  }
  const context = probeContext
  if (!context) return undefined
  const sample = 'mmmmmmmmmlli10OQ@#WwgjAaBbYy'
  let installed = false
  for (const generic of ['monospace', 'serif', 'sans-serif']) {
    context.font = `72px ${generic}`
    const baseline = context.measureText(sample).width
    context.font = `72px "${family}", ${generic}`
    if (Math.abs(context.measureText(sample).width - baseline) > 0.5) { installed = true; break }
  }
  installedCache.set(family, installed)
  return installed
}
