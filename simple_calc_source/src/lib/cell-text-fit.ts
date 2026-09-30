/**
 * Excel never truncates a number: when a formatted number does not fit its column it shows
 * "####", and a General-format number first drops decimals (or switches to scientific
 * notation) to fit. Measurement uses a shared canvas and is skipped whenever a cheap
 * character-count estimate already proves the text fits.
 */

let context: CanvasRenderingContext2D | null | undefined
const widthCache = new Map<string, number>()
const MAX_CACHE = 20_000

function measureContext() {
  if (context === undefined) {
    context = typeof document === 'undefined' ? null : document.createElement('canvas').getContext('2d')
  }
  return context
}

export function measureTextWidth(text: string, font: string): number {
  const key = `${font}\u0000${text}`
  const cached = widthCache.get(key)
  if (cached !== undefined) return cached
  const canvas = measureContext()
  let width: number
  if (canvas) {
    canvas.font = font
    width = canvas.measureText(text).width
  } else {
    const size = Number(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] || 14)
    width = text.length * size * 0.55
  }
  if (widthCache.size >= MAX_CACHE) widthCache.clear()
  widthCache.set(key, width)
  return width
}

function generalWithPrecision(value: number, digits: number): string {
  const magnitude = Math.abs(value)
  if (magnitude !== 0 && (magnitude >= 1e11 || magnitude < 1e-9)) {
    const exponent = value.toExponential(Math.max(0, digits - 1))
    const [mantissa, power] = exponent.split('e')
    const trimmed = mantissa.includes('.') ? mantissa.replace(/\.?0+$/, '') : mantissa
    const powerValue = Number(power)
    return `${trimmed}E${powerValue < 0 ? '-' : '+'}${String(Math.abs(powerValue)).padStart(2, '0')}`
  }
  const text = Number(value.toPrecision(Math.max(1, digits))).toString()
  if (!text.includes('e')) return text
  return generalWithPrecision(value, digits)
}

/**
 * The text Excel would paint for a number in a cell of `availableWidth` pixels.
 * `general` is true when the cell has no explicit number format.
 */
export function fitNumericText(text: string, value: number, general: boolean, availableWidth: number, font: string, fontPixels: number): string {
  if (!text || availableWidth <= 0) return text
  // Digits are ~0.55em in the usual UI fonts; wide glyphs rarely exceed 0.62em.
  if (text.length * fontPixels * 0.62 <= availableWidth) return text
  if (measureTextWidth(text, font) <= availableWidth) return text
  if (general && Number.isFinite(value)) {
    for (let digits = 10; digits >= 1; digits -= 1) {
      const candidate = generalWithPrecision(value, digits)
      if (candidate.length < text.length && measureTextWidth(candidate, font) <= availableWidth) return candidate
    }
  }
  const hash = measureTextWidth('#', font) || fontPixels * 0.6
  return '#'.repeat(Math.max(1, Math.floor(availableWidth / hash)))
}
