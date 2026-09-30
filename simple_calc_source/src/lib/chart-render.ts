// Framework-free chart renderer: SheetChart + resolved data -> SVG markup.
// Shared by the on-grid ChartLayer, the editor thumbnails and the print/PDF
// path, so it must not touch the DOM. Look: Excel 365 defaults (light
// gridlines, #595959 9pt axis text, 14pt title) with modern touches.
import type { ChartAnchorPoint, ChartAxis, ChartDataLabels, ChartMarkerSymbol, ChartSeries, ChartSeriesType, SheetChart, SheetData, WorkbookModel } from '../spreadsheet-types'
import { chartDisplayTitle, chartSeriesType, chartVariesColors, normalizeHex, paletteColor, type ResolvedChartData, type ResolvedChartSeries } from './charts'

export interface ChartRenderOptions {
  fontFamily?: string
  /** Number formatter, e.g. number-format.ts formatScalar. Defaults to a built-in subset of Excel formats. */
  formatNumber?: (value: number, numFmt?: string) => string
  /** Prefix for SVG ids (clip paths); must be unique per document when several charts share one page. */
  idPrefix?: string
  /** Native <title> tooltips on points (default true; skipped for very large charts). */
  tooltips?: boolean
  /** Omit all text (gallery thumbnails). */
  thumbnail?: boolean
  /** Message drawn when there is no numeric data. */
  emptyMessage?: string
}

const TEXT = '#595959'
const LABEL_TEXT = '#404040'
const GRID = '#D9D9D9'
const AXIS_LINE = '#BFBFBF'
const PT = 4 / 3
const DEFAULT_FONT = "Aptos, 'Segoe UI', Inter, Arial, sans-serif"

// ---------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------

function esc(value: unknown) {
  return String(value ?? '').replace(/[&<>"']/g, (c) => (c === '&' ? '&amp;' : c === '<' ? '&lt;' : c === '>' ? '&gt;' : c === '"' ? '&quot;' : '&#39;'))
    // Strip characters that are invalid in XML 1.0.
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
}

const r1 = (value: number) => Math.round(value * 10) / 10
const r2 = (value: number) => Math.round(value * 100) / 100

// Character width table (fraction of font size) for a Segoe UI/Aptos-like sans.
const NARROW = new Set("iljtfrI!|.,:;'`()[]{} ")
const WIDE = new Set('mwMW@%')
const CAPS = /[A-Z]/
const DIGIT = /[0-9]/
const CJK = /[ᄀ-ᇿ⺀-꓏가-힯豈-﫿︰-﹏＀-｠￠-￦]/

export function measureText(text: string, fontPx: number, bold = false) {
  let width = 0
  for (const character of text) {
    if (NARROW.has(character)) width += character === ' ' ? 0.27 : 0.28
    else if (WIDE.has(character)) width += 0.85
    else if (DIGIT.test(character)) width += 0.55
    else if (CAPS.test(character)) width += 0.64
    else if (CJK.test(character)) width += 1
    else width += 0.52
  }
  return width * fontPx * (bold ? 1.06 : 1)
}

function truncate(text: string, maxWidth: number, fontPx: number) {
  if (measureText(text, fontPx) <= maxWidth) return text
  let low = 0, high = text.length
  while (low < high) {
    const mid = Math.ceil((low + high) / 2)
    if (measureText(`${text.slice(0, mid)}…`, fontPx) <= maxWidth) low = mid
    else high = mid - 1
  }
  return low <= 0 ? '' : `${text.slice(0, low).trimEnd()}…`
}

function wrapText(text: string, maxWidth: number, fontPx: number, maxLines: number) {
  const words = text.replace(/\s+/g, ' ').trim().split(' ')
  const lines: string[] = []
  let current = ''
  for (const word of words) {
    const candidate = current ? `${current} ${word}` : word
    if (!current || measureText(candidate, fontPx) <= maxWidth) current = candidate
    else { lines.push(current); current = word }
  }
  if (current) lines.push(current)
  if (lines.length > maxLines) {
    const kept = lines.slice(0, maxLines)
    kept[maxLines - 1] = truncate(`${kept[maxLines - 1]} ${lines.slice(maxLines).join(' ')}`, maxWidth, fontPx)
    return kept
  }
  return lines.map((line) => truncate(line, maxWidth, fontPx))
}

function hashId(value: string) {
  let hash = 2166136261
  for (let index = 0; index < value.length; index += 1) hash = Math.imul(hash ^ value.charCodeAt(index), 16777619)
  return (hash >>> 0).toString(36)
}

// ---------------------------------------------------------------------------
// Number formatting (subset of Excel masks; callers may inject a full engine)
// ---------------------------------------------------------------------------

function generalNumber(value: number) {
  if (!Number.isFinite(value)) return String(value)
  if (value === 0) return '0'
  const abs = Math.abs(value)
  if (abs >= 1e11 || abs < 1e-9) return value.toExponential(4).replace(/\.?0+e/, 'E').replace('e', 'E')
  const text = Number(value.toPrecision(10)).toString()
  return text.includes('e') ? value.toExponential(4).replace(/\.?0+e/, 'E') : text
}

function excelDate(serial: number, pattern: string) {
  const ms = Math.round((serial - 25_569) * 86_400_000)
  const date = new Date(ms)
  if (Number.isNaN(date.getTime())) return generalNumber(serial)
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December']
  const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
  const y = date.getUTCFullYear(), m = date.getUTCMonth(), d = date.getUTCDate()
  const hours = date.getUTCHours(), minutes = date.getUTCMinutes(), seconds = date.getUTCSeconds()
  const twelve = /am\/pm|a\/p/i.test(pattern)
  let afterHour = false
  const output = pattern.replace(/"([^"]*)"|\\(.)|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|a\/p|\[[^\]]*\]/gi, (token, quoted, escaped) => {
    if (quoted !== undefined) return quoted
    if (escaped !== undefined) return escaped
    const lower = token.toLowerCase()
    if (lower.startsWith('[')) return ''
    if (lower === 'am/pm') return hours < 12 ? 'AM' : 'PM'
    if (lower === 'a/p') return hours < 12 ? 'A' : 'P'
    if (lower === 'hh' || lower === 'h') { afterHour = true; const value = twelve ? ((hours + 11) % 12) + 1 : hours; return lower === 'hh' ? String(value).padStart(2, '0') : String(value) }
    if (lower === 'ss' || lower === 's') return lower === 'ss' ? String(seconds).padStart(2, '0') : String(seconds)
    if ((lower === 'mm' || lower === 'm') && afterHour) { afterHour = false; return lower === 'mm' ? String(minutes).padStart(2, '0') : String(minutes) }
    switch (lower) {
      case 'yyyy': return String(y)
      case 'yy': return String(y).slice(-2)
      case 'mmmmm': return months[m][0]
      case 'mmmm': return months[m]
      case 'mmm': return months[m].slice(0, 3)
      case 'mm': return String(m + 1).padStart(2, '0')
      case 'm': return String(m + 1)
      case 'dddd': return days[date.getUTCDay()]
      case 'ddd': return days[date.getUTCDay()].slice(0, 3)
      case 'dd': return String(d).padStart(2, '0')
      default: return String(d)
    }
  })
  return output.replace(/[_*]./g, '')
}

export function defaultFormatNumber(value: number, numFmt?: string): string {
  const format = String(numFmt || '').trim()
  if (!format || /^general$/i.test(format)) return generalNumber(value)
  const sections = format.split(';')
  let section = sections[0]
  let negativeSection = false
  if (value < 0 && sections.length > 1 && sections[1].trim()) { section = sections[1]; negativeSection = true }
  else if (value === 0 && sections.length > 2 && sections[2].trim()) section = sections[2]
  // Literals: "text", \c, [$€-409], _x (spacing), *x (fill)
  let prefix = '', suffix = '', seenDigit = false, core = ''
  const cleaned = section.replace(/\[(?:Red|Black|Blue|Green|White|Yellow|Cyan|Magenta|Color\s*\d+)\]/gi, '').replace(/\[[<>=][^\]]*\]/g, '')
  const unquoted = cleaned.replace(/"[^"]*"|\\./g, '').replace(/\[[^\]]*\]/g, '')
  if (/[dmyhs]/i.test(unquoted) && !/[0#?]/.test(unquoted.replace(/[dmyhs]+/gi, ''))) return excelDate(value, cleaned).trim()
  if (/^@?$/.test(unquoted.trim()) && !/[0#?]/.test(unquoted)) return cleaned.includes('"') ? cleaned.replace(/"([^"]*)"/g, '$1').replace(/@/g, generalNumber(value)) : generalNumber(value)
  for (let index = 0; index < cleaned.length; index += 1) {
    const character = cleaned[index]
    let literal: string | null = null
    if (character === '"') { const end = cleaned.indexOf('"', index + 1); literal = cleaned.slice(index + 1, end < 0 ? undefined : end); index = end < 0 ? cleaned.length : end }
    else if (character === '\\') { literal = cleaned[index + 1] || ''; index += 1 }
    else if (character === '_' || character === '*') { index += 1; literal = character === '_' ? ' ' : '' }
    else if (character === '[') { const end = cleaned.indexOf(']', index); const inner = cleaned.slice(index + 1, end < 0 ? undefined : end); literal = inner.startsWith('$') ? inner.slice(1).split('-')[0] : ''; index = end < 0 ? cleaned.length : end }
    else if ('0#?.,Ee+-'.includes(character) && (/[0#?]/.test(character) || seenDigit || character === '.')) { core += character; if (/[0#?]/.test(character)) seenDigit = true; continue }
    else if (character === '%') { core += '%'; continue }
    else literal = character
    if (literal !== null) { if (seenDigit || core) suffix += literal; else prefix += literal }
  }
  const percent = (core.match(/%/g) || []).length
  const percentSigns = '%'.repeat(percent)
  let scaled = value * 100 ** percent
  if (/[Ee][+-]/.test(core)) {
    const decimals = (/\.([0#]+)/.exec(core)?.[1].length) ?? 2
    return (value < 0 && !negativeSection ? '-' : '') + prefix + Math.abs(scaled).toExponential(decimals).toUpperCase().replace(/E([+-])(\d)$/, 'E$10$2') + suffix
  }
  const numericPart = core.replace(/%/g, '')
  const decimals = /\.([0#?]*)/.exec(numericPart)?.[1] || ''
  const minDecimals = (decimals.match(/0/g) || []).length
  const maxDecimals = decimals.length
  // Trailing thousands separators scale by 1,000 each.
  const trailingCommas = /(,+)(?:\.|$)/.exec(numericPart.replace(/\.[0#?]*$/, ''))?.[1].length || 0
  const integerPart = numericPart.split('.')[0]
  if (trailingCommas) scaled /= 1000 ** trailingCommas
  const grouping = /[0#?],[0#?]/.test(integerPart)
  const abs = Math.abs(scaled)
  let text = abs.toFixed(maxDecimals)
  if (maxDecimals > minDecimals) text = text.replace(new RegExp(`(\\.\\d{${minDecimals}}\\d*?)0+$`), '$1').replace(/\.$/, '')
  let [whole, fraction] = text.split('.')
  const minInteger = (integerPart.match(/0/g) || []).length
  if (whole === '0' && minInteger === 0) whole = ''
  if (grouping) whole = whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',')
  const body = fraction !== undefined && fraction !== '' ? `${whole}.${fraction}` : whole
  const sign = value < 0 && !negativeSection && Number(body.replace(/[^0-9.]/g, '') || 0) !== 0 ? '-' : ''
  return `${sign}${prefix}${body}${percentSigns}${suffix}`.trim()
}

// ---------------------------------------------------------------------------
// Scales
// ---------------------------------------------------------------------------

interface Scale {
  min: number
  max: number
  step: number
  ticks: number[]
  log?: number
  /** Map a value to [0,1] along the axis. */
  norm(value: number): number
}

function niceStep(raw: number) {
  if (!(raw > 0) || !Number.isFinite(raw)) return 1
  const exponent = Math.floor(Math.log10(raw))
  const base = 10 ** exponent
  const fraction = raw / base
  const nice = fraction <= 1.0000001 ? 1 : fraction <= 2.0000001 ? 2 : fraction <= 5.0000001 ? 5 : 10
  return nice * base
}

function nextNiceStep(step: number) {
  const exponent = Math.floor(Math.log10(step))
  const base = 10 ** exponent
  const fraction = Math.round(step / base)
  return (fraction < 2 ? 2 : fraction < 5 ? 5 : 10) * base
}

function stepDecimals(step: number) {
  if (step >= 1) return 0
  return Math.min(12, Math.max(0, Math.ceil(-Math.log10(step) - 1e-9)))
}

function linearScale(dataMin: number, dataMax: number, maxIntervals: number, axis: ChartAxis | undefined, options: { percent?: boolean; zeroRule?: boolean } = {}): Scale {
  if (axis?.logBase && axis.logBase > 1) return logScale(dataMin, dataMax, axis)
  let min = Number.isFinite(dataMin) ? dataMin : 0
  let max = Number.isFinite(dataMax) ? dataMax : 1
  if (options.percent) {
    min = min < 0 ? -1 : 0
    max = max > 0 ? 1 : 0
    if (min === max) max = 1
  }
  if (min > max) [min, max] = [max, min]
  if (min === max) {
    if (min === 0) max = 1
    else if (min > 0) min = 0
    else max = 0
  }
  const zeroRule = options.zeroRule !== false
  // Excel's automatic minimum: start at zero unless the data sits in a narrow high band.
  if (zeroRule && !options.percent) {
    if (min >= 0 && !(min > max * (5 / 6))) min = 0
    else if (max <= 0 && !(max < min * (5 / 6))) max = 0
  }
  const intervals = Math.max(1, maxIntervals)
  const fixedMin = Number.isFinite(axis?.min) ? axis!.min! : undefined
  const fixedMax = Number.isFinite(axis?.max) ? axis!.max! : undefined
  const span = (fixedMax ?? max) - (fixedMin ?? min) || Math.abs(max) || 1
  let step = axis?.majorUnit && axis.majorUnit > 0 ? axis.majorUnit : options.percent ? (intervals >= 10 ? 0.1 : intervals >= 5 ? 0.2 : 0.25) : niceStep(span / intervals)
  let axisMin = 0, axisMax = 1
  for (let guard = 0; guard < 40; guard += 1) {
    axisMin = fixedMin ?? (min === 0 ? 0 : Math.floor((min < 0 ? min - 0.05 * ((max > 0 ? max : 0) - min) : min - (max - min) / 2) / step + 1e-9) * step)
    if (fixedMin === undefined && min > 0 && axisMin < 0) axisMin = 0
    if (fixedMin === undefined && min < 0 && max <= 0 && axisMin > min) axisMin = Math.floor(min / step) * step
    axisMax = fixedMax ?? (max === 0 ? 0 : Math.ceil((max + (max > 0 ? 0.05 * (max - axisMin) : 0)) / step - 1e-9) * step)
    if (options.percent) { axisMin = fixedMin ?? min; axisMax = fixedMax ?? max }
    if (axisMax <= axisMin) axisMax = axisMin + step
    const count = Math.round((axisMax - axisMin) / step)
    if (count <= intervals || axis?.majorUnit || count <= 2) break
    step = nextNiceStep(step)
  }
  const decimals = stepDecimals(step)
  const ticks: number[] = []
  const first = Math.ceil(axisMin / step - 1e-9) * step
  for (let value = first, index = 0; value <= axisMax + step * 1e-6 && index < 200; index += 1, value = first + index * step) {
    ticks.push(Number(value.toFixed(decimals + 2)))
  }
  const range = axisMax - axisMin || 1
  const reverse = axis?.reverse === true
  return { min: axisMin, max: axisMax, step, ticks, norm: (value) => { const t = (value - axisMin) / range; return reverse ? 1 - t : t } }
}

function logScale(dataMin: number, dataMax: number, axis: ChartAxis): Scale {
  const base = axis.logBase || 10
  const lo = Math.max(1e-300, Number.isFinite(axis.min) && axis.min! > 0 ? axis.min! : dataMin > 0 ? base ** Math.floor(Math.log(dataMin) / Math.log(base)) : 1)
  const hiRaw = Number.isFinite(axis.max) && axis.max! > 0 ? axis.max! : dataMax > 0 ? base ** Math.ceil(Math.log(dataMax) / Math.log(base) + 1e-9) : base
  const hi = hiRaw <= lo ? lo * base : hiRaw
  const ticks: number[] = []
  for (let value = lo; value <= hi * 1.000001 && ticks.length < 60; value *= base) ticks.push(value)
  const logLo = Math.log(lo), span = Math.log(hi) - logLo || 1
  const reverse = axis.reverse === true
  return { min: lo, max: hi, step: base, ticks, log: base, norm: (value) => { const t = value > 0 ? (Math.log(value) - logLo) / span : 0; return reverse ? 1 - t : t } }
}

// ---------------------------------------------------------------------------
// Chart model
// ---------------------------------------------------------------------------

interface SeriesPlan {
  index: number
  source: ChartSeries | undefined
  data: ResolvedChartSeries
  kind: ChartSeriesType | 'pie' | 'doughnut' | 'radar'
  secondary: boolean
}

interface Ctx {
  chart: SheetChart
  data: ResolvedChartData
  width: number
  height: number
  font: string
  axisPx: number
  titlePx: number
  textColor: string
  format: (value: number, numFmt?: string) => string
  out: string[]
  id: string
  tooltips: boolean
  thumbnail: boolean
}

function text(ctx: Ctx, x: number, y: number, value: string, options: { size?: number; anchor?: 'start' | 'middle' | 'end'; color?: string; weight?: number; rotate?: number; baseline?: 'middle' | 'hanging' | 'alphabetic'; cls?: string } = {}) {
  if (ctx.thumbnail || !value) return
  const size = options.size ?? ctx.axisPx
  const attrs = [`x="${r1(x)}"`, `y="${r1(y)}"`, `font-size="${r1(size)}"`]
  if (options.anchor && options.anchor !== 'start') attrs.push(`text-anchor="${options.anchor}"`)
  if (options.color && options.color !== ctx.textColor) attrs.push(`fill="${options.color}"`)
  if (options.weight) attrs.push(`font-weight="${options.weight}"`)
  if (options.baseline === 'middle') attrs.push('dominant-baseline="central"')
  if (options.rotate) attrs.push(`transform="rotate(${r1(options.rotate)} ${r1(x)} ${r1(y)})"`)
  if (options.cls) attrs.push(`class="${options.cls}"`)
  ctx.out.push(`<text ${attrs.join(' ')}>${esc(value)}</text>`)
}

function tooltip(ctx: Ctx, label: string) {
  return ctx.tooltips ? `<title>${esc(label)}</title>` : ''
}

function markerPath(symbol: ChartMarkerSymbol | undefined, x: number, y: number, size: number) {
  const s = size / 2
  switch (symbol) {
    case 'square': return `M${r1(x - s)} ${r1(y - s)}h${r1(size)}v${r1(size)}h${r1(-size)}Z`
    case 'diamond': return `M${r1(x)} ${r1(y - s * 1.2)}L${r1(x + s * 1.2)} ${r1(y)}L${r1(x)} ${r1(y + s * 1.2)}L${r1(x - s * 1.2)} ${r1(y)}Z`
    case 'triangle': return `M${r1(x)} ${r1(y - s * 1.15)}L${r1(x + s * 1.1)} ${r1(y + s * 0.85)}L${r1(x - s * 1.1)} ${r1(y + s * 0.85)}Z`
    case 'x': return `M${r1(x - s)} ${r1(y - s)}L${r1(x + s)} ${r1(y + s)}M${r1(x + s)} ${r1(y - s)}L${r1(x - s)} ${r1(y + s)}`
    case 'plus': return `M${r1(x - s)} ${r1(y)}H${r1(x + s)}M${r1(x)} ${r1(y - s)}V${r1(y + s)}`
    case 'star': return `M${r1(x - s)} ${r1(y)}H${r1(x + s)}M${r1(x)} ${r1(y - s)}V${r1(y + s)}M${r1(x - s * 0.7)} ${r1(y - s * 0.7)}L${r1(x + s * 0.7)} ${r1(y + s * 0.7)}M${r1(x + s * 0.7)} ${r1(y - s * 0.7)}L${r1(x - s * 0.7)} ${r1(y + s * 0.7)}`
    case 'dash': return `M${r1(x - s)} ${r1(y - 1)}h${r1(size)}v2h${r1(-size)}Z`
    case 'dot': return `M${r1(x - 1.5)} ${r1(y)}a1.5 1.5 0 1 0 3 0a1.5 1.5 0 1 0 -3 0Z`
    default: return `M${r1(x - s)} ${r1(y)}a${r1(s)} ${r1(s)} 0 1 0 ${r1(size)} 0a${r1(s)} ${r1(s)} 0 1 0 ${r1(-size)} 0Z`
  }
}

function strokedMarker(symbol: ChartMarkerSymbol | undefined) {
  return symbol === 'x' || symbol === 'plus' || symbol === 'star'
}

function effectiveMarker(chart: SheetChart, series: ChartSeries | undefined, kind: SeriesPlan['kind']): ChartMarkerSymbol {
  const marker = series?.marker
  if (marker) return marker === 'auto' ? 'circle' : marker
  if (kind === 'scatter') return 'circle'
  if (kind === 'line' || kind === 'radar') return chart.type === 'combo' ? 'none' : 'none'
  return 'none'
}

function labelsFor(series: ChartSeries | undefined): ChartDataLabels | null {
  const labels = series?.dataLabels
  if (!labels) return null
  return labels.showValue || labels.showCategory || labels.showSeriesName || labels.showPercent ? labels : null
}

function labelText(ctx: Ctx, labels: ChartDataLabels, value: number | null, category: string, seriesName: string, percent?: number, numFmt?: string) {
  const parts: string[] = []
  if (labels.showSeriesName) parts.push(seriesName)
  if (labels.showCategory) parts.push(category)
  if (labels.showValue && value !== null) parts.push(ctx.format(value, labels.numFmt || numFmt))
  if (labels.showPercent && percent !== undefined) parts.push(ctx.format(percent, '0%'))
  return parts.join(', ')
}

// ---------------------------------------------------------------------------
// Frame: background, title, legend
// ---------------------------------------------------------------------------

interface Box { x: number; y: number; w: number; h: number }

interface LegendEntry { label: string; color: string; line: boolean; marker: ChartMarkerSymbol }

function legendEntries(ctx: Ctx, plans: SeriesPlan[]): LegendEntry[] {
  const { chart, data } = ctx
  if (chartVariesColors(chart) && data.series[0]) {
    const first = data.series[0]
    return data.categories.map((label, index) => ({ label: label || String(index + 1), color: first.pointColors?.[index] || first.color, line: false, marker: 'none' as ChartMarkerSymbol }))
      .filter((_, index) => chart.type !== 'pie' && chart.type !== 'doughnut' ? true : first.values[index] !== null)
  }
  return plans.map((plan) => {
    const line = plan.kind === 'line' || plan.kind === 'scatter' || plan.kind === 'radar'
    const showLine = plan.kind === 'scatter' ? plan.source?.showLine !== false && plan.source?.showLine !== undefined : true
    return { label: plan.data.name, color: plan.data.color, line: line && showLine, marker: line ? effectiveMarker(chart, plan.source, plan.kind) : 'none' }
  })
}

function drawLegend(ctx: Ctx, entries: LegendEntry[], position: string, box: Box): Box {
  if (!entries.length || position === 'none') return box
  const size = ctx.axisPx
  const rowHeight = size * 1.45
  const swatch = (entry: LegendEntry) => (entry.line ? 18 : 8)
  const gap = 5, spacing = 14
  if (position === 'right' || position === 'left') {
    const maxWidth = Math.max(40, Math.min(box.w * 0.38, 220))
    const items = entries.slice(0, Math.max(1, Math.floor((box.h - 8) / rowHeight)))
    const widths = items.map((entry) => swatch(entry) + gap + Math.min(maxWidth - swatch(entry) - gap, measureText(entry.label, size)))
    const legendWidth = Math.min(maxWidth, Math.max(...widths)) + 6
    const legendHeight = items.length * rowHeight
    const x = position === 'right' ? box.x + box.w - legendWidth : box.x
    let y = box.y + Math.max(0, (box.h - legendHeight) / 2)
    ctx.out.push('<g class="chart-legend">')
    for (const entry of items) {
      drawLegendItem(ctx, entry, x, y + rowHeight / 2, legendWidth - swatch(entry) - gap - 2)
      y += rowHeight
    }
    ctx.out.push('</g>')
    return position === 'right' ? { ...box, w: box.w - legendWidth - 10 } : { ...box, x: box.x + legendWidth + 10, w: box.w - legendWidth - 10 }
  }
  // Top/bottom: wrap centred rows, at most three.
  const maxItemWidth = box.w * 0.9
  const rows: Array<{ entries: LegendEntry[]; widths: number[]; width: number }> = []
  let current = { entries: [] as LegendEntry[], widths: [] as number[], width: 0 }
  for (const entry of entries) {
    const width = Math.min(maxItemWidth, swatch(entry) + gap + measureText(entry.label, size))
    const needed = current.entries.length ? width + spacing : width
    if (current.entries.length && current.width + needed > box.w) { rows.push(current); current = { entries: [], widths: [], width: 0 } }
    current.entries.push(entry)
    current.widths.push(width)
    current.width += current.entries.length > 1 ? width + spacing : width
  }
  if (current.entries.length) rows.push(current)
  const maxRows = Math.max(1, Math.min(3, Math.floor(box.h / (rowHeight * 4))))
  const shown = rows.slice(0, maxRows)
  const legendHeight = shown.length * rowHeight
  let y = position === 'bottom' ? box.y + box.h - legendHeight : box.y
  ctx.out.push('<g class="chart-legend">')
  for (const row of shown) {
    let x = box.x + (box.w - row.width) / 2
    row.entries.forEach((entry, index) => {
      drawLegendItem(ctx, entry, x, y + rowHeight / 2, row.widths[index] - swatch(entry) - gap)
      x += row.widths[index] + spacing
    })
    y += rowHeight
  }
  ctx.out.push('</g>')
  return position === 'bottom' ? { ...box, h: box.h - legendHeight - 6 } : { ...box, y: box.y + legendHeight + 6, h: box.h - legendHeight - 6 }
}

function drawLegendItem(ctx: Ctx, entry: LegendEntry, x: number, cy: number, textWidth: number) {
  const size = ctx.axisPx
  if (entry.line) {
    ctx.out.push(`<path d="M${r1(x)} ${r1(cy)}h18" stroke="${entry.color}" stroke-width="2.5" stroke-linecap="round"/>`)
    if (entry.marker !== 'none') ctx.out.push(markerSvg(entry.marker, x + 9, cy, 6, entry.color))
  } else if (entry.marker !== 'none') {
    ctx.out.push(markerSvg(entry.marker, x + 4, cy, 7, entry.color))
  } else {
    ctx.out.push(`<rect x="${r1(x)}" y="${r1(cy - 4)}" width="8" height="8" rx="1" fill="${entry.color}"/>`)
  }
  const offset = entry.line ? 23 : 13
  text(ctx, x + offset, cy, truncate(entry.label, Math.max(10, textWidth), size), { baseline: 'middle' })
}

function markerSvg(symbol: ChartMarkerSymbol, x: number, y: number, size: number, color: string) {
  if (symbol === 'none') return ''
  return strokedMarker(symbol)
    ? `<path d="${markerPath(symbol, x, y, size)}" fill="none" stroke="${color}" stroke-width="1.5"/>`
    : `<path d="${markerPath(symbol, x, y, size)}" fill="${color}" stroke="${color}" stroke-width="0.75"/>`
}

// ---------------------------------------------------------------------------
// Cartesian charts
// ---------------------------------------------------------------------------

interface StackInfo { pos: number[]; neg: number[]; totals: number[] }

function stackTotals(plans: SeriesPlan[], count: number): StackInfo {
  const pos = new Array(count).fill(0), neg = new Array(count).fill(0)
  for (const plan of plans) for (let i = 0; i < count; i += 1) { const v = plan.data.values[i]; if (v === null || v === undefined) continue; if (v >= 0) pos[i] += v; else neg[i] += v }
  return { pos, neg, totals: pos.map((value, index) => value - neg[index]) }
}

function valueDomain(plans: SeriesPlan[], count: number, grouping: string, kindIsStackable: boolean) {
  let min = Infinity, max = -Infinity
  if (kindIsStackable && grouping === 'percentStacked') {
    const { pos, neg, totals } = stackTotals(plans, count)
    for (let i = 0; i < count; i += 1) {
      if (!totals[i]) continue
      max = Math.max(max, pos[i] / totals[i]); min = Math.min(min, neg[i] / totals[i])
    }
    return { min: Number.isFinite(min) ? min : 0, max: Number.isFinite(max) ? max : 1 }
  }
  if (kindIsStackable && grouping === 'stacked') {
    const { pos, neg } = stackTotals(plans, count)
    for (let i = 0; i < count; i += 1) { max = Math.max(max, pos[i]); min = Math.min(min, neg[i]) }
  } else {
    for (const plan of plans) for (const v of plan.data.values) if (v !== null && v !== undefined && Number.isFinite(v)) { if (v < min) min = v; if (v > max) max = v }
  }
  return { min: Number.isFinite(min) ? min : 0, max: Number.isFinite(max) ? max : 1 }
}

function axisFormat(axis: ChartAxis | undefined, plans: SeriesPlan[], percent: boolean) {
  if (percent) return '0%'
  if (axis?.numFmt && !/^general$/i.test(axis.numFmt)) return axis.numFmt
  return plans.find((plan) => plan.data.numFmt)?.data.numFmt
}

function smoothPath(points: Array<[number, number]>) {
  if (points.length < 3) return points.map((p, i) => `${i ? 'L' : 'M'}${r1(p[0])} ${r1(p[1])}`).join('')
  let d = `M${r1(points[0][0])} ${r1(points[0][1])}`
  for (let i = 0; i < points.length - 1; i += 1) {
    const p0 = points[i - 1] || points[i], p1 = points[i], p2 = points[i + 1], p3 = points[i + 2] || p2
    const c1x = p1[0] + (p2[0] - p0[0]) / 6, c1y = p1[1] + (p2[1] - p0[1]) / 6
    const c2x = p2[0] - (p3[0] - p1[0]) / 6, c2y = p2[1] - (p3[1] - p1[1]) / 6
    d += `C${r1(c1x)} ${r1(c1y)} ${r1(c2x)} ${r1(c2y)} ${r1(p2[0])} ${r1(p2[1])}`
  }
  return d
}

function linePath(points: Array<[number, number]>, smooth: boolean) {
  return smooth ? smoothPath(points) : points.map((p, i) => `${i ? 'L' : 'M'}${r1(p[0])} ${r1(p[1])}`).join('')
}

function renderCartesian(ctx: Ctx, plans: SeriesPlan[], box: Box) {
  const { chart, data } = ctx
  const horizontal = chart.type === 'bar'
  const scatter = chart.type === 'scatter'
  const count = scatter ? 0 : Math.max(data.categories.length, ...plans.map((plan) => plan.data.values.length), 0)
  const grouping = chart.grouping || 'clustered'
  const primary = plans.filter((plan) => !plan.secondary)
  const secondary = plans.filter((plan) => plan.secondary)
  const stackableKinds = (list: SeriesPlan[]) => list.filter((plan) => plan.kind === 'column' || plan.kind === 'bar' || plan.kind === 'area' || (plan.kind === 'line' && chart.type !== 'combo'))
  const groupingFor = (plan: SeriesPlan) => (chart.type === 'combo' && plan.kind !== 'column' && plan.kind !== 'bar' ? 'clustered' : grouping)
  const percent = grouping === 'percentStacked' && chart.type !== 'combo'
  const domainOf = (list: SeriesPlan[]) => {
    if (!list.length) return { min: 0, max: 1 }
    const stackable = stackableKinds(list)
    const others = list.filter((plan) => !stackable.includes(plan))
    const a = stackable.length ? valueDomain(stackable, count, chart.type === 'combo' ? grouping : grouping, true) : { min: Infinity, max: -Infinity }
    const b = others.length ? valueDomain(others, count, 'clustered', false) : { min: Infinity, max: -Infinity }
    return { min: Math.min(a.min, b.min), max: Math.max(a.max, b.max) }
  }
  const axes = chart.axes || {}
  const xAxis = axes.x || {}
  const yAxis = axes.y || {}
  const y2Axis = axes.y2 || {}
  const size = ctx.axisPx
  const xVisible = xAxis.visible !== false
  const yVisible = yAxis.visible !== false
  const y2Visible = secondary.length > 0 && y2Axis.visible !== false
  const yDomain = domainOf(primary.length ? primary : plans)
  const y2Domain = domainOf(secondary)
  const yFmt = axisFormat(yAxis, primary, percent)
  const y2Fmt = axisFormat(y2Axis, secondary, false)

  // Scatter X domain
  let xDomain = { min: 0, max: 1 }
  if (scatter) {
    let min = Infinity, max = -Infinity
    for (const plan of plans) for (const v of plan.data.x || []) if (v !== null && Number.isFinite(v)) { if (v < min) min = v; if (v > max) max = v }
    xDomain = { min: Number.isFinite(min) ? min : 0, max: Number.isFinite(max) ? max : 1 }
  }
  const xFmt = scatter ? (xAxis.numFmt && !/^general$/i.test(xAxis.numFmt) ? xAxis.numFmt : data.categoryNumFmt) : undefined

  // Axis title space
  const titleSize = size * 1.1
  const xTitle = xAxis.title && !ctx.thumbnail ? xAxis.title : ''
  const yTitle = yAxis.title && !ctx.thumbnail ? yAxis.title : ''
  const y2Title = secondary.length && y2Axis.title && !ctx.thumbnail ? y2Axis.title : ''
  // For horizontal bars the category axis is vertical (left) and the value axis horizontal (bottom).
  const leftTitle = horizontal ? xTitle : yTitle
  const bottomTitle = horizontal ? yTitle : xTitle
  let plot: Box = { ...box }
  if (leftTitle) { plot.x += titleSize * 1.4; plot.w -= titleSize * 1.4 }
  if (y2Title) plot.w -= titleSize * 1.4
  if (bottomTitle) plot.h -= titleSize * 1.5
  if (ctx.thumbnail) {
    const vScale = linearScale(yDomain.min, yDomain.max, 4, yAxis, { percent })
    return drawCartesianBody(ctx, plans, plot, { horizontal, scatter, count, grouping, groupingFor, yScale: vScale, y2Scale: secondary.length ? linearScale(y2Domain.min, y2Domain.max, 4, y2Axis) : vScale, xScale: scatter ? linearScale(xDomain.min, xDomain.max, 4, xAxis) : undefined, yAxis, xAxis })
  }

  const tickFont = size
  const labelOf = (value: number, fmt?: string) => ctx.format(value, fmt)
  // Pass 1: value-axis label width with a provisional length.
  const valueLength = (horizontal ? plot.w : plot.h) - 24
  const maxIntervalsFor = (length: number, labelWidth: number) => (horizontal
    ? Math.max(2, Math.min(10, Math.floor(length / Math.max(labelWidth + 14, tickFont * 3))))
    : Math.max(2, Math.min(10, Math.floor(length / (tickFont * 2.6)))))
  let yScale = linearScale(yDomain.min, yDomain.max, maxIntervalsFor(valueLength, tickFont * 3), yAxis, { percent })
  let y2Scale = secondary.length ? linearScale(y2Domain.min, y2Domain.max, maxIntervalsFor(valueLength, tickFont * 3), y2Axis) : yScale
  let xScale = scatter ? linearScale(xDomain.min, xDomain.max, Math.max(2, Math.min(10, Math.floor(plot.w / (tickFont * 4.5)))), xAxis) : undefined
  const widest = (scale: Scale, fmt?: string) => Math.max(0, ...scale.ticks.map((tick) => measureText(labelOf(tick, fmt), tickFont)))
  if (horizontal) {
    const labelWidth = widest(yScale, yFmt)
    yScale = linearScale(yDomain.min, yDomain.max, maxIntervalsFor(plot.w - 40, labelWidth), yAxis, { percent })
  }

  // Category labels (or scatter X labels)
  const categories = data.categories.slice(0, count)
  const catLabels = categories.map((label) => truncate(label, 180, tickFont))
  let leftLabelWidth = 0
  if (horizontal) leftLabelWidth = xVisible ? Math.min(plot.w * 0.4, Math.max(0, ...catLabels.map((label) => measureText(label, tickFont)))) + 8 : 0
  else leftLabelWidth = yVisible ? widest(yScale, yFmt) + 8 : 0
  const rightLabelWidth = y2Visible ? widest(y2Scale, y2Fmt) + 8 : 0
  plot = { ...plot, x: plot.x + leftLabelWidth, w: Math.max(10, plot.w - leftLabelWidth - rightLabelWidth - (rightLabelWidth ? 0 : 4)) }
  // Edge-to-edge categories (area charts) centre their first/last labels on the plot edges.
  if (!horizontal && !scatter && xVisible && count > 1 && plans.every((plan) => plan.kind === 'area')) {
    const firstHalf = Math.min(60, measureText(catLabels[0] || '', tickFont) / 2)
    const lastHalf = Math.min(60, measureText(catLabels[count - 1] || '', tickFont) / 2)
    const shiftLeft = Math.max(0, firstHalf - leftLabelWidth + 2)
    const trimRight = Math.max(0, lastHalf - (rightLabelWidth ? rightLabelWidth : 4) + 2)
    plot = { ...plot, x: plot.x + shiftLeft, w: Math.max(10, plot.w - shiftLeft - trimRight) }
  }
  // Bottom labels: category labels may need rotation or skipping.
  let bottomHeight = tickFont * 1.5
  let rotate = 0, skip = 1
  if (!horizontal && !scatter && xVisible) {
    const midCat = plans.every((plan) => plan.kind === 'area')
    const band = count ? plot.w / Math.max(1, midCat ? count - 1 || 1 : count) : plot.w
    const widestCat = Math.max(0, ...catLabels.map((label) => measureText(label, tickFont)))
    if (Number.isFinite(xAxis.labelRotation) && xAxis.labelRotation) rotate = xAxis.labelRotation!
    else if (widestCat > band - 3) rotate = -45
    if (rotate) {
      const radians = Math.abs(rotate) * Math.PI / 180
      const perLabel = tickFont * 1.25 / Math.max(0.2, Math.sin(radians))
      skip = Math.max(1, Math.ceil(perLabel / Math.max(1, band)))
      if (Math.abs(rotate) >= 89) skip = Math.max(1, Math.ceil(tickFont * 1.25 / Math.max(1, band)))
      const longest = Math.min(widestCat, box.h * 0.35 / Math.max(0.2, Math.sin(radians)))
      bottomHeight = longest * Math.sin(radians) + tickFont * Math.cos(radians) + 6
    }
  } else if (!xVisible && !horizontal) bottomHeight = 4
  if (horizontal && !yVisible) bottomHeight = 4
  plot = { ...plot, h: Math.max(10, plot.h - bottomHeight) }
  // Pass 2: final value scales with the final plot length.
  if (!horizontal) {
    yScale = linearScale(yDomain.min, yDomain.max, maxIntervalsFor(plot.h, 0), yAxis, { percent })
    if (secondary.length) y2Scale = linearScale(y2Domain.min, y2Domain.max, maxIntervalsFor(plot.h, 0), y2Axis)
    else y2Scale = yScale
  }
  if (scatter) xScale = linearScale(xDomain.min, xDomain.max, Math.max(2, Math.min(10, Math.floor(plot.w / Math.max(tickFont * 3.5, widest(xScale!, xFmt) + 14)))), xAxis)

  // Gridlines
  ctx.out.push('<g class="chart-grid">')
  const valueGrid = yAxis.gridlines !== false
  if (valueGrid) {
    for (const tick of yScale.ticks) {
      const t = yScale.norm(tick)
      if (t < -0.0001 || t > 1.0001) continue
      if (horizontal) { const x = plot.x + t * plot.w; ctx.out.push(`<path d="M${r1(x)} ${r1(plot.y)}V${r1(plot.y + plot.h)}" stroke="${GRID}" stroke-width="1"/>`) }
      else { const y = plot.y + plot.h - t * plot.h; ctx.out.push(`<path d="M${r1(plot.x)} ${r1(y)}H${r1(plot.x + plot.w)}" stroke="${GRID}" stroke-width="1"/>`) }
    }
  }
  if (xAxis.gridlines) {
    if (scatter && xScale) for (const tick of xScale.ticks) { const x = plot.x + xScale.norm(tick) * plot.w; ctx.out.push(`<path d="M${r1(x)} ${r1(plot.y)}V${r1(plot.y + plot.h)}" stroke="${GRID}" stroke-width="1"/>`) }
    else if (count) for (let i = 0; i <= count; i += 1) {
      if (horizontal) { const y = plot.y + (i / count) * plot.h; ctx.out.push(`<path d="M${r1(plot.x)} ${r1(y)}H${r1(plot.x + plot.w)}" stroke="${GRID}" stroke-width="1"/>`) }
      else { const x = plot.x + (i / count) * plot.w; ctx.out.push(`<path d="M${r1(x)} ${r1(plot.y)}V${r1(plot.y + plot.h)}" stroke="${GRID}" stroke-width="1"/>`) }
    }
  }
  ctx.out.push('</g>')

  const body = drawCartesianBody(ctx, plans, plot, { horizontal, scatter, count, grouping, groupingFor, yScale, y2Scale, xScale, yAxis, xAxis })

  // Axes and labels
  ctx.out.push('<g class="chart-axes">')
  const zeroT = yScale.log ? 0 : Math.max(0, Math.min(1, yScale.norm(Math.max(yScale.min, Math.min(yScale.max, 0)))))
  if (horizontal) {
    const x0 = plot.x + zeroT * plot.w
    if (xVisible) ctx.out.push(`<path d="M${r1(x0)} ${r1(plot.y)}V${r1(plot.y + plot.h)}" stroke="${AXIS_LINE}" stroke-width="1"/>`)
    if (yVisible) for (const tick of yScale.ticks) {
      const t = yScale.norm(tick)
      if (t < -0.0001 || t > 1.0001) continue
      text(ctx, plot.x + t * plot.w, plot.y + plot.h + tickFont * 1.15, labelOf(tick, yFmt), { anchor: 'middle' })
    }
    if (xVisible && count) {
      const band = plot.h / count
      const every = Math.max(1, Math.ceil(tickFont * 1.2 / Math.max(1, band)))
      const reverse = xAxis.reverse === true
      for (let i = 0; i < count; i += every) {
        const slot = reverse ? i : count - 1 - i // Excel draws the first category at the bottom
        text(ctx, plot.x - 6, plot.y + (slot + 0.5) * band, truncate(catLabels[i] || '', leftLabelWidth - 8, tickFont), { anchor: 'end', baseline: 'middle' })
      }
    }
  } else {
    const y0 = plot.y + plot.h - zeroT * plot.h
    if (xVisible) ctx.out.push(`<path d="M${r1(plot.x)} ${r1(scatter ? plot.y + plot.h : y0)}H${r1(plot.x + plot.w)}" stroke="${AXIS_LINE}" stroke-width="1"/>`)
    if (yVisible) for (const tick of yScale.ticks) {
      const t = yScale.norm(tick)
      if (t < -0.0001 || t > 1.0001) continue
      text(ctx, plot.x - 6, plot.y + plot.h - t * plot.h, labelOf(tick, yFmt), { anchor: 'end', baseline: 'middle' })
    }
    if (y2Visible) for (const tick of y2Scale.ticks) {
      const t = y2Scale.norm(tick)
      if (t < -0.0001 || t > 1.0001) continue
      text(ctx, plot.x + plot.w + 6, plot.y + plot.h - t * plot.h, labelOf(tick, y2Fmt), { baseline: 'middle' })
    }
    if (xVisible && scatter && xScale) {
      for (const tick of xScale.ticks) text(ctx, plot.x + xScale.norm(tick) * plot.w, plot.y + plot.h + tickFont * 1.15, labelOf(tick, xFmt), { anchor: 'middle' })
    } else if (xVisible && count) {
      const positions = body.categoryCenters
      for (let i = 0; i < count; i += skip) {
        const x = positions[xAxis.reverse ? count - 1 - i : i]
        const label = catLabels[i] || ''
        if (rotate) text(ctx, x + tickFont * 0.3, plot.y + plot.h + 6, truncate(label, (bottomHeight - 6) / Math.max(0.2, Math.sin(Math.abs(rotate) * Math.PI / 180)), tickFont), { anchor: 'end', rotate, baseline: 'hanging' })
        else text(ctx, x, plot.y + plot.h + tickFont * 1.15, truncate(label, Math.max(8, (plot.w / Math.max(1, count)) * skip - 2), tickFont), { anchor: 'middle' })
      }
    }
  }
  // Axis titles
  if (leftTitle) text(ctx, box.x + titleSize * 0.75, plot.y + plot.h / 2, truncate(leftTitle, plot.h, titleSize), { anchor: 'middle', rotate: -90, baseline: 'middle', size: titleSize })
  if (y2Title) text(ctx, box.x + box.w - titleSize * 0.6, plot.y + plot.h / 2, truncate(y2Title, plot.h, titleSize), { anchor: 'middle', rotate: 90, baseline: 'middle', size: titleSize })
  if (bottomTitle) text(ctx, plot.x + plot.w / 2, box.y + box.h - titleSize * 0.35, truncate(bottomTitle, plot.w, titleSize), { anchor: 'middle', size: titleSize })
  ctx.out.push('</g>')
  ctx.out.push(body.labels.join(''))
}

interface BodyOptions {
  horizontal: boolean
  scatter: boolean
  count: number
  grouping: string
  groupingFor: (plan: SeriesPlan) => string
  yScale: Scale
  y2Scale: Scale
  xScale?: Scale
  yAxis: ChartAxis
  xAxis: ChartAxis
}

function drawCartesianBody(ctx: Ctx, plans: SeriesPlan[], plot: Box, options: BodyOptions) {
  const { chart, data } = ctx
  const { horizontal, scatter, count, yScale, y2Scale, xScale } = options
  const labels: string[] = []
  const reverseCats = options.xAxis.reverse === true
  const midCat = plans.length > 0 && plans.every((plan) => plan.kind === 'area')
  const categoryCenters: number[] = []
  for (let i = 0; i < count; i += 1) {
    const slot = reverseCats ? count - 1 - i : i
    if (horizontal) categoryCenters.push(plot.y + (count - 1 - slot + 0.5) * (plot.h / Math.max(1, count)))
    else categoryCenters.push(midCat && count > 1 ? plot.x + slot * (plot.w / (count - 1)) : plot.x + (slot + 0.5) * (plot.w / Math.max(1, count)))
  }
  const valuePos = (scale: Scale, value: number) => {
    const t = scale.norm(scale.log && value <= 0 ? scale.min : value)
    return horizontal ? plot.x + t * plot.w : plot.y + plot.h - t * plot.h
  }
  const baseValue = (scale: Scale) => (scale.log ? scale.min : Math.max(scale.min, Math.min(scale.max, 0)))
  const band = (horizontal ? plot.h : plot.w) / Math.max(1, count)
  ctx.out.push(`<defs><clipPath id="${ctx.id}-plot"><rect x="${r1(plot.x - 1)}" y="${r1(plot.y - 1)}" width="${r1(plot.w + 2)}" height="${r1(plot.h + 2)}"/></clipPath></defs>`)
  ctx.out.push(`<g class="chart-series" clip-path="url(#${ctx.id}-plot)">`)
  const totalPoints = plans.reduce((sum, plan) => sum + plan.data.values.length, 0)
  const tooltips = ctx.tooltips && totalPoints <= 3000
  const pct = (plan: SeriesPlan) => options.groupingFor(plan) === 'percentStacked'
  const stacked = (plan: SeriesPlan) => options.groupingFor(plan) !== 'clustered'

  // Areas first (behind), then bars, then lines/scatter on top.
  const areaPlans = plans.filter((plan) => plan.kind === 'area')
  const barPlans = plans.filter((plan) => plan.kind === 'column' || plan.kind === 'bar')
  const linePlans = plans.filter((plan) => plan.kind === 'line')
  const scatterPlans = plans.filter((plan) => plan.kind === 'scatter')

  const stackGroups = (list: SeriesPlan[]) => {
    const groups = new Map<string, SeriesPlan[]>()
    for (const plan of list) { const key = plan.secondary ? 's' : 'p'; groups.set(key, [...(groups.get(key) || []), plan]) }
    return [...groups.values()]
  }

  // --- Areas
  for (const group of stackGroups(areaPlans)) {
    const scale = group[0].secondary ? y2Scale : yScale
    const lower = new Array(count).fill(0)
    const totals = stackTotals(group, count).totals
    for (const plan of group) {
      const isStacked = stacked(plan)
      const upper = new Array(count).fill(0)
      for (let i = 0; i < count; i += 1) {
        let v = plan.data.values[i] ?? 0
        if (pct(plan)) v = totals[i] ? v / totals[i] : 0
        upper[i] = (isStacked ? lower[i] : 0) + v
      }
      const top = categoryCenters.map((x, i) => [x, valuePos(scale, upper[i])] as [number, number])
      const bottom = categoryCenters.map((x, i) => [x, valuePos(scale, isStacked ? lower[i] : baseValue(scale))] as [number, number]).reverse()
      if (count === 1) { top.push([top[0][0] + band / 4, top[0][1]]); top[0] = [top[0][0] - band / 4, top[0][1]]; bottom.unshift([bottom[0][0] + band / 4, bottom[0][1]]); bottom[bottom.length - 1] = [bottom[bottom.length - 1][0] - band / 4, bottom[bottom.length - 1][1]] }
      const path = [...top, ...bottom].map((p, i) => `${i ? 'L' : 'M'}${r1(p[0])} ${r1(p[1])}`).join('') + 'Z'
      const opacity = isStacked || group.length === 1 ? '' : ' fill-opacity="0.82"'
      ctx.out.push(`<path class="chart-area" data-series="${plan.index}" d="${path}" fill="${plan.data.color}"${opacity}/>`)
      pointLabels(ctx, labels, plan, categoryCenters.map((x, i) => [x, valuePos(scale, upper[i])]), 'center', horizontal, totals)
      if (isStacked) for (let i = 0; i < count; i += 1) lower[i] = upper[i]
    }
  }

  // --- Bars
  if (barPlans.length) {
    const first = barPlans[0]
    const isStacked = stacked(first)
    const gapWidth = Math.max(0, Math.min(500, chart.gapWidth ?? (isStacked ? 150 : 219))) / 100
    const overlap = Math.max(-100, Math.min(100, chart.overlap ?? (isStacked ? 100 : -27))) / 100
    const slots = isStacked ? 1 : barPlans.length
    const barWidth = band / (slots - (slots - 1) * overlap + gapWidth)
    const clusterWidth = barWidth * (slots - (slots - 1) * overlap)
    const lowerPos = new Map<string, number[]>(), lowerNeg = new Map<string, number[]>()
    const totalsByGroup = new Map<string, number[]>()
    for (const group of stackGroups(barPlans)) totalsByGroup.set(group[0].secondary ? 's' : 'p', stackTotals(group, count).totals)
    barPlans.forEach((plan, planIndex) => {
      const scale = plan.secondary ? y2Scale : yScale
      const key = plan.secondary ? 's' : 'p'
      if (!lowerPos.has(key)) { lowerPos.set(key, new Array(count).fill(0)); lowerNeg.set(key, new Array(count).fill(0)) }
      const pos = lowerPos.get(key)!, neg = lowerNeg.get(key)!
      const totals = totalsByGroup.get(key)!
      const slot = isStacked ? 0 : planIndex
      const labelPoints: Array<[number, number] | null> = []
      const parts: string[] = []
      for (let i = 0; i < count; i += 1) {
        let v = plan.data.values[i]
        if (v === null || v === undefined) { labelPoints.push(null); continue }
        if (pct(plan)) v = totals[i] ? v / totals[i] : 0
        let startValue = baseValue(scale), endValue = v
        if (isStacked) {
          if (v >= 0) { startValue = pos[i]; endValue = pos[i] + v; pos[i] = endValue } else { startValue = neg[i]; endValue = neg[i] + v; neg[i] = endValue }
        }
        const a = valuePos(scale, startValue), b = valuePos(scale, endValue)
        const center = categoryCenters[i]
        const offset = -clusterWidth / 2 + slot * barWidth * (1 - overlap)
        const fill = plan.data.pointColors?.[i] || plan.data.color
        const fillAttr = plan.source?.invertIfNegative && v < 0 ? `fill="#FFFFFF" stroke="${fill}"` : `fill="${fill}"`
        const tip = tooltips ? tooltip(ctx, `${plan.data.name}, ${data.categories[i] ?? ''}: ${ctx.format(plan.data.values[i] as number, plan.data.numFmt)}`) : ''
        if (horizontal) {
          // Excel stacks horizontal clusters bottom-up, so the first series sits lowest.
          const y = center + clusterWidth / 2 - (slot + 1) * barWidth * (1 - overlap) - barWidth * overlap
          const x = Math.min(a, b), w = Math.max(Math.abs(b - a), v === 0 ? 0 : 0.5)
          parts.push(`<rect class="chart-bar" data-series="${plan.index}" data-point="${i}" x="${r2(x)}" y="${r2(y)}" width="${r2(w)}" height="${r2(Math.max(0.5, barWidth))}" ${fillAttr}>${tip}</rect>`)
          labelPoints.push([endValue >= startValue ? x + w : x, y + barWidth / 2])
        } else {
          const x = center + offset
          const y = Math.min(a, b), h = Math.max(Math.abs(b - a), v === 0 ? 0 : 0.5)
          parts.push(`<rect class="chart-bar" data-series="${plan.index}" data-point="${i}" x="${r2(x)}" y="${r2(y)}" width="${r2(Math.max(0.5, barWidth))}" height="${r2(h)}" ${fillAttr}>${tip}</rect>`)
          labelPoints.push([x + barWidth / 2, endValue >= startValue ? y : y + h])
        }
      }
      ctx.out.push(parts.join(''))
      barLabels(ctx, labels, plan, labelPoints, isStacked, horizontal, totals, barWidth, (i) => {
        const v = plan.data.values[i]
        if (v === null || v === undefined) return 0
        return Math.abs(valuePos(scale, pct(plan) ? (totals[i] ? v / totals[i] : 0) : v) - valuePos(scale, 0))
      })
    })
  }

  // --- Lines
  for (const group of stackGroups(linePlans)) {
    const scale = group[0].secondary ? y2Scale : yScale
    const lower = new Array(count).fill(0)
    const totals = stackTotals(group, count).totals
    for (const plan of group) {
      const isStacked = stacked(plan) && chart.type !== 'combo'
      const blanks = chart.displayBlanksAs || 'gap'
      const segments: Array<Array<[number, number]>> = [[]]
      const points: Array<[number, number] | null> = []
      for (let i = 0; i < count; i += 1) {
        let v = plan.data.values[i]
        if (v === null || v === undefined) {
          if (blanks === 'zero') v = 0
          else { points.push(null); if (blanks === 'gap' && segments[segments.length - 1].length) segments.push([]); continue }
        }
        if (pct(plan) && isStacked) v = totals[i] ? v / totals[i] : 0
        const value = isStacked ? lower[i] + v : v
        if (isStacked) lower[i] = value
        const point: [number, number] = horizontal ? [valuePos(scale, value), categoryCenters[i]] : [categoryCenters[i], valuePos(scale, value)]
        points.push(point)
        segments[segments.length - 1].push(point)
      }
      const width = (plan.source?.lineWidth ?? 2.25) * PT
      const smooth = plan.source?.smooth === true
      if (plan.source?.showLine !== false) {
        const d = segments.filter((segment) => segment.length).map((segment) => linePath(segment, smooth)).join('')
        if (d) ctx.out.push(`<path class="chart-line" data-series="${plan.index}" d="${d}" fill="none" stroke="${plan.data.color}" stroke-width="${r1(width)}" stroke-linejoin="round" stroke-linecap="round"/>`)
      }
      drawMarkers(ctx, plan, points, tooltips)
      pointLabels(ctx, labels, plan, points, 'above', horizontal, totals)
    }
  }

  // --- Scatter
  if (scatter && xScale) {
    for (const plan of scatterPlans.length ? scatterPlans : plans) {
      const scale = plan.secondary ? y2Scale : yScale
      const xs = plan.data.x || []
      const points: Array<[number, number] | null> = []
      const segments: Array<Array<[number, number]>> = [[]]
      for (let i = 0; i < plan.data.values.length; i += 1) {
        const v = plan.data.values[i], xv = xs[i]
        if (v === null || v === undefined || xv === null || xv === undefined) { points.push(null); if (segments[segments.length - 1].length) segments.push([]); continue }
        const point: [number, number] = [plot.x + xScale.norm(xv) * plot.w, valuePos(scale, v)]
        points.push(point)
        segments[segments.length - 1].push(point)
      }
      if (plan.source?.showLine === true) {
        const d = segments.filter((segment) => segment.length).map((segment) => linePath(segment, plan.source?.smooth === true)).join('')
        if (d) ctx.out.push(`<path class="chart-line" data-series="${plan.index}" d="${d}" fill="none" stroke="${plan.data.color}" stroke-width="${r1((plan.source?.lineWidth ?? 2.25) * PT)}" stroke-linejoin="round"/>`)
      }
      drawMarkers(ctx, plan, points, tooltips, (i) => `${plan.data.name}: (${ctx.format(xs[i] as number)}, ${ctx.format(plan.data.values[i] as number, plan.data.numFmt)})`)
      pointLabels(ctx, labels, plan, points, 'right', false)
    }
  }
  ctx.out.push('</g>')
  return { labels, categoryCenters }
}

function drawMarkers(ctx: Ctx, plan: SeriesPlan, points: Array<[number, number] | null>, tooltips: boolean, tipFor?: (index: number) => string) {
  const symbol = effectiveMarker(ctx.chart, plan.source, plan.kind)
  if (symbol === 'none') return
  const size = Math.max(2, Math.min(40, (plan.source?.markerSize ?? 5) * PT))
  const color = plan.data.color
  const parts: string[] = []
  const stroked = strokedMarker(symbol)
  if (!tooltips) {
    // One path for all markers keeps large series cheap.
    const d = points.filter(Boolean).map((p) => markerPath(symbol, p![0], p![1], size)).join('')
    if (d) parts.push(stroked
      ? `<path class="chart-marker" data-series="${plan.index}" d="${d}" fill="none" stroke="${color}" stroke-width="1.5"/>`
      : `<path class="chart-marker" data-series="${plan.index}" d="${d}" fill="${color}" stroke="${color}" stroke-width="0.75"/>`)
  } else {
    points.forEach((p, i) => {
      if (!p) return
      const tip = tooltip(ctx, tipFor ? tipFor(i) : `${plan.data.name}, ${ctx.data.categories[i] ?? ''}: ${ctx.format(plan.data.values[i] as number, plan.data.numFmt)}`)
      parts.push(stroked
        ? `<path class="chart-marker" data-series="${plan.index}" data-point="${i}" d="${markerPath(symbol, p[0], p[1], size)}" fill="none" stroke="${color}" stroke-width="1.5">${tip}</path>`
        : `<path class="chart-marker" data-series="${plan.index}" data-point="${i}" d="${markerPath(symbol, p[0], p[1], size)}" fill="${color}" stroke="${color}" stroke-width="0.75">${tip}</path>`)
    })
  }
  ctx.out.push(parts.join(''))
}

function pointLabels(ctx: Ctx, out: string[], plan: SeriesPlan, points: Array<[number, number] | null>, fallback: 'above' | 'center' | 'right', horizontal: boolean, totals?: number[]) {
  const labels = labelsFor(plan.source)
  if (!labels || ctx.thumbnail) return
  const size = ctx.axisPx
  const position = labels.position && labels.position !== 'auto' ? labels.position : fallback
  points.forEach((p, i) => {
    if (!p) return
    const value = plan.data.values[i]
    const content = labelText(ctx, labels, value ?? null, ctx.data.categories[i] ?? '', plan.data.name, totals?.[i] && value != null ? value / totals[i] : undefined, plan.data.numFmt)
    if (!content) return
    let x = p[0], y = p[1], anchor: 'start' | 'middle' | 'end' = 'middle'
    if (position === 'above' || position === 'outEnd' || position === 'inEnd') y -= size * 0.75
    else if (position === 'below' || position === 'inBase') y += size * 1.1
    else if (position === 'right') { x += 7; anchor = 'start' }
    else if (position === 'left') { x -= 7; anchor = 'end' }
    if (horizontal && position === 'above') { x = p[0] + 6; y = p[1]; anchor = 'start' }
    out.push(`<text x="${r1(x)}" y="${r1(y)}" font-size="${r1(size)}" fill="${LABEL_TEXT}" text-anchor="${anchor}" dominant-baseline="central" class="chart-data-label">${esc(content)}</text>`)
  })
}

function barLabels(ctx: Ctx, out: string[], plan: SeriesPlan, points: Array<[number, number] | null>, stacked: boolean, horizontal: boolean, totals: number[], barWidth: number, lengthOf: (index: number) => number) {
  const labels = labelsFor(plan.source)
  if (!labels || ctx.thumbnail) return
  const size = ctx.axisPx
  let position = labels.position && labels.position !== 'auto' ? labels.position : stacked ? 'center' : 'outEnd'
  if (stacked && position === 'outEnd') position = 'inEnd'
  points.forEach((p, i) => {
    if (!p) return
    const value = plan.data.values[i]
    if (value === null || value === undefined) return
    const content = labelText(ctx, labels, value, ctx.data.categories[i] ?? '', plan.data.name, totals[i] ? value / totals[i] : undefined, plan.data.numFmt)
    if (!content) return
    const length = lengthOf(i)
    const negative = value < 0
    const inside = position === 'center' || position === 'inEnd' || position === 'inBase'
    const fits = horizontal ? measureText(content, size) + 6 <= length : size + 4 <= length && measureText(content, size) <= barWidth + 12
    const color = inside && fits ? contrastText(plan.data.pointColors?.[i] || plan.data.color) : LABEL_TEXT
    let x = p[0], y = p[1], anchor: 'start' | 'middle' | 'end' = 'middle'
    const dir = negative ? -1 : 1
    if (horizontal) {
      if (position === 'outEnd' || !fits) { x += dir * 5; anchor = negative ? 'end' : 'start' }
      else if (position === 'inEnd') { x -= dir * 5; anchor = negative ? 'start' : 'end' }
      else if (position === 'center') { x -= dir * length / 2 }
      else { x -= dir * (length - 5); anchor = negative ? 'end' : 'start' }
    } else {
      if (position === 'outEnd' || !fits) y -= dir * size * 0.8
      else if (position === 'inEnd') y += dir * size * 0.85
      else if (position === 'center') y += dir * length / 2
      else y += dir * (length - size * 0.85)
    }
    out.push(`<text x="${r1(x)}" y="${r1(y)}" font-size="${r1(size)}" fill="${color}" text-anchor="${anchor}" dominant-baseline="central" class="chart-data-label">${esc(content)}</text>`)
  })
}

function contrastText(hex: string) {
  const value = parseInt(hex.replace('#', ''), 16)
  if (!Number.isFinite(value)) return '#FFFFFF'
  const r = (value >> 16) & 255, g = (value >> 8) & 255, b = value & 255
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? '#262626' : '#FFFFFF'
}

// ---------------------------------------------------------------------------
// Pie / doughnut
// ---------------------------------------------------------------------------

function arcPoint(cx: number, cy: number, radius: number, angle: number): [number, number] {
  return [cx + radius * Math.cos(angle), cy + radius * Math.sin(angle)]
}

function renderPie(ctx: Ctx, plans: SeriesPlan[], box: Box) {
  const { chart, data } = ctx
  const doughnut = chart.type === 'doughnut'
  const rings = doughnut ? plans : plans.slice(0, 1)
  if (!rings.length) return
  const anyLabels = rings.some((plan) => labelsFor(plan.source)) && !ctx.thumbnail
  // Reserve room for outside labels only when some slice is too thin to hold its label.
  let outsideRoom = 0
  if (anyLabels && !doughnut) {
    const plan = rings[0]
    const labels = labelsFor(plan.source)!
    const values = plan.data.values.map((v) => (v === null || !Number.isFinite(v) ? 0 : Math.abs(v)))
    const total = values.reduce((sum, v) => sum + v, 0) || 1
    const provisional = Math.min(box.w, box.h) / 2 * 0.94
    const needsOutside = labels.position === 'outEnd' || values.some((value, i) => {
      if (!value) return false
      const content = labelText(ctx, labels, plan.data.values[i], ctx.data.categories[i] ?? '', plan.data.name, value / total, plan.data.numFmt)
      const chord = 2 * provisional * 0.62 * Math.sin(Math.min(Math.PI / 2, (value / total) * Math.PI))
      return !(chord > Math.max(measureText(content, ctx.axisPx) * 0.8, ctx.axisPx * 1.6))
    })
    if (needsOutside) outsideRoom = ctx.axisPx * 2.2
  }
  const radius = Math.max(4, Math.min(box.w / 2 - outsideRoom * 1.8, box.h / 2 - outsideRoom) * (ctx.thumbnail ? 0.95 : 0.94))
  const cx = box.x + box.w / 2, cy = box.y + box.h / 2
  const hole = doughnut ? Math.max(10, Math.min(90, chart.holeSize ?? 75)) / 100 : 0
  const inner = radius * hole
  const ringWidth = (radius - inner) / rings.length
  const start = ((chart.firstSliceAngle ?? 0) - 90) * Math.PI / 180
  const labels: string[] = []
  const stroke = ctx.thumbnail ? 0.75 : 1.25
  rings.forEach((plan, ringIndex) => {
    const values = plan.data.values.map((v) => (v === null || !Number.isFinite(v) ? 0 : Math.abs(v)))
    const total = values.reduce((sum, v) => sum + v, 0)
    if (!total) return
    const outer = doughnut ? radius - ringIndex * ringWidth : radius
    const innerRadius = doughnut ? outer - ringWidth : 0
    let angle = start
    const pointLabels = labelsFor(plan.source)
    values.forEach((value, i) => {
      if (!value) return
      const sweep = (value / total) * Math.PI * 2
      const end = angle + sweep
      const color = plan.data.pointColors?.[i] || plan.data.color
      const large = sweep > Math.PI ? 1 : 0
      let d: string
      if (sweep >= Math.PI * 2 - 1e-9) {
        const [ax, ay] = arcPoint(cx, cy, outer, angle), [bx, by] = arcPoint(cx, cy, outer, angle + Math.PI)
        d = `M${r1(ax)} ${r1(ay)}A${r1(outer)} ${r1(outer)} 0 1 1 ${r1(bx)} ${r1(by)}A${r1(outer)} ${r1(outer)} 0 1 1 ${r1(ax)} ${r1(ay)}Z`
        if (innerRadius > 0) {
          const [ix, iy] = arcPoint(cx, cy, innerRadius, angle), [jx, jy] = arcPoint(cx, cy, innerRadius, angle + Math.PI)
          d += `M${r1(ix)} ${r1(iy)}A${r1(innerRadius)} ${r1(innerRadius)} 0 1 0 ${r1(jx)} ${r1(jy)}A${r1(innerRadius)} ${r1(innerRadius)} 0 1 0 ${r1(ix)} ${r1(iy)}Z`
        }
      } else if (innerRadius > 0) {
        const [ax, ay] = arcPoint(cx, cy, outer, angle), [bx, by] = arcPoint(cx, cy, outer, end)
        const [ix, iy] = arcPoint(cx, cy, innerRadius, end), [jx, jy] = arcPoint(cx, cy, innerRadius, angle)
        d = `M${r1(ax)} ${r1(ay)}A${r1(outer)} ${r1(outer)} 0 ${large} 1 ${r1(bx)} ${r1(by)}L${r1(ix)} ${r1(iy)}A${r1(innerRadius)} ${r1(innerRadius)} 0 ${large} 0 ${r1(jx)} ${r1(jy)}Z`
      } else {
        const [ax, ay] = arcPoint(cx, cy, outer, angle), [bx, by] = arcPoint(cx, cy, outer, end)
        d = `M${r1(cx)} ${r1(cy)}L${r1(ax)} ${r1(ay)}A${r1(outer)} ${r1(outer)} 0 ${large} 1 ${r1(bx)} ${r1(by)}Z`
      }
      const tip = ctx.tooltips ? tooltip(ctx, `${data.categories[i] ?? ''}: ${ctx.format(plan.data.values[i] as number, plan.data.numFmt)} (${ctx.format(value / total, '0.0%')})`) : ''
      ctx.out.push(`<path class="chart-slice" data-series="${plan.index}" data-point="${i}" d="${d}" fill="${color}" stroke="#FFFFFF" stroke-width="${stroke}" fill-rule="evenodd">${tip}</path>`)
      if (pointLabels && !ctx.thumbnail) {
        const content = labelText(ctx, pointLabels, plan.data.values[i], data.categories[i] ?? '', plan.data.name, value / total, plan.data.numFmt)
        if (content) {
          const mid = angle + sweep / 2
          const size = ctx.axisPx
          const textWidth = measureText(content, size)
          const position = pointLabels.position || 'bestFit'
          const insideRadius = innerRadius > 0 ? (outer + innerRadius) / 2 : outer * 0.62
          const [ix, iy] = arcPoint(cx, cy, insideRadius, mid)
          const chord = 2 * insideRadius * Math.sin(Math.min(Math.PI / 2, sweep / 2))
          const fitsInside = doughnut || position === 'center' || position === 'inEnd' || (position !== 'outEnd' && chord > Math.max(textWidth * 0.8, size * 1.6) && outer > 40)
          if (fitsInside) {
            labels.push(`<text x="${r1(ix)}" y="${r1(iy)}" font-size="${r1(size)}" fill="${contrastText(color)}" text-anchor="middle" dominant-baseline="central" class="chart-data-label">${esc(content)}</text>`)
          } else {
            const [ex, ey] = arcPoint(cx, cy, outer, mid)
            const [ox, oy] = arcPoint(cx, cy, outer + size * 1.1, mid)
            const right = Math.cos(mid) >= 0
            const tx = ox + (right ? 4 : -4)
            labels.push(`<path d="M${r1(ex)} ${r1(ey)}L${r1(ox)} ${r1(oy)}" stroke="#A6A6A6" stroke-width="0.75" fill="none"/>`)
            labels.push(`<text x="${r1(tx)}" y="${r1(oy)}" font-size="${r1(size)}" fill="${LABEL_TEXT}" text-anchor="${right ? 'start' : 'end'}" dominant-baseline="central" class="chart-data-label">${esc(content)}</text>`)
          }
        }
      }
      angle = end
    })
  })
  ctx.out.push(labels.join(''))
}

// ---------------------------------------------------------------------------
// Radar
// ---------------------------------------------------------------------------

function renderRadar(ctx: Ctx, plans: SeriesPlan[], box: Box) {
  const { data, chart } = ctx
  const count = Math.max(data.categories.length, ...plans.map((plan) => plan.data.values.length), 0)
  if (count < 1) return
  const size = ctx.axisPx
  const widestLabel = Math.min(110, Math.max(0, ...data.categories.slice(0, count).map((label) => measureText(label, size))))
  const labelRoom = ctx.thumbnail ? 2 : widestLabel + size
  const radius = Math.max(8, ctx.thumbnail ? Math.min(box.w, box.h) / 2 - 2 : Math.min(box.w / 2 - labelRoom, box.h / 2 - size * 1.7))
  const cx = box.x + box.w / 2, cy = box.y + box.h / 2
  let min = Infinity, max = -Infinity
  for (const plan of plans) for (const v of plan.data.values) if (v !== null && Number.isFinite(v)) { min = Math.min(min, v); max = Math.max(max, v) }
  const scale = linearScale(Number.isFinite(min) ? Math.min(0, min) : 0, Number.isFinite(max) ? max : 1, Math.max(2, Math.min(8, Math.floor(radius / (size * 1.6)))), chart.axes?.y)
  const angleOf = (i: number) => -Math.PI / 2 + (i / count) * Math.PI * 2
  ctx.out.push('<g class="chart-grid">')
  for (const tick of scale.ticks) {
    const rr = scale.norm(tick) * radius
    if (rr <= 0) continue
    const d = Array.from({ length: count }, (_, i) => arcPoint(cx, cy, rr, angleOf(i))).map((p, i) => `${i ? 'L' : 'M'}${r1(p[0])} ${r1(p[1])}`).join('') + 'Z'
    ctx.out.push(`<path d="${d}" fill="none" stroke="${GRID}" stroke-width="1"/>`)
  }
  for (let i = 0; i < count; i += 1) {
    const [x, y] = arcPoint(cx, cy, radius, angleOf(i))
    ctx.out.push(`<path d="M${r1(cx)} ${r1(cy)}L${r1(x)} ${r1(y)}" stroke="${GRID}" stroke-width="1"/>`)
  }
  ctx.out.push('</g>')
  for (const plan of plans) {
    const points = Array.from({ length: count }, (_, i) => {
      const v = plan.data.values[i]
      return v === null || v === undefined ? null : arcPoint(cx, cy, Math.max(0, scale.norm(v)) * radius, angleOf(i))
    })
    const valid = points.filter(Boolean) as Array<[number, number]>
    if (valid.length) ctx.out.push(`<path class="chart-line" data-series="${plan.index}" d="${valid.map((p, i) => `${i ? 'L' : 'M'}${r1(p[0])} ${r1(p[1])}`).join('')}Z" fill="none" stroke="${plan.data.color}" stroke-width="${r1((plan.source?.lineWidth ?? 2.25) * PT)}" stroke-linejoin="round"/>`)
    drawMarkers(ctx, plan, points, ctx.tooltips)
  }
  if (!ctx.thumbnail) {
    for (let i = 0; i < count; i += 1) {
      const angle = angleOf(i)
      const [x, y] = arcPoint(cx, cy, radius + size * 0.9, angle)
      const cos = Math.cos(angle)
      text(ctx, x, y, truncate(data.categories[i] || '', labelRoom * 2.2, size), { anchor: Math.abs(cos) < 0.2 ? 'middle' : cos > 0 ? 'start' : 'end', baseline: 'middle' })
    }
    for (const tick of scale.ticks) {
      const rr = scale.norm(tick) * radius
      text(ctx, cx + 3, cy - rr, ctx.format(tick, chart.axes?.y?.numFmt || plans[0]?.data.numFmt), { baseline: 'middle', size: size * 0.9 })
    }
  }
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

function placeholder(ctx: Ctx, box: Box, message: string, detail?: string) {
  const cx = box.x + box.w / 2, cy = box.y + box.h / 2
  const icon = Math.min(34, box.h / 3)
  ctx.out.push(`<g class="chart-placeholder" opacity="0.9"><rect x="${r1(cx - icon / 2)}" y="${r1(cy - icon - 6)}" width="${r1(icon)}" height="${r1(icon)}" rx="4" fill="#F2F2F2"/>`)
  const barW = icon / 6
  ;[0.45, 0.75, 0.3].forEach((h, i) => ctx.out.push(`<rect x="${r1(cx - icon / 2 + barW * (1 + i * 1.6))}" y="${r1(cy - 6 - icon * h - 3)}" width="${r1(barW)}" height="${r1(icon * h)}" fill="#BFBFBF"/>`))
  ctx.out.push('</g>')
  text(ctx, cx, cy + ctx.axisPx * 0.8, truncate(message, box.w - 16, ctx.axisPx * 1.05), { anchor: 'middle', size: ctx.axisPx * 1.05, color: TEXT })
  if (detail) text(ctx, cx, cy + ctx.axisPx * 2.3, truncate(detail, box.w - 16, ctx.axisPx * 0.95), { anchor: 'middle', size: ctx.axisPx * 0.95, color: '#8C8C8C' })
}

/**
 * Render a chart to standalone SVG markup (width x height CSS pixels, with a
 * matching viewBox so it scales cleanly). Pure and synchronous.
 */
export function renderChartSvg(chart: SheetChart, input: ResolvedChartData, width: number, height: number, options: ChartRenderOptions = {}): string {
  // Colours are normalised here: this markup is injected as HTML, so no caller string reaches it unescaped.
  const data: ResolvedChartData = {
    ...input,
    categories: Array.isArray(input?.categories) ? input.categories : [],
    series: (Array.isArray(input?.series) ? input.series : []).map((item, index) => {
      const color = normalizeHex(item.color) || paletteColor(index)
      return { ...item, values: Array.isArray(item.values) ? item.values : [], color, pointColors: item.pointColors?.map((point) => normalizeHex(point) || color) }
    }),
  }
  const w = Math.max(20, Math.round(Number(width) || 0))
  const h = Math.max(20, Math.round(Number(height) || 0))
  const style = chart.style || {}
  const font = options.fontFamily || style.fontFamily || DEFAULT_FONT
  const axisPx = (style.fontSize && style.fontSize > 3 && style.fontSize < 40 ? style.fontSize : 9) * PT
  const titlePx = (style.titleFontSize && style.titleFontSize > 4 && style.titleFontSize < 72 ? style.titleFontSize : 14) * PT
  const ctx: Ctx = {
    chart,
    data,
    width: w,
    height: h,
    font,
    axisPx,
    titlePx,
    textColor: normalizeHex(style.textColor) || TEXT,
    format: options.formatNumber ? (value, fmt) => { try { return options.formatNumber!(value, fmt) } catch { return defaultFormatNumber(value, fmt) } } : defaultFormatNumber,
    out: [],
    id: (options.idPrefix || `c${hashId(chart.id || 'chart')}`).replace(/[^A-Za-z0-9_-]/g, '') || 'chart',
    tooltips: options.tooltips !== false && !options.thumbnail,
    thumbnail: options.thumbnail === true,
  }
  const title = ctx.thumbnail ? '' : chartDisplayTitle(chart, data)
  // Colours are normalised: this markup is injected as HTML, so no model string reaches it unescaped.
  const background = style.background === 'transparent' ? 'none' : normalizeHex(style.background) || '#FFFFFF'
  const border = style.border === null ? 'none' : normalizeHex(style.border) || (ctx.thumbnail ? 'none' : '#D9D9D9')
  const radius = style.roundedCorners ? Math.min(10, w / 12) : 0
  const out = ctx.out
  out.push(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" font-family="${esc(font)}" fill="${esc(ctx.textColor)}" role="img" aria-label="${esc(title || 'Chart')}" class="chart-svg">`)
  out.push(`<rect class="chart-background" x="0.5" y="0.5" width="${w - 1}" height="${h - 1}"${radius ? ` rx="${r1(radius)}"` : ''} fill="${background}" stroke="${border}" stroke-width="1"/>`)
  const pad = ctx.thumbnail ? 3 : Math.max(6, Math.min(12, Math.min(w, h) * 0.035))
  let box: Box = { x: pad, y: pad, w: w - pad * 2, h: h - pad * 2 }
  if (title) {
    const lines = wrapText(title, box.w - 8, titlePx, 2)
    const lineHeight = titlePx * 1.2
    lines.forEach((line, index) => text(ctx, w / 2, box.y + titlePx * 0.95 + index * lineHeight, line, { anchor: 'middle', size: titlePx, cls: 'chart-title' }))
    const used = lines.length * lineHeight + titlePx * 0.35
    box = { ...box, y: box.y + used, h: box.h - used }
  }
  const series = Array.isArray(chart.series) ? chart.series : []
  const plans: SeriesPlan[] = ctx.data.series.map((item, index) => {
    return {
    index,
    source: series[index],
    data: item,
    kind: chartSeriesType(chart, series[index]),
    secondary: chart.type === 'combo' && series[index]?.secondaryAxis === true,
    }
  })
  if (chart.type === 'bar') plans.forEach((plan) => { plan.kind = 'bar' })
  if (chart.type === 'combo') plans.forEach((plan) => { if (plan.kind === 'bar') plan.kind = 'column'; if (plan.kind === 'scatter') plan.kind = 'line' })
  if (chart.type === 'unsupported') {
    placeholder(ctx, box, 'Unsupported chart type', chart.unsupportedKind ? `${chart.unsupportedKind} charts are kept when saved as .xlsx` : 'Kept when saved as .xlsx')
    out.push('</svg>')
    return out.join('')
  }
  const hasData = plans.some((plan) => plan.data.values.some((value) => value !== null && Number.isFinite(value)))
  if (!hasData) {
    placeholder(ctx, box, options.emptyMessage || 'No data to plot', ctx.thumbnail ? undefined : 'Choose a range with numbers')
    out.push('</svg>')
    return out.join('')
  }
  const legendPosition = chart.legend ?? (plans.length > 1 || chart.type === 'pie' || chart.type === 'doughnut' ? 'bottom' : 'none')
  if (!ctx.thumbnail && legendPosition !== 'none' && box.h > 60) box = drawLegend(ctx, legendEntries(ctx, plans), legendPosition, box)
  if (box.w < 10 || box.h < 10) { out.push('</svg>'); return out.join('') }
  if (chart.type === 'pie' || chart.type === 'doughnut') renderPie(ctx, plans, box)
  else if (chart.type === 'radar') renderRadar(ctx, plans, box)
  else renderCartesian(ctx, plans, box)
  out.push('</svg>')
  return out.join('')
}

// ---------------------------------------------------------------------------
// Print / PDF payload
// ---------------------------------------------------------------------------

/** One chart placed on printed pages: see electron/spreadsheet-print.cjs (input.charts). */
export interface PrintChartEntry {
  id: string
  svg: string
  from: ChartAnchorPoint
  to: ChartAnchorPoint
}

// Mirrors columnWidth/rowHeight in electron/spreadsheet-print.cjs so SVGs are drawn at their printed size.
function printColumnWidth(sheet: SheetData, col: number) {
  if (sheet.hiddenCols?.includes(col + 1)) return 0
  const properties = (sheet.properties || {}) as Record<string, unknown>
  const value = Number(sheet.colWidths?.[String(col + 1)] ?? properties.defaultColWidth)
  if (!Number.isFinite(value) || value <= 0) return 64
  const measured = Number(properties.printDigitWidth)
  const digit = Number.isFinite(measured) && measured >= 1 && measured <= 40 ? measured : 8
  return Math.max(2, Math.min(1_000, value * digit))
}

function printRowHeight(sheet: SheetData, row: number) {
  if (sheet.hiddenRows?.includes(row + 1)) return 0
  const properties = (sheet.properties || {}) as Record<string, unknown>
  const value = Number(sheet.rowHeights?.[String(row + 1)] ?? properties.defaultRowHeight)
  if (!Number.isFinite(value) || value <= 0) return 20
  return Math.max(2, Math.min(640, (value * 4) / 3))
}

/** Printed size (CSS px) of an anchor on `sheet`. */
export function printedAnchorSize(sheet: SheetData, from: ChartAnchorPoint, to: ChartAnchorPoint) {
  const span = (start: number, end: number, size: (index: number) => number) => {
    let total = 0
    for (let index = start; index < end && index - start < 5_000; index += 1) total += size(index)
    return total
  }
  const fromX = Math.min(printColumnWidth(sheet, from.col), (from.colOffsetEmu || 0) / 9525)
  const toX = Math.min(printColumnWidth(sheet, to.col), (to.colOffsetEmu || 0) / 9525)
  const fromY = Math.min(printRowHeight(sheet, from.row), (from.rowOffsetEmu || 0) / 9525)
  const toY = Math.min(printRowHeight(sheet, to.row), (to.rowOffsetEmu || 0) / 9525)
  return {
    width: span(from.col, to.col, (col) => printColumnWidth(sheet, col)) + toX - fromX,
    height: span(from.row, to.row, (row) => printRowHeight(sheet, row)) + toY - fromY,
  }
}

/**
 * Build `input.charts` for renderPrintPreview/printWorkbook/exportWorkbook(pdf|html):
 * `{ [sheetId]: PrintChartEntry[] }`, each SVG rendered at its printed size.
 */
export function buildPrintChartPayload(
  workbook: WorkbookModel,
  resolve: (chart: SheetChart, sheet: SheetData) => ResolvedChartData,
  options: ChartRenderOptions & { sheetIds?: string[] } = {},
): Record<string, PrintChartEntry[]> {
  const output: Record<string, PrintChartEntry[]> = {}
  let counter = 0
  for (const sheet of workbook.sheets) {
    if (!sheet.charts?.length || (options.sheetIds && !options.sheetIds.includes(sheet.id))) continue
    const entries: PrintChartEntry[] = []
    for (const chart of sheet.charts) {
      if (!chart?.anchor?.from || !chart.anchor.to) continue
      const { width, height } = printedAnchorSize(sheet, chart.anchor.from, chart.anchor.to)
      if (width < 4 || height < 4) continue
      let data: ResolvedChartData
      try { data = resolve(chart, sheet) } catch { data = { categories: [], series: [] } }
      counter += 1
      const svg = renderChartSvg(chart, data, width, height, { ...options, tooltips: false, idPrefix: `print-chart-${counter}` })
      entries.push({ id: chart.id, svg, from: chart.anchor.from, to: chart.anchor.to })
    }
    if (entries.length) output[sheet.id] = entries
  }
  return output
}

/** Small static thumbnail for the chart-type gallery (no text). */
export function renderChartThumbnail(chart: SheetChart, data: ResolvedChartData, width = 64, height = 44) {
  return renderChartSvg(chart, data, width, height, { thumbnail: true, idPrefix: `thumb-${chart.type}-${chart.grouping || 'c'}` })
}
