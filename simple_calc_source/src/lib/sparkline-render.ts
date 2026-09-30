/**
 * Draws a SPARKLINE() result as SVG markup sized to its cell: line (default), column,
 * winloss and bar chart types with Google Sheets' options (color, negcolor, highcolor,
 * lowcolor, firstcolor, lastcolor, linewidth, ymin/ymax, max, axis, axiscolor, rtl, empty,
 * color1/color2).
 */
import type { SparklineSpec } from './formula-lib-sparkline'
import type { SparklineGroup } from '../spreadsheet-types'

const NAMED: Record<string, string> = {
  red: '#d93025', green: '#188038', blue: '#1a73e8', black: '#202124', gray: '#80868b', grey: '#80868b', orange: '#e8710a',
  yellow: '#f9ab00', purple: '#9334e6', pink: '#e52592', teal: '#12b5cb', white: '#ffffff', brown: '#8d6e63', navy: '#174ea6',
}

function colorOf(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback
  const text = value.trim().toLowerCase()
  if (NAMED[text]) return NAMED[text]
  if (/^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(text)) return text
  return fallback
}

function numberOf(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) return value
  if (typeof value === 'string' && value.trim() !== '' && Number.isFinite(Number(value))) return Number(value)
  return null
}

function truthy(value: unknown) {
  return value === true || (typeof value === 'string' && value.trim().toLowerCase() === 'true') || value === 1
}

const round = (value: number) => Math.round(value * 100) / 100

export function renderSparklineSvg(spec: SparklineSpec, width: number, height: number): string {
  const w = Math.max(4, width)
  const h = Math.max(4, height)
  const options = spec.options
  const type = String(options.charttype || 'line').toLowerCase()
  const zero = String(options.empty || '').toLowerCase() === 'zero'
  let data = spec.data.map((value) => (value === null && zero ? 0 : value))
  if (truthy(options.rtl)) data = data.slice().reverse()
  const numbers = data.filter((value): value is number => value !== null)
  const pad = 1.5
  const shapes: string[] = []
  if (!numbers.length) return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"></svg>`

  if (type === 'bar') {
    const max = numberOf(options.max) ?? numbers.reduce((sum, value) => sum + Math.max(0, value), 0)
    const colors = [colorOf(options.color1, '#1a73e8'), colorOf(options.color2, '#8ab4f8')]
    let x = 0
    numbers.forEach((value, index) => {
      if (value <= 0 || max <= 0) return
      const segment = Math.min(w - x, (value / max) * w)
      if (segment <= 0) return
      shapes.push(`<rect x="${round(x)}" y="${round(h * 0.15)}" width="${round(segment)}" height="${round(h * 0.7)}" fill="${colors[index % 2]}"/>`)
      x += segment
    })
  } else if (type === 'column' || type === 'winloss') {
    const winloss = type === 'winloss'
    const count = data.length
    const slot = (w - pad * 2) / count
    const barWidth = Math.max(1, slot * 0.72)
    const ymin = winloss ? -1 : Math.min(0, numberOf(options.ymin) ?? Math.min(...numbers))
    const ymax = winloss ? 1 : Math.max(0, numberOf(options.ymax) ?? Math.max(...numbers))
    const span = ymax - ymin || 1
    const y = (value: number) => pad + (ymax - value) / span * (h - pad * 2)
    const base = y(0)
    const high = Math.max(...numbers)
    const low = Math.min(...numbers)
    const color = colorOf(options.color, '#1a73e8')
    const neg = colorOf(options.negcolor, winloss ? '#d93025' : color)
    const firstIndex = data.findIndex((value) => value !== null)
    let lastIndex = -1
    data.forEach((value, index) => { if (value !== null) lastIndex = index })
    data.forEach((value, index) => {
      if (value === null) return
      const shown = winloss ? Math.sign(value) : value
      if (winloss && shown === 0) return
      let fill = value < 0 ? neg : color
      if (options.highcolor !== undefined && value === high) fill = colorOf(options.highcolor, fill)
      if (options.lowcolor !== undefined && value === low) fill = colorOf(options.lowcolor, fill)
      if (options.firstcolor !== undefined && index === firstIndex) fill = colorOf(options.firstcolor, fill)
      if (options.lastcolor !== undefined && index === lastIndex) fill = colorOf(options.lastcolor, fill)
      const top = Math.min(y(shown), base)
      const bottom = Math.max(y(shown), base)
      shapes.push(`<rect x="${round(pad + index * slot + (slot - barWidth) / 2)}" y="${round(top)}" width="${round(barWidth)}" height="${round(Math.max(1, bottom - top))}" fill="${fill}"/>`)
    })
    if (truthy(options.axis) || winloss) shapes.push(`<line x1="0" y1="${round(base)}" x2="${w}" y2="${round(base)}" stroke="${colorOf(options.axiscolor, '#9aa0a6')}" stroke-width="0.8"/>`)
  } else {
    const count = data.length
    const ymin = numberOf(options.ymin) ?? Math.min(...numbers)
    const ymax = numberOf(options.ymax) ?? Math.max(...numbers)
    const span = ymax - ymin || 1
    const lineWidth = Math.max(0.5, Math.min(6, numberOf(options.linewidth) ?? 1.3))
    const x = (index: number) => pad + (count === 1 ? (w - pad * 2) / 2 : (index / (count - 1)) * (w - pad * 2))
    const y = (value: number) => pad + lineWidth / 2 + (ymax - value) / span * (h - pad * 2 - lineWidth)
    const segments: string[] = []
    let current: string[] = []
    data.forEach((value, index) => {
      if (value === null) { if (current.length) segments.push(current.join(' ')); current = []; return }
      current.push(`${round(x(index))},${round(y(value))}`)
    })
    if (current.length) segments.push(current.join(' '))
    const color = colorOf(options.color, '#1a73e8')
    if (truthy(options.axis) && ymin < 0 && ymax > 0) shapes.push(`<line x1="0" y1="${round(y(0))}" x2="${w}" y2="${round(y(0))}" stroke="${colorOf(options.axiscolor, '#9aa0a6')}" stroke-width="0.8"/>`)
    for (const points of segments) {
      if (points.includes(' ')) shapes.push(`<polyline points="${points}" fill="none" stroke="${color}" stroke-width="${lineWidth}" stroke-linejoin="round" stroke-linecap="round"/>`)
      else { const [px, py] = points.split(','); shapes.push(`<circle cx="${px}" cy="${py}" r="${lineWidth}" fill="${color}"/>`) }
    }
    const high = Math.max(...numbers)
    const low = Math.min(...numbers)
    const firstIndex = data.findIndex((value) => value !== null)
    let lastIndex = -1
    data.forEach((value, index) => { if (value !== null) lastIndex = index })
    if (truthy(options.markers)) {
      const fill = colorOf(options.markercolor, color)
      data.forEach((value, index) => { if (value !== null) shapes.push(`<circle cx="${round(x(index))}" cy="${round(y(value))}" r="${round(lineWidth + 0.5)}" fill="${fill}"/>`) })
    }
    const marker = (index: number, fill: string) => { const value = data[index]; if (value !== null) shapes.push(`<circle cx="${round(x(index))}" cy="${round(y(value))}" r="${round(lineWidth + 0.9)}" fill="${fill}"/>`) }
    if (options.highcolor !== undefined) marker(data.indexOf(high), colorOf(options.highcolor, color))
    if (options.lowcolor !== undefined) marker(data.indexOf(low), colorOf(options.lowcolor, color))
    if (options.firstcolor !== undefined) marker(firstIndex, colorOf(options.firstcolor, color))
    if (options.lastcolor !== undefined) marker(lastIndex, colorOf(options.lastcolor, color))
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none">${shapes.join('')}</svg>`
}

/** The renderer spec for one sparkline of an Excel group (axis: group or custom bounds). */
export function groupSparklineSpec(group: SparklineGroup, data: Array<number | null>, axis?: { min?: number; max?: number }): SparklineSpec {
  const colors = group.colors || {}
  const series = colors.series || '#376092'
  const options: SparklineSpec['options'] = {
    charttype: group.type === 'column' ? 'column' : group.type === 'stacked' ? 'winloss' : 'line',
    color: series,
    negcolor: group.negative ? colors.negative || '#D00000' : series,
  }
  if (group.high) options.highcolor = colors.high || '#D00000'
  if (group.low) options.lowcolor = colors.low || '#D00000'
  if (group.first) options.firstcolor = colors.first || '#D00000'
  if (group.last) options.lastcolor = colors.last || '#D00000'
  if (group.markers && group.type === 'line') { options.markers = true; options.markercolor = colors.markers || '#D00000' }
  if (group.displayXAxis) { options.axis = true; options.axiscolor = colors.axis || '#000000' }
  if (group.rightToLeft) options.rtl = true
  if (group.lineWeight) options.linewidth = group.lineWeight * (4 / 3)
  if (group.displayEmptyCellsAs === 'zero') options.empty = 'zero'
  if (axis?.min !== undefined) options.ymin = axis.min
  if (axis?.max !== undefined) options.ymax = axis.max
  // "Connect data points with line" drops the gaps.
  return { data: group.displayEmptyCellsAs === 'span' ? data.filter((value) => value !== null) : data, options }
}
