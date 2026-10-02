// src/shared/vector.ts (WP3)
// Text and shape drawing shared by Simple markup baking and Advanced text/shape layers (design 5.13), so
// both look identical.
//   - Text: cssFont(), layoutText() (word wrap for paragraph text, alignment, line height, letter spacing),
//     textBounds(), drawText(), rasterizeText() (waits for the font, renders into an OffscreenCanvas over
//     the transformed bounds and reads it back as straight RGBA).
//     Text-local origin = top-left of the first line box; each line box is fontSize * lineHeight tall with
//     the glyphs centred CSS-style (half leading); transform maps text-local space to the document.
//   - Shapes: shapeBounds(), drawShape() (rectangle / rounded rectangle and ellipse with fill + centred
//     stroke; lines and arrows as one filled outline with round caps and heads max(10, 4 * width) long),
//     rasterizeShape() (canvas when available, otherwise an exact pure-JS polygon rasterizer, so shape
//     layers also work in workers and Node tests).
//   - specKey(): a stable cache key. Whole-pixel translations do not change it: moving a text/shape layer
//     by whole pixels only shifts its raster cache.
// No DOM access at module load (Node tests import it); canvases are created only inside the functions.
import type { Affine, MaskBuffer, PixelBuffer, Point, Rect, Rgb8, Rgba8 } from '../imaging/types.ts'
import type { ShapeSpec, TextSpec, TextStyle } from '../advanced/types.ts'
import { LIMITS } from '../advanced/types.ts'
import { createMaskBuffer, rasterizePolygon } from '../imaging/mask.ts'

type Context2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D

/** Curated Windows families offered before queryLocalFonts() runs (design 5.13). */
export const FONT_FAMILIES: readonly string[] = Object.freeze([
  'Segoe UI', 'Arial', 'Bahnschrift', 'Calibri', 'Cambria', 'Candara', 'Consolas', 'Constantia', 'Corbel',
  'Courier New', 'Georgia', 'Impact', 'Segoe Print', 'Segoe Script', 'Sitka Text', 'Tahoma', 'Times New Roman',
  'Trebuchet MS', 'Verdana',
])

export const DEFAULT_TEXT_STYLE: TextStyle = Object.freeze({
  fontFamily: 'Segoe UI',
  fontSize: 48,
  fontWeight: 400,
  italic: false,
  underline: false,
  color: Object.freeze({ r: 0, g: 0, b: 0 }),
  align: 'left',
  lineHeight: 1.2,
  letterSpacing: 0,
})

export const IDENTITY_AFFINE: Affine = Object.freeze([1, 0, 0, 1, 0, 0] as const)

const MIN_FONT_SIZE = 1
const MAX_FONT_SIZE = 5000

export interface RasterResult {
  readonly pixels: PixelBuffer
  /** Document position of pixels (0, 0). */
  readonly offsetX: number
  readonly offsetY: number
}

// ---------------------------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------------------------

function finite(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback
}

function fontSizeOf(style: TextStyle): number {
  return Math.min(MAX_FONT_SIZE, Math.max(MIN_FONT_SIZE, finite(style.fontSize, DEFAULT_TEXT_STYLE.fontSize)))
}

function lineHeightOf(style: TextStyle): number {
  const factor = finite(style.lineHeight, 1.2)
  return fontSizeOf(style) * (factor > 0 ? factor : 1.2)
}

function byte(value: number): number {
  return Math.max(0, Math.min(255, Math.round(finite(value, 0))))
}

function rgbCss(color: Rgb8): string {
  return `rgb(${byte(color.r)}, ${byte(color.g)}, ${byte(color.b)})`
}

function rgbaCss(color: Rgba8): string {
  return `rgba(${byte(color.r)}, ${byte(color.g)}, ${byte(color.b)}, ${byte(color.a) / 255})`
}

function quoteFamily(name: string): string {
  const clean = String(name ?? '').replace(/["\\;{}]/g, '').trim()
  return `"${clean || DEFAULT_TEXT_STYLE.fontFamily}"`
}

function setLetterSpacing(context: Context2D, spacing: number): void {
  const target = context as Context2D & { letterSpacing?: string }
  if ('letterSpacing' in target) target.letterSpacing = `${finite(spacing, 0)}px`
}

function applyAffine(m: Affine, p: Point): Point {
  return { x: m[0] * p.x + m[2] * p.y + m[4], y: m[1] * p.x + m[3] * p.y + m[5] }
}

/** Axis-aligned bounds of `rect` mapped through `m`. */
export function transformRect(rect: Rect, m: Affine): Rect {
  const corners = [
    applyAffine(m, { x: rect.x, y: rect.y }),
    applyAffine(m, { x: rect.x + rect.width, y: rect.y }),
    applyAffine(m, { x: rect.x, y: rect.y + rect.height }),
    applyAffine(m, { x: rect.x + rect.width, y: rect.y + rect.height }),
  ]
  let minX = Infinity
  let minY = Infinity
  let maxX = -Infinity
  let maxY = -Infinity
  for (const corner of corners) {
    if (corner.x < minX) minX = corner.x
    if (corner.y < minY) minY = corner.y
    if (corner.x > maxX) maxX = corner.x
    if (corner.y > maxY) maxY = corner.y
  }
  if (!Number.isFinite(minX + minY + maxX + maxY)) return { x: 0, y: 0, width: 0, height: 0 }
  return { x: minX, y: minY, width: maxX - minX, height: maxY - minY }
}

function affineOf(transform: Affine | undefined | null): Affine {
  if (!transform || transform.length !== 6 || !transform.every(Number.isFinite)) return IDENTITY_AFFINE
  return transform
}

// ---------------------------------------------------------------------------------------------
// Canvas access (lazy)
// ---------------------------------------------------------------------------------------------

type AnyCanvas = OffscreenCanvas | HTMLCanvasElement

function canvasAvailable(): boolean {
  return typeof OffscreenCanvas !== 'undefined' || typeof document !== 'undefined'
}

function createCanvas(width: number, height: number): AnyCanvas | null {
  try {
    if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(width, height)
    if (typeof document !== 'undefined') {
      const canvas = document.createElement('canvas')
      canvas.width = width
      canvas.height = height
      return canvas
    }
  } catch {
    // Fall through: no canvas in this environment.
  }
  return null
}

function contextOf(canvas: AnyCanvas, readback: boolean): Context2D | null {
  const options = readback ? { willReadFrequently: true } : undefined
  return (canvas as OffscreenCanvas).getContext('2d', options) as Context2D | null
}

function release(canvas: AnyCanvas | null): void {
  if (!canvas) return
  canvas.width = 1
  canvas.height = 1
}

let measureContext: Context2D | null | undefined

function getMeasureContext(): Context2D | null {
  if (measureContext !== undefined) return measureContext
  const canvas = createCanvas(1, 1)
  measureContext = canvas ? contextOf(canvas, false) : null
  return measureContext
}

// ---------------------------------------------------------------------------------------------
// Text
// ---------------------------------------------------------------------------------------------

/** CSS font shorthand for canvas: `italic 700 48px "Segoe UI", sans-serif`. */
export function cssFont(style: TextStyle): string {
  const weight = style.fontWeight === 700 ? 700 : 400
  return `${style.italic ? 'italic ' : ''}${weight} ${fontSizeOf(style)}px ${quoteFamily(style.fontFamily)}, sans-serif`
}

export interface TextMetricsLite {
  readonly width: number
  /** Font ascent / descent above / below the baseline (px). */
  readonly ascent: number
  readonly descent: number
}

export type TextMeasure = (text: string, style: TextStyle) => TextMetricsLite

/** Canvas text measurement (with letter spacing); a plain estimate when no canvas exists. */
export function measureText(text: string, style: TextStyle): TextMetricsLite {
  const size = fontSizeOf(style)
  const context = getMeasureContext()
  if (!context) {
    const count = Array.from(text).length
    return { width: count * (size * 0.55 + finite(style.letterSpacing, 0)), ascent: size * 0.8, descent: size * 0.2 }
  }
  context.font = cssFont(style)
  setLetterSpacing(context, style.letterSpacing)
  const metrics = context.measureText(text)
  const ascent = finite(metrics.fontBoundingBoxAscent, finite(metrics.actualBoundingBoxAscent, size * 0.8))
  const descent = finite(metrics.fontBoundingBoxDescent, finite(metrics.actualBoundingBoxDescent, size * 0.2))
  return { width: finite(metrics.width, 0), ascent, descent }
}

export interface TextLine {
  readonly text: string
  /** Left edge of the line (text-local px). */
  readonly x: number
  /** Top of the line box. */
  readonly top: number
  readonly baseline: number
  readonly width: number
}

export interface TextLayout {
  readonly lines: readonly TextLine[]
  /** Layout box: the paragraph width, or the widest line for point text. */
  readonly width: number
  readonly height: number
  readonly lineHeight: number
  readonly ascent: number
  readonly descent: number
}

function breakWord(word: string, maxWidth: number, style: TextStyle, measure: TextMeasure): string[] {
  const parts: string[] = []
  let current = ''
  for (const char of Array.from(word)) {
    const candidate = current + char
    if (current && measure(candidate, style).width > maxWidth) {
      parts.push(current)
      current = char
    } else {
      current = candidate
    }
  }
  if (current) parts.push(current)
  return parts
}

/** Greedy word wrap of one paragraph; words longer than the box break between characters. */
function wrapParagraph(paragraph: string, maxWidth: number, style: TextStyle, measure: TextMeasure): string[] {
  if (!paragraph) return ['']
  const tokens = paragraph.match(/\S+|\s+/g) ?? []
  const lines: string[] = []
  let line = ''
  const fits = (text: string) => measure(text.trimEnd(), style).width <= maxWidth
  for (const token of tokens) {
    if (!/\S/.test(token)) {
      // Leading spaces of the paragraph stay; spaces at the start of a wrapped line are dropped.
      if (line || !lines.length) line += token
      continue
    }
    if (fits(line + token)) {
      line += token
      continue
    }
    if (line.trim()) {
      lines.push(line.trimEnd())
      line = ''
    }
    if (fits(line + token)) {
      line += token
      continue
    }
    // A word wider than the box breaks between characters.
    const pieces = breakWord(line + token, maxWidth, style, measure)
    for (let index = 0; index < pieces.length - 1; index += 1) lines.push(pieces[index])
    line = pieces[pieces.length - 1] ?? ''
  }
  lines.push(line.trimEnd())
  return lines
}

/** Lays text out in text-local space (no transform). `measure` defaults to canvas measurement. */
export function layoutText(spec: Pick<TextSpec, 'text' | 'style' | 'boxWidth'>, measure: TextMeasure = measureText): TextLayout {
  const style = spec.style
  const lineHeight = lineHeightOf(style)
  const reference = measure('Hg', style)
  const ascent = reference.ascent
  const descent = reference.descent
  const box = spec.boxWidth !== null && spec.boxWidth !== undefined && Number.isFinite(spec.boxWidth) && spec.boxWidth > 0 ? spec.boxWidth : null
  const paragraphs = String(spec.text ?? '').split(/\r\n|\r|\n/)
  const texts: string[] = []
  for (const paragraph of paragraphs) {
    if (box === null) texts.push(paragraph)
    else texts.push(...wrapParagraph(paragraph, box, style, measure))
  }
  const widths = texts.map((text) => (text ? measure(text, style).width : 0))
  const width = box ?? widths.reduce((max, value) => (value > max ? value : max), 0)
  const halfLeading = (lineHeight - (ascent + descent)) / 2
  const lines = texts.map((text, index): TextLine => {
    const lineWidth = widths[index]
    const x = style.align === 'center' ? (width - lineWidth) / 2 : style.align === 'right' ? width - lineWidth : 0
    const top = index * lineHeight
    return { text, x, top, baseline: top + halfLeading + ascent, width: lineWidth }
  })
  return { lines, width, height: texts.length * lineHeight, lineHeight, ascent, descent }
}

/** Document-space bounds of the text's layout box. */
export function textBounds(spec: TextSpec): Rect {
  const layout = layoutText(spec)
  return transformRect({ x: 0, y: 0, width: layout.width, height: layout.height }, affineOf(spec.transform))
}

/** Draws the text with its transform on top of the context's current transform. */
export function drawText(context: Context2D, spec: TextSpec): void {
  const layout = layoutText(spec)
  const style = spec.style
  const size = fontSizeOf(style)
  const m = affineOf(spec.transform)
  context.save()
  try {
    context.transform(m[0], m[1], m[2], m[3], m[4], m[5])
    context.font = cssFont(style)
    setLetterSpacing(context, style.letterSpacing)
    context.textBaseline = 'alphabetic'
    context.textAlign = 'left'
    context.fillStyle = rgbCss(style.color)
    for (const line of layout.lines) {
      if (!line.text) continue
      context.fillText(line.text, line.x, line.baseline)
      if (style.underline && line.width > 0) {
        context.fillRect(line.x, line.baseline + Math.max(1, size * 0.08), line.width, Math.max(1, size / 16))
      }
    }
  } finally {
    context.restore()
  }
}

function emptyResult(x: number, y: number): RasterResult {
  return { pixels: { width: 0, height: 0, data: new Uint8ClampedArray(0) }, offsetX: x, offsetY: y }
}

function integerBounds(bounds: Rect): { x0: number; y0: number; width: number; height: number } {
  const x0 = Math.floor(bounds.x)
  const y0 = Math.floor(bounds.y)
  const width = Math.max(0, Math.ceil(bounds.x + bounds.width) - x0)
  const height = Math.max(0, Math.ceil(bounds.y + bounds.height) - y0)
  return { x0, y0, width, height }
}

function checkRasterSize(width: number, height: number): void {
  if (width > LIMITS.maxDimension || height > LIMITS.maxDimension || width * height > LIMITS.maxPixels) {
    throw new RangeError('This text or shape is too large to draw. Make it smaller and try again.')
  }
}

/** Renders `draw` over `bounds` (document space) into straight RGBA pixels through a canvas. */
function renderThroughCanvas(bounds: Rect, draw: (context: Context2D) => void): RasterResult {
  const { x0, y0, width, height } = integerBounds(bounds)
  if (!width || !height) return emptyResult(x0, y0)
  checkRasterSize(width, height)
  const canvas = createCanvas(width, height)
  if (!canvas) throw new Error('Drawing text needs a canvas, which is not available here.')
  try {
    const context = contextOf(canvas, true)
    if (!context) throw new Error('A drawing surface could not be created. Close other large images and try again.')
    context.setTransform(1, 0, 0, 1, -x0, -y0)
    draw(context)
    const image = context.getImageData(0, 0, width, height)
    return { pixels: { width, height, data: image.data }, offsetX: x0, offsetY: y0 }
  } finally {
    release(canvas)
  }
}

function textRasterBounds(spec: TextSpec): Rect {
  const layout = layoutText(spec)
  const size = fontSizeOf(spec.style)
  // Glyphs overhang the layout box (italics, swashes, anti-aliasing); pad generously.
  const pad = Math.ceil(size * 0.35 + Math.abs(finite(spec.style.letterSpacing, 0))) + 2
  return transformRect({ x: -pad, y: -pad, width: layout.width + 2 * pad, height: layout.height + 2 * pad }, affineOf(spec.transform))
}

/** Synchronous rasterization with whatever fonts are available now (system fonts always are). */
export function rasterizeTextSync(spec: TextSpec): RasterResult {
  if (!String(spec.text ?? '').trim() && !spec.style.underline) {
    const origin = applyAffine(affineOf(spec.transform), { x: 0, y: 0 })
    return emptyResult(Math.floor(origin.x), Math.floor(origin.y))
  }
  return renderThroughCanvas(textRasterBounds(spec), (context) => drawText(context, spec))
}

/** Waits for the font (document.fonts / worker fonts), then rasterizes the text over its transformed bounds. */
export async function rasterizeText(spec: TextSpec): Promise<RasterResult> {
  const fonts = (globalThis as { document?: { fonts?: FontFaceSet }; fonts?: FontFaceSet }).document?.fonts
    ?? (globalThis as { fonts?: FontFaceSet }).fonts
  if (fonts && typeof fonts.load === 'function') {
    try {
      await fonts.load(cssFont(spec.style), spec.text || 'A')
    } catch {
      // A missing font falls back to the next family; rendering still works.
    }
  }
  return rasterizeTextSync(spec)
}

// ---------------------------------------------------------------------------------------------
// Shapes: geometry
// ---------------------------------------------------------------------------------------------

/** Arrow head length for a stroke width: max(10, 4 * width). */
export function arrowHeadLength(width: number): number {
  return Math.max(10, 4 * Math.max(0, finite(width, 0)))
}

const HEAD_HALF_WIDTH = 0.45

function lineColor(spec: ShapeSpec): Rgba8 | null {
  const color = spec.stroke ?? spec.fill
  return color && color.a > 0 ? color : null
}

function lineWidth(spec: ShapeSpec): number {
  return Math.max(0.5, finite(spec.strokeWidth, 1))
}

function boxOf(spec: ShapeSpec): Rect {
  const x = Math.min(spec.x1, spec.x2)
  const y = Math.min(spec.y1, spec.y2)
  return { x, y, width: Math.abs(spec.x2 - spec.x1), height: Math.abs(spec.y2 - spec.y1) }
}

function strokeWidthOf(spec: ShapeSpec): number {
  return spec.stroke && spec.stroke.a > 0 && spec.strokeWidth > 0 ? finite(spec.strokeWidth, 0) : 0
}

/** Shape-local bounds including the stroke and arrow heads. */
function localShapeBounds(spec: ShapeSpec): Rect {
  if (spec.kind === 'rectangle' || spec.kind === 'ellipse') {
    const box = boxOf(spec)
    const half = strokeWidthOf(spec) / 2
    return { x: box.x - half, y: box.y - half, width: box.width + 2 * half, height: box.height + 2 * half }
  }
  const width = lineWidth(spec)
  const heads = spec.kind === 'arrow' && spec.arrowHeads !== 'none'
  const pad = Math.max(width / 2, heads ? arrowHeadLength(width) : 0)
  const x = Math.min(spec.x1, spec.x2)
  const y = Math.min(spec.y1, spec.y2)
  return { x: x - pad, y: y - pad, width: Math.abs(spec.x2 - spec.x1) + 2 * pad, height: Math.abs(spec.y2 - spec.y1) + 2 * pad }
}

/** Document-space bounds of the drawn shape (stroke and arrow heads included). */
export function shapeBounds(spec: ShapeSpec): Rect {
  return transformRect(localShapeBounds(spec), affineOf(spec.transform))
}

function segmentsFor(radius: number): number {
  return Math.max(8, Math.min(256, Math.ceil(Math.sqrt(Math.max(0, radius)) * 6)))
}

function signedArea(points: readonly Point[]): number {
  let area = 0
  for (let i = 0, j = points.length - 1; i < points.length; j = i, i += 1) {
    area += (points[j].x * points[i].y) - (points[i].x * points[j].y)
  }
  return area / 2
}

/** Same orientation for every outline so overlapping parts of one filled path never cancel out. */
function oriented(points: Point[]): Point[] {
  return signedArea(points) < 0 ? points.reverse() : points
}

function ellipsePoints(cx: number, cy: number, rx: number, ry: number): Point[] {
  const count = segmentsFor(Math.max(rx, ry)) * 2
  const points: Point[] = []
  for (let i = 0; i < count; i += 1) {
    const angle = (i / count) * Math.PI * 2
    points.push({ x: cx + rx * Math.cos(angle), y: cy + ry * Math.sin(angle) })
  }
  return points
}

function roundRectPoints(x: number, y: number, width: number, height: number, radius: number): Point[] {
  const r = Math.max(0, Math.min(radius, width / 2, height / 2))
  if (r <= 0) return [{ x, y }, { x: x + width, y }, { x: x + width, y: y + height }, { x, y: y + height }]
  const count = Math.max(2, Math.ceil(segmentsFor(r) / 2))
  const points: Point[] = []
  const corners = [
    { cx: x + width - r, cy: y + r, start: -Math.PI / 2 },
    { cx: x + width - r, cy: y + height - r, start: 0 },
    { cx: x + r, cy: y + height - r, start: Math.PI / 2 },
    { cx: x + r, cy: y + r, start: Math.PI },
  ]
  for (const corner of corners) {
    for (let i = 0; i <= count; i += 1) {
      const angle = corner.start + (i / count) * (Math.PI / 2)
      points.push({ x: corner.cx + r * Math.cos(angle), y: corner.cy + r * Math.sin(angle) })
    }
  }
  return points
}

/** A line of width `width` from a to b with round caps (a circle when a == b). */
function capsulePoints(a: Point, b: Point, width: number): Point[] {
  const r = width / 2
  const dx = b.x - a.x
  const dy = b.y - a.y
  const length = Math.hypot(dx, dy)
  if (length < 1e-9) return ellipsePoints(a.x, a.y, r, r)
  const angle = Math.atan2(dy, dx)
  const count = Math.max(4, Math.ceil(segmentsFor(r) / 2))
  const points: Point[] = []
  for (let i = 0; i <= count; i += 1) {
    const t = angle - Math.PI / 2 + (i / count) * Math.PI
    points.push({ x: b.x + r * Math.cos(t), y: b.y + r * Math.sin(t) })
  }
  for (let i = 0; i <= count; i += 1) {
    const t = angle + Math.PI / 2 + (i / count) * Math.PI
    points.push({ x: a.x + r * Math.cos(t), y: a.y + r * Math.sin(t) })
  }
  return points
}

function headPoints(tip: Point, from: Point, length: number): Point[] {
  const dx = tip.x - from.x
  const dy = tip.y - from.y
  const distance = Math.hypot(dx, dy) || 1
  const ux = dx / distance
  const uy = dy / distance
  const baseX = tip.x - ux * length
  const baseY = tip.y - uy * length
  const half = length * HEAD_HALF_WIDTH
  return [tip, { x: baseX - uy * half, y: baseY + ux * half }, { x: baseX + uy * half, y: baseY - ux * half }]
}

/** Outlines (shape-local) of a line or arrow: one capsule for the shaft plus a triangle per head. */
export function lineOutlines(spec: ShapeSpec): Point[][] {
  const width = lineWidth(spec)
  let start: Point = { x: spec.x1, y: spec.y1 }
  let end: Point = { x: spec.x2, y: spec.y2 }
  const outlines: Point[][] = []
  if (spec.kind === 'arrow' && spec.arrowHeads !== 'none') {
    const length = Math.hypot(end.x - start.x, end.y - start.y)
    const head = arrowHeadLength(width)
    // Stop the shaft inside the head so its round cap stays hidden under it.
    const inset = Math.min(head * 0.5, length / 2)
    const ux = length > 0 ? (end.x - start.x) / length : 0
    const uy = length > 0 ? (end.y - start.y) / length : 0
    outlines.push(oriented(headPoints(end, start, head)))
    const shaftEnd = { x: end.x - ux * inset, y: end.y - uy * inset }
    if (spec.arrowHeads === 'both') {
      outlines.push(oriented(headPoints(start, end, head)))
      start = { x: start.x + ux * inset, y: start.y + uy * inset }
    }
    end = shaftEnd
  }
  outlines.unshift(oriented(capsulePoints(start, end, width)))
  return outlines
}

// ---------------------------------------------------------------------------------------------
// Shapes: drawing
// ---------------------------------------------------------------------------------------------

/** Draws the shape with its transform on top of the context's current transform. */
export function drawShape(context: Context2D, spec: ShapeSpec): void {
  const m = affineOf(spec.transform)
  context.save()
  try {
    context.transform(m[0], m[1], m[2], m[3], m[4], m[5])
    if (spec.kind === 'rectangle' || spec.kind === 'ellipse') {
      const box = boxOf(spec)
      context.beginPath()
      if (spec.kind === 'ellipse') {
        context.ellipse(box.x + box.width / 2, box.y + box.height / 2, box.width / 2, box.height / 2, 0, 0, Math.PI * 2)
      } else {
        const radius = Math.max(0, Math.min(finite(spec.cornerRadius, 0), box.width / 2, box.height / 2))
        if (radius > 0 && typeof (context as CanvasRenderingContext2D).roundRect === 'function') context.roundRect(box.x, box.y, box.width, box.height, radius)
        else context.rect(box.x, box.y, box.width, box.height)
      }
      if (spec.fill && spec.fill.a > 0) {
        context.fillStyle = rgbaCss(spec.fill)
        context.fill()
      }
      const stroke = strokeWidthOf(spec)
      if (stroke > 0 && spec.stroke) {
        context.lineWidth = stroke
        context.lineJoin = spec.kind === 'rectangle' && finite(spec.cornerRadius, 0) <= 0 ? 'miter' : 'round'
        context.strokeStyle = rgbaCss(spec.stroke)
        context.stroke()
      }
      return
    }
    const color = lineColor(spec)
    if (!color) return
    // One path, one fill: overlapping shaft and heads are painted once even with a translucent colour.
    context.beginPath()
    for (const outline of lineOutlines(spec)) {
      context.moveTo(outline[0].x, outline[0].y)
      for (let i = 1; i < outline.length; i += 1) context.lineTo(outline[i].x, outline[i].y)
      context.closePath()
    }
    context.fillStyle = rgbaCss(color)
    context.fill('nonzero')
  } finally {
    context.restore()
  }
}

// ---------------------------------------------------------------------------------------------
// Shapes: pure rasterizer (no canvas)
// ---------------------------------------------------------------------------------------------

function polygonCoverage(target: MaskBuffer, outlines: readonly Point[][], m: Affine, x0: number, y0: number): MaskBuffer {
  for (const outline of outlines) {
    const points = outline.map((point) => {
      const p = applyAffine(m, point)
      return { x: p.x - x0, y: p.y - y0 }
    })
    rasterizePolygon(target, points, true, 'nonzero')
  }
  return target
}

function paintCoverage(out: Uint8ClampedArray, coverage: Uint8Array, color: Rgba8): void {
  const cr = byte(color.r)
  const cg = byte(color.g)
  const cb = byte(color.b)
  const ca = byte(color.a) / 255
  for (let i = 0, p = 0; i < coverage.length; i += 1, p += 4) {
    const c = coverage[i]
    if (c === 0) continue
    const as = ca * (c / 255)
    const ab = out[p + 3] / 255
    const ao = as + ab * (1 - as)
    if (ao <= 0) continue
    out[p] = (cr * as + out[p] * ab * (1 - as)) / ao
    out[p + 1] = (cg * as + out[p + 1] * ab * (1 - as)) / ao
    out[p + 2] = (cb * as + out[p + 2] * ab * (1 - as)) / ao
    out[p + 3] = ao * 255
  }
}

function rasterizeShapeWithoutCanvas(spec: ShapeSpec, bounds: Rect): RasterResult {
  const { x0, y0, width, height } = integerBounds(bounds)
  if (!width || !height) return emptyResult(x0, y0)
  checkRasterSize(width, height)
  const m = affineOf(spec.transform)
  const data = new Uint8ClampedArray(width * height * 4)
  if (spec.kind === 'rectangle' || spec.kind === 'ellipse') {
    const box = boxOf(spec)
    const radius = finite(spec.cornerRadius, 0)
    const outline = (inset: number): Point[] | null => {
      const w = box.width - 2 * inset
      const h = box.height - 2 * inset
      if (!(w > 0 && h > 0)) return null
      if (spec.kind === 'ellipse') return ellipsePoints(box.x + box.width / 2, box.y + box.height / 2, w / 2, h / 2)
      return roundRectPoints(box.x + inset, box.y + inset, w, h, radius > 0 ? Math.max(0, radius - inset) : 0)
    }
    if (spec.fill && spec.fill.a > 0) {
      const body = outline(0)
      if (body) paintCoverage(data, polygonCoverage(createMaskBuffer(width, height), [body], m, x0, y0).data, spec.fill)
    }
    const stroke = strokeWidthOf(spec)
    if (stroke > 0 && spec.stroke) {
      const outer = outline(-stroke / 2)
      const inner = outline(stroke / 2)
      if (outer) {
        const ring = polygonCoverage(createMaskBuffer(width, height), [outer], m, x0, y0)
        if (inner) {
          const hole = polygonCoverage(createMaskBuffer(width, height), [inner], m, x0, y0)
          for (let i = 0; i < ring.data.length; i += 1) ring.data[i] = Math.max(0, ring.data[i] - hole.data[i])
        }
        paintCoverage(data, ring.data, spec.stroke)
      }
    }
  } else {
    const color = lineColor(spec)
    if (color) paintCoverage(data, polygonCoverage(createMaskBuffer(width, height), lineOutlines(spec), m, x0, y0).data, color)
  }
  return { pixels: { width, height, data }, offsetX: x0, offsetY: y0 }
}

/** Rasterizes a shape over its transformed bounds (canvas when available, else the pure rasterizer). */
export function rasterizeShape(spec: ShapeSpec): RasterResult {
  const bounds = shapeBounds(spec)
  // One extra pixel each side for anti-aliasing.
  const padded = { x: bounds.x - 1, y: bounds.y - 1, width: bounds.width + 2, height: bounds.height + 2 }
  if (canvasAvailable()) {
    try {
      return renderThroughCanvas(padded, (context) => drawShape(context, spec))
    } catch (error) {
      if (error instanceof RangeError) throw error
      // No usable canvas (e.g. a worker without OffscreenCanvas 2D): use the pure rasterizer.
    }
  }
  return rasterizeShapeWithoutCanvas(spec, padded)
}

/** The pure rasterizer, exposed for tests and canvas-less contexts. */
export function rasterizeShapePure(spec: ShapeSpec): RasterResult {
  const bounds = shapeBounds(spec)
  return rasterizeShapeWithoutCanvas(spec, { x: bounds.x - 1, y: bounds.y - 1, width: bounds.width + 2, height: bounds.height + 2 })
}

// ---------------------------------------------------------------------------------------------
// Cache keys
// ---------------------------------------------------------------------------------------------

function r6(value: number): number {
  return Math.round(finite(value, 0) * 1e6) / 1e6
}

/** Transform with the whole-pixel part of the translation removed. */
function keyTransform(transform: Affine): number[] {
  const m = affineOf(transform)
  return [r6(m[0]), r6(m[1]), r6(m[2]), r6(m[3]), r6(m[4] - Math.floor(m[4])), r6(m[5] - Math.floor(m[5]))]
}

function colorKey(color: Rgb8 | Rgba8 | null | undefined): string {
  if (!color) return '-'
  return `${byte(color.r)},${byte(color.g)},${byte(color.b)}${'a' in color ? `,${byte(color.a)}` : ''}`
}

export function isTextSpec(spec: TextSpec | ShapeSpec): spec is TextSpec {
  return typeof (spec as TextSpec).text === 'string' && typeof (spec as TextSpec).style === 'object'
}

/**
 * Stable key of everything that affects a text/shape raster except whole-pixel translation. A raster
 * cache is current when its specKey equals specKey(spec).
 */
export function specKey(spec: TextSpec | ShapeSpec): string {
  if (isTextSpec(spec)) {
    const s = spec.style
    return JSON.stringify(['text', spec.text, String(s.fontFamily), r6(s.fontSize), s.fontWeight === 700 ? 700 : 400,
      Boolean(s.italic), Boolean(s.underline), colorKey(s.color), s.align, r6(s.lineHeight), r6(s.letterSpacing),
      spec.boxWidth === null || spec.boxWidth === undefined ? null : r6(spec.boxWidth), keyTransform(spec.transform)])
  }
  return JSON.stringify(['shape', spec.kind, r6(spec.x1), r6(spec.y1), r6(spec.x2), r6(spec.y2), colorKey(spec.fill),
    colorKey(spec.stroke), r6(spec.strokeWidth), r6(spec.cornerRadius), spec.arrowHeads, keyTransform(spec.transform)])
}

/**
 * Whole-pixel translation between two specs with equal keys (b's position minus a's), or null when they
 * differ by more than that.
 */
export function wholePixelShift(a: TextSpec | ShapeSpec, b: TextSpec | ShapeSpec): Point | null {
  if (specKey(a) !== specKey(b)) return null
  const ma = affineOf(a.transform)
  const mb = affineOf(b.transform)
  const dx = Math.round(mb[4] - ma[4])
  const dy = Math.round(mb[5] - ma[5])
  return { x: dx, y: dy }
}
