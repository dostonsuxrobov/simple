// src/simple/markupModel.ts (WP8)
// Simple-mode markup (design 4.6): text, arrows, rectangles, ellipses, lines and pasted images that stay
// editable until they are baked into the image. Everything here is pure and DOM-free (Node-tested):
// item creation with Shift constraints (45 degree lines, squares), bounds, hit-testing of items and their
// handles, moving and resizing (pasted images keep their aspect ratio), the session's own undo stack, and
// the conversion to src/shared/vector.ts specs so baked markup looks exactly like Advanced shape/text layers.
// Coordinates are image pixels.
import type { Point, Rect, Rgb8, Rgba8 } from '../imaging/types.ts'
import type { ShapeSpec, TextSpec, TextStyle } from '../advanced/types.ts'
import type { TextMeasure } from '../shared/vector.ts'
import { arrowHeadLength, layoutText } from '../shared/vector.ts'

export type MarkupShapeKind = 'arrow' | 'rectangle' | 'ellipse' | 'line'
export type MarkupKind = 'text' | MarkupShapeKind | 'image'
/** What a click on empty canvas creates. */
export type MarkupCreateKind = 'text' | MarkupShapeKind
export type BoxHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'
export type MarkupHandle = 'start' | 'end' | BoxHandle

export const MARKUP_CREATE_KINDS: readonly MarkupCreateKind[] = Object.freeze(['text', 'arrow', 'rectangle', 'ellipse', 'line'] as const)
export const MIN_STROKE_WIDTH = 1
export const MAX_STROKE_WIDTH = 64
export const MIN_FONT_SIZE = 6
export const MAX_FONT_SIZE = 600
/** Drags shorter than this (image px) do not create a shape. */
export const MIN_SHAPE_DRAG = 3
const MAX_HISTORY = 100

export interface MarkupStyle {
  /** Stroke / text colour, #rrggbb. */
  readonly color: string
  /** Stroke width in image px (1..64). */
  readonly width: number
  /** Rectangle / ellipse fill, #rrggbb, or null. */
  readonly fill: string | null
  readonly fontFamily: string
  /** Image px. */
  readonly fontSize: number
  readonly bold: boolean
}

interface ItemBase {
  readonly id: string
}

export interface ShapeItem extends ItemBase {
  readonly kind: MarkupShapeKind
  /** Line/arrow: start -> end. Rectangle/ellipse: two opposite corners (any order). */
  readonly x1: number
  readonly y1: number
  readonly x2: number
  readonly y2: number
  readonly color: string
  readonly width: number
  readonly fill: string | null
}

export interface TextItem extends ItemBase {
  readonly kind: 'text'
  /** Top-left of the first line box. */
  readonly x: number
  readonly y: number
  readonly text: string
  readonly color: string
  readonly fontFamily: string
  readonly fontSize: number
  readonly bold: boolean
}

export interface ImageItem extends ItemBase {
  readonly kind: 'image'
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  /** Object URL (or data URL) of the pasted picture. */
  readonly src: string
  readonly naturalWidth: number
  readonly naturalHeight: number
}

export type MarkupItem = ShapeItem | TextItem | ImageItem

export const DEFAULT_MARKUP_STYLE: MarkupStyle = Object.freeze({
  color: '#e5222b',
  width: 6,
  fill: null,
  fontFamily: 'Segoe UI',
  fontSize: 48,
  bold: true,
})

let idCounter = 0
/** Unique id for a new item. */
export function nextMarkupId(): string {
  idCounter += 1
  return `markup-${idCounter}`
}

function clamp(value: number, minimum: number, maximum: number): number {
  return value < minimum ? minimum : value > maximum ? maximum : value
}

function finite(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback
}

export function isShape(item: MarkupItem): item is ShapeItem {
  return item.kind === 'arrow' || item.kind === 'rectangle' || item.kind === 'ellipse' || item.kind === 'line'
}

export function isLineKind(kind: MarkupKind): kind is 'arrow' | 'line' {
  return kind === 'arrow' || kind === 'line'
}

/** '#rgb' / '#rrggbb' to 0..255 channels (black for anything else). */
export function hexToRgb(hex: string): Rgb8 {
  const text = String(hex ?? '').trim().replace(/^#/, '')
  const full = /^[0-9a-f]{3}$/i.test(text) ? text.split('').map((c) => c + c).join('') : text
  if (!/^[0-9a-f]{6}$/i.test(full)) return { r: 0, g: 0, b: 0 }
  return { r: parseInt(full.slice(0, 2), 16), g: parseInt(full.slice(2, 4), 16), b: parseInt(full.slice(4, 6), 16) }
}

function rgba(hex: string): Rgba8 {
  return { ...hexToRgb(hex), a: 255 }
}

export function normalizeStyle(style: MarkupStyle): MarkupStyle {
  return {
    color: style.color,
    width: clamp(Math.round(finite(style.width, DEFAULT_MARKUP_STYLE.width)), MIN_STROKE_WIDTH, MAX_STROKE_WIDTH),
    fill: style.fill,
    fontFamily: style.fontFamily || DEFAULT_MARKUP_STYLE.fontFamily,
    fontSize: clamp(Math.round(finite(style.fontSize, DEFAULT_MARKUP_STYLE.fontSize)), MIN_FONT_SIZE, MAX_FONT_SIZE),
    bold: Boolean(style.bold),
  }
}

/** Snaps `end` so the segment from `start` runs at a multiple of 45 degrees (length kept). */
export function snapAngle(start: Point, end: Point): Point {
  const dx = end.x - start.x
  const dy = end.y - start.y
  const length = Math.hypot(dx, dy)
  if (length < 1e-9) return { x: end.x, y: end.y }
  const step = Math.PI / 4
  const angle = Math.round(Math.atan2(dy, dx) / step) * step
  const x = start.x + Math.cos(angle) * length
  const y = start.y + Math.sin(angle) * length
  // Exact axes: avoid -0 and 1e-15 residue from cos/sin.
  return { x: Math.abs(x - start.x) < 1e-9 ? start.x : x, y: Math.abs(y - start.y) < 1e-9 ? start.y : y }
}

/** Makes the box from `start` to `end` square (the larger side wins, direction kept). */
export function squareCorner(start: Point, end: Point): Point {
  const dx = end.x - start.x
  const dy = end.y - start.y
  const side = Math.max(Math.abs(dx), Math.abs(dy))
  return { x: start.x + (dx < 0 ? -side : side), y: start.y + (dy < 0 ? -side : side) }
}

/** A new shape dragged from `start` to `end`; `constrain` (Shift) gives 45 degree lines and squares/circles. */
export function createShapeItem(id: string, kind: MarkupShapeKind, start: Point, end: Point, style: MarkupStyle, constrain: boolean): ShapeItem {
  const s = normalizeStyle(style)
  const target = constrain ? (isLineKind(kind) ? snapAngle(start, end) : squareCorner(start, end)) : end
  return {
    id,
    kind,
    x1: start.x,
    y1: start.y,
    x2: target.x,
    y2: target.y,
    color: s.color,
    width: s.width,
    fill: kind === 'rectangle' || kind === 'ellipse' ? s.fill : null,
  }
}

export function createTextItem(id: string, at: Point, style: MarkupStyle, text = ''): TextItem {
  const s = normalizeStyle(style)
  return { id, kind: 'text', x: at.x, y: at.y, text, color: s.color, fontFamily: s.fontFamily, fontSize: s.fontSize, bold: s.bold }
}

/** A pasted picture centred in the image, at most `maxFraction` of its width and height, aspect kept. */
export function createImageItem(id: string, src: string, naturalWidth: number, naturalHeight: number, image: { width: number; height: number }, maxFraction = 0.6): ImageItem {
  const nw = Math.max(1, naturalWidth)
  const nh = Math.max(1, naturalHeight)
  const scale = Math.min(1, (image.width * maxFraction) / nw, (image.height * maxFraction) / nh)
  const width = Math.max(1, nw * scale)
  const height = Math.max(1, nh * scale)
  return { id, kind: 'image', x: (image.width - width) / 2, y: (image.height - height) / 2, width, height, src, naturalWidth: nw, naturalHeight: nh }
}

/** True when a shape drag is long enough to keep. */
export function isMeaningfulShape(item: ShapeItem): boolean {
  if (isLineKind(item.kind)) return Math.hypot(item.x2 - item.x1, item.y2 - item.y1) >= MIN_SHAPE_DRAG
  return Math.abs(item.x2 - item.x1) >= MIN_SHAPE_DRAG && Math.abs(item.y2 - item.y1) >= MIN_SHAPE_DRAG
}

// ---------------------------------------------------------------------------------------------
// Specs for drawing (src/shared/vector.ts)
// ---------------------------------------------------------------------------------------------

export function textStyleOf(item: TextItem): TextStyle {
  return {
    fontFamily: item.fontFamily,
    fontSize: item.fontSize,
    fontWeight: item.bold ? 700 : 400,
    italic: false,
    underline: false,
    color: hexToRgb(item.color),
    align: 'left',
    lineHeight: 1.2,
    letterSpacing: 0,
  }
}

export function toTextSpec(item: TextItem): TextSpec {
  return { text: item.text, style: textStyleOf(item), boxWidth: null, transform: [1, 0, 0, 1, item.x, item.y] }
}

export function toShapeSpec(item: ShapeItem): ShapeSpec {
  const box = item.kind === 'rectangle' || item.kind === 'ellipse'
  return {
    kind: item.kind,
    x1: item.x1,
    y1: item.y1,
    x2: item.x2,
    y2: item.y2,
    fill: box && item.fill ? rgba(item.fill) : null,
    stroke: rgba(item.color),
    strokeWidth: item.width,
    cornerRadius: 0,
    arrowHeads: item.kind === 'arrow' ? 'end' : 'none',
    transform: [1, 0, 0, 1, 0, 0],
  }
}

// ---------------------------------------------------------------------------------------------
// Geometry
// ---------------------------------------------------------------------------------------------

export function normalizedBox(item: ShapeItem | ImageItem): Rect {
  if (item.kind === 'image') return { x: item.x, y: item.y, width: item.width, height: item.height }
  const x = Math.min(item.x1, item.x2)
  const y = Math.min(item.y1, item.y2)
  return { x, y, width: Math.abs(item.x2 - item.x1), height: Math.abs(item.y2 - item.y1) }
}

export interface TextBox {
  readonly rect: Rect
  readonly lines: readonly { readonly text: string; readonly x: number; readonly baseline: number; readonly width: number }[]
}

/** Layout of a text item in image px (lines carry absolute baselines). */
export function textBox(item: TextItem, measure?: TextMeasure): TextBox {
  const layout = layoutText({ text: item.text || ' ', style: textStyleOf(item), boxWidth: null }, measure)
  return {
    rect: { x: item.x, y: item.y, width: Math.max(layout.width, item.fontSize * 0.3), height: layout.height },
    lines: layout.lines.map((line) => ({ text: line.text, x: item.x + line.x, baseline: item.y + line.baseline, width: line.width })),
  }
}

/** Bounds of what the item paints (stroke, arrow heads and glyph overhang included). */
export function itemBounds(item: MarkupItem, measure?: TextMeasure): Rect {
  if (item.kind === 'text') {
    const rect = textBox(item, measure).rect
    const pad = Math.ceil(item.fontSize * 0.35) + 2
    return { x: rect.x - pad, y: rect.y - pad, width: rect.width + 2 * pad, height: rect.height + 2 * pad }
  }
  if (item.kind === 'image') return normalizedBox(item)
  const half = item.width / 2
  const pad = item.kind === 'arrow' ? Math.max(half, arrowHeadLength(item.width)) : half
  const x = Math.min(item.x1, item.x2) - pad - 1
  const y = Math.min(item.y1, item.y2) - pad - 1
  return { x, y, width: Math.abs(item.x2 - item.x1) + 2 * pad + 2, height: Math.abs(item.y2 - item.y1) + 2 * pad + 2 }
}

/** Union of the items' bounds, or null for no items. */
export function markupBounds(items: readonly MarkupItem[], measure?: TextMeasure): Rect | null {
  let result: Rect | null = null
  for (const item of items) {
    const b = itemBounds(item, measure)
    if (!result) {
      result = b
      continue
    }
    const x = Math.min(result.x, b.x)
    const y = Math.min(result.y, b.y)
    result = { x, y, width: Math.max(result.x + result.width, b.x + b.width) - x, height: Math.max(result.y + result.height, b.y + b.height) - y }
  }
  return result
}

export function distanceToSegment(p: Point, a: Point, b: Point): number {
  const dx = b.x - a.x
  const dy = b.y - a.y
  const lengthSquared = dx * dx + dy * dy
  const t = lengthSquared > 0 ? clamp(((p.x - a.x) * dx + (p.y - a.y) * dy) / lengthSquared, 0, 1) : 0
  return Math.hypot(p.x - (a.x + t * dx), p.y - (a.y + t * dy))
}

function insideRect(p: Point, rect: Rect, pad = 0): boolean {
  return p.x >= rect.x - pad && p.x <= rect.x + rect.width + pad && p.y >= rect.y - pad && p.y <= rect.y + rect.height + pad
}

/** Whether `point` hits the item (strokes within half their width plus `tolerance`; filled interiors). */
export function hitItem(item: MarkupItem, point: Point, tolerance: number, measure?: TextMeasure): boolean {
  if (item.kind === 'text') return insideRect(point, textBox(item, measure).rect, tolerance)
  if (item.kind === 'image') return insideRect(point, normalizedBox(item), tolerance)
  const reach = item.width / 2 + tolerance
  if (isLineKind(item.kind)) {
    const a = { x: item.x1, y: item.y1 }
    const b = { x: item.x2, y: item.y2 }
    if (distanceToSegment(point, a, b) <= reach) return true
    if (item.kind === 'arrow') return Math.hypot(point.x - b.x, point.y - b.y) <= arrowHeadLength(item.width) * 0.75 + tolerance
    return false
  }
  const box = normalizedBox(item)
  if (item.kind === 'rectangle') {
    if (item.fill) return insideRect(point, box, reach)
    if (!insideRect(point, box, reach)) return false
    const inner = { x: box.x + reach, y: box.y + reach, width: box.width - 2 * reach, height: box.height - 2 * reach }
    return !(inner.width > 0 && inner.height > 0 && insideRect(point, inner))
  }
  const rx = box.width / 2
  const ry = box.height / 2
  const cx = box.x + rx
  const cy = box.y + ry
  if (rx < 1e-6 || ry < 1e-6) return distanceToSegment(point, { x: box.x, y: box.y }, { x: box.x + box.width, y: box.y + box.height }) <= reach
  const nx = (point.x - cx) / rx
  const ny = (point.y - cy) / ry
  const radial = Math.hypot(nx, ny)
  if (item.fill && radial <= 1) return true
  // Distance to the outline along the ray, approximated in the ellipse's own scale.
  const along = Math.abs(radial - 1) * (radial > 0 ? Math.hypot(nx * rx, ny * ry) / radial : Math.min(rx, ry))
  return along <= reach
}

/** The topmost item under `point`, or null. */
export function hitTest(items: readonly MarkupItem[], point: Point, tolerance: number, measure?: TextMeasure): MarkupItem | null {
  for (let index = items.length - 1; index >= 0; index -= 1) {
    if (hitItem(items[index], point, tolerance, measure)) return items[index]
  }
  return null
}

const BOX_HANDLES: readonly BoxHandle[] = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']
const CORNER_HANDLES: readonly BoxHandle[] = ['nw', 'ne', 'se', 'sw']

function boxHandlePoint(box: Rect, handle: BoxHandle): Point {
  const x = handle.includes('w') ? box.x : handle.includes('e') ? box.x + box.width : box.x + box.width / 2
  const y = handle.includes('n') ? box.y : handle.includes('s') ? box.y + box.height : box.y + box.height / 2
  return { x, y }
}

/** Resize handles of an item (text has none: its size is the font size). */
export function handlesOf(item: MarkupItem): { readonly handle: MarkupHandle; readonly point: Point }[] {
  if (item.kind === 'text') return []
  if (item.kind !== 'image' && isLineKind(item.kind)) {
    return [{ handle: 'start', point: { x: item.x1, y: item.y1 } }, { handle: 'end', point: { x: item.x2, y: item.y2 } }]
  }
  const box = normalizedBox(item)
  const handles = item.kind === 'image' ? CORNER_HANDLES : BOX_HANDLES
  return handles.map((handle) => ({ handle, point: boxHandlePoint(box, handle) }))
}

/** The handle of `item` under `point` (within `tolerance`), or null. */
export function hitHandle(item: MarkupItem, point: Point, tolerance: number): MarkupHandle | null {
  let best: MarkupHandle | null = null
  let bestDistance = Infinity
  for (const { handle, point: at } of handlesOf(item)) {
    const distance = Math.hypot(point.x - at.x, point.y - at.y)
    if (distance <= tolerance && distance < bestDistance) {
      best = handle
      bestDistance = distance
    }
  }
  return best
}

export function moveItem(item: MarkupItem, dx: number, dy: number): MarkupItem {
  if (item.kind === 'text' || item.kind === 'image') return { ...item, x: item.x + dx, y: item.y + dy }
  return { ...item, x1: item.x1 + dx, y1: item.y1 + dy, x2: item.x2 + dx, y2: item.y2 + dy }
}

const MIN_BOX = 2

/**
 * `start` is the item when the drag began; `point` is the pointer now. Lines move one end (Shift: 45
 * degrees); rectangles and ellipses move edges (Shift: square); pictures always keep their aspect ratio.
 */
export function resizeItem(start: MarkupItem, handle: MarkupHandle, point: Point, constrain: boolean): MarkupItem {
  if (start.kind === 'text') return start
  if (start.kind !== 'image' && isLineKind(start.kind)) {
    if (handle === 'start') {
      const p = constrain ? snapAngle({ x: start.x2, y: start.y2 }, point) : point
      return { ...start, x1: p.x, y1: p.y }
    }
    if (handle === 'end') {
      const p = constrain ? snapAngle({ x: start.x1, y: start.y1 }, point) : point
      return { ...start, x2: p.x, y2: p.y }
    }
    return start
  }
  if (handle === 'start' || handle === 'end') return start
  const box = normalizedBox(start)
  let left = box.x
  let top = box.y
  let right = box.x + box.width
  let bottom = box.y + box.height
  if (handle.includes('w')) left = Math.min(point.x, right - MIN_BOX)
  if (handle.includes('e')) right = Math.max(point.x, left + MIN_BOX)
  if (handle.includes('n')) top = Math.min(point.y, bottom - MIN_BOX)
  if (handle.includes('s')) bottom = Math.max(point.y, top + MIN_BOX)
  const corner = (handle.includes('w') || handle.includes('e')) && (handle.includes('n') || handle.includes('s'))
  const lockRatio = start.kind === 'image' ? start.naturalWidth / start.naturalHeight : constrain && corner ? 1 : null
  if (lockRatio && corner) {
    let width = right - left
    let height = bottom - top
    if (width / height > lockRatio) height = width / lockRatio
    else width = height * lockRatio
    if (handle.includes('w')) left = right - width
    else right = left + width
    if (handle.includes('n')) top = bottom - height
    else bottom = top + height
  }
  if (start.kind === 'image') return { ...start, x: left, y: top, width: right - left, height: bottom - top }
  // Keep the drawing direction of the original corners.
  const flipX = start.x2 < start.x1
  const flipY = start.y2 < start.y1
  return {
    ...start,
    x1: flipX ? right : left,
    x2: flipX ? left : right,
    y1: flipY ? bottom : top,
    y2: flipY ? top : bottom,
  }
}

/** Applies the toolbar style to an existing item (only the properties that item has). */
export function restyleItem(item: MarkupItem, style: Partial<MarkupStyle>): MarkupItem {
  if (item.kind === 'image') return item
  if (item.kind === 'text') {
    return {
      ...item,
      color: style.color ?? item.color,
      fontFamily: style.fontFamily ?? item.fontFamily,
      fontSize: style.fontSize !== undefined ? clamp(Math.round(style.fontSize), MIN_FONT_SIZE, MAX_FONT_SIZE) : item.fontSize,
      bold: style.bold ?? item.bold,
    }
  }
  return {
    ...item,
    color: style.color ?? item.color,
    width: style.width !== undefined ? clamp(Math.round(style.width), MIN_STROKE_WIDTH, MAX_STROKE_WIDTH) : item.width,
    fill: item.kind === 'rectangle' || item.kind === 'ellipse' ? (style.fill !== undefined ? style.fill : item.fill) : null,
  }
}

export function replaceItem(items: readonly MarkupItem[], next: MarkupItem): MarkupItem[] {
  return items.map((item) => (item.id === next.id ? next : item))
}

export function removeItem(items: readonly MarkupItem[], id: string): MarkupItem[] {
  return items.filter((item) => item.id !== id)
}

// ---------------------------------------------------------------------------------------------
// Session history (Undo/Redo while Markup is open)
// ---------------------------------------------------------------------------------------------

export interface MarkupHistory {
  readonly past: readonly (readonly MarkupItem[])[]
  readonly present: readonly MarkupItem[]
  readonly future: readonly (readonly MarkupItem[])[]
}

export function emptyMarkupHistory(items: readonly MarkupItem[] = []): MarkupHistory {
  return { past: [], present: items, future: [] }
}

/** A new present state (drops the redo branch); identical states are not recorded twice. */
export function commitMarkup(history: MarkupHistory, items: readonly MarkupItem[]): MarkupHistory {
  if (items === history.present) return history
  const past = [...history.past, history.present]
  if (past.length > MAX_HISTORY) past.splice(0, past.length - MAX_HISTORY)
  return { past, present: items, future: [] }
}

export function undoMarkup(history: MarkupHistory): MarkupHistory {
  if (!history.past.length) return history
  const previous = history.past[history.past.length - 1]
  return { past: history.past.slice(0, -1), present: previous, future: [history.present, ...history.future] }
}

export function redoMarkup(history: MarkupHistory): MarkupHistory {
  if (!history.future.length) return history
  const [next, ...rest] = history.future
  return { past: [...history.past, history.present], present: next, future: rest }
}

/** Items that would actually paint something when baked. */
export function paintableItems(items: readonly MarkupItem[]): MarkupItem[] {
  return items.filter((item) => (item.kind === 'text' ? item.text.trim().length > 0 : item.kind === 'image' ? item.width > 0 && item.height > 0 : true))
}
