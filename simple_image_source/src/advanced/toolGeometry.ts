// src/advanced/toolGeometry.ts (WP5)
// Pure geometry of the Advanced tools (design 5.9): selection modifiers, marquee / shape / line drags,
// angle snapping, polygon closing, the crop box (aspect presets, handles, rotation, straighten) and the
// free-transform frame (Photoshop handle semantics: proportional corners by default with Shift toggling,
// Alt about the reference point, rotation snapping, skew, distort and perspective quads).
// DOM-free and erasable TypeScript only, so Node tests load it directly.
//
// Conventions
//   - Document coordinates, y down. Angles are degrees, positive = clockwise on screen.
//   - A Quad lists its corners as top-left, top-right, bottom-right, bottom-left of the source rectangle
//     after transformation. Local (source) coordinates are mapped to the quad by a homography.
import type { Affine, Homography, Point, Rect, SelectionOp, Size } from '../imaging/types.ts'
import type { AspectPreset } from './types.ts'
import { applyHomography, homographyFromQuads, invertHomography } from '../imaging/transform.ts'
import { clampZoom } from './viewport.ts'

const EPSILON = 1e-9

function finite(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback
}

function sub(a: Point, b: Point): Point {
  return { x: a.x - b.x, y: a.y - b.y }
}

function add(a: Point, b: Point): Point {
  return { x: a.x + b.x, y: a.y + b.y }
}

function scale(a: Point, k: number): Point {
  return { x: a.x * k, y: a.y * k }
}

function dot(a: Point, b: Point): number {
  return a.x * b.x + a.y * b.y
}

function radians(degrees: number): number {
  return (degrees * Math.PI) / 180
}

function degrees(rad: number): number {
  return (rad * 180) / Math.PI
}

export function distance(a: Point, b: Point): number {
  return Math.hypot(a.x - b.x, a.y - b.y)
}

/** Rotates `p` about `center` by `angle` degrees (clockwise on screen). */
export function rotatePoint(p: Point, center: Point, angle: number): Point {
  const t = radians(angle)
  const c = Math.cos(t)
  const s = Math.sin(t)
  const dx = p.x - center.x
  const dy = p.y - center.y
  return { x: center.x + dx * c - dy * s, y: center.y + dx * s + dy * c }
}

/** Angle wrapped to (-180, 180]. */
export function normalizeAngle(angle: number): number {
  let a = finite(angle, 0) % 360
  if (a <= -180) a += 360
  if (a > 180) a -= 360
  return a
}

/** Rectangle spanned by two points (any order). */
export function rectFromPoints(a: Point, b: Point): Rect {
  return { x: Math.min(a.x, b.x), y: Math.min(a.y, b.y), width: Math.abs(b.x - a.x), height: Math.abs(b.y - a.y) }
}

/** Rectangle with edges rounded to whole pixels (Rectangular Marquee snaps to the pixel grid). */
export function snapRectToPixels(rect: Rect): Rect {
  const x0 = Math.round(rect.x)
  const y0 = Math.round(rect.y)
  const x1 = Math.round(rect.x + rect.width)
  const y1 = Math.round(rect.y + rect.height)
  return { x: Math.min(x0, x1), y: Math.min(y0, y1), width: Math.abs(x1 - x0), height: Math.abs(y1 - y0) }
}

// ---------------------------------------------------------------------------------------------
// Selection modifiers (Photoshop): with a selection, Shift = add, Alt = subtract, Shift+Alt = intersect
// at pointer-down. Keys used to pick the operation only constrain the shape after being released and
// pressed again; without a selection they constrain right away (Shift = square, Alt = from centre).
// ---------------------------------------------------------------------------------------------

export interface SelectionModifiers {
  readonly op: SelectionOp
  /** Shift chose the operation: it does not constrain until it is released. */
  readonly shiftLatched: boolean
  /** Alt chose the operation: it does not draw from the centre until it is released. */
  readonly altLatched: boolean
}

export function selectionOpFromModifiers(shift: boolean, alt: boolean, hasSelection: boolean, fallback: SelectionOp): SelectionModifiers {
  // Without a selection every mode acts as a new selection (subtracting from nothing would select nothing).
  if (!hasSelection) return { op: 'replace', shiftLatched: false, altLatched: false }
  if (shift && alt) return { op: 'intersect', shiftLatched: true, altLatched: true }
  if (shift) return { op: 'add', shiftLatched: true, altLatched: false }
  if (alt) return { op: 'subtract', shiftLatched: false, altLatched: true }
  return { op: fallback, shiftLatched: false, altLatched: false }
}

/** Updates a latched modifier with the current key state; `active` is true when it should constrain now. */
export function modifierState(latched: boolean, pressed: boolean): { readonly latched: boolean; readonly active: boolean } {
  if (latched && !pressed) return { latched: false, active: false }
  return { latched, active: pressed && !latched }
}

// ---------------------------------------------------------------------------------------------
// Rectangles from drags (marquees, shapes, crop, paragraph text)
// ---------------------------------------------------------------------------------------------

export interface DragConstraint {
  /** Square / circle (or the fixed ratio when given). */
  readonly square?: boolean
  /** The anchor is the centre instead of a corner. */
  readonly fromCenter?: boolean
  /** Width / height ratio to keep (overrides square). */
  readonly ratio?: number | null
}

/** Rectangle from a drag: anchor at a corner (or the centre), constrained to a square or ratio. */
export function constrainedRect(anchor: Point, point: Point, constraint: DragConstraint = {}): Rect {
  let dx = finite(point.x - anchor.x, 0)
  let dy = finite(point.y - anchor.y, 0)
  const ratio = constraint.ratio && constraint.ratio > 0 && Number.isFinite(constraint.ratio) ? constraint.ratio : constraint.square ? 1 : null
  if (ratio) {
    const sx = dx < 0 ? -1 : 1
    const sy = dy < 0 ? -1 : 1
    const width = Math.max(Math.abs(dx), Math.abs(dy) * ratio)
    dx = sx * width
    dy = sy * (width / ratio)
  }
  if (constraint.fromCenter) return { x: anchor.x - Math.abs(dx), y: anchor.y - Math.abs(dy), width: 2 * Math.abs(dx), height: 2 * Math.abs(dy) }
  return rectFromPoints(anchor, { x: anchor.x + dx, y: anchor.y + dy })
}

export interface MarqueeStyle {
  readonly style: 'normal' | 'fixed-ratio' | 'fixed-size'
  readonly ratio: Size
  readonly fixedSize: Size
}

/**
 * Marquee rectangle (Rectangular / Elliptical Marquee). Normal: free, Shift square, Alt from centre.
 * Fixed ratio: the drag keeps ratio.width : ratio.height. Fixed size: a rectangle of fixedSize whose top-left
 * (or centre with Alt) follows the pointer.
 */
export function marqueeRect(anchor: Point, point: Point, constraint: { readonly square: boolean; readonly fromCenter: boolean }, style?: MarqueeStyle): Rect {
  if (style && style.style === 'fixed-size') {
    const width = Math.max(1, finite(style.fixedSize.width, 1))
    const height = Math.max(1, finite(style.fixedSize.height, 1))
    if (constraint.fromCenter) return { x: point.x - width / 2, y: point.y - height / 2, width, height }
    return { x: point.x, y: point.y, width, height }
  }
  if (style && style.style === 'fixed-ratio') {
    const rw = finite(style.ratio.width, 1)
    const rh = finite(style.ratio.height, 1)
    const ratio = rw > 0 && rh > 0 ? rw / rh : 1
    return constrainedRect(anchor, point, { ratio, fromCenter: constraint.fromCenter })
  }
  return constrainedRect(anchor, point, { square: constraint.square, fromCenter: constraint.fromCenter })
}

// ---------------------------------------------------------------------------------------------
// Angles and lines
// ---------------------------------------------------------------------------------------------

/** `point` moved onto the nearest ray from `origin` at a multiple of `step` degrees (distance kept). */
export function snapToAngle(origin: Point, point: Point, step = 45): Point {
  const dx = point.x - origin.x
  const dy = point.y - origin.y
  const length = Math.hypot(dx, dy)
  if (!(length > EPSILON) || !(step > 0)) return { x: point.x, y: point.y }
  const angle = Math.round(degrees(Math.atan2(dy, dx)) / step) * step
  const t = radians(angle)
  const x = origin.x + length * Math.cos(t)
  const y = origin.y + length * Math.sin(t)
  // Remove floating dust on axis-aligned results so horizontal lines stay exactly horizontal.
  return { x: Math.abs(x - origin.x) < 1e-9 ? origin.x : x, y: Math.abs(y - origin.y) < 1e-9 ? origin.y : y }
}

/** A line drag (line / arrow shapes, gradients): Shift snaps to 45 degrees, Alt mirrors about the anchor. */
export function lineFromDrag(anchor: Point, point: Point, options: { readonly snap?: boolean; readonly fromCenter?: boolean } = {}): { readonly start: Point; readonly end: Point } {
  const end = options.snap ? snapToAngle(anchor, point, 45) : { x: point.x, y: point.y }
  const start = options.fromCenter ? { x: 2 * anchor.x - end.x, y: 2 * anchor.y - end.y } : { x: anchor.x, y: anchor.y }
  return { start, end }
}

/** Snaps an absolute angle to a multiple of `step` degrees. */
export function snapAngleValue(angle: number, step: number): number {
  if (!(step > 0)) return angle
  return Math.round(angle / step) * step
}

// ---------------------------------------------------------------------------------------------
// Polygons (lasso)
// ---------------------------------------------------------------------------------------------

/** Twice the signed area (positive = clockwise on screen). */
export function polygonArea(points: readonly Point[]): number {
  let area = 0
  for (let i = 0, j = points.length - 1; i < points.length; j = i, i += 1) {
    area += points[j].x * points[i].y - points[i].x * points[j].y
  }
  return area / 2
}

/** True when a click at `point` closes the polygon: at least three vertices and within `tolerance` of the first. */
export function closesPolygon(points: readonly Point[], point: Point, tolerance: number): boolean {
  return points.length >= 3 && distance(points[0], point) <= Math.max(0, tolerance)
}

/** True when `points` enclose no area worth selecting. */
export function isDegeneratePolygon(points: readonly Point[], minArea = 0.25): boolean {
  return points.length < 3 || Math.abs(polygonArea(points)) < minArea
}

/** Appends `point` unless it is closer than `minDistance` to the last point. Returns true when appended. */
export function appendPathPoint(points: Point[], point: Point, minDistance: number): boolean {
  if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) return false
  const last = points[points.length - 1]
  if (last && distance(last, point) < minDistance) return false
  points.push({ x: point.x, y: point.y })
  return true
}

function perpendicularDistance(p: Point, a: Point, b: Point): number {
  const d = sub(b, a)
  const length2 = dot(d, d)
  if (!(length2 > EPSILON)) return distance(p, a)
  const t = Math.max(0, Math.min(1, dot(sub(p, a), d) / length2))
  return distance(p, add(a, scale(d, t)))
}

/** Ramer-Douglas-Peucker simplification (keeps the end points). */
export function simplifyPath(points: readonly Point[], tolerance: number): Point[] {
  if (points.length <= 2 || !(tolerance > 0)) return points.map((p) => ({ x: p.x, y: p.y }))
  const keep = new Uint8Array(points.length)
  keep[0] = 1
  keep[points.length - 1] = 1
  const stack: [number, number][] = [[0, points.length - 1]]
  while (stack.length) {
    const [first, last] = stack.pop() as [number, number]
    let index = -1
    let farthest = tolerance
    for (let i = first + 1; i < last; i += 1) {
      const d = perpendicularDistance(points[i], points[first], points[last])
      if (d > farthest) {
        farthest = d
        index = i
      }
    }
    if (index >= 0) {
      keep[index] = 1
      stack.push([first, index], [index, last])
    }
  }
  const out: Point[] = []
  for (let i = 0; i < points.length; i += 1) if (keep[i]) out.push({ x: points[i].x, y: points[i].y })
  return out
}

/** Even-odd point-in-polygon test. */
export function pointInPolygon(points: readonly Point[], p: Point): boolean {
  let inside = false
  for (let i = 0, j = points.length - 1; i < points.length; j = i, i += 1) {
    const a = points[i]
    const b = points[j]
    if ((a.y > p.y) !== (b.y > p.y) && p.x < ((b.x - a.x) * (p.y - a.y)) / (b.y - a.y) + a.x) inside = !inside
  }
  return inside
}

// ---------------------------------------------------------------------------------------------
// Clicks and keys
// ---------------------------------------------------------------------------------------------

export interface ClickStamp {
  readonly time: number
  readonly x: number
  readonly y: number
}

/** Second click of a double-click: within `maxMs` and `maxDistance` (screen px) of the previous one. */
export function isDoubleClick(previous: ClickStamp | null, next: ClickStamp, maxMs = 450, maxDistance = 5): boolean {
  if (!previous) return false
  const elapsed = next.time - previous.time
  return elapsed >= 0 && elapsed <= maxMs && Math.hypot(next.x - previous.x, next.y - previous.y) <= maxDistance
}

/** Arrow-key nudge: 1 px, 10 px with Shift; null for other keys. */
export function nudgeDelta(key: string, shift: boolean): Point | null {
  const step = shift ? 10 : 1
  switch (key) {
    case 'ArrowLeft': return { x: -step, y: 0 }
    case 'ArrowRight': return { x: step, y: 0 }
    case 'ArrowUp': return { x: 0, y: -step }
    case 'ArrowDown': return { x: 0, y: step }
    default: return null
  }
}

// ---------------------------------------------------------------------------------------------
// Zoom tool
// ---------------------------------------------------------------------------------------------

/** Scrubby zoom: dragging right by 150 CSS px doubles the zoom, left halves it (clamped to 1%..3200%). */
export function scrubbyZoom(startZoom: number, dx: number, dpr: number): number {
  return clampZoom(finite(startZoom, 1) * Math.pow(2, finite(dx, 0) / 150), dpr)
}

// ---------------------------------------------------------------------------------------------
// Crop box (a rectangle of width x height centred on (cx, cy), turned by `angle` degrees)
// ---------------------------------------------------------------------------------------------

export type CropHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'

export interface CropBox {
  readonly cx: number
  readonly cy: number
  readonly width: number
  readonly height: number
  /** Degrees, clockwise on screen. */
  readonly angle: number
}

export const CROP_HANDLES: readonly CropHandle[] = Object.freeze(['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const)

/** Width / height ratios of the fixed crop presets (landscape). */
export const CROP_PRESET_RATIOS: Readonly<Record<'1:1' | '4:3' | '3:2' | '16:9' | '5:4' | '7:5', number>> = Object.freeze({
  '1:1': 1,
  '4:3': 4 / 3,
  '3:2': 3 / 2,
  '16:9': 16 / 9,
  '5:4': 5 / 4,
  '7:5': 7 / 5,
})

/**
 * Width / height ratio a preset imposes, or null for Free. 'original' uses the document. `portrait` swaps
 * the preset to its tall form (Photoshop's swap button).
 */
export function cropAspectRatio(preset: AspectPreset, portrait: boolean, doc: Size): number | null {
  let ratio: number | null = null
  if (preset === 'original') ratio = doc.width > 0 && doc.height > 0 ? doc.width / doc.height : null
  else if (preset !== 'free') ratio = CROP_PRESET_RATIOS[preset] ?? null
  if (ratio === null) return null
  if (preset === 'original') return portrait ? 1 / ratio : ratio
  const landscape = Math.max(ratio, 1 / ratio)
  return portrait ? 1 / landscape : landscape
}

export function cropBoxFromRect(rect: Rect, angle = 0): CropBox {
  return { cx: rect.x + rect.width / 2, cy: rect.y + rect.height / 2, width: Math.abs(rect.width), height: Math.abs(rect.height), angle }
}

/** Axis-aligned rectangle of an unrotated box. */
export function cropBoxRect(box: CropBox): Rect {
  return { x: box.cx - box.width / 2, y: box.cy - box.height / 2, width: box.width, height: box.height }
}

/** Box-local coordinates (origin at the centre, unrotated) of a document point. */
export function cropToLocal(box: CropBox, p: Point): Point {
  const q = rotatePoint(p, { x: box.cx, y: box.cy }, -box.angle)
  return { x: q.x - box.cx, y: q.y - box.cy }
}

/** Document point of box-local coordinates. */
export function cropFromLocal(box: CropBox, local: Point): Point {
  return rotatePoint({ x: box.cx + local.x, y: box.cy + local.y }, { x: box.cx, y: box.cy }, box.angle)
}

function handleSigns(handle: CropHandle): { sx: number; sy: number } {
  return {
    sx: handle.includes('w') ? -1 : handle.includes('e') ? 1 : 0,
    sy: handle.startsWith('n') ? -1 : handle.startsWith('s') ? 1 : 0,
  }
}

/** Document position of a handle of the box. */
export function cropHandlePoint(box: CropBox, handle: CropHandle): Point {
  const { sx, sy } = handleSigns(handle)
  return cropFromLocal(box, { x: (sx * box.width) / 2, y: (sy * box.height) / 2 })
}

/** Corners in document space: top-left, top-right, bottom-right, bottom-left (of the turned box). */
export function cropBoxCorners(box: CropBox): [Point, Point, Point, Point] {
  return [cropHandlePoint(box, 'nw'), cropHandlePoint(box, 'ne'), cropHandlePoint(box, 'se'), cropHandlePoint(box, 'sw')]
}

/** What the pointer is over: a handle (within `tolerance` document px), the inside, or the outside. */
export function hitTestCrop(box: CropBox, point: Point, tolerance: number): CropHandle | 'inside' | 'outside' {
  let best: CropHandle | null = null
  let bestDistance = Math.max(0, tolerance)
  for (const handle of CROP_HANDLES) {
    const d = distance(cropHandlePoint(box, handle), point)
    if (d <= bestDistance) {
      best = handle
      bestDistance = d
    }
  }
  if (best) return best
  const local = cropToLocal(box, point)
  return Math.abs(local.x) <= box.width / 2 && Math.abs(local.y) <= box.height / 2 ? 'inside' : 'outside'
}

export interface CropResizeOptions {
  /** Width / height to keep (a preset), or null. */
  readonly ratio: number | null
  /** Shift on a corner of a free box keeps its current ratio. */
  readonly keepRatio?: boolean
  /** Alt: resize about the centre. */
  readonly fromCenter?: boolean
  readonly minSize?: number
}

/**
 * The box after dragging `handle` to `pointer` (document point), measured in the box's own (turned) frame.
 * The opposite handle stays put (the centre with fromCenter); a preset ratio or Shift keeps the ratio.
 * Dragging past the opposite side flips the box.
 */
export function resizeCrop(start: CropBox, handle: CropHandle, pointer: Point, options: CropResizeOptions): CropBox {
  const minSize = Math.max(1, finite(options.minSize ?? 1, 1))
  const { sx, sy } = handleSigns(handle)
  const p = cropToLocal(start, pointer)
  const ratio = options.ratio && options.ratio > 0 ? options.ratio
    : options.keepRatio && sx !== 0 && sy !== 0 && start.height > 0 ? start.width / start.height : null
  const halfW = start.width / 2
  const halfH = start.height / 2
  // Fixed point (box-local) and the moving extents.
  const fx = options.fromCenter ? 0 : -sx * halfW
  const fy = options.fromCenter ? 0 : -sy * halfH
  const factor = options.fromCenter ? 2 : 1
  let width = start.width
  let height = start.height
  let dirX = 1
  let dirY = 1
  if (sx !== 0) {
    const span = (p.x - fx) * sx
    dirX = span < 0 ? -1 : 1
    width = Math.max(minSize, Math.abs(span) * factor)
  }
  if (sy !== 0) {
    const span = (p.y - fy) * sy
    dirY = span < 0 ? -1 : 1
    height = Math.max(minSize, Math.abs(span) * factor)
  }
  if (ratio) {
    if (sx !== 0 && sy !== 0) {
      if (width / height > ratio) height = width / ratio
      else width = height * ratio
    } else if (sx !== 0) {
      height = width / ratio
    } else {
      width = height * ratio
    }
    if (width < minSize) {
      width = minSize
      height = width / ratio
    }
    if (height < minSize) {
      height = minSize
      width = height * ratio
    }
  }
  // New centre (box-local): from the fixed point towards the dragged side.
  let cx = 0
  let cy = 0
  if (!options.fromCenter) {
    cx = sx !== 0 ? fx + sx * dirX * (width / 2) : 0
    cy = sy !== 0 ? fy + sy * dirY * (height / 2) : 0
    // An edge handle with a ratio grows the other axis symmetrically about the centre line.
  }
  const center = cropFromLocal(start, { x: cx, y: cy })
  return { cx: center.x, cy: center.y, width, height, angle: start.angle }
}

/** A new, unturned box dragged from `anchor` to `point` (Shift square when free, Alt from the centre). */
export function dragCrop(anchor: Point, point: Point, options: { readonly ratio: number | null; readonly square?: boolean; readonly fromCenter?: boolean }): CropBox {
  const rect = constrainedRect(anchor, point, { ratio: options.ratio, square: options.square, fromCenter: options.fromCenter })
  return cropBoxFromRect(rect)
}

export function moveCrop(start: CropBox, dx: number, dy: number): CropBox {
  return { ...start, cx: start.cx + finite(dx, 0), cy: start.cy + finite(dy, 0) }
}

/** Turns the box about its centre by the pointer's sweep from `from` to `to`; `snap` degrees snap the total angle. */
export function rotateCrop(start: CropBox, from: Point, to: Point, snap?: number): CropBox {
  const center = { x: start.cx, y: start.cy }
  const sweep = degrees(Math.atan2(to.y - center.y, to.x - center.x) - Math.atan2(from.y - center.y, from.x - center.x))
  let angle = normalizeAngle(start.angle + sweep)
  if (snap && snap > 0) angle = normalizeAngle(snapAngleValue(angle, snap))
  return { ...start, angle }
}

/**
 * Applies a ratio to a box (preset change): keeps the centre and the longer fitting dimension so the new
 * box stays inside the old one.
 */
export function fitCropRatio(box: CropBox, ratio: number | null): CropBox {
  if (!ratio || !(ratio > 0)) return box
  let width = box.width
  let height = width / ratio
  if (height > box.height) {
    height = box.height
    width = height * ratio
  }
  return { ...box, width: Math.max(1, width), height: Math.max(1, height) }
}

/**
 * Straighten: the angle (degrees) to turn the crop box by so the drawn line becomes horizontal or
 * vertical, whichever is nearer (Photoshop's Straighten).
 */
export function straightenAngle(a: Point, b: Point): number {
  if (!(distance(a, b) > EPSILON)) return 0
  const angle = degrees(Math.atan2(b.y - a.y, b.x - a.x))
  // Nearest multiple of 90 degrees; the box turns by the difference.
  const target = Math.round(angle / 90) * 90
  return normalizeAngle(angle - target)
}

/** Whole-pixel crop rectangle of an unturned box (edges rounded). */
export function cropCommitRect(box: CropBox): { readonly x: number; readonly y: number; readonly width: number; readonly height: number } {
  const rect = snapRectToPixels(cropBoxRect(box))
  return { x: rect.x, y: rect.y, width: Math.max(1, rect.width), height: Math.max(1, rect.height) }
}

/**
 * The affine that maps the document into the cropped canvas of a turned box: rotate by -angle about the
 * box centre, then put the box's top-left corner at the origin. Width and height round to whole pixels.
 */
export function cropTransform(box: CropBox): { readonly matrix: Affine; readonly width: number; readonly height: number } {
  const width = Math.max(1, Math.round(box.width))
  const height = Math.max(1, Math.round(box.height))
  const t = radians(-box.angle)
  const c = cleanUnit(Math.cos(t))
  const s = cleanUnit(Math.sin(t))
  // x' = c (x - cx) - s (y - cy) + w / 2 ; y' = s (x - cx) + c (y - cy) + h / 2
  const e = -c * box.cx + s * box.cy + width / 2
  const f = -s * box.cx - c * box.cy + height / 2
  return { matrix: [c, s, -s, c, e, f], width, height }
}

function cleanUnit(value: number): number {
  if (Math.abs(value) < 1e-12) return 0
  if (Math.abs(value - 1) < 1e-12) return 1
  if (Math.abs(value + 1) < 1e-12) return -1
  return value
}

// ---------------------------------------------------------------------------------------------
// Free transform frame
// ---------------------------------------------------------------------------------------------

/** Corners: top-left, top-right, bottom-right, bottom-left. */
export type Quad = readonly [Point, Point, Point, Point]

export type TransformHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w'

export const TRANSFORM_HANDLES: readonly TransformHandle[] = Object.freeze(['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const)

export interface TransformFrame {
  /** Untransformed bounds (document space). */
  readonly source: Rect
  /** Where the source corners are now. */
  readonly quad: Quad
  /** Reference point (document space): rotation and Alt scaling happen about it. */
  readonly pivot: Point
}

export function rectToQuad(rect: Rect): Quad {
  return [
    { x: rect.x, y: rect.y },
    { x: rect.x + rect.width, y: rect.y },
    { x: rect.x + rect.width, y: rect.y + rect.height },
    { x: rect.x, y: rect.y + rect.height },
  ]
}

export function quadCenter(quad: Quad): Point {
  return { x: (quad[0].x + quad[1].x + quad[2].x + quad[3].x) / 4, y: (quad[0].y + quad[1].y + quad[2].y + quad[3].y) / 4 }
}

export function createTransformFrame(source: Rect): TransformFrame {
  const quad = rectToQuad(source)
  return { source, quad, pivot: { x: source.x + source.width / 2, y: source.y + source.height / 2 } }
}

/** Homography mapping the source rectangle onto the quad. */
export function frameHomography(frame: Pick<TransformFrame, 'source' | 'quad'>): Homography {
  return homographyFromQuads(rectToQuad(frame.source), frame.quad)
}

/** True when the quad is a parallelogram (the transform is affine). */
export function isAffineQuad(quad: Quad, tolerance = 1e-6): boolean {
  const size = Math.max(1, distance(quad[0], quad[2]), distance(quad[1], quad[3]))
  const dx = quad[0].x + quad[2].x - quad[1].x - quad[3].x
  const dy = quad[0].y + quad[2].y - quad[1].y - quad[3].y
  return Math.hypot(dx, dy) <= tolerance * size
}

/** The affine (canvas order) that maps the source rectangle onto a parallelogram quad, or null. */
export function frameAffine(frame: Pick<TransformFrame, 'source' | 'quad'>, tolerance = 1e-6): Affine | null {
  const { source, quad } = frame
  if (!isAffineQuad(quad, tolerance) || !(source.width > 0) || !(source.height > 0)) return null
  const a = (quad[1].x - quad[0].x) / source.width
  const b = (quad[1].y - quad[0].y) / source.width
  const c = (quad[3].x - quad[0].x) / source.height
  const d = (quad[3].y - quad[0].y) / source.height
  const e = quad[0].x - a * source.x - c * source.y
  const f = quad[0].y - b * source.x - d * source.y
  return [a, b, c, d, e, f]
}

/** True when the frame has not moved, scaled, turned or distorted the source (within `tolerance` px). */
export function isIdentityFrame(frame: Pick<TransformFrame, 'source' | 'quad'>, tolerance = 1e-6): boolean {
  const original = rectToQuad(frame.source)
  for (let i = 0; i < 4; i += 1) if (distance(original[i], frame.quad[i]) > tolerance) return false
  return true
}

/** Whole-pixel translation of the frame, or null when it is not a pure integer move. */
export function frameIntegerTranslation(frame: Pick<TransformFrame, 'source' | 'quad'>, tolerance = 1e-6): Point | null {
  const original = rectToQuad(frame.source)
  const dx = frame.quad[0].x - original[0].x
  const dy = frame.quad[0].y - original[0].y
  if (Math.abs(dx - Math.round(dx)) > tolerance || Math.abs(dy - Math.round(dy)) > tolerance) return null
  for (let i = 1; i < 4; i += 1) {
    if (Math.abs(frame.quad[i].x - original[i].x - dx) > tolerance || Math.abs(frame.quad[i].y - original[i].y - dy) > tolerance) return null
  }
  return { x: Math.round(dx), y: Math.round(dy) }
}

function localHandlePoint(source: Rect, handle: TransformHandle): Point {
  const { sx, sy } = handleSigns(handle)
  return { x: source.x + ((sx + 1) / 2) * source.width, y: source.y + ((sy + 1) / 2) * source.height }
}

function mapLocal(h: Homography, p: Point): Point {
  return applyHomography(h, p)
}

/** Document position of a transform handle (edge handles sit at the mapped edge midpoints). */
export function transformHandlePoint(frame: Pick<TransformFrame, 'source' | 'quad'>, handle: TransformHandle): Point {
  switch (handle) {
    case 'nw': return frame.quad[0]
    case 'ne': return frame.quad[1]
    case 'se': return frame.quad[2]
    case 'sw': return frame.quad[3]
    default: return mapLocal(frameHomography(frame), localHandlePoint(frame.source, handle))
  }
}

export function isCornerHandle(handle: TransformHandle): boolean {
  return handle.length === 2
}

export type TransformHit =
  | { readonly kind: 'handle'; readonly handle: TransformHandle }
  | { readonly kind: 'pivot' }
  | { readonly kind: 'inside' }
  | { readonly kind: 'outside' }

/** What the pointer is over (tolerance in document px): the pivot, a handle, inside the quad or outside it. */
export function hitTestTransform(frame: TransformFrame, point: Point, tolerance: number): TransformHit {
  const t = Math.max(0, tolerance)
  if (distance(frame.pivot, point) <= t) return { kind: 'pivot' }
  let best: TransformHandle | null = null
  let bestDistance = t
  for (const handle of TRANSFORM_HANDLES) {
    const d = distance(transformHandlePoint(frame, handle), point)
    if (d <= bestDistance) {
      best = handle
      bestDistance = d
    }
  }
  if (best) return { kind: 'handle', handle: best }
  return pointInPolygon(frame.quad, point) ? { kind: 'inside' } : { kind: 'outside' }
}

/** Source (local) coordinates of a document point through the inverse homography; null beyond the horizon. */
function toLocal(h: Homography, inverse: Homography, p: Point): Point | null {
  const w = inverse[6] * p.x + inverse[7] * p.y + inverse[8]
  if (!(Math.abs(w) > 1e-12)) return null
  const local = applyHomography(inverse, p)
  if (!Number.isFinite(local.x) || !Number.isFinite(local.y)) return null
  // Reject points that map back through the line at infinity (the quad would fold over).
  const forward = h[6] * local.x + h[7] * local.y + h[8]
  if (!(forward > 0)) return null
  return local
}

export interface ScaleOptions {
  /** Keep the aspect ratio (corner default in Photoshop; Shift toggles). */
  readonly proportional: boolean
  /** Alt: scale about the reference point instead of the opposite handle. */
  readonly fromPivot: boolean
}

const MIN_SCALE = 1e-3

function safeScale(value: number): number {
  if (!Number.isFinite(value)) return 1
  if (Math.abs(value) < MIN_SCALE) return value < 0 ? -MIN_SCALE : MIN_SCALE
  return value
}

/**
 * Scales the frame by dragging `handle` to `pointer`, in the frame's own (source) coordinates, so rotated,
 * skewed and perspective frames scale along their own axes. Corner handles scale both axes (proportionally
 * when asked: the pointer is projected on the diagonal); edge handles scale one axis (both, equally, when
 * proportional). Dragging across the fixed point flips.
 */
export function scaleFrame(start: TransformFrame, handle: TransformHandle, pointer: Point, options: ScaleOptions): Quad {
  const h = frameHomography(start)
  let inverse: Homography
  try {
    inverse = invertHomography(h)
  } catch {
    return start.quad
  }
  const pointerLocal = toLocal(h, inverse, pointer)
  if (!pointerLocal) return start.quad
  const handleLocal = localHandlePoint(start.source, handle)
  let fixed: Point
  if (options.fromPivot) {
    fixed = toLocal(h, inverse, start.pivot) ?? { x: start.source.x + start.source.width / 2, y: start.source.y + start.source.height / 2 }
  } else {
    const { sx, sy } = handleSigns(handle)
    const opposite = ((sy < 0 ? 's' : sy > 0 ? 'n' : '') + (sx < 0 ? 'e' : sx > 0 ? 'w' : '')) as TransformHandle
    fixed = localHandlePoint(start.source, opposite)
  }
  const { sx: hx, sy: hy } = handleSigns(handle)
  let kx = 1
  let ky = 1
  const spanX = handleLocal.x - fixed.x
  const spanY = handleLocal.y - fixed.y
  if (hx !== 0 && hy !== 0) {
    if (options.proportional) {
      const along = { x: spanX, y: spanY }
      const length2 = dot(along, along)
      const k = length2 > EPSILON ? dot(sub(pointerLocal, fixed), along) / length2 : 1
      kx = safeScale(k)
      ky = kx
    } else {
      kx = Math.abs(spanX) > EPSILON ? safeScale((pointerLocal.x - fixed.x) / spanX) : 1
      ky = Math.abs(spanY) > EPSILON ? safeScale((pointerLocal.y - fixed.y) / spanY) : 1
    }
  } else if (hx !== 0) {
    kx = Math.abs(spanX) > EPSILON ? safeScale((pointerLocal.x - fixed.x) / spanX) : 1
    if (options.proportional) ky = Math.abs(kx)
  } else if (hy !== 0) {
    ky = Math.abs(spanY) > EPSILON ? safeScale((pointerLocal.y - fixed.y) / spanY) : 1
    if (options.proportional) kx = Math.abs(ky)
  }
  const corners = rectToQuad(start.source)
  const mapped = corners.map((c) => applyHomography(h, { x: fixed.x + (c.x - fixed.x) * kx, y: fixed.y + (c.y - fixed.y) * ky }))
  if (mapped.some((p) => !Number.isFinite(p.x) || !Number.isFinite(p.y))) return start.quad
  return mapped as unknown as Quad
}

/** Angle (degrees) of the frame's top edge. */
export function quadAngle(quad: Quad): number {
  return degrees(Math.atan2(quad[1].y - quad[0].y, quad[1].x - quad[0].x))
}

/**
 * Rotates the frame about its reference point by the pointer's sweep from `from` to `to`. With `snap`
 * (Shift: 15 degrees) the frame's total angle snaps to multiples of it.
 */
export function rotateFrame(start: TransformFrame, from: Point, to: Point, snap?: number): Quad {
  const pivot = start.pivot
  let sweep = degrees(Math.atan2(to.y - pivot.y, to.x - pivot.x) - Math.atan2(from.y - pivot.y, from.x - pivot.x))
  if (snap && snap > 0) {
    const current = quadAngle(start.quad)
    sweep = snapAngleValue(current + sweep, snap) - current
  }
  return start.quad.map((p) => rotatePoint(p, pivot, sweep)) as unknown as Quad
}

/** Ctrl+edge: slides the dragged edge along itself (skew); `symmetric` (Alt) slides the opposite edge the other way. */
export function skewFrame(start: TransformFrame, handle: TransformHandle, from: Point, to: Point, symmetric: boolean): Quad {
  const edges: Record<'n' | 's' | 'w' | 'e', [number, number]> = { n: [0, 1], s: [3, 2], w: [0, 3], e: [1, 2] }
  const opposite: Record<'n' | 's' | 'w' | 'e', 'n' | 's' | 'w' | 'e'> = { n: 's', s: 'n', w: 'e', e: 'w' }
  if (isCornerHandle(handle)) return distortFrame(start, handle, from, to)
  const side = handle as 'n' | 's' | 'w' | 'e'
  const [i, j] = edges[side]
  const quad = start.quad.map((p) => ({ x: p.x, y: p.y }))
  const edge = sub(start.quad[j], start.quad[i])
  const length2 = dot(edge, edge)
  if (!(length2 > EPSILON)) return start.quad
  const delta = sub(to, from)
  const along = scale(edge, dot(delta, edge) / length2)
  quad[i] = add(quad[i], along)
  quad[j] = add(quad[j], along)
  if (symmetric) {
    const [k, l] = edges[opposite[side]]
    quad[k] = sub(quad[k], along)
    quad[l] = sub(quad[l], along)
  }
  return quad as unknown as Quad
}

const CORNER_INDEX: Readonly<Record<'nw' | 'ne' | 'se' | 'sw', number>> = Object.freeze({ nw: 0, ne: 1, se: 2, sw: 3 })

/** Ctrl+corner: moves one corner freely (distort). */
export function distortFrame(start: TransformFrame, handle: TransformHandle, from: Point, to: Point): Quad {
  const delta = sub(to, from)
  const quad = start.quad.map((p) => ({ x: p.x, y: p.y }))
  if (isCornerHandle(handle)) {
    const index = CORNER_INDEX[handle as 'nw' | 'ne' | 'se' | 'sw']
    quad[index] = add(quad[index], delta)
    return quad as unknown as Quad
  }
  // An edge handle distorts its whole edge.
  const edges: Record<'n' | 's' | 'w' | 'e', [number, number]> = { n: [0, 1], s: [3, 2], w: [0, 3], e: [1, 2] }
  const [i, j] = edges[handle as 'n' | 's' | 'w' | 'e']
  quad[i] = add(quad[i], delta)
  quad[j] = add(quad[j], delta)
  return quad as unknown as Quad
}

/**
 * Ctrl+Alt+Shift+corner: perspective. The corner slides along the incident edge the drag follows most, and
 * the other corner of that edge slides the opposite way, so the quad stays a symmetric trapezoid.
 */
export function perspectiveFrame(start: TransformFrame, handle: TransformHandle, from: Point, to: Point): Quad {
  if (!isCornerHandle(handle)) return skewFrame(start, handle, from, to, false)
  const index = CORNER_INDEX[handle as 'nw' | 'ne' | 'se' | 'sw']
  // Neighbours: the corner on the same horizontal edge, and the one on the same vertical edge.
  const horizontal = [1, 0, 3, 2][index]
  const vertical = [3, 2, 1, 0][index]
  const quad = start.quad.map((p) => ({ x: p.x, y: p.y }))
  const delta = sub(to, from)
  const hEdge = sub(start.quad[horizontal], start.quad[index])
  const vEdge = sub(start.quad[vertical], start.quad[index])
  const hLength = Math.hypot(hEdge.x, hEdge.y)
  const vLength = Math.hypot(vEdge.x, vEdge.y)
  const hAmount = hLength > EPSILON ? dot(delta, hEdge) / hLength : 0
  const vAmount = vLength > EPSILON ? dot(delta, vEdge) / vLength : 0
  if (Math.abs(hAmount) >= Math.abs(vAmount)) {
    if (!(hLength > EPSILON)) return start.quad
    const move = scale(hEdge, hAmount / hLength)
    quad[index] = add(quad[index], move)
    quad[horizontal] = sub(quad[horizontal], move)
  } else {
    if (!(vLength > EPSILON)) return start.quad
    const move = scale(vEdge, vAmount / vLength)
    quad[index] = add(quad[index], move)
    quad[vertical] = sub(quad[vertical], move)
  }
  return quad as unknown as Quad
}

export function translateQuad(quad: Quad, dx: number, dy: number): Quad {
  return quad.map((p) => ({ x: p.x + dx, y: p.y + dy })) as unknown as Quad
}

/** Axis-aligned bounds of a quad. */
export function quadBounds(quad: Quad): Rect {
  const xs = quad.map((p) => p.x)
  const ys = quad.map((p) => p.y)
  const x = Math.min(...xs)
  const y = Math.min(...ys)
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y }
}

/** True when the quad is convex and not folded (a valid transform target). */
export function isValidQuad(quad: Quad): boolean {
  let sign = 0
  for (let i = 0; i < 4; i += 1) {
    const a = quad[i]
    const b = quad[(i + 1) % 4]
    const c = quad[(i + 2) % 4]
    const cross = (b.x - a.x) * (c.y - b.y) - (b.y - a.y) * (c.x - b.x)
    if (!Number.isFinite(cross)) return false
    if (Math.abs(cross) < 1e-9) continue
    const s = cross > 0 ? 1 : -1
    if (sign === 0) sign = s
    else if (s !== sign) return false
  }
  return sign !== 0
}

export interface FrameInfo {
  /** Reference point (document px). */
  readonly x: number
  readonly y: number
  /** Lengths of the frame's own axes (document px). */
  readonly width: number
  readonly height: number
  /** Degrees of the top edge. */
  readonly angle: number
}

/** Numeric fields of the options bar (Photoshop shows the reference point, size and angle). */
export function frameInfo(frame: TransformFrame): FrameInfo {
  const q = frame.quad
  return {
    x: frame.pivot.x,
    y: frame.pivot.y,
    width: distance(q[0], q[1]),
    height: distance(q[0], q[3]),
    angle: normalizeAngle(quadAngle(q)),
  }
}

/**
 * Rebuilds an affine frame from numeric fields: the source scaled to width x height, turned by angle and
 * placed so the reference point (kept at the same relative spot of the frame) lands on (x, y). Skew,
 * distortion and flips are reset.
 */
export function frameFromInfo(frame: TransformFrame, info: Partial<FrameInfo>): TransformFrame {
  const current = frameInfo(frame)
  const next = { ...current, ...Object.fromEntries(Object.entries(info).filter(([, value]) => Number.isFinite(value))) } as FrameInfo
  const source = frame.source
  let pivotLocal: Point = { x: source.x + source.width / 2, y: source.y + source.height / 2 }
  try {
    const h = frameHomography(frame)
    const local = toLocal(h, invertHomography(h), frame.pivot)
    if (local) pivotLocal = local
  } catch {
    // Keep the centre.
  }
  const kx = source.width > 0 ? Math.max(MIN_SCALE, next.width) / source.width : 1
  const ky = source.height > 0 ? Math.max(MIN_SCALE, next.height) / source.height : 1
  const pivot = { x: next.x, y: next.y }
  const quad = rectToQuad(source).map((c) => {
    const scaled = { x: pivot.x + (c.x - pivotLocal.x) * kx, y: pivot.y + (c.y - pivotLocal.y) * ky }
    return rotatePoint(scaled, pivot, next.angle)
  }) as unknown as Quad
  return { source, quad, pivot }
}

/** outer ∘ inner in canvas order (inner applies first). */
export function composeAffine(outer: Affine, inner: Affine): Affine {
  const [a, b, c, d, e, f] = outer
  const [g, h, i, j, k, l] = inner
  return [a * g + c * h, b * g + d * h, a * i + c * j, b * i + d * j, a * k + c * l + e, b * k + d * l + f]
}

/** Row-major 3x3 product a * b (b applies first). */
export function multiplyHomography(a: Homography, b: Homography): Homography {
  const out = new Array<number>(9)
  for (let row = 0; row < 3; row += 1) {
    for (let column = 0; column < 3; column += 1) {
      out[row * 3 + column] = a[row * 3] * b[column] + a[row * 3 + 1] * b[3 + column] + a[row * 3 + 2] * b[6 + column]
    }
  }
  return out as unknown as Homography
}

/** Translation as a homography. */
export function translationHomography(dx: number, dy: number): Homography {
  return [1, 0, dx, 0, 1, dy, 0, 0, 1]
}

/** Snaps matrix entries within 1e-9 of an integer to it (keeps flips and quarter turns pixel-exact). */
export function cleanHomography(h: Homography): Homography {
  return h.map((value) => {
    const nearest = Math.round(value)
    return Math.abs(value - nearest) < 1e-9 ? nearest : value
  }) as unknown as Homography
}

/**
 * The inverse map a warp needs: destination document coordinates to coordinates inside a source buffer
 * whose top-left pixel sits at (sourceX, sourceY), for the forward document-space homography `forward`.
 */
export function warpInverse(forward: Homography, sourceX: number, sourceY: number): Homography {
  return cleanHomography(multiplyHomography(translationHomography(-sourceX, -sourceY), invertHomography(forward)))
}

// ---------------------------------------------------------------------------------------------
// Text placement
// ---------------------------------------------------------------------------------------------

/**
 * Translation that keeps a point-text anchor fixed: the anchor sits at (alignment * width, baseline) in
 * text-local space (0 = left, 0.5 = centre, 1 = right) and must map to `anchor` in the document through the
 * linear part (a, b, c, d) of the text transform.
 */
export function anchoredTranslation(linear: readonly [number, number, number, number], anchor: Point, localAnchor: Point): Point {
  const [a, b, c, d] = linear
  return { x: anchor.x - (a * localAnchor.x + c * localAnchor.y), y: anchor.y - (b * localAnchor.x + d * localAnchor.y) }
}

export function alignmentFactor(align: 'left' | 'center' | 'right'): number {
  return align === 'center' ? 0.5 : align === 'right' ? 1 : 0
}
