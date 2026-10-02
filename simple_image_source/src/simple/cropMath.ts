// src/simple/cropMath.ts (WP8)
// Pure crop-box geometry for Simple mode (design 4.2):
//   - aspect presets (Free, Original, 1:1, 4:3, 3:2, 16:9) with an orientation swap;
//   - constrainDrag(): handle drags with a ratio lock (corner and edge handles), Alt = about the centre,
//     clamped to the bounds and to a minimum size; rectFromPoints() for a new box;
//   - arrow-key nudges, integer rounding for Apply;
//   - straighten helpers: the rotated ('expand') frame the crop box lives in while the image is turned,
//     the inside test against the turned image, the auto-crop inscribed rectangle, and a search that keeps
//     every drag inside the turned image (so a straightened crop never has transparent corners).
// DOM-free, erasable TypeScript with explicit .ts imports: Node tests load it directly.
import type { IntRect, Point, Rect, Size } from '../imaging/types.ts'
import { rotatedBounds } from '../imaging/transform.ts'

export type CropHandle = 'nw' | 'n' | 'ne' | 'e' | 'se' | 's' | 'sw' | 'w' | 'move'
export type ResizeHandle = Exclude<CropHandle, 'move'>
export const CROP_HANDLES: readonly ResizeHandle[] = Object.freeze(['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const)

/** Smallest crop box side in image pixels (smaller only when the image itself is smaller). */
export const MIN_CROP_SIZE = 8
/** Straighten limits (degrees) and slider step. */
export const STRAIGHTEN_LIMIT = 45
export const STRAIGHTEN_STEP = 0.1
/** Distance (source px) kept between a straightened crop and the image edge: bicubic edge pixels are partly transparent. */
export const STRAIGHTEN_DRAG_MARGIN = 1
export const STRAIGHTEN_AUTO_MARGIN = 2

export type AspectPreset = 'free' | 'original' | '1:1' | '4:3' | '3:2' | '16:9'

export interface AspectPresetInfo {
  readonly id: AspectPreset
  readonly label: string
  /** Landscape width / height, or null (free / follows the image). */
  readonly landscape: number | null
}

export const ASPECT_PRESETS: readonly AspectPresetInfo[] = Object.freeze([
  { id: 'free', label: 'Free', landscape: null },
  { id: 'original', label: 'Original', landscape: null },
  { id: '1:1', label: '1:1', landscape: 1 },
  { id: '4:3', label: '4:3', landscape: 4 / 3 },
  { id: '3:2', label: '3:2', landscape: 3 / 2 },
  { id: '16:9', label: '16:9', landscape: 16 / 9 },
])

export interface ConstrainOptions {
  /** Locked width / height, or null for a free box. */
  readonly ratio: number | null
  /** Resize symmetrically about the box centre (Alt). */
  readonly fromCenter: boolean
  /** The box must stay inside [0, width] x [0, height]. */
  readonly bounds: Size
  /** Default MIN_CROP_SIZE (reduced automatically when the bounds are smaller). */
  readonly minSize?: number
}

function clamp(value: number, minimum: number, maximum: number): number {
  return value < minimum ? minimum : value > maximum ? maximum : value
}

function finite(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback
}

/** True for presets whose ratio can be turned (everything except Free and 1:1). */
export function presetCanSwap(preset: AspectPreset): boolean {
  return preset !== 'free' && preset !== '1:1'
}

/**
 * Width / height of a preset for an image, or null for Free. `portrait` turns a landscape ratio upright
 * (Original follows the image, so its portrait form is the image turned).
 */
export function presetRatio(preset: AspectPreset, image: Size, portrait: boolean): number | null {
  if (preset === 'free') return null
  if (preset === 'original') {
    if (!(image.width > 0 && image.height > 0)) return null
    const ratio = image.width / image.height
    const imagePortrait = image.height > image.width
    return portrait === imagePortrait ? ratio : 1 / ratio
  }
  const info = ASPECT_PRESETS.find((entry) => entry.id === preset)
  const landscape = info?.landscape ?? null
  if (!landscape) return null
  return portrait ? 1 / landscape : landscape
}

/** The orientation a preset starts in for an image: upright for portrait images. */
export function defaultPortrait(image: Size): boolean {
  return image.height > image.width
}

/** Clamps a rectangle into the bounds (keeps its size when it fits). */
export function clampRect(rect: Rect, bounds: Size): Rect {
  const width = clamp(finite(rect.width, bounds.width), 0, bounds.width)
  const height = clamp(finite(rect.height, bounds.height), 0, bounds.height)
  return {
    x: clamp(finite(rect.x, 0), 0, bounds.width - width),
    y: clamp(finite(rect.y, 0), 0, bounds.height - height),
    width,
    height,
  }
}

/**
 * The box after dragging `handle` by (dx, dy) from `start`. Corner handles keep the opposite corner fixed
 * (or the centre with fromCenter); edge handles move one edge. With a ratio, a corner follows whichever
 * side grows the box more, and an edge changes the other side symmetrically (shifting the box to stay in
 * bounds). The result never leaves the bounds and never gets smaller than the minimum size.
 */
export function constrainDrag(start: Rect, handle: CropHandle, dx: number, dy: number, options: ConstrainOptions): Rect {
  const W = Math.max(0, options.bounds.width)
  const H = Math.max(0, options.bounds.height)
  const deltaX = finite(dx, 0)
  const deltaY = finite(dy, 0)
  if (handle === 'move') {
    return {
      x: clamp(start.x + deltaX, 0, Math.max(0, W - start.width)),
      y: clamp(start.y + deltaY, 0, Math.max(0, H - start.height)),
      width: start.width,
      height: start.height,
    }
  }
  const minimum = Math.min(options.minSize ?? MIN_CROP_SIZE, W, H)
  const west = handle.includes('w')
  const east = handle.includes('e')
  const north = handle.includes('n')
  const south = handle.includes('s')
  const horizontal = west || east
  const vertical = north || south
  const ratio = options.ratio && Number.isFinite(options.ratio) && options.ratio > 0 ? options.ratio : null
  const fromCenter = options.fromCenter
  const cx = start.x + start.width / 2
  const cy = start.y + start.height / 2
  const right = start.x + start.width
  const bottom = start.y + start.height

  let width = start.width
  let height = start.height
  if (horizontal) {
    if (fromCenter) width = 2 * (east ? right + deltaX - cx : cx - (start.x + deltaX))
    else width = east ? right + deltaX - start.x : right - (start.x + deltaX)
  }
  if (vertical) {
    if (fromCenter) height = 2 * (south ? bottom + deltaY - cy : cy - (start.y + deltaY))
    else height = south ? bottom + deltaY - start.y : bottom - (start.y + deltaY)
  }
  width = Math.max(0, width)
  height = Math.max(0, height)

  // Room on each axis: up to the bounds from the fixed edge (or centre); a derived axis may shift.
  const maxWidth = horizontal
    ? (fromCenter ? 2 * Math.min(cx, W - cx) : east ? W - start.x : right)
    : W
  const maxHeight = vertical
    ? (fromCenter ? 2 * Math.min(cy, H - cy) : south ? H - start.y : bottom)
    : H

  if (ratio) {
    if (horizontal && vertical) {
      if (height <= 0 || width / Math.max(height, 1e-9) > ratio) height = width / ratio
      else width = height * ratio
    } else if (horizontal) {
      height = width / ratio
    } else {
      width = height * ratio
    }
    let scale = Math.min(1, maxWidth / Math.max(width, 1e-9), maxHeight / Math.max(height, 1e-9))
    width *= scale
    height *= scale
    if (width < minimum || height < minimum) {
      const grow = Math.max(minimum / Math.max(width, 1e-9), minimum / Math.max(height, 1e-9))
      width *= grow
      height *= grow
    }
    scale = Math.min(1, maxWidth / Math.max(width, 1e-9), maxHeight / Math.max(height, 1e-9))
    width *= scale
    height *= scale
  } else {
    width = clamp(width, Math.min(minimum, maxWidth), Math.max(0, maxWidth))
    height = clamp(height, Math.min(minimum, maxHeight), Math.max(0, maxHeight))
  }

  let x: number
  let y: number
  if (horizontal) x = fromCenter ? cx - width / 2 : east ? start.x : right - width
  else x = clamp(cx - width / 2, 0, Math.max(0, W - width))
  if (vertical) y = fromCenter ? cy - height / 2 : south ? start.y : bottom - height
  else y = clamp(cy - height / 2, 0, Math.max(0, H - height))
  return clampRect({ x, y, width, height }, { width: W, height: H })
}

/** The handle a new box drag from `anchor` towards `point` behaves like. */
export function handleTowards(anchor: Point, point: Point): ResizeHandle {
  const east = point.x >= anchor.x
  const south = point.y >= anchor.y
  return south ? (east ? 'se' : 'sw') : (east ? 'ne' : 'nw')
}

/** A new box dragged from `anchor` to `point` (Alt: `anchor` is the centre), with the same constraints as a corner drag. */
export function rectFromPoints(anchor: Point, point: Point, options: ConstrainOptions): Rect {
  const start: Rect = { x: anchor.x, y: anchor.y, width: 0, height: 0 }
  return constrainDrag(start, handleTowards(anchor, point), point.x - anchor.x, point.y - anchor.y, options)
}

/**
 * The largest box of `ratio` inside the bounds, centred on the current box (shifted to stay inside).
 * Used when the user picks an aspect preset.
 */
export function fitAspect(current: Rect, ratio: number, bounds: Size): Rect {
  if (!(ratio > 0) || !(bounds.width > 0 && bounds.height > 0)) return clampRect(current, bounds)
  let width = bounds.width
  let height = width / ratio
  if (height > bounds.height) {
    height = bounds.height
    width = height * ratio
  }
  const cx = current.x + current.width / 2
  const cy = current.y + current.height / 2
  return clampRect({ x: cx - width / 2, y: cy - height / 2, width, height }, bounds)
}

/** Arrow-key nudge (the caller passes 1 px, or 10 px with Shift). */
export function nudgeRect(rect: Rect, dx: number, dy: number, bounds: Size): Rect {
  return constrainDrag(rect, 'move', dx, dy, { ratio: null, fromCenter: false, bounds })
}

/** Default box when the crop tool opens: 80% of the frame, centred (rounded like the original tool). */
export function defaultCropRect(bounds: Size): Rect {
  return {
    x: Math.round(bounds.width * 0.1),
    y: Math.round(bounds.height * 0.1),
    width: Math.max(1, Math.round(bounds.width * 0.8)),
    height: Math.max(1, Math.round(bounds.height * 0.8)),
  }
}

/** Whole-pixel crop rectangle inside the bounds (at least 1 x 1). */
export function toIntegerRect(rect: Rect, bounds: Size): IntRect {
  const W = Math.max(1, Math.floor(bounds.width))
  const H = Math.max(1, Math.floor(bounds.height))
  const x = clamp(Math.round(finite(rect.x, 0)), 0, W - 1)
  const y = clamp(Math.round(finite(rect.y, 0)), 0, H - 1)
  const width = clamp(Math.round(finite(rect.width, W)), 1, W - x)
  const height = clamp(Math.round(finite(rect.height, H)), 1, H - y)
  return { x, y, width, height }
}

/** True when two rectangles differ by more than `epsilon` in any coordinate. */
export function rectsDiffer(a: Rect | null, b: Rect | null, epsilon = 0.01): boolean {
  if (!a || !b) return a !== b
  return Math.abs(a.x - b.x) > epsilon || Math.abs(a.y - b.y) > epsilon
    || Math.abs(a.width - b.width) > epsilon || Math.abs(a.height - b.height) > epsilon
}

// ---------------------------------------------------------------------------------------------
// Straighten
// ---------------------------------------------------------------------------------------------

/** Folds any angle into (-45, 45] (a quarter turn is what Rotate left/right is for). */
export function foldAngle(degrees: number): number {
  if (!Number.isFinite(degrees)) return 0
  let d = ((degrees % 90) + 90) % 90
  if (d > 45) d -= 90
  return Math.round(d * 1000) / 1000
}

/** Clamps a slider value to [-45, 45] in 0.1 degree steps. */
export function clampStraighten(degrees: number): number {
  if (!Number.isFinite(degrees)) return 0
  return Math.round(clamp(degrees, -STRAIGHTEN_LIMIT, STRAIGHTEN_LIMIT) / STRAIGHTEN_STEP) * STRAIGHTEN_STEP
}

/**
 * The straighten angle after "Level by drawing a line" from `from` to `to` on the preview that is already
 * turned by `current` degrees: the line becomes horizontal (or vertical, whichever is nearer).
 */
export function levelAngle(from: Point, to: Point, current = 0): number {
  const dx = to.x - from.x
  const dy = to.y - from.y
  if (Math.hypot(dx, dy) < 1e-6) return clampStraighten(current)
  const lineDegrees = (Math.atan2(dy, dx) * 180) / Math.PI
  return clampStraighten(foldAngle(current - lineDegrees))
}

/** Size of the frame the crop box lives in: the image turned by `degrees` (the worker's 'expand' size). */
export function cropFrame(image: Size, degrees: number): Size {
  if (!degrees) return { width: image.width, height: image.height }
  return rotatedBounds(image.width, image.height, degrees)
}

/** Maps a frame point back to the unturned image (the inverse of the 'expand' rotation about the centres). */
export function frameToImage(point: Point, image: Size, degrees: number): Point {
  if (!degrees) return point
  const frame = cropFrame(image, degrees)
  const t = (degrees * Math.PI) / 180
  const cos = Math.cos(t)
  const sin = Math.sin(t)
  const px = point.x - frame.width / 2
  const py = point.y - frame.height / 2
  return { x: cos * px + sin * py + image.width / 2, y: -sin * px + cos * py + image.height / 2 }
}

/** True when every corner of `rect` (frame coordinates) lies inside the turned image, `margin` px from its edges. */
export function insideRotated(rect: Rect, image: Size, degrees: number, margin = STRAIGHTEN_DRAG_MARGIN): boolean {
  const corners: Point[] = [
    { x: rect.x, y: rect.y },
    { x: rect.x + rect.width, y: rect.y },
    { x: rect.x, y: rect.y + rect.height },
    { x: rect.x + rect.width, y: rect.y + rect.height },
  ]
  const epsilon = 1e-6
  for (const corner of corners) {
    const p = frameToImage(corner, image, degrees)
    if (p.x < margin - epsilon || p.y < margin - epsilon || p.x > image.width - margin + epsilon || p.y > image.height - margin + epsilon) return false
  }
  return true
}

/**
 * Auto-crop: the largest box of `ratio` (default: the image's own ratio) centred in the frame that stays
 * inside the image turned by `degrees`, kept `margin` px from its edges, in whole pixels.
 * For the original ratio this is design 4.2's s = min(W / (W|cos| + H|sin|), H / (W|sin| + H|cos|)).
 */
export function inscribedCrop(image: Size, degrees: number, ratio: number | null = null, margin = STRAIGHTEN_AUTO_MARGIN): Rect {
  const frame = cropFrame(image, degrees)
  const a = ratio && ratio > 0 ? ratio : image.width / Math.max(1e-9, image.height)
  const t = (degrees * Math.PI) / 180
  const c = Math.abs(Math.cos(t))
  const s = Math.abs(Math.sin(t))
  const usableWidth = Math.max(1, image.width - (degrees ? 2 * margin : 0))
  const usableHeight = Math.max(1, image.height - (degrees ? 2 * margin : 0))
  let height = Math.min(usableWidth / (a * c + s), usableHeight / (a * s + c))
  let width = a * height
  // Whole pixels, centred; shrink until the rounded box is inside (rounding can push a corner out).
  for (let attempt = 0; attempt < 64; attempt += 1) {
    const w = Math.max(1, Math.floor(width))
    const h = Math.max(1, Math.floor(height))
    const box: Rect = { x: Math.round((frame.width - w) / 2), y: Math.round((frame.height - h) / 2), width: w, height: h }
    if (!degrees || insideRotated(box, image, degrees, Math.max(0, margin - 0.5)) || (w <= 1 && h <= 1)) return clampRect(box, frame)
    width -= Math.max(1, a)
    height -= Math.max(1, 1 / a)
  }
  return clampRect({ x: frame.width / 2, y: frame.height / 2, width: 1, height: 1 }, frame)
}

function bisect(valid: (fraction: number) => boolean, low: number, high: number, steps = 22): number {
  let lo = low
  let hi = high
  for (let index = 0; index < steps; index += 1) {
    const mid = (lo + hi) / 2
    if (valid(mid)) lo = mid
    else hi = mid
  }
  return lo
}

/**
 * constrainDrag() that also keeps the box inside `inside` (the turned image while straightening). The drag
 * is shortened to the furthest valid point; free boxes and moves then slide along each axis separately so
 * the box can follow the pointer along a slanted edge.
 */
export function constrainDragInside(
  start: Rect,
  handle: CropHandle,
  dx: number,
  dy: number,
  options: ConstrainOptions,
  inside: (rect: Rect) => boolean,
): Rect {
  const full = constrainDrag(start, handle, dx, dy, options)
  if (inside(full)) return full
  if (!inside(start)) return start
  const at = (fx: number, fy: number) => constrainDrag(start, handle, dx * fx, dy * fy, options)
  const uniform = bisect((f) => inside(at(f, f)), 0, 1)
  if (handle !== 'move' && options.ratio) return at(uniform, uniform)
  const fx = bisect((f) => inside(at(f, uniform)), uniform, 1)
  const fy = bisect((f) => inside(at(fx, f)), uniform, 1)
  return at(fx, fy)
}

/** rectFromPoints() kept inside `inside`; null when even the smallest box at the anchor is outside. */
export function rectFromPointsInside(anchor: Point, point: Point, options: ConstrainOptions, inside: (rect: Rect) => boolean): Rect | null {
  const full = rectFromPoints(anchor, point, options)
  if (inside(full)) return full
  const at = (fraction: number) => rectFromPoints(anchor, { x: anchor.x + (point.x - anchor.x) * fraction, y: anchor.y + (point.y - anchor.y) * fraction }, options)
  if (!inside(at(0))) return null
  return at(bisect((fraction) => inside(at(fraction)), 0, 1))
}
