// src/advanced/viewport.ts (WP3)
// Pure view math for the Advanced canvas (design 5.6 and 5.15). DOM-free, so Node tests cover it.
//   - ViewTransform: screen (CSS px) = document px * zoom + offset. `zoom` is CSS px per document pixel.
//   - Display scale = zoom * dpr = device pixels per document pixel. Photoshop's "100%" is display scale 1
//     (one image pixel per device pixel), so every percentage shown to the user and the zoom ladder use the
//     display scale, never raw CSS zoom.
//   - Pyramid level: the coarsest level whose scale 2^-level is still >= the display scale, so a proxy is
//     never magnified (design 5.6). Level k tiles cover TILE_SIZE * 2^k document pixels.
import type { IntRect, Point, Rect, Size } from '../imaging/types.ts'
import type { ViewTransform, ViewportSize } from './types.ts'
import { TILE_SIZE } from './types.ts'

/** Photoshop's zoom presets as display scales (1 = 100%), smallest first. */
export const ZOOM_STEPS: readonly number[] = Object.freeze([
  0.01, 0.015, 0.02, 0.03, 0.04, 0.05, 0.0625, 1 / 12, 0.125, 1 / 6, 0.25, 1 / 3, 0.5, 2 / 3,
  1, 2, 3, 4, 5, 6, 7, 8, 12, 16, 24, 32,
])

/** Smallest and largest display scale (1% and 3200%). */
export const MIN_DISPLAY_SCALE = ZOOM_STEPS[0]
export const MAX_DISPLAY_SCALE = ZOOM_STEPS[ZOOM_STEPS.length - 1]

/** Relative tolerance used when comparing scales (keeps 1/3 and 0.3333 equal). */
const SCALE_EPSILON = 1e-6

function finite(value: number, fallback: number): number {
  return Number.isFinite(value) ? value : fallback
}

function positiveDpr(dpr: number): number {
  return Number.isFinite(dpr) && dpr > 0 ? dpr : 1
}

export const IDENTITY_VIEW: ViewTransform = Object.freeze({ zoom: 1, offsetX: 0, offsetY: 0 })

// ---------------------------------------------------------------------------------------------
// Scales and levels
// ---------------------------------------------------------------------------------------------

/** Device pixels per document pixel (Photoshop's zoom percentage / 100). */
export function displayScale(zoom: number, dpr: number): number {
  return finite(zoom, 1) * positiveDpr(dpr)
}

/** CSS zoom for a display scale on a screen with `dpr`. */
export function zoomForScale(scale: number, dpr: number): number {
  return finite(scale, 1) / positiveDpr(dpr)
}

/** Clamps a CSS zoom so its display scale stays within 1%..3200%. */
export function clampZoom(zoom: number, dpr: number): number {
  const ratio = positiveDpr(dpr)
  const scale = Math.min(MAX_DISPLAY_SCALE, Math.max(MIN_DISPLAY_SCALE, displayScale(zoom, ratio)))
  return scale / ratio
}

/**
 * Pyramid level for a view: clamp(floor(log2(1 / (zoom * dpr))), 0, maxLevel). Level 0 at 100% and above,
 * level 1 from 50% up to (not including) 100%, and so on; a proxy is never magnified.
 */
export function chooseLevel(zoom: number, dpr: number, maxLevel: number): number {
  const top = Math.max(0, Math.floor(finite(maxLevel, 0)))
  const scale = displayScale(zoom, dpr)
  if (!(scale > 0)) return top
  if (scale >= 1) return 0
  const level = Math.floor(Math.log2(1 / scale) + SCALE_EPSILON)
  return Math.min(top, Math.max(0, level))
}

/** Scale of pyramid level `level` relative to the document (1, 1/2, 1/4, ...). */
export function levelScale(level: number): number {
  return 2 ** -Math.max(0, Math.floor(finite(level, 0)))
}

/** Next ladder step above (direction 1) or below (-1) the current display scale, as a CSS zoom. */
export function stepZoom(zoom: number, dpr: number, direction: 1 | -1): number {
  const ratio = positiveDpr(dpr)
  const scale = displayScale(zoom, ratio)
  if (direction > 0) {
    for (const step of ZOOM_STEPS) if (step > scale * (1 + SCALE_EPSILON)) return step / ratio
    return MAX_DISPLAY_SCALE / ratio
  }
  for (let index = ZOOM_STEPS.length - 1; index >= 0; index -= 1) {
    const step = ZOOM_STEPS[index]
    if (step < scale * (1 - SCALE_EPSILON)) return step / ratio
  }
  return MIN_DISPLAY_SCALE / ratio
}

/** The ladder step nearest to a display scale (for snapping scrubby zoom and fit results). */
export function nearestZoomStep(scale: number): number {
  let best = ZOOM_STEPS[0]
  let bestDistance = Infinity
  const target = Math.log(Math.max(MIN_DISPLAY_SCALE, Math.min(MAX_DISPLAY_SCALE, finite(scale, 1))))
  for (const step of ZOOM_STEPS) {
    const distance = Math.abs(Math.log(step) - target)
    if (distance < bestDistance) {
      best = step
      bestDistance = distance
    }
  }
  return best
}

/** Photoshop-style percentage: "100%", "33.33%", "12.5%", "1600%". */
export function zoomLabel(zoom: number, dpr: number): string {
  const percent = displayScale(zoom, dpr) * 100
  const rounded = percent >= 100 ? Math.round(percent * 10) / 10 : Math.round(percent * 100) / 100
  return `${Number(rounded.toFixed(2)).toString()}%`
}

// ---------------------------------------------------------------------------------------------
// Coordinate mapping
// ---------------------------------------------------------------------------------------------

/** Document point to screen CSS px. */
export function docToScreen(view: ViewTransform, point: Point): Point {
  return { x: point.x * view.zoom + view.offsetX, y: point.y * view.zoom + view.offsetY }
}

/** Screen CSS px to document point (float). */
export function screenToDoc(view: ViewTransform, point: Point): Point {
  const zoom = view.zoom || 1
  return { x: (point.x - view.offsetX) / zoom, y: (point.y - view.offsetY) / zoom }
}

/** Screen rectangle (CSS px) of the document. */
export function documentScreenRect(view: ViewTransform, doc: Size): Rect {
  return { x: view.offsetX, y: view.offsetY, width: doc.width * view.zoom, height: doc.height * view.zoom }
}

/** Document-space rectangle (float) visible in the viewport. */
export function visibleDocRect(view: ViewTransform, viewport: Pick<ViewportSize, 'width' | 'height'>): Rect {
  const zoom = view.zoom || 1
  return {
    x: -view.offsetX / zoom,
    y: -view.offsetY / zoom,
    width: Math.max(0, viewport.width) / zoom,
    height: Math.max(0, viewport.height) / zoom,
  }
}

// ---------------------------------------------------------------------------------------------
// View operations (all return new views)
// ---------------------------------------------------------------------------------------------

/** Zoom to `zoom` (CSS px per document px) keeping the document point under `anchor` (CSS px) fixed. */
export function zoomViewAt(view: ViewTransform, zoom: number, anchor: Point): ViewTransform {
  const nextZoom = Number.isFinite(zoom) && zoom > 0 ? zoom : view.zoom
  const oldZoom = view.zoom || 1
  const ax = finite(anchor.x, 0)
  const ay = finite(anchor.y, 0)
  const docX = (ax - view.offsetX) / oldZoom
  const docY = (ay - view.offsetY) / oldZoom
  return { zoom: nextZoom, offsetX: ax - docX * nextZoom, offsetY: ay - docY * nextZoom }
}

/** Pans by (dx, dy) CSS px. */
export function panView(view: ViewTransform, dx: number, dy: number): ViewTransform {
  return { zoom: view.zoom, offsetX: view.offsetX + finite(dx, 0), offsetY: view.offsetY + finite(dy, 0) }
}

/** The document centred in the viewport at `zoom`. */
export function centerView(doc: Size, viewport: Pick<ViewportSize, 'width' | 'height'>, zoom: number): ViewTransform {
  return {
    zoom,
    offsetX: (viewport.width - doc.width * zoom) / 2,
    offsetY: (viewport.height - doc.height * zoom) / 2,
  }
}

/**
 * Fit on Screen (Ctrl+0): the largest zoom that shows the whole document inside the viewport minus
 * `padding` CSS px on each side, centred. Small documents are enlarged like Photoshop does, within
 * 1%..3200% display scale.
 */
export function fitView(doc: Size, viewport: ViewportSize, padding: number): ViewTransform {
  const pad = Math.max(0, finite(padding, 0))
  const width = Math.max(1, doc.width)
  const height = Math.max(1, doc.height)
  const availableWidth = Math.max(1, viewport.width - 2 * pad)
  const availableHeight = Math.max(1, viewport.height - 2 * pad)
  const zoom = clampZoom(Math.min(availableWidth / width, availableHeight / height), viewport.dpr)
  return centerView({ width, height }, viewport, zoom)
}

/** 100% (Ctrl+1): one document pixel per device pixel, keeping the point under `anchor` (default centre) fixed. */
export function actualPixelsView(view: ViewTransform, viewport: ViewportSize, anchor?: Point): ViewTransform {
  const point = anchor ?? { x: viewport.width / 2, y: viewport.height / 2 }
  return zoomViewAt(view, zoomForScale(1, viewport.dpr), point)
}

/** Rounds the offsets to whole device pixels so 100% views blit pixel-exact. */
export function snapView(view: ViewTransform, dpr: number): ViewTransform {
  const ratio = positiveDpr(dpr)
  return { zoom: view.zoom, offsetX: Math.round(view.offsetX * ratio) / ratio, offsetY: Math.round(view.offsetY * ratio) / ratio }
}

/**
 * Keeps at least `keep` CSS px of the document inside the viewport on each axis (or all of it when it is
 * smaller), so panning can never lose the image.
 */
export function constrainView(view: ViewTransform, doc: Size, viewport: Pick<ViewportSize, 'width' | 'height'>, keep = 64): ViewTransform {
  const margin = Math.max(0, finite(keep, 64))
  const docWidth = doc.width * view.zoom
  const docHeight = doc.height * view.zoom
  const keepX = Math.min(margin, docWidth)
  const keepY = Math.min(margin, docHeight)
  // The document's right edge must stay right of keepX; its left edge left of width - keepX.
  const minX = keepX - docWidth
  const maxX = viewport.width - keepX
  const minY = keepY - docHeight
  const maxY = viewport.height - keepY
  const offsetX = minX <= maxX ? Math.min(maxX, Math.max(minX, view.offsetX)) : (minX + maxX) / 2
  const offsetY = minY <= maxY ? Math.min(maxY, Math.max(minY, view.offsetY)) : (minY + maxY) / 2
  if (offsetX === view.offsetX && offsetY === view.offsetY) return view
  return { zoom: view.zoom, offsetX, offsetY }
}

export function viewsEqual(a: ViewTransform | null | undefined, b: ViewTransform | null | undefined): boolean {
  if (!a || !b) return a === b
  return a.zoom === b.zoom && a.offsetX === b.offsetX && a.offsetY === b.offsetY
}

// ---------------------------------------------------------------------------------------------
// Tiles on screen
// ---------------------------------------------------------------------------------------------

export interface LevelTileRange {
  /** Inclusive start, exclusive end, in level tiles. Empty when tx1 <= tx0 or ty1 <= ty0. */
  readonly tx0: number
  readonly ty0: number
  readonly tx1: number
  readonly ty1: number
}

/** Number of level-`level` tiles that cover a width x height document (columns, rows). */
export function levelTileCounts(doc: Size, level: number): { readonly columns: number; readonly rows: number } {
  const scale = 2 ** Math.max(0, Math.floor(finite(level, 0)))
  const width = Math.ceil(Math.max(0, doc.width) / scale)
  const height = Math.ceil(Math.max(0, doc.height) / scale)
  return { columns: Math.ceil(width / TILE_SIZE), rows: Math.ceil(height / TILE_SIZE) }
}

/**
 * Level tiles that intersect the viewport, grown by `margin` tiles on each side and clipped to the
 * document. Empty when the document is off screen.
 */
export function visibleTileRange(view: ViewTransform, viewport: Pick<ViewportSize, 'width' | 'height'>, doc: Size, level: number,
  margin = 0): LevelTileRange {
  const area = visibleDocRect(view, viewport)
  const span = TILE_SIZE * 2 ** Math.max(0, Math.floor(finite(level, 0)))
  const ring = Math.max(0, Math.floor(finite(margin, 0)))
  const counts = levelTileCounts(doc, level)
  if (!(area.width > 0 && area.height > 0) || !counts.columns || !counts.rows) return { tx0: 0, ty0: 0, tx1: 0, ty1: 0 }
  // Half-open: a viewport edge exactly on a tile boundary does not pull in the next tile.
  const tx0 = Math.floor(area.x / span) - ring
  const ty0 = Math.floor(area.y / span) - ring
  const tx1 = Math.ceil((area.x + area.width) / span - 1e-9) + ring
  const ty1 = Math.ceil((area.y + area.height) / span - 1e-9) + ring
  return {
    tx0: Math.max(0, tx0),
    ty0: Math.max(0, ty0),
    tx1: Math.min(counts.columns, Math.max(0, tx1)),
    ty1: Math.min(counts.rows, Math.max(0, ty1)),
  }
}

export function isEmptyTileRange(range: LevelTileRange): boolean {
  return range.tx1 <= range.tx0 || range.ty1 <= range.ty0
}

/** Level-pixel rectangle of level tile (tx, ty) clipped to the document at that level. */
export function levelTileRect(doc: Size, level: number, tx: number, ty: number): IntRect {
  const scale = 2 ** Math.max(0, Math.floor(finite(level, 0)))
  const width = Math.ceil(Math.max(0, doc.width) / scale)
  const height = Math.ceil(Math.max(0, doc.height) / scale)
  const x = tx * TILE_SIZE
  const y = ty * TILE_SIZE
  return { x, y, width: Math.max(0, Math.min(TILE_SIZE, width - x)), height: Math.max(0, Math.min(TILE_SIZE, height - y)) }
}

/**
 * Level pixels that a document-space change can alter at `level`: the scaled rectangle, one pixel wider
 * on the low side because a layer offset rounds down separately from its pixels (floor(a / n) +
 * floor(b / n) can be one less than floor((a + b) / n)). Null for an empty rectangle.
 */
export function dirtyLevelRect(rect: IntRect, level: number): IntRect | null {
  if (!(rect.width > 0 && rect.height > 0)) return null
  const scale = 2 ** Math.max(0, Math.floor(finite(level, 0)))
  const slack = scale > 1 ? 1 : 0
  const x0 = Math.floor(rect.x / scale) - slack
  const y0 = Math.floor(rect.y / scale) - slack
  const x1 = Math.ceil((rect.x + rect.width) / scale)
  const y1 = Math.ceil((rect.y + rect.height) / scale)
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
}

/** Level tiles that a document-space change touches at `level` (see dirtyLevelRect). Null when empty. */
export function dirtyTileRange(rect: IntRect, level: number): LevelTileRange | null {
  const area = dirtyLevelRect(rect, level)
  if (!area) return null
  return {
    tx0: Math.floor(area.x / TILE_SIZE),
    ty0: Math.floor(area.y / TILE_SIZE),
    tx1: Math.floor((area.x + area.width - 1) / TILE_SIZE) + 1,
    ty1: Math.floor((area.y + area.height - 1) / TILE_SIZE) + 1,
  }
}

/**
 * Tiles of `range` ordered for rendering: nearest the viewport centre first (Euclidean distance between
 * tile centres and the centre, in tile units), ties broken top-to-bottom then left-to-right.
 */
export function tilesByDistance(range: LevelTileRange, center: Point): { readonly tx: number; readonly ty: number }[] {
  const out: { tx: number; ty: number; d: number }[] = []
  for (let ty = range.ty0; ty < range.ty1; ty += 1) {
    for (let tx = range.tx0; tx < range.tx1; tx += 1) {
      const dx = tx + 0.5 - center.x
      const dy = ty + 0.5 - center.y
      out.push({ tx, ty, d: dx * dx + dy * dy })
    }
  }
  out.sort((a, b) => a.d - b.d || a.ty - b.ty || a.tx - b.tx)
  return out.map(({ tx, ty }) => ({ tx, ty }))
}
