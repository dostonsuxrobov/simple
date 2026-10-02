// src/advanced/selection.ts (WP4)
// The Advanced selection model (design 5.8). A Selection is a dense document-size coverage mask plus its
// tight bounds and a version; null means "no selection" (everything editable). Selections are immutable
// values: every function returns a new Selection (or null) and never writes into a mask it was given,
// because document states, history snapshots and previews may still reference the old one.
// History stores compact SelectionSnapshots ('none', 'all' or just the bounds region).
// Pure and DOM-free (Node tests import it directly).
import type { IntRect, MaskBuffer, PixelBuffer, SelectionOp } from '../imaging/types.ts'
import type { Selection, SelectionSnapshot } from './types.ts'
import { clipRect, combineMasks, createMaskBuffer, cropMask, maskBounds, pasteMask } from '../imaging/mask.ts'

function checkSize(width: number, height: number): void {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width <= 0 || height <= 0) {
    throw new RangeError(`Invalid selection size ${width} x ${height}.`)
  }
}

function checkMask(mask: MaskBuffer): void {
  if (!mask || !mask.data || !Number.isInteger(mask.width) || !Number.isInteger(mask.height)
    || mask.data.length !== mask.width * mask.height) {
    throw new RangeError('The selection mask does not match its size.')
  }
}

function copyMask(mask: MaskBuffer): MaskBuffer {
  return { width: mask.width, height: mask.height, data: new Uint8Array(mask.data) }
}

function intersect(a: IntRect, b: IntRect): IntRect | null {
  const x0 = Math.max(a.x, b.x)
  const y0 = Math.max(a.y, b.y)
  const x1 = Math.min(a.x + a.width, b.x + b.width)
  const y1 = Math.min(a.y + a.height, b.y + b.height)
  return x1 > x0 && y1 > y0 ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null
}

function union(a: IntRect, b: IntRect): IntRect {
  const x0 = Math.min(a.x, b.x)
  const y0 = Math.min(a.y, b.y)
  return { x: x0, y: y0, width: Math.max(a.x + a.width, b.x + b.width) - x0, height: Math.max(a.y + a.height, b.y + b.height) - y0 }
}

/** True when every byte equals `value` (word-at-a-time where the view allows it). */
function allEqual(data: Uint8Array, value: number): boolean {
  let start = 0
  if (data.byteOffset % 4 === 0 && data.length >= 64) {
    const words = data.length >> 2
    const view = new Uint32Array(data.buffer, data.byteOffset, words)
    const word = (value * 0x01010101) >>> 0
    for (let i = 0; i < words; i += 1) if (view[i] !== word) return false
    start = words << 2
  }
  for (let i = start; i < data.length; i += 1) if (data[i] !== value) return false
  return true
}

/** Wraps a document-size mask (taken over, not copied). Null when it selects nothing. */
export function selectionFromMask(mask: MaskBuffer, version: number): Selection | null {
  checkMask(mask)
  const bounds = maskBounds(mask)
  return bounds ? { mask, bounds, version } : null
}

/** Select > All. */
export function selectAll(width: number, height: number, version: number): Selection {
  checkSize(width, height)
  return { mask: createMaskBuffer(width, height, 255), bounds: { x: 0, y: 0, width, height }, version }
}

/** True when the selection covers the whole document fully. */
export function isSelectAll(selection: Selection | null): boolean {
  if (!selection) return false
  const { mask, bounds } = selection
  if (bounds.x !== 0 || bounds.y !== 0 || bounds.width !== mask.width || bounds.height !== mask.height) return false
  return allEqual(mask.data, 255)
}

/**
 * Combines a tool shape (a document-size scratch mask, left untouched) with the current selection:
 * replace, add (max), subtract (a * (255 - b) / 255) or intersect (a * b / 255). With no current
 * selection, add behaves like replace and subtract / intersect select nothing. Returns null when nothing
 * stays selected, and `current` itself when the shape is empty and the operation keeps it.
 */
export function applySelectionOp(current: Selection | null, shape: MaskBuffer, op: SelectionOp, version: number): Selection | null {
  checkMask(shape)
  if (current && (current.mask.width !== shape.width || current.mask.height !== shape.height)) {
    throw new RangeError(`The ${shape.width} x ${shape.height} shape does not match the ${current.mask.width} x ${current.mask.height} selection.`)
  }
  const shapeBounds = maskBounds(shape)
  switch (op) {
    case 'replace':
      return shapeBounds ? { mask: copyMask(shape), bounds: shapeBounds, version } : null
    case 'add': {
      if (!current) return shapeBounds ? { mask: copyMask(shape), bounds: shapeBounds, version } : null
      if (!shapeBounds) return current
      const mask = copyMask(current.mask)
      combineMasks(mask, shape, 'add', shapeBounds)
      return { mask, bounds: union(current.bounds, shapeBounds), version }
    }
    case 'subtract': {
      if (!current) return null
      if (!shapeBounds) return current
      const overlap = intersect(current.bounds, shapeBounds)
      if (!overlap) return current
      const mask = copyMask(current.mask)
      combineMasks(mask, shape, 'subtract', overlap)
      const bounds = maskBounds(mask, 1, current.bounds)
      return bounds ? { mask, bounds, version } : null
    }
    case 'intersect': {
      if (!current || !shapeBounds) return null
      const overlap = intersect(current.bounds, shapeBounds)
      if (!overlap) return null
      const mask = createMaskBuffer(shape.width, shape.height)
      const a = current.mask.data
      const b = shape.data
      const width = shape.width
      for (let y = overlap.y; y < overlap.y + overlap.height; y += 1) {
        const start = y * width + overlap.x
        for (let i = start; i < start + overlap.width; i += 1) {
          const t = a[i] * b[i] + 128
          mask.data[i] = (t + (t >> 8)) >> 8
        }
      }
      const bounds = maskBounds(mask, 1, overlap)
      return bounds ? { mask, bounds, version } : null
    }
    default:
      throw new RangeError(`Unknown selection operation "${String(op)}".`)
  }
}

/** Select > Inverse. Null when the selection covered everything. */
export function invertSelection(selection: Selection, version: number): Selection | null {
  const mask = copyMask(selection.mask)
  const data = mask.data
  for (let i = 0; i < data.length; i += 1) data[i] = 255 - data[i]
  return selectionFromMask(mask, version)
}

/** Compact history form: 'none', 'all', or a copy of the bounds region only. */
export function snapshotSelection(selection: Selection | null): SelectionSnapshot {
  if (!selection) return { kind: 'none' }
  if (isSelectAll(selection)) return { kind: 'all' }
  const rect = selection.bounds
  return { kind: 'region', rect: { ...rect }, data: cropMask(selection.mask, rect).data }
}

/**
 * Rebuilds a selection for a width x height document. A region outside the (possibly resized)
 * document is clipped; null when nothing remains.
 */
export function restoreSelection(snapshot: SelectionSnapshot, width: number, height: number, version: number): Selection | null {
  checkSize(width, height)
  switch (snapshot.kind) {
    case 'none':
      return null
    case 'all':
      return selectAll(width, height, version)
    case 'region': {
      const { rect, data } = snapshot
      if (data.length !== rect.width * rect.height) throw new RangeError('The selection snapshot is damaged.')
      const clip = clipRect(rect, width, height)
      if (!clip) return null
      const mask = createMaskBuffer(width, height)
      pasteMask(mask, { width: rect.width, height: rect.height, data }, rect.x, rect.y)
      const bounds = maskBounds(mask, 1, clip)
      return bounds ? { mask, bounds, version } : null
    }
    default:
      throw new RangeError('Unknown selection snapshot.')
  }
}

/** Bytes a snapshot keeps alive in history. */
export function snapshotBytes(snapshot: SelectionSnapshot): number {
  return snapshot.kind === 'region' ? snapshot.data.byteLength + 64 : 16
}

/**
 * Moves the selection by whole pixels (dx, dy are rounded); coverage moved past the document edge is
 * lost. Null when nothing remains on the canvas.
 */
export function translateSelection(selection: Selection, dx: number, dy: number, version: number): Selection | null {
  const ox = Math.round(Number.isFinite(dx) ? dx : 0)
  const oy = Math.round(Number.isFinite(dy) ? dy : 0)
  const { mask, bounds } = selection
  if (ox === 0 && oy === 0) return { mask, bounds, version }
  const moved = clipRect({ x: bounds.x + ox, y: bounds.y + oy, width: bounds.width, height: bounds.height }, mask.width, mask.height)
  if (!moved) return null
  const out = createMaskBuffer(mask.width, mask.height)
  for (let y = moved.y; y < moved.y + moved.height; y += 1) {
    const from = (y - oy) * mask.width + (moved.x - ox)
    out.data.set(mask.data.subarray(from, from + moved.width), y * mask.width + moved.x)
  }
  const tight = maskBounds(out, 1, moved)
  return tight ? { mask: out, bounds: tight, version } : null
}

/**
 * Selection coverage over a document-space rectangle (which may extend past the canvas): 255 everywhere
 * without a selection (everything editable), otherwise the mask with 0 outside the document.
 */
export function selectionCoverage(selection: Selection | null, rect: IntRect): MaskBuffer {
  const width = Math.max(0, Math.floor(rect.width))
  const height = Math.max(0, Math.floor(rect.height))
  if (!selection) return createMaskBuffer(width, height, 255)
  return cropMask(selection.mask, { x: Math.floor(rect.x), y: Math.floor(rect.y), width, height })
}

/**
 * Select > Load Selection from a layer (Ctrl+click on its thumbnail): the layer's alpha, placed at its
 * document offset and clipped to the width x height document.
 */
export function selectionFromAlpha(pixels: PixelBuffer, offsetX: number, offsetY: number, width: number, height: number,
  version: number): Selection | null {
  checkSize(width, height)
  if (!pixels || !pixels.data || pixels.data.length !== pixels.width * pixels.height * 4) {
    throw new RangeError('The pixel buffer does not match its size.')
  }
  const ox = Math.round(offsetX)
  const oy = Math.round(offsetY)
  const clip = clipRect({ x: ox, y: oy, width: pixels.width, height: pixels.height }, width, height)
  if (!clip) return null
  const mask = createMaskBuffer(width, height)
  for (let y = clip.y; y < clip.y + clip.height; y += 1) {
    const row = y * width
    let p = ((y - oy) * pixels.width + (clip.x - ox)) * 4 + 3
    for (let x = clip.x; x < clip.x + clip.width; x += 1, p += 4) mask.data[row + x] = pixels.data[p]
  }
  const bounds = maskBounds(mask, 1, clip)
  return bounds ? { mask, bounds, version } : null
}
