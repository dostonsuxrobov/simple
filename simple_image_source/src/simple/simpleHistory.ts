// src/simple/simpleHistory.ts (WP8)
// Simple mode's undo/redo store (design 4.10). The old store kept a full-canvas ImageData per step and
// capped the count by `160 MiB / image bytes`, so a 24 MP photo had ONE undo step and every brush stroke
// copied the whole image. Here:
//   - the budget is bytes (512 MiB), not entries; the newest entry is always kept, even when it alone is
//     larger than the budget;
//   - strokes, markup and other local edits store 256 px tile patches captured right before the first
//     draw inside each tile (a stroke over 4 tiles costs 1 MiB, whatever the image size);
//   - exact quarter turns and flips store no pixels at all: undo applies the inverse operation;
//   - geometry changes (crop, resize, straighten) and whole-image edits store one full snapshot.
// Undo and redo SWAP the stored pixels with the current ones, so an entry moves between the two stacks
// without copying. Every entry carries the document state on both sides ({ revision, hasAlpha, pristine })
// so the Modified badge, the inspector and "unchanged since open" stay exact through undo and redo.
// The stacks are plain arrays (main.tsx keeps them in refs; WP1's PSD open resets them by assignment).
// DOM-free: the canvas is reached through HistorySurface (src/simple/ops.ts), Node tests use a fake.
import type { IntRect, PixelBuffer } from '../imaging/types.ts'

export const SIMPLE_HISTORY_BUDGET_BYTES = 512 * 1024 * 1024
export const SIMPLE_HISTORY_MAX_ENTRIES = 200
export const PATCH_TILE_SIZE = 256

/** Document state on one side of an entry. */
export interface HistoryState {
  /** Content revision (drives Modified / Saved). */
  readonly revision: number
  /** Whether the image has any pixel with alpha below 255 (inspector, JPEG warning). */
  readonly hasAlpha: boolean
  /** The pixels are exactly the file as opened, so Save may write the original bytes. */
  readonly pristine: boolean
}

/** Exact, pixel-permuting geometry operations; their inverse undoes them without stored pixels. */
export type GeometryOp = 'rotate-cw' | 'rotate-ccw' | 'flip-h' | 'flip-v'

/** What the history needs from the image store (the Simple canvas). */
export interface HistorySurface {
  readonly width: number
  readonly height: number
  /** Copies a region out (straight RGBA). */
  read(rect: IntRect): PixelBuffer
  /** Writes pixels at (x, y), replacing what is there. */
  write(x: number, y: number, pixels: PixelBuffer): void
  /** Replaces the whole image; the size may change. */
  replace(pixels: PixelBuffer): void
  /** Applies an exact geometry operation in place. */
  transform(op: GeometryOp): void
}

interface EntryBase {
  readonly label: string
  readonly before: HistoryState
  readonly after: HistoryState
}

/** The other side's whole image (swapped on undo and redo). */
export interface SnapshotEntry extends EntryBase {
  readonly kind: 'snapshot'
  pixels: PixelBuffer
}

export interface TilePatch {
  readonly x: number
  readonly y: number
  /** The other side's pixels of this tile (swapped on undo and redo). */
  pixels: PixelBuffer
}

/** Tiles of a same-size edit (strokes, markup, partial edits). */
export interface PatchEntry extends EntryBase {
  readonly kind: 'patch'
  readonly width: number
  readonly height: number
  readonly tiles: TilePatch[]
}

/** A quarter turn or flip; undo applies the inverse. */
export interface GeometryEntry extends EntryBase {
  readonly kind: 'geometry'
  readonly op: GeometryOp
}

export type HistoryEntry = SnapshotEntry | PatchEntry | GeometryEntry

export interface HistoryStacks {
  readonly undo: HistoryEntry[]
  readonly redo: HistoryEntry[]
}

export function inverseGeometry(op: GeometryOp): GeometryOp {
  if (op === 'rotate-cw') return 'rotate-ccw'
  if (op === 'rotate-ccw') return 'rotate-cw'
  return op
}

/** Size of a geometry operation's result. */
export function geometrySize(width: number, height: number, op: GeometryOp): { width: number; height: number } {
  return op === 'rotate-cw' || op === 'rotate-ccw' ? { width: height, height: width } : { width, height }
}

export function entryBytes(entry: HistoryEntry): number {
  if (entry.kind === 'snapshot') return entry.pixels.data.byteLength
  if (entry.kind === 'patch') {
    let total = 0
    for (const tile of entry.tiles) total += tile.pixels.data.byteLength
    return total
  }
  return 0
}

export function historyBytes(stacks: HistoryStacks): number {
  let total = 0
  for (const entry of stacks.undo) total += entryBytes(entry)
  for (const entry of stacks.redo) total += entryBytes(entry)
  return total
}

/** Drops the oldest undo entries until the stacks fit the budget; the newest undo entry always stays. */
export function trimHistory(stacks: HistoryStacks, budget = SIMPLE_HISTORY_BUDGET_BYTES, maxEntries = SIMPLE_HISTORY_MAX_ENTRIES): number {
  let removed = 0
  let total = historyBytes(stacks)
  while (stacks.undo.length > 1 && (total > budget || stacks.undo.length > maxEntries)) {
    const oldest = stacks.undo.shift() as HistoryEntry
    total -= entryBytes(oldest)
    removed += 1
  }
  return removed
}

/** Records a new step: the redo stack is dropped, then the budget is enforced. */
export function pushEntry(stacks: HistoryStacks, entry: HistoryEntry, budget = SIMPLE_HISTORY_BUDGET_BYTES): void {
  stacks.redo.length = 0
  stacks.undo.push(entry)
  trimHistory(stacks, budget)
}

export function fullRect(surface: Pick<HistorySurface, 'width' | 'height'>): IntRect {
  return { x: 0, y: 0, width: surface.width, height: surface.height }
}

export function snapshotEntry(label: string, pixels: PixelBuffer, before: HistoryState, after: HistoryState): SnapshotEntry {
  return { kind: 'snapshot', label, pixels, before, after }
}

export function geometryEntry(label: string, op: GeometryOp, before: HistoryState, after: HistoryState): GeometryEntry {
  return { kind: 'geometry', label, op, before, after }
}

/** Collects the pre-edit pixels of every tile an edit is about to touch. */
export interface PatchRecorder {
  /** Captures the tiles under `rect` that were not captured yet. Call BEFORE drawing there. */
  touch(rect: IntRect): void
  readonly tileCount: number
  readonly bytes: number
  /** The entry, or null when nothing was touched. The recorder must not be used afterwards. */
  finish(label: string, before: HistoryState, after: HistoryState): PatchEntry | null
}

export function tileKey(tx: number, ty: number): number {
  return ty * 0x10000 + tx
}

/** Integer rectangle of `rect` clipped to the surface, or null when nothing is left. */
export function clipRect(rect: IntRect, width: number, height: number): IntRect | null {
  const x0 = Math.max(0, Math.floor(rect.x))
  const y0 = Math.max(0, Math.floor(rect.y))
  const x1 = Math.min(width, Math.ceil(rect.x + rect.width))
  const y1 = Math.min(height, Math.ceil(rect.y + rect.height))
  if (!(x1 > x0 && y1 > y0)) return null
  return { x: x0, y: y0, width: x1 - x0, height: y1 - y0 }
}

export function createPatchRecorder(surface: HistorySurface, tileSize = PATCH_TILE_SIZE): PatchRecorder {
  const width = surface.width
  const height = surface.height
  const tiles = new Map<number, TilePatch>()
  let bytes = 0
  let finished = false
  return {
    touch(rect) {
      if (finished) throw new Error('This edit was already recorded.')
      if (surface.width !== width || surface.height !== height) throw new Error('The image size changed during an edit.')
      const clipped = clipRect(rect, width, height)
      if (!clipped) return
      const tx0 = Math.floor(clipped.x / tileSize)
      const ty0 = Math.floor(clipped.y / tileSize)
      const tx1 = Math.floor((clipped.x + clipped.width - 1) / tileSize)
      const ty1 = Math.floor((clipped.y + clipped.height - 1) / tileSize)
      for (let ty = ty0; ty <= ty1; ty += 1) {
        for (let tx = tx0; tx <= tx1; tx += 1) {
          const key = tileKey(tx, ty)
          if (tiles.has(key)) continue
          const x = tx * tileSize
          const y = ty * tileSize
          const pixels = surface.read({ x, y, width: Math.min(tileSize, width - x), height: Math.min(tileSize, height - y) })
          tiles.set(key, { x, y, pixels })
          bytes += pixels.data.byteLength
        }
      }
    },
    get tileCount() { return tiles.size },
    get bytes() { return bytes },
    finish(label, before, after) {
      finished = true
      if (!tiles.size) return null
      return { kind: 'patch', label, width, height, tiles: [...tiles.values()], before, after }
    },
  }
}

function swap(entry: HistoryEntry, surface: HistorySurface, direction: 'undo' | 'redo'): void {
  if (entry.kind === 'geometry') {
    surface.transform(direction === 'undo' ? inverseGeometry(entry.op) : entry.op)
    return
  }
  if (entry.kind === 'snapshot') {
    const current = surface.read(fullRect(surface))
    surface.replace(entry.pixels)
    entry.pixels = current
    return
  }
  if (surface.width !== entry.width || surface.height !== entry.height) {
    throw new Error('The undo history no longer matches the image.')
  }
  // Read every tile first so a failure part-way leaves the image untouched.
  const current = entry.tiles.map((tile) => surface.read({ x: tile.x, y: tile.y, width: tile.pixels.width, height: tile.pixels.height }))
  entry.tiles.forEach((tile, index) => {
    surface.write(tile.x, tile.y, tile.pixels)
    tile.pixels = current[index]
  })
}

/** Undoes the newest step; returns it (apply `entry.before`), or null when there is nothing to undo. */
export function undoEntry(stacks: HistoryStacks, surface: HistorySurface): HistoryEntry | null {
  const entry = stacks.undo.pop()
  if (!entry) return null
  try {
    swap(entry, surface, 'undo')
  } catch (error) {
    stacks.undo.push(entry)
    throw error
  }
  stacks.redo.push(entry)
  return entry
}

/** Redoes the newest undone step; returns it (apply `entry.after`), or null when there is nothing to redo. */
export function redoEntry(stacks: HistoryStacks, surface: HistorySurface): HistoryEntry | null {
  const entry = stacks.redo.pop()
  if (!entry) return null
  try {
    swap(entry, surface, 'redo')
  } catch (error) {
    stacks.redo.push(entry)
    throw error
  }
  stacks.undo.push(entry)
  return entry
}
