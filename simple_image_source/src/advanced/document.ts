// src/advanced/document.ts (WP3)
// The Advanced document model and store (design 5.3).
//   - DocumentState is immutable metadata (layers bottom to top, like ag-psd); pixel contents mutate in place
//     inside surfaces whose identity is stable. React reads it with useSyncExternalStore(subscribe, getState).
//   - Every mutation goes through store.transact(label, icon, tx => ...) or store.beginStroke(...). The store
//     records inverse operations (history.ts), so tools and commands never touch history directly.
//     A transaction is atomic: tx.rollback() or an exception reverts everything it did and records nothing.
//   - Revisions (the Modified badge): a step that changes the output allocates a new host revision; selection
//     and active-layer changes do not (detected from the recorded operations, so a mislabelled call can never
//     leave a real edit without a revision). A stroke marks the document modified at its first pixel and
//     takes a fresh revision when it commits, so a Save in the middle of a stroke is never mistaken for a
//     save of the finished stroke.
//   - Safety: starting a transaction, undo/redo or another stroke first commits an open stroke; editors of
//     finished sessions become inert; editor.surface is a guarded view whose write()/setTile() snapshot the
//     touched tiles for undo automatically.
// Pure and DOM-free at module load (Node tests import it).
import type { AdjustmentSpec, IntRect, MaskBuffer, PixelBuffer, Point } from '../imaging/types.ts'
import type {
  AdjustmentLayer,
  AdvancedHost,
  DocumentChange,
  DocumentState,
  DocumentStore,
  EditTarget,
  HistoryIcon,
  Layer,
  LayerCommon,
  LayerId,
  LayerLocks,
  LayerMask,
  LayerPatch,
  MemoryUsage,
  PixelEditor,
  RasterCache,
  RasterLayer,
  Selection,
  ShapeLayer,
  ShapeSpec,
  StrokeSession,
  TextLayer,
  TextSpec,
  TiledMask,
  TiledSurface,
  Transaction,
  TransactOptions,
} from './types.ts'
import { BACKGROUND_LOCKS, DEFAULT_LOCKS, LIMITS, TILE_SIZE } from './types.ts'
import { isBlendMode } from '../imaging/blend.ts'
import { cropMask } from '../imaging/mask.ts'
import type { ActiveRef, HistoryController, HistoryEntry, HistoryOp, TilesOp } from './history.ts'
import { createHistory, opAffectsOutput } from './history.ts'
import {
  createMaskSurface,
  createSurface,
  isTiledMask,
  isTiledSurface,
  maskFromBuffer,
  surfaceFromBuffer,
  tileCoords,
  tileKey,
  tileRange,
  tileRect,
  toIntRect,
} from './tiles.ts'
import { restoreSelection, snapshotSelection } from './selection.ts'
import { documentBytes, usageOf } from './memory.ts'
import { rasterizeShape, rasterizeText, rasterizeTextSync, specKey, wholePixelShift } from '../shared/vector.ts'
import type { RasterResult } from '../shared/vector.ts'

type Mutable<T> = { -readonly [K in keyof T]: T[K] }
type TileData = Uint8ClampedArray | Uint8Array

const MAX_DIRTY_RECTS = 64

// ---------------------------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------------------------

function clamp01(value: number): number {
  return value >= 1 ? 1 : value > 0 ? value : 0
}

function isBackgroundLayer(layer: Layer | undefined): boolean {
  return Boolean(layer && layer.kind === 'raster' && layer.isBackground)
}

function union(a: IntRect, b: IntRect): IntRect {
  const x = Math.min(a.x, b.x)
  const y = Math.min(a.y, b.y)
  return { x, y, width: Math.max(a.x + a.width, b.x + b.width) - x, height: Math.max(a.y + a.height, b.y + b.height) - y }
}

/** Collects the document areas a change made stale (collapses to one box past 64 rectangles). */
class DirtySet {
  all = false
  rects: IntRect[] = []

  add(value: IntRect | 'all' | null | undefined): void {
    if (this.all || !value) return
    if (value === 'all') {
      this.all = true
      this.rects = []
      return
    }
    if (!(value.width > 0 && value.height > 0)) return
    this.rects.push(value)
    if (this.rects.length > MAX_DIRTY_RECTS) this.rects = [this.rects.reduce(union)]
  }

  get empty(): boolean {
    return !this.all && this.rects.length === 0
  }

  result(): readonly IntRect[] | 'all' {
    return this.all ? 'all' : this.rects.slice()
  }
}

interface ChangeAccumulator {
  readonly dirty: DirtySet
  structure: boolean
  selection: boolean
}

function indexOfLayer(state: Pick<DocumentState, 'layers'>, id: LayerId | null | undefined): number {
  if (id === null || id === undefined) return -1
  return state.layers.findIndex((layer) => layer.id === id)
}

function requireIndex(state: Pick<DocumentState, 'layers'>, id: LayerId): number {
  const index = indexOfLayer(state, id)
  if (index < 0) throw new Error(`The layer "${String(id)}" does not exist.`)
  return index
}

function withLayerAt(state: DocumentState, index: number, layer: Layer): DocumentState {
  const layers = state.layers.slice()
  layers[index] = layer
  return { ...state, layers }
}

function corrupt(what: string): Error {
  return new Error(`The document history is inconsistent (${what}).`)
}

/** Document-space bounds of a layer's pixels (raster surface or text/shape raster cache), or null. */
export function layerContentBounds(layer: Layer): IntRect | null {
  let bounds: IntRect | null = null
  let ox = 0
  let oy = 0
  if (layer.kind === 'raster') {
    bounds = layer.surface.contentBounds()
    ox = layer.offsetX
    oy = layer.offsetY
  } else if (layer.kind === 'text' || layer.kind === 'shape') {
    bounds = layer.raster.surface.contentBounds()
    ox = layer.raster.offsetX
    oy = layer.raster.offsetY
  }
  return bounds ? { x: bounds.x + ox, y: bounds.y + oy, width: bounds.width, height: bounds.height } : null
}

/** Area whose composite depends on `layer` ('all' for adjustment layers, which affect everything below). */
function layerDirty(layer: Layer): IntRect | 'all' | null {
  return layer.kind === 'adjustment' ? 'all' : layerContentBounds(layer)
}

/** Document position of an edited surface, or null when it is no longer the layer's live surface. */
function surfaceOrigin(state: Pick<DocumentState, 'layers'>, layerId: LayerId, target: EditTarget, surface: TiledSurface | TiledMask): Point | null {
  const layer = state.layers.find((item) => item.id === layerId)
  if (!layer) return null
  if (target === 'mask') return layer.mask && layer.mask.surface === surface ? { x: layer.mask.offsetX, y: layer.mask.offsetY } : null
  if (layer.kind === 'raster' && layer.surface === surface) return { x: layer.offsetX, y: layer.offsetY }
  return null
}

function setTileAny(surface: TiledSurface | TiledMask, tx: number, ty: number, data: TileData | null | undefined): void {
  (surface as { setTile(tx: number, ty: number, data: TileData | undefined): void }).setTile(tx, ty, data ?? undefined)
}

/** Swaps stored and live tiles (undo and redo of a pixel edit are the same operation). */
function swapTiles(op: TilesOp, origin: Point | null, dirty: DirtySet | null): void {
  for (const [key, stored] of [...op.stored]) {
    const { tx, ty } = tileCoords(key)
    const live = op.surface.tile(tx, ty) as TileData | undefined
    setTileAny(op.surface, tx, ty, stored)
    op.stored.set(key, live ?? null)
    if (origin && dirty) dirty.add({ x: origin.x + tx * TILE_SIZE, y: origin.y + ty * TILE_SIZE, width: TILE_SIZE, height: TILE_SIZE })
  }
}

// ---------------------------------------------------------------------------------------------
// Validation
// ---------------------------------------------------------------------------------------------

const LAYER_KINDS = new Set(['raster', 'text', 'shape', 'adjustment'])

function validateCanvas(width: number, height: number): void {
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
    throw new RangeError(`Invalid document size ${width} x ${height}.`)
  }
  if (width > LIMITS.maxDimension || height > LIMITS.maxDimension) {
    throw new RangeError(`Images can be at most ${LIMITS.maxDimension.toLocaleString('en-US')} pixels on each side.`)
  }
  if (width * height > LIMITS.maxPixels) {
    throw new RangeError(`Images can have at most ${Math.round(LIMITS.maxPixels / 1_000_000)} megapixels.`)
  }
}

function validateMask(mask: LayerMask): void {
  if (!mask || typeof mask !== 'object' || !isTiledMask(mask.surface)) throw new TypeError('A layer mask needs a mask surface.')
  if (!Number.isInteger(mask.offsetX) || !Number.isInteger(mask.offsetY)) throw new TypeError('A layer mask offset must be whole pixels.')
}

function validateLayer(layer: Layer): void {
  if (!layer || typeof layer !== 'object') throw new TypeError('A layer is required.')
  if (typeof layer.id !== 'string' || !layer.id) throw new TypeError('A layer needs an id.')
  if (!LAYER_KINDS.has(layer.kind)) throw new TypeError(`Unknown layer kind "${String((layer as { kind?: unknown }).kind)}".`)
  if (typeof layer.name !== 'string') throw new TypeError('A layer needs a name.')
  if (!isBlendMode(layer.blendMode)) throw new RangeError(`Unknown blend mode "${String(layer.blendMode)}".`)
  if (!(layer.opacity >= 0 && layer.opacity <= 1)) throw new RangeError('Layer opacity must be between 0 and 1.')
  if (!layer.locks || typeof layer.locks !== 'object') throw new TypeError('A layer needs lock flags.')
  if (layer.mask) validateMask(layer.mask)
  if (layer.kind === 'raster') {
    if (!isTiledSurface(layer.surface)) throw new TypeError(`"${layer.name}" needs a pixel surface.`)
    if (!Number.isInteger(layer.offsetX) || !Number.isInteger(layer.offsetY)) throw new TypeError('A layer offset must be whole pixels.')
  } else if (layer.kind === 'text' || layer.kind === 'shape') {
    if (!layer.raster || !isTiledSurface(layer.raster.surface)) throw new TypeError(`"${layer.name}" needs a raster cache.`)
    if (!Number.isInteger(layer.raster.offsetX) || !Number.isInteger(layer.raster.offsetY)) throw new TypeError('A raster cache offset must be whole pixels.')
  } else if (!layer.adjustment || typeof layer.adjustment.type !== 'string') {
    throw new TypeError(`"${layer.name}" needs an adjustment.`)
  }
}

function validateSelection(selection: Selection, width: number, height: number): void {
  const mask = selection?.mask
  if (!mask || mask.width !== width || mask.height !== height || !mask.data || mask.data.length !== width * height) {
    throw new RangeError(`A selection must be a ${width} x ${height} mask.`)
  }
  if (!selection.bounds) throw new TypeError('A selection needs its bounds.')
}

function normalizeLocks(locks: LayerLocks): LayerLocks {
  return { pixels: Boolean(locks?.pixels), position: Boolean(locks?.position), transparency: Boolean(locks?.transparency) }
}

function sameLocks(a: LayerLocks, b: LayerLocks): boolean {
  return a.pixels === b.pixels && a.position === b.position && a.transparency === b.transparency
}

/** Keeps the active layer and edit target valid after any change. */
function normalizeState(state: DocumentState): DocumentState {
  let { activeLayerId, editTarget, selection } = state
  let active = activeLayerId === null ? undefined : state.layers.find((layer) => layer.id === activeLayerId)
  if (activeLayerId !== null && !active) {
    active = state.layers[state.layers.length - 1]
    activeLayerId = active ? active.id : null
  }
  if (editTarget !== 'pixels' && editTarget !== 'mask') editTarget = 'pixels'
  if (editTarget === 'mask' && !active?.mask) editTarget = 'pixels'
  if (selection && (selection.mask.width !== state.width || selection.mask.height !== state.height)) selection = null
  if (activeLayerId === state.activeLayerId && editTarget === state.editTarget && selection === state.selection) return state
  return { ...state, activeLayerId, editTarget, selection }
}

// ---------------------------------------------------------------------------------------------
// Pixel editors
// ---------------------------------------------------------------------------------------------

interface EditorBinding {
  readonly layerId: LayerId
  readonly target: EditTarget
  readonly surface: TiledSurface | TiledMask
  readonly op: TilesOp
  isOpen(): boolean
  /** Called once, when the first tile is snapshotted. */
  onFirstTouch(): void
  /** A document-space rect changed. */
  markDirty(rect: IntRect): void
  /** Document position of the edited surface (null when it is no longer live). */
  origin(): Point | null
}

interface Raster {
  readonly width: number
  readonly height: number
  readonly data: TileData
}

/** Destination rectangle of surface.write(x, y, src, srcRect) after clipping srcRect to the source. */
function writeDestination(x: number, y: number, src: Raster, srcRect?: IntRect): IntRect {
  const left = Math.round(Number(x) || 0)
  const top = Math.round(Number(y) || 0)
  const region = srcRect ? toIntRect(srcRect) : { x: 0, y: 0, width: src.width, height: src.height }
  const sx0 = Math.max(0, region.x)
  const sy0 = Math.max(0, region.y)
  const sx1 = Math.min(src.width, region.x + region.width)
  const sy1 = Math.min(src.height, region.y + region.height)
  return { x: left + (sx0 - region.x), y: top + (sy0 - region.y), width: Math.max(0, sx1 - sx0), height: Math.max(0, sy1 - sy0) }
}

function createPixelEditor(binding: EditorBinding): PixelEditor {
  const live = binding.surface
  const channels = live.channels
  const fill = channels === 1 ? (live as TiledMask).defaultValue : 0
  const stored = binding.op.stored
  let first = true

  function touch(rect: IntRect): void {
    if (!binding.isOpen()) return
    const range = tileRange(rect)
    if (!range) return
    for (let ty = range.ty0; ty < range.ty1; ty += 1) {
      for (let tx = range.tx0; tx < range.tx1; tx += 1) {
        const key = tileKey(tx, ty)
        if (stored.has(key)) continue
        const tile = live.tile(tx, ty) as TileData | undefined
        stored.set(key, tile ? tile.slice() : null)
        if (first) {
          first = false
          binding.onFirstTouch()
        }
      }
    }
  }

  function dirtyLocal(rect: IntRect): void {
    if (!(rect.width > 0 && rect.height > 0)) return
    const origin = binding.origin()
    if (!origin) return
    binding.markDirty({ x: rect.x + origin.x, y: rect.y + origin.y, width: rect.width, height: rect.height })
  }

  function write(x: number, y: number, src: Raster, srcRect?: IntRect): IntRect {
    const dest = writeDestination(x, y, src, srcRect)
    if (!binding.isOpen()) return { x: dest.x, y: dest.y, width: 0, height: 0 }
    touch(dest)
    const result = (live as { write(x: number, y: number, src: Raster, srcRect?: IntRect): IntRect }).write(x, y, src, srcRect)
    dirtyLocal(dest)
    return result
  }

  function setTile(tx: number, ty: number, data: TileData | undefined): void {
    if (!binding.isOpen()) return
    const rect = tileRect(tx, ty)
    touch(rect)
    setTileAny(live, tx, ty, data)
    dirtyLocal(rect)
  }

  function readBefore(rect: IntRect): PixelBuffer | MaskBuffer {
    const r = toIntRect(rect)
    const out: Raster = channels === 4
      ? { width: r.width, height: r.height, data: new Uint8ClampedArray(r.width * r.height * 4) }
      : { width: r.width, height: r.height, data: new Uint8Array(r.width * r.height).fill(fill) }
    const range = tileRange(r)
    if (!range) return out as PixelBuffer | MaskBuffer
    const stride = r.width * channels
    for (let ty = range.ty0; ty < range.ty1; ty += 1) {
      const tileTop = ty * TILE_SIZE
      const y0 = Math.max(r.y, tileTop)
      const y1 = Math.min(r.y + r.height, tileTop + TILE_SIZE)
      for (let tx = range.tx0; tx < range.tx1; tx += 1) {
        const key = tileKey(tx, ty)
        const tile = stored.has(key) ? stored.get(key) : (live.tile(tx, ty) as TileData | undefined)
        if (!tile) continue
        const tileLeft = tx * TILE_SIZE
        const x0 = Math.max(r.x, tileLeft)
        const x1 = Math.min(r.x + r.width, tileLeft + TILE_SIZE)
        const span = (x1 - x0) * channels
        for (let y = y0; y < y1; y += 1) {
          const from = ((y - tileTop) * TILE_SIZE + (x0 - tileLeft)) * channels
          out.data.set(tile.subarray(from, from + span), (y - r.y) * stride + (x0 - r.x) * channels)
        }
      }
    }
    return out as PixelBuffer | MaskBuffer
  }

  // A guarded view of the live surface: reads pass through; writes snapshot tiles for undo first and are
  // ignored once the session has finished.
  const view = {
    get channels() { return live.channels },
    get version() { return live.version },
    get byteSize() { return live.byteSize },
    get tileCount() { return live.tileCount },
    contentBounds: () => live.contentBounds(),
    read: (rect: IntRect, target?: Raster) => (live as { read(rect: IntRect, target?: Raster): Raster }).read(rect, target),
    write: (x: number, y: number, src: Raster, srcRect?: IntRect) => write(x, y, src, srcRect),
    tile: (tx: number, ty: number) => live.tile(tx, ty),
    setTile: (tx: number, ty: number, data: TileData | undefined) => setTile(tx, ty, data),
    tileVersion: (tx: number, ty: number) => live.tileVersion(tx, ty),
    forEachTile: (visit: (tx: number, ty: number, data: TileData) => void) => (live as { forEachTile(visit: (tx: number, ty: number, data: TileData) => void): void }).forEachTile(visit),
    clone: () => live.clone(),
  }
  if (channels === 1) Object.defineProperty(view, 'defaultValue', { get: () => (live as TiledMask).defaultValue, enumerable: true })

  return {
    layerId: binding.layerId,
    target: binding.target,
    surface: view as unknown as TiledSurface | TiledMask,
    touch,
    writePixels(x: number, y: number, src: PixelBuffer): void {
      if (channels !== 4) throw new TypeError('This editor edits a layer mask; use writeMask.')
      write(x, y, src)
    },
    writeMask(x: number, y: number, src: MaskBuffer): void {
      if (channels !== 1) throw new TypeError('This editor edits layer pixels; use writePixels.')
      write(x, y, src)
    },
    readBefore,
    invalidate(rect: IntRect): void {
      if (!binding.isOpen()) return
      const r = toIntRect(rect)
      if (r.width > 0 && r.height > 0) binding.markDirty(r)
    },
  }
}

/** The surface an editor may change, or a clear error. */
function editableSurface(layer: Layer, target: EditTarget): TiledSurface | TiledMask {
  if (target === 'mask') {
    if (!layer.mask) throw new Error(`"${layer.name}" has no layer mask.`)
    return layer.mask.surface
  }
  if (target !== 'pixels') throw new RangeError(`Unknown edit target "${String(target)}".`)
  if (layer.kind === 'raster') return layer.surface
  if (layer.kind === 'adjustment') throw new Error(`"${layer.name}" is an adjustment layer; it has no pixels to edit.`)
  throw new Error(`Rasterize "${layer.name}" before editing its pixels.`)
}

/**
 * Why pixels of `layer` (or its mask) cannot be painted, filtered or adjusted right now, or null.
 * Tools and commands show the message; whole-document operations (image size, rotate) ignore locks.
 */
export function pixelEditBlocker(layer: Layer | null | undefined, target: EditTarget = 'pixels'): string | null {
  if (!layer) return 'Select a layer first.'
  if (target === 'mask') return layer.mask ? null : `"${layer.name}" has no layer mask.`
  if (layer.kind === 'adjustment') return `"${layer.name}" is an adjustment layer. Select a pixel layer or its mask.`
  if (layer.kind !== 'raster') return `"${layer.name}" is a ${layer.kind} layer. Rasterize it first to edit its pixels.`
  if (layer.locks.pixels) return `"${layer.name}" is locked. Unlock its pixels in the Layers panel first.`
  if (!layer.visible) return `"${layer.name}" is hidden. Show it before editing.`
  return null
}

// ---------------------------------------------------------------------------------------------
// Layer patches
// ---------------------------------------------------------------------------------------------

function rasterCacheOf(result: RasterResult, key: string): RasterCache {
  return { surface: surfaceFromBuffer(result.pixels), offsetX: result.offsetX, offsetY: result.offsetY, specKey: key }
}

/** Raster cache for `spec` (synchronous; null when text cannot be drawn here, e.g. no canvas). */
function rasterizeVector(spec: TextSpec | ShapeSpec, isText: boolean): RasterCache | null {
  try {
    const result = isText ? rasterizeTextSync(spec as TextSpec) : rasterizeShape(spec as ShapeSpec)
    return rasterCacheOf(result, specKey(spec))
  } catch (error) {
    if (error instanceof RangeError) throw error
    return null
  }
}

/** New raster cache for a changed vector spec: shifted when only whole pixels moved, else re-rendered. */
function nextRaster(current: RasterCache, before: TextSpec | ShapeSpec, after: TextSpec | ShapeSpec, isText: boolean): RasterCache {
  const shift = current.specKey === specKey(before) ? wholePixelShift(before, after) : null
  if (shift) {
    if (!shift.x && !shift.y) return current
    return { ...current, offsetX: current.offsetX + shift.x, offsetY: current.offsetY + shift.y }
  }
  // Keeps the old (stale) pixels when text cannot be drawn in this environment; the key shows it is stale.
  return rasterizeVector(after, isText) ?? current
}

function patchLayer(layer: Layer, patch: LayerPatch, index: number): Layer {
  if (!patch || typeof patch !== 'object') return layer
  const next = { ...layer } as Mutable<Layer>
  let mask: Mutable<LayerMask> | null = layer.mask ? { ...layer.mask } : null
  if (patch.name !== undefined) {
    const name = String(patch.name).trim().slice(0, 255)
    if (name) next.name = name
  }
  if (patch.visible !== undefined) next.visible = Boolean(patch.visible)
  if (patch.opacity !== undefined) {
    if (!Number.isFinite(patch.opacity)) throw new RangeError('Opacity must be a number between 0 and 1.')
    next.opacity = clamp01(patch.opacity)
  }
  if (patch.blendMode !== undefined) {
    if (!isBlendMode(patch.blendMode)) throw new RangeError(`Unknown blend mode "${String(patch.blendMode)}".`)
    next.blendMode = patch.blendMode
  }
  if (patch.locks !== undefined) {
    const locks = normalizeLocks(patch.locks)
    if (!sameLocks(locks, layer.locks)) next.locks = locks
  }
  if (patch.isBackground !== undefined) {
    if (next.kind !== 'raster') {
      if (patch.isBackground) throw new TypeError('Only a pixel layer can be the Background.')
    } else {
      const background = Boolean(patch.isBackground)
      if (background && index !== 0) throw new Error('Only the bottom layer can become the Background.')
      next.isBackground = background
    }
  }
  if (patch.clipped !== undefined) next.clipped = Boolean(patch.clipped)
  if (next.kind === 'raster' && next.isBackground) next.clipped = false
  if (patch.text !== undefined) {
    if (next.kind !== 'text' || layer.kind !== 'text') throw new TypeError(`"${layer.name}" is not a text layer.`)
    if (!patch.text || typeof patch.text.text !== 'string' || !patch.text.style) throw new TypeError('Invalid text.')
    next.raster = nextRaster(layer.raster, layer.text, patch.text, true)
    next.text = patch.text
  }
  if (patch.shape !== undefined) {
    if (next.kind !== 'shape' || layer.kind !== 'shape') throw new TypeError(`"${layer.name}" is not a shape layer.`)
    if (!patch.shape || typeof patch.shape.kind !== 'string') throw new TypeError('Invalid shape.')
    next.raster = nextRaster(layer.raster, layer.shape, patch.shape, false)
    next.shape = patch.shape
  }
  if (patch.adjustment !== undefined) {
    if (next.kind !== 'adjustment') throw new TypeError(`"${layer.name}" is not an adjustment layer.`)
    if (!patch.adjustment || typeof patch.adjustment.type !== 'string') throw new TypeError('Invalid adjustment.')
    next.adjustment = patch.adjustment
  }
  if (patch.offset !== undefined) {
    const x = Math.round(Number(patch.offset?.x))
    const y = Math.round(Number(patch.offset?.y))
    if (!Number.isFinite(x) || !Number.isFinite(y)) throw new RangeError('A layer position must be a number.')
    let dx = 0
    let dy = 0
    if (next.kind === 'raster') {
      dx = x - next.offsetX
      dy = y - next.offsetY
      next.offsetX = x
      next.offsetY = y
    } else if (next.kind === 'text' || next.kind === 'shape') {
      // The raster cache origin is the layer position; the vector transform moves with it.
      dx = x - next.raster.offsetX
      dy = y - next.raster.offsetY
      if (dx || dy) {
        next.raster = { ...next.raster, offsetX: x, offsetY: y }
        if (next.kind === 'text') {
          const m = next.text.transform
          next.text = { ...next.text, transform: [m[0], m[1], m[2], m[3], m[4] + dx, m[5] + dy] }
        } else {
          const m = next.shape.transform
          next.shape = { ...next.shape, transform: [m[0], m[1], m[2], m[3], m[4] + dx, m[5] + dy] }
        }
      }
    } else if (mask) {
      // An adjustment layer's position is its mask's.
      mask.offsetX = x
      mask.offsetY = y
    }
    if ((dx || dy) && mask && mask.linked) {
      mask.offsetX += dx
      mask.offsetY += dy
    }
  }
  if (mask) {
    if (patch.maskEnabled !== undefined) mask.enabled = Boolean(patch.maskEnabled)
    if (patch.maskLinked !== undefined) mask.linked = Boolean(patch.maskLinked)
    const before = layer.mask as LayerMask
    if (mask.enabled !== before.enabled || mask.linked !== before.linked || mask.offsetX !== before.offsetX || mask.offsetY !== before.offsetY) {
      next.mask = mask
    }
  }
  const before = layer as unknown as Record<string, unknown>
  const after = next as unknown as Record<string, unknown>
  for (const key of Object.keys(after)) if (after[key] !== before[key]) return next as Layer
  return layer
}

// ---------------------------------------------------------------------------------------------
// Store
// ---------------------------------------------------------------------------------------------

class RollbackSignal extends Error {
  constructor() {
    super('The transaction was rolled back.')
    this.name = 'RollbackSignal'
  }
}

interface TxContext extends ChangeAccumulator {
  state: DocumentState
  readonly ops: HistoryOp[]
  open: boolean
  readonly tx: Transaction
}

export interface DocumentStoreInit {
  readonly width: number
  readonly height: number
  readonly ppi: number
  /** Bottom to top. */
  readonly layers: readonly Layer[]
  /** Default: the top layer. */
  readonly activeLayerId?: LayerId | null
  readonly host: Pick<AdvancedHost, 'nextRevision' | 'setRevision'> & { readonly currentRevision?: () => number }
  /** Label of the base history entry ("Open", "Advanced editor"). */
  readonly baseLabel: string
  /** Host revision of the opening state. Default host.currentRevision?.() ?? 0. */
  readonly revision?: number
  readonly selection?: Selection | null
  /** Clock for coalescing (tests inject one). Default Date.now. */
  readonly now?: () => number
  readonly historyBudgetBytes?: number
  readonly historyMaxEntries?: number
}

/** Members the store offers on top of the shared DocumentStore contract. */
export interface DocumentStoreExtras {
  /** True while a stroke session is open. */
  hasOpenStroke(): boolean
  /** Commits the open stroke session, if any (Save and Close call this through settle()). */
  commitStroke(): void
}

export type AdvancedDocumentStore = DocumentStore & DocumentStoreExtras & { readonly history: HistoryController }

export function createDocumentStore(init: DocumentStoreInit): AdvancedDocumentStore {
  validateCanvas(init.width, init.height)
  const host = init.host
  if (!host || typeof host.nextRevision !== 'function' || typeof host.setRevision !== 'function') {
    throw new TypeError('The document store needs a host with nextRevision() and setRevision().')
  }
  const now = typeof init.now === 'function' ? init.now : () => Date.now()
  const layers = [...(init.layers ?? [])]
  if (layers.length > LIMITS.maxLayers) throw new RangeError(`A document can have at most ${LIMITS.maxLayers} layers.`)
  const ids = new Set<LayerId>()
  layers.forEach((layer, index) => {
    validateLayer(layer)
    if (ids.has(layer.id)) throw new Error(`Two layers share the id "${layer.id}".`)
    ids.add(layer.id)
    if (isBackgroundLayer(layer) && index !== 0) throw new Error('The Background layer must be the bottom layer.')
  })
  const startRevision = Number.isFinite(init.revision)
    ? Number(init.revision)
    : typeof host.currentRevision === 'function' ? host.currentRevision() : 0
  const ppi = Number.isFinite(init.ppi) && init.ppi > 0 ? init.ppi : 72
  if (init.selection) validateSelection(init.selection, init.width, init.height)

  let pixelCounter = 1
  let selectionVersion = Math.max(1, (init.selection?.version ?? 0) + 1)
  let state: DocumentState = normalizeState({
    width: init.width,
    height: init.height,
    ppi,
    layers,
    activeLayerId: init.activeLayerId === undefined ? layers[layers.length - 1]?.id ?? null : init.activeLayerId,
    editTarget: 'pixels',
    selection: init.selection ?? null,
    revision: startRevision,
    pixelVersion: pixelCounter,
  })
  const listeners = new Set<(change: DocumentChange) => void>()
  let active: TxContext | null = null
  let stroke: StrokeSession | null = null
  let navigation: ChangeAccumulator | null = null
  let disposed = false

  function nextSelectionVersion(): number {
    selectionVersion += 1
    return selectionVersion
  }

  function assertUsable(): void {
    if (disposed) throw new Error('This document has been closed.')
  }

  function emit(change: DocumentChange): void {
    for (const listener of [...listeners]) {
      try {
        listener(change)
      } catch (error) {
        // One broken subscriber must not stop the others or corrupt the store.
        console.error(error)
      }
    }
  }

  function commitStroke(): void {
    stroke?.commit()
  }

  // ----- operation application (do / undo / redo) -------------------------------------------

  function applyOp(s: DocumentState, op: HistoryOp, direction: 'undo' | 'redo', acc: ChangeAccumulator): DocumentState {
    switch (op.kind) {
      case 'tiles':
        swapTiles(op, surfaceOrigin(s, op.layerId, op.target, op.surface), acc.dirty)
        return s
      case 'surface': {
        const index = indexOfLayer(s, op.layerId)
        if (index < 0) throw corrupt('surface layer missing')
        const layer = s.layers[index]
        let next: Layer
        if (op.target === 'mask') {
          const mask = layer.mask
          if (!mask) throw corrupt('mask missing')
          next = { ...layer, mask: { ...mask, surface: op.other as TiledMask, offsetX: op.otherOffset.x, offsetY: op.otherOffset.y } } as Layer
          op.other = mask.surface
          op.otherOffset = { x: mask.offsetX, y: mask.offsetY }
        } else if (layer.kind === 'raster') {
          next = { ...layer, surface: op.other as TiledSurface, offsetX: op.otherOffset.x, offsetY: op.otherOffset.y }
          op.other = layer.surface
          op.otherOffset = { x: layer.offsetX, y: layer.offsetY }
        } else {
          throw corrupt('surface swap on a non-pixel layer')
        }
        acc.dirty.add(layerDirty(layer))
        acc.dirty.add(layerDirty(next))
        acc.structure = true
        return withLayerAt(s, index, next)
      }
      case 'insert':
      case 'remove': {
        const adding = (op.kind === 'insert') !== (direction === 'undo')
        acc.structure = true
        acc.dirty.add(layerDirty(op.layer))
        const list = s.layers.slice()
        if (adding) {
          if (indexOfLayer(s, op.layer.id) >= 0) throw corrupt('layer inserted twice')
          list.splice(Math.min(Math.max(0, op.index), list.length), 0, op.layer)
        } else {
          const index = indexOfLayer(s, op.layer.id)
          if (index < 0) throw corrupt('removed layer missing')
          list.splice(index, 1)
        }
        return { ...s, layers: list }
      }
      case 'move': {
        const to = direction === 'undo' ? op.from : op.to
        const index = indexOfLayer(s, op.layerId)
        if (index < 0) throw corrupt('moved layer missing')
        const list = s.layers.slice()
        const [layer] = list.splice(index, 1)
        list.splice(Math.min(Math.max(0, to), list.length), 0, layer)
        acc.structure = true
        // Moving can change clipping groups and what adjustments apply to: recomposite everything.
        acc.dirty.add('all')
        return { ...s, layers: list }
      }
      case 'layer': {
        const target = direction === 'undo' ? op.before : op.after
        const index = indexOfLayer(s, target.id)
        if (index < 0) throw corrupt('changed layer missing')
        acc.dirty.add(layerDirty(s.layers[index]))
        acc.dirty.add(layerDirty(target))
        acc.structure = true
        return withLayerAt(s, index, target)
      }
      case 'mask': {
        const mask = direction === 'undo' ? op.before : op.after
        const index = indexOfLayer(s, op.layerId)
        if (index < 0) throw corrupt('masked layer missing')
        const layer = s.layers[index]
        acc.dirty.add(layerDirty(layer))
        acc.structure = true
        return withLayerAt(s, index, { ...layer, mask } as Layer)
      }
      case 'active': {
        const ref = direction === 'undo' ? op.before : op.after
        acc.structure = true
        return { ...s, activeLayerId: ref.id, editTarget: ref.target }
      }
      case 'selection': {
        const snapshot = direction === 'undo' ? op.before : op.after
        acc.selection = true
        return { ...s, selection: restoreSelection(snapshot, s.width, s.height, nextSelectionVersion()) }
      }
      case 'canvas': {
        const info = direction === 'undo' ? op.before : op.after
        acc.structure = true
        acc.dirty.add('all')
        return { ...s, width: info.width, height: info.height, ppi: info.ppi }
      }
      default:
        throw corrupt('unknown operation')
    }
  }

  // ----- history ------------------------------------------------------------------------------

  const history = createHistory({
    baseLabel: init.baseLabel,
    baseRevision: startRevision,
    host,
    budgetBytes: init.historyBudgetBytes,
    maxEntries: init.historyMaxEntries,
    beforeNavigate: () => {
      if (active) throw new Error('Undo and redo are not available while a change is being made.')
      commitStroke()
    },
    apply: (entry: HistoryEntry, direction) => {
      if (!navigation) navigation = { dirty: new DirtySet(), structure: false, selection: false }
      const ops = direction === 'undo' ? [...entry.ops].reverse() : entry.ops
      let s = state
      for (const op of ops) s = applyOp(s, op, direction, navigation)
      state = s
    },
    afterNavigate: (revision: number) => {
      const acc = navigation ?? { dirty: new DirtySet(), structure: false, selection: false }
      navigation = null
      const pixelVersion = acc.dirty.empty ? state.pixelVersion : (pixelCounter += 1)
      state = normalizeState({ ...state, revision, pixelVersion })
      emit({ structure: acc.structure, selection: acc.selection, history: true, dirty: acc.dirty.result() })
    },
  })

  // ----- transactions -------------------------------------------------------------------------

  function createTransaction(ctx: TxContext): Transaction {
    const check = () => {
      if (!ctx.open) throw new Error('This change has already finished; start a new transaction.')
    }
    const record = (op: HistoryOp) => {
      ctx.state = applyOp(ctx.state, op, 'redo', ctx)
      ctx.ops.push(op)
    }
    const tx: Transaction = {
      get state() {
        return ctx.state
      },
      insertLayer(layer: Layer, index?: number): void {
        check()
        validateLayer(layer)
        const s = ctx.state
        if (indexOfLayer(s, layer.id) >= 0) throw new Error(`A layer with the id "${layer.id}" already exists.`)
        if (s.layers.length >= LIMITS.maxLayers) throw new Error(`A document can have at most ${LIMITS.maxLayers} layers.`)
        let at = index === undefined || !Number.isFinite(index) ? s.layers.length : Math.min(Math.max(0, Math.round(index)), s.layers.length)
        const hasBackground = isBackgroundLayer(s.layers[0])
        if (isBackgroundLayer(layer)) {
          if (hasBackground) throw new Error('The document already has a Background layer.')
          at = 0
        } else if (hasBackground && at === 0) {
          // Nothing goes below the Background (Photoshop).
          at = 1
        }
        record({ kind: 'insert', layer, index: at })
      },
      removeLayer(id: LayerId): void {
        check()
        const s = ctx.state
        const index = requireIndex(s, id)
        if (s.activeLayerId === id) {
          // Photoshop selects the layer below (or above when there is none).
          const replacement = s.layers[index - 1] ?? s.layers[index + 1] ?? null
          record({ kind: 'active', before: { id, target: s.editTarget }, after: { id: replacement ? replacement.id : null, target: 'pixels' } })
        }
        record({ kind: 'remove', layer: ctx.state.layers[index], index })
      },
      moveLayer(id: LayerId, toIndex: number): void {
        check()
        const s = ctx.state
        const from = requireIndex(s, id)
        const layer = s.layers[from]
        if (isBackgroundLayer(layer)) return
        let to = Math.min(Math.max(0, Math.round(Number(toIndex) || 0)), s.layers.length - 1)
        if (to === 0 && isBackgroundLayer(s.layers[0])) to = 1
        if (to === from) return
        record({ kind: 'move', layerId: id, from, to })
      },
      updateLayer(id: LayerId, patch: LayerPatch): void {
        check()
        const s = ctx.state
        const index = requireIndex(s, id)
        const layer = s.layers[index]
        const next = patchLayer(layer, patch, index)
        if (next === layer) return
        record({ kind: 'layer', before: layer, after: next })
        if (s.activeLayerId === id && s.editTarget === 'mask' && !next.mask) {
          record({ kind: 'active', before: { id, target: 'mask' }, after: { id, target: 'pixels' } })
        }
      },
      setMask(id: LayerId, mask: LayerMask | null): void {
        check()
        const s = ctx.state
        const index = requireIndex(s, id)
        const layer = s.layers[index]
        const next = mask ?? null
        if (next) {
          validateMask(next)
          if (isBackgroundLayer(layer)) throw new Error('The Background layer cannot have a mask. Convert it to a normal layer first.')
        }
        if (layer.mask === next) return
        if (!next && s.activeLayerId === id && s.editTarget === 'mask') {
          record({ kind: 'active', before: { id, target: 'mask' }, after: { id, target: 'pixels' } })
        }
        record({ kind: 'mask', layerId: id, before: layer.mask, after: next })
      },
      setActiveLayer(id: LayerId | null, target: EditTarget = 'pixels'): void {
        check()
        const s = ctx.state
        let editTarget: EditTarget = target === 'mask' ? 'mask' : 'pixels'
        if (id !== null) {
          const layer = s.layers[requireIndex(s, id)]
          if (editTarget === 'mask' && !layer.mask) editTarget = 'pixels'
        } else {
          editTarget = 'pixels'
        }
        const before: ActiveRef = { id: s.activeLayerId, target: s.editTarget }
        if (before.id === id && before.target === editTarget) return
        record({ kind: 'active', before, after: { id, target: editTarget } })
      },
      setSelection(selection: Selection | null): void {
        check()
        const s = ctx.state
        const next = selection ?? null
        if (next === s.selection) return
        if (next) validateSelection(next, s.width, s.height)
        ctx.ops.push({ kind: 'selection', before: snapshotSelection(s.selection), after: snapshotSelection(next) })
        ctx.state = { ...s, selection: next }
        ctx.selection = true
        if (next && Number.isFinite(next.version)) selectionVersion = Math.max(selectionVersion, next.version)
      },
      editPixels(id: LayerId, target: EditTarget): PixelEditor {
        check()
        const layer = ctx.state.layers[requireIndex(ctx.state, id)]
        const surface = editableSurface(layer, target)
        const op: TilesOp = { kind: 'tiles', layerId: id, target, surface, stored: new Map() }
        return createPixelEditor({
          layerId: id,
          target,
          surface,
          op,
          isOpen: () => ctx.open,
          onFirstTouch: () => { ctx.ops.push(op) },
          markDirty: (rect) => ctx.dirty.add(rect),
          origin: () => surfaceOrigin(ctx.state, id, target, surface),
        })
      },
      replaceSurface(id: LayerId, target: EditTarget, next: TiledSurface | TiledMask, offset?: Point): void {
        check()
        const s = ctx.state
        const layer = s.layers[requireIndex(s, id)]
        const point = offset ? { x: Math.round(Number(offset.x)), y: Math.round(Number(offset.y)) } : null
        if (point && (!Number.isFinite(point.x) || !Number.isFinite(point.y))) throw new RangeError('A surface offset must be a number.')
        if (target === 'mask') {
          if (!isTiledMask(next)) throw new TypeError('Replacing a mask needs a mask surface.')
          const mask = layer.mask
          if (!mask) throw new Error(`"${layer.name}" has no layer mask.`)
          const otherOffset = point ?? { x: mask.offsetX, y: mask.offsetY }
          if (next === mask.surface && otherOffset.x === mask.offsetX && otherOffset.y === mask.offsetY) return
          record({ kind: 'surface', layerId: id, target: 'mask', other: next, otherOffset })
          return
        }
        if (target !== 'pixels') throw new RangeError(`Unknown edit target "${String(target)}".`)
        if (!isTiledSurface(next)) throw new TypeError('Replacing pixels needs a pixel surface.')
        if (layer.kind === 'raster') {
          const otherOffset = point ?? { x: layer.offsetX, y: layer.offsetY }
          if (next === layer.surface && otherOffset.x === layer.offsetX && otherOffset.y === layer.offsetY) return
          record({ kind: 'surface', layerId: id, target: 'pixels', other: next, otherOffset })
          return
        }
        if (layer.kind === 'text' || layer.kind === 'shape') {
          // The pixels of a vector layer are its raster cache; the new pixels are current for its spec.
          const raster: RasterCache = {
            surface: next,
            offsetX: point ? point.x : layer.raster.offsetX,
            offsetY: point ? point.y : layer.raster.offsetY,
            specKey: specKey(layer.kind === 'text' ? layer.text : layer.shape),
          }
          record({ kind: 'layer', before: layer, after: { ...layer, raster } })
          return
        }
        throw new Error(`"${layer.name}" is an adjustment layer; it has no pixels.`)
      },
      setCanvas(width: number, height: number, ppiValue?: number): void {
        check()
        validateCanvas(width, height)
        const s = ctx.state
        let nextPpi = s.ppi
        if (ppiValue !== undefined) {
          if (!(Number.isFinite(ppiValue) && ppiValue > 0)) throw new RangeError('Resolution must be a positive number.')
          nextPpi = ppiValue
        }
        if (width === s.width && height === s.height && nextPpi === s.ppi) return
        // A selection is a document-size mask; it cannot survive a size change.
        if ((width !== s.width || height !== s.height) && s.selection) tx.setSelection(null)
        record({ kind: 'canvas', before: { width: s.width, height: s.height, ppi: s.ppi }, after: { width, height, ppi: nextPpi } })
      },
      rollback(): never {
        check()
        throw new RollbackSignal()
      },
    }
    return tx
  }

  function revert(ctx: TxContext): void {
    // Metadata lives in immutable states (the store state was not touched); only in-place pixel edits
    // need undoing, newest first.
    for (let i = ctx.ops.length - 1; i >= 0; i -= 1) {
      const op = ctx.ops[i]
      if (op.kind === 'tiles') swapTiles(op, null, null)
    }
  }

  function finish(ctx: TxContext, label: string, icon: HistoryIcon, options?: TransactOptions): void {
    const affectsOutput = ctx.ops.some(opAffectsOutput)
    const revisionBefore = state.revision
    let revision = revisionBefore
    if (affectsOutput) {
      revision = host.nextRevision()
      host.setRevision(revision)
    }
    const pixelVersion = ctx.dirty.empty ? state.pixelVersion : (pixelCounter += 1)
    state = normalizeState({ ...ctx.state, revision, pixelVersion })
    // Choosing another layer is not an undoable step (Photoshop); everything else is.
    const recorded = !ctx.ops.every((op) => op.kind === 'active')
    if (recorded) {
      const time = now()
      const target = history.coalesceTarget(options?.coalesceKey, time)
      if (target) history.extend(target, ctx.ops, revision, time)
      else history.record({ label, icon, ops: ctx.ops, revisionBefore, revisionAfter: revision, coalesceKey: options?.coalesceKey ?? null, time })
    }
    emit({ structure: ctx.structure, selection: ctx.selection, history: recorded, dirty: ctx.dirty.result() })
  }

  // ----- the store ----------------------------------------------------------------------------

  const store: AdvancedDocumentStore = {
    history,
    getState: () => state,
    subscribe(listener: (change: DocumentChange) => void): () => void {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    transact<T>(label: string, icon: HistoryIcon, run: (tx: Transaction) => T, options?: TransactOptions): T {
      assertUsable()
      if (navigation) throw new Error('The document cannot change while undo or redo is being applied.')
      // A nested call joins the running transaction (one history step, one revision).
      if (active) return run(active.tx)
      commitStroke()
      const ctx = {
        state,
        ops: [] as HistoryOp[],
        dirty: new DirtySet(),
        structure: false,
        selection: false,
        open: true,
      } as Mutable<TxContext>
      ctx.tx = createTransaction(ctx as TxContext)
      active = ctx as TxContext
      let result: T
      try {
        result = run(ctx.tx)
      } catch (error) {
        ctx.open = false
        active = null
        revert(ctx as TxContext)
        if (error instanceof RollbackSignal) return undefined as T
        throw error
      }
      ctx.open = false
      active = null
      if (ctx.ops.length) finish(ctx as TxContext, label, icon, options)
      return result
    },
    beginStroke(layerId: LayerId, target: EditTarget, label: string, icon: HistoryIcon): StrokeSession {
      assertUsable()
      if (active) throw new Error('A stroke cannot start inside a transaction.')
      if (navigation) throw new Error('A stroke cannot start while undo or redo is being applied.')
      commitStroke()
      const layer = state.layers[requireIndex(state, layerId)]
      const surface = editableSurface(layer, target)
      const op: TilesOp = { kind: 'tiles', layerId, target, surface, stored: new Map() }
      const revisionBefore = state.revision
      let open = true
      let touched = false
      const editor = createPixelEditor({
        layerId,
        target,
        surface,
        op,
        isOpen: () => open,
        onFirstTouch: () => {
          touched = true
          // Modified from the first pixel: a Save during the stroke records this provisional revision.
          const provisional = host.nextRevision()
          host.setRevision(provisional)
          state = { ...state, revision: provisional }
        },
        markDirty: (rect) => {
          pixelCounter += 1
          state = { ...state, pixelVersion: pixelCounter }
          emit({ structure: false, selection: false, history: false, dirty: [rect] })
        },
        origin: () => surfaceOrigin(state, layerId, target, surface),
      })
      const session: StrokeSession = {
        editor,
        commit(): void {
          if (!open) return
          open = false
          if (stroke === session) stroke = null
          if (!touched || !op.stored.size) return
          const revisionAfter = host.nextRevision()
          host.setRevision(revisionAfter)
          state = { ...state, revision: revisionAfter }
          history.record({ label, icon, ops: [op], revisionBefore, revisionAfter, coalesceKey: null, time: now() })
          emit({ structure: false, selection: false, history: true, dirty: [] })
        },
        cancel(): void {
          if (!open) return
          open = false
          if (stroke === session) stroke = null
          if (!touched) return
          const dirty = new DirtySet()
          const origin = surfaceOrigin(state, layerId, target, surface)
          for (const [key, tile] of op.stored) {
            const { tx, ty } = tileCoords(key)
            setTileAny(surface, tx, ty, tile)
            if (origin) dirty.add({ x: origin.x + tx * TILE_SIZE, y: origin.y + ty * TILE_SIZE, width: TILE_SIZE, height: TILE_SIZE })
          }
          op.stored.clear()
          host.setRevision(revisionBefore)
          pixelCounter += 1
          state = { ...state, revision: revisionBefore, pixelVersion: pixelCounter }
          emit({ structure: false, selection: false, history: false, dirty: dirty.result() })
        },
      }
      stroke = session
      return session
    },
    memoryUsage(): MemoryUsage {
      const bytes = documentBytes(state)
      return usageOf({ layers: bytes.layers, masks: bytes.masks, history: history.getState().totalBytes })
    },
    dispose(): void {
      if (disposed) return
      try {
        commitStroke()
      } finally {
        disposed = true
        listeners.clear()
        history.dispose()
      }
    },
    hasOpenStroke: () => stroke !== null,
    commitStroke,
  }
  return store
}

// ---------------------------------------------------------------------------------------------
// Layer factories and document helpers
// ---------------------------------------------------------------------------------------------

let layerCounter = 0

/** A new unique layer id. */
export function newLayerId(): LayerId {
  layerCounter += 1
  const crypto = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto
  const unique = typeof crypto?.randomUUID === 'function'
    ? crypto.randomUUID()
    : `${Date.now().toString(36)}-${Math.floor(Math.random() * 0x7fffffff).toString(36)}`
  return `layer-${layerCounter.toString(36)}-${unique}`
}

function commonFields(name: string, init: Partial<Omit<LayerCommon, 'id'>> = {}): LayerCommon {
  return {
    id: newLayerId(),
    name: String(name ?? '').trim() || 'Layer',
    visible: init.visible ?? true,
    opacity: clamp01(Number.isFinite(init.opacity) ? Number(init.opacity) : 1),
    blendMode: init.blendMode && isBlendMode(init.blendMode) ? init.blendMode : 'normal',
    locks: init.locks ? normalizeLocks(init.locks) : DEFAULT_LOCKS,
    mask: init.mask ?? null,
    clipped: Boolean(init.clipped),
  }
}

/** A pixel layer. A Background is opaque-normal, unmasked, unclipped and locked in position/transparency. */
export function createRasterLayer(init: { name: string; surface?: TiledSurface; offsetX?: number; offsetY?: number; isBackground?: boolean }
  & Partial<Omit<LayerCommon, 'id'>>): RasterLayer {
  const isBackground = Boolean(init.isBackground)
  const common = commonFields(init.name || (isBackground ? 'Background' : 'Layer'), init)
  return {
    kind: 'raster',
    ...common,
    ...(isBackground ? { opacity: 1, blendMode: 'normal' as const, mask: null, clipped: false, locks: init.locks ? normalizeLocks(init.locks) : BACKGROUND_LOCKS } : {}),
    isBackground,
    offsetX: Math.round(Number(init.offsetX) || 0),
    offsetY: Math.round(Number(init.offsetY) || 0),
    surface: init.surface ?? createSurface(),
  }
}

/** An adjustment layer (pass maskFromSelection(selection) to limit it to the active selection). */
export function createAdjustmentLayer(name: string, spec: AdjustmentSpec, mask: LayerMask | null = null): AdjustmentLayer {
  if (!spec || typeof spec.type !== 'string') throw new TypeError('An adjustment layer needs an adjustment.')
  return { kind: 'adjustment', ...commonFields(name, { mask }), adjustment: spec }
}

/** A text layer; rasterizes the text (after its font loads) unless a raster cache is given (PSD import). */
export async function createTextLayer(name: string, spec: TextSpec, raster?: RasterCache): Promise<TextLayer> {
  const cache = raster ?? rasterCacheOf(await rasterizeText(spec), specKey(spec))
  return { kind: 'text', ...commonFields(name), text: spec, raster: cache }
}

/** A shape layer with its raster cache. */
export function createShapeLayer(name: string, spec: ShapeSpec): ShapeLayer {
  return { kind: 'shape', ...commonFields(name), shape: spec, raster: rasterCacheOf(rasterizeShape(spec), specKey(spec)) }
}

/** A raster cache from rendered pixels (vector.ts results, or a PSD's own text pixels). */
export function rasterCacheFrom(pixels: PixelBuffer, offsetX: number, offsetY: number, key: string): RasterCache {
  return rasterCacheOf({ pixels, offsetX: Math.round(offsetX), offsetY: Math.round(offsetY) }, key)
}

/** A layer mask (default: reveal all, enabled, linked). */
export function createLayerMask(init: { surface?: TiledMask; defaultValue?: 0 | 255; offsetX?: number; offsetY?: number; enabled?: boolean; linked?: boolean } = {}): LayerMask {
  return {
    surface: init.surface ?? createMaskSurface(init.defaultValue ?? 255),
    offsetX: Math.round(Number(init.offsetX) || 0),
    offsetY: Math.round(Number(init.offsetY) || 0),
    enabled: init.enabled ?? true,
    linked: init.linked ?? true,
  }
}

/** Photoshop: a mask added with a selection active hides everything outside it. Null selection = reveal all. */
export function maskFromSelection(selection: Selection | null): LayerMask {
  if (!selection) return createLayerMask()
  const bounds = selection.bounds
  return createLayerMask({ surface: maskFromBuffer(cropMask(selection.mask, bounds), 0, bounds.x, bounds.y) })
}

/** Deep copy with a new id (Ctrl+J without a selection). A copy of the Background is a normal layer. */
export function duplicateLayer(layer: Layer, name = `${layer.name} copy`): Layer {
  const common = { ...commonFields(name), visible: layer.visible, opacity: layer.opacity, blendMode: layer.blendMode, clipped: layer.clipped }
  const mask = layer.mask ? { ...layer.mask, surface: layer.mask.surface.clone() } : null
  if (layer.kind === 'raster') {
    return {
      ...common,
      kind: 'raster',
      locks: layer.isBackground ? DEFAULT_LOCKS : layer.locks,
      mask,
      isBackground: false,
      offsetX: layer.offsetX,
      offsetY: layer.offsetY,
      surface: layer.surface.clone(),
    }
  }
  if (layer.kind === 'text') return { ...common, kind: 'text', locks: layer.locks, mask, text: layer.text, raster: { ...layer.raster, surface: layer.raster.surface.clone() } }
  if (layer.kind === 'shape') return { ...common, kind: 'shape', locks: layer.locks, mask, shape: layer.shape, raster: { ...layer.raster, surface: layer.raster.surface.clone() } }
  return { ...common, kind: 'adjustment', locks: layer.locks, mask, adjustment: layer.adjustment }
}

/** Layer > Rasterize: a pixel layer showing exactly the text/shape raster (new id; insert it in its place). */
export function rasterizeVectorLayer(layer: TextLayer | ShapeLayer): RasterLayer {
  return {
    kind: 'raster',
    ...commonFields(layer.name),
    visible: layer.visible,
    opacity: layer.opacity,
    blendMode: layer.blendMode,
    locks: layer.locks,
    mask: layer.mask,
    clipped: layer.clipped,
    isBackground: false,
    offsetX: layer.raster.offsetX,
    offsetY: layer.raster.offsetY,
    surface: layer.raster.surface.clone(),
  }
}

function escapeRegExp(text: string): string {
  return text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

/**
 * Photoshop-style names: "Layer 3" (one more than the highest "Layer n"), "Levels 1", and for copies
 * "Layer 1 copy", then "Layer 1 copy 2", "Layer 1 copy 3".
 */
export function nextLayerName(state: Pick<DocumentState, 'layers'>, base: string): string {
  const stem = String(base ?? '').trim() || 'Layer'
  const names = new Set(state.layers.map((layer) => layer.name))
  if (/ copy$/i.test(stem)) {
    if (!names.has(stem)) return stem
    for (let n = 2; ; n += 1) if (!names.has(`${stem} ${n}`)) return `${stem} ${n}`
  }
  const pattern = new RegExp(`^${escapeRegExp(stem)} (\\d+)$`)
  let highest = 0
  for (const name of names) {
    const match = pattern.exec(name)
    if (match) highest = Math.max(highest, Number(match[1]))
  }
  return `${stem} ${highest + 1}`
}

/**
 * True when flattening loses nothing (design 7.3): exactly one visible raster layer with opacity 1, normal
 * blending, no mask (even a disabled one), not clipped, and every pixel inside the canvas.
 */
export function isFlatEquivalent(state: Pick<DocumentState, 'width' | 'height' | 'layers'>): boolean {
  if (state.layers.length !== 1) return false
  const layer = state.layers[0]
  if (layer.kind !== 'raster' || !layer.visible || layer.opacity !== 1 || layer.blendMode !== 'normal' || layer.mask || layer.clipped) return false
  const bounds = layerContentBounds(layer)
  if (!bounds) return true
  return bounds.x >= 0 && bounds.y >= 0 && bounds.x + bounds.width <= state.width && bounds.y + bounds.height <= state.height
}

export { isBackgroundLayer }
