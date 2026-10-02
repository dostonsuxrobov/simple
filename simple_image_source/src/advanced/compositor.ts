// src/advanced/compositor.ts (WP3)
// The display compositor of the Advanced canvas (design 5.6). The pure per-tile compositor (composite.ts)
// computes pixels; this module decides which tiles to compute, keeps them on the GPU and draws them.
//   - Level: chooseLevel(zoom, dpr) (never a magnified proxy); while setInteractive(true) (slider drags,
//     gestures that invalidate large areas) one level coarser, refined when it ends.
//   - Level window: one canvas covering the visible tiles of the active level plus a one-tile ring.
//     Composited tiles are uploaded into it once (putImageData); every frame draws it to the screen canvas
//     with one drawImage over a checkerboard. Smoothing is on when a level pixel is drawn smaller than a
//     device pixel and off when magnifying level 0 (crisp pixels, like Photoshop). Panning past the ring
//     re-centres the window (the overlap is copied on the GPU) and refills the rest from a CPU LRU of
//     composited tiles (LIMITS.compositeTileCacheTiles = 512 tiles, 128 MB).
//   - Scheduler: requestAnimationFrame with an 8 ms compositing budget (LIMITS.frameBudgetMs): visible
//     tiles nearest the viewport centre first, then the ring; nothing off-screen. Stale work is dropped
//     whenever the view changes because the work list is rebuilt every frame. Store changes invalidate
//     composite tiles at every level; invalidated tiles keep showing their old pixels until replaced, and
//     after a level change the previous window stays underneath as a placeholder, so nothing flashes.
//     Small changes (brush dabs) mark only part of a tile out of date: that part alone is recomposited
//     into the cached tile and uploaded (composite blocks are position-stable, so the bytes are the same).
//     Where rAF does not fire (hidden or never-shown windows; measured in Electron 43) a timer takes over
//     at frame rate. settle() resolves only after a frame has drawn the latest view and state with every
//     visible tile current.
//   - flatten() / renderToCanvas() composite level 0 (never a proxy, never a preview) in tile rows,
//     yielding to the event loop and reusing current display tiles; a change during the pass restarts it
//     so the result is one consistent state. renderToCanvas streams rows into a new <canvas> that the
//     caller releases (width = height = 1).
//   - dispose() releases the GPU windows and clears the global proxy-tile cache (pyramid.ts).
// DOM access happens only inside functions, never at module load.
import type { IntRect, OpOptions, PixelBuffer, Rgba8 } from '../imaging/types.ts'
import type {
  CompositePreview,
  Compositor,
  DocumentChange,
  DocumentState,
  DocumentStore,
  Layer,
  SampleSource,
  ViewTransform,
  ViewportSize,
} from './types.ts'
import { LIMITS, TILE_SIZE } from './types.ts'
import { compositeInto, sampleDocument } from './composite.ts'
import { clearPyramidCache, maxPyramidLevel } from './pyramid.ts'
import { tileKey } from './tiles.ts'
import { registerCache } from './memory.ts'
import { layerContentBounds } from './document.ts'
import type { LevelTileRange } from './viewport.ts'
import { chooseLevel, dirtyLevelRect, isEmptyTileRange, levelTileRect, tilesByDistance, visibleTileRange } from './viewport.ts'

type Context2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D
type AnyCanvas = HTMLCanvasElement | OffscreenCanvas

const TILE_BYTES = TILE_SIZE * TILE_SIZE * 4
/** Fallback timer that keeps work going where requestAnimationFrame does not fire (hidden windows). */
const FRAME_WATCHDOG_MS = 100
/** Work slice between yields of flatten() / renderToCanvas(). */
const FLATTEN_SLICE_MS = 12
/** A window more than this many times larger than it needs to be is rebuilt smaller. */
const WINDOW_SHRINK_FACTOR = 4
const CHECKER_LIGHT = '#ffffff'
const CHECKER_DARK = '#cccccc'
const CHECKER_CSS_PX = 8

const CELL_ABSENT = 0
const CELL_STALE = 1
const CELL_CURRENT = 2

export interface CompositorStats {
  readonly frames: number
  readonly tilesComposited: number
  /** Tiles refreshed by recompositing only their out-of-date part. */
  readonly partialRefreshes: number
  readonly tilesUploaded: number
  /** Compositing time of the last frame (ms). */
  readonly lastWorkMs: number
  /** Visible tiles of the active level that are not current yet. */
  readonly pendingVisible: number
  readonly level: number
  readonly cachedTiles: number
  readonly cacheBytes: number
  readonly windowBytes: number
}

/** The shared Compositor contract plus what CanvasView and the perf smoke use. */
export interface AdvancedCompositor extends Compositor {
  /** Coarser proxy while true (slider drags, pinch); false refines at the full level. */
  setInteractive(active: boolean): void
  readonly interactive: boolean
  /** Called after each drawn frame (overlay redraws can follow it). Returns the unsubscribe function. */
  onFrame(listener: () => void): () => void
  stats(): CompositorStats
}

export interface CompositorOptions {
  /** Compositing budget per animation frame (ms). Default LIMITS.frameBudgetMs. */
  readonly frameBudgetMs?: number
  /** Composited tiles kept in the CPU cache. Default LIMITS.compositeTileCacheTiles. */
  readonly cacheTiles?: number
}

// ---------------------------------------------------------------------------------------------
// Small helpers
// ---------------------------------------------------------------------------------------------

function now(): number {
  return typeof performance !== 'undefined' && typeof performance.now === 'function' ? performance.now() : Date.now()
}

function abortError(): Error {
  const error = new Error('The operation was cancelled.')
  error.name = 'AbortError'
  return error
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal && signal.aborted) throw abortError()
}

function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => {
    if (typeof MessageChannel !== 'undefined') {
      const channel = new MessageChannel()
      channel.port1.onmessage = () => {
        channel.port1.close()
        resolve()
      }
      channel.port2.postMessage(null)
    } else {
      setTimeout(resolve, 0)
    }
  })
}

function createCanvas(width: number, height: number): AnyCanvas {
  const w = Math.max(1, Math.floor(width))
  const h = Math.max(1, Math.floor(height))
  if (typeof OffscreenCanvas !== 'undefined') return new OffscreenCanvas(w, h)
  const canvas = document.createElement('canvas')
  canvas.width = w
  canvas.height = h
  return canvas
}

function release(canvas: AnyCanvas | null | undefined): void {
  if (!canvas) return
  canvas.width = 1
  canvas.height = 1
}

function context2d(canvas: AnyCanvas): Context2D {
  const context = (canvas as HTMLCanvasElement).getContext('2d') as Context2D | null
  if (!context) throw new Error('A drawing surface could not be created. Close other large images and try again.')
  return context
}

function hasAlpha(data: Uint8ClampedArray): boolean {
  for (let p = 3; p < data.length; p += 4) if (data[p] !== 0) return true
  return false
}

function rangeContains(outer: LevelTileRange, inner: LevelTileRange): boolean {
  return inner.tx0 >= outer.tx0 && inner.ty0 >= outer.ty0 && inner.tx1 <= outer.tx1 && inner.ty1 <= outer.ty1
}

function rangeArea(range: LevelTileRange): number {
  return Math.max(0, range.tx1 - range.tx0) * Math.max(0, range.ty1 - range.ty0)
}

function grow(range: LevelTileRange, margin: number, columns: number, rows: number): LevelTileRange {
  return {
    tx0: Math.max(0, range.tx0 - margin),
    ty0: Math.max(0, range.ty0 - margin),
    tx1: Math.min(columns, range.tx1 + margin),
    ty1: Math.min(rows, range.ty1 + margin),
  }
}

function cacheKey(level: number, tx: number, ty: number): number {
  return level * 4294967296 + tileKey(tx, ty)
}

/** Document-space area a preview changes ('all' when it cannot be bounded). */
function previewArea(preview: CompositePreview | null, state: DocumentState): readonly IntRect[] | 'all' {
  if (!preview) return []
  const layer: Layer | undefined = state.layers.find((item) => item.id === preview.layerId)
  if (!layer) return 'all'
  if (preview.kind === 'layer-pixels') {
    const scale = 2 ** Math.max(0, preview.level | 0)
    const rect = preview.rect
    const area: IntRect[] = [{ x: rect.x * scale, y: rect.y * scale, width: rect.width * scale, height: rect.height * scale }]
    const bounds = layerContentBounds(layer)
    if (bounds) area.push(bounds)
    return area
  }
  // Adjustment layers change everything below them; a layer that is a clipping base shapes its group.
  if (layer.kind === 'adjustment') return 'all'
  const bounds = layerContentBounds(layer)
  return bounds ? [bounds] : []
}

// ---------------------------------------------------------------------------------------------
// Composited-tile cache (CPU, LRU)
// ---------------------------------------------------------------------------------------------

interface CachedTile {
  readonly width: number
  readonly height: number
  /** null = fully transparent. */
  data: Uint8ClampedArray | null
  readonly epoch: number
  /** Computed while a preview was active (never reused for flatten/export). */
  previewed: boolean
  /** The whole tile is out of date. */
  stale: boolean
  /** Out-of-date part of an otherwise current tile (tile-local level pixels), or null. */
  dirty: IntRect | null
}

/** The part of a level-pixel rectangle inside tile (tx, ty), in tile-local pixels, clipped to width x height. */
function localDirty(rect: IntRect, tx: number, ty: number, width: number, height: number): IntRect | null {
  const left = tx * TILE_SIZE
  const top = ty * TILE_SIZE
  const x0 = Math.max(0, rect.x - left)
  const y0 = Math.max(0, rect.y - top)
  const x1 = Math.min(width, rect.x + rect.width - left)
  const y1 = Math.min(height, rect.y + rect.height - top)
  return x1 > x0 && y1 > y0 ? { x: x0, y: y0, width: x1 - x0, height: y1 - y0 } : null
}

function unionRect(a: IntRect | null | undefined, b: IntRect): IntRect {
  if (!a) return b
  const x0 = Math.min(a.x, b.x)
  const y0 = Math.min(a.y, b.y)
  return { x: x0, y: y0, width: Math.max(a.x + a.width, b.x + b.width) - x0, height: Math.max(a.y + a.height, b.y + b.height) - y0 }
}

function covers(rect: IntRect, width: number, height: number): boolean {
  return rect.x <= 0 && rect.y <= 0 && rect.x + rect.width >= width && rect.y + rect.height >= height
}

function rectContains(outer: IntRect, inner: IntRect): boolean {
  return inner.x >= outer.x && inner.y >= outer.y && inner.x + inner.width <= outer.x + outer.width
    && inner.y + inner.height <= outer.y + outer.height
}

function rectTileRange(rect: IntRect): LevelTileRange {
  return {
    tx0: Math.floor(rect.x / TILE_SIZE),
    ty0: Math.floor(rect.y / TILE_SIZE),
    tx1: Math.floor((rect.x + rect.width - 1) / TILE_SIZE) + 1,
    ty1: Math.floor((rect.y + rect.height - 1) / TILE_SIZE) + 1,
  }
}

class TileCache {
  private readonly tiles = new Map<number, CachedTile>()
  private bytesHeld = 0
  private limit: number
  epoch = 1

  constructor(limit: number) {
    this.limit = Math.max(0, Math.floor(limit))
  }

  get size(): number {
    return this.tiles.size
  }

  get bytes(): number {
    return this.bytesHeld
  }

  /** A fully current entry (refreshing its LRU position), or null. */
  get(key: number): CachedTile | null {
    const entry = this.peek(key)
    return entry && !entry.dirty ? entry : null
  }

  /** An entry of the current epoch that is not wholly stale (it may have a dirty part), or null. */
  peek(key: number): CachedTile | null {
    const entry = this.tiles.get(key)
    if (!entry || entry.stale || entry.epoch !== this.epoch) return null
    this.tiles.delete(key)
    this.tiles.set(key, entry)
    return entry
  }

  set(key: number, entry: CachedTile): void {
    const previous = this.tiles.get(key)
    if (previous) {
      this.bytesHeld -= previous.data ? previous.data.byteLength : 0
      this.tiles.delete(key)
    }
    this.tiles.set(key, entry)
    this.bytesHeld += entry.data ? entry.data.byteLength : 0
    this.trimTo(this.limit)
  }

  /** Gives a transparent (data-less) entry a pixel buffer. */
  setData(entry: CachedTile, data: Uint8ClampedArray): void {
    this.bytesHeld += data.byteLength - (entry.data ? entry.data.byteLength : 0)
    entry.data = data
  }

  private markEntry(entry: CachedTile, rect: IntRect, tx: number, ty: number): void {
    if (entry.stale) return
    const local = localDirty(rect, tx, ty, entry.width, entry.height)
    if (!local) return
    const dirty = unionRect(entry.dirty, local)
    if (covers(dirty, entry.width, entry.height)) {
      entry.stale = true
      entry.dirty = null
    } else {
      entry.dirty = dirty
    }
  }

  /** Marks what `rect` (level pixels of `level`) covers out of date (walks whichever is smaller). */
  markRect(level: number, rect: IntRect): void {
    const range = rectTileRange(rect)
    if (rangeArea(range) <= this.tiles.size) {
      for (let ty = range.ty0; ty < range.ty1; ty += 1) {
        for (let tx = range.tx0; tx < range.tx1; tx += 1) {
          const entry = this.tiles.get(cacheKey(level, tx, ty))
          if (entry) this.markEntry(entry, rect, tx, ty)
        }
      }
      return
    }
    const low = cacheKey(level, range.tx0, range.ty0)
    const high = cacheKey(level, range.tx1 - 1, range.ty1 - 1)
    for (const [key, entry] of this.tiles) {
      if (key < low || key > high) continue
      const local = key - level * 4294967296
      const tx = (local % 0x10000) - 0x8000
      const ty = Math.floor(local / 0x10000) - 0x8000
      if (tx >= range.tx0 && tx < range.tx1 && ty >= range.ty0 && ty < range.ty1) this.markEntry(entry, rect, tx, ty)
    }
  }

  invalidateAll(): void {
    this.epoch += 1
  }

  trimTo(limit: number): void {
    for (const [key, entry] of this.tiles) {
      if (this.tiles.size <= limit) break
      this.tiles.delete(key)
      this.bytesHeld -= entry.data ? entry.data.byteLength : 0
    }
  }

  setLimit(limit: number): void {
    this.limit = Math.max(0, Math.floor(limit))
    this.trimTo(this.limit)
  }

  clear(): void {
    this.tiles.clear()
    this.bytesHeld = 0
  }
}

// ---------------------------------------------------------------------------------------------
// Level windows (GPU)
// ---------------------------------------------------------------------------------------------

interface LevelWindow {
  readonly level: number
  readonly range: LevelTileRange
  readonly canvas: AnyCanvas
  readonly context: Context2D
  /** tileKey -> CELL_STALE | CELL_CURRENT (absent = never filled). */
  readonly cells: Map<number, number>
  /** Stale cells whose out-of-date part is known (tile-local); other stale cells are wholly stale. */
  readonly partial: Map<number, IntRect>
}

function createWindow(level: number, range: LevelTileRange): LevelWindow {
  const canvas = createCanvas((range.tx1 - range.tx0) * TILE_SIZE, (range.ty1 - range.ty0) * TILE_SIZE)
  try {
    const context = context2d(canvas)
    context.imageSmoothingEnabled = false
    return { level, range, canvas, context, cells: new Map(), partial: new Map() }
  } catch (error) {
    release(canvas)
    throw error
  }
}

function windowBytes(win: LevelWindow | null): number {
  return win ? win.canvas.width * win.canvas.height * 4 : 0
}

/** Marks the window's cells under `rect` (level pixels; 'all' = everything) out of date. */
function markWindowStale(win: LevelWindow | null, rect: IntRect | 'all'): void {
  if (!win || !win.cells.size) return
  if (rect === 'all') {
    for (const key of win.cells.keys()) win.cells.set(key, CELL_STALE)
    win.partial.clear()
    return
  }
  const range = rectTileRange(rect)
  const tx0 = Math.max(range.tx0, win.range.tx0)
  const ty0 = Math.max(range.ty0, win.range.ty0)
  const tx1 = Math.min(range.tx1, win.range.tx1)
  const ty1 = Math.min(range.ty1, win.range.ty1)
  for (let ty = ty0; ty < ty1; ty += 1) {
    for (let tx = tx0; tx < tx1; tx += 1) {
      const key = tileKey(tx, ty)
      const state = win.cells.get(key)
      if (state === undefined) continue
      const local = localDirty(rect, tx, ty, TILE_SIZE, TILE_SIZE)
      if (!local) continue
      if (state === CELL_CURRENT) {
        win.cells.set(key, CELL_STALE)
        if (covers(local, TILE_SIZE, TILE_SIZE)) win.partial.delete(key)
        else win.partial.set(key, local)
        continue
      }
      // Already stale: grow a known part; a wholly stale cell stays wholly stale.
      const known = win.partial.get(key)
      if (!known) continue
      const grown = unionRect(known, local)
      if (covers(grown, TILE_SIZE, TILE_SIZE)) win.partial.delete(key)
      else win.partial.set(key, grown)
    }
  }
}

// ---------------------------------------------------------------------------------------------
// The compositor
// ---------------------------------------------------------------------------------------------

export function createCompositor(store: DocumentStore, options: CompositorOptions = {}): AdvancedCompositor {
  const budgetMs = Number.isFinite(options.frameBudgetMs) && Number(options.frameBudgetMs) > 0 ? Number(options.frameBudgetMs) : LIMITS.frameBudgetMs
  const cache = new TileCache(Number.isFinite(options.cacheTiles) ? Math.max(0, Number(options.cacheTiles)) : LIMITS.compositeTileCacheTiles)

  let screen: HTMLCanvasElement | null = null
  let screenContext: CanvasRenderingContext2D | null = null
  let view: ViewTransform = { zoom: 1, offsetX: 0, offsetY: 0 }
  let viewport: ViewportSize = { width: 0, height: 0, dpr: 1 }
  let docWidth = store.getState().width
  let docHeight = store.getState().height
  let maxLevel = maxPyramidLevel(docWidth, docHeight)
  let interactive = false
  let preview: CompositePreview | null = null
  let current: LevelWindow | null = null
  let fallback: LevelWindow | null = null
  let checker: { pattern: CanvasPattern; size: number; tile: AnyCanvas } | null = null
  let disposed = false
  let frameScheduled = false
  /** True when rAF stopped firing (hidden window): frames then run on a 16 ms timer. */
  let rafStalled = false
  /** True until a frame has drawn the latest view/state; settle() never resolves before that. */
  let needsDraw = true
  /** Consecutive failures to allocate a level window (bounded retries). */
  let windowFailures = 0
  let rafHandle: number | null = null
  let timerHandle: ReturnType<typeof setTimeout> | null = null
  /** Bumped by every document change; flatten()/renderToCanvas() restart when it moves under them. */
  let changeCounter = 0
  let settleWaiters: (() => void)[] = []
  const frameListeners = new Set<() => void>()
  const counters = { frames: 0, tilesComposited: 0, partialRefreshes: 0, tilesUploaded: 0, lastWorkMs: 0, pendingVisible: 0 }

  // ----- levels and ranges ------------------------------------------------------------------

  function baseLevel(): number {
    return chooseLevel(view.zoom, viewport.dpr, maxLevel)
  }

  function activeLevel(): number {
    const base = baseLevel()
    return interactive ? Math.min(maxLevel, base + 1) : base
  }

  function counts(level: number): { columns: number; rows: number } {
    const scale = 2 ** level
    return {
      columns: Math.ceil(Math.ceil(docWidth / scale) / TILE_SIZE),
      rows: Math.ceil(Math.ceil(docHeight / scale) / TILE_SIZE),
    }
  }

  function visibleRange(level: number): LevelTileRange {
    return visibleTileRange(view, viewport, { width: docWidth, height: docHeight }, level, 0)
  }

  function hasViewport(): boolean {
    return Boolean(screen) && viewport.width > 0 && viewport.height > 0 && view.zoom > 0
  }

  // ----- invalidation -----------------------------------------------------------------------

  function invalidateRects(dirty: readonly IntRect[] | 'all'): void {
    if (dirty === 'all') {
      cache.invalidateAll()
      markWindowStale(current, 'all')
      markWindowStale(fallback, 'all')
      return
    }
    for (const rect of dirty) {
      if (!rect || !(rect.width > 0 && rect.height > 0)) continue
      for (let level = 0; level <= maxLevel; level += 1) {
        const levelRect = dirtyLevelRect(rect, level)
        if (!levelRect) continue
        cache.markRect(level, levelRect)
        if (current && current.level === level) markWindowStale(current, levelRect)
        if (fallback && fallback.level === level) markWindowStale(fallback, levelRect)
      }
    }
  }

  function resetForSize(state: Pick<DocumentState, 'width' | 'height'>): void {
    docWidth = state.width
    docHeight = state.height
    maxLevel = maxPyramidLevel(docWidth, docHeight)
    cache.clear()
    cache.invalidateAll()
    release(current?.canvas)
    release(fallback?.canvas)
    current = null
    fallback = null
  }

  const unsubscribe = store.subscribe((change: DocumentChange) => {
    if (disposed) return
    changeCounter += 1
    const state = store.getState()
    if (state.width !== docWidth || state.height !== docHeight) resetForSize(state)
    else invalidateRects(change.dirty)
    changed()
  })

  const unregister = registerCache({
    name: 'composite',
    bytes: () => cache.bytes + windowBytes(current) + windowBytes(fallback),
    trim: () => cache.clear(),
  })

  // ----- tiles ------------------------------------------------------------------------------

  function computeTile(level: number, tx: number, ty: number): CachedTile {
    const rect = levelTileRect({ width: docWidth, height: docHeight }, level, tx, ty)
    const key = cacheKey(level, tx, ty)
    if (!(rect.width > 0 && rect.height > 0)) {
      const empty: CachedTile = { width: 0, height: 0, data: null, epoch: cache.epoch, previewed: false, stale: false, dirty: null }
      cache.set(key, empty)
      return empty
    }
    const target: PixelBuffer = { width: rect.width, height: rect.height, data: new Uint8ClampedArray(rect.width * rect.height * 4) }
    compositeInto(target, 0, 0, store.getState(), rect, { level, preview })
    counters.tilesComposited += 1
    const entry: CachedTile = {
      width: rect.width,
      height: rect.height,
      data: hasAlpha(target.data) ? target.data : null,
      epoch: cache.epoch,
      previewed: preview !== null,
      stale: false,
      dirty: null,
    }
    cache.set(key, entry)
    return entry
  }

  /**
   * Recomposites only the out-of-date part of a cached tile, in place. Composite blocks are
   * position-stable (composite.ts), so the bytes equal a whole-tile composite. Returns that part.
   */
  function refreshPart(entry: CachedTile, level: number, tx: number, ty: number): IntRect {
    const part = entry.dirty as IntRect
    let data = entry.data
    if (!data) {
      data = new Uint8ClampedArray(entry.width * entry.height * 4)
      cache.setData(entry, data)
    }
    compositeInto({ width: entry.width, height: entry.height, data }, part.x, part.y, store.getState(),
      { x: tx * TILE_SIZE + part.x, y: ty * TILE_SIZE + part.y, width: part.width, height: part.height }, { level, preview })
    counters.partialRefreshes += 1
    entry.dirty = null
    if (preview) entry.previewed = true
    return part
  }

  /** Uploads a current tile (or only `region` of it, tile-local) into its window cell. */
  function upload(win: LevelWindow, tx: number, ty: number, tile: CachedTile, region: IntRect | null): void {
    const x = (tx - win.range.tx0) * TILE_SIZE
    const y = (ty - win.range.ty0) * TILE_SIZE
    if (tile.data) {
      const image = new ImageData(tile.data as Uint8ClampedArray<ArrayBuffer>, tile.width, tile.height)
      if (region) {
        win.context.putImageData(image, x, y, region.x, region.y, region.width, region.height)
      } else {
        win.context.putImageData(image, x, y)
        // Edge tiles: keep the area outside the document transparent.
        if (tile.width < TILE_SIZE) win.context.clearRect(x + tile.width, y, TILE_SIZE - tile.width, TILE_SIZE)
        if (tile.height < TILE_SIZE) win.context.clearRect(x, y + tile.height, tile.width, TILE_SIZE - tile.height)
      }
    } else {
      win.context.clearRect(x, y, TILE_SIZE, TILE_SIZE)
    }
    counters.tilesUploaded += 1
    const key = tileKey(tx, ty)
    win.cells.set(key, CELL_CURRENT)
    win.partial.delete(key)
  }

  function refreshCell(win: LevelWindow, tx: number, ty: number): void {
    const key = tileKey(tx, ty)
    const known = win.cells.get(key) === CELL_STALE ? win.partial.get(key) ?? null : null
    let tile = cache.peek(cacheKey(win.level, tx, ty))
    let region: IntRect | null = null
    if (tile && tile.dirty) {
      // Brush dabs and small edits: recomposite and upload only what changed.
      const refreshed = refreshPart(tile, win.level, tx, ty)
      if (known && rectContains(refreshed, known)) region = refreshed
    } else if (!tile) {
      tile = computeTile(win.level, tx, ty)
    }
    upload(win, tx, ty, tile, region)
  }

  // ----- windows ----------------------------------------------------------------------------

  /** Makes sure the current window is at `level` and covers `visible`; returns it (null when nothing is visible). */
  function ensureWindow(level: number, visible: LevelTileRange): LevelWindow | null {
    if (isEmptyTileRange(visible)) return current && current.level === level ? current : null
    const { columns, rows } = counts(level)
    const wanted = grow(visible, 1, columns, rows)
    if (current && current.level === level && rangeContains(current.range, visible)
      && rangeArea(current.range) <= Math.max(16, rangeArea(wanted) * WINDOW_SHRINK_FACTOR)) {
      return current
    }
    const next = createWindow(level, wanted)
    if (current && current.level === level) {
      // Same level, new area: copy the overlap on the GPU and keep its cell states.
      const dx = (current.range.tx0 - wanted.tx0) * TILE_SIZE
      const dy = (current.range.ty0 - wanted.ty0) * TILE_SIZE
      next.context.drawImage(current.canvas, dx, dy)
      for (const [key, state] of current.cells) {
        const tx = (key % 0x10000) - 0x8000
        const ty = Math.floor(key / 0x10000) - 0x8000
        if (!(tx >= wanted.tx0 && tx < wanted.tx1 && ty >= wanted.ty0 && ty < wanted.ty1)) continue
        next.cells.set(key, state)
        const known = current.partial.get(key)
        if (known) next.partial.set(key, known)
      }
      release(current.canvas)
    } else if (current) {
      // Level change: the old window stays underneath as a placeholder until the new one is filled.
      if (current.cells.size) {
        release(fallback?.canvas)
        fallback = current
      } else {
        release(current.canvas)
      }
    }
    current = next
    return current
  }

  function dropFallbackIfCovered(visible: LevelTileRange): void {
    if (!fallback || !current) return
    for (let ty = visible.ty0; ty < visible.ty1; ty += 1) {
      for (let tx = visible.tx0; tx < visible.tx1; tx += 1) {
        if (!current.cells.has(tileKey(tx, ty))) return
      }
    }
    release(fallback.canvas)
    fallback = null
  }

  // ----- drawing ----------------------------------------------------------------------------

  function dropChecker(): void {
    release(checker?.tile)
    checker = null
  }

  /** Photoshop's transparency grid: 8 CSS px squares, anchored to the document origin. */
  function checkerPattern(context: CanvasRenderingContext2D): CanvasPattern | null {
    const square = Math.max(2, Math.round(CHECKER_CSS_PX * (viewport.dpr || 1)))
    if (checker && checker.size === square) return checker.pattern
    dropChecker()
    const tile = createCanvas(square * 2, square * 2)
    const c = context2d(tile)
    c.fillStyle = CHECKER_LIGHT
    c.fillRect(0, 0, square * 2, square * 2)
    c.fillStyle = CHECKER_DARK
    c.fillRect(square, 0, square, square)
    c.fillRect(0, square, square, square)
    const pattern = context.createPattern(tile as CanvasImageSource, 'repeat')
    if (!pattern) {
      release(tile)
      return null
    }
    // The tiny source canvas stays alive with the pattern.
    checker = { pattern, size: square, tile }
    return pattern
  }

  interface Placement {
    readonly originX: number
    readonly originY: number
    readonly scale: number
  }

  function placement(): Placement {
    const dpr = viewport.dpr || 1
    // The document origin snaps to a device pixel so 100% views blit pixel-exact.
    return { originX: Math.round(view.offsetX * dpr), originY: Math.round(view.offsetY * dpr), scale: view.zoom * dpr }
  }

  function drawWindow(context: CanvasRenderingContext2D, win: LevelWindow, place: Placement, width: number, height: number): void {
    const levelPx = 2 ** win.level * place.scale
    if (!(levelPx > 0)) return
    // Source rectangle: the part of the window that is on screen.
    const left = win.range.tx0 * TILE_SIZE
    const top = win.range.ty0 * TILE_SIZE
    const sx0 = Math.max(left, Math.floor(-place.originX / levelPx))
    const sy0 = Math.max(top, Math.floor(-place.originY / levelPx))
    const sx1 = Math.min(left + win.canvas.width, Math.ceil((width - place.originX) / levelPx))
    const sy1 = Math.min(top + win.canvas.height, Math.ceil((height - place.originY) / levelPx))
    if (sx1 <= sx0 || sy1 <= sy0) return
    context.imageSmoothingEnabled = levelPx < 1 - 1e-9 || (levelPx > 1 + 1e-9 && win.level > 0)
    if (context.imageSmoothingEnabled) context.imageSmoothingQuality = 'medium'
    context.drawImage(win.canvas as CanvasImageSource, sx0 - left, sy0 - top, sx1 - sx0, sy1 - sy0,
      place.originX + sx0 * levelPx, place.originY + sy0 * levelPx, (sx1 - sx0) * levelPx, (sy1 - sy0) * levelPx)
  }

  function draw(visible: LevelTileRange): void {
    if (!screen || !screenContext) return
    const context = screenContext
    const width = screen.width
    const height = screen.height
    context.setTransform(1, 0, 0, 1, 0, 0)
    context.globalAlpha = 1
    context.globalCompositeOperation = 'source-over'
    context.clearRect(0, 0, width, height)
    const place = placement()
    const docW = docWidth * place.scale
    const docH = docHeight * place.scale
    const x0 = Math.max(0, place.originX)
    const y0 = Math.max(0, place.originY)
    const x1 = Math.min(width, place.originX + docW)
    const y1 = Math.min(height, place.originY + docH)
    if (x1 <= x0 || y1 <= y0) return
    context.save()
    try {
      context.beginPath()
      context.rect(x0, y0, x1 - x0, y1 - y0)
      context.clip()
      const pattern = checkerPattern(context)
      if (pattern) {
        context.save()
        context.translate(place.originX, place.originY)
        context.fillStyle = pattern
        context.fillRect(x0 - place.originX, y0 - place.originY, x1 - x0, y1 - y0)
        context.restore()
      }
      if (fallback && current) {
        // The placeholder shows only where the current window has nothing yet.
        const levelPx = 2 ** current.level * place.scale
        let any = false
        context.save()
        context.beginPath()
        for (let ty = visible.ty0; ty < visible.ty1; ty += 1) {
          for (let tx = visible.tx0; tx < visible.tx1; tx += 1) {
            if (current.cells.has(tileKey(tx, ty))) continue
            any = true
            context.rect(place.originX + tx * TILE_SIZE * levelPx, place.originY + ty * TILE_SIZE * levelPx,
              TILE_SIZE * levelPx, TILE_SIZE * levelPx)
          }
        }
        if (any) {
          context.clip()
          drawWindow(context, fallback, place, width, height)
        }
        context.restore()
      } else if (fallback && !current) {
        drawWindow(context, fallback, place, width, height)
      }
      if (current) drawWindow(context, current, place, width, height)
    } finally {
      context.restore()
    }
  }

  // ----- the frame loop ---------------------------------------------------------------------

  function cancelFrame(): void {
    if (rafHandle !== null && typeof cancelAnimationFrame === 'function') cancelAnimationFrame(rafHandle)
    if (timerHandle !== null) clearTimeout(timerHandle)
    rafHandle = null
    timerHandle = null
    frameScheduled = false
  }

  /** Marks that the screen no longer shows the current state (settle() waits for the next drawn frame). */
  function changed(): void {
    needsDraw = true
    schedule()
  }

  function schedule(): void {
    if (disposed || frameScheduled) return
    frameScheduled = true
    const hidden = typeof document !== 'undefined' && document.visibilityState === 'hidden'
    const useRaf = !hidden && typeof requestAnimationFrame === 'function'
    if (useRaf) {
      rafHandle = requestAnimationFrame(() => {
        rafStalled = false
        cancelFrame()
        frame()
      })
    }
    // requestAnimationFrame does not fire in hidden, minimised or never-shown windows. A timer keeps the
    // work going there; once it had to stand in for a frame it runs at frame rate until rAF fires again.
    const delay = !useRaf || hidden || rafStalled ? 16 : FRAME_WATCHDOG_MS
    timerHandle = setTimeout(() => {
      if (useRaf) rafStalled = true
      cancelFrame()
      frame()
    }, delay)
  }

  function pendingVisibleCount(win: LevelWindow | null, visible: LevelTileRange): number {
    if (isEmptyTileRange(visible)) return 0
    if (!win) return rangeArea(visible)
    let pending = 0
    for (let ty = visible.ty0; ty < visible.ty1; ty += 1) {
      for (let tx = visible.tx0; tx < visible.tx1; tx += 1) {
        if (win.cells.get(tileKey(tx, ty)) !== CELL_CURRENT) pending += 1
      }
    }
    return pending
  }

  function resolveSettled(): void {
    const waiters = settleWaiters
    settleWaiters = []
    for (const resolve of waiters) resolve()
  }

  function frame(): void {
    if (disposed) return
    counters.frames += 1
    if (!hasViewport()) {
      counters.pendingVisible = 0
      resolveSettled()
      return
    }
    const level = activeLevel()
    const visible = visibleRange(level)
    let win: LevelWindow | null = null
    try {
      win = ensureWindow(level, visible)
      windowFailures = 0
    } catch (error) {
      // No GPU surface right now (out of memory): retry a few times with a pause instead of every frame.
      console.error(error)
      windowFailures += 1
      resolveSettled()
      if (windowFailures <= 3) {
        setTimeout(() => {
          if (!disposed) changed()
        }, 250 * windowFailures)
      }
      return
    }
    const started = now()
    let work = 0
    let failed = false
    if (win) {
      const center = {
        x: ((viewport.width / 2 - view.offsetX) / view.zoom) / (TILE_SIZE * 2 ** level),
        y: ((viewport.height / 2 - view.offsetY) / view.zoom) / (TILE_SIZE * 2 ** level),
      }
      const ring = grow(visible, 1, counts(level).columns, counts(level).rows)
      const queue = tilesByDistance(ring, center)
      // Visible tiles first, then the ring.
      queue.sort((a, b) => Number(!inRange(visible, a.tx, a.ty)) - Number(!inRange(visible, b.tx, b.ty)))
      try {
        for (const { tx, ty } of queue) {
          if (!inRange(win.range, tx, ty)) continue
          if (win.cells.get(tileKey(tx, ty)) === CELL_CURRENT) continue
          if (work > 0 && now() - started >= budgetMs) {
            work = -1
            break
          }
          refreshCell(win, tx, ty)
          work += 1
        }
      } catch (error) {
        // A tile that cannot be composited must not spin the loop or hang settle(); the next change retries.
        console.error(error)
        failed = true
      }
    }
    counters.lastWorkMs = now() - started
    counters.pendingVisible = pendingVisibleCount(win, visible)
    if (win) dropFallbackIfCovered(visible)
    draw(visible)
    needsDraw = false
    for (const listener of [...frameListeners]) {
      try {
        listener()
      } catch (error) {
        console.error(error)
      }
    }
    if (!failed && (work < 0 || counters.pendingVisible > 0 || (win && ringPending(win, visible, level)))) {
      schedule()
    }
    if (failed || counters.pendingVisible === 0) resolveSettled()
  }

  function inRange(range: LevelTileRange, tx: number, ty: number): boolean {
    return tx >= range.tx0 && tx < range.tx1 && ty >= range.ty0 && ty < range.ty1
  }

  function ringPending(win: LevelWindow, visible: LevelTileRange, level: number): boolean {
    const ring = grow(visible, 1, counts(level).columns, counts(level).rows)
    for (let ty = ring.ty0; ty < ring.ty1; ty += 1) {
      for (let tx = ring.tx0; tx < ring.tx1; tx += 1) {
        if (!inRange(win.range, tx, ty)) continue
        if (win.cells.get(tileKey(tx, ty)) !== CELL_CURRENT) return true
      }
    }
    return false
  }

  // ----- full-resolution output -------------------------------------------------------------

  interface RowSink {
    /** A pass starts (again) for a width x height document. */
    begin(width: number, height: number): void
    /** One finished tile row: `rows.height` rows starting at document row `top`. */
    row(top: number, rows: PixelBuffer): void
  }

  /**
   * Composites level 0 tile row by tile row into `sink`, yielding to the event loop every few ms. When the
   * document changes during a pass the pass starts over, so the output is always one consistent state; the
   * third attempt runs without yielding, so it always finishes. Uncommitted previews are never included.
   */
  async function streamRows(optionsIn: OpOptions | undefined, sink: RowSink): Promise<void> {
    const signal = optionsIn?.signal
    const onProgress = optionsIn?.onProgress
    for (let attempt = 0; attempt < 3; attempt += 1) {
      throwIfAborted(signal)
      const yielding = attempt < 2
      const startCounter = changeCounter
      const state = store.getState()
      const width = state.width
      const height = state.height
      sink.begin(width, height)
      let restarted = false
      let sliceStart = now()
      const rowCount = Math.ceil(height / TILE_SIZE)
      const columns = Math.ceil(width / TILE_SIZE)
      for (let row = 0; row < rowCount; row += 1) {
        const top = row * TILE_SIZE
        const rows = Math.min(TILE_SIZE, height - top)
        const buffer: PixelBuffer = { width, height: rows, data: new Uint8ClampedArray(width * rows * 4) }
        for (let tx = 0; tx < columns; tx += 1) {
          const left = tx * TILE_SIZE
          const cols = Math.min(TILE_SIZE, width - left)
          // A display tile is reusable when it is current and was not computed under a preview.
          const cached = cache.get(cacheKey(0, tx, row))
          if (cached && !cached.previewed && cached.width === cols && cached.height === rows) {
            if (cached.data) {
              for (let y = 0; y < rows; y += 1) {
                buffer.data.set(cached.data.subarray(y * cols * 4, (y + 1) * cols * 4), (y * width + left) * 4)
              }
            }
          } else {
            compositeInto(buffer, left, 0, state, { x: left, y: top, width: cols, height: rows }, { level: 0 })
          }
        }
        sink.row(top, buffer)
        if (onProgress) onProgress(Math.min(1, (top + rows) / Math.max(1, height)))
        if (yielding && now() - sliceStart >= FLATTEN_SLICE_MS) {
          await yieldToEventLoop()
          throwIfAborted(signal)
          sliceStart = now()
          if (changeCounter !== startCounter || disposed) {
            restarted = true
            break
          }
        }
      }
      if (disposed) throw new Error('This document has been closed.')
      if (!restarted) return
    }
  }

  async function flatten(optionsIn?: OpOptions): Promise<PixelBuffer> {
    if (disposed) throw new Error('This document has been closed.')
    let out: PixelBuffer = { width: 0, height: 0, data: new Uint8ClampedArray(0) }
    await streamRows(optionsIn, {
      begin(width, height) {
        if (out.width !== width || out.height !== height) out = { width, height, data: new Uint8ClampedArray(width * height * 4) }
      },
      row(top, rows) {
        out.data.set(rows.data, top * rows.width * 4)
      },
    })
    return out
  }

  async function renderToCanvas(optionsIn?: OpOptions): Promise<HTMLCanvasElement> {
    if (disposed) throw new Error('This document has been closed.')
    if (typeof document === 'undefined') throw new Error('Rendering to a canvas needs a window.')
    const canvas = document.createElement('canvas')
    try {
      let context: CanvasRenderingContext2D | null = null
      await streamRows(optionsIn, {
        begin(width, height) {
          // Setting the size also clears whatever an earlier pass drew.
          canvas.width = width
          canvas.height = height
          context = canvas.getContext('2d')
          if (!context) throw new Error('A drawing surface could not be created. Close other large images and try again.')
        },
        row(top, rows) {
          (context as CanvasRenderingContext2D).putImageData(new ImageData(rows.data as Uint8ClampedArray<ArrayBuffer>, rows.width, rows.height), 0, top)
        },
      })
      return canvas
    } catch (error) {
      release(canvas)
      throw error
    }
  }

  // ----- public object ----------------------------------------------------------------------

  const compositor: AdvancedCompositor = {
    get level() {
      return activeLevel()
    },
    get interactive() {
      return interactive
    },
    attach(canvas: HTMLCanvasElement): void {
      if (disposed) return
      screen = canvas ?? null
      screenContext = screen ? screen.getContext('2d') : null
      dropChecker()
      if (screen) {
        const width = Math.max(1, Math.round(viewport.width * (viewport.dpr || 1)))
        const height = Math.max(1, Math.round(viewport.height * (viewport.dpr || 1)))
        if (viewport.width > 0 && (screen.width !== width || screen.height !== height)) {
          screen.width = width
          screen.height = height
        }
      }
      changed()
    },
    setView(nextView: ViewTransform, size: ViewportSize): void {
      if (disposed) return
      const zoom = Number.isFinite(nextView.zoom) && nextView.zoom > 0 ? nextView.zoom : view.zoom
      view = {
        zoom,
        offsetX: Number.isFinite(nextView.offsetX) ? nextView.offsetX : 0,
        offsetY: Number.isFinite(nextView.offsetY) ? nextView.offsetY : 0,
      }
      const dpr = Number.isFinite(size.dpr) && size.dpr > 0 ? size.dpr : 1
      if (dpr !== viewport.dpr) dropChecker()
      viewport = { width: Math.max(0, Number(size.width) || 0), height: Math.max(0, Number(size.height) || 0), dpr }
      if (screen) {
        const width = Math.max(1, Math.round(viewport.width * dpr))
        const height = Math.max(1, Math.round(viewport.height * dpr))
        if (screen.width !== width || screen.height !== height) {
          screen.width = width
          screen.height = height
        }
      }
      changed()
    },
    invalidate(dirty: readonly IntRect[] | 'all'): void {
      if (disposed) return
      invalidateRects(dirty)
      changed()
    },
    setPreview(next: CompositePreview | null): void {
      if (disposed) return
      const state = store.getState()
      const before = previewArea(preview, state)
      preview = next ?? null
      const after = previewArea(preview, state)
      if (before === 'all' || after === 'all') invalidateRects('all')
      else invalidateRects([...before, ...after])
      changed()
    },
    settle(): Promise<void> {
      if (disposed || !hasViewport()) return Promise.resolve()
      if (!needsDraw && !frameScheduled) {
        const level = activeLevel()
        if (current && current.level === level && pendingVisibleCount(current, visibleRange(level)) === 0) return Promise.resolve()
      }
      return new Promise((resolve) => {
        settleWaiters.push(resolve)
        schedule()
      })
    },
    flatten,
    renderToCanvas,
    sample(x: number, y: number, size: number, source: SampleSource): Rgba8 {
      return sampleDocument(store.getState(), x, y, size, source)
    },
    setInteractive(active: boolean): void {
      if (disposed || interactive === Boolean(active)) return
      interactive = Boolean(active)
      changed()
    },
    onFrame(listener: () => void): () => void {
      frameListeners.add(listener)
      return () => { frameListeners.delete(listener) }
    },
    stats(): CompositorStats {
      return {
        frames: counters.frames,
        tilesComposited: counters.tilesComposited,
        partialRefreshes: counters.partialRefreshes,
        tilesUploaded: counters.tilesUploaded,
        lastWorkMs: counters.lastWorkMs,
        pendingVisible: counters.pendingVisible,
        level: activeLevel(),
        cachedTiles: cache.size,
        cacheBytes: cache.bytes,
        windowBytes: windowBytes(current) + windowBytes(fallback),
      }
    },
    dispose(): void {
      if (disposed) return
      disposed = true
      cancelFrame()
      unsubscribe()
      unregister()
      cache.clear()
      release(current?.canvas)
      release(fallback?.canvas)
      current = null
      fallback = null
      dropChecker()
      // Proxy tiles of this document's surfaces would otherwise stay in the global LRU after the editor closes.
      clearPyramidCache()
      screen = null
      screenContext = null
      frameListeners.clear()
      resolveSettled()
    },
  }
  return compositor
}

/** Bytes of one composited tile (for memory estimates). */
export const COMPOSITE_TILE_BYTES = TILE_BYTES
