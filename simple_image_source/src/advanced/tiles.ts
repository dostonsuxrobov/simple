// src/advanced/tiles.ts (WP3)
// Sparse, straight-alpha tiled pixel storage for the Advanced editor (design 5.2).
//   - TiledSurface: RGBA8 tiles of TILE_SIZE x TILE_SIZE (262,144 bytes each); unallocated = transparent.
//   - TiledMask: A8 tiles (65,536 bytes each); unallocated = defaultValue (255 reveals, 0 hides).
// Coordinates are layer-local and may be negative (a layer keeps pixels outside the canvas, like
// Photoshop). write() allocates tiles on demand and frees tiles that become empty (fully transparent, or
// all defaultValue), so an empty layer costs nothing.
// Versions are global and monotonic: every mutation of any grid gets a number larger than every earlier
// one. Per-tile versions feed the composite caches; per-level region versions (the newest tile version
// inside each pyramid tile footprint) let the pyramid validate a proxy tile in O(1).
// Rules for callers: never mutate an array returned by tile() in place; change pixels through write()
// or setTile() (inside a Transaction / StrokeSession editor) so versions and history stay correct.
// setTile() takes ownership of the array it is given (history swaps tiles by reference): never keep
// writing to it and never hand the same array to a second surface; clone() copies.
// Pure and DOM-free (Node tests import it directly).
import type { IntRect, MaskBuffer, PixelBuffer } from '../imaging/types.ts'
import type { TiledMask, TiledSurface } from './types.ts'
import { TILE_SHIFT, TILE_SIZE } from './types.ts'

export const TILE_PIXELS = TILE_SIZE * TILE_SIZE
/** Bytes of one RGBA tile. */
export const TILE_BYTES = TILE_PIXELS * 4
/** Bytes of one mask tile. */
export const MASK_TILE_BYTES = TILE_PIXELS
/** Tile coordinates must lie in [-TILE_COORD_LIMIT, TILE_COORD_LIMIT) (about 8.4 million px each way). */
export const TILE_COORD_LIMIT = 0x8000
/** Pyramid levels whose region versions every grid tracks (level n tiles cover 256 * 2^n px). */
export const MAX_TRACKED_LEVEL = 12

const KEY_STRIDE = 0x10000
const LITTLE_ENDIAN = new Uint8Array(new Uint32Array([0x01020304]).buffer)[0] === 0x04

let versionCounter = 0

/** Allocates a new global surface version (monotonic across every surface and mask). */
export function nextSurfaceVersion(): number {
  versionCounter += 1
  return versionCounter
}

// ---------------------------------------------------------------------------------------------
// Tile addressing
// ---------------------------------------------------------------------------------------------

/** Map key of a tile: (ty + 0x8000) * 0x10000 + (tx + 0x8000). */
export function tileKey(tx: number, ty: number): number {
  return (ty + TILE_COORD_LIMIT) * KEY_STRIDE + (tx + TILE_COORD_LIMIT)
}

/** Inverse of tileKey. */
export function tileCoords(key: number): { readonly tx: number; readonly ty: number } {
  return { tx: (key % KEY_STRIDE) - TILE_COORD_LIMIT, ty: Math.floor(key / KEY_STRIDE) - TILE_COORD_LIMIT }
}

/** Pixel rectangle of tile (tx, ty). */
export function tileRect(tx: number, ty: number): IntRect {
  return { x: tx << TILE_SHIFT, y: ty << TILE_SHIFT, width: TILE_SIZE, height: TILE_SIZE }
}

/** Tile index containing pixel coordinate `value` (floor division, correct for negatives). */
export function tileIndex(value: number): number {
  return Math.floor(value / TILE_SIZE)
}

export interface TileRange {
  /** Inclusive start, exclusive end, in tile units. */
  readonly tx0: number
  readonly ty0: number
  readonly tx1: number
  readonly ty1: number
}

/** Integer form of a rectangle (rounded like cropBuffer), width/height clamped to >= 0. */
export function toIntRect(rect: IntRect): IntRect {
  const x = Math.round(Number(rect.x) || 0)
  const y = Math.round(Number(rect.y) || 0)
  const width = Math.max(0, Math.round(Number(rect.width) || 0))
  const height = Math.max(0, Math.round(Number(rect.height) || 0))
  return { x, y, width, height }
}

/** Tiles touched by `rect`, or null for an empty rectangle. */
export function tileRange(rect: IntRect): TileRange | null {
  const r = toIntRect(rect)
  if (r.width <= 0 || r.height <= 0) return null
  return {
    tx0: tileIndex(r.x),
    ty0: tileIndex(r.y),
    tx1: tileIndex(r.x + r.width - 1) + 1,
    ty1: tileIndex(r.y + r.height - 1) + 1,
  }
}

/** Every tile (tx, ty) that `rect` touches, row by row. */
export function tilesInRect(rect: IntRect): { readonly tx: number; readonly ty: number }[] {
  const range = tileRange(rect)
  const out: { tx: number; ty: number }[] = []
  if (!range) return out
  for (let ty = range.ty0; ty < range.ty1; ty += 1) {
    for (let tx = range.tx0; tx < range.tx1; tx += 1) out.push({ tx, ty })
  }
  return out
}

/** The smallest tile-aligned rectangle containing `rect` (empty input gives an empty rect at its origin). */
export function tileAlignedRect(rect: IntRect): IntRect {
  const range = tileRange(rect)
  if (!range) {
    const r = toIntRect(rect)
    return { x: r.x, y: r.y, width: 0, height: 0 }
  }
  return {
    x: range.tx0 * TILE_SIZE,
    y: range.ty0 * TILE_SIZE,
    width: (range.tx1 - range.tx0) * TILE_SIZE,
    height: (range.ty1 - range.ty0) * TILE_SIZE,
  }
}

function checkTileCoords(tx: number, ty: number): void {
  if (!Number.isInteger(tx) || !Number.isInteger(ty) || tx < -TILE_COORD_LIMIT || ty < -TILE_COORD_LIMIT
    || tx >= TILE_COORD_LIMIT || ty >= TILE_COORD_LIMIT) {
    throw new RangeError(`Tile (${tx}, ${ty}) is too far from the canvas.`)
  }
}

// ---------------------------------------------------------------------------------------------
// Tile content tests
// ---------------------------------------------------------------------------------------------

type TileArray = Uint8ClampedArray | Uint8Array

interface Raster {
  readonly width: number
  readonly height: number
  readonly data: TileArray
}

/** True when an RGBA tile is fully transparent, or a mask tile holds only `fill`. */
export function isEmptyTile(data: TileArray, channels: 1 | 4, fill = 0): boolean {
  if (LITTLE_ENDIAN && data.byteOffset % 4 === 0 && data.length % 4 === 0) {
    const words = new Uint32Array(data.buffer, data.byteOffset, data.length >> 2)
    if (channels === 4) {
      // Alpha is the most significant byte of each little-endian word.
      for (let i = 0; i < words.length; i += 1) if (words[i] >= 0x01000000) return false
      return true
    }
    const word = (fill * 0x01010101) >>> 0
    for (let i = 0; i < words.length; i += 1) if (words[i] !== word) return false
    return true
  }
  if (channels === 4) {
    for (let i = 3; i < data.length; i += 4) if (data[i] !== 0) return false
    return true
  }
  for (let i = 0; i < data.length; i += 1) if (data[i] !== fill) return false
  return true
}

/** True when the source region holds anything other than "empty" (alpha > 0, or a value != fill). */
function regionHasContent(src: Raster, channels: 1 | 4, fill: number, sx: number, sy: number, width: number, height: number): boolean {
  const data = src.data
  const stride = src.width * channels
  if (channels === 4) {
    for (let y = sy; y < sy + height; y += 1) {
      const end = y * stride + (sx + width) * 4
      for (let i = y * stride + sx * 4 + 3; i < end; i += 4) if (data[i] !== 0) return true
    }
    return false
  }
  for (let y = sy; y < sy + height; y += 1) {
    const end = y * stride + sx + width
    for (let i = y * stride + sx; i < end; i += 1) if (data[i] !== fill) return true
  }
  return false
}

interface LocalBounds {
  readonly x0: number
  readonly y0: number
  readonly x1: number
  readonly y1: number
}

/** Tight bounds (inclusive) of non-empty pixels inside one tile, or null. */
function scanTileBounds(data: TileArray, channels: 1 | 4, fill: number): LocalBounds | null {
  let x0 = TILE_SIZE
  let y0 = -1
  let x1 = -1
  let y1 = -1
  for (let y = 0; y < TILE_SIZE; y += 1) {
    const row = y * TILE_SIZE
    let first = -1
    let last = -1
    if (channels === 4) {
      for (let x = 0; x < TILE_SIZE; x += 1) {
        if (data[(row + x) * 4 + 3] !== 0) { first = x; break }
      }
      if (first < 0) continue
      for (let x = TILE_SIZE - 1; x >= first; x -= 1) {
        if (data[(row + x) * 4 + 3] !== 0) { last = x; break }
      }
    } else {
      for (let x = 0; x < TILE_SIZE; x += 1) {
        if (data[row + x] !== fill) { first = x; break }
      }
      if (first < 0) continue
      for (let x = TILE_SIZE - 1; x >= first; x -= 1) {
        if (data[row + x] !== fill) { last = x; break }
      }
    }
    if (y0 < 0) y0 = y
    y1 = y
    if (first < x0) x0 = first
    if (last > x1) x1 = last
  }
  return y0 < 0 ? null : { x0, y0, x1, y1 }
}

// ---------------------------------------------------------------------------------------------
// The grid (one implementation serves surfaces and masks)
// ---------------------------------------------------------------------------------------------

interface TileBoundsEntry {
  readonly version: number
  readonly bounds: LocalBounds | null
}

class TileGrid {
  readonly channels: 1 | 4
  readonly defaultValue: 0 | 255
  private readonly tiles = new Map<number, TileArray>()
  private readonly versions = new Map<number, number>()
  /** regions[level - 1]: newest tile version inside each level-`level` footprint. */
  private readonly regions: Map<number, number>[] = []
  private current = 0
  private boundsAt = -1
  private boundsValue: IntRect | null = null
  private readonly tileBoundsCache = new Map<number, TileBoundsEntry>()

  constructor(channels: 1 | 4, defaultValue: 0 | 255) {
    this.channels = channels
    this.defaultValue = channels === 4 ? 0 : defaultValue
    for (let level = 1; level <= MAX_TRACKED_LEVEL; level += 1) this.regions.push(new Map())
  }

  get version(): number {
    return this.current
  }

  get byteSize(): number {
    return this.tiles.size * TILE_PIXELS * this.channels
  }

  get tileCount(): number {
    return this.tiles.size
  }

  // ----- private helpers ---------------------------------------------------------------------

  private allocate(): TileArray {
    if (this.channels === 4) return new Uint8ClampedArray(TILE_BYTES)
    const tile = new Uint8Array(MASK_TILE_BYTES)
    if (this.defaultValue !== 0) tile.fill(this.defaultValue)
    return tile
  }

  private allocateRaster(width: number, height: number): Raster {
    if (this.channels === 4) return { width, height, data: new Uint8ClampedArray(width * height * 4) }
    const data = new Uint8Array(width * height)
    if (this.defaultValue !== 0) data.fill(this.defaultValue)
    return { width, height, data }
  }

  /** Buffers passed to read()/write() hold bytes (Uint8ClampedArray or Uint8Array) of the right length. */
  private checkRaster(raster: Raster, label: string): void {
    if (!raster || !raster.data || !Number.isInteger(raster.width) || !Number.isInteger(raster.height)
      || raster.width < 0 || raster.height < 0 || raster.data.length !== raster.width * raster.height * this.channels) {
      throw new RangeError(`The ${label} does not match its size.`)
    }
    if (!(raster.data instanceof Uint8ClampedArray) && !(raster.data instanceof Uint8Array)) {
      throw new TypeError(`The ${label} must hold bytes (Uint8ClampedArray or Uint8Array).`)
    }
  }

  /** Records that the given tiles changed: one new global version for all of them. */
  private markChanged(keys: Iterable<number>): void {
    let version = 0
    for (const key of keys) {
      if (!version) version = nextSurfaceVersion()
      this.versions.set(key, version)
      const { tx, ty } = tileCoords(key)
      for (let level = 1; level <= MAX_TRACKED_LEVEL; level += 1) {
        this.regions[level - 1].set(tileKey(tx >> level, ty >> level), version)
      }
    }
    if (version) this.current = version
  }

  // ----- TiledSurface / TiledMask ------------------------------------------------------------

  contentBounds(): IntRect | null {
    if (this.boundsAt === this.current) return this.boundsValue
    let minX = Infinity
    let minY = Infinity
    let maxX = -Infinity
    let maxY = -Infinity
    for (const [key, data] of this.tiles) {
      const version = this.versions.get(key) ?? 0
      let entry = this.tileBoundsCache.get(key)
      if (!entry || entry.version !== version) {
        entry = { version, bounds: scanTileBounds(data, this.channels, this.defaultValue) }
        this.tileBoundsCache.set(key, entry)
      }
      const local = entry.bounds
      if (!local) continue
      const { tx, ty } = tileCoords(key)
      const ox = tx * TILE_SIZE
      const oy = ty * TILE_SIZE
      if (ox + local.x0 < minX) minX = ox + local.x0
      if (oy + local.y0 < minY) minY = oy + local.y0
      if (ox + local.x1 > maxX) maxX = ox + local.x1
      if (oy + local.y1 > maxY) maxY = oy + local.y1
    }
    // Forget bounds of tiles that no longer exist.
    if (this.tileBoundsCache.size > this.tiles.size) {
      for (const key of [...this.tileBoundsCache.keys()]) if (!this.tiles.has(key)) this.tileBoundsCache.delete(key)
    }
    this.boundsValue = maxX < minX ? null : { x: minX, y: minY, width: maxX - minX + 1, height: maxY - minY + 1 }
    this.boundsAt = this.current
    return this.boundsValue
  }

  read(rect: IntRect, target?: Raster): Raster {
    const r = toIntRect(rect)
    let out: Raster
    if (target) {
      this.checkRaster(target, 'target buffer')
      if (target.width !== r.width || target.height !== r.height) {
        throw new RangeError(`The target buffer is ${target.width} x ${target.height}, not ${r.width} x ${r.height}.`)
      }
      out = target
      out.data.fill(this.defaultValue)
    } else {
      out = this.allocateRaster(r.width, r.height)
    }
    const range = tileRange(r)
    if (!range || this.tiles.size === 0) return out
    const channels = this.channels
    const outStride = r.width * channels
    for (let ty = range.ty0; ty < range.ty1; ty += 1) {
      const tileTop = ty * TILE_SIZE
      const y0 = Math.max(r.y, tileTop)
      const y1 = Math.min(r.y + r.height, tileTop + TILE_SIZE)
      for (let tx = range.tx0; tx < range.tx1; tx += 1) {
        const tile = this.tiles.get(tileKey(tx, ty))
        if (!tile) continue
        const tileLeft = tx * TILE_SIZE
        const x0 = Math.max(r.x, tileLeft)
        const x1 = Math.min(r.x + r.width, tileLeft + TILE_SIZE)
        const span = (x1 - x0) * channels
        for (let y = y0; y < y1; y += 1) {
          const from = ((y - tileTop) * TILE_SIZE + (x0 - tileLeft)) * channels
          out.data.set(tile.subarray(from, from + span), (y - r.y) * outStride + (x0 - r.x) * channels)
        }
      }
    }
    return out
  }

  write(x: number, y: number, src: Raster, srcRect?: IntRect): IntRect {
    this.checkRaster(src, 'source buffer')
    const left = Math.round(Number(x) || 0)
    const top = Math.round(Number(y) || 0)
    const region = srcRect ? toIntRect(srcRect) : { x: 0, y: 0, width: src.width, height: src.height }
    // Clip the source region to the source buffer; the destination shifts with the clip.
    const sx0 = Math.max(0, region.x)
    const sy0 = Math.max(0, region.y)
    const sx1 = Math.min(src.width, region.x + region.width)
    const sy1 = Math.min(src.height, region.y + region.height)
    const destX = left + (sx0 - region.x)
    const destY = top + (sy0 - region.y)
    if (sx1 <= sx0 || sy1 <= sy0) return { x: left, y: top, width: 0, height: 0 }
    const dest: IntRect = { x: destX, y: destY, width: sx1 - sx0, height: sy1 - sy0 }
    const range = tileRange(dest) as TileRange
    checkTileCoords(range.tx0, range.ty0)
    checkTileCoords(range.tx1 - 1, range.ty1 - 1)
    const channels = this.channels
    const fill = this.defaultValue
    const srcStride = src.width * channels
    const changed: number[] = []
    for (let ty = range.ty0; ty < range.ty1; ty += 1) {
      const tileTop = ty * TILE_SIZE
      const y0 = Math.max(dest.y, tileTop)
      const y1 = Math.min(dest.y + dest.height, tileTop + TILE_SIZE)
      for (let tx = range.tx0; tx < range.tx1; tx += 1) {
        const tileLeft = tx * TILE_SIZE
        const x0 = Math.max(dest.x, tileLeft)
        const x1 = Math.min(dest.x + dest.width, tileLeft + TILE_SIZE)
        const key = tileKey(tx, ty)
        const fromX = sx0 + (x0 - dest.x)
        const fromY = sy0 + (y0 - dest.y)
        const content = regionHasContent(src, channels, fill, fromX, fromY, x1 - x0, y1 - y0)
        let tile = this.tiles.get(key)
        if (!tile) {
          // Writing "empty" into an empty tile changes nothing.
          if (!content) continue
          tile = this.allocate()
          this.tiles.set(key, tile)
        }
        const span = (x1 - x0) * channels
        for (let row = y0; row < y1; row += 1) {
          const from = (fromY + (row - y0)) * srcStride + fromX * channels
          tile.set(src.data.subarray(from, from + span), ((row - tileTop) * TILE_SIZE + (x0 - tileLeft)) * channels)
        }
        if (!content && isEmptyTile(tile, channels, fill)) this.tiles.delete(key)
        changed.push(key)
      }
    }
    this.markChanged(changed)
    return tileAlignedRect(dest)
  }

  tile(tx: number, ty: number): TileArray | undefined {
    return this.tiles.get(tileKey(tx, ty))
  }

  setTile(tx: number, ty: number, data: TileArray | undefined): void {
    checkTileCoords(tx, ty)
    const key = tileKey(tx, ty)
    if (data === undefined || data === null) {
      if (!this.tiles.has(key)) return
      this.tiles.delete(key)
      this.markChanged([key])
      return
    }
    const bytes = TILE_PIXELS * this.channels
    const expected = this.channels === 4 ? Uint8ClampedArray : Uint8Array
    if (!(data instanceof expected) || data.length !== bytes) {
      throw new RangeError(`A ${this.channels === 4 ? 'surface' : 'mask'} tile must be a ${this.channels === 4 ? 'Uint8ClampedArray' : 'Uint8Array'} of ${bytes} values.`)
    }
    this.tiles.set(key, data)
    this.markChanged([key])
  }

  tileVersion(tx: number, ty: number): number {
    return this.versions.get(tileKey(tx, ty)) ?? 0
  }

  forEachTile(visit: (tx: number, ty: number, data: TileArray) => void): void {
    // Iterate a snapshot so a visitor may call setTile safely.
    for (const [key, data] of [...this.tiles]) {
      const { tx, ty } = tileCoords(key)
      visit(tx, ty, data)
    }
  }

  clone(): TileGrid {
    const copy = new TileGrid(this.channels, this.defaultValue)
    for (const [key, data] of this.tiles) copy.tiles.set(key, data.slice())
    copy.markChanged(copy.tiles.keys())
    return copy
  }

  // ----- extensions (not part of the shared interfaces) -------------------------------------

  /** Newest tile version inside the level-`level` footprint of (tx, ty) (level 0 = tileVersion). 0 = never touched. */
  regionVersion(level: number, tx: number, ty: number): number {
    if (level <= 0) return this.tileVersion(tx, ty)
    const map = this.regions[level - 1]
    return map ? map.get(tileKey(tx, ty)) ?? 0 : -1
  }

  /** Frees allocated tiles that are empty (e.g. after setTile with a cleared tile). Returns the count freed. */
  compact(): number {
    const freed: number[] = []
    for (const [key, data] of this.tiles) {
      if (isEmptyTile(data, this.channels, this.defaultValue)) freed.push(key)
    }
    for (const key of freed) this.tiles.delete(key)
    this.markChanged(freed)
    return freed.length
  }
}

// ---------------------------------------------------------------------------------------------
// Factories and helpers
// ---------------------------------------------------------------------------------------------

/** A new, empty (fully transparent) RGBA surface. */
export function createSurface(): TiledSurface {
  return new TileGrid(4, 0) as unknown as TiledSurface
}

/** A new, empty mask: every pixel reads `defaultValue` (255 reveals all, 0 hides all). */
export function createMaskSurface(defaultValue: 0 | 255): TiledMask {
  if (defaultValue !== 0 && defaultValue !== 255) throw new RangeError('A mask default value is 0 or 255.')
  return new TileGrid(1, defaultValue) as unknown as TiledMask
}

/** A surface holding `buffer` with its top-left at layer-local (x, y). Transparent tiles are not allocated. */
export function surfaceFromBuffer(buffer: PixelBuffer, x = 0, y = 0): TiledSurface {
  const surface = createSurface()
  surface.write(x, y, buffer)
  return surface
}

/** A mask holding `buffer` at (x, y); tiles that equal `defaultValue` everywhere are not allocated. */
export function maskFromBuffer(buffer: MaskBuffer, defaultValue: 0 | 255, x = 0, y = 0): TiledMask {
  const mask = createMaskSurface(defaultValue)
  mask.write(x, y, buffer)
  return mask
}

export function isTiledSurface(value: unknown): value is TiledSurface {
  return Boolean(value) && typeof value === 'object' && (value as TiledSurface).channels === 4
    && typeof (value as TiledSurface).read === 'function'
}

export function isTiledMask(value: unknown): value is TiledMask {
  return Boolean(value) && typeof value === 'object' && (value as TiledMask).channels === 1
    && typeof (value as TiledMask).read === 'function'
}

interface RegionVersioned {
  regionVersion(level: number, tx: number, ty: number): number
}

function hasRegionVersions(grid: unknown): grid is RegionVersioned {
  return grid instanceof TileGrid
}

/**
 * Newest version of any level-0 tile inside the level-`level` footprint of tile (tx, ty); 0 when no tile
 * there was ever written. O(1) for grids made by this module; other implementations are walked.
 */
export function regionVersion(grid: TiledSurface | TiledMask, level: number, tx: number, ty: number): number {
  if (level <= 0) return grid.tileVersion(tx, ty)
  if (hasRegionVersions(grid) && level <= MAX_TRACKED_LEVEL) return grid.regionVersion(level, tx, ty)
  let newest = 0
  for (let dy = 0; dy < 2; dy += 1) {
    for (let dx = 0; dx < 2; dx += 1) {
      const version = regionVersion(grid, level - 1, tx * 2 + dx, ty * 2 + dy)
      if (version > newest) newest = version
    }
  }
  return newest
}

/** Frees empty tiles of a grid made by this module (no-op for other implementations). Returns the count. */
export function compactGrid(grid: TiledSurface | TiledMask): number {
  return grid instanceof TileGrid ? grid.compact() : 0
}
