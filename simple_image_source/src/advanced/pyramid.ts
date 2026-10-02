// src/advanced/pyramid.ts (WP3)
// Mip pyramid of tiled surfaces and masks for display at reduced zoom (design 5.5).
//   - Level k has scale 2^-k. Tile (k, tx, ty) is a 2x2 box of level k - 1 tiles (2tx..2tx+1, 2ty..2ty+1)
//     with alpha-weighted averaging (premultiply, average, unpremultiply) so transparent edges never darken.
//     Mask levels use a plain average; unallocated mask areas read as the mask's defaultValue.
//   - Levels are built lazily per tile and validated against the surface's region versions (the newest
//     level-0 tile version inside the footprint), so an edit rebuilds exactly the proxy tiles above it.
//   - Built tiles live in one global LRU (bytes reported to memory.ts); evicted tiles are rebuilt on demand.
//   - A layer at offset o sits at floor(o / 2^k) on level k (at most one proxy pixel of display error;
//     level 0 is exact). Proxy levels are for display only: flatten and export always use level 0.
// Pure and DOM-free.
import type { IntRect, MaskBuffer, PixelBuffer, Size } from '../imaging/types.ts'
import type { TiledMask, TiledSurface } from './types.ts'
import { TILE_SIZE } from './types.ts'
import { MAX_TRACKED_LEVEL, TILE_BYTES, MASK_TILE_BYTES, regionVersion, tileKey, tileRange, toIntRect } from './tiles.ts'
import { registerCache } from './memory.ts'

/** Highest pyramid level served (1 / 1024 scale). */
export const MAX_LEVEL = Math.min(10, MAX_TRACKED_LEVEL)
/** Default byte budget of the global proxy-tile cache. */
export const DEFAULT_PYRAMID_BUDGET = 256 * 1024 * 1024

const HALF = TILE_SIZE >> 1

type Grid = TiledSurface | TiledMask
type TileArray = Uint8ClampedArray | Uint8Array

// ---------------------------------------------------------------------------------------------
// Level geometry
// ---------------------------------------------------------------------------------------------

/** Smallest level at which a width x height document fits in one tile (capped at MAX_LEVEL). */
export function maxPyramidLevel(width: number, height: number): number {
  let size = Math.max(1, Math.ceil(Math.max(Number(width) || 0, Number(height) || 0)))
  let level = 0
  while (size > TILE_SIZE && level < MAX_LEVEL) {
    size = Math.ceil(size / 2)
    level += 1
  }
  return level
}

/** Size of a width x height document at `level` (ceil(size / 2^level)). */
export function levelSize(width: number, height: number, level: number): Size {
  const scale = 2 ** Math.max(0, level | 0)
  return { width: Math.ceil(width / scale), height: Math.ceil(height / scale) }
}

/** Position of a level-0 offset at `level`: floor(offset / 2^level). */
export function levelOffset(offset: number, level: number): number {
  return Math.floor(offset / 2 ** Math.max(0, level | 0))
}

// ---------------------------------------------------------------------------------------------
// 2x2 reductions
// ---------------------------------------------------------------------------------------------

/** Alpha-weighted 2x2 average of a 256x256 RGBA tile into one 128x128 quadrant of `out`. */
function reduceRgbaQuadrant(child: TileArray, out: TileArray, qx: number, qy: number): void {
  const rowBytes = TILE_SIZE * 4
  for (let y = 0; y < HALF; y += 1) {
    let s0 = 2 * y * rowBytes
    let s1 = s0 + rowBytes
    let d = ((qy * HALF + y) * TILE_SIZE + qx * HALF) * 4
    for (let x = 0; x < HALF; x += 1, s0 += 8, s1 += 8, d += 4) {
      const a0 = child[s0 + 3]
      const a1 = child[s0 + 7]
      const a2 = child[s1 + 3]
      const a3 = child[s1 + 7]
      const sum = a0 + a1 + a2 + a3
      if (sum === 0) continue
      if (sum === 1020) {
        out[d] = (child[s0] + child[s0 + 4] + child[s1] + child[s1 + 4] + 2) >> 2
        out[d + 1] = (child[s0 + 1] + child[s0 + 5] + child[s1 + 1] + child[s1 + 5] + 2) >> 2
        out[d + 2] = (child[s0 + 2] + child[s0 + 6] + child[s1 + 2] + child[s1 + 6] + 2) >> 2
        out[d + 3] = 255
        continue
      }
      const half = sum >> 1
      out[d] = Math.floor((child[s0] * a0 + child[s0 + 4] * a1 + child[s1] * a2 + child[s1 + 4] * a3 + half) / sum)
      out[d + 1] = Math.floor((child[s0 + 1] * a0 + child[s0 + 5] * a1 + child[s1 + 1] * a2 + child[s1 + 5] * a3 + half) / sum)
      out[d + 2] = Math.floor((child[s0 + 2] * a0 + child[s0 + 6] * a1 + child[s1 + 2] * a2 + child[s1 + 6] * a3 + half) / sum)
      out[d + 3] = (sum + 2) >> 2
    }
  }
}

/** Plain 2x2 average of a 256x256 mask tile into one quadrant of `out`. */
function reduceMaskQuadrant(child: TileArray, out: TileArray, qx: number, qy: number): void {
  for (let y = 0; y < HALF; y += 1) {
    let s0 = 2 * y * TILE_SIZE
    let s1 = s0 + TILE_SIZE
    let d = (qy * HALF + y) * TILE_SIZE + qx * HALF
    for (let x = 0; x < HALF; x += 1, s0 += 2, s1 += 2, d += 1) {
      out[d] = (child[s0] + child[s0 + 1] + child[s1] + child[s1 + 1] + 2) >> 2
    }
  }
}

/**
 * Alpha-weighted 2x2 downsample of a whole buffer (odd edges average the pixels that exist and count
 * the missing ones as transparent). The reference the pyramid is tested against.
 */
export function downsampleBuffer2x2(src: PixelBuffer): PixelBuffer {
  const width = Math.ceil(src.width / 2)
  const height = Math.ceil(src.height / 2)
  const data = new Uint8ClampedArray(width * height * 4)
  const s = src.data
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      let sum = 0
      let r = 0
      let g = 0
      let b = 0
      for (let dy = 0; dy < 2; dy += 1) {
        const sy = 2 * y + dy
        if (sy >= src.height) continue
        for (let dx = 0; dx < 2; dx += 1) {
          const sx = 2 * x + dx
          if (sx >= src.width) continue
          const i = (sy * src.width + sx) * 4
          const a = s[i + 3]
          sum += a
          r += s[i] * a
          g += s[i + 1] * a
          b += s[i + 2] * a
        }
      }
      if (sum === 0) continue
      const o = (y * width + x) * 4
      const half = sum >> 1
      data[o] = Math.floor((r + half) / sum)
      data[o + 1] = Math.floor((g + half) / sum)
      data[o + 2] = Math.floor((b + half) / sum)
      data[o + 3] = (sum + 2) >> 2
    }
  }
  return { width, height, data }
}

// ---------------------------------------------------------------------------------------------
// Global proxy-tile cache (LRU)
// ---------------------------------------------------------------------------------------------

interface LevelEntry {
  readonly stamp: number
  /** null = empty tile (transparent, or all defaultValue for masks). */
  readonly data: TileArray | null
  readonly bytes: number
}

const cache = new Map<string, LevelEntry>()
const gridIds = new WeakMap<object, number>()
let nextGridId = 1
let cachedBytes = 0
let budgetBytes = DEFAULT_PYRAMID_BUDGET

registerCache({ name: 'pyramid', bytes: () => cachedBytes, trim: () => clearPyramidCache() })

function gridId(grid: Grid): number {
  let id = gridIds.get(grid)
  if (id === undefined) {
    id = nextGridId
    nextGridId += 1
    gridIds.set(grid, id)
  }
  return id
}

function remember(key: string, entry: LevelEntry): void {
  const previous = cache.get(key)
  if (previous) {
    cachedBytes -= previous.bytes
    cache.delete(key)
  }
  cache.set(key, entry)
  cachedBytes += entry.bytes
  if (cachedBytes > budgetBytes) evict(budgetBytes)
}

function evict(limit: number): void {
  for (const [key, entry] of cache) {
    if (cachedBytes <= limit) break
    cache.delete(key)
    cachedBytes -= entry.bytes
  }
}

/** Bytes held by proxy tiles of every surface and mask. */
export function pyramidCacheBytes(): number {
  return cachedBytes
}

/** Drops every proxy tile (they rebuild on demand). */
export function clearPyramidCache(): void {
  cache.clear()
  cachedBytes = 0
}

/** Sets the proxy-tile byte budget (evicting least recently used tiles above it). */
export function setPyramidCacheBudget(bytes: number): void {
  budgetBytes = Math.max(0, Number(bytes) || 0)
  evict(budgetBytes)
}

// ---------------------------------------------------------------------------------------------
// Level tiles
// ---------------------------------------------------------------------------------------------

function levelTile(grid: Grid, level: number, tx: number, ty: number): TileArray | null {
  if (level <= 0) return (grid.tile(tx, ty) as TileArray | undefined) ?? null
  const stamp = regionVersion(grid, level, tx, ty)
  // A footprint where no tile was ever written is empty without looking further.
  if (stamp === 0) return null
  const key = `${gridId(grid)}:${level}:${tx}:${ty}`
  const hit = cache.get(key)
  if (hit && hit.stamp === stamp) {
    // Refresh the LRU position.
    cache.delete(key)
    cache.set(key, hit)
    return hit.data
  }
  const mask = grid.channels === 1
  let out: TileArray | null = null
  for (let qy = 0; qy < 2; qy += 1) {
    for (let qx = 0; qx < 2; qx += 1) {
      const child = levelTile(grid, level - 1, tx * 2 + qx, ty * 2 + qy)
      if (!child) continue
      if (!out) {
        if (mask) {
          out = new Uint8Array(MASK_TILE_BYTES)
          const fill = (grid as TiledMask).defaultValue
          if (fill !== 0) out.fill(fill)
        } else {
          out = new Uint8ClampedArray(TILE_BYTES)
        }
      }
      if (mask) reduceMaskQuadrant(child, out, qx, qy)
      else reduceRgbaQuadrant(child, out, qx, qy)
    }
  }
  remember(key, { stamp, data: out, bytes: out ? out.length : 0 })
  return out
}

/** RGBA tile (tx, ty) of `surface` at `level`, or null when that tile is fully transparent. */
export function surfaceLevelTile(surface: TiledSurface, level: number, tx: number, ty: number): Uint8ClampedArray | null {
  return levelTile(surface, level, tx, ty) as Uint8ClampedArray | null
}

/** Mask tile (tx, ty) at `level`, or null when every value there is the mask's defaultValue. */
export function maskLevelTile(mask: TiledMask, level: number, tx: number, ty: number): Uint8Array | null {
  return levelTile(mask, level, tx, ty) as Uint8Array | null
}

function readLevel(grid: Grid, level: number, rect: IntRect, target: { width: number; height: number; data: TileArray }, fill: number): void {
  const channels = grid.channels
  const r = toIntRect(rect)
  target.data.fill(fill)
  const range = tileRange(r)
  if (!range) return
  const stride = r.width * channels
  for (let ty = range.ty0; ty < range.ty1; ty += 1) {
    const tileTop = ty * TILE_SIZE
    const y0 = Math.max(r.y, tileTop)
    const y1 = Math.min(r.y + r.height, tileTop + TILE_SIZE)
    for (let tx = range.tx0; tx < range.tx1; tx += 1) {
      const tile = levelTile(grid, level, tx, ty)
      if (!tile) continue
      const tileLeft = tx * TILE_SIZE
      const x0 = Math.max(r.x, tileLeft)
      const x1 = Math.min(r.x + r.width, tileLeft + TILE_SIZE)
      const span = (x1 - x0) * channels
      for (let y = y0; y < y1; y += 1) {
        const from = ((y - tileTop) * TILE_SIZE + (x0 - tileLeft)) * channels
        target.data.set(tile.subarray(from, from + span), (y - r.y) * stride + (x0 - r.x) * channels)
      }
    }
  }
}

/**
 * Copies `rect` (level-`level` pixels in the surface's own coordinates) out of the surface pyramid.
 * Level 0 reads the surface itself. `target`, when given, must be exactly rect-sized.
 */
export function readSurfaceLevel(surface: TiledSurface, level: number, rect: IntRect, target?: PixelBuffer): PixelBuffer {
  if (level <= 0) return surface.read(rect, target)
  const r = toIntRect(rect)
  const out = target ?? { width: r.width, height: r.height, data: new Uint8ClampedArray(r.width * r.height * 4) }
  if (out.width !== r.width || out.height !== r.height || out.data.length !== r.width * r.height * 4) {
    throw new RangeError(`The target buffer must be ${r.width} x ${r.height}.`)
  }
  readLevel(surface, level, r, out, 0)
  return out
}

/** Mask counterpart of readSurfaceLevel (unallocated areas read as the mask's defaultValue). */
export function readMaskLevel(mask: TiledMask, level: number, rect: IntRect, target?: MaskBuffer): MaskBuffer {
  if (level <= 0) return mask.read(rect, target)
  const r = toIntRect(rect)
  const out = target ?? { width: r.width, height: r.height, data: new Uint8Array(r.width * r.height) }
  if (out.width !== r.width || out.height !== r.height || out.data.length !== r.width * r.height) {
    throw new RangeError(`The target mask must be ${r.width} x ${r.height}.`)
  }
  readLevel(mask, level, r, out, mask.defaultValue)
  return out
}

/** Unique tile key helper re-exported for caches that index level tiles. */
export function levelTileKey(level: number, tx: number, ty: number): string {
  return `${level}:${tileKey(tx, ty)}`
}
