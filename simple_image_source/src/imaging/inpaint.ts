// src/imaging/inpaint.ts (WP4)
// Spot Healing Brush and inpainting. Pure and DOM-free (runs in the imaging worker via the 'heal' op).
//
// healRegion (design 5.12):
//   1. The stroke (hole > 0) is dilated by 2 px so the boundary lies in clean pixels.
//   2. Thin holes (maximum inscribed radius <= 3 px: scratches, wires, small dots) use Telea inpainting.
//   3. Otherwise the source offset with the lowest SSD over a border band (width max(3, 0.15 * D)) is
//      searched among 16 angles on rings of 1, 1.5, 2 and 3 hole diameters D, refined by a +-4 px local
//      search; sources whose footprint overlaps the hole or leaves the image are rejected. D is the
//      hole's local width (twice the inscribed radius); elongated strokes also try rings of their length.
//      When no ring fits (small image, hole near an edge) every in-image offset is scanned on a grid;
//      only when nothing fits does healing fall back to Telea.
//   4. Seamless clone: f = g + c where g is the shifted source and c solves Laplace's equation inside the
//      hole with c = I - g on its boundary (equivalent to Poisson with source gradients as guidance).
//      A linear ramp or a flat colour is reproduced exactly.
//   5. The outermost ring of the hole is feathered 50% with the original pixels.
// The Laplace solver is geometric multigrid: V-cycles with Gauss-Seidel smoothing (2 + 2 sweeps), residual
// restriction to 2 x 2 blocks that lie fully inside the hole, bilinear prolongation, and SOR with an
// over-relaxation factor estimated from the observed Gauss-Seidel rate on the coarsest level. It stops when
// the largest Jacobi update falls below 0.002 levels (the design's single-level SOR with omega 1.85 and a
// 0.25 change limit stopped up to 11 levels short of the solution on 100-500 px holes).
// inpaintTelea: fast marching inpainting (A. Telea, 2004) with gradient-extrapolated weighted averages.
import type { HealOptions, IntRect, MaskBuffer, OpOptions, PixelBuffer } from './types.ts'
import { distanceTransform } from './distance.ts'

const DEFAULT_RINGS: readonly number[] = [1, 1.5, 2, 3]
const ANGLES = 16
const REFINE = 4
const DILATE = 2
/** Iteration cap of the single-level / coarsest-level SOR (HealOptions.maxIterations default). */
const MAX_ITERATIONS = 600
/** V-cycle cap; typical holes converge in 4-8 cycles. */
const MAX_CYCLES = 40
/** Stop when every Jacobi update |r / n| is below this many levels. */
const RESIDUAL_TOLERANCE = 0.002
/** Coarsest-level SOR stops when the largest change is below this. */
const COARSE_TOLERANCE = 0.0005
/** Stop coarsening at this span (pixels) or unknown count. */
const COARSEST_SPAN = 16
const COARSEST_COUNT = 64
const THIN_RADIUS = 3
const TELEA_RADIUS = 5

const KNOWN = 0
const BAND = 1
const INSIDE = 2

function abortError(): Error {
  const error = new Error('The operation was cancelled.')
  error.name = 'AbortError'
  return error
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted) throw abortError()
}

function checkInputs(src: PixelBuffer, hole: MaskBuffer): void {
  if (!src || !src.data || src.data.length !== src.width * src.height * 4) {
    throw new RangeError('The pixel buffer does not match its size.')
  }
  if (!hole || !hole.data || hole.width !== src.width || hole.height !== src.height || hole.data.length !== src.width * src.height) {
    throw new RangeError('The healing mask must have the same size as the pixels.')
  }
}

function copyBuffer(src: PixelBuffer): PixelBuffer {
  return { width: src.width, height: src.height, data: new Uint8ClampedArray(src.data) }
}

/** Pseudo-random fraction in [0, 1) from an integer seed (angle phase for "try again"). */
function seedPhase(seed: number | undefined): number {
  if (!Number.isFinite(seed) || !seed) return 0
  let h = Math.imul((seed as number) | 0, 0x9e3779b9)
  h = Math.imul(h ^ (h >>> 16), 0x85ebca6b)
  h ^= h >>> 13
  return (h >>> 0) / 4294967296
}

/**
 * Pixels of context the spot-healing tool should crop around a stroke with the given bounds, so every
 * candidate source ring, the band and the refinement fit (clamp the crop to the layer afterwards).
 */
export function healMargin(holeBounds: IntRect, options: HealOptions = {}): number {
  const diameter = Math.max(1, holeBounds.width, holeBounds.height) + 2 * DILATE
  const rings = (options.rings ?? DEFAULT_RINGS).filter((ring) => Number.isFinite(ring) && ring > 0)
  const farthest = rings.length ? Math.max(...rings) : 3
  const band = Math.max(3, Math.round(0.15 * diameter))
  return Math.ceil(farthest * diameter + band + REFINE + DILATE + 2)
}

// ---------------------------------------------------------------------------------------------
// Laplace (membrane) solver: geometric multigrid on a masked grid
// ---------------------------------------------------------------------------------------------

interface Membrane {
  readonly width: number
  readonly height: number
  /** 1 = unknown (solved for). */
  readonly unknown: Uint8Array
  /** 1 = fixed boundary value available (Dirichlet); other known pixels are skipped (Neumann). */
  readonly defined: Uint8Array
  /** 4 values per pixel: fixed values where defined; receives the solution where unknown. */
  readonly values: Float32Array
}

interface SolveControl {
  /** Cap of the coarsest-level SOR iterations. */
  readonly maxIterations: number
  readonly signal?: AbortSignal
  readonly progress?: (fraction: number) => void
}

/** One grid level: the unknowns in raster order with their 5-point stencil. */
interface Level {
  readonly count: number
  /** Unknown id of each neighbour (left, right, up, down) or -1. */
  readonly links: Int32Array
  /** Neighbours that are unknown or fixed (Dirichlet); the stencil's diagonal. */
  readonly diag: Float64Array
  /** Largest bounding-box side of the unknowns, px. */
  readonly span: number
}

/** Coarse-to-fine transfer: restriction parent and bilinear prolongation weights per fine unknown. */
interface Transfer {
  readonly parent: Int32Array
  readonly prolongIds: Int32Array
  readonly prolongWeights: Float64Array
}

interface Grid {
  readonly width: number
  readonly height: number
  readonly unknown: Uint8Array
  readonly fixed: Uint8Array
  readonly ids: Int32Array
  readonly pixels: Int32Array
  readonly level: Level
}

function buildGrid(width: number, height: number, unknown: Uint8Array, fixed: Uint8Array): Grid {
  const ids = new Int32Array(width * height).fill(-1)
  let count = 0
  let minX = width
  let minY = height
  let maxX = -1
  let maxY = -1
  for (let p = 0; p < unknown.length; p += 1) {
    if (!unknown[p]) continue
    ids[p] = count
    count += 1
    const x = p % width
    const y = (p / width) | 0
    if (x < minX) minX = x
    if (x > maxX) maxX = x
    if (y < minY) minY = y
    if (y > maxY) maxY = y
  }
  const pixels = new Int32Array(count)
  for (let p = 0; p < unknown.length; p += 1) if (unknown[p]) pixels[ids[p]] = p
  const links = new Int32Array(count * 4).fill(-1)
  const diag = new Float64Array(count)
  for (let k = 0; k < count; k += 1) {
    const p = pixels[k]
    const x = p % width
    const neighbours = [x > 0 ? p - 1 : -1, x + 1 < width ? p + 1 : -1, p - width, p + width]
    let n = 0
    for (let d = 0; d < 4; d += 1) {
      const q = neighbours[d]
      if (q < 0 || q >= unknown.length) continue
      if (unknown[q]) {
        links[k * 4 + d] = ids[q]
        n += 1
      } else if (fixed[q]) {
        n += 1
      }
    }
    diag[k] = n
  }
  const span = count ? Math.max(maxX - minX + 1, maxY - minY + 1) : 0
  return { width, height, unknown, fixed, ids, pixels, level: { count, links, diag, span } }
}

/**
 * Coarse grid of 2 x 2 blocks. A block is unknown only when all its children are (keeps the coarse
 * boundary inside the fine one; with "any child" the boundary drifts outwards a cell per level and the
 * V-cycle diverges); every other block is a zero (Dirichlet) cell of the correction equation.
 */
function coarsenGrid(grid: Grid): { readonly grid: Grid; readonly transfer: Transfer } {
  const { width, height, unknown } = grid
  const cw = Math.ceil(width / 2)
  const ch = Math.ceil(height / 2)
  const cu = new Uint8Array(cw * ch)
  const cf = new Uint8Array(cw * ch)
  for (let cy = 0; cy < ch; cy += 1) {
    for (let cx = 0; cx < cw; cx += 1) {
      let all = true
      let some = false
      for (let dy = 0; dy < 2; dy += 1) {
        const y = cy * 2 + dy
        if (y >= height) continue
        for (let dx = 0; dx < 2; dx += 1) {
          const x = cx * 2 + dx
          if (x >= width) continue
          if (unknown[y * width + x]) some = true
          else all = false
        }
      }
      if (some && all) cu[cy * cw + cx] = 1
      else cf[cy * cw + cx] = 1
    }
  }
  const coarse = buildGrid(cw, ch, cu, cf)
  const count = grid.level.count
  const parent = new Int32Array(count)
  const prolongIds = new Int32Array(count * 4).fill(-1)
  const prolongWeights = new Float64Array(count * 4)
  for (let k = 0; k < count; k += 1) {
    const p = grid.pixels[k]
    const x = p % width
    const y = (p / width) | 0
    parent[k] = coarse.ids[(y >> 1) * cw + (x >> 1)]
    // Bilinear interpolation between the four nearest coarse cell centres; fixed cells carry 0.
    const fx = (x + 0.5) / 2 - 0.5
    const fy = (y + 0.5) / 2 - 0.5
    const x0 = Math.floor(fx)
    const y0 = Math.floor(fy)
    const tx = fx - x0
    const ty = fy - y0
    let total = 0
    for (let j = 0; j < 2; j += 1) {
      for (let i = 0; i < 2; i += 1) {
        const cx = x0 + i
        const cy = y0 + j
        if (cx < 0 || cy < 0 || cx >= cw || cy >= ch) continue
        const w = (i ? tx : 1 - tx) * (j ? ty : 1 - ty)
        if (w <= 0) continue
        total += w
        const id = coarse.ids[cy * cw + cx]
        if (id >= 0) {
          prolongIds[k * 4 + j * 2 + i] = id
          prolongWeights[k * 4 + j * 2 + i] = w
        }
      }
    }
    if (total > 0) for (let s = 0; s < 4; s += 1) prolongWeights[k * 4 + s] /= total
  }
  return { grid: coarse, transfer: { parent, prolongIds, prolongWeights } }
}

/** Gauss-Seidel / SOR sweeps on A x = b (A = diag - links); returns the largest change of the last sweep. */
function relax(level: Level, x: Float64Array, b: Float64Array, sweeps: number, omega: number): number {
  const { count, links, diag } = level
  let largest = 0
  for (let sweep = 0; sweep < sweeps; sweep += 1) {
    largest = 0
    for (let k = 0; k < count; k += 1) {
      const n = diag[k]
      if (n === 0) continue
      const o = k * 4
      let s0 = b[o]
      let s1 = b[o + 1]
      let s2 = b[o + 2]
      let s3 = b[o + 3]
      for (let d = 0; d < 4; d += 1) {
        const j = links[o + d]
        if (j < 0) continue
        const q = j * 4
        s0 += x[q]
        s1 += x[q + 1]
        s2 += x[q + 2]
        s3 += x[q + 3]
      }
      const d0 = omega * (s0 / n - x[o])
      const d1 = omega * (s1 / n - x[o + 1])
      const d2 = omega * (s2 / n - x[o + 2])
      const d3 = omega * (s3 / n - x[o + 3])
      x[o] += d0
      x[o + 1] += d1
      x[o + 2] += d2
      x[o + 3] += d3
      const m = Math.max(Math.abs(d0), Math.abs(d1), Math.abs(d2), Math.abs(d3))
      if (m > largest) largest = m
    }
  }
  return largest
}

/** r = b - A x; returns the largest Jacobi update |r / diag|. */
function residual(level: Level, x: Float64Array, b: Float64Array, r: Float64Array): number {
  const { count, links, diag } = level
  let largest = 0
  for (let k = 0; k < count; k += 1) {
    const o = k * 4
    const n = diag[k]
    for (let c = 0; c < 4; c += 1) {
      let s = b[o + c] - n * x[o + c]
      for (let d = 0; d < 4; d += 1) {
        const j = links[o + d]
        if (j >= 0) s += x[j * 4 + c]
      }
      r[o + c] = s
      if (n > 0) {
        const m = Math.abs(s / n)
        if (m > largest) largest = m
      }
    }
  }
  return largest
}

/** SOR whose factor comes from the observed Gauss-Seidel contraction (adapts to any domain shape). */
function solveCoarsest(level: Level, x: Float64Array, b: Float64Array, maxIterations: number, tolerance: number, signal?: AbortSignal): void {
  let omega = 1
  let previous = Infinity
  for (let iteration = 0; iteration < maxIterations; iteration += 1) {
    const change = relax(level, x, b, 1, omega)
    if (change < tolerance) return
    if (iteration === 8 && omega === 1 && previous > 0 && previous < Infinity) {
      const rate = change / previous
      if (rate > 0 && rate < 1) omega = 2 / (1 + Math.sqrt(1 - rate))
    }
    previous = change
    if ((iteration & 31) === 31) throwIfAborted(signal)
  }
}

interface Scratch {
  readonly x: Float64Array
  readonly b: Float64Array
  readonly r: Float64Array
}

function vCycle(levels: readonly Level[], transfers: readonly Transfer[], index: number, scratch: readonly Scratch[], control: SolveControl): void {
  const level = levels[index]
  const { x, b, r } = scratch[index]
  if (index === levels.length - 1) {
    solveCoarsest(level, x, b, control.maxIterations, COARSE_TOLERANCE, control.signal)
    return
  }
  relax(level, x, b, 2, 1)
  residual(level, x, b, r)
  const { parent, prolongIds, prolongWeights } = transfers[index]
  const coarse = scratch[index + 1]
  coarse.b.fill(0)
  coarse.x.fill(0)
  for (let k = 0; k < level.count; k += 1) {
    const id = parent[k]
    if (id < 0) continue
    for (let c = 0; c < 4; c += 1) coarse.b[id * 4 + c] += r[k * 4 + c]
  }
  vCycle(levels, transfers, index + 1, scratch, control)
  for (let k = 0; k < level.count; k += 1) {
    for (let s = 0; s < 4; s += 1) {
      const id = prolongIds[k * 4 + s]
      if (id < 0) continue
      const w = prolongWeights[k * 4 + s]
      for (let c = 0; c < 4; c += 1) x[k * 4 + c] += w * coarse.x[id * 4 + c]
    }
  }
  relax(level, x, b, 2, 1)
}

/** Solves the discrete Laplace equation for the unknown pixels (all four channels at once). */
function solveMembrane(problem: Membrane, control: SolveControl): void {
  const { width, height, unknown, defined, values } = problem
  const top = buildGrid(width, height, unknown, defined)
  const count = top.level.count
  if (!count) return
  const levels: Level[] = [top.level]
  const transfers: Transfer[] = []
  let grid = top
  while (grid.level.span > COARSEST_SPAN && grid.level.count > COARSEST_COUNT && levels.length < 14) {
    const next = coarsenGrid(grid)
    const coarse = next.grid.level
    if (!coarse.count) break
    // A coarse level must touch a fixed cell, or its correction problem is singular.
    let anchored = false
    for (let k = 0; k < coarse.count && !anchored; k += 1) {
      let linked = 0
      for (let d = 0; d < 4; d += 1) if (coarse.links[k * 4 + d] >= 0) linked += 1
      if (coarse.diag[k] > linked) anchored = true
    }
    if (!anchored) break
    levels.push(coarse)
    transfers.push(next.transfer)
    grid = next.grid
  }
  const scratch: Scratch[] = levels.map((level) => ({
    x: new Float64Array(level.count * 4),
    b: new Float64Array(level.count * 4),
    r: new Float64Array(level.count * 4),
  }))
  const { x, b, r } = scratch[0]
  // Right-hand side from the fixed neighbours; initial guess = their mean.
  const mean = [0, 0, 0, 0]
  let fixedCount = 0
  for (let k = 0; k < count; k += 1) {
    const p = top.pixels[k]
    const px = p % width
    const neighbours = [px > 0 ? p - 1 : -1, px + 1 < width ? p + 1 : -1, p - width, p + width]
    for (const q of neighbours) {
      if (q < 0 || q >= unknown.length || unknown[q] || !defined[q]) continue
      fixedCount += 1
      for (let c = 0; c < 4; c += 1) {
        b[k * 4 + c] += values[q * 4 + c]
        mean[c] += values[q * 4 + c]
      }
    }
  }
  for (let k = 0; k < count; k += 1) for (let c = 0; c < 4; c += 1) x[k * 4 + c] = fixedCount ? mean[c] / fixedCount : 0
  if (levels.length === 1) {
    solveCoarsest(top.level, x, b, control.maxIterations, COARSE_TOLERANCE, control.signal)
  } else {
    for (let cycle = 0; cycle < MAX_CYCLES; cycle += 1) {
      vCycle(levels, transfers, 0, scratch, control)
      if (residual(top.level, x, b, r) < RESIDUAL_TOLERANCE) break
      throwIfAborted(control.signal)
      if (cycle < 8) control.progress?.((cycle + 1) / 9)
    }
  }
  for (let k = 0; k < count; k += 1) {
    const p = top.pixels[k]
    for (let c = 0; c < 4; c += 1) values[p * 4 + c] = x[k * 4 + c]
  }
}

// ---------------------------------------------------------------------------------------------
// Spot healing
// ---------------------------------------------------------------------------------------------

/**
 * Heals the pixels under `hole` (a stroke mask the size of `src`, any non-zero value counts) from a
 * matching nearby source with a seamless (Poisson) clone; thin holes use Telea inpainting. Returns a new
 * buffer; pixels away from the dilated hole are unchanged. `src` should include healMargin() of context.
 */
export function healRegion(src: PixelBuffer, hole: MaskBuffer, options: HealOptions = {}, op: OpOptions = {}): PixelBuffer {
  checkInputs(src, hole)
  const { width, height } = src
  const total = width * height
  const data = src.data
  const stroke = new Uint8Array(total)
  let strokeCount = 0
  for (let i = 0; i < total; i += 1) {
    if (hole.data[i] !== 0) {
      stroke[i] = 1
      strokeCount += 1
    }
  }
  if (strokeCount === 0 || strokeCount === total) return copyBuffer(src)
  const signal = op.signal
  throwIfAborted(signal)

  const outside = new Uint8Array(total)
  for (let i = 0; i < total; i += 1) outside[i] = stroke[i] ^ 1
  const toOutside = distanceTransform(outside, width, height)
  let inscribed = 0
  for (let i = 0; i < total; i += 1) if (stroke[i] && toOutside[i] > inscribed) inscribed = toOutside[i]
  const toStroke = distanceTransform(stroke, width, height)
  if (inscribed <= THIN_RADIUS) {
    const thin = new Uint8Array(total)
    for (let i = 0; i < total; i += 1) thin[i] = toStroke[i] <= 1 ? 255 : 0
    return inpaintTelea(src, { width, height, data: thin }, TELEA_RADIUS, op)
  }

  // Dilated hole.
  const holeSet = new Uint8Array(total)
  const holeList: number[] = []
  let hx0 = width
  let hy0 = height
  let hx1 = 0
  let hy1 = 0
  for (let i = 0; i < total; i += 1) {
    if (toStroke[i] > DILATE) continue
    holeSet[i] = 1
    holeList.push(i)
    const x = i % width
    const y = (i / width) | 0
    if (x < hx0) hx0 = x
    if (x + 1 > hx1) hx1 = x + 1
    if (y < hy0) hy0 = y
    if (y + 1 > hy1) hy1 = y + 1
  }
  if (holeList.length === total) return copyBuffer(src)
  op.onProgress?.(0.05)

  const local = 2 * (inscribed + DILATE)
  const length = Math.max(hx1 - hx0, hy1 - hy0)
  const band = Math.max(3, Math.round(0.15 * local))
  const toHole = distanceTransform(holeSet, width, height)
  const bandList: number[] = []
  for (let i = 0; i < total; i += 1) if (!holeSet[i] && toHole[i] <= band) bandList.push(i)
  const rx0 = Math.max(0, hx0 - band)
  const ry0 = Math.max(0, hy0 - band)
  const rx1 = Math.min(width, hx1 + band)
  const ry1 = Math.min(height, hy1 + band)

  const valid = (ox: number, oy: number): boolean => {
    if (ox === 0 && oy === 0) return false
    if (rx0 + ox < 0 || ry0 + oy < 0 || rx1 + ox > width || ry1 + oy > height) return false
    if (rx1 + ox <= hx0 || rx0 + ox >= hx1 || ry1 + oy <= hy0 || ry0 + oy >= hy1) return true
    const shift = oy * width + ox
    for (const p of holeList) if (holeSet[p + shift]) return false
    for (const p of bandList) if (holeSet[p + shift]) return false
    return true
  }
  const ssd = (ox: number, oy: number, limit: number): number => {
    const shift = (oy * width + ox) * 4
    let sum = 0
    for (const p of bandList) {
      const a = p * 4
      const b = a + shift
      const d0 = data[a] - data[b]
      const d1 = data[a + 1] - data[b + 1]
      const d2 = data[a + 2] - data[b + 2]
      const d3 = data[a + 3] - data[b + 3]
      sum += d0 * d0 + d1 * d1 + d2 * d2 + d3 * d3
      if (sum >= limit) return sum
    }
    return sum
  }

  const rings = (options.rings ?? DEFAULT_RINGS).filter((ring) => Number.isFinite(ring) && ring > 0)
  const scales = [local]
  if (length > 1.5 * local) scales.push(length)
  const phase = seedPhase(options.seed)
  const tried = new Set<string>()
  let bestX = 0
  let bestY = 0
  let bestScore = Infinity
  const consider = (ox: number, oy: number) => {
    const key = `${ox},${oy}`
    if (tried.has(key)) return
    tried.add(key)
    if (!valid(ox, oy)) return
    const score = ssd(ox, oy, bestScore)
    if (score < bestScore) {
      bestScore = score
      bestX = ox
      bestY = oy
    }
  }
  for (const scale of scales) {
    for (const ring of rings.length ? rings : DEFAULT_RINGS) {
      for (let k = 0; k < ANGLES; k += 1) {
        const angle = ((k + phase) * 2 * Math.PI) / ANGLES
        consider(Math.round(ring * scale * Math.cos(angle)), Math.round(ring * scale * Math.sin(angle)))
      }
    }
    throwIfAborted(signal)
  }
  if (bestScore === Infinity) {
    // No ring fits (small image, hole near an edge): scan the offsets that keep the source inside the
    // image on a grid of at most ~4096 candidates before giving up.
    const minX = -rx0
    const maxX = width - rx1
    const minY = -ry0
    const maxY = height - ry1
    if (maxX >= minX && maxY >= minY) {
      const step = Math.max(1, Math.ceil(Math.sqrt(((maxX - minX + 1) * (maxY - minY + 1)) / 4096)))
      for (let oy = minY; oy <= maxY; oy += step) {
        for (let ox = minX; ox <= maxX; ox += step) consider(ox, oy)
        throwIfAborted(signal)
      }
    }
  }
  if (bestScore === Infinity) {
    // No clean source fits around the hole at all: fall back to diffusion.
    const fallback = new Uint8Array(total)
    for (const p of holeList) fallback[p] = 255
    return inpaintTelea(src, { width, height, data: fallback }, TELEA_RADIUS, op)
  }
  const centreX = bestX
  const centreY = bestY
  for (let dy = -REFINE; dy <= REFINE; dy += 1) {
    for (let dx = -REFINE; dx <= REFINE; dx += 1) consider(centreX + dx, centreY + dy)
  }
  op.onProgress?.(0.2)
  throwIfAborted(signal)

  // Seamless clone: solve for the correction c inside the hole.
  const shift = bestY * width + bestX
  const bw = rx1 - rx0
  const bh = ry1 - ry0
  const unknown = new Uint8Array(bw * bh)
  const defined = new Uint8Array(bw * bh)
  const values = new Float32Array(bw * bh * 4)
  for (let y = 0; y < bh; y += 1) {
    for (let x = 0; x < bw; x += 1) {
      const l = y * bw + x
      const g = (ry0 + y) * width + rx0 + x
      if (holeSet[g]) {
        unknown[l] = 1
      } else if (toHole[g] <= band) {
        defined[l] = 1
        const a = g * 4
        const b = (g + shift) * 4
        for (let c = 0; c < 4; c += 1) values[l * 4 + c] = data[a + c] - data[b + c]
      }
    }
  }
  const maxIterations = Math.max(1, Math.floor(options.maxIterations ?? MAX_ITERATIONS))
  solveMembrane({ width: bw, height: bh, unknown, defined, values }, {
    maxIterations,
    signal,
    progress: op.onProgress ? (fraction) => op.onProgress?.(0.2 + 0.75 * fraction) : undefined,
  })

  const out = copyBuffer(src)
  for (const g of holeList) {
    const x = g % width
    const y = (g / width) | 0
    const edge = (x > 0 && !holeSet[g - 1]) || (x + 1 < width && !holeSet[g + 1])
      || (y > 0 && !holeSet[g - width]) || (y + 1 < height && !holeSet[g + width])
    const weight = edge ? 0.5 : 1
    const l = (y - ry0) * bw + (x - rx0)
    const a = g * 4
    const b = (g + shift) * 4
    for (let c = 0; c < 4; c += 1) {
      const healed = data[b + c] + values[l * 4 + c]
      out.data[a + c] = Math.round(data[a + c] + (healed - data[a + c]) * weight)
    }
  }
  op.onProgress?.(1)
  return out
}

// ---------------------------------------------------------------------------------------------
// Telea fast-marching inpainting
// ---------------------------------------------------------------------------------------------

class MinHeap {
  keys: Float64Array
  items: Int32Array
  size = 0

  constructor(capacity: number) {
    this.keys = new Float64Array(Math.max(1, capacity))
    this.items = new Int32Array(Math.max(1, capacity))
  }

  push(item: number, key: number): void {
    if (this.size === this.keys.length) {
      const keys = new Float64Array(this.keys.length * 2)
      keys.set(this.keys)
      const items = new Int32Array(this.items.length * 2)
      items.set(this.items)
      this.keys = keys
      this.items = items
    }
    let i = this.size++
    while (i > 0) {
      const parent = (i - 1) >> 1
      if (this.keys[parent] <= key) break
      this.keys[i] = this.keys[parent]
      this.items[i] = this.items[parent]
      i = parent
    }
    this.keys[i] = key
    this.items[i] = item
  }

  pop(): number {
    const top = this.items[0]
    const size = --this.size
    if (size > 0) {
      const key = this.keys[size]
      const item = this.items[size]
      let i = 0
      for (;;) {
        let child = 2 * i + 1
        if (child >= size) break
        if (child + 1 < size && this.keys[child + 1] < this.keys[child]) child += 1
        if (this.keys[child] >= key) break
        this.keys[i] = this.keys[child]
        this.items[i] = this.items[child]
        i = child
      }
      this.keys[i] = key
      this.items[i] = item
    }
    return top
  }
}

/**
 * Fast-marching inpainting (Telea 2004): fills the pixels under `hole` (non-zero) from the boundary
 * inwards with weighted, gradient-extrapolated averages of known pixels within `radius` (default 5).
 * Returns a new buffer; pixels outside the hole are unchanged.
 */
export function inpaintTelea(src: PixelBuffer, hole: MaskBuffer, radius = TELEA_RADIUS, op: OpOptions = {}): PixelBuffer {
  checkInputs(src, hole)
  const { width, height } = src
  const out = copyBuffer(src)
  let hx0 = width
  let hy0 = height
  let hx1 = -1
  let hy1 = -1
  for (let y = 0; y < height; y += 1) {
    const row = y * width
    for (let x = 0; x < width; x += 1) {
      if (hole.data[row + x] === 0) continue
      if (x < hx0) hx0 = x
      if (x > hx1) hx1 = x
      if (y < hy0) hy0 = y
      if (y > hy1) hy1 = y
    }
  }
  if (hx1 < 0) return out
  const eps = Math.max(1, Number.isFinite(radius) ? radius : TELEA_RADIUS)
  const reach = Math.ceil(eps)
  const bx0 = Math.max(0, hx0 - reach - 1)
  const by0 = Math.max(0, hy0 - reach - 1)
  const bx1 = Math.min(width, hx1 + reach + 2)
  const by1 = Math.min(height, hy1 + reach + 2)
  const bw = bx1 - bx0
  const bh = by1 - by0
  const size = bw * bh
  const flags = new Uint8Array(size)
  const time = new Float32Array(size)
  const work = new Float32Array(size * 4)
  let known = 0
  for (let y = 0; y < bh; y += 1) {
    for (let x = 0; x < bw; x += 1) {
      const l = y * bw + x
      const g = (by0 + y) * width + bx0 + x
      for (let c = 0; c < 4; c += 1) work[l * 4 + c] = src.data[g * 4 + c]
      if (hole.data[g] !== 0) {
        flags[l] = INSIDE
        time[l] = 1e6
      } else {
        known += 1
      }
    }
  }
  if (known === 0) return out

  const available = (x: number, y: number): boolean => x >= 0 && y >= 0 && x < bw && y < bh && flags[y * bw + x] !== INSIDE
  const heap = new MinHeap(size)
  for (let y = 0; y < bh; y += 1) {
    for (let x = 0; x < bw; x += 1) {
      const l = y * bw + x
      if (flags[l] !== KNOWN) continue
      const touches = (x > 0 && flags[l - 1] === INSIDE) || (x + 1 < bw && flags[l + 1] === INSIDE)
        || (y > 0 && flags[l - bw] === INSIDE) || (y + 1 < bh && flags[l + bw] === INSIDE)
      if (!touches) continue
      flags[l] = BAND
      heap.push(l, 0)
    }
  }

  // Neighbourhood offsets inside the disk of radius eps, with the distance weight 1 / |r|^3.
  const offX: number[] = []
  const offY: number[] = []
  const offW: number[] = []
  for (let dy = -reach; dy <= reach; dy += 1) {
    for (let dx = -reach; dx <= reach; dx += 1) {
      const lengthSq = dx * dx + dy * dy
      if (lengthSq === 0 || lengthSq > eps * eps) continue
      offX.push(dx)
      offY.push(dy)
      offW.push(1 / (lengthSq * Math.sqrt(lengthSq)))
    }
  }

  const solve = (x1: number, y1: number, x2: number, y2: number): number => {
    const a = available(x1, y1)
    const b = available(x2, y2)
    const ta = a ? time[y1 * bw + x1] : 1e6
    const tb = b ? time[y2 * bw + x2] : 1e6
    if (a && b) {
      const diff = ta - tb
      if (Math.abs(diff) >= 1) return 1 + Math.min(ta, tb)
      return (ta + tb + Math.sqrt(2 - diff * diff)) * 0.5
    }
    if (a) return 1 + ta
    if (b) return 1 + tb
    return 1e6
  }

  const sums = new Float64Array(4)
  const lows = new Float64Array(4)
  const highs = new Float64Array(4)
  const inpaintPixel = (px: number, py: number) => {
    const l = py * bw + px
    const t = time[l]
    let gx = 0
    let gy = 0
    if (available(px + 1, py)) gx = available(px - 1, py) ? (time[l + 1] - time[l - 1]) * 0.5 : time[l + 1] - t
    else if (available(px - 1, py)) gx = t - time[l - 1]
    if (available(px, py + 1)) gy = available(px, py - 1) ? (time[l + bw] - time[l - bw]) * 0.5 : time[l + bw] - t
    else if (available(px, py - 1)) gy = t - time[l - bw]
    sums.fill(0)
    lows.fill(Infinity)
    highs.fill(-Infinity)
    let weightSum = 0
    for (let k = 0; k < offX.length; k += 1) {
      const qx = px + offX[k]
      const qy = py + offY[k]
      if (!available(qx, qy)) continue
      const q = qy * bw + qx
      // r = p - q
      const rx = -offX[k]
      const ry = -offY[k]
      let direction = rx * gx + ry * gy
      if (Math.abs(direction) <= 0.01) direction = 0.000001
      const level = 1 / (1 + Math.abs(time[q] - t))
      const w = Math.abs(offW[k] * level * direction)
      const right = available(qx + 1, qy)
      const left = available(qx - 1, qy)
      const down = available(qx, qy + 1)
      const up = available(qx, qy - 1)
      for (let c = 0; c < 4; c += 1) {
        const value = work[q * 4 + c]
        let ix = 0
        let iy = 0
        if (right) ix = left ? (work[(q + 1) * 4 + c] - work[(q - 1) * 4 + c]) * 0.5 : work[(q + 1) * 4 + c] - value
        else if (left) ix = value - work[(q - 1) * 4 + c]
        if (down) iy = up ? (work[(q + bw) * 4 + c] - work[(q - bw) * 4 + c]) * 0.5 : work[(q + bw) * 4 + c] - value
        else if (up) iy = value - work[(q - bw) * 4 + c]
        sums[c] += w * (value + ix * rx + iy * ry)
        if (value < lows[c]) lows[c] = value
        if (value > highs[c]) highs[c] = value
      }
      weightSum += w
    }
    if (weightSum <= 0) return
    for (let c = 0; c < 4; c += 1) {
      let value = sums[c] / weightSum
      if (value < lows[c]) value = lows[c]
      if (value > highs[c]) value = highs[c]
      work[l * 4 + c] = value
    }
  }

  let steps = 0
  while (heap.size > 0) {
    const l = heap.pop()
    flags[l] = KNOWN
    const x = l % bw
    const y = (l / bw) | 0
    const neighbours = [[x - 1, y], [x + 1, y], [x, y - 1], [x, y + 1]]
    for (const [nx, ny] of neighbours) {
      if (nx < 0 || ny < 0 || nx >= bw || ny >= bh) continue
      const n = ny * bw + nx
      if (flags[n] !== INSIDE) continue
      time[n] = Math.min(
        solve(nx - 1, ny, nx, ny - 1),
        solve(nx + 1, ny, nx, ny - 1),
        solve(nx - 1, ny, nx, ny + 1),
        solve(nx + 1, ny, nx, ny + 1),
      )
      inpaintPixel(nx, ny)
      flags[n] = BAND
      heap.push(n, time[n])
    }
    if ((++steps & 4095) === 0) throwIfAborted(op.signal)
  }

  for (let y = 0; y < bh; y += 1) {
    for (let x = 0; x < bw; x += 1) {
      const g = (by0 + y) * width + bx0 + x
      if (hole.data[g] === 0) continue
      const l = y * bw + x
      for (let c = 0; c < 4; c += 1) out.data[g * 4 + c] = Math.round(work[l * 4 + c])
    }
  }
  return out
}
