// src/imaging/looks.ts (WP2)
// Simple mode's one-click looks: each is a short list of adjustment specs plus an optional vignette, blended
// with the original by an intensity of 0..100 %. Pure and DOM-free. The vignette depends on the position in
// the whole image, so a look must run on the whole image (never striped).
import type { AdjustmentSpec, CurvePoint, LookId, OpOptions, PixelBuffer } from './types.ts'
import { startAdjustments } from './adjustments.ts'
import { assertBuffer, chunkRowsFor, cloneBuffer, finishedRun, runRowsSync, throwIfAborted } from './buffer.ts'
import type { RowRun } from './buffer.ts'

export interface LookDefinition {
  readonly label: string
  readonly specs: readonly AdjustmentSpec[]
  /** 0..1: how much the corners darken. */
  readonly vignette: number
}

const LINEAR: readonly CurvePoint[] = [{ x: 0, y: 0 }, { x: 255, y: 255 }]

function curves(rgb: readonly CurvePoint[], blue: readonly CurvePoint[] = LINEAR): AdjustmentSpec {
  return { type: 'curves', rgb, red: LINEAR, green: LINEAR, blue }
}

const WARMING = { r: 236, g: 138, b: 0 }
const COOLING = { r: 0, g: 109, b: 255 }
const BW_DEFAULT = { reds: 40, yellows: 60, greens: 40, cyans: 60, blues: 20, magentas: 80 }

function deepFreeze<T>(value: T): T {
  if (value && typeof value === 'object' && !Object.isFrozen(value)) {
    Object.freeze(value)
    for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child)
  }
  return value
}

/** Strip order. */
export const LOOK_IDS: readonly LookId[] = Object.freeze(['none', 'vivid', 'warm', 'cool', 'mono', 'sepia', 'vintage', 'dramatic', 'fade'] as const)

export const LOOKS: Readonly<Record<LookId, LookDefinition>> = deepFreeze({
  none: { label: 'Original', specs: [], vignette: 0 },
  vivid: {
    label: 'Vivid',
    specs: [
      { type: 'vibrance', vibrance: 40, saturation: 10 },
      { type: 'brightness-contrast', brightness: 0, contrast: 20, legacy: false },
    ],
    vignette: 0,
  },
  warm: {
    label: 'Warm',
    specs: [
      { type: 'photo-filter', color: WARMING, density: 30, preserveLuminosity: true },
      { type: 'vibrance', vibrance: 10, saturation: 0 },
    ],
    vignette: 0,
  },
  cool: {
    label: 'Cool',
    specs: [{ type: 'photo-filter', color: COOLING, density: 30, preserveLuminosity: true }],
    vignette: 0,
  },
  mono: {
    label: 'Mono',
    specs: [
      { type: 'black-white', ...BW_DEFAULT, tint: null },
      { type: 'brightness-contrast', brightness: 0, contrast: 15, legacy: false },
    ],
    vignette: 0,
  },
  sepia: {
    label: 'Sepia',
    specs: [
      { type: 'black-white', ...BW_DEFAULT, tint: { r: 162, g: 128, b: 101 } },
      curves([{ x: 0, y: 12 }, { x: 255, y: 248 }]),
    ],
    vignette: 0.15,
  },
  vintage: {
    label: 'Vintage',
    specs: [
      curves([{ x: 0, y: 28 }, { x: 128, y: 132 }, { x: 255, y: 232 }], [{ x: 0, y: 24 }, { x: 255, y: 220 }]),
      { type: 'vibrance', vibrance: -10, saturation: -25 },
      { type: 'photo-filter', color: WARMING, density: 15, preserveLuminosity: true },
    ],
    vignette: 0.25,
  },
  dramatic: {
    label: 'Dramatic',
    specs: [
      curves([{ x: 0, y: 0 }, { x: 64, y: 46 }, { x: 192, y: 212 }, { x: 255, y: 255 }]),
      { type: 'vibrance', vibrance: 0, saturation: -15 },
    ],
    vignette: 0.35,
  },
  fade: {
    label: 'Fade',
    specs: [
      curves([{ x: 0, y: 45 }, { x: 255, y: 235 }]),
      { type: 'vibrance', vibrance: 0, saturation: -30 },
    ],
    vignette: 0,
  },
} satisfies Record<LookId, LookDefinition>)

/**
 * Darkens towards the corners in place, rows [startRow, endRow) only (default all): factor
 * 1 - amount * smoothstep((r - 0.35) / 0.65) with r = 1 at the corners of the whole buffer.
 */
export function applyVignette(buffer: PixelBuffer, amount: number, startRow = 0, endRow = buffer.height): void {
  const strength = amount >= 1 ? 1 : amount > 0 ? amount : 0
  if (strength === 0) return
  const { width, height, data } = buffer
  const halfX = width / 2
  const halfY = height / 2
  const columns = new Float64Array(width)
  for (let x = 0; x < width; x += 1) {
    const d = (x + 0.5 - halfX) / halfX
    columns[x] = (d * d) / 2
  }
  const first = Math.max(0, startRow)
  const last = Math.min(height, endRow)
  for (let y = first; y < last; y += 1) {
    const dy = (y + 0.5 - halfY) / halfY
    const rowTerm = (dy * dy) / 2
    for (let x = 0; x < width; x += 1) {
      const r = Math.sqrt(columns[x] + rowTerm)
      if (r <= 0.35) continue
      const t = r >= 1 ? 1 : (r - 0.35) / 0.65
      const factor = 1 - strength * t * t * (3 - 2 * t)
      const i = (y * width + x) * 4
      data[i] *= factor
      data[i + 1] *= factor
      data[i + 2] *= factor
    }
  }
}

/** The resumable form of applyLook (worker handlers run it in chunks of rows). */
export function startLook(src: PixelBuffer, look: LookId, intensity = 100): RowRun {
  assertBuffer(src)
  if (!Object.prototype.hasOwnProperty.call(LOOKS, look)) throw new RangeError(`Unknown look "${String(look)}".`)
  const definition = LOOKS[look]
  const amount = Number.isFinite(intensity) ? Math.max(0, Math.min(100, intensity)) / 100 : 1
  if (look === 'none' || amount === 0) return finishedRun(cloneBuffer(src))
  const adjust = startAdjustments(src, definition.specs)
  const output = adjust.output
  const original = src.data
  const out = output.data
  const rowBytes = src.width * 4
  return {
    output,
    rows: src.height,
    chunkRows: chunkRowsFor(src.width, 1 << 19),
    process(startRow, endRow) {
      adjust.process(startRow, endRow)
      applyVignette(output, definition.vignette, startRow, endRow)
      if (amount >= 1) return
      for (let i = startRow * rowBytes, stop = endRow * rowBytes; i < stop; i += 4) {
        out[i] = original[i] + (out[i] - original[i]) * amount
        out[i + 1] = original[i + 1] + (out[i + 1] - original[i + 1]) * amount
        out[i + 2] = original[i + 2] + (out[i + 2] - original[i + 2]) * amount
      }
    },
  }
}

/** Applies a look at `intensity` percent (0..100, default 100) and returns a new buffer; alpha is unchanged. */
export function applyLook(src: PixelBuffer, look: LookId, intensity = 100, options?: OpOptions): PixelBuffer {
  throwIfAborted(options?.signal)
  return runRowsSync(startLook(src, look, intensity), options)
}
