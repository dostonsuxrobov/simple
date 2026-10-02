// src/shared/canvas.ts (WP1)
// DOM canvas helpers shared by Simple and Advanced mode.
// Rules: a canvas is never the source of truth for layer pixels (premultiplied storage rounds
// low-alpha colours); every scratch canvas is released with releaseCanvas() (width = height = 1)
// before the operation that created it finishes (qa/composition-stress.cjs counts leftovers).
import type { PixelBuffer } from '../imaging/types.ts'

type Context2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D

function context2d(canvas: HTMLCanvasElement | OffscreenCanvas, readback: boolean): Context2D {
  // getContext returns the existing context (and ignores the options) once one was created.
  const options: CanvasRenderingContext2DSettings | undefined = readback ? { willReadFrequently: true } : undefined
  const context = (canvas as HTMLCanvasElement).getContext('2d', options) as Context2D | null
  if (!context) throw new Error('A drawing surface could not be created. Close other large images and try again.')
  return context
}

function asImageData(buffer: PixelBuffer): ImageData {
  if (typeof ImageData !== 'undefined' && buffer instanceof ImageData) return buffer
  if (buffer.data.length !== buffer.width * buffer.height * 4) {
    throw new RangeError(`Pixel buffer size mismatch: ${buffer.data.length} bytes for ${buffer.width} x ${buffer.height}.`)
  }
  // ImageData adopts the array without copying.
  return new ImageData(buffer.data as Uint8ClampedArray<ArrayBuffer>, buffer.width, buffer.height)
}

/** A detached canvas for temporary work. The caller must release it with releaseCanvas(). */
export function createScratchCanvas(width: number, height: number): HTMLCanvasElement {
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.floor(width))
  canvas.height = Math.max(1, Math.floor(height))
  return canvas
}

/** Frees the backing store immediately instead of waiting for garbage collection. */
export function releaseCanvas(canvas: HTMLCanvasElement | OffscreenCanvas | null): void {
  if (!canvas) return
  canvas.width = 1
  canvas.height = 1
}

/** Runs `use` with a scratch canvas and always releases it afterwards. */
export async function withScratchCanvas<T>(
  width: number,
  height: number,
  use: (canvas: HTMLCanvasElement) => Promise<T> | T,
): Promise<T> {
  const canvas = createScratchCanvas(width, height)
  try {
    return await use(canvas)
  } finally {
    releaseCanvas(canvas)
  }
}

/**
 * Reads the canvas top to bottom in stripes of `rows` rows (256-row stripes read a 24 MP canvas
 * about 2.5x faster than one full readback). Each stripe is a fresh straight-alpha ImageData.
 */
export function readCanvasStripes(canvas: HTMLCanvasElement, rows: number, visit: (y: number, stripe: PixelBuffer) => void): void {
  const { width, height } = canvas
  if (!width || !height) return
  const context = context2d(canvas, true)
  const step = Math.max(1, Math.floor(rows) || 1)
  for (let top = 0; top < height; top += step) {
    visit(top, context.getImageData(0, top, width, Math.min(step, height - top)))
  }
}

/** Whole-canvas copy assembled from stripes (bounded peak memory: one stripe plus the result). */
export function readCanvasPixels(canvas: HTMLCanvasElement, rows = 256): PixelBuffer {
  const { width, height } = canvas
  const data = new Uint8ClampedArray(width * height * 4)
  readCanvasStripes(canvas, rows, (top, stripe) => data.set(stripe.data, top * width * 4))
  return { width, height, data }
}

/** Draws straight-alpha pixels at (x, y), replacing what is there (putImageData semantics). */
export function putBuffer(canvas: HTMLCanvasElement | OffscreenCanvas, buffer: PixelBuffer, x: number, y: number): void {
  if (!buffer.width || !buffer.height) return
  context2d(canvas, false).putImageData(asImageData(buffer), Math.round(x), Math.round(y))
}

/** A new canvas holding exactly `buffer`. The caller must release it with releaseCanvas(). */
export function bufferToCanvas(buffer: PixelBuffer): HTMLCanvasElement {
  const canvas = createScratchCanvas(buffer.width, buffer.height)
  try {
    putBuffer(canvas, buffer, 0, 0)
  } catch (error) {
    releaseCanvas(canvas)
    throw error
  }
  return canvas
}
