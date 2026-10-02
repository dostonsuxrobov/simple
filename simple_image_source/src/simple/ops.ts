// src/simple/ops.ts (WP8)
// Canvas-side operations of Simple mode. The Simple canvas stays the image store (design section 4), so
// these helpers read and write it directly:
//   - createCanvasSurface(): the HistorySurface the undo store swaps pixels through;
//   - transformCanvas(): exact quarter turns and flips (canvas-to-canvas copies of the premultiplied
//     pixels with smoothing off: no rounding, half-transparent pixels stay exact);
//   - crop, whole-image replacement, transparency scans of a region, colour sampling, proxies for previews,
//     and encoding with per-format quality.
// Every scratch canvas is released (width = height = 1) before the operation returns; qa/composition-stress
// counts the leftovers.
import type { IntRect, PixelBuffer, Rgba8 } from '../imaging/types.ts'
import { createScratchCanvas, putBuffer, readCanvasPixels, releaseCanvas } from '../shared/canvas.ts'
import type { GeometryOp, HistorySurface } from './simpleHistory.ts'
import { geometrySize } from './simpleHistory.ts'

export type EncodeFormat = 'png' | 'jpeg' | 'webp'

/** The image canvas's 2D context. The first call creates it with willReadFrequently (CPU-backed readbacks). */
export function imageContext(canvas: HTMLCanvasElement): CanvasRenderingContext2D {
  const context = canvas.getContext('2d', { willReadFrequently: true })
  if (!context) throw new Error('The image canvas is unavailable.')
  return context
}

/** Reads a region (stripes for large regions keep the peak small). */
export function readRegion(canvas: HTMLCanvasElement, rect: IntRect): PixelBuffer {
  if (rect.x === 0 && rect.y === 0 && rect.width === canvas.width && rect.height === canvas.height && rect.width * rect.height > 4_000_000) {
    return readCanvasPixels(canvas)
  }
  return imageContext(canvas).getImageData(rect.x, rect.y, rect.width, rect.height)
}

/** Replaces the whole canvas with `pixels` (the size may change). */
export function replaceCanvasPixels(canvas: HTMLCanvasElement, pixels: PixelBuffer): void {
  canvas.width = pixels.width
  canvas.height = pixels.height
  putBuffer(canvas, pixels, 0, 0)
}

/** Exact quarter turn or flip of the canvas in place. */
export function transformCanvas(canvas: HTMLCanvasElement, op: GeometryOp): void {
  const width = canvas.width
  const height = canvas.height
  const copy = createScratchCanvas(width, height)
  try {
    const copyContext = copy.getContext('2d')
    if (!copyContext) throw new Error('The image could not be turned.')
    copyContext.drawImage(canvas, 0, 0)
    const size = geometrySize(width, height, op)
    // Assigning the size also clears the bitmap and resets the context state.
    canvas.width = size.width
    canvas.height = size.height
    const context = imageContext(canvas)
    context.save()
    context.imageSmoothingEnabled = false
    if (op === 'rotate-cw' || op === 'rotate-ccw') {
      context.translate(size.width / 2, size.height / 2)
      context.rotate((op === 'rotate-cw' ? 1 : -1) * Math.PI / 2)
      context.drawImage(copy, -width / 2, -height / 2)
    } else if (op === 'flip-h') {
      context.translate(width, 0)
      context.scale(-1, 1)
      context.drawImage(copy, 0, 0)
    } else {
      context.translate(0, height)
      context.scale(1, -1)
      context.drawImage(copy, 0, 0)
    }
    context.restore()
  } finally {
    releaseCanvas(copy)
  }
}

/** Crops the canvas to `rect` (whole pixels inside the canvas) with an exact copy. */
export function cropCanvas(canvas: HTMLCanvasElement, rect: IntRect): void {
  const copy = createScratchCanvas(rect.width, rect.height)
  try {
    const copyContext = copy.getContext('2d')
    if (!copyContext) throw new Error('The image could not be cropped.')
    copyContext.drawImage(canvas, rect.x, rect.y, rect.width, rect.height, 0, 0, rect.width, rect.height)
    canvas.width = rect.width
    canvas.height = rect.height
    const context = imageContext(canvas)
    context.imageSmoothingEnabled = false
    context.drawImage(copy, 0, 0)
    context.imageSmoothingEnabled = true
  } finally {
    releaseCanvas(copy)
  }
}

export function createCanvasSurface(canvas: HTMLCanvasElement): HistorySurface {
  return {
    get width() { return canvas.width },
    get height() { return canvas.height },
    read: (rect) => readRegion(canvas, rect),
    write: (x, y, pixels) => putBuffer(canvas, pixels, x, y),
    replace: (pixels) => replaceCanvasPixels(canvas, pixels),
    transform: (op) => transformCanvas(canvas, op),
  }
}

/** True when any pixel of `rect` (default: the whole canvas) has alpha below 255; read in stripes. */
export function regionHasTransparency(canvas: HTMLCanvasElement, rect: IntRect | null = null): boolean {
  const area = rect ?? { x: 0, y: 0, width: canvas.width, height: canvas.height }
  if (area.width <= 0 || area.height <= 0) return false
  const context = imageContext(canvas)
  const rows = Math.max(1, Math.min(256, Math.floor(16_000_000 / Math.max(4, area.width * 4))))
  for (let top = area.y; top < area.y + area.height; top += rows) {
    const height = Math.min(rows, area.y + area.height - top)
    const data = context.getImageData(area.x, top, area.width, height).data
    for (let index = 3; index < data.length; index += 4) {
      if (data[index] < 255) return true
    }
  }
  return false
}

/** Encodes the canvas. JPEG is flattened onto white first (JPEG has no transparency). */
export function encodeCanvas(canvas: HTMLCanvasElement, format: EncodeFormat, quality?: number): Promise<Blob> {
  let source = canvas
  if (format === 'jpeg') {
    source = createScratchCanvas(canvas.width, canvas.height)
    const context = source.getContext('2d', { alpha: false })
    if (!context) {
      releaseCanvas(source)
      return Promise.reject(new Error('The image canvas is unavailable.'))
    }
    context.fillStyle = '#ffffff'
    context.fillRect(0, 0, source.width, source.height)
    context.drawImage(canvas, 0, 0)
  }
  const mime = format === 'jpeg' ? 'image/jpeg' : `image/${format}`
  return new Promise<Blob>((resolve, reject) => {
    source.toBlob((blob) => (blob ? resolve(blob) : reject(new Error('The image could not be encoded.'))), mime, format === 'png' ? undefined : quality)
  }).finally(() => {
    if (source !== canvas) releaseCanvas(source)
  })
}

/** Average colour of the (2 * radius + 1)^2 pixels around (x, y), clipped to the canvas; null outside. */
export function sampleColor(canvas: HTMLCanvasElement, x: number, y: number, radius = 1): Rgba8 | null {
  const cx = Math.floor(x)
  const cy = Math.floor(y)
  if (cx < 0 || cy < 0 || cx >= canvas.width || cy >= canvas.height) return null
  const x0 = Math.max(0, cx - radius)
  const y0 = Math.max(0, cy - radius)
  const x1 = Math.min(canvas.width, cx + radius + 1)
  const y1 = Math.min(canvas.height, cy + radius + 1)
  const data = imageContext(canvas).getImageData(x0, y0, x1 - x0, y1 - y0).data
  let r = 0
  let g = 0
  let b = 0
  let a = 0
  let count = 0
  for (let index = 0; index < data.length; index += 4) {
    // Alpha-weighted so transparent neighbours do not pull the colour towards black.
    const weight = data[index + 3]
    r += data[index] * weight
    g += data[index + 1] * weight
    b += data[index + 2] * weight
    a += weight
    count += 1
  }
  if (!count) return null
  if (a === 0) return { r: 0, g: 0, b: 0, a: 0 }
  return { r: Math.round(r / a), g: Math.round(g / a), b: Math.round(b / a), a: Math.round(a / count) }
}

/** One pixel, exactly as stored (status-bar readout). */
export function pixelAt(canvas: HTMLCanvasElement, x: number, y: number): Rgba8 | null {
  if (x < 0 || y < 0 || x >= canvas.width || y >= canvas.height) return null
  const data = imageContext(canvas).getImageData(Math.floor(x), Math.floor(y), 1, 1).data
  return { r: data[0], g: data[1], b: data[2], a: data[3] }
}

export function rgbHex(color: { r: number; g: number; b: number }): string {
  const hex = (value: number) => Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, '0')
  return `#${hex(color.r)}${hex(color.g)}${hex(color.b)}`
}

/** Size of a proxy of `width` x `height` at `scale` (<= 1) that also stays under `maxPixels`. */
export function proxySize(width: number, height: number, scale: number, maxPixels: number): { width: number; height: number } {
  const byPixels = Math.sqrt(maxPixels / Math.max(1, width * height))
  const factor = Math.min(1, Math.max(1e-6, scale), byPixels)
  return { width: Math.max(1, Math.round(width * factor)), height: Math.max(1, Math.round(height * factor)) }
}

/** A downscaled copy of the canvas (high-quality smoothing) as straight RGBA. */
export function canvasProxy(canvas: HTMLCanvasElement, size: { width: number; height: number }): PixelBuffer {
  if (size.width === canvas.width && size.height === canvas.height) return readRegion(canvas, { x: 0, y: 0, width: canvas.width, height: canvas.height })
  const scratch = createScratchCanvas(size.width, size.height)
  try {
    const context = scratch.getContext('2d', { willReadFrequently: true })
    if (!context) throw new Error('The preview could not be prepared.')
    context.imageSmoothingEnabled = true
    context.imageSmoothingQuality = 'high'
    context.drawImage(canvas, 0, 0, size.width, size.height)
    return context.getImageData(0, 0, size.width, size.height)
  } finally {
    releaseCanvas(scratch)
  }
}

/** A data URL of `pixels` (for small thumbnails). */
export function pixelsToDataUrl(pixels: PixelBuffer): string {
  const scratch = createScratchCanvas(pixels.width, pixels.height)
  try {
    putBuffer(scratch, pixels, 0, 0)
    return scratch.toDataURL('image/png')
  } finally {
    releaseCanvas(scratch)
  }
}

/**
 * The crop box of a straightened preview as a new canvas: the image turned by `degrees` about the frame
 * centre (canvas resampling; used for clipboard copies of the selection only). The caller releases it.
 */
export function rotatedRegionCanvas(canvas: HTMLCanvasElement, frame: { width: number; height: number }, rect: IntRect, degrees: number): HTMLCanvasElement {
  const output = createScratchCanvas(rect.width, rect.height)
  const context = output.getContext('2d')
  if (!context) {
    releaseCanvas(output)
    throw new Error('The selected image area could not be prepared.')
  }
  context.imageSmoothingQuality = 'high'
  context.translate(frame.width / 2 - rect.x, frame.height / 2 - rect.y)
  context.rotate((degrees * Math.PI) / 180)
  context.drawImage(canvas, -canvas.width / 2, -canvas.height / 2)
  return output
}
