import type { PDFPageProxy } from 'pdfjs-dist'
import type { DetectedPageObject, PdfRect } from '../types'
import { pdfjs } from './pdf'

type Matrix = [number, number, number, number, number, number]

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0]

interface PdfImageData {
  width: number
  height: number
  data?: Uint8Array | Uint8ClampedArray
  bitmap?: CanvasImageSource
  kind?: number
}

interface PendingPageObject extends DetectedPageObject {
  matrix: Matrix
  imageObjectId?: string
  imageData?: PdfImageData
}

function asMatrix(value: unknown): Matrix | null {
  if (!Array.isArray(value) || value.length < 6) return null
  const matrix = value.slice(0, 6).map(Number)
  return matrix.every(Number.isFinite) ? matrix as Matrix : null
}

function rectFromMatrix(matrix: Matrix): PdfRect | null {
  const bounds = pdfjs.Util.getAxialAlignedBoundingBox([0, 0, 1, 1], matrix)
  const rect = {
    x: Math.min(bounds[0], bounds[2]),
    y: Math.min(bounds[1], bounds[3]),
    width: Math.abs(bounds[2] - bounds[0]),
    height: Math.abs(bounds[3] - bounds[1]),
  }
  return rect.width >= 6 && rect.height >= 6 ? rect : null
}

function matrixWith(current: Matrix, next: unknown): Matrix {
  const matrix = asMatrix(next)
  return matrix ? pdfjs.Util.transform(current, matrix) as Matrix : current
}

function imageObject(page: PDFPageProxy, objectId: string): Promise<PdfImageData | null> {
  try {
    return Promise.resolve(page.objs.get(objectId) as PdfImageData)
  } catch {
    return new Promise((resolve) => {
      try {
        page.objs.get(objectId, (value: PdfImageData) => resolve(value || null))
      } catch {
        resolve(null)
      }
    })
  }
}

function decodedImageCanvas(image: PdfImageData): HTMLCanvasElement | null {
  const width = Math.max(1, Math.floor(Number(image.width) || 0))
  const height = Math.max(1, Math.floor(Number(image.height) || 0))
  if (!width || !height) return null
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d')
  if (!context) return null
  if (image.bitmap) {
    context.drawImage(image.bitmap, 0, 0, width, height)
    return canvas
  }
  if (!image.data) return null

  const rgba = context.createImageData(width, height)
  const source = image.data
  if (image.kind === pdfjs.ImageKind.RGBA_32BPP || source.length >= width * height * 4) {
    rgba.data.set(source.subarray(0, rgba.data.length))
  } else if (image.kind === pdfjs.ImageKind.RGB_24BPP || source.length >= width * height * 3) {
    for (let sourceIndex = 0, targetIndex = 0; targetIndex < rgba.data.length; sourceIndex += 3, targetIndex += 4) {
      rgba.data[targetIndex] = source[sourceIndex]
      rgba.data[targetIndex + 1] = source[sourceIndex + 1]
      rgba.data[targetIndex + 2] = source[sourceIndex + 2]
      rgba.data[targetIndex + 3] = 255
    }
  } else if (image.kind === pdfjs.ImageKind.GRAYSCALE_1BPP) {
    const rowBytes = Math.ceil(width / 8)
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const byte = source[y * rowBytes + Math.floor(x / 8)] || 0
        const value = byte & (128 >> (x % 8)) ? 255 : 0
        const index = (y * width + x) * 4
        rgba.data[index] = value
        rgba.data[index + 1] = value
        rgba.data[index + 2] = value
        rgba.data[index + 3] = 255
      }
    }
  } else {
    return null
  }
  context.putImageData(rgba, 0, 0)
  return canvas
}

/** Render only the decoded image through its PDF transform, never the page canvas. */
function placedImageDataUrl(image: PdfImageData, matrix: Matrix, rect: PdfRect): string | undefined {
  const source = decodedImageCanvas(image)
  if (!source || rect.width <= 0 || rect.height <= 0) return undefined
  const sourceScale = Math.min(source.width / rect.width, source.height / rect.height)
  const pixelsPerPoint = Math.max(1, Math.min(4, Number.isFinite(sourceScale) ? sourceScale : 1))
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.ceil(rect.width * pixelsPerPoint))
  canvas.height = Math.max(1, Math.ceil(rect.height * pixelsPerPoint))
  const context = canvas.getContext('2d')
  if (!context) return undefined

  const [a, b, c, d, e, f] = matrix
  context.setTransform(
    pixelsPerPoint * a / source.width,
    -pixelsPerPoint * b / source.width,
    -pixelsPerPoint * c / source.height,
    pixelsPerPoint * d / source.height,
    pixelsPerPoint * (c + e - rect.x),
    pixelsPerPoint * (rect.y + rect.height - d - f),
  )
  context.imageSmoothingEnabled = true
  context.imageSmoothingQuality = 'high'
  context.drawImage(source, 0, 0)
  return canvas.toDataURL('image/png')
}

const pageObjectsCache = new WeakMap<PDFPageProxy, Promise<DetectedPageObject[]>>()

/**
 * Finds raster-image paint operations exposed by PDF.js. These bounds are a
 * rendering aid, not a mutation API; the original PDF object remains untouched.
 * Detection walks the full operator list and decodes every image, so the result
 * is shared across mounts and tool switches for the lifetime of the page proxy.
 */
export function detectPageObjects(page: PDFPageProxy, pageIndex: number): Promise<DetectedPageObject[]> {
  let request = pageObjectsCache.get(page)
  if (!request) {
    request = extractPageObjects(page, pageIndex)
    pageObjectsCache.set(page, request)
    request.catch(() => pageObjectsCache.delete(page))
  }
  return request
}

async function extractPageObjects(page: PDFPageProxy, pageIndex: number): Promise<DetectedPageObject[]> {
  const list = await page.getOperatorList()
  const candidates: PendingPageObject[] = []
  const stack: Matrix[] = []
  let matrix: Matrix = [...IDENTITY]

  const add = (
    nextMatrix: Matrix,
    operatorIndex: number,
    suffix = '',
    imageObjectId?: string,
    imageData?: PdfImageData,
  ) => {
    const rect = rectFromMatrix(nextMatrix)
    if (!rect) return
    candidates.push({
      id: `image-${pageIndex}-${operatorIndex}${suffix}`,
      pageIndex,
      kind: 'image',
      rect,
      matrix: [...nextMatrix],
      imageObjectId,
      imageData,
      label: 'Image',
    })
  }

  list.fnArray.forEach((operation, index) => {
    const args = list.argsArray[index] || []
    if (operation === pdfjs.OPS.save) {
      stack.push([...matrix])
      return
    }
    if (operation === pdfjs.OPS.restore) {
      matrix = stack.pop() || [...IDENTITY]
      return
    }
    if (operation === pdfjs.OPS.transform) {
      matrix = matrixWith(matrix, args)
      return
    }
    if (operation === pdfjs.OPS.paintFormXObjectBegin) {
      stack.push([...matrix])
      matrix = matrixWith(matrix, args[0])
      return
    }
    if (operation === pdfjs.OPS.paintFormXObjectEnd) {
      matrix = stack.pop() || [...IDENTITY]
      return
    }
    if (operation === pdfjs.OPS.paintImageXObject || operation === pdfjs.OPS.paintInlineImageXObject) {
      const source = args[0]
      add(
        matrix,
        index,
        '',
        typeof source === 'string' ? source : undefined,
        source && typeof source === 'object' ? source as PdfImageData : undefined,
      )
      return
    }
    if (operation === pdfjs.OPS.paintImageXObjectRepeat) {
      const scaleX = Number(args[1])
      const scaleY = Number(args[2])
      const positions = args[3]
      if (!Number.isFinite(scaleX) || !Number.isFinite(scaleY) || (!Array.isArray(positions) && !ArrayBuffer.isView(positions))) return
      const numericPositions = Array.from(positions as unknown as ArrayLike<number>)
      for (let offset = 0; offset + 1 < numericPositions.length; offset += 2) {
        add(
          matrixWith(matrix, [scaleX, 0, 0, scaleY, Number(numericPositions[offset]), Number(numericPositions[offset + 1])]),
          index,
          `-${offset / 2}`,
          typeof args[0] === 'string' ? args[0] : undefined,
        )
      }
      return
    }
    if (operation === pdfjs.OPS.paintInlineImageXObjectGroup) {
      const maps = args[1]
      if (!Array.isArray(maps)) return
      maps.forEach((entry, entryIndex) => add(
        matrixWith(matrix, entry?.transform),
        index,
        `-${entryIndex}`,
        undefined,
        args[0] && typeof args[0] === 'object' ? args[0] as PdfImageData : undefined,
      ))
    }
  })

  // Some producers paint the same image through nested forms. Keep the largest
  // unique bounds and avoid a stack of indistinguishable selection outlines.
  const unique = candidates.filter((candidate, index) => !candidates.some((other, otherIndex) => {
    if (otherIndex >= index) return false
    const tolerance = 0.75
    return Math.abs(other.rect.x - candidate.rect.x) < tolerance
      && Math.abs(other.rect.y - candidate.rect.y) < tolerance
      && Math.abs(other.rect.width - candidate.rect.width) < tolerance
      && Math.abs(other.rect.height - candidate.rect.height) < tolerance
  }))

  return Promise.all(unique.map(async ({ matrix, imageObjectId, imageData, ...candidate }) => {
    const decoded = imageData || (imageObjectId ? await imageObject(page, imageObjectId) : null)
    return {
      ...candidate,
      dataUrl: decoded ? placedImageDataUrl(decoded, matrix, candidate.rect) : undefined,
    }
  }))
}
