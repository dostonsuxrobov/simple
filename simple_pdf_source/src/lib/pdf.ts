import * as pdfjs from 'pdfjs-dist'
import PdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url'
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'
import type { DocumentPayload, PdfRect } from '../types'

pdfjs.GlobalWorkerOptions.workerSrc = PdfWorker

export function normalizeBytes(data: DocumentPayload['data'] | Uint8Array): Uint8Array {
  // IPC and PDF operations already hand us a fresh typed array. Reusing it here
  // avoids a second full-document memory copy before pdf.js starts loading.
  if (data instanceof Uint8Array) return data
  if (data instanceof ArrayBuffer) return new Uint8Array(data)
  if (data && data.type === 'Buffer' && Array.isArray(data.data)) return new Uint8Array(data.data)
  throw new Error('The document data could not be read.')
}

export async function loadPdf(bytes: Uint8Array, onPassword?: (update: (password: string) => void, reason: number) => void): Promise<PDFDocumentProxy> {
  const task = pdfjs.getDocument({
    data: bytes.slice(),
    // Keep dynamic evaluation disabled for untrusted PDFs. The largest opening
    // wins come from avoiding duplicate work rather than weakening this guard.
    isEvalSupported: false,
    useWorkerFetch: true,
    // Retain decoded embedded-font bytes so a text edit can reuse the PDF's
    // original face when it is written back instead of silently substituting.
    fontExtraProperties: true,
  })
  if (onPassword) task.onPassword = onPassword
  return task.promise
}

/** A PDF-native outline item (called a bookmark in Acrobat), in display order. */
export interface PdfOutlineBookmark {
  /** Stable for the lifetime of this outline tree and safe to use as a React key. */
  id: string
  title: string
  /** Zero-based destination page, or null for headings/external actions. */
  pageIndex: number | null
  /** One-based destination page, or null for headings/external actions. */
  pageNumber: number | null
  depth: number
  hasChildren: boolean
  expanded: boolean
  bold: boolean
  italic: boolean
  color: [number, number, number]
  url: string | null
}

interface PdfJsOutlineNode {
  title?: unknown
  bold?: unknown
  italic?: unknown
  color?: ArrayLike<number> | null
  dest?: string | unknown[] | null
  url?: string | null
  count?: number
  items?: PdfJsOutlineNode[]
}

type PdfPageReference = Parameters<PDFDocumentProxy['getPageIndex']>[0]

function outlineColor(value: ArrayLike<number> | null | undefined): [number, number, number] {
  return [0, 1, 2].map((index) => {
    const channel = Number(value?.[index])
    return Number.isFinite(channel) ? Math.max(0, Math.min(255, Math.round(channel))) : 0
  }) as [number, number, number]
}

function isPageReference(value: unknown): value is PdfPageReference {
  if (!value || typeof value !== 'object') return false
  const candidate = value as { num?: unknown; gen?: unknown }
  return Number.isInteger(candidate.num) && Number.isInteger(candidate.gen)
}

/**
 * Read and flatten the document's embedded outline tree.
 *
 * This resolves both direct destinations and Adobe-style named destinations.
 * A broken destination is isolated to its own item instead of hiding the rest of
 * the outline. Repeated references are memoized so large tables of contents do
 * not cause one PDF worker round-trip per item.
 */
export async function getPdfOutlineBookmarks(pdf: PDFDocumentProxy): Promise<PdfOutlineBookmark[]> {
  const outline = (await pdf.getOutline()) as PdfJsOutlineNode[] | null
  if (!Array.isArray(outline) || outline.length === 0) return []

  const flattened: Array<{ bookmark: PdfOutlineBookmark; destination: string | unknown[] | null }> = []
  const stack: Array<{ node: PdfJsOutlineNode; depth: number; path: number[] }> = []
  for (let index = outline.length - 1; index >= 0; index -= 1) {
    stack.push({ node: outline[index], depth: 0, path: [index] })
  }

  // PDF.js normally rejects cyclic outline chains itself, but keep the renderer
  // bounded if a malformed document still produces a cyclic or enormous tree.
  const visited = new WeakSet<object>()
  const maximumItems = 100_000
  const maximumDepth = 256
  while (stack.length && flattened.length < maximumItems) {
    const { node, depth, path } = stack.pop()!
    if (!node || typeof node !== 'object' || visited.has(node)) continue
    visited.add(node)

    const children = Array.isArray(node.items) ? node.items : []
    const title = typeof node.title === 'string' && node.title.trim()
      ? node.title.trim()
      : 'Untitled bookmark'
    flattened.push({
      bookmark: {
        id: `pdf-outline:${path.map((index) => index + 1).join('.')}`,
        title,
        pageIndex: null,
        pageNumber: null,
        depth,
        hasChildren: children.length > 0,
        expanded: typeof node.count !== 'number' || node.count >= 0,
        bold: node.bold === true,
        italic: node.italic === true,
        color: outlineColor(node.color),
        url: typeof node.url === 'string' ? node.url : null,
      },
      destination: node.dest ?? null,
    })

    if (depth >= maximumDepth) continue
    for (let index = children.length - 1; index >= 0; index -= 1) {
      stack.push({ node: children[index], depth: depth + 1, path: [...path, index] })
    }
  }

  const namedDestinationCache = new Map<string, Promise<unknown[] | null>>()
  const pageReferenceCache = new Map<string, Promise<number | null>>()

  const resolvePageReference = (reference: PdfPageReference): Promise<number | null> => {
    const key = `${reference.num}:${reference.gen}`
    let request = pageReferenceCache.get(key)
    if (!request) {
      request = (async () => {
        try {
          const cachedPageNumber = pdf.cachedPageNumber(reference)
          const pageIndex = cachedPageNumber === null
            ? await pdf.getPageIndex(reference)
            : cachedPageNumber - 1
          return Number.isInteger(pageIndex) && pageIndex >= 0 && pageIndex < pdf.numPages
            ? pageIndex
            : null
        } catch {
          return null
        }
      })()
      pageReferenceCache.set(key, request)
    }
    return request
  }

  const resolveDestination = async (destination: string | unknown[] | null): Promise<number | null> => {
    let explicitDestination: unknown[] | null
    if (typeof destination === 'string') {
      let request = namedDestinationCache.get(destination)
      if (!request) {
        request = pdf.getDestination(destination).catch(() => null)
        namedDestinationCache.set(destination, request)
      }
      explicitDestination = await request
    } else {
      explicitDestination = Array.isArray(destination) ? destination : null
    }

    if (!explicitDestination?.length) return null
    const pageReference = explicitDestination[0]
    // PDF.js represents the rare integer form as a zero-based page index.
    if (Number.isInteger(pageReference)) {
      const pageIndex = Number(pageReference)
      return pageIndex >= 0 && pageIndex < pdf.numPages ? pageIndex : null
    }
    return isPageReference(pageReference) ? resolvePageReference(pageReference) : null
  }

  // Avoid flooding the PDF worker with thousands of destination lookups at once.
  let cursor = 0
  const resolveNext = async () => {
    while (cursor < flattened.length) {
      const index = cursor
      cursor += 1
      const item = flattened[index]
      const pageIndex = await resolveDestination(item.destination)
      item.bookmark.pageIndex = pageIndex
      item.bookmark.pageNumber = pageIndex === null ? null : pageIndex + 1
    }
  }
  const concurrency = Math.min(12, flattened.length)
  await Promise.all(Array.from({ length: concurrency }, resolveNext))

  return flattened.map(({ bookmark }) => bookmark)
}

const textContentCache = new WeakMap<PDFPageProxy, ReturnType<PDFPageProxy['getTextContent']>>()

/** Share extraction work between the selectable text layer, search, and zooms. */
export function getPageTextContent(page: PDFPageProxy): ReturnType<PDFPageProxy['getTextContent']> {
  let request = textContentCache.get(page)
  if (!request) {
    request = page.getTextContent()
    textContentCache.set(page, request)
    request.catch(() => textContentCache.delete(page))
  }
  return request
}

export function viewportRectToPdf(
  viewport: { convertToPdfPoint: (x: number, y: number) => number[] },
  rect: { left: number; top: number; right: number; bottom: number },
): PdfRect {
  const [x1, y1] = viewport.convertToPdfPoint(rect.left, rect.top)
  const [x2, y2] = viewport.convertToPdfPoint(rect.right, rect.bottom)
  return {
    x: Math.min(x1, x2),
    y: Math.min(y1, y2),
    width: Math.abs(x2 - x1),
    height: Math.abs(y2 - y1),
  }
}

export function pdfRectToViewport(
  viewport: { convertToViewportRectangle: (rect: [number, number, number, number]) => number[] },
  rect: PdfRect,
) {
  const converted = viewport.convertToViewportRectangle([
    rect.x,
    rect.y,
    rect.x + rect.width,
    rect.y + rect.height,
  ])
  return {
    left: Math.min(converted[0], converted[2]),
    top: Math.min(converted[1], converted[3]),
    width: Math.abs(converted[2] - converted[0]),
    height: Math.abs(converted[3] - converted[1]),
  }
}

export function makeId(prefix: string) {
  return `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`
}

export { pdfjs }
