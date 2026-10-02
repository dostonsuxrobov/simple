import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { OcrPageResult } from './types'

// Recognition results, kept for the window's lifetime so a page is never read
// twice: undoing a recognition and running it again, or recognising a page
// whose pixels have not changed, is instant.
//
// Results are keyed by content (the OCR input bitmap's hash, the engine
// settings and the pixel-to-PDF mapping; see recognizePage), never by the
// pdf.js fingerprint: that survives reordering and deleting pages, so it would
// hand one page's text to another. A per-document index then tells which
// result belongs to which page of the document now on screen.

export const OCR_CACHE_LIMITS = Object.freeze({
  pages: 300,
  /** Estimated size of the stored results (UTF-16 JSON). */
  bytes: 60 * 1024 * 1024,
})

interface CacheEntry {
  result: OcrPageResult
  bytes: number
}

const entries = new Map<string, CacheEntry>()
let storedBytes = 0

function estimateBytes(result: OcrPageResult) {
  try {
    return JSON.stringify(result).length * 2
  } catch {
    return 256 * 1024
  }
}

function evict(contentKey: string) {
  const entry = entries.get(contentKey)
  if (!entry) return
  entries.delete(contentKey)
  storedBytes -= entry.bytes
}

/** The stored result for this content key (most recently used entries are kept longest). */
export function getCachedOcrResult(contentKey: string): OcrPageResult | undefined {
  const entry = entries.get(contentKey)
  if (!entry) return undefined
  entries.delete(contentKey)
  entries.set(contentKey, entry)
  return entry.result
}

/** Remember a final result; the oldest results go once the limits are reached. */
export function storeOcrResult(contentKey: string, result: OcrPageResult) {
  if (!contentKey || !result) return
  const bytes = estimateBytes(result)
  if (bytes > OCR_CACHE_LIMITS.bytes) return
  evict(contentKey)
  entries.set(contentKey, { result, bytes })
  storedBytes += bytes
  while (entries.size > OCR_CACHE_LIMITS.pages || storedBytes > OCR_CACHE_LIMITS.bytes) {
    const oldest = entries.keys().next().value
    if (oldest === undefined) break
    evict(oldest)
  }
}

// Each document proxy (every change to the bytes makes a new one) maps its
// pages to content keys. Proxies that are gone take their index with them.
const documentIndex = new WeakMap<PDFDocumentProxy, Map<number, string>>()

/** Record that `pageIndex` of this document shows the content recognised under `contentKey`. */
export function bindOcrPage(pdf: PDFDocumentProxy, pageIndex: number, contentKey: string) {
  if (!Number.isInteger(pageIndex) || pageIndex < 0 || !contentKey) return
  let pages = documentIndex.get(pdf)
  if (!pages) documentIndex.set(pdf, pages = new Map())
  pages.set(pageIndex, contentKey)
}

export function ocrContentKeyForPage(pdf: PDFDocumentProxy, pageIndex: number): string | undefined {
  return documentIndex.get(pdf)?.get(pageIndex)
}

/**
 * The recognition result for a page of this document, when it was recognised
 * in this window (exact word boxes for editing scanned text). Callers must not
 * rely on it: a result can have been evicted, and pages recognised by other
 * tools have none.
 */
export function ocrResultForPage(pdf: PDFDocumentProxy, pageIndex: number): OcrPageResult | undefined {
  const contentKey = ocrContentKeyForPage(pdf, pageIndex)
  const result = contentKey ? getCachedOcrResult(contentKey) : undefined
  if (!result) return undefined
  return result.pageIndex === pageIndex ? result : { ...result, pageIndex }
}

/** Tests and diagnostics. */
export function ocrCacheStats() {
  return { pages: entries.size, bytes: storedBytes }
}

export function clearOcrCache() {
  entries.clear()
  storedBytes = 0
}
