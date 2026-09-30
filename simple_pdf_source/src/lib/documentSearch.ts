import type { SearchResult } from '../types'
import { normalizeSearchValue } from './search'

interface IndexedPageText {
  text: string
  normalized: string
}

interface SearchOptions {
  isCancelled?: () => boolean
}

interface SearchScheduler {
  yieldToUi?: () => Promise<void>
  now?: () => number
}

/**
 * Keep the text index for the lifetime of a PDF proxy, including sidebar remounts.
 * Only completed pages are cached; a cancelled request must not prevent a later
 * query from finishing that page. The reader can stop between worker requests.
 */
export function createDocumentSearch<Document extends { numPages: number }>(
  readPageText: (pdf: Document, pageIndex: number, isCancelled: () => boolean) => Promise<string | null>,
  scheduler: SearchScheduler = {},
) {
  const documents = new WeakMap<Document, Map<number, IndexedPageText>>()
  const yieldToUi = scheduler.yieldToUi ?? (() => new Promise<void>((resolve) => setTimeout(resolve, 0)))
  const now = scheduler.now ?? (() => performance.now())

  return async function searchDocument(
    pdf: Document,
    query: string,
    { isCancelled = () => false }: SearchOptions = {},
  ): Promise<SearchResult[] | null> {
    if (isCancelled()) return null
    const needle = normalizeSearchValue(query)
    if (!needle) return []

    let pages = documents.get(pdf)
    if (!pages) {
      pages = new Map()
      documents.set(pdf, pages)
    }

    const matches: SearchResult[] = []
    let lastYield = now()
    for (let pageIndex = 0; pageIndex < pdf.numPages; pageIndex += 1) {
      if (isCancelled()) return null
      let indexed = pages.get(pageIndex)
      if (!indexed) {
        const value = await readPageText(pdf, pageIndex, isCancelled)
        if (isCancelled()) return null
        if (value === null) return null
        const text = value.replace(/\s+/g, ' ').trim()
        indexed = { text, normalized: normalizeSearchValue(text) }
        pages.set(pageIndex, indexed)
      }

      const first = indexed.normalized.indexOf(needle)
      let count = 0
      let offset = first
      while (offset >= 0) {
        count += 1
        offset = indexed.normalized.indexOf(needle, offset + needle.length)
      }
      if (count) {
        const start = Math.max(0, first - 42)
        const end = Math.min(indexed.text.length, first + needle.length + 58)
        matches.push({
          pageIndex,
          excerpt: `${start ? '…' : ''}${indexed.text.slice(start, end)}${end < indexed.text.length ? '…' : ''}`,
          count,
        })
      }

      // Cached promises otherwise keep long searches in the microtask queue,
      // delaying query edits, cancellation, scrolling and the search spinner.
      if (pageIndex + 1 < pdf.numPages && ((pageIndex + 1) % 32 === 0 || now() - lastYield >= 8)) {
        await yieldToUi()
        if (isCancelled()) return null
        lastYield = now()
      }
    }
    return isCancelled() ? null : matches
  }
}
