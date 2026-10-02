import type { Bookmark } from '../types'
import type { PdfOutlineBookmark } from './pdf'

/**
 * Bookmarks are kept as the flat, pre-order list the sidebar shows, nested by
 * `depth`. Saving hands the list to electron/pdf-outlines.cjs, which reuses
 * the document's own outline items (exact destinations, web links, headings,
 * styles) and only patches what changed, so the list must keep every item of
 * the document's outline, including those that do not go to a page.
 */

export interface BookmarkNode {
  bookmark: Bookmark
  children: BookmarkNode[]
}

/** The document's outline as editable bookmarks: every item, headings and web links included. */
export function bookmarksFromOutline(outline: PdfOutlineBookmark[]): Bookmark[] {
  return outline.map((item) => ({
    id: item.id,
    pageIndex: item.pageIndex,
    label: item.title,
    depth: item.depth,
    source: 'document' as const,
    expanded: item.expanded,
    bold: item.bold,
    italic: item.italic,
    color: item.color,
    url: item.url,
    // pdf.js numbers outline items "2.1" (first child of the second
    // top-level item); the writer finds the item it came from by that place.
    outlinePath: item.id.replace(/^pdf-outline:/, ''),
    originalLabel: item.title,
  }))
}

/**
 * Nest the flat list. A depth may only go one level deeper than the item
 * before it; deeper jumps are clamped exactly as the PDF writer clamps them.
 */
export function bookmarkTree(bookmarks: Bookmark[]): BookmarkNode[] {
  const roots: BookmarkNode[] = []
  const ancestors: BookmarkNode[] = []
  for (const bookmark of bookmarks) {
    const requested = Number.isInteger(bookmark.depth) ? Math.max(0, Number(bookmark.depth)) : 0
    const depth = Math.min(requested, ancestors.length)
    ancestors.length = depth
    const node: BookmarkNode = { bookmark, children: [] }
    if (depth === 0) roots.push(node)
    else ancestors[depth - 1].children.push(node)
    ancestors.push(node)
  }
  return roots
}

/** The pre-order list again, with depths that match the nesting. */
export function flattenBookmarkTree(nodes: BookmarkNode[], depth = 0, output: Bookmark[] = []): Bookmark[] {
  for (const node of nodes) {
    output.push((node.bookmark.depth ?? 0) === depth ? node.bookmark : { ...node.bookmark, depth })
    flattenBookmarkTree(node.children, depth + 1, output)
  }
  return output
}

function findNode(nodes: BookmarkNode[], id: string): BookmarkNode | null {
  for (const node of nodes) {
    if (node.bookmark.id === id) return node
    const nested = findNode(node.children, id)
    if (nested) return nested
  }
  return null
}

function countNodes(nodes: BookmarkNode[]): number {
  return nodes.reduce((total, node) => total + 1 + countNodes(node.children), 0)
}

/** How many bookmarks removing `id` takes away: the bookmark and everything nested under it. */
export function bookmarkSubtreeSize(bookmarks: Bookmark[], id: string) {
  const node = findNode(bookmarkTree(bookmarks), id)
  return node ? countNodes([node]) : 0
}

/** Remove a bookmark together with the bookmarks nested under it, as Acrobat does. */
export function withoutBookmark(bookmarks: Bookmark[], id: string): Bookmark[] {
  const prune = (nodes: BookmarkNode[]): BookmarkNode[] => nodes.flatMap((node) => (
    node.bookmark.id === id ? [] : [{ bookmark: node.bookmark, children: prune(node.children) }]
  ))
  return flattenBookmarkTree(prune(bookmarkTree(bookmarks)))
}

/** New page numbers after pages move or are inserted; a bookmark without a page keeps none. */
export function remapBookmarkPages(bookmarks: Bookmark[], mapIndex: (pageIndex: number) => number): Bookmark[] {
  return bookmarks.map((bookmark) => (
    typeof bookmark.pageIndex === 'number'
      ? { ...bookmark, pageIndex: mapIndex(bookmark.pageIndex) }
      : bookmark
  ))
}

/**
 * The list after `deleted` pages are removed, matching how the main process
 * prunes the document's outline: a bookmark that went to a deleted page is
 * removed, unless bookmarks are nested under it, in which case it stays as a
 * heading without a page. Other page numbers shift down.
 */
export function bookmarksAfterPageDelete(bookmarks: Bookmark[], deleted: number[]): Bookmark[] {
  const removed = new Set(deleted)
  const shift = (pageIndex: number) => pageIndex - deleted.filter((index) => index < pageIndex).length
  const prune = (nodes: BookmarkNode[]): BookmarkNode[] => nodes.flatMap((node) => {
    const children = prune(node.children)
    const { pageIndex } = node.bookmark
    if (typeof pageIndex !== 'number') return [{ bookmark: node.bookmark, children }]
    if (removed.has(pageIndex)) {
      return children.length ? [{ bookmark: { ...node.bookmark, pageIndex: null }, children }] : []
    }
    return [{ bookmark: { ...node.bookmark, pageIndex: shift(pageIndex) }, children }]
  })
  return flattenBookmarkTree(prune(bookmarkTree(bookmarks)))
}

/**
 * The bookmark the toolbar button added for this page. The button adds and
 * removes only its own "Page N" bookmarks; it never removes a chapter of the
 * document's outline that happens to start on the same page.
 */
export function toolbarBookmarkFor(bookmarks: Bookmark[], pageIndex: number) {
  return bookmarks.find((bookmark) => bookmark.source !== 'document' && bookmark.pageIndex === pageIndex)
}

/** What the sidebar shows under a bookmark's title. */
export function bookmarkTargetLabel(bookmark: Bookmark) {
  if (typeof bookmark.pageIndex === 'number') return `Page ${bookmark.pageIndex + 1}`
  if (bookmark.url) return 'Web link'
  return 'Heading'
}
