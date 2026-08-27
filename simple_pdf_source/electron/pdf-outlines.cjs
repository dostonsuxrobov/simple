'use strict'

/**
 * Replace a pdf-lib document's native /Outlines tree.
 *
 * pdf-lib 1.17 has no public bookmarks API, so this deliberately uses its
 * stable low-level PDF object model. Keeping it in a separate module lets the
 * Electron main process continue loading pdf-lib lazily.
 *
 * @param {import('pdf-lib').PDFDocument} pdfDoc
 * @param {Array<object>} entries Flattened pre-order outline entries.
 * @param {typeof import('pdf-lib')} pdfLib
 * @returns {{ written: number, skipped: number }}
 */
function replacePdfOutlines(pdfDoc, entries, pdfLib) {
  const { PDFArray, PDFDict, PDFHexString, PDFName, PDFNumber } = pdfLib
  const context = pdfDoc.context
  const pageCount = pdfDoc.getPageCount()
  const source = Array.isArray(entries) ? entries : []

  removeExistingOutlineTree(pdfDoc, pdfLib)

  if (source.length === 0) {
    return { written: 0, skipped: 0 }
  }

  const root = { children: [] }
  const ancestors = []
  const nodes = []
  let skipped = 0

  for (const entry of source) {
    if (!entry || typeof entry !== 'object') {
      skipped += 1
      continue
    }
    const titleValue = typeof entry.title === 'string' ? entry.title : entry.label
    const title = typeof titleValue === 'string' ? titleValue.trim() : ''
    const hasPageIndex = entry.pageIndex !== null && entry.pageIndex !== undefined
    const pageIndex = hasPageIndex && Number.isInteger(entry.pageIndex)
      && entry.pageIndex >= 0 && entry.pageIndex < pageCount
      ? entry.pageIndex
      : null
    if (!title || (hasPageIndex && pageIndex === null)) {
      skipped += 1
      continue
    }

    const requestedDepth = Number.isInteger(entry.depth)
      ? Math.max(0, Math.min(256, entry.depth))
      : 0
    // A depth jump can only introduce one new level; clamp malformed input to
    // the deepest parent that actually exists.
    const depth = Math.min(requestedDepth, ancestors.length)
    ancestors.length = depth
    const parent = depth === 0 ? root : ancestors[depth - 1]
    const node = {
      title,
      pageIndex,
      url: typeof entry.url === 'string' && entry.url.trim() ? entry.url.trim() : null,
      bold: entry.bold === true,
      italic: entry.italic === true,
      color: normalizeColor(entry.color),
      open: entry.expanded !== false && entry.open !== false,
      children: [],
      parent,
      ref: context.nextRef(),
    }
    parent.children.push(node)
    nodes.push(node)
    ancestors.push(node)
  }

  if (nodes.length === 0) {
    return { written: 0, skipped }
  }

  const rootRef = context.nextRef()
  assignSiblings(root.children)
  for (const node of nodes) assignSiblings(node.children)

  for (const node of nodes) {
    const map = new Map()
    map.set(PDFName.of('Title'), PDFHexString.fromText(node.title))
    map.set(PDFName.of('Parent'), node.parent === root ? rootRef : node.parent.ref)
    if (node.previous) map.set(PDFName.of('Prev'), node.previous.ref)
    if (node.next) map.set(PDFName.of('Next'), node.next.ref)

    if (node.children.length) {
      map.set(PDFName.of('First'), node.children[0].ref)
      map.set(PDFName.of('Last'), node.children[node.children.length - 1].ref)
      const visibleDescendants = countVisibleChildren(node)
      map.set(PDFName.of('Count'), PDFNumber.of(node.open ? visibleDescendants : -visibleDescendants))
    }

    if (node.pageIndex !== null) {
      const destination = PDFArray.withContext(context)
      destination.push(pdfDoc.getPage(node.pageIndex).ref)
      destination.push(PDFName.of('Fit'))
      map.set(PDFName.of('Dest'), destination)
    } else if (node.url) {
      map.set(PDFName.of('A'), context.obj({
        S: 'URI',
        URI: PDFHexString.fromText(node.url),
      }))
    }

    const flags = (node.italic ? 1 : 0) | (node.bold ? 2 : 0)
    if (flags) map.set(PDFName.of('F'), PDFNumber.of(flags))
    if (node.color && node.color.some((channel) => channel !== 0)) {
      const color = PDFArray.withContext(context)
      for (const channel of node.color) color.push(PDFNumber.of(channel))
      map.set(PDFName.of('C'), color)
    }
    context.assign(node.ref, PDFDict.fromMapWithContext(map, context))
  }

  const rootMap = new Map()
  rootMap.set(PDFName.of('Type'), PDFName.of('Outlines'))
  rootMap.set(PDFName.of('First'), root.children[0].ref)
  rootMap.set(PDFName.of('Last'), root.children[root.children.length - 1].ref)
  rootMap.set(PDFName.of('Count'), PDFNumber.of(countVisibleChildren(root)))
  context.assign(rootRef, PDFDict.fromMapWithContext(rootMap, context))
  pdfDoc.catalog.set(PDFName.of('Outlines'), rootRef)

  return { written: nodes.length, skipped }
}

/** Detach and discard the old outline dictionaries to prevent save-after-save bloat. */
function removeExistingOutlineTree(pdfDoc, pdfLib) {
  const { PDFDict, PDFName, PDFRef } = pdfLib
  const context = pdfDoc.context
  const outlinesKey = PDFName.of('Outlines')
  const existing = pdfDoc.catalog.get(outlinesKey)
  pdfDoc.catalog.delete(outlinesKey)
  if (!existing) return

  const pending = [existing]
  const references = []
  const seenReferences = new Set()
  const seenDictionaries = new WeakSet()
  // First descends into children and Next walks siblings. Parent/Prev/Last are
  // intentionally ignored so malformed cycles cannot escape this outline tree.
  while (pending.length && seenReferences.size < 250_000) {
    const candidate = pending.pop()
    let dictionary = candidate
    if (candidate instanceof PDFRef) {
      if (seenReferences.has(candidate.tag)) continue
      seenReferences.add(candidate.tag)
      references.push(candidate)
      dictionary = context.lookup(candidate)
    }
    if (!(dictionary instanceof PDFDict) || seenDictionaries.has(dictionary)) continue
    seenDictionaries.add(dictionary)
    const first = dictionary.get(PDFName.of('First'))
    const next = dictionary.get(PDFName.of('Next'))
    if (first) pending.push(first)
    if (next) pending.push(next)
  }
  for (const reference of references) context.delete(reference)
}

function assignSiblings(children) {
  for (let index = 0; index < children.length; index += 1) {
    children[index].previous = index > 0 ? children[index - 1] : null
    children[index].next = index + 1 < children.length ? children[index + 1] : null
  }
}

function countVisibleChildren(parent) {
  let count = 0
  for (const child of parent.children) {
    count += 1
    if (child.open) count += countVisibleChildren(child)
  }
  return count
}

function normalizeColor(value) {
  if (!value || typeof value.length !== 'number' || value.length < 3) return null
  const channels = [0, 1, 2].map((index) => {
    const channel = Number(value[index])
    if (!Number.isFinite(channel)) return 0
    const normalized = channel > 1 ? channel / 255 : channel
    return Math.max(0, Math.min(1, normalized))
  })
  return channels
}

module.exports = { replacePdfOutlines }
