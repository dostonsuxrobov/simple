'use strict'

const UNTITLED = 'Untitled bookmark'
const MAX_OUTLINE_ITEMS = 250_000

/**
 * Write the renderer's bookmark list into a pdf-lib document's /Outlines tree.
 *
 * `entries` is the complete outline in display (pre-) order, nested by
 * `depth`. Bookmarks read from the document carry `source: 'document'`, the
 * `outlinePath` they had when it was opened ("2.1" = first child of the second
 * top-level item, as pdf.js numbers them) and their `originalLabel`. Their
 * outline dictionaries are reused as they are: the destination (an /XYZ
 * position and zoom), actions (web links, named actions, links to other
 * files), style and every other key survive; only a changed title, or a
 * target page that moved, is patched. The tree links (/Parent /Prev /Next
 * /First /Last /Count) are rebuilt from the list, and nothing at all is written
 * when the list still matches the document. Items the list no longer contains
 * are removed; new bookmarks get a /Fit destination.
 *
 * pdf-lib 1.17 has no bookmarks API, so this uses its low-level object model.
 * Keeping it in a separate module lets the Electron main process continue
 * loading pdf-lib lazily.
 *
 * @param {import('pdf-lib').PDFDocument} pdfDoc
 * @param {Array<object>} entries Flattened pre-order outline entries.
 * @param {typeof import('pdf-lib')} pdfLib
 * @returns {{ written: number, skipped: number }}
 */
function replacePdfOutlines(pdfDoc, entries, pdfLib) {
  const { PDFDict, PDFHexString, PDFName, PDFNumber } = pdfLib
  const context = pdfDoc.context
  const pageRefs = pdfDoc.getPages().map((page) => page.ref)
  const pageCount = pageRefs.length
  const source = Array.isArray(entries) ? entries : []
  const existing = readOutlineTree(pdfDoc, pdfLib)
  const targets = destinationResolver(pdfDoc, pdfLib, pageRefs)

  // The list as a tree. A depth jump can only introduce one new level; clamp
  // malformed input to the deepest parent that actually exists.
  const root = { children: [], open: true }
  const ancestors = []
  const desired = []
  let skipped = 0
  for (const entry of source) {
    if (!entry || typeof entry !== 'object') {
      skipped += 1
      continue
    }
    const fromDocument = entry.source === 'document'
    const titleValue = typeof entry.title === 'string' ? entry.title : entry.label
    const title = typeof titleValue === 'string' ? titleValue.trim() : ''
    const hasPageIndex = entry.pageIndex !== null && entry.pageIndex !== undefined
    const pageIndex = hasPageIndex && Number.isInteger(entry.pageIndex)
      && entry.pageIndex >= 0 && entry.pageIndex < pageCount
      ? entry.pageIndex
      : null
    // A document bookmark keeps its own destination, so an unusable page
    // number never drops it (or re-parents its children); a new one needs a
    // page that exists.
    if (!title || (!fromDocument && hasPageIndex && pageIndex === null)) {
      skipped += 1
      continue
    }
    const requestedDepth = Number.isInteger(entry.depth) ? Math.max(0, Math.min(256, entry.depth)) : 0
    const depth = Math.min(requestedDepth, ancestors.length)
    ancestors.length = depth
    const parent = depth === 0 ? root : ancestors[depth - 1]
    const node = {
      entry,
      fromDocument,
      title,
      depth,
      pageIndex,
      url: typeof entry.url === 'string' && entry.url.trim() ? entry.url.trim() : null,
      bold: entry.bold === true,
      italic: entry.italic === true,
      color: normalizeColor(entry.color),
      open: entry.expanded !== false && entry.open !== false,
      children: [],
      parent,
      existing: null,
      ref: null,
      dict: null,
    }
    parent.children.push(node)
    desired.push(node)
    ancestors.push(node)
  }

  matchDocumentBookmarks(desired, existing, targets)

  if (isUnchanged(root, desired, existing, targets)) return { written: desired.length, skipped }

  for (const node of desired) {
    if (node.existing) {
      node.ref = node.existing.ref
      node.dict = node.existing.dict
      // An open item has a positive /Count; keep the reader's choice.
      const count = node.dict.lookup(PDFName.of('Count'))
      node.open = !(count instanceof PDFNumber && count.asNumber() < 0)
      if (sameTitle(node.title, node.existing.title) === false) node.dict.set(PDFName.of('Title'), PDFHexString.fromText(node.title))
      retargetIfMoved(node, targets, pageRefs, pdfLib)
    } else {
      node.ref = context.nextRef()
      node.dict = newOutlineItem(node, context, pageRefs, pdfLib)
      context.assign(node.ref, node.dict)
    }
  }

  // Items the list no longer contains were deleted by the user (or, for
  // stale entries, replaced above by fresh dictionaries).
  const reused = new Set(desired.filter((node) => node.existing).map((node) => node.existing))
  for (const item of existing.nodes) {
    if (!reused.has(item)) context.delete(item.ref)
  }

  if (!desired.length) {
    pdfDoc.catalog.delete(PDFName.of('Outlines'))
    if (existing.rootRef) context.delete(existing.rootRef)
    return { written: 0, skipped }
  }

  const rootRef = existing.rootRef || context.nextRef()
  const rootDict = existing.rootDict || PDFDict.withContext(context)
  rootDict.set(PDFName.of('Type'), PDFName.of('Outlines'))
  link(rootDict, null, root.children, true)
  for (const node of desired) {
    const parentRef = node.parent === root ? rootRef : node.parent.ref
    node.dict.set(PDFName.of('Parent'), parentRef)
    link(node.dict, node, node.children, node.open)
  }
  if (!existing.rootDict || !existing.rootRef) context.assign(rootRef, rootDict)
  pdfDoc.catalog.set(PDFName.of('Outlines'), rootRef)
  return { written: desired.length, skipped }

  function link(dict, owner, children, open) {
    children.forEach((child, index) => {
      if (index > 0) child.dict.set(PDFName.of('Prev'), children[index - 1].ref)
      else child.dict.delete(PDFName.of('Prev'))
      if (index + 1 < children.length) child.dict.set(PDFName.of('Next'), children[index + 1].ref)
      else child.dict.delete(PDFName.of('Next'))
    })
    if (!children.length) {
      dict.delete(PDFName.of('First'))
      dict.delete(PDFName.of('Last'))
      if (owner) dict.delete(PDFName.of('Count'))
      else dict.set(PDFName.of('Count'), PDFNumber.of(0))
      return
    }
    dict.set(PDFName.of('First'), children[0].ref)
    dict.set(PDFName.of('Last'), children[children.length - 1].ref)
    const visible = countVisibleChildren({ children })
    dict.set(PDFName.of('Count'), PDFNumber.of(open || !owner ? visible : -visible))
  }
}

/**
 * The document's outline items in the order pdf.js reports them (it walks
 * /First and /Next breadth first and skips an item it has already seen), with
 * the same 1-based paths the renderer received as bookmark ids.
 */
function readOutlineTree(pdfDoc, pdfLib) {
  const { PDFDict, PDFName, PDFRef } = pdfLib
  const context = pdfDoc.context
  const rootValue = pdfDoc.catalog.get(PDFName.of('Outlines'))
  const rootRef = rootValue instanceof PDFRef ? rootValue : null
  const rootDict = rootValue === undefined ? null : context.lookup(rootValue)
  const result = { rootRef, rootDict: rootDict instanceof PDFDict ? rootDict : null, nodes: [], byPath: new Map() }
  if (!result.rootDict) return result
  const first = result.rootDict.get(PDFName.of('First'))
  if (!(first instanceof PDFRef)) return result
  const top = { children: [], path: '', depth: -1 }
  const queue = [{ ref: first, parent: top }]
  const processed = new Set([first.tag])
  for (let cursor = 0; cursor < queue.length && result.nodes.length < MAX_OUTLINE_ITEMS; cursor += 1) {
    const { ref, parent } = queue[cursor]
    const dict = context.lookup(ref)
    if (!(dict instanceof PDFDict)) continue
    const node = { ref, dict, parent, children: [], depth: parent.depth + 1, path: '', title: decodeTitle(dict, pdfLib) }
    parent.children.push(node)
    node.path = parent === top ? String(parent.children.length) : `${parent.path}.${parent.children.length}`
    result.nodes.push(node)
    result.byPath.set(node.path, node)
    const child = dict.get(PDFName.of('First'))
    if (child instanceof PDFRef && !processed.has(child.tag)) {
      processed.add(child.tag)
      queue.push({ ref: child, parent: node })
    }
    const next = dict.get(PDFName.of('Next'))
    if (next instanceof PDFRef && !processed.has(next.tag)) {
      processed.add(next.tag)
      queue.push({ ref: next, parent })
    }
  }
  result.top = top
  return result
}

/** Pair each document bookmark with the outline item it was read from. */
function matchDocumentBookmarks(desired, existing, targets) {
  const unused = new Set(existing.nodes)
  const byTitle = new Map()
  for (const item of existing.nodes) {
    const key = normalizeTitle(item.title)
    if (!byTitle.has(key)) byTitle.set(key, [])
    byTitle.get(key).push(item)
  }
  const documentNodes = desired.filter((node) => node.fromDocument)
  const originalTitle = (node) => typeof node.entry.originalLabel === 'string' ? node.entry.originalLabel : node.title
  const claim = (node, item) => {
    node.existing = item
    unused.delete(item)
  }
  // Same place, same title: the usual case.
  for (const node of documentNodes) {
    const item = existing.byPath.get(String(node.entry.outlinePath ?? ''))
    if (item && unused.has(item) && sameTitle(item.title, originalTitle(node))) claim(node, item)
  }
  // Deleting a page removes bookmarks that pointed at it, which moves the
  // places of the items after them: find the same title (and page).
  for (const node of documentNodes) {
    if (node.existing) continue
    const candidates = (byTitle.get(normalizeTitle(originalTitle(node))) || []).filter((item) => unused.has(item))
    const item = candidates.find((candidate) => targets.pageOf(candidate.dict) === node.pageIndex) || candidates[0]
    if (item) claim(node, item)
  }
  // A title the two PDF readers decode differently: trust the place.
  for (const node of documentNodes) {
    if (node.existing) continue
    const item = existing.byPath.get(String(node.entry.outlinePath ?? ''))
    if (item && unused.has(item) && item.depth === node.depth) claim(node, item)
  }
}

/** True when the list describes exactly the outline the document already has. */
function isUnchanged(root, desired, existing, targets) {
  if (!desired.length && !existing.nodes.length) return true
  if (desired.length !== existing.nodes.length) return false
  if (desired.some((node) => !node.existing || sameTitle(node.title, node.existing.title) === false)) return false
  if (desired.some((node) => needsRetarget(node, targets))) return false
  const sameChildren = (wanted, actual) => wanted.length === actual.length
    && wanted.every((node, index) => node.existing === actual[index])
  if (!existing.top || !sameChildren(root.children, existing.top.children)) return false
  return desired.every((node) => sameChildren(node.children, node.existing.children))
}

function needsRetarget(node, targets) {
  if (!node.existing || node.pageIndex === null) return false
  const current = targets.pageOf(node.existing.dict)
  // Leave a destination alone when it cannot be resolved here; pdf.js may
  // know better, and rewriting it would only lose information.
  return current !== null && current !== node.pageIndex
}

/** Point a reused item at its bookmark's page, keeping the view (/XYZ …). */
function retargetIfMoved(node, targets, pageRefs, pdfLib) {
  if (!needsRetarget(node, targets)) return
  const { PDFArray, PDFName } = pdfLib
  const dict = node.dict
  const context = dict.context
  const resolved = targets.arrayOf(dict)
  const destination = PDFArray.withContext(context)
  destination.push(pageRefs[node.pageIndex])
  if (resolved && resolved.size() > 1) {
    for (let index = 1; index < resolved.size(); index += 1) destination.push(resolved.get(index))
  } else {
    destination.push(PDFName.of('Fit'))
  }
  dict.set(PDFName.of('Dest'), destination)
  // /Dest and a /GoTo action are mutually exclusive.
  if (dict.get(PDFName.of('A')) !== undefined && targets.isGoTo(dict)) dict.delete(PDFName.of('A'))
}

function newOutlineItem(node, context, pageRefs, pdfLib) {
  const { PDFArray, PDFDict, PDFHexString, PDFName, PDFNumber } = pdfLib
  const map = new Map()
  map.set(PDFName.of('Title'), PDFHexString.fromText(node.title))
  if (node.pageIndex !== null) {
    const destination = PDFArray.withContext(context)
    destination.push(pageRefs[node.pageIndex])
    destination.push(PDFName.of('Fit'))
    map.set(PDFName.of('Dest'), destination)
  } else if (node.url) {
    map.set(PDFName.of('A'), context.obj({ S: 'URI', URI: PDFHexString.fromText(node.url) }))
  }
  const flags = (node.italic ? 1 : 0) | (node.bold ? 2 : 0)
  if (flags) map.set(PDFName.of('F'), PDFNumber.of(flags))
  if (node.color && node.color.some((channel) => channel !== 0)) {
    const color = PDFArray.withContext(context)
    for (const channel of node.color) color.push(PDFNumber.of(channel))
    map.set(PDFName.of('C'), color)
  }
  return PDFDict.fromMapWithContext(map, context)
}

/**
 * Resolve an outline item's destination (/Dest, explicit or named, or a /GoTo
 * action) to a page index of `pageRefs`.
 */
function destinationResolver(pdfDoc, pdfLib, pageRefs) {
  const { PDFArray, PDFDict, PDFHexString, PDFName, PDFRef, PDFString } = pdfLib
  const context = pdfDoc.context
  const pageIndexByTag = new Map(pageRefs.map((ref, index) => [ref.tag, index]))
  let named = null
  const text = (value) => (value instanceof PDFString || value instanceof PDFHexString || value instanceof PDFName ? value.decodeText() : undefined)
  const namedDestination = (key) => {
    if (!named) {
      named = new Map()
      const dests = pdfDoc.catalog.lookupMaybe(PDFName.of('Dests'), PDFDict)
      if (dests) for (const [entryName, value] of dests.entries()) named.set(entryName.decodeText(), value)
      const tree = pdfDoc.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)?.lookupMaybe(PDFName.of('Dests'), PDFDict)
      const stack = tree ? [{ node: tree, depth: 0 }] : []
      const visited = new Set()
      while (stack.length) {
        const { node, depth } = stack.pop()
        if (!(node instanceof PDFDict) || visited.has(node) || depth > 64) continue
        visited.add(node)
        const pairs = node.lookupMaybe(PDFName.of('Names'), PDFArray)
        if (pairs) {
          for (let index = 0; index + 1 < pairs.size(); index += 2) {
            const key = text(context.lookup(pairs.get(index)))
            if (key !== undefined && !named.has(key)) named.set(key, pairs.get(index + 1))
          }
        }
        const kids = node.lookupMaybe(PDFName.of('Kids'), PDFArray)
        if (kids) for (let index = 0; index < kids.size(); index += 1) stack.push({ node: context.lookup(kids.get(index)), depth: depth + 1 })
      }
    }
    return named.get(key)
  }
  const explicit = (value, depth = 0) => {
    if (depth > 8 || value === undefined) return null
    const resolved = context.lookup(value)
    if (resolved instanceof PDFArray) return resolved
    if (resolved instanceof PDFDict) return explicit(resolved.get(PDFName.of('D')), depth + 1)
    const key = text(resolved)
    return key === undefined ? null : explicit(namedDestination(key), depth + 1)
  }
  const goTo = (dict) => {
    const action = context.lookup(dict.get(PDFName.of('A')))
    return action instanceof PDFDict && action.get(PDFName.of('S')) === PDFName.of('GoTo') ? action : null
  }
  const arrayOf = (dict) => {
    const destination = dict.get(PDFName.of('Dest'))
    if (destination !== undefined) return explicit(destination)
    const action = goTo(dict)
    return action ? explicit(action.get(PDFName.of('D'))) : null
  }
  const cache = new Map()
  return {
    arrayOf,
    isGoTo: (dict) => Boolean(goTo(dict)),
    pageOf(dict) {
      if (cache.has(dict)) return cache.get(dict)
      const target = arrayOf(dict)?.get(0)
      const index = target instanceof PDFRef && pageIndexByTag.has(target.tag) ? pageIndexByTag.get(target.tag) : null
      cache.set(dict, index)
      return index
    },
  }
}

/** The title as the renderer saw it: pdf.js decoding, trimmed, never empty. */
function decodeTitle(dict, pdfLib) {
  const { PDFHexString, PDFName, PDFString } = pdfLib
  const value = dict.lookup(PDFName.of('Title'))
  if (!(value instanceof PDFString || value instanceof PDFHexString)) return UNTITLED
  let decoded
  try {
    const bytes = value.asBytes()
    // PDF 2.0 allows UTF-8 text strings; pdf-lib reads them as PDFDocEncoding.
    decoded = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
      ? new TextDecoder('utf-8').decode(bytes.subarray(3))
      : value.decodeText()
  } catch {
    return UNTITLED
  }
  return decoded.trim() || UNTITLED
}

function normalizeTitle(title) {
  return String(title ?? '').normalize('NFC').replace(/\s+/g, ' ').trim() || UNTITLED
}

function sameTitle(left, right) {
  return normalizeTitle(left) === normalizeTitle(right)
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
