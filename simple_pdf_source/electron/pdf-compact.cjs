'use strict'

/**
 * Removing a page or a drawing from a PDF must remove its content from the
 * saved file. pdf-lib only unlinks a removed page from the page tree and
 * writes every object it parsed, so deleted pages, replaced content streams and
 * removed images stayed recoverable. This module:
 *  1. detaches removed pages from everything that still points at them
 *     (outline items, named destinations, links, form fields, the tag tree,
 *     the open action), so they become unreachable, and
 *  2. deletes every object that is no longer reachable from the trailer.
 * Pure pdf-lib (MIT); no new mupdf usage.
 */
const {
  PDFArray,
  PDFDict,
  PDFHexString,
  PDFName,
  PDFNull,
  PDFNumber,
  PDFObjectCopier,
  PDFPage,
  PDFPageLeaf,
  PDFRef,
  PDFStream,
  PDFString,
} = require('pdf-lib')

const name = (value) => PDFName.of(value)
const MAX_DEPTH = 512
// Arrays whose elements are independent members: a removed member is taken
// out. In all other arrays (destinations, number/name tree pairs) a removed
// reference becomes null so the positions of the remaining values hold.
const MEMBER_LIST_KEYS = new Set(['Kids', 'Annots', 'Fields', 'CO', 'K'])

function nameOf(value) {
  return value instanceof PDFName ? value.decodeText() : undefined
}

function textOf(value) {
  if (value instanceof PDFString || value instanceof PDFHexString) return value.decodeText()
  if (value instanceof PDFName) return value.decodeText()
  return undefined
}

/** Leaves of the page tree in document order, read from the tree itself. */
function pageLeaves(pdfDoc) {
  const { context } = pdfDoc
  const leaves = []
  const visited = new Set()
  const root = pdfDoc.catalog.get(name('Pages'))
  const stack = [root]
  while (stack.length) {
    const value = stack.pop()
    if (value instanceof PDFRef) {
      if (visited.has(value.tag)) continue
      visited.add(value.tag)
    }
    const node = context.lookup(value)
    if (!(node instanceof PDFDict)) continue
    const kids = node.lookupMaybe(name('Kids'), PDFArray)
    if (nameOf(node.get(name('Type'))) === 'Pages' || (kids && nameOf(node.get(name('Type'))) !== 'Page')) {
      if (kids) for (let index = kids.size() - 1; index >= 0; index -= 1) stack.push(kids.get(index))
    } else if (value instanceof PDFRef) {
      leaves.push({ ref: value, node })
    }
  }
  return leaves
}

/** name → destination, from the catalog /Dests dictionary and /Names /Dests tree. */
function namedDestinations(pdfDoc) {
  const { context } = pdfDoc
  let map = null
  return (key) => {
    if (!map) {
      map = new Map()
      const dests = pdfDoc.catalog.lookupMaybe(name('Dests'), PDFDict)
      if (dests) for (const [entryName, value] of dests.entries()) map.set(entryName.decodeText(), value)
      const names = pdfDoc.catalog.lookupMaybe(name('Names'), PDFDict)
      const tree = names?.lookupMaybe(name('Dests'), PDFDict)
      if (tree) forEachTreeEntry(context, tree, 'Names', (entryKey, value) => {
        const text = textOf(entryKey)
        if (text !== undefined && !map.has(text)) map.set(text, value)
      })
    }
    return map.get(key)
  }
}

function forEachTreeEntry(context, root, arrayKey, visit) {
  const visited = new Set()
  const stack = [{ node: root, depth: 0 }]
  while (stack.length) {
    const { node, depth } = stack.pop()
    if (!(node instanceof PDFDict) || visited.has(node) || depth > MAX_DEPTH) continue
    visited.add(node)
    const entries = node.lookupMaybe(name(arrayKey), PDFArray)
    if (entries) {
      for (let index = 0; index + 1 < entries.size(); index += 2) visit(context.lookup(entries.get(index)), entries.get(index + 1))
    }
    const kids = node.lookupMaybe(name('Kids'), PDFArray)
    if (kids) for (let index = 0; index < kids.size(); index += 1) stack.push({ node: context.lookup(kids.get(index)), depth: depth + 1 })
  }
}

/** The page reference a destination (explicit, named, or /D wrapper) points to. */
function destinationPage(context, value, resolveNamed, depth = 0) {
  if (depth > 8) return undefined
  const resolved = context.lookup(value)
  if (resolved instanceof PDFArray) {
    const target = resolved.get(0)
    return target instanceof PDFRef ? target : undefined
  }
  if (resolved instanceof PDFDict) return destinationPage(context, resolved.get(name('D')), resolveNamed, depth + 1)
  const named = textOf(resolved)
  if (named !== undefined) {
    const destination = resolveNamed(named)
    return destination === undefined ? undefined : destinationPage(context, destination, resolveNamed, depth + 1)
  }
  return undefined
}

function goToDestination(context, actionValue) {
  const action = context.lookup(actionValue)
  if (!(action instanceof PDFDict) || nameOf(action.get(name('S'))) !== 'GoTo') return undefined
  return action.get(name('D'))
}

function outlineChildren(context, parent) {
  const children = []
  const visited = new Set()
  let value = parent.get(name('First'))
  while (value instanceof PDFRef && !visited.has(value.tag) && children.length < 1_000_000) {
    visited.add(value.tag)
    const dict = context.lookup(value)
    if (!(dict instanceof PDFDict)) break
    children.push({ ref: value, dict })
    value = dict.get(name('Next'))
  }
  return children
}

function relinkOutline(parent, kept) {
  if (!kept.length) {
    parent.delete(name('First'))
    parent.delete(name('Last'))
    parent.delete(name('Count'))
    return
  }
  parent.set(name('First'), kept[0].ref)
  parent.set(name('Last'), kept[kept.length - 1].ref)
  kept.forEach((item, index) => {
    if (index) item.dict.set(name('Prev'), kept[index - 1].ref)
    else item.dict.delete(name('Prev'))
    if (index + 1 < kept.length) item.dict.set(name('Next'), kept[index + 1].ref)
    else item.dict.delete(name('Next'))
  })
}

function recountOutline(context, node, isRoot, depth = 0) {
  const children = outlineChildren(context, node)
  if (!children.length || depth > MAX_DEPTH) {
    if (!isRoot) node.delete(name('Count'))
    else node.set(name('Count'), PDFNumber.of(0))
    return 0
  }
  let visible = 0
  for (const child of children) {
    const count = child.dict.lookupMaybe(name('Count'), PDFNumber)
    const open = Boolean(count && count.asNumber() > 0)
    const inner = recountOutline(context, child.dict, false, depth + 1)
    visible += 1 + (open ? inner : 0)
  }
  const ownCount = node.lookupMaybe(name('Count'), PDFNumber)
  const open = isRoot || Boolean(ownCount && ownCount.asNumber() > 0)
  node.set(name('Count'), PDFNumber.of(open ? visible : -visible))
  return visible
}

/**
 * Outline items that point at a removed page: leaves are removed, items with
 * children lose their destination and stay as headings.
 */
function pruneOutlines(pdfDoc, pointsToRemoved) {
  const { context } = pdfDoc
  const root = pdfDoc.catalog.lookupMaybe(name('Outlines'), PDFDict)
  if (!root) return 0
  let removedItems = 0
  const dangling = (item) => {
    const destination = item.get(name('Dest'))
    if (destination !== undefined) return pointsToRemoved(destination)
    const action = goToDestination(context, item.get(name('A')))
    return action !== undefined && pointsToRemoved(action)
  }
  const prune = (parent, depth) => {
    if (depth > MAX_DEPTH) return
    const kept = []
    for (const child of outlineChildren(context, parent)) {
      prune(child.dict, depth + 1)
      if (dangling(child.dict)) {
        if (!child.dict.get(name('First'))) {
          removedItems += 1
          continue
        }
        child.dict.delete(name('Dest'))
        child.dict.delete(name('A'))
      }
      kept.push(child)
    }
    relinkOutline(parent, kept)
  }
  prune(root, 0)
  if (!removedItems) return 0
  if (!root.get(name('First'))) pdfDoc.catalog.delete(name('Outlines'))
  else recountOutline(context, root, true)
  return removedItems
}

function setTreeLimits(node, firstKey, lastKey) {
  if (!node.has(name('Limits'))) return
  if (firstKey === undefined) node.delete(name('Limits'))
  else node.set(name('Limits'), node.context.obj([firstKey, lastKey]))
}

/**
 * Filter a name tree ('Names') or number tree ('Nums'). `decide(key, value)`
 * returns false to drop the entry. Empty leaves are unlinked and /Limits of
 * intermediate nodes are recomputed.
 * @returns {{ first: any, last: any, empty: boolean }}
 */
function pruneTree(context, node, arrayKey, decide, depth = 0, visited = new Set()) {
  if (!(node instanceof PDFDict) || visited.has(node) || depth > MAX_DEPTH) return { first: undefined, last: undefined, empty: false }
  visited.add(node)
  let first
  let last
  const kids = node.lookupMaybe(name('Kids'), PDFArray)
  if (kids) {
    for (let index = 0; index < kids.size(); index += 1) {
      const child = context.lookup(kids.get(index))
      const result = pruneTree(context, child, arrayKey, decide, depth + 1, visited)
      if (result.empty) {
        kids.remove(index)
        index -= 1
        continue
      }
      if (first === undefined) first = result.first
      if (result.last !== undefined) last = result.last
    }
  }
  const entries = node.lookupMaybe(name(arrayKey), PDFArray)
  if (entries) {
    const kept = []
    for (let index = 0; index + 1 < entries.size(); index += 2) {
      const key = entries.get(index)
      const value = entries.get(index + 1)
      if (decide(context.lookup(key), value) !== false) kept.push(key, value)
    }
    if (kept.length !== entries.size()) {
      while (entries.size()) entries.remove(entries.size() - 1)
      for (const value of kept) entries.push(value)
    }
    if (kept.length) {
      if (first === undefined) first = kept[0]
      last = kept[kept.length - 2]
    }
  }
  setTreeLimits(node, first, last)
  const empty = (!kids || kids.size() === 0) && (!entries || entries.size() === 0) && depth > 0
  return { first, last, empty }
}

function pruneNamedDestinations(pdfDoc, pointsToRemoved) {
  const { context } = pdfDoc
  let removedEntries = 0
  const dests = pdfDoc.catalog.lookupMaybe(name('Dests'), PDFDict)
  if (dests) {
    for (const [key, value] of dests.entries()) {
      if (pointsToRemoved(value)) {
        dests.delete(key)
        removedEntries += 1
      }
    }
  }
  const names = pdfDoc.catalog.lookupMaybe(name('Names'), PDFDict)
  const tree = names?.lookupMaybe(name('Dests'), PDFDict)
  if (tree) {
    pruneTree(context, tree, 'Names', (_key, value) => {
      if (!pointsToRemoved(value)) return true
      removedEntries += 1
      return false
    })
  }
  return removedEntries
}

/** Links on the remaining pages that jump to a removed page. */
function pruneAnnotations(pdfDoc, leaves, pointsToRemoved, removedAnnotations) {
  const { context } = pdfDoc
  let removedCount = 0
  for (const { node } of leaves) {
    const annots = node.lookupMaybe(name('Annots'), PDFArray)
    if (!annots) continue
    for (let index = annots.size() - 1; index >= 0; index -= 1) {
      const value = annots.get(index)
      if (value instanceof PDFRef && removedAnnotations.has(value.tag)) {
        annots.remove(index)
        removedCount += 1
        continue
      }
      const annotation = context.lookup(value)
      if (!(annotation instanceof PDFDict)) continue
      const isLink = nameOf(annotation.get(name('Subtype'))) === 'Link'
      const destination = annotation.get(name('Dest'))
      const action = goToDestination(context, annotation.get(name('A')))
      const deadDestination = destination !== undefined && pointsToRemoved(destination)
      const deadAction = action !== undefined && pointsToRemoved(action)
      if ((deadDestination || deadAction) && isLink) {
        annots.remove(index)
        removedCount += 1
        continue
      }
      if (deadDestination) annotation.delete(name('Dest'))
      if (deadAction) annotation.delete(name('A'))
      const inReplyTo = annotation.get(name('IRT'))
      if (inReplyTo instanceof PDFRef && removedAnnotations.has(inReplyTo.tag)) annotation.delete(name('IRT'))
    }
  }
  return removedCount
}

/** Fields (and field kids) whose widgets lived only on removed pages. */
function pruneFormFields(pdfDoc, removedAnnotations) {
  const { context } = pdfDoc
  const acroForm = pdfDoc.catalog.lookupMaybe(name('AcroForm'), PDFDict)
  const fields = acroForm?.lookupMaybe(name('Fields'), PDFArray)
  if (!fields) return 0
  const removedFields = new Set()
  const visited = new Set()
  const filter = (array, depth) => {
    for (let index = array.size() - 1; index >= 0; index -= 1) {
      const value = array.get(index)
      if (value instanceof PDFRef) {
        if (removedAnnotations.has(value.tag)) {
          array.remove(index)
          removedFields.add(value.tag)
          continue
        }
        if (visited.has(value.tag)) continue
        visited.add(value.tag)
      }
      const field = context.lookup(value)
      const kids = field instanceof PDFDict ? field.lookupMaybe(name('Kids'), PDFArray) : undefined
      if (!kids || depth > MAX_DEPTH) continue
      const before = kids.size()
      filter(kids, depth + 1)
      if (before && !kids.size()) {
        array.remove(index)
        if (value instanceof PDFRef) removedFields.add(value.tag)
      }
    }
  }
  filter(fields, 0)
  const order = acroForm.lookupMaybe(name('CO'), PDFArray)
  if (order && removedFields.size) {
    for (let index = order.size() - 1; index >= 0; index -= 1) {
      const value = order.get(index)
      if (value instanceof PDFRef && removedFields.has(value.tag)) order.remove(index)
    }
  }
  return removedFields.size
}

/**
 * Tagged PDFs (Word, Acrobat) reference pages from structure elements (/Pg),
 * marked-content references, object references and the /ParentTree. Remove
 * the parts that describe removed pages so the tag tree stays valid and does
 * not keep those pages reachable.
 */
function pruneStructureTree(pdfDoc, removedPages, removedAnnotations) {
  const { context } = pdfDoc
  const root = pdfDoc.catalog.lookupMaybe(name('StructTreeRoot'), PDFDict)
  if (!root) return 0
  const removedPage = (value) => value instanceof PDFRef && removedPages.has(value.tag)
  const droppedElements = new Set()
  const visited = new Set()

  const keepKid = (item, page, depth) => {
    if (item instanceof PDFNumber) return !removedPage(page)
    if (item instanceof PDFRef && removedAnnotations.has(item.tag)) return false
    const dict = context.lookup(item)
    if (!(dict instanceof PDFDict)) return true
    if (item instanceof PDFRef) {
      if (visited.has(item.tag)) return !droppedElements.has(item.tag)
      visited.add(item.tag)
    }
    const type = nameOf(dict.get(name('Type')))
    const ownPage = dict.get(name('Pg'))
    const effectivePage = ownPage instanceof PDFRef ? ownPage : page
    if (type === 'MCR') return !removedPage(effectivePage)
    if (type === 'OBJR') {
      const object = dict.get(name('Obj'))
      if (object instanceof PDFRef && (removedAnnotations.has(object.tag) || removedPages.has(object.tag))) return false
      return !removedPage(effectivePage)
    }
    if (depth > MAX_DEPTH) return true
    const drop = pruneElement(dict, page, depth + 1)
    if (drop && item instanceof PDFRef) droppedElements.add(item.tag)
    return !drop
  }

  // Returns true when the element should be removed from its parent.
  const pruneElement = (element, inheritedPage, depth) => {
    const ownPage = element.get(name('Pg'))
    const page = ownPage instanceof PDFRef ? ownPage : inheritedPage
    const kidsValue = element.get(name('K'))
    if (kidsValue === undefined) return removedPage(ownPage)
    const kidsArray = context.lookup(kidsValue) instanceof PDFArray ? context.lookup(kidsValue) : null
    const items = kidsArray
      ? Array.from({ length: kidsArray.size() }, (_, index) => kidsArray.get(index))
      : [kidsValue]
    const keep = items.map((item) => keepKid(item, page, depth))
    if (keep.every(Boolean)) return false
    if (!keep.some(Boolean)) {
      element.delete(name('K'))
      return true
    }
    if (kidsArray) {
      for (let index = items.length - 1; index >= 0; index -= 1) if (!keep[index]) kidsArray.remove(index)
    }
    if (removedPage(ownPage)) element.delete(name('Pg'))
    return false
  }

  pruneElement(root, undefined, 0)

  const removedKeys = new Set()
  for (const tag of [...removedPages, ...removedAnnotations]) {
    const [objectNumber, generation] = tag.split(' ').map(Number)
    const object = context.lookup(PDFRef.of(objectNumber, generation))
    if (!(object instanceof PDFDict)) continue
    for (const key of ['StructParents', 'StructParent']) {
      const value = object.lookupMaybe(name(key), PDFNumber)
      if (value) removedKeys.add(value.asNumber())
    }
  }
  const dropped = (value) => value instanceof PDFRef && droppedElements.has(value.tag)
  const parentTree = root.lookupMaybe(name('ParentTree'), PDFDict)
  if (parentTree) {
    pruneTree(context, parentTree, 'Nums', (key, value) => {
      if (key instanceof PDFNumber && removedKeys.has(key.asNumber())) return false
      if (dropped(value)) return false
      const array = context.lookup(value)
      if (array instanceof PDFArray) {
        for (let index = 0; index < array.size(); index += 1) if (dropped(array.get(index))) array.set(index, PDFNull)
      }
      return true
    })
  }
  const idTree = root.lookupMaybe(name('IDTree'), PDFDict)
  if (idTree) pruneTree(context, idTree, 'Names', (_key, value) => !dropped(value))
  return droppedElements.size
}

function pruneOpenAction(pdfDoc, pointsToRemoved) {
  const openAction = pdfDoc.catalog.get(name('OpenAction'))
  if (openAction === undefined) return false
  const resolved = pdfDoc.context.lookup(openAction)
  const destination = resolved instanceof PDFArray ? openAction : goToDestination(pdfDoc.context, openAction)
  if (destination === undefined || !pointsToRemoved(destination)) return false
  pdfDoc.catalog.delete(name('OpenAction'))
  return true
}

/** Annotations listed on removed pages, or whose /P names a removed page. */
function annotationsOfRemovedPages(pdfDoc, removedRefs, removedPages) {
  const { context } = pdfDoc
  const annotations = new Set()
  for (const ref of removedRefs) {
    const page = context.lookup(ref)
    const annots = page instanceof PDFDict ? page.lookupMaybe(name('Annots'), PDFArray) : undefined
    if (!annots) continue
    for (let index = 0; index < annots.size(); index += 1) {
      const value = annots.get(index)
      if (value instanceof PDFRef) annotations.add(value.tag)
    }
  }
  for (const [ref, object] of context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFDict) || !object.has(name('Subtype'))) continue
    const page = object.get(name('P'))
    if (page instanceof PDFRef && removedPages.has(page.tag)) annotations.add(ref.tag)
  }
  return annotations
}

/**
 * Walk everything reachable from the trailer. References to removed objects
 * are cut on the way (removed from member lists, nulled elsewhere), so the
 * walk never enters a removed page. With `compact`, every object the walk did
 * not reach is deleted.
 */
function scrubAndCompact(pdfDoc, removedTags, { compact = true } = {}) {
  const { context } = pdfDoc
  const reached = new Set()
  const seenDirect = new Set()
  const stack = []
  const { Root, Info, Encrypt, ID } = context.trailerInfo
  for (const value of [Root, Info, Encrypt, ID]) if (value) stack.push({ value, key: undefined })
  let cut = 0

  const visitDict = (dict) => {
    for (const [key, child] of dict.entries()) {
      if (child instanceof PDFRef && removedTags.has(child.tag)) {
        dict.delete(key)
        cut += 1
        continue
      }
      stack.push({ value: child, key: key.decodeText() })
    }
  }
  const visitArray = (array, key) => {
    const memberList = MEMBER_LIST_KEYS.has(key)
    for (let index = array.size() - 1; index >= 0; index -= 1) {
      const child = array.get(index)
      if (child instanceof PDFRef && removedTags.has(child.tag)) {
        if (memberList) array.remove(index)
        else array.set(index, PDFNull)
        cut += 1
        continue
      }
      stack.push({ value: child, key: undefined })
    }
  }

  while (stack.length) {
    const { value, key } = stack.pop()
    if (value instanceof PDFRef) {
      if (reached.has(value.tag)) continue
      reached.add(value.tag)
      const object = context.lookup(value)
      if (object !== undefined) stack.push({ value: object, key })
      continue
    }
    if (!(value instanceof PDFDict) && !(value instanceof PDFArray) && !(value instanceof PDFStream)) continue
    if (seenDirect.has(value)) continue
    seenDirect.add(value)
    if (value instanceof PDFDict) visitDict(value)
    else if (value instanceof PDFArray) visitArray(value, key)
    else visitDict(value.dict)
  }

  let deleted = 0
  if (compact) {
    for (const [ref] of context.enumerateIndirectObjects()) {
      if (reached.has(ref.tag)) continue
      context.delete(ref)
      deleted += 1
    }
  }
  return { deleted, cut }
}

/** Delete every object that cannot be reached from the trailer. */
function compactUnreachable(pdfDoc) {
  return scrubAndCompact(pdfDoc, new Set(), { compact: true })
}

/**
 * Make removed pages unreachable and (by default) drop every unreachable
 * object. `removedRefs` are page references that are no longer in the page
 * tree: pages removed with removePage(), or placeholders standing in for
 * pages that were deliberately not copied (see copyPagesRemapped).
 */
function detachRemovedPages(pdfDoc, removedRefs, { compact = true } = {}) {
  const refs = (removedRefs || []).filter((ref) => ref instanceof PDFRef)
  const removedPages = new Set(refs.map((ref) => ref.tag))
  if (!removedPages.size) return { ...(compact ? compactUnreachable(pdfDoc) : { deleted: 0, cut: 0 }), outlineItems: 0, links: 0, fields: 0 }
  const { context } = pdfDoc
  const resolveNamed = namedDestinations(pdfDoc)
  const pointsToRemoved = (destination) => {
    const page = destinationPage(context, destination, resolveNamed)
    return page instanceof PDFRef && removedPages.has(page.tag)
  }
  const leaves = pageLeaves(pdfDoc).filter(({ ref }) => !removedPages.has(ref.tag))
  const removedAnnotations = annotationsOfRemovedPages(pdfDoc, refs, removedPages)
  const outlineItems = pruneOutlines(pdfDoc, pointsToRemoved)
  pruneNamedDestinations(pdfDoc, pointsToRemoved)
  const links = pruneAnnotations(pdfDoc, leaves, pointsToRemoved, removedAnnotations)
  const fields = pruneFormFields(pdfDoc, removedAnnotations)
  pruneStructureTree(pdfDoc, removedPages, removedAnnotations)
  pruneOpenAction(pdfDoc, pointsToRemoved)
  const result = scrubAndCompact(pdfDoc, new Set([...removedPages, ...removedAnnotations]), { compact })
  return { ...result, outlineItems, links, fields }
}

/**
 * Copy pages between documents without dragging other pages along.
 * pdf-lib's copyPages() deep-copies everything a page references; a link to
 * another page (or an annotation's /P) therefore copied that page's content as
 * a hidden orphan, and links between copied pages pointed at those orphans.
 * Here references to selected pages resolve to their copies, and references to
 * every other source page resolve to `excludedRefs` placeholders that
 * detachRemovedPages() then cuts.
 */
function copyPagesRemapped(destDoc, srcDoc, indices) {
  const copier = PDFObjectCopier.for(srcDoc.context, destDoc.context)
  const leaves = pageLeaves(srcDoc)
  const reserved = new Map()
  const selected = []
  for (const index of indices) {
    const leaf = leaves[index]
    if (!leaf) throw new RangeError(`Page ${index + 1} does not exist.`)
    let destRef = reserved.get(leaf.ref.tag)
    const duplicate = Boolean(destRef)
    if (!destRef) {
      destRef = destDoc.context.nextRef()
      reserved.set(leaf.ref.tag, destRef)
      copier.traversedObjects.set(leaf.ref, destRef)
    }
    selected.push({ leaf, destRef: duplicate ? destDoc.context.nextRef() : destRef })
  }
  const excludedRefs = []
  for (const leaf of leaves) {
    if (reserved.has(leaf.ref.tag)) continue
    const placeholder = destDoc.context.nextRef()
    copier.traversedObjects.set(leaf.ref, placeholder)
    excludedRefs.push(placeholder)
  }
  const pages = selected.map(({ leaf, destRef }) => {
    if (!(leaf.node instanceof PDFPageLeaf)) throw new Error('A page of the source document is malformed.')
    const copied = copier.copy(leaf.node)
    destDoc.context.assign(destRef, copied)
    return PDFPage.of(copied, destRef, destDoc)
  })
  return { pages, excludedRefs }
}

module.exports = {
  compactUnreachable,
  copyPagesRemapped,
  detachRemovedPages,
  pageLeaves,
}
