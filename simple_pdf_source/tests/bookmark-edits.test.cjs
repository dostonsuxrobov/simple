'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const test = require('node:test')
const ts = require('typescript')
const { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString } = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')

// src/lib/bookmarks.ts is the renderer's bookmark list logic (no DOM);
// transpile it on the fly like tests/viewer-clipboard.test.cjs does.
const bookmarks = (async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-bookmark-edits-'))
  const sourcePath = path.resolve(__dirname, '..', 'src', 'lib', 'bookmarks.ts')
  const compiled = ts.transpileModule(await fs.readFile(sourcePath, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 },
    fileName: sourcePath,
    reportDiagnostics: true,
  })
  const errors = (compiled.diagnostics || []).filter((item) => item.category === ts.DiagnosticCategory.Error)
  assert.equal(errors.length, 0, errors.map((item) => item.messageText).join('\n'))
  const target = path.join(directory, 'bookmarks.mjs')
  await fs.writeFile(target, compiled.outputText)
  process.once('exit', () => { try { require('node:fs').rmSync(directory, { recursive: true, force: true }) } catch { /* best effort */ } })
  return import(pathToFileURL(target).href)
})()

const entry = (id, label, depth, pageIndex, extra = {}) => ({ id, label, depth, pageIndex, source: 'document', ...extra })

/**
 * Five pages; the outline the bookmark bugs used to destroy:
 *   1   Chapter 1      /XYZ 72 700 2 (bold)
 *   2   Website        URI action
 *   3   Part II        no destination, closed
 *   3.1   Chapter 2    /XYZ 0 500 null, page 2
 *   3.2   Chapter 3    /FitH 300, page 3
 *   4   Notes          GoTo action to page 5
 */
async function outlineFixture() {
  const doc = await PDFDocument.create()
  const pages = Array.from({ length: 5 }, () => doc.addPage([400, 600]))
  const { context } = doc
  const [chapter1, website, part2, chapter2, chapter3, notes] = Array.from({ length: 6 }, () => context.nextRef())
  const root = context.nextRef()
  const item = (title, extra) => context.obj({ Title: PDFHexString.fromText(title), ...extra })
  const dest = (...parts) => context.obj(parts)
  context.assign(chapter1, item('Chapter 1', { Parent: root, Next: website, F: 2, Dest: dest(pages[0].ref, PDFName.of('XYZ'), 72, 700, 2) }))
  context.assign(website, item('Website', { Parent: root, Prev: chapter1, Next: part2, A: { S: 'URI', URI: PDFString.of('https://example.com/guide') } }))
  context.assign(part2, item('Part II', { Parent: root, Prev: website, Next: notes, First: chapter2, Last: chapter3, Count: -2 }))
  context.assign(chapter2, item('Chapter 2', { Parent: part2, Next: chapter3, Dest: dest(pages[1].ref, PDFName.of('XYZ'), 0, 500, null) }))
  context.assign(chapter3, item('Chapter 3', { Parent: part2, Prev: chapter2, Dest: dest(pages[2].ref, PDFName.of('FitH'), 300) }))
  context.assign(notes, item('Notes', { Parent: root, Prev: part2, A: { S: 'GoTo', D: dest(pages[4].ref, PDFName.of('Fit')) } }))
  context.assign(root, context.obj({ Type: 'Outlines', First: chapter1, Last: notes, Count: 4 }))
  doc.catalog.set(PDFName.of('Outlines'), root)
  return doc.save({ useObjectStreams: false })
}

/** The outline as src/lib/pdf.ts getPdfOutlineBookmarks() reports it. */
async function pdfjsOutline(bytes) {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const document = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise
  try {
    const items = []
    const visit = async (nodes, depth, prefix) => {
      for (const [index, node] of nodes.entries()) {
        const itemPath = [...prefix, index + 1]
        let destination = node.dest
        if (typeof destination === 'string') destination = await document.getDestination(destination)
        let pageIndex = null
        if (Array.isArray(destination) && destination[0] && typeof destination[0] === 'object') {
          pageIndex = await document.getPageIndex(destination[0]).catch(() => null)
        }
        const children = node.items || []
        items.push({
          id: `pdf-outline:${itemPath.join('.')}`,
          title: typeof node.title === 'string' && node.title.trim() ? node.title.trim() : 'Untitled bookmark',
          pageIndex,
          pageNumber: pageIndex === null ? null : pageIndex + 1,
          depth,
          hasChildren: children.length > 0,
          expanded: typeof node.count !== 'number' || node.count >= 0,
          bold: node.bold === true,
          italic: node.italic === true,
          color: [0, 0, 0],
          url: typeof node.url === 'string' ? node.url : null,
        })
        await visit(children, depth + 1, itemPath)
      }
    }
    await visit((await document.getOutline()) || [], 0, [])
    return items
  } finally {
    await document.destroy()
  }
}

/** The saved outline: title, depth and what it points at. */
async function savedOutline(bytes) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const { context } = doc
  const pageNumber = new Map(doc.getPages().map((page, index) => [page.ref.tag, index + 1]))
  const describe = (value) => {
    const resolved = context.lookup(value)
    if (!(resolved instanceof PDFArray)) return resolved === undefined ? null : String(resolved)
    return resolved.asArray().map((part, index) => {
      if (index === 0 && part instanceof PDFRef) return pageNumber.has(part.tag) ? `page${pageNumber.get(part.tag)}` : 'missing'
      if (part instanceof PDFName) return part.decodeText()
      if (part instanceof PDFNumber) return part.asNumber()
      return String(part)
    }).join(' ')
  }
  const items = []
  const walk = (parent, depth) => {
    let value = parent.get(PDFName.of('First'))
    while (value instanceof PDFRef) {
      const dict = context.lookup(value, PDFDict)
      const action = dict.lookupMaybe(PDFName.of('A'), PDFDict)
      items.push({
        title: dict.lookup(PDFName.of('Title')).decodeText(),
        depth,
        target: dict.get(PDFName.of('Dest')) !== undefined
          ? describe(dict.get(PDFName.of('Dest')))
          : action
            ? `${action.get(PDFName.of('S')).decodeText()}:${action.get(PDFName.of('URI')) ? action.lookup(PDFName.of('URI')).decodeText() : describe(action.get(PDFName.of('D')))}`
            : null,
      })
      walk(dict, depth + 1)
      value = dict.get(PDFName.of('Next'))
    }
  }
  const root = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)
  if (root) walk(root, 0)
  return items
}

async function saveWithBookmarks(bytes, list) {
  const { invoke } = loadMain()
  const result = await invoke('pdf:flatten-overlays', bytes, [], {}, { bookmarks: list, report: true })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  return result.data
}

test('nesting follows depth, deeper jumps are clamped like the writer does, and flattening keeps the order', async () => {
  const { bookmarkTree, flattenBookmarkTree } = await bookmarks
  const list = [entry('a', 'A', 0, 0), entry('b', 'B', 3, 1), entry('c', 'C', 1, 2), entry('d', 'D', 0, null)]
  const tree = bookmarkTree(list)
  assert.deepEqual(tree.map((node) => [node.bookmark.id, node.children.map((child) => child.bookmark.id)]), [['a', ['b', 'c']], ['d', []]])
  assert.deepEqual(flattenBookmarkTree(tree).map((item) => [item.id, item.depth]), [['a', 0], ['b', 1], ['c', 1], ['d', 0]])
})

test('removing a bookmark removes the bookmarks nested under it and nothing else', async () => {
  const { bookmarkSubtreeSize, withoutBookmark } = await bookmarks
  const list = [
    entry('1', 'Chapter 1', 0, 0), entry('2', 'Part II', 0, null), entry('2.1', 'Chapter 2', 1, 1),
    entry('2.1.1', 'Section', 2, 1), entry('2.2', 'Chapter 3', 1, 2), entry('3', 'Notes', 0, 4),
  ]
  assert.equal(bookmarkSubtreeSize(list, '2'), 4)
  assert.equal(bookmarkSubtreeSize(list, '2.1'), 2)
  assert.equal(bookmarkSubtreeSize(list, 'missing'), 0)
  assert.deepEqual(withoutBookmark(list, '2').map((item) => item.id), ['1', '3'])
  assert.deepEqual(withoutBookmark(list, '2.1').map((item) => item.id), ['1', '2', '2.2', '3'])
  assert.deepEqual(withoutBookmark(list, '3').map((item) => item.id), ['1', '2', '2.1', '2.1.1', '2.2'])
})

test('after pages are deleted, bookmarks follow the way the outline is pruned', async () => {
  const { bookmarksAfterPageDelete } = await bookmarks
  const list = [
    entry('1', 'Chapter 1', 0, 0),
    entry('2', 'Part II', 0, 1), // on a deleted page, but has children: becomes a heading
    entry('2.1', 'Chapter 2', 1, 1), // on a deleted page: removed
    entry('2.2', 'Chapter 3', 1, 3),
    entry('3', 'Website', 0, null, { url: 'https://example.com' }),
    entry('4', 'Appendix', 0, 2), // on a deleted page with nothing below: removed
    entry('5', 'Notes', 0, 4),
  ]
  const after = bookmarksAfterPageDelete(list, [1, 2])
  assert.deepEqual(after.map((item) => [item.id, item.pageIndex, item.depth]), [
    ['1', 0, 0], ['2', null, 0], ['2.2', 1, 1], ['3', null, 0], ['5', 2, 0],
  ])
  assert.equal(after.find((item) => item.id === '3').url, 'https://example.com')
})

test('page moves keep bookmarks without a page as they are', async () => {
  const { remapBookmarkPages } = await bookmarks
  const list = [entry('a', 'A', 0, 0), entry('b', 'Heading', 0, null), entry('c', 'C', 1, 2)]
  const moved = remapBookmarkPages(list, (index) => [2, 0, 1][index])
  assert.deepEqual(moved.map((item) => item.pageIndex), [2, null, 1])
  assert.equal(moved[1], list[1], 'an item without a page is not copied')
})

test('the toolbar button only toggles its own bookmarks, never a chapter of the document', async () => {
  const { toolbarBookmarkFor, bookmarkTargetLabel } = await bookmarks
  const list = [entry('chapter', 'Chapter 1', 0, 3), { id: 'mine', label: 'Page 4', depth: 0, pageIndex: 3, source: 'simple' }]
  assert.equal(toolbarBookmarkFor(list, 3).id, 'mine')
  assert.equal(toolbarBookmarkFor(list.slice(0, 1), 3), undefined)
  assert.equal(bookmarkTargetLabel(list[0]), 'Page 4')
  assert.equal(bookmarkTargetLabel(entry('w', 'Site', 0, null, { url: 'https://example.com' })), 'Web link')
  assert.equal(bookmarkTargetLabel(entry('h', 'Part', 0, null)), 'Heading')
})

test('editing one bookmark in the app keeps every other outline item exactly (rename, add, remove a subtree)', async () => {
  const { bookmarksFromOutline, withoutBookmark } = await bookmarks
  const bytes = await outlineFixture()
  const list = bookmarksFromOutline(await pdfjsOutline(bytes))
  assert.deepEqual(list.map((item) => [item.label, item.depth, item.pageIndex, item.outlinePath]), [
    ['Chapter 1', 0, 0, '1'], ['Website', 0, null, '2'], ['Part II', 0, null, '3'],
    ['Chapter 2', 1, 1, '3.1'], ['Chapter 3', 1, 2, '3.2'], ['Notes', 0, 4, '4'],
  ])
  assert.equal(list[1].url, 'https://example.com/guide')

  // Rename one chapter and add a toolbar bookmark on page 4.
  const renamed = list.map((item) => item.label === 'Chapter 2' ? { ...item, label: 'Chapter Two' } : item)
  const edited = [...renamed, { id: 'bookmark-new', pageIndex: 3, label: 'Page 4', depth: 0, source: 'simple' }]
  assert.deepEqual(await savedOutline(await saveWithBookmarks(bytes, edited)), [
    { title: 'Chapter 1', depth: 0, target: 'page1 XYZ 72 700 2' },
    { title: 'Website', depth: 0, target: 'URI:https://example.com/guide' },
    { title: 'Part II', depth: 0, target: null },
    { title: 'Chapter Two', depth: 1, target: 'page2 XYZ 0 500 null' },
    { title: 'Chapter 3', depth: 1, target: 'page3 FitH 300' },
    { title: 'Notes', depth: 0, target: 'GoTo:page5 Fit' },
    { title: 'Page 4', depth: 0, target: 'page4 Fit' },
  ])

  // Removing the heading removes its chapters with it, and only them.
  const removed = withoutBookmark(list, list.find((item) => item.label === 'Part II').id)
  assert.deepEqual((await savedOutline(await saveWithBookmarks(bytes, removed))).map((item) => [item.title, item.target]), [
    ['Chapter 1', 'page1 XYZ 72 700 2'], ['Website', 'URI:https://example.com/guide'], ['Notes', 'GoTo:page5 Fit'],
  ])
})

test('bookmarks edited after a page delete save without dangling destinations', async () => {
  const { bookmarksAfterPageDelete, bookmarksFromOutline } = await bookmarks
  const { invoke } = loadMain()
  const bytes = await outlineFixture()
  const list = bookmarksFromOutline(await pdfjsOutline(bytes))
  // Delete page 2 (Chapter 2's page), as the Pages panel does.
  const mutated = await invoke('pdf:mutate', bytes, { type: 'delete', indices: [1] })
  const remaining = bookmarksAfterPageDelete(list, [1])
  assert.deepEqual(remaining.map((item) => [item.label, item.pageIndex]), [
    ['Chapter 1', 0], ['Website', null], ['Part II', null], ['Chapter 3', 1], ['Notes', 3],
  ])
  const edited = remaining.map((item) => item.label === 'Notes' ? { ...item, label: 'Notes (renamed)' } : item)
  const saved = await savedOutline(await saveWithBookmarks(mutated, edited))
  assert.deepEqual(saved.map((item) => [item.title, item.depth, item.target]), [
    ['Chapter 1', 0, 'page1 XYZ 72 700 2'],
    ['Website', 0, 'URI:https://example.com/guide'],
    ['Part II', 0, null],
    ['Chapter 3', 1, 'page2 FitH 300'],
    ['Notes (renamed)', 0, 'GoTo:page4 Fit'],
  ])
  assert.ok(saved.every((item) => !String(item.target).includes('missing')))
})
