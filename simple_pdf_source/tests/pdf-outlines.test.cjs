const assert = require('node:assert/strict')
const test = require('node:test')
const pdfLib = require('pdf-lib')
const { replacePdfOutlines } = require('../electron/pdf-outlines.cjs')
const { loadMain } = require('./helpers/electron-harness.cjs')

const { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString } = pdfLib
const REPORT = { report: true }

/**
 * Five pages and an outline with everything a rebuild used to lose:
 *   1   Chapter 1      /XYZ 72 700 2, bold, red
 *   2   Website        URI action
 *   3   Part II        no destination, closed
 *   3.1   Chapter 2    /XYZ 0 500 null
 *   3.2   Chapter 3    /FitH 300
 *   4   Appendix       named destination "appendix" (/XYZ 50 400 1.5)
 *   5   Notes          GoTo action to page 5
 */
async function outlineFixture() {
  const doc = await PDFDocument.create()
  const pages = Array.from({ length: 5 }, () => doc.addPage([400, 600]))
  const { context } = doc
  const refs = Array.from({ length: 7 }, () => context.nextRef())
  const [chapter1, website, part2, chapter2, chapter3, appendix, notes] = refs
  const outlinesRef = context.nextRef()
  const item = (title, extra) => {
    const dict = context.obj({ Title: PDFHexString.fromText(title), ...extra })
    return dict
  }
  const dest = (...parts) => context.obj(parts)
  context.assign(chapter1, item('Chapter 1', {
    Parent: outlinesRef, Next: website,
    Dest: dest(pages[0].ref, PDFName.of('XYZ'), 72, 700, 2), F: 2, C: [1, 0, 0],
  }))
  context.assign(website, item('Website', {
    Parent: outlinesRef, Prev: chapter1, Next: part2,
    A: { S: 'URI', URI: PDFString.of('https://example.com/guide') },
  }))
  context.assign(part2, item('Part II', {
    Parent: outlinesRef, Prev: website, Next: appendix, First: chapter2, Last: chapter3, Count: -2,
  }))
  context.assign(chapter2, item('Chapter 2', {
    Parent: part2, Next: chapter3,
    Dest: dest(pages[1].ref, PDFName.of('XYZ'), 0, 500, null),
  }))
  context.assign(chapter3, item('Chapter 3', {
    Parent: part2, Prev: chapter2,
    Dest: dest(pages[2].ref, PDFName.of('FitH'), 300),
  }))
  context.assign(appendix, item('Appendix', {
    Parent: outlinesRef, Prev: part2, Next: notes, Dest: PDFString.of('appendix'),
  }))
  context.assign(notes, item('Notes', {
    Parent: outlinesRef, Prev: appendix,
    A: { S: 'GoTo', D: dest(pages[4].ref, PDFName.of('Fit')) },
  }))
  context.assign(outlinesRef, context.obj({ Type: 'Outlines', First: chapter1, Last: notes, Count: 5 }))
  doc.catalog.set(PDFName.of('Outlines'), outlinesRef)
  const names = context.obj({ Names: [PDFString.of('appendix'), dest(pages[3].ref, PDFName.of('XYZ'), 50, 400, 1.5)] })
  doc.catalog.set(PDFName.of('Names'), context.obj({ Dests: names }))
  return doc.save({ useObjectStreams: false })
}

async function pdfjs() {
  return import('pdfjs-dist/legacy/build/pdf.mjs')
}

/** The bookmark list the renderer builds from pdf.js (see src/lib/pdf.ts). */
async function rendererEntries(bytes) {
  const { getDocument } = await pdfjs()
  const document = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise
  try {
    const entries = []
    const visit = async (items, depth, path) => {
      for (const [index, node] of items.entries()) {
        const itemPath = [...path, index + 1]
        let destination = node.dest
        if (typeof destination === 'string') destination = await document.getDestination(destination)
        let pageIndex = null
        if (Array.isArray(destination) && destination[0] && typeof destination[0] === 'object') {
          pageIndex = await document.getPageIndex(destination[0]).catch(() => null)
        }
        const title = typeof node.title === 'string' && node.title.trim() ? node.title.trim() : 'Untitled bookmark'
        entries.push({
          id: `pdf-outline:${itemPath.join('.')}`,
          pageIndex,
          label: title,
          depth,
          source: 'document',
          expanded: typeof node.count !== 'number' || node.count >= 0,
          bold: node.bold === true,
          italic: node.italic === true,
          color: Array.from(node.color || [0, 0, 0]),
          url: typeof node.url === 'string' ? node.url : null,
          outlinePath: itemPath.join('.'),
          originalLabel: title,
          originalPageIndex: pageIndex,
        })
        await visit(node.items || [], depth + 1, itemPath)
      }
    }
    await visit((await document.getOutline()) || [], 0, [])
    return entries
  } finally {
    await document.destroy()
  }
}

/** The saved outline read back with pdf-lib, destinations by page number. */
async function savedOutline(bytes) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const { context } = doc
  const pageIndex = new Map(doc.getPages().map((page, index) => [page.ref.tag, index]))
  const describe = (value) => {
    const resolved = context.lookup(value)
    if (resolved instanceof PDFArray) {
      return resolved.asArray().map((part, index) => {
        if (index === 0 && part instanceof PDFRef) return pageIndex.has(part.tag) ? `page${pageIndex.get(part.tag) + 1}` : `missing:${part.tag}`
        if (part instanceof PDFName) return part.decodeText()
        if (part instanceof PDFNumber) return part.asNumber()
        return String(part)
      }).join(' ')
    }
    if (resolved instanceof PDFString || resolved instanceof PDFHexString) return `named:${resolved.decodeText()}`
    return resolved === undefined ? null : String(resolved)
  }
  const items = []
  const walk = (parent, depth) => {
    let value = parent.get(PDFName.of('First'))
    while (value instanceof PDFRef) {
      const dict = context.lookup(value, PDFDict)
      const action = dict.lookupMaybe(PDFName.of('A'), PDFDict)
      const count = dict.lookupMaybe(PDFName.of('Count'), PDFNumber)
      items.push({
        title: dict.lookup(PDFName.of('Title')).decodeText(),
        depth,
        dest: describe(dict.get(PDFName.of('Dest'))),
        action: action ? action.get(PDFName.of('S')).decodeText() : null,
        uri: action?.lookup(PDFName.of('URI'))?.decodeText?.() ?? null,
        goTo: action ? describe(action.get(PDFName.of('D'))) : null,
        count: count ? count.asNumber() : null,
        flags: dict.lookupMaybe(PDFName.of('F'), PDFNumber)?.asNumber() ?? null,
        parentOk: context.lookup(dict.get(PDFName.of('Parent'))) === parent,
      })
      walk(dict, depth + 1)
      value = dict.get(PDFName.of('Next'))
    }
  }
  const root = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)
  if (root) walk(root, 0)
  return { items, rootCount: root?.lookupMaybe(PDFName.of('Count'), PDFNumber)?.asNumber() ?? null, pageCount: doc.getPageCount() }
}

async function flattenWithBookmarks(bytes, bookmarks) {
  const { invoke } = loadMain()
  const result = await invoke('pdf:flatten-overlays', bytes, [], {}, { bookmarks, ...REPORT })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  return result.data
}

test('an unchanged bookmark list leaves every outline object exactly as it was', async () => {
  const bytes = await outlineFixture()
  const entries = await rendererEntries(bytes)
  assert.equal(entries.length, 7)
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const before = new Map(doc.context.enumerateIndirectObjects().map(([ref, object]) => [ref.tag, object.toString()]))
  assert.deepEqual(replacePdfOutlines(doc, entries, pdfLib), { written: 7, skipped: 0 })
  const after = new Map(doc.context.enumerateIndirectObjects().map(([ref, object]) => [ref.tag, object.toString()]))
  assert.deepEqual([...after.entries()], [...before.entries()])
})

test('adding one bookmark keeps headings, links, nesting and exact destinations', async () => {
  const bytes = await outlineFixture()
  const entries = await rendererEntries(bytes)
  const output = await flattenWithBookmarks(bytes, [...entries, { id: 'bookmark-new', pageIndex: 4, label: 'Page 5', depth: 0, source: 'simple' }])
  const { items, rootCount } = await savedOutline(output)
  assert.deepEqual(items.map(({ title, depth }) => [title, depth]), [
    ['Chapter 1', 0], ['Website', 0], ['Part II', 0], ['Chapter 2', 1], ['Chapter 3', 1], ['Appendix', 0], ['Notes', 0], ['Page 5', 0],
  ])
  const byTitle = Object.fromEntries(items.map((item) => [item.title, item]))
  assert.equal(byTitle['Chapter 1'].dest, 'page1 XYZ 72 700 2')
  assert.equal(byTitle['Chapter 1'].flags, 2)
  assert.equal(byTitle.Website.action, 'URI')
  assert.equal(byTitle.Website.uri, 'https://example.com/guide')
  assert.equal(byTitle['Part II'].dest, null)
  assert.equal(byTitle['Part II'].count, -2, 'a closed heading stays closed')
  assert.equal(byTitle['Chapter 2'].dest, 'page2 XYZ 0 500 null')
  assert.equal(byTitle['Chapter 3'].dest, 'page3 FitH 300')
  assert.equal(byTitle.Appendix.dest, 'named:appendix')
  assert.equal(byTitle.Notes.action, 'GoTo')
  assert.equal(byTitle.Notes.goTo, 'page5 Fit')
  assert.equal(byTitle['Page 5'].dest, 'page5 Fit')
  assert.ok(items.every((item) => item.parentOk), 'every item points at its parent')
  // Part II is closed, so its two children are not visible: 6 visible items.
  assert.equal(rootCount, 6)
  const reread = await rendererEntries(output)
  assert.deepEqual(reread.map((entry) => entry.pageIndex), [0, null, null, 1, 2, 3, 4, 4])
})

test('renaming changes only that title', async () => {
  const bytes = await outlineFixture()
  const entries = await rendererEntries(bytes)
  const renamed = entries.map((entry) => entry.label === 'Chapter 2' ? { ...entry, label: 'Chapter Two — Ünïcode' } : entry)
  const { items } = await savedOutline(await flattenWithBookmarks(bytes, renamed))
  const original = (await savedOutline(bytes)).items
  assert.deepEqual(items.map((item) => item.title), original.map((item) => item.title === 'Chapter 2' ? 'Chapter Two — Ünïcode' : item.title))
  assert.deepEqual(items.map(({ dest, action, count }) => [dest, action, count]), original.map(({ dest, action, count }) => [dest, action, count]))
})

test('deleting a heading removes its subtree and nothing else', async () => {
  const bytes = await outlineFixture()
  const entries = await rendererEntries(bytes)
  const kept = entries.filter((entry) => !entry.outlinePath.startsWith('3'))
  const output = await flattenWithBookmarks(bytes, kept)
  const { items, rootCount } = await savedOutline(output)
  assert.deepEqual(items.map((item) => item.title), ['Chapter 1', 'Website', 'Appendix', 'Notes'])
  assert.equal(items[0].dest, 'page1 XYZ 72 700 2')
  assert.equal(rootCount, 4)
  const doc = await PDFDocument.load(output, { updateMetadata: false })
  const titles = doc.context.enumerateIndirectObjects()
    .filter(([, object]) => object instanceof PDFDict && object.has(PDFName.of('Title')))
    .map(([, object]) => object.lookup(PDFName.of('Title')).decodeText())
  assert.ok(!titles.includes('Chapter 2') && !titles.includes('Part II'), 'deleted outline items are not left behind in the file')
})

test('removing every bookmark removes the outline', async () => {
  const bytes = await outlineFixture()
  const output = await flattenWithBookmarks(bytes, [])
  const { items } = await savedOutline(output)
  assert.equal(items.length, 0)
  const doc = await PDFDocument.load(output, { updateMetadata: false })
  assert.equal(doc.catalog.get(PDFName.of('Outlines')), undefined)
})

test('after a page is deleted, bookmark edits still keep the remaining destinations', async () => {
  const { invoke } = loadMain()
  const bytes = await outlineFixture()
  const entries = await rendererEntries(bytes)
  // Delete page 2 (Chapter 2's page). The main process prunes that bookmark
  // from the base document; the renderer mirrors it: a bookmark on a deleted
  // page goes when it has no children, other page numbers shift down.
  const mutated = await invoke('pdf:mutate', bytes, { type: 'delete', indices: [1] })
  const remaining = entries
    .filter((entry) => entry.pageIndex !== 1)
    .map((entry) => ({ ...entry, pageIndex: entry.pageIndex === null ? null : entry.pageIndex - (entry.pageIndex > 1 ? 1 : 0) }))
  const edited = [...remaining.map((entry) => entry.label === 'Chapter 3' ? { ...entry, label: 'Chapter 3 (renamed)' } : entry),
    { id: 'bookmark-new', pageIndex: 0, label: 'Page 1', depth: 0, source: 'simple' }]
  const output = await flattenWithBookmarks(mutated, edited)
  const { items, pageCount } = await savedOutline(output)
  assert.equal(pageCount, 4)
  assert.deepEqual(items.map(({ title, depth }) => [title, depth]), [
    ['Chapter 1', 0], ['Website', 0], ['Part II', 0], ['Chapter 3 (renamed)', 1], ['Appendix', 0], ['Notes', 0], ['Page 1', 0],
  ])
  const byTitle = Object.fromEntries(items.map((item) => [item.title, item]))
  assert.equal(byTitle['Chapter 1'].dest, 'page1 XYZ 72 700 2')
  assert.equal(byTitle['Chapter 3 (renamed)'].dest, 'page2 FitH 300')
  assert.equal(byTitle.Notes.goTo, 'page4 Fit')
  assert.ok(items.every((item) => !String(item.dest).startsWith('missing') && !String(item.goTo).startsWith('missing')), 'no destination points at a deleted page')
  const reread = await rendererEntries(output)
  assert.deepEqual(reread.map((entry) => entry.pageIndex), [0, null, null, 1, 2, 3, 0])
})

test('new bookmarks without any document outline are written with /Fit destinations', async () => {
  const doc = await PDFDocument.create()
  doc.addPage([300, 300])
  doc.addPage([300, 300])
  const result = replacePdfOutlines(doc, [
    { title: 'Start', pageIndex: 0, depth: 0 },
    { title: 'Inner', pageIndex: 1, depth: 3 },
    { title: '', pageIndex: 0, depth: 0 },
    { title: 'Off the end', pageIndex: 9, depth: 0 },
  ], pdfLib)
  assert.deepEqual(result, { written: 2, skipped: 2 })
  const { items } = await savedOutline(await doc.save())
  assert.deepEqual(items.map(({ title, depth, dest }) => [title, depth, dest]), [['Start', 0, 'page1 Fit'], ['Inner', 1, 'page2 Fit']])
})
