'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const test = require('node:test')
const { createCanvas } = require('@napi-rs/canvas')
const {
  PDFArray, PDFDict, PDFDocument, PDFName, PDFNumber, PDFRawStream, PDFRef, PDFString, StandardFonts,
} = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { decodedStreams, documentContainsText, pageObjectCount, pageTexts } = require('./helpers/pdf-test-utils.cjs')

const SECRET = 'CONFIDENTIAL-SALARY-123456'
const FIELD_SECRET = 'SECRET-FIELD-VALUE'

function marker(width, height) {
  const canvas = createCanvas(width, height)
  const context = canvas.getContext('2d')
  context.fillStyle = '#c03'
  context.fillRect(0, 0, width, height)
  return canvas.toBuffer('image/png')
}

function outlineTree(doc, items) {
  const { context } = doc
  const rootRef = context.nextRef()
  const build = (entries, parentRef) => {
    const refs = entries.map(() => context.nextRef())
    entries.forEach((entry, index) => {
      const map = { Title: PDFString.of(entry.title), Parent: parentRef }
      if (index) map.Prev = refs[index - 1]
      if (index + 1 < refs.length) map.Next = refs[index + 1]
      if (entry.dest) map.Dest = entry.dest
      if (entry.children?.length) {
        const childRefs = build(entry.children, refs[index])
        map.First = childRefs[0]
        map.Last = childRefs[childRefs.length - 1]
        map.Count = entry.children.length
      }
      context.assign(refs[index], context.obj(map))
    })
    return refs
  }
  const refs = build(items, rootRef)
  context.assign(rootRef, context.obj({ Type: 'Outlines', First: refs[0], Last: refs[refs.length - 1], Count: 99 }))
  doc.catalog.set(PDFName.of('Outlines'), rootRef)
}

/** Three pages; page 3 holds a secret, a unique image and a form value, and everything points at it. */
async function richDocument() {
  const doc = await PDFDocument.create()
  const { context } = doc
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const pages = [doc.addPage([400, 400]), doc.addPage([400, 400]), doc.addPage([400, 400])]
  pages[0].drawText('Table of contents', { x: 20, y: 350, size: 12, font })
  pages[1].drawText('Chapter two body', { x: 20, y: 350, size: 12, font })
  pages[2].drawText(SECRET, { x: 20, y: 350, size: 12, font })
  pages[2].drawImage(await doc.embedPng(marker(7, 5)), { x: 20, y: 200, width: 70, height: 50 })
  const [p1, p2, p3] = pages.map((page) => page.ref)

  const link = (rect, entry) => context.register(context.obj({ Type: 'Annot', Subtype: 'Link', Rect: rect, Border: [0, 0, 0], ...entry }))
  pages[0].node.set(PDFName.of('Annots'), context.obj([
    link([20, 300, 200, 315], { Dest: [p3, 'XYZ', 0, 400, 0] }),
    link([20, 280, 200, 295], { Dest: PDFString.of('chapter2') }),
    link([20, 260, 200, 275], { A: { S: 'GoTo', D: [p3, 'Fit'] } }),
    link([20, 240, 200, 255], { A: { S: 'URI', URI: PDFString.of('https://example.com') } }),
  ]))

  outlineTree(doc, [
    { title: 'Contents', dest: context.obj([p1, 'Fit']) },
    { title: 'Appendix', dest: context.obj([p3, 'Fit']) },
    { title: 'Part II', dest: context.obj([p3, 'Fit']), children: [{ title: 'Chapter 2', dest: context.obj([p2, 'Fit']) }] },
  ])
  doc.catalog.set(PDFName.of('Names'), context.obj({
    Dests: { Names: [PDFString.of('appendix'), [p3, 'Fit'], PDFString.of('chapter2'), [p2, 'Fit']] },
  }))
  doc.catalog.set(PDFName.of('OpenAction'), context.obj([p3, 'Fit']))

  const structRoot = context.nextRef()
  const documentElement = context.nextRef()
  const first = context.register(context.obj({ Type: 'StructElem', S: 'P', P: documentElement, Pg: p1, K: 0 }))
  const third = context.register(context.obj({ Type: 'StructElem', S: 'P', P: documentElement, Pg: p3, K: 0 }))
  context.assign(documentElement, context.obj({ Type: 'StructElem', S: 'Document', P: structRoot, K: [first, third] }))
  context.assign(structRoot, context.obj({
    Type: 'StructTreeRoot', K: documentElement, ParentTree: { Nums: [0, [first], 1, [third]] }, ParentTreeNextKey: 2,
  }))
  pages[0].node.set(PDFName.of('StructParents'), PDFNumber.of(0))
  pages[2].node.set(PDFName.of('StructParents'), PDFNumber.of(1))
  doc.catalog.set(PDFName.of('StructTreeRoot'), structRoot)

  const form = doc.getForm()
  const secret = form.createTextField('secret')
  secret.setText(FIELD_SECRET)
  secret.addToPage(pages[2], { x: 20, y: 100, width: 200, height: 20 })
  const shared = form.createTextField('shared')
  shared.setText('shared value')
  shared.addToPage(pages[0], { x: 20, y: 100, width: 200, height: 20 })
  shared.addToPage(pages[2], { x: 20, y: 60, width: 200, height: 20 })
  return doc.save()
}

function hasImageOfSize(streams, width, height) {
  return streams.some(({ dict }) => dict.get(PDFName.of('Subtype'))?.toString() === '/Image'
    && dict.get(PDFName.of('Width'))?.asNumber?.() === width && dict.get(PDFName.of('Height'))?.asNumber?.() === height)
}

function outlineTitles(doc) {
  const titles = []
  const walk = (dict, depth) => {
    let ref = dict.get(PDFName.of('First'))
    while (ref instanceof PDFRef) {
      const item = doc.context.lookup(ref)
      titles.push({ depth, title: item.lookup(PDFName.of('Title')).decodeText(), dest: item.get(PDFName.of('Dest')) })
      walk(item, depth + 1)
      ref = item.get(PDFName.of('Next'))
    }
  }
  const root = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)
  if (root) walk(root, 0)
  return titles
}

test('a deleted page leaves nothing behind: content, image, field value, bookmarks, links, tags', async () => {
  const { invoke } = loadMain()
  const input = await richDocument()
  assert.equal(await documentContainsText(input, SECRET), true, 'the fixture contains the secret')
  const output = await invoke('pdf:mutate', input, { type: 'delete', indices: [2] })

  assert.equal(await documentContainsText(output, SECRET), false, 'deleted page text is gone')
  assert.equal(await documentContainsText(output, FIELD_SECRET), false, 'deleted form value is gone')
  assert.equal(hasImageOfSize(await decodedStreams(output), 7, 5), false, 'deleted page image is gone')
  assert.equal(await pageObjectCount(output), 2)

  const doc = await PDFDocument.load(output)
  assert.equal(doc.getPageCount(), 2)
  const pageRefs = new Set(doc.getPages().map((page) => page.ref.tag))
  const titles = outlineTitles(doc)
  assert.deepEqual(titles.map(({ depth, title }) => `${depth}:${title}`), ['0:Contents', '0:Part II', '1:Chapter 2'])
  assert.equal(titles[1].dest, undefined, 'a heading whose page was deleted keeps its children, not its destination')
  for (const { dest } of titles.filter((entry) => entry.dest)) assert.ok(pageRefs.has(doc.context.lookup(dest).get(0).tag))

  const annots = doc.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray).asArray().map((ref) => doc.context.lookup(ref))
  const links = annots.filter((annot) => annot.get(PDFName.of('Subtype'))?.toString() === '/Link')
  assert.equal(links.length, 2, 'links to the deleted page are removed; the named and URI links stay')
  const names = doc.catalog.lookup(PDFName.of('Names'), PDFDict).lookup(PDFName.of('Dests'), PDFDict).lookup(PDFName.of('Names'), PDFArray)
  assert.deepEqual(names.asArray().filter((value) => value instanceof PDFString).map((value) => value.decodeText()), ['chapter2'])
  assert.equal(doc.catalog.has(PDFName.of('OpenAction')), false)

  const fieldNames = doc.getForm().getFields().map((field) => field.getName())
  assert.deepEqual(fieldNames, ['shared'])
  assert.equal(doc.getForm().getTextField('shared').acroField.getWidgets().length, 1)

  const structRoot = doc.catalog.lookup(PDFName.of('StructTreeRoot'), PDFDict)
  const documentElement = structRoot.lookup(PDFName.of('K'), PDFDict)
  assert.equal(documentElement.lookup(PDFName.of('K'), PDFArray).size(), 1)
  const nums = structRoot.lookup(PDFName.of('ParentTree'), PDFDict).lookup(PDFName.of('Nums'), PDFArray)
  assert.deepEqual(nums.asArray().filter((value) => value instanceof PDFNumber).map((value) => value.asNumber()), [0])

  assert.deepEqual(await pageTexts(output), ['Table of contents', 'Chapter two body'])

  // Saving afterwards (with an unrelated edit) keeps it gone.
  const saved = await invoke('pdf:flatten-overlays', output, [{
    id: 'h', type: 'highlight', pageIndex: 0, rect: { x: 10, y: 340, width: 100, height: 20 }, color: [1, 1, 0], opacity: 0.3,
  }], {}, {})
  assert.equal(await documentContainsText(saved, SECRET), false)
  assert.equal(await pageObjectCount(saved), 2)
})

test('deleting an image overlay removes the image pixels from the saved file', async () => {
  const { invoke } = loadMain()
  const doc = await PDFDocument.create()
  const page = doc.addPage([400, 400])
  page.drawText('Caption stays', { x: 20, y: 350, size: 12, font: await doc.embedFont(StandardFonts.Helvetica) })
  page.drawImage(await doc.embedPng(marker(7, 5)), { x: 20, y: 200, width: 70, height: 50 })
  const input = await doc.save()
  assert.equal(hasImageOfSize(await decodedStreams(input), 7, 5), true)
  const output = await invoke('pdf:flatten-overlays', input, [{
    id: 'delete-photo', type: 'object', kind: 'image', pageIndex: 0,
    rect: { x: 20, y: 200, width: 70, height: 50 }, originalRect: { x: 20, y: 200, width: 70, height: 50 },
    opacity: 1, cover: true,
  }], {}, {})
  assert.equal(hasImageOfSize(await decodedStreams(output), 7, 5), false)
  assert.deepEqual(await pageTexts(output), ['Caption stays'])
})

test('replaced text is not kept in an orphaned content stream', async () => {
  const { invoke } = loadMain()
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const page = doc.addPage([400, 400])
  page.drawText('Account holder', { x: 40, y: 360, size: 12, font })
  page.drawText('SSN 123-45-6789', { x: 40, y: 300, size: 12, font })
  const input = await doc.save()
  const width = font.widthOfTextAtSize('SSN 123-45-6789', 12)
  const output = await invoke('pdf:flatten-overlays', input, [{
    id: 'ssn', type: 'text', pageIndex: 0,
    rect: { x: 40, y: 298, width: 120, height: 14 },
    originalRect: { x: 40, y: 298, width, height: 12 },
    originalText: 'SSN 123-45-6789', text: 'SSN on file',
    fontSize: 12, fontFamily: 'Helvetica', color: [0, 0, 0], align: 'left', cover: true, textFit: 'fit',
  }], {}, {})
  assert.equal(await documentContainsText(output, '123-45-6789'), false)
  const [text] = await pageTexts(output)
  assert.match(text, /Account holder/)
  assert.match(text, /SSN on file/)
})

/** Page 1 is a table of contents linking to pages 2 and 3. */
async function linkedDocument() {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const pages = [doc.addPage([400, 400]), doc.addPage([400, 400]), doc.addPage([400, 400])]
  pages[0].drawText('Contents page', { x: 20, y: 350, size: 12, font })
  pages[1].drawText('Public chapter', { x: 20, y: 350, size: 12, font })
  pages[2].drawText('CONFIDENTIAL-APPENDIX-PAYROLL', { x: 20, y: 350, size: 12, font })
  const { context } = doc
  pages[0].node.set(PDFName.of('Annots'), context.obj([
    context.register(context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [20, 300, 200, 315], Dest: [pages[1].ref, 'Fit'] })),
    context.register(context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [20, 280, 200, 295], Dest: [pages[2].ref, 'Fit'] })),
  ]))
  return doc.save()
}

async function exportWithHarness(data, indices) {
  const { invoke, saveAnswers, tempFile } = loadMain()
  const target = tempFile(`export-${indices.join('-')}-${Date.now()}.pdf`)
  saveAnswers.push(target)
  assert.equal(await invoke('pdf:export-as-pdf', { data, indices, fullDocument: false, suggestedName: 'Export' }), target)
  return fs.readFileSync(target)
}

test('exporting selected pages does not carry linked pages along', async () => {
  const input = await linkedDocument()
  const single = await exportWithHarness(input, [0])
  assert.equal(await documentContainsText(single, 'CONFIDENTIAL-APPENDIX-PAYROLL'), false)
  assert.equal(await documentContainsText(single, 'Public chapter'), false)
  assert.equal(await pageObjectCount(single), 1)
  const singleDoc = await PDFDocument.load(single)
  assert.equal(singleDoc.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray).size(), 0, 'links to pages not exported are removed')

  const two = await exportWithHarness(input, [0, 1])
  assert.equal(await documentContainsText(two, 'CONFIDENTIAL-APPENDIX-PAYROLL'), false)
  assert.equal(await pageObjectCount(two), 2)
  const twoDoc = await PDFDocument.load(two)
  const annots = twoDoc.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
  assert.equal(annots.size(), 1)
  const destination = twoDoc.context.lookup(annots.get(0)).lookup(PDFName.of('Dest'), PDFArray)
  assert.equal(destination.get(0).tag, twoDoc.getPage(1).ref.tag, 'the remaining link points at the exported copy of page 2')
})

test('inserted documents keep working internal links and no hidden page copies', async () => {
  const { invoke } = loadMain()
  const base = await PDFDocument.create()
  base.addPage([400, 400])
  const inserted = await invoke('pdf:insert-dropped-files', await base.save(), 1, [{ name: 'linked.pdf', data: await linkedDocument() }])
  assert.equal(inserted.added, 3)
  assert.equal(await pageObjectCount(inserted.data), 4)
  const doc = await PDFDocument.load(inserted.data)
  const annots = doc.getPage(1).node.lookup(PDFName.of('Annots'), PDFArray)
  const targets = annots.asArray().map((ref) => doc.context.lookup(ref).lookup(PDFName.of('Dest'), PDFArray).get(0).tag)
  assert.deepEqual(targets, [doc.getPage(2).ref.tag, doc.getPage(3).ref.tag])
})

test('compaction keeps every reachable object and drops orphans', async () => {
  const { compactUnreachable } = require('../electron/pdf-compact.cjs')
  const input = await richDocument()
  const doc = await PDFDocument.load(input)
  const orphan = doc.context.register(doc.context.flateStream('BT (ORPHANED-OLD-CONTENT) Tj ET'))
  const before = doc.context.enumerateIndirectObjects().length
  // pdf-lib's own form API also leaves superseded appearance streams behind.
  const { deleted } = compactUnreachable(doc)
  assert.ok(deleted >= 1)
  assert.equal(doc.context.lookup(orphan), undefined)
  assert.equal(doc.context.enumerateIndirectObjects().length, before - deleted)
  // Nothing that is still referenced was deleted: every reference resolves.
  const dangling = []
  const check = (value) => {
    if (value instanceof PDFRef) { if (doc.context.lookup(value) === undefined) dangling.push(value.tag); return }
    if (value instanceof PDFDict) value.entries().forEach(([, child]) => check(child))
    else if (value instanceof PDFArray) value.asArray().forEach(check)
  }
  for (const [, object] of doc.context.enumerateIndirectObjects()) check(object instanceof PDFRawStream ? object.dict : object)
  assert.deepEqual(dangling, [])
  const output = await doc.save()
  assert.deepEqual(await pageTexts(output), await pageTexts(input))
  assert.equal(await documentContainsText(output, 'ORPHANED-OLD-CONTENT'), false)
  const reloaded = await PDFDocument.load(output)
  assert.equal(reloaded.getForm().getFields().length, 2)
  assert.equal(outlineTitles(reloaded).length, 4)
})
