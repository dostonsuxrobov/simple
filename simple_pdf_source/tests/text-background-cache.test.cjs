'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { PDFDocument, PDFName, StandardFonts } = require('pdf-lib')
const { createTextBackgroundCache } = require('../electron/text-background-cache.cjs')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { documentContainsText, pageObjectCount, pageTexts } = require('./helpers/pdf-test-utils.cjs')

function countingCache() {
  const calls = { parse: 0, extract: 0, remove: 0 }
  const cache = createTextBackgroundCache({
    parse: async (bytes) => { calls.parse += 1; return { length: bytes.length } },
    extractPage: async (_doc, pageIndex) => { calls.extract += 1; return new Uint8Array([pageIndex]) },
    removeText: async (page, edits) => { calls.remove += 1; return new Uint8Array([page[0], edits.length]) },
  })
  return { cache, calls }
}

const edit = (text) => ({ type: 'text', cover: true, originalText: text, originalRect: { x: 1, y: 2, width: 3, height: 4 }, pageIndex: 5 })

test('one parse per document revision, one extraction per page, one result per edit set', async () => {
  const { cache, calls } = countingCache()
  const revision = new Uint8Array([1, 2, 3])
  const first = await cache.render(revision, 0, [edit('a')])
  assert.deepEqual(await cache.render(new Uint8Array([1, 2, 3]), 0, [edit('a')]), first)
  assert.deepEqual(calls, { parse: 1, extract: 1, remove: 1 }, 'a repeated request (same bytes, page, edits) is served from cache')
  await cache.render(revision, 0, [edit('b')])
  assert.deepEqual(calls, { parse: 1, extract: 1, remove: 2 })
  await cache.render(revision, 1, [edit('a')])
  assert.deepEqual(calls, { parse: 1, extract: 2, remove: 3 })
  await cache.render(new Uint8Array([1, 2, 4]), 0, [edit('a')])
  assert.deepEqual(calls, { parse: 2, extract: 3, remove: 4 }, 'a new revision is parsed again')
})

test('failures are not cached and concurrent requests share one computation', async () => {
  let attempts = 0
  const cache = createTextBackgroundCache({
    parse: async () => ({}),
    extractPage: async () => new Uint8Array([7]),
    removeText: async () => {
      attempts += 1
      if (attempts === 1) throw new Error('not located')
      return new Uint8Array([attempts])
    },
  })
  await assert.rejects(cache.render(new Uint8Array([9]), 0, []), /not located/)
  const [left, right] = await Promise.all([cache.render(new Uint8Array([9]), 0, []), cache.render(new Uint8Array([9]), 0, [])])
  assert.deepEqual(left, right)
  assert.equal(attempts, 2)
})

test('the preview page is the selected page only, with the edited text removed', async () => {
  const { invoke } = loadMain()
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const first = doc.addPage([300, 300])
  const second = doc.addPage([300, 300])
  first.drawText('Editable line', { x: 20, y: 250, size: 12, font })
  first.drawText('Neighbour line', { x: 20, y: 200, size: 12, font })
  second.drawText('OTHER-PAGE-CONTENT', { x: 20, y: 250, size: 12, font })
  first.node.set(PDFName.of('Annots'), doc.context.obj([
    doc.context.register(doc.context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [20, 100, 100, 120], Dest: [second.ref, 'Fit'] })),
  ]))
  const bytes = await doc.save()
  const edits = [{ type: 'text', cover: true, originalText: 'Editable line', originalRect: { x: 20, y: 248, width: font.widthOfTextAtSize('Editable line', 12), height: 12 } }]
  const preview = await invoke('pdf:text-background', bytes, 0, edits)
  assert.deepEqual(await pageTexts(preview), ['Neighbour line'])
  assert.equal(await pageObjectCount(preview), 1)
  assert.equal(await documentContainsText(preview, 'OTHER-PAGE-CONTENT'), false, 'a linked page is not copied into the preview')
  assert.deepEqual(await invoke('pdf:text-background', bytes, 0, edits), preview)
})
