'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { changedOnDisk, createWindowDocuments, defaultFileKey } = require('../electron/window-documents.cjs')

test('a file that is already open is found again, whatever the case of its path', () => {
  const documents = createWindowDocuments()
  documents.setCurrent(7, 'C:\\Docs\\Report.pdf')
  assert.equal(documents.findWindow('c:\\docs\\REPORT.PDF'), 7)
  assert.equal(documents.findWindow('C:\\Docs\\Other.pdf'), null)
  assert.equal(defaultFileKey('C:\\Docs\\Report.pdf'), defaultFileKey('c:\\DOCS\\report.PDF'))
})

test('the window that asks is never told to focus itself', () => {
  const documents = createWindowDocuments()
  documents.setCurrent(7, 'C:\\Docs\\Report.pdf')
  assert.equal(documents.findWindow('C:\\Docs\\Report.pdf', 7), null)
  documents.setCurrent(8, 'C:\\Docs\\Report.pdf')
  assert.equal(documents.findWindow('C:\\Docs\\Report.pdf', 7), 8)
})

test('the record follows the window: another file, Save As, back home, closed', () => {
  const documents = createWindowDocuments()
  documents.setCurrent(1, 'C:\\a.pdf')
  documents.setCurrent(1, 'C:\\b.pdf')
  assert.equal(documents.findWindow('C:\\a.pdf'), null, 'the first file is free again')
  assert.equal(documents.findWindow('C:\\b.pdf'), 1)
  documents.setCurrent(1, null)
  assert.equal(documents.findWindow('C:\\b.pdf'), null, 'the home screen shows no file')
  documents.setCurrent(1, 'C:\\c.pdf')
  documents.forget(1)
  assert.equal(documents.findWindow('C:\\c.pdf'), null, 'a closed window shows nothing')
  assert.equal(documents.stampFor(1, 'C:\\c.pdf'), null)
})

test('a window created for a file counts as showing it only until the claim runs out', () => {
  let time = 1_000
  const documents = createWindowDocuments({ now: () => time, pendingMs: 500 })
  documents.expect(3, 'C:\\a.pdf')
  assert.equal(documents.findWindow('C:\\a.pdf'), 3, 'a second open right away focuses the new window')
  time += 501
  assert.equal(documents.findWindow('C:\\a.pdf'), null, 'an open that never happened does not block the file')
  documents.expect(4, 'C:\\b.pdf')
  documents.setCurrent(4, 'C:\\b.pdf')
  time += 10_000
  assert.equal(documents.findWindow('C:\\b.pdf'), 4, 'a confirmed open does not run out')
})

test('stamps belong to one window and one file, and follow the file the window shows', () => {
  const documents = createWindowDocuments()
  documents.noteStamp(1, 'C:\\a.pdf', { size: 10, mtimeMs: 5, extra: true })
  assert.deepEqual(documents.stampFor(1, 'c:\\A.PDF'), { size: 10, mtimeMs: 5 })
  assert.equal(documents.stampFor(2, 'C:\\a.pdf'), null)
  // The window reads a second file (a drop) and then shows it: the first stamp is dropped.
  documents.noteStamp(1, 'C:\\b.pdf', { size: 20, mtimeMs: 6 })
  documents.setCurrent(1, 'C:\\b.pdf')
  assert.equal(documents.stampFor(1, 'C:\\a.pdf'), null)
  assert.deepEqual(documents.stampFor(1, 'C:\\b.pdf'), { size: 20, mtimeMs: 6 })
  // A save whose result could not be read back leaves no stamp (no false alarm later).
  documents.noteStamp(1, 'C:\\b.pdf', null)
  assert.equal(documents.stampFor(1, 'C:\\b.pdf'), null)
})

test('a file counts as changed on disk when its size or modified time differs', () => {
  const stamp = { size: 10, mtimeMs: 5 }
  assert.equal(changedOnDisk(stamp, { size: 10, mtimeMs: 5 }), false)
  assert.equal(changedOnDisk(stamp, { size: 11, mtimeMs: 5 }), true)
  assert.equal(changedOnDisk(stamp, { size: 10, mtimeMs: 6 }), true)
  // Nothing to compare (never read by this window, or the file is gone): no question is asked.
  assert.equal(changedOnDisk(null, { size: 10, mtimeMs: 5 }), false)
  assert.equal(changedOnDisk(stamp, null), false)
})
