'use strict'

const assert = require('node:assert/strict')
const path = require('node:path')
const test = require('node:test')
const { RECENT_LIMIT, mergeRecent, sanitizeRecents } = require('../electron/recent-files.cjs')

test('opening a video puts it first and removes a case-insensitive duplicate on Windows', () => {
  const firstPath = path.resolve('C:\\Media\\First.mp4')
  const existing = [
    { path: firstPath.toUpperCase(), name: 'stale name', openedAt: 10 },
    { path: path.resolve('C:\\Media\\Other.webm'), name: 'Other.webm', openedAt: 9 },
  ]
  const result = mergeRecent(existing, firstPath, 20, 'win32')
  assert.equal(result.length, 2)
  assert.deepEqual(result[0], { path: firstPath, name: path.basename(firstPath), openedAt: 20 })
  assert.equal(result[1].name, 'Other.webm')
})

test('invalid history entries are discarded and the list is bounded', () => {
  const entries = Array.from({ length: RECENT_LIMIT + 8 }, (_, index) => ({
    path: path.resolve(`video-${index}.mp4`),
    name: `wrong-${index}`,
    openedAt: index,
  }))
  entries.splice(2, 0, { path: '', name: '', openedAt: Number.NaN })
  const result = sanitizeRecents(entries)
  assert.equal(result.length, RECENT_LIMIT)
  assert.equal(result[0].name, path.basename(entries[0].path))
})

test('malformed stored state becomes an empty recent list', () => {
  assert.deepEqual(sanitizeRecents(null), [])
  assert.deepEqual(sanitizeRecents({}), [])
})
