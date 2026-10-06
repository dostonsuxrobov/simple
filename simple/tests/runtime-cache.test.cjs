'use strict'

// The portable launcher's runtime cache: a running build removes the unpacked
// folders of earlier builds, never its own, a newer one or one that is in use.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { pruneRuntimeCache } = require('../electron/runtime-cache.cjs')

function scratch(t) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-runtime-cache-'))
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }))
  return directory
}

/** One unpacked build: simple.exe and resources/app.asar dated `builtAt`. */
function build(root, name, builtAt) {
  const directory = path.join(root, name)
  fs.mkdirSync(path.join(directory, 'resources'), { recursive: true })
  fs.writeFileSync(path.join(directory, 'simple.exe'), 'exe')
  fs.writeFileSync(path.join(directory, '.simple-unified-runtime-ready'), 'ready')
  if (builtAt !== null) {
    const asar = path.join(directory, 'resources', 'app.asar')
    fs.writeFileSync(asar, 'asar')
    fs.utimesSync(asar, builtAt, builtAt)
  }
  return path.join(directory, 'simple.exe')
}

test('removes earlier builds and keeps the running, same and newer ones', async (t) => {
  const cache = path.join(scratch(t), 'simple', 'cache')
  const day = (n) => new Date(Date.UTC(2026, 9, n))
  const execPath = build(cache, 'current', day(5))
  build(cache, 'older-a', day(1))
  build(cache, 'older-b', day(4))
  build(cache, 'same', day(5))
  build(cache, 'newer', day(6))
  build(cache, 'incomplete', null)
  fs.mkdirSync(path.join(cache, 'older-c.simple-trash-123', 'resources'), { recursive: true })

  const result = await pruneRuntimeCache({ execPath })
  assert.deepEqual(result.removed.sort(), ['incomplete', 'older-a', 'older-b'])
  assert.deepEqual(result.kept.sort(), ['newer', 'same'])
  assert.deepEqual(fs.readdirSync(cache).sort(), ['current', 'newer', 'same'])
  assert.ok(fs.existsSync(path.join(cache, 'current', 'resources', 'app.asar')), 'the running build is untouched')
})

test('does nothing outside the launcher cache', async (t) => {
  const root = path.join(scratch(t), 'Program Files', 'simple')
  const execPath = build(root, 'current', new Date(Date.UTC(2026, 9, 5)))
  build(root, 'older', new Date(Date.UTC(2026, 9, 1)))
  const result = await pruneRuntimeCache({ execPath })
  assert.deepEqual(result, { removed: [], kept: [] })
  assert.deepEqual(fs.readdirSync(root).sort(), ['current', 'older'])
})

test('keeps an earlier build whose files are open', { skip: process.platform !== 'win32' && 'Windows refuses the rename only there' }, async (t) => {
  const cache = path.join(scratch(t), 'simple', 'cache')
  const execPath = build(cache, 'current', new Date(Date.UTC(2026, 9, 5)))
  build(cache, 'running', new Date(Date.UTC(2026, 9, 1)))
  // A running runtime holds app.asar open, as this handle does.
  const handle = fs.openSync(path.join(cache, 'running', 'resources', 'app.asar'), 'r')
  t.after(() => fs.closeSync(handle))
  const result = await pruneRuntimeCache({ execPath })
  assert.deepEqual(result, { removed: [], kept: ['running'] })
  assert.ok(fs.existsSync(path.join(cache, 'running', 'simple.exe')), 'nothing inside it was deleted')
})
