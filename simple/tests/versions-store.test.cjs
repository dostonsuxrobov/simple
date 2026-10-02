'use strict'

// The versions store of the shared Save layer (design §3.8): a copy of a file
// as it was on disk before its first overwrite in each session, kept per file
// (10 versions), for 60 days and within 500 MB. Every case uses its own temp
// folder; nothing outside it is touched.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')

const SHARED = path.resolve(__dirname, '..', 'shared')
const core = require(path.join(SHARED, 'electron', 'io-core.cjs'))
const stores = require(path.join(SHARED, 'electron', 'stores.cjs'))

const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-versions-store-'))
let caseCount = 0

core.configureIo({ journalDir: path.join(ROOT, 'journal'), logger: { warn() {}, info() {} } })
stores.configureStores({ logger: { warn() {} } })
test.after(() => fs.rmSync(ROOT, { recursive: true, force: true }))

function scenario(name, options = {}) {
  caseCount += 1
  const folder = path.join(ROOT, `${String(caseCount).padStart(2, '0')}-${name}`)
  fs.mkdirSync(folder, { recursive: true })
  return { folder, store: stores.createVersionsStore({ root: path.join(folder, 'versions'), ...options }) }
}

function versionTexts(entries) {
  return entries.map((entry) => fs.readFileSync(entry.path, 'utf8'))
}

test('the first overwrite in a session keeps one version, and later saves add none', async () => {
  const { folder, store } = scenario('once')
  const file = path.join(folder, 'Budget.xlsx')
  fs.writeFileSync(file, 'version one')
  const first = await store.backupOnce(file)
  assert.equal(first.ok, true)
  assert.match(first.id, /^[0-9a-f]{40}\/\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}-\d{3}Z-Budget\.xlsx$/)
  fs.writeFileSync(file, 'version two')
  assert.deepEqual(await store.backupOnce(file), { ok: true, skipped: 'already-backed-up' })
  if (process.platform === 'win32') {
    assert.deepEqual(await store.backupOnce(file.toUpperCase()), { ok: true, skipped: 'already-backed-up' },
      'the same file under another spelling is the same file')
  }
  const listed = await store.list(file)
  assert.equal(listed.length, 1)
  assert.deepEqual(versionTexts(listed), ['version one'])
  assert.equal(listed[0].name, 'Budget.xlsx')
  assert.equal(listed[0].sourcePath, path.resolve(file))
  assert.equal(listed[0].size, 'version one'.length)
  assert.ok(store.hasBackup(file))

  store.resetSession()
  assert.equal((await store.backupOnce(file)).ok, true, 'a new session copies the file again')
  assert.deepEqual(versionTexts(await store.list(file)), ['version two', 'version one'], 'newest first')
})

test('concurrent first saves of one file make a single copy', async () => {
  const { folder, store } = scenario('concurrent')
  const file = path.join(folder, 'Report.docx')
  fs.writeFileSync(file, 'original')
  const results = await Promise.all([store.backupOnce(file), store.backupOnce(file), store.backupOnce(file)])
  assert.equal(results.filter((result) => result.id).length, 1)
  assert.equal((await store.list(file)).length, 1)
})

test('the eleventh version removes the oldest', async () => {
  const { folder, store } = scenario('eleven')
  const file = path.join(folder, 'Notes.txt')
  for (let index = 1; index <= 11; index += 1) {
    fs.writeFileSync(file, `content ${index}`)
    const result = await store.backup(file)
    assert.equal(result.ok, true, JSON.stringify(result))
  }
  await store.idle()
  const listed = await store.list(file)
  assert.equal(listed.length, 10)
  assert.deepEqual(versionTexts(listed), [11, 10, 9, 8, 7, 6, 5, 4, 3, 2].map((index) => `content ${index}`))
})

test('copies made in the same millisecond keep their order, even for names that start with digits', async () => {
  const { folder, store } = scenario('same-ms')
  const file = path.join(folder, '2-report.txt')
  const realNow = Date.now
  const frozen = Math.floor(Date.now() / 1000) * 1000
  Date.now = () => frozen
  const RealDate = Date
  globalThis.Date = class extends RealDate {
    constructor(...args) { super(...(args.length ? args : [frozen])) }
    static now() { return frozen }
  }
  try {
    for (let index = 1; index <= 3; index += 1) {
      fs.writeFileSync(file, `copy ${index}`)
      assert.equal((await store.backup(file)).ok, true)
    }
  } finally {
    globalThis.Date = RealDate
    Date.now = realNow
  }
  const listed = await store.list(file)
  assert.deepEqual(versionTexts(listed), ['copy 3', 'copy 2', 'copy 1'])
  assert.ok(listed.every((entry) => entry.name === '2-report.txt'), 'the original name is kept exactly')
  assert.deepEqual(listed.map((entry) => entry.createdAt), Array(3).fill(new Date(frozen).toISOString()))
})

test('version ids resolve only inside the store', async () => {
  const { folder, store } = scenario('resolve')
  const file = path.join(folder, 'Plan.md')
  fs.writeFileSync(file, '# plan')
  const { id } = await store.backup(file)
  const resolved = await store.resolve(id)
  assert.equal(fs.readFileSync(resolved.path, 'utf8'), '# plan')
  assert.equal(resolved.name, 'Plan.md')
  assert.equal(resolved.sourcePath, path.resolve(file))
  const key = id.split('/')[0]
  for (const bad of ['', '../x', `${key}/../../outside.txt`, `${key}\\..\\x`, `${key}/source.json`, `${'0'.repeat(40)}/2026-10-02T09-00-00-000Z-missing.txt`, null, 42]) {
    assert.equal(await store.resolve(bad), null, `${bad} must not resolve`)
  }
})

test('old versions expire, and the whole store stays under its size cap', async () => {
  const { folder, store } = scenario('retention', { maxTotalBytes: 25_000 })
  const day = 24 * 60 * 60 * 1000
  const files = ['a.bin', 'b.bin', 'c.bin'].map((name) => path.join(folder, name))
  for (const file of files) {
    fs.writeFileSync(file, Buffer.alloc(10_000, 1))
    assert.equal((await store.backup(file)).ok, true)
  }
  await store.idle()
  const total = async () => (await Promise.all(files.map((file) => store.list(file)))).flat().reduce((sum, entry) => sum + entry.size, 0)
  assert.ok(await total() <= 25_000, 'the size cap removed the oldest copy')
  assert.equal((await store.list(files[0])).length, 0)
  assert.equal((await store.list(files[2])).length, 1, 'the newest copy is kept')

  const report = await store.prune({ now: Date.now() + 61 * day })
  assert.ok(report.removed >= 1)
  assert.equal(await total(), 0, 'versions older than 60 days are removed')
  assert.deepEqual(fs.readdirSync(path.join(folder, 'versions')), [], 'empty version folders are removed too')
})

test('missing, oversized and unreadable files are reported, never thrown', async () => {
  const { folder, store } = scenario('skips', { maxFileBytes: 1000 })
  assert.deepEqual(await store.backupOnce(path.join(folder, 'new.docx')), { ok: true, skipped: 'missing' })
  const big = path.join(folder, 'big.pdf')
  fs.writeFileSync(big, Buffer.alloc(2000))
  const tooLarge = await store.backupOnce(big)
  assert.equal(tooLarge.ok, false)
  assert.equal(tooLarge.skipped, 'too-large')
  assert.match(tooLarge.reason, /larger than/)
  assert.equal(store.hasBackup(big), false, 'a skipped copy is tried again next time')
  assert.deepEqual(await store.backupOnce(folder), { ok: true, skipped: 'not-a-file' })
  assert.deepEqual(await store.list(big), [])
})
