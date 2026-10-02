'use strict'

// The recovery journal of the shared Save layer (design §4): one folder per
// document, generation-numbered parts and a manifest written last. Every case
// runs in its own temp folder; "another process" is simulated with a second
// store that has a different process token. One case kills a real child
// process between a part write and the manifest write.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawn } = require('node:child_process')

const SHARED = path.resolve(__dirname, '..', 'shared')
const core = require(path.join(SHARED, 'electron', 'io-core.cjs'))
const stores = require(path.join(SHARED, 'electron', 'stores.cjs'))

const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-recovery-store-'))
const JOURNAL = path.join(ROOT, 'journal')
let caseCount = 0

core.configureIo({ journalDir: JOURNAL, logger: { warn() {}, info() {} } })
stores.configureStores({ logger: { warn() {} } })
test.after(() => fs.rmSync(ROOT, { recursive: true, force: true }))

function scenario(name) {
  caseCount += 1
  const folder = path.join(ROOT, `${String(caseCount).padStart(2, '0')}-${name}`)
  fs.mkdirSync(folder, { recursive: true })
  return folder
}

function storesFor(folder) {
  const root = path.join(folder, 'userData', 'recovery')
  return {
    root,
    here: stores.createRecoveryStore({ root, processToken: 'process-here' }),
    later: stores.createRecoveryStore({ root, processToken: 'process-later' }),
  }
}

function snapshot(revision, extra = {}) {
  return {
    docId: 'doc-1',
    revision,
    module: 'calc',
    title: 'Budget.xlsx',
    kind: 'spreadsheet',
    format: 'xlsx',
    parts: [
      { name: 'workbook.json', data: `{"revision":${revision}}` },
      { name: 'view.json', data: Buffer.from('{"sheet":1}') },
    ],
    ...extra,
  }
}

function partText(result, name) {
  const part = result.parts.find((item) => item.name === name)
  assert.ok(part, `part ${name} is missing`)
  return typeof part.data === 'string' ? part.data : part.data.toString('utf8')
}

test('write, list, read and discard round-trip exactly, and only orphaned entries are offered', async () => {
  const folder = scenario('round-trip')
  const { root, here, later } = storesFor(folder)
  const source = path.join(folder, 'Budget.xlsx')
  fs.writeFileSync(source, 'original workbook bytes')
  const sourceStamp = await core.stampFile(source)
  const canvas = Uint8Array.from([137, 80, 78, 71, 0, 1, 2, 3, 255])
  const big = Buffer.from('cell;'.repeat(5000))
  const written = await here.write({
    docId: 'doc-1',
    revision: 7,
    module: 'calc',
    title: 'Budget.xlsx',
    kind: 'spreadsheet',
    format: 'xlsx',
    sourcePath: source,
    sourceStamp,
    meta: { sheet: 2, selection: 'B2:D5' },
    parts: [
      { name: 'state.json', data: '{"edits":["B2=42"]}' },
      { name: 'canvas.png', data: canvas },
      { name: 'workbook.json', data: big, compress: true },
      { name: 'base', ref: { path: source, stamp: sourceStamp }, required: true },
    ],
  })
  assert.equal(written.ok, true, JSON.stringify(written))
  assert.deepEqual(written.written.sort(), ['canvas.png', 'state.json', 'workbook.json'])
  const stored = fs.statSync(path.join(root, 'doc-1', 'workbook.json.1')).size
  assert.ok(stored < big.length / 10, 'a part asked to be compressed is stored compressed')

  assert.deepEqual(await here.list(), [], 'the live entry of this process is not offered')
  const listed = await later.list()
  assert.equal(listed.length, 1)
  const entry = listed[0]
  assert.equal(entry.id, 'doc-1')
  assert.equal(entry.title, 'Budget.xlsx')
  assert.equal(entry.orphaned, true)
  assert.equal(entry.state, 'orphaned')
  assert.equal(entry.sourcePath, source)
  assert.equal(entry.sourceState, 'same')
  assert.equal(entry.sourceChanged, false)
  assert.equal(entry.sourceMissing, false)
  assert.equal(entry.restorable, true)
  assert.equal(entry.revision, 7)
  assert.ok(entry.offeredAt, 'listing marks the entry as offered')

  const read = await later.read('doc-1')
  assert.equal(read.ok, true)
  assert.equal(read.revision, 7)
  assert.equal(read.fellBack, false)
  assert.deepEqual(read.extra, { sheet: 2, selection: 'B2:D5' })
  assert.deepEqual(read.parts.map((part) => part.name), ['state.json', 'canvas.png', 'workbook.json', 'base'])
  assert.equal(read.parts[0].data, '{"edits":["B2=42"]}', 'text parts come back as text')
  assert.deepEqual(new Uint8Array(read.parts[1].data), canvas)
  assert.ok(Buffer.isBuffer(read.parts[2].data) && read.parts[2].data.equals(big))
  assert.equal(read.parts[3].ref.path, source)
  assert.equal(read.parts[3].ref.matches, true)
  assert.deepEqual(read.entry.sourceStamp, sourceStamp)

  const discarded = await later.discard('doc-1')
  assert.deepEqual(discarded, { ok: true, removed: true })
  assert.equal(fs.existsSync(path.join(root, 'doc-1')), false)
  assert.deepEqual(await later.list(), [])
  assert.deepEqual(await later.read('doc-1'), { ok: false, code: 'NOT_FOUND' })
})

test('an unchanged part is not rewritten, and only the previous generation is kept', async () => {
  const folder = scenario('generations')
  const { root, here } = storesFor(folder)
  const files = () => fs.readdirSync(path.join(root, 'doc-1')).filter((name) => !name.startsWith('~simple-')).sort()
  const first = await here.write({ docId: 'doc-1', revision: 1, parts: [{ name: 'a.json', data: 'A1' }, { name: 'b.json', data: 'B1' }] })
  assert.deepEqual(first.written, ['a.json', 'b.json'])
  const second = await here.write({ docId: 'doc-1', revision: 2, parts: [{ name: 'a.json', data: 'A1' }, { name: 'b.json', data: 'B2' }] })
  assert.equal(second.generation, 2)
  assert.deepEqual(second.reused, ['a.json'])
  assert.deepEqual(second.written, ['b.json'])
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'doc-1', 'manifest.json'), 'utf8'))
  assert.equal(manifest.parts.find((part) => part.name === 'a.json').file, 'a.json.1', 'the unchanged part keeps its file name')
  assert.equal(manifest.parts.find((part) => part.name === 'b.json').file, 'b.json.2')
  await here.write({ docId: 'doc-1', revision: 3, parts: [{ name: 'a.json', data: 'A3' }, { name: 'b.json', data: 'B2' }] })
  assert.deepEqual(files(), ['a.json.1', 'a.json.3', 'b.json.2', 'manifest.json'], 'generation 1 files that generation 2 no longer used are gone')
  await here.write({ docId: 'doc-1', revision: 4, parts: [{ name: 'a.json', data: 'A3' }, { name: 'b.json', data: 'B4' }] })
  assert.deepEqual(files(), ['a.json.3', 'b.json.2', 'b.json.4', 'manifest.json'], 'files older than the previous generation are removed')
  const read = await here.read('doc-1')
  assert.equal(partText(read, 'a.json'), 'A3')
  assert.equal(partText(read, 'b.json'), 'B4')
})

test('a process killed between a part write and the manifest write leaves the previous generation readable', async (t) => {
  const folder = scenario('killed')
  const { root, later } = storesFor(folder)
  const first = await stores.createRecoveryStore({ root, processToken: 'process-first' }).write(snapshot(1))
  assert.equal(first.ok, true)
  const script = `
    const core = require(${JSON.stringify(path.join(SHARED, 'electron', 'io-core.cjs'))})
    const stores = require(${JSON.stringify(path.join(SHARED, 'electron', 'stores.cjs'))})
    core.configureIo({ journalDir: ${JSON.stringify(path.join(folder, 'child-journal'))}, logger: { warn() {}, info() {} } })
    const store = stores.createRecoveryStore({ root: ${JSON.stringify(root)}, processToken: 'process-child', testHooks: {
      beforeManifestWrite: () => { process.stdout.write('READY\\n'); return new Promise(() => {}) },
    } })
    store.write({ docId: 'doc-1', revision: 2, parts: [{ name: 'workbook.json', data: 'SECOND' }, { name: 'view.json', data: 'VIEW2' }] })
    setTimeout(() => process.exit(3), 60000)
  `
  const child = spawn(process.execPath, ['-e', script], { stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true })
  t.after(() => { try { child.kill('SIGKILL') } catch {} })
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('the child never reached the manifest write')), 30000)
    let output = ''
    child.stdout.on('data', (chunk) => {
      output += chunk
      if (output.includes('READY')) {
        clearTimeout(timer)
        resolve()
      }
    })
    child.on('exit', (code) => reject(new Error(`the child exited early with ${code}`)))
  })
  child.removeAllListeners('exit')
  const exited = new Promise((resolve) => child.once('exit', resolve))
  child.kill('SIGKILL')
  await exited
  assert.ok(fs.existsSync(path.join(root, 'doc-1', 'workbook.json.2')), 'the killed write left its new part behind')
  const read = await later.read('doc-1')
  assert.equal(read.ok, true)
  assert.equal(read.revision, 1)
  assert.equal(partText(read, 'workbook.json'), '{"revision":1}', 'the previous generation is intact')
  const next = await later.write(snapshot(2))
  assert.equal(next.ok, true, 'the next write replaces the abandoned part file')
  assert.equal(partText(await later.read('doc-1'), 'workbook.json'), '{"revision":2}')
})

test('a write at or below a discarded revision is dropped, and so is a stale snapshot', async () => {
  const folder = scenario('ordering')
  const { root, here } = storesFor(folder)
  assert.equal((await here.write(snapshot(5))).ok, true)
  assert.deepEqual(await here.discard('doc-1', 5), { ok: true, removed: true })
  assert.deepEqual(await here.write(snapshot(4)), { ok: false, docId: 'doc-1', dropped: 'discarded' })
  assert.deepEqual(await here.write(snapshot(5)), { ok: false, docId: 'doc-1', dropped: 'discarded' })
  assert.equal(fs.existsSync(path.join(root, 'doc-1')), false, 'a late snapshot never brings the entry back')
  assert.equal((await here.write(snapshot(6))).ok, true, 'newer edits are journaled again')

  assert.equal((await here.write(snapshot(9))).ok, true)
  assert.deepEqual(await here.write(snapshot(8)), { ok: false, docId: 'doc-1', dropped: 'stale' })
  assert.deepEqual(await here.discard('doc-1', 8), { ok: true, removed: false, kept: true }, 'an entry holding newer edits is kept')
  assert.equal(partText(await here.read('doc-1'), 'workbook.json'), '{"revision":9}')
})

test('writes and discards of one document run in call order', async () => {
  const folder = scenario('queue')
  const { here } = storesFor(folder)
  const results = await Promise.all([here.write(snapshot(1)), here.write(snapshot(2)), here.discard('doc-1', 2), here.write(snapshot(2)), here.write(snapshot(3))])
  assert.equal(results[0].ok, true)
  assert.equal(results[1].ok, true)
  assert.equal(results[2].removed, true)
  assert.equal(results[3].dropped, 'discarded')
  assert.equal(results[4].ok, true)
  assert.equal((await here.read('doc-1')).revision, 3)
})

test('the Documents workspace legacy layout migrates once and stays migrated', async () => {
  const folder = scenario('docs-legacy')
  const userData = path.join(folder, 'userData')
  const root = path.join(userData, 'recovery')
  fs.mkdirSync(root, { recursive: true })
  const docx = Buffer.from('PK\u0003\u0004 pretend docx bytes')
  fs.writeFileSync(path.join(root, 'abc-123.docx'), docx)
  fs.writeFileSync(path.join(userData, 'recoveries.json'), JSON.stringify([
    { id: 'abc-123', title: 'Report', sourcePath: 'C:\\Users\\someone\\Documents\\Report.docx', updatedAt: Date.parse('2026-09-30T10:00:00Z') },
    { id: 'gone-456', title: 'Missing', sourcePath: null, updatedAt: Date.now() },
    { id: '../../evil', title: 'Evil', sourcePath: null, updatedAt: Date.now() },
  ]))
  const store = stores.createRecoveryStore({ root, processToken: 'process-new' })
  const report = await store.migrateDocsLegacy({ userData })
  assert.deepEqual(report, { migrated: 1, skipped: 2, indexRemoved: true })
  assert.equal(fs.existsSync(path.join(userData, 'recoveries.json')), false)
  assert.equal(fs.existsSync(path.join(root, 'abc-123.docx')), false)
  const listed = await store.list()
  assert.equal(listed.length, 1)
  assert.equal(listed[0].id, 'abc-123')
  assert.equal(listed[0].title, 'Report.docx')
  assert.equal(listed[0].sourcePath, null, 'migrated entries restore as untitled documents')
  assert.equal(listed[0].suggestedPath, path.resolve('C:\\Users\\someone\\Documents\\Report.docx'))
  assert.equal(listed[0].module, 'docs')
  assert.equal(listed[0].updatedAt, '2026-09-30T10:00:00.000Z')
  const read = await store.read('abc-123')
  assert.ok(read.parts[0].data.equals(docx))
  assert.equal(read.parts[0].name, 'document.docx')

  assert.deepEqual(await store.migrateDocsLegacy({ userData }), { migrated: 0, skipped: 0, indexRemoved: false }, 'a second run does nothing')
  // A crash after the manifest but before the old files were removed: the rerun only cleans up.
  fs.writeFileSync(path.join(root, 'abc-123.docx'), docx)
  fs.writeFileSync(path.join(userData, 'recoveries.json'), JSON.stringify([{ id: 'abc-123', title: 'Report', sourcePath: null, updatedAt: 1 }]))
  assert.deepEqual(await store.migrateDocsLegacy({ userData }), { migrated: 0, skipped: 1, indexRemoved: true })
  assert.equal(fs.existsSync(path.join(root, 'abc-123.docx')), false)
  assert.ok((await store.read('abc-123')).parts[0].data.equals(docx), 'the migrated copy is untouched')
})

test('pruning removes old offered entries, then the oldest offered ones over the size cap, never unoffered work', async () => {
  const folder = scenario('prune')
  const { root, here, later } = storesFor(folder)
  const day = 24 * 60 * 60 * 1000
  for (const [index, id] of ['old-1', 'old-2', 'new-3'].entries()) {
    assert.equal((await here.write({ docId: id, revision: 1, parts: [{ name: 'data.bin', data: Buffer.alloc(4000, index) }] })).ok, true)
  }
  const setUpdated = (id, iso) => {
    const file = path.join(root, id, 'manifest.json')
    const manifest = JSON.parse(fs.readFileSync(file, 'utf8'))
    fs.writeFileSync(file, JSON.stringify({ ...manifest, updatedAt: iso }))
  }
  setUpdated('old-1', new Date(Date.now() - 40 * day).toISOString())
  setUpdated('old-2', new Date(Date.now() - 35 * day).toISOString())
  // Nothing was offered yet: age alone never removes work.
  assert.deepEqual((await later.prune()).removed, [])
  assert.equal((await later.list()).length, 3, 'listing offers every orphaned entry')
  const byAge = await later.prune()
  assert.deepEqual(byAge.removed.sort(), ['old-1', 'old-2'])
  assert.ok(fs.existsSync(path.join(root, 'new-3')))

  for (const id of ['cap-a', 'cap-b']) assert.equal((await here.write({ docId: id, revision: 1, parts: [{ name: 'data.bin', data: Buffer.alloc(4000, 7) }] })).ok, true)
  setUpdated('cap-a', new Date(Date.now() - 3 * day).toISOString())
  setUpdated('cap-b', new Date(Date.now() - 2 * day).toISOString())
  await later.list()
  const byCap = await later.prune({ maxTotalBytes: 12_000 })
  assert.deepEqual(byCap.removed, ['cap-a'], 'the oldest offered entry goes first')
  assert.ok(byCap.totalBytes <= 12_000)
  assert.ok(fs.existsSync(path.join(root, 'cap-b')) && fs.existsSync(path.join(root, 'new-3')))
})

test('an unreadable manifest is moved to _damaged and removed after a week', async () => {
  const folder = scenario('damaged')
  const { root, here, later } = storesFor(folder)
  assert.equal((await here.write(snapshot(1))).ok, true)
  fs.writeFileSync(path.join(root, 'doc-1', 'manifest.json'), '{ "schema": 2, "docId": ')
  assert.deepEqual(await later.list(), [])
  const damaged = fs.readdirSync(path.join(root, '_damaged'))
  assert.equal(damaged.length, 1)
  assert.match(damaged[0], /^doc-1-\d+$/)
  assert.equal(fs.existsSync(path.join(root, 'doc-1')), false)
  assert.equal((await later.prune({ now: Date.now() + 8 * 24 * 60 * 60 * 1000 })).damagedRemoved, 1)
})

test('a damaged newest part falls back to the previous generation', async () => {
  const folder = scenario('fallback')
  const { root, here } = storesFor(folder)
  await here.write({ docId: 'doc-1', revision: 1, parts: [{ name: 'state.json', data: 'ONE' }] })
  await here.write({ docId: 'doc-1', revision: 2, parts: [{ name: 'state.json', data: 'TWO' }] })
  fs.writeFileSync(path.join(root, 'doc-1', 'state.json.2'), 'TW0')
  const read = await here.read('doc-1')
  assert.equal(read.ok, true)
  assert.equal(read.fellBack, true)
  assert.equal(read.revision, 1)
  assert.equal(partText(read, 'state.json'), 'ONE')
  fs.writeFileSync(path.join(root, 'doc-1', 'state.json.1'), ' NE')
  assert.equal((await here.read('doc-1')).code, 'DAMAGED')
})

test('a crashed window\'s entries are offered by this process, and a restored entry is claimed', async () => {
  const folder = scenario('orphans')
  const { root, here } = storesFor(folder)
  await here.write(snapshot(3))
  assert.deepEqual(await here.list(), [])
  await here.markOrphaned(['doc-1'])
  const listed = await here.list()
  assert.equal(listed.length, 1)
  assert.equal(listed[0].orphaned, true)
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'doc-1', 'manifest.json'), 'utf8'))
  assert.equal(manifest.state, 'orphaned', 'the orphaned state survives a restart')
  assert.equal(await here.claim('doc-1'), true)
  assert.deepEqual(await here.list(), [], 'a claimed entry is no longer offered')
  assert.ok(here.lastWriteFor(['doc-1']) > 0)
  assert.equal(here.lastWriteFor(['other']), null)
})

test('an entry whose required base file changed is listed as not restorable', async () => {
  const folder = scenario('base-changed')
  const { here, later } = storesFor(folder)
  const base = path.join(folder, 'Scan.pdf')
  fs.writeFileSync(base, '%PDF-1.7 original')
  const stamp = await core.stampFile(base)
  await here.write({
    docId: 'doc-1', revision: 2, title: 'Scan.pdf', sourcePath: base, sourceStamp: stamp,
    parts: [{ name: 'state.json', data: '{}' }, { name: 'base', ref: { path: base, stamp }, required: true }],
  })
  fs.writeFileSync(base, '%PDF-1.7 changed by another program')
  const [entry] = await later.list()
  assert.equal(entry.sourceState, 'changed')
  assert.equal(entry.sourceChanged, true)
  assert.equal(entry.restorable, false)
  assert.equal(entry.reason, 'base-changed')
  fs.rmSync(base)
  const [missing] = await later.list()
  assert.equal(missing.sourceMissing, true)
  assert.equal(missing.reason, 'base-missing')
})

test('invalid ids, part names and oversized meta are refused before anything is written', async () => {
  const folder = scenario('validation')
  const { root, here } = storesFor(folder)
  await assert.rejects(() => here.write({ ...snapshot(1), docId: '../escape' }), TypeError)
  await assert.rejects(() => here.write({ ...snapshot(1), parts: [{ name: 'manifest.json', data: 'x' }] }), TypeError)
  await assert.rejects(() => here.write({ ...snapshot(1), parts: [{ name: 'a', data: 'x' }, { name: 'A', data: 'y' }] }), TypeError)
  await assert.rejects(() => here.write({ ...snapshot(1), parts: [{ name: 'ref', ref: { path: 'relative.pdf' } }] }), TypeError)
  await assert.rejects(() => here.write({ ...snapshot(1), meta: { blob: 'x'.repeat(300 * 1024) } }), TypeError)
  await assert.rejects(() => here.write({ ...snapshot(-1) }), TypeError)
  assert.throws(() => here.discard('doc-1', -2), TypeError)
  assert.equal(fs.existsSync(root), false)
})
