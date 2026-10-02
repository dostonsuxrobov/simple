// DocumentSession and io-client (simple/shared/renderer) against a fake
// window.simpleIO: revision-based dirty state, coalesced saves, lossy state,
// the pristine shortcut, failure prompts, close requests, the recovery
// journal's cadence (design §3.6, §4.2-§4.5) and the drop and paste helpers
// (§6.5, §6.6). Timers are fake; nothing touches the disk or the network.

import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { registerHooks } from 'node:module'
import { fileURLToPath } from 'node:url'

// The shared renderer files are TypeScript ES modules written for Vite:
// extensionless relative imports and a plain JSON import, inside a package
// whose package.json says "type": "commonjs". These hooks let Node's built-in
// type stripping load them unchanged.
const RENDERER = new URL('../shared/renderer/', import.meta.url)
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (/^\.\.?\//.test(specifier) && !/\.[a-z0-9]+$/i.test(specifier) && context.parentURL?.startsWith(RENDERER.href)) {
      return nextResolve(`${specifier}.ts`, context)
    }
    return nextResolve(specifier, context)
  },
  load(url, context, nextLoad) {
    if (url.startsWith(RENDERER.href) && url.endsWith('.json')) {
      return { format: 'module', source: `export default ${readFileSync(fileURLToPath(url), 'utf8')};`, shortCircuit: true }
    }
    if (url.startsWith(RENDERER.href) && url.endsWith('.ts')) {
      return { format: 'module-typescript', source: readFileSync(fileURLToPath(url), 'utf8'), shortCircuit: true }
    }
    return nextLoad(url, context)
  },
})

const {
  DocumentSession,
  describeRecoveryEntry,
  discardRecoveryEntry,
  findRecoveryForPath,
  listRecoveryEntries,
  nextRequest,
  orphanBannerText,
  recoveryBannerText,
  snapshotSize,
  windowTitleFor,
} = await import(new URL('document-session.ts', RENDERER).href)
const client = await import(new URL('io-client.ts', RENDERER).href)
const { catalogText } = client
const catalog = JSON.parse(readFileSync(new URL('io-catalog.json', RENDERER), 'utf8'))

// ---------------------------------------------------------------------------
// Fakes
// ---------------------------------------------------------------------------

const flush = () => new Promise((resolve) => setImmediate(resolve))

function deferred() {
  let resolve
  let reject
  const promise = new Promise((done, fail) => {
    resolve = done
    reject = fail
  })
  return { promise, resolve, reject }
}

/** A fake clock: timers run only when advance() passes their due time. */
function createClock(start = 1_000_000) {
  let now = start
  let sequence = 0
  const timers = new Map()
  return {
    timers: {
      setTimeout(callback, ms) {
        sequence += 1
        timers.set(sequence, { id: sequence, at: now + Math.max(0, ms), callback })
        return sequence
      },
      clearTimeout(handle) {
        timers.delete(handle)
      },
      now: () => now,
    },
    get now() {
      return now
    },
    async advance(ms) {
      const end = now + ms
      await flush()
      for (;;) {
        let next = null
        for (const timer of timers.values()) {
          if (timer.at <= end && (!next || timer.at < next.at || (timer.at === next.at && timer.id < next.id))) next = timer
        }
        if (!next) break
        timers.delete(next.id)
        now = next.at
        next.callback()
        await flush()
      }
      now = end
      await flush()
    },
  }
}

function createIO(options = {}) {
  const calls = { prompts: [], writes: [], discards: [], shown: [], opened: [], openPath: [], chosen: [] }
  const handlers = new Map()
  const io = {
    version: 1,
    module: options.module || 'calc',
    calls,
    handlers,
    answers: [...(options.answers || [])],
    failWrites: false,
    entries: [],
    payloads: {},
    paths: new Map(),
    clipboardPayload: { files: [] },
    async capabilities() {
      return { officeEngine: { available: false, source: null, path: null, version: null, verified: false, reason: 'not-installed' }, platform: 'win32' }
    },
    onCapabilitiesChanged() {
      return () => {}
    },
    pathForFile(file) {
      return io.paths.get(file) ?? ''
    },
    async chooseSavePath(request) {
      calls.chosen.push(request)
      return null
    },
    async chooseOpenPaths() {
      return []
    },
    async prompt(key, vars) {
      calls.prompts.push({ key, vars })
      return io.answers.length ? io.answers.shift() : 'cancel'
    },
    // Shaped like io:recovery-* in simple/shared/electron/io-ipc.cjs.
    writeResult: null,
    recovery: {
      async write(snapshot) {
        if (io.failWrites) throw new Error('EIO write')
        if (io.writeResult) return io.writeResult
        calls.writes.push(structuredClone(snapshot))
        return { ok: true, docId: snapshot.docId, revision: snapshot.revision, generation: calls.writes.length }
      },
      async list() {
        return io.entries
      },
      async read(id) {
        return io.payloads[id] || { ok: false, code: 'NOT_FOUND' }
      },
      async discard(id, upToRevision) {
        calls.discards.push({ id, upToRevision })
        return { ok: true, removed: true }
      },
    },
    versions: {
      async list() {
        return []
      },
      async open() {},
    },
    clipboard: {
      async read() {
        return io.clipboardPayload
      },
    },
    shell: {
      async showItem(path) {
        calls.shown.push(path)
      },
      // Present only to prove nothing calls it: documents open inside Simple.
      async openPath(path) {
        calls.openPath.push(path)
      },
    },
    prefs: {
      async get() {
        return undefined
      },
      async set() {},
    },
    officeEngine: {
      async status() {
        return { available: false, source: null, path: null, version: null, verified: false, reason: 'not-installed' }
      },
    },
    onRequest(type, handler) {
      handlers.set(type, handler)
      return () => {
        if (handlers.get(type) === handler) handlers.delete(type)
      }
    },
    async openInSimple(path) {
      calls.opened.push(path)
      return { ok: true, mode: 'pdf' }
    },
  }
  return io
}

function saved(request, extra = {}) {
  return { ok: true, path: 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx', name: 'Budget.xlsx', format: 'xlsx', strategy: 'rename', warnings: [], folderChanged: false, ...extra }
}

function failed(code, extra = {}) {
  return { ok: false, code, name: 'Budget.xlsx', path: 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx', technical: `${code} test`, ...extra }
}

function createAdapter(overrides = {}) {
  const adapter = {
    name: 'Budget.xlsx',
    saves: [],
    results: [],
    toasts: [],
    commits: 0,
    blocked: null,
    focused: 0,
    pristine: false,
    snapshotCount: 0,
    snapshotDelay: null,
    snapshotData: null,
    restored: [],
    onCommit: null,
    async commitPendingEdits() {
      adapter.commits += 1
      if (adapter.onCommit) adapter.onCommit()
      if (adapter.blocked) return { blocked: adapter.blocked, focus: () => { adapter.focused += 1 } }
      return undefined
    },
    isPristine() {
      return adapter.pristine
    },
    save(request) {
      adapter.saves.push({ ...request })
      const next = adapter.results.shift()
      if (typeof next === 'function') return next(request)
      return Promise.resolve(next ?? saved(request))
    },
    async snapshotForRecovery() {
      adapter.snapshotCount += 1
      if (adapter.snapshotDelay) await adapter.snapshotDelay()
      return {
        format: 'xlsx',
        extra: { activeSheet: 0 },
        parts: [
          { name: 'workbook.json', data: adapter.snapshotData ?? JSON.stringify({ n: adapter.snapshotCount }), compress: true },
          { name: 'base', ref: 'source', required: false },
        ],
      }
    },
    async restoreFromRecovery(entry, payload) {
      adapter.restored.push({ entry, payload })
    },
    title() {
      return adapter.name
    },
    notify(kind, text, actions) {
      adapter.toasts.push({ kind, text, actions })
    },
    ...overrides,
  }
  return adapter
}

function lastToast(adapter) {
  return adapter.toasts[adapter.toasts.length - 1]
}

function quietSession(adapter, io, options = {}) {
  return new DocumentSession(adapter, io, { recovery: { enabled: false }, log: () => {}, ...options })
}

// ---------------------------------------------------------------------------
// Revision-based dirty state and saving
// ---------------------------------------------------------------------------

test('an edit made while a save runs stays dirty, and the next save writes it', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  const pending = deferred()
  adapter.results.push(() => pending.promise)

  session.noteChange()
  assert.equal(session.dirty, true)
  assert.equal(session.state.status, 'unsaved')
  const saving = session.save()
  await flush()
  assert.equal(adapter.commits, 1, 'pending edits are committed before saving')
  assert.equal(adapter.saves.length, 1)
  assert.equal(adapter.saves[0].revision, 1)
  assert.equal(adapter.saves[0].mode, 'save')
  assert.equal(session.state.saving, true)
  assert.equal(session.state.statusText, catalog.status.saving)

  session.noteChange()
  pending.resolve(saved())
  assert.equal(await saving, false, 'the newer edit is not on disk yet')
  assert.equal(session.savedRevision, 1)
  assert.equal(session.revision, 2)
  assert.equal(session.dirty, true)
  assert.equal(session.state.status, 'unsaved')
  assert.equal(lastToast(adapter).text, catalogText('toast.savedNewerPending'))

  assert.equal(await session.save(), true)
  assert.equal(adapter.saves[1].revision, 2)
  assert.equal(session.dirty, false)
  assert.equal(session.state.status, 'saved')
  assert.equal(lastToast(adapter).text, catalog.toast.saved)
})

test('a second Ctrl+S while a save runs coalesces into one follow-up', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)

  // No edit during the save: the queued saves resolve without another write.
  const first = deferred()
  adapter.results.push(() => first.promise)
  session.noteChange()
  const one = session.save()
  await flush()
  const two = session.save()
  const three = session.save()
  assert.equal(two, three, 'repeated requests share one queued save')
  assert.equal(session.state.queued, true)
  assert.equal(session.state.status, 'willSaveAgain')
  assert.equal(session.state.statusText, catalog.status.willSaveAgain)
  first.resolve(saved())
  assert.deepEqual(await Promise.all([one, two, three]), [true, true, true])
  assert.equal(adapter.saves.length, 1)

  // An edit during the save: exactly one follow-up write, at the newer revision.
  const second = deferred()
  adapter.results.push(() => second.promise)
  session.noteChange()
  const four = session.save()
  await flush()
  session.noteChange()
  const five = session.save()
  const six = session.save()
  second.resolve(saved())
  assert.equal(await four, false)
  assert.deepEqual(await Promise.all([five, six]), [true, true])
  assert.equal(adapter.saves.length, 3)
  assert.equal(adapter.saves[2].revision, session.revision)
  assert.equal(session.dirty, false)
  assert.equal(session.state.queued, false)
})

test('a queued save does not reopen a failure the user just answered', async () => {
  const io = createIO({ answers: ['cancel'] })
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  const pending = deferred()
  adapter.results.push(() => pending.promise)
  session.noteChange()
  const first = session.save()
  await flush()
  const queued = session.save()
  pending.resolve(failed('LOCKED'))
  assert.equal(await first, false)
  assert.equal(await queued, false)
  assert.equal(io.calls.prompts.length, 1)
  assert.equal(session.state.status, 'notSaved')
})

test('the pristine shortcut is offered only while the model equals the bound file', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)

  await session.save('save-as', 'xlsx')
  assert.equal(adapter.saves[0].pristine, true, 'untouched since open')
  assert.equal(adapter.saves[0].format, 'xlsx')

  session.noteChange()
  await session.save()
  assert.equal(adapter.saves[1].pristine, false)

  await session.save('save-copy')
  assert.equal(adapter.saves[2].pristine, true, 'unchanged since the last save')

  session.noteChange()
  adapter.pristine = true // undo back to the saved state
  session.noteChange()
  assert.equal(session.dirty, false)
  await session.save()
  assert.equal(adapter.saves[3].pristine, true)
})

test('Save a Copy writes a copy and leaves dirty state and binding alone', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  session.noteChange()
  adapter.results.push(saved(null, { path: 'C:\\Users\\me\\Desktop\\Budget copy.xlsx', name: 'Budget copy.xlsx', folderChanged: true }))
  assert.equal(await session.save('save-copy'), true)
  assert.equal(session.dirty, true)
  assert.equal(session.savedRevision, 0)
  assert.equal(lastToast(adapter).text, catalogText('toast.savedTo', { folder: 'me › Desktop' }))
  assert.equal(lastToast(adapter).actions[0].label, catalog.toast.actions.showInFolder)
  lastToast(adapter).actions[0].run()
  await flush()
  assert.deepEqual(io.calls.shown, ['C:\\Users\\me\\Desktop\\Budget copy.xlsx'])
})

test('a blocked commit stops the save with the workspace message and focuses the problem', async () => {
  const io = createIO()
  const adapter = createAdapter({})
  const session = quietSession(adapter, io)
  session.noteChange()
  adapter.blocked = 'Apply or cancel the crop first.'
  assert.equal(await session.save(), false)
  assert.equal(adapter.saves.length, 0)
  assert.equal(adapter.focused, 1)
  assert.deepEqual(lastToast(adapter), { kind: 'warning', text: 'Apply or cancel the crop first.', actions: undefined })
  assert.equal(await session.prepare('export'), false)
  assert.equal(session.dirty, true)

  // A commit that applies a pending edit is part of the captured revision.
  adapter.blocked = null
  adapter.onCommit = () => {
    adapter.onCommit = null
    session.noteChange()
  }
  assert.equal(await session.save(), true)
  assert.equal(adapter.saves[0].revision, 2)
  assert.equal(session.dirty, false)
})

test('a commit that throws blocks the save instead of losing the pending edit', async () => {
  const io = createIO()
  const adapter = createAdapter({ async commitPendingEdits() { throw new Error('editor gone') } })
  const session = quietSession(adapter, io)
  session.noteChange()
  assert.equal(await session.save(), false)
  assert.equal(adapter.saves.length, 0)
  assert.equal(lastToast(adapter).text, catalog.notices.BLOCKED)
})

// ---------------------------------------------------------------------------
// Failure prompts
// ---------------------------------------------------------------------------

test('save failures show the catalog prompt and follow the answer', async () => {
  const io = createIO({ answers: ['save-as'] })
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  session.noteChange()
  adapter.results.push(failed('LOCKED', { technical: 'EBUSY rename after 7 attempts' }))
  adapter.results.push(saved(null, { path: 'C:\\Users\\me\\Documents\\Budget.xlsx', folderChanged: true }))
  assert.equal(await session.save(), true)
  assert.equal(io.calls.prompts[0].key, 'saveFailed.LOCKED')
  assert.equal(io.calls.prompts[0].vars.name, 'Budget.xlsx')
  assert.equal(io.calls.prompts[0].vars.technical, 'EBUSY rename after 7 attempts')
  assert.equal(io.calls.prompts[0].vars.folder, 'C:\\Users\\me\\Documents\\Finance')
  assert.equal(adapter.saves[1].mode, 'save-as')
  assert.equal(adapter.saves[1].fallbackReason, 'LOCKED')
  assert.equal(adapter.saves[1].failedPath, 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx')
  assert.equal(adapter.commits, 2, 'pending edits are committed again before each attempt')
  assert.equal(session.state.status, 'saved')
  assert.equal(session.state.lastFailure, null)

  // Changed on disk → Replace sends force; moved away → Save Here Again sends recreate.
  io.answers.push('replace', 'recreate')
  session.noteChange()
  adapter.results.push(failed('CHANGED_ON_DISK'), failed('SOURCE_MISSING'), saved())
  assert.equal(await session.save(), true)
  assert.equal(adapter.saves[3].force, true)
  assert.equal(adapter.saves[4].recreate, true)
  assert.equal(adapter.saves[4].force, undefined)

  // The engine-only format: save the alternative format next to the original.
  io.answers.push('save-alternative')
  session.noteChange()
  adapter.results.push(failed('NEEDS_OFFICE_ENGINE', { altFormat: 'xlsx', altExt: '.xlsx', formatLabel: 'Excel 97–2003 workbook' }), saved())
  assert.equal(await session.save(), true)
  assert.equal(adapter.saves[6].format, 'xlsx')
  assert.equal(adapter.saves[6].mode, 'save')
})

test('a canceled failure prompt keeps the document dirty with "Not saved", and the chip reopens it', async () => {
  const io = createIO({ answers: ['cancel'] })
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  session.noteChange()
  adapter.results.push(failed('READ_ONLY'))
  assert.equal(await session.save(), false)
  assert.equal(session.dirty, true)
  assert.equal(session.state.status, 'notSaved')
  assert.equal(session.state.statusText, catalog.status.notSaved)
  assert.equal(session.state.lastFailure.code, 'READ_ONLY')
  assert.equal(adapter.toasts.length, 0, 'failures are prompts, never toasts')

  io.answers.push('save-as')
  adapter.results.push(saved(null, { folderChanged: true }))
  assert.equal(await session.reopenFailure(), true)
  assert.equal(io.calls.prompts.length, 2)
  assert.equal(io.calls.prompts[1].key, 'saveFailed.READ_ONLY')
  assert.equal(adapter.saves[1].fallbackReason, 'READ_ONLY')
  assert.equal(session.state.status, 'saved')
})

test('canceled dialogs, unknown codes, thrown errors and the original-file button', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  session.noteChange()
  adapter.results.push({ ok: false, code: 'CANCELED' })
  assert.equal(await session.save('save-as'), false)
  assert.equal(io.calls.prompts.length, 0, 'a canceled dialog is not a failure')
  assert.equal(session.state.status, 'unsaved')

  adapter.results.push(() => Promise.reject(new Error('Error invoking remote method')))
  assert.equal(await session.save(), false)
  assert.equal(io.calls.prompts[0].key, 'saveFailed.UNKNOWN')
  assert.match(io.calls.prompts[0].vars.technical, /Error invoking remote method/)

  adapter.results.push({ ok: false, code: 'SOMETHING_NEW' })
  assert.equal(await session.save(), false)
  assert.equal(io.calls.prompts[1].key, 'saveFailed.UNKNOWN')

  adapter.results.push(undefined)
  adapter.results.push(() => Promise.resolve({}))
  assert.equal(await session.save(), true, 'the default fake result is a success')
  session.noteChange()
  assert.equal(await session.save(), false, 'a result without ok is a failure, never a silent success')

  io.answers.push('show-original', 'cancel')
  adapter.results.push(failed('RESTORE_NEEDED', { backupPath: 'C:\\Users\\me\\Documents\\Finance\\~simple-1a2b3c4d.old' }))
  assert.equal(await session.save(), false)
  assert.deepEqual(io.calls.shown, ['C:\\Users\\me\\Documents\\Finance\\~simple-1a2b3c4d.old'])
  assert.equal(io.calls.prompts.filter((prompt) => prompt.key === 'saveFailed.RESTORE_NEEDED').length, 2, 'the prompt returns after showing the original')
  assert.equal(io.calls.openPath.length, 0)
})

test('nextRequest maps every catalog button', () => {
  const request = { mode: 'save', revision: 4, pristine: false, docId: 'd', format: 'xlsx' }
  const failure = { ok: false, code: 'LOCKED', path: 'C:\\a\\b.xlsx', altFormat: 'xlsx' }
  assert.deepEqual(nextRequest('retry', request, failure), request)
  assert.deepEqual(nextRequest('save-as', request, failure), { ...request, mode: 'save-as', fallbackReason: 'LOCKED', failedPath: 'C:\\a\\b.xlsx' })
  assert.equal(nextRequest('save-as', { ...request, mode: 'save-copy' }, failure).mode, 'save-copy')
  assert.equal(nextRequest('replace', request, failure).force, true)
  assert.equal(nextRequest('recreate', request, failure).recreate, true)
  assert.deepEqual(nextRequest('save-other-format', request, failure), { mode: 'save-as', revision: 4, pristine: false, docId: 'd' })
  assert.deepEqual(nextRequest('save-alternative', request, failure), { ...request, format: 'xlsx' })
  assert.equal(nextRequest('cancel', request, failure), null)
  assert.equal(nextRequest('', request, failure), null)
})

// ---------------------------------------------------------------------------
// Lossy state, sibling saves, exports and close requests
// ---------------------------------------------------------------------------

test('a lossy save keeps the lossy state until a full-format save', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = new DocumentSession(adapter, io, { log: () => {}, recovery: { debounceMs: 60_000, maxDelayMs: 60_000 } })
  session.noteChange()
  adapter.results.push(saved(null, {
    path: 'C:\\Users\\me\\Documents\\Finance\\Budget.csv',
    name: 'Budget.csv',
    format: 'csv',
    lossy: { format: 'csv', formatLabel: 'CSV', lost: ['formulas', 'formatting'], fullFormat: 'xlsx', fullLabel: 'Excel workbook' },
  }))
  assert.equal(await session.save('save-as', 'csv'), true)
  assert.equal(session.dirty, false)
  assert.deepEqual([...session.state.lossy.lost], ['formulas', 'formatting'])
  assert.equal(session.state.lossy.fullFormat, 'xlsx')
  assert.equal(session.state.status, 'savedAs')
  assert.equal(session.state.statusText, 'Saved as CSV')
  assert.equal(lastToast(adapter).text, catalogText('toast.savedLossy', { formatLabel: 'CSV', lostShort: 'Formulas and formatting' }))
  assert.equal(io.calls.discards.length, 0, 'the recovery copy keeps the full workbook after a lossy save')

  const answer = await session.closeQuery()
  assert.equal(answer.dirty, false)
  assert.equal(answer.lossy.format, 'csv')
  assert.equal(answer.title, 'Budget.xlsx')

  // Saving again as CSV keeps the state; a full-format save clears it.
  adapter.results.push(saved(null, { format: 'csv', lossy: { format: 'csv', lost: ['formulas'] } }))
  await session.save()
  assert.equal(session.state.lossy.formatLabel, 'CSV')
  adapter.results.push(saved())
  await session.save('save-as', 'xlsx')
  assert.equal(session.state.lossy, null)
  assert.equal(session.state.status, 'saved')
  assert.equal(io.calls.discards.length, 1)
  session.dispose()
})

test('a full-format copy also clears the lossy state', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  session.noteChange()
  adapter.results.push(saved(null, { format: 'csv', lossy: { format: 'csv', lost: ['formulas'] } }))
  await session.save('save-as', 'csv')
  assert.ok(session.state.lossy)
  adapter.results.push(saved(null, { name: 'Budget full.xlsx', folderChanged: false }))
  await session.save('save-copy', 'xlsx')
  assert.equal(session.state.lossy, null)
})

test('sibling saves say which file was written and that the original was not changed', async () => {
  const io = createIO()
  const adapter = createAdapter()
  adapter.name = 'Budget.xls'
  const session = quietSession(adapter, io)
  session.noteChange()
  adapter.results.push(saved(null, { name: 'Budget.xlsx', sibling: true, originalName: 'Budget.xls', originalFormatLabel: 'Excel 97–2003 workbook' }))
  await session.save()
  assert.equal(lastToast(adapter).text, catalogText('toast.savedModernCopy', { siblingName: 'Budget.xlsx', name: 'Budget.xls', formatLabel: 'Excel 97–2003 workbook' }))
  assert.equal(lastToast(adapter).actions[0].label, catalog.toast.actions.showInFolder)
  assert.doesNotMatch(lastToast(adapter).text, /install|download/i)

  session.noteChange()
  adapter.results.push(saved(null, { name: 'photo (edited).png', sibling: true }))
  await session.save()
  assert.equal(lastToast(adapter).text, catalogText('toast.savedSibling', { siblingName: 'photo (edited).png', name: 'Budget.xls' }))
})

test('a save waits for a running export or print instead of being dropped', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = quietSession(adapter, io, { timers: clock.timers, activityWaitMs: 60_000 })
  const exportDone = deferred()
  const exporting = session.track('export', () => exportDone.promise)
  assert.equal(session.state.activity, 'export')
  assert.equal(session.state.status, 'exporting')
  session.noteChange()
  const saving = session.save()
  await flush()
  assert.equal(adapter.saves.length, 0)
  assert.equal(session.state.status, 'waitingForExport')
  assert.equal(session.state.statusText, catalog.status.waitingForExport)
  exportDone.resolve('exported')
  assert.equal(await exporting, 'exported')
  assert.equal(await saving, true)
  assert.equal(adapter.saves.length, 1)

  // A print that never ends delays the save only up to the cap.
  const printing = session.track('print', () => new Promise(() => {}))
  void printing
  session.noteChange()
  const later = session.save()
  await clock.advance(59_999)
  assert.equal(adapter.saves.length, 1)
  assert.equal(session.state.status, 'waitingForPrint')
  await clock.advance(1)
  assert.equal(await later, true)
  assert.equal(adapter.saves.length, 2)
})

test('main requests: close-query, save-now, discard and recovery-flush', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  assert.deepEqual([...io.handlers.keys()].sort(), ['close-query', 'discard', 'recovery-flush', 'save-now'])

  let answer = await io.handlers.get('close-query')()
  assert.equal(answer.dirty, false)
  assert.equal(answer.saving, false)
  assert.equal(answer.kind, 'spreadsheet')
  assert.equal(answer.docId, session.docId)
  assert.equal(adapter.commits, 1, 'close-query commits pending edits first')

  adapter.blocked = 'Finish editing the cell first.'
  answer = await io.handlers.get('close-query')()
  assert.equal(answer.dirty, true, 'uncommitted input counts as unsaved')
  assert.equal(answer.blocked, 'Finish editing the cell first.')
  assert.equal(adapter.toasts.length, 0, 'close-query never shows a toast; the Save prompt follows')
  adapter.blocked = null

  session.noteChange()
  assert.equal(await io.handlers.get('recovery-flush')({ reason: 'session-end', budgetMs: 2000 }).then((result) => result.ok), true)
  assert.equal(io.calls.writes.length, 1, 'flushed right away, without waiting for the debounce')

  // Close while a save runs: the answer says so and the chip says "Finishing save…".
  const pending = deferred()
  adapter.results.push(() => pending.promise)
  const saving = io.handlers.get('save-now')({ mode: 'save-as', format: 'xlsx' })
  await flush()
  assert.equal(adapter.saves[0].mode, 'save-as')
  assert.equal(adapter.saves[0].format, 'xlsx')
  answer = await io.handlers.get('close-query')()
  assert.equal(answer.saving, true)
  assert.equal(session.state.status, 'finishingSave')
  pending.resolve(saved(null, { folderChanged: true }))
  assert.equal(await saving, true)

  session.noteChange()
  await io.handlers.get('discard')()
  assert.deepEqual(io.calls.discards[io.calls.discards.length - 1], { id: session.docId, upToRevision: session.revision })

  session.dispose()
  assert.equal(io.handlers.size, 0)
})

test('the window title shows unsaved changes', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const titles = []
  const session = quietSession(adapter, io, { windowTitle: (title) => titles.push(title) })
  assert.equal(titles[0], 'Budget.xlsx — Simple Spreadsheets')
  session.noteChange()
  session.noteChange()
  assert.deepEqual(titles, ['Budget.xlsx — Simple Spreadsheets', '• Budget.xlsx — Simple Spreadsheets'])
  await session.save()
  assert.equal(titles[titles.length - 1], 'Budget.xlsx — Simple Spreadsheets')
  assert.equal(windowTitleFor('', true, 'Simple PDF'), '• Simple PDF')
})

test('subscribers get a stable state object that changes only when something visible changes', () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  let calls = 0
  const unsubscribe = session.subscribe(() => { calls += 1 })
  const before = session.state
  assert.equal(session.state, before)
  session.noteChange()
  assert.equal(calls, 1)
  const dirtyState = session.state
  session.noteChange()
  session.noteChange()
  assert.equal(calls, 1, 'more edits do not change what the UI shows')
  assert.equal(session.state, dirtyState)
  assert.ok(Object.isFrozen(dirtyState))
  unsubscribe()
  session.setReadOnly(true)
  assert.equal(calls, 1)
  assert.equal(session.state.status, 'readOnly')
})

// ---------------------------------------------------------------------------
// Recovery journal
// ---------------------------------------------------------------------------

test('recovery cadence: 3 s after the last change, at most 30 s while editing continues', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })

  session.noteChange()
  await clock.advance(2999)
  assert.equal(io.calls.writes.length, 0)
  await clock.advance(1)
  assert.equal(io.calls.writes.length, 1)
  const first = io.calls.writes[0]
  assert.equal(first.docId, session.docId)
  assert.equal(first.revision, 1)
  assert.equal(first.title, 'Budget.xlsx')
  assert.equal(first.module, 'calc')
  assert.equal(first.kind, 'spreadsheet')
  assert.equal(first.parts[0].name, 'workbook.json')
  assert.equal(first.parts[1].ref, 'source', 'the base file is referenced, never named by the renderer')
  assert.deepEqual(first.extra, { activeSheet: 0 })
  assert.equal(first.sourcePath, undefined)
  assert.equal(session.lastRecoveryAt, clock.now)

  // A change every second for 40 s: the first unsnapshotted change is written 30 s later.
  const start = clock.now
  for (let second = 0; second < 40; second += 1) {
    session.noteChange()
    await clock.advance(1000)
    if (second === 28) assert.equal(io.calls.writes.length, 1, 'nothing before 30 s of continuous editing')
  }
  assert.equal(io.calls.writes.length, 2)
  assert.equal(io.calls.writes[1].revision, 31)
  assert.ok(clock.now - start >= 30_000)

  // Editing stops: the rest is written 3 s later, then nothing more.
  await clock.advance(3000)
  assert.equal(io.calls.writes.length, 3)
  assert.equal(io.calls.writes[2].revision, session.revision)
  await clock.advance(60_000)
  assert.equal(io.calls.writes.length, 3)

  // A successful save removes the entry up to the saved revision.
  await session.save()
  assert.deepEqual(io.calls.discards, [{ id: session.docId, upToRevision: session.revision }])
  session.dispose()
})

test('recovery copies continue while a save runs', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  const pending = deferred()
  adapter.results.push(() => pending.promise)
  session.noteChange()
  const saving = session.save()
  await flush() // the save has committed and captured revision 1
  session.noteChange()
  await clock.advance(3000)
  assert.equal(io.calls.writes.length, 1, 'not paused by the save')
  assert.equal(io.calls.writes[0].revision, 2)
  pending.resolve(saved())
  assert.equal(await saving, false)
  // Revision 1 was saved; the journal keeps revision 2 because main discards only up to 1.
  assert.deepEqual(io.calls.discards, [{ id: session.docId, upToRevision: 1 }])
  session.dispose()
})

test('the gap between snapshots adapts to how long a snapshot takes', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  adapter.snapshotDelay = () => new Promise((resolve) => clock.timers.setTimeout(resolve, 1000))
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  const start = clock.now
  session.noteChange()
  await clock.advance(4000) // snapshot starts at +3000 and takes 1000 ms
  assert.equal(io.calls.writes.length, 1)
  session.noteChange()
  await clock.advance(start + 10_999 - clock.now)
  assert.equal(adapter.snapshotCount, 1, 'the next snapshot waits 8 × 1000 ms after the last one started')
  await clock.advance(1)
  assert.equal(adapter.snapshotCount, 2)
  await clock.advance(1000)
  assert.equal(io.calls.writes.length, 2)
  session.dispose()
})

test('autosave failures are retried with backoff and shown after three in a row', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  io.failWrites = true
  session.noteChange()
  await clock.advance(3000)
  assert.equal(adapter.snapshotCount, 1)
  assert.equal(session.state.autosave, 'on')
  await clock.advance(3000)
  assert.equal(adapter.snapshotCount, 2)
  await clock.advance(5999)
  assert.equal(adapter.snapshotCount, 2, 'the third try waits for the backoff')
  await clock.advance(1)
  assert.equal(adapter.snapshotCount, 3)
  assert.equal(session.state.autosave, 'failing')
  assert.equal(session.state.status, 'autosaveFailing')
  assert.equal(session.state.statusText, catalog.status.autosaveFailing)

  io.failWrites = false
  await clock.advance(12_000)
  assert.equal(io.calls.writes.length, 1)
  assert.equal(session.state.autosave, 'on')
  assert.equal(session.state.status, 'unsaved')
  session.dispose()
})

test('what main answers to a recovery write decides what happens next', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })

  // Another window holds this id: a failure that is retried with backoff.
  io.writeResult = { ok: false, code: 'ALREADY_OPEN', message: 'This document is shown in another window.' }
  session.noteChange()
  await clock.advance(3000)
  assert.equal(adapter.snapshotCount, 1)
  await clock.advance(3000)
  assert.equal(adapter.snapshotCount, 2, 'retried')

  // Dropped because a save already covered it: done, no retry.
  io.writeResult = { ok: false, docId: session.docId, dropped: 'discarded' }
  await clock.advance(6000)
  assert.equal(adapter.snapshotCount, 3)
  await clock.advance(60_000)
  assert.equal(adapter.snapshotCount, 3)
  assert.equal(session.state.autosave, 'on')

  // Too large for main: autosave turns off for this document.
  io.writeResult = { ok: false, docId: session.docId, code: 'TOO_LARGE', message: 'too large' }
  session.noteChange()
  await clock.advance(3000)
  assert.equal(session.state.autosave, 'off')
  session.noteChange()
  await clock.advance(60_000)
  assert.equal(adapter.snapshotCount, 4)
  session.dispose()

  // A workspace without a recovery store: recovery stops quietly.
  const quiet = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  io.writeResult = { ok: false, code: 'UNSUPPORTED', message: 'This workspace does not keep recovery copies.' }
  quiet.noteChange()
  await clock.advance(3000)
  quiet.noteChange()
  await clock.advance(60_000)
  assert.equal(adapter.snapshotCount, 5)
  assert.equal(quiet.state.autosave, 'on')
  assert.equal(quiet.state.status, 'unsaved')
  quiet.dispose()
})

test('after a lossy save the recovery journal keeps the full model until a full-format save', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  session.noteChange()
  adapter.results.push(saved(null, { format: 'csv', lossy: { format: 'csv', formatLabel: 'CSV', lost: ['formulas'] } }))
  await session.save('save-as', 'csv')
  assert.equal(session.dirty, false)
  await clock.advance(3000)
  assert.equal(io.calls.writes.length, 1, 'the full model is journaled although the CSV is saved')
  await session.flushRecovery('blur')
  assert.equal(io.calls.discards.length, 0, 'a clean but lossy document keeps its recovery copy')
  adapter.results.push(saved())
  await session.save('save-as', 'xlsx')
  assert.deepEqual(io.calls.discards, [{ id: session.docId, upToRevision: 1 }])
  session.dispose()
})

test('a snapshot that never finishes times out instead of stopping autosave', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  let hang = true
  adapter.snapshotDelay = () => (hang ? new Promise(() => {}) : Promise.resolve())
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {}, recovery: { snapshotTimeoutMs: 10_000 } })
  session.noteChange()
  await clock.advance(3000)
  assert.equal(adapter.snapshotCount, 1)
  await clock.advance(9999)
  assert.equal(adapter.snapshotCount, 1)
  hang = false
  await clock.advance(1)
  // Timed out: the loop is free again, and the retry writes the copy.
  await clock.advance(80_000)
  assert.equal(adapter.snapshotCount, 2)
  assert.equal(io.calls.writes.length, 1)
  assert.equal(session.state.autosave, 'on')
  session.dispose()
})

test('autosave turns off for a very large document and says so', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  adapter.snapshotData = 'x'.repeat(64)
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {}, recovery: { maxBytes: 32 } })
  session.noteChange()
  await clock.advance(3000)
  assert.equal(io.calls.writes.length, 0)
  assert.equal(session.state.autosave, 'off')
  assert.equal(session.state.statusText, catalog.status.autosaveOff)
  session.noteChange()
  await clock.advance(60_000)
  assert.equal(adapter.snapshotCount, 1, 'no more snapshots for this document')
  assert.equal(snapshotSize({ parts: [{ name: 'a', data: 'abc' }, { name: 'b', data: new Uint8Array(5) }, { name: 'c', ref: { path: 'x' } }] }), 8)
  session.dispose()
})

test('undoing back to the saved state removes the recovery copy', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  session.noteChange()
  await clock.advance(3000)
  assert.equal(io.calls.writes.length, 1)
  adapter.pristine = true
  session.noteChange()
  await clock.advance(3000)
  assert.equal(io.calls.writes.length, 1)
  assert.deepEqual(io.calls.discards, [{ id: session.docId, upToRevision: 2 }])
  session.dispose()
})

test('flushRecovery writes now, prepare(print) flushes, and a deadline bounds the wait', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  session.noteChange()
  await session.flushRecovery('blur')
  assert.equal(io.calls.writes.length, 1)
  await session.flushRecovery('blur')
  assert.equal(io.calls.writes.length, 1, 'nothing new to write')

  session.noteChange()
  assert.equal(await session.prepare('print'), true)
  assert.equal(io.calls.writes.length, 2)

  session.noteChange()
  adapter.snapshotDelay = () => new Promise(() => {})
  let done = false
  const flushing = session.flushRecovery('session-end', 2000).then(() => { done = true })
  await clock.advance(1999)
  assert.equal(done, false)
  await clock.advance(1)
  await flushing
  assert.equal(done, true)
  session.dispose()
})

test('default timers follow the global clock (node:test mock timers)', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'], now: 0 })
  const io = createIO()
  const adapter = createAdapter()
  const session = new DocumentSession(adapter, io, { log: () => {} })
  session.noteChange()
  t.mock.timers.tick(2999)
  await flush()
  assert.equal(io.calls.writes.length, 0)
  t.mock.timers.tick(1)
  await flush()
  await flush()
  assert.equal(io.calls.writes.length, 1)
  assert.equal(session.lastRecoveryAt, 3000)
  session.dispose()
})

test('restore: the entry becomes this document, dirty and "Recovered, not saved yet"', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const clock = createClock()
  const entry = {
    id: 'old-doc-id',
    title: 'Budget.xlsx',
    module: 'calc',
    sourcePath: 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx',
    revision: 42,
    updatedAt: new Date(2026, 9, 2, 14, 2).toISOString(),
    orphaned: true,
  }
  io.entries = [entry, { id: 'live', title: 'Open now.xlsx', orphaned: false }]
  io.payloads['old-doc-id'] = { entry, parts: [{ name: 'workbook.json', data: '{"sheets":[]}' }] }
  const entries = await listRecoveryEntries(io)
  assert.deepEqual(entries.map((item) => item.id), ['old-doc-id'], 'documents open right now are not offered')

  const session = new DocumentSession(adapter, io, { timers: clock.timers, log: () => {} })
  assert.equal(session.canRestoreHere(), true)
  assert.equal(await session.restore(entries[0]), true)
  assert.equal(adapter.restored[0].payload.parts[0].data, '{"sheets":[]}')
  assert.equal(session.docId, 'old-doc-id')
  assert.equal(session.dirty, true)
  assert.equal(session.state.recovered, true)
  assert.equal(session.state.statusText, catalog.status.recovered)
  assert.equal(session.canRestoreHere(), false)
  await clock.advance(60_000)
  assert.equal(io.calls.writes.length, 0, 'the journal already holds what was restored')

  session.noteChange()
  await clock.advance(3000)
  assert.equal(io.calls.writes[0].docId, 'old-doc-id')
  assert.ok(io.calls.writes[0].revision > 42)

  await session.save()
  assert.deepEqual(io.calls.discards, [{ id: 'old-doc-id', upToRevision: session.revision }])
  assert.equal(session.state.recovered, false)
  assert.equal(session.state.status, 'saved')
  session.dispose()
})

test('a recovery entry that cannot be read is reported, not restored', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const session = quietSession(adapter, io)
  assert.equal(await session.restore({ id: 'gone', title: 'Lost.xlsx' }), false)
  assert.equal(lastToast(adapter).text, catalog.recovery.cannotRestore)
  assert.equal(session.dirty, false)
  assert.equal(session.docId === 'gone', false)
  io.payloads.damaged = { ok: false, code: 'DAMAGED', reason: 'state.json: content does not match its checksum' }
  assert.equal(await session.restore({ id: 'damaged', title: 'Damaged.xlsx' }), false)
  assert.equal(adapter.restored.length, 0)
  io.recovery.read = async () => { throw new Error('IPC failed') }
  assert.equal(await session.restore({ id: 'x', title: 'x' }), false)
})

test('restore accepts the keyed payload main sends and adopts main\'s document id', async () => {
  const io = createIO()
  const adapter = createAdapter()
  const entry = { id: 'entry-1', docId: 'entry-1', title: 'Photo.png', revision: 3, orphaned: true }
  io.payloads['entry-1'] = { ok: true, docId: 'entry-1', entry, revision: 3, generation: 2, fellBack: false, parts: { 'canvas.png': new Uint8Array([137, 80]) }, extra: null }
  const session = quietSession(adapter, io, { docId: 'fresh-doc' })

  // Unsaved edits are never replaced: restore belongs in a new window then.
  session.noteChange()
  assert.equal(await session.restore(entry), false)
  assert.equal(adapter.restored.length, 0)
  await session.save()
  assert.equal(await session.restore(entry), true)
  assert.equal(client.recoveredPart(adapter.restored[0].payload, 'canvas.png', 'bytes').length, 2)
  assert.equal(session.docId, 'entry-1')
  assert.equal(session.dirty, true)
  assert.equal(session.state.status, 'recovered')
})

test('a restored copy stays unsaved even when the adapter calls the restored model pristine', async () => {
  const io = createIO()
  // Like Docs (model === baseline) and Image (an explicit flag): loading the
  // recovered model through the open path resets the adapter's baseline.
  const adapter = createAdapter({
    async restoreFromRecovery(entry, payload) {
      adapter.restored.push({ entry, payload })
      adapter.pristine = true
    },
  })
  const entry = { id: 'rec-1', docId: 'rec-1', title: 'Budget.xlsx', revision: 7, orphaned: true }
  io.payloads['rec-1'] = { ok: true, docId: 'rec-1', entry, parts: [{ name: 'workbook.json', data: '{}' }] }
  const session = quietSession(adapter, io, { docId: 'opened-doc' })
  assert.equal(await session.restore(entry), true)
  assert.equal(adapter.pristine, true)
  assert.equal(session.dirty, true, 'recovered work is not on disk yet')
  assert.equal(session.state.status, 'recovered')
  const answer = await io.handlers.get('close-query')()
  assert.equal(answer.dirty, true, 'closing asks first')
  assert.equal(await session.save(), true)
  assert.equal(adapter.saves[0].pristine, false, 'the save always writes the recovered model')
  assert.equal(session.dirty, false)
  assert.equal(session.state.recovered, false)
  // After a real write the adapter's pristine flag counts again.
  session.noteChange()
  assert.equal(session.dirty, false)
  session.dispose()
})

test('restore takes the entry over in main only after the model was rebuilt', async () => {
  const io = createIO()
  const order = []
  io.recovery.adopt = async (id) => {
    order.push(`adopt:${id}`)
    return { ok: true, docId: id, untitled: false, path: 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx' }
  }
  const adapter = createAdapter({
    async restoreFromRecovery(entry) {
      order.push(`restore:${entry.id}`)
      if (entry.id === 'broken') throw new Error('a part could not be read')
    },
  })
  const entry = { id: 'rec-2', docId: 'rec-2', title: 'Budget.xlsx', revision: 2, orphaned: true }
  io.payloads['rec-2'] = { ok: true, docId: 'rec-2', entry, parts: [{ name: 'workbook.json', data: '{}' }] }
  io.payloads.broken = { ok: true, docId: 'broken', entry: { id: 'broken' }, parts: [] }
  const session = quietSession(adapter, io, { docId: 'opened-doc' })
  assert.equal(await session.restore({ id: 'broken', docId: 'broken', title: 'Broken.xlsx' }), false)
  assert.deepEqual(order, ['restore:broken'], 'a failed restore never binds the entry in main')
  assert.equal(session.docId, 'opened-doc')
  assert.equal(await session.restore(entry), true)
  assert.deepEqual(order, ['restore:broken', 'restore:rec-2', 'adopt:rec-2'])
  assert.equal(session.docId, 'rec-2')

  // If another window took the file meanwhile, the restored content stays, as a new untitled document.
  const second = quietSession(createAdapter(), io, { docId: 'other-doc' })
  io.recovery.adopt = async () => ({ ok: false, code: 'LOCKED', reason: 'open-in-another-window' })
  assert.equal(await second.restore(entry), true)
  assert.notEqual(second.docId, 'rec-2')
  assert.notEqual(second.docId, 'other-doc', 'never saved under the id of what the window showed before')
  assert.equal(second.dirty, true)
  session.dispose()
  second.dispose()
})

test('a Save a Copy over the document\'s own file leaves the document changed', async () => {
  const io = createIO()
  const adapter = createAdapter()
  adapter.pristine = true
  const session = quietSession(adapter, io)
  adapter.results.push((request) => Promise.resolve(saved(request, { ownFileChanged: true })))
  assert.equal(await session.save('save-copy'), true)
  assert.equal(session.dirty, true, 'the file holds the copy now, not the document')
  assert.equal(await session.save(), true)
  assert.equal(adapter.saves[1].pristine, false)
  assert.equal(session.dirty, false)
  session.noteOwnFileChanged()
  assert.equal(session.dirty, true)
  session.dispose()
})

test('another Simple window holding the file gets its own prompt, not "open in another program"', async () => {
  const io = createIO({ answers: ['cancel'] })
  const adapter = createAdapter()
  adapter.results.push(failed('LOCKED', { reason: 'open-in-another-window' }))
  const session = quietSession(adapter, io)
  session.noteChange()
  assert.equal(await session.save(), false)
  assert.equal(io.calls.prompts[0].key, 'saveFailed.OPEN_IN_ANOTHER_WINDOW')
  assert.match(client.catalogText('saveFailed.OPEN_IN_ANOTHER_WINDOW.message', { name: 'Budget.xlsx' }), /another Simple window/)
  session.dispose()
})

test('recovery card rows and banners', async () => {
  const now = new Date(2026, 9, 2, 16, 0)
  const row = describeRecoveryEntry({
    id: 'a',
    title: 'Budget.xlsx',
    sourcePath: 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx',
    updatedAt: new Date(2026, 9, 2, 14, 2).getTime(),
    sourceChanged: true,
  }, now)
  assert.equal(row.title, 'Budget.xlsx')
  assert.match(row.subtitle, /^Unsaved changes from \S/)
  assert.match(row.subtitle, /14|2/)
  assert.equal(row.folder, 'C:\\Users\\me\\Documents\\Finance')
  assert.deepEqual(row.notes, [catalog.recovery.changedSince])
  assert.equal(row.restorable, true)
  assert.equal(row.untitled, false)

  const untitled = describeRecoveryEntry({ id: 'b', title: '', module: 'calc', updatedAt: '2026-09-01T10:00:00Z' }, now)
  assert.equal(untitled.title, 'Untitled spreadsheet')
  assert.equal(untitled.folder, '')
  assert.equal(untitled.untitled, true)
  assert.deepEqual(describeRecoveryEntry({ id: 'c', title: 'x', restorable: false }, now).notes, [catalog.recovery.cannotRestore])
  assert.deepEqual(describeRecoveryEntry({ id: 'd', title: 'x', sourcePath: 'C:\\x.pdf', sourceMissing: true }, now).notes, [catalog.recovery.sourceMissing])

  const io = createIO()
  io.entries = [
    { id: 'older', title: 'A', sourcePath: 'C:\\Docs\\A.xlsx', updatedAt: '2026-10-01T10:00:00Z' },
    { id: 'newer', title: 'B', sourcePath: 'C:\\Docs\\B.xlsx', updatedAt: '2026-10-02T10:00:00Z' },
  ]
  assert.deepEqual((await listRecoveryEntries(io)).map((entry) => entry.id), ['newer', 'older'])
  assert.deepEqual((await listRecoveryEntries(io, { sourcePath: 'c:/docs/a.XLSX' })).map((entry) => entry.id), ['older'])
  assert.equal(findRecoveryForPath(io.entries, 'C:\\DOCS\\B.xlsx').id, 'newer')
  assert.equal(findRecoveryForPath(io.entries, 'C:\\Docs\\C.xlsx'), null)
  assert.match(recoveryBannerText(io.entries[1], now), /^You have unsaved changes to this file from /)
  assert.equal(orphanBannerText(1), catalog.recovery.orphanBannerOne)
  assert.equal(orphanBannerText(3), catalogText('recovery.orphanBanner', { count: 3 }))
  assert.equal(orphanBannerText(0), '')
  assert.equal(await discardRecoveryEntry(io, io.entries[0]), true)
  assert.deepEqual(io.calls.discards, [{ id: 'older', upToRevision: undefined }])

  io.recovery.list = async () => { throw new Error('unreadable') }
  assert.deepEqual(await listRecoveryEntries(io), [])
  assert.deepEqual(await listRecoveryEntries(null), [])
})

test('a session without the bridge still tracks dirty state and never throws', async () => {
  const adapter = createAdapter()
  const session = new DocumentSession(adapter, null, { log: () => {} })
  session.noteChange()
  adapter.results.push(failed('LOCKED'))
  assert.equal(await session.save(), false)
  assert.equal(session.state.status, 'notSaved')
  await session.flushRecovery('blur')
  assert.equal(await session.restore({ id: 'x', title: 'x' }), false)
  session.dispose()
  assert.equal(await session.save(), false)
})

// ---------------------------------------------------------------------------
// io-client
// ---------------------------------------------------------------------------

test('catalog lookup and formatting', () => {
  assert.equal(client.catalogText('status.saving'), catalog.status.saving)
  assert.equal(client.catalogText('status.savedAs', { formatLabel: 'CSV' }), 'Saved as CSV')
  assert.equal(client.catalogText('no.such.key'), '')
  assert.equal(client.formatTemplate('Saved to {folder}.', {}), `Saved to ${catalog.placeholders.folder}.`)
  assert.equal(client.formatTemplate('{unknownPlaceholder}', {}), '{unknownPlaceholder}')
  const prompt = client.catalogPrompt('saveFailed.LOCKED', { name: 'Budget.xlsx' })
  assert.equal(prompt.message, '"Budget.xlsx" is open in another program.')
  assert.deepEqual(prompt.buttons.map((button) => button.id), ['retry', 'save-as', 'cancel'])
  assert.equal(prompt.defaultId, 'retry')
  assert.equal(prompt.cancelId, 'cancel')
  assert.equal(client.catalogPrompt('status.saving'), null)
  assert.equal(client.appName('calc'), 'Simple Spreadsheets')
  assert.equal(client.appName('unknown-module'), catalog.placeholders.appName)
  assert.equal(client.kindName('docs'), 'document')
  assert.equal(client.statusText(null), '')
  assert.equal(client.statusText('exportingPercent', { percent: 40 }), 'Exporting… 40%')
})

test('failures: variables, in-app descriptions and conversion of thrown errors', () => {
  const failure = { ok: false, code: 'RESTORE_NEEDED', path: 'C:\\Docs\\A.pdf', asidePath: 'C:\\Docs\\~simple-1.old', technical: 'EPERM rename', neededBytes: 5 }
  const vars = client.failureVars(failure)
  assert.equal(vars.name, 'A.pdf')
  assert.equal(vars.folder, 'C:\\Docs')
  assert.equal(vars.backupPath, 'C:\\Docs\\~simple-1.old')
  assert.equal(vars.code, 'RESTORE_NEEDED')
  assert.ok(Object.values(vars).every((value) => typeof value === 'string'))

  const open = client.describeFailure({ ok: false, code: 'NOT_FOUND', context: 'open', name: 'Report.pdf', technical: 'ENOENT open' })
  assert.equal(open.message, '"Report.pdf" isn\'t there anymore.')
  assert.equal(open.details, 'Details: ENOENT open')
  assert.equal(client.describeFailure({ ok: false, code: 'NOT_A_CODE' }).message, client.catalogText('saveFailed.UNKNOWN.message', { name: catalog.placeholders.name }))

  assert.equal(client.failureFromError(new Error('boom')).code, 'UNKNOWN')
  assert.equal(client.failureFromError(new Error('boom')).technical, 'boom')
  assert.equal(client.failureFromError(Object.assign(new Error('held'), { code: 'LOCKED' })).code, 'LOCKED')
  assert.equal(client.failureFromError(Object.assign(new Error('stop'), { name: 'AbortError' })).code, 'CANCELED')
  assert.equal(client.failureFromError('text').technical, 'text')
  assert.equal(client.toIoResult(undefined).ok, false)
  assert.equal(client.toIoResult({ ok: true }).ok, false, 'a success needs a path')
  assert.equal(client.toIoResult({ ok: false, code: 'WHAT' }).code, 'UNKNOWN')
  assert.equal(client.isIoCode('DISK_FULL'), true)
  assert.equal(client.isIoCode('CLOUD'), false)
})

test('promptFailure asks main and treats a missing answer as Cancel', async () => {
  const io = createIO({ answers: ['retry'] })
  assert.equal(await client.promptFailure(io, { ok: false, code: 'DISK_FULL', drive: 'drive C:', needed: '2 MB' }), 'retry')
  assert.deepEqual(io.calls.prompts[0], { key: 'saveFailed.DISK_FULL', vars: { code: 'DISK_FULL', drive: 'drive C:', needed: '2 MB' } })
  assert.equal(await client.promptFailure(io, { ok: false, code: 'NOT_FOUND' }), 'cancel')
  assert.equal(io.calls.prompts[1].key, 'saveFailed.UNKNOWN', 'open-only codes fall back to the generic save prompt')
  io.prompt = async () => { throw new Error('window closed') }
  assert.equal(await client.promptFailure(io, { ok: false, code: 'LOCKED' }), 'cancel')
  assert.equal(await client.promptFailure(null, { ok: false, code: 'LOCKED' }), 'cancel')
})

test('getSimpleIO finds the bridge only when the preload exposed it', () => {
  assert.equal(client.getSimpleIO(), null)
  assert.throws(() => client.requireSimpleIO(), /simple-io bridge/)
  const io = createIO()
  globalThis.simpleIO = io
  try {
    assert.equal(client.getSimpleIO(), io)
    assert.equal(client.requireSimpleIO(), io)
    globalThis.simpleIO = { version: 2 }
    assert.equal(client.getSimpleIO(), null)
  } finally {
    delete globalThis.simpleIO
  }
})

test('dropped files keep their real paths', () => {
  const io = createIO()
  const local = new File(['%PDF-1.7'], 'Report.PDF', { type: 'application/pdf' })
  const fromBrowser = new File(['x'], 'image.png', { type: 'image/png' })
  const other = new File(['x'], 'notes.xyz')
  io.paths.set(local, 'C:\\Users\\me\\Downloads\\Report.PDF')
  io.paths.set(other, 'C:\\Users\\me\\notes.xyz')
  const transfer = { types: ['Files'], files: [local, fromBrowser, other] }
  assert.equal(client.dragHasFiles(transfer), true)
  assert.equal(client.dragHasFiles({ types: ['text/plain'], files: [] }), false)
  assert.equal(client.dragHasFiles(null), false)

  const files = client.filesFromTransfer(transfer, io)
  assert.deepEqual(files.map((file) => [file.name, file.path, file.extension]), [
    ['Report.PDF', 'C:\\Users\\me\\Downloads\\Report.PDF', '.pdf'],
    ['image.png', null, '.png'],
    ['notes.xyz', 'C:\\Users\\me\\notes.xyz', '.xyz'],
  ])
  assert.equal(files[0].file, local)
  assert.equal(files[0].size, 8)

  const { accepted, rejected } = client.partitionFiles(files, ['.PDF', '.png'])
  assert.deepEqual(accepted.map((file) => file.name), ['Report.PDF', 'image.png'])
  assert.deepEqual(rejected.map((file) => file.name), ['notes.xyz'])
  assert.equal(client.partitionFiles(files, (file) => file.path !== null).accepted.length, 2)

  assert.equal(client.notOpenedMessage([]), null)
  assert.equal(client.notOpenedMessage(['a.xyz']), catalogText('toast.notOpenedOne', { name: 'a.xyz' }))
  assert.equal(client.notOpenedMessage(['a.xyz', 'b.abc']), '2 files weren\'t opened: a.xyz, b.abc (not supported).')

  assert.equal(client.openPlacement({ dirty: false, onWelcomeScreen: true }), 'this-window')
  assert.equal(client.openPlacement({ dirty: true, onWelcomeScreen: true }), 'new-window')
  assert.equal(client.openPlacement({ dirty: false, onWelcomeScreen: false }), 'new-window')
  assert.equal(client.dropAction({ onInsertTarget: true, dirty: true, onWelcomeScreen: false }), 'insert')
  assert.equal(client.dropAction({ onInsertTarget: false, dirty: false, onWelcomeScreen: true }), 'open')
  assert.equal(client.dropAction({ onInsertTarget: false, dirty: true, onWelcomeScreen: true }), 'open-new-window')
})

test('paste keeps file paths and falls back to the system clipboard', async () => {
  const io = createIO()
  const copied = new File(['a,b'], 'table.csv', { type: 'text/csv' })
  io.paths.set(copied, 'C:\\Data\\table.csv')
  const data = { 'text/plain': 'a,b', 'text/html': '', 'text/rtf': '' }
  const event = { clipboardData: { types: ['Files', 'text/plain'], files: [copied], getData: (type) => data[type] ?? '' } }
  io.clipboardPayload = { files: ['C:\\Ignored\\other.csv'], png: new Uint8Array([137, 80, 78, 71]), html: '<b>x</b>' }
  const pasted = await client.readPaste(event, io)
  assert.deepEqual(pasted.files.map((file) => file.path), ['C:\\Data\\table.csv'])
  assert.equal(pasted.text, 'a,b')
  assert.equal(pasted.html, '<b>x</b>')
  assert.equal(pasted.png.length, 4)

  io.clipboardPayload = { files: ['C:\\Data\\one.xlsx', 'C:\\Data\\two.xlsx'] }
  const fromSystem = await client.readPaste(null, io)
  assert.deepEqual(fromSystem.files.map((file) => [file.name, file.path, file.file]), [
    ['one.xlsx', 'C:\\Data\\one.xlsx', null],
    ['two.xlsx', 'C:\\Data\\two.xlsx', null],
  ])
  io.clipboard.read = async () => { throw new Error('busy') }
  assert.deepEqual((await client.readPaste(event, io)).files.length, 1)
  assert.equal((await client.readPaste(null, null)).text, null)

  assert.deepEqual([...await client.readFileBytes(copied)], [97, 44, 98])
  assert.equal(await client.readFileBytes(copied, 2), null)
})

test('openInSimple hands files to Simple and never to another program', async () => {
  const io = createIO()
  const opened = await client.openInSimple(io, 'C:\\Out\\Budget.pdf')
  assert.equal(client.openedInSimple(opened), true)
  assert.deepEqual(io.calls.opened, ['C:\\Out\\Budget.pdf'])
  io.openInSimple = async (path) => ({ ok: true, action: 'shown-in-folder', mode: null, appName: null, path, reason: 'unsupported' })
  assert.equal(client.openedInSimple(await client.openInSimple(io, 'C:\\Out\\data.json')), false)
  io.openInSimple = async (path) => ({ ok: false, code: 'NOT_FOUND', message: 'gone', path })
  assert.equal(client.openedInSimple(await client.openInSimple(io, 'C:\\Out\\gone.pdf')), false)
  io.openInSimple = async () => undefined
  assert.equal(client.openedInSimple(await client.openInSimple(io, 'C:\\Out\\Budget.pdf')), true)
  delete io.openInSimple
  const fallback = await client.openInSimple(io, 'C:\\Out\\Budget.pdf')
  assert.deepEqual(fallback, { ok: true, action: 'shown-in-folder', shownInFolder: true, path: 'C:\\Out\\Budget.pdf', reason: 'no-bridge' })
  assert.equal(client.openedInSimple({ ok: true, shownInFolder: true }), false)
  assert.deepEqual(io.calls.shown, ['C:\\Out\\Budget.pdf'], 'an older bridge only shows the file in its folder')
  assert.equal(io.calls.openPath.length, 0)
  assert.equal(client.openedInSimple(await client.openInSimple(null, 'C:\\x.pdf')), false)
})

test('recovered parts are read the same whichever shape main sends', () => {
  const keyed = { ok: true, docId: 'd', entry: { id: 'd', title: 'x' }, parts: { 'state.json': new TextEncoder().encode('{"a":1}'), 'note.txt': 'hi' } }
  assert.deepEqual(client.recoveredParts(keyed).map((part) => part.name), ['state.json', 'note.txt'])
  assert.equal(client.recoveredPart(keyed, 'state.json'), '{"a":1}')
  assert.deepEqual([...client.recoveredPart(keyed, 'note.txt', 'bytes')], [104, 105])
  assert.equal(client.recoveredPart(keyed, 'missing.json'), null)
  const listed = { entry: { id: 'd', title: 'x' }, parts: [{ name: 'canvas.png', data: new Uint8Array([1, 2]) }, { name: 'base', ref: { path: 'C:\\a.png', matches: true } }] }
  assert.deepEqual([...client.recoveredPart(listed, 'canvas.png', 'bytes')], [1, 2])
  assert.equal(client.recoveredPart(listed, 'base'), null)
  assert.deepEqual(client.recoveredParts({ ok: false, code: 'NOT_FOUND' }), [])
  assert.deepEqual(client.recoveredParts(null), [])
})

test('path and text helpers', () => {
  assert.equal(client.baseName('C:\\Users\\me\\Budget.xlsx'), 'Budget.xlsx')
  assert.equal(client.baseName('/home/me/a.txt'), 'a.txt')
  assert.equal(client.folderOf('C:\\Users\\me\\Budget.xlsx'), 'C:\\Users\\me')
  assert.equal(client.folderOf('C:\\Budget.xlsx'), 'C:\\')
  assert.equal(client.folderOf('Budget.xlsx'), '')
  assert.equal(client.folderLabel('C:\\Users\\me\\Documents\\Finance'), 'Documents › Finance')
  assert.equal(client.folderLabel('C:\\'), 'C:')
  assert.equal(client.fileExtension('Q3 plan v2.1'), '.1')
  assert.equal(client.fileExtension('archive.TAR.GZ'), '.gz')
  assert.equal(client.fileExtension('.gitignore'), '')
  assert.equal(client.fileExtension('name.'), '')
  assert.equal(client.samePath('C:/Docs/A.xlsx', 'c:\\docs\\a.XLSX'), true)
  assert.equal(client.samePath('C:\\Docs\\A.xlsx', null), false)
  assert.equal(client.joinList(['formulas']), 'formulas')
  assert.equal(client.joinList(['formulas', 'formatting', 'charts']), 'formulas, formatting and charts')
  assert.equal(client.capitalize('formulas'), 'Formulas')
})
