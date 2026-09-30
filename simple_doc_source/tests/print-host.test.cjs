const test = require('node:test')
const assert = require('node:assert/strict')
const { EventEmitter } = require('node:events')
const {
  isStalePrintDirectory,
  loadPdfForPrinting,
  printWebContentsSilently,
  PRINT_DIRECTORY_PREFIX,
  STALE_PRINT_DIRECTORY_AGE_MS,
} = require('../electron/print-host.cjs')

class FakeContents extends EventEmitter {
  isDestroyed() { return false }
}

class FakeWindow extends EventEmitter {
  constructor(load) {
    super()
    this.webContents = new FakeContents()
    this.load = load
  }
  isDestroyed() { return false }
  loadURL(url) { return this.load(url, this) }
}

test('PDF print host waits for explicit PDF parsing readiness and registers before navigation', async () => {
  let navigationResolved = false
  const window = new FakeWindow(async (url, target) => {
    assert.equal(url, 'file:///print.pdf')
    assert.equal(target.webContents.listenerCount('page-title-updated'), 1)
    target.webContents.emit('page-title-updated', {}, 'loading', false)
    await Promise.resolve()
    navigationResolved = true
    target.webContents.emit('page-title-updated', {}, 'print.pdf', true)
  })

  await loadPdfForPrinting(window, 'file:///print.pdf', { timeoutMs: 100 })
  assert.equal(navigationResolved, true)
  assert.equal(window.webContents.listenerCount('page-title-updated'), 0)
  assert.equal(window.webContents.listenerCount('render-process-gone'), 0)
  assert.equal(window.listenerCount('closed'), 0)
})

test('PDF print host fails closed instead of printing before readiness', async () => {
  const window = new FakeWindow(async () => {})
  await assert.rejects(
    loadPdfForPrinting(window, 'file:///never-ready.pdf', { timeoutMs: 15 }),
    /did not finish loading/,
  )
  assert.equal(window.webContents.listenerCount('page-title-updated'), 0)
})

test('stale print cleanup never selects a fresh concurrent print directory', () => {
  const now = Date.now()
  const directory = (mtimeMs) => ({ isDirectory: () => true, mtimeMs })
  assert.equal(isStalePrintDirectory(`${PRINT_DIRECTORY_PREFIX}active`, directory(now - 1_000), now), false)
  assert.equal(isStalePrintDirectory(`${PRINT_DIRECTORY_PREFIX}old`, directory(now - STALE_PRINT_DIRECTORY_AGE_MS - 1), now), true)
  assert.equal(isStalePrintDirectory('another-app-print-old', directory(0), now), false)
  assert.equal(isStalePrintDirectory(`${PRINT_DIRECTORY_PREFIX}file`, { isDirectory: () => false, mtimeMs: 0 }, now), false)
})

test('final print submission requires silent mode and reports driver failures without opening a dialog', async () => {
  let receivedOptions = null
  const contents = {
    isDestroyed: () => false,
    print(options, callback) {
      receivedOptions = options
      callback(false, 'printer offline')
    },
  }
  await assert.rejects(printWebContentsSilently(contents, { silent: false }), /requires silent mode/)
  const result = await printWebContentsSilently(contents, { silent: true, deviceName: 'Office Printer' })
  assert.equal(receivedOptions.silent, true)
  assert.equal(receivedOptions.deviceName, 'Office Printer')
  assert.deepEqual(result, { success: false, failureReason: 'printer offline' })
})

test('a printer that never confirms returns an actionable timeout and ignores a late callback', async () => {
  const contents = new FakeContents()
  let callback
  contents.print = (_options, value) => { callback = value }
  const result = await printWebContentsSilently(contents, { silent: true }, { timeoutMs: 15 })
  assert.equal(result.success, false)
  assert.match(result.failureReason, /Check the print queue/)
  assert.equal(contents.listenerCount('render-process-gone'), 0)
  callback(true)
})

test('renderer crashes release the print job and clean up listeners', async () => {
  const contents = new FakeContents()
  contents.print = () => { contents.emit('render-process-gone', {}, { reason: 'crashed' }) }
  const result = await printWebContentsSilently(contents, { silent: true })
  assert.equal(result.success, false)
  assert.match(result.failureReason, /renderer stopped/)
  assert.equal(contents.listenerCount('destroyed'), 0)
})
