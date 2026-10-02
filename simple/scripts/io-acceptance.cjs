'use strict'

/**
 * I/O acceptance gate (design §10.3; `npm run test:io`).
 *
 * Always runs:
 *   contract  Export As with "Open the file after exporting": the real
 *             ExportModel writes through safeWriteFile and hands the file to
 *             the real io:open-in-simple handler, which starts Simple's own
 *             executable for it. Electron's shell.openPath / openExternal are
 *             never called.
 *   harness   (Windows) the unpackaged app starts with an isolated profile, the
 *             QA driver runs inside it and reports back, and the process tree
 *             is stopped.
 *
 * Then, for every workspace ENABLED in simple/shared/manifest.json, the UI
 * scenarios run in the real app with that workspace's adapter
 * (scripts/io-adapters/<mode>.cjs):
 *   1 open, edit, save        Ctrl+S writes the file in place; it validates; no recovery copy is left
 *   2 locked                  another program holds the file (FileShare.Read); the LOCKED prompt
 *                             appears; Save As writes the same name in Documents; the original is untouched
 *   3 killed                  edit, wait, kill; relaunch offers 1 recovery entry; Restore + Ctrl+S saves the edit
 *   4 close while saving      a slow write (SIMPLE_QA_WRITE_DELAY) is never cut short; no prompt; no temp left
 *   5 drop                    a dropped file keeps its path; Ctrl+S saves in place without a dialog
 *   6 export                  Export As in a remembered format, "open after" goes through openInSimple only
 *
 *   node scripts/io-acceptance.cjs [--module=calc] [--keep] [--contract-only]
 * --contract-only runs only the Electron-free contract (npm test runs it that way).
 * Exit code 0 when every scenario passed.
 */

const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { registerHooks } = require('node:module')
const { fileURLToPath, pathToFileURL } = require('node:url')
const harness = require('./io-harness.cjs')

const ROOT = path.resolve(__dirname, '..')
const RENDERER = pathToFileURL(path.join(ROOT, 'shared', 'renderer') + path.sep).href
const SETUP_WORDING = /\b(?:download|install|installing)\b/i

// ---------------------------------------------------------------------------
// Contract: Export As → openInSimple, without Electron
// ---------------------------------------------------------------------------

let hooksInstalled = false
/** Lets Node load the shared renderer TypeScript unchanged (as the renderer tests do). */
function installRendererHooks() {
  if (hooksInstalled) return
  hooksInstalled = true
  registerHooks({
    resolve(specifier, context, nextResolve) {
      if (/^\.\.?\//.test(specifier) && !/\.[a-z0-9]+$/i.test(specifier) && context.parentURL && context.parentURL.startsWith(RENDERER)) {
        return nextResolve(`${specifier}.ts`, context)
      }
      return nextResolve(specifier, context)
    },
    load(url, context, nextLoad) {
      if (url.startsWith(RENDERER) && url.endsWith('.json')) {
        return { format: 'module', source: `export default ${fs.readFileSync(fileURLToPath(url), 'utf8')};`, shortCircuit: true }
      }
      if (url.startsWith(RENDERER) && url.endsWith('.ts')) {
        return { format: 'module-typescript', source: fs.readFileSync(fileURLToPath(url), 'utf8'), shortCircuit: true }
      }
      return nextLoad(url, context)
    },
  })
}

/**
 * A stand-in for Electron in this Node process: shell calls are recorded and
 * fail the scenario; app answers like the unpackaged unified app.
 */
function fakeElectron(record) {
  const refuse = (name) => (...args) => {
    record.shell.push({ call: name, args: args.map(String) })
    return Promise.resolve('')
  }
  return {
    app: { getAppPath: () => ROOT, isPackaged: false, on() {}, removeListener() {}, getPath: () => os.tmpdir() },
    shell: { openPath: refuse('openPath'), openExternal: refuse('openExternal'), openItem: refuse('openItem'), showItemInFolder: (target) => record.shownInFolder.push(String(target)) },
    BrowserWindow: { getAllWindows: () => [], fromWebContents: () => null },
    dialog: {},
    clipboard: {},
  }
}

/**
 * Runs the export contract scenario.
 * @param {{work: string}} run
 * @returns {Promise<string>} a one-line summary
 */
async function exportContract(run) {
  const record = { shell: [], shownInFolder: [], launched: [] }
  const electronPath = require.resolve('electron')
  const previous = require.cache[electronPath]
  require.cache[electronPath] = { id: electronPath, filename: electronPath, loaded: true, exports: fakeElectron(record) }
  const ioIpc = require('../shared/electron/io-ipc.cjs')
  const { safeWriteFile } = require('../shared/electron/safe-write.cjs')
  const formats = require('../shared/electron/formats.cjs')
  const handlers = new Map()
  const ipcMain = { handle: (channel, handler) => handlers.set(channel, handler), removeHandler: (channel) => handlers.delete(channel) }
  const registration = ioIpc.registerSharedIo({
    ipcMain,
    module: 'calc',
    trust: () => true,
    unified: true,
    maintenance: false,
    launch: (command, args) => record.launched.push({ command, args }),
  })
  try {
    installRendererHooks()
    const { ExportModel } = await import(new URL('export-model.ts', RENDERER).href)
    const event = { sender: { id: 1 }, senderFrame: null }
    const call = (channel, ...args) => handlers.get(channel)(event, ...args)
    const target = path.join(run.work, 'Budget.pdf')
    const prefs = new Map()
    const io = {
      version: 1,
      module: 'calc',
      prefs: { get: async (key) => prefs.get(key), set: async (key, value) => { prefs.set(key, value); return true } },
      chooseSavePath: async () => ({ path: target, format: 'pdf' }),
      prompt: async () => 'cancel',
      shell: { showItem: async (filePath) => call('io:shell-show-item', filePath), openPath: async (filePath) => call('io:open-in-simple', filePath) },
      openInSimple: (filePath) => call('io:open-in-simple', filePath),
      capabilities: async () => call('io:capabilities'),
      onCapabilitiesChanged: () => () => {},
    }
    const rows = formats.exportFormats('calc', { engine: false })
    const model = new ExportModel({
      io,
      module: 'calc',
      documentName: 'Budget.xlsx',
      documentPath: path.join(run.work, 'Budget.xlsx'),
      formats: rows,
      exporter: async (request) => {
        const pdf = Buffer.from('%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Kids[]/Count 0>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF\n', 'latin1')
        const written = await safeWriteFile(request.path, pdf)
        return { ok: true, path: written.path, name: path.basename(written.path), format: request.format, warnings: [] }
      },
      notify: () => {},
    })
    await model.load()
    assert.equal(model.select('pdf'), true, 'PDF must be offered without the office engine')
    model.setOpenAfter(true)
    const outcome = await model.run()
    assert.equal(outcome.ok, true, JSON.stringify(outcome))
    assert.equal(outcome.opened, true, 'the exported file must be opened')
    assert.ok(fs.existsSync(target), 'the exported file exists')
    assert.equal(record.launched.length, 1, 'open after export starts Simple for the file')
    assert.ok(record.launched[0].args.includes(target), `Simple is started with the exported file: ${JSON.stringify(record.launched)}`)
    assert.deepEqual(record.shell, [], 'nothing is handed to another program (shell.openPath / openExternal)')
    assert.equal(prefs.get('openAfterExport'), true, '"open after exporting" is remembered')
    return `export → ${path.basename(target)} → openInSimple → Simple relaunched with the file; shell untouched`
  } finally {
    registration.dispose()
    if (previous) require.cache[electronPath] = previous
    else delete require.cache[electronPath]
  }
}

// ---------------------------------------------------------------------------
// Workspace scenarios (enabled workspaces only)
// ---------------------------------------------------------------------------

function noSetupWording(launch, label) {
  for (const entry of launch.dialogLog) {
    assert.ok(!SETUP_WORDING.test(String(entry.message || '')), `${label}: a prompt asks to download or install something: ${entry.message}`)
  }
}

function stepOk(launch, label) {
  assert.equal(launch.timedOut, false, `${label} timed out. ${launch.output.slice(-1500)}`)
  assert.ok(launch.result, `${label}: the driver wrote no result. ${launch.output.slice(-1500)}`)
  assert.equal(launch.result.ok, true, `${label}: ${launch.result.error}`)
  noSetupWording(launch, label)
  return launch.result.value
}

function recoveryEntries(run) {
  try { return fs.readdirSync(run.recovery, { withFileTypes: true }).filter((entry) => entry.isDirectory() && !entry.name.startsWith('_')).map((entry) => entry.name) } catch { return [] }
}

/** The shared structural validator for a written file (format id first, as validators.cjs expects). */
function validates(filePath, format) {
  const { validateFile } = require('../shared/electron/validators.cjs')
  return validateFile(format, filePath)
}

/**
 * Runs scenarios 1–6 for one enabled workspace.
 * @param {string} mode
 * @param {{keep?: boolean}} options
 * @returns {Promise<Array<{name: string, ok: boolean, detail: string}>>}
 */
async function workspaceScenarios(mode, options) {
  const { path: adapterPath, adapter } = harness.loadAdapter(mode)
  const sample = adapter.sample
  const format = adapter.format
  const outcomes = []
  const scenario = async (name, work) => {
    const run = harness.makeRunFolder(`${mode}-${name.replace(/\W+/g, '-')}`)
    try {
      const detail = await work(run)
      outcomes.push({ name: `${mode}: ${name}`, ok: true, detail: detail || '' })
    } catch (error) {
      outcomes.push({ name: `${mode}: ${name}`, ok: false, detail: String((error && error.message) || error) })
    } finally {
      if (!options.keep) await harness.removeRunFolder(run)
    }
  }
  const launch = (run, extra) => harness.launchSimple({ run, mode, adapterPath, timeoutMs: 240_000, ...extra })

  await scenario('open, edit, save', async (run) => {
    const file = harness.copySample(run, sample)
    const before = harness.hashFile(file)
    const result = await launch(run, { files: [file], step: { scenario: 'editSave', marker: 'QA-EDIT-1', watch: file, watchBefore: before } })
    const value = stepOk(result, 'open, edit, save')
    assert.notEqual(harness.hashFile(file), before, 'the file changed')
    const valid = await validates(file, format)
    assert.equal(valid.ok, true, `the saved file validates: ${valid.reason || ''}`)
    if (typeof adapter.containsEdit === 'function') assert.equal(await adapter.containsEdit(file, 'QA-EDIT-1'), true, 'the edit is in the file')
    if (value.status !== null && value.status !== undefined) assert.match(String(value.status), /Saved/, 'the chip says Saved')
    assert.deepEqual(recoveryEntries(run), [], 'no recovery copy is left after the save')
    return `saved in place (${value.status || 'no chip read'})`
  })

  await scenario('locked by another program', async (run) => {
    const file = harness.copySample(run, sample)
    const before = harness.hashFile(file)
    const fallback = path.join(run.documents, path.basename(file))
    const holder = await harness.holdFile(file, 'Read')
    let result
    try {
      result = await launch(run, {
        files: [file],
        step: { scenario: 'editSave', marker: 'QA-EDIT-2', watch: fallback },
        dialogs: { prompts: { 'saveFailed.LOCKED': ['save-as'] }, save: ['default'] },
      })
    } finally {
      await holder.release()
    }
    stepOk(result, 'locked')
    assert.ok(result.dialogLog.some((entry) => entry.type === 'prompt' && entry.key === 'saveFailed.LOCKED'), 'the LOCKED prompt was shown')
    assert.ok(fs.existsSync(fallback), 'Save As wrote the same name in Documents')
    assert.equal(harness.hashFile(file), before, 'the original is untouched')
    return 'LOCKED → Save As → Documents\\same name; original unchanged'
  })

  await scenario('killed, then recovered', async (run) => {
    const file = harness.copySample(run, sample)
    const before = harness.hashFile(file)
    const first = await launch(run, { files: [file], wait: false, step: { scenario: 'editAndWait', marker: 'QA-EDIT-3', waitMs: 4000, keepRunning: true } })
    const started = Date.now()
    while (Date.now() - started < 90_000 && !recoveryEntries(run).length) await new Promise((resolve) => setTimeout(resolve, 250))
    await new Promise((resolve) => setTimeout(resolve, 1000))
    await first.stop()
    await first.done
    assert.equal(recoveryEntries(run).length, 1, 'one recovery entry was written before the kill')
    const second = await launch(run, { step: { scenario: 'restoreSave', watch: file, watchBefore: before } })
    const value = stepOk(second, 'restore')
    assert.equal(value.offered, 1, 'the welcome card offers 1 entry')
    if (value.statusAfterRestore !== null && value.statusAfterRestore !== undefined) assert.match(String(value.statusAfterRestore), /Recovered/, 'the chip says Recovered, not saved yet')
    if (typeof adapter.containsEdit === 'function') assert.equal(await adapter.containsEdit(file, 'QA-EDIT-3'), true, 'the recovered edit is in the file')
    return 'kill → relaunch → Restore → Ctrl+S saved the edit'
  })

  await scenario('close while saving', async (run) => {
    const file = harness.copySample(run, sample)
    const before = harness.hashFile(file)
    const result = await launch(run, {
      files: [file],
      env: { SIMPLE_QA_WRITE_DELAY: '1500' },
      step: { scenario: 'closeWhileSaving', marker: 'QA-EDIT-4', keepRunning: true },
    })
    assert.equal(result.timedOut, false, `the app did not exit after the save. ${result.output.slice(-1500)}`)
    assert.equal(result.result && result.result.ok, true, result.result && result.result.error)
    assert.ok(!result.dialogLog.some((entry) => entry.type === 'prompt'), `no prompt is shown: ${JSON.stringify(result.dialogLog)}`)
    assert.notEqual(harness.hashFile(file), before, 'the save completed before the app exited')
    assert.deepEqual(fs.readdirSync(path.dirname(file)).filter((name) => /^~simple-/.test(name)), [], 'no temp file is left')
    return 'the window closed only after the write; no prompt, no temp'
  })

  await scenario('drop keeps the path', async (run) => {
    if (typeof adapter.drop !== 'function') throw new Error(`the ${mode} adapter must implement drop(win, h, filePath)`)
    const file = harness.copySample(run, sample)
    const before = harness.hashFile(file)
    const result = await launch(run, { step: { scenario: 'dropSave', dropFile: file, marker: 'QA-EDIT-5', watch: file, watchBefore: before } })
    const value = stepOk(result, 'drop')
    if (value.droppedPath) assert.equal(path.resolve(value.droppedPath).toLowerCase(), path.resolve(file).toLowerCase(), 'the dropped file keeps its real path')
    assert.ok(!result.dialogLog.some((entry) => entry.type === 'save'), 'Ctrl+S after a drop needs no Save dialog')
    return 'drop → Ctrl+S saved in place without a dialog'
  })

  await scenario('export, then open in Simple', async (run) => {
    if (typeof adapter.exportAs !== 'function') throw new Error(`the ${mode} adapter must implement exportAs(win, h, {format, openAfter})`)
    const file = harness.copySample(run, sample)
    const exportFormat = adapter.exportFormat || 'pdf'
    const formats = require('../shared/electron/formats.cjs')
    const extension = (formats.formatById(exportFormat).extensions || ['.pdf'])[0]
    const exported = path.join(path.dirname(file), `${path.parse(file).name}${extension}`)
    const result = await launch(run, { files: [file], step: { scenario: 'exportOpen', format: exportFormat, watch: exported }, dialogs: { save: ['default'] } })
    stepOk(result, 'export')
    assert.ok(fs.existsSync(exported), `the export was written: ${exported}`)
    const handOffs = result.result.handOffs
    assert.deepEqual(handOffs.shell, [], 'nothing is handed to another program')
    assert.ok(handOffs.openInSimple.some((filePath) => path.resolve(filePath).toLowerCase() === exported.toLowerCase()), `open after export goes through openInSimple: ${JSON.stringify(handOffs)}`)
    return `exported ${path.basename(exported)}; opened through openInSimple`
  })

  return outcomes
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

/**
 * Runs the acceptance gate.
 * @param {{modules?: string[], keep?: boolean, harnessSmoke?: boolean, contractOnly?: boolean}} [options]
 *   contractOnly: run only the Electron-free export contract (no app is started)
 * @returns {Promise<{ok: boolean, outcomes: Array<{name: string, ok: boolean, detail: string}>, enabled: string[]}>}
 */
async function runAcceptance(options = {}) {
  const outcomes = []
  const record = async (name, work) => {
    try {
      outcomes.push({ name, ok: true, detail: (await work()) || '' })
    } catch (error) {
      outcomes.push({ name, ok: false, detail: String((error && error.stack) || error) })
    }
  }

  const contractRun = harness.makeRunFolder('contract')
  await record('contract: export → open after export goes through openInSimple', () => exportContract(contractRun))
  if (!options.keep) await harness.removeRunFolder(contractRun)
  if (options.contractOnly) return { ok: outcomes.every((outcome) => outcome.ok), outcomes, enabled: [] }

  if (process.platform === 'win32' && options.harnessSmoke !== false) {
    await record('harness: the app starts isolated and the QA driver reports back', async () => {
      const run = harness.makeRunFolder('harness')
      try {
        const result = await harness.launchSimple({ run, mode: 'launcher', step: { scenario: 'smoke' }, timeoutMs: 90_000 })
        const value = stepOk(result, 'harness smoke')
        assert.ok(value.windows >= 1, 'a window opened')
        assert.deepEqual(result.result.handOffs, { openInSimple: [], shell: [] })
        return `launcher window "${value.title}"`
      } finally {
        if (!options.keep) await harness.removeRunFolder(run)
      }
    })
  }

  const enabled = harness.enabledWorkspaces()
  const selected = options.modules && options.modules.length ? enabled.filter((name) => options.modules.includes(name)) : enabled
  for (const mode of selected) {
    if (process.platform !== 'win32') {
      outcomes.push({ name: `${mode}: UI scenarios`, ok: true, detail: 'skipped (Windows only)' })
      continue
    }
    try {
      outcomes.push(...await workspaceScenarios(mode, options))
    } catch (error) {
      outcomes.push({ name: `${mode}: UI scenarios`, ok: false, detail: String((error && error.message) || error) })
    }
  }
  return { ok: outcomes.every((outcome) => outcome.ok), outcomes, enabled }
}

async function main(argv) {
  const options = { modules: [] }
  for (const argument of argv) {
    if (argument === '--keep') options.keep = true
    else if (argument === '--contract-only') options.contractOnly = true
    else if (argument.startsWith('--module=')) options.modules.push(...argument.slice(9).split(',').filter(Boolean))
    else {
      console.error(`Unknown option: ${argument}`)
      return 2
    }
  }
  const result = await runAcceptance(options)
  for (const outcome of result.outcomes) console.log(`${outcome.ok ? 'ok  ' : 'FAIL'} ${outcome.name}${outcome.detail ? `: ${outcome.detail}` : ''}`)
  if (options.contractOnly) console.log('Contract only: no app was started.')
  else console.log(result.enabled.length
    ? `Workspaces checked: ${result.enabled.join(', ')}.`
    : 'No workspace is enabled in simple/shared/manifest.json yet; only the shared scenarios ran.')
  return result.ok ? 0 : 1
}

if (require.main === module) main(process.argv.slice(2)).then((code) => { process.exitCode = code })

module.exports = { exportContract, runAcceptance, workspaceScenarios }
