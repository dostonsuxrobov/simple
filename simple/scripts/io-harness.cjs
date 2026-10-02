'use strict'

/**
 * Shared plumbing for the I/O acceptance gates (scripts/io-acceptance.cjs and
 * scripts/no-office-matrix.cjs). It starts the unpackaged unified app the way
 * a user does (`electron simple --simple-mode=<mode> <file>`), with:
 *   - an isolated --user-data-dir, recovery store and save journal in a run folder;
 *   - SIMPLE_QA_DIALOGS answering every native dialog and prompt from a script,
 *     and logging each one;
 *   - SIMPLE_QA_DOCUMENTS_DIR / SIMPLE_QA_DESKTOP_DIR redirecting the fallback folders;
 *   - SIMPLE_QA_DRIVER loading scripts/io-harness-driver.cjs inside the app's
 *     main process, which runs one scenario step and writes a JSON result;
 *   - a hard timeout, after which the whole process tree is stopped
 *     (taskkill /T /F), and the tree is always stopped at the end.
 * Samples are copied from "Simple test examples" into the run folder; the
 * originals are only read.
 *
 * Workspace-specific steps (how to make an edit, click Restore, drop a file,
 * run Export As) come from an adapter, scripts/io-adapters/<mode>.cjs, that a
 * workspace integrator writes when the workspace is enabled in
 * simple/shared/manifest.json. See the integration guide for its contract.
 */

const crypto = require('node:crypto')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { execFile, spawn } = require('node:child_process')
const { loadManifest } = require('./sync-shared.cjs')

const ROOT = path.resolve(__dirname, '..')
const SAMPLES = path.resolve(ROOT, '..', 'Simple test examples')
const DRIVER = path.join(__dirname, 'io-harness-driver.cjs')
const ADAPTERS = path.join(__dirname, 'io-adapters')

/** Workspaces switched on in simple/shared/manifest.json, in manifest order. */
function enabledWorkspaces(sharedRoot) {
  const manifest = loadManifest(sharedRoot)
  return Object.entries(manifest.workspaces).filter(([, workspace]) => workspace.enabled).map(([name]) => name)
}

/**
 * The adapter of an enabled workspace.
 * @param {string} mode
 * @param {string} [folder] adapter folder (tests point it elsewhere)
 * @returns {{path: string, adapter: object}}
 * @throws {Error} naming the file to write when it is missing
 */
function loadAdapter(mode, folder = ADAPTERS) {
  const file = path.join(folder, `${mode}.cjs`)
  if (!fs.existsSync(file)) {
    throw new Error(`${mode} is enabled in simple/shared/manifest.json but has no acceptance adapter. Write ${path.relative(ROOT, file)} (see the shared I/O integration guide).`)
  }
  delete require.cache[require.resolve(file)]
  return { path: file, adapter: require(file) }
}

/** SHA-256 of a file, or null when it cannot be read. */
function hashFile(filePath) {
  try { return crypto.createHash('sha256').update(fs.readFileSync(filePath)).digest('hex') } catch { return null }
}

/**
 * A fresh run folder with documents, desktop, profile and journal folders.
 * @param {string} name label for the folder
 * @param {string} [parent] default: %TEMP%
 */
function makeRunFolder(name, parent = os.tmpdir()) {
  fs.mkdirSync(parent, { recursive: true })
  const root = fs.mkdtempSync(path.join(parent, `simple-io-${name}-`))
  const folders = {}
  for (const key of ['work', 'documents', 'desktop', 'profile', 'journal', 'recovery']) {
    folders[key] = path.join(root, key)
    fs.mkdirSync(folders[key], { recursive: true })
  }
  return { root, ...folders }
}

/**
 * Copies a sample into the run's work folder.
 * @param {{work: string}} run
 * @param {string} sample file name in "Simple test examples"
 * @param {string} [as] name of the copy
 * @returns {string} the copy's path
 */
function copySample(run, sample, as) {
  const from = path.join(SAMPLES, sample)
  if (!fs.existsSync(from)) throw new Error(`The sample "${sample}" is not in ${SAMPLES}.`)
  const to = path.join(run.work, as || sample)
  fs.copyFileSync(from, to)
  return to
}

function stopTree(pid) {
  return new Promise((resolve) => {
    if (!pid) return resolve()
    if (process.platform === 'win32') execFile('taskkill.exe', ['/PID', String(pid), '/T', '/F'], { windowsHide: true }, () => resolve())
    else {
      try { process.kill(-pid, 'SIGKILL') } catch {}
      resolve()
    }
  })
}

/**
 * Starts Simple in a workspace with the QA driver and waits for it to finish.
 *
 * @param {object} options
 * @param {ReturnType<typeof makeRunFolder>} options.run
 * @param {string} options.mode workspace ('pdf', 'calc', 'docs', 'image', 'video') or 'launcher'
 * @param {string[]} [options.files] files passed on the command line
 * @param {object} options.step what the driver runs: {scenario, ...data}; see io-harness-driver.cjs
 * @param {string|null} [options.adapterPath] the workspace adapter, loaded by the driver
 * @param {object} [options.dialogs] SIMPLE_QA_DIALOGS script ({save, open, prompts}); a log file is added
 * @param {Record<string, string>} [options.env] extra environment
 * @param {number} [options.timeoutMs=120000]
 * @param {boolean} [options.wait=true] false: return {child, done} at once (for scenarios that kill the app)
 * @returns {Promise<{code: number|null, timedOut: boolean, output: string, result: object|null, dialogLog: object[]}>}
 */
async function launchSimple(options) {
  const { run, mode } = options
  const id = crypto.randomBytes(4).toString('hex')
  const resultPath = path.join(run.root, `result-${id}.json`)
  const configPath = path.join(run.root, `driver-${id}.json`)
  const dialogsPath = path.join(run.root, `dialogs-${id}.json`)
  const dialogLog = path.join(run.root, `dialogs-${id}.log`)
  fs.writeFileSync(dialogsPath, JSON.stringify({ save: [], open: [], prompts: {}, ...(options.dialogs || {}), log: dialogLog }))
  fs.writeFileSync(configPath, JSON.stringify({ ...options.step, mode, adapterPath: options.adapterPath || null, resultPath, run }))
  const env = {
    ...process.env,
    SIMPLE_QA_DRIVER: DRIVER,
    SIMPLE_QA_DRIVER_CONFIG: configPath,
    SIMPLE_QA_DIALOGS: dialogsPath,
    SIMPLE_QA_DOCUMENTS_DIR: run.documents,
    SIMPLE_QA_DESKTOP_DIR: run.desktop,
    SIMPLE_IO_JOURNAL_DIR: run.journal,
    SIMPLE_RECOVERY_DIR: run.recovery,
    SIMPLE_USER_DATA_DIR: run.profile,
    ...(options.env || {}),
  }
  delete env.ELECTRON_RUN_AS_NODE
  const args = [ROOT, ...(mode === 'launcher' ? [] : [`--simple-mode=${mode}`]), ...(options.files || []), `--user-data-dir=${run.profile}`]
  const child = spawn(require('electron'), args, { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
  let output = ''
  child.stdout.on('data', (data) => { output = (output + data).slice(-8000) })
  child.stderr.on('data', (data) => { output = (output + data).slice(-8000) })
  const timeoutMs = options.timeoutMs || 120_000
  let timedOut = false
  const timer = setTimeout(() => { timedOut = true; void stopTree(child.pid) }, timeoutMs)
  const done = new Promise((resolve) => child.once('exit', resolve)).then(async (code) => {
    clearTimeout(timer)
    await stopTree(child.pid)
    let result = null
    try { result = JSON.parse(fs.readFileSync(resultPath, 'utf8')) } catch {}
    let log = []
    try { log = fs.readFileSync(dialogLog, 'utf8').split(/\r?\n/).filter(Boolean).map((line) => JSON.parse(line)) } catch {}
    return { code, timedOut, output, result, dialogLog: log }
  })
  if (options.wait === false) return { child, done, stop: () => stopTree(child.pid) }
  return done
}

/**
 * Holds a file open from a separate PowerShell process with the given share
 * mode (Read = how Word, Excel and Acrobat hold files), until release().
 * @param {string} filePath
 * @param {'None'|'Read'|'ReadWrite'|'Delete'} [share='Read']
 * @returns {Promise<{release: () => Promise<void>}>}
 */
async function holdFile(filePath, share = 'Read') {
  const script = `$f = [System.IO.File]::Open('${filePath.replaceAll("'", "''")}', 'Open', 'Read', '${share}'); [Console]::Out.WriteLine('held'); [Console]::In.ReadLine() | Out-Null; $f.Close()`
  const child = spawn('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script], { windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] })
  await new Promise((resolve, reject) => {
    const timer = setTimeout(() => reject(new Error('The file holder did not start.')), 20_000)
    child.stdout.on('data', (data) => { if (String(data).includes('held')) { clearTimeout(timer); resolve() } })
    child.once('exit', (code) => { clearTimeout(timer); reject(new Error(`The file holder stopped (${code}).`)) })
  })
  return {
    release: async () => {
      try { child.stdin.end('\n') } catch {}
      await new Promise((resolve) => { const timer = setTimeout(resolve, 3000); child.once('exit', () => { clearTimeout(timer); resolve() }) })
      await stopTree(child.pid)
    },
  }
}

/** Removes a run folder (best effort; Electron may hold files for a moment). */
async function removeRunFolder(run) {
  for (let attempt = 0; attempt < 10; attempt += 1) {
    try {
      fs.rmSync(run.root, { recursive: true, force: true })
      return
    } catch {
      await new Promise((resolve) => setTimeout(resolve, 300))
    }
  }
}

module.exports = {
  ADAPTERS,
  DRIVER,
  SAMPLES,
  copySample,
  enabledWorkspaces,
  hashFile,
  holdFile,
  launchSimple,
  loadAdapter,
  makeRunFolder,
  removeRunFolder,
  stopTree,
}
