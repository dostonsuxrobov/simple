'use strict'

/**
 * Dry run of vendoring the shared Save/Import/Export layer into every
 * workspace, without touching any workspace folder:
 *
 *  1. copies simple/shared to a temporary folder and switches every workspace
 *     on in that copy of manifest.json;
 *  2. builds a temporary tree with each workspace's real electron/preload.cjs
 *     and a generated electron/main.cjs wired the way the integration guide says;
 *  3. vendors into that tree with sync-shared --target-root, then runs --check;
 *  4. runs each preload (sandbox rules, fake electron) and checks that it
 *     exposes window.simpleIO next to the workspace's own bridge;
 *  5. loads every vendored electron/simple-io/*.cjs and parses every .json, and
 *     bundles them with esbuild exactly as sync-build.cjs bundles a workspace's
 *     electron/main.cjs for the unified app;
 *  6. optionally (--types) type-checks src/simple-io/*.ts with each workspace's
 *     own TypeScript and tsconfig (scripts/check-shared-renderer.cjs);
 *  7. optionally (--electron) loads the vendored modules inside Electron's main
 *     process, registers the shared IPC and guards a hidden window per workspace,
 *     once from the raw vendored files and once from the esbuild bundle.
 *
 *   node scripts/check-vendoring.cjs [--types] [--electron] [--module=pdf,calc] [--temp-root <dir>] [--keep]
 *
 * Workspaces are read, never written. Exit codes: 0 clean, 1 problems, 2 setup error.
 */

const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const vm = require('node:vm')
const { execFile, spawn } = require('node:child_process')
const { loadManifest, syncShared } = require('./sync-shared.cjs')

const ROOT = path.resolve(__dirname, '..')
const REPO_ROOT = path.resolve(ROOT, '..')
const SHARED_ROOT = path.join(ROOT, 'shared')

/**
 * The electron/main.cjs an integrated workspace would have, reduced to the
 * shared I/O wiring (see the integration guide).
 * @param {string} name workspace id
 * @param {object} workspace manifest entry
 */
function generatedMain(name, workspace) {
  const has = (file) => workspace.electron.includes(`electron/${file}`)
  const lines = [
    "'use strict'",
    "const { app, BrowserWindow, ipcMain } = require('electron')",
    "const { registerSharedIo, bridgeArguments } = require('./simple-io/io-ipc.cjs')",
    "const { sweep } = require('./simple-io/io-core.cjs')",
  ]
  if (has('document-guard.cjs')) lines.push("const guard = require('./simple-io/document-guard.cjs')", "const stores = require('./simple-io/stores.cjs')")
  if (has('office-engine.cjs')) lines.push("const officeEngine = require('./simple-io/office-engine.cjs')")
  const extras = [has('document-guard.cjs') ? 'guard, stores' : '', has('office-engine.cjs') ? 'officeEngine' : ''].filter(Boolean).join(', ')
  lines.push(
    'app.whenReady().then(async () => {',
    '  await sweep()',
    `  registerSharedIo({ ipcMain, module: '${name}'${extras ? `, ${extras}` : ''} })`,
    `  const win = new BrowserWindow({ show: false, webPreferences: { preload: __dirname + '/preload.cjs', sandbox: true, contextIsolation: true, additionalArguments: bridgeArguments('${name}') } })`,
  )
  if (has('document-guard.cjs')) lines.push('  guard.installWindowGuard(win)')
  lines.push('})', '')
  return lines.join('\n')
}

/**
 * Runs a preload the way Electron's sandbox does (only `electron` can be
 * required) with a fake electron, and returns the names it exposes.
 * @param {string} text preload source
 * @param {string} filename for stack traces
 * @returns {string[]}
 */
function exposedNames(text, filename) {
  const exposed = {}
  const fakeElectron = {
    contextBridge: { exposeInMainWorld: (key, api) => { exposed[key] = api } },
    ipcRenderer: { on() {}, once() {}, send() {}, invoke: async () => null, removeListener() {}, removeAllListeners() {} },
    webUtils: { getPathForFile: () => '' },
    webFrame: { setZoomFactor() {}, setVisualZoomLevelLimits() {} },
  }
  const sandboxRequire = (id) => {
    if (id === 'electron') return fakeElectron
    if (['events', 'timers', 'url'].includes(id)) return require(id)
    throw new Error(`a sandboxed preload cannot require "${id}"`)
  }
  vm.runInNewContext(text, {
    require: sandboxRequire,
    process: { argv: ['--simple-io-module=x'], env: {}, platform: process.platform, versions: process.versions },
    Buffer, console, setTimeout, clearTimeout, URL, TextDecoder, TextEncoder,
    window: {}, document: { addEventListener() {} }, addEventListener() {},
  }, { filename })
  return Object.keys(exposed)
}

/**
 * Bundles a workspace's vendored electron modules with the esbuild options
 * sync-build.cjs uses for electron/main.cjs, so a module that only works
 * unbundled (a runtime path, a dynamic require) fails here.
 * @param {string} electronRoot the temporary <workspace>/electron folder
 * @param {string[]} files vendored .cjs file names
 * @returns {Promise<string>} the bundle's path
 */
async function bundleVendored(electronRoot, files) {
  const esbuild = require('esbuild')
  const entry = path.join(electronRoot, 'simple-io-bundle-entry.cjs')
  const outfile = path.join(electronRoot, 'simple-io-bundle.cjs')
  const lines = files.map((file) => `  ${JSON.stringify(file)}: require('./simple-io/${file}'),`)
  fs.writeFileSync(entry, `'use strict'\nmodule.exports = {\n${lines.join('\n')}\n}\n`)
  await esbuild.build({
    entryPoints: [entry],
    outfile,
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node22',
    external: ['electron', 'mupdf'],
    legalComments: 'none',
    sourcemap: false,
    logLevel: 'silent',
  })
  return outfile
}

function freshRequire(file) {
  delete require.cache[require.resolve(file)]
  return require(file)
}

/** Loads the vendored modules inside Electron's main process (one process, every workspace). */
function electronScenario(config) {
  const fs = require('node:fs')
  const { app, BrowserWindow } = require('electron')
  const result = { workspaces: {} }
  app.on('window-all-closed', () => {})
  app.whenReady().then(async () => {
    try {
      for (const { name, folder, files, guard, bundle } of config.workspaces) {
        const report = { loaded: 0 }
        const base = `${folder}/electron/simple-io/`
        const modules = {}
        for (const file of files) { modules[file] = require(base + file); report.loaded += 1 }
        // The raw vendored files first, then the esbuild bundle the unified app ships.
        const sources = [['channels', (file) => modules[file]]]
        if (bundle) sources.push(['bundledChannels', (file) => require(bundle)[file]])
        for (const [key, load] of sources) {
          const io = load('io-ipc.cjs')
          const handlers = new Map()
          const ipcMain = { handle: (channel, fn) => handlers.set(channel, fn), removeHandler: (channel) => handlers.delete(channel), on() {}, removeListener() {} }
          const options = { ipcMain, module: name, maintenance: false }
          if (guard) {
            options.guard = load('document-guard.cjs')
            options.stores = load('stores.cjs')
          }
          if (files.includes('office-engine.cjs')) options.officeEngine = load('office-engine.cjs')
          const registration = io.registerSharedIo(options)
          report[key] = registration.channels.length
          const sweep = await load('io-core.cjs').sweep()
          report.sweep = sweep.journals
          if (guard) {
            const win = new BrowserWindow({ show: false, webPreferences: { sandbox: true, contextIsolation: true } })
            options.guard.installWindowGuard(win, { appGuard: false })
            win.destroy()
            report.guarded = true
          }
          registration.dispose()
        }
        result.workspaces[name] = report
      }
    } catch (error) {
      result.error = String((error && error.stack) || error)
    } finally {
      fs.writeFileSync(config.resultPath, JSON.stringify(result))
      app.exit(0)
    }
  })
}

async function runInElectron(temp, workspaces, timeoutMs = 120_000) {
  const electronBinary = require('electron')
  const script = path.join(temp, 'electron-scenario.cjs')
  const resultPath = path.join(temp, 'electron-result.json')
  const profile = path.join(temp, 'electron-profile')
  fs.writeFileSync(script, `(${electronScenario.toString()})(${JSON.stringify({ workspaces, resultPath })})\n`)
  const env = { ...process.env, SIMPLE_FORCE_NO_OFFICE: '1', SIMPLE_IO_JOURNAL_DIR: path.join(temp, 'journal'), SIMPLE_RECOVERY_DIR: path.join(temp, 'recovery') }
  delete env.ELECTRON_RUN_AS_NODE
  const child = spawn(electronBinary, [script, `--user-data-dir=${profile}`], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
  let output = ''
  child.stdout.on('data', (data) => { output = (output + data).slice(-4000) })
  child.stderr.on('data', (data) => { output = (output + data).slice(-4000) })
  const stopTree = () => new Promise((resolve) => {
    if (process.platform === 'win32') execFile('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => resolve())
    else { child.kill('SIGKILL'); resolve() }
  })
  let timedOut = false
  const timer = setTimeout(() => { timedOut = true; void stopTree() }, timeoutMs)
  try {
    await new Promise((resolve) => child.once('exit', resolve))
  } finally {
    clearTimeout(timer)
    await stopTree()
  }
  if (timedOut) return { error: `Electron did not finish within ${timeoutMs / 1000} s. ${output}` }
  try { return JSON.parse(fs.readFileSync(resultPath, 'utf8')) } catch { return { error: `Electron wrote no result. ${output}` } }
}

/**
 * Runs the dry run.
 * @param {object} [options]
 * @param {string[]} [options.modules] workspaces to include (default: all in the manifest)
 * @param {boolean} [options.types] also type-check the renderer files per workspace
 * @param {boolean} [options.electron] also load the vendored modules inside Electron
 * @param {string} [options.tempRoot] parent folder of the temporary tree (default: the system temp folder)
 * @param {boolean} [options.keep] keep the temporary tree
 * @param {string} [options.repoRoot] folder holding the real workspaces (read only)
 * @returns {Promise<{ok: boolean, problems: string[], lines: string[], temp: string}>}
 */
async function checkVendoring(options = {}) {
  const repoRoot = path.resolve(options.repoRoot || REPO_ROOT)
  const tempParent = path.resolve(options.tempRoot || os.tmpdir())
  fs.mkdirSync(tempParent, { recursive: true })
  const temp = fs.mkdtempSync(path.join(tempParent, 'simple-vendoring-'))
  const sharedRoot = path.join(temp, 'shared')
  const targetRoot = path.join(temp, 'repo')
  const problems = []
  const lines = []
  try {
    fs.cpSync(SHARED_ROOT, sharedRoot, { recursive: true })
    const manifest = loadManifest(sharedRoot)
    const names = options.modules && options.modules.length ? options.modules : Object.keys(manifest.workspaces)
    for (const name of names) if (!manifest.workspaces[name]) throw new Error(`Unknown module "${name}".`)
    for (const [name, workspace] of Object.entries(manifest.workspaces)) {
      workspace.enabled = names.includes(name)
      if (!workspace.enabled) continue
      const electronRoot = path.join(targetRoot, workspace.folder, 'electron')
      fs.mkdirSync(electronRoot, { recursive: true })
      fs.copyFileSync(path.join(repoRoot, workspace.folder, 'electron', 'preload.cjs'), path.join(electronRoot, 'preload.cjs'))
      fs.writeFileSync(path.join(electronRoot, 'main.cjs'), generatedMain(name, workspace))
    }
    fs.writeFileSync(path.join(sharedRoot, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`)

    const write = syncShared({ sharedRoot, targetRoot })
    if (!write.ok) problems.push(`vendoring: ${write.message}`)
    const check = syncShared({ sharedRoot, targetRoot, check: true })
    if (!check.ok) problems.push(`check after vendoring: ${check.message}`)
    else lines.push(`Vendored ${write.changes.length} files and blocks into ${names.join(', ')}; --check is clean.`)

    const electronRuns = []
    for (const name of names) {
      const workspace = manifest.workspaces[name]
      const base = path.join(targetRoot, workspace.folder)
      const preloadPath = path.join(base, 'electron', 'preload.cjs')
      try {
        const preload = fs.readFileSync(preloadPath, 'utf8')
        const own = exposedNames(fs.readFileSync(path.join(repoRoot, workspace.folder, 'electron', 'preload.cjs'), 'utf8'), `${workspace.folder}/electron/preload.cjs (original)`)
        const exposed = exposedNames(preload, `${workspace.folder}/electron/preload.cjs`)
        if (!exposed.includes('simpleIO')) problems.push(`${workspace.folder}: the vendored preload does not expose simpleIO`)
        const lost = own.filter((key) => !exposed.includes(key))
        if (lost.length) problems.push(`${workspace.folder}: the vendored preload no longer exposes ${lost.join(', ')}`)
        lines.push(`${name}: preload exposes ${exposed.join(', ')}`)
      } catch (error) {
        problems.push(`${workspace.folder}: the vendored preload does not run: ${error.message}`)
      }
      const files = workspace.electron.map((source) => path.basename(source))
      let loaded = 0
      for (const file of files) {
        const full = path.join(base, 'electron', 'simple-io', file)
        try {
          if (file.endsWith('.json')) JSON.parse(fs.readFileSync(full, 'utf8'))
          else freshRequire(full)
          loaded += 1
        } catch (error) {
          problems.push(`${workspace.folder}/electron/simple-io/${file} does not load: ${error.message}`)
        }
      }
      const rendererFiles = workspace.renderer.map((source) => path.basename(source))
      const missingRenderer = rendererFiles.filter((file) => !fs.existsSync(path.join(base, 'src', 'simple-io', file)))
      if (missingRenderer.length) problems.push(`${workspace.folder}/src/simple-io is missing ${missingRenderer.join(', ')}`)
      lines.push(`${name}: ${loaded}/${files.length} electron/simple-io files load in Node, ${rendererFiles.length - missingRenderer.length} renderer files vendored`)
      const cjsFiles = files.filter((file) => file.endsWith('.cjs'))
      let bundle = null
      try {
        bundle = await bundleVendored(path.join(base, 'electron'), cjsFiles)
        const bundled = freshRequire(bundle)
        const empty = cjsFiles.filter((file) => !bundled[file] || !Object.keys(bundled[file]).length)
        if (empty.length) problems.push(`${workspace.folder}: the esbuild bundle exports nothing for ${empty.join(', ')}`)
        else lines.push(`${name}: ${cjsFiles.length} electron/simple-io modules bundle with esbuild (as sync-build does) and load`)
      } catch (error) {
        problems.push(`${workspace.folder}: electron/simple-io does not bundle with esbuild: ${error.message}`)
        bundle = null
      }
      electronRuns.push({ name, folder: base.replaceAll('\\', '/'), files: cjsFiles, guard: files.includes('document-guard.cjs'), bundle: bundle ? bundle.replaceAll('\\', '/') : null })
    }

    if (options.types) {
      const { checkSharedRenderer } = require('./check-shared-renderer.cjs')
      const withRenderer = names.filter((name) => manifest.workspaces[name].renderer.length)
      if (withRenderer.length) {
        const typed = checkSharedRenderer({ modules: withRenderer, sharedRoot, targetRoot: repoRoot, tempRoot: temp })
        lines.push(...typed.message.split('\n'))
        if (!typed.ok) problems.push('src/simple-io does not type-check in every workspace (see above)')
      }
    }

    if (options.electron) {
      const result = await runInElectron(temp, electronRuns)
      if (result.error) problems.push(`Electron: ${result.error}`)
      for (const [name, report] of Object.entries(result.workspaces || {})) {
        lines.push(`${name}: Electron loaded ${report.loaded} modules, registered ${report.channels} channels${report.bundledChannels ? ` (bundled: ${report.bundledChannels})` : ''}${report.guarded ? ', guarded a window' : ''}`)
      }
    }
  } finally {
    if (!options.keep) fs.rmSync(temp, { recursive: true, force: true })
  }
  return { ok: !problems.length, problems, lines, temp: options.keep ? temp : '' }
}

async function main(argv) {
  const options = { modules: [] }
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index]
    if (argument === '--types') options.types = true
    else if (argument === '--electron') options.electron = true
    else if (argument === '--keep') options.keep = true
    else if (argument.startsWith('--module=')) options.modules.push(...argument.slice(9).split(',').filter(Boolean))
    else if (argument === '--temp-root') options.tempRoot = argv[++index]
    else if (argument.startsWith('--temp-root=')) options.tempRoot = argument.slice(12)
    else {
      console.error(`Unknown option: ${argument}`)
      return 2
    }
  }
  try {
    const result = await checkVendoring(options)
    for (const line of result.lines) console.log(line)
    if (result.temp) console.log(`Temporary tree kept at ${result.temp}`)
    if (!result.ok) {
      console.error(result.problems.join('\n'))
      return 1
    }
    console.log('Vendoring dry run: clean.')
    return 0
  } catch (error) {
    console.error((error && error.stack) || error)
    return 2
  }
}

if (require.main === module) main(process.argv.slice(2)).then((code) => { process.exitCode = code })

module.exports = { checkVendoring, exposedNames, generatedMain }
