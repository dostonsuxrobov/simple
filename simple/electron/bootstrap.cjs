'use strict'

const fs = require('node:fs')
const path = require('node:path')
const electron = require('electron')
const { app, BrowserWindow, clipboard, Menu } = electron
const { filterSecondInstanceArgv, keepOwnPaths, MODES, routeCommandLine } = require('./routing.cjs')
const { launchDetached } = require('./launch.cjs')
const { expandOpenLists } = require('./open-list.cjs')
const { createHandOff } = require('./hand-off.cjs')
const formats = require('../shared/electron/formats.cjs')
const { installContextMenus } = require('./context-menu.cjs')

const APP_USER_MODEL_ID = 'com.simple.unified'
installContextMenus({ app, BrowserWindow, clipboard, Menu })
const selfTest = process.argv.includes('--simple-self-test')
// Honor explicit profiles so isolated checks and managed sessions cannot join
// the user's live workspace. Detached modes inherit the same profile root.
const requestedStateRoot = app.commandLine.getSwitchValue('user-data-dir') || process.env.SIMPLE_USER_DATA_DIR
const stateRoot = requestedStateRoot ? path.resolve(requestedStateRoot) : path.join(app.getPath('appData'), 'simple')
process.env.SIMPLE_USER_DATA_DIR = stateRoot
// A long list of files arrives as --simple-open-list=<file> (the portable
// wrapper cuts command lines near 8,192 characters); read it, delete it and
// route its paths like any others. Lists older than a day that no process
// picked up are removed on every start.
process.argv = expandOpenLists(process.argv, stateRoot)
// Paths are routed once, here. A process started for another workspace gets
// the decision with --simple-mode, and a running workspace receives it with
// the second-instance message, so no process routes the same path twice
// (a content sniff could answer differently from another folder or while
// the file is being written).
const routed = routeCommandLine(process.argv)
const grouped = routed.groups
const mode = routed.mode

// Each internal mode keeps its own lock and state directory. That lets every
// workspace coexist while retaining its own lifecycle and IPC registrations,
// with no duplicate handlers in a shared process.
app.setPath('userData', path.join(stateRoot, mode))
app.setName('simple')

if (selfTest) {
  // Every workspace bundle, shared I/O module and converter must load from
  // this layout (see self-test.cjs); one JSON line reports the result.
  // The hidden print window closing must not end the run early.
  app.on('window-all-closed', () => {})
  app.whenReady().then(async () => {
    let report
    try {
      report = await require('./self-test.cjs').runSelfTest({ root: path.join(__dirname, '..'), print: true })
    } catch (error) {
      report = { healthy: false, checks: [], failed: [`self-test: ${error && error.message}`] }
    }
    const healthy = report.healthy && MODES.every((name) => fs.existsSync(path.join(__dirname, '..', 'modules', name, 'electron', 'main.cjs')))
    process.stdout.write(`${JSON.stringify({ healthy, mode, routes: [...grouped.keys()], checks: report.checks.length, failed: report.failed })}\n`)
    app.exit(healthy ? 0 : 1)
  })
} else {
  // Acceptance tests (scripts/io-acceptance.cjs) drive an unpackaged app
  // through a driver script in the main process. Never in a packaged app.
  if (!app.isPackaged && process.env.SIMPLE_QA_DRIVER) {
    const driver = path.resolve(process.env.SIMPLE_QA_DRIVER)
    app.whenReady().then(() => require(driver)({ app, mode })).catch((error) => {
      process.stderr.write(`QA driver failed: ${error && error.stack}\n`)
      app.exit(3)
    })
  }
  // A single invocation may contain several file types. Keep the first mode in
  // this process and relaunch the same portable executable once per other mode.
  // Each process is created before the workspace module below loads (that
  // module quits this process when its workspace is already running). A
  // failure to start (the EXE was moved while running) never ends this
  // process, which still opens its own files: it is logged and shown in a
  // message once the app is ready, and a quit waits for that message.
  const handOff = createHandOff({
    app,
    // Looked up when a message is shown.
    dialog: { showMessageBox: (options) => electron.dialog.showMessageBox(options) },
    launch: launchDetached,
    label: (name) => formats.data.workspaces?.[name]?.label || name,
  })
  for (const [otherMode, pathsForMode] of grouped) {
    if (otherMode === mode) continue
    void handOff.start(otherMode, pathsForMode.map((filePath) => path.resolve(filePath)))
  }
  // Remove the unpacked runtimes of earlier builds (see runtime-cache.cjs), well
  // after start-up so it never competes with opening a file.
  if (app.isPackaged && process.env.PORTABLE_EXECUTABLE_FILE) {
    app.whenReady().then(() => {
      setTimeout(() => { void require('./runtime-cache.cjs').pruneRuntimeCache({ execPath: process.execPath }) }, 20_000)
    })
  }

  if (mode === 'launcher') {
    require('../launcher/main.cjs')
  } else {
    const moduleMain = path.join(__dirname, '..', 'modules', mode, 'electron', 'main.cjs')
    if (!fs.existsSync(moduleMain)) {
      throw new Error(`The ${mode} module has not been built. Run npm run sync first.`)
    }
    // Modules re-parse argv with their own broader extension lists, so a path
    // routed to another mode must never reach them. The second-instance relay
    // carries the raw command line of the new instance (not its filtered
    // process.argv), so the new instance also sends its routing decision as
    // the lock's additional data. Every listener shares one argv array, so
    // filtering it here runs before the module's own handler sees it. The
    // running workspace never reads the file to decide again.
    const ownPaths = (grouped.get(mode) || []).map((filePath) => path.resolve(filePath))
    process.argv = keepOwnPaths(process.argv, ownPaths)
    const requestLock = app.requestSingleInstanceLock
    if (typeof requestLock === 'function') {
      app.requestSingleInstanceLock = (additionalData) => requestLock.call(app, {
        ...(additionalData && typeof additionalData === 'object' ? additionalData : {}),
        simpleRouting: { mode, paths: ownPaths },
      })
    }
    app.on('second-instance', (_event, argv, workingDirectory, additionalData) => {
      filterSecondInstanceArgv(argv, { mode, workingDirectory, additionalData })
    })
    require(moduleMain)
    // Calc changes the app name synchronously and Docs sets a legacy taskbar ID
    // when ready. Re-apply unified branding after their startup hooks run.
    app.setName('simple')
    app.whenReady().then(() => app.setAppUserModelId(APP_USER_MODEL_ID))
  }
}
