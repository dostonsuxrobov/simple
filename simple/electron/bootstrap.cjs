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

// The launcher and the text workspaces share one process (see host.cjs): every
// process tree costs about 105 MB before a document opens, so three of them
// open cost that three times over. Workspaces that need the GPU (below) keep a
// process of their own, as does everything with SIMPLE_SEPARATE_PROCESSES=1
// and the self-test.
const GPU_MODES = new Set(['image', 'video'])
const selfTestRun = process.argv.includes('--simple-self-test')
const separate = process.env.SIMPLE_SEPARATE_PROCESSES === '1' || selfTestRun
const isHosted = (name) => !separate && !GPU_MODES.has(name)
const hosted = isHosted(mode)

// A process of its own keeps its own lock and state directory. The shared
// process has one lock ("host"); its workspaces keep their state folders
// (<state>/<mode>) through the host.
app.setPath('userData', path.join(stateRoot, hosted ? 'host' : mode))
app.setName('simple')

// Every workspace is its own process tree, and Chromium's GPU process costs
// about 100 MB in each (measured: 104 MB with it, 20 MB without). Text, grids
// and pages draw just as well in software, so only the workspaces that push
// pixels (image compositing, video decoding) keep it. SIMPLE_FORCE_GPU=1
// turns it back on everywhere.
if (!GPU_MODES.has(mode) && process.env.SIMPLE_FORCE_GPU !== '1') app.disableHardwareAcceleration()

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
  // The start screen lists what was opened, whichever workspace opens it.
  if (grouped.size) {
    require('./recent-files.cjs').recordRecent(stateRoot, [...grouped].flatMap(([name, pathsForMode]) => pathsForMode.map((filePath) => ({ path: path.resolve(filePath), mode: name }))))
  }
  // What this invocation opens in the shared process: every hosted workspace
  // with files, or just this process's own workspace when there are none.
  const hostGroups = []
  if (hosted) {
    for (const [name, pathsForMode] of grouped) {
      if (isHosted(name)) hostGroups.push([name, pathsForMode.map((filePath) => path.resolve(filePath))])
    }
    if (!hostGroups.some(([name]) => name === mode)) hostGroups.unshift([mode, []])
  }
  for (const [otherMode, pathsForMode] of grouped) {
    if (otherMode === mode || (hosted && isHosted(otherMode))) continue
    void handOff.start(otherMode, pathsForMode.map((filePath) => path.resolve(filePath)))
  }
  // Remove the unpacked runtimes of earlier builds (see runtime-cache.cjs), well
  // after start-up so it never competes with opening a file.
  if (app.isPackaged && process.env.PORTABLE_EXECUTABLE_FILE) {
    app.whenReady().then(() => {
      setTimeout(() => { void require('./runtime-cache.cjs').pruneRuntimeCache({ execPath: process.execPath }) }, 20_000)
    })
  }

  if (hosted) {
    // The first instance becomes the shared process; a later one only passes
    // on what it was asked to open (already routed, as absolute paths) and ends.
    // --simple-new (the start screen's New): a new empty window, not the workspace's open one.
    const fresh = process.argv.includes('--simple-new')
    if (!app.requestSingleInstanceLock({ simpleHost: { groups: hostGroups, fresh } })) {
      app.quit()
    } else {
      const { createHost } = require('./host.cjs')
      const host = createHost({
        appRoot: path.join(__dirname, '..'),
        stateRoot,
        // A workspace that could not start is said in a message (the start screen already
        // reported the file as opened), and the host waits for it before the process may end.
        onOpenFailure: async (name, pathsForMode, error) => {
          process.stderr.write(`Could not open ${pathsForMode.length} file(s) in ${name}: ${error && error.message}\n`)
          // A file that did not open is not a recent file.
          for (const filePath of pathsForMode) require('./recent-files.cjs').forgetRecent(stateRoot, filePath)
          const label = formats.data.workspaces?.[name]?.label || name
          await app.whenReady()
          await electron.dialog.showMessageBox({
            type: 'warning', title: 'simple', buttons: ['OK'], noLink: true,
            message: pathsForMode.length ? `Simple couldn't open ${pathsForMode.length === 1 ? path.basename(pathsForMode[0]) : `${pathsForMode.length} files`}.` : `Simple couldn't start ${label}.`,
            detail: `${label} did not start. Try again; if it keeps happening, start Simple again.`,
          })
        },
      })
      host.install()
      const openGroups = (groups, workingDirectory, flags = []) => {
        for (const [name, pathsForMode] of groups) host.open(name, pathsForMode, workingDirectory, flags).catch(() => {})
      }
      app.on('second-instance', (_event, _argv, workingDirectory, additionalData) => {
        const sent = additionalData && typeof additionalData === 'object' && additionalData.simpleHost ? additionalData.simpleHost.groups : null
        const groups = Array.isArray(sent)
          ? sent.filter((group) => Array.isArray(group) && isHosted(group[0]) && (group[0] === 'launcher' || MODES.includes(group[0])) && Array.isArray(group[1]))
            .map(([name, pathsForMode]) => [name, pathsForMode.filter((item) => typeof item === 'string' && path.isAbsolute(item))])
          : []
        const flags = additionalData && typeof additionalData === 'object' && additionalData.simpleHost && additionalData.simpleHost.fresh === true ? ['--simple-new'] : []
        openGroups(groups.length ? groups : [['launcher', []]], typeof workingDirectory === 'string' && workingDirectory ? workingDirectory : undefined, flags)
      })
      app.whenReady().then(() => app.setAppUserModelId(APP_USER_MODEL_ID))
      openGroups(hostGroups, undefined, fresh ? ['--simple-new'] : [])
    }
  } else if (mode === 'launcher') {
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
