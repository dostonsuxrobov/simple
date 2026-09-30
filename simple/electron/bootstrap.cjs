'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { app, BrowserWindow, clipboard, Menu } = require('electron')
const { explicitMode, groupPathsByMode, MODES, modeForPath, supportedPaths } = require('./routing.cjs')
const { launchDetached } = require('./launch.cjs')
const { installContextMenus } = require('./context-menu.cjs')

const APP_USER_MODEL_ID = 'com.simple.unified'
installContextMenus({ app, BrowserWindow, clipboard, Menu })
const selfTest = process.argv.includes('--simple-self-test')
const incomingPaths = supportedPaths(process.argv)
const grouped = groupPathsByMode(incomingPaths)
const firstMode = grouped.keys().next().value || null
const mode = firstMode || explicitMode(process.argv) || 'launcher'

// Each internal mode keeps its own lock and state directory. That lets every
// workspace coexist while retaining its own lifecycle and IPC registrations,
// with no duplicate handlers in a shared process.
// Honor explicit profiles so isolated checks and managed sessions cannot join
// the user's live workspace. Detached modes inherit the same profile root.
const requestedStateRoot = app.commandLine.getSwitchValue('user-data-dir') || process.env.SIMPLE_USER_DATA_DIR
const stateRoot = requestedStateRoot ? path.resolve(requestedStateRoot) : path.join(app.getPath('appData'), 'simple')
process.env.SIMPLE_USER_DATA_DIR = stateRoot
app.setPath('userData', path.join(stateRoot, mode))
app.setName('simple')

if (selfTest) {
  app.whenReady().then(() => {
    const required = MODES.map((name) => path.join(__dirname, '..', 'modules', name, 'electron', 'main.cjs'))
    const healthy = required.every((filePath) => fs.existsSync(filePath))
    process.stdout.write(`${JSON.stringify({ healthy, mode, routes: [...grouped.keys()] })}\n`)
    app.exit(healthy ? 0 : 1)
  })
} else {
  // A single invocation may contain several file types. Keep the first mode in
  // this process and relaunch the same portable executable once per other mode.
  for (const [otherMode, pathsForMode] of grouped) {
    if (otherMode !== mode) launchDetached(pathsForMode)
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
    // process.argv), and every listener shares one argv array, so splicing it
    // here runs before the module's own handler sees it.
    const foreignPath = (argument) => {
      if (typeof argument !== 'string' || argument.startsWith('-')) return false
      const owner = modeForPath(argument)
      return Boolean(owner) && owner !== mode
    }
    process.argv = process.argv.filter((argument) => !foreignPath(argument))
    app.on('second-instance', (_event, argv) => {
      if (!Array.isArray(argv)) return
      for (let index = argv.length - 1; index >= 0; index -= 1) {
        if (foreignPath(argv[index])) argv.splice(index, 1)
      }
    })
    require(moduleMain)
    // Calc changes the app name synchronously and Docs sets a legacy taskbar ID
    // when ready. Re-apply unified branding after their startup hooks run.
    app.setName('simple')
    app.whenReady().then(() => app.setAppUserModelId(APP_USER_MODEL_ID))
  }
}
