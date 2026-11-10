'use strict'

// One process for several workspaces.
//
// Each workspace module (modules/<mode>/electron/main.cjs, and the launcher)
// was written as a whole Electron app: it takes the single-instance lock,
// registers its IPC channels, keeps its state in userData, and quits when its
// last window closes. Running each one in its own process tree costs about
// 105 MB per workspace before a document is opened (main process 74 MB, GPU
// process 20 MB, network service 11 MB, measured).
//
// The host loads those same modules, unchanged, into one process. Each module
// sees a scoped `electron`: its own windows, its own IPC handlers (two modules
// may use the same channel name; a message is answered by the module that owns
// the sending window), its own userData folder and browser storage, and an
// `app` whose quit() ends only that workspace. When a workspace's last window
// closes it quits as before, and the host forgets the module, so the next file
// of that kind starts it fresh. The process ends when no workspace is left.

const path = require('node:path')
const { EventEmitter } = require('node:events')
const Module = require('node:module')
const electron = require('electron')
const { OPEN_LIST_SWITCH, needsOpenList, writeOpenListSync } = require('./open-list.cjs')

/** App events a workspace owns: the host raises them for that workspace only. */
const SCOPED_APP_EVENTS = new Set(['second-instance', 'window-all-closed', 'before-quit', 'will-quit', 'quit', 'activate'])
/** App events about one window: a workspace hears them for its own windows only. */
const WINDOW_APP_EVENTS = new Set(['browser-window-created', 'browser-window-focus', 'browser-window-blur'])
/**
 * How long a starting workspace may take to open its first window. After that
 * it is taken for a start that failed (see createScope) and the next one starts.
 */
const START_TIMEOUT_MS = Number(process.env.SIMPLE_HOST_START_TIMEOUT_MS) || 20_000

/**
 * Custom URL schemes a hosted workspace declares before the app is ready.
 * A workspace loaded later cannot do that itself, so the host does it at start
 * (the entry must equal the workspace's own protocol.registerSchemesAsPrivileged call).
 */
const PRIVILEGED_SCHEMES = [
  // Documents: installed fonts (simple_doc_source/electron/main.cjs).
  { scheme: 'simple-font', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true } },
]

function createHost({ appRoot, stateRoot, onOpenFailure = null, log = (text) => process.stderr.write(text) }) {
  const { app, ipcMain, session } = electron
  const RealBrowserWindow = electron.BrowserWindow
  /** @type {Map<string, object>} live workspaces by mode */
  const scopes = new Map()
  /** @type {Map<number, object>} webContents id -> workspace */
  const ownerByContents = new Map()
  const handledChannels = new Set()
  const listenedChannels = new Set()
  let loading = Promise.resolve()
  let pending = 0
  let appQuitting = false
  let hookInstalled = false
  // Tests only (tests/bootstrap-handoff.test.cjs), never in a packaged app: their fake
  // workspaces open no window on purpose and must not be taken for failed starts.
  const windowlessAllowed = !app.isPackaged && process.env.SIMPLE_HOST_ALLOW_WINDOWLESS === '1'

  const entryFor = (mode) => (mode === 'launcher'
    ? path.join(appRoot, 'launcher', 'main.cjs')
    : path.join(appRoot, 'modules', mode, 'electron', 'main.cjs'))
  /** Folders whose loaded code belongs to a workspace (forgotten when it ends). */
  const foldersFor = (mode) => (mode === 'launcher' ? [path.join(appRoot, 'launcher')] : [path.join(appRoot, 'modules', mode)])

  const inside = (folder, file) => {
    const relative = path.relative(folder, file)
    return relative !== '' && !relative.startsWith('..') && !path.isAbsolute(relative)
  }

  /** The workspace whose code is in `file` (the launcher also owns simple/shared). */
  function modeOfFile(file) {
    if (typeof file !== 'string') return null
    const modules = path.join(appRoot, 'modules')
    if (inside(modules, file)) return path.relative(modules, file).split(path.sep)[0]
    if (inside(path.join(appRoot, 'launcher'), file) || inside(path.join(appRoot, 'shared'), file)) return 'launcher'
    return null
  }

  function ownerOf(contents) {
    return contents && !contents.isDestroyed?.() ? ownerByContents.get(contents.id) || null : null
  }

  // ---- IPC: one real handler per channel, answered by the sender's workspace ----

  function ensureHandled(channel) {
    if (handledChannels.has(channel)) return
    handledChannels.add(channel)
    ipcMain.handle(channel, (event, ...args) => {
      const scope = ownerOf(event.sender)
      const handler = scope && scope.handlers.get(channel)
      if (!handler) throw new Error(`No handler registered for '${channel}'`)
      if (handler.once) scope.handlers.delete(channel)
      return handler.run(event, ...args)
    })
  }

  function ensureListened(channel) {
    if (listenedChannels.has(channel)) return
    listenedChannels.add(channel)
    ipcMain.on(channel, (event, ...args) => {
      const scope = ownerOf(event.sender)
      if (scope) scope.messages.emit(channel, event, ...args)
    })
  }

  function scopedIpcMain(scope) {
    const listen = (method) => (channel, listener) => { ensureListened(channel); scope.messages[method](channel, listener); return facade }
    const facade = {
      handle(channel, run) {
        if (scope.handlers.has(channel)) throw new Error(`Attempted to register a second handler for '${channel}'`)
        ensureHandled(channel)
        scope.handlers.set(channel, { run, once: false })
      },
      handleOnce(channel, run) {
        if (scope.handlers.has(channel)) throw new Error(`Attempted to register a second handler for '${channel}'`)
        ensureHandled(channel)
        scope.handlers.set(channel, { run, once: true })
      },
      removeHandler(channel) { scope.handlers.delete(channel) },
      on: listen('on'),
      addListener: listen('on'),
      once: listen('once'),
      off(channel, listener) { scope.messages.removeListener(channel, listener); return facade },
      removeListener(channel, listener) { scope.messages.removeListener(channel, listener); return facade },
      removeAllListeners(channel) { if (channel === undefined) scope.messages.removeAllListeners(); else scope.messages.removeAllListeners(channel); return facade },
      listenerCount: (channel) => scope.messages.listenerCount(channel),
      listeners: (channel) => scope.messages.listeners(channel),
    }
    return facade
  }

  // ---- Windows ----

  function scopedBrowserWindow(scope) {
    // Electron's own window lookups (getAllWindows, fromWebContents) only count objects
    // whose class is named BrowserWindow, so the subclass keeps that name.
    class BrowserWindow extends RealBrowserWindow {
      constructor(options = {}) {
        // A workspace the host gave up on (it opened no window in time) has no IPC
        // handlers left and nobody waiting for it: a late window would be a dead one.
        if (scope.failed) throw new Error(`${scope.mode} took too long to start and was ended.`)
        const preferences = { ...(options.webPreferences || {}) }
        // The workspace's own browser storage (what used to be its process's default session).
        if (!preferences.session && !preferences.partition) preferences.session = scope.session
        super({ ...options, webPreferences: preferences })
        const contentsId = this.webContents.id
        ownerByContents.set(contentsId, scope)
        scope.windows.add(this)
        if (process.env.SIMPLE_HOST_DEBUG) log(`host: ${Date.now() % 100000} ${scope.mode} window ${contentsId} created (show=${options.show})\n`)
        this.once('closed', () => {
          ownerByContents.delete(contentsId)
          scope.windows.delete(this)
          if (process.env.SIMPLE_HOST_DEBUG) log(`host: ${Date.now() % 100000} ${scope.mode} window ${contentsId} closed\n`)
          windowClosed(scope)
        })
        scope.markStarted()
      }

      static getAllWindows() {
        return RealBrowserWindow.getAllWindows().filter((win) => scope.windows.has(win))
      }

      static getFocusedWindow() {
        const focused = RealBrowserWindow.getFocusedWindow()
        return focused && scope.windows.has(focused) ? focused : null
      }
    }
    return BrowserWindow
  }

  // ---- app ----

  function scopedApp(scope) {
    const windowListeners = new Map()
    const add = (method) => (name, listener) => {
      // Code of a workspace that already ended (a start that failed late) registers nothing more.
      if (scope.ended) return proxy
      if (SCOPED_APP_EVENTS.has(name)) scope.events[method](name, listener)
      else if (WINDOW_APP_EVENTS.has(name)) {
        const wrapped = (event, win, ...rest) => { if (scope.windows.has(win)) listener(event, win, ...rest) }
        if (!windowListeners.has(listener)) windowListeners.set(listener, new Map())
        windowListeners.get(listener).set(name, wrapped)
        scope.appListeners.push([name, wrapped])
        app[method](name, wrapped)
      } else {
        scope.appListeners.push([name, listener])
        app[method](name, listener)
      }
      return proxy
    }
    const remove = (name, listener) => {
      if (SCOPED_APP_EVENTS.has(name)) scope.events.removeListener(name, listener)
      else if (WINDOW_APP_EVENTS.has(name)) {
        const wrapped = windowListeners.get(listener)?.get(name)
        if (wrapped) app.removeListener(name, wrapped)
      } else app.removeListener(name, listener)
      return proxy
    }
    const overrides = {
      on: add('on'),
      addListener: add('on'),
      once: add('once'),
      off: remove,
      removeListener: remove,
      // The host holds the one real lock; a hosted workspace is always the running instance.
      requestSingleInstanceLock: () => true,
      hasSingleInstanceLock: () => true,
      releaseSingleInstanceLock: () => {},
      quit: () => { quitScope(scope) },
      exit: () => { quitScope(scope, { force: true }) },
      relaunch: () => {},
      getPath: (name) => (name === 'userData' || name === 'sessionData' ? scope.userData : app.getPath(name)),
      setPath: (name, value) => { if (name !== 'userData' && name !== 'sessionData') app.setPath(name, value) },
      // Branding stays the host's.
      setName: () => {},
      getName: () => app.getName(),
      setAppUserModelId: () => {},
    }
    const bound = new Map()
    const proxy = new Proxy(app, {
      get(target, property) {
        if (Object.prototype.hasOwnProperty.call(overrides, property)) return overrides[property]
        const value = target[property]
        if (typeof value !== 'function') return value
        if (!bound.has(property)) bound.set(property, value.bind(target))
        return bound.get(property)
      },
      set(target, property, value) {
        if (property === 'name') return true
        target[property] = value
        return true
      },
    })
    return proxy
  }

  function scopedProtocol(scope) {
    const real = scope.session.protocol
    return {
      // Declared by the host before the app was ready (PRIVILEGED_SCHEMES).
      registerSchemesAsPrivileged: () => {},
      handle: (scheme, handler) => { real.handle(scheme, handler); scope.schemes.add(scheme) },
      unhandle: (scheme) => { real.unhandle(scheme); scope.schemes.delete(scheme) },
      isProtocolHandled: (scheme) => real.isProtocolHandled(scheme),
    }
  }

  function scopedElectron(scope) {
    const overrides = {
      app: scopedApp(scope),
      BrowserWindow: scopedBrowserWindow(scope),
      ipcMain: scopedIpcMain(scope),
      protocol: scopedProtocol(scope),
      session: new Proxy(session, {
        get(target, property) {
          if (property === 'defaultSession') return scope.session
          const value = target[property]
          return typeof value === 'function' ? value.bind(target) : value
        },
      }),
    }
    return new Proxy(electron, {
      get(target, property) {
        return Object.prototype.hasOwnProperty.call(overrides, property) ? overrides[property] : target[property]
      },
    })
  }

  // ---- Lifecycle ----

  function createScope(mode) {
    const userData = path.join(stateRoot, mode)
    const scope = {
      mode,
      userData,
      // The folder the workspace's own process used as its profile, so recent
      // files, settings and window state carry over.
      session: session.fromPath(userData),
      windows: new Set(),
      handlers: new Map(),
      messages: new EventEmitter(),
      events: new EventEmitter(),
      appListeners: [],
      schemes: new Set(),
      quitting: false,
      leaving: false,
      ended: false,
      started: false,
      failed: false,
      markStarted: () => {},
    }
    scope.messages.setMaxListeners(0)
    scope.events.setMaxListeners(0)
    // A workspace has started once its first window exists. One that has no window when
    // the time is up never will: its start failed after its code loaded (a rejected
    // promise on the way to its first window), and left alone it would keep this process
    // alive, unseen, for good. So it is ended and whoever asked for it is told.
    // The rule is the window and nothing else. Every workspace registers its app
    // listeners while its code loads, before anything can fail, so "it wired itself up"
    // says nothing about being alive; and none of them waits for a person before its
    // first window (only for its own files on disk), so 20 seconds without one is a failure.
    scope.whenStarted = new Promise((resolve, reject) => {
      let timer = null
      // Armed by startScope once the workspace's code has loaded: loading it (a cold disk, a
      // virus scan of the bundle) is not part of the time it has to open its window.
      scope.armStartTimer = () => { if (!scope.started && !scope.ended) timer = setTimeout(() => {
        if (scope.started || scope.ended || scope.windows.size || windowlessAllowed) { resolve(); return }
        log(`${mode} opened no window within ${START_TIMEOUT_MS} ms; ending it.\n`)
        scope.failed = true
        reject(new Error(`${mode} did not open a window in time (${START_TIMEOUT_MS} ms).`))
        endScope(scope)
      }, START_TIMEOUT_MS) }
      scope.markStarted = () => { if (!scope.started) { scope.started = true; if (timer) clearTimeout(timer); resolve() } }
    })
    // Every caller that waits handles the failure itself; with no caller it is not an unhandled rejection.
    scope.whenStarted.catch(() => {})
    scope.electron = scopedElectron(scope)
    return scope
  }

  function installRequireHook() {
    if (hookInstalled) return
    hookInstalled = true
    const load = Module._load
    Module._load = function simpleHostLoad(request, parent, ...rest) {
      if (request === 'electron' && parent && parent.filename) {
        const scope = scopes.get(modeOfFile(parent.filename))
        if (scope) return scope.electron
      }
      return load.call(this, request, parent, ...rest)
    }
  }

  function raise(scope, name, ...args) {
    let prevented = false
    const event = { preventDefault: () => { prevented = true }, get defaultPrevented() { return prevented } }
    for (const listener of scope.events.rawListeners(name)) {
      try { listener.call(scope.electron.app, event, ...args) } catch (error) { log(`${scope.mode}: a '${name}' listener failed: ${error && error.stack}\n`) }
    }
    return prevented
  }

  /** What app.quit() means for a workspace: ask (before-quit), close its windows, then end it. */
  function quitScope(scope, { force = false } = {}) {
    if (process.env.SIMPLE_HOST_DEBUG) log(`host: ${Date.now() % 100000} quit requested by ${scope.mode}: ${new Error().stack.split('\n').slice(2, 5).join(' <- ')}\n`)
    if (scope.ended) return
    // With no window left this quit cannot be called off by anyone: the workspace is leaving
    // (its guard may still be finishing a write). open() waits for that instead of handing
    // it a file it would close again.
    if (!scope.windows.size) scope.leaving = true
    if (!force && raise(scope, 'before-quit')) return
    scope.quitting = true
    const windows = [...scope.windows]
    if (!windows.length) { endScope(scope); return }
    for (const win of windows) {
      try { if (force) win.destroy(); else win.close() } catch {}
    }
    // A window that refused to close (unsaved changes, Cancel) cancels the quit.
    setImmediate(() => { if (!scope.ended && scope.windows.size) scope.quitting = false })
  }

  function windowClosed(scope) {
    if (scope.ended || scope.windows.size) return
    if (scope.quitting) { endScope(scope); return }
    if (scope.events.listenerCount('window-all-closed')) raise(scope, 'window-all-closed')
    else quitScope(scope)
  }

  function endScope(scope) {
    if (scope.ended) return
    scope.ended = true
    raise(scope, 'will-quit')
    raise(scope, 'quit', 0)
    for (const [name, listener] of scope.appListeners) app.removeListener(name, listener)
    for (const scheme of scope.schemes) { try { scope.session.protocol.unhandle(scheme) } catch {} }
    scope.handlers.clear()
    scope.messages.removeAllListeners()
    scope.events.removeAllListeners()
    scope.markStarted()
    if (scopes.get(scope.mode) === scope) scopes.delete(scope.mode)
    // Forget the workspace's code, so the next start runs it from the top.
    const folders = foldersFor(scope.mode)
    for (const file of Object.keys(require.cache)) {
      if (folders.some((folder) => inside(folder, file))) delete require.cache[file]
    }
    quitWhenEmpty()
  }

  function quitWhenEmpty() {
    if (scopes.size || pending) return
    appQuitting = true
    app.quit()
  }

  function startScope(mode, paths, flags = []) {
    const scope = createScope(mode)
    scopes.set(mode, scope)
    // The module reads its files from the command line once, while it starts.
    process.argv = [process.argv[0], ...flags, ...paths]
    try {
      if (process.env.SIMPLE_HOST_DEBUG) log(`host: starting ${mode} from ${entryFor(mode)} with ${paths.length} path(s)\n`)
      require(entryFor(mode))
      if (process.env.SIMPLE_HOST_DEBUG) log(`host: loaded ${mode}\n`)
    } catch (error) {
      log(`Could not start ${mode}: ${error && error.stack}\n`)
      endScope(scope)
      throw error
    }
    scope.armStartTimer()
    return scope.whenStarted
  }

  /**
   * Hands a request to a new instance that starts once this one has exited
   * (app.relaunch waits for the exit, so the new instance gets the lock). Each
   * request is its own relaunch: Electron starts one instance per call, and
   * they sort themselves out like any instances started together (the first
   * takes the lock, the others pass their files on to it).
   */
  function openAfterExit(mode, paths) {
    const execPath = process.env.PORTABLE_EXECUTABLE_FILE || process.execPath
    const flags = []
    // The same profile as this instance, when it was given one.
    const profile = app.commandLine.getSwitchValue('user-data-dir')
    if (profile) flags.push(`--user-data-dir=${profile}`)
    // Already routed: the new instance must not decide again (the launcher is what starts without a mode).
    if (mode !== 'launcher') flags.push(`--simple-mode=${mode}`)
    // An unpackaged Electron takes the app folder as its first argument.
    const withApp = (args) => (app.isPackaged ? args : [app.getAppPath(), ...args])
    let args = withApp([...flags, ...paths])
    // The portable wrapper cuts long command lines (see open-list.cjs).
    if (needsOpenList(execPath, args)) args = withApp([...flags, `${OPEN_LIST_SWITCH}${writeOpenListSync(stateRoot, paths)}`])
    app.relaunch({ execPath, args })
  }

  /**
   * Opens files in a workspace (starting it when needed); without files it
   * shows the workspace, as starting it a second time always did.
   * @param {string} mode
   * @param {string[]} paths absolute
   * @param {string} [workingDirectory]
   * @param {string[]} [flags] switches for the workspace (--simple-new: a new empty window, not the open one)
   * @returns {Promise<void>} settles once the request reached the workspace
   */
  function open(mode, paths = [], workingDirectory = process.cwd(), flags = []) {
    // A request that arrives after the last workspace ended, while the process is on
    // its way out: a workspace started now would end with the process and the file
    // would never open. A new instance opens it instead.
    if (appQuitting) {
      try { openAfterExit(mode, paths) } catch (error) { return Promise.reject(error) }
      return Promise.resolve()
    }
    pending += 1
    // One workspace starts at a time: each reads process.argv while it starts.
    const turn = loading.then(async () => {
      // Browser storage folders (createScope) exist only once the app is ready.
      await app.whenReady()
      const running = scopes.get(mode)
      if (running && !running.ended) {
        // A start that failed is reported to whoever asked for it; this request starts the workspace again.
        await running.whenStarted.catch(() => {})
        // Leaving (see quitScope): let it end, at most as long as its guard waits for writes.
        for (let waited = 0; running.leaving && !running.ended && !running.windows.size && waited < 35_000; waited += 100) await new Promise((resolve) => setTimeout(resolve, 100))
        if (!running.ended) { raise(running, 'second-instance', [process.argv[0], ...flags, ...paths], workingDirectory, {}); return }
      }
      await startScope(mode, paths, flags)
    })
    loading = turn.catch(() => {})
    const settle = () => { pending -= 1; quitWhenEmpty() }
    // A failure is reported (onOpenFailure may show a message) before the process is allowed
    // to end: when the failed workspace was the only one, quitting first would take the
    // message down with it and the file would just never appear.
    return turn.then((value) => { settle(); return value }, async (error) => {
      try { await onOpenFailure?.(mode, paths, error) } catch { /* reporting is best effort */ }
      settle()
      throw error
    })
  }

  /** Call before the app is ready. */
  function install() {
    electron.protocol.registerSchemesAsPrivileged(PRIVILEGED_SCHEMES)
    installRequireHook()
    // The host decides when the process ends (see quitWhenEmpty).
    app.on('window-all-closed', () => {})
    app.on('before-quit', (event) => {
      if (appQuitting || !scopes.size) return
      // Quit requested from outside a workspace: each one is asked in turn, as when it quits itself.
      event.preventDefault()
      for (const scope of [...scopes.values()]) quitScope(scope)
    })
  }

  return { install, open, modes: () => [...scopes.keys()] }
}

module.exports = { createHost, PRIVILEGED_SCHEMES }
