'use strict'

// Starting Simple again from the launcher and the bootstrap: long file lists
// travel in a list file (the portable wrapper cuts command lines near 8,192
// characters), and a process that can't start is reported, never thrown.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { EventEmitter } = require('node:events')
const openList = require('../electron/open-list.cjs')
const routing = require('../electron/routing.cjs')

function scratch(t) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-launch-'))
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }))
  return directory
}

/** Loads electron/launch.cjs for a portable EXE path and profile, restoring the environment afterwards. */
function loadLaunch(t, { exe, profile }) {
  const previous = { exe: process.env.PORTABLE_EXECUTABLE_FILE, profile: process.env.SIMPLE_USER_DATA_DIR }
  process.env.PORTABLE_EXECUTABLE_FILE = exe
  process.env.SIMPLE_USER_DATA_DIR = profile
  t.after(() => {
    for (const [key, value] of [['PORTABLE_EXECUTABLE_FILE', previous.exe], ['SIMPLE_USER_DATA_DIR', previous.profile]]) {
      if (value === undefined) delete process.env[key]
      else process.env[key] = value
    }
  })
  return require('../electron/launch.cjs')
}

/** A spawn stand-in: records the call, then starts or fails on the next tick like child_process does. */
function fakeSpawn(outcome) {
  const calls = []
  const spawn = (command, args, options) => {
    const child = new EventEmitter()
    child.pid = 4242
    child.unref = () => { child.unrefed = true }
    calls.push({ command, args, options, child })
    process.nextTick(() => {
      if (outcome === 'spawn') child.emit('spawn')
      else child.emit('error', Object.assign(new Error(`spawn ${command} ${outcome}`), { code: outcome }))
    })
    return child
  }
  return { spawn, calls }
}

function longPaths(folder, count) {
  return Array.from({ length: count }, (_, index) => path.join(folder, 'Camera Roll', '2026', `A rather long holiday photo name number ${String(index).padStart(4, '0')}.jpg`))
}

test('a long file list is written to a list file in the profile, read once and deleted', (t) => {
  const profile = scratch(t)
  const paths = longPaths(profile, 3)
  return openList.writeOpenList(profile, paths).then((file) => {
    assert.equal(path.dirname(file), openList.openListFolder(profile))
    assert.match(path.basename(file), /^open-[0-9a-f]{32}\.json$/)
    const argv = ['simple.exe', '--simple-mode=image', `${openList.OPEN_LIST_SWITCH}${file}`, '--flag']
    assert.deepEqual(openList.expandOpenLists(argv, profile), ['simple.exe', '--simple-mode=image', ...paths, '--flag'])
    assert.equal(fs.existsSync(file), false, 'the list is deleted once read')
    assert.deepEqual(openList.expandOpenLists(argv, profile), ['simple.exe', '--simple-mode=image', '--flag'], 'a list is never read twice')
  })
})

test('only list files of this profile are read or deleted; anything else on the command line is dropped unread', (t) => {
  const profile = scratch(t)
  const outside = path.join(profile, 'important.json')
  fs.writeFileSync(outside, JSON.stringify({ version: 1, paths: [path.join(profile, 'x.pdf')] }))
  const folder = openList.openListFolder(profile)
  fs.mkdirSync(folder, { recursive: true })
  const misnamed = path.join(folder, 'notes.json')
  fs.writeFileSync(misnamed, JSON.stringify({ version: 1, paths: [path.join(profile, 'y.pdf')] }))
  const relative = path.join(folder, `open-${'a'.repeat(32)}.json`)
  fs.writeFileSync(relative, JSON.stringify({ version: 1, paths: ['relative.pdf', 7, path.join(profile, 'z.pdf')] }))
  const argv = ['simple.exe', `${openList.OPEN_LIST_SWITCH}${outside}`, `${openList.OPEN_LIST_SWITCH}${misnamed}`, `${openList.OPEN_LIST_SWITCH}${relative}`]
  assert.deepEqual(openList.expandOpenLists(argv, profile), ['simple.exe', path.join(profile, 'z.pdf')])
  assert.equal(fs.existsSync(outside), true, 'a file outside the list folder is never deleted')
  assert.equal(fs.existsSync(misnamed), true)
  assert.equal(fs.existsSync(relative), false)
  // A list nobody picked up (its process failed to start) is swept after a day.
  const stale = path.join(folder, `open-${'b'.repeat(32)}.json`)
  fs.writeFileSync(stale, '{}')
  const old = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000)
  fs.utimesSync(stale, old, old)
  openList.sweepOpenLists(folder)
  assert.equal(fs.existsSync(stale), false)
  assert.equal(fs.existsSync(misnamed), true)
})

test('launchDetached passes short lists on the command line and long ones through a list file', async (t) => {
  const profile = scratch(t)
  const { launchDetached } = loadLaunch(t, { exe: path.join(profile, 'simple.exe'), profile })
  const fake = fakeSpawn('spawn')
  const few = longPaths(profile, 3)
  assert.equal(await launchDetached(['--simple-mode=image', ...few], { spawn: fake.spawn }), 4242)
  assert.deepEqual(fake.calls[0].args, ['--simple-mode=image', ...few])
  assert.equal(fake.calls[0].options.detached, true)
  assert.equal(fake.calls[0].child.unrefed, true)

  const many = longPaths(profile, 120)
  assert.ok(openList.commandLineLength(path.join(profile, 'simple.exe'), many) > 8192, 'the fixture is longer than the portable wrapper allows')
  await launchDetached(['--simple-mode=image', ...many], { spawn: fake.spawn })
  const { args } = fake.calls[1]
  assert.equal(args.length, 2)
  assert.equal(args[0], '--simple-mode=image')
  assert.ok(args[1].startsWith(openList.OPEN_LIST_SWITCH))
  assert.ok(openList.commandLineLength(path.join(profile, 'simple.exe'), args) < openList.MAX_COMMAND_LINE_CHARS)
  // What the new process sees, and how it routes them: every file, in order.
  const argv = openList.expandOpenLists(['simple.exe', ...args], profile)
  assert.deepEqual(argv.slice(2), many)
  const routed = routing.routeCommandLine(argv)
  assert.equal(routed.mode, 'image')
  assert.equal(routed.groups.get('image').length, 120)
})

test('a process that cannot start is reported with a plain message, and its list file is removed', async (t) => {
  const profile = scratch(t)
  const { launchDetached } = loadLaunch(t, { exe: path.join(profile, 'simple.exe'), profile })
  const fake = fakeSpawn('ENOENT')
  await assert.rejects(launchDetached(['--simple-mode=image', ...longPaths(profile, 120)], { spawn: fake.spawn }), (error) => {
    assert.equal(error.code, 'LAUNCH_FAILED')
    assert.equal(error.cause.code, 'ENOENT')
    assert.match(error.message, /simple\.exe was moved, renamed or deleted\. Start Simple again from its new place\./)
    return true
  })
  await new Promise((resolve) => setTimeout(resolve, 50))
  const folder = openList.openListFolder(profile)
  assert.deepEqual(fs.existsSync(folder) ? fs.readdirSync(folder) : [], [], 'no list file is left behind')
  // A later error event (after a failed start) is swallowed, never uncaught.
  fake.calls[0].child.emit('error', new Error('late'))
})

test('a real spawn of a missing EXE rejects instead of throwing in the main process', async (t) => {
  const profile = scratch(t)
  const { launchDetached } = loadLaunch(t, { exe: path.join(profile, 'moved', 'simple.exe'), profile })
  await assert.rejects(launchDetached([path.join(profile, 'a.pdf')]), (error) => error.code === 'LAUNCH_FAILED' && /was moved, renamed or deleted/.test(error.message))
})

test('a running workspace receives every routed path, also those that came in a list file', () => {
  const folder = path.join(os.tmpdir(), 'simple-launch-routing')
  const paths = longPaths(folder, 4)
  // The raw command line of the new instance holds only the (already deleted) list.
  const argv = ['simple.exe', '--simple-mode=image', `${openList.OPEN_LIST_SWITCH}${path.join(folder, 'open-lists', `open-${'c'.repeat(32)}.json`)}`]
  routing.filterSecondInstanceArgv(argv, { mode: 'image', workingDirectory: folder, additionalData: { simpleRouting: { mode: 'image', paths } } })
  assert.deepEqual(argv, ['simple.exe', '--simple-mode=image', ...paths])
  // An ordinary relay is unchanged: nothing is added twice.
  const relayed = ['simple.exe', 'a.jpg', 'b.pdf']
  routing.filterSecondInstanceArgv(relayed, { mode: 'image', workingDirectory: folder, additionalData: { simpleRouting: { mode: 'image', paths: [path.join(folder, 'a.jpg')] } } })
  assert.deepEqual(relayed, ['simple.exe', path.join(folder, 'a.jpg')])
  // Paths routed to another workspace are never added.
  const other = ['simple.exe']
  routing.filterSecondInstanceArgv(other, { mode: 'pdf', additionalData: { simpleRouting: { mode: 'image', paths } } })
  assert.deepEqual(other, ['simple.exe'])
})

test('the launcher Open and Combine pickers list what the format registry says', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'launcher', 'main.cjs'), 'utf8')
  assert.match(source, /formats\.dialogFilters\('launcher'/)
  assert.match(source, /formats\.dialogFilters\('combine', \{ engine/)
  assert.doesNotMatch(source, /EXTENSIONS_BY_MODE\.\w+\.map|\.\.\.COMBINE_EXTENSIONS/, 'no hand-written Open filter lists')
  const formats = require('../shared/electron/formats.cjs')
  const { COMBINE_EXTENSIONS } = require('../launcher/combine-policy.cjs')
  // Combine accepts exactly what the registry offers with an engine; without one, .doc and .xls are not offered.
  assert.deepEqual(new Set(formats.openExtensions('combine', { engine: true }).map((extension) => extension.slice(1))), new Set(COMBINE_EXTENSIONS))
  assert.deepEqual(new Set(formats.openExtensions('combine', { engine: false }).map((extension) => extension.slice(1))), new Set(COMBINE_EXTENSIONS.filter((extension) => extension !== 'doc' && extension !== 'xls')))
})

test('launchDetached creates the process before it returns, also when the paths travel in a list file', (t) => {
  const profile = scratch(t)
  const { launchDetached } = loadLaunch(t, { exe: path.join(profile, 'simple.exe'), profile })
  const fake = fakeSpawn('spawn')
  // Not awaited: the bootstrap loads its workspace module right after, and a
  // workspace that is already running quits this process while it loads.
  const started = launchDetached(['--simple-mode=pdf', ...longPaths(profile, 120)], { spawn: fake.spawn })
  assert.equal(fake.calls.length, 1, 'spawned in the same tick')
  const listArgument = fake.calls[0].args.find((argument) => argument.startsWith(openList.OPEN_LIST_SWITCH))
  assert.ok(listArgument && fs.existsSync(listArgument.slice(openList.OPEN_LIST_SWITCH.length)), 'the list file is on disk before the process starts')
  return started
})

test('every start removes list files that no process picked up, also without a list of its own', (t) => {
  const profile = scratch(t)
  const folder = openList.openListFolder(profile)
  fs.mkdirSync(folder, { recursive: true })
  const stale = path.join(folder, `open-${'d'.repeat(32)}.json`)
  const fresh = path.join(folder, `open-${'e'.repeat(32)}.json`)
  for (const file of [stale, fresh]) fs.writeFileSync(file, JSON.stringify({ version: 1, paths: [path.join(profile, 'private.pdf')] }))
  const old = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000)
  fs.utimesSync(stale, old, old)
  assert.deepEqual(openList.expandOpenLists(['simple.exe', 'a.pdf'], profile), ['simple.exe', 'a.pdf'])
  assert.equal(fs.existsSync(stale), false, 'the paths of a failed hand-off do not stay in the profile')
  assert.equal(fs.existsSync(fresh), true, 'a list another process is about to read is kept')
})

/** A stand-in for Electron's app: ready on demand, records quits and lets will-quit be prevented. */
function fakeApp() {
  const app = new EventEmitter()
  let markReady
  const ready = new Promise((resolve) => { markReady = resolve })
  app.whenReady = () => ready
  app.ready = () => markReady()
  app.quits = 0
  app.exited = false
  app.quit = () => {
    app.quits += 1
    const event = { defaultPrevented: false, preventDefault() { this.defaultPrevented = true } }
    app.emit('will-quit', event)
    if (!event.defaultPrevented) app.exited = true
  }
  return app
}

const tick = () => new Promise((resolve) => setImmediate(resolve))

test('a workspace that cannot start is shown in a message once the app is ready, naming its files', async () => {
  const { createHandOff, MAX_NAMED_FILES } = require('../electron/hand-off.cjs')
  const app = fakeApp()
  const shown = []
  const logged = []
  const paths = Array.from({ length: MAX_NAMED_FILES + 2 }, (_, index) => path.join(os.tmpdir(), `Statement ${index + 1}.pdf`))
  const handOff = createHandOff({
    app,
    dialog: { showMessageBox: async (options) => { shown.push(options); return { response: 0 } } },
    launch: async () => { throw Object.assign(new Error("Simple couldn't start another window because simple.exe was moved, renamed or deleted. Start Simple again from its new place."), { code: 'LAUNCH_FAILED' }) },
    label: (mode) => ({ pdf: 'Simple PDF' })[mode],
    log: (text) => logged.push(text),
  })
  assert.equal(await handOff.start('pdf', paths), false)
  await tick()
  assert.equal(shown.length, 0, 'nothing is shown before the app is ready')
  assert.match(logged.join(''), /Could not open 12 file\(s\) in pdf: Simple couldn't start another window/)
  app.ready()
  await handOff.settled()
  assert.equal(shown.length, 1)
  assert.equal(shown[0].type, 'warning')
  assert.equal(shown[0].message, "Simple couldn't open 12 files in Simple PDF.")
  assert.match(shown[0].detail, /simple\.exe was moved, renamed or deleted/)
  assert.match(shown[0].detail, /Statement 1\.pdf\nStatement 2\.pdf/)
  assert.match(shown[0].detail, /Statement 10\.pdf\nand 2 more$/)
  assert.doesNotMatch(shown[0].detail, /Statement 11\.pdf/)
})

test('a quit waits for every start and for the message about files that did not open', async () => {
  const { createHandOff } = require('../electron/hand-off.cjs')
  const app = fakeApp()
  const shown = []
  let finishLaunch
  let closeMessage
  const handOff = createHandOff({
    app,
    dialog: { showMessageBox: (options) => new Promise((resolve) => { shown.push(options); closeMessage = resolve }) },
    launch: () => new Promise((resolve, reject) => { finishLaunch = reject }),
    log: () => {},
  })
  void handOff.start('calc', [path.join(os.tmpdir(), 'Budget.xlsx')])
  // The workspace module quits at once (its workspace is already running).
  app.quit()
  assert.equal(app.exited, false, 'the quit is held while a start is pending')
  finishLaunch(Object.assign(new Error('Simple couldn\'t start another window (EPERM). Try again.'), { code: 'LAUNCH_FAILED' }))
  app.ready()
  for (let index = 0; index < 5 && !shown.length; index += 1) await tick()
  assert.equal(shown.length, 1, 'the message is shown before the process ends')
  assert.equal(app.exited, false)
  assert.equal(shown[0].message, "Simple couldn't open 1 file in calc.")
  closeMessage({ response: 0 })
  for (let index = 0; index < 5 && !app.exited; index += 1) await tick()
  assert.equal(app.exited, true, 'then the quit goes ahead')

  // A start that succeeds holds nothing and shows nothing.
  const calm = fakeApp()
  calm.ready()
  const quiet = createHandOff({ app: calm, dialog: { showMessageBox: async () => assert.fail('no message') }, launch: async () => 7, log: () => {} })
  assert.equal(await quiet.start('pdf', [path.join(os.tmpdir(), 'a.pdf')]), true)
  calm.quit()
  assert.equal(calm.exited, true)
})
