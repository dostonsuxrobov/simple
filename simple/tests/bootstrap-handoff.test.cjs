'use strict'

// The bootstrap in Electron: how a command line reaches the workspaces when
// one of them is already running. A scratch copy of electron/*.cjs and the
// format registry runs with fake workspaces that open no window, log what
// they receive and quit by themselves; each scenario has its own profile, and
// every Electron process still running at the end is ended.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { execFile, execFileSync, spawn } = require('node:child_process')

const ROOT = path.resolve(__dirname, '..')
const MODES = ['docs', 'calc', 'pdf', 'image', 'video']
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

/** A fake workspace: takes the lock like the real ones, logs argv and every relay, and never opens a window. */
function fakeModule(mode) {
  return `'use strict'
const electron = require('electron')
const fs = require('node:fs')
const { app } = electron
const log = (entry) => fs.appendFileSync(process.env.SIMPLE_E2E_LOG, JSON.stringify({ mode: ${JSON.stringify(mode)}, pid: process.pid, ...entry }) + '\\n')
// A real message box would wait for a person; record it instead.
electron.dialog.showMessageBox = async (options) => { log({ event: 'message', message: options.message, detail: options.detail }); return { response: 0 } }
const gotLock = app.requestSingleInstanceLock()
log({ event: 'start', gotLock, argv: process.argv.slice(1) })
if (!gotLock) app.quit()
else {
  app.on('second-instance', (_event, argv, workingDirectory, additionalData) => log({ event: 'second-instance', argv: argv.slice(1), additionalData }))
  setTimeout(() => app.exit(0), Number(process.env.SIMPLE_E2E_HOLD_MS || 4000))
}
`
}

function makeApp(root) {
  const app = path.join(root, 'app')
  fs.mkdirSync(path.join(app, 'electron'), { recursive: true })
  for (const name of fs.readdirSync(path.join(ROOT, 'electron'))) {
    if (name.endsWith('.cjs')) fs.copyFileSync(path.join(ROOT, 'electron', name), path.join(app, 'electron', name))
  }
  fs.cpSync(path.join(ROOT, 'shared', 'electron'), path.join(app, 'shared', 'electron'), { recursive: true })
  for (const mode of MODES) {
    fs.mkdirSync(path.join(app, 'modules', mode, 'electron'), { recursive: true })
    fs.writeFileSync(path.join(app, 'modules', mode, 'electron', 'main.cjs'), fakeModule(mode))
  }
  fs.writeFileSync(path.join(app, 'package.json'), JSON.stringify({ name: 'simple-bootstrap-e2e', version: '0.0.0', main: 'electron/bootstrap.cjs' }))
  return app
}

/**
 * Ends every Electron process whose command line mentions the scratch folder
 * (also the ones the bootstrap started detached), and waits until they are
 * gone so the folder can be removed.
 * @returns {number} how many were still running at the last look
 */
function endLeftovers(marker) {
  const script = [
    `$deadline = (Get-Date).AddSeconds(15)`,
    `do {`,
    `  $left = @(Get-CimInstance Win32_Process -Filter "Name='electron.exe'" | Where-Object { $_.CommandLine -like '*${marker}*' })`,
    `  $left | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }`,
    `  if ($left.Count) { Start-Sleep -Milliseconds 250 }`,
    `} while ($left.Count -and (Get-Date) -lt $deadline)`,
    `$left.Count`,
  ].join('\n')
  try {
    const output = execFileSync(path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'), ['-NoProfile', '-NonInteractive', '-Command', script], {
      encoding: 'utf8', windowsHide: true, timeout: 40_000,
    })
    return Number.parseInt(output.trim(), 10) || 0
  } catch {
    return -1
  }
}

function harness(t) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-bootstrap-e2e-'))
  const marker = path.basename(root)
  const app = makeApp(root)
  const files = path.join(root, 'files')
  fs.mkdirSync(files)
  const children = []
  t.after(async () => {
    for (const child of children) {
      await new Promise((resolve) => execFile('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => resolve()))
    }
    assert.equal(endLeftovers(marker), 0, 'every Electron process this test started has ended')
    fs.rmSync(root, { recursive: true, force: true, maxRetries: 25, retryDelay: 200 })
  })
  const scenario = (name) => {
    const profile = path.join(root, `profile-${name}`)
    const logFile = path.join(root, `log-${name}.jsonl`)
    fs.mkdirSync(profile)
    const env = { ...process.env, SIMPLE_E2E_LOG: logFile }
    for (const key of ['ELECTRON_RUN_AS_NODE', 'SIMPLE_USER_DATA_DIR', 'PORTABLE_EXECUTABLE_FILE']) delete env[key]
    const start = (args, options = {}) => {
      const child = spawn(require('electron'), [app, `--user-data-dir=${profile}`, ...args], {
        cwd: options.cwd || root, env: { ...env, ...options.env }, windowsHide: true, stdio: 'ignore',
      })
      children.push(child)
      child.exited = new Promise((resolve) => child.once('exit', resolve))
      return child
    }
    const read = () => {
      try { return fs.readFileSync(logFile, 'utf8').trim().split('\n').filter(Boolean).map((line) => JSON.parse(line)) } catch { return [] }
    }
    const waitFor = async (predicate, what, timeoutMs = 20_000) => {
      const end = Date.now() + timeoutMs
      while (Date.now() < end) {
        const entries = read()
        if (predicate(entries)) return entries
        await sleep(100)
      }
      assert.fail(`timed out waiting for ${what}: ${JSON.stringify(read()).slice(0, 2000)}`)
    }
    const lists = () => { try { return fs.readdirSync(path.join(profile, 'open-lists')) } catch { return [] } }
    return { profile, start, read, waitFor, lists }
  }
  return { files, scenario }
}

const running = (entries, mode) => entries.some((entry) => entry.mode === mode && entry.event === 'start' && entry.gotLock)

function longPdfPaths(folder, count) {
  return Array.from({ length: count }, (_, index) => path.join(folder, 'A folder with a fairly long name for scanned statements', `Statement number ${String(index).padStart(4, '0')} for the account review.pdf`))
}

test('Electron: files for a running workspace and for others all open, also when they travel in list files', { skip: process.platform !== 'win32' && 'Windows only', timeout: 240_000 }, async (t) => {
  const { files, scenario } = harness(t)

  await t.test('a relay: the running workspace gets its own files, made absolute; another workspace starts for the rest', async () => {
    const run = scenario('relay')
    run.start(['--simple-mode=calc'], { env: { SIMPLE_E2E_HOLD_MS: '15000' } })
    await run.waitFor((entries) => running(entries, 'calc'), 'Spreadsheets to start')
    run.start(['Budget.xlsx', 'Other.xlsx', 'Report.pdf'], { cwd: files, env: { SIMPLE_E2E_HOLD_MS: '3000' } })
    const entries = await run.waitFor((log) => log.some((entry) => entry.event === 'second-instance') && running(log, 'pdf'), 'the relay and the PDF workspace')
    const relay = entries.find((entry) => entry.event === 'second-instance')
    assert.equal(relay.mode, 'calc')
    assert.deepEqual(relay.argv.filter((argument) => !argument.startsWith('-')).slice(1), [path.join(files, 'Budget.xlsx'), path.join(files, 'Other.xlsx')])
    assert.deepEqual(relay.additionalData.simpleRouting, { mode: 'calc', paths: [path.join(files, 'Budget.xlsx'), path.join(files, 'Other.xlsx')] })
    const pdf = entries.find((entry) => entry.mode === 'pdf' && entry.event === 'start' && entry.gotLock)
    assert.ok(pdf.argv.includes('--simple-mode=pdf'))
    assert.ok(pdf.argv.includes(path.join(files, 'Report.pdf')))
    assert.ok(!pdf.argv.some((argument) => argument.endsWith('.xlsx')), 'a workspace never sees another workspace\'s files')
  })

  await t.test('a list file relayed to the running workspace arrives as every path', async () => {
    const run = scenario('list-relay')
    run.start(['--simple-mode=calc'], { env: { SIMPLE_E2E_HOLD_MS: '15000' } })
    await run.waitFor((entries) => running(entries, 'calc'), 'Spreadsheets to start')
    const { writeOpenList } = require('../electron/open-list.cjs')
    const workbooks = Array.from({ length: 100 }, (_, index) => path.join(files, 'Quarterly workbooks', `Workbook ${String(index).padStart(4, '0')}.xlsx`))
    const list = await writeOpenList(run.profile, workbooks)
    run.start(['--simple-mode=calc', `--simple-open-list=${list}`], { env: { SIMPLE_E2E_HOLD_MS: '3000' } })
    const entries = await run.waitFor((log) => log.some((entry) => entry.event === 'second-instance'), 'the relay')
    const relay = entries.find((entry) => entry.event === 'second-instance')
    assert.deepEqual(relay.argv.filter((argument) => argument.endsWith('.xlsx')), workbooks)
    assert.ok(!relay.argv.some((argument) => argument.startsWith('--simple-open-list=')))
    assert.deepEqual(run.lists(), [], 'the list file is read once and deleted')
  })

  await t.test('a long list for another workspace still opens when this process quits at once (its own workspace is running)', async () => {
    const run = scenario('long-other')
    run.start(['--simple-mode=calc'], { env: { SIMPLE_E2E_HOLD_MS: '15000' } })
    await run.waitFor((entries) => running(entries, 'calc'), 'Spreadsheets to start')
    const statements = longPdfPaths(files, 60)
    const second = run.start([path.join(files, 'Budget.xlsx'), ...statements], { env: { SIMPLE_E2E_HOLD_MS: '3000' } })
    const entries = await run.waitFor((log) => running(log, 'pdf') && log.some((entry) => entry.event === 'second-instance'), 'the PDF workspace and the relay')
    const pdf = entries.find((entry) => entry.mode === 'pdf' && entry.event === 'start' && entry.gotLock)
    assert.deepEqual(pdf.argv.filter((argument) => argument.endsWith('.pdf')), statements, 'all 60 PDFs reach the PDF workspace')
    assert.ok(!pdf.argv.some((argument) => argument.startsWith('--simple-open-list=')))
    const relay = entries.find((entry) => entry.event === 'second-instance')
    assert.deepEqual(relay.additionalData.simpleRouting.paths, [path.join(files, 'Budget.xlsx')])
    await second.exited
    assert.deepEqual(run.lists(), [], 'no list file is left in the profile')
  })

  await t.test('a workspace that cannot start is reported in a message before the process ends', async () => {
    const run = scenario('failed-start')
    run.start(['--simple-mode=calc'], { env: { SIMPLE_E2E_HOLD_MS: '15000' } })
    await run.waitFor((entries) => running(entries, 'calc'), 'Spreadsheets to start')
    const second = run.start([path.join(files, 'Budget.xlsx'), path.join(files, 'Report.pdf')], {
      env: { SIMPLE_E2E_HOLD_MS: '3000', PORTABLE_EXECUTABLE_FILE: path.join(files, 'moved', 'simple.exe') },
    })
    const entries = await run.waitFor((log) => log.some((entry) => entry.event === 'message'), 'the message')
    const message = entries.find((entry) => entry.event === 'message')
    assert.equal(message.message, "Simple couldn't open 1 file in Simple PDF.")
    assert.match(message.detail, /simple\.exe was moved, renamed or deleted/)
    assert.match(message.detail, /Report\.pdf/)
    assert.ok(entries.some((entry) => entry.event === 'second-instance'), 'its own file still reached the running workspace')
    assert.equal(await Promise.race([second.exited, sleep(15_000).then(() => 'still running')]) !== 'still running', true, 'then the process ends')
  })
})
