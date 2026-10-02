'use strict'

// `simple --simple-self-test` (electron/self-test.cjs): every workspace bundle,
// shared I/O module and converter loads from the app's layout, and a missing
// one is named.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { execFile, spawn } = require('node:child_process')
const { COMBINE_CONVERTERS, checkEntryFile, runSelfTest, sharedFiles } = require('../electron/self-test.cjs')

const ROOT = path.resolve(__dirname, '..')

test('the self-test covers every shared module the manifest vendors', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'shared', 'manifest.json'), 'utf8'))
  const { modules, data } = sharedFiles(ROOT)
  for (const workspace of Object.values(manifest.workspaces)) {
    for (const source of workspace.electron) {
      const relative = `shared/${source}`
      assert.ok(source.endsWith('.json') ? data.includes(relative) : modules.includes(relative), `${relative} is not checked`)
    }
  }
  for (const required of ['shared/electron/html-to-pdf.cjs', 'shared/electron/office-engine.cjs', 'shared/electron/safe-write.cjs']) assert.ok(modules.includes(required), required)
})

test('the unpackaged layout passes the self-test in Node', async () => {
  const report = await runSelfTest({ root: ROOT })
  assert.equal(report.healthy, true, report.failed.join('\n'))
  assert.ok(report.checks.some((check) => check.name === 'Combine worker bundle compiles' && check.ok))
  assert.ok(report.checks.some((check) => check.name === 'safeWriteFile writes and verifies' && check.ok))
})

test('a layout missing a shared module, a bundle or the Combine worker fails and names it', async (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-self-test-layout-'))
  t.after(() => fs.rmSync(root, { recursive: true, force: true }))
  for (const folder of ['shared', 'launcher', 'electron']) fs.cpSync(path.join(ROOT, folder), path.join(root, folder), { recursive: true })
  fs.rmSync(path.join(root, 'shared', 'electron', 'html-to-pdf.cjs'))
  const report = await runSelfTest({ root })
  assert.equal(report.healthy, false)
  const failed = report.failed.join('\n')
  assert.match(failed, /load shared\/electron\/html-to-pdf\.cjs/)
  assert.match(failed, /module pdf: missing dist\/index\.html, electron\/main\.cjs/)
  assert.match(failed, /Combine worker bundle compiles: missing/)
})

test('a Combine worker bundled without one of its converters fails and names it', async (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-self-test-bundle-'))
  t.after(() => fs.rmSync(root, { recursive: true, force: true }))
  for (const folder of ['shared', 'launcher', 'electron']) fs.cpSync(path.join(ROOT, folder), path.join(root, folder), { recursive: true })
  fs.mkdirSync(path.join(root, 'modules', 'shared'), { recursive: true })
  const markers = Object.values(COMBINE_CONVERTERS).filter((marker) => marker !== 'convertToHtml')
  fs.writeFileSync(path.join(root, 'modules', 'shared', 'combine-worker.cjs'), `'use strict'\n// ${markers.join(' ')}\n`)
  const report = await runSelfTest({ root })
  const failed = report.failed.join('\n')
  assert.match(failed, /Combine worker bundle compiles: the bundle has no Word layout \(mammoth\)/)
})

test('Electron: simple --simple-self-test reports a healthy app, including the PDF printer', { skip: process.platform !== 'win32' && 'Windows only', timeout: 150_000 }, async (t) => {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-self-test-profile-'))
  t.after(() => fs.rmSync(profile, { recursive: true, force: true }))
  const env = { ...process.env, SIMPLE_FORCE_NO_OFFICE: '1' }
  delete env.ELECTRON_RUN_AS_NODE
  const child = spawn(require('electron'), [ROOT, '--simple-self-test', `--user-data-dir=${profile}`], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
  let output = ''
  child.stdout.on('data', (data) => { output += data })
  child.stderr.on('data', (data) => { output += data })
  const stopTree = () => new Promise((resolve) => execFile('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => resolve()))
  const timer = setTimeout(() => { void stopTree() }, 120_000)
  const code = await new Promise((resolve) => child.once('exit', resolve))
  clearTimeout(timer)
  await stopTree()
  const line = output.split(/\r?\n/).find((text) => text.startsWith('{"healthy"'))
  assert.ok(line, `no report. ${output.slice(-2000)}`)
  const report = JSON.parse(line)
  assert.deepEqual(report.failed, [])
  assert.equal(report.healthy, true)
  assert.equal(code, 0)
  assert.ok(report.checks > 20)
})

test('the launcher entry point is compiled and every module it requires must be there; MuPDF must really unlock', async (t) => {
  const report = await runSelfTest({ root: ROOT })
  for (const name of ['compile launcher/main.cjs and resolve its requires', 'compile electron/bootstrap.cjs and resolve its requires', 'Combine can unlock permissions-only PDFs (MuPDF)']) {
    assert.ok(report.checks.some((check) => check.name === name && check.ok), `${name}: ${JSON.stringify(report.checks.find((check) => check.name === name))}`)
  }
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-self-test-entry-'))
  t.after(() => fs.rmSync(root, { recursive: true, force: true }))
  for (const folder of ['shared', 'launcher', 'electron']) fs.cpSync(path.join(ROOT, folder), path.join(root, folder), { recursive: true })
  // A module the launcher requires was moved: the launcher window would stay blank.
  fs.renameSync(path.join(root, 'launcher', 'open-paths.cjs'), path.join(root, 'launcher', 'open-paths-moved.cjs'))
  fs.appendFileSync(path.join(root, 'electron', 'bootstrap.cjs'), '\nconst broken = (\n')
  const broken = await runSelfTest({ root })
  const failed = broken.failed.join('\n')
  assert.match(failed, /compile launcher\/main\.cjs and resolve its requires: requires \.\/open-paths\.cjs, which is missing/)
  assert.match(failed, /compile electron\/bootstrap\.cjs and resolve its requires: /)
  assert.throws(() => checkEntryFile(path.join(root, 'electron', 'bootstrap.cjs')), SyntaxError)
})
