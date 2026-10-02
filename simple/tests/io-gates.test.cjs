'use strict'

// The I/O QA gates (design §10.2, §10.3): the fast, Electron-free parts run
// here on every `npm test`; `npm run test:io` runs them in full, including the
// real app for every workspace enabled in shared/manifest.json.
// - scripts/no-office-matrix.cjs: nothing Simple offers needs LibreOffice, and
//   legacy files are saved to a modern sibling with the original untouched;
// - scripts/io-acceptance.cjs: "Open the file after exporting" opens the file
//   in Simple (openInSimple), never in another program.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { execFile } = require('node:child_process')

const ROOT = path.resolve(__dirname, '..')

function runScript(script, args, timeoutMs = 180_000) {
  return new Promise((resolve) => {
    const env = { ...process.env }
    delete env.ELECTRON_RUN_AS_NODE
    execFile(process.execPath, [path.join(ROOT, 'scripts', script), ...args], { cwd: ROOT, env, timeout: timeoutMs, windowsHide: true, maxBuffer: 16 * 1024 * 1024 }, (error, stdout, stderr) => {
      resolve({ code: error ? (typeof error.code === 'number' ? error.code : 1) : 0, stdout: String(stdout), stderr: String(stderr) })
    })
  })
}

test('engine-absent matrix: registry, engine and sibling saves pass for every workspace', { timeout: 200_000 }, async () => {
  const { code, stdout, stderr } = await runScript('no-office-matrix.cjs', ['--shared-only', '--json'])
  const line = stdout.split(/\r?\n/).find((text) => text.startsWith('{'))
  assert.ok(line, `no result. ${stderr.slice(-2000)}`)
  const result = JSON.parse(line)
  const failed = result.outcomes.filter((outcome) => !outcome.ok)
  assert.deepEqual(failed, [], failed.map((outcome) => `${outcome.name}: ${outcome.detail}`).join('\n'))
  assert.equal(code, 0)
  const names = result.outcomes.map((outcome) => outcome.name)
  for (const workspace of ['pdf', 'calc', 'docs', 'image', 'video']) assert.ok(names.some((name) => name.startsWith(`registry ${workspace}:`)), workspace)
  assert.ok(names.some((name) => name.startsWith('engine: probe says forced-off')))
  for (const row of ['save calc xls → sibling .xlsx', 'save calc ods → sibling .xlsx', 'save docs doc → sibling .docx']) assert.ok(names.includes(row), row)
})

test('the registry check catches a legacy binary saved in place or exported without its loss list', () => {
  const { registryProblems } = require('../scripts/no-office-matrix.cjs')
  const formats = require('../shared/electron/formats.cjs')
  const data = JSON.parse(fs.readFileSync(path.join(ROOT, 'shared', 'electron', 'formats.json'), 'utf8'))
  assert.deepEqual(registryProblems('calc'), [])
  const xls = data.formats.find((format) => format.id === 'xls')
  xls.workspaces.calc.save = 'in-place'
  xls.workspaces.calc.export = 'native'
  const broken = formats.createRegistry(data)
  const problems = registryProblems('calc', broken).join('\n')
  assert.match(problems, /Save writes the legacy binary xls/)
  assert.match(problems, /Export offers the legacy binary xls without the explicit "values only" loss list/)
})

test('acceptance contract: open after export goes through openInSimple and never through the shell', { timeout: 120_000 }, async () => {
  const { code, stdout, stderr } = await runScript('io-acceptance.cjs', ['--contract-only'])
  assert.equal(code, 0, `${stdout}\n${stderr}`)
  assert.match(stdout, /^ok {3}contract: export → open after export goes through openInSimple/m)
  assert.match(stdout, /shell untouched/)
})

test('an enabled workspace without an acceptance adapter fails with the file to write', (t) => {
  const harness = require('../scripts/io-harness.cjs')
  const empty = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-io-adapters-'))
  t.after(() => fs.rmSync(empty, { recursive: true, force: true }))
  assert.throws(() => harness.loadAdapter('calc', empty), /calc is enabled in simple\/shared\/manifest\.json but has no acceptance adapter\. Write .*calc\.cjs/)
})

test('the gates follow the enabled flags in shared/manifest.json', (t) => {
  const harness = require('../scripts/io-harness.cjs')
  const shared = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-io-manifest-'))
  t.after(() => fs.rmSync(shared, { recursive: true, force: true }))
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'shared', 'manifest.json'), 'utf8'))
  const real = Object.entries(manifest.workspaces).filter(([, workspace]) => workspace.enabled).map(([name]) => name)
  assert.deepEqual(harness.enabledWorkspaces(), real)
  for (const [name, workspace] of Object.entries(manifest.workspaces)) workspace.enabled = name === 'docs'
  fs.writeFileSync(path.join(shared, 'manifest.json'), JSON.stringify(manifest))
  assert.deepEqual(harness.enabledWorkspaces(shared), ['docs'])
})

test('npm run test:io runs both gates in full', () => {
  const scripts = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8')).scripts
  assert.match(scripts['test:io'], /no-office-matrix\.cjs/)
  assert.match(scripts['test:io'], /io-acceptance\.cjs/)
})
