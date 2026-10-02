'use strict'

// The static LOCAL-ONLY guard (scripts/local-only-guard.cjs, run by verify.cjs):
// Simple's app code never talks to the network, never names a cloud service,
// and never hands a document to another program.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { checkLocalOnly, checkNamesOnly, checkText, findServiceNames, stripComments } = require('../scripts/local-only-guard.cjs')

const ROOT = path.resolve(__dirname, '..')

test('the real app folders pass the local-only guard', () => {
  const result = checkLocalOnly()
  assert.equal(result.ok, true, result.message)
  assert.ok(result.files > 20, `checked only ${result.files} files`)
})

test('network modules and requests are reported with file and line', () => {
  const cases = [
    "const https = require('https')",
    "const net = require('node:net')",
    "import tls from 'node:tls'",
    "const dgram = await import('dgram')",
    "await fetch('https://example.com/x')",
    'await fetch(url)',
    "const request = net.request({ url: 'x' })",
    'session.defaultSession.downloadURL(address)',
    "const { autoUpdater } = require('electron')",
    'const socket = new WebSocket(address)',
    'navigator.sendBeacon(address, body)',
  ]
  for (const line of cases) {
    const problems = checkText('shared/electron/x.cjs', `'use strict'\n${line}\n`)
    assert.ok(problems.length, `not reported: ${line}`)
    assert.match(problems[0], /^shared\/electron\/x\.cjs:2: /, line)
  }
})

test('other network calls, download programs and download commands are reported', () => {
  const cases = [
    "const dns = require('node:dns')\ndns.lookup(host, done)",
    "import { lookup } from 'dns'",
    'const response = await session.defaultSession.fetch(url)',
    'const response = await window.fetch(url)',
    'await win.loadURL(remoteUrl)',
    "await win.loadURL('https://example.com/')",
    "spawn('powershell.exe', ['-Command', script])",
    "execFile('C:\\\\Windows\\\\System32\\\\curl.exe', [address])",
    "const script = 'Invoke-WebRequest -Uri $u -OutFile $f'",
    "execFileSync(certutil, ['-urlcache', '-f', address, target])",
  ]
  for (const line of cases) {
    const problems = checkText('launcher/x.cjs', `'use strict'\n${line}\n`)
    assert.ok(problems.length, `not reported: ${line}`)
  }
  assert.deepEqual(checkText('launcher/x.cjs', "await win.loadURL(pathToFileURL(page).href)\nawait win.loadURL('data:text/html,ok')\nwin.loadFile(page)\n"), [])
})

test('programs are started by absolute path, never by a bare name', () => {
  for (const line of ["execFileSync('reg.exe', ['query', key])", "execFile('taskkill.exe', ['/PID', pid])", "spawn('icacls', [target])"]) {
    const problems = checkText('shared/electron/x.cjs', `${line}\n`)
    assert.equal(problems.length, 1, line)
    assert.match(problems[0], /bare name/)
  }
  assert.deepEqual(checkText('shared/electron/x.cjs', "execFile(systemProgram('reg.exe'), args)\nspawn(executable, args)\nexecFile('C:\\\\Windows\\\\System32\\\\reg.exe', args)\n"), [])
  // The real spawn sites use System32 by absolute path.
  for (const file of ['shared/electron/text-codec.cjs', 'shared/electron/office-engine.cjs', 'shared/electron/safe-write.cjs', 'launcher/associations.cjs']) {
    assert.match(fs.readFileSync(path.join(ROOT, file), 'utf8'), /systemProgram\('[a-z]+\.exe'\)/, file)
  }
})

test('messages in app code never ask the user to install or download anything', () => {
  assert.equal(checkText('launcher/combine-policy.cjs', "throw new CombineError('ENGINE', 'Install LibreOffice to combine this file.')\n").length, 1)
  assert.equal(checkText('shared/electron/x.cjs', 'const message = `Download the converter for ${name}.`\n').length, 1)
  // Event names and keys are not prose.
  assert.deepEqual(checkText('shared/electron/x.cjs', "ses.on('will-download', stop)\nconst key = 'installedAt'\n"), [])
})

test('local fetches, XML namespaces and the development server are allowed', () => {
  const text = [
    "const bytes = await fetch('data:application/pdf;base64,AAAA')",
    "const blob = await fetch(`blob:${id}`)",
    "const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'",
    "const registry = '<oor:items xmlns:oor=\"http://openoffice.org/2001/registry\">'",
    "const dev = 'http://localhost:5173/'",
    "const svg = 'http://www.w3.org/2000/svg'",
  ].join('\n')
  assert.deepEqual(checkText('shared/electron/x.cjs', text), [])
})

test('remote addresses are reported, in code but not in comments', () => {
  assert.equal(checkText('launcher/x.cjs', "const page = 'https://example.com/help'\n").length, 1)
  assert.deepEqual(checkText('launcher/x.cjs', '// see https://example.com/help\nconst a = 1\n'), [])
})

// The service names are never written out in Simple's source, tests included:
// the fixtures are ROT13-encoded and decoded only while the test runs.
function rot13(text) {
  return text.replace(/[a-z]/gi, (char) => {
    const base = char <= 'Z' ? 65 : 97
    return String.fromCharCode(((char.charCodeAt(0) - base + 13) % 26) + base)
  })
}
const ENCODED_NAMES = ['BarQevir', 'barqevir', 'Bar Qevir', 'Bar-Qevir', 'BARQEVIR', 'Qebcobk', 'Tbbtyr Qevir', 'vPybhq', 'FunerCbvag', 'obk.pbz', 'Obk Qevir', 'cPybhq']

test('cloud service names are reported anywhere, comments included, in any case', () => {
  for (const name of ENCODED_NAMES.map(rot13)) {
    const problems = checkText('shared/electron/io-core.cjs', `// Files in ${name} folders\nconst a = 1\n`)
    assert.equal(problems.length, 1, rot13(name))
    assert.match(problems[0], /another program is using this file/)
  }
  // Inside identifiers too: camelCase and snake_case parts are compared.
  assert.equal(checkText('shared/electron/x.cjs', `const ${rot13('vfBarQevirCngu')} = true\n`).length, 1)
  assert.equal(checkText('shared/electron/x.cjs', `const ${rot13('VF_BAR_QEVIR')} = true\n`).length, 1)
  assert.equal(checkText('shared/electron/io-catalog.json', `{"a": "Saved to ${rot13('BarQevir')}"}`).length, 1)
  assert.deepEqual(checkText('shared/electron/x.cjs', 'const inbox = mailbox.count\nconst drive = describeDrive(file)\n'), [])
})

test('the names live only as hashes, and tests and scripts are checked for them as well', () => {
  const source = fs.readFileSync(path.join(ROOT, 'scripts', 'local-only-guard.cjs'), 'utf8')
  assert.deepEqual(findServiceNames(source), [], 'the guard itself names no service')
  const data = JSON.parse(fs.readFileSync(path.join(ROOT, 'scripts', 'local-only-names.json'), 'utf8'))
  assert.ok(data.hashes.length >= 10 && data.hashes.every((hash) => /^[0-9a-f]{64}$/.test(hash)))
  assert.equal(findServiceNames(`see ${rot13('Qebcobk')}`).length, 1)
  assert.deepEqual(checkNamesOnly('tests/x.test.cjs', "require('node:http')\n"), [], 'tests are checked only for names')
  assert.equal(checkNamesOnly('tests/x.test.cjs', `// ${rot13('vPybhq')}\n`).length, 1)
})

test('documents are never handed to another program', () => {
  for (const line of ['shell.openPath(filePath)', 'electron.shell.openPath(target)', 'shell.openItem(target)', 'shell.openExternal(pathToFileURL(target).href)', "shell.openExternal('https://example.com')"]) {
    assert.ok(checkText('launcher/main.cjs', `${line}\n`).length, `not reported: ${line}`)
  }
  assert.deepEqual(checkText('launcher/main.cjs', "shell.openExternal('ms-settings:defaultapps')\nshell.showItemInFolder(target)\n"), [])
  // The preload bridge's own method name is not a call to Electron's shell.
  assert.deepEqual(checkText('shared/preload/io-bridge.cjs', "openPath: (filePath) => invoke('io:open-in-simple', filePath),\n"), [])
})

test('user-facing text never asks the user to download or install anything', () => {
  const catalog = JSON.stringify({ notices: { engine: 'Install the compatibility engine to open this file.' }, ok: { text: 'Saved a modern copy next to the original.' } })
  const problems = checkText('shared/electron/io-catalog.json', catalog)
  assert.equal(problems.length, 1)
  assert.match(problems[0], /notices\.engine/)
  assert.equal(checkText('launcher/index.html', '<p>Download the converter</p>').length, 1)
  assert.equal(checkText('launcher/renderer.js', "status.textContent = 'Installing…'\n").length, 1)
  assert.deepEqual(checkText('launcher/renderer.js', '// installs the click handler\nbutton.onclick = run\n'), [])
})

test('comment stripping keeps strings and line numbers', () => {
  const source = "const a = '// not a comment' // a comment\n/* block\ncomment */ const b = \"/* kept */\"\n"
  const stripped = stripComments(source)
  assert.equal(stripped.split('\n').length, source.split('\n').length)
  assert.match(stripped, /'\/\/ not a comment'/)
  assert.match(stripped, /"\/\* kept \*\/"/)
  assert.doesNotMatch(stripped, /\/\/ a comment|block/)
})

test('tests and fixtures are not app code; a tree with problems fails with every file named', (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-local-only-'))
  t.after(() => fs.rmSync(root, { recursive: true, force: true }))
  fs.mkdirSync(path.join(root, 'shared', 'electron'), { recursive: true })
  fs.mkdirSync(path.join(root, 'shared', 'tests'), { recursive: true })
  fs.mkdirSync(path.join(root, 'launcher'), { recursive: true })
  fs.writeFileSync(path.join(root, 'shared', 'electron', 'ok.cjs'), "'use strict'\nmodule.exports = {}\n")
  fs.writeFileSync(path.join(root, 'shared', 'tests', 'server.cjs'), "require('node:http')\n")
  fs.writeFileSync(path.join(root, 'shared', 'electron', 'thing.test.cjs'), "require('node:http')\n")
  fs.writeFileSync(path.join(root, 'launcher', 'main.cjs'), "const { shell } = require('electron')\nshell.openPath(file)\n")
  const result = checkLocalOnly({ root })
  assert.equal(result.ok, false)
  assert.deepEqual(result.problems.map((problem) => problem.split(':').slice(0, 2).join(':')), ['launcher/main.cjs:2'])
  assert.match(result.message, /local-only-guard\.cjs/)
  assert.ok(fs.existsSync(path.join(ROOT, 'scripts', 'local-only-guard.cjs')))
})
