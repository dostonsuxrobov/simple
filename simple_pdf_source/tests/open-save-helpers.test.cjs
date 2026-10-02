'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const test = require('node:test')
const ts = require('typescript')
const { PDFDocument } = require('pdf-lib')
const { encryptedPdf, linearizedLayout } = require('./helpers/encrypted-fixtures.cjs')
const { hasEncryptionDictionary } = require('../electron/pdf-unlock.cjs')

// src/lib/openSave.ts holds the renderer's open/save decisions (no DOM);
// transpile it on the fly like tests/viewer-clipboard.test.cjs does.
const helpers = (async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-open-save-'))
  const sourcePath = path.resolve(__dirname, '..', 'src', 'lib', 'openSave.ts')
  const compiled = ts.transpileModule(await fs.readFile(sourcePath, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 },
    fileName: sourcePath,
    reportDiagnostics: true,
  })
  const errors = (compiled.diagnostics || []).filter((item) => item.category === ts.DiagnosticCategory.Error)
  assert.equal(errors.length, 0, errors.map((item) => item.messageText).join('\n'))
  const target = path.join(directory, 'openSave.mjs')
  await fs.writeFile(target, compiled.outputText)
  process.once('exit', () => { try { require('node:fs').rmSync(directory, { recursive: true, force: true }) } catch { /* best effort */ } })
  return import(pathToFileURL(target).href)
})()

test('a password-protected PDF is recognised anywhere in the file, as the main process does', async () => {
  const { mayBeEncrypted } = await helpers
  const classic = await encryptedPdf('Locked', 'encrypt=aes-256,user-password=pw,owner-password=owner')
  // Linearized layout: /Encrypt only near the start, 300 KB of padding after it.
  const linearized = linearizedLayout(classic)
  assert.ok(linearized.length > 300 * 1024)
  assert.equal(linearized.subarray(linearized.length - 262_144).includes(Buffer.from('/Encrypt')), false, 'the tail does not mention /Encrypt')
  for (const bytes of [classic, linearized]) {
    assert.equal(mayBeEncrypted(new Uint8Array(bytes)), true)
    assert.equal(hasEncryptionDictionary(bytes), true, 'and the main process agrees')
  }
})

test('plain files, /EncryptMetadata and a /Encrypt split across scan chunks are judged correctly', async () => {
  const { mayBeEncrypted } = await helpers
  const plain = await (await PDFDocument.create()).save()
  assert.equal(mayBeEncrypted(plain), false)
  const metadataOnly = Buffer.concat([Buffer.from(plain), Buffer.from('\n% /EncryptMetadata false\n')])
  assert.equal(mayBeEncrypted(new Uint8Array(metadataOnly)), false)
  // The name straddles the 8 MB chunk boundary.
  const big = Buffer.alloc(8 * 1024 * 1024 + 64, 0x20)
  Buffer.from('/Encrypt 5 0 R').copy(big, 8 * 1024 * 1024 - 4)
  assert.equal(mayBeEncrypted(new Uint8Array(big)), true)
})

test('a signed original is saved as an "(edited)" copy beside it', async () => {
  const { signedCopyName } = await helpers
  assert.equal(signedCopyName({ path: 'C:\\Contracts\\Lease.PDF', name: 'Lease.PDF' }), 'C:\\Contracts\\Lease (edited).pdf')
  assert.equal(signedCopyName({ path: null, name: 'scan.pdf' }), 'scan (edited).pdf')
})

test('only failures and changed values block a save; notes are summarised in one line', async () => {
  const { blockingProblems, summarizeProblems, problemText } = await helpers
  const failure = { code: 'FORM_FIELD_NOT_FOUND', message: 'Form field “x” was not found in this PDF, so its value was not saved.' }
  const truncated = { code: 'FORM_VALUE_TRUNCATED', message: 'only “10001” was saved', blockingMessage: 'Form field “zip” allows at most 5 characters (you typed 6). Shorten it and save again.', dataLoss: true }
  const shrunk = { code: 'TEXT_SHRUNK', message: 'Text on page 1 was reduced from 12 pt to 10 pt.' }
  assert.deepEqual(blockingProblems({ failures: [failure], warnings: [truncated, shrunk] }), [failure, truncated])
  assert.deepEqual(blockingProblems({ failures: [], warnings: [shrunk] }), [])
  assert.equal(problemText(truncated), truncated.blockingMessage)
  assert.equal(summarizeProblems([failure, truncated, failure]), `${failure.message} (+1 more)`)
  assert.equal(summarizeProblems([]), '')
})
