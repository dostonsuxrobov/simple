'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { PDFDocument } = require('pdf-lib')
const { hasEncryptionDictionary, unlockPdf } = require('../electron/pdf-unlock.cjs')
const { preparePrintPdf } = require('../electron/pdf-print.cjs')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { encryptedPdf: textPdf, linearizedLayout } = require('./helpers/encrypted-fixtures.cjs')
const { pageTexts } = require('./helpers/pdf-test-utils.cjs')

async function encryptedPdf(options) {
  const mupdf = await import('mupdf')
  const doc = new mupdf.PDFDocument()
  doc.insertPage(-1, doc.addPage([0, 0, 200, 200], 0, doc.newDictionary(), ''))
  return new Uint8Array(doc.saveToBuffer(options).asUint8Array())
}

// The check this replaces only searched the final 256 KB of the file.
function tailMentionsEncrypt(bytes) {
  return Buffer.from(bytes).subarray(Math.max(0, bytes.length - 262_144)).includes('/Encrypt')
}

test('plain PDFs are left alone', async () => {
  const { status } = await unlockPdf(await encryptedPdf(''))
  assert.equal(status, 'none')
})

test('user-password PDFs need the right password and unlock to an editable copy', async () => {
  const bytes = await encryptedPdf('encrypt=aes-256,user-password=abc,owner-password=own')
  assert.equal((await unlockPdf(bytes, '')).status, 'needs-password')
  assert.equal((await unlockPdf(bytes, 'nope')).status, 'wrong-password')
  const result = await unlockPdf(bytes, 'abc')
  assert.equal(result.status, 'unlocked')
  assert.equal((await PDFDocument.load(result.data)).getPageCount(), 1)
})

test('owner-password-only PDFs unlock without a prompt', async () => {
  const bytes = await encryptedPdf('encrypt=aes-128,owner-password=own')
  const result = await unlockPdf(bytes, '')
  assert.equal(result.status, 'unlocked')
  assert.equal((await PDFDocument.load(result.data)).getPageCount(), 1)
})

test('linearized owner-locked files over 256 KB are detected and unlocked', async () => {
  const bytes = linearizedLayout(await textPdf('Protected statement', 'encrypt=aes-128,owner-password=own'))
  assert.ok(bytes.length > 262_144)
  assert.equal(tailMentionsEncrypt(bytes), false, 'the fixture keeps /Encrypt out of the last 256 KB')
  assert.equal(hasEncryptionDictionary(bytes), true)
  const result = await unlockPdf(bytes, '')
  assert.equal(result.status, 'unlocked')
  const [text] = await pageTexts(result.data)
  assert.match(text, /Protected statement/)
})

test('linearized user-password files ask for the password', async () => {
  const bytes = linearizedLayout(await textPdf('Protected statement', 'encrypt=aes-256,user-password=pw,owner-password=own'))
  assert.equal(tailMentionsEncrypt(bytes), false)
  assert.equal((await unlockPdf(bytes, '')).status, 'needs-password')
  assert.equal((await unlockPdf(bytes, 'nope')).status, 'wrong-password')
  assert.equal((await unlockPdf(bytes, 'pw')).status, 'unlocked')
})

test('only a trailer /Encrypt entry counts, not /EncryptMetadata or loose text', () => {
  assert.equal(hasEncryptionDictionary(Buffer.from('<< /EncryptMetadata false >> (/Encrypt) /Encrypted')), false)
  assert.equal(hasEncryptionDictionary(Buffer.from('trailer\n<</Size 9/Root 1 0 R/Encrypt 7 0 R>>')), true)
  assert.equal(hasEncryptionDictionary(Buffer.from('trailer << /Encrypt << /Filter /Standard >> >>')), true)
})

test('Open reports encryption, and edits of locked bytes fail with coded errors instead of pdf-lib errors', async () => {
  const { invoke, tempFile } = loadMain()
  const ownerLocked = linearizedLayout(await textPdf('Protected statement', 'encrypt=aes-128,owner-password=own'))
  const payload = await invoke('file:open-path', tempFile('owner-locked.pdf', ownerLocked))
  assert.equal(payload.encrypted, true)
  const plainPayload = await invoke('file:open-path', tempFile('plain.pdf', await (await PDFDocument.create()).save()))
  assert.equal(plainPayload.encrypted, false)

  const rotate = { type: 'rotate', indices: [0], degrees: 90 }
  await assert.rejects(invoke('pdf:mutate', ownerLocked, rotate), (error) => {
    assert.match(error.message, /OWNER_LOCKED: /)
    assert.doesNotMatch(error.message, /PDFDocument\.load|ignoreEncryption/)
    return true
  })
  const passwordLocked = linearizedLayout(await textPdf('Protected statement', 'encrypt=aes-256,user-password=pw,owner-password=own'))
  await assert.rejects(invoke('pdf:flatten-overlays', passwordLocked, [], {}, {}), /PASSWORD_REQUIRED: /)

  // The renderer's unlock flow then makes every operation work.
  const unlocked = await invoke('pdf:unlock', ownerLocked, '')
  assert.equal(unlocked.status, 'unlocked')
  const rotated = await invoke('pdf:mutate', unlocked.data, rotate)
  assert.equal((await PDFDocument.load(rotated)).getPage(0).getRotation().angle, 90)
  const flattened = await invoke('pdf:flatten-overlays', unlocked.data, [{
    id: 'mark', type: 'highlight', pageIndex: 0, rect: { x: 10, y: 10, width: 50, height: 10 }, color: [1, 1, 0], opacity: 0.3,
  }], {}, {})
  assert.match((await pageTexts(flattened))[0], /Protected statement/)
})

test('adding pages from protected PDFs: owner-locked sources are decrypted, password sources are refused by name', async () => {
  const { invoke } = loadMain()
  const base = await PDFDocument.create()
  base.addPage([300, 300])
  const baseBytes = await base.save()
  const ownerLocked = linearizedLayout(await textPdf('Inserted protected page', 'encrypt=aes-128,owner-password=own'))
  const inserted = await invoke('pdf:insert-dropped-files', baseBytes, 1, [{ name: 'locked.pdf', data: ownerLocked }])
  assert.equal(inserted.added, 1)
  assert.match((await pageTexts(inserted.data))[1], /Inserted protected page/)

  const passwordLocked = await textPdf('Secret', 'encrypt=aes-256,user-password=pw,owner-password=own')
  await assert.rejects(
    invoke('pdf:insert-dropped-files', baseBytes, 1, [{ name: 'needs-password.pdf', data: passwordLocked }]),
    /PASSWORD_REQUIRED: “needs-password\.pdf” is password-protected/,
  )
})

test('printing an owner-locked PDF prints its decrypted content, not blank pages', async () => {
  const ownerLocked = linearizedLayout(await textPdf('Protected statement', 'encrypt=aes-128,owner-password=own'))
  const printable = await preparePrintPdf(ownerLocked, { paperSize: 'A4' })
  assert.match((await pageTexts(printable))[0], /Protected statement/)
  const passwordLocked = await textPdf('Protected statement', 'encrypt=aes-256,user-password=pw,owner-password=own')
  await assert.rejects(preparePrintPdf(passwordLocked, {}), (error) => error.code === 'PASSWORD_REQUIRED')
})
