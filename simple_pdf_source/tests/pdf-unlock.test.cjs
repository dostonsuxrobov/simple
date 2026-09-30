'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { unlockPdf } = require('../electron/pdf-unlock.cjs')

async function encryptedPdf(options) {
  const mupdf = await import('mupdf')
  const doc = new mupdf.PDFDocument()
  doc.insertPage(-1, doc.addPage([0, 0, 200, 200], 0, doc.newDictionary(), ''))
  return new Uint8Array(doc.saveToBuffer(options).asUint8Array())
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
  const { PDFDocument } = require('pdf-lib')
  assert.equal((await PDFDocument.load(result.data)).getPageCount(), 1)
})

test('owner-password-only PDFs unlock without a prompt', async () => {
  const bytes = await encryptedPdf('encrypt=aes-128,owner-password=own')
  const result = await unlockPdf(bytes, '')
  assert.equal(result.status, 'unlocked')
  const { PDFDocument } = require('pdf-lib')
  assert.equal((await PDFDocument.load(result.data)).getPageCount(), 1)
})
