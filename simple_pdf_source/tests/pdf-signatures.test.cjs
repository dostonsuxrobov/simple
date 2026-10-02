'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const test = require('node:test')
const { PDFDocument } = require('pdf-lib')
const { detectSignatures, documentSignatureStatus, fileSignatureStatus } = require('../electron/pdf-signatures.cjs')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { buildRawPdf, patchByteRange } = require('./helpers/raw-pdf.cjs')

const CONTENTS = '0'.repeat(128)

function signedPdf(signatureDictionary, { sigFlags = 3 } = {}) {
  return patchByteRange(buildRawPdf([
    `<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] /SigFlags ${sigFlags} >> >>`,
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Annots [4 0 R] >>',
    '<< /FT /Sig /T (Signature1) /Type /Annot /Subtype /Widget /Rect [0 0 0 0] /P 3 0 R /V 5 0 R /F 132 >>',
    signatureDictionary,
  ]))
}

// iText and Acrobat write dictionaries without spaces between names.
const COMPACT = `<</Type/Sig/Filter/Adobe.PPKLite/SubFilter/adbe.pkcs7.detached/ByteRange[0 0000000000 0000000000 0000000000]/Contents<${CONTENTS}>>>`
const SPACED = `<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 0000000000 0000000000 0000000000] /Contents <${CONTENTS}> >>`

function plainPdf(catalogExtra = '') {
  return buildRawPdf([
    `<< /Type /Catalog /Pages 2 0 R ${catalogExtra}>>`,
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>',
  ])
}

test('a compact /Type/Sig signature with /ByteRange is detected and intact', () => {
  const status = detectSignatures(signedPdf(COMPACT, { sigFlags: 1 }))
  assert.equal(status.signed, true)
  assert.equal(status.intact, true)
})

test('a spaced /Type /Sig signature is detected', () => {
  assert.equal(detectSignatures(signedPdf(SPACED)).signed, true)
})

test('an AcroForm with /SigFlags 3 counts as signed', () => {
  assert.equal(detectSignatures(plainPdf('/AcroForm << /Fields [] /SigFlags 3 >> ')).signed, true)
})

test('/SigFlags text without an AcroForm, and unsigned signature fields, are not signatures', () => {
  assert.equal(detectSignatures(plainPdf('/Note (/SigFlags 3) ')).signed, false)
  assert.equal(detectSignatures(plainPdf('/AcroForm << /Fields [] /SigFlags 1 >> ')).signed, false)
  // Prepared for signing but never signed: the range is still a placeholder.
  const prepared = buildRawPdf([
    '<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [] /SigFlags 1 >> >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>',
    `<< /Type /Sig /ByteRange [0 /********** /********** /**********] /Contents <${CONTENTS}> >>`,
  ])
  assert.equal(detectSignatures(prepared).signed, false)
})

test('a plain PDF is not signed', () => {
  assert.deepEqual(detectSignatures(plainPdf()), { signed: false, count: 0, intact: false })
})

test('a signature inside a compressed object stream is still found', async () => {
  const rewritten = await (await PDFDocument.load(signedPdf(COMPACT, { sigFlags: 1 }))).save({ useObjectStreams: true })
  assert.equal(Buffer.from(rewritten).includes('/ByteRange'), false, 'the fixture really compresses it')
  const status = detectSignatures(rewritten)
  assert.equal(status.signed, true)
  assert.equal(status.intact, false, 'a full rewrite no longer matches the signed byte ranges')
})

test('structural detection: signed /FT /Sig field, /SigFlags AppendOnly, /Perms', async () => {
  assert.equal(documentSignatureStatus(await PDFDocument.load(signedPdf(SPACED, { sigFlags: 1 }))).signed, true)
  assert.equal(documentSignatureStatus(await PDFDocument.load(plainPdf('/AcroForm << /Fields [] /SigFlags 3 >> '))).signed, true)
  assert.equal(documentSignatureStatus(await PDFDocument.load(plainPdf('/Perms << /DocMDP << /Type /Sig >> >> '))).signed, true)
  assert.equal(documentSignatureStatus(await PDFDocument.load(plainPdf('/AcroForm << /Fields [] /SigFlags 1 >> '))).signed, false)
  assert.equal(documentSignatureStatus(await PDFDocument.load(plainPdf())).signed, false)
})

test('Open reports signatureDetected for compact signatures, and Save never overwrites a signed original in place', async () => {
  const { invoke, saveDialogs, tempFile } = loadMain()
  const original = signedPdf(COMPACT)
  const target = tempFile('signed.pdf', original)
  const payload = await invoke('file:open-path', target)
  assert.equal(payload.signatureDetected, true)
  assert.equal((await fileSignatureStatus(target)).intact, true)

  const result = await invoke('pdf:save', { data: new Uint8Array(Buffer.from('%PDF-1.7 edited')), path: target, name: 'signed.pdf', forceDialog: false })
  assert.equal(result, null, 'the save dialog was cancelled')
  assert.equal(saveDialogs.length, 1, 'a Save As dialog was shown instead of overwriting')
  assert.match(saveDialogs[0].defaultPath, /signed \(edited\)\.pdf$/)
  assert.deepEqual(fs.readFileSync(target), original, 'the signed original is untouched')
})

test('a file whose signature is already broken can be saved in place', async () => {
  const { invoke, saveDialogs, tempFile } = loadMain()
  const before = saveDialogs.length
  const rewritten = await (await PDFDocument.load(signedPdf(COMPACT))).save({ useObjectStreams: false })
  const target = tempFile('rewritten.pdf', rewritten)
  const payload = await invoke('file:open-path', target)
  assert.equal(payload.signatureDetected, true, 'still reported, so the renderer warns')
  const data = new Uint8Array(Buffer.from('%PDF-1.7 next'))
  const result = await invoke('pdf:save', { data, path: target, name: 'rewritten.pdf', forceDialog: false })
  assert.equal(result.path, target)
  assert.equal(saveDialogs.length, before)
  assert.deepEqual(fs.readFileSync(target), Buffer.from(data))
})
