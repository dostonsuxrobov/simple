const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const JSZip = require('jszip')
const {
  IMPORT_FORMATS,
  OPEN_DIALOG_FILTERS,
  OPEN_FORMATS,
  documentOpenKind,
  isOpenableDocumentPath,
  sniffDocumentBytes,
  validateDocxBytes,
  wordPackageInfo,
} = require('../electron/docx-files.cjs')

// DOC-017 / DOC-SIE-15: every format Simple Docs opens, found by name and corrected by
// content, without LibreOffice.

const MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
const TEMPLATE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'
const MACRO = 'application/vnd.ms-word.document.macroEnabled.main+xml'
const MACRO_TEMPLATE = 'application/vnd.ms-word.template.macroEnabledTemplate.main+xml'

async function wordPackage(contentType, extra = {}) {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${contentType}"/></Types>`)
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  zip.file('word/document.xml', '<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body></w:document>')
  for (const [name, data] of Object.entries(extra)) zip.file(name, data)
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function odt(mimetype = 'application/vnd.oasis.opendocument.text', first = true) {
  const zip = new JSZip()
  if (first) zip.file('mimetype', mimetype, { compression: 'STORE' })
  zip.file('content.xml', '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>')
  if (!first) zip.file('mimetype', mimetype)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const OLE = Buffer.concat([Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]), Buffer.alloc(600)])
const kind = (name, bytes) => documentOpenKind(name, Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes))

test('Word packages, templates, legacy Word, RTF, OpenDocument, web pages, Markdown and text are openable', () => {
  for (const name of ['a.docx', 'a.DOCM', 'a.dotx', 'a.dotm', 'a.doc', 'a.rtf', 'a.odt', 'a.html', 'a.htm', 'a.md', 'a.markdown', 'a.txt', String.raw`C:\Users\x\Notes.TXT`]) {
    assert.equal(isOpenableDocumentPath(name), true, name)
  }
  for (const name of ['a.pdf', 'a.xlsx', 'a.odp', 'a', 'a.docx.zip', null, undefined, 42]) assert.equal(isOpenableDocumentPath(name), false, String(name))
  assert.deepEqual([...IMPORT_FORMATS].sort(), ['html', 'md', 'odt', 'rtf', 'txt'])
})

test('the Open dialog lists "All supported documents" first, then each family, then All files', () => {
  const [all, ...rest] = OPEN_DIALOG_FILTERS
  assert.equal(all.name, 'All supported documents')
  assert.deepEqual([...all.extensions].sort(), Object.keys(OPEN_FORMATS).map((extension) => extension.slice(1)).sort())
  assert.deepEqual(rest.map((filter) => filter.name), ['Word documents', 'Word templates', 'OpenDocument text', 'Rich Text', 'Web pages', 'Markdown', 'Plain text', 'All files'])
  const covered = new Set(rest.flatMap((filter) => filter.extensions))
  for (const extension of all.extensions) assert.ok(covered.has(extension), `${extension} has its own filter`)
  assert.deepEqual(rest.at(-1).extensions, ['*'])
  // Both open dialogs in main use these filters.
  const main = fs.readFileSync(path.join(__dirname, '..', 'electron', 'main.cjs'), 'utf8')
  assert.equal((main.match(/filters: OPEN_DIALOG_FILTERS/g) || []).length, 2)
})

test('the renderer\'s drop check and title cleanup know the same extensions as main', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'src', 'main.ts'), 'utf8')
  const expected = Object.keys(OPEN_FORMATS).map((extension) => extension.slice(1)).sort()
  const extensionsOf = (pattern) => {
    const regex = new RegExp(pattern)
    return expected.filter((extension) => regex.test(`file.${extension}`) && regex.test(`FILE.${extension.toUpperCase()}`))
  }
  const drop = /const OPENABLE_DOCUMENT = \/(.+?)\/i;/.exec(source)
  assert.ok(drop, 'OPENABLE_DOCUMENT is declared in src/main.ts')
  assert.deepEqual(extensionsOf(new RegExp(drop[1], 'i')), expected)
  for (const extension of ['pdf', 'xlsx', 'png', 'zip']) assert.equal(new RegExp(drop[1], 'i').test(`file.${extension}`), false, extension)
  const title = /name\.replace\(\/(.+?)\/i, ""\)/.exec(source)
  assert.ok(title, 'fileNameWithoutExtension strips document extensions')
  assert.deepEqual(extensionsOf(new RegExp(title[1], 'i')), expected)
})

test('content decides the format: renamed files open as what they are', async () => {
  const docx = await wordPackage(MAIN)
  const open = await odt()
  assert.equal(sniffDocumentBytes(docx), 'word')
  assert.equal(sniffDocumentBytes(open), 'odt')
  assert.equal(sniffDocumentBytes(await odt(undefined, false)), 'odf')
  assert.equal(sniffDocumentBytes(await odt('application/vnd.oasis.opendocument.spreadsheet')), 'zip')
  assert.equal(sniffDocumentBytes(OLE), 'ole')
  assert.equal(sniffDocumentBytes(Buffer.from(String.raw`{\rtf1\ansi x}`)), 'rtf')
  assert.equal(sniffDocumentBytes(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(' \r\n<!DOCTYPE html><p>x')])), 'html')
  assert.equal(sniffDocumentBytes(Buffer.from('\ufeff<html>x', 'utf16le')), 'html')
  assert.equal(sniffDocumentBytes(Buffer.from('%PDF-1.7\n1 0 obj')), 'binary')
  assert.equal(sniffDocumentBytes(Buffer.from([0x89, 0x50, 0x4e, 0x47, 0, 0, 0, 0])), 'binary')
  assert.equal(sniffDocumentBytes(Buffer.alloc(0)), 'empty')
  assert.equal(sniffDocumentBytes(Buffer.from('# Notes\n\nplain')), 'text')

  assert.deepEqual(kind('report.docx', docx), { kind: 'docx', renamed: false })
  assert.deepEqual(kind('report.dotm', docx), { kind: 'dotm', renamed: false })
  assert.deepEqual(kind('report.rtf', docx), { kind: 'docx', renamed: true })
  assert.deepEqual(kind('notes.txt', String.raw`{\rtf1\ansi x}`), { kind: 'rtf', renamed: true })
  assert.deepEqual(kind('letter.rtf', String.raw`{\rtf1\ansi x}`), { kind: 'rtf', renamed: false })
  assert.deepEqual(kind('minutes.docx', open), { kind: 'odt', renamed: true })
  assert.deepEqual(kind('minutes.odt', await odt(undefined, false)), { kind: 'odt', renamed: false })
  assert.deepEqual(kind('page.docx', '<!doctype html><html><body>x'), { kind: 'html', renamed: true })
  assert.deepEqual(kind('old.rtf', OLE), { kind: 'doc', renamed: true })
  // An OLE file named .docx is an encrypted package: the DOCX reader explains.
  assert.deepEqual(kind('locked.docx', OLE), { kind: 'docx', renamed: false })
  assert.throws(() => validateDocxBytes(OLE), /Password-protected/)
  // .doc keeps its own reader, which sniffs RTF, web pages and DOCX itself.
  assert.deepEqual(kind('legacy.doc', String.raw`{\rtf1 x}`), { kind: 'doc', renamed: false })
  // Markdown and text stay what their names say, even when they start with markup.
  assert.deepEqual(kind('readme.md', '<!-- note --><p>hi</p>'), { kind: 'md', renamed: false })
  assert.deepEqual(kind('page.html', 'just text'), { kind: 'html', renamed: false })
  assert.deepEqual(kind('flat.odt', '<?xml version="1.0"?><office:document>'), { kind: 'odt', renamed: false })
  // Through "All files": a file without a known extension opens when its content says what it is.
  assert.deepEqual(kind('README', 'Plain words'), { kind: 'txt', renamed: false })
  assert.deepEqual(kind('saved-page', '<html><body>x'), { kind: 'html', renamed: false })
  assert.deepEqual(kind('archive', docx), { kind: 'docx', renamed: true })
})

test('what Simple Docs cannot read is refused with a plain message', async () => {
  assert.throws(() => kind('a.pdf', '%PDF-1.4\n1 0 obj'), /can’t open this kind of file/)
  const workbook = new JSZip()
  workbook.file('xl/workbook.xml', '<workbook/>')
  const xlsx = await workbook.generateAsync({ type: 'nodebuffer' })
  assert.throws(() => kind('a.xlsx', xlsx), /can’t open this kind of file/)
  const spreadsheet = await odt('application/vnd.oasis.opendocument.spreadsheet')
  assert.throws(() => kind('sheet.odt', spreadsheet), /not a readable OpenDocument document/)
  assert.throws(() => kind('a.md', Buffer.from([0x50, 0x4b, 0x03, 0x04, 0, 0])), /not a readable text document/)
  assert.throws(() => kind('empty.txt', ''), /This file is empty/)
  assert.throws(() => kind('photo', Buffer.from([0xff, 0xd8, 0xff, 0xe0, 0, 0x10])), /can’t open this kind of file/)
})

test('templates and macro-enabled packages are recognized from their parts', async () => {
  assert.deepEqual(wordPackageInfo(await wordPackage(MAIN)), { template: false, macros: false })
  assert.deepEqual(wordPackageInfo(await wordPackage(TEMPLATE)), { template: true, macros: false })
  assert.deepEqual(wordPackageInfo(await wordPackage(MACRO, { 'word/vbaProject.bin': 'vba' })), { template: false, macros: true })
  assert.deepEqual(wordPackageInfo(await wordPackage(MACRO_TEMPLATE, { 'word/vbaProject.bin': 'vba' })), { template: true, macros: true })
  // The DOCX reader accepts every Word package variant.
  for (const contentType of [MAIN, TEMPLATE, MACRO, MACRO_TEMPLATE]) assert.ok(validateDocxBytes(await wordPackage(contentType)).length > 0, contentType)
})
