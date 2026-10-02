'use strict'

// Opening and adding non-PDF files without an office suite. The unit tests
// run in Node; the last test drives electron/main.cjs and the real hidden
// print windows in Electron (with the office engine forced off) and checks
// the PDFs that come back.
process.env.SIMPLE_FORCE_NO_OFFICE = '1'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')
const test = require('node:test')
const JSZip = require('jszip')
const { PDFDocument, PDFDict, PDFName, PDFRawStream } = require('pdf-lib')
const importer = require('../electron/office-import.cjs')
const fixtures = require('./helpers/import-fixtures.cjs')

const SAMPLES = path.join(__dirname, '..', '..', 'Simple test examples')
const ELECTRON = path.join(__dirname, '..', 'node_modules', 'electron', 'dist', process.platform === 'win32' ? 'electron.exe' : 'electron')
const hasSamples = fs.existsSync(path.join(SAMPLES, 'Complex document.docx'))

function temporaryDirectory(label) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), `simple-pdf-${label}-`))
  process.once('exit', () => { try { fs.rmSync(directory, { recursive: true, force: true, maxRetries: 3 }) } catch { /* best effort */ } })
  return directory
}

/** Samples are only ever read from a private copy. */
function sampleCopy(name, directory) {
  const target = path.join(directory, name)
  fs.copyFileSync(path.join(SAMPLES, name), target)
  return target
}

let pdfjsModule
async function pdfjs() {
  pdfjsModule ||= await import('pdfjs-dist/legacy/build/pdf.mjs')
  return pdfjsModule
}

/** Text (and the text items) of every page. */
async function readPdf(bytes) {
  const { getDocument } = await pdfjs()
  const document = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, disableFontFace: true }).promise
  try {
    const pages = []
    for (let number = 1; number <= document.numPages; number += 1) {
      const page = await document.getPage(number)
      const content = await page.getTextContent()
      const items = content.items.filter((item) => typeof item.str === 'string')
      pages.push({ text: items.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim(), items, view: page.view })
    }
    return pages
  } finally {
    await document.destroy()
  }
}

/** Left and right colours of page 1 rendered small. */
async function sideColors(bytes) {
  const { createCanvas } = require('@napi-rs/canvas')
  const { getDocument } = await pdfjs()
  const document = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, disableFontFace: true }).promise
  try {
    const page = await document.getPage(1)
    const unit = page.getViewport({ scale: 1 })
    const viewport = page.getViewport({ scale: 80 / unit.width })
    const canvas = createCanvas(Math.round(viewport.width), Math.max(1, Math.round(viewport.height)))
    const context = canvas.getContext('2d')
    context.fillStyle = '#fff'
    context.fillRect(0, 0, canvas.width, canvas.height)
    await page.render({ canvasContext: context, viewport }).promise
    const sample = (x) => {
      const [r, g, b] = context.getImageData(Math.floor(canvas.width * x), Math.floor(canvas.height / 2), 1, 1).data
      return r > 150 && g < 100 && b < 100 ? 'red' : b > 150 && r < 100 ? 'blue' : `${r},${g},${b}`
    }
    return [sample(0.2), sample(0.8)]
  } finally {
    await document.destroy()
  }
}

/** Image XObjects reachable from the pages (through form XObjects too). */
async function imageCount(bytes) {
  const document = await PDFDocument.load(bytes)
  const seen = new Set()
  let count = 0
  const visit = (resources) => {
    const xobjects = resources?.lookupMaybe?.(PDFName.of('XObject'), PDFDict)
    if (!xobjects) return
    for (const key of xobjects.keys()) {
      const reference = xobjects.get(key)
      if (seen.has(String(reference))) continue
      seen.add(String(reference))
      const object = xobjects.lookup(key)
      if (!(object instanceof PDFRawStream)) continue
      const subtype = object.dict.get(PDFName.of('Subtype'))
      if (subtype === PDFName.of('Image')) count += 1
      else if (subtype === PDFName.of('Form')) visit(object.dict.lookupMaybe(PDFName.of('Resources'), PDFDict))
    }
  }
  for (const page of document.getPages()) visit(page.node.Resources())
  return count
}

test('file types are recognised from their bytes', () => {
  const zip = (entry) => Buffer.concat([Buffer.from('PK\u0003\u0004', 'latin1'), Buffer.alloc(26), Buffer.from(entry, 'latin1')])
  const cases = [
    [Buffer.concat([Buffer.from('junk before the header\n'), Buffer.from('%PDF-1.7\n')]), 'pdf'],
    [Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex'), 'png'],
    [Buffer.from('ffd8ffe000104a464946', 'hex'), 'jpeg'],
    [fixtures.gif(), 'gif'],
    [Buffer.concat([Buffer.from('RIFF\u0000\u0000\u0000\u0000WEBPVP8 ', 'latin1')]), 'webp'],
    [fixtures.bmp(), 'bmp'],
    [fixtures.tiff([4]), 'tiff'],
    [fixtures.svg(), 'svg'],
    [zip('word/document.xml'), 'docx'],
    [zip('xl/workbook.xml'), 'xlsx'],
    [zip('ppt/presentation.xml'), 'pptx'],
    [Buffer.concat([Buffer.from('PK\u0003\u0004', 'latin1'), Buffer.alloc(26), Buffer.from('mimetypeapplication/vnd.oasis.opendocument.text', 'latin1')]), 'odt'],
    [fixtures.legacyDoc(), 'doc'],
    [fixtures.rtf(), 'rtf'],
    [Buffer.from('\ufeff  <!DOCTYPE html><html><body>x</body></html>'), 'html'],
    [Buffer.from('plain words'), null],
  ]
  for (const [bytes, expected] of cases) assert.equal(importer.sniffType(bytes), expected, `expected ${expected}`)
})

test('text files decode like Notepad: BOM, BOM-less UTF-16, UTF-8, then Cyrillic or Western code pages', () => {
  const decode = (bytes) => importer.decodeText(bytes)
  assert.deepEqual(decode(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('Hé')])), { text: 'Hé', encoding: 'utf-8', bom: true })
  assert.equal(decode(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('UTF16 Notepad text Привет', 'utf16le')])).text, 'UTF16 Notepad text Привет')
  assert.equal(decode(Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from('Big', 'utf16le').swap16()])).text, 'Big')
  assert.deepEqual(decode(Buffer.from('Plain words without a mark', 'utf16le')), { text: 'Plain words without a mark', encoding: 'utf-16le', bom: false })
  assert.equal(decode(Buffer.from('ПриветмирВтораястрока', 'utf16le')).encoding, 'utf-16le')
  assert.equal(decode(Buffer.from('Grüße 你好 ✅')).text, 'Grüße 你好 ✅')
  const western = Buffer.from([0x43, 0x61, 0x66, 0xe9, 0x20, 0x80, 0x20, 0x6e, 0x61, 0xef, 0x76, 0x65])
  assert.deepEqual(decode(western), { text: 'Café € naïve', encoding: 'windows-1252', bom: false })
  const cyrillic = Buffer.from([0xcf, 0xf0, 0xe8, 0xe2, 0xe5, 0xf2, 0x20, 0xec, 0xe8, 0xf0])
  assert.deepEqual(decode(cyrillic), { text: 'Привет мир', encoding: 'windows-1251', bom: false })
  // A UTF-8 file cut inside a character is still UTF-8.
  assert.equal(decode(Buffer.from('Grüße ✅').subarray(0, -1)).encoding, 'utf-8')
})

test('the bytes decide how a file opens; text extensions keep markup as text', () => {
  const kind = (bytes, name) => importer.detectImportKind(bytes, name).kind
  assert.equal(kind(Buffer.from('89504e470d0a1a0a0000000d49484452', 'hex'), 'photo.jpg'), 'image')
  assert.equal(kind(Buffer.from('%PDF-1.4\n'), 'scan'), 'pdf')
  assert.equal(kind(fixtures.rtf(), 'letter.doc'), 'rtf')
  assert.equal(kind(Buffer.from('<!doctype html><p>Hi</p>'), 'saved.doc'), 'html')
  assert.equal(kind(Buffer.from('<!doctype html><p>Hi</p>'), 'notes.txt'), 'text')
  assert.equal(kind(fixtures.svg(), 'drawing.md'), 'markdown')
  assert.equal(kind(fixtures.legacyDoc(), 'report.docx'), 'word-legacy')
  assert.equal(kind(Buffer.from('a,b\n1,2\n'), 'table.csv'), 'spreadsheet')
  assert.equal(kind(Buffer.from([0, 1, 2, 3, 0, 9]), 'mystery.bin'), null)
})

test('pickers list office-engine formats only when an engine is installed', () => {
  const extensions = (filters) => new Set(filters.flatMap((filter) => filter.extensions))
  const without = extensions(importer.importDialogFilters({ engine: false }))
  for (const extension of ['pdf', 'docx', 'odt', 'rtf', 'xlsx', 'xls', 'ods', 'csv', 'md', 'html', 'txt', 'png', 'webp', 'gif', 'bmp', 'tif', 'svg']) assert.ok(without.has(extension), extension)
  for (const extension of ['doc', 'pptx', 'ppt']) assert.ok(!without.has(extension), extension)
  const withEngine = extensions(importer.importDialogFilters({ engine: true }))
  for (const extension of ['doc', 'pptx']) assert.ok(withEngine.has(extension), extension)
  assert.deepEqual(importer.importDialogFilters({ imagesOnly: true }).map((filter) => filter.name), ['Images'])
  assert.ok(importer.isImportableName('C:\\x\\Report.XLSX'))
  assert.ok(!importer.isImportableName('C:\\x\\archive.zip'))
})

test('Markdown becomes headings and emphasis, keeps inline HTML and drops scripts', () => {
  const html = importer.markdownToHtml('---\ntitle: x\n---\n# Title\n\nSome **bold** and _italic_ text<br>next <sup>2</sup>\n\n<script>alert(1)</script>\n\n- item\n\n| a | b |\n|---|---|\n| 1 | 2 |\n', { title: 'Doc' })
  assert.match(html, /<h1>Title<\/h1>/)
  assert.match(html, /<strong>bold<\/strong>/)
  assert.match(html, /<em>italic<\/em>/)
  assert.match(html, /<sup>2<\/sup>/)
  assert.match(html, /<table>/)
  assert.match(html, /<pre class="front-matter">title: x<\/pre>/)
  assert.doesNotMatch(html, /<script|alert\(1\)/)
  assert.match(html, /Content-Security-Policy/)
})

test('web pages are sanitised before they are printed', () => {
  const html = importer.sanitizeHtmlDocument('<html><head><meta http-equiv="refresh" content="0;url=https://x"><base href="https://x/"></head><body onload="x()"><script>bad()</script><iframe src="https://x"></iframe><a href="javascript:bad()">link</a><img src="pic.png" onerror="bad()"></body></html>', { title: 'Page', base: 'file:///C:/docs/' })
  assert.doesNotMatch(html, /<script|bad\(\)|<iframe|http-equiv="refresh"|https:\/\/x\//)
  assert.match(html, /<base href="file:\/\/\/C:\/docs\/">/)
  assert.match(html, /@page \{ size:/)
  assert.match(html, /<img src="pic.png">/)
})

test('conversion windows load only data, blob and files inside their folders', () => {
  const root = path.join(os.tmpdir(), 'simple-job')
  const inside = require('node:url').pathToFileURL(path.join(root, 'picture-1.png')).href
  const outside = require('node:url').pathToFileURL(path.join(os.tmpdir(), 'other', 'secret.txt')).href
  assert.equal(importer.isAllowedRequest('data:image/png;base64,AAAA', [root]), true)
  assert.equal(importer.isAllowedRequest('blob:file:///abc', [root]), true)
  assert.equal(importer.isAllowedRequest(inside, [root]), true)
  assert.equal(importer.isAllowedRequest(outside, [root]), false)
  assert.equal(importer.isAllowedRequest('https://example.com/x.png', [root]), false)
  assert.equal(importer.isAllowedRequest('file://server/share/x.png', [root]), false)
  assert.equal(importer.isAllowedRequest(inside, []), false)
})

test('RTF keeps text, emphasis, code pages, Unicode, tables and page breaks', () => {
  const { html } = importer.rtfToHtml(fixtures.rtf(), { title: 'Notes' })
  assert.match(html, /text-align: center[^>]*><span style="font-weight: 700; font-size: 18pt">RTF_TITLE/)
  assert.match(html, /<span style="font-weight: 700">RTF_BOLD<\/span>/)
  assert.match(html, /café/)
  assert.match(html, /Привет/)
  assert.match(html, /€ euro/)
  assert.match(html, /color: #c80000">Red words/)
  assert.match(html, /<table class="rtf-table">.*RTF_CELL_A.*RTF_CELL_B/s)
  assert.match(html, /class="page-break"/)
})

test('OpenDocument text keeps headings, emphasis, lists and tables', async () => {
  const { html } = await importer.odtToHtml(await fixtures.odt(), { title: 'Notes' })
  assert.match(html, /<h1[^>]*>ODT_HEADING<\/h1>/)
  assert.match(html, /font-weight: 700">ODT_BOLD/)
  assert.match(html, /paragraph &amp; more/)
  assert.match(html, /<ul><li>.*ODT_ITEM_ONE/s)
  assert.match(html, /<td>.*ODT_CELL_A.*<\/td><td>.*ODT_CELL_B/s)
})

test('Word documents keep sections, colours, pictures and tables', { skip: !hasSamples && 'sample files are not available' }, async () => {
  const directory = temporaryDirectory('docx')
  const result = await importer.docxToHtml(fs.readFileSync(sampleCopy('Complex document.docx', directory)), { title: 'Complex document' })
  assert.deepEqual(result.layout.sections.map((section) => [section.width, section.height]), [[612, 792], [792, 612]])
  assert.match(result.html, /@page section\d+ \{ size: 792pt 612pt;/)
  assert.equal((result.html.match(/<div class="docx-section"/g) || []).length, 2)
  assert.match(result.html, /color: #003388">Mixed Arial blue/)
  assert.match(result.html, /<p>(<span style="color: #[0-9a-f]{6}">)?Table one follows\.(<\/span>)?<\/p>/)
  assert.equal((result.html.match(/<table>/g) || []).length, 2)
  assert.deepEqual([...result.files.keys()], ['picture-1.png'])
  assert.match(result.html, /<img src="picture-1\.png"[^>]*style="width: 240pt; height: 105pt"/)
  assert.doesNotMatch(result.html, /[\uE000-\uE00F]/)
})

test('spreadsheets become one table per visible sheet', { skip: !hasSamples && 'sample files are not available' }, async () => {
  const directory = temporaryDirectory('sheets')
  for (const name of ['Complex workbook.xlsx', 'Complex workbook.xls', 'Complex workbook.ods']) {
    const extension = path.extname(name).slice(1)
    const result = importer.spreadsheetToHtml(fs.readFileSync(sampleCopy(name, directory)), { title: 'Workbook', extension })
    assert.match(result.html, /16,490\.00/, name)
    assert.ok(result.sheets >= 3, name)
  }
  const csv = importer.spreadsheetToHtml(Buffer.from('Name,Amount\nTea,"1,234.50"\n'), { extension: 'csv' })
  assert.match(csv.html, /<td>Tea<\/td><td[^>]*>1,234\.50<\/td>/)
})

test('formats that need an office engine fail with a coded, plain message', async () => {
  for (const [name, bytes, wording] of [['legacy.doc', fixtures.legacyDoc(), /Simple Documents or Word, save it as a \.docx/], ['slides.pptx', await fixtures.pptx(), /save the presentation as a PDF/]]) {
    await assert.rejects(importer.convertToPdf(bytes, { name }), (error) => {
      assert.equal(error.code, 'NEEDS_OFFICE_ENGINE')
      assert.equal(error.name, 'NEEDS_OFFICE_ENGINE')
      assert.match(error.message, new RegExp(`^${name.replace('.', '\\.')}: `))
      assert.match(error.message, wording)
      return true
    })
  }
  await assert.rejects(importer.convertToPdf(Buffer.from([0, 1, 2, 3, 0, 5]), { name: 'bad.xyz' }), { code: 'UNSUPPORTED_FORMAT', message: /^bad\.xyz: Simple can't open \.xyz files/ })
  await assert.rejects(importer.convertToPdf(Buffer.alloc(0), { name: 'empty.txt' }), { code: 'EMPTY_FILE' })
})

test('PDF bytes pass through; PNG, JPEG and TIFF convert without a browser', async () => {
  const pdf = Buffer.from('%PDF-1.7\n%%EOF\n')
  const passed = await importer.convertToPdf(pdf, { name: 'scan' })
  assert.equal(passed.converted, false)
  assert.equal(passed.data, pdf)
  const { createCanvas } = require('@napi-rs/canvas')
  const canvas = createCanvas(40, 20)
  const context = canvas.getContext('2d')
  context.fillStyle = 'rgb(220,20,30)'; context.fillRect(0, 0, 20, 20)
  context.fillStyle = 'rgb(20,40,210)'; context.fillRect(20, 0, 20, 20)
  const misnamed = await importer.convertToPdf(canvas.toBuffer('image/png'), { name: 'png_named_as.jpg' })
  assert.deepEqual(await sideColors(misnamed.data), ['red', 'blue'])
  const tiff = await importer.convertToPdf(fixtures.tiff([40, 30, 20]), { name: 'pages.tif' })
  const document = await PDFDocument.load(tiff.data)
  assert.deepEqual(document.getPages().map((page) => Math.round(page.getWidth())), [40, 30, 20])
  assert.deepEqual(await sideColors(tiff.data), ['red', 'blue'])
})

test('in Electron, every format opens and adds without an office suite', { skip: !fs.existsSync(ELECTRON) && 'Electron is not installed', timeout: 300_000 }, async () => {
  const directory = temporaryDirectory('electron')
  const input = path.join(directory, 'in')
  const output = path.join(directory, 'out')
  fs.mkdirSync(input)
  fs.mkdirSync(output)
  const { createCanvas } = require('@napi-rs/canvas')
  const canvas = createCanvas(40, 20)
  const context = canvas.getContext('2d')
  context.fillStyle = 'rgb(220,20,30)'; context.fillRect(0, 0, 20, 20)
  context.fillStyle = 'rgb(20,40,210)'; context.fillRect(20, 0, 20, 20)
  const files = await fixtures.writeImportFixtures(input, { webp: canvas.toBuffer('image/webp'), png: canvas.toBuffer('image/png'), large: true })
  const samples = hasSamples
    ? Object.fromEntries(['Complex document.docx', 'Complex document.md', 'Complex document.html', 'Complex workbook.xlsx', 'Complex workbook.xls', 'Complex workbook.ods', 'Complex workbook.pdf']
      .map((name) => [name, sampleCopy(name, input)]))
    : {}
  const convertNames = ['utf16-bom.txt', 'utf16-nobom.txt', 'cp1252.txt', 'cp1251.txt', 'multi.txt', 'picture.webp', 'picture.gif', 'picture.bmp', 'picture.svg', 'notes.rtf', 'notes.odt', 'large.txt']
  const job = {
    outDir: output,
    openAnswers: [[files['picture.webp']]],
    convert: [
      ...convertNames.map((name) => ({ id: name, file: files[name] })),
      ...['Complex document.md', 'Complex document.html', 'Complex workbook.xlsx', 'Complex workbook.xls', 'Complex workbook.ods'].filter((name) => samples[name]).map((name) => ({ id: name, file: samples[name] })),
    ],
    ipc: [
      { id: 'open-doc', channel: 'file:open-path', args: [files['legacy.doc']] },
      { id: 'open-pptx', channel: 'file:open-path', args: [files['slides.pptx']] },
      { id: 'pick-image', channel: 'file:pick-image', args: [] },
      { id: 'open-dialog', channel: 'file:open-dialog', args: [] },
      ...(hasSamples ? [
        { id: 'open-docx', channel: 'file:open-path', args: [samples['Complex document.docx']] },
        { id: 'open-bytes-md', channel: 'file:open-bytes', args: [{ name: 'Complex document.md', data: { $file: samples['Complex document.md'] } }] },
        { id: 'insert', channel: 'pdf:insert-dropped-files', args: [{ $file: samples['Complex workbook.pdf'] }, 1, { $inputs: [
          { name: 'Complex workbook.pdf', file: samples['Complex workbook.pdf'] },
          { name: 'Complex document.docx', file: samples['Complex document.docx'] },
          { name: 'bad.xyz', file: files['bad.xyz'] },
        ] }] },
      ] : []),
    ],
  }
  const jobFile = path.join(directory, 'job.json')
  fs.writeFileSync(jobFile, JSON.stringify(job))
  const env = { ...process.env, SIMPLE_IMPORT_JOB: jobFile, SIMPLE_FORCE_NO_OFFICE: '1' }
  delete env.ELECTRON_RUN_AS_NODE
  const child = spawn(ELECTRON, [`--user-data-dir=${path.join(directory, 'profile')}`, path.join(__dirname, 'helpers', 'office-import-runner.cjs')], { env, windowsHide: true, stdio: ['ignore', 'ignore', 'pipe'] })
  let stderr = ''
  child.stderr.on('data', (chunk) => { stderr = (stderr + chunk).slice(-4000) })
  const exited = await new Promise((resolve) => {
    const timer = setTimeout(() => {
      if (process.platform === 'win32') spawnSync('taskkill.exe', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
      else child.kill('SIGKILL')
      resolve(false)
    }, 240_000)
    child.once('exit', () => { clearTimeout(timer); resolve(true) })
  })
  assert.ok(exited, `Electron did not finish in time. ${stderr}`)
  const results = JSON.parse(fs.readFileSync(path.join(output, 'results.json'), 'utf8'))
  assert.equal(results.fatal, undefined, JSON.stringify(results.fatal))
  const converted = async (id) => {
    const entry = results.convert[id]
    assert.ok(entry?.ok, `${id}: ${JSON.stringify(entry?.error)}`)
    return fs.readFileSync(entry.file)
  }

  // Text in every encoding comes out exactly.
  assert.equal((await readPdf(await converted('utf16-bom.txt')))[0].text, 'UTF16 Notepad text Привет Second line')
  assert.equal((await readPdf(await converted('utf16-nobom.txt')))[0].text, 'UTF16 without mark: plain words here and Привет')
  assert.equal((await readPdf(await converted('cp1252.txt')))[0].text, 'Café naïve résumé € price Second — line')
  assert.equal((await readPdf(await converted('cp1251.txt')))[0].text, 'Привет мир Вторая строка')
  const multi = await readPdf(await converted('multi.txt'))
  assert.equal(multi.length, 2, 'a form feed starts a new page')
  assert.match(multi[0].text, /你好 ✅/)
  assert.doesNotMatch(multi[0].text, /\u0000/)

  // Every picture format arrives the right way round; SVG stays vector text.
  for (const name of ['picture.webp', 'picture.gif', 'picture.bmp', 'picture.svg']) {
    const bytes = await converted(name)
    assert.equal((await PDFDocument.load(bytes)).getPageCount(), 1, name)
    assert.deepEqual(await sideColors(bytes), ['red', 'blue'], name)
  }
  assert.match((await readPdf(await converted('picture.svg')))[0].text, /SVG_VECTOR_TEXT/)

  const rtf = await readPdf(await converted('notes.rtf'))
  assert.equal(rtf.length, 2)
  for (const expected of ['RTF_TITLE', 'RTF_BOLD', 'café', 'Привет', '€ euro', 'RTF_CELL_A', 'RTF_CELL_B']) assert.ok(rtf[0].text.includes(expected), expected)
  assert.match(rtf[1].text, /After the page break/)
  const odt = (await readPdf(await converted('notes.odt')))[0].text
  for (const expected of ['ODT_HEADING', 'ODT_BOLD', 'ODT_ITEM_ONE', 'ODT_CELL_B']) assert.ok(odt.includes(expected), expected)

  // A 5 MB text file converts quickly and never freezes the main process.
  const large = results.convert['large.txt']
  assert.ok(large.ok && large.ms < 20_000, `large text took ${large.ms} ms`)
  assert.ok(large.stallMs < 200, `the main process stalled for ${large.stallMs} ms`)
  assert.ok((await PDFDocument.load(fs.readFileSync(large.file))).getPageCount() > 50)

  // Office-only formats explain themselves; pickers do not offer them.
  for (const id of ['open-doc', 'open-pptx']) {
    assert.equal(results.ipc[id].ok, false)
    assert.match(results.ipc[id].rendererText, /NEEDS_OFFICE_ENGINE: /)
  }
  assert.match(results.ipc['pick-image'].value.dataUrl, /^data:image\/(png|jpeg);base64,/)
  const documentFilters = results.openDialogs.find((options) => options.title === 'Open a document').filters
  const offered = new Set(documentFilters.flatMap((filter) => filter.extensions))
  for (const extension of ['docx', 'xlsx', 'md', 'html', 'txt', 'rtf', 'odt', 'webp', 'tif', 'svg']) assert.ok(offered.has(extension), extension)
  for (const extension of ['doc', 'pptx']) assert.ok(!offered.has(extension), extension)

  if (!hasSamples) return
  const markdown = await readPdf(await converted('Complex document.md'))
  const markdownText = markdown.map((page) => page.text).join(' ')
  assert.doesNotMatch(markdownText, /(^|\s)# |\*\*/)
  const sizes = markdown[0].items.filter((item) => item.str.trim()).map((item) => ({ text: item.str, height: item.height }))
  const heading = sizes.find((item) => item.text === 'created')
  const body = sizes.find((item) => item.text.startsWith('Mixed Arial blue'))
  assert.ok(heading && body && heading.height > body.height, 'the H1 is larger than body text')
  assert.match((await readPdf(await converted('Complex document.html'))).map((page) => page.text).join(' '), /STRESS_START/)
  for (const id of ['Complex workbook.xlsx', 'Complex workbook.xls', 'Complex workbook.ods']) {
    assert.match((await readPdf(await converted(id))).map((page) => page.text).join(' '), /16,490\.00/, id)
  }

  const opened = results.ipc['open-docx']
  assert.ok(opened.ok, opened.rendererText)
  assert.equal(opened.value.converted, true)
  assert.equal(opened.value.path, null)
  assert.ok(opened.ms < 10_000, `opening the Word document took ${opened.ms} ms`)
  const word = fs.readFileSync(opened.value.data.file)
  const wordPages = await readPdf(word)
  assert.ok(wordPages.length >= 3)
  const wordText = wordPages.map((page) => page.text).join(' ')
  for (const expected of ['Table one follows.', 'TABLE_ONE_1 value', 'TABLE_ONE_9 value', 'TABLE_TWO_4', 'FOOTNOTE_CONTENT']) assert.ok(wordText.includes(expected), expected)
  assert.ok(await imageCount(word) >= 1, 'the picture is kept')
  assert.deepEqual(wordPages.at(-1).view.slice(2).map(Math.round), [792, 612], 'the landscape section stays landscape')

  const bytes = results.ipc['open-bytes-md']
  assert.ok(bytes.ok, bytes.rendererText)
  assert.equal(bytes.value.name, 'Complex document.pdf')
  assert.equal(bytes.value.converted, true)

  const inserted = results.ipc.insert
  assert.ok(inserted.ok, inserted.rendererText)
  const workbookPages = (await PDFDocument.load(fs.readFileSync(samples['Complex workbook.pdf']))).getPageCount()
  assert.equal(inserted.value.added, workbookPages + wordPages.length)
  assert.deepEqual(inserted.value.skipped.map((item) => [item.name, item.code]), [['bad.xyz', 'UNSUPPORTED_FORMAT']])
  assert.equal((await PDFDocument.load(fs.readFileSync(inserted.value.data.file))).getPageCount(), workbookPages * 2 + wordPages.length)
})
