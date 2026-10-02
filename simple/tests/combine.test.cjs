const test = require('node:test')
const { after } = require('node:test')
const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const { spawn, execFile } = require('node:child_process')
const { pathToFileURL } = require('node:url')
const { PDFDict, PDFDocument, PDFName, PDFPageLeaf, PDFString } = require('../../simple_pdf_source/node_modules/pdf-lib')
const JSZip = require('../../simple_pdf_source/node_modules/jszip')
const { combineFiles, pageIndices, LEGACY_DOC_POLICY } = require('../launcher/combine-service.cjs')
const { atomicWrite, mupdfPath, runCombine, unlockPdfBytes } = require('../launcher/combine-host.cjs')
const { legacyFormatError, COMBINE_EXTENSIONS } = require('../launcher/combine-policy.cjs')
const native = require('../launcher/combine-native.cjs')
const htmlToPdf = require('../shared/electron/html-to-pdf.cjs')

const ROOT = path.resolve(__dirname, '..')
const SAMPLES = path.resolve(ROOT, '..', 'Simple test examples')
// The combined PDF is saved through the shared safe-write path; keep its save journal in a scratch folder.
const JOURNAL = require('node:fs').mkdtempSync(path.join(os.tmpdir(), 'simple-combine-journal-'))
process.env.SIMPLE_IO_JOURNAL_DIR = JOURNAL
const SETUP_WORDING = /\b(install|installs|installed|installing|installer|download|downloads|downloading|libreoffice)\b/i
const hash = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex')

let bundledWorker = null
/**
 * Bundles the Combine worker from the current sources with the same esbuild
 * options as scripts/sync-build.cjs, so these tests run what ships.
 */
function workerBundle() {
  bundledWorker ??= (async () => {
    const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-combine-bundle-'))
    const outfile = path.join(directory, 'combine-worker.cjs')
    await require('esbuild').build({
      entryPoints: [path.join(ROOT, 'launcher', 'combine-worker.cjs')],
      outfile, bundle: true, platform: 'node', format: 'cjs', target: 'node22', legalComments: 'none', logLevel: 'silent',
    })
    return { directory, outfile }
  })()
  return bundledWorker
}

after(async () => {
  assert.ok(path.basename(JOURNAL).startsWith('simple-combine-journal-'))
  await fs.rm(JOURNAL, { recursive: true, force: true })
  if (!bundledWorker) return
  const { directory } = await bundledWorker
  assert.ok(path.basename(directory).startsWith('simple-combine-bundle-'))
  await fs.rm(directory, { recursive: true, force: true })
})

async function scratch(prefix) {
  return fs.mkdtemp(path.join(os.tmpdir(), `simple-combine-${prefix}-`))
}

async function removeScratch(directory) {
  const absolute = path.resolve(directory)
  assert.ok(absolute.startsWith(path.resolve(os.tmpdir()) + path.sep) && path.basename(absolute).startsWith('simple-combine-'))
  await fs.rm(absolute, { recursive: true, force: true })
}

async function copySamples(directory, names) {
  const copies = {}
  for (const name of names) {
    copies[name] = path.join(directory, name)
    await fs.copyFile(path.join(SAMPLES, name), copies[name])
  }
  return copies
}

async function withForcedOff(action) {
  const previous = process.env.SIMPLE_FORCE_NO_OFFICE
  process.env.SIMPLE_FORCE_NO_OFFICE = '1'
  try { return await action() } finally {
    if (previous === undefined) delete process.env.SIMPLE_FORCE_NO_OFFICE
    else process.env.SIMPLE_FORCE_NO_OFFICE = previous
  }
}

/** Stands in for html-to-pdf: one page per job, sized like the job asks. */
function recordingPrinter(jobs) {
  return async (html, options) => {
    jobs.push({ html, options })
    const pdf = await PDFDocument.create()
    const size = options.pageSize || { width: 8.5, height: 11 }
    const [width, height] = options.landscape ? [size.height, size.width] : [size.width, size.height]
    pdf.addPage([width * 72, height * 72])
    return pdf.save()
  }
}

async function legacyFixture(filePath) {
  const bytes = Buffer.alloc(1024)
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(bytes)
  await fs.writeFile(filePath, bytes)
}

test('page selections preserve user order, support ranges, and reject missing pages', () => {
  assert.deepEqual(pageIndices('3, 1-2', 3), [2, 0, 1])
  assert.deepEqual(pageIndices('', 2), [0, 1])
  for (const value of ['0', '3-1', '4', '1,', 'cat', '1-10000000']) assert.throws(() => pageIndices(value, 3))
})

test('combine worker preserves file order, page selections and page sizes without touching inputs', async () => {
  const { outfile } = await workerBundle()
  const dir = await scratch('test')
  try {
    const first = await PDFDocument.create()
    first.addPage([200, 300])
    first.addPage([400, 500])
    const second = await PDFDocument.create()
    second.addPage([600, 200])
    const a = path.join(dir, 'one.pdf')
    const b = path.join(dir, 'two.pdf')
    const original = Buffer.from(await first.save())
    await fs.writeFile(a, original)
    await fs.writeFile(b, await second.save())
    const progress = []
    const result = await runCombine([{ path: b }, { path: a, pages: '2,1' }], (value) => progress.push(value.name), { workerPath: outfile })
    const combined = await PDFDocument.load(result.bytes)
    assert.deepEqual(combined.getPages().map((page) => [page.getWidth(), page.getHeight()]), [[600, 200], [400, 500], [200, 300]])
    assert.equal(result.pageCount, 3)
    assert.deepEqual(result.builtIn, [])
    assert.deepEqual(progress, ['two.pdf', 'one.pdf'])
    assert.deepEqual(await fs.readFile(a), original)
    await assert.rejects(combineFiles([{ path: a, pages: '99' }, { path: b }]), /one.pdf: Choose pages/)
    await assert.rejects(runCombine([{ path: a, pages: '99' }, { path: b }], () => {}, { workerPath: outfile }), (error) => error.code === 'INVALID_PAGES' && /one.pdf: Choose pages/.test(error.message))
    const destination = path.join(dir, 'output.pdf')
    await fs.writeFile(destination, 'existing')
    await atomicWrite(destination, result.bytes)
    assert.equal((await PDFDocument.load(await fs.readFile(destination))).getPageCount(), 3)
    assert.equal((await fs.readdir(dir)).some((name) => name.endsWith('.tmp')), false)
  } finally { await removeScratch(dir) }
})

test('the combined PDF is saved through safe-write: never over a read-only file, never as damaged bytes', async () => {
  const dir = await scratch('readonly')
  const target = path.join(dir, 'Combined.pdf')
  try {
    await fs.writeFile(target, 'keep me')
    await fs.chmod(target, 0o444)
    const pdf = await PDFDocument.create()
    pdf.addPage([100, 100])
    await assert.rejects(atomicWrite(target, await pdf.save()), (error) => error.code === 'READ_ONLY' && /read-only\. Save the combined PDF with a different name/.test(error.message))
    assert.equal(await fs.readFile(target, 'utf8'), 'keep me')
    assert.deepEqual((await fs.readdir(dir)).filter((name) => name !== 'Combined.pdf'), [], 'no temp file is left behind')
    await fs.chmod(target, 0o666)
    await assert.rejects(atomicWrite(target, Buffer.from('not a PDF')), (error) => error.code === 'VALIDATION_FAILED', 'damaged output never replaces a file')
    assert.equal(await fs.readFile(target, 'utf8'), 'keep me')
    assert.deepEqual(await atomicWrite(target, await pdf.save()).then((result) => result.strategy), 'rename')
  } finally {
    await fs.chmod(target, 0o666).catch(() => {})
    await removeScratch(dir)
  }
})

test('Word combination uses the layout engine and preserves its page geometry', async () => {
  const dir = await scratch('doc')
  try {
    const source = path.join(dir, 'legacy.doc')
    await fs.writeFile(source, 'test source')
    const pdf = await PDFDocument.create()
    pdf.addPage([595.28, 841.89])
    const output = await pdf.save()
    let calls = 0
    const result = await combineFiles([{ path: source }, { path: source }], () => {}, {
      convertOfficeBytes: async (input, options) => {
        calls++
        assert.equal(input.inputExtension, 'doc')
        assert.equal(input.outputExtension, 'pdf')
        assert.equal(input.filter, 'writer_pdf_Export')
        assert.equal(options.policyId, LEGACY_DOC_POLICY, 'legacy date fields are still locked in the private copy')
        assert.equal(typeof options.prepareInput, 'function')
        assert.deepEqual(options.prepareInput(Buffer.from('not a cfb file')).bytes, Buffer.from('not a cfb file'))
        return output
      },
    })
    assert.equal(calls, 2)
    assert.equal(result.pageCount, 2)
    assert.deepEqual(result.builtIn, [])
    const combined = await PDFDocument.load(result.bytes)
    assert.equal(combined.getPage(0).getWidth(), 595.28)
  } finally { await removeScratch(dir) }
})

test('XLS, XLSX and ODS use spreadsheet print layout, retain page ranges and preserve source bytes', async () => {
  const dir = await scratch('sheet')
  try {
    const pdf = await PDFDocument.create()
    pdf.addPage([792, 612])
    pdf.addPage([595.28, 841.89])
    const output = await pdf.save()
    const calls = []
    const sources = await Promise.all(['XLS', 'xlsx', 'ods'].map(async (extension) => {
      const source = path.join(dir, `schedule.${extension}`)
      await fs.writeFile(source, `original ${extension} print settings`)
      return source
    }))
    const originals = await Promise.all(sources.map((source) => fs.readFile(source)))
    const options = { convertOfficeBytes: async (input, conversion) => {
      calls.push(input.inputExtension)
      assert.equal(input.outputExtension, 'pdf')
      assert.equal(input.filter, 'calc_pdf_Export')
      assert.equal(conversion.prepareInput, undefined)
      assert.ok(input.bytes.toString().includes('print settings'))
      return output
    } }
    const result = await combineFiles(sources.map((source) => ({ path: source, pages: '2,1' })), () => {}, options)
    assert.deepEqual(calls, ['xls', 'xlsx', 'ods'])
    assert.equal(result.pageCount, 6)
    const combined = await PDFDocument.load(result.bytes)
    assert.deepEqual(combined.getPages().map((page) => [page.getWidth(), page.getHeight()]), Array.from({ length: 3 }, () => [[595.28, 841.89], [792, 612]]).flat())
    await assert.rejects(combineFiles([{ path: sources[0], pages: '3' }, { path: sources[1] }], () => {}, options), /schedule.XLS: Choose pages between 1 and 2/)
    assert.deepEqual(await Promise.all(sources.map((source) => fs.readFile(source))), originals)
  } finally { await removeScratch(dir) }
})

test('interactive and encryption-marked PDFs are rejected before pages are copied', async () => {
  const dir = await scratch('protected')
  try {
    const formPdf = await PDFDocument.create()
    const page = formPdf.addPage([300, 300])
    const field = formPdf.getForm().createTextField('Name')
    field.setText('Keep this value'); field.addToPage(page, { x: 10, y: 40, width: 200, height: 30 })
    const source = path.join(dir, 'form.pdf')
    const original = await formPdf.save()
    await fs.writeFile(source, original)
    await assert.rejects(combineFiles([{ path: source }, { path: source }]), /form.pdf:.*interactive form fields.*flatten/)
    assert.deepEqual(await fs.readFile(source), Buffer.from(original))
    const protectedPdf = await PDFDocument.create(); protectedPdf.addPage([300, 300])
    protectedPdf.context.trailerInfo.Encrypt = protectedPdf.context.register(protectedPdf.context.obj({ Filter: 'Standard', V: 1, R: 2, P: -4 }))
    await fs.writeFile(source, await protectedPdf.save())
    await assert.rejects(combineFiles([{ path: source }, { path: source }]), (error) => error.code === 'ENCRYPTED' && /^form.pdf: This PDF is protected/.test(error.message))
    // An encryption dictionary MuPDF can't read is refused the same way, never copied as is.
    await assert.rejects(combineFiles([{ path: source }, { path: source }], () => {}, { unlockPdf: unlockPdfBytes }), (error) => error.code === 'ENCRYPTED' && /^form.pdf: /.test(error.message))
  } finally { await removeScratch(dir) }
})

/** Encrypts PDF bytes with MuPDF (the library the PDF workspace and Combine unlock with). */
async function mupdfEncrypt(bytes, options) {
  const mupdf = await import(pathToFileURL(mupdfPath()).href)
  const document = mupdf.Document.openDocument(bytes, 'application/pdf')
  try { return Buffer.from(document.saveToBuffer(options).asUint8Array()) } finally { document.destroy?.() }
}

test('PDFs with only a permissions password combine like the PDF workspace opens them; a password-protected PDF is refused', async () => {
  const { outfile } = await workerBundle()
  const dir = await scratch('owner')
  try {
    const plain = await PDFDocument.create()
    plain.addPage([300, 400]).drawText('Monthly statement', { x: 20, y: 300 })
    plain.addPage([300, 400])
    const raw = await plain.save()
    const statement = path.join(dir, 'statement.pdf')
    const locked = path.join(dir, 'locked.pdf')
    const statementBytes = await mupdfEncrypt(raw, 'encrypt=aes-128,owner-password=owner,user-password=,permissions=-3904')
    await fs.writeFile(statement, statementBytes)
    await fs.writeFile(locked, await mupdfEncrypt(raw, 'encrypt=aes-256,owner-password=owner,user-password=secret'))
    assert.equal((await PDFDocument.load(statementBytes, { ignoreEncryption: true })).isEncrypted, true)

    const direct = await combineFiles([{ path: statement, pages: '2' }, { path: statement }], () => {}, { unlockPdf: unlockPdfBytes })
    assert.equal(direct.pageCount, 3)
    // Through the bundled worker: it asks this (main) thread to unlock.
    const viaWorker = await runCombine([{ path: statement }, { path: statement, pages: '1' }], () => {}, { workerPath: outfile })
    assert.equal(viaWorker.pageCount, 3)
    assert.equal((await PDFDocument.load(viaWorker.bytes)).isEncrypted, false)
    assert.deepEqual(await fs.readFile(statement), statementBytes, 'the protected original is only read')

    await assert.rejects(runCombine([{ path: statement }, { path: locked }], () => {}, { workerPath: outfile }),
      (error) => error.code === 'ENCRYPTED' && /^locked\.pdf: This PDF needs a password to open\. Open it in Simple with its password/.test(error.message))
  } finally { await removeScratch(dir) }
})

test('internal links point at the copied pages, links to pages left out are removed, and no unselected page comes along', async () => {
  const dir = await scratch('links')
  try {
    const doc = await PDFDocument.create()
    const pages = [1, 2, 3, 4].map((number) => {
      const page = doc.addPage([300, 300])
      page.drawText(`Page ${number}`, { x: 20, y: 250 })
      return page
    })
    const context = doc.context
    const link = (extra) => context.register(context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [0, 0, 50, 50], P: pages[0].ref, ...extra }))
    pages[0].node.set(PDFName.of('Annots'), context.obj([
      link({ Dest: [pages[2].ref, 'XYZ', null, null, null] }),            // page 3: selected
      link({ A: { S: 'GoTo', D: [pages[1].ref, 'Fit'] } }),                // page 2: left out
      link({ Dest: PDFString.of('chapter4') }),                           // named, page 4: left out
      link({ A: { S: 'GoTo', D: PDFString.of('chapter3') } }),            // named, page 3: selected
      link({ A: { S: 'URI', URI: PDFString.of('mailto:someone@example.invalid') } }),
    ]))
    // Page 3 links back to page 1 through its own annotation.
    pages[2].node.set(PDFName.of('Annots'), context.obj([link({ Dest: [pages[0].ref, 'Fit'], P: pages[2].ref })]))
    doc.catalog.set(PDFName.of('Names'), context.obj({ Dests: context.obj({ Names: [
      PDFString.of('chapter3'), context.obj([pages[2].ref, 'Fit']),
      PDFString.of('chapter4'), context.obj([pages[3].ref, 'Fit']),
    ] }) }))
    const source = path.join(dir, 'manual.pdf')
    const original = Buffer.from(await doc.save())
    await fs.writeFile(source, original)
    const other = await PDFDocument.create(); other.addPage([200, 200])
    const second = path.join(dir, 'cover.pdf')
    await fs.writeFile(second, await other.save())

    const result = await combineFiles([{ path: second }, { path: source, pages: '3,1' }], () => {}, { title: 'Manual extract' })
    const combined = await PDFDocument.load(result.bytes, { updateMetadata: false })
    assert.equal(combined.getPageCount(), 3)
    const pageObjects = combined.context.enumerateIndirectObjects()
      .filter(([, object]) => object instanceof PDFPageLeaf || (object instanceof PDFDict && object.get(PDFName.of('Type')) === PDFName.of('Page')))
    assert.equal(pageObjects.length, 3, 'no hidden copies of linked pages')
    const [, pageThree, pageOne] = combined.getPages()
    const annotations = (page) => {
      const annots = page.node.Annots()
      return annots ? Array.from({ length: annots.size() }, (_, index) => combined.context.lookup(annots.get(index))) : []
    }
    const firstPageLinks = annotations(pageOne)
    assert.equal(firstPageLinks.length, 3, 'links to pages that were left out are removed')
    assert.equal(String(firstPageLinks[0].get(PDFName.of('Dest'))), `[ ${pageThree.ref} /XYZ null null null ]`)
    assert.equal(String(firstPageLinks[1].get(PDFName.of('Dest'))), `[ ${pageThree.ref} /Fit ]`)
    assert.equal(combined.context.lookup(firstPageLinks[2].get(PDFName.of('A'))).get(PDFName.of('S')), PDFName.of('URI'), 'other links are kept')
    assert.equal(String(annotations(pageThree)[0].get(PDFName.of('Dest'))), `[ ${pageOne.ref} /Fit ]`)
    assert.ok(firstPageLinks.every((annotation) => !annotation.has(PDFName.of('P'))))
    assert.equal(combined.getTitle(), 'Manual extract')
    assert.equal(combined.getProducer(), 'simple')
    assert.equal(combined.getCreator(), 'simple')
    assert.deepEqual(await fs.readFile(source), original, 'the source file is only read')
  } finally { await removeScratch(dir) }
})

test('SIMPLE_FORCE_NO_OFFICE=1: Word and spreadsheet copies are laid out by Simple, hidden sheets stay out, sources stay untouched', async () => {
  const dir = await scratch('no-office')
  try {
    const copies = await copySamples(dir, ['Complex document.pdf', 'Complex document.docx', 'Complex workbook.xlsx', 'Complex workbook.ods'])
    const csv = path.join(dir, 'people.csv')
    await fs.writeFile(csv, Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('sep=;\nName;Amount;Note\nAlice;1 234,50;"multi\nline; quoted"\nБорис;42;Привет\n', 'utf8')]))
    const sources = [...Object.values(copies), csv]
    const before = await Promise.all(sources.map(async (file) => hash(await fs.readFile(file))))
    const jobs = []
    const progress = []
    const result = await withForcedOff(() => combineFiles(sources.map((file) => ({ path: file })), (value) => progress.push(value.name), { printHtml: recordingPrinter(jobs) }))
    const pdfPages = (await PDFDocument.load(await fs.readFile(copies['Complex document.pdf']))).getPageCount()
    assert.equal(result.pageCount, pdfPages + jobs.length, 'every PDF page plus one stub page per print job')
    assert.ok(result.pageCount >= pdfPages + 4)
    assert.deepEqual(result.builtIn, ['Complex document.docx', 'Complex workbook.xlsx', 'Complex workbook.ods', 'people.csv'])
    assert.deepEqual(progress, sources.map((file) => path.basename(file)))

    const [word, ...sheets] = jobs
    assert.match(word.html, /STRESS_START/)
    assert.match(word.html, /LONG_CONTENT_START/)
    assert.match(word.html, /@page \{ size: 8\.5in 11in;/, 'the first section keeps its Letter page size')
    assert.equal(word.options.preferCSSPageSize, true)
    assert.doesNotMatch(word.html, /<script|javascript:/i)
    assert.ok(sheets.every((job) => !job.html.includes('HIDDEN_AUDIT_SENTINEL')), 'a hidden sheet is never printed')
    const workbook = sheets.filter((job) => /Complex workbook - /.test(job.options.title))
    assert.deepEqual(workbook.map((job) => job.options.title), [
      'Complex workbook - Summary', 'Complex workbook - Transactions', 'Complex workbook - Inputs & notes',
      'Complex workbook - Summary', 'Complex workbook - Transactions', 'Complex workbook - Inputs & notes',
    ])
    const transactions = workbook[1]
    assert.equal(transactions.options.landscape, true)
    assert.deepEqual(transactions.options.pageSize, { width: 8.27, height: 11.69 })
    assert.match(transactions.html, /<thead><tr>.*Invoice.*<\/tr><\/thead>/, 'print titles repeat on every page')
    assert.match(transactions.html, /INV-00060/)
    assert.doesNotMatch(transactions.html, /INV-00061/, 'the print area ends at row 63')
    assert.match(transactions.options.footerTemplate, /Confidential sample[\s\S]*class="pageNumber"[\s\S]*class="totalPages"/)
    assert.match(transactions.options.headerTemplate, /Transactions/)
    const people = sheets.at(-1)
    assert.match(people.html, /Борис/)
    assert.match(people.html, /multi\nline; quoted/)
    assert.deepEqual(await Promise.all(sources.map(async (file) => hash(await fs.readFile(file)))), before)
  } finally { await removeScratch(dir) }
})

test('SIMPLE_FORCE_NO_OFFICE=1: .doc and .xls get a coded error that names the native alternative', async () => {
  const dir = await scratch('legacy')
  try {
    const { 'Complex workbook.xls': xls, 'Complex document.pdf': pdf } = await copySamples(dir, ['Complex workbook.xls', 'Complex document.pdf'])
    const doc = path.join(dir, 'notes.doc')
    await legacyFixture(doc)
    const disguised = path.join(dir, 'disguised.docx')
    await legacyFixture(disguised)
    const printer = recordingPrinter([])
    await withForcedOff(async () => {
      await assert.rejects(combineFiles([{ path: pdf }, { path: xls }], () => {}, { printHtml: printer }), (error) => {
        assert.equal(error.code, 'NEEDS_OFFICE_ENGINE')
        assert.match(error.message, /^Complex workbook\.xls: Simple can't combine older Excel workbooks \(\.xls\) on this PC\. Open it in Simple and save it as \.xlsx first, then add the \.xlsx file\.$/)
        return true
      })
      await assert.rejects(combineFiles([{ path: doc }, { path: pdf }], () => {}, { printHtml: printer }), (error) => (
        error.code === 'NEEDS_OFFICE_ENGINE' && /^notes\.doc: .*save it as \.docx first/.test(error.message) && !SETUP_WORDING.test(error.message)
      ))
      await assert.rejects(combineFiles([{ path: disguised }, { path: pdf }], () => {}, { printHtml: printer }), (error) => (
        error.code === 'NEEDS_OFFICE_ENGINE' && /^disguised\.docx: .*\.doc\)/.test(error.message)
      ), 'an older binary Word file named .docx is recognised by its content')
    })
    for (const name of ['report.doc', 'budget.xls']) {
      const error = legacyFormatError(name)
      assert.equal(error.code, 'NEEDS_OFFICE_ENGINE')
      assert.doesNotMatch(error.message, SETUP_WORDING)
      assert.doesNotMatch(error.message, /report|budget/, 'the file name is added by the caller')
    }
    assert.ok(['doc', 'xls', 'csv', 'docx', 'xlsx', 'ods'].every((extension) => COMBINE_EXTENSIONS.includes(extension)))
  } finally { await removeScratch(dir) }
})

test("Simple's own Word layout: page breaks, safe links, the document's font and margins", async () => {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
  zip.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  zip.file('word/_rels/document.xml.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/a" TargetMode="External"/><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/></Relationships>')
  zip.file('word/styles.xml', '<?xml version="1.0"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Cambria"/><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>')
  zip.file('word/document.xml', '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>'
    + '<w:p><w:r><w:t>Page one</w:t></w:r></w:p>'
    + '<w:p><w:r><w:rPr><w:b/></w:rPr><w:br w:type="page"/></w:r><w:r><w:t>Page two</w:t></w:r></w:p>'
    + '<w:p><w:hyperlink r:id="rId8"><w:r><w:t>web</w:t></w:r></w:hyperlink><w:hyperlink r:id="rId9"><w:r><w:t>script</w:t></w:r></w:hyperlink></w:p>'
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="720" w:right="1080" w:bottom="720" w:left="1080"/></w:sectPr></w:body></w:document>')
  const bytes = await zip.generateAsync({ type: 'nodebuffer' })
  const { jobs } = await native.docxPrintJobs(bytes, { name: 'Letter.docx' })
  assert.equal(jobs.length, 1)
  const { html, options } = jobs[0]
  assert.equal(options.title, 'Letter')
  assert.match(html, /@page \{ size: 8\.268in 11\.693in; margin: 0\.5in 0\.75in 0\.5in 0\.75in; \}/)
  assert.match(html, /font-family: "Cambria", Calibri/)
  assert.match(html, /font-size: 12pt/)
  assert.match(html, /<p>Page one<\/p><div class="page-break"><\/div><p>Page two<\/p>/)
  assert.match(html, /<a href="https:\/\/example\.com\/a">web<\/a>/)
  assert.match(html, /<a href="#">script<\/a>/)
  assert.doesNotMatch(html, /javascript:/)
  await assert.rejects(native.docxPrintJobs(Buffer.from('not a zip at all'), { name: 'x.docx' }), (error) => error.code === 'DAMAGED')
  assert.equal(native.sniffContainer(bytes), 'zip')
})

test('CSV, Excel header codes and page-break helpers', () => {
  assert.deepEqual(native.parseCsv('a,b\r\n"x, y","say ""hi"""\n').rows, [['a', 'b'], ['x, y', 'say "hi"']])
  assert.deepEqual(native.parseCsv('a;b;c\n1;2;3\n4;5;6').delimiter, ';')
  assert.deepEqual(native.parseCsv('a\tb\n1\t2').delimiter, '\t')
  assert.deepEqual(native.parseCsv('sep=|\na|b\n').rows, [['a', 'b']])
  assert.deepEqual(native.parseCsv('"line\nbreak",2\n').rows, [['line\nbreak', '2']])
  const sections = native.excelHeaderSections('&L&"Arial,Bold"&12Budget && plan&C&KFF0000Page &P of &N&R&A', { sheetName: 'Q<1>' })
  assert.deepEqual(sections, { left: 'Budget &amp; plan', center: 'Page <span class="pageNumber"></span> of <span class="totalPages"></span>', right: 'Q&lt;1&gt;' })
  assert.equal(native.excelHeaderSections('', {}), null)
  assert.equal(native.applyPageBreaks('<li>a<br class="page-break" />b</li>'), '<li>ab</li>', 'breaks inside list items are dropped, never mis-nested')
  assert.equal(native.sanitizeLinks('<a href="mailto:a@b.c">m</a><a href="file:///C:/x">f</a>'), '<a href="mailto:a@b.c">m</a><a href="#">f</a>')
})

test('html-to-pdf request policy and options: data and local job files only, never remote or shared paths', () => {
  const root = path.join(os.tmpdir(), 'simple-html-pdf-policy')
  const inside = pathToFileURL(path.join(root, 'document.html')).href
  assert.equal(htmlToPdf.isAllowedRequest(inside, [root]), true)
  assert.equal(htmlToPdf.isAllowedRequest('data:image/png;base64,AAAA', [root]), true)
  assert.equal(htmlToPdf.isAllowedRequest(pathToFileURL(path.join(os.tmpdir(), 'other', 'x.png')).href, [root]), false)
  assert.equal(htmlToPdf.isAllowedRequest(`${pathToFileURL(root).href}/../outside.png`, [root]), false)
  for (const url of ['https://example.com/a.png', 'http://127.0.0.1/a', 'ws://127.0.0.1/', 'ftp://host/x', 'blob:null/1', 'javascript:alert(1)', 'file://server/share/x.png', 'file://127.0.0.1/c$/x.png', 'chrome://settings', 'not a url']) {
    assert.equal(htmlToPdf.isAllowedRequest(url, [root]), false, url)
  }
  assert.equal(htmlToPdf.isAllowedRequest(inside, []), false, 'no roots: nothing local loads')

  const wrapped = htmlToPdf.prepareHtmlDocument('<p>Hi</p>', { title: 'A <b>' })
  assert.match(wrapped, /^<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none';/)
  assert.match(wrapped, /<title>A &lt;b&gt;<\/title>/)
  const full = htmlToPdf.prepareHtmlDocument('<!doctype html><html lang="en"><head><title>T</title><header></header></head><body></body></html>')
  assert.match(full, /<head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy"[^>]*><title>T<\/title><header><\/header>/)

  assert.deepEqual(htmlToPdf.printOptionsFor({ pageSize: { width: 8.27, height: 11.69 }, landscape: true, margins: { top: 0.6 }, scale: 0.5, footerTemplate: '<b>f</b>' }), {
    printBackground: true, landscape: true, preferCSSPageSize: false, generateTaggedPDF: true, generateDocumentOutline: true,
    pageSize: { width: 8.27, height: 11.69 }, margins: { top: 0.6, right: 0, bottom: 0, left: 0 }, scale: 0.5,
    displayHeaderFooter: true, headerTemplate: '<span></span>', footerTemplate: '<b>f</b>',
  })
  assert.equal(htmlToPdf.printOptionsFor({}).preferCSSPageSize, true)
  for (const bad of [{ scale: 3 }, { pageSize: 'B9' }, { margins: { top: -1 } }, { pageSize: { width: 0, height: 5 } }]) {
    assert.throws(() => htmlToPdf.printOptionsFor(bad), (error) => error.code === 'INVALID_OPTIONS')
  }
})

test('shared html-to-pdf source: built-ins and electron only, the vendoring header, no remote addresses', async () => {
  const source = await fs.readFile(path.join(ROOT, 'shared', 'electron', 'html-to-pdf.cjs'), 'utf8')
  assert.equal(source.split(/\r?\n/, 1)[0], '// Vendored from simple/shared/electron/html-to-pdf.cjs by simple/scripts/sync-shared.cjs. Do not edit here.')
  for (const match of source.matchAll(/require\(\s*['"]([^'"]+)['"]\s*\)/g)) assert.match(match[1], /^(node:|electron$)/, match[1])
  assert.deepEqual(source.match(/\b(?:https?|ftp|wss?):\/\/[^\s'"`)]+/gi) || [], [])
  assert.match(source, /javascript: false/)
  assert.match(source, /sandbox: true/)
  assert.match(source, /DEFAULT_TIMEOUT_MS = 30_000/)
})

/**
 * Runs inside Electron's main process (serialized by runElectronScenario):
 * prints hostile HTML against a local listener, then combines copies of the
 * samples through the bundled worker and the real html-to-pdf printer.
 */
function electronScenario(config) {
  const { app } = require('electron')
  const fs = require('node:fs')
  const http = require('node:http')
  app.setPath('userData', config.profile)
  app.on('window-all-closed', () => {})
  app.whenReady().then(async () => {
    const result = {}
    const hits = []
    const server = http.createServer((request, response) => { hits.push(request.url); response.end('x') })
    try {
      await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve))
      const { printHtmlToPdf } = require(config.htmlToPdf)
      const { runCombine } = require(config.combineHost)
      const remote = `http://127.0.0.1:${server.address().port}`
      const hostile = `<!doctype html><html><head><link rel="stylesheet" href="${remote}/a.css"><meta http-equiv="refresh" content="0;url=${remote}/refresh">`
        + `<style>@import url(${remote}/b.css); p { background: url(${remote}/c.png) }</style></head><body><p>Hostile page</p>`
        + `<img src="${remote}/d.png"><iframe src="${remote}/e"></iframe><object data="${remote}/f"></object><img src="file://127.0.0.1/c$/simple-unc-probe.png">`
        + `<script>fetch('${remote}/g')</script></body></html>`
      const hostilePdf = await printHtmlToPdf(hostile)
      await new Promise((resolve) => setTimeout(resolve, 300))
      result.hostile = { signature: hostilePdf.subarray(0, 5).toString('latin1'), hits: [...hits] }
      try {
        await printHtmlToPdf('<p>too slow</p>', { timeoutMs: 1 })
        result.timeout = 'none'
      } catch (error) { result.timeout = error.code }
      const started = Date.now()
      const combined = await runCombine(config.entries, () => {}, { printHtml: printHtmlToPdf, workerPath: config.workerPath })
      fs.writeFileSync(config.outputPath, combined.bytes)
      result.combined = { pageCount: combined.pageCount, builtIn: combined.builtIn, milliseconds: Date.now() - started }
      result.hitsAfter = hits.length
    } catch (error) {
      result.error = { code: error.code, message: error.message, stack: error.stack }
    } finally {
      server.close()
      fs.writeFileSync(config.resultPath, JSON.stringify(result))
      app.exit(0)
    }
  })
}

/** Starts Electron with an isolated profile and stops the whole process tree when done or late. */
async function runElectronScenario(directory, config, timeoutMs) {
  const electronBinary = require('electron')
  const script = path.join(directory, 'scenario.cjs')
  const resultPath = path.join(directory, 'scenario-result.json')
  const profile = path.join(directory, 'profile')
  await fs.writeFile(script, `(${electronScenario.toString()})(${JSON.stringify({ ...config, resultPath, profile })})\n`)
  const env = { ...process.env, SIMPLE_FORCE_NO_OFFICE: '1' }
  delete env.ELECTRON_RUN_AS_NODE
  const child = spawn(electronBinary, [script, `--user-data-dir=${profile}`], { env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
  let output = ''
  child.stdout.on('data', (data) => { output = (output + data).slice(-4000) })
  child.stderr.on('data', (data) => { output = (output + data).slice(-4000) })
  const stopTree = () => new Promise((resolve) => {
    if (process.platform === 'win32') execFile('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => resolve())
    else { child.kill('SIGKILL'); resolve() }
  })
  let timedOut = false
  const timer = setTimeout(() => { timedOut = true; void stopTree() }, timeoutMs)
  try {
    await new Promise((resolve) => child.once('exit', resolve))
  } finally {
    clearTimeout(timer)
    await stopTree()
  }
  assert.equal(timedOut, false, `Electron scenario timed out. ${output}`)
  return JSON.parse(await fs.readFile(resultPath, 'utf8').catch(() => `{"error":{"message":${JSON.stringify(`no result. ${output}`)}}}`))
}

async function pageTexts(pdfBytes) {
  const pdfjs = await import(pathToFileURL(path.join(ROOT, '..', 'simple_pdf_source', 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.mjs')).href)
  const document = await pdfjs.getDocument({ data: new Uint8Array(pdfBytes), disableWorker: true, isEvalSupported: false, verbosity: 0 }).promise
  try {
    const texts = []
    for (let number = 1; number <= document.numPages; number += 1) {
      const page = await document.getPage(number)
      texts.push((await page.getTextContent()).items.map((item) => item.str || '').join(' ').replace(/\s+/g, ' '))
    }
    return texts
  } finally { await document.destroy() }
}

test('Electron, SIMPLE_FORCE_NO_OFFICE=1: PDF + DOCX + XLSX copies combine through the hardened printer', { skip: process.platform !== 'win32' && 'Windows only', timeout: 240_000 }, async () => {
  const { outfile } = await workerBundle()
  const dir = await scratch('electron')
  try {
    const copies = await copySamples(dir, ['Complex document.pdf', 'Complex document.docx', 'Complex workbook.xlsx', 'Complex workbook.ods'])
    const csv = path.join(dir, 'people.csv')
    await fs.writeFile(csv, 'Name,City\nБорис,Toshkent\n')
    const order = ['Complex document.pdf', 'Complex document.docx', 'Complex workbook.xlsx', 'Complex workbook.ods']
    const entries = [...order.map((name) => ({ path: copies[name] })), { path: csv }]
    const before = await Promise.all(entries.map(async (entry) => hash(await fs.readFile(entry.path))))
    const outputPath = path.join(dir, 'Combined.pdf')
    const result = await runElectronScenario(dir, {
      htmlToPdf: path.join(ROOT, 'shared', 'electron', 'html-to-pdf.cjs'),
      combineHost: path.join(ROOT, 'launcher', 'combine-host.cjs'),
      workerPath: outfile,
      entries,
      outputPath,
    }, 180_000)
    assert.equal(result.error, undefined, JSON.stringify(result.error))
    assert.equal(result.hostile.signature, '%PDF-')
    assert.deepEqual(result.hostile.hits, [], 'the print window never reached the network')
    assert.equal(result.hitsAfter, 0)
    assert.equal(result.timeout, 'PRINT_TIMEOUT')
    assert.deepEqual(result.combined.builtIn, ['Complex document.docx', 'Complex workbook.xlsx', 'Complex workbook.ods', 'people.csv'])

    const output = await fs.readFile(outputPath)
    const combined = await PDFDocument.load(output)
    const sourcePages = (await PDFDocument.load(await fs.readFile(copies['Complex document.pdf']))).getPageCount()
    assert.equal(combined.getPageCount(), result.combined.pageCount)
    assert.ok(combined.getPageCount() >= sourcePages + 4, `expected at least ${sourcePages + 4} pages, got ${combined.getPageCount()}`)
    const texts = await pageTexts(output)
    const converted = texts.slice(sourcePages)
    assert.match(converted[0], /STRESS_START/, 'the Word document follows the PDF pages')
    assert.deepEqual(combined.getPage(sourcePages).getSize(), { width: 612, height: 792 }, 'Letter, from the document')
    assert.ok(converted.some((text) => /Quarterly operations/.test(text)), 'the workbook summary is printed')
    const transactions = converted.findIndex((text) => /Transaction detail/.test(text))
    assert.ok(transactions >= 0)
    const transactionsPage = combined.getPage(sourcePages + transactions).getSize()
    assert.ok(transactionsPage.width > transactionsPage.height, 'the Transactions sheet keeps its landscape setup')
    assert.ok(converted.some((text) => /Con\s*fi\s*dential sample/.test(text) && /Page 1 of \d/.test(text)), 'Excel headers and footers are printed')
    assert.ok(texts.every((text) => !/HIDDEN_AUDIT_SENTINEL/.test(text)), 'hidden sheets never reach the PDF')
    assert.match(converted.at(-1), /Борис/)
    assert.deepEqual(await Promise.all(entries.map(async (entry) => hash(await fs.readFile(entry.path)))), before, 'sources are only read')
  } finally { await removeScratch(dir) }
})

/** The printer as Chromium writes a job: a link inside the job is a named destination in its catalog. */
function namedLinkPrinter(jobs) {
  return async (html, options) => {
    jobs.push({ html, options })
    const pdf = await PDFDocument.create()
    const first = pdf.addPage([300, 300])
    const second = pdf.addPage([300, 300])
    const context = pdf.context
    first.node.set(PDFName.of('Annots'), context.obj([
      context.register(context.obj({ Type: 'Annot', Subtype: 'Link', Rect: [0, 0, 50, 50], P: first.ref, Dest: PDFName.of('footnote-1') })),
    ]))
    pdf.catalog.set(PDFName.of('Dests'), context.obj({ 'footnote-1': context.obj([second.ref, 'XYZ', 0, 300, 0]) }))
    return pdf.save()
  }
}

test('links inside a file Simple lays out itself (bookmarks, footnotes) still work in the combined PDF', async () => {
  const dir = await scratch('print-links')
  try {
    const { 'Complex document.docx': docx } = await copySamples(dir, ['Complex document.docx'])
    const cover = path.join(dir, 'cover.pdf')
    const coverPdf = await PDFDocument.create(); coverPdf.addPage([200, 200])
    await fs.writeFile(cover, await coverPdf.save())
    const jobs = []
    const result = await withForcedOff(() => combineFiles([{ path: cover }, { path: docx }], () => {}, { printHtml: namedLinkPrinter(jobs) }))
    assert.ok(jobs.length >= 1)
    const combined = await PDFDocument.load(result.bytes, { updateMetadata: false })
    const pages = combined.getPages()
    assert.equal(pages.length, 1 + jobs.length * 2)
    const links = []
    pages.forEach((page, index) => {
      const annots = page.node.Annots()
      for (let position = 0; annots && position < annots.size(); position += 1) links.push({ index, annotation: combined.context.lookup(annots.get(position)) })
    })
    assert.equal(links.length, jobs.length, 'every job keeps its internal link')
    for (const { index, annotation } of links) {
      assert.equal(String(annotation.get(PDFName.of('Dest'))), `[ ${pages[index + 1].ref} /XYZ 0 300 0 ]`, 'the link leads to the page after it, as printed')
    }
  } finally { await removeScratch(dir) }
})

/** A password-protected .docx or .xlsx: an encrypted package inside a compound file. */
function encryptedOfficeBytes() {
  const bytes = Buffer.alloc(1024)
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(bytes)
  const entry = (offset, name, type) => {
    bytes.write(name, offset, 'utf16le')
    bytes.writeUInt16LE((name.length + 1) * 2, offset + 64)
    bytes[offset + 66] = type
  }
  entry(512, 'Root Entry', 5)
  entry(640, 'EncryptionInfo', 2)
  entry(768, 'EncryptedPackage', 2)
  return bytes
}

test('a password-protected .docx or .xlsx is named as protected, never as an older .doc or .xls', async () => {
  const dir = await scratch('protected-office')
  try {
    const pdf = path.join(dir, 'cover.pdf')
    const coverPdf = await PDFDocument.create(); coverPdf.addPage([200, 200])
    await fs.writeFile(pdf, await coverPdf.save())
    const budget = path.join(dir, 'Budget.xlsx')
    const report = path.join(dir, 'Report.docx')
    await fs.writeFile(budget, encryptedOfficeBytes())
    await fs.writeFile(report, encryptedOfficeBytes())
    const protectedError = (name) => (error) => {
      assert.equal(error.code, 'ENCRYPTED')
      assert.equal(error.message, `${name}: This file is password-protected, and Simple can't read password-protected Word or Excel files. Add a copy saved without a password.`)
      assert.doesNotMatch(error.message, /older|\.doc\)|\.xls\)|save it as/)
      return true
    }
    for (const [file, name] of [[budget, 'Budget.xlsx'], [report, 'Report.docx']]) {
      // Without an engine Simple would lay it out itself; with one, it never reaches the engine.
      await withForcedOff(() => assert.rejects(combineFiles([{ path: pdf }, { path: file }], () => {}, { printHtml: recordingPrinter([]) }), protectedError(name)))
      const convert = async () => { throw new Error('a protected file never reaches the office engine') }
      await assert.rejects(combineFiles([{ path: file }, { path: pdf }], () => {}, { convertOfficeBytes: convert }), protectedError(name))
    }
  } finally { await removeScratch(dir) }
})

test('files picked for Combine that cannot be reached are named with a plain reason, never Node text', async () => {
  const { assertTargetIsNotASource, describeCombinePaths } = require('../launcher/combine-paths.cjs')
  const dir = await scratch('picked')
  try {
    const kept = path.join(dir, 'kept.pdf')
    await fs.writeFile(kept, '%PDF-1.4\n')
    const gone = path.join(dir, 'gone.pdf')
    await assert.rejects(describeCombinePaths([kept, gone]), (error) => (
      error.code === 'NOT_FOUND' && error.message === 'gone.pdf: This file was moved, renamed or deleted. Add it again from its new place.'
    ))
    const failing = (code) => ({ stat: async () => { throw Object.assign(new Error(`${code}: raw text, stat '${kept}'`), { code, syscall: 'stat' }) } })
    await assert.rejects(describeCombinePaths([kept], { fs: failing('EACCES') }), (error) => error.code === 'LOCKED' && /^kept\.pdf: Simple couldn't read this file\./.test(error.message) && error.technical === 'EACCES stat')
    await assert.rejects(describeCombinePaths([kept], { fs: failing('EIO') }), (error) => error.code === 'FILE_UNAVAILABLE' && /^kept\.pdf: This file isn't available right now\./.test(error.message))
    await assert.rejects(describeCombinePaths([kept], { fs: failing('ELOOP') }), (error) => error.code === 'UNREADABLE' && !/ELOOP|stat/.test(error.message))
    const legacy = path.join(dir, 'old.doc')
    await legacyFixture(legacy)
    const described = await describeCombinePaths([kept, legacy], { officeEngineStatus: async () => ({ available: false }) })
    assert.deepEqual(described.entries.map((entry) => entry.name), ['kept.pdf'])
    assert.equal(described.skipped[0].code, 'NEEDS_OFFICE_ENGINE')

    // A source deleted while the Save dialog was open.
    await assert.rejects(assertTargetIsNotASource([{ path: kept }, { path: gone }], path.join(dir, 'Combined.pdf')), (error) => (
      error.code === 'NOT_FOUND' && error.message.startsWith('gone.pdf: ') && !/ENOENT|realpath/.test(error.message)
    ))
    await assert.rejects(assertTargetIsNotASource([{ path: kept }], path.join(dir, 'KEPT.pdf')), (error) => error.code === 'SOURCE_TARGET')
    await assertTargetIsNotASource([{ path: kept }], path.join(dir, 'Combined.pdf'))
  } finally { await removeScratch(dir) }
})
