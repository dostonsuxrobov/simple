const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const { PDFDocument } = require('../../simple_pdf_source/node_modules/pdf-lib')
const { combineFiles, pageIndices } = require('../launcher/combine-service.cjs')
const { atomicWrite, runCombine } = require('../launcher/combine-host.cjs')

test('page selections preserve user order, support ranges, and reject missing pages', () => {
  assert.deepEqual(pageIndices('3, 1-2', 3), [2, 0, 1])
  assert.deepEqual(pageIndices('', 2), [0, 1])
  for (const value of ['0', '3-1', '4', '1,', 'cat', '1-10000000']) assert.throws(() => pageIndices(value, 3))
})

test('combine worker preserves file order, page selections and page sizes without touching inputs', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-combine-test-'))
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
    const result = await runCombine([{ path: b }, { path: a, pages: '2,1' }], (value) => progress.push(value.name))
    const combined = await PDFDocument.load(result.bytes)
    assert.deepEqual(combined.getPages().map((page) => [page.getWidth(), page.getHeight()]), [[600, 200], [400, 500], [200, 300]])
    assert.equal(result.pageCount, 3)
    assert.deepEqual(progress, ['two.pdf', 'one.pdf'])
    assert.deepEqual(await fs.readFile(a), original)
    await assert.rejects(combineFiles([{ path: a, pages: '99' }, { path: b }]), /one.pdf: Choose pages/)
    const destination = path.join(dir, 'output.pdf')
    await fs.writeFile(destination, 'existing')
    await atomicWrite(destination, result.bytes)
    assert.equal((await PDFDocument.load(await fs.readFile(destination))).getPageCount(), 3)
    assert.equal((await fs.readdir(dir)).some((name) => name.endsWith('.tmp')), false)
  } finally { await fs.rm(dir, { recursive: true, force: true }) }
})

test('Word combination uses the layout engine and preserves its page geometry', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-combine-doc-'))
  try {
    const source = path.join(dir, 'legacy.doc')
    await fs.writeFile(source, 'test source')
    const pdf = await PDFDocument.create()
    pdf.addPage([595.28, 841.89])
    const output = await pdf.save()
    let calls = 0
    const result = await combineFiles([{ path: source }, { path: source }], () => {}, {
      convertOfficeBytes: async (input) => {
        calls++
        assert.equal(input.inputExtension, 'doc')
        assert.equal(input.outputExtension, 'pdf')
        assert.equal(input.filter, 'writer_pdf_Export')
        return output
      },
    })
    assert.equal(calls, 2)
    assert.equal(result.pageCount, 2)
    const combined = await PDFDocument.load(result.bytes)
    assert.equal(combined.getPage(0).getWidth(), 595.28)
  } finally { await fs.rm(dir, { recursive: true, force: true }) }
})

test('XLS, XLSX and ODS use spreadsheet print layout, retain page ranges and preserve source bytes', async () => {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-combine-sheet-'))
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
    const options = { convertOfficeBytes: async (input) => {
      calls.push(input.inputExtension)
      assert.equal(input.outputExtension, 'pdf')
      assert.equal(input.filter, 'calc_pdf_Export')
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
  } finally {
    const absolute = path.resolve(dir)
    assert.ok(absolute.startsWith(path.resolve(os.tmpdir()) + path.sep) && path.basename(absolute).startsWith('simple-combine-sheet-'))
    await fs.rm(absolute, { recursive: true, force: true })
  }
})

test('interactive and encryption-marked PDFs are rejected before pages are copied', async () => {
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'simple-combine-protected-'))
  try {
    const formPdf=await PDFDocument.create()
    const page=formPdf.addPage([300,300])
    const field=formPdf.getForm().createTextField('Name')
    field.setText('Keep this value');field.addToPage(page,{x:10,y:40,width:200,height:30})
    const source=path.join(dir,'form.pdf')
    const original=await formPdf.save()
    await fs.writeFile(source,original)
    await assert.rejects(combineFiles([{path:source},{path:source}]),/form.pdf:.*interactive form fields.*flatten/)
    assert.deepEqual(await fs.readFile(source),Buffer.from(original))
    const protectedPdf=await PDFDocument.create();protectedPdf.addPage([300,300])
    protectedPdf.context.trailerInfo.Encrypt=protectedPdf.context.register(protectedPdf.context.obj({Filter:'Standard',V:1,R:2,P:-4}))
    await fs.writeFile(source,await protectedPdf.save())
    await assert.rejects(combineFiles([{path:source},{path:source}]),/form.pdf: Unlock this PDF/)
  } finally {
    const absolute=path.resolve(dir)
    assert.ok(absolute.startsWith(path.resolve(os.tmpdir())+path.sep)&&path.basename(absolute).startsWith('simple-combine-protected-'))
    await fs.rm(absolute,{recursive:true,force:true})
  }
})
