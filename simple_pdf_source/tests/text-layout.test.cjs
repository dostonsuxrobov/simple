const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const { createRequire } = require('node:module')
const { pathToFileURL } = require('node:url')
const test = require('node:test')
const { PDFDocument, StandardFonts } = require('pdf-lib')

const layoutModule = import('../electron/text-layout.mjs')

test('fitting defaults preserve native lines and wrap newly added text', async () => {
  const { resolveTextFit, layoutText } = await layoutModule
  assert.equal(resolveTextFit({ originalText: 'Existing PDF line' }), 'fit')
  assert.equal(resolveTextFit({ originalText: '' }), 'wrap')
  assert.equal(resolveTextFit({ originalText: 'line one\nline two' }), 'wrap')
  assert.equal(resolveTextFit({ originalText: 'Existing PDF line', textFit: 'wrap' }), 'wrap')
  const fit = layoutText('A longer line', 30, (value) => value.length * 10, 'fit')
  assert.deepEqual(fit.lines, ['A longer line'])
  assert.equal(fit.fitScale, 30 / 130)
})

test('wrap handles long identifiers, blank paragraphs, and Unicode characters', async () => {
  const { layoutText } = await layoutModule
  const measure = (value) => Array.from(value).length * 10
  const wrapped = layoutText('hello world\n\nabcdefgh\n😀😀😀', 40, measure, 'wrap')
  assert.deepEqual(wrapped.lines, ['hell', 'o', 'worl', 'd', '', 'abcd', 'efgh', '😀😀😀'])
  assert.ok(wrapped.lines.every((line) => measure(line) <= 40))
})

function pdfWriter() {
  // Load the real writer functions without launching Electron or registering
  // application IPC. The writer itself is exercised unchanged against PDFs.
  const entry = path.resolve(__dirname, '../electron/main.cjs')
  const source = fs.readFileSync(entry, 'utf8').split('\nconst gotLock =')[0]
    .replace("import('./text-layout.mjs')", `import(${JSON.stringify(pathToFileURL(path.resolve(__dirname, '../electron/text-layout.mjs')).href)})`)
  const localRequire = createRequire(entry)
  return Function('require', `${source}\nreturn { flattenOverlays };`)((id) => id === 'electron' ? {} : localRequire(id))
}

async function fixture() {
  const document = await PDFDocument.create()
  const page = document.addPage([400, 500])
  const font = await document.embedFont(StandardFonts.Courier)
  const originalText = 'SOURCE'
  page.drawText(originalText, { x: 40, y: 400, size: 12, font })
  return {
    data: await document.save(),
    edit: {
      id: 'replacement', type: 'text', pageIndex: 0,
      rect: { x: 40, y: 180, width: 120, height: 230 },
      originalRect: { x: 40, y: 398, width: font.widthOfTextAtSize(originalText, 12), height: 12 },
      originalText, text: 'Bigger words wrap across several lines',
      fontSize: 24, fontFamily: 'Courier New', fontWeight: 400, fontStyle: 'normal',
      lineHeight: 29, scaleX: 1, letterSpacing: 1,
      textFit: 'wrap', preserveSourceMetrics: false,
      color: [0, 0, 0], align: 'left', cover: true,
    },
  }
}

async function savedText(data) {
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const document = await pdfjs.getDocument({ data: new Uint8Array(data), isEvalSupported: false }).promise
  try {
    return (await (await document.getPage(1)).getTextContent()).items.filter((item) => 'str' in item)
  } finally { await document.destroy() }
}

test('saving native text with intentional formatting retains size, spacing, and every wrapped word', async () => {
  const { data, edit } = await fixture()
  const result = await pdfWriter().flattenOverlays(data, [edit])
  const items = (await savedText(result)).filter((item) => item.str && item.str !== 'SOURCE')
  assert.equal(items.map((item) => item.str).join('').replace(/\s+/g, ''), edit.text.replace(/\s+/g, ''))
  assert.ok(new Set(items.map((item) => Math.round(item.transform[5]))).size > 1)
  for (const item of items) {
    assert.ok(Math.abs(item.transform[0] - 24) < 0.01, 'Font width was recalibrated back to the original source')
    assert.ok(Math.abs(item.transform[3] - 24) < 0.01, 'Chosen font size was not preserved')
    assert.ok(item.width <= edit.rect.width + 2, 'A wrapped line exceeds the text box')
  }
})

test('fitting a much longer native replacement writes one complete line within the box', async () => {
  const { data, edit } = await fixture()
  edit.text = 'A much longer replacement line'
  edit.textFit = 'fit'
  const result = await pdfWriter().flattenOverlays(data, [edit])
  const items = (await savedText(result)).filter((item) => item.str && item.str !== 'SOURCE')
  assert.equal(items.map((item) => item.str).join(''), edit.text)
  assert.equal(new Set(items.map((item) => Math.round(item.transform[5]))).size, 1)
  assert.ok(items.reduce((width, item) => width + item.width, 0) <= edit.rect.width + 2)
})

test('saving overflowing wrapped text reports the box instead of dropping typed words', async () => {
  const { data, edit } = await fixture()
  edit.rect.height = 25
  await assert.rejects(pdfWriter().flattenOverlays(data, [edit]), /Text on page 1 does not fit its box/)
})

test('a native baseline below its selection box can be saved and exported', async () => {
  const { data, edit } = await fixture()
  edit.rect = { ...edit.originalRect, y: 402, height: 8 }
  edit.baselineOffset = -2
  edit.text = 'UPDATED'
  edit.textFit = 'fit'
  edit.fontSize = 12
  const output = await pdfWriter().flattenOverlays(data, [edit])
  const items = await savedText(output)
  assert.equal(items.map(item => item.str).join(''), 'UPDATED')
  assert.ok(Math.abs(items[0].transform[5] - 400) < 0.01)
})
