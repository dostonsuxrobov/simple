const test = require('node:test')
const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawnSync } = require('node:child_process')
const {
  PDFArray, PDFDict, PDFDocument, PDFName, PDFRawStream, StandardFonts, decodePDFRawStream, rgb,
  beginText, endText, popGraphicsState, pushGraphicsState, setFontAndSize, setTextMatrix,
  setTextRenderingMode, showText, TextRenderingMode,
} = require('pdf-lib')
const fixture = require('./fixtures/scan-fixture.cjs')
const layer = require('../electron/ocr-text-layer.cjs')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { documentContainsText } = require('./helpers/pdf-test-utils.cjs')

const pdfjsModule = () => import('pdfjs-dist/legacy/build/pdf.mjs')
const mupdfModule = () => import('mupdf')
const META = { engine: 'tesseract.js 7.0.0', language: 'eng' }
const collapse = (text) => String(text).replace(/\s+/g, ' ').trim()

let cleanScan
const clean300 = () => (cleanScan ??= fixture.buildScanVariant('clean300'))

/**
 * The 'ocr-text-layer' payload a perfect recognition of the vector source page
 * would produce, read from mupdf's structured text of that page: each word's
 * baseline origin and advance extent, the gaps between words, and the line
 * box height as the font size (Tesseract's row height).
 */
async function vectorPayload() {
  const mupdf = await mupdfModule()
  const vector = await fixture.vectorPage()
  const doc = mupdf.Document.openDocument(vector.bytes, 'application/pdf')
  const page = doc.loadPage(0)
  const inverse = mupdf.Matrix.invert(page.getTransform())
  const toPdf = ([x, y]) => ({ x: x * inverse[0] + y * inverse[2] + inverse[4], y: x * inverse[1] + y * inverse[3] + inverse[5] })
  const lines = []
  let chars
  const text = page.toStructuredText('preserve-whitespace')
  try {
    text.walk({
      beginLine() { chars = [] },
      onChar(c, origin, _font, _size, quad) {
        chars.push({ c, origin: toPdf(origin), top: toPdf([quad[0], quad[1]]).y, bottom: toPdf([quad[4], quad[5]]).y, left: toPdf([quad[4], quad[5]]).x, right: toPdf([quad[6], quad[7]]).x })
      },
      endLine() { lines.push(chars) },
    })
  } finally {
    text.destroy()
    page.destroy()
    doc.destroy()
  }
  return {
    vector,
    lines: lines.map((line) => {
      const words = []
      let word = null
      for (const char of line) {
        if (/\s/.test(char.c)) { word = null; continue }
        if (!word) words.push(word = [])
        word.push(char)
      }
      return {
        fontSize: Math.max(...line.map((char) => char.top - char.bottom)),
        words: words.map((wordChars, index) => {
          const first = wordChars[0]
          const last = wordChars[wordChars.length - 1]
          const next = words[index + 1]
          return {
            text: wordChars.map((char) => char.c).join(''),
            x: first.origin.x,
            y: first.origin.y,
            dx: 1,
            dy: 0,
            width: last.right - first.left,
            gap: next ? next[0].left - last.right : 0,
          }
        }),
      }
    }),
  }
}

async function withLayer(bytes, pages, meta = META) {
  const doc = await PDFDocument.load(bytes)
  const op = layer.validateOcrLayerOperation({ type: 'ocr-text-layer', meta, pages }, doc.getPageCount())
  const fontRef = layer.addGlyphlessFont(doc)
  for (const page of op.pages) layer.addOcrTextLayer(doc, page.pageIndex, page.lines, { fontRef, meta: op.meta })
  return doc.save({ useObjectStreams: true })
}

async function pdfjsPage(bytes, run) {
  const pdfjs = await pdfjsModule()
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false, verbosity: 0 }).promise
  try {
    return await run(await pdf.getPage(1), pdfjs)
  } finally {
    await pdf.destroy()
  }
}

const pageText = (bytes) => pdfjsPage(bytes, async (page) => (await page.getTextContent()).items.map((item) => item.str + (item.hasEOL ? '\n' : '')).join(''))

/** The first page's content streams, decoded and joined. */
async function pageContent(bytes) {
  const doc = await PDFDocument.load(bytes)
  const contents = doc.getPage(0).node.Contents()
  const streams = contents instanceof PDFArray ? contents.asArray().map((ref) => doc.context.lookup(ref)) : [contents]
  return streams.map((stream) => Buffer.from(decodePDFRawStream(stream).decode()).toString('latin1')).join('\n')
}

/** sha256 of every image XObject's stored (encoded) bytes. */
async function imageStreamHashes(bytes) {
  const doc = await PDFDocument.load(bytes)
  return doc.context.enumerateIndirectObjects()
    .filter(([, object]) => object instanceof PDFRawStream && object.dict.get(PDFName.of('Subtype'))?.toString() === '/Image')
    .map(([, object]) => crypto.createHash('sha256').update(object.getContents()).digest('hex'))
}

async function rendered(bytes, scale = 0.5) {
  const mupdf = await mupdfModule()
  const doc = mupdf.Document.openDocument(bytes, 'application/pdf')
  const page = doc.loadPage(0)
  const pixmap = page.toPixmap(mupdf.Matrix.scale(scale, scale), mupdf.ColorSpace.DeviceRGB, false)
  try { return Buffer.from(pixmap.getPixels()) } finally { pixmap.destroy(); page.destroy(); doc.destroy() }
}

/** Draw `text` in render mode 3 with a standard font, the way OCR tools other than Simple do. */
async function drawInvisible(doc, page, text, x, y, size = 11, fontName = StandardFonts.TimesRoman) {
  const font = await doc.embedFont(fontName)
  const key = page.node.newFontDictionary('Ocr', font.ref)
  page.pushOperators(pushGraphicsState(), beginText(), setTextRenderingMode(TextRenderingMode.Invisible),
    setFontAndSize(key, size), setTextMatrix(1, 0, 0, 1, x, y), showText(font.encodeText(text)), endText(), popGraphicsState())
}

const truthWords = (truth) => truth.lines.flatMap((line) => line.words.map((word) => word.text))

test('the layer reads back as the scanned text on the source baselines and widths, in GlyphLessFont render mode 3', async () => {
  const { vector, lines } = await vectorPayload()
  const scan = await clean300()
  const output = await withLayer(scan.pdf, [{ pageIndex: 0, lines }])
  const nonEmpty = (items) => items.filter((item) => item.str)
  const vectorItems = await pdfjsPage(vector.bytes, async (page) => nonEmpty((await page.getTextContent()).items))
  const result = await pdfjsPage(output, async (page, pdfjs) => {
    const content = await page.getTextContent()
    const list = await page.getOperatorList()
    const { textPaintResolver } = await import('../electron/text-appearance.mjs')
    const paint = textPaintResolver(list, pdfjs.OPS)
    const items = nonEmpty(content.items)
    return {
      items,
      text: content.items.map((item) => item.str + (item.hasEOL ? '\n' : '')).join(''),
      styles: content.styles,
      fontObject: page.commonObjs.get(items[0].fontName),
      modes: new Set(list.fnArray.flatMap((fn, index) => (fn === pdfjs.OPS.setTextRenderingMode ? [list.argsArray[index][0]] : []))),
      paints: items.map((item) => paint(item.fontName, item.str)),
    }
  })

  assert.equal(collapse(result.text), collapse(vector.lines.map((line) => line.text).join(' ')), 'pdf.js text equals the truth')
  // One item per line, except where wide table gaps split the source line too.
  for (const line of vector.lines.filter((entry) => !/\s{2}/.test(entry.text))) {
    assert.equal(result.items.filter((item) => item.str === line.text).length, 1, `one item for "${line.text.slice(0, 30)}…"`)
  }
  assert.deepEqual(result.items.map((item) => collapse(item.str) || ' '), vectorItems.map((item) => collapse(item.str) || ' '))
  result.items.forEach((item, index) => {
    const source = vectorItems[index]
    assert.ok(Math.abs(item.transform[5] - source.transform[5]) <= 0.5, `baseline of "${item.str}"`)
    assert.ok(Math.abs(item.transform[4] - source.transform[4]) <= 0.5, `start of "${item.str}"`)
    assert.ok(Math.abs(item.width - source.width) <= source.width * 0.015, `width of "${item.str}": ${item.width} vs ${source.width}`)
  })
  const style = result.styles[result.items[0].fontName]
  assert.ok(Math.abs(style.ascent - 0.8) <= 0.01, `ascent ${style.ascent}`)
  assert.ok(Math.abs(style.descent + 0.2) <= 0.01, `descent ${style.descent}`)
  assert.equal(result.fontObject.name, 'GlyphLessFont')
  assert.deepEqual([...result.modes], [3], 'only invisible text is drawn')
  assert.ok(result.paints.every((paint) => paint.invisible && !paint.color), 'every item is recognised as invisible text')
  assert.ok(output.length - scan.pdf.length < 8000, `the layer adds ${output.length - scan.pdf.length} bytes`)
})

test('mupdf extracts and finds the recognised words where the source page has them', async () => {
  const mupdf = await mupdfModule()
  const { vector, lines } = await vectorPayload()
  const output = await withLayer((await clean300()).pdf, [{ pageIndex: 0, lines }])
  const open = (bytes) => {
    const doc = mupdf.Document.openDocument(bytes, 'application/pdf')
    const page = doc.loadPage(0)
    const text = page.toStructuredText('preserve-whitespace')
    return { text, close: () => { text.destroy(); page.destroy(); doc.destroy() } }
  }
  const layerText = open(output)
  const sourceText = open(vector.bytes)
  try {
    assert.equal(collapse(layerText.text.asText()), collapse(vector.lines.map((line) => line.text).join(' ')))
    // Whole words only: inside a word the layer spaces glyphs evenly.
    const words = ['Quarterly', 'committee', 'negotiations', 'warehouse', 'September', '7741-0093', '$1,250.75', 'SKU-0311', 'Gasket', 'Alvarez', 'inspection;', 'expansion']
    let worst = 0
    for (const word of words) {
      const found = layerText.text.search(word)
      const expected = sourceText.text.search(word)
      assert.equal(found.length, 1, `${word} found once`)
      assert.equal(found.length, expected.length)
      found.forEach((hit, index) => hit.forEach((quad, part) => quad.forEach((value, coordinate) => {
        worst = Math.max(worst, Math.abs(value - expected[index][part][coordinate]))
      })))
    }
    assert.ok(worst <= 0.75, `search quads within ${worst.toFixed(3)} pt`)
  } finally {
    layerText.close()
    sourceText.close()
  }
})

test('on a /Rotate 90 page with a CropBox offset the words read in order at the scanned positions', async () => {
  const truth = await fixture.groundTruth()
  const scan = await fixture.buildScanVariant('rotated90Crop')
  // The fixture's portrait layout turned onto the landscape media (scan-fixture.cjs).
  const map = (u, v) => ({ x: 50 + (fixture.PAGE_HEIGHT - v), y: 40 + u })
  const lines = truth.lines.map((line) => ({
    fontSize: line.size,
    words: line.words.map((word, index) => {
      const origin = map(word.x0, line.baseline)
      const along = map(word.x0 + 1, line.baseline)
      const next = line.words[index + 1]
      return { text: word.text, x: origin.x, y: origin.y, dx: along.x - origin.x, dy: along.y - origin.y, width: word.x1 - word.x0, gap: next ? next.x0 - word.x1 : 0 }
    }),
  }))
  const output = await withLayer(scan.pdf, [{ pageIndex: 0, lines }])
  const { items, toViewport } = await pdfjsPage(output, async (page) => {
    const viewport = page.getViewport({ scale: 300 / 72, rotation: page.rotate })
    assert.equal(page.rotate, 90)
    assert.deepEqual(page.view, [50, 40, 842, 652])
    return { items: (await page.getTextContent()).items.filter((item) => item.str.trim()), toViewport: (x, y) => viewport.convertToViewportPoint(x, y) }
  })
  assert.deepEqual(items.map((item) => item.str).join(' ').split(/\s+/).filter(Boolean), truthWords(truth), 'reading order equals the truth')
  const sourceWords = lines.flatMap((line) => line.words)
  let next = 0
  let worst = 0
  for (const item of items) {
    const words = item.str.trim().split(/\s+/)
    const word = sourceWords[next]
    assert.equal(word.text, words[0])
    const [ix, iy] = toViewport(item.transform[4], item.transform[5])
    const [wx, wy] = toViewport(word.x, word.y)
    worst = Math.max(worst, Math.hypot(ix - wx, iy - wy))
    next += words.length
  }
  assert.ok(worst <= 0.5, `item origins within ${worst.toFixed(3)} px at 300 DPI`)
})

test('recognising again with replaceExisting keeps each word once and leaves the scan image untouched', async () => {
  const { invoke } = loadMain()
  const { vector, lines } = await vectorPayload()
  const scan = await clean300()
  const operation = { type: 'ocr-text-layer', replaceExisting: true, meta: META, pages: [{ pageIndex: 0, lines }] }
  const once = await invoke('pdf:mutate', scan.pdf, operation)
  const twice = await invoke('pdf:mutate', once, operation)
  const expected = collapse(vector.lines.map((line) => line.text).join(' '))
  assert.equal(collapse(await pageText(once)), expected)
  assert.equal(collapse(await pageText(twice)), expected, 'every word exactly once after replacing')
  const glyphs = lines.reduce((sum, line) => sum + line.words.reduce((count, word) => count + word.text.length, 0) + line.words.length - 1, 0)
  assert.deepEqual(await layer.invisibleTextCensus(twice), [{ pageIndex: 0, invisible: glyphs, visible: 0 }])
  assert.deepEqual(await imageStreamHashes(twice), await imageStreamHashes(scan.pdf), 'the image stream bytes are unchanged')
  assert.deepEqual(await rendered(twice), await rendered(scan.pdf), 'the page looks exactly as before')
  assert.ok((await pageContent(twice)).includes('/SimpleOCR <</Engine (tesseract.js 7.0.0) /Lang (eng) /Version 1>> BDC'), 'layer marker written')
  // The replaced layer's font and content streams are not kept as orphans.
  const fontPrograms = async (bytes) => (await PDFDocument.load(bytes)).context.enumerateIndirectObjects()
    .filter(([, object]) => object instanceof PDFDict && object.get(PDFName.of('FontName'))?.toString() === '/GlyphLessFont').length
  assert.equal(await fontPrograms(twice), 1)
  assert.ok(Math.abs(twice.length - once.length) < 200, `replacing again does not grow the file (${once.length} -> ${twice.length})`)

  // Without replaceExisting the runner skips searchable pages; the writer itself appends.
  const appended = await invoke('pdf:mutate', twice, { ...operation, replaceExisting: false })
  assert.equal((await pageText(appended)).split('negotiations').length - 1, 2)
})

test('a layer from another OCR tool is replaced while visible text on the page stays', async () => {
  const { invoke } = loadMain()
  const scan = await clean300()
  const doc = await PDFDocument.load(scan.pdf)
  const page = doc.getPage(0)
  const helvetica = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Visible native footer line', { x: 72, y: 40, size: 10, font: helvetica, color: rgb(0, 0, 0) })
  await drawInvisible(doc, page, 'Third party OCR words', 72, 681.3)
  const input = await doc.save()
  assert.deepEqual(await layer.invisibleTextCensus(input), [{ pageIndex: 0, invisible: 21, visible: 26 }])

  const lines = [{ fontSize: 11, words: [{ text: 'Recognized', x: 72, y: 681.3, dx: 1, dy: 0, width: 50, gap: 3 }, { text: 'again', x: 125, y: 681.3, dx: 1, dy: 0, width: 25, gap: 0 }] }]
  const output = await invoke('pdf:mutate', input, { type: 'ocr-text-layer', replaceExisting: true, meta: META, pages: [{ pageIndex: 0, lines }] })
  const text = await pageText(output)
  assert.match(text, /Visible native footer line/)
  assert.match(text, /Recognized again/)
  assert.doesNotMatch(text, /Third party/)
  assert.deepEqual(await layer.invisibleTextCensus(output), [{ pageIndex: 0, invisible: 16, visible: 26 }])
  assert.deepEqual(await rendered(output), await rendered(input), 'visible content and the scan are unchanged')
  assert.deepEqual(await imageStreamHashes(output), await imageStreamHashes(input))
  assert.equal(await documentContainsText(output, 'Third party OCR words'), false, 'the replaced text is not left in the file')
  assert.equal(await documentContainsText(output, 'Visible native footer line'), true)

  // Pages listed without words only lose their old layer.
  const cleared = await invoke('pdf:mutate', output, { type: 'ocr-text-layer', replaceExisting: true, meta: META, pages: [{ pageIndex: 0, lines: [] }] })
  assert.deepEqual(await layer.invisibleTextCensus(cleared), [{ pageIndex: 0, invisible: 0, visible: 26 }])
  assert.equal(collapse(await pageText(cleared)), 'Visible native footer line')
})

test('invisible text drawn exactly over visible text is kept instead of taking the visible text with it', async () => {
  const doc = await PDFDocument.create()
  const page = doc.addPage([300, 200])
  const helvetica = await doc.embedFont(StandardFonts.Helvetica)
  page.drawText('Signed total 1,250.75', { x: 30, y: 120, size: 12, font: helvetica, color: rgb(0, 0, 0) })
  await drawInvisible(doc, page, 'Signed total 1,250.75', 30, 120, 12, StandardFonts.Helvetica)
  await drawInvisible(doc, page, 'stale words', 30, 60)
  const input = await doc.save()
  const kept = []
  const output = await layer.removeInvisibleText(input, [0], { onKept: (entry) => kept.push(entry) })
  assert.deepEqual(kept, [{ pageIndex: 0, count: 21 }])
  assert.deepEqual(await layer.invisibleTextCensus(output), [{ pageIndex: 0, invisible: 21, visible: 21 }])
  assert.deepEqual(await rendered(output, 2), await rendered(input, 2))
  assert.doesNotMatch(await pageText(output), /stale/)

  const untouched = await PDFDocument.create()
  untouched.addPage([200, 200]).drawText('Only visible text', { x: 20, y: 100, size: 12 })
  const plain = await untouched.save()
  assert.equal(await layer.removeInvisibleText(plain, [0]), plain, 'bytes without invisible text are returned as they are')
})

test('the embedded GlyphLessFont, its CID map and ToUnicode CMap are well formed', async () => {
  const fontkit = require('@pdf-lib/fontkit')
  const output = await withLayer((await clean300()).pdf, [{ pageIndex: 0, lines: [{ fontSize: 12, words: [{ text: 'Ünïcode ✓ 中文', x: 72, y: 700, dx: 1, dy: 0, width: 90, gap: 0 }] }] }])
  const doc = await PDFDocument.load(output)
  const objects = doc.context.enumerateIndirectObjects().map(([, object]) => object)
  const dicts = objects.map((object) => (object instanceof PDFRawStream ? object.dict : object)).filter((object) => object instanceof PDFDict)
  const named = (type, key, value) => dicts.find((dict) => dict.get(PDFName.of('Type'))?.toString() === type && dict.get(PDFName.of(key))?.toString() === value)
  const decoded = (ref) => Buffer.from(decodePDFRawStream(doc.context.lookup(ref)).decode())

  const type0 = named('/Font', 'Subtype', '/Type0')
  assert.equal(type0.get(PDFName.of('BaseFont')).toString(), '/GlyphLessFont')
  assert.equal(type0.get(PDFName.of('Encoding')).toString(), '/Identity-H')
  const cidFont = doc.context.lookup(doc.context.lookup(type0.get(PDFName.of('DescendantFonts'))).get(0))
  assert.equal(cidFont.get(PDFName.of('Subtype')).toString(), '/CIDFontType2')
  assert.equal(cidFont.get(PDFName.of('DW')).toString(), '500')
  const map = decoded(cidFont.get(PDFName.of('CIDToGIDMap')))
  assert.equal(map.length, 131072)
  for (let index = 0; index < map.length; index += 2) assert.ok(map[index] === 0 && map[index + 1] === 1)

  const descriptor = named('/FontDescriptor', 'FontName', '/GlyphLessFont')
  assert.equal(descriptor.get(PDFName.of('Ascent')).toString(), '800')
  assert.equal(descriptor.get(PDFName.of('Descent')).toString(), '-200')
  const program = decoded(descriptor.get(PDFName.of('FontFile2')))
  assert.deepEqual(program, Buffer.from(layer.GLYPHLESS_TTF_BASE64, 'base64'))
  const font = fontkit.create(program)
  assert.equal(font.unitsPerEm, 2048)
  assert.equal(font.ascent, 1638)
  assert.equal(font.descent, -410)
  // Table checksums and head.checkSumAdjustment match the patched bytes.
  const checksum = (bytes, offset, length) => {
    let total = 0
    for (let index = 0; index < ((length + 3) & ~3); index += 4) {
      const byte = (k) => (offset + index + k < bytes.length ? bytes[offset + index + k] : 0)
      total = (total + ((byte(0) << 24) | (byte(1) << 16) | (byte(2) << 8) | byte(3))) >>> 0
    }
    return total
  }
  for (let index = 0; index < program.readUInt16BE(4); index += 1) {
    const entry = 12 + index * 16
    const tag = program.toString('latin1', entry, entry + 4)
    const offset = program.readUInt32BE(entry + 8)
    const length = program.readUInt32BE(entry + 12)
    // head's own checksum is taken with checkSumAdjustment set to 0.
    const bytes = Buffer.from(program)
    if (tag === 'head') bytes.writeUInt32BE(0, offset + 8)
    assert.equal(program.readUInt32BE(entry + 4), checksum(bytes, offset, length), `${tag} checksum`)
  }
  assert.equal(checksum(program, 0, program.length), 0xb1b0afba, 'whole-font checksum')

  const cmap = decoded(type0.get(PDFName.of('ToUnicode'))).toString('latin1')
  const ranges = [...cmap.matchAll(/<([0-9A-F]{4})> <([0-9A-F]{4})> <([0-9A-F]{4})>/g)]
  assert.equal(ranges.length, 248)
  assert.ok(ranges.every(([, start, end, target]) => start === target && start.slice(0, 2) === end.slice(0, 2) && start.endsWith('00') && end.endsWith('FF')))
  assert.ok(ranges.every(([, start]) => parseInt(start.slice(0, 2), 16) < 0xd8 || parseInt(start.slice(0, 2), 16) > 0xdf), 'no surrogate ranges')
  assert.ok([...cmap.matchAll(/(\d+) beginbfrange/g)].every(([, count]) => Number(count) <= 100))
  assert.equal(collapse(await pageText(output)), 'Ünïcode ✓ 中文')
})

test('the layer request is validated and normalised before anything is written', () => {
  const word = (extra = {}) => ({ text: 'word', x: 72, y: 700, dx: 1, dy: 0, width: 30, gap: 4, ...extra })
  const request = (pages, extra = {}) => ({ type: 'ocr-text-layer', meta: META, pages, ...extra })
  const page = (words, extra = {}) => ({ pageIndex: 0, lines: [{ fontSize: 12, words, ...extra }] })
  const rejects = (operation, pattern, pageCount = 2) => assert.throws(() => layer.validateOcrLayerOperation(operation, pageCount), (error) => error.code === 'OCR_LAYER_INVALID' && pattern.test(error.message))

  rejects(request([page([word({ x: NaN })])]), /invalid x/)
  rejects(request([page([word({ y: Infinity })])]), /invalid y/)
  rejects(request([page([word()], { fontSize: 5000 })]), /font size/)
  rejects(request([page([word()], { fontSize: 0.1 })]), /font size/)
  rejects(request([{ pageIndex: 2, lines: [] }]), /not in the document/)
  rejects(request([{ pageIndex: -1, lines: [] }]), /not in the document/)
  rejects(request([{ pageIndex: 0.5, lines: [] }]), /not in the document/)
  rejects(request([{ pageIndex: 0, lines: [] }, { pageIndex: 0, lines: [] }]), /listed twice/)
  rejects(request([page([word({ dx: 0, dy: 0 })])]), /direction/)
  rejects(request([page([word({ dx: 2, dy: 0 })])]), /direction/)
  rejects(request([page([word({ width: -1 })])]), /width or gap/)
  rejects(request([page([word({ gap: -0.5 })])]), /width or gap/)
  rejects(request([page([word({ x: 1e9 })])]), /outside the page/)
  rejects(request([page([word({ text: 42 })])]), /no text/)
  rejects(request([page([word()], { words: 'word' })]), /no list of words/)
  rejects({ ...request([]), type: 'crop' }, /not a text layer/)
  rejects({ type: 'ocr-text-layer' }, /lists no pages/)
  rejects(request([page(Array.from({ length: 25_001 }, () => word()))]), /more than 25,000 words/)
  rejects(request([page([word({ text: 'x'.repeat(2 * 1024 * 1024 + 1) })])]), /too much text/)

  const normalised = layer.validateOcrLayerOperation(request([page([
    word({ text: ' ﬁnal\u0007 ', dx: 0.9995, dy: 0.0001 }),
    word({ text: 'a\u{1F600}b' }),
    word({ text: ' \u0001 ' }),
  ]), { pageIndex: 1, lines: [{ fontSize: 9, words: [word({ text: '' })] }] }], { replaceExisting: 'yes', meta: { engine: 'tesseract.js (7)\\', language: 'eng+deu; drop table' } }), 2)
  assert.equal(normalised.replaceExisting, false, 'only true replaces')
  assert.deepEqual(normalised.meta, { engine: 'tesseract.js (7)\\', language: 'eng+deudroptable' })
  assert.deepEqual(normalised.pages[0].lines[0].words.map((entry) => entry.text), ['final', 'a\uFFFDb'])
  assert.ok(Math.abs(Math.hypot(normalised.pages[0].lines[0].words[0].dx, normalised.pages[0].lines[0].words[0].dy) - 1) < 1e-12)
  assert.deepEqual(normalised.pages[1], { pageIndex: 1, lines: [] }, 'lines without words are dropped')
  // The same text rules as the renderer's ocr-geometry normalisation.
  return import('../electron/ocr-geometry.mjs').then(({ normalizeOcrText }) => {
    for (const sample of ['ﬁnancial ﬃce ﬀ ﬂ ﬄ ﬅ ﬆ', 'Café', 'a\u{1F600}b', 'x\uD800y\uDC00z', 'a\u0001b\u007f\u0085c\td', 'e\u0301']) {
      assert.equal(layer.normalizeLayerText(sample), normalizeOcrText(sample), JSON.stringify(sample))
    }
  })
})

test('pages that inherit one /Resources dictionary each get their layer through a single font entry', async () => {
  const source = await PDFDocument.create()
  const helvetica = await source.embedFont(StandardFonts.Helvetica)
  for (const label of ['First page heading', 'Second page heading']) source.addPage([400, 300]).drawText(label, { x: 30, y: 250, size: 14, font: helvetica })
  // Move page 1's resources up to the page tree and let both pages inherit them.
  const shared = source.getPage(0).node.Resources()
  source.getPage(1).node.Resources().get(PDFName.of('Font')).entries().forEach(([key, value]) => shared.get(PDFName.of('Font')).set(key, value))
  source.catalog.Pages().set(PDFName.of('Resources'), source.context.register(shared))
  for (const page of source.getPages()) page.node.delete(PDFName.of('Resources'))
  const input = await source.save()
  const word = (text) => ({ fontSize: 12, words: [{ text, x: 30, y: 100, dx: 1, dy: 0, width: 60, gap: 0 }] })
  const output = await withLayer(input, [{ pageIndex: 0, lines: [word('alpha')] }, { pageIndex: 1, lines: [word('beta')] }])
  const pdfjs = await pdfjsModule()
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(output), isEvalSupported: false, verbosity: 0 }).promise
  try {
    const texts = []
    for (let number = 1; number <= 2; number += 1) texts.push(collapse((await (await pdf.getPage(number)).getTextContent()).items.map((item) => item.str).join(' ')))
    assert.deepEqual(texts, ['First page heading alpha', 'Second page heading beta'])
  } finally {
    await pdf.destroy()
  }
  const doc = await PDFDocument.load(output)
  const fonts = doc.getPages().map((page) => page.node.Resources().lookup(PDFName.of('Font'), PDFDict))
  const glyphless = fonts.map((dict) => dict.entries().filter(([key]) => key.toString().startsWith('/SimpleOCR')).length)
  assert.deepEqual(glyphless, [1, 1], 'one GlyphLessFont entry, reused by the page that shares the dictionary')
})

test('touching words and skewed lines still read as one line with single spaces', async () => {
  const blank = await PDFDocument.create()
  blank.addPage([612, 792])
  const word = (text, x, width, y = 700, dx = 1, dy = 0) => ({ text, x, y, dx, dy, width, gap: 0 })
  // Ink widths differ per letter, so every word needs its own stretch.
  const slanted = (degrees, y) => {
    const angle = (degrees * Math.PI) / 180
    let along = 0
    return [['The', 15], ['committee', 52], ['reviewed', 38], ['the', 13], ['budget', 33], ['WMW', 30]].map(([text, width]) => {
      const entry = word(text, 72 + along * Math.cos(angle), width, y + along * Math.sin(angle), Math.cos(angle), Math.sin(angle))
      along += width + 3
      return entry
    })
  }
  const output = await withLayer(await blank.save(), [{ pageIndex: 0, lines: [
    // "beta" ends exactly where "gamma" starts; "gamma" overlaps "delta".
    { fontSize: 10, words: [word('alpha', 72, 30), word('beta', 102, 20), word('gamma', 122, 30), word('delta', 150, 25)] },
    { fontSize: 10, words: slanted(4, 500) },
    { fontSize: 10, words: slanted(-1.2, 300) },
  ] }])
  const items = await pdfjsPage(output, async (page) => (await page.getTextContent()).items.filter((item) => item.str))
  assert.deepEqual(items.map((item) => item.str), ['alpha beta gamma delta', 'The committee reviewed the budget WMW', 'The committee reviewed the budget WMW'])
})

test('the content stream escapes metadata, clamps scaling and fits slanted words with one scale', () => {
  const along = (text, distance, width) => ({ text, x: 30 + 0.6 * distance, y: 40 + 0.8 * distance, dx: 0.6, dy: 0.8, width, gap: 0 })
  const content = layer.ocrLayerContent('/SimpleOCR-1', [
    { fontSize: 10, words: [
      { text: 'ab', x: 10, y: 20, dx: 1, dy: 0, width: 0, gap: 0 },
      { text: 'c', x: 10.0004, y: 20, dx: 1, dy: 0, width: 1e5, gap: 0 },
    ] },
    { fontSize: 10, words: [along('abc', 0, 18), along('de', 22, 8)] },
  ], { engine: 'eng (x)\\', language: 'eng' })
  assert.equal(content, [
    'q',
    '/SimpleOCR <</Engine (eng \\(x\\)\\\\) /Lang (eng) /Version 1>> BDC',
    'BT',
    '3 Tr 0 Tc 0 Tw 0 Ts',
    '/SimpleOCR-1 10 Tf',
    // Touching words: the first is narrowed as far as allowed and a space still follows.
    '1 0 0 1 10 20 Tm',
    '1 Tz <00610062> Tj',
    '<0020> Tj',
    '1 0 0 1 10 20 Tm',
    '5000 Tz <0063> Tj',
    '/SimpleOCR-1 10 Tf',
    // A slanted line: one Tz for its words, each fitted with character spacing
    // inside the range pdf.js still reads as one word (+0.09 / -0.18 em).
    '0.6 0.8 -0.8 0.6 30 40 Tm',
    '0.9 Tc 100 Tz <006100620063> Tj',
    '0 Tc 86 Tz <0020> Tj',
    '0.6 0.8 -0.8 0.6 43.2 57.6 Tm',
    '-1.8 Tc 100 Tz <00640065> Tj',
    'ET',
    'EMC',
    'Q',
    '',
  ].join('\n'))
})

test('writeInvisibleRun adds one stretched invisible run that pdf.js reads and the caller can roll back', async () => {
  const doc = await PDFDocument.create()
  const page = doc.addPage([400, 300])
  const fontRef = layer.addGlyphlessFont(doc)
  const key = layer.glyphlessFontKey(page, fontRef)
  assert.equal(layer.glyphlessFontKey(page, fontRef), key, 'one resource entry per font')
  assert.equal(layer.writeInvisibleRun(page, key, { text: 'Corrected total', origin: { x: 40, y: 150 }, dir: { x: 1, y: 0 }, width: 120, fontSize: 11, meta: META }), true)
  assert.equal(layer.writeInvisibleRun(page, key, { text: ' \u0002 ', origin: { x: 40, y: 100 }, dir: { x: 1, y: 0 }, width: 10, fontSize: 11 }), false, 'nothing to write')
  assert.throws(() => layer.writeInvisibleRun(page, key, { text: 'x', origin: { x: NaN, y: 1 }, dir: { x: 1, y: 0 }, width: 10, fontSize: 11 }), /invalid x/)
  const bytes = await doc.save()
  const { items, modes } = await pdfjsPage(bytes, async (pdfPage, pdfjs) => {
    const list = await pdfPage.getOperatorList()
    return {
      items: (await pdfPage.getTextContent()).items.filter((item) => item.str),
      modes: list.fnArray.flatMap((fn, index) => (fn === pdfjs.OPS.setTextRenderingMode ? [list.argsArray[index][0]] : [])),
    }
  })
  assert.equal(items.length, 1)
  assert.equal(items[0].str, 'Corrected total')
  assert.ok(Math.abs(items[0].width - 120) < 0.01)
  assert.deepEqual(items[0].transform.slice(4), [40, 150])
  assert.deepEqual(modes, [3])
  assert.ok(page.contentStream.operators.some((operator) => operator.toString().includes('BDC')))
})

test('text drawn after recognition is still visible: the layer does not leak its render mode', async () => {
  const { invoke } = loadMain()
  const blank = await PDFDocument.create()
  blank.addPage([300, 200])
  const recognised = await withLayer(await blank.save(), [{ pageIndex: 0, lines: [{ fontSize: 12, words: [{ text: 'hidden', x: 20, y: 150, dx: 1, dy: 0, width: 40, gap: 0 }] }] }])
  const flattened = await invoke('pdf:flatten-overlays', recognised, [{
    id: 'added', type: 'text', pageIndex: 0, text: 'Visible addition', fontFamily: 'Arial', fontSize: 14,
    rect: { x: 20, y: 60, width: 200, height: 24 }, color: [0, 0, 0],
  }], {}, {})
  const mupdf = await mupdfModule()
  const doc = mupdf.Document.openDocument(flattened, 'application/pdf')
  const page = doc.loadPage(0)
  const pixmap = page.toPixmap(mupdf.Matrix.identity, mupdf.ColorSpace.DeviceGray, false)
  try {
    const pixels = pixmap.getPixels()
    const width = pixmap.getWidth()
    let ink = 0
    for (let y = 200 - 84; y < 200 - 60; y += 1) for (let x = 20; x < 220; x += 1) if (pixels[y * width + x] < 128) ink += 1
    assert.ok(ink > 50, `the added text is painted (${ink} dark pixels)`)
  } finally {
    pixmap.destroy()
    page.destroy()
    doc.destroy()
  }
  assert.match(collapse(await pageText(flattened)), /hidden.*Visible addition|Visible addition.*hidden/)
})

test('the bundled backend writes and replaces the layer without reading font files from disk', { timeout: 120_000 }, async () => {
  const esbuild = require('esbuild')
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-ocr-bundle-'))
  try {
    // The same bundling as simple/scripts/sync-build.cjs: one CommonJS backend,
    // electron and mupdf external, mupdf copied to vendor/ beside it.
    await esbuild.build({
      entryPoints: [path.join(__dirname, '..', 'electron', 'main.cjs')],
      outfile: path.join(temp, 'electron', 'main.cjs'),
      bundle: true,
      platform: 'node',
      format: 'cjs',
      target: 'node22',
      external: ['electron', 'mupdf'],
      legalComments: 'none',
      logLevel: 'silent',
    })
    const mupdfRoot = path.dirname(path.dirname(require.resolve('mupdf')))
    const vendor = path.join(temp, 'vendor', 'mupdf')
    fs.mkdirSync(path.join(vendor, 'dist'), { recursive: true })
    fs.copyFileSync(path.join(mupdfRoot, 'package.json'), path.join(vendor, 'package.json'))
    for (const name of ['mupdf.js', 'mupdf-wasm.js', 'mupdf-wasm.wasm']) fs.copyFileSync(path.join(mupdfRoot, 'dist', name), path.join(vendor, 'dist', name))

    const scan = await clean300()
    const doc = await PDFDocument.load(scan.pdf)
    await drawInvisible(doc, doc.getPage(0), 'Older layer words', 72, 681.3)
    fs.writeFileSync(path.join(temp, 'input.pdf'), await doc.save())
    fs.writeFileSync(path.join(temp, 'operation.json'), JSON.stringify({
      type: 'ocr-text-layer', replaceExisting: true, meta: META,
      pages: [{ pageIndex: 0, lines: [{ fontSize: 11, words: [{ text: 'Bundled', x: 72, y: 681.3, dx: 1, dy: 0, width: 40, gap: 3 }, { text: 'layer', x: 115, y: 681.3, dx: 1, dy: 0, width: 25, gap: 0 }] }] }],
    }))
    fs.writeFileSync(path.join(temp, 'run.cjs'), `
const fs = require('node:fs')
const Module = require('node:module')
const handlers = new Map()
class WebContents { constructor() { this.id = 1 } on() {} once() {} send() {} }
class BrowserWindow {
  constructor() { this.webContents = new WebContents() }
  static fromWebContents() { return null }
  static getAllWindows() { return [] }
  static getFocusedWindow() { return null }
  removeMenu() {} loadURL() {} loadFile() {} once() {} on() {} show() {} close() {} destroy() {}
  isDestroyed() { return false }
}
const electron = {
  app: { isPackaged: true, requestSingleInstanceLock: () => true, on() {}, whenReady: () => Promise.resolve(), getVersion: () => 'test', quit() {} },
  BrowserWindow, dialog: {}, nativeImage: { createFromPath: () => ({}) }, shell: {},
  ipcMain: { handle: (channel, handler) => handlers.set(channel, handler), on() {} },
}
const load = Module._load
Module._load = function (request, ...rest) { return request === 'electron' ? electron : load.call(this, request, ...rest) }
require('./electron/main.cjs')
setImmediate(async () => {
  try {
    const output = await handlers.get('pdf:mutate')({ sender: new WebContents() }, fs.readFileSync('input.pdf'), JSON.parse(fs.readFileSync('operation.json', 'utf8')))
    fs.writeFileSync('output.pdf', output)
  } catch (error) {
    console.error(error)
    process.exitCode = 1
  }
})
`)
    const run = spawnSync(process.execPath, ['run.cjs'], { cwd: temp, env: { ...process.env, TEMP: temp, TMP: temp, NODE_PATH: '' }, encoding: 'utf8', timeout: 90_000 })
    assert.equal(run.status, 0, run.stderr || run.stdout)
    const output = fs.readFileSync(path.join(temp, 'output.pdf'))
    assert.equal(collapse(await pageText(output)), 'Bundled layer')
    assert.deepEqual(await imageStreamHashes(output), await imageStreamHashes(scan.pdf))
  } finally {
    fs.rmSync(temp, { recursive: true, force: true, maxRetries: 3 })
  }
})
