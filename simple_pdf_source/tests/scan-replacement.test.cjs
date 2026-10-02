const test = require('node:test')
const assert = require('node:assert/strict')

// Minimal-footprint replacement of edited scanned text (design 4.8.5 / 6.2(7)).
const planner = () => import('../electron/scan-replacement.mjs')

// "Invoice number 48213 was paid in full on July 14." laid out at 10 pt with
// a 5 pt-per-character measure and 3 pt spaces: easy to reason about.
const CHAR = 5
const SPACE = 3
const measure = (text) => Array.from(text).reduce((width, character) => width + (character === ' ' ? SPACE : CHAR), 0)
function lineOf(text, x = 72) {
  const words = []
  let cursor = x
  for (const token of text.split(' ')) {
    words.push({ text: token, rect: { x: cursor, y: 700, width: measure(token), height: 10 } })
    cursor += measure(token) + SPACE
  }
  return words
}
const LINE = 'Invoice number 48213 was paid in full on July 14.'
const words = lineOf(LINE)
const plan = async (newText, extra = {}) => (await planner()).planScanReplacement({ words, originalText: LINE, newText, measure, spaceWidth: SPACE, ...extra })

test('identical text (or only different spacing) needs nothing', async () => {
  assert.deepEqual(await plan(LINE), { kind: 'none' })
  assert.deepEqual(await plan(`  ${LINE.replace(/ /g, '   ')}  `), { kind: 'none' })
})

test('a one-word substitution replaces just that word', async () => {
  const result = await plan(LINE.replace('48213', '48214'))
  assert.equal(result.kind, 'words')
  assert.equal(result.first, 2)
  assert.equal(result.last, 2)
  assert.equal(result.originalText, '48213')
  assert.equal(result.text, '48214')
  assert.deepEqual(result.rect, words[2].rect)
  // Room runs to the next word, less one space.
  assert.equal(result.start, words[2].rect.x)
  assert.equal(result.room, words[3].rect.x - SPACE - words[2].rect.x)
})

test('a slightly longer word still fits its room (up to 8 % squeeze)', async () => {
  // Room for "48213" is 25 pt (to the next word, less a space); "482130"
  // measures 30 pt > 25 * 1.08, so the rest of the line is re-typeset, which
  // fits once it may be squeezed by up to 8 % (161 pt into 156 pt).
  const longer = await plan(LINE.replace('48213', '482130'))
  assert.equal(longer.kind, 'words')
  assert.equal(longer.first, 2)
  assert.equal(longer.last, words.length - 1)
  assert.equal(longer.text, '482130 was paid in full on July 14.')
  // 5 % wider than its room: squeezed in place.
  const narrow = await plan(LINE.replace(' in ', ' im '), { measure: (text) => (text === 'im' ? 10.5 : measure(text)) })
  assert.equal(narrow.kind, 'words')
  assert.equal(narrow.first, 5)
  assert.equal(narrow.last, 5)
  assert.ok(10.5 <= narrow.room * 1.08)
})

test('a replacement too long for the rest of the line becomes a whole-line edit', async () => {
  const result = await plan(LINE.replace('July', 'September and October'), { lineEnd: words.at(-1).rect.x + words.at(-1).rect.width })
  assert.deepEqual(result, { kind: 'line' })
  // With room to the right of the line it is re-typeset from the change to the line end.
  const roomy = await plan(LINE.replace('July', 'Sept'), { lineEnd: 1000 })
  assert.equal(roomy.kind, 'words')
})

test('insertion and deletion', async () => {
  // A single space has no room for a word: the line end needs some.
  assert.deepEqual(await plan(LINE.replace('paid in', 'paid back in')), { kind: 'line' })
  const inserted = await plan(LINE.replace('paid in', 'paid back in'), { lineEnd: 1000 })
  assert.equal(inserted.kind, 'words')
  // No room in a single space: from the insertion point to the line end.
  assert.equal(inserted.first, 5)
  assert.equal(inserted.last, words.length - 1)
  assert.equal(inserted.text, 'back in full on July 14.')
  assert.equal(inserted.originalText, 'in full on July 14.')

  const deleted = await plan(LINE.replace(' full', ''))
  assert.equal(deleted.kind, 'words')
  assert.equal(deleted.first, 6)
  assert.equal(deleted.last, 6)
  assert.equal(deleted.text, '')
  assert.equal(deleted.originalText, 'full')

  const cleared = await plan('')
  assert.equal(cleared.kind, 'words')
  assert.equal(cleared.first, 0)
  assert.equal(cleared.last, words.length - 1)
  assert.equal(cleared.text, '')
})

test('punctuation attached to a word changes with it', async () => {
  const result = await plan(LINE.replace('14.', '15.'))
  assert.equal(result.kind, 'words')
  assert.equal(result.first, 9)
  assert.equal(result.originalText, '14.')
  assert.equal(result.text, '15.')
  const comma = await plan(LINE.replace('July 14.', 'July, 14.'), { lineEnd: 1000 })
  assert.equal(comma.kind, 'words')
  assert.equal(comma.originalText, 'July 14.')
  assert.equal(comma.text, 'July, 14.')
})

test('leading and trailing whitespace is ignored', async () => {
  const result = await plan(`   ${LINE.replace('was', 'is')}\t `)
  assert.equal(result.kind, 'words')
  assert.equal(result.first, 3)
  assert.equal(result.text, 'is')
})

test('line breaks, missing or misaligned word boxes edit the whole line', async () => {
  assert.deepEqual(await plan(LINE.replace(' on ', '\non ')), { kind: 'line' })
  const { planScanReplacement } = await planner()
  assert.deepEqual(planScanReplacement({ words: words.slice(1), originalText: LINE, newText: LINE.replace('was', 'is'), measure, spaceWidth: SPACE }), { kind: 'line' })
  assert.deepEqual(planScanReplacement({ words: undefined, originalText: LINE, newText: 'x', measure, spaceWidth: SPACE }), { kind: 'line' })
})

test('positions along a slanted baseline are honoured', async () => {
  const along = words.map((word, index) => ({ start: index * 40, end: index * 40 + measure(word.text) }))
  const result = await plan(LINE.replace('was', 'were'), { along })
  assert.equal(result.kind, 'words')
  assert.equal(result.start, 120)
  // Room to the next word along the line: 160 - 3 - 120.
  assert.equal(result.room, 37)
})

// ---------------------------------------------------------------------------
// Saving edits of scanned text (electron/main.cjs flattenOverlays): the
// retouch patch is drawn under the new text, only the replaced words' OCR
// text is removed, and nothing outside the patch changes.

const zlib = require('node:zlib')
const fixture = require('./fixtures/scan-fixture.cjs')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { pageTexts } = require('./helpers/pdf-test-utils.cjs')

const DPI = 300
const SCALE = DPI / 72

/** A PNG of RGBA pixels (what the renderer's canvas would encode). */
function encodePng(rgba, width, height) {
  const chunk = (type, data) => {
    const length = Buffer.alloc(4)
    length.writeUInt32BE(data.length)
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data])
    const crc = Buffer.alloc(4)
    crc.writeUInt32BE(zlib.crc32(body) >>> 0)
    return Buffer.concat([length, body, crc])
  }
  const header = Buffer.alloc(13)
  header.writeUInt32BE(width, 0)
  header.writeUInt32BE(height, 4)
  header[8] = 8
  header[9] = 6
  const rows = Buffer.alloc((width * 4 + 1) * height)
  for (let y = 0; y < height; y += 1) Buffer.from(rgba.buffer, rgba.byteOffset + y * width * 4, width * 4).copy(rows, y * (width * 4 + 1) + 1)
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', zlib.deflateSync(rows)), chunk('IEND', Buffer.alloc(0))])
}

/** The fixture scan with Simple's OCR text layer written from the ground truth (no OCR run needed). */
async function recognisedScan(name) {
  const layer = require('../electron/ocr-text-layer.cjs')
  const { PDFDocument } = require('pdf-lib')
  const variant = await fixture.buildScanVariant(name)
  const doc = await PDFDocument.load(variant.pdf)
  const lines = variant.truth.lines.map((line) => ({
    fontSize: line.size,
    words: line.words.map((word, index) => {
      const next = line.words[index + 1]
      const end = { x: word.baseline.origin.x + word.baseline.dir.x * word.inkWidth, y: word.baseline.origin.y + word.baseline.dir.y * word.inkWidth }
      return {
        text: word.text, x: word.baseline.origin.x, y: word.baseline.origin.y, dx: word.baseline.dir.x, dy: word.baseline.dir.y,
        width: word.inkWidth, gap: next ? Math.hypot(next.baseline.origin.x - end.x, next.baseline.origin.y - end.y) : 0,
      }
    }),
  }))
  const fontRef = layer.addGlyphlessFont(doc)
  layer.addOcrTextLayer(doc, 0, lines, { fontRef, meta: { engine: 'tesseract.js 7.0.0', language: 'eng' } })
  return { variant, bytes: await doc.save({ useObjectStreams: true }) }
}

/**
 * The overlay the renderer commits for an edit of a line where `word`
 * becomes `replacement` (or the whole line becomes `lineText`), built with
 * the same modules the preparation worker runs.
 */
async function scanOverlay({ variant, bytes }, startsWith, { word, replacement, mode = 'appearance', lineText } = {}) {
  const { estimateScanStyle } = await import('../electron/scan-style.mjs')
  const { retouchLine, composePatch } = await import('../electron/ocr-retouch.mjs')
  const line = variant.truth.lines.find((item) => item.text.startsWith(startsWith))
  const size = line.size
  const first = line.words[0]
  const last = line.words.at(-1)
  const dir = first.baseline.dir
  const origin = first.baseline.origin
  const length = Math.hypot(last.baseline.origin.x + dir.x * last.inkWidth - origin.x, last.baseline.origin.y + dir.y * last.inkWidth - origin.y)
  const up = { x: -dir.y, y: dir.x }
  const corner = (u, v) => ({ x: origin.x + dir.x * u + up.x * v, y: origin.y + dir.y * u + up.y * v })
  const corners = [corner(-0.25 * size, -0.65 * size), corner(length + 0.25 * size, -0.65 * size), corner(-0.25 * size, 1.25 * size), corner(length + 0.25 * size, 1.25 * size)]
  const px0 = Math.floor(Math.min(...corners.map((c) => c.x)) * SCALE)
  const px1 = Math.ceil(Math.max(...corners.map((c) => c.x)) * SCALE)
  const py0 = Math.floor((792 - Math.max(...corners.map((c) => c.y))) * SCALE)
  const py1 = Math.ceil((792 - Math.min(...corners.map((c) => c.y))) * SCALE)
  const raster = await fixture.rasterize(bytes, DPI)
  const width = px1 - px0
  const height = py1 - py0
  const pixels = new Uint8Array(width * height)
  for (let y = 0; y < height; y += 1) pixels.set(raster.data.subarray((y + py0) * raster.width + px0, (y + py0) * raster.width + px1), y * width)
  const toPdfRect = (box) => ({ x: (px0 + box.x0) / SCALE, y: 792 - (py0 + box.y1) / SCALE, width: (box.x1 - box.x0) / SCALE, height: (box.y1 - box.y0) / SCALE })
  const common = { channels: 1, dpi: DPI, baseline: { x: origin.x * SCALE - px0, y: (792 - origin.y) * SCALE - py0, dx: dir.x, dy: -dir.y }, length: length * SCALE, fontSize: size * SCALE, text: line.text }
  const { style, features } = estimateScanStyle(pixels, width, height, common)
  const retouch = retouchLine(pixels, width, height, { ...common, xHeight: features.xHeight, fontClass: style.fontClass, segment: true, seed: 'test-line' })
  assert.equal(retouch.segmented, true)
  const patchOf = (labels) => {
    const patch = composePatch(retouch, labels)
    return { rect: toPdfRect({ x0: patch.x, y0: patch.y, x1: patch.x + patch.width, y1: patch.y + patch.height }), dataUrl: `data:image/png;base64,${encodePng(patch.rgba, patch.width, patch.height).toString('base64')}`, dpi: DPI }
  }
  const words = retouch.words.map((item) => ({ text: item.text, rect: toPdfRect(item.box), start: item.box.u0 / SCALE, end: item.box.u1 / SCALE }))
  const run = { origin, dir, length, fontSize: size }
  const frame = (start, frameWidth) => ({ x: origin.x + start, y: origin.y - 0.3 * size, width: frameWidth, height: 1.2 * size })
  const scan = { key: 'd1|test', mode, status: 'ready', lineRect: frame(0, length), baseline: { origin, dir }, lineText: line.text, run, words, patch: patchOf(null), style, paper: style.background }
  let text = lineText ?? line.text
  if (word !== undefined) {
    const index = words.findIndex((item) => item.text === word)
    text = line.text.split(' ').map((token, position) => (position === index ? replacement : token)).join(' ')
    const target = words[index]
    scan.replace = {
      first: index,
      last: index,
      originalText: word,
      text: replacement,
      rect: frame(target.start, words[index + 1].start - target.start),
      originalRect: target.rect,
      baselineOffset: 0.3 * size,
      run: { origin: { x: origin.x + dir.x * target.start, y: origin.y + dir.y * target.start }, dir, length: target.end - target.start, fontSize: size },
      patch: patchOf(new Set(retouch.words[index].labels)),
    }
  }
  return {
    id: 'scan-edit', type: 'text', pageIndex: 0, rect: frame(0, length), originalRect: scan.lineRect, originalText: line.text, text,
    fontSize: style.fontSize, fontFamily: style.fontFamily, fontWeight: style.fontWeight, fontStyle: 'normal', textFit: 'fit', preserveSourceMetrics: false,
    lineHeight: style.fontSize * 1.18, scaleX: 1, angle: 0, displayRotation: 0, align: 'left', color: style.color, baselineOffset: 0.3 * size, cover: true, scan,
  }
}

const render = (bytes, dpi = 150) => fixture.rasterize(bytes, dpi)

/** Bounding box (PDF points) of the pixels that differ between two renders, or null. */
function differenceBounds(a, b, dpi = 150) {
  let x0 = Infinity
  let y0 = Infinity
  let x1 = -Infinity
  let y1 = -Infinity
  for (let y = 0; y < a.height; y += 1) {
    for (let x = 0; x < a.width; x += 1) {
      if (a.data[y * a.width + x] === b.data[y * b.width + x]) continue
      x0 = Math.min(x0, x)
      y0 = Math.min(y0, y)
      x1 = Math.max(x1, x + 1)
      y1 = Math.max(y1, y + 1)
    }
  }
  if (!Number.isFinite(x0)) return null
  const scale = 72 / dpi
  return { x: x0 * scale, y: 792 - y1 * scale, width: (x1 - x0) * scale, height: (y1 - y0) * scale }
}

const inside = (inner, outer, slack) => inner.x >= outer.x - slack && inner.y >= outer.y - slack
  && inner.x + inner.width <= outer.x + outer.width + slack && inner.y + inner.height <= outer.y + outer.height + slack
const union = (a, b) => {
  const x = Math.min(a.x, b.x)
  const y = Math.min(a.y, b.y)
  return { x, y, width: Math.max(a.x + a.width, b.x + b.width) - x, height: Math.max(a.y + a.height, b.y + b.height) - y }
}
const savedText = async (bytes) => (await pageTexts(bytes))[0]

test('saving a one-word edit of a scan: patch under the new word, the old word gone from text and pixels, nothing else changed', { timeout: 120_000 }, async () => {
  const { invoke } = loadMain()
  const scan = await recognisedScan('clean300')
  const overlay = await scanOverlay(scan, 'Invoice', { word: '48213', replacement: '48214' })
  const result = await invoke('pdf:flatten-overlays', scan.bytes, [overlay], {}, { report: true })
  assert.deepEqual(result.failures, [])
  const text = await savedText(result.data)
  assert.equal(text.split('48214').length - 1, 1, 'the new word is in the text once')
  assert.equal(text.includes('48213'), false, 'the old word is gone from the text')
  for (const kept of ['Invoice', 'number', 'paid', 'negotiations', '7741-0093']) assert.ok(text.includes(kept), `${kept} keeps its recognised text`)
  const changed = differenceBounds(await render(scan.bytes), await render(result.data))
  assert.ok(changed, 'the page changed')
  const allowed = union(overlay.scan.replace.patch.rect, overlay.scan.replace.rect)
  assert.ok(inside(changed, allowed, 72 / 150), `pixels changed outside the edited word: ${JSON.stringify({ changed, allowed })}`)
})

test('"recognized text only" corrects the searchable text and leaves every pixel as it was', { timeout: 120_000 }, async () => {
  const { invoke } = loadMain()
  const scan = await recognisedScan('clean300')
  const overlay = await scanOverlay(scan, 'Invoice', { word: '48213', replacement: '48214', mode: 'recognized-text' })
  const result = await invoke('pdf:flatten-overlays', scan.bytes, [overlay], {}, { report: true })
  assert.deepEqual(result.failures, [])
  const text = await savedText(result.data)
  assert.ok(text.includes('48214'))
  assert.equal(text.includes('48213'), false)
  assert.equal(differenceBounds(await render(scan.bytes), await render(result.data)), null, 'the page looks exactly the same')
})

test('deleting a scanned line erases its words from the picture and the text', { timeout: 120_000 }, async () => {
  const { invoke } = loadMain()
  const scan = await recognisedScan('clean300')
  const overlay = await scanOverlay(scan, 'Invoice', { lineText: '' })
  const result = await invoke('pdf:flatten-overlays', scan.bytes, [overlay], {}, { report: true })
  assert.deepEqual(result.failures, [])
  const text = await savedText(result.data)
  for (const gone of ['Invoice', '48213', '$1,250.75']) assert.equal(text.includes(gone), false, `${gone} is gone`)
  for (const kept of ['negotiations', '7741-0093']) assert.ok(text.includes(kept), `${kept} stays`)
  const changed = differenceBounds(await render(scan.bytes), await render(result.data))
  assert.ok(changed && inside(changed, overlay.scan.patch.rect, 72 / 150), 'only the line\'s pixels changed')
})

test('a slanted scanned line loses only its own recognised words', { timeout: 120_000 }, async () => {
  const { invoke } = loadMain()
  const scan = await recognisedScan('noisy300')
  const overlay = await scanOverlay(scan, 'Invoice', { lineText: '' })
  overlay.angle = Math.atan2(overlay.scan.run.dir.y, overlay.scan.run.dir.x)
  const result = await invoke('pdf:flatten-overlays', scan.bytes, [overlay], {}, { report: true })
  assert.deepEqual(result.failures, [])
  const text = await savedText(result.data)
  for (const gone of ['Invoice', '48213', 'remaining', '$1,250.75']) assert.equal(text.includes(gone), false, `${gone} is gone`)
  // The lines above and below run into this line's bounding box at its ends.
  for (const kept of ['immediately.', 'negotiations', 'transferred', '7741-0093', 'month.']) assert.ok(text.includes(kept), `${kept} stays`)
})

test('an edit of scanned text without its position is reported, not drawn over the old words', { timeout: 60_000 }, async () => {
  const { invoke } = loadMain()
  const scan = await recognisedScan('clean300')
  const overlay = await scanOverlay(scan, 'Invoice', { word: '48213', replacement: '48214' })
  delete overlay.scan.replace.run
  const result = await invoke('pdf:flatten-overlays', scan.bytes, [overlay], {}, { report: true })
  assert.equal(result.failures.length, 1)
  assert.equal(result.failures[0].overlayId, 'scan-edit')
  assert.equal(differenceBounds(await render(scan.bytes), await render(result.data)), null)
  assert.ok((await savedText(result.data)).includes('48213'))
})

test('without a retouch patch (its preparation failed) the old words are covered with the paper colour, never left under the new text', { timeout: 120_000 }, async () => {
  const { invoke } = loadMain()
  const scan = await recognisedScan('clean300')
  const overlay = await scanOverlay(scan, 'Invoice', { lineText: 'Invoice number 48214' })
  overlay.scan.status = 'failed'
  overlay.scan.patch = undefined
  overlay.inkRect = overlay.scan.lineRect
  const result = await invoke('pdf:flatten-overlays', scan.bytes, [overlay], {}, { report: true })
  assert.deepEqual(result.failures, [])
  const text = await savedText(result.data)
  assert.ok(text.includes('Invoice number 48214'))
  assert.equal(text.includes('48213'), false)
  // The old line is gone from the picture: nothing dark is left where "$1,250.75" was.
  const word = scan.variant.truth.words.find((item) => item.text === '$1,250.75')
  const raster = await render(result.data)
  const s = 150 / 72
  const x = Math.round(word.centre.x * s)
  const y = Math.round((792 - word.centre.y) * s)
  let darkest = 255
  for (let dy = -4; dy <= 4; dy += 1) for (let dx = -20; dx <= 20; dx += 1) darkest = Math.min(darkest, raster.data[(y + dy) * raster.width + x + dx])
  assert.ok(darkest > 200, `old ink left under the edit (darkest ${darkest})`)
})
