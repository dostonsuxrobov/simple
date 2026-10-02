'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { PDFDocument, StandardFonts, degrees } = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { textItems } = require('./helpers/pdf-test-utils.cjs')

const SIZE = 12

/** A page with one run drawn at `origin` and turned by `angle` degrees. */
async function rotatedRun(text, origin, angle, { pageRotation = 0, extra } = {}) {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const page = doc.addPage([400, 500])
  page.setRotation(degrees(pageRotation))
  page.drawText(text, { x: origin.x, y: origin.y, size: SIZE, font, rotate: degrees(angle) })
  if (extra) extra(page, font)
  return { bytes: await doc.save(), width: font.widthOfTextAtSize(text, SIZE), font }
}

/** The axis-aligned box the editor stores for a vertical run (ascent 0.8). */
function verticalRunRect(origin, width) {
  return { x: origin.x - 0.8 * SIZE, y: origin.y, width: SIZE, height: width }
}

function nativeEdit(id, rect, originalText, text, angle) {
  return {
    id, type: 'text', pageIndex: 0, rect, originalRect: rect, originalText, text, angle,
    fontSize: SIZE, fontFamily: 'Helvetica', color: [0, 0, 0], align: 'left', cover: true,
    textFit: 'fit', preserveSourceMetrics: true, displayRotation: 0,
  }
}

function runGeometry(item) {
  const [a, b] = item.transform
  return { x: item.transform[4], y: item.transform[5], scale: Math.hypot(a, b), angle: Math.atan2(b, a) * 180 / Math.PI }
}

async function savedRun(output, pattern) {
  const [items] = await textItems(output)
  const match = items.find((item) => pattern.test(item.str))
  assert.ok(match, `no saved run matches ${pattern}: ${JSON.stringify(items.map((item) => item.str))}`)
  return { items, run: runGeometry(match) }
}

for (const pageRotation of [0, 90]) {
  test(`editing a vertical run (page /Rotate ${pageRotation}) keeps its size, direction and place`, async () => {
    const { invoke } = loadMain()
    const origin = { x: 100, y: 100 }
    const { bytes, width } = await rotatedRun('VERTICAL LABEL', origin, 90, { pageRotation })
    const edit = nativeEdit('vertical', verticalRunRect(origin, width), 'VERTICAL LABEL', 'VERTICAL LABEL 2', Math.PI / 2)
    const output = await invoke('pdf:flatten-overlays', bytes, [edit], {}, {})
    const { items, run } = await savedRun(output, /VERTICAL LABEL 2/)
    assert.ok(Math.abs(run.angle - 90) < 0.5, `direction ${run.angle}`)
    // The old writer measured the run against the box's 12 pt width and
    // squeezed it to a sliver (about 13% here).
    assert.ok(run.scale > SIZE * 0.8, `glyph scale ${run.scale}`)
    assert.ok(Math.abs(run.x - origin.x) < 1.5 && Math.abs(run.y - origin.y) < 1.5, `origin ${run.x},${run.y}`)
    assert.equal(items.filter((item) => item.str === 'VERTICAL LABEL').length, 0, 'the original run was replaced')
  })
}

test('a box rotated in the inspector is saved rotated about its centre, where the editor shows it', async () => {
  const { invoke } = loadMain()
  const doc = await PDFDocument.create()
  doc.addPage([400, 500])
  const rect = { x: 150, y: 300, width: 200, height: 30 }
  const overlay = {
    id: 'turned', type: 'text', pageIndex: 0, rect, text: 'Hello', angle: Math.PI / 2,
    fontSize: 16, fontFamily: 'Segoe UI', color: [0, 0, 0], align: 'left', cover: false, textFit: 'wrap',
  }
  const output = await invoke('pdf:flatten-overlays', await doc.save(), [overlay], {}, {})
  const { run } = await savedRun(output, /Hello/)
  const centre = { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 }
  assert.ok(Math.abs(run.angle - 90) < 0.5)
  // Reading frame 200 × 30 centred on the box: the baseline starts at the
  // frame's start edge (centre − 100 along the text) and 16 pt below its top.
  assert.ok(Math.abs(run.x - (centre.x + 16 - 15)) < 1, `x ${run.x}`)
  assert.ok(Math.abs(run.y - (centre.y - 100)) < 1, `y ${run.y}`)
})

test('editing a diagonal watermark does not erase horizontal text inside its bounding box', async () => {
  const { invoke } = loadMain()
  const origin = { x: 100, y: 250 }
  const doc = await PDFDocument.create()
  const helvetica = await doc.embedFont(StandardFonts.Helvetica)
  const page = doc.addPage([400, 500])
  page.drawText('Body line inside', { x: 120, y: 300, size: 12, font: helvetica })
  page.drawText('DRAFT COPY', { x: origin.x, y: origin.y, size: 40, font: helvetica, rotate: degrees(45), opacity: 0.4 })
  const input = await doc.save()
  const width = helvetica.widthOfTextAtSize('DRAFT COPY', 40)
  const c = Math.SQRT1_2
  const corners = [[0, -8], [width, -8], [width, 32], [0, 32]].map(([u, v]) => [origin.x + u * c - v * c, origin.y + u * c + v * c])
  const xs = corners.map(([x]) => x)
  const ys = corners.map(([, y]) => y)
  const rect = { x: Math.min(...xs), y: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) }
  assert.ok(rect.x < 120 && rect.y < 300 && rect.x + rect.width > 200 && rect.y + rect.height > 310, 'the body line is inside the box')
  const edit = { ...nativeEdit('stamp', rect, 'DRAFT COPY', 'FINAL', Math.PI / 4), fontSize: 40 }
  const output = await invoke('pdf:flatten-overlays', input, [edit], {}, {})
  const [items] = await textItems(output)
  const strings = items.map((item) => item.str)
  assert.ok(strings.includes('Body line inside'), `body text kept: ${JSON.stringify(strings)}`)
  assert.ok(strings.some((value) => /FINAL/.test(value)))
  assert.equal(strings.filter((value) => value !== 'Body line inside').join('').replace(/FINAL|\s/g, ''), '', `old watermark letters removed: ${JSON.stringify(strings)}`)
})
