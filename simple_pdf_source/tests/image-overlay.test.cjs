'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const { createCanvas } = require('@napi-rs/canvas')
const { PDFDocument } = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { pdfjs } = require('./helpers/pdf-test-utils.cjs')

function withOrientation(bytes, orientation) {
  const metadata = Buffer.alloc(32)
  metadata.write('Exif\0\0', 0, 'ascii')
  metadata.write('II', 6, 'ascii')
  metadata.writeUInt16LE(42, 8)
  metadata.writeUInt32LE(8, 10)
  metadata.writeUInt16LE(1, 14)
  metadata.writeUInt16LE(0x112, 16)
  metadata.writeUInt16LE(3, 18)
  metadata.writeUInt32LE(1, 20)
  metadata.writeUInt16LE(orientation, 24)
  return Buffer.concat([bytes.subarray(0, 2), Buffer.from([0xff, 0xe1, 0, 34]), metadata, bytes.subarray(2)])
}

/** 80×40 JPEG: red, green over blue, yellow quadrants. */
function quadrants() {
  const canvas = createCanvas(80, 40)
  const context = canvas.getContext('2d')
  for (const [color, x, y] of [['#ff0000', 0, 0], ['#00ff00', 40, 0], ['#0000ff', 0, 20], ['#ffff00', 40, 20]]) {
    context.fillStyle = color
    context.fillRect(x, y, 40, 20)
  }
  return canvas.toBuffer('image/jpeg', 100)
}

function imageOverlay(dataUrl, rect) {
  return { id: 'photo', type: 'object', kind: 'image', pageIndex: 0, rect, dataUrl, opacity: 1, cover: false }
}

async function pageWithSize(width, height) {
  const doc = await PDFDocument.create()
  doc.addPage([width, height])
  return doc.save()
}

async function cornerColors(bytes) {
  const { getDocument } = await pdfjs()
  const pdf = await getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise
  try {
    const page = await pdf.getPage(1)
    const viewport = page.getViewport({ scale: 2 })
    const canvas = createCanvas(viewport.width, viewport.height)
    const context = canvas.getContext('2d')
    await page.render({ canvasContext: context, viewport }).promise
    return [[0.25, 0.25], [0.75, 0.25], [0.25, 0.75], [0.75, 0.75]].map(([x, y]) => {
      const [r, g, b] = context.getImageData(Math.floor(x * canvas.width), Math.floor(y * canvas.height), 1, 1).data
      return r > 150 && g > 150 ? 'y' : r > 150 ? 'r' : g > 150 ? 'g' : b > 150 ? 'b' : '?'
    })
  } finally {
    await pdf.destroy()
  }
}

test('a small JPEG decoded from Node’s shared buffer pool can be saved', async () => {
  const { invoke } = loadMain()
  const jpeg = quadrants()
  assert.ok(jpeg.length < 4096, 'small enough to come from the pool')
  // Move the pool offset away from zero, as in a long-running main process.
  for (let index = 0; index < 8; index += 1) Buffer.from('c2ltcGxl', 'base64')
  const output = await invoke('pdf:flatten-overlays', await pageWithSize(80, 40),
    [imageOverlay(`data:image/jpeg;base64,${jpeg.toString('base64')}`, { x: 0, y: 0, width: 80, height: 40 })], {}, {})
  assert.deepEqual(await cornerColors(output), ['r', 'g', 'b', 'y'])
})

test('a camera JPEG is saved the way it is shown: turned by its EXIF orientation, not stretched sideways', async () => {
  const { invoke } = loadMain()
  // Chromium shows orientation 6 as 40×80 and the editor sizes the box so.
  const expected = { 3: ['y', 'b', 'g', 'r'], 6: ['b', 'r', 'y', 'g'], 8: ['g', 'y', 'r', 'b'] }
  for (const [orientation, colors] of Object.entries(expected)) {
    const rotated = Number(orientation) >= 5
    const [width, height] = rotated ? [40, 80] : [80, 40]
    const jpeg = withOrientation(quadrants(), Number(orientation))
    const output = await invoke('pdf:flatten-overlays', await pageWithSize(width, height),
      [imageOverlay(`data:image/jpeg;base64,${jpeg.toString('base64')}`, { x: 0, y: 0, width, height })], {}, {})
    assert.deepEqual(await cornerColors(output), colors, `EXIF orientation ${orientation}`)
  }
})
