const test = require('node:test')
const assert = require('node:assert/strict')

async function printLayout() {
  return import('../electron/print-layout.mjs')
}

test('fit uses the printable paper area and preserves image aspect ratio', async () => {
  const { computePrintLayout } = await printLayout()
  const layout = computePrintLayout(1600, 900, {
    paper: 'letter',
    orientation: 'portrait',
    marginMm: 12.7,
    scaleMode: 'fit',
  })
  assert.equal(layout.paper.widthMm, 215.9)
  assert.equal(layout.paper.heightMm, 279.4)
  assert.ok(Math.abs(layout.printable.widthMm - 190.5) < 0.0001)
  assert.ok(Math.abs(layout.printable.heightMm - 254) < 0.0001)
  assert.ok(Math.abs(layout.image.widthMm - 190.5) < 0.0001)
  assert.ok(Math.abs(layout.image.heightMm - 107.15625) < 0.0001)
  assert.ok(Math.abs(layout.image.yMm - 73.421875) < 0.0001)
  assert.equal(layout.clipped, false)
})

test('paper orientation, fill crop, positioning, and actual-size scaling are deterministic', async () => {
  const { computePrintLayout } = await printLayout()
  const fill = computePrintLayout(1000, 1000, {
    paper: 'a4', orientation: 'landscape', marginMm: 0, scaleMode: 'fill', position: 'bottom-right',
  })
  assert.deepEqual([fill.paper.widthMm, fill.paper.heightMm], [297, 210])
  assert.equal(fill.image.widthMm, 297)
  assert.equal(fill.image.heightMm, 297)
  assert.equal(fill.image.xMm, 0)
  assert.equal(fill.image.yMm, -87)
  assert.equal(fill.clipped, true)

  const actual = computePrintLayout(960, 480, { scaleMode: 'actual', marginMm: 0 })
  assert.ok(Math.abs(actual.image.widthMm - 254) < 0.0001)
  assert.ok(Math.abs(actual.image.heightMm - 127) < 0.0001)
  assert.ok(Math.abs(actual.effectiveDpi - 96) < 0.0001)

  const custom = computePrintLayout(960, 480, { scaleMode: 'custom', scalePercent: 50, marginMm: 0, position: 'top-left' })
  assert.ok(Math.abs(custom.image.widthMm - 127) < 0.0001)
  assert.ok(Math.abs(custom.image.heightMm - 63.5) < 0.0001)
  assert.equal(custom.image.xMm, 0)
  assert.equal(custom.image.yMm, 0)
})

test('settings normalization contains malformed renderer input', async () => {
  const { computePrintLayout, normalizePrintSettings } = await printLayout()
  assert.deepEqual(normalizePrintSettings({
    paper: 'poster', orientation: 'sideways', marginMm: 999, scaleMode: 'stretch', scalePercent: -20,
    position: 'outside', background: 'url(file:///secret)', grayscale: 'yes', copies: 10000,
  }), {
    paper: 'letter', orientation: 'portrait', marginMm: 50, scaleMode: 'fit', scalePercent: 10,
    position: 'center', background: '#ffffff', grayscale: false, copies: 99,
  })
  assert.throws(() => computePrintLayout(0, 100), /dimensions are invalid/)
})

test('print HTML carries the same physical layout, appearance, and one-page contract', async () => {
  const { buildPrintHtml, electronPrintOptions } = await printLayout()
  const html = buildPrintHtml('file:///C:/Temp/image.png', 1000, 500, {
    paper: 'a4', orientation: 'landscape', marginMm: 10, scaleMode: 'custom', scalePercent: 50,
    position: 'top-right', background: '#112233', grayscale: true, copies: 2,
  }, 'A <safe> image')
  assert.match(html, /@page \{ size: A4 landscape; margin: 0; \}/)
  assert.match(html, /width: 297mm; height: 210mm/)
  assert.match(html, /left: 10mm; top: 10mm; width: 277mm; height: 190mm/)
  assert.match(html, /\.image-frame \{[^}]*background: #112233; filter: grayscale\(1\)/)
  assert.doesNotMatch(html, /img \{[^}]*filter:/)
  assert.match(html, /<title>A &lt;safe&gt; image<\/title>/)
  assert.match(html, /page-break-after: avoid/)
  assert.deepEqual(electronPrintOptions({ paper: 'a4', orientation: 'landscape', copies: 3 }), {
    silent: true,
    printBackground: true,
    color: true,
    landscape: true,
    scaleFactor: 100,
    copies: 3,
    collate: true,
    margins: { marginType: 'none' },
    pageSize: 'A4',
  })
  assert.equal(electronPrintOptions({ grayscale: true }).color, false)
  assert.equal(electronPrintOptions({ grayscale: false }).color, true)
  assert.equal(Object.hasOwn(electronPrintOptions({}), 'deviceName'), false)
  assert.throws(() => buildPrintHtml('https://example.com/tracker.png', 10, 10), /data is invalid/)
})
