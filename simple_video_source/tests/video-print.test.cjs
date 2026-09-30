'use strict'

const assert = require('node:assert/strict')
const test = require('node:test')
const {
  assertPrintersInstalled,
  createVideoPrintDocument,
  ensurePrinterAvailable,
  nativePrintOptions,
  pngDimensions,
  runSilentPrintJob,
  safePrintOptions,
} = require('../electron/video-print.cjs')

function pngHeader(width, height) {
  const bytes = Buffer.alloc(32)
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]).copy(bytes)
  bytes.writeUInt32BE(13, 8)
  bytes.write('IHDR', 12, 'ascii')
  bytes.writeUInt32BE(width, 16)
  bytes.writeUInt32BE(height, 20)
  return bytes
}

function request(options = {}) {
  return {
    bytes: pngHeader(1600, 900),
    sourceName: 'Demo <frame>.webm',
    seconds: 65.8,
    options: {
      paperSize: 'letter',
      orientation: 'landscape',
      margins: 'normal',
      scaleMode: 'fit',
      customScale: 100,
      colorMode: 'color',
      metadata: true,
      ...options,
    },
  }
}

test('fit creates one exact Letter landscape page without cropping', () => {
  const result = createVideoPrintDocument(request())
  assert.equal(result.page.label, 'Letter')
  assert.equal(result.page.widthInches, 11)
  assert.equal(result.page.heightInches, 8.5)
  assert.equal(result.frame.timestamp, '1:05')
  assert.equal(result.placement.scale, 0.6)
  assert.equal(result.placement.cropped, false)
  assert.match(result.html, /@page \{ size: Letter landscape; margin: 0; \}/)
  assert.match(result.html, /class="print-page"/)
  assert.match(result.html, /data-scale-mode="fit"/)
  assert.match(result.html, /Frame at 1:05 · 1600 × 900px/)
  assert.match(result.html, /Demo &lt;frame&gt;\.webm/)
  assert.doesNotMatch(result.html, /Demo <frame>/)
})

test('fill, actual, and custom sizing report clipping honestly', () => {
  const fill = createVideoPrintDocument(request({ scaleMode: 'fill' }))
  assert.equal(fill.placement.cropped, true)
  assert.ok(fill.placement.renderedWidth > fill.placement.availableWidth)

  const actual = createVideoPrintDocument(request({ orientation: 'portrait', margins: 'narrow', scaleMode: 'actual', metadata: false }))
  assert.equal(actual.placement.scale, 1)
  assert.equal(actual.placement.cropped, true)

  const custom = createVideoPrintDocument(request({ scaleMode: 'custom', customScale: 50, colorMode: 'grayscale', metadata: false }))
  assert.equal(custom.placement.scale, 0.5)
  assert.equal(custom.options.colorMode, 'grayscale')
  assert.match(custom.html, /filter:grayscale\(1\)/)
  assert.match(custom.html, /data-metadata="false"/)
  assert.doesNotMatch(custom.html, /class="frame-metadata"/)
})

test('print options are allowlisted and custom scaling is bounded', () => {
  assert.deepEqual(safePrintOptions({ paperSize: 'executive', orientation: 'sideways', margins: 'zero', scaleMode: 'stretch', customScale: 9999, colorMode: 'sepia', metadata: 'yes' }), {
    paperSize: 'letter',
    orientation: 'portrait',
    margins: 'normal',
    scaleMode: 'fit',
    customScale: 400,
    colorMode: 'color',
    metadata: false,
  })
})

test('native printing is silent, lets Electron select the system default, and preserves the page model', () => {
  const document = createVideoPrintDocument(request({ paperSize: 'a4', orientation: 'portrait', colorMode: 'grayscale' }))
  assert.deepEqual(nativePrintOptions(document), {
    silent: true,
    printBackground: true,
    color: false,
    landscape: false,
    margins: { marginType: 'none' },
    pageSize: 'A4',
    scaleFactor: 100,
    pagesPerSheet: 1,
    collate: true,
  })
  assert.equal(Object.hasOwn(nativePrintOptions(document), 'deviceName'), false)
})

test('printer preflight explains an empty or unreadable Windows printer list', async () => {
  const printers = [
    { name: 'First_System_Name', displayName: 'Office printer' },
    { name: 'Second_System_Name', displayName: 'Home printer' },
  ]
  assert.equal(assertPrintersInstalled(printers), 2)
  assert.equal(await ensurePrinterAvailable({ getPrintersAsync: async () => printers }), 2)
  assert.throws(() => assertPrintersInstalled([]), /No printers are installed/i)
  await assert.rejects(() => ensurePrinterAvailable({ getPrintersAsync: async () => { throw new Error('spooler') } }), /could not read the installed printers/i)
})

test('silent print helper passes silent:true to webContents.print and reports driver failures', async () => {
  const document = createVideoPrintDocument(request())
  let receivedOptions = null
  const result = await runSilentPrintJob({
    print(options, callback) {
      receivedOptions = options
      callback(true, '')
    },
  }, document)
  assert.equal(result.printed, true)
  assert.equal(receivedOptions.silent, true)
  assert.equal(Object.hasOwn(receivedOptions, 'deviceName'), false)

  await assert.rejects(() => runSilentPrintJob({
    print(_options, callback) { callback(false, 'Printer is offline') },
  }, document), /default printer did not accept the print job: Printer is offline/)
})

test('invalid or unsafe PNG dimensions are rejected', () => {
  assert.throws(() => pngDimensions(Buffer.from('not a png')), /valid PNG/i)
  assert.throws(() => pngDimensions(pngHeader(0, 900)), /dimensions/i)
  assert.throws(() => pngDimensions(pngHeader(16_385, 1)), /dimensions/i)
  assert.throws(() => createVideoPrintDocument(null), /invalid/i)
})

test('silent print releases a missing callback and crashed renderer without retrying', async () => {
  const {EventEmitter}=require('node:events')
  const renderer=new EventEmitter()
  let submissions=0
  renderer.print=()=>{submissions++}
  await assert.rejects(runSilentPrintJob(renderer,createVideoPrintDocument(request()),{timeoutMs:5}),/Check the print queue.*duplicate/)
  assert.equal(submissions,1)
  assert.equal(renderer.listenerCount('destroyed'),0)
  renderer.print=()=>queueMicrotask(()=>renderer.emit('render-process-gone'))
  await assert.rejects(runSilentPrintJob(renderer,createVideoPrintDocument(request())),/renderer stopped/)
  assert.equal(renderer.listenerCount('render-process-gone'),0)
})
