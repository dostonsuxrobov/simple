'use strict'

// Regression smoke for two Advanced basics, in the built renderer (npm run build:web first):
//   1. an opaque image opens as an ordinary layer: opacity and "lock transparent pixels"
//      are available, and Delete on a selection leaves transparent pixels (not white);
//   2. dropping image files adds each one as a new layer and keeps the document open;
//   3. PDFs: one opened in the viewer is shown as page 1 (with Open in PDF), and one dropped on
//      the Advanced editor adds one layer per page. Next/Previous stay with images.
// Run: electron qa/advanced-drop-smoke.cjs   (prints one JSON line; exit 1 on failure)

const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const { app, BrowserWindow, ipcMain } = require('electron')

const projectRoot = path.join(__dirname, '..')
const workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-image-drop-'))
app.setPath('userData', path.join(workDir, 'profile'))

const W = 320
const H = 200

function crc32(buffer) {
  let crc = ~0
  for (const byte of buffer) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1))
  }
  return ~crc >>> 0
}

function chunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body))
  return Buffer.concat([length, body, crc])
}

function encodePng(width, height, pixel) {
  const raw = Buffer.alloc((width * 4 + 1) * height)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) raw.set(pixel(x, y), y * (width * 4 + 1) + 1 + x * 4)
  }
  const header = Buffer.alloc(13)
  header.writeUInt32BE(width, 0)
  header.writeUInt32BE(height, 4)
  header[8] = 8
  header[9] = 6
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', zlib.deflateSync(raw)), chunk('IEND', Buffer.alloc(0))])
}

const fixtureBytes = encodePng(W, H, (x) => [x % 256, 120, 60, 255])
const fixturePath = path.join(workDir, 'opaque.png')
fs.writeFileSync(fixturePath, fixtureBytes)
const dropA = encodePng(40, 30, () => [255, 0, 0, 255]).toString('base64')
const dropB = encodePng(20, 20, () => [0, 0, 255, 128]).toString('base64')

ipcMain.handle('file:open-path', () => ({ data: new Uint8Array(fixtureBytes), name: 'opaque.png', path: fixturePath, size: fixtureBytes.length, format: 'png', mime: 'image/png', directSave: true }))
ipcMain.handle('file:open-dialog', () => null)
const pdfPath = path.join(workDir, 'two pages.pdf')
let pdfBytes = null
ipcMain.handle('file:list-siblings', () => [fixturePath])
ipcMain.handle('file:read-pdf', () => ({ data: new Uint8Array(pdfBytes), name: 'two pages.pdf', path: pdfPath, size: pdfBytes.length }))
ipcMain.handle('clipboard:read-image', () => null)
ipcMain.handle('clipboard:has-image', () => false)
ipcMain.handle('clipboard:write-png', () => ({ width: 1, height: 1 }))
ipcMain.handle('app:get-version', () => '0.0.0-smoke')
for (const channel of ['window:set-title', 'window:minimize', 'window:toggle-maximize', 'window:confirm-close']) ipcMain.on(channel, () => {})

const result = { checks: [], failures: [], consoleErrors: [] }
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
function check(condition, message) {
  result.checks.push(message)
  if (!condition) result.failures.push(message)
}

app.whenReady().then(async () => {
  // A two-page PDF: page 1 red, page 2 blue (200 x 100 pt).
  {
    const { PDFDocument, rgb } = require('pdf-lib')
    const doc = await PDFDocument.create()
    for (const colour of [rgb(1, 0, 0), rgb(0, 0, 1)]) doc.addPage([200, 100]).drawRectangle({ x: 0, y: 0, width: 200, height: 100, color: colour })
    pdfBytes = Buffer.from(await doc.save())
    fs.writeFileSync(pdfPath, pdfBytes)
  }
  const window = new BrowserWindow({
    width: 1365, height: 860, show: false,
    webPreferences: { preload: path.join(projectRoot, 'electron', 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true },
  })
  window.webContents.on('console-message', (event) => {
    const level = event.level ?? event.params?.level
    if (level === 'error' || level === 3) result.consoleErrors.push(String(event.message ?? event.params?.message).slice(0, 300))
  })
  const js = (expression) => window.webContents.executeJavaScript(expression)
  async function waitFor(expression, label, timeout = 15_000) {
    const started = Date.now()
    let last
    while (Date.now() - started < timeout) {
      try { last = await js(expression); if (last) return last } catch (error) { last = String(error) }
      await sleep(70)
    }
    throw new Error(`Timed out waiting for ${label} (last: ${JSON.stringify(last)})`)
  }
  const layers = () => js(`window.__simpleAdvanced.store.getState().layers.map((layer) => ({ name: layer.name, background: Boolean(layer.isBackground), kind: layer.kind }))`)
  const sample = (x, y) => js(`(() => { const c = window.__simpleAdvanced.compositor.sample(${x}, ${y}, 1, 'all'); return [c.r, c.g, c.b, c.a] })()`)
  const settle = () => js(`window.__simpleAdvanced.compositor.settle().then(() => true)`)

  try {
    await window.loadFile(path.join(projectRoot, 'dist', 'index.html'))
    await sleep(200)
    window.webContents.focus()
    await js('window.__SIMPLE_IMAGE_QA__ = true; true')
    window.webContents.send('file:open-external', fixturePath)
    await waitFor(`document.querySelector('canvas')?.width === ${W} && document.querySelector('.advanced-button:not(:disabled)') !== null`, 'Simple shows the image')
    // 3a. The viewer: a PDF handed to the window is shown as a picture of its first page.
    window.webContents.send('file:open-external', pdfPath)
    await waitFor(`document.querySelector('.document-title')?.textContent === 'two pages.pdf' && document.querySelector('canvas')?.width > 300`, 'the PDF shown as a picture', 30_000)
    const preview = await js(`(() => { const c = document.querySelector('canvas'); const p = c.getContext('2d').getImageData(Math.floor(c.width / 2), Math.floor(c.height / 2), 1, 1).data; return { width: c.width, height: c.height, pixel: Array.from(p), state: document.querySelector('.save-state')?.textContent } })()`)
    check(preview.width > 300 && Math.abs(preview.width / preview.height - 2) < 0.05, `page 1 at reading size (${preview.width} x ${preview.height})`)
    check(preview.pixel[0] > 200 && preview.pixel[2] < 60, `page 1 is the red page (${preview.pixel})`)
    check(preview.state === 'Saved', `a PDF preview is not an unsaved picture (${preview.state})`)
    await js(`document.querySelector('.more-button').click(); true`)
    check(await js(`Array.from(document.querySelectorAll('.more-menu button')).some((b) => b.textContent.includes('Open in PDF'))`), 'More offers Open in PDF')
    await js(`document.querySelector('.more-button').click(); true`)
    window.webContents.send('file:open-external', fixturePath)
    await waitFor(`document.querySelector('.document-title')?.textContent === 'opaque.png' && document.querySelector('canvas')?.width === ${W}`, 'back on the picture')
    check(!(await js(`Array.from(document.querySelectorAll('.more-menu button')).some((b) => b.textContent.includes('Open in PDF'))`)), 'Open in PDF is gone for an ordinary picture')

    await js(`document.querySelector('.advanced-button').click(); true`)
    await waitFor(`document.querySelector('.ae-shell .ae-layer-row') !== null && Boolean(window.__simpleAdvanced)`, 'the Advanced editor')
    await settle()

    // 1. The opaque image is an ordinary layer.
    const opened = await layers()
    check(opened.length === 1 && !opened[0].background, `an opaque image opens as an ordinary layer (${JSON.stringify(opened)})`)
    const disabled = await js(`Array.from(document.querySelectorAll('.ae-layers-panel input, .ae-layers-panel button, .ae-panel input, .ae-panel button')).filter((el) => el.disabled && /opacity|transparent/i.test((el.getAttribute('aria-label') || '') + ' ' + (el.title || ''))).map((el) => el.getAttribute('aria-label') || el.title)`)
    check(disabled.length === 0, `opacity and transparency lock are enabled (disabled: ${JSON.stringify(disabled)})`)

    // Select all, Delete: the pixels become transparent, not white.
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'a', code: 'KeyA', ctrlKey: true, bubbles: true, cancelable: true })); true`)
    await waitFor(`window.__simpleAdvanced.store.getState().selection !== null`, 'a selection')
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Delete', code: 'Delete', bubbles: true, cancelable: true })); true`)
    await settle()
    const cleared = await sample(160, 100)
    check(cleared[3] === 0, `Delete on a selection leaves transparency (${cleared})`)
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', code: 'KeyZ', ctrlKey: true, bubbles: true, cancelable: true })); true`)
    await settle()
    const restored = await sample(160, 100)
    check(restored[3] === 255, `undo restores the pixels (${restored})`)
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'd', code: 'KeyD', ctrlKey: true, bubbles: true, cancelable: true })); true`)

    // 2. Dropping two files adds two layers; the document stays open in Advanced.
    await js(`(() => {
      const bytes = (base64) => Uint8Array.from(atob(base64), (c) => c.charCodeAt(0))
      const transfer = new DataTransfer()
      transfer.items.add(new File([bytes(${JSON.stringify(dropA)})], 'red patch.png', { type: 'image/png' }))
      transfer.items.add(new File([bytes(${JSON.stringify(dropB)})], 'blue glass.png', { type: 'image/png' }))
      const target = document.querySelector('main')
      for (const type of ['dragenter', 'dragover', 'drop']) target.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: transfer }))
      return true
    })()`)
    await waitFor(`window.__simpleAdvanced && window.__simpleAdvanced.store.getState().layers.length === 3`, 'three layers after the drop')
    await settle()
    const after = await layers()
    check(JSON.stringify(after.map((layer) => layer.name)) === JSON.stringify(['Background', 'red patch 1', 'blue glass 1']), `dropped files become named layers (${JSON.stringify(after.map((layer) => layer.name))})`)
    check(await js(`document.querySelector('.ae-shell') !== null`), 'the editor stays in Advanced')
    const state = await js(`(() => { const s = window.__simpleAdvanced.store.getState(); return { width: s.width, height: s.height, active: s.layers.find((l) => l.id === s.activeLayerId)?.name } })()`)
    check(state.width === W && state.height === H, `the document keeps its size (${state.width} x ${state.height})`)
    check(state.active === 'blue glass 1', `the last dropped layer is active (${state.active})`)
    const centre = await sample(160, 100)
    check(centre[2] > centre[0] && centre[3] === 255, `the dropped layers composite over the image (${centre})`)
    // One undo step per dropped layer.
    await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'z', code: 'KeyZ', ctrlKey: true, bubbles: true, cancelable: true })); true`)
    await waitFor(`window.__simpleAdvanced.store.getState().layers.length === 2`, 'undo removes one dropped layer')
    check(true, 'undo removes one dropped layer')

    // 3b. A PDF dropped on the Advanced editor: one layer per page.
    const before = (await layers()).length
    await js(`(() => {
      const bytes = Uint8Array.from(atob(${JSON.stringify(pdfBytes.toString('base64'))}), (c) => c.charCodeAt(0))
      const transfer = new DataTransfer()
      transfer.items.add(new File([bytes], 'two pages.pdf', { type: 'application/pdf' }))
      const target = document.querySelector('main')
      for (const type of ['dragenter', 'dragover', 'drop']) target.dispatchEvent(new DragEvent(type, { bubbles: true, cancelable: true, dataTransfer: transfer }))
      return true
    })()`)
    await waitFor(`window.__simpleAdvanced.store.getState().layers.length === ${before + 2}`, 'two layers from the two-page PDF', 30_000)
    await settle()
    const pdfLayers = (await layers()).slice(-2).map((layer) => layer.name)
    check(pdfLayers[0].startsWith('two pages p1') && pdfLayers[1].startsWith('two pages p2'), `PDF pages become named layers (${JSON.stringify(pdfLayers)})`)
    const top = await sample(160, 100)
    check(top[2] > 200 && top[0] < 60 && top[3] === 255, `page 2 (blue) is on top (${top})`)
    check(await js(`document.querySelector('.ae-export') !== null`), 'the Advanced editor has its own Export button')
  } catch (error) {
    result.failures.push(String(error && error.message || error))
  }
  result.ok = result.failures.length === 0
  process.stdout.write(`${JSON.stringify(result)}\n`)
  app.exit(result.ok ? 0 : 1)
})
