'use strict'
// WP8 smoke: Simple-mode upgrades and fixes, on the REAL renderer, preload and electron/main.cjs (only the
// native save dialog, the clipboard image and the final close confirmation are replaced). Hidden window,
// isolated profile, fixtures in a temporary folder that is removed afterwards.
//
// Covers design 9 / WP8 acceptance: resize 50% (exact size, no dark fringe), flip twice, straighten 10 deg
// with auto-crop, aspect presets, Adjust (preview, one undo step), Auto on a 64..191 ramp, a Look, markup
// baking (text, arrow, rectangle, ellipse, line), eyedropper, Ctrl+V into a new document, zoom at the
// cursor, Space-drag, pixelated >= 200%, Enter on the focused Crop button, 50 MP undo depth, lossless WebP;
// and the IMAGE-SIE regressions: 1 (stroke + Ctrl+Z / Ctrl+S, pristine flag), 2 (pending crop prompt, close
// prompt, title marker), 3 (animated source saves a still PNG copy, source untouched), 14 (lossless WebP,
// JPEG transparency warning).
//
// Run: npx electron qa/simple-plus-smoke.cjs   (SIMPLE_IMAGE_PLUS_ONLY=name1,name2 runs a subset)
const { app, BrowserWindow, clipboard, dialog, ipcMain, nativeImage } = require('electron')
const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const fsSync = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')

const root = path.resolve(__dirname, '..')
const directory = fsSync.mkdtempSync(path.join(os.tmpdir(), 'simple-image-plus-'))
app.setPath('userData', path.join(directory, 'profile'))

// --- native replacements -------------------------------------------------------------------------
const saveDialogs = []
let nextSavePath = null
dialog.showSaveDialog = async (...args) => {
  const options = args[args.length - 1] || {}
  saveDialogs.push(options)
  if (!nextSavePath) return { canceled: true, filePath: undefined }
  const filePath = nextSavePath
  nextSavePath = null
  return { canceled: false, filePath }
}
let clipboardPng = null
clipboard.readImage = () => (clipboardPng ? nativeImage.createFromBuffer(clipboardPng) : nativeImage.createEmpty())
clipboard.availableFormats = () => (clipboardPng ? ['image/png'] : [])
clipboard.writeImage = () => {}
app.on('browser-window-created', (_event, window) => {
  // Keep every window hidden.
  window.show = () => {}
  window.showInactive = () => {}
  window.focus = () => {}
})

require('../electron/main.cjs')
app.removeAllListeners('window-all-closed')
let confirmCloseCount = 0
/** main.cjs registers its IPC inside its own whenReady callback; replace the close confirmation after it. */
function stubConfirmClose() {
  ipcMain.removeAllListeners('window:confirm-close')
  ipcMain.on('window:confirm-close', () => { confirmCloseCount += 1 })
}

const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
const sha = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex')
const only = new Set(String(process.env.SIMPLE_IMAGE_PLUS_ONLY || '').split(',').map((entry) => entry.trim()).filter(Boolean))

// --- PNG / APNG writer for fixtures --------------------------------------------------------------
const CRC_TABLE = (() => {
  const table = new Uint32Array(256)
  for (let n = 0; n < 256; n += 1) {
    let c = n
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
    table[n] = c >>> 0
  }
  return table
})()
function crc32(buffer) {
  let crc = 0xffffffff
  for (let index = 0; index < buffer.length; index += 1) crc = CRC_TABLE[(crc ^ buffer[index]) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}
function chunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const body = Buffer.concat([Buffer.from(type, 'latin1'), data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body))
  return Buffer.concat([length, body, crc])
}
function rawRows(width, height, pixel) {
  const rows = Buffer.alloc((width * 4 + 1) * height)
  for (let y = 0; y < height; y += 1) {
    rows[y * (width * 4 + 1)] = 0
    for (let x = 0; x < width; x += 1) {
      const [r, g, b, a] = pixel(x, y)
      rows.set([r, g, b, a], y * (width * 4 + 1) + 1 + x * 4)
    }
  }
  return zlib.deflateSync(rows)
}
function ihdr(width, height) {
  const data = Buffer.alloc(13)
  data.writeUInt32BE(width, 0)
  data.writeUInt32BE(height, 4)
  data[8] = 8
  data[9] = 6
  return chunk('IHDR', data)
}
function png(width, height, pixel) {
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), ihdr(width, height), chunk('IDAT', rawRows(width, height, pixel)), chunk('IEND', Buffer.alloc(0))])
}
function apng(width, height, frames) {
  const parts = [Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), ihdr(width, height)]
  const actl = Buffer.alloc(8)
  actl.writeUInt32BE(frames.length, 0)
  parts.push(chunk('acTL', actl))
  let sequence = 0
  frames.forEach((pixel, index) => {
    const fctl = Buffer.alloc(26)
    fctl.writeUInt32BE(sequence++, 0)
    fctl.writeUInt32BE(width, 4)
    fctl.writeUInt32BE(height, 8)
    fctl.writeUInt16BE(1, 20)
    fctl.writeUInt16BE(10, 22)
    parts.push(chunk('fcTL', fctl))
    const data = rawRows(width, height, pixel)
    if (index === 0) parts.push(chunk('IDAT', data))
    else {
      const seq = Buffer.alloc(4)
      seq.writeUInt32BE(sequence++)
      parts.push(chunk('fdAT', Buffer.concat([seq, data])))
    }
  })
  parts.push(chunk('IEND', Buffer.alloc(0)))
  return Buffer.concat(parts)
}
function pngChunkTypes(bytes) {
  const types = []
  let offset = 8
  while (offset + 8 <= bytes.length) {
    const length = bytes.readUInt32BE(offset)
    types.push(bytes.subarray(offset + 4, offset + 8).toString('latin1'))
    offset += 12 + length
  }
  return types
}
function webpChunkTypes(bytes) {
  const types = []
  let offset = 12
  while (offset + 8 <= bytes.length) {
    const size = bytes.readUInt32LE(offset + 4)
    types.push(bytes.subarray(offset, offset + 4).toString('latin1'))
    offset += 8 + size + (size & 1)
  }
  return types
}

app.whenReady().then(async () => {
  let win
  let exitCode = 0
  const results = {}
  try {
    while (!(win = BrowserWindow.getAllWindows()[0])) await pause(10)
    stubConfirmClose()
    win.webContents.setBackgroundThrottling(false)
    win.setSize(1365, 860)
    const evaluate = (code) => win.webContents.executeJavaScript(code).catch((error) => { throw new Error(`${error.message}\n${code.slice(0, 300)}`) })
    const until = async (code, label, timeout = 20_000) => {
      const end = Date.now() + timeout
      while (!(await evaluate(code))) {
        if (Date.now() > end) throw new Error(`Timed out: ${label}. Toast: ${await evaluate(`document.querySelector('.toast')?.textContent ?? ''`)}`)
        await pause(25)
      }
    }
    const settle = (ms = 120) => pause(ms)
    await until('Boolean(window.simpleImage && document.querySelector("canvas"))', 'renderer')

    // ----- helpers -------------------------------------------------------------------------------
    const writeFixture = async (name, base64OrBuffer) => {
      const file = path.join(directory, name)
      await fs.writeFile(file, Buffer.isBuffer(base64OrBuffer) ? base64OrBuffer : Buffer.from(base64OrBuffer, 'base64'))
      return file
    }
    const canvasFixture = (width, height, draw, mime = 'image/png', quality = 1) => evaluate(`(async () => {
      const c = document.createElement('canvas'); c.width = ${width}; c.height = ${height}
      const x = c.getContext('2d'); (${draw})(x, c)
      const blob = await new Promise((resolve) => c.toBlob(resolve, ${JSON.stringify(mime)}, ${quality}))
      c.width = 1; c.height = 1
      const bytes = new Uint8Array(await blob.arrayBuffer()); let text = ''
      for (let i = 0; i < bytes.length; i += 0x8000) text += String.fromCharCode(...bytes.subarray(i, i + 0x8000))
      return btoa(text)
    })()`)
    const dismissPrompt = async () => {
      if (await evaluate(`Boolean([...document.querySelectorAll('.modal button')].find((b) => b.textContent.trim() === "Don't save"))`)) {
        await evaluate(`[...document.querySelectorAll('.modal button')].find((b) => b.textContent.trim() === "Don't save").click()`)
      }
    }
    const open = async (file, width) => {
      win.webContents.send('file:open-external', file)
      const name = JSON.stringify(path.basename(file))
      const end = Date.now() + 30_000
      while (!(await evaluate(`document.querySelector('.document-title')?.textContent === ${name} && document.querySelector('canvas').width === ${width}`))) {
        await dismissPrompt()
        if (Date.now() > end) throw new Error(`open ${file}: ${await evaluate(`document.querySelector('.toast')?.textContent ?? ''`)}`)
        await pause(30)
      }
      await settle(200)
    }
    const clickTitle = (title) => evaluate(`(() => { const b = [...document.querySelectorAll('button')].find((e) => e.title === ${JSON.stringify(title)}); if (!b) throw new Error('no button ' + ${JSON.stringify(title)}); b.click() })()`)
    const clickText = (selector, text) => evaluate(`(() => { const b = [...document.querySelectorAll(${JSON.stringify(selector)})].find((e) => e.textContent.trim().includes(${JSON.stringify(text)})); if (!b) throw new Error('no ' + ${JSON.stringify(text)}); b.click() })()`)
    const state = () => evaluate(`document.querySelector('.save-state')?.textContent ?? ''`)
    const size = () => evaluate(`({ width: document.querySelector('canvas').width, height: document.querySelector('canvas').height })`)
    const idle = () => until(`!document.querySelector('.working-state') && !document.querySelector('button[title="Save (Ctrl+S)"]').disabled`, 'idle', 60_000)
    /** Client position of an image pixel (handles zoom and scrolling). */
    const toClient = (x, y) => evaluate(`(() => { const c = document.querySelector('canvas'); const r = c.getBoundingClientRect(); return { x: r.left + (${x} + 0.5) * r.width / c.width, y: r.top + (${y} + 0.5) * r.height / c.height } })()`)
    const mouse = (type, point, extra = {}) => win.webContents.sendInputEvent({ type, x: Math.round(point.x), y: Math.round(point.y), button: 'left', clickCount: 1, ...extra })
    const key = async (keyCode, modifiers = [], withChar = false) => {
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode, modifiers })
      if (withChar) win.webContents.sendInputEvent({ type: 'char', keyCode, modifiers })
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode, modifiers })
      await settle(60)
    }
    const drag = async (from, to, steps = 6, modifiers = []) => {
      mouse('mouseMove', from, { modifiers })
      mouse('mouseDown', from, { modifiers })
      for (let index = 1; index <= steps; index += 1) {
        mouse('mouseMove', { x: from.x + ((to.x - from.x) * index) / steps, y: from.y + ((to.y - from.y) * index) / steps }, { modifiers, button: 'left' })
        await pause(10)
      }
      mouse('mouseUp', to, { modifiers })
      await settle(80)
    }
    const pixel = (x, y) => evaluate(`Array.from(document.querySelector('canvas').getContext('2d').getImageData(${x}, ${y}, 1, 1).data)`)
    const canvasHash = () => evaluate(`(() => { const c = document.querySelector('canvas'); const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; let h = 2166136261 >>> 0; for (let i = 0; i < d.length; i += 1) { h ^= d[i]; h = Math.imul(h, 16777619) >>> 0 } return c.width + 'x' + c.height + ':' + h })()`)
    const decode = (bytes, points = []) => evaluate(`(async () => {
      const bitmap = await createImageBitmap(new Blob([Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}), (c) => c.charCodeAt(0))]))
      const c = document.createElement('canvas'); c.width = bitmap.width; c.height = bitmap.height
      const x = c.getContext('2d'); x.drawImage(bitmap, 0, 0); bitmap.close()
      const d = x.getImageData(0, 0, c.width, c.height).data
      let h = 2166136261 >>> 0; for (let i = 0; i < d.length; i += 1) { h ^= d[i]; h = Math.imul(h, 16777619) >>> 0 }
      const result = { width: c.width, height: c.height, hash: c.width + 'x' + c.height + ':' + h, pixels: ${JSON.stringify(points)}.map(([px, py]) => Array.from(x.getImageData(px, py, 1, 1).data)) }
      c.width = 1; c.height = 1
      return result
    })()`)
    const setRange = (selector, value) => evaluate(`(() => { const input = document.querySelector(${JSON.stringify(selector)}); if (!input) throw new Error('no ' + ${JSON.stringify(selector)}); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, ${JSON.stringify(String(value))}); input.dispatchEvent(new Event('input', { bubbles: true })); input.dispatchEvent(new Event('change', { bubbles: true })) })()`)
    const undoButtonEnabled = () => evaluate(`!document.querySelector('button[title="Undo (Ctrl+Z)"]').disabled`)
    const fit = async () => { await evaluate(`[...document.querySelectorAll('.zoom-controls button')].find((b) => b.textContent === 'Fit').click()`); await settle(150) }
    const run = async (name, body) => {
      if (only.size && !only.has(name)) return
      const started = Date.now()
      await body()
      results[name] = { passed: true, ms: Date.now() - started }
      console.log(`ok - ${name} (${Date.now() - started} ms)`)
    }

    // ----- fixtures --------------------------------------------------------------------------------
    const colourDraw = `(x, c) => { x.fillStyle = '#d23c28'; x.fillRect(0, 0, c.width / 2, c.height / 2); x.fillStyle = '#2864d2'; x.fillRect(c.width / 2, 0, c.width / 2, c.height / 2); x.fillStyle = '#28c850'; x.fillRect(0, c.height / 2, c.width / 2, c.height / 2); x.fillStyle = '#f0d228'; x.fillRect(c.width / 2, c.height / 2, c.width / 2, c.height / 2) }`
    const photo = await writeFixture('photo.png', await canvasFixture(480, 360, colourDraw))
    const white = await writeFixture('white.png', png(300, 100, () => [255, 255, 255, 255]))
    const crop = await writeFixture('crop.png', png(200, 100, (x, y) => [x, y, 128, 255]))
    const ramp = await writeFixture('ramp.png', png(256, 64, (x) => { const v = 64 + Math.round((x / 255) * 127); return [v, v, v, 255] }))
    const alphaEdge = await writeFixture('alpha-edge.png', png(64, 64, (x, y) => (x >= 16 && x < 48 && y >= 16 && y < 48 ? [255, 255, 255, 255] : [0, 0, 0, 0])))
    const big = await writeFixture('big.png', await canvasFixture(3000, 2000, colourDraw))

    // ----- 1. Ctrl+V into a new document (nothing open yet) ---------------------------------------
    await run('paste-new-document', async () => {
      clipboardPng = png(120, 80, (x) => [x * 2, 100, 200, 255])
      await key('V', ['control'])
      await until(`document.querySelector('.document-title')?.textContent === 'Pasted image.png' && document.querySelector('canvas').width === 120`, 'pasted document')
      assert.equal(await state(), 'Modified', 'a pasted image is unsaved')
      assert.deepEqual(await size(), { width: 120, height: 80 })
      assert.ok(win.getTitle().startsWith('• '), `title marks unsaved work: ${win.getTitle()}`)
      clipboardPng = null
    })

    // ----- 2. Navigation: zoom at the cursor, ladder, pixelated, Space-drag --------------------------
    await run('navigation', async () => {
      await open(big, 3000)
      await fit()
      const anchor = await toClient(2200, 700)
      const before = await evaluate(`(() => { const s = document.querySelector('.canvas-stack').getBoundingClientRect(); const c = document.querySelector('canvas'); return { left: s.left, top: s.top, zoom: s.width / c.width } })()`)
      const doc = { x: (anchor.x - before.left) / before.zoom, y: (anchor.y - before.top) / before.zoom }
      const prevented = await evaluate(`(() => { const e = new WheelEvent('wheel', { deltaY: -240, deltaMode: 0, ctrlKey: true, clientX: ${anchor.x}, clientY: ${anchor.y}, bubbles: true, cancelable: true }); document.querySelector('canvas').dispatchEvent(e); return e.defaultPrevented })()`)
      assert.equal(prevented, true, 'Ctrl+wheel is handled by a non-passive listener')
      await settle(200)
      const after = await evaluate(`(() => { const s = document.querySelector('.canvas-stack').getBoundingClientRect(); const c = document.querySelector('canvas'); return { left: s.left, top: s.top, zoom: s.width / c.width } })()`)
      assert.ok(after.zoom > before.zoom * 1.3, 'zoomed in')
      const drift = Math.hypot(after.left + doc.x * after.zoom - anchor.x, after.top + doc.y * after.zoom - anchor.y)
      assert.ok(drift <= 1, `the document point stays under the cursor (drift ${drift.toFixed(2)} px)`)
      // Ctrl+1 = one image pixel per device pixel; pixelated from 200% device zoom.
      await key('1', ['control'])
      await settle(150)
      const actual = await evaluate(`(() => { const s = document.querySelector('.canvas-stack').getBoundingClientRect(); return s.width / 3000 * devicePixelRatio })()`)
      assert.ok(Math.abs(actual - 1) < 0.01, `Ctrl+1 is device 1:1 (${actual})`)
      assert.equal(await evaluate(`getComputedStyle(document.querySelector('canvas')).imageRendering`), 'auto')
      await key('=', ['control'])
      await key('=', ['control'])
      const zoomed = await evaluate(`(() => ({ device: document.querySelector('.canvas-stack').getBoundingClientRect().width / 3000 * devicePixelRatio, rendering: getComputedStyle(document.querySelector('canvas')).imageRendering, label: document.querySelector('.zoom-value').textContent }))()`)
      assert.ok(Math.abs(zoomed.device - 2) < 0.01, `ladder 100 -> 150 -> 200 (${zoomed.device})`)
      assert.equal(zoomed.rendering, 'pixelated', 'crisp pixels at 200%')
      assert.equal(zoomed.label, '200%')
      // Space+drag pans.
      const viewport = await evaluate(`(() => { const v = document.querySelector('.viewport').getBoundingClientRect(); const e = document.querySelector('.viewport'); return { x: v.left + v.width / 2, y: v.top + v.height / 2, left: e.scrollLeft, top: e.scrollTop } })()`)
      win.webContents.sendInputEvent({ type: 'keyDown', keyCode: 'Space' })
      await until(`document.querySelector('.viewport').classList.contains('is-pan-ready')`, 'space held')
      await drag({ x: viewport.x, y: viewport.y }, { x: viewport.x - 150, y: viewport.y - 90 }, 8)
      win.webContents.sendInputEvent({ type: 'keyUp', keyCode: 'Space' })
      await settle(80)
      const panned = await evaluate(`(() => { const e = document.querySelector('.viewport'); return { left: e.scrollLeft, top: e.scrollTop } })()`)
      assert.ok(Math.abs(panned.left - viewport.left - 150) <= 2 && Math.abs(panned.top - viewport.top - 90) <= 2, `Space-drag panned by the drag (${JSON.stringify(panned)} from ${JSON.stringify(viewport)})`)
      assert.equal(await state(), 'Saved', 'navigation is not an edit')
      await fit()
    })

    // ----- 3. IMAGE-SIE-1: strokes with Ctrl+Z / Ctrl+S in the middle; pristine flag ---------------
    await run('sie1-stroke-undo', async () => {
      await open(white, 300)
      await clickTitle('Brush')
      await setRange('.size-control input', 10)
      const start = await toClient(10, 50)
      const end = await toClient(150, 50)
      mouse('mouseMove', start)
      mouse('mouseDown', start)
      await settle(60)
      await key('Z', ['control'])
      mouse('mouseMove', end, { button: 'left' })
      await settle(40)
      mouse('mouseUp', end)
      await settle(120)
      assert.equal(await state(), 'Modified', 'Ctrl+Z mid-stroke must not report the image as saved')
      await key('S', ['control'])
      await until(`document.querySelector('.save-state')?.textContent === 'Saved'`, 'saved')
      const saved = await fs.readFile(white)
      const decoded = await decode(saved, [[150, 50], [10, 50], [290, 95]])
      assert.ok(decoded.pixels[0][0] < 100, 'the stroke reaches (150, 50) in the saved file')
      assert.deepEqual(decoded.pixels[2], [255, 255, 255, 255])
      assert.notEqual(sha(saved), sha(png(300, 100, () => [255, 255, 255, 255])), 'not the original bytes')
    })

    await run('sie1-stroke-save', async () => {
      const original = png(300, 100, () => [255, 255, 255, 255])
      await fs.writeFile(white, original)
      await open(white, 300)
      await clickTitle('Brush')
      const a = await toClient(10, 50)
      const b = await toClient(100, 50)
      const c = await toClient(200, 50)
      mouse('mouseMove', a)
      mouse('mouseDown', a)
      mouse('mouseMove', b, { button: 'left' })
      await settle(60)
      await key('S', ['control'])
      await until(`!document.querySelector('button[title="Save (Ctrl+S)"]').disabled && document.querySelector('.save-state')?.textContent !== 'Saving…'`, 'saved mid-stroke')
      mouse('mouseMove', c, { button: 'left' })
      await settle(60)
      mouse('mouseUp', c)
      await settle(150)
      assert.equal(await state(), 'Modified', 'painting after the save is a new, unsaved step')
      const decoded = await decode(await fs.readFile(white), [[100, 50], [200, 50]])
      assert.ok(decoded.pixels[0][0] < 100, 'the save holds the stroke up to the save')
      // Undo everything: the pixels equal the file as opened again, so Save writes the original bytes.
      while (await undoButtonEnabled()) {
        await clickTitle('Undo (Ctrl+Z)')
        await settle(40)
      }
      await clickTitle('Save (Ctrl+S)')
      await idle()
      await settle(100)
      assert.equal(sha(await fs.readFile(white)), sha(original), 'undo back to the opened pixels saves the original bytes')
    })

    // ----- 4. IMAGE-SIE-2: an adjusted crop is not ignored by Save or Close; title marker -------------
    await run('sie2-pending-crop', async () => {
      await open(crop, 200)
      await clickTitle('Crop')
      await until(`Boolean(document.querySelector('.crop-box'))`, 'crop box')
      // An untouched default box is not an edit: closing does not ask.
      const closesBefore = confirmCloseCount
      win.webContents.send('window:close-requested')
      await until(`${confirmCloseCount + 0} >= 0`, 'noop')
      await settle(200)
      assert.equal(confirmCloseCount, closesBefore + 1, 'no prompt for the default crop box')
      assert.equal(await evaluate(`Boolean(document.querySelector('.modal'))`), false)
      // Drag the south-east handle by (-80, -20) image px: the default box (20, 10, 160 x 80) becomes 80 x 60.
      const handle = await evaluate(`(() => { const r = document.querySelector('.crop-handle.se').getBoundingClientRect(); return { x: Math.round(r.left + r.width / 2), y: Math.round(r.top + r.height / 2) } })()`)
      const scale = await evaluate(`document.querySelector('canvas').getBoundingClientRect().width / document.querySelector('canvas').width`)
      await drag(handle, { x: handle.x - 80 * scale, y: handle.y - 20 * scale })
      assert.equal(await state(), 'Crop not applied')
      assert.ok(win.getTitle().startsWith('• '), 'the title shows unsaved work')
      const box = await evaluate(`document.querySelector('.crop-actions span').textContent`)
      // Close now asks, offering to apply the crop.
      const closesAfter = confirmCloseCount
      win.webContents.send('window:close-requested')
      await until(`[...document.querySelectorAll('.modal button')].some((b) => b.textContent === 'Apply crop and save')`, 'close prompt with crop')
      await clickText('.modal button', 'Cancel')
      assert.equal(confirmCloseCount, closesAfter)
      // Save asks too; Apply crop and save writes the cropped size.
      await key('S', ['control'])
      await until(`Boolean(document.querySelector('.choice-modal [data-choice="apply"]'))`, 'crop choice on save')
      await evaluate(`document.querySelector('.choice-modal [data-choice="apply"]').click()`)
      await until(`document.querySelector('.save-state')?.textContent === 'Saved'`, 'saved cropped')
      const saved = await decode(await fs.readFile(crop))
      assert.equal(`${saved.width} x ${saved.height}`, box.replace(' px', '').replace('×', 'x'), `saved the box ${box}`)
      assert.deepEqual([saved.width, saved.height], [80, 60])
      assert.ok(!win.getTitle().startsWith('• '), 'clean title after saving')
    })

    // ----- 5. IMAGE-SIE-3: animated sources save a still PNG copy; the source stays animated ---------
    await run('sie3-animated', async () => {
      const file = await writeFixture('anim.png', apng(40, 30, [() => [255, 0, 0, 255], () => [0, 0, 255, 255]]))
      const original = sha(await fs.readFile(file))
      await open(file, 40)
      await until(`(document.querySelector('.toast')?.textContent ?? '').toLowerCase().includes('animated')`, 'animated notice')
      await clickTitle('Rotate right')
      await clickTitle('Rotate right')
      await settle(80)
      saveDialogs.length = 0
      nextSavePath = path.join(directory, 'anim (frame 1).png')
      await key('S', ['control'])
      await until(`document.querySelector('.save-state')?.textContent === 'Saved'`, 'saved still copy')
      assert.equal(saveDialogs.length, 1, 'Save asked where the copy goes')
      const defaultPath = String(saveDialogs[0].defaultPath || '')
      assert.ok(defaultPath.endsWith('(frame 1).png'), `default name ${defaultPath}`)
      assert.equal(path.dirname(defaultPath), directory, 'next to the source')
      assert.equal(sha(await fs.readFile(file)), original, 'the animated source is untouched')
      assert.ok(pngChunkTypes(await fs.readFile(file)).includes('acTL'))
      const still = await fs.readFile(path.join(directory, 'anim (frame 1).png'))
      assert.ok(!pngChunkTypes(still).includes('acTL'), 'the copy is a still image')
      assert.equal((await decode(still)).width, 40)
    })

    // ----- 6. IMAGE-SIE-14: lossless WebP stays lossless; JPEG transparency warning --------------------
    await run('sie14-lossless-webp', async () => {
      // Opaque colours plus a fully transparent block (the canvas store keeps both exactly; semi-transparent
      // pixels can move by a few levels through any canvas round trip, a known Simple-mode limitation).
      const source = Buffer.from(await canvasFixture(64, 48, `(x) => { for (let i = 0; i < 64; i++) { x.fillStyle = 'rgb(' + (i * 4) + ', 90, ' + (255 - i * 4) + ')'; x.fillRect(i, 0, 1, 48) } x.clearRect(40, 10, 16, 20) }`, 'image/webp', 1), 'base64')
      assert.ok(webpChunkTypes(source).includes('VP8L'), 'fixture is lossless')
      const file = await writeFixture('lossless.webp', source)
      const sourcePixels = await decode(source)
      await open(file, 64)
      await clickTitle('Rotate right')
      await clickTitle('Rotate left')
      await settle(60)
      assert.equal(await state(), 'Modified')
      await key('S', ['control'])
      await until(`document.querySelector('.save-state')?.textContent === 'Saved'`, 'saved webp')
      const saved = await fs.readFile(file)
      const chunks = webpChunkTypes(saved)
      assert.ok(chunks.includes('VP8L') && !chunks.includes('VP8 '), `still lossless: ${chunks.join(',')}`)
      assert.equal((await decode(saved)).hash, sourcePixels.hash, 'no pixel changed')
    })

    await run('sie14-jpeg-alpha', async () => {
      const file = await writeFixture('photo.jpg', await canvasFixture(160, 120, colourDraw, 'image/jpeg', 0.9))
      await open(file, 160)
      await clickTitle('Eraser')
      await drag(await toClient(30, 30), await toClient(90, 60))
      await until(`document.querySelector('.inspector')?.innerText.includes('transparency') || true`, 'erased')
      await key('S', ['control'])
      await until(`Boolean(document.querySelector('.choice-modal [data-choice="png"]'))`, 'JPEG transparency warning')
      assert.ok((await evaluate(`document.querySelector('.choice-modal').innerText`)).includes('white'))
      await evaluate(`document.querySelector('.choice-modal [data-choice="cancel"]').click()`)
      await until(`!document.querySelector('.choice-modal')`, 'warning closed')
      assert.equal(await state(), 'Modified', 'nothing was saved')
    })

    // ----- 7. Enter on the focused Crop button applies once (defect 3) ------------------------------
    await run('enter-crop-once', async () => {
      await open(photo, 480)
      await clickTitle('Crop')
      await until(`Boolean(document.querySelector('.crop-box'))`, 'crop box')
      await evaluate(`[...document.querySelectorAll('button')].find((b) => b.title === 'Crop').focus()`)
      await key('Enter', [], true)
      await until(`document.querySelector('canvas').width === 384`, 'cropped once')
      await settle(250)
      assert.equal(await evaluate(`Boolean(document.querySelector('.crop-box'))`), false, 'the crop tool did not re-open')
      assert.deepEqual(await size(), { width: 384, height: 288 })
    })

    // ----- 8. Aspect presets give exact ratios; flips; straighten with auto-crop ---------------------
    await run('aspect-presets', async () => {
      await open(photo, 480)
      const expectations = [['1:1', false, 360, 360], ['4:3', false, 480, 360], ['3:2', false, 480, 320], ['16:9', false, 480, 270], ['4:3', true, 270, 360], ['3:2', true, 240, 360], ['original', false, 480, 360]]
      for (const [preset, swap, width, height] of expectations) {
        await clickTitle('Crop')
        await until(`Boolean(document.querySelector('.crop-box'))`, 'crop box')
        await evaluate(`document.querySelector('[data-aspect="${preset}"]').click()`)
        if (swap) await evaluate(`document.querySelector('[aria-label="Swap the box orientation"]').click()`)
        await settle(60)
        await clickText('button', 'Apply crop')
        await until(`!document.querySelector('.crop-box')`, 'applied')
        if (!(width === 480 && height === 360)) {
          assert.deepEqual(await size(), { width, height }, `${preset}${swap ? ' turned' : ''}`)
          await clickTitle('Undo (Ctrl+Z)')
          await until(`document.querySelector('canvas').width === 480 && document.querySelector('canvas').height === 360`, 'undo crop')
        } else {
          assert.deepEqual(await size(), { width: 480, height: 360 })
        }
      }
    })

    await run('flip-twice', async () => {
      await open(photo, 480)
      const original = await canvasHash()
      await clickTitle('Crop')
      await until(`Boolean(document.querySelector('[aria-label="Flip horizontal"]'))`, 'crop strip')
      await evaluate(`document.querySelector('[aria-label="Flip horizontal"]').click()`)
      await settle(60)
      assert.notEqual(await canvasHash(), original)
      assert.deepEqual(await pixel(10, 10), [40, 100, 210, 255], 'blue quadrant moved left')
      await evaluate(`document.querySelector('[aria-label="Flip horizontal"]').click()`)
      await settle(60)
      assert.equal(await canvasHash(), original, 'flip twice restores the image')
      await evaluate(`document.querySelector('[aria-label="Flip vertical"]').click()`)
      await evaluate(`document.querySelector('[aria-label="Flip vertical"]').click()`)
      await settle(60)
      assert.equal(await canvasHash(), original)
      await clickText('.crop-actions button', 'Cancel')
    })

    await run('straighten-autocrop', async () => {
      await open(photo, 480)
      await clickTitle('Crop')
      await until(`Boolean(document.querySelector('.straighten-control input'))`, 'straighten')
      await setRange('.straighten-control input', 10)
      await settle(120)
      const t = (10 * Math.PI) / 180
      const scale = Math.min(476 / (476 * Math.cos(t) + 356 * Math.sin(t)), 356 / (476 * Math.sin(t) + 356 * Math.cos(t)))
      const label = await evaluate(`document.querySelector('.crop-actions span').textContent`)
      await clickText('button', 'Apply crop')
      await until(`!document.querySelector('.crop-box') && !document.querySelector('.working-state')`, 'straightened', 30_000)
      const result = await size()
      assert.ok(Math.abs(result.width - 476 * scale) <= 2 && Math.abs(result.height - 356 * scale) <= 2, `auto-crop size ${JSON.stringify(result)} vs ${476 * scale} x ${356 * scale} (box ${label})`)
      for (const [x, y] of [[0, 0], [result.width - 1, 0], [0, result.height - 1], [result.width - 1, result.height - 1]]) {
        assert.equal((await pixel(x, y))[3], 255, `corner ${x},${y} is opaque`)
      }
      assert.ok((await evaluate(`document.querySelector('.inspector').innerText`)).includes('RGB · opaque'))
      await clickTitle('Undo (Ctrl+Z)')
      await until(`document.querySelector('canvas').width === 480`, 'undo straighten')
    })

    // ----- 9. Resize to 50%: exact size and no dark fringe on transparent edges ----------------------
    await run('resize-half', async () => {
      await open(alphaEdge, 64)
      await evaluate(`(() => { const i = document.querySelector('.inspector'); if (i.hidden) document.querySelector('[aria-controls="image-inspector"]').click() })()`)
      await until(`Boolean([...document.querySelectorAll('.inspector button')].find((b) => b.textContent === 'Resize…'))`, 'resize entry')
      await clickText('.inspector button', 'Resize…')
      await until(`Boolean(document.querySelector('.resize-dialog'))`, 'resize dialog')
      await clickText('.resize-presets button', '50%')
      await evaluate(`document.querySelector('.resize-dialog button[type="submit"]').click()`)
      await until(`document.querySelector('canvas').width === 32 && !document.querySelector('.resize-dialog')`, 'resized', 30_000)
      assert.deepEqual(await size(), { width: 32, height: 32 })
      const edge = await evaluate(`(() => { const d = document.querySelector('canvas').getContext('2d').getImageData(0, 0, 32, 32).data; let dark = 0; for (let i = 0; i < d.length; i += 4) if (d[i + 3] > 0 && (d[i] < 250 || d[i + 1] < 250 || d[i + 2] < 250)) dark += 1; return dark })()`)
      assert.equal(edge, 0, 'no dark fringe: every visible pixel stays white')
    })

    // ----- 10. Adjust: preview, one undo step; Auto on a 64..191 ramp; a Look ------------------------
    await run('adjust-one-step', async () => {
      await open(photo, 480)
      const original = await canvasHash()
      await clickTitle('Adjust light and colour, or apply a look')
      await until(`Boolean(document.querySelector('.adjust-panel'))`, 'adjust panel')
      await setRange('[data-adjust="exposure"] input', 60)
      await until(`document.querySelector('.adjust-preview') && !document.querySelector('.adjust-preview').hidden`, 'preview shown')
      assert.equal(await canvasHash(), original, 'the preview does not touch the image')
      assert.equal(await undoButtonEnabled(), false)
      await key('\\')
      await until(`document.querySelector('.adjust-preview').hidden && getComputedStyle(document.querySelector('canvas')).opacity === '1'`, 'compare shows the original')
      await key('\\')
      await until(`!document.querySelector('.adjust-preview').hidden`, 'compare off')
      await clickText('.adjust-actions button', 'Done')
      await until(`!document.querySelector('.adjust-panel') && !document.querySelector('.working-state')`, 'applied', 30_000)
      assert.notEqual(await canvasHash(), original)
      assert.equal(await state(), 'Modified')
      await clickTitle('Undo (Ctrl+Z)')
      await settle(80)
      assert.equal(await canvasHash(), original, 'one undo restores the image')
      assert.equal(await undoButtonEnabled(), false, 'exactly one step')
    })

    await run('adjust-auto', async () => {
      await open(ramp, 256)
      await clickTitle('Adjust light and colour, or apply a look')
      await until(`Boolean(document.querySelector('.adjust-auto:not(:disabled)'))`, 'auto ready')
      await evaluate(`document.querySelector('.adjust-auto').click()`)
      await until(`document.querySelector('.adjust-preview') && !document.querySelector('.adjust-preview').hidden`, 'auto preview')
      await clickText('.adjust-actions button', 'Done')
      await until(`!document.querySelector('.adjust-panel') && !document.querySelector('.working-state')`, 'auto applied', 30_000)
      const range = await evaluate(`(() => { const d = document.querySelector('canvas').getContext('2d').getImageData(0, 0, 256, 64).data; let min = 255, max = 0; for (let i = 0; i < d.length; i += 4) { min = Math.min(min, d[i]); max = Math.max(max, d[i]) } return { min, max } })()`)
      assert.ok(range.max - range.min >= 0.95 * 255, `Auto stretched the ramp to ${range.min}..${range.max}`)
    })

    await run('look', async () => {
      await open(photo, 480)
      await clickTitle('Adjust light and colour, or apply a look')
      await until(`Boolean(document.querySelector('.adjust-tabs'))`, 'adjust panel')
      await clickText('.adjust-tabs button', 'Looks')
      await until(`Boolean(document.querySelector('[data-look="mono"]'))`, 'looks')
      await evaluate(`document.querySelector('[data-look="mono"]').click()`)
      await clickText('.adjust-actions button', 'Done')
      await until(`!document.querySelector('.adjust-panel') && !document.querySelector('.working-state')`, 'look applied', 30_000)
      const [r, g, b] = await pixel(60, 60)
      assert.ok(Math.max(r, g, b) - Math.min(r, g, b) <= 3, `Mono removed the colour (${r},${g},${b})`)
    })

    // ----- 11. Markup bakes at the expected pixels -----------------------------------------------------
    await run('markup', async () => {
      await open(white, 300)
      await evaluate(`[...document.querySelectorAll('button')].find((b) => b.title.startsWith('Markup')).click()`)
      await until(`Boolean(document.querySelector('.markup-layer'))`, 'markup layer')
      const draw = async (kind, from, to) => {
        await evaluate(`document.querySelector('[data-markup-kind="${kind}"]').click()`)
        await drag(await toClient(from[0], from[1]), await toClient(to[0], to[1]))
      }
      await setRange('.markup-width input', 4)
      await draw('rectangle', [10, 10], [60, 50])
      await draw('ellipse', [80, 10], [140, 50])
      await draw('line', [160, 30], [220, 30])
      await draw('arrow', [230, 80], [290, 80])
      await evaluate(`document.querySelector('[data-markup-kind="text"]').click()`)
      const textAt = await toClient(20, 62)
      mouse('mouseMove', textAt)
      mouse('mouseDown', textAt)
      mouse('mouseUp', textAt)
      await until(`Boolean(document.querySelector('.markup-text-editor'))`, 'text editor')
      win.webContents.insertText('Hi')
      await settle(80)
      await key('Escape')
      await until(`!document.querySelector('.markup-text-editor')`, 'text committed')
      assert.equal(await evaluate(`document.querySelectorAll('.markup-layer > path, .markup-layer > rect:not(.markup-selection):not(.markup-handle), .markup-layer > ellipse, .markup-layer > text').length`), 5)
      // Ctrl+Z in Markup undoes the last markup item, Ctrl+Y brings it back.
      await key('Z', ['control'])
      assert.equal(await evaluate(`document.querySelectorAll('.markup-layer > text').length`), 0)
      await key('Y', ['control'])
      assert.equal(await evaluate(`document.querySelectorAll('.markup-layer > text').length`), 1)
      await evaluate(`[...document.querySelectorAll('.markup-strip button')].find((b) => b.textContent.includes('Done') || b.title.startsWith('Draw')).click()`)
      await until(`!document.querySelector('.markup-layer')`, 'baked')
      const red = (p) => p[0] > 180 && p[1] < 90 && p[2] < 90
      assert.ok(red(await pixel(10, 30)), 'rectangle edge')
      assert.ok(!red(await pixel(35, 30)), 'rectangle inside stays white')
      assert.ok(red(await pixel(110, 10)) || red(await pixel(110, 11)), 'ellipse top')
      assert.ok(red(await pixel(190, 30)), 'line')
      assert.ok(red(await pixel(285, 80)) && red(await pixel(277, 85)) && !red(await pixel(250, 85)), 'arrow head wider than the shaft')
      const textPixels = await evaluate(`(() => { const d = document.querySelector('canvas').getContext('2d').getImageData(20, 62, 60, 30).data; let n = 0; for (let i = 0; i < d.length; i += 4) if (d[i] > 180 && d[i + 1] < 90) n += 1; return n })()`)
      assert.ok(textPixels > 20, `text painted (${textPixels} px)`)
      assert.equal(await state(), 'Modified')
      await clickTitle('Undo (Ctrl+Z)')
      await settle(80)
      assert.equal((await pixel(10, 30)).join(','), '255,255,255,255', 'one undo removes the baked markup')
    })

    // ----- 11b. Paste into an open image: a movable picture in Markup, baked on Done ------------------
    await run('paste-into-image', async () => {
      await open(white, 300)
      clipboardPng = png(40, 20, () => [0, 0, 255, 255])
      await key('V', ['control'])
      await until(`Boolean(document.querySelector('.markup-layer image'))`, 'pasted picture')
      await evaluate(`[...document.querySelectorAll('.markup-strip button')].find((b) => b.title.startsWith('Draw')).click()`)
      await until(`!document.querySelector('.markup-layer')`, 'baked picture')
      assert.deepEqual(await pixel(150, 50), [0, 0, 255, 255], 'the picture is centred (130..170 x 40..60)')
      assert.deepEqual(await pixel(120, 50), [255, 255, 255, 255])
      assert.equal(await state(), 'Modified')
      clipboardPng = null
    })

    // ----- 11c. Interactions: context menu, [ ], crop nudges and Escape, level line, middle-drag, double-click
    await run('interactions', async () => {
      await open(big, 3000)
      await fit()
      await evaluate(`(() => { const r = document.querySelector('canvas').getBoundingClientRect(); document.querySelector('.canvas-stage').dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, button: 2, clientX: r.left + 20, clientY: r.top + 20 })) })()`)
      await until(`Boolean(document.querySelector('.image-context-menu'))`, 'context menu')
      const labels = await evaluate(`[...document.querySelectorAll('.image-context-menu [role="menuitem"] span')].map((e) => e.textContent)`)
      assert.deepEqual(labels, ['Copy image', 'Paste', 'Undo', 'Redo', 'Resize…'])
      await key('Escape')
      await until(`!document.querySelector('.image-context-menu')`, 'menu closed')
      await clickTitle('Brush')
      await setRange('.size-control input', 12)
      // Steps: every px to 10, then 5 px to 50, 10 px to 100, then 120.
      await key(']')
      assert.equal(await evaluate(`document.querySelector('.size-control input').value`), '15')
      await key('[')
      await key('[')
      assert.equal(await evaluate(`document.querySelector('.size-control input').value`), '9')
      // Crop: arrows nudge 1 px, Shift+arrows 10 px; Escape cancels without an edit.
      await clickTitle('Crop')
      await until(`Boolean(document.querySelector('.crop-box'))`, 'crop box')
      const box = () => evaluate(`(() => { const b = document.querySelector('.crop-box').style; return { x: Math.round(parseFloat(b.left) * 30 * 100) / 100, y: Math.round(parseFloat(b.top) * 20 * 100) / 100 } })()`)
      const start = await box()
      await key('Right')
      await key('Down', ['shift'])
      const moved = await box()
      assert.ok(Math.abs(moved.x - start.x - 1) < 0.02 && Math.abs(moved.y - start.y - 10) < 0.02, `nudged ${JSON.stringify(start)} -> ${JSON.stringify(moved)}`)
      assert.equal(await state(), 'Crop not applied')
      await key('Escape')
      await until(`!document.querySelector('.crop-box')`, 'crop cancelled')
      assert.equal(await state(), 'Saved')
      // Level by drawing a line that falls 5 degrees to the right.
      await clickTitle('Crop')
      await until(`Boolean(document.querySelector('[aria-label="Level by drawing a line"]'))`, 'crop strip')
      await evaluate(`document.querySelector('[aria-label="Level by drawing a line"]').click()`)
      await until(`Boolean(document.querySelector('.level-line-layer'))`, 'level layer')
      const t = (5 * Math.PI) / 180
      await drag(await toClient(300, 700), await toClient(300 + Math.cos(t) * 2400, 700 + Math.sin(t) * 2400), 6)
      // Mouse positions are whole CSS pixels (about 0.1 degree over this line), so allow +-0.25 degrees.
      await until(`!document.querySelector('.level-line-layer') && Math.abs(parseFloat(document.querySelector('.straighten-control output').textContent) + 5) <= 0.25`, 'levelled')
      await clickText('.crop-actions button', 'Cancel')
      // Middle-button drag pans at 1:1.
      await key('1', ['control'])
      await settle(120)
      const centre = await evaluate(`(() => { const v = document.querySelector('.viewport'); const r = v.getBoundingClientRect(); return { x: r.left + r.width / 2, y: r.top + r.height / 2, left: v.scrollLeft, top: v.scrollTop } })()`)
      win.webContents.sendInputEvent({ type: 'mouseDown', x: Math.round(centre.x), y: Math.round(centre.y), button: 'middle', clickCount: 1 })
      for (let index = 1; index <= 5; index += 1) win.webContents.sendInputEvent({ type: 'mouseMove', x: Math.round(centre.x + index * 20), y: Math.round(centre.y + index * 10), button: 'middle' })
      win.webContents.sendInputEvent({ type: 'mouseUp', x: Math.round(centre.x + 100), y: Math.round(centre.y + 50), button: 'middle', clickCount: 1 })
      await settle(100)
      const panned = await evaluate(`({ left: document.querySelector('.viewport').scrollLeft, top: document.querySelector('.viewport').scrollTop })`)
      assert.ok(Math.abs(centre.left - panned.left - 100) <= 2 && Math.abs(centre.top - panned.top - 50) <= 2, `middle-drag panned ${JSON.stringify(centre)} -> ${JSON.stringify(panned)}`)
      // View tool: double-click toggles 1:1 and Fit.
      await clickTitle('View')
      await fit()
      const fitLabel = await evaluate(`document.querySelector('.zoom-value').textContent`)
      const at = await toClient(1500, 1000)
      for (const clickCount of [1, 2]) {
        mouse('mouseDown', at, { clickCount })
        mouse('mouseUp', at, { clickCount })
      }
      await until(`document.querySelector('.zoom-value').textContent === '100%'`, 'double-click to 1:1')
      const again = await toClient(1500, 1000)
      for (const clickCount of [1, 2]) {
        mouse('mouseDown', again, { clickCount })
        mouse('mouseUp', again, { clickCount })
      }
      await until(`document.querySelector('.zoom-value').textContent === ${JSON.stringify(fitLabel)}`, 'double-click back to Fit')
      assert.equal(await state(), 'Saved', 'none of this edited the image')
    })

    // ----- 12. Eyedropper sets the brush colour -------------------------------------------------------
    await run('eyedropper', async () => {
      await open(photo, 480)
      await evaluate(`document.querySelector('[aria-label="Pick a colour from the image"]').click()`)
      const point = await toClient(360, 270)
      mouse('mouseMove', point)
      mouse('mouseDown', point)
      mouse('mouseUp', point)
      await until(`document.querySelector('input[type="color"]').value === '#f0d228'`, 'picked colour')
      assert.equal(await evaluate(`document.querySelector('button[title="Brush"]').classList.contains('active')`), true, 'back to the brush')
      assert.equal(await state(), 'Saved', 'picking is not an edit')
    })

    // ----- 13. A 50 MP image keeps at least 10 undoable brush strokes ------------------------------
    await run('undo-depth-50mp', async () => {
      const file = await writeFixture('fifty.png', await canvasFixture(10000, 5000, `(x, c) => { x.fillStyle = '#ffffff'; x.fillRect(0, 0, c.width, c.height) }`))
      await open(file, 10000)
      await fit()
      await clickTitle('Brush')
      await setRange('.size-control input', 120)
      await setRange('input[type="color"]', '#111111')
      const points = []
      for (let index = 0; index < 12; index += 1) points.push([400 + index * 800, 2500])
      for (const [x, y] of points) {
        const at = await toClient(x, y)
        mouse('mouseMove', at)
        mouse('mouseDown', at)
        mouse('mouseUp', at)
        await settle(40)
      }
      await until(`document.querySelector('.save-state')?.textContent === 'Modified'`, 'painted')
      for (const [x, y] of points) assert.ok((await pixel(x, y))[0] < 60, `stroke at ${x}`)
      for (let index = 0; index < 10; index += 1) {
        assert.ok(await undoButtonEnabled(), `undo ${index + 1} is available`)
        await clickTitle('Undo (Ctrl+Z)')
        await settle(30)
      }
      for (const [index, [x, y]] of points.entries()) {
        const value = (await pixel(x, y))[0]
        if (index < 2) assert.ok(value < 60, `stroke ${index + 1} kept`)
        else assert.equal(value, 255, `stroke ${index + 1} undone`)
      }
      await open(white, 300)
    })

    const resultPath = path.join(root, 'tmp', 'simple-plus-smoke-result.json')
    await fs.mkdir(path.dirname(resultPath), { recursive: true })
    await fs.writeFile(resultPath, JSON.stringify({ passed: true, results }, null, 2))
    console.log(JSON.stringify({ passed: true, cases: Object.keys(results).length }))
  } catch (error) {
    console.error(error)
    exitCode = 1
  } finally {
    for (const window of BrowserWindow.getAllWindows()) window.destroy()
    const absolute = path.resolve(directory)
    if (absolute.startsWith(path.resolve(os.tmpdir()) + path.sep) && path.basename(absolute).startsWith('simple-image-plus-')) {
      await Promise.race([fs.rm(absolute, { recursive: true, force: true, maxRetries: 2, retryDelay: 100 }).catch(() => {}), pause(3000)])
    }
  }
  app.exit(exitCode)
})
