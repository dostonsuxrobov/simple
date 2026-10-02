// qa/advanced-smoke.cjs (WP6): npm run smoke:advanced (after npm run build:web)
// Drives the production renderer through the Advanced editor the way a user would, in a hidden window with
// fake IPC (like qa/smoke.cjs: nothing is written outside the work folder, nothing is ever printed):
//   1. open a PNG, press Advanced: the image is the Background layer, nothing is modified;
//   2. New layer, paint on it, set Multiply (Layers panel);
//   3. a marquee selection, then a Levels adjustment layer: it takes the selection as its mask; its
//      settings change in the Properties panel (Levels editor);
//   4. a marquee on the Background and Ctrl+J (layer via copy); Free Transform (Ctrl+T, W field, Enter);
//   5. Undo / Redo, a History panel jump and back, a snapshot and reverting to it;
//   6. panels and dialogs: the colour picker, Image Size (applied, undone), a destructive Levels dialog with
//      a live canvas preview (cancelled), a Gaussian Blur dialog (100% preview, OK, undone), Feather, the
//      Layers panel's rename, drag to reorder, lock, Ctrl+click selection and context menu;
//   7. Save (Ctrl+S) writes a PSD: parsed here with ag-psd (layer count, names, order, blend modes, mask);
//   8. Export As PNG: the size matches; Print opens and is cancelled (no print job);
//   9. Simple: the flatten prompt, Flatten; the Simple canvas equals the exported composite and Ctrl+Z
//      brings back the image as it was before Advanced.
// Environment: SIMPLE_IMAGE_QA_DIR (work folder; default a new temp folder), SIMPLE_IMAGE_SMOKE_TIMEOUT_MS.
'use strict'

const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const { app, BrowserWindow, ipcMain } = require('electron')
const { imageDimensions } = require('../electron/image-files.cjs')

const projectRoot = path.join(__dirname, '..')
const workDir = process.env.SIMPLE_IMAGE_QA_DIR
  ? path.resolve(process.env.SIMPLE_IMAGE_QA_DIR)
  : fs.mkdtempSync(path.join(os.tmpdir(), 'simple-image-advanced-smoke-'))
fs.mkdirSync(workDir, { recursive: true })
app.setPath('userData', path.join(workDir, 'user-data'))

const timeoutMs = Number(process.env.SIMPLE_IMAGE_SMOKE_TIMEOUT_MS) || 240_000
const hardTimeout = setTimeout(() => {
  console.error(`advanced-smoke: timed out after ${timeoutMs} ms (artifacts: ${workDir})`)
  app.exit(2)
}, timeoutMs)
hardTimeout.unref?.()

// #region fixture: a 320 x 200 opaque gradient PNG

const W = 320
const H = 200
const fixtureColor = (x, y) => [Math.round(x * 0.7), y, 180, 255]

function crc32(buffer) {
  let table = crc32.table
  if (!table) {
    table = crc32.table = Array.from({ length: 256 }, (_, n) => {
      let c = n
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      return c >>> 0
    })
  }
  let crc = 0xffffffff
  for (const byte of buffer) crc = table[(crc ^ byte) & 0xff] ^ (crc >>> 8)
  return (crc ^ 0xffffffff) >>> 0
}

function pngChunk(type, data) {
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
    for (let x = 0; x < width; x += 1) {
      const p = y * (width * 4 + 1) + 1 + x * 4
      const [r, g, b, a] = pixel(x, y)
      raw[p] = r
      raw[p + 1] = g
      raw[p + 2] = b
      raw[p + 3] = a
    }
  }
  const header = Buffer.alloc(13)
  header.writeUInt32BE(width, 0)
  header.writeUInt32BE(height, 4)
  header[8] = 8
  header[9] = 6
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    pngChunk('IHDR', header),
    pngChunk('IDAT', zlib.deflateSync(raw)),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}

const fixturePath = path.join(workDir, 'gradient.png')
const fixtureBytes = encodePng(W, H, fixtureColor)
fs.writeFileSync(fixturePath, fixtureBytes)

// #endregion

// #region fake backend (captures every save; never prints)

const ag = require('ag-psd')
ag.initializeCanvas(() => { throw new Error('no canvas in the main process') }, (width, height) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) }))

const saves = []
const printJobs = []
ipcMain.handle('file:open-path', () => ({ data: new Uint8Array(fixtureBytes), name: 'gradient.png', path: fixturePath, size: fixtureBytes.length, format: 'png', mime: 'image/png', directSave: true }))
ipcMain.handle('file:open-dialog', () => null)
ipcMain.handle('file:list-siblings', () => [])
ipcMain.handle('file:save', (_event, input) => {
  const format = input && input.format
  const bytes = Buffer.from(input && input.data ? input.data : [])
  const extension = format === 'jpeg' ? '.jpg' : `.${format}`
  const target = path.join(workDir, `saved-${saves.length + 1}${extension}`)
  fs.writeFileSync(target, bytes)
  saves.push({ format, purpose: input && input.purpose, forceDialog: Boolean(input && input.forceDialog), bytes, path: target })
  return { path: target, name: path.basename(target), size: bytes.length, format }
})
ipcMain.handle('image:convert-to-pdf', () => null)
ipcMain.handle('image:print', (_event, input) => {
  // Never print from a smoke test: record the request and report failure.
  printJobs.push({ name: input && input.name })
  return false
})
ipcMain.handle('clipboard:write-png', () => ({ width: 1, height: 1 }))
ipcMain.handle('clipboard:read-image', () => null)
ipcMain.handle('clipboard:has-image', () => false)
ipcMain.handle('app:get-version', () => '0.0.0-smoke')
ipcMain.handle('app:new-window', () => false)
ipcMain.handle('file:open-in-new-window', () => false)
ipcMain.on('window:set-title', () => {})
ipcMain.on('window:minimize', () => {})
ipcMain.on('window:toggle-maximize', () => {})
ipcMain.on('window:confirm-close', () => {})

// #endregion

const result = { workDir, steps: {}, failures: [], consoleErrors: [] }
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))

function check(condition, message) {
  if (!condition) result.failures.push(message)
  return Boolean(condition)
}

function near(a, b, tolerance = 2) {
  return Array.isArray(a) && Array.isArray(b) && a.length === b.length && a.every((value, index) => Math.abs(value - b[index]) <= tolerance)
}

function psdSummary(bytes) {
  const psd = ag.readPsd(bytes, { useImageData: true, skipThumbnail: true })
  return {
    width: psd.width,
    height: psd.height,
    layers: (psd.children || []).map((layer) => ({
      name: layer.name,
      blendMode: layer.blendMode,
      opacity: layer.opacity,
      hidden: Boolean(layer.hidden),
      mask: Boolean(layer.mask),
      adjustment: layer.adjustment ? layer.adjustment.type : null,
    })),
  }
}

app.whenReady().then(async () => {
  const window = new BrowserWindow({
    width: 1365,
    height: 860,
    show: false,
    backgroundColor: '#e9e9e9',
    webPreferences: { preload: path.join(projectRoot, 'electron', 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true },
  })
  window.webContents.on('console-message', (event) => {
    const level = event.level ?? event.params?.level
    const message = event.message ?? event.params?.message
    if (level === 'error' || level === 3) result.consoleErrors.push(String(message).slice(0, 500))
  })
  const js = (expression) => window.webContents.executeJavaScript(expression)
  async function waitFor(expression, timeout = 15_000, label = expression) {
    const started = Date.now()
    let last
    while (Date.now() - started < timeout) {
      try {
        last = await js(expression)
        if (last) return last
      } catch (error) {
        last = String(error)
      }
      await sleep(70)
    }
    throw new Error(`Timed out waiting for: ${label} (last: ${JSON.stringify(last)})`)
  }
  const key = (init) => js(`window.dispatchEvent(new KeyboardEvent('keydown', ${JSON.stringify({ bubbles: true, cancelable: true, ...init })})); true`)
  const keyUp = (init) => js(`window.dispatchEvent(new KeyboardEvent('keyup', ${JSON.stringify({ bubbles: true, cancelable: true, ...init })})); true`)
  const click = (selector, label = selector) => js(`(() => { const el = document.querySelector(${JSON.stringify(selector)}); if (!el) return false; el.click(); return true })()`)
    .then((ok) => { if (!ok) throw new Error(`Nothing to click: ${label}`) })
  // Document point -> client point through the stage's data-view hook (CanvasView).
  const toClient = (x, y) => js(`(() => {
    const stage = document.querySelector('.ae-stage')
    const [zoom, ox, oy] = stage.dataset.view.split(' ').map(Number)
    const rect = stage.getBoundingClientRect()
    return { x: rect.left + ${x} * zoom + ox, y: rect.top + ${y} * zoom + oy }
  })()`)
  const pointer = (type, point, extra = {}) => js(`(() => {
    const stage = document.querySelector('.ae-stage')
    stage.dispatchEvent(new PointerEvent(${JSON.stringify(type)}, ${JSON.stringify({ bubbles: true, cancelable: true, composed: true, pointerId: 7, pointerType: 'mouse', isPrimary: true, clientX: point.x, clientY: point.y, button: type === 'pointermove' ? -1 : 0, buttons: type === 'pointerup' ? 0 : 1, pressure: type === 'pointerup' ? 0 : 0.5, ...extra })}))
    return true
  })()`)
  async function dragOnCanvas(from, to, steps = 8, extra = {}) {
    const a = await toClient(from[0], from[1])
    const b = await toClient(to[0], to[1])
    await pointer('pointermove', a, { buttons: 0, pressure: 0, ...extra })
    await pointer('pointerdown', a, extra)
    for (let i = 1; i <= steps; i += 1) await pointer('pointermove', { x: a.x + ((b.x - a.x) * i) / steps, y: a.y + ((b.y - a.y) * i) / steps }, extra)
    await pointer('pointerup', b, extra)
  }
  // Sets a field the way typing does (React sees an input event), then presses Enter in it.
  const typeInto = (selector, value, enter = true) => js(`(() => {
    const input = document.querySelector(${JSON.stringify(selector)})
    if (!input) return false
    input.focus()
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    setter.call(input, ${JSON.stringify(String(value))})
    input.dispatchEvent(new Event('input', { bubbles: true }))
    if (${enter}) input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true }))
    return true
  })()`).then((ok) => { if (!ok) throw new Error(`No field: ${selector}`) })
  const selectValue = (selector, value) => js(`(() => {
    const select = document.querySelector(${JSON.stringify(selector)})
    if (!select) return false
    const setter = Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set
    setter.call(select, ${JSON.stringify(value)})
    select.dispatchEvent(new Event('change', { bubbles: true }))
    return true
  })()`).then((ok) => { if (!ok) throw new Error(`No select: ${selector}`) })
  const historyLabels = () => js(`Array.from(document.querySelectorAll('.ae-history-row')).map((row) => row.textContent.trim())`)
  const historyCurrent = () => js(`document.querySelector('.ae-history-row.is-current')?.textContent.trim() ?? null`)
  const layerNames = () => js(`Array.from(document.querySelectorAll('.ae-layer-row .ae-layer-name')).map((row) => row.textContent.trim())`)
  const saveState = () => js(`document.querySelector('.save-state')?.textContent ?? null`)
  const settle = () => js(`window.__simpleAdvanced.compositor.settle().then(() => true)`)
  // Document state through the QA hook (AdvancedEditor exposes it when __SIMPLE_IMAGE_QA__ is set).
  const docState = () => js(`(() => {
    const state = window.__simpleAdvanced.store.getState()
    const bounds = (layer) => {
      const source = layer.kind === 'raster' ? { s: layer.surface, x: layer.offsetX, y: layer.offsetY } : layer.raster ? { s: layer.raster.surface, x: layer.raster.offsetX, y: layer.raster.offsetY } : null
      const b = source && source.s.contentBounds()
      return b ? { x: b.x + source.x, y: b.y + source.y, width: b.width, height: b.height } : null
    }
    return {
      width: state.width,
      height: state.height,
      active: state.layers.find((layer) => layer.id === state.activeLayerId)?.name ?? null,
      target: state.editTarget,
      selection: state.selection ? state.selection.bounds : null,
      layers: state.layers.map((layer) => ({ name: layer.name, kind: layer.kind, blendMode: layer.blendMode, opacity: layer.opacity, visible: layer.visible, mask: Boolean(layer.mask), clipped: layer.clipped, locks: layer.locks, bounds: bounds(layer), adjustment: layer.kind === 'adjustment' ? layer.adjustment : null })),
    }
  })()`)
  // Composite colour at a document point (what Save and Export flatten).
  const sample = (x, y) => js(`(() => { const c = window.__simpleAdvanced.compositor.sample(${x}, ${y}, 1, 'all'); return [c.r, c.g, c.b, c.a] })()`)
  // What the screen shows at a document point.
  const displayPixel = (x, y) => js(`(() => {
    const stage = document.querySelector('.ae-stage')
    const [zoom, ox, oy, dpr] = stage.dataset.view.split(' ').map(Number)
    const canvas = stage.querySelector('.ae-display')
    const px = Math.floor(Math.round(ox * dpr) + (${x} + 0.5) * zoom * dpr)
    const py = Math.floor(Math.round(oy * dpr) + (${y} + 0.5) * zoom * dpr)
    return Array.from(canvas.getContext('2d').getImageData(px, py, 1, 1).data)
  })()`)
  const simplePixel = (x, y) => js(`Array.from(document.querySelector('canvas').getContext('2d').getImageData(${x}, ${y}, 1, 1).data)`)
  const menuClick = async (menu, ...items) => {
    await js(`Array.from(document.querySelectorAll('.ae-menubar-title')).find((b) => b.textContent === ${JSON.stringify(menu)}).click(); true`)
    for (const item of items) {
      await waitFor(`Array.from(document.querySelectorAll('.ae-menu .ae-menu-item')).some((el) => el.querySelector('.ae-menu-label')?.textContent === ${JSON.stringify(item)})`, 4000, `menu item ${item}`)
      await js(`Array.from(document.querySelectorAll('.ae-menu .ae-menu-item')).find((el) => el.querySelector('.ae-menu-label')?.textContent === ${JSON.stringify(item)}).click(); true`)
    }
  }
  const step = async (name, run) => {
    const started = Date.now()
    try {
      result.steps[name] = { ...(await run()), ms: Date.now() - started }
    } catch (error) {
      result.steps[name] = { error: String(error && error.stack || error), ms: Date.now() - started }
      result.failures.push(`${name}: ${String(error && error.message || error)}`)
      throw error
    }
  }

  try {
    await window.loadFile(path.join(projectRoot, 'dist', 'index.html'))
    await sleep(200)
    // The page gets keyboard focus although the window stays hidden (fields then blur like for a user).
    window.webContents.focus()
    await js('window.__SIMPLE_IMAGE_QA__ = true; true')
    window.webContents.send('file:open-external', fixturePath)
    await waitFor(`document.querySelector('canvas')?.width === ${W} && document.querySelector('.advanced-button:not(:disabled)') !== null`, 20_000, 'Simple shows the image')

    // 1. Press Advanced.
    await step('enterAdvanced', async () => {
      const original = await simplePixel(160, 100)
      check(near(original, fixtureColor(160, 100), 1), `Simple shows the fixture (${original})`)
      await click('.advanced-button')
      await waitFor(`document.querySelector('.ae-shell .ae-layer-row') !== null && Boolean(document.querySelector('.ae-stage')?.dataset.view) && Boolean(window.__simpleAdvanced)`, 20_000, 'the Advanced editor')
      await settle()
      const opened = { layers: await layerNames(), history: await historyLabels(), saveState: await saveState(), panels: await js(`Array.from(document.querySelectorAll('.ae-panel h3')).map((h) => h.textContent.trim())`) }
      check(JSON.stringify(opened.layers) === JSON.stringify(['Background']), `one Background layer (${opened.layers})`)
      check(opened.saveState === 'Saved', `entering does not modify (${opened.saveState})`)
      check(JSON.stringify(opened.history) === JSON.stringify(['Advanced editor']), `history base entry (${opened.history})`)
      check(JSON.stringify(opened.panels) === JSON.stringify(['Layers', 'Properties', 'History', 'Color']), `four panels (${opened.panels})`)
      check(await js(`document.querySelector('.ae-lock-row') !== null && document.querySelector('.ae-cpanel') !== null && document.querySelector('.ae-history-panel') !== null`), 'the full panels are mounted (not the placeholders)')
      check(near(await displayPixel(160, 100), fixtureColor(160, 100), 3), 'the display shows the image')
      return opened
    })

    // 2. New layer, paint, Multiply.
    await step('newLayerPaintMultiply', async () => {
      await click('.ae-layers-footer button[aria-label="New layer"]', 'New layer button')
      await waitFor(`document.querySelectorAll('.ae-layer-row').length === 2`, 5000, 'the new layer')
      check(JSON.stringify(await layerNames()) === JSON.stringify(['Layer 1', 'Background']), `Layer 1 above the Background (${await layerNames()})`)
      // Foreground colour through the Color panel's hex field.
      await typeInto('.ae-cpanel-hex input', 'DC1E1E')
      await waitFor(`window.__simpleAdvanced.editor.getState().foreground.r === 220`, 3000, 'the foreground colour')
      await key({ key: 'b', code: 'KeyB' })
      await waitFor(`document.querySelector('.ae-tool.is-active')?.getAttribute('aria-label') === 'Brush Tool'`, 4000, 'the Brush tool')
      await dragOnCanvas([40, 60], [280, 60], 14)
      await waitFor(`Array.from(document.querySelectorAll('.ae-history-row')).some((row) => row.textContent.includes('Brush Tool'))`, 8000, 'the stroke step')
      let state = await docState()
      const painted = state.layers[1]
      check(painted.name === 'Layer 1' && painted.bounds && painted.bounds.width > 200, `the stroke is on Layer 1 (${JSON.stringify(painted.bounds)})`)
      check(state.layers[0].bounds && state.layers[0].bounds.width === W, 'the Background is untouched')
      // Multiply from the Layers panel's blend menu.
      await selectValue('.ae-layer-props select[aria-label="Blend mode"]', 'multiply')
      await waitFor(`window.__simpleAdvanced.store.getState().layers[1].blendMode === 'multiply'`, 3000, 'Multiply')
      state = await docState()
      await settle()
      const composite = await sample(160, 60)
      const base = fixtureColor(160, 60)
      const expected = [Math.round(base[0] * 220 / 255), Math.round(base[1] * 30 / 255), Math.round(base[2] * 30 / 255)]
      check(near(composite.slice(0, 3), expected, 3), `Multiply composites (got ${composite}, want about ${expected})`)
      check(await saveState() === 'Modified', 'painting marks the document Modified')
      return { history: await historyLabels(), composite, expected }
    })

    // 3. Levels adjustment layer with a mask from a marquee selection; edited in Properties.
    await step('levelsLayerWithMask', async () => {
      await key({ key: 'm', code: 'KeyM' })
      await waitFor(`document.querySelector('.ae-tool.is-active')?.getAttribute('aria-label') === 'Rectangular Marquee Tool'`, 4000, 'the Marquee tool')
      await dragOnCanvas([20, 100], [140, 180])
      await waitFor(`window.__simpleAdvanced.store.getState().selection !== null`, 4000, 'a selection')
      const selection = (await docState()).selection
      check(selection && Math.abs(selection.width - 120) <= 2 && Math.abs(selection.height - 80) <= 2, `marquee 120 x 80 (${JSON.stringify(selection)})`)
      await click('.ae-layers-footer button[aria-label="New adjustment layer"]', 'New adjustment layer button')
      await waitFor(`document.querySelector('.ae-popup-menu') !== null`, 3000, 'the adjustment menu')
      await js(`Array.from(document.querySelectorAll('.ae-popup-menu .ae-menu-item')).find((item) => item.textContent.includes('Levels')).click(); true`)
      await waitFor(`document.querySelectorAll('.ae-layer-row').length === 3`, 4000, 'the Levels layer')
      let state = await docState()
      const levels = state.layers[2]
      check(levels.kind === 'adjustment' && levels.name === 'Levels 1' && levels.mask, `Levels 1 with a mask (${JSON.stringify({ kind: levels.kind, name: levels.name, mask: levels.mask })})`)
      check(state.selection === null, 'the selection became the mask (deselected)')
      check(await js(`document.querySelector('.ae-layer-row.is-active .ae-mask-thumb') !== null`), 'the mask thumbnail shows in the row')
      // Properties shows the Levels editor; set the input black point.
      await waitFor(`document.querySelector('.ae-levels .ae-lv-fields .ae-num') !== null`, 4000, 'the Levels editor in Properties')
      const before = await sample(60, 150)
      await typeInto('.ae-levels .ae-lv-fields .ae-num', '90')
      await waitFor(`window.__simpleAdvanced.store.getState().layers[2].adjustment.rgb.inBlack === 90`, 3000, 'input black 90')
      await settle()
      const inside = await sample(60, 150)
      const outside = await sample(240, 150)
      check(inside[0] < before[0] || inside[1] < before[1], `the adjustment darkens inside its mask (${before} -> ${inside})`)
      check(near(outside, fixtureColor(240, 150), 1), `outside the mask nothing changes (${outside})`)
      check((await historyLabels()).includes('Modify Levels Layer'), 'the edit is a history step')
      // Histogram drawn behind the Levels editor.
      await waitFor(`(() => { const c = document.querySelector('.ae-levels .ae-histogram canvas'); if (!c) return false; const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; for (let i = 3; i < d.length; i += 4) if (d[i]) return true; return false })()`, 4000, 'the Levels histogram (of the image below the layer) is drawn')
      state = await docState()
      return { levels: state.layers[2].adjustment.rgb, before, inside, outside }
    })

    // 4. Marquee on the Background + Ctrl+J, then Free Transform.
    await step('layerViaCopyAndTransform', async () => {
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((row) => row.textContent.includes('Background')).click(); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().activeLayerId === window.__simpleAdvanced.store.getState().layers[0].id`, 3000, 'the Background is active')
      await dragOnCanvas([200, 20], [260, 70])
      await waitFor(`window.__simpleAdvanced.store.getState().selection !== null`, 3000, 'a selection on the Background')
      await key({ key: 'j', code: 'KeyJ', ctrlKey: true })
      await waitFor(`document.querySelectorAll('.ae-layer-row').length === 4`, 4000, 'Ctrl+J')
      let state = await docState()
      const copy = state.layers[1]
      check(copy.name === 'Layer 2' && copy.bounds && Math.abs(copy.bounds.width - 60) <= 2 && Math.abs(copy.bounds.x - 200) <= 2, `Layer 2 holds the copied 60 x 50 pixels (${JSON.stringify(copy)})`)
      // Ctrl+T: the options bar shows X Y W H angle; type a new width and commit with Enter.
      await key({ key: 't', code: 'KeyT', ctrlKey: true })
      await waitFor(`document.querySelector('.ae-options-group[aria-label="Free transform"]') !== null`, 4000, 'Free Transform')
      await typeInto('.ae-options-group[aria-label="Free transform"] label:nth-of-type(3) .ae-opt-num', '120')
      await sleep(150)
      await key({ key: 'Enter', code: 'Enter' })
      await waitFor(`document.querySelector('.ae-options-group[aria-label="Free transform"]') === null && Array.from(document.querySelectorAll('.ae-history-row')).some((row) => row.textContent.includes('Free Transform'))`, 15_000, 'the transform commits')
      state = await docState()
      const transformed = state.layers[1]
      check(transformed.bounds && transformed.bounds.width >= 110 && transformed.bounds.width <= 130, `the layer is about 120 px wide after the transform (${JSON.stringify(transformed.bounds)})`)
      return { copy: copy.bounds, transformed: transformed.bounds, history: await historyLabels() }
    })

    // 5. Undo / redo, History jump, snapshot.
    await step('historyAndSnapshot', async () => {
      const widthOf = async () => (await docState()).layers[1].bounds.width
      const transformedWidth = await widthOf()
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().layers[1].surface.contentBounds()?.width === 60`, 5000, 'undo the transform')
      await waitFor(`document.querySelector('.ae-history-row.is-current')?.textContent.trim() === 'Layer Via Copy'`, 3000, 'the History panel shows the undone step')
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true, shiftKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().layers[1].surface.contentBounds()?.width === ${transformedWidth}`, 5000, 'redo the transform')
      // Jump back to the stroke: two layers then; then jump to the newest state again.
      const labels = await historyLabels()
      const strokeIndex = labels.indexOf('Brush Tool')
      await js(`document.querySelectorAll('.ae-history-row')[${strokeIndex}].click(); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().layers.length === 2`, 5000, 'History jump to the stroke')
      const future = await waitFor(`document.querySelectorAll('.ae-history-row.is-future').length === ${labels.length - 1 - strokeIndex} && document.querySelectorAll('.ae-history-row.is-future').length`, 3000, 'later states are dimmed, not removed')
      await js(`Array.from(document.querySelectorAll('.ae-history-row')).at(-1).click(); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().layers.length === 4`, 5000, 'History jump back to the newest state')
      // Snapshot, a change, then revert to the snapshot as one undoable step.
      await click('.ae-history-footer button[aria-label="New snapshot"]', 'New snapshot')
      await waitFor(`document.querySelectorAll('.ae-snapshot').length === 1`, 4000, 'the snapshot')
      await js(`document.querySelector('.ae-layer-row .ae-eye').click(); true`)
      await waitFor(`!window.__simpleAdvanced.store.getState().layers[3].visible`, 3000, 'hide the top layer')
      await click('.ae-snapshot-main', 'Snapshot 1')
      await waitFor(`window.__simpleAdvanced.store.getState().layers.length === 4 && window.__simpleAdvanced.store.getState().layers[3].visible`, 5000, 'revert to the snapshot')
      check((await historyLabels()).at(-1) === 'Revert to Snapshot 1', `the revert is a step (${(await historyLabels()).at(-1)})`)
      const names = (await docState()).layers.map((layer) => layer.name)
      check(JSON.stringify(names) === JSON.stringify(['Background', 'Layer 2', 'Layer 1', 'Levels 1']), `the snapshot restores every layer (${names})`)
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true })
      await waitFor(`!window.__simpleAdvanced.store.getState().layers[3].visible`, 3000, 'undo the revert')
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().layers[3].visible`, 3000, 'undo the hide')
      return { labels: await historyLabels(), current: await historyCurrent() }
    })

    // 6. Panels and dialogs.
    await step('colorPicker', async () => {
      await click('.ae-toolbox .ae-chip.is-foreground', 'the foreground chip')
      await waitFor(`document.querySelector('.ae-cp-popover') !== null`, 3000, 'the colour picker')
      // Drag in the saturation / brightness field: the colour follows live; Escape puts the original back.
      await js(`(() => {
        const field = document.querySelector('.ae-cp-popover .ae-cp-field-area')
        const box = field.getBoundingClientRect()
        const opts = (fx, fy, buttons) => ({ bubbles: true, cancelable: true, pointerId: 11, pointerType: 'mouse', isPrimary: true, clientX: box.left + box.width * fx, clientY: box.top + box.height * fy, button: 0, buttons })
        field.dispatchEvent(new PointerEvent('pointerdown', opts(0.2, 0.8, 1)))
        field.dispatchEvent(new PointerEvent('pointermove', opts(0.1, 0.9, 1)))
        field.dispatchEvent(new PointerEvent('pointerup', opts(0.1, 0.9, 0)))
        return true
      })()`)
      const dragged = await waitFor(`(() => { const c = window.__simpleAdvanced.editor.getState().foreground; return c.r !== 220 || c.g !== 30 ? [c.r, c.g, c.b] : null })()`, 3000, 'the field changes the colour')
      check(dragged[0] < 60 && dragged[1] < 60, `a low-brightness point gives a dark colour (${dragged})`)
      await js(`document.querySelector('.ae-cp-popover').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })); true`)
      await waitFor(`document.querySelector('.ae-cp-popover') === null && window.__simpleAdvanced.editor.getState().foreground.r === 220 && window.__simpleAdvanced.editor.getState().foreground.g === 30`, 3000, 'Escape restores the colour')
      // A typed hex value and Enter keep the new colour and close the picker.
      await click('.ae-toolbox .ae-chip.is-foreground', 'the foreground chip')
      await waitFor(`document.querySelector('.ae-cp-popover') !== null`, 3000, 'the colour picker again')
      await typeInto('.ae-cp-popover .ae-cp-hex input', '00FF00')
      await waitFor(`document.querySelector('.ae-cp-popover') === null && window.__simpleAdvanced.editor.getState().foreground.g === 255 && window.__simpleAdvanced.editor.getState().foreground.r === 0`, 3000, 'hex and Enter set the colour')
      const picked = await waitFor(`document.querySelector('.ae-toolbox .ae-chip.is-foreground').dataset.color === '#00ff00' && document.querySelector('.ae-toolbox .ae-chip.is-foreground').dataset.color`, 3000, 'the chip shows the colour')
      const recent = await js(`window.__simpleAdvanced.editor.getState().recentColors.map((c) => [c.r, c.g, c.b])`)
      check(recent.length > 0 && recent[0][1] === 255, `the picked colour is first in the recent colours (${JSON.stringify(recent)})`)
      const eyeDropper = await js(`Boolean(window.EyeDropper) && document.querySelector('.ae-cpanel-pick') !== null`)
      check(eyeDropper, 'pick from screen (EyeDropper API) is offered in the Color panel')
      return { dragged, picked, eyeDropper }
    })

    await step('imageSizeDialog', async () => {
      await key({ key: 'i', code: 'KeyI', ctrlKey: true, altKey: true })
      await waitFor(`document.querySelector('.ae-dframe h2')?.textContent === 'Image Size'`, 4000, 'the Image Size dialog')
      await typeInto('#ae-is-width', '160')
      await waitFor(`document.querySelector('.ae-dframe') === null && window.__simpleAdvanced.store.getState().width === 160`, 15_000, 'Image Size applies')
      const size = await js(`[window.__simpleAdvanced.store.getState().width, window.__simpleAdvanced.store.getState().height]`)
      check(size[0] === 160 && size[1] === 100, `constrained proportions give 160 x 100 (${size})`)
      check((await historyLabels()).at(-1) === 'Image Size', 'Image Size is one step')
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().width === ${W}`, 5000, 'undo Image Size')
      // Canvas Size opens with the anchor grid; Escape cancels.
      await key({ key: 'c', code: 'KeyC', ctrlKey: true, altKey: true })
      await waitFor(`document.querySelector('.ae-dframe h2')?.textContent === 'Canvas Size' && document.querySelectorAll('.ae-anchor button').length === 9`, 4000, 'the Canvas Size dialog')
      await js(`document.querySelector('.ae-dframe').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true })); true`)
      await waitFor(`document.querySelector('.ae-dframe') === null`, 3000, 'Escape closes Canvas Size')
      return { size }
    })

    await step('adjustmentDialogPreview', async () => {
      // Destructive Levels (Ctrl+L) on the whole Background: the canvas previews, Cancel changes nothing.
      await key({ key: 'd', code: 'KeyD', ctrlKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().selection === null`, 3000, 'Deselect')
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((row) => row.textContent.includes('Background')).click(); true`)
      await settle()
      const historyBefore = (await historyLabels()).length
      const before = await displayPixel(300, 190)
      await key({ key: 'l', code: 'KeyL', ctrlKey: true })
      await waitFor(`document.querySelector('.ae-dframe h2')?.textContent === 'Levels'`, 4000, 'the Levels dialog')
      await waitFor(`(() => { const c = document.querySelector('.ae-dframe .ae-histogram canvas'); if (!c) return false; const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; for (let i = 3; i < d.length; i += 4) if (d[i]) return true; return false })()`, 4000, 'the dialog histogram')
      // Shift+Up raises the input black point by 10 (one key press per task, as typing would arrive).
      for (let i = 0; i < 12; i += 1) {
        await js(`(() => { const input = document.querySelector('.ae-dframe .ae-lv-fields .ae-num'); input.focus(); input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowUp', code: 'ArrowUp', shiftKey: true, bubbles: true, cancelable: true })); return true })()`)
        await sleep(15)
      }
      await waitFor(`document.querySelector('.ae-dframe .ae-lv-fields .ae-num').value === '120'`, 3000, 'input black 120')
      await sleep(120)
      await settle()
      const previewed = await displayPixel(300, 190)
      check(previewed[0] < before[0] - 20, `the canvas previews the adjustment (${before} -> ${previewed})`)
      check((await historyLabels()).length === historyBefore, 'the preview is not a history step')
      await click('.ae-dframe .ae-dframe-cancel', 'Cancel')
      await waitFor(`document.querySelector('.ae-dframe') === null`, 3000, 'Cancel closes the dialog')
      await settle()
      const after = await displayPixel(300, 190)
      check(near(after, before, 2), `Cancel restores the canvas (${after} vs ${before})`)
      return { before, previewed, after }
    })

    await step('filterDialog', async () => {
      await menuClick('Filter', 'Blur', 'Gaussian Blur…')
      await waitFor(`document.querySelector('.ae-dframe h2')?.textContent === 'Gaussian Blur'`, 4000, 'the Gaussian Blur dialog')
      await waitFor(`(() => { const c = document.querySelector('.ae-filter-preview canvas'); if (!c) return false; const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; for (let i = 3; i < d.length; i += 4) if (d[i]) return true; return false })()`, 8000, 'the 100% preview')
      await typeInto('.ae-dframe .ae-sf .ae-num', '4')
      await waitFor(`document.querySelector('.ae-dframe') === null && Array.from(document.querySelectorAll('.ae-history-row')).some((row) => row.textContent.includes('Gaussian Blur'))`, 15_000, 'the blur applies')
      const label = (await historyLabels()).at(-1)
      const last = await js(`window.__simpleAdvanced.editor.getState().lastFilter`)
      check(last && last.type === 'gaussian-blur' && last.radius === 4, `Last Filter remembers the radius (${JSON.stringify(last)})`)
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true })
      await waitFor(`!Array.from(document.querySelectorAll('.ae-history-row:not(.is-future)')).some((row) => row.textContent.includes('Gaussian Blur'))`, 4000, 'undo the blur')
      return { label }
    })

    await step('featherSelection', async () => {
      await key({ key: 'a', code: 'KeyA', ctrlKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().selection !== null`, 3000, 'Select All')
      await key({ key: 'F6', code: 'F6', shiftKey: true })
      await waitFor(`document.querySelector('.ae-dframe h2')?.textContent === 'Feather Selection'`, 4000, 'the Feather dialog')
      await typeInto('#ae-modify-amount', '5')
      await waitFor(`document.querySelector('.ae-dframe') === null && Array.from(document.querySelectorAll('.ae-history-row')).some((row) => row.textContent.trim() === 'Feather')`, 10_000, 'Feather applies')
      await key({ key: 'd', code: 'KeyD', ctrlKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().selection === null`, 3000, 'Deselect')
      return {}
    })

    await step('layersPanel', async () => {
      // Rename by double-click.
      await js(`(() => { const name = Array.from(document.querySelectorAll('.ae-layer-row .ae-layer-name')).find((el) => el.textContent.trim() === 'Layer 2'); name.dispatchEvent(new MouseEvent('dblclick', { bubbles: true })); return true })()`)
      await waitFor(`document.querySelector('.ae-rename') !== null`, 3000, 'the rename field')
      await typeInto('.ae-rename', 'Copied patch')
      await waitFor(`window.__simpleAdvanced.store.getState().layers[1].name === 'Copied patch'`, 3000, 'rename')
      // Drag "Copied patch" above "Layer 1".
      const order = async () => (await docState()).layers.map((layer) => layer.name)
      await js(`(() => {
        const rows = Array.from(document.querySelectorAll('.ae-layer-row'))
        const row = rows.find((el) => el.textContent.includes('Copied patch'))
        const target = rows.find((el) => el.textContent.includes('Layer 1'))
        const from = row.querySelector('.ae-layer-name').getBoundingClientRect()
        const to = target.getBoundingClientRect()
        const opts = (x, y, buttons) => ({ bubbles: true, cancelable: true, pointerId: 9, pointerType: 'mouse', isPrimary: true, clientX: x, clientY: y, button: 0, buttons })
        row.querySelector('.ae-layer-name').dispatchEvent(new PointerEvent('pointerdown', opts(from.left + 10, from.top + 5, 1)))
        window.dispatchEvent(new PointerEvent('pointermove', opts(from.left + 10, from.top - 10, 1)))
        window.dispatchEvent(new PointerEvent('pointermove', opts(from.left + 10, to.top + 3, 1)))
        window.dispatchEvent(new PointerEvent('pointerup', opts(from.left + 10, to.top + 3, 0)))
        return true
      })()`)
      await waitFor(`window.__simpleAdvanced.store.getState().layers[2].name === 'Copied patch'`, 3000, 'drag to reorder')
      check((await historyLabels()).at(-1) === 'Layer Order', 'reordering is a "Layer Order" step')
      // Lock all, then unlock with the badge.
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((el) => el.textContent.includes('Copied patch')).click(); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().activeLayerId === window.__simpleAdvanced.store.getState().layers[2].id`, 3000, 'select Copied patch')
      await click('.ae-lock-row button[aria-label="Lock all"]', 'Lock all')
      await waitFor(`(() => { const l = window.__simpleAdvanced.store.getState().layers[2].locks; return l.pixels && l.position && l.transparency })()`, 3000, 'lock all')
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((el) => el.textContent.includes('Copied patch')).querySelector('.ae-lock-badge').click(); true`)
      await waitFor(`(() => { const l = window.__simpleAdvanced.store.getState().layers[2].locks; return !l.pixels && !l.position && !l.transparency })()`, 3000, 'unlock with the badge')
      // Ctrl+click the thumbnail: the layer's pixels become the selection.
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((el) => el.textContent.includes('Copied patch')).querySelector('.ae-thumb-wrap').dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, ctrlKey: true })); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().selection !== null`, 3000, 'Ctrl+click loads the selection')
      const selected = (await docState()).selection
      const patch = (await docState()).layers[2].bounds
      check(selected && patch && Math.abs(selected.width - patch.width) <= 1 && Math.abs(selected.x - patch.x) <= 1, `the selection matches the layer pixels (${JSON.stringify(selected)} vs ${JSON.stringify(patch)})`)
      await key({ key: 'd', code: 'KeyD', ctrlKey: true })
      // Context menu: Merge Down onto Layer 1, then undo it.
      await js(`(() => { const row = Array.from(document.querySelectorAll('.ae-layer-row')).find((el) => el.textContent.includes('Copied patch')); const box = row.getBoundingClientRect(); row.dispatchEvent(new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: box.left + 40, clientY: box.top + 10 })); return true })()`)
      await waitFor(`document.querySelector('.ae-popup-menu[aria-label="Layer"]') !== null`, 3000, 'the layer context menu')
      const items = await js(`Array.from(document.querySelectorAll('.ae-popup-menu .ae-menu-item')).map((item) => ({ label: item.querySelector('.ae-menu-label').textContent, keys: item.querySelector('.ae-menu-keys').textContent, enabled: item.getAttribute('aria-disabled') !== 'true' }))`)
      check(items.some((item) => item.label === 'Merge Down' && item.enabled && item.keys === 'Ctrl+E'), `Merge Down in the context menu (${JSON.stringify(items.map((item) => item.label))})`)
      await js(`Array.from(document.querySelectorAll('.ae-popup-menu .ae-menu-item')).find((item) => item.querySelector('.ae-menu-label').textContent === 'Merge Down').click(); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().layers.length === 3`, 4000, 'Merge Down')
      check((await order())[1] === 'Layer 1', `Merge Down keeps the lower layer's name (${await order()})`)
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true })
      await waitFor(`window.__simpleAdvanced.store.getState().layers.length === 4`, 4000, 'undo Merge Down')
      // Alt+click an eye shows only that layer; Alt+click again brings the others back.
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((el) => el.textContent.includes('Background')).querySelector('.ae-eye').dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, altKey: true })); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().layers.filter((layer) => layer.visible).length === 1`, 3000, 'Alt+click solos')
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((el) => el.textContent.includes('Background')).querySelector('.ae-eye').dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, altKey: true })); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().layers.every((layer) => layer.visible)`, 3000, 'Alt+click again restores')
      return { order: await order(), contextMenu: items.map((item) => item.label) }
    })

    // Every adjustment type's controls in Properties change their layer (then the layer is removed again).
    await step('adjustmentTypes', async () => {
      const run = (command) => js(`window.__simpleAdvanced.run(${JSON.stringify(command)}); true`)
      // New adjustment layers go above the active layer (Photoshop), so read the active one.
      const top = () => js(`(() => { const state = window.__simpleAdvanced.store.getState(); const layer = state.layers.find((entry) => entry.id === state.activeLayerId); return layer && layer.kind === 'adjustment' ? layer.adjustment : null })()`)
      const firstNumber = '.ae-properties-panel .ae-sf .ae-num'
      const cases = [
        ['brightness-contrast', 'Brightness/Contrast', () => typeInto(firstNumber, '30'), (a) => a.brightness === 30],
        ['curves', 'Curves', () => js(`(() => { const svg = document.querySelector('.ae-properties-panel .ae-curves-graph svg'); const box = svg.getBoundingClientRect(); const o = (fx, fy, b) => ({ bubbles: true, cancelable: true, pointerId: 5, pointerType: 'mouse', isPrimary: true, clientX: box.left + box.width * fx, clientY: box.top + box.height * fy, button: 0, buttons: b }); svg.dispatchEvent(new PointerEvent('pointerdown', o(0.5, 0.3, 1))); svg.dispatchEvent(new PointerEvent('pointerup', o(0.5, 0.3, 0))); return true })()`), (a) => a.rgb.length === 3 && a.rgb[1].y > 150],
        ['exposure', 'Exposure', () => typeInto(firstNumber, '1.5'), (a) => a.exposure === 1.5],
        ['vibrance', 'Vibrance', () => typeInto(firstNumber, '40'), (a) => a.vibrance === 40],
        ['hue-saturation', 'Hue/Saturation', () => js(`Array.from(document.querySelectorAll('.ae-properties-panel .ae-check')).find((el) => el.textContent.includes('Colorize')).querySelector('input').click(); true`), (a) => a.colorize === true && a.master.saturation === 25],
        ['color-balance', 'Color Balance', async () => {
          await js(`Array.from(document.querySelectorAll('.ae-properties-panel .ae-seg button')).find((el) => el.textContent === 'Shadows').click(); true`)
          await sleep(60)
          await typeInto(firstNumber, '20')
        }, (a) => a.shadows.cyanRed === 20 && a.midtones.cyanRed === 0],
        ['black-white', 'Black & White', () => js(`Array.from(document.querySelectorAll('.ae-properties-panel .ae-check')).find((el) => el.textContent.includes('Tint')).querySelector('input').click(); true`), (a) => a.tint !== null],
        ['photo-filter', 'Photo Filter', () => selectValue('.ae-properties-panel .ae-field select', 'cooling-80'), (a) => a.color.r === 0 && a.color.g === 109 && a.color.b === 255],
        ['posterize', 'Posterize', () => typeInto(firstNumber, '8'), (a) => a.levels === 8],
        ['threshold', 'Threshold', () => typeInto(firstNumber, '100'), (a) => a.level === 100],
        ['gradient-map', 'Gradient Map', () => js(`(() => { const strip = document.querySelector('.ae-properties-panel .ae-stops-strip'); const box = strip.getBoundingClientRect(); strip.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, pointerId: 6, pointerType: 'mouse', isPrimary: true, clientX: box.left + box.width * 0.5, clientY: box.top + 8, button: 0, buttons: 1 })); return true })()`), (a) => a.stops.length === 3],
        ['invert', 'Invert', async () => {}, () => Boolean(true)],
      ]
      const done = []
      for (const [type, title, act, verify] of cases) {
        await run(`adjustment-layer.${type}`)
        await waitFor(`document.querySelector('.ae-properties-panel .ae-prop-name')?.textContent === ${JSON.stringify(title)}`, 4000, `${title} in Properties`)
        await act()
        let ok = false
        for (let i = 0; i < 40 && !ok; i += 1) {
          const adjustment = await top()
          ok = Boolean(adjustment && adjustment.type === type && verify(adjustment))
          if (!ok) await sleep(50)
        }
        check(ok, `${title}: the Properties control changes the adjustment layer (${JSON.stringify(await top())})`)
        if (type === 'invert') check(await js(`document.querySelector('.ae-properties-panel .ae-note') !== null`), 'Invert explains that it has no settings')
        if (type === 'threshold') {
          await waitFor(`(() => { const c = document.querySelector('.ae-properties-panel .ae-histogram canvas'); if (!c) return false; const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; for (let i = 3; i < d.length; i += 4) if (d[i]) return true; return false })()`, 4000, 'the Threshold histogram')
        }
        done.push(type)
        await run('layer.delete')
        await waitFor(`!window.__simpleAdvanced.store.getState().layers.some((layer) => layer.kind === 'adjustment' && layer.adjustment.type === ${JSON.stringify(type)})`, 3000, `remove the ${title} layer`)
      }
      check(JSON.stringify((await docState()).layers.map((layer) => layer.name)) === JSON.stringify(['Background', 'Layer 1', 'Copied patch', 'Levels 1']), 'the document is back to its four layers')
      return { done }
    })

    // 7. Save writes a PSD with the layers; a value typed but not confirmed with Enter is saved too.
    await step('savePsd', async () => {
      await js(`Array.from(document.querySelectorAll('.ae-layer-row')).find((row) => row.textContent.includes('Layer 1')).click(); true`)
      await waitFor(`window.__simpleAdvanced.store.getState().activeLayerId === window.__simpleAdvanced.store.getState().layers[1].id`, 3000, 'select Layer 1')
      await typeInto('.ae-layer-props .ae-num', '65', false)
      check(await js(`document.activeElement === document.querySelector('.ae-layer-props .ae-num')`), 'the opacity field has the keyboard (typed, not confirmed)')
      const before = saves.length
      await key({ key: 's', code: 'KeyS', ctrlKey: true })
      await waitFor(`document.querySelector('.save-state')?.textContent === 'Saved'`, 30_000, 'the layered document saves')
      const written = saves.slice(before)
      check(written.length === 1 && written[0].format === 'psd', `Save writes a PSD (${JSON.stringify(written.map((save) => save.format))})`)
      const summary = psdSummary(written[0].bytes)
      const names = summary.layers.map((layer) => layer.name)
      check(summary.width === W && summary.height === H, `PSD size ${summary.width} x ${summary.height}`)
      check(JSON.stringify(names) === JSON.stringify(['Background', 'Layer 1', 'Copied patch', 'Levels 1']), `PSD layers bottom to top (${names})`)
      const byName = Object.fromEntries(summary.layers.map((layer) => [layer.name, layer]))
      check(byName['Layer 1'] && byName['Layer 1'].blendMode === 'multiply', `Layer 1 is Multiply in the PSD (${byName['Layer 1'] && byName['Layer 1'].blendMode})`)
      check(byName['Layer 1'] && Math.abs(byName['Layer 1'].opacity - 0.65) <= 1 / 255, `the typed opacity 65% was committed before saving (${byName['Layer 1'] && byName['Layer 1'].opacity})`)
      check(byName['Levels 1'] && byName['Levels 1'].adjustment === 'levels' && byName['Levels 1'].mask, 'Levels 1 is a masked levels adjustment layer in the PSD')
      return { summary }
    })

    // 8. Export As PNG; Print opens and is cancelled.
    let exported = null
    await step('exportAndPrint', async () => {
      const before = saves.length
      await key({ key: 'w', code: 'KeyW', ctrlKey: true, altKey: true, shiftKey: true })
      await waitFor(`document.querySelector('.export-menu button[data-export-format="png"]') !== null`, 4000, 'the Export menu')
      const formats = await js(`Array.from(document.querySelectorAll('.export-menu [data-export-format]')).map((b) => b.dataset.exportFormat)`)
      check(formats.includes('psd'), `Export offers PSD in Advanced (${formats})`)
      await click('.export-menu button[data-export-format="png"]', 'PNG export')
      await waitFor(`document.querySelector('.toast')?.textContent.includes('Exported')`, 20_000, 'the export finishes')
      const written = saves.slice(before)
      check(written.length === 1 && written[0].format === 'png' && written[0].purpose === 'export', `Export As PNG (${JSON.stringify(written.map((save) => [save.format, save.purpose]))})`)
      exported = written[0]
      const dims = imageDimensions(written[0].bytes, '.png')
      check(dims && dims.width === W && dims.height === H, `the exported PNG is ${W} x ${H} (${JSON.stringify(dims)})`)
      await key({ key: 'p', code: 'KeyP', ctrlKey: true })
      await waitFor(`document.querySelector('.image-print-dialog') !== null`, 15_000, 'the print dialog')
      await js(`Array.from(document.querySelectorAll('.image-print-dialog button')).find((b) => b.textContent.trim() === 'Cancel').click(); true`)
      await waitFor(`document.querySelector('.image-print-dialog') === null`, 4000, 'Cancel closes the print dialog')
      check(printJobs.length === 0, 'no print job was sent')
      return { formats, dims }
    })

    // 9. Back to Simple with the flatten prompt; Simple holds the composite; Ctrl+Z restores the original.
    await step('flattenToSimple', async () => {
      await click('.ae-simple', 'the Simple button')
      await waitFor(`document.querySelector('.advanced-dialog') !== null`, 5000, 'the flatten prompt')
      await js(`Array.from(document.querySelectorAll('.advanced-dialog button')).find((b) => b.textContent.trim() === 'Flatten').click(); true`)
      await waitFor(`document.querySelector('.advanced-root') === null && document.querySelector('.toolbar').hidden === false`, 15_000, 'back in Simple')
      const comparison = await js(`(async () => {
        const bytes = new Uint8Array(${JSON.stringify(Array.from(exported.bytes))})
        const bitmap = await createImageBitmap(new Blob([bytes], { type: 'image/png' }), { premultiplyAlpha: 'none', colorSpaceConversion: 'none' })
        const scratch = new OffscreenCanvas(bitmap.width, bitmap.height)
        const sctx = scratch.getContext('2d')
        sctx.drawImage(bitmap, 0, 0)
        const a = sctx.getImageData(0, 0, bitmap.width, bitmap.height).data
        const simple = document.querySelector('canvas')
        const b = simple.getContext('2d').getImageData(0, 0, simple.width, simple.height).data
        let max = 0
        for (let i = 0; i < a.length; i += 1) max = Math.max(max, Math.abs(a[i] - b[i]))
        return { width: simple.width, height: simple.height, maxDifference: max }
      })()`)
      check(comparison.width === W && comparison.height === H && comparison.maxDifference <= 1, `the Simple canvas equals the flattened composite (${JSON.stringify(comparison)})`)
      const stateAfterExit = await saveState()
      await key({ key: 'z', code: 'KeyZ', ctrlKey: true })
      await waitFor(`(() => { const d = document.querySelector('canvas').getContext('2d').getImageData(160, 60, 1, 1).data; return Math.abs(d[0] - ${fixtureColor(160, 60)[0]}) <= 1 && Math.abs(d[1] - ${fixtureColor(160, 60)[1]}) <= 1 })()`, 5000, 'Ctrl+Z in Simple restores the pre-Advanced image')
      const restored = await simplePixel(160, 60)
      check(near(restored, fixtureColor(160, 60), 1), `pre-Advanced pixels are back (${restored})`)
      return { comparison, restored, stateAfterExit }
    })
  } catch (error) {
    result.error = String(error && error.stack || error)
    try { fs.writeFileSync(path.join(workDir, 'advanced-smoke-failure.png'), (await window.webContents.capturePage()).toPNG()) } catch {}
  }
  try { fs.writeFileSync(path.join(workDir, 'advanced-smoke.png'), (await window.webContents.capturePage()).toPNG()) } catch {}
  result.saves = saves.map((save) => ({ format: save.format, purpose: save.purpose, bytes: save.bytes.length }))
  result.printJobs = printJobs.length
  // Console errors are reported; React development warnings do not appear in the production build.
  fs.writeFileSync(path.join(workDir, 'advanced-smoke-result.json'), JSON.stringify(result, (key, value) => (key === 'bytes' && typeof value === 'object' ? undefined : value), 2), 'utf8')
  const ok = !result.error && result.failures.length === 0
  if (!ok) {
    console.error(`advanced-smoke failed:\n- ${[...result.failures, ...(result.error ? [result.error] : [])].join('\n- ')}\nArtifacts: ${workDir}`)
    app.exit(1)
    return
  }
  console.log(`advanced-smoke passed (${Object.keys(result.steps).length} steps). Artifacts: ${workDir}`)
  app.exit(0)
})
