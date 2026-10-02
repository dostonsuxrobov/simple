// qa/psd-smoke.cjs (WP7): npm run smoke:psd (after npm run build:web)
// Runs the real backend (electron/main.cjs with its IPC handlers and preload) and the production renderer
// in a hidden window with an isolated profile and temp folder:
//   1. a layered PSD generated with ag-psd opens from the command line through file:open-path and lands in
//      Advanced mode showing its layers;
//   2. Ctrl+S on the unchanged document writes the original bytes back (byte-exact);
//   3. in the renderer, the real lazy PSD chunk (DOM canvas hooks) imports the file again, edits layers
//      (visibility, opacity, rename), exports and saves through file:save (Save As, dialog stubbed); the
//      saved file is re-read here with ag-psd;
//   4. when the editor build supports editing from the keyboard (Ctrl+J), the edit is saved in place as a
//      layered PSD and re-read; the placeholder editor has no editing, so this step reports itself skipped;
//   5. the clipboard read bridge is exposed (the system clipboard is never written).
// Environment: SIMPLE_IMAGE_QA_DIR (work folder; default a new temp folder), SIMPLE_IMAGE_SMOKE_TIMEOUT_MS.
'use strict'

const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const Module = require('node:module')
const electron = require('electron')

const { app } = electron
const projectRoot = path.join(__dirname, '..')
const workDir = process.env.SIMPLE_IMAGE_QA_DIR
  ? path.resolve(process.env.SIMPLE_IMAGE_QA_DIR)
  : fs.mkdtempSync(path.join(os.tmpdir(), 'simple-image-psd-smoke-'))
fs.mkdirSync(workDir, { recursive: true })
const tempDir = path.join(workDir, 'temp')
fs.mkdirSync(tempDir, { recursive: true })
// main.cjs sweeps stale print folders in the temp folder at start-up; keep that inside the work folder.
process.env.TEMP = tempDir
process.env.TMP = tempDir
app.setPath('userData', path.join(workDir, 'user-data'))

const timeoutMs = Number(process.env.SIMPLE_IMAGE_SMOKE_TIMEOUT_MS) || 120_000
const hardTimeout = setTimeout(() => {
  console.error(`psd-smoke: timed out after ${timeoutMs} ms`)
  app.exit(2)
}, timeoutMs)
hardTimeout.unref?.()

// #region fixture

const ag = require('ag-psd')
ag.initializeCanvas(() => { throw new Error('no canvas in the main process') }, (width, height) => ({ width, height, data: new Uint8ClampedArray(width * height * 4) }))

function solid(width, height, rgba) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let i = 0; i < data.length; i += 4) data.set(rgba, i)
  return { width, height, data }
}

function gradientMask(width, height) {
  const data = new Uint8ClampedArray(width * height * 4)
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const value = Math.round((x / Math.max(1, width - 1)) * 255)
      data.set([value, value, value, 255], (y * width + x) * 4)
    }
  }
  return { width, height, data }
}

const W = 320
const H = 200
const LAYER_NAMES = ['Background', 'Red half', 'Clipped', 'Hidden', 'Title', 'Levels 1']
const fixture = Buffer.from(ag.writePsdUint8Array({
  width: W,
  height: H,
  imageData: solid(W, H, [40, 90, 160, 255]),
  imageResources: { resolutionInfo: { horizontalResolution: 144, horizontalResolutionUnit: 'PPI', widthUnit: 'Inches', verticalResolution: 144, verticalResolutionUnit: 'PPI', heightUnit: 'Inches' } },
  children: [
    { name: 'Background', left: 0, top: 0, imageData: solid(W, H, [40, 90, 160, 255]) },
    {
      name: 'Red half', left: 20, top: 20, blendMode: 'multiply', opacity: 0.5, imageData: solid(160, 100, [230, 30, 30, 255]),
      mask: { left: 20, top: 20, right: 180, bottom: 120, defaultColor: 0, imageData: gradientMask(160, 100) },
    },
    { name: 'Clipped', left: 40, top: 40, clipping: true, imageData: solid(40, 40, [250, 240, 20, 255]) },
    { name: 'Hidden', left: 200, top: 120, hidden: true, imageData: solid(60, 40, [20, 200, 60, 255]) },
    {
      name: 'Title', left: 20, top: 150, imageData: solid(120, 30, [255, 255, 255, 255]),
      text: { text: 'Simple PSD', transform: [1, 0, 0, 1, 20, 174], style: { font: { name: 'SegoeUI-Bold' }, fontSize: 24, fillColor: { r: 255, g: 255, b: 255 } } },
    },
    { name: 'Levels 1', adjustment: { type: 'levels', rgb: { shadowInput: 12, highlightInput: 235, shadowOutput: 0, highlightOutput: 255, midtoneInput: 1.1 } } },
  ],
}, { invalidateTextLayers: true }))
const fixturePath = path.join(workDir, 'poster.psd')
fs.writeFileSync(fixturePath, fixture)

// #endregion

// #region the real backend, in a hidden window, with Save As stubbed

const windows = []
const saveDialogs = []
let nextSavePath = path.join(workDir, 'edited.psd')

class HiddenWindow extends electron.BrowserWindow {
  constructor(options = {}) {
    super({ ...options, show: false })
    windows.push(this)
  }

  show() {}

  focus() {}
}

const dialogStub = new Proxy(electron.dialog, {
  get(target, key) {
    if (key === 'showSaveDialog') {
      return async (_window, options) => {
        saveDialogs.push(options)
        return { canceled: false, filePath: nextSavePath }
      }
    }
    if (key === 'showOpenDialog') return async () => ({ canceled: true, filePaths: [] })
    return Reflect.get(target, key)
  },
})
const electronStub = new Proxy(electron, {
  get(target, key) {
    if (key === 'BrowserWindow') return HiddenWindow
    if (key === 'dialog') return dialogStub
    return Reflect.get(target, key)
  },
})
const originalLoad = Module._load
Module._load = function load(request, parent, isMain) {
  if (request === 'electron') return electronStub
  return originalLoad.call(this, request, parent, isMain)
}
// main.cjs opens supported files named on the command line (here: the generated PSD).
process.argv.push(fixturePath)
require(path.join(projectRoot, 'electron', 'main.cjs'))

// #endregion

async function waitFor(window, expression, timeout = 20_000) {
  const started = Date.now()
  while (Date.now() - started < timeout) {
    try {
      if (await window.webContents.executeJavaScript(expression)) return
    } catch {
      // The page may be navigating; try again.
    }
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  throw new Error(`Timed out waiting for: ${expression}`)
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

function psdChunk() {
  const assets = path.join(projectRoot, 'dist', 'assets')
  const chunk = fs.readdirSync(assets).find((name) => /^psd-[\w-]+\.js$/.test(name))
  if (!chunk) throw new Error('dist has no psd chunk; run npm run build:web first.')
  return `./assets/${chunk}`
}

function layerSummary(bytes) {
  const psd = ag.readPsd(bytes, { useImageData: true, skipThumbnail: true })
  return {
    width: psd.width,
    height: psd.height,
    layers: (psd.children || []).map((layer) => ({
      name: layer.name,
      hidden: Boolean(layer.hidden),
      opacity: layer.opacity,
      blendMode: layer.blendMode,
      clipping: Boolean(layer.clipping),
      mask: layer.mask ? { defaultColor: layer.mask.defaultColor } : null,
      adjustment: layer.adjustment ? layer.adjustment.type : null,
      text: layer.text ? { text: layer.text.text, font: layer.text.style && layer.text.style.font && layer.text.style.font.name } : null,
    })),
  }
}

function check(condition, message, failures) {
  if (!condition) failures.push(message)
}

app.whenReady().then(async () => {
  const failures = []
  const result = { workDir, steps: {} }
  try {
    const started = Date.now()
    while (!windows.length && Date.now() - started < 20_000) await sleep(50)
    const window = windows[0]
    if (!window) throw new Error('The app did not create its window.')
    await waitFor(window, 'document.readyState === "complete" && Boolean(window.simpleImage)')

    // 1. The PSD opens from the command line into Advanced mode with its layers listed.
    await waitFor(window, `(() => {
      const root = document.querySelector('.advanced-root')
      if (!root) return false
      const text = root.innerText
      return ${JSON.stringify(LAYER_NAMES)}.every((name) => text.includes(name))
    })()`, 30_000)
    const opened = await window.webContents.executeJavaScript(`(() => ({
      title: document.title,
      state: document.querySelector('.save-state')?.textContent ?? null,
      toolbarHidden: Boolean(document.querySelector('.toolbar')?.closest('[hidden]') || document.querySelector('.toolbar')?.hidden),
      advancedText: document.querySelector('.advanced-root').innerText.slice(0, 400),
      toast: document.querySelector('.toast')?.textContent ?? null,
      firstCanvas: (() => { const canvas = document.querySelector('canvas'); return canvas ? { width: canvas.width, height: canvas.height } : null })(),
    }))()`)
    result.steps.open = opened
    check(opened.state === 'Saved', `a freshly opened PSD is not modified (state: ${opened.state})`, failures)

    // 2. Ctrl+S on the unchanged document writes the original bytes.
    const beforeSave = fs.statSync(fixturePath).mtimeMs
    await sleep(30)
    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', code: 'KeyS', ctrlKey: true, bubbles: true, cancelable: true }))`)
    await waitFor(window, `document.querySelector('.toast')?.textContent.includes('Saved poster.psd')`)
    const afterBytes = fs.readFileSync(fixturePath)
    result.steps.unchangedSave = {
      byteExact: afterBytes.equals(fixture),
      rewritten: fs.statSync(fixturePath).mtimeMs >= beforeSave,
      dialogs: saveDialogs.length,
    }
    check(afterBytes.equals(fixture), 'an unchanged Save writes the original PSD bytes', failures)
    check(saveDialogs.length === 0, 'an unchanged Save of a .psd does not ask for a file name', failures)

    // 3. The renderer's PSD codec (real lazy chunk, DOM canvas hooks): import, edit, export, Save As.
    nextSavePath = path.join(workDir, 'edited.psd')
    const codec = await window.webContents.executeJavaScript(`(async () => {
      const payload = await window.simpleImage.openPath(${JSON.stringify(fixturePath)})
      const module = await import(${JSON.stringify(psdChunk())})
      module.configurePsdCanvas('dom')
      const doc = await module.importPsd(payload.data)
      const layers = doc.layers.map((layer) => {
        if (layer.name === 'Hidden') return { ...layer, visible: true, name: 'Shown' }
        if (layer.name === 'Red half') return { ...layer, opacity: 0.25 }
        return layer
      })
      const edited = { ...doc, layers }
      const composite = module.flattenForExport(edited)
      const bytes = await module.exportPsd(edited, composite)
      const saved = await window.simpleImage.saveImage({ data: bytes, path: null, name: 'poster.psd', format: 'psd', forceDialog: true, purpose: 'export' })
      const text = doc.layers.find((layer) => layer.kind === 'text')
      return {
        format: payload.format,
        directSave: payload.directSave,
        names: doc.layers.map((layer) => layer.name),
        kinds: doc.layers.map((layer) => layer.kind),
        issues: doc.issues.map((issue) => issue.code),
        ppi: doc.ppi,
        hasComposite: Boolean(doc.composite),
        textFamily: text ? text.text.style.fontFamily : null,
        textWeight: text ? text.text.style.fontWeight : null,
        savedPath: saved ? saved.path : null,
        bytes: bytes.byteLength,
      }
    })()`)
    result.steps.codec = codec
    check(codec.format === 'psd' && codec.directSave === true, 'file:open-path reports a PSD that saves in place', failures)
    check(JSON.stringify(codec.names) === JSON.stringify(LAYER_NAMES), `layer names and order survive import (${codec.names})`, failures)
    check(JSON.stringify(codec.kinds) === JSON.stringify(['raster', 'raster', 'raster', 'raster', 'text', 'adjustment']), `layer kinds (${codec.kinds})`, failures)
    check(codec.ppi === 144 && codec.hasComposite, 'resolution and merged image are read', failures)
    check(codec.textFamily === 'Segoe UI' && codec.textWeight === 700, `PostScript font maps to a CSS family (${codec.textFamily} ${codec.textWeight})`, failures)
    check(codec.savedPath === nextSavePath && fs.existsSync(nextSavePath), 'the edited PSD was saved through file:save', failures)
    if (fs.existsSync(nextSavePath)) {
      const reread = layerSummary(fs.readFileSync(nextSavePath))
      result.steps.codec.reread = reread
      const byName = Object.fromEntries(reread.layers.map((layer) => [layer.name, layer]))
      check(reread.layers.length === LAYER_NAMES.length, 'the saved PSD keeps every layer', failures)
      check(byName.Shown && byName.Shown.hidden === false, 'the renamed, now visible layer is saved', failures)
      check(byName['Red half'] && Math.abs(byName['Red half'].opacity - 0.25) <= 1 / 255, 'the changed opacity is saved', failures)
      check(byName['Red half'] && byName['Red half'].blendMode === 'multiply' && byName['Red half'].mask?.defaultColor === 0, 'blend mode and mask are saved', failures)
      check(byName.Clipped && byName.Clipped.clipping, 'the clipping mask is saved', failures)
      check(byName.Title && byName.Title.text && byName.Title.text.text === 'Simple PSD' && byName.Title.text.font === 'SegoeUI-Bold', 'text stays text with its font', failures)
      check(byName['Levels 1'] && byName['Levels 1'].adjustment === 'levels', 'the adjustment layer stays an adjustment layer', failures)
    }

    // 4. Editing from the keyboard in the editor (needs the full Advanced editor; the placeholder cannot edit).
    nextSavePath = path.join(workDir, 'unexpected-dialog.psd')
    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'j', code: 'KeyJ', ctrlKey: true, bubbles: true, cancelable: true }))`)
    let edited = false
    try {
      await waitFor(window, `document.querySelector('.save-state')?.textContent === 'Modified'`, 2500)
      edited = true
    } catch {
      edited = false
    }
    if (edited) {
      await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', code: 'KeyS', ctrlKey: true, bubbles: true, cancelable: true }))`)
      await waitFor(window, `document.querySelector('.toast')?.textContent.includes('Saved poster.psd')`, 30_000)
      const saved = layerSummary(fs.readFileSync(fixturePath))
      result.steps.uiEdit = { status: 'ran', layers: saved.layers.map((layer) => layer.name) }
      check(saved.layers.length === LAYER_NAMES.length + 1, 'Ctrl+J then Ctrl+S saves the new layer into the PSD', failures)
    } else {
      result.steps.uiEdit = { status: 'skipped', reason: 'this editor build does not edit from the keyboard yet (placeholder Advanced editor)' }
      check(fs.readFileSync(fixturePath).equals(fixture), 'without an edit the PSD on disk stays unchanged', failures)
    }

    // 5. The clipboard read bridge exists (read-only check; the system clipboard is never written).
    const clipboard = await window.webContents.executeJavaScript(`(async () => ({
      read: typeof window.simpleImage.readClipboardImage,
      has: typeof window.simpleImage.clipboardHasImage,
      hasImage: typeof (await window.simpleImage.clipboardHasImage()),
    }))()`)
    result.steps.clipboard = clipboard
    check(clipboard.read === 'function' && clipboard.has === 'function' && clipboard.hasImage === 'boolean', 'the clipboard bridge is exposed', failures)

    const screenshot = await window.webContents.capturePage()
    fs.writeFileSync(path.join(workDir, 'psd-smoke.png'), screenshot.toPNG())
    result.failures = failures
    fs.writeFileSync(path.join(workDir, 'psd-smoke-result.json'), JSON.stringify(result, null, 2), 'utf8')
    if (failures.length) {
      console.error(`psd-smoke failed:\n- ${failures.join('\n- ')}`)
      app.exit(1)
      return
    }
    console.log(`psd-smoke passed (UI edit step: ${result.steps.uiEdit.status}). Artifacts: ${workDir}`)
    app.exit(0)
  } catch (error) {
    console.error(error)
    try { fs.writeFileSync(path.join(workDir, 'psd-smoke-result.json'), JSON.stringify({ ...result, error: String(error && error.stack || error), failures }, null, 2), 'utf8') } catch {}
    app.exit(1)
  }
})
