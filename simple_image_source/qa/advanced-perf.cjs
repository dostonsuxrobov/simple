// qa/advanced-perf.cjs (WP6): npm run perf:advanced (after npm run build:web)
// Measures the Advanced editor against the performance targets of design section 5.7 on a synthetic 24 MP
// (6000 x 4000) photo-like PNG, in an offscreen-rendered window (real animation frames, never shown) with
// fake IPC (nothing leaves the work folder, nothing is printed). The numbers are informational: the run
// fails only when a step cannot be completed. Targets:
//   enter Advanced, first frame < 700 ms; brush stroke at 100% with 5 layers, median frame < 16 ms;
//   visibility toggle at Fit, fully current < 250 ms; adjustment slider drag at Fit >= 20 fps;
//   flatten + PNG save (Export As PNG) < 3 s; Gaussian blur radius 20 on a 24 MP layer < 4 s with progress.
// Results: <work folder>/advanced-perf-result.json and a table on stdout. The machine's load matters: run it
// on an idle machine for comparable numbers.
// Environment: SIMPLE_IMAGE_QA_DIR (work folder; default a new temp folder), SIMPLE_IMAGE_PERF_TIMEOUT_MS,
// SIMPLE_IMAGE_PERF_SIZE ("6000x4000" by default).
'use strict'

const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const { app, BrowserWindow, ipcMain } = require('electron')

const projectRoot = path.join(__dirname, '..')
const workDir = process.env.SIMPLE_IMAGE_QA_DIR
  ? path.resolve(process.env.SIMPLE_IMAGE_QA_DIR)
  : fs.mkdtempSync(path.join(os.tmpdir(), 'simple-image-advanced-perf-'))
fs.mkdirSync(workDir, { recursive: true })
app.setPath('userData', path.join(workDir, 'user-data'))

const timeoutMs = Number(process.env.SIMPLE_IMAGE_PERF_TIMEOUT_MS) || 420_000
const hardTimeout = setTimeout(() => {
  console.error(`advanced-perf: timed out after ${timeoutMs} ms (artifacts: ${workDir})`)
  app.exit(2)
}, timeoutMs)
hardTimeout.unref?.()

const [W, H] = (process.env.SIMPLE_IMAGE_PERF_SIZE || '6000x4000').split('x').map((value) => Math.max(64, Number(value) || 0))

// #region fixture: a photo-like 24 MP PNG (smooth gradients, soft shapes, fine noise)

function crc32(buffer) {
  let table = crc32.table
  if (!table) {
    table = crc32.table = new Int32Array(256)
    for (let n = 0; n < 256; n += 1) {
      let c = n
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
      table[n] = c
    }
  }
  let crc = -1
  for (let i = 0; i < buffer.length; i += 1) crc = table[(crc ^ buffer[i]) & 0xff] ^ (crc >>> 8)
  return (crc ^ -1) >>> 0
}

function pngChunk(type, data) {
  const length = Buffer.alloc(4)
  length.writeUInt32BE(data.length)
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data])
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(body))
  return Buffer.concat([length, body, crc])
}

function makeFixture(width, height) {
  const stride = width * 4 + 1
  const raw = Buffer.alloc(stride * height)
  let seed = 12345
  for (let y = 0; y < height; y += 1) {
    const row = y * stride
    raw[row] = 0
    for (let x = 0; x < width; x += 1) {
      seed = (seed * 1103515245 + 12345) >>> 0
      const noise = (seed >>> 24) % 9 - 4
      const dx = x / width - 0.62
      const dy = y / height - 0.4
      const blob = Math.max(0, 1 - Math.sqrt(dx * dx + dy * dy) * 2.6)
      const p = row + 1 + x * 4
      raw[p] = Math.max(0, Math.min(255, Math.round(40 + 170 * (x / width) + 60 * blob + noise)))
      raw[p + 1] = Math.max(0, Math.min(255, Math.round(70 + 120 * (y / height) + 30 * blob + noise)))
      raw[p + 2] = Math.max(0, Math.min(255, Math.round(150 - 90 * (y / height) + 80 * blob + noise)))
      raw[p + 3] = 255
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
    pngChunk('IDAT', zlib.deflateSync(raw, { level: 1 })),
    pngChunk('IEND', Buffer.alloc(0)),
  ])
}

const fixtureStarted = Date.now()
const fixtureBytes = makeFixture(W, H)
const fixturePath = path.join(workDir, `photo-${W}x${H}.png`)
fs.writeFileSync(fixturePath, fixtureBytes)
const fixtureMs = Date.now() - fixtureStarted

// #endregion

// #region fake backend

const saves = []
ipcMain.handle('file:open-path', () => ({ data: new Uint8Array(fixtureBytes), name: path.basename(fixturePath), path: fixturePath, size: fixtureBytes.length, format: 'png', mime: 'image/png', directSave: true }))
ipcMain.handle('file:open-dialog', () => null)
ipcMain.handle('file:list-siblings', () => [])
ipcMain.handle('file:save', (_event, input) => {
  saves.push({ format: input && input.format, purpose: input && input.purpose, bytes: input && input.data ? input.data.byteLength : 0, at: Date.now() })
  return { path: path.join(workDir, `saved.${input && input.format}`), name: `saved.${input && input.format}`, size: 1, format: input && input.format }
})
ipcMain.handle('image:convert-to-pdf', () => null)
ipcMain.handle('image:print', () => false)
ipcMain.handle('clipboard:write-png', () => ({ width: 1, height: 1 }))
ipcMain.handle('clipboard:read-image', () => null)
ipcMain.handle('clipboard:has-image', () => false)
ipcMain.handle('app:get-version', () => '0.0.0-perf')
ipcMain.handle('app:new-window', () => false)
ipcMain.handle('file:open-in-new-window', () => false)
ipcMain.on('window:set-title', () => {})
ipcMain.on('window:minimize', () => {})
ipcMain.on('window:toggle-maximize', () => {})
ipcMain.on('window:confirm-close', () => {})

// #endregion

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
const result = { workDir, size: `${W} x ${H}`, megapixels: Number(((W * H) / 1e6).toFixed(1)), fixtureMs, cpus: os.cpus().length, loadAverage: os.loadavg(), measurements: {}, failures: [] }

function median(values) {
  if (!values.length) return null
  const sorted = [...values].sort((a, b) => a - b)
  const middle = Math.floor(sorted.length / 2)
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2
}

function percentile(values, fraction) {
  if (!values.length) return null
  const sorted = [...values].sort((a, b) => a - b)
  return sorted[Math.min(sorted.length - 1, Math.floor(sorted.length * fraction))]
}

function round(value) {
  return value === null || value === undefined ? null : Math.round(value * 10) / 10
}

app.whenReady().then(async () => {
  const window = new BrowserWindow({
    width: 1600,
    height: 1000,
    show: false,
    backgroundColor: '#e9e9e9',
    webPreferences: { preload: path.join(projectRoot, 'electron', 'preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true, offscreen: true, backgroundThrottling: false },
  })
  window.webContents.setFrameRate(60)
  const consoleErrors = []
  window.webContents.on('console-message', (event) => {
    const level = event.level ?? event.params?.level
    const message = event.message ?? event.params?.message
    if (level === 'error' || level === 3) consoleErrors.push(String(message).slice(0, 400))
  })
  const js = (expression) => window.webContents.executeJavaScript(expression)
  async function waitFor(expression, timeout = 30_000, label = expression) {
    const started = Date.now()
    let last
    while (Date.now() - started < timeout) {
      try {
        last = await js(expression)
        if (last) return last
      } catch (error) {
        last = String(error)
      }
      await sleep(50)
    }
    throw new Error(`Timed out waiting for: ${label} (last: ${JSON.stringify(last)})`)
  }
  const measure = async (name, target, run) => {
    try {
      const value = await run()
      result.measurements[name] = { ...value, target }
    } catch (error) {
      result.measurements[name] = { error: String(error && error.message || error), target }
      result.failures.push(`${name}: ${String(error && error.message || error)}`)
      throw error
    }
  }

  try {
    await window.loadFile(path.join(projectRoot, 'dist', 'index.html'))
    await sleep(200)
    await js('window.__SIMPLE_IMAGE_QA__ = true; true')
    const openStarted = Date.now()
    window.webContents.send('file:open-external', fixturePath)
    await waitFor(`document.querySelector('canvas')?.width === ${W} && document.querySelector('.advanced-button:not(:disabled)') !== null`, 120_000, 'Simple shows the 24 MP image')
    result.simpleOpenMs = Date.now() - openStarted

    // Enter Advanced: click to editor ready, and to the first fully drawn frame.
    await measure('enterAdvanced', 'first frame < 700 ms', () => js(`(async () => {
      const started = performance.now()
      document.querySelector('.advanced-button').click()
      while (!(window.__simpleAdvanced && document.querySelector('.ae-stage')?.dataset.view)) await new Promise((resolve) => setTimeout(resolve, 4))
      const ready = performance.now() - started
      await window.__simpleAdvanced.compositor.settle()
      return { readyMs: ready, firstFrameMs: performance.now() - started, level: window.__simpleAdvanced.compositor.level }
    })()`))

    // Five full layers with different blend modes and opacities.
    await js(`(async () => {
      const a = window.__simpleAdvanced
      for (let i = 0; i < 4; i += 1) {
        a.run('layer.duplicate')
        await new Promise((resolve) => setTimeout(resolve, 30))
      }
      const layers = a.store.getState().layers
      a.store.transact('Perf setup', 'layer', (tx) => {
        tx.updateLayer(layers[1].id, { blendMode: 'multiply', opacity: 0.5 })
        tx.updateLayer(layers[2].id, { blendMode: 'screen', opacity: 0.6 })
        tx.updateLayer(layers[3].id, { blendMode: 'overlay', opacity: 0.4 })
        tx.updateLayer(layers[4].id, { blendMode: 'normal', opacity: 0.3 })
      })
      return a.store.getState().layers.length
    })()`)
    const layerCount = await js(`window.__simpleAdvanced.store.getState().layers.length`)
    if (layerCount !== 5) throw new Error(`expected 5 layers, got ${layerCount}`)

    // Brush stroke at 100% on the top layer: compositor work per frame and frame intervals.
    await measure('brushStroke100', 'median frame < 16 ms', () => js(`(async () => {
      const a = window.__simpleAdvanced
      const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
      a.run('view.actual-pixels')
      a.run('tool.brush')
      a.editor.updateOptions('brush', { size: 60, hardness: 0.8 })
      await sleep(100)
      await a.compositor.settle()
      const stage = document.querySelector('.ae-stage')
      const box = stage.getBoundingClientRect()
      const work = []
      const intervals = []
      let last = performance.now()
      const off = a.compositor.onFrame(() => {
        const now = performance.now()
        intervals.push(now - last)
        last = now
        work.push(a.compositor.stats().lastWorkMs)
      })
      const send = (type, x, y, buttons) => stage.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerId: 21, pointerType: 'mouse', isPrimary: true, clientX: x, clientY: y, button: type === 'pointermove' ? -1 : 0, buttons, pressure: buttons ? 0.6 : 0 }))
      const y0 = box.top + box.height * 0.3
      const x0 = box.left + box.width * 0.15
      const x1 = box.left + box.width * 0.85
      send('pointermove', x0, y0, 0)
      send('pointerdown', x0, y0, 1)
      const started = performance.now()
      const moves = 90
      for (let i = 1; i <= moves; i += 1) {
        const t = i / moves
        send('pointermove', x0 + (x1 - x0) * t, y0 + Math.sin(t * Math.PI * 3) * box.height * 0.2, 1)
        await sleep(11)
      }
      send('pointerup', x1, y0, 0)
      const strokeMs = performance.now() - started
      await a.compositor.settle()
      off()
      const sorted = (values) => [...values].sort((p, q) => p - q)
      const med = (values) => { const s = sorted(values); return s.length ? s[Math.floor(s.length / 2)] : null }
      const p95 = (values) => { const s = sorted(values); return s.length ? s[Math.min(s.length - 1, Math.floor(s.length * 0.95))] : null }
      return { zoom: a.view.getView().zoom, frames: work.length, strokeMs, medianWorkMs: med(work), p95WorkMs: p95(work), medianFrameIntervalMs: med(intervals), p95FrameIntervalMs: p95(intervals), history: a.store.history.getState().entries.at(-1).label }
    })()`))

    // Visibility toggle at Fit (Layers panel eye), until every visible tile is current.
    await measure('visibilityToggleFit', 'fully current < 250 ms', () => js(`(async () => {
      const a = window.__simpleAdvanced
      a.run('view.fit')
      await new Promise((resolve) => setTimeout(resolve, 150))
      await a.compositor.settle()
      const times = []
      for (let i = 0; i < 4; i += 1) {
        const started = performance.now()
        document.querySelectorAll('.ae-layer-row .ae-eye')[2].click()
        await new Promise((resolve) => setTimeout(resolve, 0))
        await a.compositor.settle()
        times.push(performance.now() - started)
      }
      return { level: a.compositor.level, timesMs: times, medianMs: [...times].sort((p, q) => p - q)[Math.floor(times.length / 2)] }
    })()`))

    // Adjustment layer slider drag at Fit: Levels black handle in Properties, each step fully drawn.
    await measure('adjustmentSliderFit', '>= 20 fps', () => js(`(async () => {
      const a = window.__simpleAdvanced
      a.run('adjustment-layer.levels')
      let handle = null
      for (let i = 0; i < 100 && !handle; i += 1) {
        await new Promise((resolve) => setTimeout(resolve, 20))
        handle = document.querySelector('.ae-properties-panel .ae-lv-handle.is-black')
      }
      if (!handle) throw new Error('no Levels handle')
      await a.compositor.settle()
      const track = handle.parentElement.getBoundingClientRect()
      const send = (type, fraction, buttons) => handle.dispatchEvent(new PointerEvent(type, { bubbles: true, cancelable: true, pointerId: 31, pointerType: 'mouse', isPrimary: true, clientX: track.left + track.width * fraction, clientY: track.top + 8, button: type === 'pointermove' ? -1 : 0, buttons }))
      send('pointerdown', 0, 1)
      const steps = 30
      const started = performance.now()
      for (let i = 1; i <= steps; i += 1) {
        send('pointermove', (i / steps) * 0.35, 1)
        await new Promise((resolve) => setTimeout(resolve, 0))
        await a.compositor.settle()
      }
      const elapsed = performance.now() - started
      send('pointerup', 0.35, 0)
      await a.compositor.settle()
      const spec = a.store.getState().layers.at(-1).adjustment
      return { steps, elapsedMs: elapsed, fps: steps / (elapsed / 1000), inBlack: spec.rgb.inBlack, historySteps: a.store.history.getState().entries.filter((entry) => entry.label === 'Modify Levels Layer').length }
    })()`))

    // Flatten + PNG save: Export As PNG, from the click to the bytes arriving at the backend.
    await measure('flattenPngExport', '< 3000 ms', async () => {
      await js(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'w', code: 'KeyW', ctrlKey: true, altKey: true, shiftKey: true, bubbles: true, cancelable: true })); true`)
      await waitFor(`document.querySelector('.export-menu button[data-export-format="png"]') !== null`, 10_000, 'the Export menu')
      const before = saves.length
      const started = Date.now()
      await js(`document.querySelector('.export-menu button[data-export-format="png"]').click(); true`)
      const deadline = Date.now() + 120_000
      while (saves.length === before && Date.now() < deadline) await sleep(20)
      if (saves.length === before) throw new Error('the PNG never arrived')
      const save = saves[saves.length - 1]
      return { ms: save.at - started, bytes: save.bytes, format: save.format }
    })

    // Gaussian blur radius 20 on a full 24 MP layer, through the Filter dialog.
    await measure('gaussianBlur20', '< 4000 ms with progress', async () => {
      await js(`(() => { const a = window.__simpleAdvanced; const layers = a.store.getState().layers; a.store.transact('Select Layer', 'layer', (tx) => tx.setActiveLayer(layers[1].id), { affectsOutput: false }); a.run('filter.gaussian-blur'); return true })()`)
      await waitFor(`document.querySelector('.ae-dframe h2')?.textContent === 'Gaussian Blur'`, 10_000, 'the Gaussian Blur dialog')
      await js(`(() => {
        const input = document.querySelector('.ae-dframe .ae-sf .ae-num')
        input.focus()
        Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '20')
        input.dispatchEvent(new Event('input', { bubbles: true }))
        return true
      })()`)
      await sleep(50)
      const timing = await js(`(async () => {
        const input = document.querySelector('.ae-dframe .ae-sf .ae-num')
        const started = performance.now()
        input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true, cancelable: true }))
        let busySeen = false
        let workingButton = false
        while (true) {
          const hint = document.querySelector('.ae-status-hint')?.textContent ?? ''
          if (hint.includes('Working')) busySeen = true
          if (document.querySelector('.ae-dframe-ok')?.textContent.includes('Working')) workingButton = true
          const labels = Array.from(document.querySelectorAll('.ae-history-row:not(.is-future)')).map((row) => row.textContent.trim())
          if (labels.some((label) => label.startsWith('Gaussian Blur')) && !document.querySelector('.ae-dframe')) break
          if (performance.now() - started > 120000) throw new Error('the blur did not finish')
          await new Promise((resolve) => setTimeout(resolve, 15))
        }
        return { ms: performance.now() - started, busyIndicator: busySeen, workingButton }
      })()`)
      return timing
    })

    result.memory = await js(`window.__simpleAdvanced.store.memoryUsage()`)
  } catch (error) {
    result.error = String(error && error.stack || error)
  }
  result.consoleErrors = consoleErrors
  fs.writeFileSync(path.join(workDir, 'advanced-perf-result.json'), JSON.stringify(result, null, 2), 'utf8')
  const m = result.measurements
  const rows = [
    ['Enter Advanced, first frame', m.enterAdvanced && round(m.enterAdvanced.firstFrameMs), 'ms', '< 700'],
    ['Brush at 100%, median compositor work per frame', m.brushStroke100 && round(m.brushStroke100.medianWorkMs), 'ms', '< 16'],
    ['Brush at 100%, median frame interval', m.brushStroke100 && round(m.brushStroke100.medianFrameIntervalMs), 'ms', '< 16'],
    ['Visibility toggle at Fit, fully current', m.visibilityToggleFit && round(m.visibilityToggleFit.medianMs), 'ms', '< 250'],
    ['Adjustment slider at Fit', m.adjustmentSliderFit && round(m.adjustmentSliderFit.fps), 'fps', '>= 20'],
    ['Flatten + PNG export', m.flattenPngExport && round(m.flattenPngExport.ms), 'ms', '< 3000'],
    ['Gaussian blur r20 on 24 MP', m.gaussianBlur20 && round(m.gaussianBlur20.ms), 'ms', '< 4000'],
  ]
  console.log(`advanced-perf on ${result.size} (${result.megapixels} MP), ${result.cpus} CPUs, load ${result.loadAverage.map((v) => v.toFixed(1)).join(' ')}`)
  for (const [label, value, unit, target] of rows) console.log(`  ${label.padEnd(50)} ${String(value ?? 'n/a').padStart(9)} ${unit.padEnd(4)} target ${target}`)
  console.log(`Artifacts: ${workDir}`)
  if (result.error || result.failures.length) {
    console.error(`advanced-perf could not finish:\n- ${[...result.failures, ...(result.error ? [result.error] : [])].join('\n- ')}`)
    app.exit(1)
    return
  }
  app.exit(0)
})
