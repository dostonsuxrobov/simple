// Render already-reviewed local print HTML through the real Electron PDF path.
// This writes a PDF only; it never sends a physical printer job.
//
//   electron scripts/qa-print-render.cjs <input.html> <output.pdf>
//   electron scripts/qa-print-render.cjs <config.json>
//
// Config mode: { input, output, landscape?, probes?: [{ name, selector, colors: ['#RRGGBB'], tolerance? }] }
// renders a hidden window, reports for every probe which of its colours are painted inside the
// element (screen rendering of the same HTML), prints the PDF with the options main.cjs uses for
// PDF export, and writes `${output}.json` with { pdfBytes, probes }.
const { app, BrowserWindow } = require('electron')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')

app.commandLine.appendSwitch('force-device-scale-factor', '1')
app.disableHardwareAcceleration()

let window, profile, ownedProfile = false

function rgb(hex) {
  const value = Number.parseInt(String(hex).replace(/^#/, ''), 16)
  return [(value >> 16) & 255, (value >> 8) & 255, value & 255]
}

async function probeColors(probes) {
  const rects = await window.webContents.executeJavaScript(`(${JSON.stringify(probes.map((probe) => probe.selector))}).map((selector) => {
    const element = document.querySelector(selector)
    if (!element) return null
    const box = element.getBoundingClientRect()
    return { x: box.left, y: box.top, width: box.width, height: box.height }
  })`)
  const image = await window.webContents.capturePage()
  const { width, height } = image.getSize()
  const bitmap = image.toBitmap()
  const results = {}
  probes.forEach((probe, index) => {
    const rect = rects[index]
    const found = {}
    for (const color of probe.colors) found[color] = false
    if (!rect || width === 0) { results[probe.name] = { found, rect: null }; return }
    const tolerance = probe.tolerance ?? 6
    const targets = probe.colors.map((color) => [color, rgb(color)])
    const left = Math.max(0, Math.floor(rect.x)), top = Math.max(0, Math.floor(rect.y))
    const right = Math.min(width, Math.ceil(rect.x + rect.width)), bottom = Math.min(height, Math.ceil(rect.y + rect.height))
    for (let y = top; y < bottom; y += 1) {
      for (let x = left; x < right; x += 1) {
        const offset = (y * width + x) * 4
        // Native bitmaps are BGRA.
        const pixel = [bitmap[offset + 2], bitmap[offset + 1], bitmap[offset]]
        for (const [color, target] of targets) {
          if (!found[color] && Math.abs(pixel[0] - target[0]) <= tolerance && Math.abs(pixel[1] - target[1]) <= tolerance && Math.abs(pixel[2] - target[2]) <= tolerance) found[color] = true
        }
      }
    }
    results[probe.name] = { found, rect }
  })
  return results
}

;(async () => {
  const configMode = String(process.argv[2] || '').toLowerCase().endsWith('.json')
  const config = configMode
    ? JSON.parse(await fs.readFile(path.resolve(process.argv[2]), 'utf8'))
    : { input: process.argv[2], output: process.argv[3] }
  const input = path.resolve(config.input)
  const output = path.resolve(config.output)
  // A caller-owned profile is removed by the caller once this process has exited (Windows
  // keeps Chromium's files locked until then); otherwise use and remove a private one.
  profile = config.profile ? path.resolve(config.profile) : await fs.mkdtemp(path.join(os.tmpdir(), 'calc-pdf-qa-'))
  ownedProfile = !config.profile
  app.setPath('userData', profile)
  await app.whenReady()
  window = new BrowserWindow({ show: false, width: 1400, height: 1100, backgroundColor: '#ffffff', webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false } })
  await window.loadFile(input)
  await window.webContents.executeJavaScript('document.fonts.ready.then(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))))')
  const probes = Array.isArray(config.probes) && config.probes.length ? await probeColors(config.probes) : {}
  const pdf = configMode
    ? await window.webContents.printToPDF({ printBackground: true, landscape: config.landscape === true, preferCSSPageSize: true, generateTaggedPDF: true, generateDocumentOutline: true })
    : await window.webContents.printToPDF({ preferCSSPageSize: true, printBackground: true, margins: { top: 0, left: 0, right: 0, bottom: 0 } })
  await fs.writeFile(output, pdf)
  if (configMode) await fs.writeFile(`${output}.json`, JSON.stringify({ pdfBytes: pdf.length, probes }, null, 2))
  console.log(`PDF render verified: ${pdf.length} bytes`)
})().catch(error => { console.error(error); process.exitCode = 1 }).finally(async () => {
  window?.destroy()
  if (profile && ownedProfile) await fs.rm(profile, { recursive: true, force: true }).catch(() => {})
  app.exit(process.exitCode || 0)
})
