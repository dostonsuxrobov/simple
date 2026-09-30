const { app, BrowserWindow } = require('electron')
const { PDFDocument } = require('pdf-lib')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const { loadPdfForPrinting } = require('../electron/print-host.cjs')

let directory
let window

async function run() {
  directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-docs-print-host-smoke-'))
  app.setPath('userData', path.join(directory, 'profile'))
  await app.whenReady()

  const document = await PDFDocument.create()
  document.addPage([612, 792]).drawText('Print host readiness smoke', { x: 72, y: 700, size: 20 })
  const target = path.join(directory, 'readiness.pdf')
  await fs.writeFile(target, await document.save())

  window = new BrowserWindow({
    show: false,
    webPreferences: { nodeIntegration: false, contextIsolation: true, sandbox: true, plugins: true, devTools: false },
  })
  const started = Date.now()
  await loadPdfForPrinting(window, pathToFileURL(target).href)
  const result = {
    ok: true,
    elapsedMs: Date.now() - started,
    title: window.webContents.getTitle(),
    loading: window.webContents.isLoading(),
  }
  if (result.loading) throw new Error('The PDF host still reports active navigation after readiness.')
  console.log(JSON.stringify(result, null, 2))
}

run().then(async () => {
  if (window && !window.isDestroyed()) window.destroy()
  if (directory) await fs.rm(directory, { recursive: true, force: true })
  app.quit()
}, async (error) => {
  console.error(error)
  if (window && !window.isDestroyed()) window.destroy()
  if (directory) await fs.rm(directory, { recursive: true, force: true }).catch(() => {})
  app.exit(1)
})
