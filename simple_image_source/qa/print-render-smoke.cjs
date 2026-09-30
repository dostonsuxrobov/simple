const { app, BrowserWindow } = require('electron')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const { PDFDocument } = require('pdf-lib')
const { imageDimensions, validateImageBytes } = require('../electron/image-files.cjs')

app.whenReady().then(async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-print-render-'))
  let window = null
  let exitCode = 0
  try {
    const { buildPrintHtml } = await import('../electron/print-layout.mjs')
    const sourcePath = path.join(__dirname, '..', 'public', 'brand-icon.png')
    const png = validateImageBytes(await fs.readFile(sourcePath), '.png')
    const dimensions = imageDimensions(png, '.png')
    const settings = {
      paper: 'a4', orientation: 'landscape', marginMm: 10, scaleMode: 'fill', scalePercent: 100,
      position: 'bottom-right', background: '#000000', grayscale: true, copies: 2,
    }
    const htmlPath = path.join(directory, 'print.html')
    await fs.writeFile(htmlPath, buildPrintHtml(pathToFileURL(sourcePath).href, dimensions.width, dimensions.height, settings, 'Render smoke'), 'utf8')
    window = new BrowserWindow({
      width: 1123,
      height: 794,
      show: false,
      webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true },
    })
    await window.loadFile(htmlPath)
    const decoded = await window.webContents.executeJavaScript(`document.querySelector('img').decode().then(() => ({ width: document.querySelector('img').naturalWidth, height: document.querySelector('img').naturalHeight }))`)
    if (decoded.width !== dimensions.width || decoded.height !== dimensions.height) throw new Error('The print renderer decoded the wrong image dimensions.')
    const bytes = await window.webContents.printToPDF({
      printBackground: true,
      landscape: true,
      pageSize: 'A4',
      preferCSSPageSize: true,
      margins: { marginType: 'none' },
    })
    const document = await PDFDocument.load(bytes)
    if (document.getPageCount() !== 1) throw new Error(`Expected one print page, received ${document.getPageCount()}.`)
    const page = document.getPage(0)
    const width = page.getWidth()
    const height = page.getHeight()
    const expectedWidth = 297 / 25.4 * 72
    const expectedHeight = 210 / 25.4 * 72
    if (Math.abs(width - expectedWidth) > 1 || Math.abs(height - expectedHeight) > 1) {
      throw new Error(`Expected landscape A4, received ${width.toFixed(2)} × ${height.toFixed(2)} points.`)
    }
    console.log(JSON.stringify({ pages: 1, widthPoints: width, heightPoints: height, decoded, bytes: bytes.byteLength }))
  } catch (error) {
    console.error(error)
    exitCode = 1
  } finally {
    if (window && !window.isDestroyed()) window.destroy()
    await fs.rm(directory, { recursive: true, force: true }).catch(() => {})
  }
  app.exit(exitCode)
})
