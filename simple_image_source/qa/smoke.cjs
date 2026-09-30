const { app, BrowserWindow, ipcMain } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const { MIME_BY_EXTENSION, extensionOf, imageDimensions, outputExtension, validateImageBytes } = require('../electron/image-files.cjs')

const projectRoot = path.join(__dirname, '..')
const fixturePath = process.env.SIMPLE_IMAGE_SMOKE_FILE || path.join(projectRoot, 'public', 'brand-icon.png')
let convertedImage = null
const exportedImages = []

async function payload(filePath) {
  const resolved = path.resolve(filePath)
  const extension = extensionOf(resolved)
  const data = validateImageBytes(await fs.readFile(resolved), extension)
  return {
    data: new Uint8Array(data),
    name: path.basename(resolved),
    path: resolved,
    size: data.byteLength,
    format: extension.slice(1),
    mime: MIME_BY_EXTENSION[extension],
    directSave: true,
  }
}

async function waitFor(window, expression, timeout = 10_000) {
  const started = Date.now()
  while (Date.now() - started < timeout) {
    if (await window.webContents.executeJavaScript(expression)) return
    await new Promise((resolve) => setTimeout(resolve, 100))
  }
  throw new Error(`Timed out waiting for: ${expression}`)
}

ipcMain.handle('file:open-path', (_event, filePath) => payload(filePath))
ipcMain.handle('file:save', (_event, input) => {
  const extension = outputExtension(input?.format)
  const data = validateImageBytes(input?.data, extension)
  const dimensions = imageDimensions(data, extension)
  exportedImages.push({
    format: input?.format,
    dimensions,
    forceDialog: input?.forceDialog,
    path: input?.path,
    purpose: input?.purpose,
  })
  return {
    path: path.join(projectRoot, 'qa', `exported${extension}`),
    name: `exported${extension}`,
    size: data.byteLength,
    format: extension.slice(1),
  }
})
ipcMain.handle('image:convert-to-pdf', (_event, input) => {
  const data = validateImageBytes(input?.data, '.png')
  convertedImage = { name: input?.name, dimensions: imageDimensions(data, '.png') }
  return { path: path.join(projectRoot, 'qa', 'converted.pdf'), name: 'converted.pdf', size: data.byteLength }
})
ipcMain.on('window:set-title', () => {})

app.whenReady().then(async () => {
  const window = new BrowserWindow({
    width: 1365,
    height: 840,
    show: false,
    backgroundColor: '#e9e9e9',
    webPreferences: {
      preload: path.join(projectRoot, 'electron', 'preload.cjs'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  })
  try {
    await window.loadFile(path.join(projectRoot, 'dist', 'index.html'))
    await new Promise((resolve) => setTimeout(resolve, 250))
    window.webContents.send('file:open-external', fixturePath)
    await waitFor(window, 'document.querySelector("canvas")?.width > 0 && document.querySelector(".inspector") !== null')
    await window.webContents.executeJavaScript(`if (document.querySelector('.inspector').hidden) document.querySelector('[aria-controls="image-inspector"]').click()`)
    await waitFor(window, `document.querySelector('.inspector').hidden === false`)
    const initial = await window.webContents.executeJavaScript(`(() => {
      const canvas = document.querySelector('canvas')
      return { width: canvas.width, height: canvas.height, details: document.querySelector('.inspector').innerText }
    })()`)
    await window.webContents.executeJavaScript(`document.querySelector('.inspector-close').click()`)
    await waitFor(window, `document.querySelector('.inspector').hidden && localStorage.getItem('simple-image:inspector-open') === 'closed'`)
    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'F4', bubbles: true }))`)
    await waitFor(window, `!document.querySelector('.inspector').hidden && document.querySelector('[aria-controls="image-inspector"]').getAttribute('aria-expanded') === 'true' && localStorage.getItem('simple-image:inspector-open') === 'open'`)
    window.setSize(760, 840)
    await waitFor(window, `matchMedia('(max-width: 820px)').matches && getComputedStyle(document.querySelector('.inspector')).position === 'absolute'`)
    const inspectorControl = await window.webContents.executeJavaScript(`(() => {
      const inspector = document.querySelector('.inspector')
      const toggle = document.querySelector('[aria-controls="image-inspector"]')
      return {
        f4Reopened: !inspector.hidden,
        ariaExpanded: toggle.getAttribute('aria-expanded'),
        storedPreference: localStorage.getItem('simple-image:inspector-open'),
        responsivePosition: getComputedStyle(inspector).position,
        responsiveWidth: inspector.getBoundingClientRect().width,
      }
    })()`)
    window.setSize(1365, 840)
    await waitFor(window, `!matchMedia('(max-width: 820px)').matches && getComputedStyle(document.querySelector('.inspector')).position === 'static'`)
    await window.webContents.executeJavaScript(`document.querySelector('button[title="Crop"]').click()`)
    await waitFor(window, 'document.querySelector(".crop-box") !== null')
    await window.webContents.executeJavaScript(`Array.from(document.querySelectorAll('button')).find((button) => button.textContent.includes('Apply crop')).click()`)
    await waitFor(window, `document.querySelector('canvas').width < ${initial.width} && document.querySelector('.crop-box') === null`)
    const edited = await window.webContents.executeJavaScript(`(() => {
      const canvas = document.querySelector('canvas')
      return { width: canvas.width, height: canvas.height, modified: document.body.innerText.includes('Modified') }
    })()`)
    await window.webContents.executeJavaScript(`document.querySelector('button[title="Rotate right"]').click()`)
    await waitFor(window, `document.querySelector('canvas').width === ${edited.height}`)
    const rotated = await window.webContents.executeJavaScript(`(() => {
      const canvas = document.querySelector('canvas')
      return { width: canvas.width, height: canvas.height }
    })()`)
    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'e', ctrlKey: true, shiftKey: true, bubbles: true }))`)
    await waitFor(window, `document.querySelector('.export-menu') !== null`)
    window.setSize(760, 840)
    await waitFor(window, `matchMedia('(max-width: 1080px)').matches`)
    const exportMenu = await window.webContents.executeJavaScript(`(() => ({
      formats: Array.from(document.querySelectorAll('[data-export-format]')).map((button) => button.dataset.exportFormat),
      text: document.querySelector('.export-menu').innerText,
      visibleOptions: Array.from(document.querySelectorAll('[data-export-format]')).every((button) => button.getBoundingClientRect().height > 0 && getComputedStyle(button.querySelector('span')).display !== 'none'),
      fitsWindow: document.querySelector('.export-menu').getBoundingClientRect().left >= 0 && document.querySelector('.export-menu').getBoundingClientRect().right <= innerWidth,
    }))()`)
    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }))`)
    await waitFor(window, `document.querySelector('.export-menu') === null`)
    window.setSize(1365, 840)
    await window.webContents.executeJavaScript(`document.querySelector('button[title="Export As (Ctrl+Shift+E)"]').click()`)
    await waitFor(window, `document.querySelector('.export-menu') !== null`)
    for (const format of ['png', 'jpeg', 'webp']) {
      await window.webContents.executeJavaScript(`document.querySelector('[data-export-format="${format}"]').click()`)
      await waitFor(window, `document.querySelector('.toast')?.textContent.includes('Exported exported.${format === 'jpeg' ? 'jpg' : format}')`)
      await window.webContents.executeJavaScript(`document.querySelector('button[title="Export As (Ctrl+Shift+E)"]').click()`)
      await waitFor(window, `document.querySelector('.export-menu') !== null`)
    }
    await window.webContents.executeJavaScript(`document.querySelector('[data-export-format="pdf"]').click()`)
    await waitFor(window, `document.querySelector('.toast')?.textContent.includes('Exported converted.pdf')`)
    const conversionKeptEdits = await window.webContents.executeJavaScript(`document.querySelector('.save-state')?.textContent === 'Modified'`)
    await new Promise((resolve) => setTimeout(resolve, 300))
    const screenshot = await window.webContents.capturePage()
    if (!process.env.SIMPLE_IMAGE_SMOKE_NO_ARTIFACTS) {
      await fs.writeFile(path.join(__dirname, 'smoke.png'), screenshot.toPNG())
      await fs.writeFile(path.join(__dirname, 'smoke-result.json'), JSON.stringify({ initial, inspectorControl, edited, rotated, exportMenu, exportedImages, convertedImage, conversionKeptEdits }, null, 2), 'utf8')
    }
    if (
      !edited.modified
      || edited.width >= initial.width
      || rotated.width !== edited.height
      || rotated.height !== edited.width
      || convertedImage?.name !== path.basename(fixturePath)
      || convertedImage?.dimensions?.width !== rotated.width
      || convertedImage?.dimensions?.height !== rotated.height
      || exportMenu.formats.join(',') !== 'png,jpeg,webp,pdf'
      || !exportMenu.visibleOptions
      || !exportMenu.fitsWindow
      || exportMenu.text.includes('SVG image')
      || !exportMenu.text.includes('Vector paths are not preserved')
      || exportedImages.length !== 3
      || exportedImages.some((entry) => entry.forceDialog !== true || entry.path !== null || entry.purpose !== 'export')
      || exportedImages.some((entry) => entry.dimensions?.width !== rotated.width || entry.dimensions?.height !== rotated.height)
      || !conversionKeptEdits
      || !initial.details.includes('PNG')
      || !inspectorControl.f4Reopened
      || inspectorControl.ariaExpanded !== 'true'
      || inspectorControl.storedPreference !== 'open'
      || inspectorControl.responsivePosition !== 'absolute'
      || inspectorControl.responsiveWidth <= 0
      || inspectorControl.responsiveWidth > 280
    ) {
      throw new Error('The editor smoke assertions failed.')
    }
    app.exit(0)
  } catch (error) {
    console.error(error)
    app.exit(1)
  }
})
