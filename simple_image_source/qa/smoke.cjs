const { app, BrowserWindow, ipcMain } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const { MIME_BY_EXTENSION, extensionOf, imageDimensions, validateImageBytes } = require('../electron/image-files.cjs')

const projectRoot = path.join(__dirname, '..')
const fixturePath = process.env.SIMPLE_IMAGE_SMOKE_FILE || path.join(projectRoot, 'public', 'brand-icon.png')
let convertedImage = null

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
    const initial = await window.webContents.executeJavaScript(`(() => {
      const canvas = document.querySelector('canvas')
      return { width: canvas.width, height: canvas.height, details: document.querySelector('.inspector').innerText }
    })()`)
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
    await window.webContents.executeJavaScript(`document.querySelector('button[title="Convert image to PDF"]').click()`)
    await waitFor(window, `document.querySelector('.toast')?.textContent.includes('Created converted.pdf')`)
    const conversionKeptEdits = await window.webContents.executeJavaScript(`document.querySelector('.save-state')?.textContent === 'Modified'`)
    await new Promise((resolve) => setTimeout(resolve, 300))
    const screenshot = await window.webContents.capturePage()
    await fs.writeFile(path.join(__dirname, 'smoke.png'), screenshot.toPNG())
    await fs.writeFile(path.join(__dirname, 'smoke-result.json'), JSON.stringify({ initial, edited, rotated, convertedImage, conversionKeptEdits }, null, 2), 'utf8')
    if (
      !edited.modified
      || edited.width >= initial.width
      || rotated.width !== edited.height
      || rotated.height !== edited.width
      || convertedImage?.name !== path.basename(fixturePath)
      || convertedImage?.dimensions?.width !== rotated.width
      || convertedImage?.dimensions?.height !== rotated.height
      || !conversionKeptEdits
      || !initial.details.includes('PNG')
    ) {
      throw new Error('The editor smoke assertions failed.')
    }
    app.exit(0)
  } catch (error) {
    console.error(error)
    app.exit(1)
  }
})
