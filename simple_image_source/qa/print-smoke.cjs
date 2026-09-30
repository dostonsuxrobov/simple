const { app, BrowserWindow, ipcMain } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const { MIME_BY_EXTENSION, extensionOf, imageDimensions, validateImageBytes } = require('../electron/image-files.cjs')

const projectRoot = path.join(__dirname, '..')
const fixturePath = process.env.SIMPLE_IMAGE_SMOKE_FILE || path.join(projectRoot, 'public', 'brand-icon.png')
let printRequest = null
let openDialogRequests = 0

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
    await new Promise((resolve) => setTimeout(resolve, 80))
  }
  throw new Error(`Timed out waiting for: ${expression}`)
}

ipcMain.handle('file:open-path', (_event, filePath) => payload(filePath))
ipcMain.handle('file:open-dialog', () => { openDialogRequests += 1; return null })
ipcMain.handle('image:print', async (_event, input) => {
  const png = validateImageBytes(input?.data, '.png')
  printRequest = {
    name: input?.name,
    dimensions: imageDimensions(png, '.png'),
    requestedDimensions: { width: input?.width, height: input?.height },
    settings: input?.settings,
  }
  await new Promise((resolve) => setTimeout(resolve, 500))
  return true
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
    await window.webContents.executeJavaScript(`localStorage.removeItem('simple-image:print-settings:v2')`)
    window.webContents.send('file:open-external', fixturePath)
    await waitFor(window, `document.querySelector('.save-state')?.textContent === 'Saved' && document.querySelector('button[title="Print (Ctrl+P)"]')?.disabled === false`)
    const source = await window.webContents.executeJavaScript(`(() => ({ width: document.querySelector('canvas').width, height: document.querySelector('canvas').height }))()`)

    await window.webContents.executeJavaScript(`document.querySelector('button[title="Print (Ctrl+P)"]').click()`)
    await waitFor(window, `document.querySelector('.image-print-dialog') !== null && document.querySelector('.print-preview-image-frame img')?.complete && document.querySelector('.print-preview-image-frame img')?.naturalWidth > 0`)
    const defaultLayout = await window.webContents.executeJavaScript(`(() => {
      const dialog = document.querySelector('.image-print-dialog').getBoundingClientRect()
      const controls = document.querySelector('.image-print-controls').getBoundingClientRect()
      const preview = document.querySelector('.image-print-preview').getBoundingClientRect()
      const sheet = document.querySelector('.print-preview-sheet').getBoundingClientRect()
      return {
        title: document.querySelector('#image-print-title').textContent,
        paper: document.querySelector('[data-print-setting="paper"]').value,
        orientation: document.querySelector('[data-print-setting="orientation"]').value,
        scaleMode: document.querySelector('[data-print-setting="scale-mode"]').value,
        panesSideBySide: controls.right <= preview.left + 1,
        dialogFits: dialog.left >= 0 && dialog.top >= 0 && dialog.right <= innerWidth && dialog.bottom <= innerHeight,
        portraitSheet: sheet.height > sheet.width,
        imageLoaded: document.querySelector('.print-preview-image-frame img').complete,
      }
    })()`)

    const backgroundIsolation = await window.webContents.executeJavaScript(`(() => {
      const canvas = document.querySelector('canvas')
      const before = { width: canvas.width, height: canvas.height, title: document.querySelector('.document-title').textContent }
      document.querySelector('button[title="Rotate right"]').click()
      const transfer = new DataTransfer()
      transfer.items.add(new File(['not an image'], 'intruder.png', { type: 'image/png' }))
      const overlay = document.querySelector('.image-print-overlay')
      overlay.dispatchEvent(new DragEvent('dragenter', { bubbles: true, cancelable: true, dataTransfer: transfer }))
      overlay.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: transfer }))
      return {
        dimensionsUnchanged: canvas.width === before.width && canvas.height === before.height,
        titleUnchanged: document.querySelector('.document-title').textContent === before.title,
        noDropOverlay: document.querySelector('.drop-overlay') === null,
        backgroundInert: ['.titlebar', '.toolbar', '.content', '.statusbar'].every((selector) => document.querySelector(selector).inert),
      }
    })()`)

    const openRequestsBeforeShortcut = openDialogRequests
    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'o', ctrlKey: true }))`)
    await new Promise((resolve) => setTimeout(resolve, 120))
    if (openDialogRequests !== openRequestsBeforeShortcut) throw new Error('Background Ctrl+O escaped the print modal.')

    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))`)
    await waitFor(window, `document.querySelector('.image-print-dialog') === null && document.activeElement?.getAttribute('title') === 'Print (Ctrl+P)'`)
    const toolbarFocusRestored = await window.webContents.executeJavaScript(`document.activeElement?.getAttribute('title') === 'Print (Ctrl+P)'`)

    await window.webContents.executeJavaScript(`(() => {
      const details = document.querySelector('[aria-controls="image-inspector"]')
      details.focus()
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true }))
    })()`)
    await waitFor(window, `document.querySelector('.image-print-dialog') !== null`)
    await window.webContents.executeJavaScript(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))`)
    await waitFor(window, `document.querySelector('.image-print-dialog') === null && document.activeElement?.getAttribute('aria-controls') === 'image-inspector'`)
    const shortcutFocusRestored = await window.webContents.executeJavaScript(`document.activeElement?.getAttribute('aria-controls') === 'image-inspector'`)

    await window.webContents.executeJavaScript(`document.querySelector('button[title="Print (Ctrl+P)"]').click()`)
    await waitFor(window, `document.querySelector('.image-print-dialog') !== null`)

    await window.webContents.executeJavaScript(`(() => {
      const choose = (setting, value) => {
        const element = document.querySelector('[data-print-setting="' + setting + '"]')
        element.value = value
        element.dispatchEvent(new Event('change', { bubbles: true }))
      }
      choose('paper', 'a4')
      choose('orientation', 'landscape')
      choose('margin-preset', '0')
      choose('scale-mode', 'fill')
      choose('position', 'bottom-right')
      document.querySelector('button[aria-label="Black background"]').click()
      document.querySelector('[data-print-setting="grayscale"]').click()
      const copies = document.querySelector('[data-print-setting="copies"]')
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
      setter.call(copies, '3')
      copies.dispatchEvent(new Event('input', { bubbles: true }))
    })()`)
    await waitFor(window, `document.querySelector('.print-preview-sheet')?.dataset.orientation === 'landscape' && document.querySelector('.print-preview-warning') !== null && document.querySelector('[data-print-setting="copies"]').value === '3'`)
    const configuredLayout = await window.webContents.executeJavaScript(`(() => {
      const sheet = document.querySelector('.print-preview-sheet').getBoundingClientRect()
      const frame = document.querySelector('.print-preview-image-frame')
      return {
        landscapeSheet: sheet.width > sheet.height,
        clippedWarning: document.querySelector('.print-preview-warning').textContent.includes('cropped'),
        blackBackground: getComputedStyle(frame).backgroundColor === 'rgb(0, 0, 0)',
        grayscaleComposition: getComputedStyle(frame).filter === 'grayscale(1)' && getComputedStyle(frame.querySelector('img')).filter === 'none',
        grayscaleLabelHonest: document.querySelector('[data-print-setting="grayscale"]').parentElement.textContent.includes('transparency background'),
        summary: document.querySelector('.print-preview-summary').innerText,
      }
    })()`)

    if (!process.env.SIMPLE_IMAGE_SMOKE_NO_ARTIFACTS) {
      await window.webContents.executeJavaScript(`new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)))`)
      const screenshot = await window.webContents.capturePage()
      await fs.writeFile(path.join(__dirname, 'print-smoke.png'), screenshot.toPNG())
    }

    window.setSize(760, 840)
    await waitFor(window, `matchMedia('(max-width: 820px)').matches && getComputedStyle(document.querySelector('.image-print-body')).display === 'block'`)
    const responsiveLayout = await window.webContents.executeJavaScript(`(() => ({
      stacked: getComputedStyle(document.querySelector('.image-print-body')).display === 'block',
      dialogFits: document.querySelector('.image-print-dialog').getBoundingClientRect().right <= innerWidth,
      bodyScrollable: ['auto', 'scroll'].includes(getComputedStyle(document.querySelector('.image-print-body')).overflowY),
    }))()`)
    window.setSize(1365, 840)
    await waitFor(window, `!matchMedia('(max-width: 820px)').matches`)

    await window.webContents.executeJavaScript(`(() => {
      const submit = document.querySelector('.print-submit')
      submit.focus()
      submit.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }))
    })()`)
    const focusTrapped = await window.webContents.executeJavaScript(`document.activeElement?.getAttribute('aria-label') === 'Close print setup'`)

    await window.webContents.executeJavaScript(`document.querySelector('.print-submit').click()`)
    await waitFor(window, `document.querySelector('.print-submit')?.disabled === true`)
    const frozenJob = await window.webContents.executeJavaScript(`(() => {
      const before = {
        paper: document.querySelector('[data-print-setting="paper"]').value,
        background: document.querySelector('[data-print-setting="background"]').value,
        orientation: document.querySelector('.print-preview-sheet').dataset.orientation,
      }
      const paper = document.querySelector('[data-print-setting="paper"]')
      paper.value = 'letter'
      paper.dispatchEvent(new Event('change', { bubbles: true }))
      document.querySelector('button[aria-label="White background"]').click()
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }))
      return new Promise((resolve) => setTimeout(() => resolve({
        allOptionsDisabled: Array.from(document.querySelectorAll('[data-print-setting], .print-background-control button')).every((control) => control.matches(':disabled')),
        unchanged: document.querySelector('[data-print-setting="background"]').value === before.background
          && document.querySelector('.print-preview-sheet').dataset.orientation === before.orientation,
        dialogStayedOpen: Boolean(document.querySelector('.image-print-dialog')),
        busyText: document.querySelector('.print-submit').textContent,
      }), 80))
    })()`)
    await waitFor(window, `document.querySelector('.image-print-dialog') === null && document.activeElement?.getAttribute('title') === 'Print (Ctrl+P)'`)
    const submitFocusRestored = await window.webContents.executeJavaScript(`document.activeElement?.getAttribute('title') === 'Print (Ctrl+P)'`)

    const result = { source, defaultLayout, backgroundIsolation, toolbarFocusRestored, shortcutFocusRestored, configuredLayout, responsiveLayout, focusTrapped, frozenJob, submitFocusRestored, printRequest }
    if (!process.env.SIMPLE_IMAGE_SMOKE_NO_ARTIFACTS) {
      await fs.writeFile(path.join(__dirname, 'print-smoke-result.json'), JSON.stringify(result, null, 2), 'utf8')
    }
    if (
      defaultLayout.title !== 'Set up your print'
      || defaultLayout.paper !== 'letter'
      || defaultLayout.orientation !== 'portrait'
      || defaultLayout.scaleMode !== 'fit'
      || !defaultLayout.panesSideBySide
      || !defaultLayout.dialogFits
      || !defaultLayout.portraitSheet
      || !defaultLayout.imageLoaded
      || !backgroundIsolation.dimensionsUnchanged
      || !backgroundIsolation.titleUnchanged
      || !backgroundIsolation.noDropOverlay
      || !backgroundIsolation.backgroundInert
      || !toolbarFocusRestored
      || !shortcutFocusRestored
      || !configuredLayout.landscapeSheet
      || !configuredLayout.clippedWarning
      || !configuredLayout.blackBackground
      || !configuredLayout.grayscaleComposition
      || !configuredLayout.grayscaleLabelHonest
      || !configuredLayout.summary.includes('Effective resolution')
      || !responsiveLayout.stacked
      || !responsiveLayout.dialogFits
      || !responsiveLayout.bodyScrollable
      || !focusTrapped
      || !frozenJob.allOptionsDisabled
      || !frozenJob.unchanged
      || !frozenJob.dialogStayedOpen
      || !frozenJob.busyText.includes('Sending to printer')
      || !submitFocusRestored
      || printRequest?.name !== path.basename(fixturePath)
      || printRequest?.dimensions?.width !== source.width
      || printRequest?.dimensions?.height !== source.height
      || printRequest?.requestedDimensions?.width !== source.width
      || printRequest?.requestedDimensions?.height !== source.height
      || printRequest?.settings?.paper !== 'a4'
      || printRequest?.settings?.orientation !== 'landscape'
      || printRequest?.settings?.marginMm !== 0
      || printRequest?.settings?.scaleMode !== 'fill'
      || printRequest?.settings?.position !== 'bottom-right'
      || printRequest?.settings?.background !== '#000000'
      || printRequest?.settings?.grayscale !== true
      || printRequest?.settings?.copies !== 3
    ) throw new Error(`Print smoke assertions failed: ${JSON.stringify(result)}`)
    app.exit(0)
  } catch (error) {
    console.error(error)
    app.exit(1)
  }
})
