'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')

const projectDirectory = path.resolve(__dirname, '..')
const electronPath = require('electron')
const port = Number(process.env.SIMPLE_IMAGE_CONTEXT_QA_PORT || 9876)

function delay(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds))
}

async function terminateChild(child) {
  if (!child || child.exitCode !== null) return
  if (process.platform === 'win32' && child.pid) {
    spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], {
      windowsHide: true,
      stdio: 'ignore',
      timeout: 5_000,
    })
    return
  }
  child.kill('SIGKILL')
  await Promise.race([
    new Promise((resolve) => child.once('exit', resolve)),
    delay(2_000),
  ])
}

async function waitForDebugger() {
  const deadline = Date.now() + 20_000
  while (Date.now() < deadline) {
    try {
      const pages = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      const page = pages.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
      if (page) return page
    } catch {}
    await delay(125)
  }
  throw new Error('Timed out waiting for the Simple Image renderer.')
}

async function connectCdp(url) {
  const socket = new WebSocket(url)
  const pending = new Map()
  let sequence = 0
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data))
    if (!message.id || !pending.has(message.id)) return
    const { resolve, reject } = pending.get(message.id)
    pending.delete(message.id)
    if (message.error) reject(new Error(message.error.message))
    else resolve(message.result)
  })
  return {
    close: () => socket.close(),
    send(method, params = {}) {
      const id = ++sequence
      return new Promise((resolve, reject) => {
        pending.set(id, { resolve, reject })
        socket.send(JSON.stringify({ id, method, params }))
      })
    },
  }
}

async function main() {
  const temporaryRoot = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-context-'))
  const profileDirectory = path.join(temporaryRoot, 'profile')
  const fixturePath = process.env.SIMPLE_IMAGE_CONTEXT_QA_FILE || path.join(projectDirectory, 'public', 'brand-icon.png')
  let child
  let cdp
  let stderr = ''

  try {
    child = spawn(electronPath, [
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${profileDirectory}`,
      projectDirectory,
      fixturePath,
    ], {
      cwd: projectDirectory,
      env: { ...process.env, ELECTRON_DISABLE_SECURITY_WARNINGS: 'true' },
      stdio: ['ignore', 'ignore', 'pipe'],
      windowsHide: true,
    })
    child.stderr.on('data', (chunk) => { stderr = `${stderr}${chunk}`.slice(-12_000) })

    const page = await waitForDebugger()
    cdp = await connectCdp(page.webSocketDebuggerUrl)
    await cdp.send('Runtime.enable')
    await cdp.send('Page.bringToFront')

    const evaluate = async (expression) => {
      const response = await cdp.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true })
      if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text)
      return response.result?.value
    }
    const waitFor = async (expression, label) => {
      const deadline = Date.now() + 15_000
      while (Date.now() < deadline) {
        if (await evaluate(expression)) return
        await delay(100)
      }
      throw new Error(`Timed out waiting for ${label}.`)
    }
    const openContextMenu = async () => {
      const dispatch = await evaluate(`(() => {
        const target = document.querySelector('.canvas-stage')
        if (!target) throw new Error('The image canvas stage is unavailable.')
        const event = new MouseEvent('contextmenu', {
          bubbles: true,
          cancelable: true,
          button: 2,
          buttons: 2,
          clientX: innerWidth - 1,
          clientY: innerHeight - 1,
        })
        target.dispatchEvent(event)
        return { prevented: event.defaultPrevented, width: innerWidth, height: innerHeight }
      })()`)
      assert.equal(dispatch.prevented, true)
      await waitFor(`Boolean(document.querySelector('.image-context-menu'))`, 'the image context menu')
      await waitFor(`Boolean(document.activeElement?.closest('.image-context-menu'))`, 'context-menu focus')
      return dispatch
    }

    await waitFor(`document.querySelector('canvas')?.width > 0 && !document.querySelector('.canvas-stage')?.hidden`, 'the decoded image')
    const original = await evaluate(`(() => {
      const canvas = document.querySelector('canvas')
      return { width: canvas.width, height: canvas.height }
    })()`)

    await evaluate(`document.querySelector('button[title="Crop"]')?.click()`)
    await waitFor(`Boolean(document.querySelector('.crop-box'))`, 'the non-destructive crop selection')
    const selection = await evaluate(`(() => {
      const canvas = document.querySelector('canvas')
      const box = document.querySelector('.crop-box').getBoundingClientRect()
      return { canvasWidth: canvas.width, canvasHeight: canvas.height, visibleWidth: box.width, visibleHeight: box.height }
    })()`)

    const viewport = await openContextMenu()
    const menu = await evaluate(`(() => {
      const menu = document.querySelector('.image-context-menu')
      const rect = menu.getBoundingClientRect()
      return {
        label: menu.getAttribute('aria-label'),
        items: [...menu.querySelectorAll('[role="menuitem"]')].map((item) => ({
          label: item.querySelector('span')?.textContent?.trim() || '',
          disabled: item.disabled,
        })),
        active: document.activeElement?.querySelector('span')?.textContent?.trim() || '',
        bounds: { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom },
      }
    })()`)
    assert.equal(menu.label, 'Image selection menu')
    assert.deepEqual(menu.items, [
      { label: 'Copy selected area', disabled: false },
      { label: 'Undo', disabled: true },
      { label: 'Redo', disabled: true },
      { label: 'Apply crop', disabled: false },
    ])
    assert.equal(menu.active, 'Copy selected area')
    assert(menu.bounds.left >= 0 && menu.bounds.top >= 0)
    assert(menu.bounds.right <= viewport.width && menu.bounds.bottom <= viewport.height)

    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', code: 'ArrowDown', bubbles: true, cancelable: true }))`)
    assert.equal(await evaluate(`document.activeElement?.querySelector('span')?.textContent?.trim()`), 'Apply crop')
    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', code: 'Home', bubbles: true, cancelable: true }))`)
    assert.equal(await evaluate(`document.activeElement?.querySelector('span')?.textContent?.trim()`), 'Copy selected area')
    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', bubbles: true, cancelable: true }))`)
    assert.equal(await evaluate(`document.activeElement?.querySelector('span')?.textContent?.trim()`), 'Apply crop')
    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', code: 'Tab', shiftKey: true, bubbles: true, cancelable: true }))`)
    assert.equal(await evaluate(`document.activeElement?.querySelector('span')?.textContent?.trim()`), 'Copy selected area')
    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', bubbles: true, cancelable: true }))`)
    await waitFor(`!document.querySelector('.image-context-menu')`, 'the image context menu to close with Escape')

    await openContextMenu()
    await evaluate(`(() => {
      const button = [...document.querySelectorAll('.image-context-menu [role="menuitem"]')]
        .find((item) => item.querySelector('span')?.textContent?.trim() === 'Copy selected area')
      if (!button || button.disabled) throw new Error('Copy selected area is unavailable.')
      button.click()
    })()`)
    await waitFor(`document.querySelector('.toast')?.textContent?.includes('Selected area copied to the clipboard.')`, 'the selected-area clipboard confirmation')
    const clipboardBridge = await evaluate(`(async () => {
      const canvas = document.querySelector('canvas')
      const blob = await new Promise((resolve, reject) => canvas.toBlob((value) => value ? resolve(value) : reject(new Error('PNG encoding failed.')), 'image/png'))
      return window.simpleImage.copyPng(new Uint8Array(await blob.arrayBuffer()))
    })()`)
    assert.deepEqual(clipboardBridge, { width: original.width, height: original.height })

    const result = {
      fixture: path.basename(fixturePath),
      original,
      selection,
      menu: { labels: menu.items.map((item) => item.label), keyboardNavigation: true, viewportClamped: true },
      clipboardBridge: { ...clipboardBridge, actionConfirmed: true },
    }
    process.stdout.write(`${JSON.stringify(result, null, 2)}\n`)
  } catch (error) {
    if (stderr) process.stderr.write(stderr)
    throw error
  } finally {
    cdp?.close()
    await terminateChild(child)
    const safePrefix = path.join(os.tmpdir(), 'simple-image-context-')
    if (temporaryRoot.startsWith(safePrefix)) await fs.rm(temporaryRoot, { recursive: true, force: true, maxRetries: 5, retryDelay: 120 })
  }
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
