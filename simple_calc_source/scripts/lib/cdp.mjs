// Minimal Chrome DevTools Protocol driver for the running dev app
// (electron . --remote-debugging-port=9385 with VITE_DEV_SERVER_URL set).
import fs from 'node:fs/promises'

const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

export async function connect({ port = Number(process.env.SIMPLE_CALC_QA_PORT || 9385), timeout = 60_000 } = {}) {
  const deadline = Date.now() + timeout
  let target
  while (Date.now() < deadline) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      target = targets.find((item) => item.type === 'page' && item.url.startsWith('http://127.0.0.1:5173'))
      if (target?.webSocketDebuggerUrl) break
    } catch {}
    await pause(250)
  }
  if (!target?.webSocketDebuggerUrl) throw new Error('No simple_calc page target is available.')
  const socket = new WebSocket(target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  let requestId = 0
  const pending = new Map()
  const errors = []
  const logs = []
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data))
    if (message.method === 'Runtime.exceptionThrown') errors.push(message.params?.exceptionDetails?.exception?.description || message.params?.exceptionDetails?.text)
    if (message.method === 'Runtime.consoleAPICalled') logs.push(message.params.args.map((arg) => arg.value ?? arg.description).join(' '))
    const request = pending.get(message.id)
    if (!request) return
    pending.delete(message.id)
    if (message.error) request.reject(new Error(message.error.message))
    else request.resolve(message.result)
  })
  const call = (method, params = {}) => {
    const id = ++requestId
    socket.send(JSON.stringify({ id, method, params }))
    return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
  }
  await call('Runtime.enable')
  // Native alert/confirm dialogs would block the renderer; accept them automatically.
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data))
    if (message.method === 'Page.javascriptDialogOpening') {
      logs.push(`[dialog] ${message.params.message}`)
      socket.send(JSON.stringify({ id: ++requestId, method: 'Page.handleJavaScriptDialog', params: { accept: true } }))
    }
  })
  await call('Page.enable')
  const evaluate = async (expression) => {
    const result = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text)
    return result.result?.value
  }
  const waitFor = async (expression, label, limit = 20_000) => {
    const end = Date.now() + limit
    while (Date.now() < end) {
      if (await evaluate(`Boolean(${expression})`)) return
      await pause(80)
    }
    throw new Error(`Timed out waiting for ${label}`)
  }
  const screenshot = async (file) => {
    const { data } = await call('Page.captureScreenshot', { format: 'png' })
    await fs.writeFile(file, Buffer.from(data, 'base64'))
  }
  const key = async (keyName, { code, text, modifiers = 0, keyCode } = {}) => {
    const base = { key: keyName, code: code || keyName, modifiers, windowsVirtualKeyCode: keyCode, nativeVirtualKeyCode: keyCode }
    await call('Input.dispatchKeyEvent', { type: text ? 'keyDown' : 'rawKeyDown', ...base, text })
    await call('Input.dispatchKeyEvent', { type: 'keyUp', ...base })
  }
  const type = (text) => call('Input.insertText', { text })
  const mouse = async (type, x, y, { button = 'left', buttons = 1, clickCount = 1, modifiers = 0 } = {}) => {
    await call('Input.dispatchMouseEvent', { type, x, y, button, buttons: type === 'mouseReleased' ? 0 : buttons, clickCount, modifiers })
  }
  const click = async (x, y, options) => {
    await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y, button: 'none', buttons: 0, modifiers: options?.modifiers || 0 })
    await mouse('mousePressed', x, y, options)
    await mouse('mouseReleased', x, y, options)
  }
  const cellCenter = (address) => evaluate(`(() => {
    const cell = [...document.querySelectorAll('.grid-cell')].find((item) => item.dataset.cellAddress === ${JSON.stringify(address)} && !item.getAttribute('aria-hidden'));
    if (!cell) return null;
    const box = cell.getBoundingClientRect();
    return { x: box.left + box.width / 2, y: box.top + box.height / 2 };
  })()`)
  const cellText = (address) => evaluate(`(() => {
    const cell = [...document.querySelectorAll('.grid-cell')].find((item) => item.dataset.cellAddress === ${JSON.stringify(address)} && !item.getAttribute('aria-hidden'));
    return cell ? cell.textContent : null;
  })()`)
  const close = () => socket.close()
  const onMessage = (listener) => socket.addEventListener('message', (event) => listener(JSON.parse(String(event.data))))
  return { call, evaluate, waitFor, screenshot, key, type, mouse, click, cellCenter, cellText, errors, logs, close, pause, onMessage }
}

export { pause }
