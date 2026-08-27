import process from 'node:process'

const port = Number(process.env.SIMPLE_QA_PORT || 9383)
const appUrl = process.env.SIMPLE_QA_URL || 'http://127.0.0.1:5174/'
const deadline = Date.now() + 120_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
const target = targets.find((item) => item.type === 'page' && (item.url === 'about:blank' || item.url.startsWith(appUrl)))
if (!target?.webSocketDebuggerUrl) throw new Error('No QA browser target is available.')

const socket = new WebSocket(target.webSocketDebuggerUrl)
await new Promise((resolve, reject) => {
  socket.addEventListener('open', resolve, { once: true })
  socket.addEventListener('error', reject, { once: true })
})

let requestId = 0
const pending = new Map()
socket.addEventListener('message', (event) => {
  const message = JSON.parse(String(event.data))
  const request = pending.get(message.id)
  if (!request) return
  pending.delete(message.id)
  if (message.error) request.reject(new Error(message.error.message))
  else request.resolve(message.result)
})

function call(method, params = {}) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method, params }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

async function evaluate(expression) {
  const result = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  if (result.exceptionDetails) {
    throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text)
  }
  return result.result?.value
}

async function waitFor(expression, label) {
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(60)
  }
  throw new Error(`Timed out waiting for ${label}.`)
}

async function pressKey(key, code, virtualKeyCode, modifiers = 0) {
  const params = { key, code, windowsVirtualKeyCode: virtualKeyCode, nativeVirtualKeyCode: virtualKeyCode, modifiers }
  await call('Input.dispatchKeyEvent', { ...params, type: 'rawKeyDown' })
  await call('Input.dispatchKeyEvent', { ...params, type: 'keyUp' })
}

async function click(selector) {
  const point = await evaluate(`(() => {
    const element = document.querySelector(${JSON.stringify(selector)})
    if (!element) return null
    const bounds = element.getBoundingClientRect()
    return { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2, width: bounds.width, height: bounds.height }
  })()`)
  if (!point?.width || !point?.height) throw new Error(`Element is not clickable: ${selector}`)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: point.x, y: point.y })
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: point.x, y: point.y, button: 'left', clickCount: 1 })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: point.x, y: point.y, button: 'left', clickCount: 1 })
  return point
}

async function setInput(selector, value) {
  await evaluate(`(() => {
    const input = document.querySelector(${JSON.stringify(selector)})
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    setter.call(input, ${JSON.stringify(value)})
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.focus()
  })()`)
}

const bridge = `Object.defineProperty(window, 'simple', {
  configurable: true,
  value: {
    onMaximized: () => () => {}, onOpenExternal: () => () => {}, onCloseRequested: () => () => {},
    minimize: () => {}, toggleMaximize: () => {}, close: () => {},
    openFile: async () => {
      const response = await fetch('${appUrl}@fs/C:/Users/dosto/Documents/.archive/pdf_app/tmp/pdfs/simple-interaction-fixture.pdf')
      return { data: await response.arrayBuffer(), name: 'simple-interaction-fixture.pdf', converted: false, signatureDetected: false }
    },
    openInNewWindow: async () => false, openPath: async () => { throw new Error('unused') }, openBytes: async () => { throw new Error('unused') },
    savePdf: async () => ({ canceled: true }), showItem: () => {}, flattenOverlays: async (bytes) => bytes,
    printPdf: async () => {}, insertFiles: async (bytes) => bytes, insertDroppedFiles: async (bytes) => bytes,
    mutatePdf: async (bytes) => bytes, exportPages: async () => null, startPageDrag: async () => {}, pickImage: async () => null,
  },
})`

await call('Page.enable')
await call('Page.addScriptToEvaluateOnNewDocument', { source: bridge })
await call('Page.navigate', { url: `${appUrl}?search-qa=${Date.now()}` })
await waitFor(`Boolean(document.querySelector('.drop-card .button'))`, 'welcome screen')
await click('.drop-card .button')
await waitFor(`document.querySelector('.page-field span')?.textContent?.includes('/ 8')`, 'eight-page PDF')

const sentence = 'Select only these four words, not the complete line.'
const sourcePage = 1

await pressKey('F4', 'F4', 115)
await waitFor(`!document.querySelector('.sidebar')`, 'closed sidebar')
await pressKey('f', 'KeyF', 70, 2)
await waitFor(`document.activeElement?.getAttribute('aria-label') === 'Find in document'`, 'Ctrl+F search focus')
const opened = await evaluate(`({
  sidebar: Boolean(document.querySelector('.sidebar')),
  selected: document.querySelector('button[aria-label="Search"]')?.getAttribute('aria-selected'),
  focused: document.activeElement?.getAttribute('aria-label'),
})`)

await setInput('input[aria-label="Find in document"]', sentence)
await pressKey('Enter', 'Enter', 13)
await waitFor(`!document.querySelector('.search-status')?.textContent?.includes('Searching') && document.querySelectorAll('.search-results > button').length > 0`, 'sentence search results')
const sentenceResult = await evaluate(`({
  query: document.querySelector('input[aria-label="Find in document"]')?.value,
  status: document.querySelector('.search-status')?.textContent,
  pages: Array.from(document.querySelectorAll('.search-results > button > span')).map((item) => item.textContent),
})`)

await pressKey('f', 'KeyF', 70, 2)
await pause(100)
const repeated = await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Find in document"]')
  return { focused: document.activeElement === input, selection: [input.selectionStart, input.selectionEnd], length: input.value.length }
})()`)

await setInput('input[aria-label="Find in document"]', 'the')
await pressKey('Enter', 'Enter', 13)
await waitFor(`!document.querySelector('.search-status')?.textContent?.includes('Searching') && document.querySelectorAll('.search-results > button').length > 1`, 'multi-page search results')
const first = await evaluate(`({ page: document.querySelector('input[aria-label="Current page"]')?.value, active: document.querySelector('.search-results > button.is-current > span')?.textContent })`)
await pressKey('Enter', 'Enter', 13)
await pause(400)
const next = await evaluate(`({ page: document.querySelector('input[aria-label="Current page"]')?.value, active: document.querySelector('.search-results > button.is-current > span')?.textContent })`)
await pressKey('Enter', 'Enter', 13, 1)
await pause(400)
const previous = await evaluate(`({ page: document.querySelector('input[aria-label="Current page"]')?.value, active: document.querySelector('.search-results > button.is-current > span')?.textContent })`)

await pressKey('Escape', 'Escape', 27)
await pause(150)
const dismissed = await evaluate(`({
  input: Boolean(document.querySelector('input[aria-label="Find in document"]')),
  searchSelected: document.querySelector('button[aria-label="Search"]')?.getAttribute('aria-selected'),
  selectedTab: document.querySelector('.sidebar-tabs button[aria-selected="true"]')?.textContent?.trim(),
})`)

const closePoint = await evaluate(`(() => {
  const bounds = document.querySelector('button[aria-label="Close"]').getBoundingClientRect()
  return { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
})()`)
await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: closePoint.x, y: closePoint.y })
await pause(100)
const closeHover = await evaluate(`(() => {
  const style = getComputedStyle(document.querySelector('button[aria-label="Close"]'))
  return { background: style.backgroundColor, color: style.color }
})()`)

await click('button[aria-label="Immersive reading (Esc to exit)"]')
await waitFor(`Boolean(document.fullscreenElement)`, 'immersive fullscreen')
await pressKey('Escape', 'Escape', 27)
await waitFor(`!document.fullscreenElement && !document.querySelector('.document-app').classList.contains('is-immersive')`, 'immersive Escape')
const immersiveEscape = await evaluate(`({ fullscreen: Boolean(document.fullscreenElement), toolbar: getComputedStyle(document.querySelector('.toolbar')).display })`)

const checks = {
  ctrlF: opened.sidebar && opened.selected === 'true' && opened.focused === 'Find in document',
  sentence: sentenceResult.query === sentence && sentenceResult.pages.length > 0,
  repeatedCtrlF: repeated.focused && repeated.selection[0] === 0 && repeated.selection[1] === repeated.length,
  next: next.page !== first.page,
  previous: previous.page === first.page,
  escapeDismiss: !dismissed.input && dismissed.searchSelected === 'false',
  closeHover: closeHover.background === 'rgb(232, 17, 35)' && closeHover.color === 'rgb(255, 255, 255)',
  immersiveEscape: !immersiveEscape.fullscreen && immersiveEscape.toolbar === 'grid',
}

socket.close()
console.log(JSON.stringify({ sourcePage, sentence, opened, sentenceResult, repeated, first, next, previous, dismissed, closeHover, immersiveEscape, checks }, null, 2))
if (Object.values(checks).some((passed) => !passed)) process.exitCode = 1
