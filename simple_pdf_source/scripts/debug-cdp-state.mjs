const port = Number(process.env.SIMPLE_BENCH_PORT || 9380)
const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
const target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
if (!target) throw new Error('Renderer target missing')

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
  request(message)
})

function call(method, params = {}) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method, params }))
  return new Promise((resolve) => pending.set(id, resolve))
}

async function evaluate(expression) {
  const message = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  return message.result?.result?.value
}

console.log(JSON.stringify(await evaluate(`(async () => {
  const event = new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, cancelable: true })
  const dispatched = window.dispatchEvent(event)
  document.querySelector('.save-button')?.click()
  await new Promise((resolve) => setTimeout(resolve, 2500))
  return ({
  dispatched,
  prevented: event.defaultPrevented,
  eventKey: event.key,
  eventCtrl: event.ctrlKey,
  title: document.title,
  shell: Boolean(document.querySelector('.document-app')),
  slots: document.querySelectorAll('.continuous-page-slot').length,
  sidebar: Boolean(document.querySelector('.sidebar')),
  search: Boolean(document.querySelector('.search-pane')),
  dirty: Boolean(document.querySelector('.dirty-dot')),
  busy: document.querySelector('.busy-overlay')?.textContent?.trim() || '',
  toast: document.querySelector('.toast')?.textContent?.trim() || '',
  saveDisabled: document.querySelector('.save-button')?.disabled,
  activeTag: document.activeElement?.tagName,
  activeLabel: document.activeElement?.getAttribute('aria-label'),
  activeHtml: document.activeElement?.outerHTML?.slice(0, 300),
  })
})()`), null, 2))
const windowObject = await call('Runtime.evaluate', { expression: 'window', returnByValue: false })
const listeners = await call('DOMDebugger.getEventListeners', { objectId: windowObject.result.result.objectId })
console.log(JSON.stringify((listeners.result.listeners || []).filter((listener) => listener.type === 'keydown'), null, 2))
socket.close()
