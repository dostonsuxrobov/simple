import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9333)
const deadline = Date.now() + 30_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Wait for Electron's debugging endpoint.
  }
  if (!target) await pause(25)
}
if (!target) throw new Error('Renderer did not start.')

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
  else request.resolve(message.result.result?.value)
})

function evaluate(expression) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, returnByValue: true } }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

async function waitFor(expression) {
  while (Date.now() < deadline) {
    const value = await evaluate(expression)
    if (value) return value
    await pause(30)
  }
  throw new Error(`Timed out: ${expression}`)
}

const readyExpression = `(() => {
  const current = document.querySelector('.continuous-page-slot.is-current')
  const canvas = current?.querySelector('.page-canvas')
  return Boolean(canvas && canvas.width > 1 && !current?.querySelector('.page-rendering'))
})()`
await waitFor(readyExpression)

const initial = await evaluate(`(() => ({
  page: document.querySelector('input[aria-label="Current page"]')?.value,
  pages: document.querySelector('.page-field span')?.textContent,
  thumbnails: document.querySelectorAll('.thumbnail-item').length,
  pageSlots: document.querySelectorAll('.continuous-page-slot').length,
  mountedCanvases: document.querySelectorAll('.continuous-page-slot .page-canvas').length,
}))()`)

await evaluate(`document.querySelector('button[aria-label="Next page (Page Down)"]')?.click()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '2' && ${readyExpression}`)

await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Current page"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, '700')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
})()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '700' && ${readyExpression}`)
await pause(350)

const final = await evaluate(`(() => ({
  page: document.querySelector('input[aria-label="Current page"]')?.value,
  thumbnails: document.querySelectorAll('.thumbnail-item').length,
  thumbnailLabels: Array.from(document.querySelectorAll('.thumbnail-meta > span')).map((item) => item.textContent),
  mountedCanvases: document.querySelectorAll('.continuous-page-slot .page-canvas').length,
  canvas: (() => {
    const canvas = document.querySelector('.continuous-page-slot.is-current .page-canvas')
    return [canvas?.width || 0, canvas?.height || 0]
  })(),
}))()`)

socket.close()
console.log(JSON.stringify({ initial, final }))
