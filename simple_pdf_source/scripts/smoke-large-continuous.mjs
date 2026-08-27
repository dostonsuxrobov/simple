import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9335)
const expectedPages = Number(process.env.SIMPLE_EXPECTED_PAGES || 1557)
const targetPage = Number(process.env.SIMPLE_TARGET_PAGE || Math.max(1, expectedPages - 57))
const deadline = Date.now() + 90_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Electron is still opening its debugging endpoint.
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
  else if (message.result.exceptionDetails) request.reject(new Error(message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text))
  else request.resolve(message.result.result?.value)
})

function evaluate(expression) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, returnByValue: true, awaitPromise: true } }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

async function waitFor(expression, label) {
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(30)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`Boolean(document.querySelector('.document-app')) || Boolean(document.querySelector('.document-error'))`, 'document shell')
const loadedState = await evaluate(`(() => ({
  slots: document.querySelectorAll('.continuous-page-slot').length,
  total: document.querySelector('.page-field span')?.textContent || '',
  error: document.querySelector('.document-error')?.textContent?.trim() || '',
  loading: Boolean(document.querySelector('.busy-overlay')),
}))()`)
if (loadedState.error) throw new Error(loadedState.error)
if (loadedState.slots !== expectedPages) throw new Error(`Expected ${expectedPages} page slots, found ${loadedState.slots}; state=${JSON.stringify(loadedState)}`)
await waitFor(`Boolean(document.querySelector('.continuous-page-slot.is-current .page-canvas'))`, 'first page canvas')

const initial = await evaluate(`(() => ({
  page: Number(document.querySelector('input[aria-label="Current page"]')?.value),
  slots: document.querySelectorAll('.continuous-page-slot').length,
  mountedCanvases: document.querySelectorAll('.continuous-page-slot .page-canvas').length,
  thumbnailRows: document.querySelectorAll('.thumbnail-virtual-row').length,
}))()`)

const jumpStartedAt = Date.now()
await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Current page"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, '${targetPage}')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
})()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '${targetPage}'`, 'target page navigation')
await waitFor(`Boolean(document.querySelector('.continuous-page-slot[data-page-index="${targetPage - 1}"] .page-canvas'))`, 'target page canvas')
await pause(250)

const final = await evaluate(`(() => ({
  page: Number(document.querySelector('input[aria-label="Current page"]')?.value),
  mountedCanvases: document.querySelectorAll('.continuous-page-slot .page-canvas').length,
  thumbnailRows: document.querySelectorAll('.thumbnail-virtual-row').length,
  targetIsCurrent: document.querySelector('.continuous-page-slot[data-page-index="${targetPage - 1}"]')?.classList.contains('is-current'),
}))()`)

if (initial.mountedCanvases > 12 || final.mountedCanvases > 12) throw new Error(`Canvas virtualization regressed: ${initial.mountedCanvases} -> ${final.mountedCanvases}`)
if (initial.thumbnailRows > 12 || final.thumbnailRows > 12) throw new Error(`Thumbnail virtualization regressed: ${initial.thumbnailRows} -> ${final.thumbnailRows}`)
if (!final.targetIsCurrent) throw new Error('Jumped page is not the current continuous-view page.')

socket.close()
console.log(JSON.stringify({ initial, targetPage, jumpMs: Date.now() - jumpStartedAt, final }))
