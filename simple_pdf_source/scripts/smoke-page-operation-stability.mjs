import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9397)
const deadline = Date.now() + 35_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Wait for Electron's debugging endpoint.
  }
  if (!target) await pause(30)
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
  else if (message.result.exceptionDetails) request.reject(new Error(`${message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text}\nExpression: ${request.expression}`))
  else request.resolve(message.result.result?.value)
})

function evaluate(expression) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, returnByValue: true, awaitPromise: true } }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject, expression }))
}

async function waitFor(expression, label = expression) {
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(25)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 8 && document.querySelector('.continuous-page-slot[data-page-index="0"] .page-canvas')?.width > 300 && !document.querySelector('.continuous-page-slot[data-page-index="0"] .page-rendering')`, 'initial PDF pages')
await pause(250)

const before = await evaluate(`(() => {
  const canvas = document.querySelector('.continuous-page-slot[data-page-index="0"] .page-canvas')
  return { width: canvas.width, height: canvas.height }
})()`)

await evaluate(`document.querySelector('button[aria-label="Rotate page 1 right"]')?.click()`)
await waitFor(`(() => {
  const canvas = document.querySelector('.continuous-page-slot[data-page-index="0"] .page-canvas')
  return canvas && canvas.width > canvas.height
})()`, 'rotated first page')

await evaluate(`(() => {
  window.confirm = () => true
  const shell = document.querySelector('.continuous-pages')
  const canvas = document.querySelector('.continuous-page-slot[data-page-index="0"] .page-canvas')
  const viewer = document.querySelector('.viewer')
  const firstSlot = document.querySelector('.continuous-page-slot[data-page-index="0"]')
  const viewerBounds = viewer.getBoundingClientRect()
  const firstSlotBounds = firstSlot.getBoundingClientRect()
  window.__pageOperationStability = {
    shell,
    canvas,
    minimumPageSlots: document.querySelectorAll('.continuous-page-slot').length,
    shellRemoved: false,
    canvasRemoved: false,
    initialLoaderSeen: false,
    firstPageLoadingSeen: false,
    firstPageRenderingVeilSeen: false,
    updatingSeen: false,
    startViewportOffset: (viewerBounds.top - firstSlotBounds.top) / firstSlotBounds.height,
  }
  const sample = () => {
    const result = window.__pageOperationStability
    result.minimumPageSlots = Math.min(result.minimumPageSlots, document.querySelectorAll('.continuous-page-slot').length)
    result.shellRemoved ||= !result.shell?.isConnected
    result.canvasRemoved ||= !result.canvas?.isConnected
    result.initialLoaderSeen ||= Boolean(document.querySelector('.document-loading'))
    result.firstPageLoadingSeen ||= Boolean(document.querySelector('.continuous-page-slot[data-page-index="0"] .page-loading'))
    result.firstPageRenderingVeilSeen ||= Boolean(document.querySelector('.continuous-page-slot[data-page-index="0"] .page-rendering'))
    result.updatingSeen ||= Boolean(document.querySelector('.document-refreshing'))
  }
  window.__pageOperationObserver = new MutationObserver(sample)
  window.__pageOperationObserver.observe(document.body, { childList: true, subtree: true })
  window.__pageOperationSampler = window.setInterval(sample, 1)
  sample()
})()`)

await evaluate(`document.querySelector('button[aria-label="Delete page 1"]')?.click()`)
await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 7 && !document.querySelector('.busy-overlay') && !document.querySelector('.document-refreshing') && !document.querySelector('.continuous-page-slot[data-page-index="0"] .is-page-transitioning')`, 'page deletion and replacement render')
await pause(80)

const deletion = await evaluate(`(() => {
  window.__pageOperationObserver?.disconnect()
  window.clearInterval(window.__pageOperationSampler)
  const result = window.__pageOperationStability
  const viewerBounds = (result.shell.closest('.viewer') || result.shell.parentElement).getBoundingClientRect()
  const firstSlotBounds = result.shell.querySelector('.continuous-page-slot[data-page-index="0"]').getBoundingClientRect()
  return {
    minimumPageSlots: result.minimumPageSlots,
    shellRemoved: result.shellRemoved,
    canvasRemoved: result.canvasRemoved,
    initialLoaderSeen: result.initialLoaderSeen,
    firstPageLoadingSeen: result.firstPageLoadingSeen,
    firstPageRenderingVeilSeen: result.firstPageRenderingVeilSeen,
    updatingSeen: result.updatingSeen,
    startViewportOffset: result.startViewportOffset,
    finalViewportOffset: (viewerBounds.top - firstSlotBounds.top) / firstSlotBounds.height,
    pageSlots: document.querySelectorAll('.continuous-page-slot').length,
    shellSame: result.shell === document.querySelector('.continuous-pages'),
    canvasSame: result.canvas === document.querySelector('.continuous-page-slot[data-page-index="0"] .page-canvas'),
    currentPage: document.querySelector('input[aria-label="Current page"]')?.value,
  }
})()`)

if (deletion.minimumPageSlots < 1 || deletion.shellRemoved || deletion.canvasRemoved || deletion.initialLoaderSeen || deletion.firstPageLoadingSeen || deletion.firstPageRenderingVeilSeen || !deletion.shellSame || !deletion.canvasSame) {
  throw new Error(`Delete tore down the visible document surface: ${JSON.stringify({ before, deletion })}`)
}
if (Math.abs(deletion.finalViewportOffset - deletion.startViewportOffset) > 0.08) {
  throw new Error(`Delete changed the viewport anchor: ${JSON.stringify({ before, deletion })}`)
}
if (deletion.pageSlots !== 7 || deletion.currentPage !== '1') {
  throw new Error(`Delete state was incorrect: ${JSON.stringify({ before, deletion })}`)
}

await evaluate(`document.querySelector('button[aria-label="More page actions"]')?.click()`)
await waitFor(`Array.from(document.querySelectorAll('[role="menuitem"]')).some((button) => button.textContent.includes('Insert blank page'))`, 'blank-page command')
await evaluate(`Array.from(document.querySelectorAll('[role="menuitem"]')).find((button) => button.textContent.includes('Insert blank page'))?.click()`)
await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 8 && document.querySelector('input[aria-label="Current page"]')?.value === '2' && !document.querySelector('.document-update-guard')`, 'blank page atomic target')

await evaluate(`document.querySelector('button[aria-label="More page actions"]')?.click()`)
await waitFor(`Array.from(document.querySelectorAll('[role="menuitem"]')).some((button) => button.textContent.includes('Duplicate selected'))`, 'duplicate-page command')
await evaluate(`Array.from(document.querySelectorAll('[role="menuitem"]')).find((button) => button.textContent.includes('Duplicate selected'))?.click()`)
await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 9 && document.querySelector('input[aria-label="Current page"]')?.value === '3' && !document.querySelector('.document-update-guard')`, 'duplicate page atomic target')
await pause(80)

const after = await evaluate(`(() => {
  return {
    finalPageSlots: document.querySelectorAll('.continuous-page-slot').length,
    currentPage: document.querySelector('input[aria-label="Current page"]')?.value,
    dirty: Boolean(document.querySelector('.dirty-dot')),
  }
})()`)

socket.close()

if (after.finalPageSlots !== 9 || after.currentPage !== '3' || !after.dirty) {
  throw new Error(`Structural operation state was incorrect: ${JSON.stringify({ before, deletion, after })}`)
}

console.log(JSON.stringify({ before, deletion, after }))
