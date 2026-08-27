import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9385)
const requestedStart = Number(process.env.SIMPLE_SELECTION_START_INDEX || 10)
const pagesToTraverse = Number(process.env.SIMPLE_SELECTION_PAGE_COUNT || 6)
const deadline = Date.now() + 60_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Electron is still opening its debugging endpoint.
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
    await pause(35)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`document.querySelectorAll('.continuous-page-slot').length > ${requestedStart + pagesToTraverse}`, 'continuous document')

let startIndex = requestedStart
for (; startIndex < requestedStart + 30; startIndex += 1) {
  await evaluate(`document.querySelector('.continuous-page-slot[data-page-index="${startIndex}"]')?.scrollIntoView({ block: 'center' })`)
  await pause(450)
  if (await evaluate(`document.querySelectorAll('.continuous-page-slot[data-page-index="${startIndex}"] .text-layer span').length > 0`)) break
}
if (startIndex >= requestedStart + 30) throw new Error('Could not find a textual page for the selection test.')

await waitFor(`document.querySelector('.continuous-page-slot[data-page-index="${startIndex}"]')?.classList.contains('is-current')`, 'start page tracking')
await pause(800)

const initial = await evaluate(`(() => {
  const slot = document.querySelector('.continuous-page-slot[data-page-index="${startIndex}"]')
  const node = slot?.querySelector('.text-layer span')?.firstChild
  if (!(node instanceof Text)) return null
  const range = document.createRange()
  range.setStart(node, 0)
  range.setEnd(node, Math.min(5, node.length))
  const selection = getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  window.__simpleLongSelectionAnchor = node
  return {
    current: Number(document.querySelector('.continuous-page-slot.is-current')?.dataset.pageIndex),
    selectedLength: selection.toString().length,
    anchorConnected: node.isConnected,
  }
})()`)
if (!initial?.anchorConnected || initial.selectedLength < 1) throw new Error(`Could not start selection: ${JSON.stringify(initial)}`)
await pause(100)

const samples = []
for (let index = startIndex + 1; index <= startIndex + pagesToTraverse; index += 1) {
  await evaluate(`document.querySelector('.continuous-page-slot[data-page-index="${index}"]')?.scrollIntoView({ block: 'center' })`)
  await waitFor(`document.querySelectorAll('.continuous-page-slot[data-page-index="${index}"] .text-layer span').length > 0`, `text layer on page ${index + 1}`)
  await evaluate(`(() => {
    const spans = document.querySelectorAll('.continuous-page-slot[data-page-index="${index}"] .text-layer span')
    const node = spans[spans.length - 1]?.firstChild
    const selection = getSelection()
    if (node instanceof Text && selection?.rangeCount) selection.extend(node, node.length)
  })()`)
  await pause(350)
  const state = await evaluate(`(() => {
    const selection = getSelection()
    const mounted = [...document.querySelectorAll('.continuous-page-slot .page-surface')]
      .map((surface) => Number(surface.closest('.continuous-page-slot')?.dataset.pageIndex))
    return {
      target: ${index},
      current: Number(document.querySelector('.continuous-page-slot.is-current')?.dataset.pageIndex),
      selectedLength: selection?.toString().length || 0,
      collapsed: selection?.isCollapsed ?? true,
      anchorConnected: Boolean(window.__simpleLongSelectionAnchor?.isConnected),
      anchorMounted: mounted.includes(${startIndex}),
      mountedCount: mounted.length,
    }
  })()`)
  samples.push(state)
  if (!state.anchorConnected || !state.anchorMounted || state.collapsed || state.selectedLength < initial.selectedLength) {
    throw new Error(`Selection destabilized while traversing pages: ${JSON.stringify({ initial, samples })}`)
  }
}

const finalTarget = startIndex + pagesToTraverse
await evaluate(`getSelection()?.removeAllRanges()`)
await waitFor(`Number(document.querySelector('.continuous-page-slot.is-current')?.dataset.pageIndex) === ${finalTarget}`, 'current page catch-up after selection')
await pause(500)

const released = await evaluate(`(() => {
  const mounted = [...document.querySelectorAll('.continuous-page-slot .page-surface')]
    .map((surface) => Number(surface.closest('.continuous-page-slot')?.dataset.pageIndex))
  return {
    current: Number(document.querySelector('.continuous-page-slot.is-current')?.dataset.pageIndex),
    anchorConnected: Boolean(window.__simpleLongSelectionAnchor?.isConnected),
    anchorMounted: mounted.includes(${startIndex}),
    mountedCount: mounted.length,
  }
})()`)
if (released.anchorConnected || released.anchorMounted) throw new Error(`Selection pages were not pruned after release: ${JSON.stringify(released)}`)

socket.close()
console.log(JSON.stringify({ startIndex, pagesToTraverse, initial, samples, released }))
