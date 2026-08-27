import fs from 'node:fs/promises'
import path from 'node:path'
import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9394)
const deadline = Date.now() + 35_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))
const droppedPageBase64 = (await fs.readFile(path.resolve('tmp/pdfs/simple-drop-page.pdf'))).toString('base64')

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
    await pause(35)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 8 && document.querySelector('.page-canvas')?.width > 1`, 'initial PDF pages')

await evaluate(`(() => {
  const pageThree = document.querySelector('.continuous-page-slot[data-page-index="2"] .page-transfer-handle')
  pageThree.dispatchEvent(new MouseEvent('click', { bubbles: true, ctrlKey: true }))
})()`)
await waitFor(`document.querySelectorAll('.continuous-page-slot.is-selected').length === 2`, 'view-area multi-selection')

const selection = await evaluate(`(() => ({
  indices: Array.from(document.querySelectorAll('.continuous-page-slot.is-selected')).map((slot) => Number(slot.dataset.pageIndex)),
  badge: document.querySelector('.continuous-page-slot[data-page-index="2"] .page-transfer-handle strong')?.textContent,
  title: document.querySelector('.continuous-page-slot[data-page-index="2"] .page-transfer-handle')?.title,
}))()`)

const dragFeedback = await evaluate(`(() => {
  const binary = atob('${droppedPageBase64}')
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
  const transfer = new DataTransfer()
  transfer.items.add(new File([bytes], 'first-selection.pdf', { type: '' }))
  transfer.items.add(new File([bytes], 'second-selection.PDF', { type: '' }))
  const target = document.querySelector('.continuous-page-slot[data-page-index="1"]')
  const bounds = target.getBoundingClientRect()
  target.dispatchEvent(new DragEvent('dragover', {
    bubbles: true,
    cancelable: true,
    dataTransfer: transfer,
    clientY: bounds.top + 2,
  }))
  return new Promise((resolve) => requestAnimationFrame(() => resolve({
    edge: target.dataset.dropEdge,
    label: target.querySelector('.page-drop-indicator')?.textContent,
  })))
})()`)

await evaluate(`(() => {
  const binary = atob('${droppedPageBase64}')
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
  const transfer = new DataTransfer()
  transfer.items.add(new File([bytes], 'first-selection.pdf', { type: '' }))
  transfer.items.add(new File([bytes], 'second-selection.PDF', { type: '' }))
  const target = document.querySelector('.continuous-page-slot[data-page-index="1"]')
  const bounds = target.getBoundingClientRect()
  target.dispatchEvent(new DragEvent('drop', {
    bubbles: true,
    cancelable: true,
    dataTransfer: transfer,
    clientY: bounds.top + 2,
  }))
})()`)

await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 10 && !document.querySelector('.busy-overlay')`, 'two dropped PDFs imported')
await waitFor(`document.querySelector('.continuous-page-slot[data-page-index="1"] [data-text-item="true"]')`, 'first inserted page text')

const imported = await evaluate(`(() => ({
  pageCount: document.querySelectorAll('.continuous-page-slot').length,
  currentPage: document.querySelector('input[aria-label="Current page"]')?.value,
  selected: Array.from(document.querySelectorAll('.continuous-page-slot.is-selected')).map((slot) => Number(slot.dataset.pageIndex)),
  insertedText: Array.from(document.querySelectorAll('.continuous-page-slot[data-page-index="1"] [data-text-item="true"]')).map((item) => item.textContent).join(' '),
  dirty: Boolean(document.querySelector('.dirty-dot')),
  toast: document.querySelector('.toast')?.textContent,
  dropIndicatorCleared: !document.querySelector('[data-drop-edge]') && !document.querySelector('.page-drop-indicator'),
}))()`)

await evaluate(`document.querySelector('button[aria-label="Undo (Ctrl+Z)"]').click()`)
await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 8 && !document.querySelector('.busy-overlay')`, 'drop import undo')
await evaluate(`document.querySelector('button[aria-label="More page actions"]')?.click()`)
await waitFor(`Boolean(document.querySelector('button[aria-label="Export selected pages"]') && document.querySelector('button[aria-label="Add pages from files"]'))`, 'page transfer menu actions')

const restored = await evaluate(`(() => ({
  pageCount: document.querySelectorAll('.continuous-page-slot').length,
  currentPage: document.querySelector('input[aria-label="Current page"]')?.value,
  fallbacks: {
    export: Boolean(document.querySelector('button[aria-label="Export selected pages"]')),
    import: Boolean(document.querySelector('button[aria-label="Add pages from files"]')),
  },
}))()`)

socket.close()

if (JSON.stringify(selection.indices) !== JSON.stringify([0, 2]) || selection.badge !== '2' || !selection.title?.includes('2 selected pages')) {
  throw new Error(`View-area multi-selection feedback failed: ${JSON.stringify(selection)}`)
}
if (dragFeedback.edge !== 'before' || !dragFeedback.label?.includes('before page 2')) {
  throw new Error(`View-area drop targeting failed: ${JSON.stringify(dragFeedback)}`)
}
if (imported.pageCount !== 10 || imported.currentPage !== '2' || JSON.stringify(imported.selected) !== JSON.stringify([1, 2])) {
  throw new Error(`Multi-file page import state failed: ${JSON.stringify(imported)}`)
}
if (!imported.insertedText.includes('Dropped page') || !imported.dirty || !imported.toast?.includes('2 pages added') || !imported.dropIndicatorCleared) {
  throw new Error(`Multi-file page import feedback failed: ${JSON.stringify(imported)}`)
}
if (restored.pageCount !== 8 || !restored.fallbacks.export || !restored.fallbacks.import) {
  throw new Error(`Page import undo/accessibility fallback failed: ${JSON.stringify(restored)}`)
}

console.log(JSON.stringify({ selection, dragFeedback, imported, restored }))
