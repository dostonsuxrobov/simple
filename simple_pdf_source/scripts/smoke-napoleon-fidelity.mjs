import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9370)
const deadline = Date.now() + 120_000
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
    await pause(40)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`document.querySelector('.page-field span')?.textContent?.includes('1277')`, 'Napoleon document')
await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Current page"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, '30')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
})()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '30'`, 'page 30 navigation')
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((item) => item.textContent?.startsWith('Napoleon Bonaparte'))`, 'page 30 text')

const render = await evaluate(`(() => {
  const canvas = document.querySelector('.continuous-page-slot.is-current .page-canvas')
  return {
    backing: [canvas.width, canvas.height],
    css: [Number.parseFloat(canvas.style.width), Number.parseFloat(canvas.style.height)],
    outputScale: [canvas.width / Number.parseFloat(canvas.style.width), canvas.height / Number.parseFloat(canvas.style.height)],
  }
})()`)

await evaluate(`document.querySelector('button[aria-label="Edit text (E)"]').click()`)
const selection = await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent?.startsWith('Napoleon Bonaparte'))
  const source = span.textContent
  const start = source.indexOf('founder')
  const end = start + 'founder of modern France'.length
  const range = document.createRange()
  range.setStart(span.firstChild, start)
  range.setEnd(span.firstChild, end)
  const bounds = range.getBoundingClientRect()
  const selection = window.getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  span.dispatchEvent(new MouseEvent('click', {
    bubbles: true,
    cancelable: true,
    clientX: bounds.left + bounds.width / 2,
    clientY: bounds.top + bounds.height / 2,
  }))
  return { text: selection.toString(), lineWidth: span.getBoundingClientRect().width, selectionWidth: bounds.width }
})()`)

await waitFor(`document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')?.value === 'founder of modern France'`, 'selected-substring editor')
const editor = await evaluate(`(() => {
  const input = document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')
  const frame = input.closest('.inline-text-frame')
  const style = getComputedStyle(input)
  return {
    value: input.value,
    frameWidth: frame.getBoundingClientRect().width,
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    lineHeight: style.lineHeight,
    color: style.color,
  }
})()`)

await evaluate(`(() => {
  const input = document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set
  setter.call(input, 'architect of modern France')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }))
})()`)
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current .text-overlay')).some((item) => item.textContent === 'architect of modern France')`, 'committed replacement')

const committed = await evaluate(`(() => {
  const overlay = Array.from(document.querySelectorAll('.continuous-page-slot.is-current .text-overlay')).find((item) => item.textContent === 'architect of modern France')
  const cover = overlay.parentElement.querySelector('.text-original-cover')
  return {
    overlayWidth: overlay.getBoundingClientRect().width,
    coverWidth: cover.getBoundingClientRect().width,
    fontFamily: getComputedStyle(overlay).fontFamily,
    fontSize: getComputedStyle(overlay).fontSize,
    color: getComputedStyle(overlay).color,
  }
})()`)

await evaluate(`document.querySelector('.save-button').click()`)
await waitFor(`!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'saved Napoleon edit')
await pause(400)

socket.close()
console.log(JSON.stringify({ render, selection, editor, committed }))
