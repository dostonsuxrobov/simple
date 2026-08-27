import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9390)
const pageNumber = Number(process.env.SIMPLE_FIDELITY_PAGE || 1)
const original = process.env.SIMPLE_FIDELITY_ORIGINAL || ''
const replacement = process.env.SIMPLE_FIDELITY_REPLACEMENT || ''
const deadline = Date.now() + 120_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

if (!original || !replacement) throw new Error('Set SIMPLE_FIDELITY_ORIGINAL and SIMPLE_FIDELITY_REPLACEMENT.')

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Electron is still opening its debugging endpoint.
  }
  if (!target) await pause(40)
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
    await pause(50)
  }
  throw new Error(`Timed out: ${label}`)
}

const originalJson = JSON.stringify(original)
const replacementJson = JSON.stringify(replacement)
await waitFor(`Boolean(document.querySelector('.page-canvas'))`, 'document canvas')

if (pageNumber !== 1) {
  await evaluate(`(() => {
    const input = document.querySelector('input[aria-label="Current page"]')
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    setter.call(input, ${JSON.stringify(String(pageNumber))})
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
  })()`)
  await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === ${JSON.stringify(String(pageNumber))}`, `page ${pageNumber}`)
}

await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((item) => item.textContent?.includes(${originalJson}))`, 'source text')
await evaluate(`document.querySelector('button[aria-label*="Edit text"]')?.click()`)

const source = await evaluate(`(() => {
  const needle = ${originalJson}
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent?.includes(needle))
  const start = span.textContent.indexOf(needle)
  const range = document.createRange()
  range.setStart(span.firstChild, start)
  range.setEnd(span.firstChild, start + needle.length)
  const bounds = range.getBoundingClientRect()
  const selection = window.getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  const style = getComputedStyle(span)
  const details = {
    sourceItem: span.textContent,
    selectedText: selection.toString(),
    selectedRect: { x: bounds.x, y: bounds.y, width: bounds.width, height: bounds.height },
    dataset: { ...span.dataset },
    computed: {
      fontFamily: style.fontFamily,
      fontSize: style.fontSize,
      fontWeight: style.fontWeight,
      fontStyle: style.fontStyle,
      color: style.color,
      transform: style.transform,
    },
  }
  span.dispatchEvent(new MouseEvent('click', {
    bubbles: true,
    cancelable: true,
    clientX: bounds.left + bounds.width / 2,
    clientY: bounds.top + bounds.height / 2,
  }))
  return details
})()`)

await waitFor(`document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')?.value === ${originalJson}`, 'direct editor')
const editor = await evaluate(`(() => {
  const input = document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')
  const frame = input.closest('.inline-text-frame').getBoundingClientRect()
  const style = getComputedStyle(input)
  return {
    value: input.value,
    rect: { x: frame.x, y: frame.y, width: frame.width, height: frame.height },
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    lineHeight: style.lineHeight,
    color: style.color,
    transform: style.transform,
  }
})()`)

await evaluate(`(() => {
  const input = document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set
  setter.call(input, ${replacementJson})
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }))
})()`)
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current .text-overlay')).some((item) => item.textContent === ${replacementJson})`, 'committed replacement')

const committed = await evaluate(`(() => {
  const overlay = Array.from(document.querySelectorAll('.continuous-page-slot.is-current .text-overlay')).find((item) => item.textContent === ${replacementJson})
  const cover = overlay.parentElement.querySelector('.text-original-cover')
  const box = overlay.getBoundingClientRect()
  const coverBox = cover?.getBoundingClientRect()
  const style = getComputedStyle(overlay)
  return {
    rect: { x: box.x, y: box.y, width: box.width, height: box.height },
    coverRect: coverBox && { x: coverBox.x, y: coverBox.y, width: coverBox.width, height: coverBox.height },
    fontFamily: style.fontFamily,
    fontSize: style.fontSize,
    fontWeight: style.fontWeight,
    fontStyle: style.fontStyle,
    color: style.color,
    transform: style.transform,
  }
})()`)

await evaluate(`document.querySelector('.save-button')?.click()`)
await waitFor(`!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'saved edit')
await waitFor(`!document.querySelector('.continuous-page-slot.is-current .text-overlay')`, 'flattened overlay reload')
const sourceBaseline = Number(source.dataset.pdfBaselineY)
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).filter((item) => Math.abs(Number(item.dataset.pdfBaselineY) - ${JSON.stringify(sourceBaseline)}) < 0.02).length > 1`, 'saved word runs')

const saved = await evaluate(`(() => {
  return Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]'))
    .filter((item) => Math.abs(Number(item.dataset.pdfBaselineY) - ${JSON.stringify(sourceBaseline)}) < 0.02)
    .map((span) => {
      const style = getComputedStyle(span)
      return {
        text: span.textContent,
        dataset: { ...span.dataset },
        computed: { fontFamily: style.fontFamily, fontSize: style.fontSize, fontWeight: style.fontWeight, fontStyle: style.fontStyle, color: style.color },
      }
    })
})()`)

socket.close()
console.log(JSON.stringify({ pageNumber, original, replacement, source, editor, committed, saved }))
