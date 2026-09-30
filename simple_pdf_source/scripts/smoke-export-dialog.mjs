const port = Number(process.env.SIMPLE_EXPORT_TEST_PORT || 9398)
const deadline = Date.now() + 35_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {}
  if (!target) await pause(40)
}
if (!target) throw new Error('PDF renderer did not start.')

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
  socket.send(JSON.stringify({ id, method: 'Runtime.evaluate', params: { expression, returnByValue: true, awaitPromise: true, userGesture: true } }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

async function waitFor(expression, label = expression) {
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(40)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`Boolean(document.querySelector('.page-canvas')?.width > 300 && Array.from(document.querySelectorAll('button')).some((button) => button.textContent.trim() === 'Export As'))`, 'PDF and Export As toolbar')
await evaluate(`Array.from(document.querySelectorAll('button')).find((button) => button.textContent.trim() === 'Export As').click()`)
await waitFor(`Boolean(document.querySelector('[role="dialog"][aria-label="Export As"]'))`, 'Export As dialog')

const initial = await evaluate(`(() => {
  const dialog = document.querySelector('[role="dialog"][aria-label="Export As"]')
  const formats = Array.from(dialog.querySelectorAll('[role="radio"]')).map((button) => ({ label: button.querySelector('strong')?.textContent, selected: button.getAttribute('aria-checked') === 'true' }))
  return {
    formats,
    pageSummary: dialog.querySelector('.export-dialog-footer > span')?.textContent,
    resolutionDisabled: dialog.querySelector('select')?.disabled,
  }
})()`)

const expectedFormats = ['PDF', 'PNG images', 'JPEG images', 'WebP images', 'Word document', 'Plain text', 'Markdown', 'Web page']
if (JSON.stringify(initial.formats.map((item) => item.label)) !== JSON.stringify(expectedFormats)) throw new Error(`Unexpected formats: ${JSON.stringify(initial)}`)
if (!initial.formats.find((item) => item.label === 'PNG images')?.selected || initial.pageSummary?.trim() !== '8 pages selected' || initial.resolutionDisabled) {
  throw new Error(`Unexpected default export state: ${JSON.stringify(initial)}`)
}

const codecs = await evaluate(`(async () => {
  const canvas = document.querySelector('.page-canvas')
  const encode = (type, quality) => new Promise((resolve, reject) => canvas.toBlob(async (blob) => {
    if (!blob) { reject(new Error(type + ' encoding failed')); return }
    const bytes = new Uint8Array(await blob.arrayBuffer())
    resolve({ type: blob.type, size: blob.size, head: Array.from(bytes.slice(0, 12)) })
  }, type, quality))
  return {
    png: await encode('image/png'),
    jpeg: await encode('image/jpeg', .9),
    webp: await encode('image/webp', .9),
  }
})()`)
if (codecs.png.size < 1_000 || codecs.png.head.slice(0, 4).join(',') !== '137,80,78,71') throw new Error(`PNG encoding failed: ${JSON.stringify(codecs.png)}`)
if (codecs.jpeg.size < 1_000 || codecs.jpeg.head.slice(0, 2).join(',') !== '255,216') throw new Error(`JPEG encoding failed: ${JSON.stringify(codecs.jpeg)}`)
if (codecs.webp.size < 1_000 || String.fromCharCode(...codecs.webp.head.slice(0, 4)) !== 'RIFF' || String.fromCharCode(...codecs.webp.head.slice(8, 12)) !== 'WEBP') throw new Error(`WebP encoding failed: ${JSON.stringify(codecs.webp)}`)

await evaluate(`Array.from(document.querySelectorAll('[role="radio"]')).find((button) => button.querySelector('strong')?.textContent === 'Word document').click()`)
const wordState = await evaluate(`({ note: document.querySelector('.export-layout-note')?.textContent, imageControlsDisabled: Array.from(document.querySelectorAll('.export-image-options select')).every((input) => input.disabled) })`)
if (!wordState.note?.includes('Scans need OCR') || !wordState.imageControlsDisabled) throw new Error(`Word limitations are not explicit: ${JSON.stringify(wordState)}`)

await evaluate(`(() => { const input = document.querySelector('input[aria-label="Custom export page range"]'); input.focus(); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '1-3, 5'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
await waitFor(`document.querySelector('.export-dialog-footer > span')?.textContent.trim() === '4 pages selected'`, 'custom export range')

await evaluate(`Array.from(document.querySelectorAll('[role="radio"]')).find((button) => button.querySelector('strong')?.textContent === 'WebP images').click()`)
const webpState = await evaluate(`Array.from(document.querySelectorAll('.export-image-options select')).map((input) => input.disabled)`)
if (webpState.some(Boolean)) throw new Error(`WebP quality controls should be enabled: ${JSON.stringify(webpState)}`)

await evaluate(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))`)
await waitFor(`!document.querySelector('[role="dialog"][aria-label="Export As"]')`, 'Escape closes export')
const shortcutState = await evaluate(`(() => {
  const opened = !window.dispatchEvent(new KeyboardEvent('keydown', { key: 'E', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }))
  return { opened }
})()`)
await waitFor(`Boolean(document.querySelector('[role="dialog"][aria-label="Export As"]'))`, 'Ctrl+Shift+E opens export')
const modalIsolation = await evaluate(`(() => {
  const prevented = !window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true }))
  return { prevented, exportDialogs: document.querySelectorAll('[aria-label="Export As"]').length, printDialogs: document.querySelectorAll('[aria-label="Print"]').length }
})()`)
socket.close()

if (!shortcutState.opened || !modalIsolation.prevented || modalIsolation.exportDialogs !== 1 || modalIsolation.printDialogs !== 0) {
  throw new Error(`Export shortcut/modal isolation failed: ${JSON.stringify({ shortcutState, modalIsolation })}`)
}

console.log(JSON.stringify({ initial, codecs, wordState, webpState, shortcutState, modalIsolation }))
