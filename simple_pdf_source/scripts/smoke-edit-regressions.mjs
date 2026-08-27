import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9334)
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

await waitFor(`Boolean(document.querySelector('.continuous-page-slot.is-current .page-canvas'))`, 'page canvas')
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((item) => item.textContent.startsWith('Select only'))`, 'source text')
await evaluate(`document.querySelector('button[aria-label="Edit text (E)"]').click()`)
await waitFor(`Boolean(document.querySelector('.continuous-page-slot.is-current .page-surface.tool-edit'))`, 'edit mode')

const opened = await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent.startsWith('Select only'))
  const needle = 'only these four'
  const start = span.textContent.indexOf(needle)
  const range = document.createRange()
  range.setStart(span.firstChild, start)
  range.setEnd(span.firstChild, start + needle.length)
  const selection = getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  const bounds = range.getBoundingClientRect()
  const spanBounds = span.getBoundingClientRect()
  span.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, clientX: bounds.left + bounds.width / 2, clientY: bounds.top + bounds.height / 2 }))
  return { itemText: span.textContent, start, end: start + needle.length, width: spanBounds.width }
})()`)
await waitFor(`document.querySelector('.inline-pdf-text-editor')?.value === ${JSON.stringify(opened.itemText)}`, 'whole-run text editor')
await waitFor(`document.querySelector('.inline-pdf-text-editor')?.selectionStart === ${opened.start} && document.querySelector('.inline-pdf-text-editor')?.selectionEnd === ${opened.end}`, 'selected substring in whole run')

const firstReplacement = opened.itemText.replace('only these four', 'exactly four')
const finalReplacement = firstReplacement.replace('exactly four', 'precisely four')
const textEdit = await evaluate(`(() => {
  const input = document.querySelector('.inline-pdf-text-editor')
  const frame = input.closest('.inline-text-frame').getBoundingClientRect()
  input.setRangeText('exactly four', input.selectionStart, input.selectionEnd, 'end')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true, cancelable: true }))
  return { value: input.value, width: frame.width }
})()`)
await waitFor(`Array.from(document.querySelectorAll('.text-overlay')).some((item) => item.textContent === ${JSON.stringify(firstReplacement)})`, 'committed whole-run replacement')
await waitFor(`Array.from(document.querySelectorAll('[data-text-item="true"]')).some((item) => item.textContent === ${JSON.stringify(opened.itemText)} && item.dataset.editCovered === 'true')`, 'covered source line suppression')

await evaluate(`Array.from(document.querySelectorAll('.text-overlay')).find((item) => item.textContent === ${JSON.stringify(firstReplacement)}).click()`)
await waitFor(`document.querySelector('.inline-pdf-text-editor')?.value === ${JSON.stringify(firstReplacement)}`, 'select committed text box')
const copyPaste = await evaluate(`(async () => {
  const buttons = () => Array.from(document.querySelectorAll('.edit-inspector button'))
  const source = document.querySelector('.inline-text-frame').getBoundingClientRect()
  buttons().find((button) => button.textContent.trim() === 'Copy').click()
  let paste
  for (let index = 0; index < 60; index += 1) {
    await new Promise(requestAnimationFrame)
    paste = buttons().find((button) => button.textContent.trim() === 'Paste')
    if (paste && !paste.disabled) break
  }
  if (!paste || paste.disabled) throw new Error('Paste did not enable after Copy')
  paste.click()
  for (let index = 0; index < 120; index += 1) {
    await new Promise(requestAnimationFrame)
    const input = document.querySelector('.inline-pdf-text-editor')
    const frame = input?.closest('.inline-text-frame')?.getBoundingClientRect()
    if (input && frame && (Math.abs(frame.left - source.left) > 2 || Math.abs(frame.top - source.top) > 2)) {
      const result = { value: input.value, x: frame.left - source.left, y: frame.top - source.top }
      buttons().find((button) => button.textContent.trim() === 'Delete').click()
      return result
    }
  }
  throw new Error('Pasted text box did not appear at an offset')
})()`)
await waitFor(`!document.querySelector('.inline-pdf-text-editor')`, 'discard pasted text box')

const overlap = await evaluate(`(() => {
  const candidate = document.querySelector('.detected-object')
  const text = Array.from(document.querySelectorAll('[data-text-item="true"]')).find((item) => item.textContent.startsWith('Text remains selectable'))
  if (!candidate || !text) throw new Error('Image overlap fixture was not detected')
  const bounds = text.getBoundingClientRect()
  const hit = document.elementFromPoint(bounds.left + bounds.width / 2, bounds.top + bounds.height / 2)
  const zIndex = getComputedStyle(candidate).zIndex
  candidate.click()
  return { textWinsHitTest: hit === text || text.contains(hit), zIndex }
})()`)
await waitFor(`Boolean(document.querySelector('.object-edit-frame.is-native-selection'))`, 'native image selection')
const imageSelection = await evaluate(`(() => ({
  previewImages: document.querySelectorAll('.object-edit-frame.is-native-selection img').length,
  textStillPresent: Array.from(document.querySelectorAll('[data-text-item="true"]')).some((item) => item.textContent.startsWith('Text remains selectable')),
}))()`)
await evaluate(`Array.from(document.querySelectorAll('.edit-inspector button')).find((button) => button.textContent.trim() === 'Delete').click()`)
await waitFor(`!document.querySelector('.object-edit-frame') && !document.querySelector('.detected-object')`, 'native image deletion')

await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('[data-text-item="true"]')).find((item) => item.textContent.startsWith('Inline style'))
  const bounds = span.getBoundingClientRect()
  span.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, clientX: bounds.left + bounds.width / 2, clientY: bounds.top + bounds.height / 2 }))
})()`)
await waitFor(`document.querySelector('.inline-pdf-text-editor')?.value === 'Inline style-preserving edit sample'`, 'deletion text selection')
await evaluate(`Array.from(document.querySelectorAll('.edit-inspector button')).find((button) => button.textContent.trim() === 'Delete').click()`)
await waitFor(`Array.from(document.querySelectorAll('[data-text-item="true"]')).some((item) => item.textContent === 'Inline style-preserving edit sample' && item.dataset.editCovered === 'true')`, 'deleted line suppression')

await evaluate(`document.querySelector('.save-button').click()`)
await waitFor(`!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'first save')
await evaluate(`Array.from(document.querySelectorAll('.text-overlay')).find((item) => item.textContent === ${JSON.stringify(firstReplacement)}).click()`)
await waitFor(`document.querySelector('.inline-pdf-text-editor')?.value === ${JSON.stringify(firstReplacement)}`, 're-edit after save')
await evaluate(`(() => {
  const input = document.querySelector('.inline-pdf-text-editor')
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set
  setter.call(input, ${JSON.stringify(finalReplacement)})
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true, cancelable: true }))
})()`)
await waitFor(`Array.from(document.querySelectorAll('.text-overlay')).some((item) => item.textContent === ${JSON.stringify(finalReplacement)})`, 'second replacement')
await evaluate(`document.querySelector('.save-button').click()`)
await waitFor(`!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'second save')

if (textEdit.value !== firstReplacement) throw new Error(`Selection replacement produced the wrong complete run: ${JSON.stringify(textEdit)}`)
if (Math.abs(textEdit.width - opened.width) > 2) throw new Error(`Editor frame did not match the highlighted run: ${JSON.stringify({ opened, textEdit })}`)
if (copyPaste.value !== firstReplacement || Math.abs(copyPaste.x) < 2 || Math.abs(copyPaste.y) < 2) throw new Error(`Text box copy/paste failed: ${JSON.stringify(copyPaste)}`)
if (!overlap.textWinsHitTest || Number(overlap.zIndex) >= 2) throw new Error(`Image selection still blocked overlaid text: ${JSON.stringify(overlap)}`)
if (imageSelection.previewImages !== 0 || !imageSelection.textStillPresent) throw new Error(`Native image selection obscured page text: ${JSON.stringify(imageSelection)}`)

socket.close()
console.log(JSON.stringify({ opened, textEdit, copyPaste, overlap, imageSelection, firstReplacement, finalReplacement }))
