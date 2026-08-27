import process from 'node:process'
import fs from 'node:fs/promises'
import path from 'node:path'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9334)
const deadline = Date.now() + 45_000
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

const currentCanvasReady = `(() => {
  const canvas = document.querySelector('.continuous-page-slot.is-current .page-canvas')
  return Boolean(canvas && canvas.width > 1)
})()`
await waitFor(currentCanvasReady, 'current page canvas')

const initial = await evaluate(`(() => ({
  page: document.querySelector('input[aria-label="Current page"]')?.value,
  pageSlots: document.querySelectorAll('.continuous-page-slot').length,
  mountedCanvases: document.querySelectorAll('.continuous-page-slot .page-canvas').length,
  selectedActions: document.querySelectorAll('.thumbnail-item.is-selected .thumbnail-actions button').length,
  hasExportButton: Boolean(document.querySelector('button[aria-label*="Export"]')),
  hasRemoteRotateButton: Boolean(document.querySelector('.toolbar button[aria-label^="Rotate"]')),
  hasRemoteDeleteButton: Boolean(document.querySelector('.toolbar button[aria-label^="Delete"]')),
  scrollable: document.querySelector('.viewer').scrollHeight > document.querySelector('.viewer').clientHeight,
}))()`)

const rotation = await evaluate(`(async () => {
  const canvas = document.querySelector('.continuous-page-slot.is-current .page-canvas')
  const original = [canvas.width, canvas.height]
  const started = performance.now()
  document.querySelector('button[aria-label="Rotate page 1 right"]').click()
  while (performance.now() - started < 3000) {
    const next = document.querySelector('.continuous-page-slot.is-current .page-canvas')
    if (next && next.width === original[1] && next.height === original[0]) {
      const elapsed = performance.now() - started
      document.querySelector('button[aria-label="Rotate page 1 left"]').click()
      return { original, rotated: [next.width, next.height], elapsed }
    }
    await new Promise(requestAnimationFrame)
  }
  throw new Error('Rotation did not repaint in time')
})()`)

await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((item) => item.textContent.startsWith('Select only'))`, 'selectable text after rotation')

await evaluate(`document.querySelector('button[aria-label="Select text (V)"]').click()`)
const selectedText = await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent.startsWith('Select only'))
  const text = span.firstChild
  const range = document.createRange()
  range.setStart(text, 7)
  range.setEnd(text, 11)
  const selection = getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  return selection.toString()
})()`)

await evaluate(`document.querySelector('button[aria-label="Highlight text"]').click()`)
await waitFor(`Boolean(document.querySelector('.continuous-page-slot.is-current .page-surface.tool-highlight'))`, 'highlight tool activation')
await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent.startsWith('Select only'))
  const text = span.firstChild
  const range = document.createRange()
  range.setStart(text, 7)
  range.setEnd(text, 11)
  const selection = getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  span.dispatchEvent(new MouseEvent('mouseup', { bubbles: true }))
})()`)
await waitFor(`Boolean(document.querySelector('.continuous-page-slot.is-current .highlight-overlay'))`, 'focused highlight')
const highlight = await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent.startsWith('Select only'))
  const overlay = document.querySelector('.continuous-page-slot.is-current .highlight-overlay')
  return { spanWidth: span.getBoundingClientRect().width, overlayWidth: overlay.getBoundingClientRect().width }
})()`)

await evaluate(`Array.from(document.querySelectorAll('.sidebar-tabs button')).find((button) => button.textContent === 'Bookmarks').click()`)
await waitFor(`document.querySelectorAll('.bookmark-row').length === 4`, 'native bookmarks')
const bookmarks = await evaluate(`Array.from(document.querySelectorAll('.bookmark-row strong')).map((item) => item.textContent)`)
await evaluate(`Array.from(document.querySelectorAll('.sidebar-tabs button')).find((button) => button.textContent === 'Pages').click()`)

await evaluate(`document.querySelector('.continuous-page-slot[data-page-index="3"]').scrollIntoView({ block: 'center' })`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '4'`, 'continuous scroll page tracking')
await waitFor(currentCanvasReady, 'landscape page canvas')
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((item) => item.textContent.startsWith('Inline style'))`, 'landscape selectable text')

await evaluate(`document.querySelector('button[aria-label="Edit text (E)"]').click()`)
// Automatic fit can repaint the current page when the inspector opens. Wait
// for that replacement text layer before targeting its inline-edit span.
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((item) => item.textContent.startsWith('Inline style'))`, 'landscape text after edit mode')
await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent.startsWith('Inline style'))
  const bounds = span.getBoundingClientRect()
  span.dispatchEvent(new MouseEvent('click', { bubbles: true, clientX: bounds.left + bounds.width * .35, clientY: bounds.top + bounds.height / 2 }))
})()`)
await waitFor(`Boolean(document.querySelector('.continuous-page-slot.is-current .inline-pdf-text-editor'))`, 'inline text editor')
const editStyle = await evaluate(`(() => {
  const editor = document.querySelector('.continuous-page-slot.is-current .inline-pdf-text-editor')
  const source = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).find((item) => item.textContent.startsWith('Inline style'))
  const frame = editor.closest('.inline-text-frame')
  const before = getComputedStyle(editor)
  const initialValue = editor.value
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set
  setter.call(editor, 'Direct style-preserving edit sample')
  editor.dispatchEvent(new Event('input', { bubbles: true }))
  editor.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true }))
  return {
    initialValue,
    frameWidth: frame.getBoundingClientRect().width,
    sourceWidth: source.getBoundingClientRect().width,
    family: before.fontFamily,
    size: before.fontSize,
    style: before.fontStyle,
    weight: before.fontWeight,
    color: before.color,
    background: before.backgroundColor,
  }
})()`)
await waitFor(`Boolean(document.querySelector('.continuous-page-slot.is-current .text-overlay'))`, 'committed inline edit')

await evaluate(`document.querySelector('.continuous-page-slot.is-current .text-overlay').click()`)
await waitFor(`Boolean(document.querySelector('.edit-inspector button'))`, 'selected committed text')
const copyPaste = await evaluate(`(async () => {
  const inspectorButtons = () => Array.from(document.querySelectorAll('.edit-inspector button'))
  inspectorButtons().find((button) => button.textContent.trim() === 'Copy').click()
  const paste = inspectorButtons().find((button) => button.textContent.trim() === 'Paste')
  if (!paste || paste.disabled) throw new Error('Paste was not enabled after copying a text box')
  const sourceFrame = document.querySelector('.inline-text-frame').getBoundingClientRect()
  paste.click()
  for (let index = 0; index < 120; index += 1) {
    await new Promise(requestAnimationFrame)
    const editor = document.querySelector('.inline-pdf-text-editor')
    const frame = editor?.closest('.inline-text-frame')?.getBoundingClientRect()
    if (editor && frame && (Math.abs(frame.left - sourceFrame.left) > 2 || Math.abs(frame.top - sourceFrame.top) > 2)) {
      const result = {
        value: editor.value,
        offset: [frame.left - sourceFrame.left, frame.top - sourceFrame.top],
      }
      inspectorButtons().find((button) => button.textContent.trim() === 'Delete').click()
      return result
    }
  }
  throw new Error('Pasted text box did not appear at an offset')
})()`)
await waitFor(`!document.querySelector('.continuous-page-slot.is-current .inline-pdf-text-editor')`, 'discard pasted copy')

await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Current page"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, '5')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }))
})()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '5'`, 'jump to page five')
await waitFor(currentCanvasReady, 'page five canvas')
await evaluate(`document.querySelector('button[aria-label="Bookmark this page"]').click()`)
await evaluate(`document.querySelector('button[aria-label="Rotate page 5 right"]').click()`)
await evaluate(`document.querySelector('.save-button').click()`)
await waitFor(`!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'saved document')
await pause(250)

const final = await evaluate(`(() => ({
  page: document.querySelector('input[aria-label="Current page"]')?.value,
  mountedCanvases: document.querySelectorAll('.continuous-page-slot .page-canvas').length,
  currentCanvas: (() => { const canvas = document.querySelector('.continuous-page-slot.is-current .page-canvas'); return [canvas?.width || 0, canvas?.height || 0] })(),
  bookmarked: Boolean(document.querySelector('button[aria-label="Remove bookmark"]')),
}))()`)

await evaluate(`(() => {
  const binary = atob('${droppedPageBase64}')
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
  const transfer = new DataTransfer()
  transfer.items.add(new File([bytes], 'dropped-page.pdf', { type: 'application/pdf' }))
  const target = document.querySelector('.thumbnail-item.is-selected')
  const bounds = target.getBoundingClientRect()
  target.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, dataTransfer: transfer, clientY: bounds.bottom - 2 }))
})()`)
await waitFor(`document.querySelector('.page-field span')?.textContent?.includes('9')`, 'dropped PDF insertion')
const dropImport = await evaluate(`(() => ({
  pages: document.querySelector('.page-field span')?.textContent,
  page: document.querySelector('input[aria-label="Current page"]')?.value,
  selectedActions: document.querySelectorAll('.thumbnail-item.is-selected .thumbnail-actions button').length,
}))()`)

socket.close()
if (editStyle.initialValue !== 'Inline style-preserving edit sample') {
  throw new Error(`Edit mode opened a word instead of the complete highlighted text run: ${JSON.stringify(editStyle)}`)
}
if (Math.abs(editStyle.frameWidth - editStyle.sourceWidth) > 2) {
  throw new Error(`Text editor did not cover the complete highlighted run: ${JSON.stringify(editStyle)}`)
}
if (copyPaste.value !== 'Direct style-preserving edit sample') {
  throw new Error(`Copy/paste did not preserve the complete text box: ${JSON.stringify(copyPaste)}`)
}
console.log(JSON.stringify({ initial, rotation, selectedText, highlight, bookmarks, editStyle, copyPaste, final, dropImport }))
