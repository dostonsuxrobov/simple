import assert from 'node:assert/strict'
import { writeFile } from 'node:fs/promises'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9495)
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))
const startupDeadline = Date.now() + 90_000
let target
while (!target && Date.now() < startupDeadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch { /* Electron is opening its debugging endpoint. */ }
  if (!target) await pause(40)
}
if (!target) throw new Error('Renderer did not start')

const socket = new WebSocket(target.webSocketDebuggerUrl)
await new Promise((resolve, reject) => {
  socket.addEventListener('open', resolve, { once: true })
  socket.addEventListener('error', reject, { once: true })
})
let requestId = 0
const pending = new Map()
const runtimeErrors = []
socket.addEventListener('message', (event) => {
  const message = JSON.parse(String(event.data))
  if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params.exceptionDetails)
  if (message.method === 'Runtime.consoleAPICalled' && message.params.type === 'error') runtimeErrors.push(message.params.args)
  const request = pending.get(message.id)
  if (!request) return
  pending.delete(message.id)
  clearTimeout(request.timeout)
  if (message.error) request.reject(new Error(message.error.message))
  else if (message.result?.exceptionDetails) request.reject(new Error(message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text))
  else request.resolve(message.result?.result?.value ?? message.result)
})

function call(method, params = {}) {
  return new Promise((resolve, reject) => {
    const id = ++requestId
    const timeout = setTimeout(() => { pending.delete(id); reject(new Error(`Timed out: ${method}`)) }, 20_000)
    pending.set(id, { resolve, reject, timeout })
    socket.send(JSON.stringify({ id, method, params }))
  })
}
const evaluate = (expression) => call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
await call('Runtime.enable')

async function waitFor(expression, label, milliseconds = 15_000) {
  const deadline = Date.now() + milliseconds
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(40)
  }
  const state = await evaluate(`({ active: document.activeElement?.outerHTML.slice(0, 500), toast: document.querySelector('.toast')?.textContent, page: document.querySelector('input[aria-label="Current page"]')?.value })`)
  throw new Error(`Timed out: ${label}; ${JSON.stringify(state)}`)
}

async function key(key, code, modifiers = 0) {
  const virtual = ({ Enter: 13, Escape: 27, ArrowLeft: 37, ArrowUp: 38, ArrowRight: 39, ArrowDown: 40, Delete: 46, Tab: 9 })[key] || key.toUpperCase().charCodeAt(0)
  const parameters = { key, code, modifiers, windowsVirtualKeyCode: virtual, nativeVirtualKeyCode: virtual }
  await call('Input.dispatchKeyEvent', { type: 'keyDown', ...parameters })
  await call('Input.dispatchKeyEvent', { type: 'keyUp', ...parameters })
}

async function clickExpression(expression, xFraction = 0.5, yFraction = 0.5) {
  await waitFor(`Boolean(${expression})`, `click target ${expression}`)
  const point = await evaluate(`(() => {
    const element = ${expression}
    if (!element) throw new Error('Click target is missing')
    element.scrollIntoView({ block: 'nearest', inline: 'nearest' })
    const box = element.getBoundingClientRect()
    return { x: box.left + box.width * ${xFraction}, y: box.top + box.height * ${yFraction} }
  })()`)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...point })
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', ...point, button: 'left', buttons: 1, clickCount: 1 })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', ...point, button: 'left', buttons: 0, clickCount: 1 })
}
const click = (selector) => clickExpression(`document.querySelector(${JSON.stringify(selector)})`)
const inspectorButton = (text) => clickExpression(`[...document.querySelectorAll('.edit-inspector button')].find((button) => button.textContent.trim() === ${JSON.stringify(text)})`)
const field = (label) => `[...document.querySelectorAll('.edit-inspector label')].find((label) => label.querySelector('span')?.textContent.trim() === ${JSON.stringify(label)})?.querySelector('input, select')`

async function fillExpression(expression, text) {
  await clickExpression(expression)
  await key('a', 'KeyA', 2)
  await call('Input.insertText', { text: String(text) })
}

async function selectExpression(expression, value) {
  await evaluate(`(() => {
    const select = ${expression}
    if (!select || ![...select.options].some((option) => option.value === ${JSON.stringify(value)})) throw new Error('Requested select option is missing')
    select.focus()
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(select, ${JSON.stringify(value)})
    select.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
}

const pageText = (prefix) => `[...document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')].find((item) => item.textContent.startsWith(${JSON.stringify(prefix)}))`
const editor = `document.querySelector('.inline-pdf-text-editor')`
const geometry = () => evaluate(`(() => Object.fromEntries([...document.querySelectorAll('.edit-inspector .geometry-grid label')].map((label) => [label.querySelector('span').textContent, Number(label.querySelector('input').value)])))()`)
const closeEnough = (value, expected, message, tolerance = 0.2) => assert.ok(Math.abs(value - expected) < tolerance, `${message}: ${value} versus ${expected}`)

try {
  await waitFor(`Boolean(${pageText('Select only')})`, 'initial PDF text', 45_000)
  const toolbar = await evaluate(`({ editLabel: document.querySelector('button[aria-label="Edit text (E)"]')?.textContent, zoom: document.querySelector('input[aria-label="Zoom percentage"]')?.value })`)
  assert.ok(toolbar.editLabel.includes('Edit PDF'), 'Edit PDF must have a visible toolbar label')

  await fillExpression(`document.querySelector('input[aria-label="Zoom percentage"]')`, '125')
  await key('Enter', 'Enter')
  await waitFor(`document.querySelector('select[aria-label="Page fit"]')?.value === 'custom' && document.querySelector('input[aria-label="Zoom percentage"]')?.value === '125'`, 'custom zoom')
  await selectExpression(`document.querySelector('select[aria-label="Page fit"]')`, 'width')
  await waitFor(`document.querySelector('select[aria-label="Page fit"]')?.value === 'width' && document.querySelector('input[aria-label="Zoom percentage"]')?.value !== '125'`, 'fit width')
  const widthZoom = await evaluate(`Number(document.querySelector('input[aria-label="Zoom percentage"]').value)`)
  await selectExpression(`document.querySelector('select[aria-label="Page fit"]')`, 'fit')
  await waitFor(`Number(document.querySelector('input[aria-label="Zoom percentage"]')?.value) < ${widthZoom}`, 'fit page')
  const fitZoom = await evaluate(`Number(document.querySelector('input[aria-label="Zoom percentage"]').value)`)

  await evaluate(`document.activeElement?.blur()`)
  await key('1', 'Digit1', 2)
  await waitFor(`document.querySelector('select[aria-label="Page fit"]')?.value === 'actual' && document.querySelector('input[aria-label="Zoom percentage"]')?.value === '100'`, 'actual-size shortcut')
  await key('2', 'Digit2', 2)
  await waitFor(`document.querySelector('select[aria-label="Page fit"]')?.value === 'width'`, 'fit-width shortcut')
  await key('0', 'Digit0', 2)
  await waitFor(`document.querySelector('select[aria-label="Page fit"]')?.value === 'fit'`, 'fit-page shortcut')

  await call('Emulation.setDeviceMetricsOverride', { width: 1100, height: 900, deviceScaleFactor: 1, mobile: false })
  await pause(150)
  const narrow = await evaluate(`(() => {
    const selectors = ['button[aria-label="Edit text (E)"]', 'input[aria-label="Current page"]', 'input[aria-label="Zoom percentage"]', 'select[aria-label="Page fit"]']
    return selectors.map((selector) => { const element = document.querySelector(selector); const bounds = element.getBoundingClientRect(); return { selector, visible: bounds.width > 0 && bounds.height > 0 && bounds.left >= 0 && bounds.right <= innerWidth && bounds.bottom <= innerHeight } })
  })()`)
  assert.ok(narrow.every((item) => item.visible), `Essential toolbar controls disappeared at 1100px: ${JSON.stringify(narrow)}`)
  await call('Emulation.clearDeviceMetricsOverride')
  await pause(150)

  await click('button[aria-label="Edit text (E)"]')
  await waitFor(`Boolean(document.querySelector('.page-surface.tool-edit')) && Boolean(${pageText('Inline style')})`, 'edit mode text')
  await clickExpression(pageText('Inline style'))
  await waitFor(`${editor}?.value === 'Inline style-preserving edit sample'`, 'direct text editor')
  const sourceFont = await evaluate(`(() => { const select = ${field('Font')}; return { value: select.value, selected: select.selectedOptions[0]?.textContent } })()`)
  assert.ok(sourceFont.value && sourceFont.selected && !/\bg_|pdfjs/i.test(sourceFont.selected), 'The actual source font must have a readable, selected option')

  const originalGeometry = await geometry()
  await evaluate(`${editor}.focus(); ${editor}.setSelectionRange(3, 3)`)
  await key('ArrowRight', 'ArrowRight')
  const caret = await evaluate(`${editor}.selectionStart`)
  assert.equal(caret, 4, 'Arrow key must move the text caret while typing')
  await key('d', 'KeyD', 2)
  await pause(100)
  assert.deepEqual(await geometry(), originalGeometry, 'Typing shortcuts must not move or duplicate the text box')
  assert.equal(await evaluate(`document.querySelector('input[aria-label="Current page"]').value`), '1', 'Typing must not change pages')

  const sizedText = 'Font size retained'
  const fontSize = 24
  await fillExpression(editor, sizedText)
  await fillExpression(field('Size'), fontSize)
  await key('Tab', 'Tab')
  await waitFor(`Number(${field('Size')}?.value) === ${fontSize}`, 'font size selection')
  await waitFor(`(() => { const input = ${editor}; const page = document.querySelector('.continuous-page-slot.is-current .page-surface'); return input && Math.abs(parseFloat(getComputedStyle(input).fontSize) / (page.getBoundingClientRect().width / 500) - ${fontSize}) < 0.2 })()`, 'font size rendered')
  const sized = await evaluate(`(() => { const input = ${editor}; const style = getComputedStyle(input); const page = document.querySelector('.continuous-page-slot.is-current .page-surface'); return { size: parseFloat(style.fontSize), zoom: page.getBoundingClientRect().width / 500 } })()`)
  closeEnough(sized.size / sized.zoom, fontSize, 'Editor font size did not update')
  await inspectorButton('Done')
  await waitFor(`!${editor}`, 'font-size edit committed')

  await clickExpression(pageText('Select only'))
  await waitFor(`${editor}?.value?.startsWith('Select only')`, 'wrap source edit')
  await waitFor(`document.querySelector('.continuous-page-slot.is-current .page-canvas')?.dataset.textRemovals?.includes('Select only')`, 'text-only preview')
  const sourceRemovals = await evaluate(`document.querySelector('.continuous-page-slot.is-current .page-canvas').dataset.textRemovals`)
  await selectExpression(field('Text fitting'), 'wrap')
  await fillExpression(field('W'), 220)
  await key('Tab', 'Tab')
  const wrappedText = 'Quicksilver amber lanterns illuminate curious explorers carrying velvet notebooks across sapphire valleys. Every carefully chosen sentence remains readable while the paragraph expands inside its own text box.'
  await fillExpression(editor, wrappedText)
  await fillExpression(field('Line spacing'), 22)
  await key('Tab', 'Tab')
  const wrapped = await evaluate(`(() => {
    const input = ${editor}
    const style = getComputedStyle(input)
    const frame = input.closest('.inline-text-frame')
    const box = frame.getBoundingClientRect()
    const page = document.querySelector('.continuous-page-slot.is-current .page-surface').getBoundingClientRect()
    const covers = document.querySelectorAll('.text-original-cover').length
    return { text: input.value, height: box.height, width: box.width, fontSize: parseFloat(style.fontSize), lineHeight: parseFloat(style.lineHeight), wrap: input.wrap, whiteSpace: style.whiteSpace, scrollHeight: input.scrollHeight, clientHeight: input.clientHeight,
      frameBackground: getComputedStyle(frame).backgroundColor, editorBackground: style.backgroundColor,
      covers }
  })()`)
  assert.equal(wrapped.text, wrappedText, 'Real text input did not retain the full paragraph')
  assert.ok(wrapped.height > wrapped.lineHeight * 3, 'Wrapped editor did not grow for multiline content')
  assert.ok(wrapped.scrollHeight <= wrapped.clientHeight + 2, 'Wrapped editor clips part of the paragraph')
  assert.equal(wrapped.frameBackground, 'rgba(0, 0, 0, 0)', 'Expanded native editor frame must not paint over surrounding content')
  assert.equal(wrapped.editorBackground, 'rgba(0, 0, 0, 0)', 'Expanded native textarea must not paint over surrounding content')
  assert.equal(wrapped.covers, 0, 'Editing must never paint a source cover')
  assert.equal(await evaluate(`document.querySelector('.continuous-page-slot.is-current .page-canvas').dataset.textRemovals`), sourceRemovals, 'Wrapping must keep the original source selection')
  if (process.env.SIMPLE_EDIT_QA_SCREENSHOT) {
    const capture = await call('Page.captureScreenshot', { format: 'png' })
    await writeFile(process.env.SIMPLE_EDIT_QA_SCREENSHOT, Buffer.from(capture.data, 'base64'))
  }
  await inspectorButton('Done')
  await waitFor(`!${editor}`, 'wrap edit committed')
  const committedPreview = await evaluate(`(() => {
    const overlay = [...document.querySelectorAll('.continuous-page-slot.is-current .text-overlay')].find((item) => item.textContent.startsWith('Quicksilver'))
    const page = document.querySelector('.continuous-page-slot.is-current .page-surface').getBoundingClientRect()
    const covers = document.querySelectorAll('.text-original-cover').length
    return { background: getComputedStyle(overlay).backgroundColor, covers }
  })()`)
  assert.equal(committedPreview.background, 'rgba(0, 0, 0, 0)', 'Committed native replacement must not hide surrounding page content')
  assert.equal(committedPreview.covers, 0, 'Done must not add a background patch')

  await clickExpression(`[...document.querySelectorAll('.continuous-page-slot.is-current .text-overlay')].find((item) => item.textContent.startsWith('Quicksilver'))`)
  await waitFor(`${editor}?.value === ${JSON.stringify(wrappedText)}`, 'wrapped paragraph re-edit')
  await fillExpression(editor, `${wrappedText} `.repeat(30))
  await waitFor(`Boolean(document.querySelector('.text-fit-feedback.has-overflow'))`, 'page-edge overflow warning')
  await inspectorButton('Cancel')
  await waitFor(`!${editor}`, 'cancel overflow draft')

  // The center of this fixture image deliberately contains selectable text.
  // Hit an exposed corner to exercise actual image selection through the UI.
  await clickExpression(`document.querySelector('.continuous-page-slot.is-current .detected-object')`, 0.12, 0.12)
  await waitFor(`Boolean(document.querySelector('.object-edit-frame'))`, 'native image selection')
  const before = await geometry()
  // Restore non-field focus without disturbing the selected image.
  await evaluate(`document.activeElement?.blur()`)
  await key('ArrowRight', 'ArrowRight')
  await key('ArrowDown', 'ArrowDown', 8)
  const after = await geometry()
  closeEnough(after.X - before.X, 1, 'Arrow-right image nudge')
  closeEnough(before.Y - after.Y, 10, 'Shift-arrow-down image nudge')
  await key('d', 'KeyD', 2)
  await waitFor(`(() => { const labels = [...document.querySelectorAll('.edit-inspector .geometry-grid label')]; const x = labels.find((label) => label.querySelector('span').textContent === 'X'); return Math.abs(Number(x?.querySelector('input').value) - ${after.X}) > 2 })()`, 'duplicate object shortcut')
  const duplicate = await geometry()
  assert.ok(duplicate.X > after.X && duplicate.Y < after.Y, 'Duplicate should be offset down and right')
  await inspectorButton('Delete')
  await waitFor(`!document.querySelector('.object-edit-frame')`, 'remove duplicate')

  await click('.save-button')
  await waitFor(`!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'saved document', 45_000)
  console.log(JSON.stringify({ toolbar, zoom: { custom: 125, width: widthZoom, fit: fitZoom }, narrow, sourceFont, caret, sizedText, fontSize, sized, wrappedText, sourceRemovals, wrapped, committedPreview, nudge: { before, after }, duplicate }))
} catch (error) {
  const state = await evaluate(`({ editor: document.querySelector('.inline-pdf-text-editor')?.outerHTML.slice(0, 2000), fields: [...document.querySelectorAll('.edit-inspector label')].map((label) => ({ name: label.querySelector('span')?.textContent, value: label.querySelector('input,select')?.value })) })`).catch(() => null)
  if (process.env.SIMPLE_EDIT_QA_SCREENSHOT) {
    const capture = await call('Page.captureScreenshot', { format: 'png' }).catch(() => null)
    if (capture?.data) await writeFile(`${process.env.SIMPLE_EDIT_QA_SCREENSHOT}.failed.png`, Buffer.from(capture.data, 'base64'))
  }
  throw new Error(`${error.message}\nUI state: ${JSON.stringify(state)}\nRuntime errors: ${JSON.stringify(runtimeErrors)}`, { cause: error })
} finally {
  for (const request of pending.values()) clearTimeout(request.timeout)
  socket.close()
}
