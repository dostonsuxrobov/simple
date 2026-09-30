const port = Number(process.env.SIMPLE_PRINT_TEST_PORT || 9399)
const deadline = Date.now() + 40_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {}
  if (!target) await pause(50)
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
    await pause(50)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`Boolean(document.querySelector('.page-canvas')?.width > 300 && document.querySelector('button[aria-label="Print (Ctrl+P)"]'))`, 'loaded PDF and Print toolbar action')
const shortcutPrevented = await evaluate(`(() => { document.querySelector('button[aria-label="Print (Ctrl+P)"]').focus(); return !window.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', ctrlKey: true, bubbles: true, cancelable: true })) })()`)
await waitFor(`Boolean(document.querySelector('[role="dialog"][aria-label="Print"]') && document.querySelector('.print-page-image'))`, 'live Print dialog')

const initial = await evaluate(`(() => {
  const dialog = document.querySelector('[role="dialog"][aria-label="Print"]')
  const paper = dialog.querySelector('.print-paper')
  const options = dialog.querySelector('[aria-label="Print options"]')
  const preview = dialog.querySelector('[aria-label="Live print preview"]')
  const labels = Array.from(options.querySelectorAll('.print-field > span, fieldset > span')).map((node) => node.textContent.trim())
  return {
    shortcutPrevented: ${shortcutPrevented},
    twoPane: Boolean(options && preview),
    focusInside: dialog.contains(document.activeElement),
    optionsLeftOfPreview: options.getBoundingClientRect().left < preview.getBoundingClientRect().left,
    paperRatio: paper.getBoundingClientRect().width / paper.getBoundingClientRect().height,
    labels,
    printerValue: dialog.querySelector('.print-options-pane select')?.value,
    printerLabel: dialog.querySelector('.print-options-pane select option:checked')?.textContent,
    systemPreviewPresent: /system preview/i.test(dialog.textContent),
    summary: dialog.querySelector('.print-preview-summary')?.textContent,
    nav: dialog.querySelector('.print-preview-navigation span')?.textContent,
  }
})()`)
if (!initial.shortcutPrevented || !initial.twoPane || !initial.focusInside || !initial.optionsLeftOfPreview || initial.paperRatio >= 1 || initial.printerValue !== '' || initial.printerLabel !== 'Default Windows printer' || initial.systemPreviewPresent) throw new Error(`Unexpected initial two-pane state: ${JSON.stringify(initial)}`)
for (const label of ['Printer', 'Copies', 'Pages', 'Paper', 'Orientation', 'Margins', 'Color', 'Page sizing']) {
  if (!initial.labels.includes(label)) throw new Error(`Missing print option ${label}: ${JSON.stringify(initial.labels)}`)
}
if (!initial.summary?.includes('Content fits') || initial.nav?.trim() !== '1 / 8') throw new Error(`Unexpected live preview summary: ${JSON.stringify(initial)}`)

await evaluate(`document.querySelector('button[aria-label="Next print page"]').click(); document.querySelector('button[aria-label="Next print page"]').click(); document.querySelector('button[aria-label="Next print page"]').click()`)
await waitFor(`document.querySelector('.print-preview-summary')?.textContent.includes('PDF page 4')`, 'mixed-size page navigation')
const autoOrientationStable = await evaluate(`document.querySelector('.print-preview-toolbar span')?.textContent.includes('Letter portrait')`)
if (!autoOrientationStable) throw new Error('Auto orientation changed with preview navigation instead of staying job-wide.')

await evaluate(`(() => {
  const findSelect = (name) => Array.from(document.querySelectorAll('.print-field')).find((label) => label.querySelector(':scope > span')?.textContent.trim() === name)?.querySelector('select')
  const setSelect = (name, value) => { const select = findSelect(name); Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(select, value); select.dispatchEvent(new Event('change', { bubbles: true })) }
  setSelect('Paper', 'Legal')
  setSelect('Orientation', 'landscape')
  setSelect('Margins', 'none')
  setSelect('Color', 'bw')
  setSelect('Page sizing', 'custom')
})()`)
await waitFor(`document.querySelector('.print-preview-toolbar span')?.textContent.includes('Legal landscape') && Boolean(document.querySelector('.print-scale-field'))`, 'live paper/orientation/custom controls refresh')
await evaluate(`(() => { const input = document.querySelector('.print-scale-field input'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '175'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
await waitFor(`document.querySelector('.print-preview-summary')?.textContent.includes('cropped')`, 'custom scale crop preview')

const changed = await evaluate(`(() => {
  const paper = document.querySelector('.print-paper').getBoundingClientRect()
  const image = document.querySelector('.print-page-image')
  const paperElement = document.querySelector('.print-paper')
  const paperComputed = getComputedStyle(paperElement)
  const stageElement = document.querySelector('.print-preview-stage')
  const stageRect = stageElement.getBoundingClientRect()
  return {
    paperRatio: paper.width / paper.height,
    paperRect: { width: paper.width, height: paper.height },
    paperBox: { offsetWidth: paperElement.offsetWidth, offsetHeight: paperElement.offsetHeight, computedWidth: paperComputed.width, computedHeight: paperComputed.height },
    stageBox: { width: stageRect.width, height: stageRect.height, clientWidth: stageElement.clientWidth, clientHeight: stageElement.clientHeight },
    paperStyle: paperElement.getAttribute('style'),
    computedAspectRatio: paperComputed.aspectRatio,
    grayscale: getComputedStyle(image).filter,
    imageStyle: image.getAttribute('style'),
    selectedValues: Array.from(document.querySelectorAll('.print-options-pane select')).map((select) => select.value),
    cropped: document.querySelector('.print-preview-summary')?.textContent,
    scale: document.querySelector('.print-preview-toolbar span')?.textContent,
  }
})()`)
if (changed.paperRatio <= 1 || !changed.grayscale.includes('grayscale') || !changed.cropped?.includes('cropped') || !changed.scale?.includes('175%')) throw new Error(`Live print preview did not follow controls: ${JSON.stringify(changed)}`)

await evaluate(`(() => { const input = document.querySelector('input[aria-label="Custom page range"]'); input.focus(); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '4, 7'); input.dispatchEvent(new Event('input', { bubbles: true })); })()`)
await waitFor(`document.querySelector('.print-preview-navigation span')?.textContent.trim() === '1 / 2' && document.querySelector('.print-preview-summary')?.textContent.includes('PDF page 4')`, 'custom range drives preview pages')
await evaluate(`document.querySelector('button[aria-label="Next print page"]').click()`)
await waitFor(`document.querySelector('.print-preview-summary')?.textContent.includes('PDF page 7')`, 'preview navigation')

const focusTrap = await evaluate(`(() => {
  const dialog = document.querySelector('[role="dialog"][aria-label="Print"]')
  const focusable = Array.from(dialog.querySelectorAll('button, input, select, [tabindex="0"]')).filter((element) => !element.disabled && element.getClientRects().length > 0)
  focusable.at(-1).focus()
  const prevented = !focusable.at(-1).dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }))
  return { prevented, wrapped: document.activeElement === focusable[0] }
})()`)
if (!focusTrap.prevented || !focusTrap.wrapped) throw new Error(`Print focus trap failed: ${JSON.stringify(focusTrap)}`)

await evaluate(`document.querySelector('[role="dialog"][aria-label="Print"]').dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))`)
await waitFor(`!document.querySelector('[role="dialog"][aria-label="Print"]')`, 'Escape closes Print dialog')
const focusRestored = await evaluate(`document.activeElement?.getAttribute('aria-label') === 'Print (Ctrl+P)'`)
socket.close()

if (!focusRestored) throw new Error('Print dialog did not restore focus to its toolbar trigger.')
console.log(JSON.stringify({ initial, changed, focusTrap, focusRestored }))
