// Viewer interaction regression (run via scripts/run-viewer-interaction-regression.mjs
// against its generated 40-page reading fixture). Covers navigation alignment,
// screen-wise PageDown, focus return, undo/redo keys, Delete scoping, search
// hit visibility, page tracking at low zoom and with an idle selection, copy
// line breaks, Select All, edits on a non-current page and paste routing.
// The system clipboard and window.confirm are stubbed inside the page.
import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9412)
const deadline = Date.now() + 200_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Electron has not exposed its renderer yet.
  }
  if (!target) await pause(50)
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
  else if (message.result?.exceptionDetails) request.reject(new Error(message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text))
  else request.resolve(message.result?.result?.value ?? message.result)
})
function call(method, params = {}) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method, params }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}
const evaluate = (expression) => call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
async function waitFor(expression, label, timeout = 20_000) {
  const until = Math.min(deadline, Date.now() + timeout)
  while (Date.now() < until) {
    if (await evaluate(expression)) return
    await pause(40)
  }
  throw new Error(`Timed out: ${label}`)
}

const failures = []
const results = {}
function check(label, condition, details) {
  results[label] = { passed: Boolean(condition), ...(details && typeof details === 'object' ? details : { details }) }
  if (!condition) failures.push(label)
}

const key = (name, options = {}) => evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', ${JSON.stringify({ key: name, bubbles: true, cancelable: true, ...options })}))`)
async function realKey(name, code, keyCode, text) {
  await call('Input.dispatchKeyEvent', { type: text ? 'keyDown' : 'rawKeyDown', key: name, code, windowsVirtualKeyCode: keyCode, ...(text ? { text } : {}) })
  await call('Input.dispatchKeyEvent', { type: 'keyUp', key: name, code, windowsVirtualKeyCode: keyCode })
}
async function click(x, y) {
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x, y })
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', buttons: 1, clickCount: 1 })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', buttons: 0, clickCount: 1 })
}
const pageValue = () => evaluate(`document.querySelector('input[aria-label="Current page"]')?.value`)
async function goToPageField(number, refocus = true) {
  await evaluate(`(() => {
    const input = document.querySelector('input[aria-label="Current page"]')
    input.focus()
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
    setter.call(input, '${number}')
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
  })()`)
  await pause(700)
  // Back to the document, as a click on the page does.
  if (refocus) await evaluate(`document.querySelector('.viewer').focus({ preventScroll: true })`)
}
const slotOffset = (index) => evaluate(`(() => {
  const viewer = document.querySelector('.viewer')
  const slot = document.querySelector('.continuous-page-slot[data-page-index="${index}"]')
  return slot.getBoundingClientRect().top - viewer.getBoundingClientRect().top
})()`)
const viewerState = () => evaluate(`(() => {
  const viewer = document.querySelector('.viewer')
  return { scrollTop: viewer.scrollTop, clientHeight: viewer.clientHeight, scrollHeight: viewer.scrollHeight, zoom: document.querySelector('input[aria-label="Zoom percentage"]')?.value }
})()`)
async function fitWidth() {
  await key('2', { ctrlKey: true })
  await pause(900)
}
async function setSidebar(open) {
  const isOpen = `Boolean(document.querySelector('.sidebar [role="tab"]'))`
  if (await evaluate(isOpen) !== open) await key('F4')
  await waitFor(`${isOpen} === ${open}`, open ? 'sidebar open' : 'sidebar closed')
  await pause(300)
}
const toast = () => evaluate(`document.querySelector('.toast')?.textContent || ''`)

try {
await call('Emulation.setFocusEmulationEnabled', { enabled: true }).catch(() => {})
await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 40`, 'continuous 40-page document', 60_000)
await waitFor(`document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]').length >= 3`, 'first page text layer', 60_000)
await evaluate(`(() => {
  window.__clip = { text: '', writes: [] }
  const clipboard = navigator.clipboard
  Object.defineProperty(clipboard, 'writeText', { configurable: true, value: async (text) => { window.__clip.text = String(text); window.__clip.writes.push(String(text)) } })
  Object.defineProperty(clipboard, 'readText', { configurable: true, value: async () => window.__clip.text })
  Object.defineProperty(clipboard, 'read', { configurable: true, value: async () => { throw new Error('No image on the stub clipboard') } })
  window.__confirms = []
  window.confirm = (message) => { window.__confirms.push(String(message)); return false }
})()`)

// pdf-viewer-interaction-6: navigation aligns the page top; PageDown moves by a screen.
await fitWidth()
await goToPageField(5, false)
const alignOffset = await slotOffset(4)
check('navigationAlignsPageTop', Math.abs(alignOffset) <= 3 && await pageValue() === '5', { alignOffset })

// pdf-viewer-interaction-4: a click on the page takes focus from the page box.
const pageBox = await evaluate(`(() => {
  const span = [...document.querySelectorAll('.continuous-page-slot[data-page-index="4"] [data-text-item="true"]')].find((item) => item.textContent.includes('Body line 3'))
  const rect = span.getBoundingClientRect()
  return { x: rect.left + 4, y: rect.top + rect.height / 2, before: document.activeElement?.getAttribute('aria-label') }
})()`)
await click(pageBox.x + 300, pageBox.y)
const focusAfterPageClick = await evaluate(`document.activeElement?.classList.contains('viewer')`)
const pageDownSteps = []
let previousState = await viewerState()
for (let step = 0; step < 6 && (await pageValue()) === '5'; step += 1) {
  await realKey('PageDown', 'PageDown', 34)
  await pause(450)
  const state = await viewerState()
  pageDownSteps.push({ delta: Math.round(state.scrollTop - previousState.scrollTop), page: await pageValue(), offset6: Math.round(await slotOffset(5)) })
  previousState = state
}
const maxStep = Math.max(...pageDownSteps.map((item) => item.delta))
const lastStep = pageDownSteps.at(-1)
check('clickFocusesDocument', focusAfterPageClick, { before: pageBox.before })
check('pageDownScrollsByScreen', pageDownSteps.length >= 2 && maxStep <= previousState.clientHeight + 2
  && lastStep?.page === '6' && Math.abs(lastStep.offset6) <= 3, { pageDownSteps, clientHeight: previousState.clientHeight })

// pdf-viewer-interaction-11: Ctrl+Shift+Z redoes.
await evaluate(`document.querySelector('button[aria-label="Bookmark this page"]')?.click()`)
await waitFor(`Boolean(document.querySelector('button[aria-label="Remove bookmark"]'))`, 'bookmark added')
await key('z', { ctrlKey: true })
await waitFor(`Boolean(document.querySelector('button[aria-label="Bookmark this page"]'))`, 'undo removes bookmark')
await key('Z', { ctrlKey: true, shiftKey: true })
await pause(300)
check('ctrlShiftZRedoes', await evaluate(`Boolean(document.querySelector('button[aria-label="Remove bookmark"]'))`))

// pdf-viewer-interaction-10: Delete on a bookmark never offers to delete pages.
await setSidebar(true)
await evaluate(`[...document.querySelectorAll('.sidebar [role="tab"]')].find((tab) => tab.textContent === 'Bookmarks').click()`)
await waitFor(`Boolean(document.querySelector('.bookmark-main'))`, 'bookmark row')
await evaluate(`document.querySelector('.bookmark-main').focus()`)
await key('Delete')
await pause(200)
const confirmsAfterBookmarkDelete = await evaluate(`window.__confirms.length`)
await evaluate(`[...document.querySelectorAll('.sidebar [role="tab"]')].find((tab) => tab.textContent === 'Pages').click()`)
await waitFor(`Boolean(document.querySelector('.thumbnails button'))`, 'thumbnail buttons')
await evaluate(`document.querySelector('.thumbnails button').focus()`)
await key('Delete')
await pause(200)
const confirmsAfterThumbnailDelete = await evaluate(`window.__confirms.length`)
check('deleteScopedToPagesPanel', confirmsAfterBookmarkDelete === 0 && confirmsAfterThumbnailDelete === 1, { confirmsAfterBookmarkDelete, confirmsAfterThumbnailDelete })
await setSidebar(false)
await fitWidth()

// pdf-viewer-interaction-7: search hits at page edges stay visible.
await key('f', { ctrlKey: true })
await waitFor(`document.activeElement?.matches('input[aria-label="Find in document"]')`, 'Ctrl+F search focus')
await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Find in document"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, 'edge marker')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
})()`)
await waitFor(`Boolean(document.querySelector('[data-search-highlight="true"][data-search-active="true"] [data-search-match-rect="true"]'))`, 'first edge hit')
const hitVisible = () => evaluate(`(() => {
  const rect = document.querySelector('[data-search-highlight="true"][data-search-active="true"] [data-search-match-rect="true"]')?.getBoundingClientRect()
  const view = document.querySelector('.viewer').getBoundingClientRect()
  return { visible: Boolean(rect && rect.bottom > view.top && rect.top < view.bottom), top: rect ? Math.round(rect.top - view.top) : null, page: document.querySelector('[data-search-highlight="true"][data-search-active="true"]')?.getAttribute('data-page-index') }
})()`)
await pause(700)
const hits = [await hitVisible()]
for (let index = 0; index < 3; index += 1) {
  await evaluate(`document.querySelector('input[aria-label="Find in document"]').dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))`)
  await pause(800)
  hits.push(await hitVisible())
}
check('searchHitsStayVisible', hits.every((hit) => hit.visible) && new Set(hits.map((hit) => hit.page)).size >= 3, { hits })

// pdf-viewer-interaction-4: clicking the page after searching returns keyboard focus.
const textPoint = await evaluate(`(() => {
  const view = document.querySelector('.viewer').getBoundingClientRect()
  const span = [...document.querySelectorAll('.viewer [data-text-item="true"]')].find((item) => {
    const rect = item.getBoundingClientRect()
    return item.textContent.includes('Body line') && rect.top > view.top + 40 && rect.bottom < view.bottom - 40
  })
  const rect = span.getBoundingClientRect()
  return { x: rect.left + 6, y: rect.top + rect.height / 2 }
})()`)
await click(textPoint.x, textPoint.y)
await pause(150)
const focusAfterSearchClick = await evaluate(`document.activeElement?.classList.contains('viewer')`)
await realKey('h', 'KeyH', 72, 'h')
await pause(200)
const handActive = await evaluate(`document.querySelector('button[aria-label="Hand tool (H)"]')?.classList.contains('is-active')`)
await realKey('v', 'KeyV', 86, 'v')
check('searchThenClickRestoresShortcuts', focusAfterSearchClick && handActive, { focusAfterSearchClick, handActive })

// pdf-viewer-interaction-8: zooming after scrolling away never jumps back to the old hit.
const hitPage = Number((await hitVisible()).page)
const farPage = hitPage > 20 ? 5 : 30
await evaluate(`(() => {
  const viewer = document.querySelector('.viewer')
  const slot = document.querySelector('.continuous-page-slot[data-page-index="${farPage}"]')
  viewer.scrollTop += slot.getBoundingClientRect().top - viewer.getBoundingClientRect().top + 40
})()`)
await pause(900)
const beforeZoom = await pageValue()
await key('-', { ctrlKey: true })
await pause(1400)
const afterZoom = await evaluate(`(() => {
  const view = document.querySelector('.viewer').getBoundingClientRect()
  const visible = (index) => {
    const rect = document.querySelector('.continuous-page-slot[data-page-index="' + index + '"]')?.getBoundingClientRect()
    return Boolean(rect && rect.bottom > view.top && rect.top < view.bottom)
  }
  const page = Number(document.querySelector('input[aria-label="Current page"]').value)
  return { page, hitVisible: visible(${hitPage}), indicatorPageVisible: visible(page - 1) }
})()`)
check('zoomKeepsReadingPosition', beforeZoom === String(farPage + 1) && !afterZoom.hitVisible && afterZoom.indicatorPageVisible, { hitPage, farPage, beforeZoom, afterZoom })
await setSidebar(false)

// pdf-viewer-interaction-12: first and last pages become current at low zoom.
for (let index = 0; index < 20; index += 1) await key('-', { ctrlKey: true })
await pause(900)
await evaluate(`document.querySelector('.viewer').scrollTop = 0`)
await pause(500)
const topPage = await pageValue()
await evaluate(`(() => { const viewer = document.querySelector('.viewer'); viewer.scrollTop = viewer.scrollHeight })()`)
await pause(500)
const bottomPage = await pageValue()
check('lowZoomEdgePages', topPage === '1' && bottomPage === '40', { topPage, bottomPage, zoom: (await viewerState()).zoom })

// pdf-viewer-interaction-1: an idle selection neither pins every page nor freezes tracking.
await fitWidth()
await goToPageField(1)
const selectionStart = await evaluate(`(() => {
  const span = [...document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]')].find((item) => item.textContent.startsWith('Section 1'))
  const node = span.firstChild
  const range = document.createRange()
  range.setStart(node, 0)
  range.setEnd(node, 7)
  getSelection().removeAllRanges()
  getSelection().addRange(range)
  window.__selectionNode = node
  return getSelection().toString()
})()`)
for (let index = 1; index <= 25; index += 1) {
  await evaluate(`(() => {
    const viewer = document.querySelector('.viewer')
    const slot = document.querySelector('.continuous-page-slot[data-page-index="${index}"]')
    viewer.scrollTop += slot.getBoundingClientRect().top - viewer.getBoundingClientRect().top
  })()`)
  await pause(140)
}
await pause(900)
const afterSelectionScroll = await evaluate(`(() => ({
  mounted: document.querySelectorAll('.continuous-page-slot .page-surface').length,
  page: document.querySelector('input[aria-label="Current page"]').value,
  selection: getSelection().toString(),
  anchorConnected: Boolean(window.__selectionNode?.isConnected),
}))()`)
check('idleSelectionKeepsMemoryBounded', afterSelectionScroll.mounted <= 10 && Math.abs(Number(afterSelectionScroll.page) - 26) <= 1
  && afterSelectionScroll.selection === selectionStart && afterSelectionScroll.anchorConnected, { selectionStart, ...afterSelectionScroll })
await evaluate(`getSelection().removeAllRanges()`)

// pdf-viewer-interaction-5: copied text keeps line breaks.
await goToPageField(3)
await waitFor(`[...document.querySelectorAll('.continuous-page-slot[data-page-index="2"] [data-text-item="true"]')].some((item) => item.textContent.startsWith('the lazy dog'))`, 'page 3 text')
const copied = await evaluate(`(() => {
  const spans = [...document.querySelectorAll('.continuous-page-slot[data-page-index="2"] [data-text-item="true"]')]
  const first = spans.find((item) => item.textContent.startsWith('The quick')).firstChild
  const last = spans.find((item) => item.textContent.startsWith('the lazy dog')).firstChild
  const range = document.createRange()
  range.setStart(first, 4)
  range.setEnd(last, 12)
  getSelection().removeAllRanges()
  getSelection().addRange(range)
  const data = new DataTransfer()
  document.dispatchEvent(new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }))
  const text = data.getData('text/plain')
  getSelection().removeAllRanges()
  return text
})()`)
check('copyKeepsLineBreaks', copied === 'quick brown fox jumps over\nthe lazy dog', { copied })

// pdf-viewer-interaction-9: Ctrl+A + copy yields the whole document.
await evaluate(`window.__clip.writes = []`)
await key('a', { ctrlKey: true })
const selectAllSync = await evaluate(`(() => {
  const data = new DataTransfer()
  document.dispatchEvent(new ClipboardEvent('copy', { clipboardData: data, bubbles: true, cancelable: true }))
  return data.getData('text/plain')
})()`)
await waitFor(`window.__clip.writes.length > 0`, 'whole-document copy')
const selectAllText = await evaluate(`window.__clip.writes.at(-1)`)
const headings = (selectAllText.match(/Section \d+ heading/g) || []).length
check('selectAllCopiesWholeDocument', headings === 40 && selectAllText.includes('jumps over\nthe lazy dog')
  && selectAllText.includes('Section 40 heading') && selectAllSync.length > 0, { headings, syncLength: selectAllSync.length, length: selectAllText.length })
await key('Escape')
await evaluate(`getSelection().removeAllRanges()`)

// pdf-viewer-interaction-3: editing a heading on the next (non-current) page never scrolls.
await goToPageField(10)
await evaluate(`(() => {
  const viewer = document.querySelector('.viewer')
  const slot = document.querySelector('.continuous-page-slot[data-page-index="10"]')
  const view = viewer.getBoundingClientRect()
  viewer.scrollTop += slot.getBoundingClientRect().top - (view.top + view.height * 0.62)
})()`)
await pause(700)
await waitFor(`[...document.querySelectorAll('.continuous-page-slot[data-page-index="10"] [data-text-item="true"]')].some((item) => item.textContent.startsWith('Section 11'))`, 'page 11 heading text')
const pageBeforeEdit = await pageValue()
await evaluate(`document.querySelector('button[aria-label="Edit text (E)"]').click()`)
await pause(300)
await waitFor(`[...document.querySelectorAll('.continuous-page-slot[data-page-index="10"] [data-text-item="true"]')].some((item) => item.textContent.startsWith('Section 11'))`, 'page 11 heading text in edit mode')
const heading = await evaluate(`(() => {
  const span = [...document.querySelectorAll('.continuous-page-slot[data-page-index="10"] [data-text-item="true"]')].find((item) => item.textContent.startsWith('Section 11'))
  const rect = span.getBoundingClientRect()
  return { x: rect.left + 10, y: rect.top + rect.height / 2, scrollTop: document.querySelector('.viewer').scrollTop }
})()`)
await click(heading.x, heading.y)
await pause(900)
const editState = await evaluate(`(() => {
  const viewer = document.querySelector('.viewer')
  const editor = document.querySelector('.viewer textarea')
  const view = viewer.getBoundingClientRect()
  const rect = editor?.getBoundingClientRect()
  return { scrollTop: viewer.scrollTop, editorVisible: Boolean(rect && rect.top >= view.top && rect.bottom <= view.bottom), page: document.querySelector('input[aria-label="Current page"]').value }
})()`)
check('editOnNextPageDoesNotScroll', pageBeforeEdit === '10' && Math.abs(editState.scrollTop - heading.scrollTop) <= 3 && editState.editorVisible,
  { pageBeforeEdit, scrollBefore: heading.scrollTop, ...editState })

// pdf-edit-pipeline-12: Ctrl+V honours newer system clipboard content.
await evaluate(`[...document.querySelectorAll('.edit-inspector button')].find((button) => button.textContent.trim() === 'Copy')?.click()`)
await pause(200)
const internalText = await evaluate(`window.__clip.text`)
await evaluate(`document.querySelector('.viewer').focus()`)
await key('v', { ctrlKey: true })
await pause(700)
const internalToast = await toast()
await key('Escape')
await evaluate(`(() => { window.__clip.text = 'Invoice #4471'; document.querySelector('.viewer').focus() })()`)
await key('v', { ctrlKey: true })
await pause(700)
const systemPaste = await evaluate(`({ value: document.querySelector('.viewer textarea')?.value, toast: document.querySelector('.toast')?.textContent || '' })`)
check('pasteHonoursSystemClipboard', Boolean(internalText) && internalToast.includes('Text box pasted')
  && systemPaste.value === 'Invoice #4471' && systemPaste.toast.includes('Text pasted'), { internalText, internalToast, systemPaste })
await key('Escape')

} catch (error) {
  failures.push(`aborted: ${error.message}`)
}
socket.close()
const output = { passed: failures.length === 0, failures, results }
console.log(JSON.stringify(output))
if (failures.length) process.exit(1)
