// Recognize text (OCR) UI smoke over the Chrome DevTools Protocol. Started by
// run-ocr-regression.mjs, which launches the built app on a fixture and sets:
//   SIMPLE_OCR_UI_PORT      the app's remote debugging port
//   SIMPLE_OCR_UI_SCENARIO  main | stop | mixed (see run-ocr-regression.mjs)
//   SIMPLE_OCR_UI_TRUTH     JSON: the fixture's lines and words (PDF points)
// Prints one JSON line with what it measured; any failed check throws.
import assert from 'node:assert/strict'
import { readFile } from 'node:fs/promises'

const port = Number(process.env.SIMPLE_OCR_UI_PORT || 9547)
const scenario = process.env.SIMPLE_OCR_UI_SCENARIO || 'main'
const truth = JSON.parse(await readFile(process.env.SIMPLE_OCR_UI_TRUTH, 'utf8'))
const relaxedTiming = process.env.SIMPLE_OCR_UI_RELAXED_TIMING === '1'
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))
const log = (...parts) => console.error(`[smoke-ocr ${scenario}]`, ...parts.map((part) => typeof part === 'string' ? part : JSON.stringify(part)))
const collapse = (text) => String(text || '').replace(/\s+/g, ' ').trim()

const startupDeadline = Date.now() + 90_000
let target
while (!target && Date.now() < startupDeadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl && /index\.html/.test(item.url))
  } catch { /* the app is opening its debugging endpoint */ }
  if (!target) await pause(50)
}
if (!target) throw new Error('The renderer did not start')

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
  if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params.exceptionDetails?.exception?.description || message.params.exceptionDetails?.text)
  const request = pending.get(message.id)
  if (!request) return
  pending.delete(message.id)
  clearTimeout(request.timeout)
  if (message.error) request.reject(new Error(message.error.message))
  else if (message.result?.exceptionDetails) request.reject(new Error(message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text))
  else request.resolve(message.result?.result?.value ?? message.result)
})

function call(method, params = {}, timeoutMs = 30_000) {
  return new Promise((resolve, reject) => {
    const id = ++requestId
    const timeout = setTimeout(() => { pending.delete(id); reject(new Error(`Timed out: ${method}`)) }, timeoutMs)
    pending.set(id, { resolve, reject, timeout })
    socket.send(JSON.stringify({ id, method, params }))
  })
}
const evaluate = (expression, options = {}) => call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true, ...options })
await call('Runtime.enable')

async function state() {
  return evaluate(`({
    toast: document.querySelector('.toast')?.textContent || '',
    dialog: document.querySelector('[role="dialog"], [role="alertdialog"]')?.getAttribute('aria-label') || '',
    progress: document.querySelector('.ocr-progress-status')?.textContent || '',
    inspector: document.querySelector('.edit-inspector-empty > strong')?.textContent || '',
    page: document.querySelector('input[aria-label="Current page"]')?.value,
    dialogText: (document.querySelector('.ocr-dialog')?.textContent || '').slice(0, 400),
  })`)
}

async function waitFor(expression, label, milliseconds = 20_000) {
  const deadline = Date.now() + milliseconds
  while (Date.now() < deadline) {
    if (await evaluate(`Boolean(${expression})`)) return
    await pause(30)
  }
  throw new Error(`Timed out: ${label}; ${JSON.stringify(await state())}`)
}

async function key(keyName, code, modifiers = 0) {
  const virtual = ({ Enter: 13, Escape: 27, Delete: 46, Tab: 9 })[keyName] || keyName.toUpperCase().charCodeAt(0)
  const parameters = { key: keyName, code, modifiers, windowsVirtualKeyCode: virtual, nativeVirtualKeyCode: virtual }
  await call('Input.dispatchKeyEvent', { type: 'keyDown', ...parameters })
  await call('Input.dispatchKeyEvent', { type: 'keyUp', ...parameters })
}

async function mouse(type, point, extra = {}) {
  await call('Input.dispatchMouseEvent', { type, x: point.x, y: point.y, button: 'left', buttons: type === 'mouseReleased' ? 0 : 1, clickCount: 1, ...extra })
}

async function clickAt(point) {
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: point.x, y: point.y })
  await mouse('mousePressed', point)
  await mouse('mouseReleased', point)
}

async function clickExpression(expression, label = expression) {
  await waitFor(`Boolean(${expression})`, `click target ${label}`)
  const point = await evaluate(`(() => {
    const element = ${expression}
    element.scrollIntoView({ block: 'nearest', inline: 'nearest' })
    const box = element.getBoundingClientRect()
    return { x: box.left + box.width / 2, y: box.top + box.height / 2 }
  })()`)
  await clickAt(point)
}

const buttonWithText = (text, scope = 'document') => `[...${scope}.querySelectorAll('button')].find((button) => button.textContent.trim() === ${JSON.stringify(text)} && !button.disabled)`
const surface = `document.querySelector('.continuous-page-slot.is-current .page-surface')`
const spans = `[...document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')]`

/** PDF point -> window point on the current page (unrotated fixtures). */
async function pagePoint(x, y) {
  const geometry = await evaluate(`(() => { const box = ${surface}.getBoundingClientRect(); return { left: box.left, top: box.top, width: box.width, height: box.height } })()`)
  const scale = geometry.width / (truth.pageWidth || 612)
  return { x: geometry.left + x * scale, y: geometry.top + ((truth.pageHeight || 792) - y) * scale, scale, geometry }
}

async function waitForPageReady(label) {
  await waitFor(`Boolean(${surface}) && !${surface}.classList.contains('is-page-transitioning') && !document.querySelector('.document-refreshing') && !document.querySelector('.busy-overlay')`, label, 45_000)
}

async function search(query) {
  await evaluate(`document.activeElement?.blur?.()`)
  await key('f', 'KeyF', 2)
  await waitFor(`document.activeElement?.getAttribute('aria-label') === 'Find in document'`, 'search box focused')
  await key('a', 'KeyA', 2)
  await call('Input.insertText', { text: query })
  await key('Enter', 'Enter')
  // Done once matches are listed, or the pane says what it found nothing on.
  await waitFor(`(() => {
    const status = document.querySelector('.search-status')?.textContent || ''
    if (/Searching/.test(status)) return false
    return /matches/.test(status) || Boolean(document.querySelector('.search-ocr-hint'))
  })()`, `search ${query}`, 30_000)
  return evaluate(`(() => ({
    status: document.querySelector('.search-status')?.textContent || '',
    hint: document.querySelector('.search-ocr-hint span')?.textContent || '',
    results: [...document.querySelectorAll('.search-results > button')].map((button) => ({
      page: Number((button.querySelector('span')?.textContent || '').replace(/\\D/g, '')),
      count: Number((button.querySelector('small')?.textContent || '').replace(/\\D/g, '')),
    })),
  }))()`)
}

const totalMatches = (found) => found.results.reduce((sum, result) => sum + result.count, 0)

async function toastMatching(pattern, label, milliseconds = 60_000) {
  await waitFor(`${pattern}.test(document.querySelector('.toast')?.textContent || '')`, label, milliseconds)
  return (await state()).toast
}

async function dismissToast() {
  await evaluate(`document.querySelector('.toast .toast-close')?.click()`)
  await waitFor(`!document.querySelector('.toast')`, 'toast dismissed', 5_000).catch(() => {})
}

function iou(a, b) {
  const x0 = Math.max(a.left, b.left)
  const y0 = Math.max(a.top, b.top)
  const x1 = Math.min(a.left + a.width, b.left + b.width)
  const y1 = Math.min(a.top + a.height, b.top + b.height)
  const overlap = Math.max(0, x1 - x0) * Math.max(0, y1 - y0)
  const union = a.width * a.height + b.width * b.height - overlap
  return union > 0 ? overlap / union : 0
}

/** Watch for the progress dialog while a run goes, until `doneExpression` holds. */
async function watchRun(doneExpression, label, milliseconds = 90_000) {
  const started = Date.now()
  let sawProgress = false
  let maxValue = 0
  const deadline = started + milliseconds
  while (Date.now() < deadline) {
    const snapshot = await evaluate(`({
      progress: Boolean(document.querySelector('.ocr-dialog .ocr-progress')),
      value: Number(document.querySelector('.ocr-progress-bar')?.getAttribute('aria-valuenow') || 0),
      done: Boolean(${doneExpression}),
    })`)
    sawProgress ||= snapshot.progress
    maxValue = Math.max(maxValue, snapshot.value)
    if (snapshot.done) return { ms: Date.now() - started, sawProgress, maxValue }
    await pause(20)
  }
  throw new Error(`Timed out: ${label}; ${JSON.stringify(await state())}`)
}

async function openRecognizeDialog() {
  await clickExpression(`document.querySelector('button[aria-label="Recognize text (OCR)"]')`, 'Recognize text toolbar button')
  await waitFor(`document.querySelector('.ocr-dialog')?.getAttribute('aria-label') === 'Recognize text'`, 'Recognize text dialog')
}

async function scenarioMain() {
  const report = {}
  const word = truth.words.find((item) => item.text === 'negotiations')
  assert.ok(word, 'fixture word "negotiations" is missing')
  await waitForPageReady('scanned page')
  report.textBefore = await evaluate(`${spans}.length`)
  assert.equal(report.textBefore, 0, 'the scanned page must start without text')

  // Search before recognising: nothing, and the hint offers recognition.
  const before = await search('negotiations')
  assert.equal(totalMatches(before), 0, 'a scan has no searchable text yet')
  assert.match(before.hint, /1 page has no searchable text/, `search hint: ${JSON.stringify(before)}`)
  report.searchHint = before.hint

  // Text export of a scan asks first.
  await evaluate(`document.activeElement?.blur?.()`)
  await key('e', 'KeyE', 2 | 8)
  await waitFor(`document.querySelector('.export-dialog[aria-label="Export As"]')`, 'Export As dialog')
  await clickExpression(`[...document.querySelectorAll('.export-format-grid button')].find((button) => button.textContent.includes('Plain text'))`, 'Plain text format')
  await clickExpression(buttonWithText('Export', `document.querySelector('.export-dialog')`), 'Export button')
  await waitFor(`document.querySelector('[aria-label="Scanned pages in this export"]')`, 'export scan prompt')
  report.exportPrompt = collapse(await evaluate(`document.querySelector('[aria-label="Scanned pages in this export"] .ocr-dialog-intro').textContent`))
  assert.match(report.exportPrompt, /1 page in this export is a scanned image/)
  await clickExpression(buttonWithText('Cancel', `document.querySelector('[aria-label="Scanned pages in this export"]')`), 'Cancel export prompt')
  await waitFor(`!document.querySelector('[aria-label="Scanned pages in this export"]')`, 'export prompt closed')

  // Edit mode: the inspector says what the page is; the scan is not a selectable picture.
  await clickExpression(`document.querySelector('button[aria-label="Edit text (E)"]')`, 'Edit PDF')
  await waitFor(`document.querySelector('.edit-inspector-empty > strong')?.textContent === 'This page is a scanned image'`, 'scanned-page inspector state')
  report.inspectorBefore = 'This page is a scanned image'
  await pause(300)
  report.pageImageSelectable = await evaluate(`document.querySelectorAll('.continuous-page-slot.is-current .detected-object').length`)
  assert.equal(report.pageImageSelectable, 0, 'the page-covering scan must not be offered as a selectable image')
  const margin = await pagePoint(18, 774)
  await clickAt(margin)
  await waitFor(`document.querySelector('.ocr-offer')`, 'offer after a click on the margin')
  assert.equal(await evaluate(`Boolean(document.querySelector('.object-edit-frame'))`), false, 'a click on the scan selected the page image (D1)')
  await key('Escape', 'Escape')
  await waitFor(`!document.querySelector('.ocr-offer')`, 'offer closed by Escape')

  // Click the word: offer -> Recognize text -> progress -> toast -> the clicked line opens for editing.
  const wordPoint = await pagePoint(word.centre.x, word.centre.y)
  await clickAt(wordPoint)
  await waitFor(`document.querySelector('.ocr-offer')`, 'offer after a click on a scanned word')
  report.offer = collapse(await evaluate(`document.querySelector('.ocr-offer p').textContent`))
  assert.match(report.offer, /This page is a scanned image\. Recognize its text to edit words directly\./)
  await clickExpression(buttonWithText('Recognize text', `document.querySelector('.ocr-offer')`), 'offer Recognize text')
  const recognized = await watchRun(`${spans}.some((span) => span.textContent.includes('negotiations'))`, 'recognised text in the text layer')
  report.recognizeToSearchableMs = recognized.ms
  report.progressDialogSeen = recognized.sawProgress
  assert.ok(recognized.sawProgress, 'the progress dialog did not appear')
  report.toast = await toastMatching(/Text recognized on 1 page/, 'completion toast')
  assert.match(report.toast, /Text recognized on 1 page\. You can now search, select, and edit it\./)
  if (!relaxedTiming) assert.ok(recognized.ms <= 4_000, `recognising one scanned page took ${recognized.ms} ms (budget 4 s)`)
  await waitFor(`document.querySelector('.inline-pdf-text-editor')?.value?.includes('negotiations')`, 'the clicked line opens for editing', 10_000)
  report.editAtLine = await evaluate(`document.querySelector('.inline-pdf-text-editor').value`)
  await key('Escape', 'Escape')
  await waitFor(`!document.querySelector('.inline-pdf-text-editor')`, 'edit closed')
  await waitFor(`document.querySelector('.edit-inspector-empty > strong')?.textContent === 'Scanned page with recognized text'`, 'recognised-scan inspector state')
  report.inspectorAfter = 'Scanned page with recognized text'
  // D1 on the recognised scan: the margin selects nothing and offers nothing.
  await clickAt(margin)
  await pause(400)
  assert.equal(await evaluate(`Boolean(document.querySelector('.object-edit-frame, .ocr-offer, .inline-pdf-text-editor'))`), false, 'a margin click on the recognised scan selected something')
  await dismissToast()

  // Ctrl+F finds the word, and the highlight sits on the scanned word.
  const found = await search('negotiations')
  assert.equal(totalMatches(found), 1, `search after recognising: ${JSON.stringify(found)}`)
  await waitFor(`document.querySelector('.continuous-page-slot.is-current .search-match-rect')`, 'search highlight')
  const highlight = await evaluate(`(() => { const box = document.querySelector('.continuous-page-slot.is-current .search-match-rect').getBoundingClientRect(); return { left: box.left, top: box.top, width: box.width, height: box.height } })()`)
  const centre = await pagePoint(word.centre.x, word.centre.y)
  const truthBox = { left: centre.x - (word.width * centre.scale) / 2, top: centre.y - (word.height * centre.scale) / 2, width: word.width * centre.scale, height: word.height * centre.scale }
  report.highlightIoU = Math.round(iou(highlight, truthBox) * 1000) / 1000
  log('recognised', { ms: report.recognizeToSearchableMs, iou: report.highlightIoU, highlight, truthBox })
  assert.ok(report.highlightIoU >= 0.5, `highlight IoU ${report.highlightIoU} < 0.5 (${JSON.stringify({ highlight, truthBox })})`)

  // Drag-select line 2 and copy it.
  await key('Escape', 'Escape')
  await evaluate(`document.activeElement?.blur?.()`)
  await clickExpression(`document.querySelector('button[aria-label="Select text (V)"]')`, 'Select text tool')
  const lineText = truth.lines[1]
  const firstWord = lineText.split(' ')[0]
  // From inside the first glyph of the line to inside its last one (the scan
  // is skewed, so the line's bounding box is taller than the line).
  const bounds = await evaluate(`(() => {
    const line = ${spans}.find((span) => span.textContent.trim().startsWith(${JSON.stringify(firstWord)}) && span.textContent.length > ${Math.floor(lineText.length / 2)})
    if (!line) return null
    const lineId = line.dataset.textLine
    const parts = ${spans}.filter((span) => span.dataset.textLine === lineId && span.textContent.trim())
      .sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)
    // Just inside the glyph's outer edge: the caret lands before the first glyph and after the last.
    const glyph = (span, index, side) => {
      const range = document.createRange()
      range.setStart(span.firstChild, index)
      range.setEnd(span.firstChild, index + 1)
      const box = range.getBoundingClientRect()
      const inset = Math.min(0.4, box.width * 0.15)
      return { x: side === 'start' ? box.left + inset : box.right - inset, y: box.top + box.height / 2 }
    }
    const first = parts[0]
    const last = parts.at(-1)
    const lastIndex = last.textContent.trimEnd().length - 1
    return { spans: parts.length, start: glyph(first, Math.max(0, first.textContent.search(/\\S/)), 'start'), end: glyph(last, lastIndex, 'end') }
  })()`)
  assert.ok(bounds, `line 2 ("${lineText}") is not in the text layer`)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: bounds.start.x, y: bounds.start.y })
  await mouse('mousePressed', bounds.start)
  for (let step = 1; step <= 8; step += 1) {
    await mouse('mouseMoved', { x: bounds.start.x + (bounds.end.x - bounds.start.x) * step / 8, y: bounds.start.y + (bounds.end.y - bounds.start.y) * step / 8 })
  }
  await mouse('mouseReleased', bounds.end)
  report.selected = await evaluate(`window.getSelection()?.toString() || ''`)
  log('selection', { selected: report.selected, bounds, active: await evaluate(`document.activeElement?.tagName + '.' + (document.activeElement?.className || '')`), tool: await evaluate(`document.querySelector('.page-surface')?.className || ''`) })
  // What the app's copy handler writes (clipboardData cannot be read back during the event).
  await evaluate(`(() => {
    window.__ocrCopied = null
    const setData = DataTransfer.prototype.setData
    DataTransfer.prototype.setData = function (type, value) {
      if (type === 'text/plain') window.__ocrCopied = value
      return setData.call(this, type, value)
    }
  })()`)
  await evaluate(`document.execCommand('copy')`, { userGesture: true })
  report.copied = await evaluate(`String(window.__ocrCopied ?? '')`)
  assert.equal(collapse(report.copied), collapse(lineText), 'copied line differs from the truth line')
  await evaluate(`window.getSelection()?.removeAllRanges()`)

  // Undo removes the text; Redo brings it back without recognising again.
  await evaluate(`document.activeElement?.blur?.()`)
  await key('z', 'KeyZ', 2)
  await waitFor(`${spans}.length === 0 && Boolean(${surface})`, 'undo removed the recognised text')
  await waitForPageReady('page after undo')
  const afterUndo = await search('negotiations')
  report.undoMatches = totalMatches(afterUndo)
  assert.equal(report.undoMatches, 0, 'undo must remove the recognised text')
  await evaluate(`document.activeElement?.blur?.()`)
  const redoStarted = Date.now()
  await key('y', 'KeyY', 2)
  await waitFor(`${spans}.some((span) => span.textContent.includes('negotiations'))`, 'redo restored the text')
  report.redoToSearchableMs = Date.now() - redoStarted
  const afterRedo = await search('negotiations')
  report.redoToResultMs = Date.now() - redoStarted
  report.redoMatches = totalMatches(afterRedo)
  assert.equal(report.redoMatches, 1, 'redo must restore the recognised text')
  if (!relaxedTiming) assert.ok(report.redoToSearchableMs < 1_000, `redo took ${report.redoToSearchableMs} ms to make the text searchable`)

  // Undo, then Recognize again: the cached result makes it instant.
  await evaluate(`document.activeElement?.blur?.()`)
  await key('z', 'KeyZ', 2)
  await waitFor(`${spans}.length === 0`, 'second undo')
  await waitForPageReady('page after second undo')
  await dismissToast()
  await openRecognizeDialog()
  await waitFor(`/All pages that need it \\(1 of 1\\)/.test(document.querySelector('.ocr-scope')?.textContent || '')`, 'needed count')
  report.dialogScope = collapse(await evaluate(`document.querySelector('.ocr-scope').textContent`))
  await clickExpression(buttonWithText('Recognize', `document.querySelector('.ocr-dialog')`), 'dialog Recognize')
  const rerun = await watchRun(`${spans}.some((span) => span.textContent.includes('negotiations'))`, 'cached re-run')
  report.cachedRerunMs = rerun.ms
  await toastMatching(/Text recognized on 1 page/, 'cached re-run toast')
  if (!relaxedTiming) assert.ok(rerun.ms < 2_000, `a cached re-run took ${rerun.ms} ms`)

  // Save in place; the runner reads the saved file back.
  await dismissToast()
  await evaluate(`document.activeElement?.blur?.()`)
  await key('s', 'KeyS', 2)
  report.saveToast = await toastMatching(/Saved/, 'save toast', 30_000)
  return report
}

async function scenarioStop() {
  const report = {}
  await waitForPageReady('first scanned page')
  await openRecognizeDialog()
  await waitFor(`/All pages that need it \\(6 of 6\\)/.test(document.querySelector('.ocr-scope')?.textContent || '') && !document.querySelector('.ocr-counting')`, 'six pages need it', 60_000)
  report.estimate = await evaluate(`document.querySelector('.ocr-dialog .export-dialog-footer > span')?.textContent || ''`)
  await clickExpression(buttonWithText('Recognize', `document.querySelector('.ocr-dialog')`), 'dialog Recognize')
  // Stop once at least one page is finished and others are still running.
  await waitFor(`/Recognizing page [2-5] of 6/.test(document.querySelector('.ocr-progress-status')?.textContent || '')`, 'first page finished', 90_000)
  report.progressAtStop = await evaluate(`({ status: document.querySelector('.ocr-progress-status')?.textContent, value: document.querySelector('.ocr-progress-bar')?.getAttribute('aria-valuenow'), left: document.querySelector('.ocr-progress small')?.textContent })`)
  await clickExpression(buttonWithText('Stop', `document.querySelector('.ocr-dialog')`), 'Stop')
  report.toast = await toastMatching(/Stopped/, 'stopped toast', 60_000)
  const match = /text added to (\d+) of 6 pages/.exec(report.toast)
  assert.ok(match, `stop toast: ${report.toast}`)
  report.completed = Number(match[1])
  assert.ok(report.completed >= 1 && report.completed < 6, `pages completed before Stop: ${report.completed}`)
  await waitForPageReady('document after stop')
  await dismissToast()
  const found = await search('negotiations')
  report.searchablePages = found.results.map((result) => result.page)
  assert.equal(found.results.length, report.completed, `searchable pages ${JSON.stringify(found)} != completed ${report.completed}`)
  assert.ok(found.results.every((result) => result.count === 1), 'each finished page holds the word once')
  assert.match(found.hint, new RegExp(`${6 - report.completed} pages? ha(s|ve) no searchable text`), `hint: ${found.hint}`)
  return report
}

async function scenarioMixed() {
  const report = {}
  await waitForPageReady('mixed page')
  await openRecognizeDialog()
  await waitFor(`/All pages that need it \\(0 of 1\\)/.test(document.querySelector('.ocr-scope')?.textContent || '') && !document.querySelector('.ocr-counting')`, 'a mixed page is not counted as needing it')
  report.noteForNeeded = await evaluate(`document.querySelector('.ocr-note')?.textContent || ''`)
  const thisPage = `[...document.querySelectorAll('.ocr-scope label')].find((label) => label.textContent.includes('This page'))`
  await clickExpression(`${thisPage}?.querySelector('input')`, 'This page scope')
  if (!await evaluate(`Boolean(${thisPage}?.querySelector('input')?.checked)`)) await clickExpression(thisPage, 'This page label')
  await waitFor(`${thisPage}?.querySelector('input')?.checked`, 'This page scope chosen')
  await clickExpression(buttonWithText('Recognize', `document.querySelector('.ocr-dialog')`), 'dialog Recognize')
  report.toast = await toastMatching(/Text recognized on 1 page/, 'completion toast', 90_000)
  await waitForPageReady('mixed page after recognising')
  await dismissToast()
  const typedWords = [...new Set(truth.typedLines.join(' ').split(' ').filter((item) => /^[A-Za-z]{5,}$/.test(item)))].slice(0, 6)
  report.typedWordMatches = {}
  for (const typedWord of typedWords) {
    const found = await search(typedWord)
    report.typedWordMatches[typedWord] = totalMatches(found)
  }
  const duplicated = Object.entries(report.typedWordMatches).filter(([, count]) => count !== 1)
  assert.equal(duplicated.length, 0, `typed words found more (or less) than once: ${JSON.stringify(report.typedWordMatches)}`)
  const scanOnly = await search('Invoice')
  report.scanOnlyMatches = totalMatches(scanOnly)
  assert.equal(report.scanOnlyMatches, 1, 'a word only in the scanned part must be found once')
  return report
}

let report
try {
  if (scenario === 'main') report = await scenarioMain()
  else if (scenario === 'stop') report = await scenarioStop()
  else if (scenario === 'mixed') report = await scenarioMixed()
  else throw new Error(`Unknown scenario ${scenario}`)
  const appErrors = runtimeErrors.filter((message) => message && !/ResizeObserver loop/.test(message))
  assert.deepEqual(appErrors, [], 'uncaught renderer errors')
  console.log(JSON.stringify({ scenario, ...report }))
} finally {
  socket.close()
}
