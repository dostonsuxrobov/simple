import process from 'node:process'
import fs from 'node:fs/promises'
import path from 'node:path'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9380)
const query = process.env.SIMPLE_SEARCH_QUERY || 'scrolling verification'
const screenshotPath = process.env.SIMPLE_SEARCH_SCREENSHOT
  || path.resolve('qa-pdfs', 'search-highlight-runtime.png')
const deadline = Date.now() + 60_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Electron has not exposed its renderer yet.
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
  else if (message.result?.exceptionDetails) request.reject(new Error(message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text))
  else request.resolve(message.result?.result?.value ?? message.result)
})

function call(method, params = {}) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method, params }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

function evaluate(expression) {
  return call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
}

function highlightSnapshot() {
  return evaluate(`(() => {
    const group = document.querySelector('[data-search-highlight="true"][data-search-active="true"]')
    const rects = group ? Array.from(group.querySelectorAll('[data-search-match-rect="true"]')) : []
    const pageIndex = Number(group?.getAttribute('data-page-index'))
    const needle = ${JSON.stringify(query)}.toLocaleLowerCase()
    const sourceSpans = Number.isInteger(pageIndex)
      ? Array.from(document.querySelectorAll(\`.continuous-page-slot[data-page-index="\${pageIndex}"] [data-text-item="true"]\`))
      : []
    const sourceText = sourceSpans.map((span) => span.textContent || '').join(' ')
    const sourceOffset = sourceText.toLocaleLowerCase().indexOf(needle)
    const sourceEnd = sourceOffset + ${JSON.stringify(query)}.length
    let cursor = 0
    let startPoint = null
    let endPoint = null
    let startSpanIndex = -1
    let endSpanIndex = -1
    sourceSpans.forEach((span, spanIndex) => {
      const text = span.textContent || ''
      const spanStart = cursor
      const spanEnd = spanStart + text.length
      if (!startPoint && sourceOffset >= spanStart && sourceOffset < spanEnd && span.firstChild instanceof Text) {
        startPoint = { node: span.firstChild, offset: sourceOffset - spanStart }
        startSpanIndex = spanIndex
      }
      if (!endPoint && sourceEnd > spanStart && sourceEnd <= spanEnd && span.firstChild instanceof Text) {
        endPoint = { node: span.firstChild, offset: sourceEnd - spanStart }
        endSpanIndex = spanIndex
      }
      cursor = spanEnd + 1
    })
    const expectedRange = document.createRange()
    if (startPoint && endPoint && sourceOffset >= 0) {
      expectedRange.setStart(startPoint.node, startPoint.offset)
      expectedRange.setEnd(endPoint.node, endPoint.offset)
    }
    const expected = startPoint && endPoint
      ? Array.from(expectedRange.getClientRects()).map((rect) => ({ left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom, width: rect.width, height: rect.height }))
      : []
    const actual = rects.map((element) => {
      const rect = element.getBoundingClientRect()
      return { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom, width: rect.width, height: rect.height }
    })
    const deltas = actual.flatMap((rect, index) => {
      const wanted = expected[index]
      return wanted ? [Math.abs(rect.left - wanted.left), Math.abs(rect.top - wanted.top), Math.abs(rect.right - wanted.right), Math.abs(rect.bottom - wanted.bottom)] : [Infinity]
    })
    const firstRectStyle = rects[0] ? getComputedStyle(rects[0]) : null
    const groupStyle = group ? getComputedStyle(group) : null
    const viewer = document.querySelector('.viewer')?.getBoundingClientRect()
    const activeBounds = actual[0]
    const visible = Boolean(viewer && activeBounds
      && activeBounds.right > viewer.left && activeBounds.left < viewer.right
      && activeBounds.bottom > viewer.top && activeBounds.top < viewer.bottom)
    return {
      groups: document.querySelectorAll('[data-search-highlight="true"][data-search-active="true"]').length,
      pageIndex,
      occurrenceIndex: Number(group?.getAttribute('data-occurrence-index')),
      query: group?.getAttribute('data-search-query'),
      rectCount: rects.length,
      expectedCount: expected.length,
      maxEdgeDelta: deltas.length ? Math.max(...deltas) : Infinity,
      background: firstRectStyle?.backgroundColor,
      pointerEvents: groupStyle?.pointerEvents,
      sourceWidth: startSpanIndex >= 0 && endSpanIndex >= startSpanIndex
        ? sourceSpans.slice(startSpanIndex, endSpanIndex + 1).reduce((sum, span) => sum + span.getBoundingClientRect().width, 0)
        : 0,
      highlightWidth: actual.reduce((sum, rect) => sum + rect.width, 0),
      visible,
      actual,
      expected,
    }
  })()`)
}

async function waitFor(expression, label) {
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(40)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`document.querySelectorAll('.continuous-page-slot').length >= 5`, 'continuous document')
await waitFor(`document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]').length >= 3`, 'first page text layer')

await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true, cancelable: true }))`)
await waitFor(`document.activeElement?.matches('input[aria-label="Find in document"]')`, 'Ctrl+F search focus')

await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Find in document"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, ${JSON.stringify(query)})
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
})()`)
await waitFor(`document.querySelectorAll('[data-search-result]').length > 0 && !document.querySelector('.search-status')?.textContent?.includes('Searching')`, 'phrase search results')
await waitFor(`Boolean(document.querySelector('[data-search-highlight="true"][data-search-active="true"][data-page-index="0"] [data-search-match-rect="true"]'))`, 'orange search highlight')

const highlightFirst = await highlightSnapshot()
const screenshot = await call('Page.captureScreenshot', { format: 'png', fromSurface: true })
await fs.mkdir(path.dirname(screenshotPath), { recursive: true })
await fs.writeFile(screenshotPath, Buffer.from(screenshot.data, 'base64'))

const searchBefore = await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Find in document"]')
  return {
    focused: document.activeElement === input,
    query: input.value,
    status: document.querySelector('.search-status')?.textContent,
    resultPages: Array.from(document.querySelectorAll('[data-search-result] span')).map((item) => item.textContent),
    page: document.querySelector('input[aria-label="Current page"]')?.value,
  }
})()`)

await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Find in document"]')
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
})()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '2'`, 'Enter navigates to next search page')
await waitFor(`Boolean(document.querySelector('[data-search-highlight="true"][data-search-active="true"][data-page-index="1"] [data-search-match-rect="true"]'))`, 'next orange search highlight')
const highlightNext = await highlightSnapshot()

await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Find in document"]')
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true, bubbles: true, cancelable: true }))
})()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === '1'`, 'Shift+Enter navigates to previous search page')
await waitFor(`Boolean(document.querySelector('[data-search-highlight="true"][data-search-active="true"][data-page-index="0"] [data-search-match-rect="true"]'))`, 'previous orange search highlight')
const highlightPrevious = await highlightSnapshot()

const search = {
  ...searchBefore,
  navigation: [searchBefore.page, '2', '1'],
}

const closeBounds = await evaluate(`(() => {
  const button = document.querySelector('button[aria-label="Close"]')
  const bounds = button.getBoundingClientRect()
  return { x: bounds.left + bounds.width / 2, y: bounds.top + bounds.height / 2 }
})()`)
await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: closeBounds.x, y: closeBounds.y })
await pause(80)
const closeHover = await evaluate(`(() => {
  const button = document.querySelector('button[aria-label="Close"]')
  const style = getComputedStyle(button)
  const iconStyle = getComputedStyle(button.querySelector('svg'))
  return { hovered: button.matches(':hover'), background: style.backgroundColor, color: style.color, iconColor: iconStyle.color }
})()`)

await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Find in document"]')
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }))
})()`)
await waitFor(`!document.querySelector('.search-pane')`, 'Escape closes search pane')
await waitFor(`!document.querySelector('[data-search-highlight="true"][data-search-active="true"]')`, 'search highlight cleanup')

await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Current page"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, '1')
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
})()`)
await waitFor(`document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]').length >= 3`, 'first page after search navigation')

const selectionBefore = await evaluate(`(() => {
  document.querySelector('button[aria-label="Select text (V)"]')?.click()
  const spans = Array.from(document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]'))
    .filter((span) => span.firstChild instanceof Text && span.firstChild.length)
  const first = spans[0].firstChild
  const last = spans[Math.min(spans.length - 1, 3)].firstChild
  const range = document.createRange()
  range.setStart(first, 0)
  range.setEnd(last, last.length)
  const selection = getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  const rects = Array.from(range.getClientRects())
  return {
    text: selection.toString(),
    length: selection.toString().length,
    rectCount: rects.length,
    endpointsConnected: first.isConnected && last.isConnected,
    pageSurfaceWidth: document.querySelector('.continuous-page-slot[data-page-index="0"] .page-surface')?.getBoundingClientRect().width,
  }
})()`)

await evaluate(`document.querySelector('.continuous-page-slot[data-page-index="4"]')?.scrollIntoView({ block: 'center', behavior: 'auto' })`)
await pause(900)
const selectionAfter = await evaluate(`(() => {
  const selection = getSelection()
  const range = selection?.rangeCount ? selection.getRangeAt(0) : null
  const rects = range ? Array.from(range.getClientRects()) : []
  return {
    text: selection?.toString() || '',
    length: selection?.toString().length || 0,
    rangeCount: selection?.rangeCount || 0,
    collapsed: selection?.isCollapsed ?? true,
    endpointsConnected: Boolean(range?.startContainer?.isConnected && range?.endContainer?.isConnected),
    rectCount: rects.length,
    firstPageMounted: Boolean(document.querySelector('.continuous-page-slot[data-page-index="0"] .text-layer [data-text-item="true"]')),
    pageSurfaceWidth: document.querySelector('.continuous-page-slot[data-page-index="0"] .page-surface')?.getBoundingClientRect().width,
    currentPage: document.querySelector('input[aria-label="Current page"]')?.value,
  }
})()`)

if (!search.focused || !search.resultPages.length) throw new Error(`Search failed: ${JSON.stringify(search)}`)
for (const [label, highlight] of Object.entries({ highlightFirst, highlightNext, highlightPrevious })) {
  if (highlight.groups !== 1 || !highlight.rectCount || highlight.rectCount !== highlight.expectedCount
    || highlight.maxEdgeDelta > 1.25 || !highlight.visible || highlight.pointerEvents !== 'none'
    || highlight.background !== 'rgba(249, 115, 22, 0.44)'
    || highlight.highlightWidth >= highlight.sourceWidth - 8) {
    throw new Error(`${label} failed: ${JSON.stringify(highlight)}`)
  }
}
if (!closeHover.hovered || closeHover.background !== 'rgb(232, 17, 35)' || closeHover.color !== 'rgb(255, 255, 255)') {
  throw new Error(`Close hover styling failed: ${JSON.stringify(closeHover)}`)
}
if (selectionBefore.length < 20 || selectionAfter.text !== selectionBefore.text || selectionAfter.collapsed || !selectionAfter.endpointsConnected || !selectionAfter.firstPageMounted) {
  throw new Error(`Selection was invalidated: ${JSON.stringify({ selectionBefore, selectionAfter })}`)
}
if (selectionBefore.pageSurfaceWidth !== selectionAfter.pageSurfaceWidth) {
  throw new Error(`Automatic fit changed while text was selected: ${selectionBefore.pageSurfaceWidth} -> ${selectionAfter.pageSurfaceWidth}`)
}

await evaluate(`getSelection()?.removeAllRanges()`)
socket.close()
console.log(JSON.stringify({ search, highlightFirst, highlightNext, highlightPrevious, screenshotPath, closeHover, selectionBefore, selectionAfter }))
