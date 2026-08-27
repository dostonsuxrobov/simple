import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9386)
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
  if (!target) await pause(35)
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

async function waitFor(expression, label) {
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(40)
  }
  throw new Error(`Timed out: ${label}`)
}

function selectionSnapshot() {
  return evaluate(`(() => {
    const selection = getSelection()
    const focusSpan = selection?.focusNode?.parentElement?.closest?.('[data-text-item="true"]')
    const anchorSpan = selection?.anchorNode?.parentElement?.closest?.('[data-text-item="true"]')
    const range = selection?.rangeCount ? selection.getRangeAt(0) : null
    return {
      text: selection?.toString() || '',
      collapsed: selection?.isCollapsed ?? true,
      anchorId: anchorSpan?.dataset.textId || '',
      anchorOffset: selection?.anchorOffset ?? -1,
      focusId: focusSpan?.dataset.textId || '',
      focusOffset: selection?.focusOffset ?? -1,
      rectCount: range ? range.getClientRects().length : 0,
      pointerDriver: document.querySelector('.viewer')?.dataset.pointerSelectingText || '',
      scrollTop: document.querySelector('.viewer')?.scrollTop || 0,
    }
  })()`)
}

async function mouse(type, point, pressed = false) {
  await call('Input.dispatchMouseEvent', {
    type,
    x: point.x,
    y: point.y,
    button: 'left',
    buttons: pressed ? 1 : 0,
    clickCount: 1,
  })
}

await waitFor(`document.querySelectorAll('.continuous-page-slot').length === 8`, 'continuous fixture')
await evaluate(`document.querySelector('button[aria-label="Select text (V)"]')?.click()`)
await evaluate(`document.querySelector('.continuous-page-slot[data-page-index="0"]')?.scrollIntoView({ block: 'center', behavior: 'auto' })`)
await waitFor(`document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]').length >= 4`, 'page-one text layer')
await pause(250)

const points = await evaluate(`(() => {
  const spans = [...document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]')]
  const first = spans.find((span) => span.textContent?.startsWith('Select only these four words'))
  const second = spans.find((span) => span.textContent?.startsWith('Inline style-preserving'))
  if (!first || !second || !(first.firstChild instanceof Text) || !(second.firstChild instanceof Text)) return null
  const pointAtOffset = (span, offset, side) => {
    const node = span.firstChild
    const safe = Math.max(0, Math.min(node.length - 1, side === 'start' ? offset : offset - 1))
    const range = document.createRange()
    range.setStart(node, safe)
    range.setEnd(node, safe + 1)
    const rect = range.getBoundingClientRect()
    return {
      x: side === 'start' ? rect.left + Math.min(0.35, rect.width * 0.15) : rect.right - Math.min(0.35, rect.width * 0.15),
      y: rect.top + rect.height / 2,
    }
  }
  const firstRect = first.getBoundingClientRect()
  const secondRect = second.getBoundingClientRect()
  const gap = { x: firstRect.left + firstRect.width * 0.55, y: (firstRect.bottom + secondRect.top) / 2 }
  return {
    firstId: first.dataset.textId,
    secondId: second.dataset.textId,
    start: pointAtOffset(first, first.textContent.indexOf('only'), 'start'),
    firstEnd: pointAtOffset(first, first.textContent.indexOf(','), 'end'),
    secondEnd: pointAtOffset(second, Math.min(18, second.textContent.length), 'end'),
    gap,
    gapHit: document.elementsFromPoint(gap.x, gap.y).some((element) => element.closest?.('[data-text-item="true"]')),
  }
})()`)
if (!points || points.gapHit) throw new Error(`Could not find deterministic text/gap points: ${JSON.stringify(points)}`)

await mouse('mouseMoved', points.start)
await mouse('mousePressed', points.start, true)
await mouse('mouseMoved', points.firstEnd, true)
await pause(80)
const firstLine = await selectionSnapshot()
if (firstLine.pointerDriver !== 'true' || firstLine.collapsed || !firstLine.text.includes('only these four words')) {
  throw new Error(`Pointer selection did not follow the first line: ${JSON.stringify(firstLine)}`)
}

// The pointer is over no glyph here. The selection endpoint must remain on the
// last hovered character instead of Chromium guessing an adjacent PDF line.
await mouse('mouseMoved', points.gap, true)
await pause(80)
const gap = await selectionSnapshot()
if (gap.text !== firstLine.text || gap.focusId !== firstLine.focusId || gap.focusOffset !== firstLine.focusOffset
  || Math.abs(gap.scrollTop - firstLine.scrollTop) > 1) {
  throw new Error(`Selection jumped while the pointer crossed whitespace: ${JSON.stringify({ firstLine, gap })}`)
}

await mouse('mouseMoved', points.secondEnd, true)
await pause(80)
const secondLine = await selectionSnapshot()
if (secondLine.focusId !== points.secondId || secondLine.collapsed || !secondLine.text.includes('Inline style')) {
  throw new Error(`Selection did not advance to the hovered second line: ${JSON.stringify(secondLine)}`)
}
await mouse('mouseReleased', points.secondEnd)
await pause(80)
const released = await selectionSnapshot()
if (released.pointerDriver || released.text !== secondLine.text || released.collapsed) {
  throw new Error(`Selection changed on pointer release: ${JSON.stringify({ secondLine, released })}`)
}

await evaluate(`getSelection()?.removeAllRanges()`)
await evaluate(`document.querySelector('button[aria-label="Highlight text"]')?.click()`)
await mouse('mouseMoved', points.start)
await mouse('mousePressed', points.start, true)
await mouse('mouseMoved', points.firstEnd, true)
await mouse('mouseReleased', points.firstEnd)
await waitFor(`Boolean(document.querySelector('.continuous-page-slot[data-page-index="0"] .highlight-overlay'))`, 'pointer-created highlight')
const highlight = await evaluate(`(() => {
  const overlay = document.querySelector('.continuous-page-slot[data-page-index="0"] .highlight-overlay')
  const line = [...document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]')]
    .find((span) => span.textContent?.startsWith('Select only these four words'))
  const rect = overlay?.getBoundingClientRect()
  const source = line?.getBoundingClientRect()
  return {
    exists: Boolean(overlay),
    width: rect?.width || 0,
    sourceWidth: source?.width || 0,
    selectionCleared: getSelection()?.isCollapsed ?? true,
  }
})()`)
if (!highlight.exists || !highlight.selectionCleared || highlight.width <= 20 || highlight.width >= highlight.sourceWidth - 20) {
  throw new Error(`Focused pointer highlight failed: ${JSON.stringify(highlight)}`)
}
await evaluate(`document.querySelector('button[aria-label="Undo (Ctrl+Z)"]')?.click()`)
await waitFor(`!document.querySelector('.continuous-page-slot[data-page-index="0"] .highlight-overlay')`, 'highlight undo')
await evaluate(`document.querySelector('button[aria-label="Select text (V)"]')?.click()`)

await evaluate(`document.querySelector('button[aria-label="Rotate page 1 right"]')?.click()`)
await waitFor(`Math.abs(Number(document.querySelector('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]')?.dataset.textAngle || 0)) > 1`, 'rotated text layer')
await pause(180)

const rotated = await evaluate(`(() => {
  const surface = document.querySelector('.continuous-page-slot[data-page-index="0"] .page-surface')
  if (!surface) return null
  const page = surface.getBoundingClientRect()
  const spans = [...surface.querySelectorAll('[data-text-item="true"]')]
  const metrics = spans.map((span) => {
    const rect = span.getBoundingClientRect()
    return {
      id: span.dataset.textId,
      text: span.textContent,
      width: rect.width,
      height: rect.height,
      scaleX: Number(span.dataset.textScaleX),
      inside: rect.left >= page.left - 1.5 && rect.top >= page.top - 1.5
        && rect.right <= page.right + 1.5 && rect.bottom <= page.bottom + 1.5,
    }
  })
  return {
    page: { width: page.width, height: page.height },
    count: spans.length,
    maxWidth: Math.max(...metrics.map((item) => item.width)),
    maxHeight: Math.max(...metrics.map((item) => item.height)),
    maxScaleX: Math.max(...metrics.map((item) => item.scaleX)),
    outside: metrics.filter((item) => !item.inside),
  }
})()`)
if (!rotated?.count || rotated.maxWidth > rotated.page.width + 2 || rotated.maxHeight > rotated.page.height + 2
  || rotated.maxScaleX > 4 || rotated.outside.length) {
  throw new Error(`Rotated text hit boxes escaped the page: ${JSON.stringify(rotated)}`)
}

// Repeat the focused phrase drag after rotation; the caret hit-test must track
// the vertical glyph run without selecting neighbouring text.
const rotatedPoints = await evaluate(`(() => {
  const span = [...document.querySelectorAll('.continuous-page-slot[data-page-index="0"] [data-text-item="true"]')]
    .find((item) => item.textContent?.startsWith('Select only these four words'))
  if (!span || !(span.firstChild instanceof Text)) return null
  const node = span.firstChild
  const point = (offset, side) => {
    const safe = Math.max(0, Math.min(node.length - 1, side === 'start' ? offset : offset - 1))
    const range = document.createRange()
    range.setStart(node, safe)
    range.setEnd(node, safe + 1)
    const rect = range.getBoundingClientRect()
    return { x: rect.left + rect.width / 2, y: side === 'start' ? rect.top + 0.2 : rect.bottom - 0.2 }
  }
  return { start: point(span.textContent.indexOf('only'), 'start'), end: point(span.textContent.indexOf(','), 'end') }
})()`)
if (!rotatedPoints) throw new Error('Could not resolve rotated selection points.')
await mouse('mouseMoved', rotatedPoints.start)
await mouse('mousePressed', rotatedPoints.start, true)
await mouse('mouseMoved', rotatedPoints.end, true)
await pause(80)
const rotatedSelection = await selectionSnapshot()
await mouse('mouseReleased', rotatedPoints.end)
if (rotatedSelection.collapsed || !rotatedSelection.text.includes('only these four words')) {
  throw new Error(`Rotated pointer selection failed: ${JSON.stringify(rotatedSelection)}`)
}

await evaluate(`getSelection()?.removeAllRanges()`)
await evaluate(`document.querySelector('button[aria-label="Rotate page 1 left"]')?.click()`)
socket.close()
console.log(JSON.stringify({ firstLine, gap, secondLine, released, highlight, rotated, rotatedSelection }))
