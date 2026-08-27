import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9390)
const targetPage = Math.max(1, Number(process.env.SIMPLE_TARGET_PAGE || 1))
const sourceText = process.env.SIMPLE_SOURCE_TEXT || 'Machine intelligence'
const replacementText = process.env.SIMPLE_REPLACEMENT_TEXT || 'Machine reasoning'
const deadline = Date.now() + 120_000
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

function evaluate(expression) {
  const id = ++requestId
  socket.send(JSON.stringify({
    id,
    method: 'Runtime.evaluate',
    params: { expression, returnByValue: true, awaitPromise: true },
  }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

async function waitFor(expression, label) {
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return
    await pause(40)
  }
  throw new Error(`Timed out: ${label}`)
}

await waitFor(`Boolean(document.querySelector('.document-app'))`, 'document shell')
await evaluate(`(() => {
  const input = document.querySelector('input[aria-label="Current page"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set
  setter.call(input, ${JSON.stringify(String(targetPage))})
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }))
})()`)
await waitFor(`document.querySelector('input[aria-label="Current page"]')?.value === ${JSON.stringify(String(targetPage))}`, 'target page')
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((span) => span.textContent?.includes(${JSON.stringify(sourceText)}))`, 'source text')

await evaluate(`document.querySelector('button[aria-label="Edit text (E)"]')?.click()`)
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((span) => span.textContent?.includes(${JSON.stringify(sourceText)}))`, 'text layer after edit-mode fit')

const source = await evaluate(`(() => {
  const span = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]'))
    .find((item) => item.textContent?.includes(${JSON.stringify(sourceText)}))
  const start = span.textContent.indexOf(${JSON.stringify(sourceText)})
  const range = document.createRange()
  range.setStart(span.firstChild, start)
  range.setEnd(span.firstChild, start + ${JSON.stringify(sourceText)}.length)
  const bounds = range.getBoundingClientRect()
  const itemBounds = span.getBoundingClientRect()
  const suffixRange = document.createRange()
  suffixRange.setStart(span.firstChild, start + ${JSON.stringify(sourceText)}.length)
  suffixRange.setEnd(span.firstChild, span.firstChild.length)
  const suffixBounds = suffixRange.getBoundingClientRect()
  const selection = getSelection()
  selection.removeAllRanges()
  selection.addRange(range)
  const computed = getComputedStyle(span)
  const data = { ...span.dataset }
  span.dispatchEvent(new MouseEvent('click', {
    bubbles: true,
    cancelable: true,
    clientX: bounds.left + bounds.width / 2,
    clientY: bounds.top + bounds.height / 2,
  }))
  return {
    text: selection.toString(),
    itemText: span.textContent,
    selectionStart: start,
    selectionEnd: start + ${JSON.stringify(sourceText)}.length,
    bounds: { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height },
    itemBounds: { left: itemBounds.left, top: itemBounds.top, width: itemBounds.width, height: itemBounds.height },
    suffixBounds: { left: suffixBounds.left, top: suffixBounds.top, width: suffixBounds.width, height: suffixBounds.height },
    style: {
      family: computed.fontFamily,
      size: computed.fontSize,
      weight: computed.fontWeight,
      fontStyle: computed.fontStyle,
      letterSpacing: computed.letterSpacing,
      direction: computed.direction,
    },
    data,
  }
})()`)

const committedText = `${source.itemText.slice(0, source.selectionStart)}${replacementText}${source.itemText.slice(source.selectionEnd)}`

await waitFor(`document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')?.value === ${JSON.stringify(sourceText)}`, 'inline editor')
const editor = await evaluate(`(() => {
  const input = document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')
  const style = getComputedStyle(input)
  const frame = input.closest('.inline-text-frame').getBoundingClientRect()
  return {
    value: input.value,
    bounds: { left: frame.left, top: frame.top, width: frame.width, height: frame.height },
    style: {
      family: style.fontFamily,
      size: style.fontSize,
      weight: style.fontWeight,
      fontStyle: style.fontStyle,
      lineHeight: style.lineHeight,
      letterSpacing: style.letterSpacing,
      color: style.color,
      background: style.backgroundColor,
      transform: style.transform,
    },
  }
})()`)

await evaluate(`(() => {
  const input = document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')
  const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value').set
  setter.call(input, ${JSON.stringify(replacementText)})
  input.dispatchEvent(new Event('input', { bubbles: true }))
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', ctrlKey: true, bubbles: true, cancelable: true }))
})()`)
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current .text-overlay')).some((item) => item.textContent === ${JSON.stringify(committedText)})`, 'committed full-item reflow')

const committed = await evaluate(`(() => {
  const overlay = Array.from(document.querySelectorAll('.continuous-page-slot.is-current .text-overlay'))
    .find((item) => item.textContent === ${JSON.stringify(committedText)})
  const style = getComputedStyle(overlay)
  const bounds = overlay.getBoundingClientRect()
  const cover = overlay.parentElement.querySelector('.text-original-cover')?.getBoundingClientRect()
  return {
    text: overlay.textContent,
    bounds: { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height },
    cover: cover ? { left: cover.left, top: cover.top, width: cover.width, height: cover.height } : null,
    style: {
      family: style.fontFamily,
      size: style.fontSize,
      weight: style.fontWeight,
      fontStyle: style.fontStyle,
      lineHeight: style.lineHeight,
      letterSpacing: style.letterSpacing,
      color: style.color,
      background: style.backgroundColor,
      transform: style.transform,
    },
  }
})()`)

await evaluate(`document.querySelector('.save-button')?.click()`)
await waitFor(`!document.querySelector('.busy-overlay') && !document.querySelector('.dirty-dot')`, 'saved edit')
const replacementNeedle = replacementText.split(/\s+/).find((word) => !sourceText.includes(word)) || replacementText.split(/\s+/).at(-1) || replacementText
await waitFor(`Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]')).some((span) => span.textContent?.includes(${JSON.stringify(replacementNeedle)}))`, 'saved replacement text layer')

const saved = await evaluate(`(() => {
  const baseline = Number(${JSON.stringify(source.data.pdfBaselineY)})
  const baselineItems = Array.from(document.querySelectorAll('.continuous-page-slot.is-current [data-text-item="true"]'))
    .filter((item) => Math.abs(Number(item.dataset.pdfBaselineY) - baseline) < 0.02)
  const needleItem = baselineItems.find((item) => item.textContent?.includes(${JSON.stringify(replacementNeedle)}))
  const runItems = baselineItems.filter((item) => item.dataset.pdfFontName === needleItem?.dataset.pdfFontName)
  const suffixItem = runItems.find((item) => item.textContent?.trim() === 'part')
  const needleBounds = needleItem?.getBoundingClientRect()
  const suffixBounds = suffixItem?.getBoundingClientRect()
  return {
    text: runItems.map((item) => item.textContent).join(''),
    items: runItems.map((item) => ({
      text: item.textContent,
      bounds: (() => { const bounds = item.getBoundingClientRect(); return { left: bounds.left, top: bounds.top, width: bounds.width, height: bounds.height } })(),
      data: { ...item.dataset },
    })),
    suffixGap: needleBounds && suffixBounds ? suffixBounds.left - needleBounds.right : null,
  }
})()`)

if (source.text !== sourceText) throw new Error(`Wrong source selection: ${JSON.stringify(source)}`)
if (editor.style.weight !== source.style.weight || editor.style.fontStyle !== source.style.fontStyle) {
  throw new Error(`Editor style did not inherit source style: ${JSON.stringify({ source, editor })}`)
}
if (editor.value !== sourceText) throw new Error(`Editor did not remain selection-only: ${JSON.stringify(editor)}`)
if (committed.text !== committedText) throw new Error(`Commit did not reflow the complete source item: ${JSON.stringify({ committedText, committed })}`)
if (!committed.cover || Math.abs(committed.cover.width - source.itemBounds.width) > 1) {
  throw new Error(`Commit did not cover the complete source item: ${JSON.stringify({ source, committed })}`)
}
if (saved.text !== committedText) throw new Error(`Replacement and suffix were not saved as one reflowed run: ${JSON.stringify({ committedText, saved })}`)
if (!Number.isFinite(saved.suffixGap) || saved.suffixGap < -0.5 || saved.suffixGap > Number.parseFloat(editor.style.size)) {
  throw new Error(`Saved suffix did not move naturally after the replacement: ${JSON.stringify(saved)}`)
}

socket.close()
console.log(JSON.stringify({ targetPage, source, editor, committed, saved }))
