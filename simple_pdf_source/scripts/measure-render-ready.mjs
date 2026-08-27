import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9333)
const startedAt = Number(process.env.SIMPLE_BENCH_STARTED_AT || Date.now())
const timeoutAt = startedAt + 45_000

const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let target
while (!target && Date.now() < timeoutAt) {
  try {
    const response = await fetch(`http://127.0.0.1:${port}/json/list`)
    const targets = await response.json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // Electron has not opened the debugging endpoint yet.
  }
  if (!target) await pause(25)
}

if (!target) throw new Error('Timed out waiting for the Electron renderer.')

const socket = new WebSocket(target.webSocketDebuggerUrl)
await new Promise((resolve, reject) => {
  socket.addEventListener('open', resolve, { once: true })
  socket.addEventListener('error', reject, { once: true })
})

let requestId = 0
const pending = new Map()
socket.addEventListener('message', (event) => {
  const message = JSON.parse(String(event.data))
  if (!message.id || !pending.has(message.id)) return
  const { resolve, reject } = pending.get(message.id)
  pending.delete(message.id)
  if (message.error) reject(new Error(message.error.message))
  else resolve(message.result)
})

function evaluate(expression) {
  const id = ++requestId
  socket.send(JSON.stringify({
    id,
    method: 'Runtime.evaluate',
    params: { expression, returnByValue: true },
  }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}

let state
while (Date.now() < timeoutAt) {
  const response = await evaluate(`(() => {
    const canvas = document.querySelector('.page-canvas')
    const pageNumber = document.querySelector('.status-bar')?.textContent || ''
    return {
      ready: Boolean(canvas && canvas.width > 1 && !document.querySelector('.page-rendering') && !document.querySelector('.document-loading')),
      canvasWidth: canvas?.width || 0,
      canvasHeight: canvas?.height || 0,
      pageNumber,
      bodyText: document.body?.innerText?.slice(0, 160) || '',
    }
  })()`)
  state = response.result?.value
  if (state?.ready) break
  await pause(25)
}

socket.close()

if (!state?.ready) throw new Error(`Timed out waiting for the first rendered page: ${JSON.stringify(state)}`)

console.log(JSON.stringify({
  readyMs: Date.now() - startedAt,
  ...state,
}))
