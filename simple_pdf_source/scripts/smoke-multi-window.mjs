import process from 'node:process'

const port = Number(process.env.SIMPLE_BENCH_PORT || 9335)
const filePath = process.env.SIMPLE_SECOND_FILE
if (!filePath) throw new Error('Set SIMPLE_SECOND_FILE.')
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))
const deadline = Date.now() + 30_000

let target
while (!target && Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
  } catch {
    // The renderer is still starting.
  }
  if (!target) await pause(40)
}
if (!target) throw new Error('Renderer did not start.')

const socket = new WebSocket(target.webSocketDebuggerUrl)
await new Promise((resolve, reject) => {
  socket.addEventListener('open', resolve, { once: true })
  socket.addEventListener('error', reject, { once: true })
})
const response = new Promise((resolve, reject) => {
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data))
    if (message.id !== 1) return
    if (message.error || message.result?.exceptionDetails) reject(new Error(message.error?.message || message.result.exceptionDetails.text))
    else resolve(message.result?.result?.value)
  })
})
socket.send(JSON.stringify({
  id: 1,
  method: 'Runtime.evaluate',
  params: {
    expression: `window.simple.openInNewWindow(${JSON.stringify(filePath)})`,
    awaitPromise: true,
    returnByValue: true,
  },
}))
if (!await response) throw new Error('The new-window request was rejected.')

let pages = []
while (Date.now() < deadline) {
  pages = (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).filter((item) => item.type === 'page')
  if (pages.length >= 2) break
  await pause(40)
}
socket.close()
if (pages.length < 2) throw new Error(`Expected two PDF windows, found ${pages.length}.`)
console.log(JSON.stringify({ windowCount: pages.length, titles: pages.map((page) => page.title) }))
