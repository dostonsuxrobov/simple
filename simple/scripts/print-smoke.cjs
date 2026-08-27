'use strict'

const assert = require('node:assert/strict')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')

function argument(name) {
  const prefix = `--${name}=`
  return process.argv.find((value) => value.startsWith(prefix))?.slice(prefix.length) || null
}

const executable = path.resolve(argument('exe') || '')
const imagePath = path.resolve(argument('image') || '')
const port = Number(argument('port') || 9599)
assert.ok(executable && imagePath)

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

async function target() {
  const started = Date.now()
  while (Date.now() - started < 60_000) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      const page = targets.find((item) => item.type === 'page' && item.url.includes('/modules/image/dist/index.html'))
      if (page) return page
    } catch {}
    await delay(150)
  }
  throw new Error('The packaged image workspace did not start.')
}

async function main() {
  const child = spawn(executable, [`--remote-debugging-port=${port}`, imagePath], { stdio: ['ignore', 'ignore', 'ignore'], windowsHide: true })
  let socket
  let nextId = 0
  const pending = new Map()
  try {
    const page = await target()
    socket = new WebSocket(page.webSocketDebuggerUrl)
    await new Promise((resolve, reject) => {
      socket.addEventListener('open', resolve, { once: true })
      socket.addEventListener('error', reject, { once: true })
    })
    socket.addEventListener('message', (event) => {
      const message = JSON.parse(String(event.data))
      if (!pending.has(message.id)) return
      const callback = pending.get(message.id)
      pending.delete(message.id)
      callback(message)
    })
    const call = (method, params = {}) => new Promise((resolve, reject) => {
      const id = ++nextId
      pending.set(id, (message) => message.error ? reject(new Error(message.error.message)) : resolve(message.result))
      socket.send(JSON.stringify({ id, method, params }))
    })
    const evaluate = async (expression) => {
      const response = await call('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true })
      if (response.exceptionDetails) throw new Error(response.exceptionDetails.text)
      return response.result.value
    }
    const waitFor = async (expression, timeout = 20_000) => {
      const started = Date.now()
      while (Date.now() - started < timeout) {
        if (await evaluate(expression).catch(() => false)) return
        await delay(100)
      }
      throw new Error(`Timed out: ${expression}`)
    }
    await waitFor(`document.querySelector('canvas')?.width > 0`)
    await waitFor(`document.querySelector('button[title^="Print"]') && !document.querySelector('button[title^="Print"]').disabled`)
    await evaluate(`document.querySelector('button[title^="Print"]').click()`)
    let status = { printing: false, toast: '' }
    for (let attempt = 0; attempt < 30; attempt += 1) {
      status = await evaluate(`({ printing: document.body.innerText.includes('Printing…'), toast: document.querySelector('.toast')?.textContent || '' })`)
      if (status.printing || /cancel|print|Windows/i.test(status.toast)) break
      await delay(100)
    }
    if (status.printing) {
      process.stdout.write('Native print request opened.\n')
    } else if (/cancel/i.test(status.toast)) {
      process.stdout.write('Native print request reached Windows and was canceled.\n')
    } else {
      throw new Error(`Printing did not reach the native handoff: ${status.toast || 'no status'}`)
    }
    await evaluate('window.simpleImage.confirmClose()').catch(() => {})
    await delay(800)
  } finally {
    try { socket?.close() } catch {}
    if (child.exitCode === null) spawnSync('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
  }
}

main().catch((error) => {
  console.error(error.stack || error)
  process.exitCode = 1
})
