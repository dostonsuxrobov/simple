import { spawn, spawnSync } from 'node:child_process'
import { mkdtemp, rm } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { performance } from 'node:perf_hooks'

function parseArguments(argv) {
  const options = {
    exe: path.resolve('release/win-unpacked/simple.exe'),
    pdf: '',
    iterations: 3,
    port: 9320,
    timeout: 90_000,
  }
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index]
    const value = argv[index + 1]
    if (argument === '--exe' && value) options.exe = path.resolve(value), index += 1
    else if (argument === '--pdf' && value) options.pdf = path.resolve(value), index += 1
    else if (argument === '--iterations' && value) options.iterations = Math.max(1, Number(value)), index += 1
    else if (argument === '--port' && value) options.port = Number(value), index += 1
    else if (argument === '--timeout' && value) options.timeout = Math.max(1_000, Number(value)), index += 1
    else if (argument === '--help') {
      console.log('Usage: node scripts/benchmark-startup.mjs [--exe path] [--pdf path] [--iterations n] [--port n] [--timeout ms]')
      process.exit(0)
    }
  }
  return options
}

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

async function waitForTargets(port, deadline) {
  while (performance.now() < deadline) {
    try {
      const response = await fetch(`http://127.0.0.1:${port}/json/list`)
      if (response.ok) {
        const targets = await response.json()
        const page = targets.find((target) => target.type === 'page' && target.webSocketDebuggerUrl)
        if (page) return page
      }
    } catch {
      // Chromium has not opened its debugging endpoint yet.
    }
    await delay(15)
  }
  throw new Error('Timed out waiting for the Electron renderer.')
}

async function connectToTarget(url) {
  const socket = new WebSocket(url)
  const pending = new Map()
  let nextId = 1
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data)
    if (!message.id || !pending.has(message.id)) return
    const { resolve, reject } = pending.get(message.id)
    pending.delete(message.id)
    if (message.error) reject(new Error(message.error.message))
    else resolve(message.result)
  })
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  return {
    call(method, params = {}) {
      const id = nextId++
      socket.send(JSON.stringify({ id, method, params }))
      return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
    },
    close() {
      socket.close()
    },
  }
}

async function inspectPage(client) {
  const expression = `(() => {
    const canvas = document.querySelector('.page-canvas')
    const textLayer = document.querySelector('.text-layer')
    return {
      shell: Boolean(document.querySelector('.app-root')),
      documentMounted: Boolean(document.querySelector('.document-app')),
      pagePrepared: Boolean(document.querySelector('.page-surface')),
      firstPagePainted: Boolean(canvas && canvas.width > 1 && canvas.height > 1 && !document.querySelector('.page-rendering')),
      textLayerReady: Boolean(textLayer && (textLayer.children.length > 0 || !document.querySelector('.page-rendering'))),
      documentError: document.querySelector('.document-error')?.textContent?.trim() || '',
    }
  })()`
  const result = await client.call('Runtime.evaluate', { expression, returnByValue: true })
  return result.result.value
}

async function stopApp(client, child) {
  try {
    await client?.call('Runtime.evaluate', { expression: 'window.simple?.close?.()' })
  } catch {
    // The renderer may already have closed.
  }
  const exited = await Promise.race([
    new Promise((resolve) => child.once('exit', () => resolve(true))),
    delay(2_500).then(() => false),
  ])
  if (!exited) spawnSync('taskkill.exe', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
  client?.close()
}

async function runIteration(options, iteration) {
  const profileDirectory = await mkdtemp(path.join(os.tmpdir(), 'simple-benchmark-profile-'))
  const port = options.port + iteration
  const arguments_ = [
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${profileDirectory}`,
  ]
  if (options.pdf) arguments_.push(options.pdf)
  const startedAt = performance.now()
  const child = spawn(options.exe, arguments_, {
    stdio: 'ignore',
    windowsHide: true,
  })
  const deadline = startedAt + options.timeout
  let client
  const metrics = {
    iteration: iteration + 1,
    rendererEndpointMs: null,
    shellReadyMs: null,
    documentMountedMs: null,
    pagePreparedMs: null,
    firstPagePaintedMs: null,
    textLayerReadyMs: null,
  }
  try {
    const target = await waitForTargets(port, deadline)
    metrics.rendererEndpointMs = performance.now() - startedAt
    client = await connectToTarget(target.webSocketDebuggerUrl)
    while (performance.now() < deadline) {
      const state = await inspectPage(client)
      const elapsed = performance.now() - startedAt
      if (state.shell && metrics.shellReadyMs === null) metrics.shellReadyMs = elapsed
      if (state.documentMounted && metrics.documentMountedMs === null) metrics.documentMountedMs = elapsed
      if (state.pagePrepared && metrics.pagePreparedMs === null) metrics.pagePreparedMs = elapsed
      if (state.firstPagePainted && metrics.firstPagePaintedMs === null) metrics.firstPagePaintedMs = elapsed
      if (state.textLayerReady && metrics.textLayerReadyMs === null) metrics.textLayerReadyMs = elapsed
      if (state.documentError) throw new Error(state.documentError)
      const complete = options.pdf ? metrics.firstPagePaintedMs !== null : metrics.shellReadyMs !== null
      if (complete) break
      await delay(15)
    }
    const complete = options.pdf ? metrics.firstPagePaintedMs !== null : metrics.shellReadyMs !== null
    if (!complete) throw new Error('Timed out waiting for the requested ready state.')
    return metrics
  } finally {
    await stopApp(client, child)
    await rm(profileDirectory, { recursive: true, force: true })
    await delay(600)
  }
}

function rounded(value) {
  return typeof value === 'number' ? Math.round(value) : null
}

function median(values) {
  const ordered = values.filter((value) => typeof value === 'number').sort((a, b) => a - b)
  if (!ordered.length) return null
  const middle = Math.floor(ordered.length / 2)
  return ordered.length % 2 ? ordered[middle] : (ordered[middle - 1] + ordered[middle]) / 2
}

const options = parseArguments(process.argv.slice(2))
const results = []
for (let iteration = 0; iteration < options.iterations; iteration += 1) {
  results.push(await runIteration(options, iteration))
}

const displayRows = results.map((result) => Object.fromEntries(
  Object.entries(result).map(([key, value]) => [key, rounded(value)]),
))
console.table(displayRows)
const timingKeys = Object.keys(results[0]).filter((key) => key.endsWith('Ms'))
console.log(JSON.stringify({
  executable: options.exe,
  pdf: options.pdf || null,
  iterations: options.iterations,
  medianMs: Object.fromEntries(timingKeys.map((key) => [key, rounded(median(results.map((result) => result[key])))])),
  samples: displayRows,
}, null, 2))
