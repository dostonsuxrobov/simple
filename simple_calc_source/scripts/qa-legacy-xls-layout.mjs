import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import process from 'node:process'
import { spawn, spawnSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'

const require = createRequire(import.meta.url)
const electronPath = require('electron')
const scriptDirectory = path.dirname(fileURLToPath(import.meta.url))
const root = path.resolve(scriptDirectory, '..')
const fixturePath = path.join(scriptDirectory, 'fixtures', 'schedule_template.xls')
const basePort = 9_500 + (process.pid % 200)
const vitePort = basePort
const debugPort = basePort + 200
const appUrl = `http://127.0.0.1:${vitePort}/`
const profilePath = path.resolve(root, 'tmp', `legacy-xls-layout-profile-${process.pid}`)
const deadline = Date.now() + 60_000
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

let viteOutput = ''
let electronOutput = ''
let vite = null
let electron = null
let client = null

function collectOutput(child, assign) {
  const append = (chunk) => assign(String(chunk))
  child.stdout?.on('data', append)
  child.stderr?.on('data', append)
}

async function waitForUrl(url) {
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url)
      if (response.ok) return
    } catch {}
    await pause(100)
  }
  throw new Error(`Timed out waiting for ${url}.\n${viteOutput}`)
}

async function waitForTarget() {
  while (Date.now() < deadline) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${debugPort}/json/list`)).json()
      const target = targets.find((item) => item.type === 'page' && item.url.startsWith(appUrl))
      if (target?.webSocketDebuggerUrl) return target
    } catch {}
    if (electron?.exitCode !== null) throw new Error(`Electron exited before its page was ready.\n${electronOutput}`)
    await pause(100)
  }
  throw new Error(`Timed out waiting for the simple_calc debug target.\n${electronOutput}`)
}

class CdpClient {
  constructor(url) {
    this.url = url
    this.nextId = 0
    this.pending = new Map()
    this.socket = null
  }

  async connect() {
    this.socket = new WebSocket(this.url)
    await new Promise((resolve, reject) => {
      this.socket.addEventListener('open', resolve, { once: true })
      this.socket.addEventListener('error', reject, { once: true })
    })
    this.socket.addEventListener('message', (event) => {
      const message = JSON.parse(String(event.data))
      const request = this.pending.get(message.id)
      if (!request) return
      this.pending.delete(message.id)
      if (message.error) request.reject(new Error(message.error.message))
      else request.resolve(message.result)
    })
    await this.call('Runtime.enable')
  }

  call(method, params = {}) {
    const id = ++this.nextId
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject })
      this.socket.send(JSON.stringify({ id, method, params }))
    })
  }

  async evaluate(expression) {
    const response = await this.call('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
    })
    if (response.exceptionDetails) {
      throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text || 'Renderer evaluation failed.')
    }
    return response.result.value
  }

  async waitFor(expression, label) {
    while (Date.now() < deadline) {
      try {
        if (await this.evaluate(`Boolean(${expression})`)) return
      } catch {}
      await pause(100)
    }
    throw new Error(`Timed out waiting for ${label}.`)
  }

  close() {
    try { this.socket?.close() } catch {}
  }
}

function forceStop(child) {
  if (!child || child.exitCode !== null) return
  child.kill()
  if (process.platform === 'win32' && child.exitCode === null) {
    spawnSync('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
  }
}

function assertLabelPlacement(label, box, name) {
  assert(label, `${name} label cell should be rendered`)
  assert(box, `${name} entry box should be rendered`)
  assert.equal(label.textAlign, 'left', `${name} should be left-aligned in its source cell`)
  assert.ok(Math.abs(label.cellLeft - box.cellLeft) <= 1, `${name} should share the entry box's left edge`)
  assert.ok(Math.abs(label.cellBottom - box.cellTop) <= 1, `${name} should sit in the row immediately above its entry box`)
  const visibleInset = label.contentLeft - box.cellLeft
  assert.ok(
    visibleInset >= 0 && visibleInset <= 12,
    `${name} text should start beside its entry box, not ${visibleInset.toFixed(1)}px to the right`,
  )
}

async function main() {
  await fs.access(fixturePath)
  await fs.mkdir(profilePath, { recursive: true })

  const viteBin = path.join(root, 'node_modules', 'vite', 'bin', 'vite.js')
  vite = spawn(process.execPath, [viteBin, '--host', '127.0.0.1', '--port', String(vitePort), '--strictPort'], {
    cwd: root,
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  collectOutput(vite, (chunk) => { viteOutput = `${viteOutput}${chunk}`.slice(-8_000) })
  await waitForUrl(appUrl)

  electron = spawn(electronPath, [
    `--remote-debugging-port=${debugPort}`,
    `--user-data-dir=${profilePath}`,
    '.',
    fixturePath,
  ], {
    cwd: root,
    env: { ...process.env, VITE_DEV_SERVER_URL: appUrl },
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  collectOutput(electron, (chunk) => { electronOutput = `${electronOutput}${chunk}`.slice(-8_000) })

  const target = await waitForTarget()
  client = new CdpClient(target.webSocketDebuggerUrl)
  await client.connect()
  await client.waitFor(
    `document.querySelector('.document-title')?.textContent.includes('schedule_template') && document.querySelector('.sheet-viewport')`,
    'the imported legacy schedule',
  )
  await client.evaluate(`(() => {
    const viewport = document.querySelector('.sheet-viewport');
    const targetTop = Math.min(360, Math.max(0, viewport.scrollHeight - viewport.clientHeight));
    viewport.scrollTo({ top: targetTop, left: 0 });
  })()`)
  await client.waitFor(
    `[...document.querySelectorAll('.grid-cell')].some((cell) => cell.id.endsWith('-A34')) && [...document.querySelectorAll('.grid-cell')].some((cell) => cell.id.endsWith('-D37'))`,
    'rows 34 through 37',
  )

  const layout = await client.evaluate(`(() => {
    const metric = (address) => {
      const cell = [...document.querySelectorAll('.grid-cell')].find((item) => item.id.endsWith('-' + address));
      if (!cell) return null;
      const content = cell.querySelector('.cell-content');
      const cellBox = cell.getBoundingClientRect();
      const contentBox = content?.getBoundingClientRect();
      return {
        address,
        text: cell.textContent.trim(),
        textAlign: getComputedStyle(cell).textAlign,
        transform: content ? getComputedStyle(content).transform : 'none',
        cellLeft: cellBox.left,
        cellTop: cellBox.top,
        cellBottom: cellBox.bottom,
        contentLeft: contentBox?.left ?? cellBox.left,
      };
    };
    return {
      company: metric('A34'),
      companyBox: metric('A35'),
      authorization: metric('A36'),
      authorizationBox: metric('A37'),
      assignmentDate: metric('D36'),
      assignmentDateBox: metric('D37'),
    };
  })()`)

  process.stdout.write(`${JSON.stringify(layout, null, 2)}\n`)
  assert.equal(layout.company.text, 'Company Name')
  assert.equal(layout.authorization.text, 'Authorization Signature')
  assert.equal(layout.assignmentDate.text, 'Date of Assignment')
  assertLabelPlacement(layout.company, layout.companyBox, 'Company Name')
  assertLabelPlacement(layout.authorization, layout.authorizationBox, 'Authorization Signature')
  assertLabelPlacement(layout.assignmentDate, layout.assignmentDateBox, 'Date of Assignment')

  process.stdout.write('Legacy XLS lower-form layout QA passed.\n')
}

try {
  await main()
} finally {
  try { await client?.evaluate('window.simpleCalc.close()') } catch {}
  client?.close()
  forceStop(electron)
  forceStop(vite)
  const safeTmpRoot = `${path.resolve(root, 'tmp')}${path.sep}`
  if (profilePath.startsWith(safeTmpRoot)) await fs.rm(profilePath, { recursive: true, force: true }).catch(() => {})
}
