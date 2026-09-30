const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')
const { fileURLToPath } = require('node:url')

const executable = path.resolve(process.argv[2] || path.join(__dirname, '..', 'release', 'win-unpacked', 'Simple Image.exe'))
const fixture = path.resolve(process.argv[3] || path.join(__dirname, '..', 'public', 'brand-icon.png'))
const port = Number(process.env.SIMPLE_IMAGE_DEBUG_PORT || 9647)
assert.ok(fs.statSync(executable).isFile(), `Packaged executable not found: ${executable}`)
assert.ok(fs.statSync(fixture).isFile(), `Image fixture not found: ${fixture}`)

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

async function pageTargets() {
  return (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).filter((target) => target.type === 'page')
}

async function waitForTarget(predicate, timeout = 45_000) {
  const started = Date.now()
  while (Date.now() - started < timeout) {
    try {
      const target = (await pageTargets()).find(predicate)
      if (target?.webSocketDebuggerUrl) return target
    } catch {}
    await delay(120)
  }
  throw new Error('The expected packaged browser surface did not appear.')
}

class CdpClient {
  constructor(url) {
    this.url = url
    this.socket = null
    this.nextId = 0
    this.pending = new Map()
  }

  async connect() {
    this.socket = new WebSocket(this.url)
    await new Promise((resolve, reject) => {
      this.socket.addEventListener('open', resolve, { once: true })
      this.socket.addEventListener('error', reject, { once: true })
    })
    this.socket.addEventListener('message', (event) => {
      const message = JSON.parse(String(event.data))
      if (!message.id || !this.pending.has(message.id)) return
      const pending = this.pending.get(message.id)
      this.pending.delete(message.id)
      if (message.error) pending.reject(new Error(message.error.message))
      else pending.resolve(message.result)
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
    const result = await this.call('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true })
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text)
    return result.result.value
  }

  async waitFor(expression, timeout = 30_000) {
    const started = Date.now()
    while (Date.now() - started < timeout) {
      try { if (await this.evaluate(expression)) return } catch {}
      await delay(100)
    }
    throw new Error(`Timed out waiting for: ${expression}`)
  }

  close() {
    try { this.socket?.close() } catch {}
  }
}

function ownedPrintDirectory(url) {
  try {
    const directory = path.resolve(path.dirname(fileURLToPath(url)))
    return path.dirname(directory) === path.resolve(os.tmpdir()) && path.basename(directory).startsWith('simple-image-print-') ? directory : null
  } catch {
    return null
  }
}

async function main() {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-image-packaged-print-profile-'))
  const child = spawn(executable, [`--remote-debugging-port=${port}`, `--user-data-dir=${profile}`, fixture], {
    env: { ...process.env, APPDATA: profile },
    stdio: 'ignore',
    windowsHide: true,
  })
  let workspaceClient = null
  let printDirectory = null
  try {
    const workspace = await waitForTarget((target) => /\/dist\/index\.html(?:$|[?#])/i.test(target.url))
    workspaceClient = new CdpClient(workspace.webSocketDebuggerUrl)
    await workspaceClient.connect()
    await workspaceClient.waitFor(`document.querySelector('.save-state')?.textContent === 'Saved' && !document.querySelector('button[title="Print (Ctrl+P)"]')?.disabled`)
    await workspaceClient.evaluate(`localStorage.removeItem('simple-image:print-settings:v2')`)
    await workspaceClient.evaluate(`document.querySelector('button[title="Print (Ctrl+P)"]').click()`)
    try {
      await workspaceClient.waitFor(`Boolean(document.querySelector('.image-print-dialog') && document.querySelector('.image-print-controls') && document.querySelector('.image-print-preview'))`)
    } catch (error) {
      const state = await workspaceClient.evaluate(`({
        title: document.title,
        printButton: document.querySelector('button[title="Print (Ctrl+P)"]')?.outerHTML,
        toast: document.querySelector('.toast')?.textContent,
        body: document.body.innerText.slice(0, 1200),
      })`)
      throw new Error(`${error.message}\nPackaged renderer state: ${JSON.stringify(state)}`)
    }
    const dialog = await workspaceClient.evaluate(`(() => {
      const controls = document.querySelector('.image-print-controls').getBoundingClientRect()
      const preview = document.querySelector('.image-print-preview').getBoundingClientRect()
      return { sideBySide: controls.right <= preview.left + 1, title: document.querySelector('#image-print-title').textContent }
    })()`)
    await workspaceClient.evaluate(`(() => {
      const choose = (setting, value) => {
        const element = document.querySelector('[data-print-setting="' + setting + '"]')
        element.value = value
        element.dispatchEvent(new Event('change', { bubbles: true }))
      }
      choose('paper', 'a4')
      choose('orientation', 'landscape')
      choose('margin-preset', '0')
      choose('scale-mode', 'fill')
    })()`)
    await workspaceClient.waitFor(`document.querySelector('.print-preview-sheet')?.dataset.orientation === 'landscape' && document.querySelector('[data-print-setting="scale-mode"]')?.value === 'fill'`)
    const clicked = await workspaceClient.evaluate(`(() => { const button = document.querySelector('.print-submit'); if (!button || button.disabled) return false; button.click(); return true })()`)
    assert.equal(clicked, true, 'The packaged Print button was unavailable.')
    await delay(250)
    const immediateHandoff = await workspaceClient.evaluate(`({
      busy: document.querySelector('.print-submit')?.disabled,
      submitText: document.querySelector('.print-submit')?.textContent,
      toast: document.querySelector('.toast')?.textContent,
      stored: localStorage.getItem('simple-image:print-settings:v2'),
      optionsFrozen: Array.from(document.querySelectorAll('[data-print-setting], .print-background-control button')).every((control) => control.matches(':disabled')),
    })`)
    if (immediateHandoff.toast && /could not|invalid|failed|error|cannot|not found/i.test(immediateHandoff.toast)) {
      throw new Error(`Packaged print handoff failed: ${JSON.stringify(immediateHandoff)}`)
    }
    let printSurface
    try {
      printSurface = await waitForTarget((target) => target.id !== workspace.id && /simple-image-print-.*print\.html/i.test(target.url), 15_000)
    } catch (error) {
      const state = await workspaceClient.evaluate(`({
        dialogOpen: Boolean(document.querySelector('.image-print-dialog')),
        busy: document.querySelector('.print-submit')?.disabled,
        submitText: document.querySelector('.print-submit')?.textContent,
        toast: document.querySelector('.toast')?.textContent,
      })`)
      const targets = await pageTargets().catch(() => [])
      throw new Error(`${error.message}\nPrint handoff state: ${JSON.stringify(state)}\nTargets: ${JSON.stringify(targets.map((target) => ({ title: target.title, url: target.url })))}`)
    }
    printDirectory = ownedPrintDirectory(printSurface.url)
    assert.equal(dialog.title, 'Set up your print')
    assert.equal(dialog.sideBySide, true)
    assert.equal(immediateHandoff.optionsFrozen, true)
    assert.ok(printDirectory, `The packaged print surface was not in an owned temporary directory: ${printSurface.url}`)
    process.stdout.write(`${JSON.stringify({ dialog, immediateHandoff, printSurface: printSurface.url })}\n`)
  } finally {
    if (workspaceClient) {
      try { await Promise.race([workspaceClient.call('Browser.close'), delay(800)]) } catch {}
      workspaceClient.close()
    }
    spawnSync('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
    await delay(250)
    if (printDirectory) {
      try { fs.rmSync(printDirectory, { recursive: true, force: true }) } catch {}
    }
    try { fs.rmSync(profile, { recursive: true, force: true }) } catch {}
  }
}

main().catch((error) => {
  console.error(error.stack || error)
  process.exitCode = 1
})
