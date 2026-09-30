'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')

function argument(name) {
  const prefix = `--${name}=`
  return process.argv.find((value) => value.startsWith(prefix))?.slice(prefix.length) || null
}

function writeMinimalPdf(targetPath) {
  const content = 'BT /F1 18 Tf 72 720 Td (Simple export smoke) Tj ET\n'
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${Buffer.byteLength(content, 'ascii')} >>\nstream\n${content}endstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  let output = '%PDF-1.4\n'
  const offsets = [0]
  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(output, 'ascii'))
    output += `${index + 1} 0 obj\n${object}\nendobj\n`
  })
  const xrefOffset = Buffer.byteLength(output, 'ascii')
  output += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  output += offsets.slice(1).map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')
  output += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`
  fs.writeFileSync(targetPath, output, 'ascii')
}

function writeVideoFixture(targetPath) {
  const ffmpeg = argument('ffmpeg') || 'ffmpeg'
  const generated = spawnSync(ffmpeg, [
    '-hide_banner', '-loglevel', 'error', '-y',
    '-f', 'lavfi', '-i', 'color=c=0x356a9f:s=320x180:d=1:r=12',
    '-c:v', 'libvpx-vp9', '-pix_fmt', 'yuv420p', '-an', targetPath,
  ], { encoding: 'utf8', windowsHide: true })
  if (generated.error || generated.status !== 0) {
    throw new Error(`Could not generate the WebM export fixture with ffmpeg. ${generated.error?.message || generated.stderr || ''}`.trim())
  }
}

const executable = path.resolve(argument('exe') || path.join(__dirname, '..', 'release', 'simple.exe'))
const firstPort = Number(argument('port') || 9699)
const onlyMode = argument('only')
const supportedModes = ['image', 'docs', 'calc', 'pdf', 'video']

assert.ok(fs.existsSync(executable) && fs.statSync(executable).isFile(), `Missing packaged executable: ${executable}`)
assert.ok(Number.isInteger(firstPort) && firstPort >= 1024 && firstPort <= 65_530, `Invalid first debugging port: ${firstPort}`)
assert.ok(!onlyMode || supportedModes.includes(onlyMode), `Unknown --only mode: ${onlyMode}`)

const generatedFixtureDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-export-smoke-fixtures-'))
const fixturePaths = {
  image: path.resolve(argument('image') || path.join(__dirname, '..', '..', 'simple_image_source', 'public', 'brand-icon.png')),
  docs: path.resolve(argument('doc') || path.join(__dirname, '..', '..', 'simple_doc_source', 'qa', 'fixtures', 'simple-docs-roundtrip-fixture.docx')),
  calc: path.resolve(argument('sheet') || path.join(__dirname, '..', '..', 'simple_calc_source', 'scripts', 'fixtures', 'print-smoke.csv')),
  pdf: path.resolve(argument('pdf') || path.join(generatedFixtureDirectory, 'export-smoke.pdf')),
  video: path.resolve(argument('video') || path.join(generatedFixtureDirectory, 'export-smoke.webm')),
}

if (!argument('pdf') && (!onlyMode || onlyMode === 'pdf')) writeMinimalPdf(fixturePaths.pdf)
if (!argument('video') && (!onlyMode || onlyMode === 'video')) writeVideoFixture(fixturePaths.video)

for (const mode of supportedModes.filter((value) => !onlyMode || value === onlyMode)) {
  const filePath = fixturePaths[mode]
  assert.ok(fs.existsSync(filePath) && fs.statSync(filePath).isFile(), `Missing ${mode} fixture: ${filePath}`)
}

const modeContracts = {
  image: {
    ready: `document.querySelector('canvas')?.width > 0 && !document.querySelector('button[title="Export As (Ctrl+Shift+E)"]').disabled`,
    control: `button[title="Export As (Ctrl+Shift+E)"]`,
    surface: `Boolean(document.querySelector('[role="menu"][aria-label="Export image as"]'))`,
    formats: `[...document.querySelectorAll('[role="menu"][aria-label="Export image as"] [data-export-format]')].map((item) => item.dataset.exportFormat)`,
    expected: ['png', 'jpeg', 'webp', 'pdf'],
  },
  docs: {
    ready: `document.querySelector('#welcome')?.hidden && document.querySelector('#loading-overlay')?.hidden && Boolean(document.querySelector('#export-as-button'))`,
    control: `#export-as-button`,
    surface: `Boolean(document.querySelector('#export-menu:not([hidden])'))`,
    formats: `[...document.querySelectorAll('#export-menu:not([hidden]) [data-export-format]')].map((item) => item.dataset.exportFormat)`,
    expected: ['docx', 'pdf', 'html', 'md', 'txt'],
  },
  calc: {
    ready: `Boolean(document.querySelector('.workbook-app .sheet-viewport') && document.querySelector('.export-as-command'))`,
    control: `.export-as-command`,
    surface: `Boolean(document.querySelector('[role="dialog"][aria-labelledby="export-dialog-title"]'))`,
    formats: `[...document.querySelectorAll('[role="dialog"][aria-labelledby="export-dialog-title"] [data-export-format]')].map((item) => item.dataset.exportFormat)`,
    expected: ['pdf', 'xlsx', 'xls', 'html', 'ods', 'csv', 'tsv'],
  },
  pdf: {
    ready: `document.querySelector('.page-canvas')?.width > 300 && Boolean(document.querySelector('.export-as-button'))`,
    control: `.export-as-button`,
    surface: `Boolean(document.querySelector('[role="dialog"][aria-label="Export As"]'))`,
    formats: `[...document.querySelectorAll('[role="dialog"][aria-label="Export As"] [role="radio"] strong')].map((item) => item.textContent.trim())`,
    expected: ['PDF', 'PNG images', 'JPEG images', 'WebP images', 'Word document', 'Plain text', 'Markdown', 'Web page'],
  },
  video: {
    ready: `document.querySelector('video')?.readyState >= 2 && document.querySelector('video')?.videoWidth > 0 && Boolean(document.querySelector('.export-toolbar-button'))`,
    control: `.export-toolbar-button`,
    surface: `Boolean(document.querySelector('.export-modal'))`,
    formats: `[...document.querySelectorAll('.export-modal .export-options .export-option-copy strong')].map((item) => item.textContent.trim())`,
    expected: ['PNG image', 'JPEG image', 'Original video copy'],
  },
}

const delay = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

async function workspaceTarget(port, mode) {
  const started = Date.now()
  while (Date.now() - started < 60_000) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      const target = targets.find((item) => item.type === 'page' && item.url.includes(`/modules/${mode}/dist/index.html`))
      if (target?.webSocketDebuggerUrl) return target
    } catch {}
    await delay(150)
  }
  throw new Error(`The packaged ${mode} workspace did not start.`)
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
      const { resolve, reject } = this.pending.get(message.id)
      this.pending.delete(message.id)
      if (message.error) reject(new Error(message.error.message))
      else resolve(message.result)
    })
    this.socket.addEventListener('close', () => {
      for (const { reject } of this.pending.values()) reject(new Error('The packaged workspace closed.'))
      this.pending.clear()
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
      userGesture: true,
    })
    if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text)
    return response.result.value
  }

  async waitFor(expression, timeout = 30_000) {
    const started = Date.now()
    while (Date.now() - started < timeout) {
      try {
        if (await this.evaluate(expression)) return
      } catch {}
      await delay(100)
    }
    throw new Error(`Timed out waiting for renderer state: ${expression}`)
  }

  close() {
    try { this.socket?.close() } catch {}
  }
}

async function readFormats(client, contract) {
  const formats = await client.evaluate(contract.formats)
  assert.deepEqual(formats, contract.expected)
  return formats
}

async function closeExportSurface(client, contract) {
  await client.evaluate(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', bubbles: true, cancelable: true }))`)
  await client.waitFor(`!(${contract.surface})`)
}

async function probe(mode, filePath, port) {
  const contract = modeContracts[mode]
  const profileDirectory = fs.mkdtempSync(path.join(os.tmpdir(), `simple-${mode}-export-profile-`))
  const child = spawn(executable, [`--remote-debugging-port=${port}`, `--user-data-dir=${profileDirectory}`, filePath], {
    env: { ...process.env, APPDATA: profileDirectory },
    stdio: 'ignore',
    windowsHide: true,
  })
  let client = null
  try {
    const workspace = await workspaceTarget(port, mode)
    client = new CdpClient(workspace.webSocketDebuggerUrl)
    await client.connect()
    await client.waitFor(contract.ready, mode === 'video' ? 45_000 : 30_000)

    let documentState = null
    if (mode === 'docs') {
      documentState = await client.evaluate(`({
        title: document.title,
        saveState: document.querySelector('#save-state')?.textContent?.trim() || '',
        editorActive: document.querySelector('#editor')?.classList.contains('is-active') || false,
        welcomeHidden: document.querySelector('#welcome')?.hidden || false,
        compatibility: document.querySelector('#document-compatibility')?.textContent || '',
        pageView: document.querySelector('#original-layout-button')?.getAttribute('aria-pressed') === 'true',
      })`)
      assert.equal(documentState.editorActive, true, 'The packaged Word document did not enter the editable Docs surface.')
      assert.equal(documentState.welcomeHidden, true, 'The packaged Docs welcome surface still covers the opened document.')
      if (path.extname(filePath).toLowerCase() === '.doc') {
        if (documentState.compatibility.includes('Text-only import')) {
          assert.equal(documentState.saveState, 'Unsaved', 'A text-only fallback must require a separate save.')
          assert.match(documentState.title, /^• /, 'A text-only fallback must be marked unsaved.')
        } else {
          assert.equal(documentState.saveState, 'Saved locally', 'Opening a formatted legacy document must not create unsaved changes.')
          assert.doesNotMatch(documentState.title, /^• /, 'An unchanged formatted legacy document must remain clean.')
          assert.equal(documentState.pageView, true, 'A formatted legacy document must offer its source Page view.')
        }
      }
    }

    const toolbar = await client.evaluate(`(() => {
      const button = document.querySelector(${JSON.stringify(contract.control)})
      if (!button) return null
      const bounds = button.getBoundingClientRect()
      const style = getComputedStyle(button)
      return {
        text: button.textContent.replace(/\\s+/g, ' ').trim(),
        title: button.getAttribute('title') || '',
        disabled: Boolean(button.disabled),
        visible: style.display !== 'none' && style.visibility !== 'hidden' && bounds.width > 0 && bounds.height > 0 && bounds.right > 0 && bounds.bottom > 0 && bounds.left < innerWidth && bounds.top < innerHeight,
      }
    })()`)
    assert.ok(toolbar, `The packaged ${mode} workspace has no prominent Export As control.`)
    assert.equal(toolbar.disabled, false, `The packaged ${mode} Export As control is disabled.`)
    assert.equal(toolbar.visible, true, `The packaged ${mode} Export As control is outside the visible workspace.`)
    assert.match(toolbar.text, /export as/i, `The packaged ${mode} Export As control is not clearly labelled.`)

    await client.evaluate(`document.querySelector(${JSON.stringify(contract.control)}).click()`)
    await client.waitFor(contract.surface)
    const toolbarFormats = await readFormats(client, contract)
    await closeExportSurface(client, contract)

    await client.evaluate(`document.dispatchEvent(new KeyboardEvent('keydown', { key: 'E', code: 'KeyE', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true }))`)
    await client.waitFor(contract.surface)
    const shortcutFormats = await readFormats(client, contract)
    await closeExportSurface(client, contract)

    const result = { mode, ...(documentState ? { documentState } : {}), toolbar, toolbarFormats, shortcutFormats }
    process.stdout.write(`${mode}: packaged Export As toolbar and Ctrl+Shift+E passed (${toolbarFormats.join(', ')}).\n`)
    return result
  } catch (error) {
    if (client) {
      const state = await client.evaluate(`({ title: document.title, text: document.body.innerText.slice(0, 1600), buttons: [...document.querySelectorAll('button')].map((button) => button.getAttribute('aria-label') || button.title || button.textContent.replace(/\\s+/g, ' ').trim()).filter(Boolean).slice(0, 100) })`).catch(() => null)
      if (state) error.message = `${error.message}\n${mode} renderer state: ${JSON.stringify(state)}`
    }
    throw error
  } finally {
    if (client) {
      try { await Promise.race([client.call('Browser.close'), delay(1_000)]) } catch {}
      client.close()
    }
    spawnSync('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore', windowsHide: true })
    await delay(300)
    try { fs.rmSync(profileDirectory, { recursive: true, force: true }) } catch {}
  }
}

async function main() {
  const modes = supportedModes.filter((mode) => !onlyMode || mode === onlyMode)
  const results = []
  for (let index = 0; index < modes.length; index += 1) {
    const mode = modes[index]
    results.push(await probe(mode, fixturePaths[mode], firstPort + index))
  }
  process.stdout.write(`${JSON.stringify({ executable, results }, null, 2)}\n`)
}

main().catch((error) => {
  console.error(error.stack || error)
  process.exitCode = 1
}).finally(() => {
  try { fs.rmSync(generatedFixtureDirectory, { recursive: true, force: true }) } catch {}
})
