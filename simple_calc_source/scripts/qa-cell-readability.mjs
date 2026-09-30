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
const fixturePath = path.join(root, 'tmp', 'cell-readability.xlsx')
const ExcelJS = require('exceljs')
const { workbookPayloadFromBytes, serializeWorkbook } = require('../electron/workbooks.cjs')
const basePort = 9_500 + (process.pid % 200)
const vitePort = basePort
const debugPort = basePort + 200
const appUrl = `http://127.0.0.1:${vitePort}/`
const profilePath = path.resolve(root, 'tmp', `cell-readability-profile-${process.pid}`)
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
  await fs.mkdir(path.dirname(fixturePath), { recursive: true })
  const source = new ExcelJS.Workbook()
  const sheet = source.addWorksheet('Readability', { properties: { defaultRowHeight: 24, defaultColWidth: 18 } })
  for (const [row, family] of [[1, 'Arial'], [2, 'Times New Roman'], [3, 'Calibri'], [4, 'Noto Sans CJK SC']]) {
    sheet.getRow(row).height = 15
    for (const col of [1, 2, 3]) {
      const cell = sheet.getCell(row, col)
      cell.value = 9876.54
      cell.font = { name: family, size: 12 }
      cell.alignment = { vertical: ['top', 'middle', 'bottom'][col - 1] }
      cell.border = { top: { style: 'thin' }, bottom: { style: 'thin' } }
    }
  }
  sheet.getCell('A5').value = 'Default row'
  sheet.getCell('B5').value = 1234
  sheet.getRow(6).height = 34
  sheet.getCell('A6').value = 'First line\nSecond line'
  sheet.getCell('A6').font = { name: 'Arial', size: 12 }
  sheet.getCell('A6').alignment = { wrapText: true, vertical: 'top' }
  const sourceBytes = Buffer.from(await source.xlsx.writeBuffer())
  const opened = await workbookPayloadFromBytes('cell-readability.xlsx', sourceBytes)
  const roundtripBytes = await serializeWorkbook(opened.workbook, 'xlsx', { baseBytes: sourceBytes, sourceFormat: 'xlsx' })
  await fs.writeFile(fixturePath, roundtripBytes)
  const roundtrip = new ExcelJS.Workbook()
  await roundtrip.xlsx.load(roundtripBytes)
  assert.equal(roundtrip.worksheets[0].properties.defaultRowHeight, 24)
  assert.equal(roundtrip.worksheets[0].properties.defaultColWidth, 18)
  for (const row of [1, 2, 3, 4]) {
    assert.equal(roundtrip.worksheets[0].getRow(row).height, 15)
    for (const col of [1, 2, 3]) assert.equal(roundtrip.worksheets[0].getCell(row, col).alignment.vertical, ['top', 'middle', 'bottom'][col - 1])
  }
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
  await client.waitFor("document.querySelector('.document-title')?.textContent.includes('cell-readability') && document.querySelector('.sheet-viewport')", 'readability fixture')
  for (const zoom of [100, 50, 150, 200]) {
    await client.evaluate(`(() => {
      const input = document.querySelector('input[aria-label="Zoom"]');
      Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(input, '${zoom}');
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
    })()`)
    await client.waitFor(`document.querySelector('.zoom-control span')?.textContent === '${zoom}%'`, 'zoom ' + zoom)
    await client.waitFor(`Math.abs([...document.querySelectorAll('.grid-cell')].find(item => item.id.endsWith('-A1')).getBoundingClientRect().height - ${20 * zoom / 100}) < 0.5`, 'cell geometry at zoom ' + zoom)
    const metrics = await client.evaluate(`(() => {
      const metric = (address) => {
        const cell = [...document.querySelectorAll('.grid-cell')].find(item => item.id.endsWith('-' + address));
        const content = cell.querySelector('.cell-content');
        const box = cell.getBoundingClientRect(), text = content.getBoundingClientRect(), style = getComputedStyle(cell);
        return { address, top: box.top, bottom: box.bottom, height: box.height, width: box.width, textTop: text.top, textBottom: text.bottom, align: style.alignItems, lineHeight: style.lineHeight };
      };
      return { cells: ['A1','B1','C1','A2','B2','C2','A3','B3','C3','A4','B4','C4'].map(metric), defaultCell: metric('A5'), wrapped: metric('A6') };
    })()`)
    for (const cell of metrics.cells) {
      assert.ok(cell.textTop >= cell.top + 0.5, JSON.stringify({ zoom, ...cell }) + ': text must stay below top border')
      assert.ok(cell.textBottom <= cell.bottom - 0.5, JSON.stringify({ zoom, ...cell }) + ': text must stay above bottom border')
      assert.equal(Math.round(cell.height), Math.round(20 * zoom / 100), 'source row height must be preserved')
      assert.ok(cell.align.includes(cell.address.startsWith('A') ? 'flex-start' : cell.address.startsWith('B') ? 'center' : 'flex-end'))
    }
    assert.ok(Math.abs(metrics.defaultCell.height - 32 * zoom / 100) < 1, 'source default row height')
    // Excel draws a width-18 column at Truncate(((256 * 18 + 18) / 256) * 7) = 126 px.
    assert.ok(Math.abs(metrics.defaultCell.width - 126 * zoom / 100) < 1, 'source default column width')
    assert.ok(metrics.wrapped.textBottom <= metrics.wrapped.bottom, 'wrapped text fits explicit row height')
    process.stdout.write('Readability and dimensions passed at ' + zoom + '% zoom.\n')
  }
  const shot = await client.call('Page.captureScreenshot', { format: 'png' })
  await fs.writeFile(path.join(root, 'tmp', 'cell-readability.png'), Buffer.from(shot.data, 'base64'))
  process.stdout.write('Cell readability UI and formatting round-trip QA passed.\n')
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
