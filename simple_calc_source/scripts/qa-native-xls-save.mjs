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
const originalFixture = path.join(scriptDirectory, 'fixtures', 'schedule_template.xls')
const fixturePath = path.join(root, 'tmp', 'native-save-ui.xls')
const pdfPath = path.join(root, 'tmp', 'schedule-export.pdf')
const entryPath = path.join(root, 'tmp', `native-save-entry-${process.pid}.cjs`)
const { workbookPayloadFromPath } = require('../electron/workbooks.cjs')
const crypto = require('node:crypto')
const basePort = 9_500 + (process.pid % 200)
const vitePort = basePort
const debugPort = basePort + 200
const appUrl = `http://127.0.0.1:${vitePort}/`
const profilePath = path.resolve(root, 'tmp', `native-save-profile-${process.pid}`)
const deadline = Date.now() + 180_000
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
  const timings = {}
  await fs.mkdir(path.dirname(fixturePath), { recursive: true })
  await fs.copyFile(originalFixture, fixturePath)
  const originalBytes = await fs.readFile(fixturePath)
  const before = await workbookPayloadFromPath(fixturePath)
  assert.equal(before.requiresSaveAs, false)
  await fs.mkdir(profilePath, { recursive: true })

  const viteBin = path.join(root, 'node_modules', 'vite', 'bin', 'vite.js')
  vite = spawn(process.execPath, [viteBin, '--host', '127.0.0.1', '--port', String(vitePort), '--strictPort'], {
    cwd: root,
    windowsHide: true,
    stdio: ['ignore', 'pipe', 'pipe'],
  })
  collectOutput(vite, (chunk) => { viteOutput = `${viteOutput}${chunk}`.slice(-8_000) })
  await waitForUrl(appUrl)

  // The production Export handler and PDF renderer run normally. Only the native
  // filename picker is substituted with a disposable QA output path.
  await fs.writeFile(entryPath, `const {dialog}=require('electron');dialog.showSaveDialog=async(_w,o)=>{if(o.title!=='Export spreadsheet as PDF')throw new Error('Unexpected save dialog '+o.title);return {canceled:false,filePath:${JSON.stringify(pdfPath)}}};require('../electron/main.cjs');`)
  const openStarted = Date.now()
  electron = spawn(electronPath, [
    `--remote-debugging-port=${debugPort}`,
    `--user-data-dir=${profilePath}`,
    entryPath,
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
  await client.waitFor("document.querySelector('.document-title')?.textContent.includes('native-save-ui.xls') && document.querySelector('.sheet-viewport')", 'XLS fixture')
  await client.waitFor("document.querySelector('[id$=\"-E2\"]')", 'first visible data row')
  timings.freshOpenMs = Date.now() - openStarted
  const screenshot = async (name) => {
    const shot = await client.call('Page.captureScreenshot', { format: 'png' })
    await fs.writeFile(path.join(root, 'tmp', name), Buffer.from(shot.data, 'base64'))
  }
  const setZoom = async (value) => {
    await client.evaluate(`(() => { const s=document.querySelector('.toolbar-zoom-select'); s.value=${JSON.stringify(value)}; s.dispatchEvent(new Event('change',{bubbles:true})); })()`)
    await pause(150)
  }
  const sourceRowHeight = await client.evaluate("document.querySelector('[id$=\"-E2\"]').getBoundingClientRect().height")
  assert.ok(Math.abs(sourceRowHeight - 16 * 4 / 3) <= .6, `automatic row height must be16pt with pixel rounding, got${sourceRowHeight}px`)
  await screenshot('schedule-normal.png')
  const fitStarted = Date.now()
  await setZoom('fit-sheet')
  await client.waitFor("document.querySelector('[id$=\"-D37\"]')", 'whole form fitted')
  timings.fitSheetMs = Date.now() - fitStarted
  const form = await client.evaluate(`(() => { const a=document.querySelector('[id$="-A1"]').getBoundingClientRect(); const b=document.querySelector('[id$="-D37"]').getBoundingClientRect(); const v=document.querySelector('.sheet-viewport').getBoundingClientRect(); return {firstTop:a.top,lastBottom:b.bottom,viewportTop:v.top,viewportBottom:v.bottom,zoom:document.querySelector('.toolbar-zoom-select').value,text:document.querySelector('[id$="-A32"]').textContent}; })()`)
  assert.ok(form.firstTop >= form.viewportTop && form.lastBottom <= form.viewportBottom - 10, 'Fit sheet shows headers, full prose and signature boxes together')
  assert.match(form.text, /incorporated herein by reference/)
  await screenshot('schedule-fit-sheet.png')
  await setZoom('fit-width')
  const widthFit = await client.evaluate(`(() => { const b=document.querySelector('[id$="-E1"]').getBoundingClientRect(); const v=document.querySelector('.sheet-viewport').getBoundingClientRect();return{right:b.right,viewportRight:v.right,zoom:document.querySelector('.toolbar-zoom-select').value} })()`)
  assert.ok(widthFit.right <= widthFit.viewportRight - 10, 'Fit width keeps all5 data columns in view')
  await setZoom('100')
  const edit = async (address, value) => {
    await client.evaluate(`document.querySelector('[id$="-${address}"]').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, button: 0 }))`)
    await pause(50)
    await client.evaluate(`document.querySelector('[id$="-${address}"]').dispatchEvent(new MouseEvent('dblclick', { bubbles: true, button: 0 }))`)
    await client.waitFor("document.querySelector('.cell-editor')", 'cell editor')
    await client.evaluate("document.querySelector('.cell-editor').select()")
    await client.call('Input.insertText', { text: value })
    await client.evaluate("document.querySelector('.cell-editor').blur()")
    await client.waitFor("!document.querySelector('.cell-editor')", 'committed cell')
  }
  await edit('E2', '123.45')
  await edit('E3', '=E2*2')
  const saveStarted = Date.now()
  await client.evaluate("document.querySelector('.save-command').click()")
  await client.waitFor("document.body.textContent.includes('Saved native-save-ui.xls')", 'native XLS save')
  timings.saveXlsMs = Date.now() - saveStarted
  assert.equal(await client.evaluate("document.querySelector('.document-title').textContent.includes('.xlsx')"), false, 'Save must retain XLS filename')
  const saved = await workbookPayloadFromPath(fixturePath)
  const sheet = saved.workbook.sheets[0]
  assert.equal(saved.sourceFormat, 'xls')
  assert.equal(sheet.cells.E2.value, 123.45)
  assert.equal(sheet.cells.E3.formula.replaceAll('$',''), 'E2*2')
  assert.equal(sheet.cells.E3.result, 246.9)
  assert.equal(sheet.cells.E30.formula, 'SUM(E2:E26)')
  assert.equal(sheet.cells.E30.result, 370.35)
  assert.equal(sheet.cells.A1.style.font.name, before.workbook.sheets[0].cells.A1.style.font.name)
  assert.equal(sheet.cells.A1.style.font.size, before.workbook.sheets[0].cells.A1.style.font.size)
  assert.equal(sheet.cells.A1.style.alignment.vertical, 'bottom')
  assert.equal(sheet.cells.E2.style.border.bottom.style, 'thin')
  assert.equal(sheet.rowHeights['32'], before.workbook.sheets[0].rowHeights['32'])
  assert.equal(sheet.rowHeights['2'], 16)
  assert.equal(saved.workbook.metadata.normalFont.name, 'Arial')
  assert.equal(saved.workbook.metadata.normalFont.size, 10)
  assert.equal(sheet.pageSetup.scale, 91)
  assert.equal(sheet.pageSetup.fitToPage, false)
  assert.equal(sheet.colCount, before.workbook.sheets[0].colCount, 'hidden legacy tail must not expand the editable grid')
  for (let col = 1; col <= 6; col++) assert.equal(sheet.colWidths[col], before.workbook.sheets[0].colWidths[col], `column ${col} width survives native conversion exactly`)
  assert.deepEqual(sheet.pageSetup.margins, before.workbook.sheets[0].pageSetup.margins, 'source margins survive native conversion exactly')
  assert.match(sheet.headerFooter.oddHeader, /SCHEDULE OF ACCOUNTS/)
  assert.deepEqual([...sheet.merges].sort(), [...before.workbook.sheets[0].merges].sort())
  assert.ok(await client.evaluate("document.querySelector('[aria-label=\"Show original backup\"]') !== null"), 'original backup must be easy to find')
  assert.ok((await fs.readFile(fixturePath)).subarray(0,8).equals(Buffer.from('d0cf11e0a1b11ae1','hex')))
  const copy = await client.evaluate(`window.simpleCalc.openPath(${JSON.stringify(fixturePath)})`)
  const savedBytes = await fs.readFile(fixturePath)
  const result = await client.evaluate(`window.simpleCalc.saveWorkbook(${JSON.stringify({ documentId: copy.documentId, workbook: copy.workbook, saveAs: false, format: 'xls', sourceUnmodified: true })})`)
  assert.equal(result.format, 'xls')
  assert.ok(savedBytes.equals(await fs.readFile(fixturePath)), 'unchanged XLS save must be byte-for-byte')
  const reopened = await client.evaluate(`window.simpleCalc.openPath(${JSON.stringify(fixturePath)})`)
  await fs.utimes(fixturePath, new Date(), new Date(Date.now() + 5000))
  const conflict = await client.evaluate(`window.simpleCalc.saveWorkbook(${JSON.stringify({ documentId: reopened.documentId, workbook: reopened.workbook, saveAs: false, format: 'xls', sourceUnmodified: true })}).then(() => '', error => error.message)`)
  assert.match(conflict, /changed outside Simple/)
  assert.ok(savedBytes.equals(await fs.readFile(fixturePath)), 'external-change rejection must leave source alone')
  const printStarted = Date.now()
  await client.evaluate("document.querySelector('[aria-label=\"Print (Ctrl+P)\"]').click()")
  await client.waitFor("document.querySelector('.print-preview-frame')?.contentDocument?.querySelector('.sheet-table')", 'saved print layout')
  timings.printPreviewMs = Date.now() - printStarted
  await pause(200)
  const paper = await client.evaluate(`(() => {
    const d=document.querySelector('.print-preview-frame').contentDocument;
    const p=d.querySelector('.print-page'), cell=d.querySelector('[data-address="D37"]'), content=d.querySelector('.page-content');
    const box=e=>{const r=e.getBoundingClientRect();return{left:r.left,right:r.right,top:r.top,bottom:r.bottom,width:r.width,height:r.height,scrollHeight:e.scrollHeight,clientHeight:e.clientHeight,scrollWidth:e.scrollWidth,clientWidth:e.clientWidth,text:e.textContent}};
    const account=address=>{const c=d.querySelector('[data-address="'+address+'"]');return{symbol:box(c.querySelector('.accounting-symbol')),amount:box(c.querySelector('.accounting-amount')),cell:box(c)}};
    return{pages:d.querySelectorAll('.print-page').length,scale:p.dataset.scale,lastBottom:cell?.getBoundingClientRect().bottom,contentBottom:content.getBoundingClientRect().bottom,text:d.body.textContent,html:d.documentElement.outerHTML,
      prose:box(d.querySelector('[data-address="A32"]>span')),label:box(d.querySelector('[data-address="B30"]>span')),accounts:['E2','E3','E30'].map(account),
      boxes:['A35','A37','D37'].map(a=>{const c=d.querySelector('[data-address="'+a+'"]'),s=d.defaultView.getComputedStyle(c);return{address:a,top:s.borderTopStyle,right:s.borderRightStyle,bottom:s.borderBottomStyle,left:s.borderLeftStyle}}),
      warnings:[...document.querySelectorAll('.print-warning')].map(e=>e.textContent)};
  })()`)
  await fs.writeFile(path.join(root, 'tmp', 'schedule-print.html'), paper.html)
  await screenshot('schedule-print-preview.png')
  process.stdout.write(`${JSON.stringify({timings,paper:{...paper,html:undefined,text:undefined}})}\n`)
  assert.equal(paper.pages, 1, 'saved font metrics/91% print layout produce one page')
  assert.equal(paper.scale, '0.9100')
  assert.ok(paper.lastBottom <= paper.contentBottom + 1, 'print must not clip signature boxes beyond page content')
  assert.match(paper.text, /SCHEDULE OF ACCOUNTS/)
  assert.match(paper.text, /incorporated herein by reference/)
  assert.ok(paper.prose.scrollHeight <= paper.prose.clientHeight + 1, 'entire contract paragraph must fit without vertical clipping')
  assert.match(paper.label.text, /INVOICES - THIS SCHEDULE:/)
  assert.ok(paper.label.scrollWidth <= paper.label.clientWidth + 1, 'full invoice label must remain visible across empty adjacent cells')
  for (const account of paper.accounts) {
    assert.ok(account.symbol.right < account.amount.left, 'accounting symbol and right-aligned amount must be separated')
    assert.ok(account.amount.right <= account.cell.right && account.amount.right >= account.cell.right - 8, 'amount must align with its cell right edge')
  }
  assert.equal(paper.accounts[2].amount.text, '370.35')
  for (const box of paper.boxes) for (const side of ['top','right','bottom','left']) assert.equal(box[side], 'solid', `${box.address} ${side} border must close the signature box`)
  await client.evaluate("document.querySelector('.print-preview-frame').contentWindow.scrollTo(0,10000)")
  await pause(100)
  await screenshot('schedule-print-lower.png')
  await client.evaluate("document.querySelector('[aria-label=\"Close print dialog\"]').click()")
  const exportStarted = Date.now()
  await client.evaluate("document.querySelector('.export-as-command').click()")
  assert.equal(await client.evaluate("document.querySelector('[aria-label=\"Use saved export layout\"]').checked"), true, 'PDF export defaults to the same saved layout as Print')
  await client.evaluate("document.querySelector('.export-card .primary-action').click()")
  await client.waitFor("!document.querySelector('.export-card')", 'PDF export completed')
  timings.pdfExportMs = Date.now() - exportStarted
  const pdf = await fs.readFile(pdfPath)
  assert.equal(pdf.subarray(0,5).toString(), '%PDF-')
  assert.ok(pdf.length > 10000, 'real export must contain a rendered PDF')
  assert.ok(savedBytes.equals(await fs.readFile(fixturePath)), 'PDF export must not change the open XLS')
  await setZoom('fit-sheet')
  await screenshot('schedule-edited-fit-sheet.png')
  await fs.writeFile(path.join(root, 'tmp', 'schedule-qa.json'), JSON.stringify({timings,form,widthFit,paper:{...paper,html:undefined,text:undefined}},null,2))
  process.stdout.write(`${JSON.stringify({timings,form,widthFit,paper:{...paper,html:undefined,text:undefined}})}\n`)
  assert.ok(originalBytes.equals(await fs.readFile(originalFixture)), 'reference fixture must remain untouched')
  const shot = await client.call('Page.captureScreenshot', { format: 'png' })
  await fs.writeFile(path.join(root, 'tmp', 'native-save-ui.png'), Buffer.from(shot.data, 'base64'))
  process.stdout.write('Native XLS UI Save passed: original extension/path, new and existing formulas, recalculated values, fonts, borders, row heights, merges, exact unchanged bytes, external-change protection.\n')
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
  await fs.rm(entryPath, { force: true }).catch(() => {})
}
