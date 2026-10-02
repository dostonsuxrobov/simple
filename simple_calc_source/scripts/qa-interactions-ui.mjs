/*
 * Grid interactions driven through the real app in Electron over the Chrome DevTools Protocol:
 * in-cell list dropdowns (CALC-018: the arrow on the active cell, the searchable list, keyboard
 * picks that store typed values, Esc), dragging the selection border to move cells and Ctrl-drag
 * to copy them (CALC-029: references follow, one undo step, "There's already data here"), and
 * View > Page break preview (CALC-014).
 *
 * The script starts its own Vite dev server and an off-screen Electron window with an isolated
 * profile and an in-memory clipboard (the system clipboard is never touched), runs the checks,
 * and stops both process trees. Nothing is printed and no network is used beyond the local
 * dev server.
 *
 *   npm run test:interactions-ui
 *   QA_WORK_DIR=<dir> QA_VITE_PORT=5317 QA_CDP_PORT=9417 node scripts/qa-interactions-ui.mjs
 *   QA_SCREENSHOTS=1 also saves interactions-*.png screenshots into QA_WORK_DIR.
 */
import assert from 'node:assert/strict'
import { execFileSync, spawn } from 'node:child_process'
import fs from 'node:fs/promises'
import { createRequire } from 'node:module'
import path from 'node:path'
import process from 'node:process'
import ExcelJS from 'exceljs'

const require = createRequire(import.meta.url)
const root = process.cwd()
const vitePort = Number(process.env.QA_VITE_PORT || 5317)
const cdpPort = Number(process.env.QA_CDP_PORT || 9417)
const workDir = path.resolve(process.env.QA_WORK_DIR || path.join(root, 'tmp', 'interactions-ui'))
const deadline = Date.now() + Number(process.env.QA_TIMEOUT_MS || 300_000)
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))
const children = []
let passed = 0

function stopChildren() {
  for (const child of children.splice(0)) {
    if (!child.pid || child.exitCode !== null) continue
    try {
      if (process.platform === 'win32') execFileSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore' })
      else process.kill(-child.pid, 'SIGKILL')
    } catch {
      // Already gone.
    }
  }
}
process.on('exit', stopChildren)
process.on('SIGINT', () => { stopChildren(); process.exit(130) })
const watchdog = setTimeout(() => { console.error('Interactions UI QA timed out.'); stopChildren(); process.exit(1) }, Math.max(1_000, deadline - Date.now() + 5_000))
watchdog.unref()

async function waitForHttp(url) {
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url)
      if (response.ok || response.status === 404) return
    } catch {
      // Not up yet.
    }
    await pause(250)
  }
  throw new Error(`Timed out waiting for ${url}`)
}

// ---- Fixture ----------------------------------------------------------------------------
await fs.mkdir(workDir, { recursive: true })
const fixturePath = path.join(workDir, 'interactions-fixture.xlsx')
{
  const book = new ExcelJS.Workbook()
  const data = book.addWorksheet('Data')
  data.getCell('A1').value = 1
  data.getCell('A2').value = 2
  data.getCell('A3').value = 3
  data.getCell('B1').value = { formula: 'A1*10', result: 10 }
  data.getCell('C1').value = { formula: 'SUM(A1:A3)', result: 6 }
  data.getCell('D1').dataValidation = { type: 'list', allowBlank: true, formulae: ['"Low,Medium,High"'] }
  for (const [index, amount] of [10, 20, 30].entries()) {
    const cell = data.getCell(`E${index + 1}`)
    cell.value = amount
    cell.numFmt = '$#,##0.00'
  }
  data.getCell('F1').dataValidation = { type: 'list', allowBlank: true, formulae: ['$E$1:$E$3'] }
  data.getCell('G1').dataValidation = { type: 'list', allowBlank: true, formulae: ['"1,2,3"'] }
  data.getCell('K1').value = 'keep'
  for (let row = 20; row <= 140; row += 1) data.getCell(`A${row}`).value = row
  await fs.writeFile(fixturePath, Buffer.from(await book.xlsx.writeBuffer()))
}

// ---- Processes ---------------------------------------------------------------------------
const harnessPath = path.join(workDir, 'electron-harness.cjs')
await fs.writeFile(harnessPath, `'use strict'
const { app, clipboard, shell } = require('electron')
app.setPath('userData', process.env.QA_USER_DATA)
app.commandLine.appendSwitch('remote-debugging-port', process.env.QA_CDP_PORT)
// The owner's system clipboard stays untouched: the app talks to an in-memory one.
const memory = { text: '', html: '' }
const replacements = {
  readText: () => memory.text,
  readHTML: () => memory.html,
  readBuffer: () => Buffer.alloc(0),
  readImage: () => ({ isEmpty: () => true, toDataURL: () => '' }),
  write: (data) => { memory.text = String((data && data.text) || ''); memory.html = String((data && data.html) || '') },
  writeText: (text) => { memory.text = String(text); memory.html = '' },
}
for (const [name, value] of Object.entries(replacements)) Object.defineProperty(clipboard, name, { value, configurable: true, writable: true })
try {
  Object.defineProperty(shell, 'openExternal', { value: async () => {}, configurable: true, writable: true })
} catch {}
app.on('browser-window-created', (_event, win) => {
  win.webContents.setBackgroundThrottling(false)
  const offScreen = () => { try { win.setPosition(-3400, -3400) } catch {} }
  offScreen()
  win.on('show', offScreen)
})
require(process.env.QA_MAIN)
`)

const vite = spawn(process.execPath, [path.join(root, 'node_modules', 'vite', 'bin', 'vite.js'), '--port', String(vitePort), '--strictPort', '--host', '127.0.0.1', '--clearScreen', 'false'], {
  cwd: root,
  stdio: ['ignore', 'pipe', 'pipe'],
  detached: process.platform !== 'win32',
})
children.push(vite)
vite.stderr.on('data', (chunk) => { if (process.env.QA_VERBOSE) process.stderr.write(chunk) })
const appUrl = `http://127.0.0.1:${vitePort}/`
await waitForHttp(appUrl)

const electronBinary = require('electron')
const electron = spawn(electronBinary, [harnessPath], {
  cwd: root,
  stdio: ['ignore', 'pipe', 'pipe'],
  detached: process.platform !== 'win32',
  env: {
    ...process.env,
    VITE_DEV_SERVER_URL: appUrl,
    QA_USER_DATA: path.join(workDir, 'electron-profile'),
    QA_CDP_PORT: String(cdpPort),
    QA_MAIN: path.join(root, 'electron', 'main.cjs'),
    ELECTRON_ENABLE_LOGGING: '',
  },
})
children.push(electron)
electron.stderr.on('data', (chunk) => { if (process.env.QA_VERBOSE) process.stderr.write(chunk) })

// ---- CDP ---------------------------------------------------------------------------------
let target
while (Date.now() < deadline) {
  try {
    const targets = await (await fetch(`http://127.0.0.1:${cdpPort}/json/list`)).json()
    target = targets.find((item) => item.type === 'page' && item.url.startsWith(appUrl))
    if (target?.webSocketDebuggerUrl) break
  } catch {
    // Electron is still starting.
  }
  await pause(250)
}
if (!target?.webSocketDebuggerUrl) throw new Error('No simple_calc page target is available.')
const socket = new WebSocket(target.webSocketDebuggerUrl)
await new Promise((resolve, reject) => {
  socket.addEventListener('open', resolve, { once: true })
  socket.addEventListener('error', reject, { once: true })
})
let requestId = 0
const pending = new Map()
const runtimeErrors = []
socket.addEventListener('message', (event) => {
  const message = JSON.parse(String(event.data))
  if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params?.exceptionDetails?.exception?.description || message.params?.exceptionDetails?.text || 'Runtime exception')
  const request = pending.get(message.id)
  if (!request) return
  pending.delete(message.id)
  if (message.error) request.reject(new Error(message.error.message))
  else request.resolve(message.result)
})
function call(method, params = {}) {
  const id = ++requestId
  socket.send(JSON.stringify({ id, method, params }))
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }))
}
async function evaluate(expression) {
  const result = await call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text)
  return result.result?.value
}
async function waitFor(expression, label, timeout = 10_000) {
  const until = Math.min(deadline, Date.now() + timeout)
  while (Date.now() < until) {
    if (await evaluate(`Boolean(${expression})`)) return
    await pause(50)
  }
  throw new Error(`Timed out waiting for ${label}`)
}
const settle = () => evaluate('new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve(true))))')
async function snap(name) {
  if (!process.env.QA_SCREENSHOTS) return
  await settle()
  const shot = await call('Page.captureScreenshot', { format: 'png' })
  await fs.writeFile(path.join(workDir, `interactions-${name}.png`), Buffer.from(shot.data, 'base64'))
}

await call('Runtime.enable')
await call('Page.enable')
await call('Emulation.setFocusEmulationEnabled', { enabled: true })
await waitFor('window.__calcQA && document.readyState === "complete"', 'the app and its QA hook', 60_000)

// ---- Input helpers -------------------------------------------------------------------------
const KEYS = {
  Enter: { code: 'Enter', vk: 13 }, Escape: { code: 'Escape', vk: 27 }, F5: { code: 'F5', vk: 116 },
  ArrowUp: { code: 'ArrowUp', vk: 38 }, ArrowDown: { code: 'ArrowDown', vk: 40 },
}
const MODIFIERS = { alt: 1, ctrl: 2, meta: 4, shift: 8 }
const modifierMask = (modifiers = []) => modifiers.reduce((mask, name) => mask | MODIFIERS[name], 0)
function keyInfo(key) {
  if (KEYS[key]) return { key, ...KEYS[key] }
  if (/^[a-z]$/i.test(key)) return { key, code: `Key${key.toUpperCase()}`, vk: key.toUpperCase().charCodeAt(0) }
  return { key, code: '', vk: 0 }
}
async function press(key, modifiers = []) {
  const info = keyInfo(key)
  const mask = modifierMask(modifiers)
  const printable = info.key.length === 1 && !(mask & (MODIFIERS.ctrl | MODIFIERS.alt | MODIFIERS.meta))
  await call('Input.dispatchKeyEvent', { type: printable ? 'keyDown' : 'rawKeyDown', key: info.key, code: info.code, windowsVirtualKeyCode: info.vk, nativeVirtualKeyCode: info.vk, modifiers: mask, ...(printable ? { text: info.key, unmodifiedText: info.key } : {}) })
  await call('Input.dispatchKeyEvent', { type: 'keyUp', key: info.key, code: info.code, windowsVirtualKeyCode: info.vk, nativeVirtualKeyCode: info.vk, modifiers: mask })
  await settle()
}
const cell = (address) => evaluate(`(() => { const book = window.__calcQA.workbook(); const sheet = book.sheets.find((item) => item.id === book.activeSheetId); return sheet.cells[${JSON.stringify(address)}] || null })()`)
async function goTo(reference) {
  await press('F5')
  await waitFor("document.activeElement === document.querySelector('.name-box')", 'the name box (F5)')
  await call('Input.insertText', { text: reference })
  await press('Enter')
  await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", `the grid after Go To ${reference}`)
  await settle()
}
async function centerOf(selector) {
  const box = await evaluate(`(() => { const element = document.querySelector(${JSON.stringify(selector)}); if (!element) return null; const rect = element.getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()`)
  assert.ok(box, `${selector} is on screen`)
  return box
}
async function clickAt(box, modifiers = []) {
  const mask = modifierMask(modifiers)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: box.x, y: box.y, modifiers: mask })
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: box.x, y: box.y, button: 'left', buttons: 1, clickCount: 1, modifiers: mask })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: box.x, y: box.y, button: 'left', buttons: 0, clickCount: 1, modifiers: mask })
  await settle()
}
const clickSelector = async (selector) => clickAt(await centerOf(selector))
/** Presses on a side of the selection border and drags it onto a cell. */
async function dragBorder(side, toAddress, modifiers = []) {
  const from = await centerOf(`.selection-outline[data-selection-pane="body"] [data-move-edge="${side}"]`)
  const to = await centerOf(`[data-cell-address="${toAddress}"][data-cell-pane="body"]`)
  const mask = modifierMask(modifiers)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: from.x, y: from.y, modifiers: mask })
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: from.x, y: from.y, button: 'left', buttons: 1, clickCount: 1, modifiers: mask })
  for (let step = 1; step <= 8; step += 1) {
    const progress = step / 8
    await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: from.x + (to.x - from.x) * progress, y: from.y + (to.y - from.y) * progress, button: 'left', buttons: 1, modifiers: mask })
    await pause(15)
  }
  await settle()
  const preview = await evaluate("document.querySelector('.move-preview')?.dataset.movePreview || ''")
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: to.x, y: to.y, button: 'left', buttons: 0, clickCount: 1, modifiers: mask })
  await pause(80)
  await settle()
  return preview
}
async function chooseAlertButton(label) {
  await waitFor(`[...document.querySelectorAll('[role="alertdialog"] button')].some((button) => button.textContent.trim() === ${JSON.stringify(label)})`, `the "${label}" button`)
  await evaluate(`[...document.querySelectorAll('[role="alertdialog"] button')].find((button) => button.textContent.trim() === ${JSON.stringify(label)}).click()`)
  await waitFor("!document.querySelector('[role=\"alertdialog\"]')", 'the alert to close')
  await settle()
}
const activeOption = () => evaluate("document.querySelector('[data-validation-dropdown] [role=\"option\"][aria-selected=\"true\"]')?.dataset.option ?? null")
async function check(label, run) {
  try {
    await run()
    passed += 1
    if (process.env.QA_VERBOSE) console.log(`ok - ${label}`)
  } catch (error) {
    console.error(`FAILED: ${label}`)
    await snap(`failed-${passed + 1}`).catch(() => {})
    throw error
  }
}

try {
  await evaluate(`window.__calcQA.openPath(${JSON.stringify(fixturePath)})`)
  await waitFor("document.querySelector('.sheet-viewport') && window.__calcQA.workbook()?.sheets?.length === 1", 'the fixture workbook', 20_000)
  await settle()

  // ---- CALC-018: in-cell dropdowns ------------------------------------------------------------
  await check('a list cell shows its value as text, with the arrow only on the active cell', async () => {
    assert.equal(await evaluate("document.querySelectorAll('select.cell-dropdown').length"), 0, 'no native <select> pills')
    await goTo('A1')
    assert.equal(await evaluate("Boolean(document.querySelector('[data-dropdown-button]'))"), false, 'no arrow on an ordinary cell')
    await goTo('D1')
    await waitFor("document.querySelector('[data-dropdown-button=\"D1\"]')", 'the arrow on D1')
    await snap('arrow')
  })
  await check('the arrow opens a searchable list; typing filters and Enter picks', async () => {
    await clickSelector('[data-dropdown-button="D1"]')
    await waitFor("document.querySelector('[data-validation-dropdown]')", 'the dropdown list')
    assert.deepEqual(await evaluate("[...document.querySelectorAll('[data-validation-dropdown] [role=\"option\"]')].map((item) => item.dataset.option)"), ['Low', 'Medium', 'High'])
    await call('Input.insertText', { text: 'hi' })
    await waitFor("document.querySelectorAll('[data-validation-dropdown] [role=\"option\"]').length === 1", 'the filtered list')
    assert.equal(await activeOption(), 'High')
    await snap('search')
    await press('Enter')
    await waitFor("!document.querySelector('[data-validation-dropdown]')", 'the list to close')
    assert.equal((await cell('D1'))?.value, 'High')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'focus back on the grid')
  })
  await check('Alt+Down opens the list at the current value; Esc cancels', async () => {
    await press('ArrowDown', ['alt'])
    await waitFor("document.querySelector('[data-validation-dropdown]')", 'the dropdown list (Alt+Down)')
    assert.equal(await activeOption(), 'High', 'the current value is highlighted')
    await press('ArrowUp')
    assert.equal(await activeOption(), 'Medium')
    await press('Escape')
    await waitFor("!document.querySelector('[data-validation-dropdown]')", 'the list to close')
    assert.equal((await cell('D1'))?.value, 'High', 'Esc changes nothing')
  })
  await check('keyboard picks from a literal list store numbers', async () => {
    await goTo('G1')
    await press('ArrowDown', ['alt'])
    await waitFor("document.querySelector('[data-validation-dropdown]')", 'the dropdown list')
    await press('ArrowDown')
    await press('ArrowDown')
    assert.equal(await activeOption(), '3')
    await press('Enter')
    await waitFor("!document.querySelector('[data-validation-dropdown]')", 'the list to close')
    assert.equal((await cell('G1'))?.value, 3, 'the item 3 is the number 3')
  })
  await check('a mouse pick from a list of cells stores the cell value and its format', async () => {
    await goTo('F1')
    await clickSelector('[data-dropdown-button="F1"]')
    await waitFor("document.querySelector('[data-validation-dropdown] [data-option=\"$20.00\"]')", 'the $20.00 option')
    await clickSelector('[data-validation-dropdown] [data-option="$20.00"]')
    await waitFor("!document.querySelector('[data-validation-dropdown]')", 'the list to close')
    const picked = await cell('F1')
    assert.equal(picked?.value, 20)
    assert.equal(picked?.numFmt, '$#,##0.00')
    assert.equal(await evaluate("document.querySelector('[data-cell-address=\"F1\"] .cell-content')?.textContent"), '$20.00', 'shown with its number format')
  })

  // ---- CALC-029: drag the selection border ------------------------------------------------------
  await check('dragging the border moves the block and references follow', async () => {
    await goTo('A1:A3')
    const preview = await dragBorder('left', 'I6')
    assert.equal(preview, 'I5:I7', 'the outline follows the pointer, the grabbed cell under it')
    await waitFor("window.__calcQA.workbook().sheets[0].cells.I5?.value === 1", 'the moved block')
    assert.deepEqual([(await cell('I5'))?.value, (await cell('I6'))?.value, (await cell('I7'))?.value], [1, 2, 3])
    assert.equal(await cell('A1'), null, 'the source is emptied')
    assert.equal((await cell('B1'))?.formula, 'I5*10')
    assert.equal((await cell('C1'))?.formula, 'SUM(I5:I7)')
    assert.equal(await evaluate("document.querySelector('.name-box')?.value"), 'I5:I7', 'the selection travels with the block')
    assert.ok(await evaluate(`(() => {
      const outline = document.querySelector('.selection-outline[data-selection-pane="body"]').getBoundingClientRect()
      const first = document.querySelector('[data-cell-address="I5"][data-cell-pane="body"]').getBoundingClientRect()
      const last = document.querySelector('[data-cell-address="I7"][data-cell-pane="body"]').getBoundingClientRect()
      return Math.abs(outline.top - first.top) < 3 && Math.abs(outline.bottom - last.bottom) < 3 && Math.abs(outline.left - first.left) < 3
    })()`), 'the selection is the moved block I5:I7')
    await waitFor("window.__calcQA.value(window.__calcQA.workbook().sheets[0].id, 'C1') === 6", 'the formulas to recalculate')
    await snap('moved')
  })
  await check('one undo step puts the block and the formulas back', async () => {
    await press('z', ['ctrl'])
    await waitFor("window.__calcQA.workbook().sheets[0].cells.A1?.value === 1", 'the undone move')
    assert.equal(await cell('I5'), null)
    assert.equal((await cell('B1'))?.formula, 'A1*10')
    assert.equal((await cell('C1'))?.formula, 'SUM(A1:A3)')
  })
  await check('Ctrl-drag copies, shifting relative references', async () => {
    await goTo('B1')
    const preview = await dragBorder('top', 'B10', ['ctrl'])
    assert.equal(preview, 'B10')
    await waitFor("window.__calcQA.workbook().sheets[0].cells.B10?.formula === 'A10*10'", 'the copied formula')
    assert.equal((await cell('B1'))?.formula, 'A1*10', 'the source stays')
  })
  await check('dropping on data asks first; Cancel leaves everything', async () => {
    await goTo('A1')
    await dragBorder('top', 'K1')
    await waitFor("/already data here/.test(document.querySelector('[role=\"alertdialog\"]')?.textContent || '')", "Excel's question")
    await chooseAlertButton('Cancel')
    assert.equal((await cell('K1'))?.value, 'keep')
    assert.equal((await cell('A1'))?.value, 1)
    await goTo('A1')
    await dragBorder('top', 'K1')
    await chooseAlertButton('Continue')
    await waitFor("window.__calcQA.workbook().sheets[0].cells.K1?.value === 1", 'the replaced cell')
    assert.equal((await cell('B1'))?.formula, 'K1*10')
  })
  await check('a click on the border without dragging just selects', async () => {
    await goTo('D5')
    await clickSelector('.selection-outline[data-selection-pane="body"] [data-move-edge="right"]')
    assert.equal(await evaluate("Boolean(document.querySelector('[role=\"alertdialog\"]'))"), false)
    assert.equal((await cell('D1'))?.value, 'High', 'nothing moved')
  })

  // ---- CALC-014: View > Page break preview ------------------------------------------------------
  await check('View > Page break preview outlines and numbers the printed pages', async () => {
    await clickSelector('[data-menu-trigger="view"]')
    await waitFor("document.querySelector('[data-menu-action=\"view-page-break-preview\"]')", 'the View menu')
    await clickSelector('[data-menu-action="view-page-break-preview"]')
    await waitFor("document.querySelector('.page-preview-page[data-preview-page=\"1\"]')", 'page 1 outlined', 30_000)
    await waitFor("document.querySelector('.page-preview-page[data-preview-page=\"2\"]')", 'page 2 outlined (140 rows)', 10_000)
    await snap('page-break-preview')
    await clickSelector('[data-menu-trigger="view"]')
    await waitFor("document.querySelector('[data-menu-action=\"view-page-break-preview\"]')", 'the View menu')
    await clickSelector('[data-menu-action="view-page-break-preview"]')
    await waitFor("!document.querySelector('.page-preview-page')", 'the preview to turn off')
  })

  assert.deepEqual(runtimeErrors, [], 'no uncaught exceptions in the renderer')
  console.log(`Interactions UI QA passed: ${passed} checks (list dropdown arrow, search, keyboard and mouse picks with typed values, Esc; drag-to-move with references and undo, Ctrl-drag copy, overwrite question, border click; page break preview).`)
} finally {
  try { socket.close() } catch {}
  stopChildren()
}
