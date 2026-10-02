/*
 * Data operations driven through the real app in Electron over the Chrome DevTools Protocol:
 * visible-cells-only commands in filtered lists, filters across row inserts, Fill Down / Right,
 * the fill handle (double-click, Ctrl+drag, Auto Fill Options), dropdown picks, the Sort
 * Warning, Alt+Enter wrapping, undo/redo polish, and duplicating a sheet with a table.
 *
 * The script starts its own Vite dev server and an off-screen Electron window with an isolated
 * profile and an in-memory clipboard (the system clipboard is never touched), runs the checks,
 * and stops both process trees. No network access is used beyond the local dev server.
 *
 *   npm run test:data-ops-ui
 *   QA_WORK_DIR=<dir> QA_VITE_PORT=5297 QA_CDP_PORT=9397 node scripts/qa-data-ops-ui.mjs
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
const vitePort = Number(process.env.QA_VITE_PORT || 5297)
const cdpPort = Number(process.env.QA_CDP_PORT || 9397)
const workDir = path.resolve(process.env.QA_WORK_DIR || path.join(root, 'tmp', 'data-ops-ui'))
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
const fixturePath = path.join(workDir, 'data-ops-fixture.xlsx')
const JAN_15_2026 = 46037
{
  const book = new ExcelJS.Workbook()
  const data = book.addWorksheet('Data', { views: [{ state: 'normal', activeCell: 'A1' }] })
  data.getCell('H10').value = new Date(Date.UTC(2026, 0, 15))
  data.getCell('H10').numFmt = 'm/d/yyyy'
  for (let row = 20; row <= 29; row += 1) data.getCell(`J${row}`).value = row - 19
  data.getCell('K20').value = { formula: 'J20*2' }
  data.getCell('M20').value = 5
  for (const [index, amount] of [10, 20, 30].entries()) {
    data.getCell(`P${index + 1}`).value = amount
    data.getCell(`P${index + 1}`).numFmt = '$#,##0.00'
  }
  for (let row = 1; row <= 3; row += 1) data.getCell(`N${row}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['$P$1:$P$3'] }
  data.getCell('N5').dataValidation = { type: 'list', allowBlank: true, formulae: ['"1,2,3"'] }
  const sortRows = [['Name', 'Qty', 'Price'], ['b', 2, 20], ['a', 1, 10], ['d', 4, 40], ['c', 3, 30]]
  sortRows.forEach((values, offset) => values.forEach((value, column) => { data.getCell(offset + 1, 28 + column).value = value }))
  const second = book.addWorksheet('Second')
  second.getCell('A1').value = 'second'
  const filtered = book.addWorksheet('Filter')
  const listRows = [['Region', 'Sales', 'Note', 'Double'], ['East', 10, 'a'], ['West', 20, 'b'], ['East', 30, 'c'], ['West', 40, 'd'], ['East', 50, 'e'], ['West', 60, 'f']]
  listRows.forEach((values, offset) => {
    values.forEach((value, column) => { filtered.getCell(offset + 1, column + 1).value = value })
    if (offset) filtered.getCell(offset + 1, 4).value = { formula: `B${offset + 1}*2` }
  })
  const tableSheet = book.addWorksheet('Tbl')
  tableSheet.addTable({ name: 'Sales', ref: 'A1', headerRow: true, columns: [{ name: 'Name' }, { name: 'Qty' }], rows: [['a', 5], ['b', 7]] })
  tableSheet.getCell('D1').value = { formula: 'SUM(Sales[Qty])' }
  await book.xlsx.writeFile(fixturePath)
}

// ---- Processes ---------------------------------------------------------------------------
const harnessPath = path.join(workDir, 'electron-harness.cjs')
await fs.writeFile(harnessPath, `'use strict'
const { app, clipboard } = require('electron')
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

await call('Runtime.enable')
await call('Page.enable')
await call('Emulation.setFocusEmulationEnabled', { enabled: true })
await waitFor('window.__calcQA && document.readyState === "complete"', 'the app and its QA hook')

// ---- Input helpers -------------------------------------------------------------------------
const KEYS = {
  Enter: { code: 'Enter', vk: 13 }, Tab: { code: 'Tab', vk: 9 }, Escape: { code: 'Escape', vk: 27 }, Backspace: { code: 'Backspace', vk: 8 }, Delete: { code: 'Delete', vk: 46 },
  ArrowUp: { code: 'ArrowUp', vk: 38 }, ArrowDown: { code: 'ArrowDown', vk: 40 }, ArrowLeft: { code: 'ArrowLeft', vk: 37 }, ArrowRight: { code: 'ArrowRight', vk: 39 }, F5: { code: 'F5', vk: 116 },
}
const MODIFIERS = { alt: 1, ctrl: 2, meta: 4, shift: 8 }
function modifierMask(modifiers = []) { return modifiers.reduce((mask, name) => mask | MODIFIERS[name], 0) }
function keyInfo(key) {
  if (KEYS[key]) return { key, ...KEYS[key] }
  if (/^[a-z]$/i.test(key)) return { key, code: `Key${key.toUpperCase()}`, vk: key.toUpperCase().charCodeAt(0) }
  if (/^[0-9]$/.test(key)) return { key, code: `Digit${key}`, vk: key.charCodeAt(0) }
  const punctuation = { ' ': ['Space', 32], '.': ['Period', 190], ',': ['Comma', 188], '=': ['Equal', 187], '*': ['Digit8', 56], '(': ['Digit9', 57], ')': ['Digit0', 48], ':': ['Semicolon', 186], '-': ['Minus', 189], '/': ['Slash', 191] }
  if (punctuation[key]) return { key, code: punctuation[key][0], vk: punctuation[key][1] }
  return { key, code: '', vk: 0 }
}
async function press(key, modifiers = []) {
  const info = keyInfo(key)
  const mask = modifierMask(modifiers)
  const printable = info.key.length === 1 && !(mask & (MODIFIERS.ctrl | MODIFIERS.alt | MODIFIERS.meta))
  await call('Input.dispatchKeyEvent', {
    type: printable ? 'keyDown' : 'rawKeyDown',
    key: info.key,
    code: info.code,
    windowsVirtualKeyCode: info.vk,
    nativeVirtualKeyCode: info.vk,
    modifiers: mask,
    ...(printable ? { text: info.key, unmodifiedText: info.key } : {}),
  })
  await call('Input.dispatchKeyEvent', { type: 'keyUp', key: info.key, code: info.code, windowsVirtualKeyCode: info.vk, nativeVirtualKeyCode: info.vk, modifiers: mask })
  await settle()
}
async function type(text) { for (const character of text) await press(character, character === '*' || character === '(' || character === ')' || character === ':' ? ['shift'] : []) }
const activeAddress = () => evaluate("(document.querySelector('.sheet-viewport')?.getAttribute('aria-activedescendant') || '').split('-').pop()")
const nameBox = () => evaluate("document.querySelector('.name-box')?.value")
const sheetExpr = (sheetName) => sheetName
  ? `window.__calcQA.workbook().sheets.find((item) => item.name === ${JSON.stringify(sheetName)})`
  : '(() => { const book = window.__calcQA.workbook(); return book.sheets.find((item) => item.id === book.activeSheetId) })()'
const cell = (address, sheetName) => evaluate(`(() => { const sheet = ${sheetExpr(sheetName)}; return sheet ? (sheet.cells[${JSON.stringify(address)}] || null) : null })()`)
const sheetState = (sheetName) => evaluate(`(() => { const sheet = ${sheetExpr(sheetName)}; return sheet ? JSON.parse(JSON.stringify({ name: sheet.name, filter: sheet.filter, autoFilter: sheet.autoFilter, filteredRows: sheet.filteredRows, hiddenRows: sheet.hiddenRows, rowHeights: sheet.rowHeights, tables: sheet.tables })) : null })()`)
const activeSheetName = () => evaluate('(() => { const book = window.__calcQA.workbook(); return book.sheets.find((item) => item.id === book.activeSheetId)?.name })()')
const dirty = () => evaluate("Boolean(document.querySelector('.dirty-dot'))")
const statusStat = (key) => evaluate(`document.querySelector('[data-status-stat="${key}"] strong')?.textContent.trim() ?? null`)
async function goTo(reference) {
  await press('F5')
  await waitFor("document.activeElement === document.querySelector('.name-box')", 'the name box (F5)')
  await call('Input.insertText', { text: reference })
  await press('Enter')
  await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", `the grid after Go To ${reference}`)
}
async function enter(reference, text) {
  await goTo(reference)
  await type(text)
  await press('Enter')
}
async function clickSelector(selector, filter) {
  const box = await evaluate(`(() => { const filter = ${filter ? String(filter) : 'null'}; const element = [...document.querySelectorAll(${JSON.stringify(selector)})].find((item) => !filter || filter(item)); if (!element) return null; element.scrollIntoView?.({ block: 'nearest', inline: 'nearest' }); const rect = element.getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()`)
  assert.ok(box, `${selector} is on screen`)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: box.x, y: box.y, button: 'left', buttons: 1, clickCount: 1 })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: box.x, y: box.y, button: 'left', buttons: 0, clickCount: 1 })
  await settle()
}
async function switchSheet(name) {
  await clickSelector('.sheet-tab', new Function('item', `return item.textContent.trim() === ${JSON.stringify(name)}`))
  await waitFor(`(() => { const book = window.__calcQA.workbook(); return book.sheets.find((item) => item.id === book.activeSheetId)?.name === ${JSON.stringify(name)} })()`, `sheet ${name}`)
  await evaluate("document.querySelector('.sheet-viewport')?.focus({ preventScroll: true })")
  await settle()
}
/** Center of the fill handle and of a cell, in viewport pixels. */
async function fillPoints(address) {
  return evaluate(`(() => {
    const handle = document.querySelector('.fill-handle')
    const element = [...document.querySelectorAll('[data-cell-address="${address}"]')].find((item) => item.dataset.cellPane === 'body') || document.querySelector('[data-cell-address="${address}"]')
    if (!handle || !element) return null
    const a = handle.getBoundingClientRect()
    const b = element.getBoundingClientRect()
    return { from: { x: a.left + a.width / 2, y: a.top + a.height / 2 }, to: { x: b.left + b.width / 2, y: b.top + b.height / 2 } }
  })()`)
}
async function dragFillHandle(address, modifiers = []) {
  const points = await fillPoints(address)
  assert.ok(points, `the fill handle and ${address} are on screen`)
  const mask = modifierMask(modifiers)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: points.from.x, y: points.from.y, button: 'left', buttons: 1, clickCount: 1, modifiers: mask })
  for (let step = 1; step <= 6; step += 1) {
    const progress = step / 6
    await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: points.from.x + (points.to.x - points.from.x) * progress, y: points.from.y + (points.to.y - points.from.y) * progress, button: 'left', buttons: 1, modifiers: mask })
    await pause(20)
  }
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: points.to.x, y: points.to.y, button: 'left', buttons: 0, clickCount: 1, modifiers: mask })
  await pause(120)
  await settle()
}
async function doubleClickFillHandle() {
  const box = await evaluate("(() => { const handle = document.querySelector('.fill-handle'); if (!handle) return null; const rect = handle.getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()")
  assert.ok(box, 'the fill handle is on screen')
  for (const clickCount of [1, 2]) {
    await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: box.x, y: box.y, button: 'left', buttons: 1, clickCount })
    await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: box.x, y: box.y, button: 'left', buttons: 0, clickCount })
  }
  await pause(120)
  await settle()
}
async function chooseAlertButton(label) {
  await waitFor(`[...document.querySelectorAll('[role="alertdialog"] button')].some((button) => button.textContent.trim() === ${JSON.stringify(label)})`, `the "${label}" button`)
  await evaluate(`[...document.querySelectorAll('[role="alertdialog"] button')].find((button) => button.textContent.trim() === ${JSON.stringify(label)}).click()`)
  await waitFor("!document.querySelector('[role=\"alertdialog\"]')", 'the alert to close')
  await settle()
}
async function pickDropdown(address, option) {
  const done = await evaluate(`(() => {
    const select = [...document.querySelectorAll('[data-cell-address="${address}"] select.cell-dropdown')].find((item) => !item.inert)
    if (!select) return false
    select.value = ${JSON.stringify(option)}
    select.dispatchEvent(new Event('change', { bubbles: true }))
    return true
  })()`)
  assert.ok(done, `the dropdown in ${address}`)
  await settle()
}
/** Data > Sort range A→Z. */
async function sortAscending() {
  await clickSelector('[data-menu-trigger]', (element) => element.textContent.trim() === 'Data')
  await waitFor("document.querySelector('[data-menu-action=\"data-sort-asc\"]')", 'the Data menu')
  await clickSelector('[data-menu-action="data-sort-asc"]')
}
async function check(label, run) {
  try {
    await run()
    passed += 1
    if (process.env.QA_VERBOSE) console.log(`ok - ${label}`)
  } catch (error) {
    console.error(`FAILED: ${label}`)
    throw error
  }
}

try {
  await evaluate(`window.__calcQA.openPath(${JSON.stringify(fixturePath)})`)
  await waitFor("document.querySelector('.sheet-viewport') && window.__calcQA.workbook()?.sheets?.length === 4", 'the fixture workbook')
  await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after open')

  // ---- Undo polish (CALC-004) -----------------------------------------------------------------
  await check('Undo brings the selection back to the change; Ctrl+Shift+Z redoes', async () => {
    await enter('V1', 'x')
    assert.equal((await cell('V1'))?.value, 'x')
    await goTo('Z50')
    await press('z', ['ctrl'])
    assert.equal((await cell('V1')), null, 'the entry is undone')
    assert.equal(await activeAddress(), 'V1', 'the selection returns to the undone cell')
    await press('z', ['ctrl', 'shift'])
    assert.equal((await cell('V1'))?.value, 'x', 'Ctrl+Shift+Z redoes (it used to undo)')
    assert.equal(await activeAddress(), 'Z50', 'redo returns to where the user was')
  })
  await check('Undoing back to the save point leaves the workbook clean', async () => {
    await press('s', ['ctrl'])
    await waitFor("/Saved/.test(document.querySelector('.toast')?.textContent || '') || document.querySelector('[role=\"alertdialog\"]')", 'the save', 20_000)
    if (await evaluate("Boolean(document.querySelector('[role=\"alertdialog\"]'))")) await chooseAlertButton('Continue')
    await waitFor("!document.querySelector('.dirty-dot')", 'a clean workbook after saving', 20_000)
    await enter('W1', 'y')
    assert.equal(await dirty(), true, 'an edit makes it dirty')
    await press('z', ['ctrl'])
    assert.equal(await dirty(), false, 'undoing to the save point makes it clean again')
    await press('y', ['ctrl'])
    assert.equal(await dirty(), true, 'redoing past the save point makes it dirty again')
    await press('z', ['ctrl'])
    await press('z', ['ctrl'])
    assert.equal(await dirty(), true, 'undoing before the save point is a change too')
    await press('y', ['ctrl'])
    assert.equal(await dirty(), false)
  })

  // ---- Fill Down / Fill Right copy (calc-grid-interaction-8) --------------------------------
  await check('Ctrl+D copies a date instead of making a series', async () => {
    await goTo('H10:H13')
    await press('d', ['ctrl'])
    for (const address of ['H11', 'H12', 'H13']) assert.equal((await cell(address))?.value, JAN_15_2026, `${address} holds the same date`)
  })
  await check('Ctrl+D on one row copies the row above; Ctrl+R on one column copies the column to the left', async () => {
    await goTo('H14')
    await press('d', ['ctrl'])
    assert.equal((await cell('H14'))?.value, JAN_15_2026)
    await goTo('I10')
    await press('r', ['ctrl'])
    assert.equal((await cell('I10'))?.value, JAN_15_2026)
    assert.equal((await cell('I10'))?.numFmt, 'm/d/yyyy')
  })

  // ---- Fill handle: double-click, Auto Fill Options, Ctrl+drag (CALC-006, CALC-019) ----------
  await check('Double-clicking the fill handle fills down beside the data', async () => {
    await goTo('K20')
    await doubleClickFillHandle()
    await waitFor(`${sheetExpr()}.cells.K29?.formula === 'J29*2'`, 'K21:K29 filled down', 5000)
    assert.equal((await cell('K30')), null, 'the fill stops with the data beside it')
    assert.equal(await nameBox(), 'K20:K29')
  })
  await check('Day names continue as a series; Auto Fill Options can copy instead (one undo step)', async () => {
    await enter('L20', 'Mon')
    await goTo('L20')
    await doubleClickFillHandle()
    await waitFor(`${sheetExpr()}.cells.L21?.value === 'Tue'`, 'the day series', 5000)
    assert.equal((await cell('L26'))?.value, 'Sun')
    assert.equal((await cell('L27'))?.value, 'Mon')
    await waitFor("document.querySelector('[data-autofill-options]')", 'the Auto Fill Options button')
    await clickSelector('[data-autofill-options]')
    await waitFor("document.querySelector('[data-menu-action=\"autofill-copy\"]')", 'the options menu')
    assert.ok(await evaluate("Boolean(document.querySelector('[data-menu-action=\"autofill-weekdays\"]')) === false"), 'no date options for day names')
    await clickSelector('[data-menu-action="autofill-copy"]')
    await waitFor(`${sheetExpr()}.cells.L21?.value === 'Mon'`, 'the copy replacing the series', 5000)
    assert.equal((await cell('L29'))?.value, 'Mon')
    await press('z', ['ctrl'])
    assert.equal((await cell('L21')), null, 'one Ctrl+Z removes the whole fill')
    assert.equal((await cell('L20'))?.value, 'Mon')
  })
  await check('Ctrl+drag on the fill handle turns a copied number into a series', async () => {
    await goTo('M20')
    await dragFillHandle('M23')
    assert.deepEqual([(await cell('M21'))?.value, (await cell('M23'))?.value], [5, 5], 'a lone number copies')
    await press('z', ['ctrl'])
    await goTo('M20')
    await dragFillHandle('M23', ['ctrl'])
    assert.deepEqual([(await cell('M21'))?.value, (await cell('M22'))?.value, (await cell('M23'))?.value], [6, 7, 8], 'with Ctrl it counts up')
  })

  // ---- Dropdown picks store typed values (calc-grid-interaction-9) ---------------------------
  await check('Picking from a dropdown stores numbers, not text', async () => {
    await goTo('N5')
    await pickDropdown('N5', '2')
    assert.equal((await cell('N5'))?.value, 2, 'the literal list item 2 is the number 2')
    await goTo('N1')
    await pickDropdown('N1', '$20.00')
    assert.equal((await cell('N1'))?.value, 20, 'a pick from cells stores the cell value')
    assert.equal((await cell('N1'))?.numFmt, '$#,##0.00', 'with its number format')
    await settle()
    assert.equal(await evaluate("[...document.querySelectorAll('[data-cell-address=\"N1\"] select.cell-dropdown')].find((item) => !item.inert)?.value"), '$20.00', 'the dropdown still shows the picked option')
    assert.equal(await evaluate("[...document.querySelectorAll('[data-cell-address=\"N5\"] select.cell-dropdown')].find((item) => !item.inert)?.value"), '2')
  })

  // ---- Sort Warning (CALC-001) ---------------------------------------------------------------
  await check('Sorting one column beside data asks first; Expand keeps the rows together', async () => {
    await goTo('AC1:AC5')
    await sortAscending()
    await waitFor("document.querySelector('[role=\"alertdialog\"]')?.textContent.includes('Expand the selection')", 'the Sort Warning')
    assert.equal(await evaluate("document.activeElement?.textContent.trim()"), 'Expand the selection', 'Expand is the default')
    await chooseAlertButton('Expand the selection')
    assert.deepEqual([(await cell('AB2'))?.value, (await cell('AC2'))?.value, (await cell('AD2'))?.value], ['a', 1, 10])
    assert.deepEqual([(await cell('AB5'))?.value, (await cell('AC5'))?.value, (await cell('AD5'))?.value], ['d', 4, 40])
    assert.equal((await cell('AB1'))?.value, 'Name', 'the header row stays on top')
  })
  await check('Continue with the current selection sorts only the column; Cancel sorts nothing', async () => {
    for (const [address, value] of [['AD2', '40'], ['AD3', '30'], ['AD4', '20'], ['AD5', '10']]) await enter(address, value)
    await goTo('AD2:AD5')
    await sortAscending()
    await chooseAlertButton('Cancel')
    assert.equal((await cell('AD2'))?.value, 40, 'cancelled: nothing moved')
    await goTo('AD2:AD5')
    await sortAscending()
    await chooseAlertButton('Continue with the current selection')
    assert.deepEqual([(await cell('AD2'))?.value, (await cell('AD5'))?.value], [10, 40])
    assert.deepEqual([(await cell('AB2'))?.value, (await cell('AC2'))?.value], ['a', 1], 'other columns stay put')
  })

  // ---- Alt+Enter wraps and grows the row (CALC-011) -----------------------------------------
  await check('Alt+Enter line breaks turn on Wrap Text and grow the row', async () => {
    await goTo('T2')
    await type('Line1')
    await press('Enter', ['alt'])
    await type('Line2')
    await press('Enter')
    const stored = await cell('T2')
    assert.equal(stored?.value, 'Line1\nLine2')
    assert.equal(stored?.style?.alignment?.wrapText, true)
    const rows = (await sheetState()).rowHeights
    assert.ok(Number(rows?.['2']) > 20, `row 2 grew to show both lines (${rows?.['2']})`)
  })

  // ---- Duplicate a sheet with a table (calc-file-io-objects-10) ------------------------------
  await check('Duplicating a sheet gives its table a unique name', async () => {
    await switchSheet('Tbl')
    await evaluate("document.querySelector('button[aria-label=\"Duplicate active sheet\"]').click()")
    await waitFor("window.__calcQA.workbook().sheets.some((item) => item.name === 'Tbl copy')", 'the copy')
    assert.equal(await activeSheetName(), 'Tbl copy')
    const copy = await sheetState('Tbl copy')
    assert.equal(copy.tables[0].name, 'Sales2')
    assert.equal((await cell('D1', 'Tbl copy'))?.formula, 'SUM(Sales2[Qty])')
    assert.equal((await cell('D1', 'Tbl'))?.formula, 'SUM(Sales[Qty])')
    await press('z', ['ctrl'])
    await waitFor("!window.__calcQA.workbook().sheets.some((item) => item.name === 'Tbl copy')", 'the copy undone')
    assert.equal(await activeSheetName(), 'Tbl', 'undo returns to the sheet that was active')
  })

  // ---- Filtered lists: visible cells only (calc-grid-interaction-2, CALC-002, -5) ------------
  await check('A filter hides the West rows', async () => {
    await switchSheet('Filter')
    await goTo('A1')
    await press('l', ['ctrl', 'shift'])
    await waitFor(`${sheetExpr('Filter')}.filter?.ref === 'A1:D7'`, 'the filter range')
    await clickSelector('[data-filter-column="A"][data-filter-key="sheet"]')
    await waitFor("document.querySelector('.dt-filter-menu')", 'the filter menu')
    await evaluate("[...document.querySelectorAll('.dt-values-row')].find((row) => row.textContent.includes('West')).querySelector('input').click()")
    await settle()
    await evaluate("document.querySelector('.dt-filter-footer .dt-button.is-primary').click()")
    await waitFor(`JSON.stringify(${sheetExpr('Filter')}.filteredRows) === '[3,5,7]'`, 'rows 3, 5 and 7 filtered out')
  })
  await check('The status bar sums the visible rows only', async () => {
    await goTo('B2:B7')
    await waitFor(`document.querySelector('[data-status-stat="sum"] strong')?.textContent.trim() === '90'`, 'Sum 90 (10+30+50)', 3000)
  })
  await check('Delete clears the visible rows only', async () => {
    await goTo('C2:C7')
    await press('Delete')
    assert.deepEqual([(await cell('C2'))?.value, (await cell('C4'))?.value, (await cell('C6'))?.value], [undefined, undefined, undefined])
    assert.deepEqual([(await cell('C3'))?.value, (await cell('C5'))?.value, (await cell('C7'))?.value], ['b', 'd', 'f'], 'hidden rows keep their data')
    await press('z', ['ctrl'])
    assert.equal((await cell('C2'))?.value, 'a')
  })
  await check('Ctrl+D and Ctrl+Enter fill the visible rows only', async () => {
    await enter('C2', 'Done')
    await goTo('C2:C7')
    await press('d', ['ctrl'])
    assert.deepEqual([(await cell('C4'))?.value, (await cell('C6'))?.value], ['Done', 'Done'])
    assert.deepEqual([(await cell('C3'))?.value, (await cell('C5'))?.value, (await cell('C7'))?.value], ['b', 'd', 'f'])
    await press('z', ['ctrl'])
    await press('z', ['ctrl'])
    await goTo('C2:C7')
    await type('Z')
    await press('Enter', ['ctrl'])
    assert.deepEqual([(await cell('C2'))?.value, (await cell('C6'))?.value, (await cell('C5'))?.value], ['Z', 'Z', 'd'])
    await press('z', ['ctrl'])
  })
  await check('Copy takes the visible rows; their formulas paste from where they were', async () => {
    await goTo('A1:D7')
    await press('c', ['ctrl'])
    await goTo('F10')
    await press('v', ['ctrl'])
    assert.deepEqual([(await cell('F10'))?.value, (await cell('F11'))?.value, (await cell('F12'))?.value, (await cell('F13'))?.value], ['Region', 'East', 'East', 'East'])
    assert.equal((await cell('G12'))?.value, 30)
    assert.equal((await cell('I12'))?.formula, 'G12*2', 'B4*2 copied from row 4 lands on row 12 as G12*2')
    assert.equal((await cell('F14')), null, 'nothing from the hidden rows')
    await press('z', ['ctrl'])
  })
  await check('Sorting a filtered column from its menu keeps the hidden rows where they are', async () => {
    await clickSelector('[data-filter-column="B"][data-filter-key="sheet"]')
    await waitFor("document.querySelector('.dt-filter-menu')", 'the filter menu')
    await evaluate("document.querySelectorAll('.dt-filter-menu .dt-menu-group')[0].querySelectorAll('.dt-menu-item')[1].click()")
    await waitFor(`${sheetExpr('Filter')}.cells.B2?.value === 50`, 'the visible rows sorted largest first')
    assert.deepEqual([(await cell('B4'))?.value, (await cell('B6'))?.value], [30, 10])
    assert.deepEqual([(await cell('A3'))?.value, (await cell('B3'))?.value, (await cell('B5'))?.value, (await cell('B7'))?.value], ['West', 20, 40, 60], 'hidden rows stay put')
    await evaluate("document.querySelector('.sheet-viewport')?.focus({ preventScroll: true })")
    await press('z', ['ctrl'])
    assert.equal((await cell('B2'))?.value, 10)
  })
  await check('Inserting a row above moves the filter with its rows; removing it shows every row', async () => {
    // Shift+Space selects row 1; Ctrl+Shift+= inserts a row above it.
    await goTo('A1')
    await press(' ', ['shift'])
    await press('=', ['ctrl', 'shift'])
    await waitFor(`${sheetExpr('Filter')}.filter?.ref === 'A2:D8'`, 'the filter range below the new row')
    const state = await sheetState('Filter')
    assert.deepEqual(state.filteredRows, [4, 6, 8])
    assert.deepEqual(state.hiddenRows, [4, 6, 8])
    await press('l', ['ctrl', 'shift'])
    await waitFor(`!${sheetExpr('Filter')}.filter`, 'the filter removed')
    assert.deepEqual((await sheetState('Filter')).hiddenRows, [], 'every row shows again')
  })

  assert.deepEqual(runtimeErrors, [], 'no uncaught exceptions in the renderer')
  console.log(`Data operations UI QA passed: ${passed} checks (undo selection, Ctrl+Shift+Z, clean at the save point, Fill Down/Right copies, fill handle double-click, named series and Auto Fill Options, Ctrl+drag, dropdown picks, Sort Warning, Alt+Enter wrap, duplicate sheet tables, filtered Delete/Ctrl+D/Ctrl+Enter/Copy/status bar, filter-menu sort keeps hidden rows, filters across row inserts).`)
} finally {
  try { socket.close() } catch {}
  stopChildren()
}
