/*
 * Excel/Sheets polish driven through the real app in Electron over the Chrome DevTools
 * Protocol: links to places in the workbook (cells, HYPERLINK(), names, Ctrl+K), the "There's a
 * problem with this formula" check, error explanations and quiet tooltips, AutoComplete and
 * Alt+Down, Ctrl+mouse-wheel zoom kept per sheet and saved, Edit/Help menus with the keyboard
 * reference and an arrow-key menu search, manual calculation with F9, the circular-reference
 * indicator, sort by colour, the chart range picker, and File > Page setup.
 *
 * The script starts its own Vite dev server and an off-screen Electron window with an isolated
 * profile and an in-memory clipboard (the system clipboard is never touched), runs the checks,
 * and stops both process trees. No network access is used beyond the local dev server; no
 * link is opened outside the app.
 *
 *   npm run test:polish-ui
 *   QA_WORK_DIR=<dir> QA_VITE_PORT=5303 QA_CDP_PORT=9403 node scripts/qa-polish-ui.mjs
 *   QA_SCREENSHOTS=1 also saves polish-*.png screenshots of each new surface into QA_WORK_DIR.
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
const vitePort = Number(process.env.QA_VITE_PORT || 5303)
const cdpPort = Number(process.env.QA_CDP_PORT || 9403)
const workDir = path.resolve(process.env.QA_WORK_DIR || path.join(root, 'tmp', 'polish-ui'))
const deadline = Date.now() + Number(process.env.QA_TIMEOUT_MS || 360_000)
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
// Built with ExcelJS, then written by the app's own serializer so the in-workbook links are
// stored exactly as a file from Excel stores them (<hyperlink location="...">).
await fs.mkdir(workDir, { recursive: true })
const fixturePath = path.join(workDir, 'polish-fixture.xlsx')
{
  const { workbookPayloadFromBytes, serializeWorkbook } = require('../electron/workbooks.cjs')
  const book = new ExcelJS.Workbook()
  const summary = book.addWorksheet('Summary')
  summary.getCell('A1').value = 'Go to detail'
  summary.getCell('A2').value = { formula: 'HYPERLINK("#\'Data detail\'!B2","Back to B2")', result: 'Back to B2' }
  summary.getCell('A3').value = 'To the name'
  summary.getCell('A4').value = 'Broken link'
  for (const [index, text] of ['Fruit', 'Apple', 'Banana', 'Cherry'].entries()) summary.getCell(`C${index + 1}`).value = text
  const colors = [null, 'FFFFFF00', null, 'FFFFFF00']
  summary.getCell('G1').value = 'Item'
  for (const [index, text] of ['one', 'two', 'three', 'four'].entries()) {
    const cell = summary.getCell(`G${index + 2}`)
    cell.value = text
    if (colors[index]) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: colors[index] } }
  }
  summary.autoFilter = 'G1:G5'
  for (const [row, values] of [[1, ['Month', 'Sales']], [2, ['Jan', 10]], [3, ['Feb', 14]], [4, ['Mar', 9]]]) {
    summary.getCell(`J${row}`).value = values[0]
    summary.getCell(`K${row}`).value = values[1]
  }
  const detail = book.addWorksheet('Data detail')
  detail.getCell('B2').value = 'B2 here'
  detail.getCell('C5').value = 'C5 here'
  detail.getCell('D4').value = 'D4 here'
  book.definedNames.add("'Data detail'!$D$4", 'TargetCell')
  const base = Buffer.from(await book.xlsx.writeBuffer())
  const model = (await workbookPayloadFromBytes('polish-fixture.xlsx', base)).workbook
  const cells = model.sheets[0].cells
  cells.A1 = { ...cells.A1, hyperlink: "#'Data detail'!C5" }
  cells.A3 = { ...cells.A3, hyperlink: '#TargetCell' }
  cells.A4 = { ...cells.A4, hyperlink: '#NoSuchSheet!A1' }
  await fs.writeFile(fixturePath, await serializeWorkbook(model, 'xlsx', { baseBytes: base }))
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
// Nothing leaves the app during the run: an external link would only be recorded.
try {
  Object.defineProperty(shell, 'openExternal', { value: async (url) => { process.stdout.write('QA-OPEN-EXTERNAL ' + url + '\\n') }, configurable: true, writable: true })
} catch (error) {
  process.stderr.write('QA: could not stub shell.openExternal: ' + error + '\\n')
}
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
const externalOpens = []
electron.stdout.on('data', (chunk) => {
  for (const line of String(chunk).split(/\r?\n/)) if (line.startsWith('QA-OPEN-EXTERNAL ')) externalOpens.push(line.slice(17))
})
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
  const focused = await evaluate("(() => { const element = document.activeElement; return element ? `${element.tagName.toLowerCase()}${element.className ? '.' + String(element.className).replace(/\\s+/g, '.') : ''}${element.getAttribute('aria-label') ? `[${element.getAttribute('aria-label')}]` : ''}` : 'none' })()").catch(() => 'unknown')
  throw new Error(`Timed out waiting for ${label} (focus on ${focused})`)
}
const settle = () => evaluate('new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve(true))))')
/** QA_SCREENSHOTS=1 saves a PNG of the window at each named step (for reviewing the look). */
async function snap(name) {
  if (!process.env.QA_SCREENSHOTS) return
  await settle()
  const shot = await call('Page.captureScreenshot', { format: 'png' })
  await fs.writeFile(path.join(workDir, `polish-${name}.png`), Buffer.from(shot.data, 'base64'))
}

await call('Runtime.enable')
await call('Page.enable')
await call('Emulation.setFocusEmulationEnabled', { enabled: true })
await waitFor('window.__calcQA && document.readyState === "complete"', 'the app and its QA hook')

// ---- Input helpers -------------------------------------------------------------------------
const KEYS = {
  Enter: { code: 'Enter', vk: 13 }, Tab: { code: 'Tab', vk: 9 }, Escape: { code: 'Escape', vk: 27 }, Backspace: { code: 'Backspace', vk: 8 }, Delete: { code: 'Delete', vk: 46 },
  ArrowUp: { code: 'ArrowUp', vk: 38 }, ArrowDown: { code: 'ArrowDown', vk: 40 }, ArrowLeft: { code: 'ArrowLeft', vk: 37 }, ArrowRight: { code: 'ArrowRight', vk: 39 },
  F5: { code: 'F5', vk: 116 }, F9: { code: 'F9', vk: 120 },
}
const MODIFIERS = { alt: 1, ctrl: 2, meta: 4, shift: 8 }
function modifierMask(modifiers = []) { return modifiers.reduce((mask, name) => mask | MODIFIERS[name], 0) }
function keyInfo(key) {
  if (KEYS[key]) return { key, ...KEYS[key] }
  if (/^[a-z]$/i.test(key)) return { key, code: `Key${key.toUpperCase()}`, vk: key.toUpperCase().charCodeAt(0) }
  if (/^[0-9]$/.test(key)) return { key, code: `Digit${key}`, vk: key.charCodeAt(0) }
  const punctuation = { ' ': ['Space', 32], '.': ['Period', 190], ',': ['Comma', 188], '=': ['Equal', 187], '+': ['Equal', 187], '*': ['Digit8', 56], '(': ['Digit9', 57], ')': ['Digit0', 48], ':': ['Semicolon', 186], '-': ['Minus', 189], '!': ['Digit1', 49], '/': ['Slash', 191] }
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
async function type(text) { for (const character of text) await press(character, '*()+:!'.includes(character) ? ['shift'] : []) }
const activeAddress = () => evaluate("(document.querySelector('.sheet-viewport')?.getAttribute('aria-activedescendant') || '').split('-').pop()")
const nameBox = () => evaluate("document.querySelector('.name-box')?.value")
const editorOpen = () => evaluate("Boolean(document.querySelector('textarea.cell-editor'))")
const activeSheetName = () => evaluate('(() => { const book = window.__calcQA.workbook(); return book.sheets.find((item) => item.id === book.activeSheetId)?.name })()')
const sheetOf = (name) => `window.__calcQA.workbook().sheets.find((item) => item.name === ${JSON.stringify(name)})`
const cell = (address, sheetName = 'Summary') => evaluate(`(() => { const sheet = ${sheetOf(sheetName)}; return sheet ? (sheet.cells[${JSON.stringify(address)}] || null) : null })()`)
const value = (address, sheetName = 'Summary') => evaluate(`(() => { const sheet = ${sheetOf(sheetName)}; return sheet ? window.__calcQA.value(sheet.id, ${JSON.stringify(address)}) : undefined })()`)
const zoomLabel = () => evaluate("document.querySelector('.zoom-control > span')?.textContent")
const toastText = () => evaluate("document.querySelector('.toast > span')?.textContent || ''")
async function goTo(reference) {
  await press('F5')
  await waitFor("document.activeElement === document.querySelector('.name-box')", 'the name box (F5)')
  await call('Input.insertText', { text: reference })
  await press('Enter')
  await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", `the grid after Go To ${reference}`)
}
async function pointAt(selector, filter) {
  const box = await evaluate(`(() => { const filter = ${filter ? String(filter) : 'null'}; const element = [...document.querySelectorAll(${JSON.stringify(selector)})].find((item) => !filter || filter(item)); if (!element) return null; element.scrollIntoView?.({ block: 'nearest', inline: 'nearest' }); const rect = element.getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()`)
  assert.ok(box, `${selector} is on screen`)
  return box
}
/** A real mouse click (it moves focus, unlike element.click()). */
async function clickSelector(selector, filter, modifiers = []) {
  const box = await pointAt(selector, filter)
  const mask = modifierMask(modifiers)
  await call('Input.dispatchMouseEvent', { type: 'mouseMoved', x: box.x, y: box.y, modifiers: mask })
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: box.x, y: box.y, button: 'left', buttons: 1, clickCount: 1, modifiers: mask })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: box.x, y: box.y, button: 'left', buttons: 0, clickCount: 1, modifiers: mask })
  await settle()
}
const clickCell = (address, modifiers = []) => clickSelector(`[data-cell-address="${address}"]`, null, modifiers)
const clickTab = (name) => clickSelector('.sheet-tab', `(element) => element.textContent.trim() === ${JSON.stringify(name)}`)
async function wheel(deltaY, modifiers = ['ctrl']) {
  const box = await evaluate("(() => { const rect = document.querySelector('.sheet-viewport').getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()")
  await call('Input.dispatchMouseEvent', { type: 'mouseWheel', x: box.x, y: box.y, deltaX: 0, deltaY, modifiers: modifierMask(modifiers) })
  await settle()
}
async function searchMenus(query) {
  await press('/', ['alt'])
  await waitFor("document.activeElement?.getAttribute('aria-label') === 'Search menus'", 'the menu search (Alt+/)')
  await call('Input.insertText', { text: query })
  await settle()
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
  await waitFor("document.querySelector('.sheet-viewport') && window.__calcQA.workbook()?.sheets?.length === 2", 'the fixture workbook')
  await settle()

  // ---- CALC-005: links to places in this workbook ------------------------------------------
  await check('a link to a cell on another sheet selects that cell there', async () => {
    await clickSelector('[data-cell-address="A1"] .cell-hyperlink')
    await waitFor(`(${sheetOf('Data detail')})?.id === window.__calcQA.workbook().activeSheetId`, 'the Data detail sheet')
    await waitFor("document.querySelector('.name-box')?.value === 'C5'", 'C5 selected')
    assert.deepEqual(externalOpens, [], 'nothing is handed to the browser')
  })
  await check('a HYPERLINK() formula cell follows its location', async () => {
    await clickTab('Summary')
    await waitFor("document.querySelector('[data-cell-address=\"A2\"] .cell-hyperlink')", 'the HYPERLINK cell drawn as a link')
    assert.equal(await evaluate("document.querySelector('[data-cell-address=\"A2\"] .cell-hyperlink').textContent"), 'Back to B2')
    await clickSelector('[data-cell-address="A2"] .cell-hyperlink')
    await waitFor(`(${sheetOf('Data detail')})?.id === window.__calcQA.workbook().activeSheetId && document.querySelector('.name-box')?.value === 'B2'`, 'B2 on Data detail')
  })
  await check('a link to a defined name selects the name’s cell', async () => {
    await clickTab('Summary')
    await clickSelector('[data-cell-address="A3"] .cell-hyperlink')
    await waitFor(`(${sheetOf('Data detail')})?.id === window.__calcQA.workbook().activeSheetId && document.querySelector('.name-box')?.value === 'D4'`, 'D4 through TargetCell')
  })
  await check('a link to a missing place says so instead of failing', async () => {
    await clickTab('Summary')
    await clickSelector('[data-cell-address="A4"] .cell-hyperlink')
    await waitFor("/Reference isn.t valid/.test(document.querySelector('.toast > span')?.textContent || '')", 'the invalid reference message')
    assert.equal(await activeSheetName(), 'Summary')
    assert.match(await evaluate("document.querySelector('[data-cell-address=\"A1\"]').title"), /Go to 'Data detail'!C5/, 'a link’s tooltip names its destination')
  })
  await check('Ctrl+K links a cell to a place in this workbook', async () => {
    await goTo('B6')
    await press('k', ['ctrl'])
    await waitFor("document.querySelector('.link-dialog')", 'the link dialog')
    await snap('link-dialog-web')
    await clickSelector('.link-dialog [role="radio"]', "(element) => element.textContent.includes('Place in this workbook')")
    await waitFor("document.activeElement?.getAttribute('aria-label') === 'Cell reference'", 'the cell reference field')
    await evaluate(`(() => {
      const select = document.querySelector('.link-dialog select')
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(select, 'sheet:Data detail')
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })()`)
    await evaluate("(() => { const field = document.querySelector('.link-dialog [aria-label=\"Cell reference\"]'); field.focus(); field.select() })()")
    await call('Input.insertText', { text: 'E7' })
    await evaluate("(() => { const field = document.querySelector('.link-dialog [aria-label=\"Text to display\"]'); field.focus(); field.select() })()")
    await call('Input.insertText', { text: 'Jump' })
    await snap('link-dialog-place')
    // (A synthetic Enter has no keypress, so it cannot submit a form: click OK.)
    await clickSelector('.link-dialog button[type="submit"]')
    await waitFor("!document.querySelector('.link-dialog')", 'the dialog to close')
    const linked = await cell('B6')
    assert.equal(linked?.hyperlink, "#'Data detail'!E7")
    assert.equal(linked?.value, 'Jump')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after the dialog')
    await clickSelector('[data-cell-address="B6"] .cell-hyperlink')
    await waitFor(`(${sheetOf('Data detail')})?.id === window.__calcQA.workbook().activeSheetId && document.querySelector('.name-box')?.value === 'E7'`, 'E7 through the new link')
    await clickTab('Summary')
  })

  // ---- CALC-007 / CALC-046: formula problems and error help -------------------------------
  await check('an unreadable formula keeps the editor open on the problem', async () => {
    await goTo('E10')
    await type('=1+*2')
    await press('Enter')
    await waitFor("document.querySelector('[data-formula-problem]')", 'the formula problem note')
    await snap('formula-problem')
    assert.match(await evaluate("document.querySelector('[data-formula-problem]').textContent"), /There.s a problem with this formula/)
    assert.equal(await editorOpen(), true, 'the editor stays open')
    assert.equal(await cell('E10'), null, 'nothing was entered')
    const selected = await evaluate("(() => { const editor = document.querySelector('textarea.cell-editor'); return { start: editor.selectionStart, end: editor.selectionEnd, value: editor.value } })()")
    assert.equal(selected.value, '=1+*2')
    assert.ok(selected.start >= 1 && selected.start <= 4, `the problem is highlighted (caret at ${selected.start})`)
    await press('Escape')
    assert.equal(await editorOpen(), false)
    assert.equal(await cell('E10'), null)
    assert.equal(await evaluate("Boolean(document.querySelector('[data-formula-problem]'))"), false)
  })
  await check('a missing parenthesis is offered as a correction', async () => {
    await goTo('E11')
    await type('=SUM(1,2')
    await press('Enter')
    await waitFor("document.querySelector('[role=\"alertdialog\"]')?.textContent.includes('=SUM(1,2)')", 'the correction offer')
    assert.equal(await evaluate("document.activeElement?.textContent.trim()"), 'Yes', 'accepting is the default button')
    await clickSelector('[role="alertdialog"] button', "(element) => element.textContent.trim() === 'Yes'")
    await waitFor("!document.querySelector('[role=\"alertdialog\"]')", 'the offer to close')
    assert.equal((await cell('E11'))?.formula, 'SUM(1,2)')
    assert.equal(await value('E11'), 3)
    await waitFor("document.querySelector('.name-box')?.value === 'E12'", 'the move Enter makes, after the correction')
  })
  await check('an unknown function still enters (#NAME?) and explains itself', async () => {
    await goTo('E12')
    await type('=SUMM(1)')
    await press('Enter')
    assert.equal((await cell('E12'))?.formula, 'SUMM(1)')
    assert.equal(await value('E12'), '#NAME?')
    await goTo('E12')
    // Review F11: selecting an error cell shows a small indicator; its card opens from it.
    await waitFor("document.querySelector('[data-error-indicator=\"#NAME?\"]')", 'the error indicator')
    assert.equal(await evaluate("Boolean(document.querySelector('[data-error-tip]'))"), false, 'no card covers the cells until asked for')
    await clickSelector('[data-error-indicator]')
    await waitFor("document.querySelector('[data-error-tip=\"#NAME?\"]')", 'the error card')
    assert.equal(await nameBox(), 'E12', 'the indicator does not move the selection')
    assert.match(await evaluate("document.querySelector('[data-error-tip]').textContent"), /SUMM/)
  })
  await check('error cells explain themselves; ordinary cells have no tooltip', async () => {
    await goTo('E13')
    await type('=1/0')
    await press('Enter')
    await goTo('E13')
    await clickSelector('[data-error-indicator="#DIV/0!"]')
    await waitFor("document.querySelector('[data-error-tip=\"#DIV/0!\"]')", 'the #DIV/0! card')
    await snap('error-card')
    assert.match(await evaluate("document.querySelector('[data-error-tip]').textContent"), /Divide by zero/)
    assert.match(await evaluate("document.querySelector('[data-cell-address=\"E13\"]').title"), /#DIV\/0!.*divides by zero/i)
    assert.equal(await evaluate("document.querySelector('[data-cell-address=\"C2\"]').hasAttribute('title')"), false, 'a text cell has no tooltip')
    assert.equal(await evaluate("document.querySelector('[data-cell-address=\"E11\"]').hasAttribute('title')"), false, 'a formula cell has no tooltip')
    await goTo('A20')
    assert.equal(await evaluate("Boolean(document.querySelector('[data-error-tip]') || document.querySelector('[data-error-indicator]'))"), false, 'the card goes with the selection')
  })

  // ---- CALC-017: AutoComplete and Alt+Down -------------------------------------------------
  await check('AutoComplete completes a unique entry of the column', async () => {
    await goTo('C5')
    await type('Ba')
    await snap('autocomplete')
    const editor = await evaluate("(() => { const editor = document.querySelector('textarea.cell-editor'); return { value: editor.value, start: editor.selectionStart, end: editor.selectionEnd } })()")
    assert.deepEqual(editor, { value: 'Banana', start: 2, end: 6 }, 'the suggestion follows the typed text, selected')
    await press('Enter')
    assert.equal((await cell('C5'))?.value, 'Banana')
  })
  await check('accepting a suggestion uses the entry’s own capitalisation', async () => {
    await goTo('C6')
    await type('ch')
    assert.equal(await evaluate("document.querySelector('textarea.cell-editor').value"), 'cherry')
    await press('Enter')
    assert.equal((await cell('C6'))?.value, 'Cherry')
  })
  await check('Backspace rejects a suggestion; numbers are never completed', async () => {
    await goTo('C7')
    await type('Ap')
    assert.equal(await evaluate("document.querySelector('textarea.cell-editor').value"), 'Apple')
    await press('Backspace')
    assert.equal(await evaluate("document.querySelector('textarea.cell-editor').value"), 'Ap')
    await press('Enter')
    assert.equal((await cell('C7'))?.value, 'Ap')
    await goTo('C8')
    await type('1')
    assert.equal(await evaluate("document.querySelector('textarea.cell-editor').value"), '1')
    await press('Escape')
  })
  await check('Alt+Down picks from the column’s entries, with type-to-filter', async () => {
    await goTo('C8')
    await press('ArrowDown', ['alt'])
    await waitFor("document.querySelector('.pick-list')", 'the pick list')
    await snap('pick-list')
    const items = await evaluate("[...document.querySelectorAll('.pick-list-item')].map((item) => item.textContent)")
    assert.deepEqual(items, ['Ap', 'Apple', 'Banana', 'Cherry', 'Fruit'])
    await call('Input.insertText', { text: 'che' })
    await waitFor("document.querySelectorAll('.pick-list-item').length === 1", 'the filtered list')
    await press('Enter')
    await waitFor("!document.querySelector('.pick-list')", 'the list to close')
    assert.equal((await cell('C8'))?.value, 'Cherry')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after the pick')
  })

  // ---- CALC-023: zoom ------------------------------------------------------------------------
  await check('Ctrl+mouse wheel zooms the sheet (not the page), in 15% steps', async () => {
    const before = await evaluate('({ dpr: window.devicePixelRatio, scale: window.visualViewport?.scale ?? 1 })')
    assert.equal(await zoomLabel(), '100%')
    await wheel(-100)
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '115%'", 'zoom 115%')
    await wheel(-100)
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '130%'", 'zoom 130%')
    await wheel(100)
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '115%'", 'back to 115%')
    const after = await evaluate('({ dpr: window.devicePixelRatio, scale: window.visualViewport?.scale ?? 1 })')
    assert.deepEqual(after, before, 'the page itself is not zoomed')
    await wheel(-100, [])
    assert.equal(await zoomLabel(), '115%', 'a plain wheel scrolls, it does not zoom')
  })
  await check('each sheet keeps its zoom', async () => {
    await clickTab('Data detail')
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '100%'", 'Data detail at 100%')
    await clickTab('Summary')
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '115%'", 'Summary back at 115%')
  })

  // ---- CALC-031: menus, keyboard reference, menu search ------------------------------------
  await check('the menu bar has Edit and Help', async () => {
    assert.deepEqual(await evaluate("[...document.querySelectorAll('[data-menu-trigger]')].map((item) => item.textContent.trim())"), ['File', 'Edit', 'View', 'Insert', 'Format', 'Data', 'Tools', 'Help'])
  })
  await check('Ctrl+/ opens the keyboard reference, searchable', async () => {
    await press('/', ['ctrl'])
    await waitFor("document.querySelector('.shortcuts-card')", 'the keyboard reference')
    await snap('shortcuts')
    await call('Input.insertText', { text: 'zoom' })
    await waitFor("document.querySelectorAll('.shortcut-row').length === 1", 'one matching shortcut')
    assert.match(await evaluate("document.querySelector('.shortcut-row').textContent"), /Zoom in or out/)
    await press('Escape')
    await waitFor("!document.querySelector('.shortcuts-card')", 'the reference to close')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after the reference')
  })
  await check('Alt+/ searches the menus and toolbar; arrows choose, Enter runs', async () => {
    await searchMenus('zoom')
    const first = await evaluate("document.querySelector('input[aria-label=\"Search menus\"]').getAttribute('aria-activedescendant')")
    assert.equal(first, 'command-search-palette-zoom-in', 'toolbar commands are searchable, best match first')
    await press('ArrowDown')
    await press('ArrowDown')
    assert.equal(await evaluate("document.querySelector('input[aria-label=\"Search menus\"]').getAttribute('aria-activedescendant')"), 'command-search-palette-zoom-reset')
    await press('Enter')
    await waitFor("!document.querySelector('.command-search-panel')", 'the search to close')
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '100%'", 'Zoom to 100% ran')
    await wheel(-100)
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '115%'", 'zoom 115% again')
  })

  // ---- CALC-027: calculation -------------------------------------------------------------------
  await check('manual calculation waits for F9; the status bar offers Calculate', async () => {
    await searchMenus('calculation manual')
    await press('Enter')
    await waitFor("window.__calcQA.workbook().metadata?.calcProperties?.calcMode === 'manual'", 'manual calculation')
    await goTo('H1')
    await type('5')
    await press('Enter')
    await goTo('H2')
    await type('=H1*2')
    await press('Enter')
    assert.equal(await value('H2'), 10)
    await goTo('H1')
    await type('7')
    await press('Enter')
    assert.equal(await value('H2'), 10, 'dependents wait in manual mode')
    await waitFor("document.querySelector('[data-status-calculate]')", 'the Calculate indicator')
    await press('F9')
    await waitFor(`window.__calcQA.value(${sheetOf('Summary')}.id, 'H2') === 14`, 'F9 recalculates')
    await waitFor("!document.querySelector('[data-status-calculate]')", 'the indicator to clear')
    await searchMenus('calculation automatic')
    await press('Enter')
    await waitFor("window.__calcQA.workbook().metadata?.calcProperties?.calcMode === 'auto'", 'automatic calculation again')
  })
  await check('a circular reference is named in the status bar and jumps to its cell', async () => {
    await goTo('H5')
    await type('=H5+1')
    await press('Enter')
    await waitFor("document.querySelector('[data-circular-reference]')", 'the circular reference indicator')
    await snap('circular-status')
    assert.match(await evaluate("document.querySelector('[data-circular-reference]').textContent"), /Circular references: H5/)
    await goTo('A20')
    await clickSelector('[data-circular-reference]')
    await waitFor("document.querySelector('.name-box')?.value === 'H5'", 'H5 selected from the status bar')
    await press('Delete')
    await waitFor("!document.querySelector('[data-circular-reference]')", 'the indicator to go')
  })

  // ---- CALC-010: sort by colour ----------------------------------------------------------------
  await check('Sort by color in the filter menu puts that color first', async () => {
    await clickSelector('.filter-button[data-filter-column="G"]')
    await waitFor("document.querySelector('.dt-filter-menu')", 'the filter menu')
    await clickSelector('.dt-filter-menu .dt-disclosure-toggle', "(element) => element.textContent.includes('Sort by color')")
    await clickSelector('.dt-filter-menu .dt-color-row', "(element) => !/no fill|automatic/i.test(element.textContent) && element.closest('.dt-disclosure')?.textContent.includes('Sort by color')")
    await waitFor("!document.querySelector('.dt-filter-menu')", 'the menu to close')
    const order = []
    for (const address of ['G2', 'G3', 'G4', 'G5']) order.push((await cell(address))?.value)
    assert.deepEqual(order, ['two', 'four', 'one', 'three'])
    assert.equal((await cell('G1'))?.value, 'Item', 'the header stays')
  })

  // ---- CALC-009: chart range picker ------------------------------------------------------------
  await check('the chart editor picks its data range on the sheet', async () => {
    await goTo('J1:K3')
    await searchMenus('chart')
    await waitFor("document.querySelector('[data-menu-action=\"search-insert-chart\"]')", 'Insert › Chart')
    await clickSelector('[data-menu-action="search-insert-chart"]')
    await waitFor("document.querySelector('aside.chart-editor')", 'the chart editor')
    await clickSelector('[aria-label="Select range on the sheet"]')
    await waitFor("document.querySelector('.range-pick-bar')", 'the range pick bar')
    await snap('range-pick')
    await clickCell('J1')
    await clickCell('K4', ['shift'])
    assert.equal(await evaluate("document.querySelector('[data-range-pick]').textContent"), 'J1:K4')
    await press('x')
    assert.equal(await editorOpen(), false, 'typing does not edit cells while picking')
    await press('Enter')
    await waitFor("!document.querySelector('.range-pick-bar')", 'the pick to finish')
    await waitFor(`(${sheetOf('Summary')}.charts || []).some((chart) => chart.dataRange === 'Summary!J1:K4')`, 'the chart range set from the sheet')
    await clickSelector('aside.chart-editor .chart-primary-button')
  })

  // ---- Print setup from the menus ---------------------------------------------------------------
  await check('File > Page setup sets the print area and page breaks', async () => {
    await goTo('A1:C10')
    await clickSelector('[data-menu-trigger="file"]')
    await clickSelector('[data-menu-action="file-page-setup"]')
    await clickSelector('[data-menu-action="print-area-set"]')
    await waitFor(`${sheetOf('Summary')}.pageSetup?.printArea === 'A1:C10'`, 'the print area')
    await goTo('A5')
    await clickSelector('[data-menu-trigger="file"]')
    await clickSelector('[data-menu-action="file-page-setup"]')
    await clickSelector('[data-menu-action="page-break-insert"]')
    await waitFor(`(${sheetOf('Summary')}.rowBreaks || []).some((entry) => entry.id === 4)`, 'a manual break above row 5')
    await waitFor("document.querySelector('.page-break-line.is-manual[data-page-break-row=\"5\"]')", 'the break drawn on the sheet')
  })

  // ---- Saved and reopened -------------------------------------------------------------------------
  await check('the zoom and the new link are saved with the file', async () => {
    await clickTab('Summary')
    await press('s', ['ctrl'])
    await waitFor("/Saved/.test(document.querySelector('.toast')?.textContent || '') || document.querySelector('[role=\"alertdialog\"]')", 'the save', 20_000)
    if (await evaluate("Boolean(document.querySelector('[role=\"alertdialog\"]'))")) {
      await evaluate("[...document.querySelectorAll('[role=\"alertdialog\"] button')].find((button) => /continue|save|yes/i.test(button.textContent))?.click()")
      await waitFor("/Saved/.test(document.querySelector('.toast')?.textContent || '')", 'the save', 20_000)
    }
    await evaluate(`window.__calcQA.openPath(${JSON.stringify(fixturePath)})`)
    await waitFor(`(${sheetOf('Summary')})?.cells?.B6?.hyperlink === "#'Data detail'!E7"`, 'the reopened link', 20_000)
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '115%'", 'the reopened zoom')
    await clickTab('Data detail')
    await waitFor("document.querySelector('.zoom-control > span')?.textContent === '100%'", 'the other sheet’s zoom')
  })

  assert.deepEqual(externalOpens, [], 'no link left the app')
  assert.deepEqual(runtimeErrors, [], 'no uncaught exceptions in the renderer')
  console.log(`Polish UI QA passed: ${passed} checks (in-workbook links, HYPERLINK(), names, Ctrl+K places, formula problems and corrections, error cards and quiet tooltips, AutoComplete, Alt+Down, Ctrl+wheel zoom per sheet and saved, Edit/Help menus, Ctrl+/ reference, Alt+/ search with arrows, manual calculation and F9, circular references, sort by color, chart range picker, page setup menus).`)
} finally {
  try { socket.close() } catch {}
  stopChildren()
}
