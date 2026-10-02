/*
 * Keyboard, selection, clipboard and Find flows, driven through the real app in Electron over
 * the Chrome DevTools Protocol.
 *
 * The script starts its own Vite dev server and an off-screen Electron window with an isolated
 * profile and an in-memory clipboard (the system clipboard is never touched), runs the checks,
 * and stops both process trees. No network access is used beyond the local dev server.
 *
 *   npm run test:keyboard-ui
 *   QA_WORK_DIR=<dir> QA_VITE_PORT=5291 QA_CDP_PORT=9391 node scripts/qa-keyboard-ui.mjs
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
const vitePort = Number(process.env.QA_VITE_PORT || 5291)
const cdpPort = Number(process.env.QA_CDP_PORT || 9391)
const workDir = path.resolve(process.env.QA_WORK_DIR || path.join(root, 'tmp', 'keyboard-ui'))
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
const fixturePath = path.join(workDir, 'keyboard-fixture.xlsx')
{
  const book = new ExcelJS.Workbook()
  const data = book.addWorksheet('Data', { views: [{ state: 'normal', activeCell: 'C7' }] })
  for (let row = 2; row <= 20; row += 1) {
    data.getCell(`B${row}`).dataValidation = { type: 'list', allowBlank: true, formulae: ['"Yes,No"'], showErrorMessage: true, errorStyle: 'stop', error: 'Pick Yes or No.' }
  }
  data.getCell('C10').value = 'Total'
  data.getCell('C20').value = 'Total'
  data.getCell('C30').value = 'Total'
  const second = book.addWorksheet('Second')
  second.getCell('B2').value = 'Total'
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
  const focused = await evaluate("(() => { const element = document.activeElement; return element ? `${element.tagName.toLowerCase()}${element.className ? '.' + String(element.className).replace(/\\s+/g, '.') : ''}${element.getAttribute('aria-label') ? `[${element.getAttribute('aria-label')}]` : ''}` : 'none' })()").catch(() => 'unknown')
  throw new Error(`Timed out waiting for ${label} (focus on ${focused})`)
}
const settle = () => evaluate('new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(() => resolve(true))))')

await call('Runtime.enable')
await call('Page.enable')
await call('Emulation.setFocusEmulationEnabled', { enabled: true })
await waitFor('window.__calcQA && document.readyState === "complete"', 'the app and its QA hook')

// ---- Input helpers -------------------------------------------------------------------------
const KEYS = {
  Enter: { code: 'Enter', vk: 13 }, Tab: { code: 'Tab', vk: 9 }, Escape: { code: 'Escape', vk: 27 }, Backspace: { code: 'Backspace', vk: 8 },
  ArrowUp: { code: 'ArrowUp', vk: 38 }, ArrowDown: { code: 'ArrowDown', vk: 40 }, ArrowLeft: { code: 'ArrowLeft', vk: 37 }, ArrowRight: { code: 'ArrowRight', vk: 39 },
  PageUp: { code: 'PageUp', vk: 33 }, PageDown: { code: 'PageDown', vk: 34 }, F5: { code: 'F5', vk: 116 }, F10: { code: 'F10', vk: 121 }, ContextMenu: { code: 'ContextMenu', vk: 93 },
}
const MODIFIERS = { alt: 1, ctrl: 2, meta: 4, shift: 8 }
function modifierMask(modifiers = []) { return modifiers.reduce((mask, name) => mask | MODIFIERS[name], 0) }
function keyInfo(key) {
  if (KEYS[key]) return { key, ...KEYS[key] }
  if (/^[a-z]$/i.test(key)) return { key, code: `Key${key.toUpperCase()}`, vk: key.toUpperCase().charCodeAt(0) }
  if (/^[0-9]$/.test(key)) return { key, code: `Digit${key}`, vk: key.charCodeAt(0) }
  const punctuation = { ' ': ['Space', 32], '.': ['Period', 190], ',': ['Comma', 188], '[': ['BracketLeft', 219], '=': ['Equal', 187], '+': ['Equal', 187], '*': ['Digit8', 56], '(': ['Digit9', 57], ')': ['Digit0', 48], ':': ['Semicolon', 186], '-': ['Minus', 189], '!': ['Digit1', 49] }
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
async function type(text) { for (const character of text) await press(character, character === '*' || character === '(' || character === ')' || character === ':' || character === '!' || character === '+' ? ['shift'] : []) }
const activeAddress = () => evaluate("(document.querySelector('.sheet-viewport')?.getAttribute('aria-activedescendant') || '').split('-').pop()")
const nameBox = () => evaluate("document.querySelector('.name-box')?.value")
const gridHasFocus = () => evaluate("document.activeElement === document.querySelector('.sheet-viewport')")
const editorOpen = () => evaluate("Boolean(document.querySelector('textarea.cell-editor'))")
const activeSheetId = () => evaluate('window.__calcQA.workbook().activeSheetId')
const cell = (address, sheetName) => evaluate(`(() => { const book = window.__calcQA.workbook(); const sheet = ${sheetName ? `book.sheets.find((item) => item.name === ${JSON.stringify(sheetName)})` : 'book.sheets.find((item) => item.id === book.activeSheetId)'}; return sheet ? (sheet.cells[${JSON.stringify(address)}] || null) : null })()`)
const value = (address, sheetName) => evaluate(`(() => { const book = window.__calcQA.workbook(); const sheet = ${sheetName ? `book.sheets.find((item) => item.name === ${JSON.stringify(sheetName)})` : 'book.sheets.find((item) => item.id === book.activeSheetId)'}; return sheet ? window.__calcQA.value(sheet.id, ${JSON.stringify(address)}) : undefined })()`)
const marquee = () => evaluate("Boolean(document.querySelector('[data-clipboard-marquee]'))")
async function focusGrid() {
  await evaluate("document.querySelector('.sheet-viewport')?.focus({ preventScroll: true })")
  await settle()
}
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
async function clickCell(address) {
  const box = await evaluate(`(() => { const element = [...document.querySelectorAll('[data-cell-address="${address}"]')].find((item) => item.dataset.cellPane === 'body') || document.querySelector('[data-cell-address="${address}"]'); if (!element) return null; const rect = element.getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()`)
  assert.ok(box, `cell ${address} is rendered`)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: box.x, y: box.y, button: 'left', buttons: 1, clickCount: 1 })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: box.x, y: box.y, button: 'left', buttons: 0, clickCount: 1 })
  await settle()
}
/** A real mouse click (it moves focus, unlike element.click()). */
async function clickSelector(selector, filter) {
  const box = await evaluate(`(() => { const filter = ${filter ? String(filter) : 'null'}; const element = [...document.querySelectorAll(${JSON.stringify(selector)})].find((item) => !filter || filter(item)); if (!element) return null; const rect = element.getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()`)
  assert.ok(box, `${selector} is on screen`)
  await call('Input.dispatchMouseEvent', { type: 'mousePressed', x: box.x, y: box.y, button: 'left', buttons: 1, clickCount: 1 })
  await call('Input.dispatchMouseEvent', { type: 'mouseReleased', x: box.x, y: box.y, button: 'left', buttons: 0, clickCount: 1 })
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
  // ---- Open a file: its saved active cell is the selection (CALC-024) --------------------
  await evaluate(`window.__calcQA.openPath(${JSON.stringify(fixturePath)})`)
  await waitFor("document.querySelector('.sheet-viewport') && window.__calcQA.workbook()?.sheets?.length === 2", 'the fixture workbook')
  await settle()
  await check('a file opens on its saved active cell', async () => {
    assert.equal(await activeAddress(), 'C7')
  })
  await check('the grid has the keyboard after opening', async () => {
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after open')
  })

  // ---- Escape returns the keyboard to the grid (calc-grid-interaction-7) -----------------
  await check('Esc in the cell editor cancels and leaves the grid responding to arrows', async () => {
    await goTo('E5')
    await type('abc')
    assert.equal(await editorOpen(), true)
    await press('Escape')
    assert.equal(await editorOpen(), false)
    assert.equal(await gridHasFocus(), true, 'focus is back on the grid, not <body>')
    assert.equal(await cell('E5'), null, 'the cancelled entry is not stored')
    await press('ArrowDown')
    assert.equal(await activeAddress(), 'E6')
  })
  await check('Esc in the formula bar returns the keyboard to the grid', async () => {
    await evaluate("document.querySelector('.formula-bar-input')?.focus()")
    await press('Escape')
    assert.equal(await gridHasFocus(), true)
  })
  await check('closing a prompt dialog returns the keyboard to the grid', async () => {
    await evaluate("document.querySelector('.sheet-tab.is-active')?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }))")
    await waitFor("document.querySelector('.prompt-card')", 'the rename prompt')
    await press('Escape')
    await waitFor("!document.querySelector('.prompt-card')", 'the prompt to close')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after the prompt')
  })

  // ---- Arrow-key commits are validated and make one undo step (calc-grid-interaction-6) ---
  await check('an arrow-key commit runs data validation', async () => {
    await goTo('B2')
    await type('Maybe')
    await press('ArrowDown')
    await waitFor("document.querySelector('[role=\"alertdialog\"]')", 'the validation alert')
    assert.equal(await cell('B2'), null, 'the rejected entry is not stored')
    await evaluate("[...document.querySelectorAll('[role=\"alertdialog\"] button')].find((button) => button.textContent.trim() === 'Cancel')?.click()")
    await waitFor("!document.querySelector('[role=\"alertdialog\"]')", 'the alert to close')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after the alert')
  })
  await check('an arrow-key commit of 1,000 is one undo step', async () => {
    await goTo('D2')
    await type('1,000')
    await press('ArrowDown')
    assert.equal(await activeAddress(), 'D3')
    assert.equal((await cell('D2'))?.value, 1000)
    await press('z', ['ctrl'])
    assert.equal(await cell('D2'), null, 'a single Ctrl+Z removes the entry')
  })

  // ---- Enter / Tab inside a selection, Tab runs (CALC-020) --------------------------------
  await check('Enter fills a selected block down its columns and keeps the selection', async () => {
    await goTo('G5:H6')
    for (const text of ['1', '2', '3', '4']) { await type(text); await press('Enter') }
    assert.equal((await cell('G5'))?.value, 1)
    assert.equal((await cell('G6'))?.value, 2)
    assert.equal((await cell('H5'))?.value, 3)
    assert.equal((await cell('H6'))?.value, 4)
    assert.equal(await nameBox(), 'G5:H6')
    assert.equal(await activeAddress(), 'G5', 'the active cell wrapped back to the start')
  })
  await check('Tab moves along the rows of a selected block without collapsing it', async () => {
    await press('Tab')
    assert.equal(await activeAddress(), 'H5')
    await press('Tab')
    assert.equal(await activeAddress(), 'G6')
    assert.equal(await nameBox(), 'G5:H6')
  })
  await check('Enter after a Tab run returns to the column where it started', async () => {
    await goTo('G10')
    await type('a'); await press('Tab')
    await type('b'); await press('Tab')
    await type('c'); await press('Enter')
    assert.equal((await cell('I10'))?.value, 'c')
    assert.equal(await activeAddress(), 'G11')
  })

  // ---- Excel keyboard staples (CALC-021) ---------------------------------------------------
  await check('Shift+Backspace collapses the selection to the active cell', async () => {
    await goTo('G5:H6')
    await press('Tab')
    await press('Backspace', ['shift'])
    assert.equal(await nameBox(), 'H5')
  })
  await check('Ctrl+. cycles the active cell clockwise through the corners', async () => {
    await goTo('G5:H6')
    await press('.', ['ctrl'])
    assert.equal(await activeAddress(), 'H5')
    await press('.', ['ctrl'])
    assert.equal(await activeAddress(), 'H6')
    await press('.', ['ctrl'])
    assert.equal(await activeAddress(), 'G6')
    assert.equal(await nameBox(), 'G5:H6')
  })
  await check('Ctrl+A selects the current region, then the whole sheet', async () => {
    await goTo('G5')
    await press('a', ['ctrl'])
    assert.equal(await nameBox(), 'G5:H6')
    assert.equal(await activeAddress(), 'G5', 'the active cell stays put')
    await press('a', ['ctrl'])
    assert.match(await nameBox(), /^A1:/)
  })
  await check('Ctrl+Shift+8 selects the current region', async () => {
    await goTo('H6')
    await press('*', ['ctrl', 'shift'])
    assert.equal(await nameBox(), 'G5:H6')
  })
  await check('Ctrl+[ goes to the precedents of the active cell', async () => {
    await enter('J5', '=G5+H6')
    await goTo('J5')
    await press('[', ['ctrl'])
    assert.equal(await nameBox(), 'G5')
  })
  await check('Shift+F10 opens the cell menu at the active cell; Esc returns to the grid', async () => {
    await goTo('H6')
    await press('F10', ['shift'])
    await waitFor("document.querySelector('.sheet-context-menu')", 'the context menu')
    assert.equal(await nameBox(), 'H6', 'the selection is unchanged')
    await press('Escape')
    await waitFor("!document.querySelector('.sheet-context-menu')", 'the menu to close')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after the menu')
  })
  await check('Ctrl+Backspace brings the active cell back into view', async () => {
    await goTo('H6')
    await evaluate("document.querySelector('.sheet-viewport').scrollTop = 4000")
    await settle()
    await press('Backspace', ['ctrl'])
    await waitFor("(() => { const viewport = document.querySelector('.sheet-viewport'); const element = [...document.querySelectorAll('[data-cell-address=\"H6\"]')].find((item) => item.dataset.cellPane === 'body'); if (!element) return false; const a = viewport.getBoundingClientRect(); const b = element.getBoundingClientRect(); return b.top >= a.top && b.bottom <= a.bottom })()", 'H6 back in view')
    assert.equal(await nameBox(), 'H6')
  })

  // ---- Copy marquee, Paste Values (CALC-022, calc-grid-interaction-1) ----------------------
  await check('Paste Values pastes the copied formula’s value, not the cell at the same offset from A1', async () => {
    await enter('L1', 'Name')
    await enter('M2', '10')
    await enter('N2', '5')
    await enter('O2', '=M2*N2')
    await goTo('O2')
    await press('c', ['ctrl'])
    await waitFor("document.querySelector('[data-clipboard-marquee]')", 'the copy marquee')
    await goTo('Q2')
    await press('v', ['ctrl', 'shift'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).cells.Q2`, 'the pasted value')
    assert.deepEqual(await cell('Q2'), { value: 50 })
    assert.equal(await marquee(), true, 'copy mode survives the paste')
    await press('Escape')
    assert.equal(await marquee(), false, 'Esc clears the marquee')
  })

  // ---- Cut mode ends on structural edits (calc-grid-interaction-3) -------------------------
  await check('inserting a row cancels cut mode; Ctrl+V then pastes a copy and moves nothing', async () => {
    await enter('S1', '1')
    await enter('S2', '2')
    await goTo('S1:S2')
    await press('x', ['ctrl'])
    await waitFor("document.querySelector('[data-clipboard-marquee]')", 'the cut marquee')
    await goTo('A1')
    await press(' ', ['shift'])
    await press('+', ['ctrl', 'shift'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).cells.S3?.value === 2`, 'the inserted row')
    assert.equal(await marquee(), false, 'the cut marquee is gone')
    await goTo('U1')
    await press('v', ['ctrl'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).cells.U2`, 'the paste')
    assert.equal((await cell('U1'))?.value, 1)
    assert.equal((await cell('U2'))?.value, 2)
    assert.equal((await cell('S2'))?.value, 1, 'the source cells stay where they are')
    assert.equal((await cell('S3'))?.value, 2)
  })

  // ---- Cross-sheet cut keeps formulas right (calc-grid-interaction-4) ----------------------
  await check('cutting to another sheet moves inside references and pins outside ones to the source sheet', async () => {
    await enter('W1', '10')
    await enter('W2', '=W1*2')
    await enter('W3', '=X5+1')
    await enter('X5', '100')
    await goTo('W1:W3')
    await press('x', ['ctrl'])
    await press('PageDown', ['ctrl'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).name === 'Second'`, 'the second sheet')
    await goTo('C1')
    await press('v', ['ctrl'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.name === 'Second').cells.C3`, 'the moved block')
    assert.equal((await cell('C2', 'Second'))?.formula, 'C1*2')
    assert.equal((await cell('C3', 'Second'))?.formula, 'Data!X5+1')
    assert.equal(await value('C2', 'Second'), 20)
    assert.equal(await value('C3', 'Second'), 101)
    assert.equal(await cell('W1', 'Data'), null, 'the block left the source sheet')
  })

  // ---- Each sheet keeps its selection and scroll (CALC-024) --------------------------------
  await check('switching sheets keeps each sheet’s selection and scroll position', async () => {
    const secondActive = await activeAddress()
    await press('PageUp', ['ctrl'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).name === 'Data'`, 'the data sheet')
    await goTo('D60')
    await settle()
    const scrolled = await evaluate("document.querySelector('.sheet-viewport').scrollTop")
    assert.ok(scrolled > 200, 'D60 scrolled the sheet')
    await press('PageDown', ['ctrl'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).name === 'Second'`, 'the second sheet')
    assert.equal(await activeAddress(), secondActive, 'Second kept its own selection')
    await press('PageUp', ['ctrl'])
    await waitFor(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).name === 'Data'`, 'back on the data sheet')
    await settle()
    assert.equal(await activeAddress(), 'D60')
    const restored = await evaluate("document.querySelector('.sheet-viewport').scrollTop")
    assert.ok(Math.abs(restored - scrolled) <= 2, `scroll restored (${restored} vs ${scrolled})`)
  })

  // ---- An effectively unbounded grid (calc-grid-interaction-12, CALC-015) ------------------
  await check('Go To A5000 selects A5000 and renders it', async () => {
    await goTo('A5000')
    assert.equal(await activeAddress(), 'A5000')
    await waitFor("[...document.querySelectorAll('[data-cell-address=\"A5000\"]')].length > 0", 'row 5000 rendered')
  })
  await check('Go To A2:A5000 then Ctrl+Enter fills all of it', async () => {
    await goTo('A2:A5000')
    assert.equal(await nameBox(), 'A2:A5000')
    await type('7')
    await press('Enter', ['ctrl'])
    assert.equal((await cell('A5000'))?.value, 7)
    assert.equal((await cell('A2'))?.value, 7)
  })
  await check('Ctrl+Down in an empty column goes to row 1,048,576', async () => {
    await goTo('AB1')
    await press('ArrowDown', ['ctrl'])
    assert.equal(await activeAddress(), 'AB1048576')
    await waitFor("[...document.querySelectorAll('[data-cell-address=\"AB1048576\"]')].length > 0", 'the last row rendered')
    await press('ArrowUp', ['ctrl'])
    assert.equal(await activeAddress(), 'AB1')
  })

  // ---- Scrolling while editing (calc-grid-interaction-10) ----------------------------------
  await check('the mouse wheel scrolls the sheet during formula entry', async () => {
    await goTo('A1')
    await type('=SUM(')
    assert.equal(await editorOpen(), true)
    const box = await evaluate("(() => { const rect = document.querySelector('.sheet-viewport').getBoundingClientRect(); return { x: rect.left + rect.width / 2, y: rect.top + rect.height / 2 } })()")
    await call('Input.dispatchMouseEvent', { type: 'mouseWheel', x: box.x, y: box.y, deltaX: 0, deltaY: 1500 })
    await pause(400)
    await settle()
    const top = await evaluate("document.querySelector('.sheet-viewport').scrollTop")
    assert.ok(top > 300, `the sheet scrolled (${top})`)
    await pause(200)
    await settle()
    assert.ok(Math.abs((await evaluate("document.querySelector('.sheet-viewport').scrollTop")) - top) <= 1, 'the scroll is not snapped back')
    assert.equal(await editorOpen(), true, 'the entry is still open')
  })
  await check('pointing with arrows past the edge scrolls the view to the pointed cell', async () => {
    await evaluate("document.querySelector('.sheet-viewport').scrollTop = 0")
    await settle()
    for (let index = 0; index < 60; index += 1) await press('ArrowDown')
    const draft = await evaluate("document.querySelector('textarea.cell-editor')?.value")
    assert.equal(draft, '=SUM(A61')
    await waitFor("(() => { const viewport = document.querySelector('.sheet-viewport'); const element = [...document.querySelectorAll('[data-cell-address=\"A61\"]')].find((item) => item.dataset.cellPane === 'body'); if (!element) return false; const a = viewport.getBoundingClientRect(); const b = element.getBoundingClientRect(); return b.top >= a.top && b.bottom <= a.bottom })()", 'A61 in view')
    await press('Escape')
    assert.equal(await gridHasFocus(), true)
  })

  // ---- Find stays put on edits; All sheets reaches every sheet (calc-grid-interaction-11) ---
  await check('Find does not pull the selection back when you edit elsewhere', async () => {
    await goTo('A1')
    await press('f', ['ctrl'])
    await waitFor("document.activeElement?.getAttribute('aria-label') === 'Find in this sheet'", 'the find box')
    await call('Input.insertText', { text: 'Total' })
    await waitFor("document.querySelector('.sheet-viewport').getAttribute('aria-activedescendant').endsWith('-C11')", 'the first match')
    await press('Enter')
    await press('Enter')
    assert.equal(await activeAddress(), 'C31')
    await clickCell('E5')
    await type('7')
    await press('Enter')
    assert.equal((await cell('E5'))?.value, 7)
    assert.equal(await activeAddress(), 'E6', 'the selection stays where the edit left it')
  })
  await check('Find Next with All sheets reaches the other sheet and wraps back', async () => {
    const sheetName = () => evaluate(`window.__calcQA.workbook().sheets.find((item) => item.id === window.__calcQA.workbook().activeSheetId).name`)
    await evaluate("document.querySelector('[aria-label=\"Search options\"]')?.click()")
    await waitFor("document.querySelector('.search-options-row')", 'the options row')
    await evaluate("[...document.querySelectorAll('.search-options-row label')].find((label) => label.textContent.includes('All sheets'))?.querySelector('input')?.click()")
    // New options re-anchor on the match being shown (Data!C31), they do not jump around.
    await waitFor("document.querySelector('.sheet-viewport').getAttribute('aria-activedescendant').endsWith('-C31')", 'the current match kept')
    assert.equal(await sheetName(), 'Data')
    await evaluate("document.querySelector('[aria-label=\"Next match\"]').click()")
    await settle()
    assert.equal(await sheetName(), 'Second', 'Next goes on to the other sheet')
    assert.equal(await activeAddress(), 'B2')
    await settle()
    assert.equal(await sheetName(), 'Second', 'and stays there (no bounce back)')
    await evaluate("document.querySelector('[aria-label=\"Next match\"]').click()")
    await settle()
    assert.equal(await sheetName(), 'Data', 'Next wraps to the first sheet')
    assert.equal(await activeAddress(), 'C11')
    await evaluate("document.querySelector('[aria-label=\"Close find\"]').click()")
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after closing Find')
  })
  // Review F12: number-like text ('100) is found by 100 with Look in: Formulas + Entire cell.
  await check('Find with Look in Formulas and Entire cell finds number-like text', async () => {
    await goTo('M40')
    await type('1')
    await press('a', ['ctrl'])
    await call('Input.insertText', { text: "'100" })
    await press('Enter')
    assert.equal((await cell('M40'))?.value, '100', 'the entry is text')
    await goTo('A1')
    await press('f', ['ctrl'])
    await waitFor("document.activeElement?.getAttribute('aria-label') === 'Find in this sheet' || document.activeElement?.getAttribute('aria-label') === 'Find in all sheets'", 'the find box')
    if (!(await evaluate("Boolean(document.querySelector('.search-options-row'))"))) await evaluate("document.querySelector('[aria-label=\"Search options\"]')?.click()")
    await waitFor("document.querySelector('.search-options-row')", 'the options row')
    const setOptions = (entire, lookIn) => evaluate(`(() => {
      const box = [...document.querySelectorAll('.search-options-row label')].find((label) => label.textContent.includes('Entire cell')).querySelector('input')
      if (box.checked !== ${entire}) box.click()
      const select = document.querySelector('.search-options-row select')
      Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(select, ${JSON.stringify(lookIn)})
      select.dispatchEvent(new Event('change', { bubbles: true }))
    })()`)
    await setOptions(true, 'formulas')
    await evaluate("(() => { const input = document.activeElement?.closest?.('.search-panel, .search-bar') ? document.activeElement : document.querySelector('input[aria-label^=\"Find in\"]'); input.focus(); input.select() })()")
    await call('Input.insertText', { text: '100' })
    await waitFor("document.querySelector('.sheet-viewport').getAttribute('aria-activedescendant').endsWith('-M40')", 'the text cell found')
    await setOptions(false, 'values')
    await evaluate("document.querySelector('[aria-label=\"Close find\"]').click()")
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after closing Find')
  })

  // ---- Save commits the entry in progress, with validation (CALC-SIE-5) --------------------
  await check('Ctrl+S during an entry saves the entry and the active cell', async () => {
    await goTo('K2')
    await type('pending')
    assert.equal(await editorOpen(), true)
    await press('s', ['ctrl'])
    await waitFor("/Saved/.test(document.body.textContent || '')", 'the save to finish', 20_000)
    const saved = new ExcelJS.Workbook()
    await saved.xlsx.readFile(fixturePath)
    assert.equal(saved.getWorksheet('Data').getCell('K2').value, 'pending')
    assert.equal(saved.getWorksheet('Data').views?.[0]?.activeCell, 'K2', 'the saved file remembers the active cell')
  })
  await check('Ctrl+S during an entry that fails validation shows the alert and does not save', async () => {
    const before = (await fs.stat(fixturePath)).mtimeMs
    await goTo('B3')
    await type('Maybe')
    await press('s', ['ctrl'])
    await waitFor("document.querySelector('[role=\"alertdialog\"]')", 'the validation alert')
    await pause(500)
    assert.equal((await fs.stat(fixturePath)).mtimeMs, before, 'nothing was written')
    await evaluate("[...document.querySelectorAll('[role=\"alertdialog\"] button')].find((button) => button.textContent.trim() === 'Cancel')?.click()")
    await waitFor("!document.querySelector('[role=\"alertdialog\"]')", 'the alert to close')
  })

  // ---- Shift+Arrow grows whole rows / columns (review F3) ---------------------------------
  const fullySelected = (role, label) => evaluate(`[...document.querySelectorAll('[role="${role}"][aria-label="${label}"]')].some((element) => element.classList.contains('is-fully-selected'))`)
  await check('Shift+Space then Shift+Down twice selects rows 5:7, not C5:C7', async () => {
    await goTo('C5')
    await press(' ', ['shift'])
    await press('ArrowDown', ['shift'])
    await press('ArrowDown', ['shift'])
    for (const row of [5, 6, 7]) assert.equal(await fullySelected('rowheader', `Row ${row}`), true, `row ${row} is fully selected`)
    assert.equal(await fullySelected('rowheader', 'Row 8'), false)
  })
  await check('Ctrl+Space then Shift+Right selects columns C:D', async () => {
    await goTo('C5')
    await press(' ', ['ctrl'])
    await press('ArrowRight', ['shift'])
    assert.equal(await fullySelected('columnheader', 'Column C'), true)
    assert.equal(await fullySelected('columnheader', 'Column D'), true)
  })

  await check('a command chosen from the menus hands the keyboard back to the grid', async () => {
    await goTo('G5:H6')
    await clickSelector('.menu-search-trigger')
    await waitFor("document.activeElement?.getAttribute('aria-label') === 'Search menus'", 'the command search')
    await call('Input.insertText', { text: 'Gridlines' })
    await waitFor("document.querySelector('[data-menu-action=\"search-view-show-gridlines\"]')", 'the Gridlines command')
    await clickSelector('[data-menu-action="search-view-show-gridlines"]')
    await waitFor("!document.querySelector('.command-search-panel')", 'the menu to close')
    await waitFor("document.activeElement === document.querySelector('.sheet-viewport')", 'grid focus after the command')
    await press('ArrowDown')
    assert.equal(await activeAddress(), 'G6', 'arrows work straight after the command')
  })

  await check('Ctrl+S commits a side-panel field that is still being typed (chart title)', async () => {
    await goTo('G5:H6')
    await clickSelector('[data-menu-trigger]', (element) => element.textContent.trim() === 'Insert')
    await waitFor("document.querySelector('[data-menu-action=\"insert-chart\"]')", 'the Insert menu')
    await clickSelector('[data-menu-action="insert-chart"]')
    await waitFor("document.querySelector('aside.chart-editor')", 'the chart editor')
    await evaluate("[...document.querySelectorAll('.chart-editor-tabs [role=\"tab\"]')].find((tab) => tab.textContent.trim() === 'Customize')?.click()")
    await waitFor("[...document.querySelectorAll('aside.chart-editor label')].some((label) => label.textContent.includes('Chart title'))", 'the chart title field')
    await evaluate("(() => { const field = [...document.querySelectorAll('aside.chart-editor label')].find((label) => label.textContent.includes('Chart title')).querySelector('input'); field.focus(); field.select() })()")
    await call('Input.insertText', { text: 'Quarterly' })
    await press('s', ['ctrl'])
    await waitFor(`(window.__calcQA.workbook().sheets.find((item) => item.name === 'Data').charts || []).some((chart) => chart.title === 'Quarterly')`, 'the typed chart title in the workbook', 15_000)
  })

  assert.deepEqual(runtimeErrors, [], 'no uncaught exceptions in the renderer')
  console.log(`Keyboard UI QA passed: ${passed} checks (Esc/prompt/menu focus, validated arrow commits, Enter/Tab in blocks and Tab runs, Shift+Backspace, Ctrl+., Ctrl+A, Ctrl+Shift+8, Ctrl+[, Shift+F10, Ctrl+Backspace, copy marquee and Paste Values, cut cancelled by inserts, cross-sheet cut, per-sheet view, Go To beyond the drawn area, Ctrl+Down to the last row, scrolling while editing, Find, Save mid-entry, menu commands, side-panel fields on save).`)
} finally {
  try { socket.close() } catch {}
  stopChildren()
}
