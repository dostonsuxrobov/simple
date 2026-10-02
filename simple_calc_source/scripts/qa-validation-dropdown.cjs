'use strict'

// In-cell list dropdown QA (CALC-018) in a hidden Electron window: the real ValidationDropdown
// with Excel's single choice (current value highlighted, Home/End/arrows, Enter, Clear, Esc)
// and Sheets' coloured chips with several picks (checkboxes, Done, click outside keeps the
// picks, Esc cancels). Nothing is saved; the window is never shown.
//
//   npm run test:validation-dropdown
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..')
const OUT = path.join(ROOT, 'tmp', 'qa-validation-dropdown')

async function buildHarness() {
  const esbuild = require('esbuild')
  fs.mkdirSync(OUT, { recursive: true })
  await esbuild.build({
    entryPoints: [path.join(__dirname, 'qa-validation-dropdown-entry.tsx')],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    jsx: 'automatic',
    outfile: path.join(OUT, 'bundle.js'),
    loader: { '.png': 'dataurl' },
    define: { 'process.env.NODE_ENV': '"production"' },
    logLevel: 'silent',
  })
  fs.writeFileSync(path.join(OUT, 'index.html'), '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Dropdown QA</title><link rel="stylesheet" href="./bundle.css"><style>html,body,#root{width:100%;height:100%;margin:0}</style></head><body><div id="root"></div><script src="./bundle.js"></script></body></html>')
}

function killTree(child) {
  if (!child || child.exitCode !== null) return
  if (process.platform === 'win32') require('node:child_process').spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore' })
  else child.kill('SIGKILL')
}

function runElectron(timeoutMs) {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'calc-dropdown-'))
  const env = { ...process.env, SIMPLE_CALC_QA_PROFILE: profile }
  delete env.ELECTRON_RUN_AS_NODE
  const child = require('node:child_process').spawn(require('electron'), [__filename, `--user-data-dir=${profile}`], { env, windowsHide: true, stdio: ['ignore', 'inherit', 'inherit'] })
  return new Promise((resolve) => {
    const timer = setTimeout(() => { killTree(child); resolve('timeout') }, timeoutMs)
    child.on('error', () => { clearTimeout(timer); resolve('error') })
    child.on('exit', (code) => { clearTimeout(timer); resolve(code) })
  }).finally(() => {
    killTree(child)
    try { fs.rmSync(profile, { recursive: true, force: true, maxRetries: 10, retryDelay: 200 }) } catch {}
  })
}

async function run() {
  await buildHarness()
  let code = await runElectron(90_000)
  // A loaded machine can starve the hidden window; one retry with more time.
  if (code === 'timeout') code = await runElectron(180_000)
  if (code !== 0) throw new Error(`Validation dropdown QA failed (${code}).`)
}

// ---------------------------------------------------------------------------
// Driver (runs inside Electron)
// ---------------------------------------------------------------------------
async function drive() {
  const assert = require('node:assert/strict')
  const { app, BrowserWindow } = require('electron')
  app.setPath('userData', process.env.SIMPLE_CALC_QA_PROFILE || path.join(os.tmpdir(), 'calc-dropdown-profile'))
  let window
  const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
  const js = (code) => window.webContents.executeJavaScript(code)
  const waitFor = async (code, label, timeout = 20_000) => {
    const end = Date.now() + timeout
    while (Date.now() < end) {
      if (await js(`Boolean(${code})`).catch(() => false)) return
      await pause(40)
    }
    throw new Error(`timed out waiting for ${label}`)
  }
  const key = async (name, extra = {}) => {
    await js(`document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: ${JSON.stringify(name)}, bubbles: true, cancelable: true, ...${JSON.stringify(extra)} })); true`)
    await pause(30)
  }
  const press = async (selector) => {
    await js(`(() => { const element = document.querySelector(${JSON.stringify(selector)}); if (!element) throw new Error('missing ' + ${JSON.stringify(selector)}); element.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); return true })()`)
    await pause(30)
  }
  const clickText = (text) => js(`(() => { const element = [...document.querySelectorAll('button')].find((item) => item.textContent.trim() === ${JSON.stringify(text)}); if (!element) throw new Error('missing button ' + ${JSON.stringify(text)}); element.click(); return true })()`)
  const active = () => js("document.querySelector('[role=\"option\"][aria-selected=\"true\"]')?.dataset.option ?? null")
  const open = async (mode) => {
    await window.loadFile(path.join(OUT, 'index.html'), { query: { mode } })
    await waitFor("document.querySelector('[data-validation-dropdown]') && document.activeElement?.getAttribute('aria-label') === 'Search the list'", `the ${mode} dropdown`)
  }

  try {
    await app.whenReady()
    window = new BrowserWindow({ show: false, width: 800, height: 600, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false } })

    // Excel: one choice, the current value highlighted, keys move, Enter picks.
    await open('single')
    assert.equal(await active(), 'Medium', 'the current value is highlighted')
    assert.equal(await js("document.querySelector('[data-option=\"Medium\"]').classList.contains('is-picked')"), true, 'and ticked')
    await key('End')
    assert.equal(await active(), 'High')
    await key('Home')
    assert.equal(await active(), 'Low')
    await key('ArrowUp')
    assert.equal(await active(), 'High', 'the arrows wrap')
    await key('Enter')
    await waitFor("document.querySelector('.closed')", 'the list to close')
    assert.deepEqual(await js('window.harness.picks'), ['High'])

    await open('single')
    await clickText('Clear')
    await waitFor("document.querySelector('.closed')", 'the list to close')
    assert.deepEqual(await js('window.harness.picks'), [''], 'Clear empties the cell')

    await open('single')
    await key('Escape')
    await waitFor("document.querySelector('.closed')", 'the list to close')
    assert.deepEqual([await js('window.harness.picks.length'), await js('window.harness.closed')], [0, 1], 'Esc changes nothing')

    // Sheets: coloured chips, several picks.
    await open('multiple')
    assert.equal(await js("document.querySelector('[role=\"listbox\"]').getAttribute('aria-multiselectable')"), 'true')
    assert.deepEqual(await js("[...document.querySelectorAll('[role=\"option\"]')].map((item) => item.getAttribute('aria-checked'))"), ['true', 'false', 'false'], 'the current pick is ticked')
    assert.equal(await js("getComputedStyle(document.querySelector('[data-option=\"Done\"] .validation-chip')).backgroundColor"), 'rgb(183, 225, 205)', 'options wear their colours')
    assert.equal(await js("getComputedStyle(document.querySelector('[data-option=\"Open\"] .validation-chip')).backgroundColor"), 'rgb(232, 234, 229)', 'an option without a colour is a grey chip')
    await press('[data-option="Done"]')
    await press('[data-option="Open"]')
    assert.deepEqual(await js("[...document.querySelectorAll('[role=\"option\"]')].map((item) => item.getAttribute('aria-checked'))"), ['false', 'false', 'true'], 'clicks toggle')
    await key('ArrowDown')
    await key('Enter')
    assert.equal(await js("document.querySelector('[data-option=\"In progress\"]').getAttribute('aria-checked')"), 'true', 'Enter toggles the active option')
    await clickText('Done')
    await waitFor("document.querySelector('.closed')", 'the list to close')
    assert.deepEqual(await js('window.harness.many'), [['Done', 'In progress']], 'Done keeps the picks')

    await open('multiple')
    await press('[data-option="Done"]')
    await key('Escape')
    await waitFor("document.querySelector('.closed')", 'the list to close')
    assert.deepEqual([await js('window.harness.many.length'), await js('window.harness.closed')], [0, 1], 'Esc cancels the picks')

    await open('multiple')
    await press('[data-option="Done"]')
    await js("document.querySelector('.validation-dropdown-layer').dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 })); true")
    await waitFor("document.querySelector('.closed')", 'the list to close')
    assert.deepEqual(await js('window.harness.many'), [['Open', 'Done']], 'a click outside keeps the picks, as Sheets applies them')

    await open('multiple')
    const input = "document.querySelector('[aria-label=\"Search the list\"]')"
    await js(`(() => { const element = ${input}; Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(element, 'prog'); element.dispatchEvent(new Event('input', { bubbles: true })); return true })()`)
    await waitFor("document.querySelectorAll('[role=\"option\"]').length === 1", 'the filtered list')
    assert.equal(await active(), 'In progress')

    console.log('Validation dropdown QA passed: single choice (current value, Home/End/arrows, Enter, Clear, Esc) and coloured chips with several picks (toggle by click and Enter, Done, click outside keeps, Esc cancels, search).')
  } catch (error) {
    console.error(error)
    process.exitCode = 1
  } finally {
    window?.destroy()
    app.exit(process.exitCode || 0)
  }
}

module.exports = { run }

if (process.versions.electron) {
  void drive()
} else if (require.main === module) {
  run().catch((error) => { console.error(error); process.exitCode = 1 })
}
