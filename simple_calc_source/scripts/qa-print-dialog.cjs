'use strict'

// Print and Export dialog QA in a hidden Electron window (run by npm run test:print through
// qa-print-layout.cjs). The real dialogs render the real print engine's previews; the test
// drives the page range, custom scale and margins, printer, copies and collation, Save as
// PDF, the collapsed page setup (print area, titles, header field, page break) and the Export
// dialog, and checks what each callback receives. Nothing is printed or saved.
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..')
const OUT = path.join(ROOT, 'tmp', 'qa-print-dialog')

async function buildHarness() {
  const esbuild = require('esbuild')
  fs.mkdirSync(OUT, { recursive: true })
  await esbuild.build({
    entryPoints: [path.join(__dirname, 'qa-print-dialog-entry.tsx')],
    bundle: true,
    platform: 'browser',
    format: 'iife',
    jsx: 'automatic',
    outfile: path.join(OUT, 'bundle.js'),
    loader: { '.png': 'dataurl' },
    define: { 'process.env.NODE_ENV': '"production"' },
    logLevel: 'silent',
  })
  fs.writeFileSync(path.join(OUT, 'index.html'), '<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Print dialog QA</title><link rel="stylesheet" href="./bundle.css"><style>html,body,#root{width:100%;height:100%;margin:0}</style></head><body><div id="root"></div><script src="./bundle.js"></script></body></html>')
}

function killTree(child) {
  if (!child || child.exitCode !== null) return
  if (process.platform === 'win32') require('node:child_process').spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore' })
  else child.kill('SIGKILL')
}

function runElectron(timeoutMs) {
  const profile = fs.mkdtempSync(path.join(os.tmpdir(), 'calc-print-dialog-'))
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
  let code = await runElectron(120_000)
  // A loaded machine can starve the hidden window; one retry with more time.
  if (code === 'timeout') code = await runElectron(240_000)
  if (code !== 0) throw new Error(`Print dialog QA failed (${code}).`)
}

// ---------------------------------------------------------------------------
// Driver (runs inside Electron)
// ---------------------------------------------------------------------------
async function drive() {
  const assert = require('node:assert/strict')
  const { app, BrowserWindow } = require('electron')
  app.commandLine.appendSwitch('force-device-scale-factor', '1')
  app.setPath('userData', process.env.SIMPLE_CALC_QA_PROFILE || path.join(os.tmpdir(), 'calc-print-dialog-profile'))
  let window
  const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
  const js = (code) => window.webContents.executeJavaScript(code)
  const waitFor = async (code, label, timeout = 30_000) => {
    const end = Date.now() + timeout
    while (Date.now() < end) {
      if (await js(`Boolean(${code})`).catch(() => false)) return
      await pause(60)
    }
    throw new Error(`timed out waiting for ${label}`)
  }
  const idle = () => waitFor("document.querySelector('.print-preview-pane')?.getAttribute('aria-busy') === 'false' && document.querySelector('.print-preview-frame')", 'the preview')
  const setValue = (selector, value) => js(`(() => {
    const element = document.querySelector(${JSON.stringify(selector)})
    if (!element) throw new Error('missing ' + ${JSON.stringify(selector)})
    const proto = element instanceof HTMLSelectElement ? HTMLSelectElement.prototype : element instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(element, ${JSON.stringify(value)})
    element.dispatchEvent(new Event(element instanceof HTMLSelectElement ? 'change' : 'input', { bubbles: true }))
    return true
  })()`)
  const click = (selector) => js(`(() => { const element = document.querySelector(${JSON.stringify(selector)}); if (!element) throw new Error('missing ' + ${JSON.stringify(selector)}); element.click(); return true })()`)
  const clickText = (text) => js(`(() => { const element = [...document.querySelectorAll('button')].find((item) => item.textContent.trim() === ${JSON.stringify(text)}); if (!element) throw new Error('missing button ' + ${JSON.stringify(text)}); element.click(); return true })()`)
  const open = async (mode) => {
    await window.loadFile(path.join(OUT, 'index.html'), { query: { mode } })
    await waitFor("document.querySelector('.print-layout-dialog, .export-card')", `the ${mode} dialog`)
  }

  async function printMode() {
    await open('print')
    await idle()
    assert.deepEqual(await js("[...document.querySelector('[aria-label=\"Printer\"]').options].map((option) => option.value + '=' + option.textContent)"), ['pdf=Save as PDF', '=Default printer', 'printer:Office Laser=Office Laser', 'printer:Label=Label Printer'])
    assert.equal(await js("document.querySelector('[aria-label=\"Printer\"]').value"), '', 'the default printer is preselected')
    assert.equal(await js("document.querySelector('.print-controls-actions .primary-action').textContent.trim()"), 'Print now')
    assert.match(await js("document.querySelector('.print-preview-frame').srcdoc"), /<circle[^>]*fill="#D6392B"/, 'the preview shows conditional icons')
    const firstModel = await js('window.harness.pages[window.harness.pages.length - 1]')
    assert.deepEqual(firstModel.map((page) => [page.pageNumber, page.sheetId, page.firstRow, page.lastRow, page.firstColumn, page.lastColumn]).slice(0, 2), [[1, 'report', 0, 46, 0, 9], [2, 'report', 47, 93, 0, 9]], 'the page model reports where pages break')

    await setValue('[aria-label="First page"]', '2')
    await setValue('[aria-label="Last page"]', '3')
    await waitFor("document.querySelector('.print-preview-summary span')?.textContent.startsWith('Pages 2–3 of')", 'the page range summary')
    await idle()
    assert.equal(await js("(document.querySelector('.print-preview-frame').srcdoc.match(/class=\"print-page print-sheet\"/g) || []).length"), 2, 'the preview shows only the chosen pages')

    await click('[aria-label="Use saved page layout"]')
    await waitFor("document.querySelector('[aria-label=\"Print scaling\"]')", 'manual layout controls')
    await setValue('[aria-label="Print scaling"]', 'custom')
    await waitFor("document.querySelector('[aria-label=\"Custom scale percent\"]')", 'the scale field')
    await setValue('[aria-label="Custom scale percent"]', '50')
    await waitFor("document.querySelector('.print-preview-frame')?.srcdoc.includes('data-scale=\"0.5000\"')", 'the custom scale in the preview')
    await setValue('[aria-label="Print margins"]', 'custom')
    await waitFor("document.querySelector('[aria-label=\"Top margin in inches\"]')", 'the margin fields')
    await setValue('[aria-label="Top margin in inches"]', '1.5')
    await waitFor("document.querySelector('.print-preview-frame')?.srcdoc.includes('--margin-top:144.00px')", 'the custom margin in the preview')
    await setValue('[aria-label="Left margin in inches"]', '11')
    await waitFor("document.querySelector('.print-error')?.textContent.includes('Enter each margin in inches')", 'the margin error')
    assert.equal(await js("document.querySelector('.print-controls-actions .primary-action').disabled"), true, 'an invalid margin blocks printing')
    await setValue('[aria-label="Left margin in inches"]', '0.4')
    await waitFor("!document.querySelector('.print-error')", 'the margin error to clear')

    await setValue('[aria-label="Copies"]', '0')
    await waitFor("document.querySelector('.print-controls-actions .primary-action').disabled", 'invalid copies to block printing')
    await setValue('[aria-label="Copies"]', '2')
    await waitFor("[...document.querySelectorAll('.print-check span')].some((item) => item.textContent === 'Collate')", 'the collate option')
    await click('.print-check input')
    await setValue('[aria-label="Printer"]', 'printer:Label')
    await waitFor("document.querySelector('.print-native-note').textContent.includes('Label Printer')", 'the printer note')

    assert.equal(await js("document.querySelector('.print-advanced').open"), false, 'page setup starts collapsed')
    await click('.print-advanced > summary')
    await waitFor("document.querySelector('.print-advanced').open", 'page setup to open')
    await clickText('Set to B20:D30')
    await waitFor("window.harness.patches.some((patch) => patch.printArea === 'B20:D30')", 'the print area change')
    await setValue('[aria-label="First page"]', '')
    await setValue('[aria-label="Last page"]', '')
    await waitFor("document.querySelector('.print-preview-frame')?.srcdoc.includes('data-row-range=\"20:30\"')", 'the print area in the preview')
    await setValue('[aria-label="Rows to repeat at top"]', '1')
    await js("document.querySelector('[aria-label=\"Rows to repeat at top\"]').dispatchEvent(new FocusEvent('focusout', { bubbles: true }))")
    await waitFor("window.harness.patches.some((patch) => patch.printTitlesRow === '1:1')", 'the title rows change')
    await clickText('Repeat frozen columns')
    await waitFor("window.harness.patches.some((patch) => patch.printTitlesColumn === 'A:A')", 'the title columns change')
    await js("document.querySelector('[aria-label=\"Header center\"]').focus()")
    await setValue('[aria-label="Header center"]', 'Quarterly report, page ')
    await clickText('Page')
    await waitFor("window.harness.patches.some((patch) => patch.oddHeader === '&CQuarterly report, page &P')", 'the header with a page field')
    await waitFor("document.querySelector('.print-preview-frame')?.srcdoc.includes('Quarterly report, page 1')", 'the header in the preview')
    await clickText('Insert break above row 20')
    await waitFor("window.harness.patches.some((patch) => Array.isArray(patch.rowBreaks) && patch.rowBreaks.join() === '20')", 'the page break change')
    await waitFor("document.querySelector('.print-advanced').textContent.includes('Above row 20')", 'the listed page break')

    await setValue('[aria-label="First page"]', '1')
    await setValue('[aria-label="Last page"]', '1')
    await idle()
    await waitFor("!document.querySelector('.print-controls-actions .primary-action').disabled", 'Print to be enabled')
    await click('.print-controls-actions .primary-action')
    await waitFor('window.harness.printed.length === 1', 'the print request')
    const printed = await js('window.harness.printed[0]')
    assert.deepEqual(printed.printer, { deviceName: 'Label', copies: 2, collate: false }, 'the printer, copies and collation reach the print request')
    assert.equal(printed.scaling, 'custom')
    assert.equal(printed.scalePercent, 50)
    assert.deepEqual(printed.customMargins, { top: 1.5, right: 0.5, bottom: 0.5, left: 0.4 })
    assert.deepEqual(printed.pageRange, { from: 1, to: 1 })
    await waitFor("document.getElementById('closed')", 'the dialog to close after printing')
  }

  async function pdfMode() {
    await open('print')
    await idle()
    await setValue('[aria-label="Printer"]', 'pdf')
    await waitFor("document.querySelector('.print-controls-actions .primary-action').textContent.includes('Save PDF')", 'the Save PDF action')
    assert.equal(await js("Boolean(document.querySelector('[aria-label=\"Copies\"]'))"), false, 'a PDF has no copies')
    await click('.print-controls-actions .primary-action')
    await waitFor('window.harness.pdf.length === 1', 'the PDF request')
    assert.equal(await js("Boolean(document.querySelector('.print-layout-dialog'))"), true, 'a cancelled save keeps the dialog open')
    assert.equal(await js('window.harness.pdf[0].printer'), undefined, 'a PDF carries no printer job')
  }

  async function legacyMode() {
    await open('noprinters')
    await idle()
    assert.equal(await js("Boolean(document.querySelector('[aria-label=\"Printer\"]'))"), false, 'without the printer IPC the dialog prints to the default printer, as before')
    assert.equal(await js("Boolean(document.querySelector('.print-advanced'))"), false, 'page setup editing appears only with the app callbacks')
    assert.equal(await js("document.querySelector('.print-native-note').textContent"), 'Print sends these pages directly to your default Windows printer.')
    await js("document.querySelector('.print-close').focus()")
    assert.equal(await js("(() => { document.activeElement.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true, cancelable: true })); return document.activeElement?.getAttribute('aria-label') })()"), 'Preview zoom', 'focus wraps inside the dialog')
    await click('.print-controls-actions .primary-action')
    await waitFor('window.harness.printed.length === 1', 'the print request')
    assert.equal(await js('window.harness.printed[0].printer'), undefined, 'no job settings without a printer list')
  }

  async function exportMode() {
    await open('export')
    await click('[aria-label="Use saved export layout"]')
    await waitFor("document.querySelector('[aria-label=\"Export scaling\"]')", 'export layout controls')
    await setValue('[aria-label="Export scaling"]', 'fit-pages')
    await waitFor("document.querySelector('[aria-label=\"Pages wide\"]')", 'the fit fields')
    await setValue('[aria-label="Pages tall"]', '2')
    await setValue('[aria-label="Export margins"]', 'custom')
    await waitFor("document.querySelector('[aria-label=\"Top margin in inches\"]')", 'export margin fields')
    await click('.print-actions .primary-action')
    await waitFor('window.harness.pdf.length === 1', 'the export request')
    const request = await js('window.harness.pdf[0]')
    assert.equal(request.format, 'pdf')
    assert.equal(request.options.scaling, 'fit-pages')
    assert.equal(request.options.fitWidth, 1)
    assert.equal(request.options.fitHeight, 2)
    assert.deepEqual(request.options.customMargins, { top: 0.5, right: 0.5, bottom: 0.5, left: 0.5 })
  }

  try {
    await app.whenReady()
    window = new BrowserWindow({ show: false, width: 1280, height: 860, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false, backgroundThrottling: false } })
    await printMode()
    await pdfMode()
    await legacyMode()
    await exportMode()
    console.log('Print dialog QA passed: page range preview, custom scale and margins, input validation, printer, copies and collation in the print request, Save as PDF, collapsed page setup (print area, title rows/columns, header field, page break), the legacy default-printer path, focus wrap, and Export As custom layout.')
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
