// Real Chromium layout/PDF regression; never submits a physical print job.
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
if (!process.versions.electron) {
  const env = { ...process.env }; delete env.ELECTRON_RUN_AS_NODE
  const result = require('node:child_process').spawnSync(require('electron'), [__filename, ...process.argv.slice(2)], { env, stdio: 'inherit', windowsHide: true })
  process.exit(result.status ?? 1)
}
const { app, BrowserWindow } = require('electron')
const { createSpreadsheetPrintDocument } = require('../electron/spreadsheet-print.cjs')
const { complexWorkbook } = require('./complex-workbook-fixture.cjs')
const { workbookPayloadFromPath } = require('../electron/workbooks.cjs')
const SSF = require('xlsx').SSF
let window, profile
;(async () => {
  profile = await fs.mkdtemp(path.join(os.tmpdir(), 'calc-print-geometry-'))
  app.setPath('userData', profile)
  await app.whenReady()
  window = new BrowserWindow({ show: false, width: 1250, height: 1600, webPreferences: { sandbox: true, contextIsolation: true, nodeIntegration: false } })
  const model = process.argv[2] ? (await workbookPayloadFromPath(path.resolve(process.argv[2]))).workbook : complexWorkbook().workbook
  await window.loadURL('about:blank')
  const digitWidth = await window.webContents.executeJavaScript(`(()=>{const c=document.createElement('canvas').getContext('2d');c.font='10pt Arial';return Math.max(...'0123456789'.split('').map(n=>c.measureText(n).width))})()`)
  const displayValues = {}
  for (const sheet of model.sheets) {
    sheet.properties = { ...sheet.properties, printDigitWidth: digitWidth }
    displayValues[sheet.id] = Object.fromEntries(Object.entries(sheet.cells).map(([address, cell]) => {
      const value = cell.formula ? cell.result : cell.value
      let display = value ?? ''
      try { if (typeof value === 'number') display = SSF.format(cell.numFmt || 'General', value) } catch {}
      return [address, String(display)]
    }))
  }
  const directory = path.resolve('tmp/complex-print')
  await fs.mkdir(directory, { recursive: true })
  async function render(workbook, name, values) {
    const result = createSpreadsheetPrintDocument({ name, workbook, displayValues: values, options: { scope: 'workbook', useSavedLayout: true } })
    const input = path.join(directory, `${name}.html`)
    await fs.writeFile(input, result.html)
    await window.loadFile(input)
    await window.webContents.executeJavaScript('document.fonts.ready.then(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))))')
    const geometry = await window.webContents.executeJavaScript(`Array.from(document.querySelectorAll('.print-page')).map(p=>{const c=p.querySelector('.page-content').getBoundingClientRect(),t=p.querySelector('table').getBoundingClientRect();return{name:p.dataset.sheetName,scale:p.dataset.scale,contentRight:c.right,tableRight:t.right,contentBottom:c.bottom,tableBottom:t.bottom}})`)
    for (const page of geometry) {
      assert.ok(page.tableRight <= page.contentRight + .1, `${page.name}: outside right border is clipped: ${JSON.stringify(page)}`)
      assert.ok(page.tableBottom <= page.contentBottom + .1, `${page.name}: outside bottom border is clipped`)
    }
    await fs.writeFile(path.join(directory, `${name}.pdf`), await window.webContents.printToPDF({ preferCSSPageSize: true, printBackground: true, margins: { top: 0, left: 0, right: 0, bottom: 0 } }))
    return geometry
  }
  const geometry = await render(model, 'complex', displayValues)
  // Independently cover the maximum supported border stroke at a full-width fit.
  const probe = complexWorkbook().workbook
  probe.sheets = [probe.sheets[0]]
  for (const side of Object.values(probe.sheets[0].cells.A13.style.border)) side.style = 'thick'
  await render(probe, 'thick-border', {})
  await fs.writeFile(path.join(directory, 'geometry.json'), JSON.stringify(geometry, null, 2))
  console.log(`Print geometry passed: ${geometry.length} mixed-paper pages, full-width thin and thick merged borders inside the printable rectangle; digit width ${digitWidth}.`)
})().catch(error => { console.error(error); process.exitCode = 1 }).finally(async () => {
  window?.destroy()
  if (profile) await fs.rm(profile, { recursive: true, force: true }).catch(() => {})
  app.exit(process.exitCode || 0)
})
