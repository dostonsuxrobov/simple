import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { spawn, spawnSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'

const require = createRequire(import.meta.url)
const { complexWorkbook } = require('./complex-workbook-fixture.cjs')
const { workbookPayloadFromPath } = require('../electron/workbooks.cjs')
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const directory = path.join(root, 'tmp', 'complex-stress')
const profile = path.join(directory, `profile-${process.pid}`)
const entry = path.join(directory, 'entry.cjs'), source = path.join(directory, 'created.xlsx')
const port = 10200 + process.pid % 100, debugPort = port + 100
const pause = ms => new Promise(resolve => setTimeout(resolve, ms))
const report = { checks: [], exports: {}, timings: {} }, children = []
let client
class Cdp {
  constructor(url) { this.socket = new WebSocket(url); this.id = 0; this.pending = new Map(); this.socket.addEventListener('message', event => { const m = JSON.parse(event.data), p = this.pending.get(m.id); if (p) { this.pending.delete(m.id); m.error ? p.reject(new Error(m.error.message)) : p.resolve(m.result) } }) }
  connect() { return new Promise((resolve, reject) => { this.socket.addEventListener('open', resolve, { once: true }); this.socket.addEventListener('error', reject, { once: true }) }) }
  call(method, params = {}) { return new Promise((resolve, reject) => { const id = ++this.id; this.pending.set(id, { resolve, reject }); this.socket.send(JSON.stringify({ id, method, params })) }) }
  async evaluate(expression) { const r = await this.call('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true }); if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description || r.exceptionDetails.text); return r.result.value }
  async wait(expression, label, timeout = 45000) { const start = Date.now(); while (Date.now() - start < timeout) { if (await this.evaluate(`Boolean(${expression})`).catch(() => false)) return; await pause(100) } throw new Error(`Timed out: ${label}; ${await this.evaluate('document.body.innerText.slice(-600)')}`) }
}
function start(executable, args, env) { const c = spawn(executable, args, { cwd: root, env: { ...process.env, ...env }, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] }); c.output = ''; c.stdout.on('data', v => c.output += v); c.stderr.on('data', v => c.output += v); children.push(c); return c }
async function screenshot(name) { const r = await client.call('Page.captureScreenshot', { format: 'png' }); await fs.writeFile(path.join(directory, name), Buffer.from(r.data, 'base64')) }
async function selectSheet(name) { await client.evaluate(`[...document.querySelectorAll('.sheet-tab')].find(e=>e.textContent.includes(${JSON.stringify(name)})).click()`); await pause(120) }
async function edit(address, value) {
  await client.evaluate(`document.querySelector('[id$="-${address}"]').dispatchEvent(new MouseEvent('mousedown',{bubbles:true,button:0}))`)
  await pause(40)
  await client.evaluate(`document.querySelector('[id$="-${address}"]').dispatchEvent(new MouseEvent('dblclick',{bubbles:true,button:0}))`)
  await client.wait("document.querySelector('.cell-editor')", 'cell editor')
  await client.evaluate("document.querySelector('.cell-editor').select()")
  await client.call('Input.insertText', { text: value })
  await client.evaluate("document.querySelector('.cell-editor').blur()")
  await client.wait("!document.querySelector('.cell-editor')", 'edit committed')
}
async function inspectRoundtrip(file, format) {
  const payload = await workbookPayloadFromPath(file), model = payload.workbook
  const transaction = model.sheets.find(s => s.name === 'Transactions'), summary = model.sheets.find(s => s.name === 'Summary'), notes = model.sheets.find(s => s.name === 'Inputs & notes')
  const result = { bytes: (await fs.stat(file)).size, sheets: model.sheets.length, formulas: model.sheets.reduce((n,s)=>n+Object.values(s.cells).filter(c=>c.formula).length,0), hidden: model.sheets.find(s=>s.name==='Audit hidden')?.state, normalFont:model.metadata.normalFont, heading:summary?.cells.A1?.style, noteHeight:notes?.rowHeights['9'], error:notes?.cells.B12, total:summary?.cells.B5, printArea:summary?.pageSetup?.printArea }
  if (format !== 'csv') {
    assert.equal(model.sheets.length, 4)
    assert.equal(transaction.cells.D4.value, 7)
    assert.equal(notes.cells.B11.value, '00001234567890123456')
    assert.equal(notes.cells.B6.value, 'فاتورة تجريبية باللغة العربية')
    assert.ok(result.formulas >= 609, 'all calculated cells must survive; native formats may represent a constant error as an error formula')
    for (const originalSheet of complexWorkbook().workbook.sheets) for (const [address, cell] of Object.entries(originalSheet.cells)) if (cell.formula) {
      assert.ok(model.sheets.find(s => s.name === originalSheet.name)?.cells[address]?.formula, `${originalSheet.name}!${address} formula must survive`)
    }
    assert.equal(result.hidden, 'hidden', 'hidden audit sheet must remain hidden')
    assert.equal(result.normalFont.name, 'Arial')
    assert.equal(result.normalFont.size, 10)
    assert.equal(result.heading.font.name, 'Georgia')
    assert.equal(result.heading.font.size, 21)
    assert.equal(result.noteHeight, 12)
    assert.equal(result.error.value ?? result.error.result, '#DIV/0!')
    if (result.error.formula) assert.equal(result.error.formula, '#DIV/0!', 'a converted error formula must agree with its cached error and source')
  }
  return result
}
try {
  await fs.mkdir(directory, { recursive: true })
  await fs.writeFile(entry, `const{dialog}=require('electron');const path=require('node:path');dialog.showSaveDialog=async(_w,o)=>({canceled:false,filePath:o.title.startsWith('Export')?path.join(${JSON.stringify(directory)},'export.'+o.filters[0].extensions[0]):${JSON.stringify(source)}});dialog.showOpenDialog=async()=>({canceled:false,filePaths:[${JSON.stringify(source)}]});require('../../electron/main.cjs');`)
  start(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--host', '127.0.0.1', '--port', String(port), '--strictPort'])
  for (let i=0;i<150;i++) { if (await fetch(`http://127.0.0.1:${port}`).then(r=>r.ok).catch(()=>false)) break; await pause(100) }
  start(require('electron'), [`--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, entry], { VITE_DEV_SERVER_URL:`http://127.0.0.1:${port}` })
  for(let i=0;i<150;i++) { const targets=await fetch(`http://127.0.0.1:${debugPort}/json/list`).then(r=>r.json()).catch(()=>[]);const target=targets.find(t=>t.type==='page'&&t.url.includes(`:${port}`));if(target){client=new Cdp(target.webSocketDebuggerUrl);await client.connect();break}await pause(100) }
  assert.ok(client, 'Electron UI must start')
  await client.wait('window.simpleCalc', 'creation API')
  const fixture = complexWorkbook()
  const initialStart=Date.now()
  const saved=await client.evaluate(`(async()=>{const{documentId}=await window.simpleCalc.createWorkbook();return window.simpleCalc.saveWorkbook({documentId,workbook:${JSON.stringify(fixture.workbook)},format:'xlsx',saveAs:true})})()`)
  assert.equal(saved.path, source)
  report.timings.createAndSaveMs=Date.now()-initialStart
  await client.wait("document.querySelector('.welcome-actions') || document.querySelector('.welcome')", 'welcome')
  await client.evaluate("[...document.querySelectorAll('button')].find(b=>/Open/.test(b.textContent)).click()")
  await client.wait("document.querySelector('.sheet-viewport') && document.querySelector('[id$=\"-B5\"]')", 'created workbook open')
  await client.evaluate('window.confirm=()=>true')
  await screenshot('summary-before.png')
  await selectSheet('Transactions')
  await edit('D4','7')
  // Apply a real toolbar change, then verify native storage of the direct style.
  await client.evaluate("document.querySelector('[aria-label=\"Bold (Ctrl+B)\"]').click()")
  await selectSheet('Summary')
  const extra=2*15 // original D4=5; the edit adds two units at $15.
  await client.wait(`document.querySelector('[id$="-B5"]').textContent.includes(${JSON.stringify((fixture.expected.net+extra).toLocaleString('en-US',{minimumFractionDigits:2,maximumFractionDigits:2}))})`, 'cross-sheet total recalculated')
  await screenshot('summary-after.png')
  await selectSheet('Transactions'); await edit('E4','invalid'); await selectSheet('Summary')
  await client.wait("document.querySelector('[id$=\"-B5\"]').textContent.includes('#VALUE!')", 'real errors replace stale numeric results')
  await client.evaluate("document.querySelector('.save-command').click()"); await pause(500)
  const errorWorkbook = await workbookPayloadFromPath(source)
  assert.equal(errorWorkbook.workbook.sheets[0].cells.B5.result,'#VALUE!', 'Save stores the current error instead of a previous successful total')
  assert.equal(errorWorkbook.workbook.sheets[0].cells.B5.resultType,'error')
  await client.evaluate("document.querySelector('.export-as-command').click()")
  await client.wait("document.querySelector('[data-export-format=\"html\"]')",'HTML export options')
  await client.evaluate("document.querySelector('[data-export-format=\"html\"]').click()")
  await client.evaluate("document.querySelector('.export-card .primary-action').click()")
  await client.wait("!document.querySelector('.export-card')",'error HTML export')
  assert.match(await fs.readFile(path.join(directory,'export.html'),'utf8'),/#VALUE!/, 'HTML/print data contains the visible error')
  await selectSheet('Transactions'); await edit('E4','15'); await selectSheet('Summary')
  report.checks.push('live formula errors replace stale cached totals in grid, Save and HTML; corrected inputs recalculate successfully')
  await client.evaluate("document.querySelector('.save-command').click()")
  await client.wait("document.body.textContent.includes('Saved created.xlsx')", 'native XLSX save')
  await pause(350)
  report.source = await inspectRoundtrip(source,'xlsx')
  assert.equal(report.source.total.result, fixture.expected.net+extra)
  assert.equal((await workbookPayloadFromPath(source)).workbook.sheets.find(s=>s.name==='Transactions').cells.D4.style.font.bold,true)
  const sourceBytes = await fs.readFile(source)
  for (const format of ['pdf','html','ods','csv','xls']) {
    const started=Date.now()
    await client.evaluate("document.querySelector('.export-as-command').click()")
    await client.evaluate(`document.querySelector('[data-export-format="${format}"]').click()`)
    await client.evaluate("document.querySelector('.export-card .primary-action').click()")
    await client.wait("!document.querySelector('.export-card') || document.querySelector('.export-card .print-error')", `${format} export`,90000)
    const error=await client.evaluate("document.querySelector('.export-card .print-error')?.textContent")
    if(error){report.exports[format]={error};await client.evaluate("document.querySelector('[aria-label=\"Close export options\"]').click()");continue}
    const file=path.join(directory,`export.${format}`)
    report.exports[format]={elapsedMs:Date.now()-started,bytes:(await fs.stat(file)).size}
    if(['ods','xls'].includes(format)) report.exports[format].roundtrip=await inspectRoundtrip(file,format)
    if(format==='html') {
      const html=await fs.readFile(file,'utf8');assert.ok(!html.includes('HIDDEN_AUDIT_SENTINEL'))
      report.exports.html.pages=(html.match(/class="print-page print-sheet"/g)||[]).length
      const transactionPages = html.split('<section ').slice(1).filter(section=>section.includes('data-sheet-name="Transactions"'))
      assert.ok(transactionPages.length > 1)
      for (const page of transactionPages) assert.match(page, />Invoice<\/span>/, 'saved heading row repeats on every table page')
      assert.match(html, new RegExp(`Page ${report.exports.html.pages} of ${report.exports.html.pages}`), 'the last sheet continues automatic workbook page numbering')
    }
    if(format==='csv') {const csv=await fs.readFile(file,'utf8');assert.ok(!csv.includes('INV-00001'));assert.ok(!csv.includes('HIDDEN_AUDIT_SENTINEL'));assert.ok(csv.includes('Quarterly operations'))}
    assert.ok(sourceBytes.equals(await fs.readFile(source)),`${format} export must leave original unchanged`)
  }
  await selectSheet('Inputs & notes');await client.evaluate("(()=>{const z=document.querySelector('.toolbar-zoom-select');z.value='fit-sheet';z.dispatchEvent(new Event('change',{bubbles:true}))})()");await pause(150);await screenshot('notes-fit.png')
  report.checks.push('real create/save API, UI open/input edit/formatting, 609 formulas across 4 sheets, cross-sheet recalculation, exports leave original unchanged')
  assert.ok(Object.values(report.exports).every(value => !value.error), 'every requested export must complete')
  const originalClient = client
  for (const format of ['xls','ods']) {
    const oldTargets=await fetch(`http://127.0.0.1:${debugPort}/json/list`).then(r=>r.json())
    await originalClient.evaluate(`window.simpleCalc.openInNewWindow(${JSON.stringify(path.join(directory,`export.${format}`))})`)
    let next
    for(let i=0;i<450;i++){const targets=await fetch(`http://127.0.0.1:${debugPort}/json/list`).then(r=>r.json());next=targets.find(t=>t.type==='page'&&!oldTargets.some(o=>o.id===t.id));if(next)break;await pause(100)}
    assert.ok(next, 'native reopened window must start')
    client=new Cdp(next.webSocketDebuggerUrl);await client.connect()
    await client.wait("document.querySelector('.sheet-tab') && document.querySelector('.sheet-viewport')", `${format} workbook visible`,90000)
    await client.evaluate('window.confirm=()=>true')
    await selectSheet('Inputs & notes')
    await client.wait("document.querySelector('[id$=\"-B12\"]')?.textContent.includes('#DIV/0!')", 'error remains correct after live recalculation')
    await screenshot(`reopened-${format}.png`)
    if(format==='ods') {
      await edit('B3','8%')
      await client.evaluate("document.querySelector('.save-command').click()")
      await client.wait("document.body.textContent.includes('Saved export.ods')", 'native ODS save',90000)
      assert.ok(await client.evaluate("Boolean(document.querySelector('[aria-label=\"Show original backup\"]'))"), 'ODS overwrite retains an original backup')
      const savedOds=await workbookPayloadFromPath(path.join(directory,'export.ods'))
      assert.equal(savedOds.workbook.sheets.find(s=>s.name==='Inputs & notes').cells.B3.value,.08)
      assert.equal(savedOds.workbook.sheets.find(s=>s.name==='Audit hidden').state,'hidden')
      report.checks.push('ODS edited in real UI, original format saved, original backup created, updated rate and hidden state reopened')
    }
    await client.evaluate('window.simpleCalc.close()');client.socket.close();client=originalClient
  }
  assert.ok(sourceBytes.equals(await fs.readFile(source)), 'editing exported copies must leave the synthetic original unchanged')
  await fs.writeFile(path.join(directory,'report.json'),JSON.stringify(report,null,2))
  console.log(JSON.stringify(report,null,2))
} catch (error) {
  report.failure = error.stack || String(error)
  await fs.writeFile(path.join(directory,'report.json'),JSON.stringify(report,null,2))
  console.error(report.failure)
  process.exitCode = 1
} finally {
  try {await Promise.race([client?.call('Browser.close'),pause(1000)])}catch{}
  client?.socket.close()
  for(const child of children) if(child.exitCode===null) spawnSync('taskkill.exe',['/PID',String(child.pid),'/T','/F'],{windowsHide:true,stdio:'ignore'})
  if(profile.startsWith(directory+path.sep))await fs.rm(profile,{recursive:true,force:true}).catch(()=>{})
}
