import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import path from 'node:path'
import { spawn, spawnSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'

// Real Chromium layout and pointer tests; only the native file chooser is stubbed.
// Every workbook is synthetic and stays in the ignored tmp directory.
const require = createRequire(import.meta.url)
const { complexWorkbook } = require('./complex-workbook-fixture.cjs')
const { workbookPayloadFromPath } = require('../electron/workbooks.cjs')
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const directory = path.join(root, 'tmp', 'frozen-merges')
const profile = path.join(directory, `profile-${process.pid}`)
const source = path.join(directory, 'created.xlsx'), entry = path.join(directory, 'entry.cjs')
const port = 10600 + process.pid % 100, debugPort = port + 100
const pause = ms => new Promise(resolve => setTimeout(resolve, ms))
const report = { checks: [], visibility: {}, scroll: [] }, children = []
let client

class Cdp {
  constructor(url) {
    this.socket = new WebSocket(url); this.id = 0; this.pending = new Map()
    this.socket.addEventListener('message', event => {
      const message = JSON.parse(event.data), pending = this.pending.get(message.id)
      if (pending) { this.pending.delete(message.id); message.error ? pending.reject(new Error(message.error.message)) : pending.resolve(message.result) }
    })
  }
  connect() { return new Promise((resolve, reject) => { this.socket.addEventListener('open', resolve, { once: true }); this.socket.addEventListener('error', reject, { once: true }) }) }
  call(method, params = {}) { return new Promise((resolve, reject) => { const id = ++this.id; this.pending.set(id, { resolve, reject }); this.socket.send(JSON.stringify({ id, method, params })) }) }
  async evaluate(expression) {
    const response = await this.call('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true })
    if (response.exceptionDetails) throw new Error(response.exceptionDetails.exception?.description || response.exceptionDetails.text)
    return response.result.value
  }
  async wait(expression, label, timeout = 45000) {
    const started = Date.now()
    while (Date.now() - started < timeout) { if (await this.evaluate(`Boolean(${expression})`).catch(() => false)) return; await pause(100) }
    throw new Error(`Timed out: ${label}`)
  }
}
function start(executable, args, env = {}) {
  const child = spawn(executable, args, { cwd: root, env: { ...process.env, ...env }, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
  child.output = ''; child.stdout.on('data', value => child.output += value); child.stderr.on('data', value => child.output += value)
  children.push(child); return child
}
async function screenshot(name) {
  // Request a fresh compositor frame after switching sheets.
  await client.call('Emulation.setDeviceMetricsOverride', { width: 1440, height: 900, deviceScaleFactor: 1, mobile: false })
  let timer
  try {
    const result = await Promise.race([client.call('Page.captureScreenshot', { format: 'png', fromSurface: true }), new Promise((_,reject)=>{timer=setTimeout(()=>reject(new Error(`Screenshot timed out: ${name}`)),15000)})])
    await fs.writeFile(path.join(directory, name), Buffer.from(result.data, 'base64'))
  } finally { clearTimeout(timer) }
}
async function sheet(name) {
  await client.evaluate(`[...document.querySelectorAll('.sheet-tab')].find(node=>node.textContent.trim()===${JSON.stringify(name)}).click()`)
  await client.wait(`document.querySelector('.sheet-viewport')?.getAttribute('aria-label') === ${JSON.stringify(`${name} spreadsheet grid`)}`, 'sheet selected')
  await scroll(0, 0)
}
async function scroll(left, top) {
  await client.evaluate(`(()=>{const node=document.querySelector('.sheet-viewport');node.scrollLeft=${left};node.scrollTop=${top};node.dispatchEvent(new Event('scroll',{bubbles:true}))})()`)
  await pause(160)
}
async function click(x, y, count = 1) {
  await client.call('Input.dispatchMouseEvent', { type: 'mousePressed', x, y, button: 'left', clickCount: count })
  await client.call('Input.dispatchMouseEvent', { type: 'mouseReleased', x, y, button: 'left', clickCount: count })
  await pause(100)
}
async function key(key, code = key, keyCode = 27) {
  await client.call('Input.dispatchKeyEvent', { type: 'keyDown', key, code, windowsVirtualKeyCode: keyCode })
  await client.call('Input.dispatchKeyEvent', { type: 'keyUp', key, code, windowsVirtualKeyCode: keyCode })
}
async function measure(address) {
  return client.evaluate(`(() => {
    const cells=[...document.querySelectorAll('[data-cell-address="${address}"]')];
    const text=cells[0]?.textContent||'', counts=Array.from(text,()=>0);
    const rect=r=>({left:r.left,top:r.top,width:r.width,height:r.height,right:r.right,bottom:r.bottom});
    const viewport=document.querySelector('.sheet-viewport').getBoundingClientRect();
    const parts=cells.map(cell=>{
      const content=cell.querySelector('.cell-content'), node=content?.firstChild;
      if(node?.nodeType===Node.TEXT_NODE)for(let i=0;i<node.length;i++){
        const range=document.createRange();range.setStart(node,i);range.setEnd(node,i+1);
        const r=range.getBoundingClientRect(), x=(r.left+r.right)/2,y=(r.top+r.bottom)/2;
        // The narrow draggable freeze divider deliberately sits above cells.
        // Ignore only that handle, retaining opaque pane layers in the hit test.
        const hit=document.elementsFromPoint(x,y).find(element=>!element.closest('.freeze-divider'))?.closest('[data-cell-address]');
        if(r.width&&x>viewport.left+46&&y>viewport.top+27&&x<viewport.right&&y<viewport.bottom&&hit===cell)counts[i]++;
      }
      return {pane:cell.dataset.cellPane,id:cell.id,role:cell.getAttribute('role'),rect:rect(cell.getBoundingClientRect()),text:cell.textContent};
    });
    return {text,counts,parts,viewport:rect(viewport),selection:[...document.querySelectorAll('.selection-outline')].map(node=>({pane:node.dataset.selectionPane,rect:rect(node.getBoundingClientRect())})),handles:document.querySelectorAll('.fill-handle').length};
  })()`)
}
async function verifyVisible(address, expectedText, expectedPanes) {
  const measured = await measure(address)
  assert.equal(measured.text, expectedText)
  assert.deepEqual(measured.parts.map(part => part.pane).sort(), [...expectedPanes].sort())
  assert.equal(measured.parts.filter(part => part.role === 'gridcell').length, 1, 'one accessible master per merge')
  assert.equal(new Set(measured.parts.map(part => part.id)).size, measured.parts.length, 'pane fragment IDs must be unique')
  for (let index=0; index<expectedText.length; index++) if (!/\s/.test(expectedText[index])) assert.equal(measured.counts[index], 1, `each visible character paints once: ${address} character ${index} (${expectedText[index]})`)
  report.visibility[address] = { characters: expectedText.length, panes: expectedPanes, readableOnce: true }
  return measured
}

try {
  await fs.mkdir(directory, { recursive: true })
  // An isolated, non-activating window keeps native Chromium capture dependable.
  await fs.writeFile(entry, `const{BrowserWindow,dialog}=require('electron');BrowserWindow.prototype.show=function(){this.webContents.setBackgroundThrottling(false);this.showInactive()};dialog.showSaveDialog=async()=>({canceled:false,filePath:${JSON.stringify(source)}});dialog.showOpenDialog=async()=>({canceled:false,filePaths:[${JSON.stringify(source)}]});require(${JSON.stringify(path.join(root, 'electron', 'main.cjs'))});`)
  start(process.execPath, [path.join(root, 'node_modules/vite/bin/vite.js'), '--host', '127.0.0.1', '--port', String(port), '--strictPort'])
  for (let attempt=0; attempt<150; attempt++) { if (await fetch(`http://127.0.0.1:${port}`).then(response=>response.ok).catch(()=>false)) break; await pause(100) }
  start(require('electron'), [`--remote-debugging-port=${debugPort}`, `--user-data-dir=${profile}`, entry], { VITE_DEV_SERVER_URL: `http://127.0.0.1:${port}` })
  for (let attempt=0; attempt<150; attempt++) {
    const targets=await fetch(`http://127.0.0.1:${debugPort}/json/list`).then(response=>response.json()).catch(()=>[])
    const target=targets.find(item=>item.type==='page'&&item.url.includes(`:${port}`))
    if(target){client=new Cdp(target.webSocketDebuggerUrl);await client.connect();break} await pause(100)
  }
  assert.ok(client, 'hidden Electron UI starts')
  await client.wait('window.simpleCalc', 'native creation API')
  const fixture = complexWorkbook().workbook
  const summary = fixture.sheets[0]
  const crossings = { ...structuredClone(summary), id: 'crossings', name: 'Crossings', cells: {}, rowCount: 80, colCount: 20, frozen: { rows: 3, columns: 2 }, merges: ['A2:D6', 'F2:H6', 'A12:D16'], rowHeights: {}, colWidths: {} }
  const style = { font: { name: 'Arial', size: 12 }, alignment: { wrapText: true, vertical: 'middle' }, fill: { pattern: 'solid', fgColor: { argb: 'FFE0EFE9' } } }
  crossings.cells.A2 = { value: 'Both axes cross here.\nEvery pane shows its own part.\nEnd of merged content.', style }
  crossings.cells.F2 = { value: 'Only frozen rows cross here.\nThe lower lines scroll normally.\nFinal row marker.', style }
  crossings.cells.A12 = { value: 'Only frozen columns cross here.\nThe right side scrolls normally.\nFinal column marker.', style }
  crossings.cells.A12.style = { ...style, font: { name: 'Cambria', size: 13.5 } }
  crossings.merges.push('A20:D20', 'A22:D22', 'A24:D24')
  crossings.cells.A20 = { value: 'First', style }
  crossings.cells.A22 = { value: true, style }
  crossings.cells.A24 = { value: 'Example link', hyperlink: 'https://example.com', style }
  crossings.dataValidations = { A20: { type: 'list', formulae: ['"First,Second"'] }, A22: { type: 'list', formulae: ['"TRUE,FALSE"'] } }
  fixture.sheets.push(crossings)
  await client.evaluate(`(async()=>{const{documentId}=await window.simpleCalc.createWorkbook();await window.simpleCalc.saveWorkbook({documentId,workbook:${JSON.stringify(fixture)},format:'xlsx',saveAs:true})})()`)
  const beforeBytes = await fs.readFile(source)
  const beforeModel = (await workbookPayloadFromPath(source)).workbook
  await client.wait("document.querySelector('.welcome-actions') || document.querySelector('.welcome')", 'welcome')
  await client.evaluate("[...document.querySelectorAll('button')].find(node=>/Open/.test(node.textContent)).click()")
  await client.wait("document.querySelector('[data-cell-address=\"A13\"]')", 'workbook visible')
  const title = await verifyVisible('A1', summary.cells.A1.value, ['frozen-row', 'frozen-corner'])
  await verifyVisible('A13', summary.cells.A13.value, ['body', 'frozen-column'])
  assert.equal(title.selection.length, 2, 'initial single-address focus outlines the full merged title')
  assert.ok(title.selection.every(part=>Math.abs(part.rect.width-title.parts[0].rect.width)<1&&Math.abs(part.rect.height-title.parts[0].rect.height)<1))
  assert.equal(title.handles, 1)
  assert.deepEqual(await client.evaluate("[document.querySelector('[aria-label=\"Font\"]').value,document.querySelector('[aria-label=\"Font size\"]').value]"), ['Georgia','21'], 'imported 21-point font is accurately shown in toolbar')
  await screenshot('summary-fixed.png')
  // Both visible halves use the same merged master and select the full range.
  for (const part of title.parts) {
    const x = part.pane==='frozen-corner' ? part.rect.left+20 : part.rect.left+300
    await click(x, part.rect.top+20)
    const selected=await measure('A1')
    assert.equal(selected.selection.length,2);assert.equal(selected.handles,1)
    assert.ok(selected.selection.every(item=>Math.abs(item.rect.width-part.rect.width)<1))
    await click(x,part.rect.top+20,2)
    await client.wait("document.querySelector('.cell-editor')", 'merged editor from either pane')
    assert.equal(await client.evaluate("document.querySelector('.cell-editor').value"), summary.cells.A1.value)
    await key('Escape');await client.wait("!document.querySelector('.cell-editor')", 'edit canceled')
  }
  report.checks.push('full original Summary title/note visible exactly once; initial focus and clicks on both sides select the entire merge; both sides edit the same master')
  await scroll(120, 48)
  const scrolled=await measure('A1')
  for (const before of title.parts) {
    const after=scrolled.parts.find(part=>part.pane===before.pane)
    assert.ok(Math.abs(after.rect.top-before.rect.top)<1, 'frozen title remains vertically pinned')
    assert.ok(Math.abs(after.rect.left-before.rect.left-(before.pane==='frozen-corner'?0:-120))<1, 'frozen and scrolling title segments retain the correct horizontal offset')
  }
  await screenshot('summary-scrolled.png')
  await sheet('Crossings')
  const cases=[['A2', ['body','frozen-row','frozen-column','frozen-corner']], ['F2',['body','frozen-row']], ['A12',['body','frozen-column']]]
  const baseline={}
  for(const [address,panes] of cases)baseline[address]=await verifyVisible(address,crossings.cells[address].value,panes)
  const note=baseline.A12.parts.find(part=>part.pane==='frozen-column')
  await click(note.rect.left+12,note.rect.top+12)
  assert.deepEqual(await client.evaluate("[document.querySelector('[aria-label=\"Font\"]').value,document.querySelector('[aria-label=\"Font size\"]').value]"), ['Cambria','13.5'], 'custom imported font family and fractional size remain selectable without changing style')
  assert.ok(await client.evaluate("(()=>{const node=document.querySelector('[aria-label=\"Font size\"]'),style=getComputedStyle(node),context=document.createElement('canvas').getContext('2d');context.font=style.font;return context.measureText(node.value).width<=node.clientWidth-parseFloat(style.paddingLeft)-parseFloat(style.paddingRight)})()"), 'fractional font size is visibly readable within its dropdown, including the final digit')
  await screenshot('crossings-fixed.png')
  for(const [left,top] of [[64,0],[0,24],[64,24],[1200,800],[0,0]]) {
    await scroll(left,top)
    for(const [address] of cases) {
      const current=await measure(address)
      for(const before of baseline[address].parts) {
        const after=current.parts.find(part=>part.pane===before.pane)
        if(!after)continue // a fully scrolled-away merge may be virtualized out
        const pinnedX=['frozen-corner','frozen-column'].includes(before.pane),pinnedY=['frozen-corner','frozen-row'].includes(before.pane)
        assert.ok(Math.abs(after.rect.left-before.rect.left+(pinnedX?0:left))<1, `${address}/${before.pane} horizontal pane geometry`)
        assert.ok(Math.abs(after.rect.top-before.rect.top+(pinnedY?0:top))<1, `${address}/${before.pane} vertical pane geometry`)
      }
      assert.ok(current.counts.every(count=>count<=1), 'scrolling must not show a text character twice')
    }
    report.scroll.push({left,top,passed:true})
    if(left===64&&top===24)await screenshot('crossings-scrolled.png')
  }
  await scroll(64,24)
  const corner=await client.evaluate("(()=>{const r=document.querySelector('.frozen-pane-layer.is-corner').getBoundingClientRect();return {right:r.right,bottom:r.bottom}})()")
  const pointer={x:corner.right+35,y:corner.bottom+12}
  assert.equal(await client.evaluate(`document.elementFromPoint(${pointer.x},${pointer.y})?.closest('[data-cell-address]')?.dataset.cellPane`),'body')
  await click(pointer.x,pointer.y)
  let selected=await measure('A2');assert.equal(selected.selection.length,4);assert.equal(selected.handles,1)
  await click(pointer.x,pointer.y,2)
  await client.wait("document.querySelector('.cell-editor')",'editor from scrolling fragment')
  const editor=await client.evaluate("(()=>{const e=document.querySelector('.cell-editor'),r=e.getBoundingClientRect(),v=document.querySelector('.sheet-viewport').getBoundingClientRect();return {text:e.value,left:r.left,top:r.top,width:r.width,viewportLeft:v.left,viewportTop:v.top,viewportRight:v.right}})()")
  assert.equal(editor.text,crossings.cells.A2.value)
  assert.ok(editor.left>=editor.viewportLeft+45&&editor.top>=editor.viewportTop+26&&editor.left+editor.width<=editor.viewportRight,'editor remains fully visible while its scrolling copy has passed behind the divider')
  await screenshot('crossings-scrolled-editor.png')
  await key('Escape');await client.wait("!document.querySelector('.cell-editor')",'scrolled edit canceled')
  await scroll(0,0)
  for(const [address,panes] of cases)await verifyVisible(address,crossings.cells[address].value,panes)
  const controls=await client.evaluate(`(() => {
    const data={};for(const address of ['A20','A22','A24']){
      const copies=[...document.querySelectorAll('[data-cell-address="'+address+'"]')];
      data[address]=copies.map(cell=>{const input=cell.querySelector('select,input,button');return {pane:cell.dataset.cellPane,continuation:cell.getAttribute('aria-hidden')==='true',tabIndex:input.tabIndex,inert:input.inert}});
    }return data;
  })()`)
  for(const entries of Object.values(controls)){
    assert.equal(entries.filter(entry=>entry.tabIndex>=0).length,1,'only one focusable control for a merged cell')
    assert.ok(entries.filter(entry=>entry.continuation).every(entry=>entry.tabIndex===-1))
  }
  assert.ok(controls.A20.filter(entry=>entry.continuation).every(entry=>entry.inert),'dropdown copies are inert')
  await client.evaluate("window.pickerCalls=[];const original=HTMLSelectElement.prototype.showPicker;HTMLSelectElement.prototype.showPicker=function(){const call={id:this.closest('[data-cell-address]').id,hidden:!!this.closest('[aria-hidden=true]')};window.pickerCalls.push(call);try{const result=original.call(this);call.opened=true;return result}catch(error){call.error=String(error);throw error}}")
  const selectPoint=await client.evaluate("(()=>{const e=document.querySelector('[data-cell-address=\"A20\"][data-cell-pane=\"body\"]'),r=e.getBoundingClientRect(),f=document.querySelector('.frozen-pane-layer.is-column').getBoundingClientRect();return {x:f.right+25,y:r.top+r.height/2}})()")
  await click(selectPoint.x,selectPoint.y)
  assert.equal(await client.evaluate('window.pickerCalls.length'),1,'continuation click invokes the native master picker')
  assert.equal(await client.evaluate('window.pickerCalls[0].hidden'),false)
  assert.equal(await client.evaluate('window.pickerCalls[0].opened'),true,'native picker actually opens successfully')
  assert.equal(await client.evaluate("document.activeElement?.closest('[data-cell-address]')?.getAttribute('aria-hidden')"),null,'focus never remains inside an aria-hidden continuation')
  await key('Escape')
  report.checks.push('merged dropdown, checkbox and hyperlink expose only one tab stop; continuation dropdown uses the accessible native master picker')
  for(const percent of [75,125]){
    await sheet('Summary')
    await client.evaluate(`(()=>{const e=document.querySelector('.toolbar-zoom-select');e.value='${percent}';e.dispatchEvent(new Event('change',{bubbles:true}))})()`)
    await pause(160);await scroll(0,0)
    await verifyVisible('A1',summary.cells.A1.value,['frozen-row','frozen-corner'])
    await verifyVisible('A13',summary.cells.A13.value,['body','frozen-column'])
    await screenshot(`summary-${percent}.png`)
  }
  report.checks.push('75% and 125% zoom retain readable title and paragraph; scrolled four-pane merge selects and opens a visible sticky editor; Escape restores content; toolbar reports Georgia21 and Cambria13.5 accurately')
  report.checks.push('column-only, row-only, and four-pane merged cells; independent and combined scrolling; far-scroll virtualization and return; no duplicate visible text or IDs')
  await client.evaluate("document.querySelector('.save-command').click()")
  await pause(400)
  assert.ok(beforeBytes.equals(await fs.readFile(source)), 'viewing, scrolling, selection and canceling edits leave source bytes untouched')
  const afterModel=(await workbookPayloadFromPath(source)).workbook
  for(let index=0;index<beforeModel.sheets.length;index++) {
    for(const property of ['cells','merges','frozen'])assert.deepEqual(afterModel.sheets[index][property],beforeModel.sheets[index][property], `${property} must remain unchanged`)
  }
  report.checks.push('original saved cells, merge ranges, freeze metadata and workbook bytes unchanged')
  await fs.writeFile(path.join(directory,'report.json'),JSON.stringify(report,null,2))
  console.log(JSON.stringify(report,null,2))
} catch(error) {
  report.failure=error.stack||String(error)
  try { await screenshot('failure.png') } catch {}
  await fs.writeFile(path.join(directory,'report.json'),JSON.stringify(report,null,2))
  console.error(report.failure); process.exitCode=1
} finally {
  try {await Promise.race([client?.call('Browser.close'),pause(1000)])}catch{}
  client?.socket.close()
  for(const child of children)if(child.exitCode===null)spawnSync('taskkill.exe',['/PID',String(child.pid),'/T','/F'],{windowsHide:true,stdio:'ignore'})
  const resolvedProfile=path.resolve(profile)
  if(resolvedProfile.startsWith(path.resolve(directory)+path.sep))await fs.rm(resolvedProfile,{recursive:true,force:true}).catch(()=>{})
}
