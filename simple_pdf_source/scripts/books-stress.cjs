'use strict'
// Real renderer/preload/save tests on disposable copies. Never writes Books.
const { app, BrowserWindow, dialog } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const directory = path.resolve(__dirname, '../../.codex-tmp/books-stress-20260905')
const booksDirectory = process.env.SIMPLE_BOOKS_SOURCE || path.join(app.getPath('documents'), 'books')
const run = process.env.SIMPLE_BOOKS_RUN || 'baseline'
assert.match(run, /^[a-z0-9-]+$/i, 'Run names must stay inside the test workspace')
const output = path.join(directory, run)
app.setPath('userData', path.join(directory, `profile-${run}`))
const choices = []
dialog.showOpenDialog = async () => ({ canceled: false, filePaths: choices.shift() || [] })
dialog.showSaveDialog = async () => ({ canceled: true })
require('../electron/main.cjs')
const cases = [
  { id: 'napoleon', file: 'Napoleon A Life.pdf', page: 30, original: 'founder of modern France', replacement: 'builder of modern France' },
  { id: 'garamond', file: 'Elon Musk by Walter Isacson.pdf', page: 30, original: 'very large head', replacement: 'very small head' },
  { id: 'colored-cover', file: 'Building Engines for Growth and Competitiveness in China.pdf', page: 1, original: 'Growth', replacement: 'Change' },
  { id: 'berling', file: 'Building Engines for Growth and Competitiveness in China.pdf', page: 30, original: 'average annual GDP', replacement: 'average yearly GDP' },
  { id: 'mixed-style', file: 'Insurance_Explained_Simply.pdf', page: 30, original: 'The house itself.', replacement: 'The home itself.' },
  { id: 'cyrillic', file: 'Воспоминания и размышления.pdf', page: 30, original: 'без крова', replacement: 'без жилья' },
  { id: 'tall-page', file: 'Machine intelligence, part 1.pdf', page: 1, original: 'part 1', replacement: 'part 2' },
  { id: 'calibri', file: 'Chinggis Khan Letter to Changchun 1219.pdf', page: 1, original: 'extravagant luxury', replacement: 'extravagant wealth' },
  { id: 'scan', file: 'The Innovator’s Dilemma 2000.pdf', page: 30, scan: true },
]
const pause = ms => new Promise(resolve => setTimeout(resolve, ms))
const evaluate = (win, code) => win.webContents.executeJavaScript(code)
async function until(win, code, label, timeout = 45000) {
  const deadline = Date.now() + timeout
  while (!(await evaluate(win, code))) {
    if (Date.now() > deadline) throw new Error(`Timeout ${label}: ${await evaluate(win, "document.querySelector('.toast')?.textContent || document.body.innerText.slice(-500)")}`)
    await pause(40)
  }
}
const results = []
app.whenReady().then(async () => {
  await fs.mkdir(output, { recursive: true })
  let last
  try {
    while (!(last = BrowserWindow.getAllWindows()[0])) await pause(20)
    await until(last, 'Boolean(window.simple)', 'initial preload')
    for (const test of cases.filter(c => !process.env.SIMPLE_BOOKS_CASE || c.id === process.env.SIMPLE_BOOKS_CASE)) {
      const result = { ...test }
      let win
      try {
        const target = path.join(output, `${test.id}.pdf`)
        await fs.copyFile(path.join(directory, 'originals', test.file), target)
        const ids = new Set(BrowserWindow.getAllWindows().map(w => w.id))
        const started = performance.now()
        await evaluate(last, `window.simple.openInNewWindow(${JSON.stringify(target)})`)
        while (!(win = BrowserWindow.getAllWindows().find(w => !ids.has(w.id)))) await pause(20)
        win.show(); win.focus()
        await until(win, 'Boolean(document.querySelector(".page-canvas"))', 'canvas')
        if (test.page !== 1) {
          await evaluate(win, `(()=>{const i=document.querySelector('input[aria-label="Current page"]'); Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set.call(i,${JSON.stringify(String(test.page))});i.dispatchEvent(new Event('input',{bubbles:true}));i.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))})()`)
        }
        const sourceSelector = '.continuous-page-slot.is-current [data-text-item="true"]'
        if (test.scan) {
          await evaluate(win, `document.querySelector('button[aria-label*="Edit text"]').click()`)
          await until(win, `document.querySelector('.edit-inspector')?.textContent.includes('This page has no editable text')`, 'scan explanation')
          assert.equal(await evaluate(win, `document.querySelectorAll(${JSON.stringify(sourceSelector)}).length`), 0)
          await fs.writeFile(path.join(output, 'scan-explanation.png'), (await win.webContents.capturePage()).toPNG())
          await evaluate(win, `[...document.querySelectorAll('.edit-inspector button')].find(b=>b.textContent==='Add text').click()`)
          await evaluate(win, `(()=>{const c=document.querySelector('.continuous-page-slot.is-current .page-canvas'),b=c.getBoundingClientRect();c.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,clientX:b.left+25,clientY:b.top+25}))})()`)
          await until(win, `Boolean(document.querySelector('textarea[aria-label="Edit text directly on the PDF"]'))`, 'scan note editor')
          await evaluate(win, `(()=>{const i=document.querySelector('textarea[aria-label="Edit text directly on the PDF"]');Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype,'value').set.call(i,'TEST COPY - scan note');i.dispatchEvent(new Event('input',{bubbles:true}));i.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',ctrlKey:true,bubbles:true}))})()`)
          await until(win, `Boolean(document.querySelector('.text-overlay'))`, 'scan note committed')
          await evaluate(win, `document.querySelector('.save-button').click()`)
          await until(win, `!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'scan note saved')
          result.passed = true; result.scanNoteAdded = true
          await fs.writeFile(path.join(output, 'scan-note.png'), (await win.webContents.capturePage()).toPNG())
          results.push(result); await fs.writeFile(path.join(output, 'results.json'), JSON.stringify(results, null, 2))
          win.destroy(); console.log(JSON.stringify({id:test.id,passed:true,scanNoteAdded:true})); continue
        }
        await until(win, `[...document.querySelectorAll(${JSON.stringify(sourceSelector)})].some(s=>s.textContent.includes(${JSON.stringify(test.original)}))`, 'source text')
        result.openMs = Math.round(performance.now() - started)
        await evaluate(win, `document.querySelector('button[aria-label*="Edit text"]').click()`)
        result.source = await evaluate(win, `(()=>{const s=[...document.querySelectorAll(${JSON.stringify(sourceSelector)})].find(s=>s.textContent.includes(${JSON.stringify(test.original)}));s.scrollIntoView({block:'center'});const start=s.textContent.indexOf(${JSON.stringify(test.original)}),range=document.createRange();range.setStart(s.firstChild,start);range.setEnd(s.firstChild,start+${test.original.length});window.getSelection().removeAllRanges();window.getSelection().addRange(range);const b=range.getBoundingClientRect(), page=s.closest('.pdf-page')||s.closest('.page-surface');const data={text:s.textContent,dataset:{...s.dataset},selectedRect:{x:b.x,y:b.y,width:b.width,height:b.height}};s.dispatchEvent(new MouseEvent('click',{bubbles:true,cancelable:true,clientX:b.x+b.width/2,clientY:b.y+b.height/2}));return data})()`)
        const replacementText = result.source.text.replace(test.original, test.replacement)
        await until(win, `document.querySelector('textarea[aria-label="Edit text directly on the PDF"]')?.value === ${JSON.stringify(result.source.text)}`, 'editor')
        result.editor = await evaluate(win, `(()=>{const i=document.querySelector('textarea[aria-label="Edit text directly on the PDF"]'),s=getComputedStyle(i);return {color:s.color,font:s.font,fontFamily:s.fontFamily}})()`)
        await evaluate(win, `(()=>{const i=document.querySelector('textarea[aria-label="Edit text directly on the PDF"]');Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype,'value').set.call(i,${JSON.stringify(replacementText)});i.dispatchEvent(new Event('input',{bubbles:true}));i.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',ctrlKey:true,bubbles:true}))})()`)
        await until(win, `[...document.querySelectorAll('.text-overlay')].some(s=>s.textContent === ${JSON.stringify(replacementText)})`, 'committed overlay')
        result.overlay = await evaluate(win, `(()=>{const e=[...document.querySelectorAll('.text-overlay')].find(s=>s.textContent === ${JSON.stringify(replacementText)});return {style:e.getAttribute('style'),cover:e.parentElement.querySelector('.text-original-cover')?.getAttribute('style')}})()`)
        await fs.writeFile(path.join(output, `${test.id}-editor.png`), (await win.webContents.capturePage()).toPNG())
        const saving = performance.now()
        await evaluate(win, `document.querySelector('.save-button').click()`)
        await until(win, `!document.querySelector('.dirty-dot') && !document.querySelector('.busy-overlay')`, 'save', 90000)
        result.saveMs = Math.round(performance.now() - saving)
        const oldWindow = win
        const savedIds = new Set(BrowserWindow.getAllWindows().map(w => w.id))
        await evaluate(oldWindow, `window.simple.openInNewWindow(${JSON.stringify(target)})`)
        while (!(win = BrowserWindow.getAllWindows().find(w => !savedIds.has(w.id)))) await pause(20)
        oldWindow.destroy(); win.show(); win.focus()
        await until(win, 'Boolean(document.querySelector(".page-canvas"))', 'reopened canvas')
        if (test.page !== 1) await evaluate(win, `(()=>{const i=document.querySelector('input[aria-label="Current page"]');Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set.call(i,${JSON.stringify(String(test.page))});i.dispatchEvent(new Event('input',{bubbles:true}));i.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))})()`)
        // Covered source text is still extractable and can interleave with the
        // new words in DOM reading order. Verify the replacement face's runs
        // together, rather than accidentally accepting only the old text.
        await until(win, `(()=>{if(document.querySelector('input[aria-label="Current page"]')?.value !== ${JSON.stringify(String(test.page))})return false;const fonts={};for(const s of document.querySelectorAll(${JSON.stringify(sourceSelector)}))fonts[s.dataset.pdfFontName]=(fonts[s.dataset.pdfFontName]||'')+s.textContent;return Object.values(fonts).some(t=>t.replace(/\\s/g,'').includes(${JSON.stringify(replacementText.replace(/\s/g, ''))}))})()`, 'reopened replacement text on the requested page')
        result.reopenedReplacementVerified = true
        result.savedRuns = await evaluate(win, `[...document.querySelectorAll(${JSON.stringify(sourceSelector)})].filter(s=>Math.abs(Number(s.dataset.pdfBaselineY)-${Number(result.source.dataset.pdfBaselineY)})<0.1).map(s=>({text:s.textContent,dataset:{...s.dataset}}))`)
        if (test.id === 'tall-page') {
          await evaluate(win, `(()=>{const i=document.querySelector('input[aria-label="Zoom percentage"]');Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value').set.call(i,'400');i.dispatchEvent(new Event('input',{bubbles:true}));i.dispatchEvent(new KeyboardEvent('keydown',{key:'Enter',bubbles:true}))})()`)
          await until(win, `document.querySelector('.page-canvas')?.width > 1000 && document.querySelector('.page-canvas')?.height > 3000`, '400% tall page')
          result.zoom400 = await evaluate(win, `(()=>{const c=document.querySelector('.page-canvas');return {width:c.width,height:c.height,pixels:c.width*c.height}})()`)
        }
        await fs.writeFile(path.join(output, `${test.id}-saved.png`), (await win.webContents.capturePage()).toPNG())
        result.bytes = (await fs.stat(target)).size
        result.passed = true
      } catch (error) { result.error = error.stack; result.passed = false }
      if (win && !win.isDestroyed()) win.destroy()
      results.push(result)
      await fs.writeFile(path.join(output, 'results.json'), JSON.stringify(results, null, 2))
      console.log(JSON.stringify({ id: result.id, passed: result.passed, openMs: result.openMs, saveMs: result.saveMs, error: result.error }))
    }
    const inventory = JSON.parse(await fs.readFile(path.join(directory, 'inventory.json'), 'utf8'))
    for (const file of inventory) {
      const bytes = await fs.readFile(path.join(booksDirectory, file.file))
      assert.equal(crypto.createHash('sha256').update(bytes).digest('hex'), file.sha256, `Original changed: ${file.file}`)
    }
    console.log('All 27 Books originals remain byte-identical.')
  } finally { for (const win of BrowserWindow.getAllWindows()) win.destroy() }
  app.exit(results.every(r => r.passed) ? 0 : 1)
}).catch(error => { console.error(error); app.exit(1) })
