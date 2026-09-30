'use strict'
// Uses disposable files and automatic save-dialog answers to exercise the
// actual renderer, preload, save writer, and all eight export formats.
const { app, BrowserWindow, dialog, nativeImage } = require('electron')
const fs = require('node:fs/promises')
const path = require('node:path')
const assert = require('node:assert/strict')
const { PDFDocument, rgb } = require('pdf-lib')
const output = path.resolve(__dirname, '../../.codex-tmp/pdf-save-export-regression')
app.setPath('userData', path.join(output, `profile-${Date.now()}`))
let saveTarget = null
dialog.showSaveDialog = async () => saveTarget ? { canceled: false, filePath: saveTarget } : { canceled: true }
const backend = process.env.SIMPLE_TEST_BACKEND || path.resolve(__dirname, '../electron/main.cjs')
require(backend)
const pause = ms => new Promise(resolve => setTimeout(resolve, ms))
const evaluate = (win, code) => win.webContents.executeJavaScript(code)
async function until(win, code, label) {
  const deadline = Date.now() + 45000
  while (!(await evaluate(win, code))) {
    if (Date.now() > deadline) throw new Error(`${label}: ${await evaluate(win, "document.querySelector('.toast')?.textContent || document.body.innerText.slice(-400)")}`)
    await pause(40)
  }
}
async function textOf(file) {
  const mupdf = await import('mupdf')
  const doc = mupdf.Document.openDocument(await fs.readFile(file), 'application/pdf')
  const page = doc.loadPage(0), text = page.toStructuredText()
  try { return text.asText() } finally { text.destroy(); page.destroy(); doc.destroy() }
}

app.whenReady().then(async () => {
  const timer = setTimeout(() => { console.error('Regression timed out'); app.exit(1) }, 180000)
  try {
    await fs.mkdir(output, { recursive: true })
    const input = path.join(output, 'input.pdf')
    const doc = await PDFDocument.create(), page = doc.addPage([400, 400])
    for (let x = 0; x < 400; x += 8) page.drawRectangle({ x, y: 0, width: 8, height: 400, color: rgb(x / 400, .6, .8) })
    page.drawText('Original editable text', { x: 40, y: 250, size: 16 })
    page.drawText('Neighbor stays unchanged', { x: 40, y: 275, size: 16 })
    await fs.writeFile(input, await doc.save())
    let home
    while (!(home = BrowserWindow.getAllWindows()[0])) await pause(20)
    await until(home, 'Boolean(window.simple)', 'preload')
    const ids = new Set(BrowserWindow.getAllWindows().map(win => win.id))
    await evaluate(home, `window.simple.openInNewWindow(${JSON.stringify(input)})`)
    let win
    while (!(win = BrowserWindow.getAllWindows().find(win => !ids.has(win.id)))) await pause(20)
    win.webContents.on('console-message', (_event, level, message) => { if (level >= 2) console.error(message) })
    await until(win, 'Boolean(document.querySelector("[data-text-item=true]"))', 'loaded')
    await evaluate(win, `document.querySelector('button[aria-label="Edit text (E)"]').click()`)
    await evaluate(win, `(() => {
      const s=[...document.querySelectorAll('[data-text-item=true]')].find(s=>s.textContent==='Original editable text')
      const b=s.getBoundingClientRect()
      s.dispatchEvent(new MouseEvent('click',{bubbles:true,clientX:b.x+b.width/2,clientY:b.y+b.height/2}))
    })()`)
    await until(win, 'Boolean(document.querySelector("textarea.inline-pdf-text-editor"))', 'editor')
    await until(win, `document.querySelector('.page-canvas')?.dataset.textRemovals?.includes('Original editable text')`, 'background preview')
    await fs.writeFile(path.join(output, 'selected.png'), (await win.webContents.capturePage()).toPNG())
    await evaluate(win, `(() => {
      const e=document.querySelector('textarea.inline-pdf-text-editor')
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype,'value').set.call(e,'First saved edit')
      e.dispatchEvent(new Event('input',{bubbles:true}))
    })()`)
    // Save while the editor is still open, then save the committed state again.
    await evaluate(win, `document.querySelector('.save-button').click()`)
    await until(win, `!document.querySelector('.busy-overlay') && !document.querySelector('.dirty-dot')`, 'first save')
    assert.match(await textOf(input), /First saved edit/)
    assert.doesNotMatch(await textOf(input), /Original editable text/)
    await evaluate(win, `document.querySelector('.save-button').click()`)
    await until(win, `!document.querySelector('.busy-overlay')`, 'repeat save')
    assert.equal((await textOf(input)).match(/First saved edit/g)?.length, 1)
    await evaluate(win, `document.querySelector('.text-overlay').click()`)
    await until(win, 'Boolean(document.querySelector("textarea.inline-pdf-text-editor"))', 're-edit')
    await evaluate(win, `(() => {
      const e=document.querySelector('textarea.inline-pdf-text-editor')
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype,'value').set.call(e,'Final exported edit')
      e.dispatchEvent(new Event('input',{bubbles:true}))
    })()`)
    // Export pending edits without requiring Done or a save first.
    const formats = { pdf:'PDF', png:'PNG images', jpeg:'JPEG images', webp:'WebP images', docx:'Word document', txt:'Plain text', md:'Markdown', html:'Web page' }
    for (const [format, label] of Object.entries(formats)) {
      saveTarget = path.join(output, `edited.${format === 'jpeg' ? 'jpg' : format}`)
      await fs.rm(saveTarget, { force: true })
      await evaluate(win, `document.querySelector('.export-as-button').click()`)
      await until(win, `Boolean(document.querySelector('[aria-label="Export As"]'))`, 'export dialog')
      await evaluate(win, `[...document.querySelectorAll('[role=radio]')].find(e=>e.querySelector('strong')?.textContent===${JSON.stringify(label)}).click()`)
      await evaluate(win, `[...document.querySelectorAll('.export-dialog-footer button')].find(e=>e.textContent==='Export').click()`)
      await until(win, `!document.querySelector('[aria-label="Export As"]') && !document.querySelector('.busy-overlay')`, `${format} export`)
      assert.ok((await fs.stat(saveTarget)).size > 10)
      if (format === 'pdf') {
        const text = await textOf(saveTarget)
        assert.match(text, /Final exported edit/)
        assert.doesNotMatch(text, /Original editable text|First saved edit/)
        assert.match(text, /Neighbor stays unchanged/)
      } else if (format === 'docx') {
        const text = (await require('mammoth').extractRawText({ path: saveTarget })).value
        assert.match(text, /Final exported edit/)
        assert.doesNotMatch(text, /Original editable text|First saved edit/)
      } else if (['txt','md','html'].includes(format)) {
        const text = await fs.readFile(saveTarget, 'utf8')
        assert.match(text, /Final exported edit/)
        assert.doesNotMatch(text, /Original editable text|First saved edit/)
      } else if (format !== 'webp') {
        assert.equal(nativeImage.createFromPath(saveTarget).isEmpty(), false)
      } else assert.equal((await fs.readFile(saveTarget)).toString('ascii', 8, 12), 'WEBP')
      console.log(`PASS ${format}`)
    }
    saveTarget = null // Cancel Save As must keep pending edits.
    await evaluate(win, `window.dispatchEvent(new KeyboardEvent('keydown',{key:'S',ctrlKey:true,shiftKey:true,bubbles:true}))`)
    await until(win, `!document.querySelector('.busy-overlay')`, 'cancel Save As')
    assert.match(await textOf(input), /First saved edit/)
    assert.equal(await evaluate(win, `Boolean(document.querySelector('.dirty-dot'))`), true)
    await evaluate(win, `document.querySelector('.save-button').click()`)
    await until(win, `!document.querySelector('.busy-overlay') && !document.querySelector('.dirty-dot')`, 'final save')
    assert.match(await textOf(input), /Final exported edit/)
    await fs.writeFile(path.join(output, 'done.png'), (await win.webContents.capturePage()).toPNG())
    await fs.writeFile(path.join(output, 'results.json'), JSON.stringify({ passed: true, backend, formats: Object.keys(formats), repeatedSave: true, pendingEdits: true, canceledSaveAs: true }, null, 2))
    console.log('PASS repeated save, pending edits, canceled Save As, text extraction, and all exports')
    app.exit(0)
  } catch (error) { console.error(error.stack); app.exit(1) }
  finally { clearTimeout(timer) }
})
