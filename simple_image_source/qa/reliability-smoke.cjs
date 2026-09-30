'use strict'
const { app, BrowserWindow, ipcMain } = require('electron')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { preparePrintableImage } = require('../electron/print-image.cjs')
const { imageDimensions } = require('../electron/image-files.cjs')
const { buildPrintHtml } = require('../electron/print-layout.cjs')
const { pathToFileURL } = require('node:url')
const { PDFDocument } = require('pdf-lib')

const root = path.resolve(__dirname, '..')
const fixtures = new Map()
let saved
let printed
const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
ipcMain.handle('file:open-path', async (_event, name) => {
  if (name === 'slow.jpg') await pause(450)
  return fixtures.get(name)
})
ipcMain.handle('file:save', (_event, input) => {
  saved = input
  return { name: input.name, path: input.path, size: input.data.byteLength, format: input.format }
})
ipcMain.handle('image:print', (_event, input) => { printed = preparePrintableImage(input); return true })
ipcMain.handle('app:get-version', () => 'test')
ipcMain.on('window:set-title', () => {})

async function until(win, expression, timeout = 15_000) {
  const start = performance.now()
  while (!(await win.webContents.executeJavaScript(expression))) {
    if (performance.now() - start > timeout) throw new Error(`Timed out: ${expression}`)
    await pause(15)
  }
  return performance.now() - start
}

app.whenReady().then(async () => {
  let win
  let printWin
  let directory
  let exitCode = 0
  try {
    win = new BrowserWindow({ width: 1365, height: 840, show: false, webPreferences: {
      preload: path.join(root, 'electron/preload.cjs'), contextIsolation: true, nodeIntegration: false, sandbox: true,
    } })
    await win.loadFile(path.join(root, 'dist/index.html'))
    const encoded = await win.webContents.executeJavaScript(`(async () => {
      const canvas = document.createElement('canvas'); canvas.width = 6000; canvas.height = 4000;
      const context = canvas.getContext('2d');
      const gradient = context.createLinearGradient(0,0,6000,4000); gradient.addColorStop(0,'#14649e'); gradient.addColorStop(1,'#ffd6a0');
      context.fillStyle=gradient; context.fillRect(0,0,6000,4000);
      for(let index=0;index<1800;index++) { context.fillStyle='hsl('+(index*17%360)+' 50% 50%)'; context.fillRect(index*137%5990,index*97%3990,40,20); }
      return canvas.toDataURL('image/jpeg',0.92).split(',')[1];
    })()`)
    const bytes = Buffer.from(encoded, 'base64')
    const exif = Buffer.from('45786966000049492a0008000000010012010300010000000600000000000000','hex')
    const portraitBytes = Buffer.concat([bytes.subarray(0,2),Buffer.from([0xff,0xe1,0,34]),exif,bytes.subarray(2)])
    for (const name of ['large.jpg', 'slow.jpg', 'latest.jpg']) fixtures.set(name, {
      data: new Uint8Array(bytes), name, path: name, size: bytes.length, format: 'jpg', mime: 'image/jpeg', directSave: true,
    })
    fixtures.set('portrait.jpg',{data:new Uint8Array(portraitBytes),name:'portrait.jpg',path:'portrait.jpg',size:portraitBytes.length,format:'jpg',mime:'image/jpeg',directSave:true})
    await win.webContents.executeJavaScript(`(() => {
      window.__pixelReadbacks = 0;
      const read = CanvasRenderingContext2D.prototype.getImageData;
      CanvasRenderingContext2D.prototype.getImageData = function(...args) { window.__pixelReadbacks++; return read.apply(this,args); };
      window.__encodes=0;
      const encode=HTMLCanvasElement.prototype.toBlob;
      HTMLCanvasElement.prototype.toBlob=function(...args) { window.__encodes++; return encode.apply(this,args); };
    })()`)
    const openStart = performance.now()
    win.webContents.send('file:open-external', 'large.jpg')
    await until(win, `document.querySelector('canvas')?.width === 6000 && document.body.innerText.includes('large.jpg')`)
    const openMs = performance.now() - openStart
    assert.equal(await win.webContents.executeJavaScript('window.__pixelReadbacks'), 0, 'JPEG opening must not read back every pixel')

    await win.webContents.executeJavaScript(`document.querySelector('button[title="Save (Ctrl+S)"]').click()`)
    await until(win, `document.querySelector('button[title="Save (Ctrl+S)"]').disabled === false`)
    assert.deepEqual(Buffer.from(saved.data), bytes, 'Unchanged save must preserve the exact source JPEG')
    assert.equal(await win.webContents.executeJavaScript('window.__encodes'), 0)

    const printStart = performance.now()
    await win.webContents.executeJavaScript(`document.querySelector('button[title="Print (Ctrl+P)"]').click()`)
    await until(win, `document.querySelector('.print-preview-image-frame img')?.naturalWidth === 6000`)
    const previewMs = performance.now() - printStart
    await win.webContents.executeJavaScript(`document.querySelector('.print-submit').click()`)
    await until(win, `!document.querySelector('.image-print-dialog')`)
    assert.deepEqual(printed.bytes, bytes, 'Untouched photo printing must preserve compressed source bytes')
    assert.equal(printed.extension, '.jpg')
    assert.equal(await win.webContents.executeJavaScript('window.__encodes'), 0)

    const legacyCost = await win.webContents.executeJavaScript(`(async () => {
      const canvas=document.querySelector('canvas'); const context=canvas.getContext('2d');
      const start=performance.now(); let alpha=false;
      for(let top=0;top<canvas.height;top+=256) {
        const data=context.getImageData(0,top,canvas.width,Math.min(256,canvas.height-top)).data;
        for(let index=3;index<data.length;index+=4) if(data[index]<255) alpha=true;
      }
      const scanMs=performance.now()-start; const encodeStart=performance.now();
      const blob=await new Promise(resolve=>canvas.toBlob(resolve,'image/png'));
      await blob.arrayBuffer();
      return {scanMs, pngEncodeMs:performance.now()-encodeStart,pngBytes:blob.size,alpha};
    })()`)

    win.webContents.send('file:open-external', 'slow.jpg')
    await pause(20)
    win.webContents.send('file:open-external', 'latest.jpg')
    await until(win, `document.body.innerText.includes('latest.jpg')`)
    await pause(600)
    assert.equal(await win.webContents.executeJavaScript(`document.body.innerText.includes('latest.jpg') && !document.body.innerText.includes('slow.jpg')`), true, 'Slow previous opens must not replace a newer image')

    // Verify the revised JPEG print payload reaches Chromium at correct page size;
    // printToPDF writes to memory and never submits a physical print job.
    directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-reliability-'))
    const jpegPath = path.join(directory, 'photo.jpg')
    const htmlPath = path.join(directory, 'print.html')
    await fs.writeFile(jpegPath, portraitBytes)
    await fs.writeFile(htmlPath, buildPrintHtml(pathToFileURL(jpegPath).href, 4000, 6000, {paper:'a4',orientation:'landscape'}))
    printWin = new BrowserWindow({ show:false, webPreferences:{contextIsolation:true,nodeIntegration:false,sandbox:true} })
    await printWin.loadFile(htmlPath)
    assert.equal(await printWin.webContents.executeJavaScript(`document.querySelector('img').decode().then(()=>document.querySelector('img').naturalWidth)`), 4000)
    const printedPdf = await PDFDocument.load(await printWin.webContents.printToPDF({printBackground:true,preferCSSPageSize:true,landscape:true,pageSize:'A4'}))
    assert.equal(printedPdf.getPageCount(), 1)
    assert.ok(Math.abs(printedPdf.getPage(0).getWidth() - 841.89) < 1)

    // Edits intentionally use the current canvas, and undo restores exact bytes.
    await win.webContents.executeJavaScript(`document.querySelector('button[title="Rotate right"]').click()`)
    await until(win, `document.querySelector('canvas').width === 4000`)
    await win.webContents.executeJavaScript(`document.querySelector('button[title="Print (Ctrl+P)"]').click()`)
    await until(win, `document.querySelector('.image-print-dialog') !== null`)
    await win.webContents.executeJavaScript(`document.querySelector('.print-submit').click()`)
    await until(win, `!document.querySelector('.image-print-dialog')`)
    assert.equal(printed.extension, '.png')
    assert.deepEqual(imageDimensions(printed.bytes,'.png'), {width:4000,height:6000})
    await win.webContents.executeJavaScript(`document.querySelector('button[title="Undo (Ctrl+Z)"]').click()`)
    await until(win, `document.querySelector('canvas').width === 6000`)
    await win.webContents.executeJavaScript(`document.querySelector('button[title="Save (Ctrl+S)"]').click()`)
    await until(win, `document.querySelector('button[title="Save (Ctrl+S)"]').disabled === false`)
    assert.deepEqual(Buffer.from(saved.data),bytes)

    win.webContents.send('file:open-external','portrait.jpg')
    await until(win,`document.body.innerText.includes('portrait.jpg') && document.querySelector('canvas')?.width===4000`)
    await win.webContents.executeJavaScript(`document.querySelector('button[title="Print (Ctrl+P)"]').click()`)
    await until(win,`document.querySelector('.print-preview-image-frame img')?.naturalWidth===4000`)
    await win.webContents.executeJavaScript(`document.querySelector('.print-submit').click()`)
    await until(win,`!document.querySelector('.image-print-dialog')`)
    assert.deepEqual(printed.bytes,portraitBytes)
    assert.deepEqual(printed.dimensions,{width:4000,height:6000})

    const result={passed:true,fixture:'6000 x 4000 JPEG (24 MP)',openMs,previewMs,sourceBytes:bytes.length,...legacyCost,unchangedSaveAndPrintExact:true,staleOpenIgnored:true,printPdfPages:1,editedPrintAndUndo:true,exifPortraitPrint:true}
    await fs.writeFile(path.join(root,'qa/reliability-result.json'),JSON.stringify(result,null,2))
    console.log(JSON.stringify(result))
  } catch(error) { console.error(error); exitCode = 1 }
  finally {
    if(win&&!win.isDestroyed()) win.destroy()
    if(printWin&&!printWin.isDestroyed()) printWin.destroy()
    if(directory) {
      const absolute=path.resolve(directory)
      assert.ok(absolute.startsWith(path.resolve(os.tmpdir())+path.sep)&&path.basename(absolute).startsWith('simple-image-reliability-'))
      await fs.rm(absolute,{recursive:true,force:true})
    }
  }
  app.exit(exitCode)
})
