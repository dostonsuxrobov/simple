'use strict'
// Real Image renderer/preload/main routes; only native file choices are replaced.
const {app,BrowserWindow,dialog}=require('electron')
const fs=require('node:fs/promises'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict'),crypto=require('node:crypto')
const root=path.resolve(__dirname,'..'),directory=require('node:fs').mkdtempSync(path.join(os.tmpdir(),'simple-image-composition-'))
app.setPath('userData',path.join(directory,'profile'))
let destination=''
dialog.showSaveDialog=async()=>({canceled:false,filePath:destination})
require('../electron/main.cjs')
app.removeAllListeners('window-all-closed')
const pause=ms=>new Promise(r=>setTimeout(r,ms)),hash=b=>crypto.createHash('sha256').update(b).digest('hex')
app.whenReady().then(async()=>{
  let win,exitCode=0
  try{
    while(!(win=BrowserWindow.getAllWindows()[0]))await pause(10)
    const evaluate=code=>win.webContents.executeJavaScript(code).catch(error=>{throw new Error(`${error.message}\n${code.slice(0,400)}`)})
    const until=async(code,label,timeout=30_000)=>{const end=Date.now()+timeout;while(!(await evaluate(code))){if(Date.now()>end)throw new Error(`${label}: ${await evaluate(`document.querySelector('.toast')?.textContent`)}`);await pause(20)}}
    await until('Boolean(window.simpleImage && document.querySelector("canvas"))','renderer')
    const fixtures=await evaluate(`(()=>{const c=document.createElement('canvas');c.width=4096;c.height=3072;const x=c.getContext('2d');for(const[color,left,top]of[['red',0,0],['lime',2048,0],['blue',0,1536],['yellow',2048,1536]]){x.fillStyle=color;x.fillRect(left,top,2048,1536)}x.clearRect(0,0,128,3072);x.clearRect(0,0,4096,128);x.clearRect(3968,0,128,3072);x.clearRect(0,2944,4096,128);x.clearRect(128,128,128,128);x.fillStyle='rgba(255,0,0,.5)';x.fillRect(128,128,128,128);return {png:c.toDataURL('image/png').split(',')[1],webp:c.toDataURL('image/webp',.92).split(',')[1],jpeg:c.toDataURL('image/jpeg',.92).split(',')[1]}})()`)
    const png=path.join(directory,'alpha.png'),webp=path.join(directory,'alpha.webp')
    await fs.writeFile(png,Buffer.from(fixtures.png,'base64'));await fs.writeFile(webp,Buffer.from(fixtures.webp,'base64'))
    const jpeg=path.join(directory,'camera.jpg'),encodedJpeg=Buffer.from(fixtures.jpeg,'base64'),exif=Buffer.from('45786966000049492a0008000000010012010300010000000600000000000000','hex')
    await fs.writeFile(jpeg,Buffer.concat([encodedJpeg.subarray(0,2),Buffer.from([0xff,0xe1,0,34]),exif,encodedJpeg.subarray(2)]))
    const originals=await Promise.all([png,webp,jpeg].map(async p=>hash(await fs.readFile(p))))
    const open=async(file,width=4096)=>{win.webContents.send('file:open-external',file);await until(`document.querySelector('canvas')?.width===${width} && document.body.innerText.includes(${JSON.stringify(path.basename(file))})`,'open image')}
    const click=async(title)=>evaluate(`[...document.querySelectorAll('button')].find(b=>b.title===${JSON.stringify(title)}).click()`)
    const exportAs=async(format,name)=>{destination=path.join(directory,name);await evaluate(`document.querySelector('.export-button').click()`);await until(`Boolean(document.querySelector('[data-export-format="${format}"]'))`,'export menu');await evaluate(`document.querySelector('[data-export-format="${format}"]').click()`);await until(`!document.querySelector('.export-button').disabled`,'export completed');return fs.readFile(destination)}
    const saveAs=async(name)=>{destination=path.join(directory,name);await evaluate(`document.querySelector('[title="Save options"]').click()`);await until(`Boolean(document.querySelector('.save-menu'))`,'save options');await evaluate(`document.querySelector('.save-menu button').click()`);await until(`!document.querySelector('button[title="Save (Ctrl+S)"]').disabled`,'save completed');return fs.readFile(destination)}
    const readPixels=async(bytes,points)=>evaluate(`(async()=>{const bitmap=await createImageBitmap(new Blob([Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}),c=>c.charCodeAt(0))]));const c=document.createElement('canvas');c.width=bitmap.width;c.height=bitmap.height;const x=c.getContext('2d');x.drawImage(bitmap,0,0);bitmap.close();const result={width:c.width,height:c.height,pixels:${JSON.stringify(points)}.map(([px,py])=>Array.from(x.getImageData(px,py,1,1).data))};c.width=1;c.height=1;return result})()`)
    const began=performance.now();await open(png);const openMilliseconds=performance.now()-began
    await evaluate(`(()=>{window.__temporaryCanvases=[];const create=document.createElement.bind(document);document.createElement=function(name,...args){const element=create(name,...args);if(name==='canvas')window.__temporaryCanvases.push(element);return element}})()`)
    assert.deepEqual(await saveAs('copy.png'),await fs.readFile(png),'unchanged PNG Save As must preserve source bytes')
    await click('Crop');await evaluate(`document.querySelector('.apply-button').click()`)
    await until(`document.querySelector('canvas').width===3277`,'default crop')
    assert.ok(await evaluate(`document.querySelector('.inspector').innerText.includes('RGB · opaque')`))
    await click('Undo (Ctrl+Z)')
    await until(`document.querySelector('canvas').width===4096`,'undo crop')
    assert.ok(await evaluate(`document.querySelector('.inspector').innerText.includes('RGBA · transparency')`),'Undo must restore transparency details together with pixels')
    const fit=await evaluate(`(()=>{const c=document.querySelector('canvas').getBoundingClientRect(),v=document.querySelector('.viewport').getBoundingClientRect();return {width:c.width,height:c.height,availableWidth:v.width-88,availableHeight:v.height-88}})()`)
    assert.ok(fit.width<=fit.availableWidth+2&&fit.height<=fit.availableHeight+2,'Undo in Fit mode must keep the restored full image inside the viewport')
    await click('Rotate right');await until(`document.querySelector('canvas').width===3072`,'rotate')
    const rotated=await exportAs('png','rotated.png')
    const colorSamples=await readPixels(rotated,[[500,500],[2500,500],[500,3500],[2500,3500],[5,5],[2871,200]])
    assert.deepEqual(colorSamples,{width:3072,height:4096,pixels:[[0,0,255,255],[255,0,0,255],[255,255,0,255],[0,255,0,255],[0,0,0,0],[255,0,0,128]]})
    await click('Undo (Ctrl+Z)');await until(`document.querySelector('canvas').width===4096`,'undo rotation')
    await click('Brush')
    const point=await evaluate(`(()=>{const r=document.querySelector('canvas').getBoundingClientRect();return {x:Math.round(r.left+r.width*.25),y:Math.round(r.top+r.height*.25)}})()`)
    win.webContents.sendInputEvent({type:'mouseMove',...point});win.webContents.sendInputEvent({type:'mouseDown',button:'left',clickCount:1,...point});win.webContents.sendInputEvent({type:'mouseUp',button:'left',clickCount:1,...point})
    await until(`document.querySelector('.save-state').textContent==='Modified'`,'paint')
    const painted=await exportAs('png','painted.png'),brushPixel=(await readPixels(painted,[[1024,768]])).pixels[0]
    assert.ok(brushPixel[0]<40&&brushPixel[1]<40&&brushPixel[2]<40,'brush mark must appear at the intended source pixel')
    await click('Undo (Ctrl+Z)');await until(`document.querySelector('.save-state').textContent==='Saved'`,'undo paint')
    await open(webp)
    assert.deepEqual(await saveAs('copy.webp'),await fs.readFile(webp),'unchanged WebP Save As must retain its native encoding')
    await click('Rotate left');await until(`document.querySelector('canvas').width===3072`,'rotate WebP')
    const editedWebp=await saveAs('edited.webp'),webpPixels=await readPixels(editedWebp,[[5,5],[500,500]])
    assert.equal(webpPixels.pixels[0][3],0,'edited WebP must retain transparent edges')
    assert.ok(webpPixels.pixels[1][1]>240&&webpPixels.pixels[1][0]<15,'rotated WebP must retain green quadrant')
    const pdf=await exportAs('pdf','composition.pdf')
    const {PDFDocument}=require('pdf-lib'),document=await PDFDocument.load(pdf)
    assert.equal(document.getPageCount(),1)
    assert.ok(Math.abs(document.getPage(0).getWidth()/document.getPage(0).getHeight()-.75)<.00001)
    const images=document.context.enumerateIndirectObjects().filter(([,o])=>o.dict?.get?.(require('pdf-lib').PDFName.of('Subtype'))?.toString()==='/Image')
    assert.ok(images.some(([,o])=>o.dict.has(require('pdf-lib').PDFName.of('SMask'))),'PDF must preserve transparent image pixels with a soft mask')
    const pdfjs=await import(require('node:url').pathToFileURL(require.resolve('../../simple_pdf_source/node_modules/pdfjs-dist/legacy/build/pdf.mjs')).href)
    const renderedPdf=await pdfjs.getDocument({data:new Uint8Array(pdf),disableWorker:true,isEvalSupported:false}).promise
    try{
      const page=await renderedPdf.getPage(1),viewport=page.getViewport({scale:1}),canvas=require('../../simple_pdf_source/node_modules/@napi-rs/canvas').createCanvas(viewport.width,viewport.height),context=canvas.getContext('2d')
      await page.render({canvasContext:context,viewport}).promise
      const corner=Array.from(context.getImageData(2,2,1,1).data),green=Array.from(context.getImageData(Math.floor(viewport.width*.25),Math.floor(viewport.height*.25),1,1).data)
      assert.deepEqual(corner,[255,255,255,255],'transparent PDF corner must render as white paper, not black')
      assert.ok(green[1]>240&&green[0]<15&&green[2]<15,'PDF rendered quadrant must match the edited WebP orientation')
      await fs.mkdir(path.join(root,'tmp'),{recursive:true});await fs.writeFile(path.join(root,'tmp','composition-pdf.png'),canvas.toBuffer('image/png'))
    }finally{await renderedPdf.destroy()}
    await open(jpeg,3072)
    assert.deepEqual(await saveAs('camera-copy.jpeg'),await fs.readFile(jpeg),'unchanged EXIF JPEG Save As must retain orientation and compressed bytes')
    await click('Rotate left');await until(`document.querySelector('canvas').width===4096`,'manual rotation after EXIF orientation')
    const cameraEdited=await saveAs('camera-edited.jpeg'),cameraPixels=await readPixels(cameraEdited,[[500,500],[3500,500],[500,2500],[3500,2500]])
    assert.deepEqual([cameraPixels.width,cameraPixels.height],[4096,3072])
    for(const [index,channel]of [[0,0],[1,1],[2,2]])assert.ok(cameraPixels.pixels[index][channel]>240,'EXIF plus manual rotation must retain source quadrant order')
    assert.ok(cameraPixels.pixels[3][0]>240&&cameraPixels.pixels[3][1]>240)
    const temporaryCanvasPixels=await evaluate(`window.__temporaryCanvases.reduce((sum,c)=>sum+c.width*c.height,0)`)
    assert.ok(temporaryCanvasPixels<100,'finished edits/encodes must release large temporary canvases even before garbage collection')
    assert.deepEqual(await Promise.all([png,webp,jpeg].map(async p=>hash(await fs.readFile(p)))),originals)
    await fs.mkdir(path.join(root,'tmp'),{recursive:true});await fs.writeFile(path.join(root,'tmp','composition-stress.png'),(await win.webContents.capturePage()).toPNG())
    const result={passed:true,pixels:4096*3072,openMilliseconds,pngAndWebpUnchangedSaveExact:true,undoRestoresAlphaAndFit:true,rotatedQuadrantsAndHalfAlphaExact:true,paintPixelCorrect:true,editedWebpAlpha:true,exifAndManualJpegRotation:true,pdfPages:1,pdfSoftMask:true,pdfRenderedColorAndAlpha:true,temporaryCanvasPixels,originalsUnchanged:true}
    await fs.writeFile(path.join(root,'tmp','composition-stress-result.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result))
  }catch(error){console.error(error);exitCode=1}
  finally{for(const window of BrowserWindow.getAllWindows())window.destroy();const absolute=path.resolve(directory);assert.ok(absolute.startsWith(path.resolve(os.tmpdir())+path.sep)&&path.basename(absolute).startsWith('simple-image-composition-'));await Promise.race([fs.rm(absolute,{recursive:true,force:true,maxRetries:2,retryDelay:100}).catch(()=>{}),pause(2000)])}
  app.exit(exitCode)
})
