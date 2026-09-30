'use strict'

// Runs the real launcher, preload, IPC handlers, bundled worker and office
// conversion engine; only native file choices and opening the result are stubbed.
const {app,dialog,BrowserWindow}=require('electron')
const assert=require('node:assert/strict')
const fs=require('node:fs/promises')
const path=require('node:path')
const os=require('node:os')
const crypto=require('node:crypto')
const {PDFDocument}=require('../../simple_pdf_source/node_modules/pdf-lib')
const {createCanvas}=require('../../simple_pdf_source/node_modules/@napi-rs/canvas')
const {pathToFileURL}=require('node:url')
const {prepareLegacySheetPreview}=require('../launcher/legacy-sheet-preview.cjs')
const {convertOfficeBytes}=require('../../simple_doc_source/electron/office-converter.cjs')
const root=path.resolve(__dirname,'..')
const runtimeRoot=process.env.SIMPLE_COMBINE_APP_ROOT ? path.resolve(process.env.SIMPLE_COMBINE_APP_ROOT) : root
const directory=require('node:fs').mkdtempSync(path.join(os.tmpdir(),'simple-combine-ui-'))
app.setPath('userData',path.join(directory,'profile'))
const selections=[]
const pickerOptions=[]
const launched=[]
let destination=path.join(directory,'combined.pdf')
dialog.showOpenDialog=async(...args)=>{pickerOptions.push(args.at(-1));return {canceled:false,filePaths:selections.shift()||[]}}
dialog.showSaveDialog=async()=>({canceled:false,filePath:destination})
require(path.join(runtimeRoot,'electron','launch.cjs')).launchDetached=(paths)=>{launched.push(paths)}
require(path.join(runtimeRoot,'launcher','main.cjs'))

const pause=(ms)=>new Promise(resolve=>setTimeout(resolve,ms))
const hash=(bytes)=>crypto.createHash('sha256').update(bytes).digest('hex')

app.whenReady().then(async()=>{
  let window
  let exitCode=0
  try {
    while(!(window=BrowserWindow.getAllWindows()[0])) await pause(10)
    const evaluate=(code)=>window.webContents.executeJavaScript(code)
    const until=async(code,label,timeout=75_000)=>{
      const deadline=Date.now()+timeout
      while(!(await evaluate(code))) {if(Date.now()>deadline)throw new Error(`Timed out: ${label}; ${await evaluate(`document.querySelector('#combine-message')?.textContent`)}`);await pause(20)}
    }
    await until(`Boolean(document.querySelector('#combine'))`,'launcher loaded')
    const first=await PDFDocument.create()
    for(let index=1;index<=3;index++) first.addPage([612,792]).drawText(`Fixture page ${index}`,{x:50,y:700})
    const second=await PDFDocument.create();second.addPage([200,300])
    const a=path.join(directory,'three-pages.pdf')
    const b=path.join(directory,'remove-me.pdf')
    const doc=path.join(directory,'reference.doc')
    const sheet=path.join(directory,'schedule_template.xls')
    const image=path.join(directory,'phone-photo.jpg')
    const originalDoc=path.join(root,'..','.codex-tmp','reliability-reference','test_doc.doc')
    const originalSheet=path.join(root,'..','.codex-tmp','reliability-reference','schedule_template.xls')
    await fs.writeFile(a,await first.save());await fs.writeFile(b,await second.save());await fs.copyFile(originalDoc,doc);await fs.copyFile(originalSheet,sheet)
    const canvas=createCanvas(80,40),context=canvas.getContext('2d')
    for(const [color,x,y]of[['red',0,0],['lime',40,0],['blue',0,20],['yellow',40,20]]){context.fillStyle=color;context.fillRect(x,y,40,20)}
    const jpeg=canvas.toBuffer('image/jpeg',100)
    const exif=Buffer.from('45786966000049492a0008000000010012010300010000000600000000000000','hex')
    await fs.writeFile(image,Buffer.concat([jpeg.subarray(0,2),Buffer.from([0xff,0xe1,0,34]),exif,jpeg.subarray(2)]))
    const protectedSources=[a,b,doc,sheet,image,originalDoc,originalSheet]
    const originalHashes=await Promise.all(protectedSources.map(async(file)=>hash(await fs.readFile(file))))

    await evaluate(`document.querySelector('#combine').click()`)
    assert.equal(await evaluate(`document.querySelector('#combine-dialog').open && document.activeElement.id==='combine-add'`),true)
    selections.push([a,b])
    await evaluate(`document.querySelector('#combine-add').click()`)
    await until(`document.querySelectorAll('.combine-row').length===2`,'two selected files')
    for(const extension of ['xls','xlsx','ods'])assert.ok(pickerOptions.at(-1).filters[0].extensions.includes(extension),`picker must accept ${extension}`)
    await evaluate(`document.querySelector('[aria-label="Move up remove-me.pdf"]').click()`)
    assert.match(await evaluate(`document.querySelector('.combine-name strong').textContent`),/remove-me/)
    assert.equal(await evaluate(`document.activeElement.getAttribute('aria-label')`),'Preview remove-me.pdf')
    await evaluate(`document.querySelector('[aria-label="Remove remove-me.pdf"]').click()`)
    assert.equal(await evaluate(`document.querySelector('#combine-save').disabled`),true)
    selections.push([doc,sheet,image])
    await evaluate(`document.querySelector('#combine-add').click()`)
    await until(`document.querySelectorAll('.combine-row').length===4`,'Word, spreadsheet and photo added')
    await evaluate(`document.querySelector('[aria-label="Move up reference.doc"]').click()`)
    assert.equal(await evaluate(`document.querySelector('[aria-label="Pages from phone-photo.jpg"]').disabled`),true)
    assert.equal(await evaluate(`document.querySelector('[aria-label="Pages from schedule_template.xls"]').disabled`),false)
    await evaluate(`document.querySelector('[aria-label="Preview schedule_template.xls"]').click()`)
    const previewDeadline=Date.now()+10_000
    while(!launched.length){if(Date.now()>previewDeadline)throw new Error('Spreadsheet Preview did not route to its source');await pause(10)}
    assert.deepEqual(launched,[[sheet]])
    launched.length=0
    await evaluate(`(()=>{const input=document.querySelector('[aria-label="Pages from three-pages.pdf"]');input.value='99';input.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#combine-save').click()})()`)
    await until(`!document.querySelector('#combine-save').disabled && document.querySelector('#combine-message').dataset.kind==='error'`,'invalid range rejected')
    assert.match(await evaluate(`document.querySelector('#combine-message').textContent`),/three-pages.pdf: Choose pages/)
    assert.equal(await fs.stat(destination).then(()=>true,()=>false),false)

    await evaluate(`(()=>{for(const [name,value]of[['three-pages.pdf','3,1'],['schedule_template.xls','2']]){const input=document.querySelector('[aria-label="Pages from '+name+'"]');input.value=value;input.dispatchEvent(new Event('input',{bubbles:true}))}document.querySelector('#combine-save').click()})()`)
    await until(`!document.querySelector('#combine-save').disabled && document.querySelector('#combine-message').textContent.includes('schedule_template.xls: Choose pages between 1 and 1')`,'spreadsheet page range checked against its native one-page print layout')
    assert.equal(await fs.stat(destination).then(()=>true,()=>false),false)
    await evaluate(`(()=>{const input=document.querySelector('[aria-label="Pages from schedule_template.xls"]');input.value='1';input.dispatchEvent(new Event('input',{bubbles:true}))})()`)

    destination=a
    await evaluate(`(()=>{const input=document.querySelector('[aria-label="Pages from three-pages.pdf"]');input.value='3,1';input.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('#combine-save').click()})()`)
    await until(`!document.querySelector('#combine-save').disabled && document.querySelector('#combine-message').textContent.includes('source file intact')`,'source overwrite refused')
    destination=path.join(directory,'combined.pdf')
    const started=performance.now()
    await evaluate(`document.querySelector('#combine-save').click()`)
    assert.equal(await evaluate(`document.querySelector('#combine-add').disabled && !document.querySelector('#combine-close').disabled && document.querySelector('#combine-dialog').dispatchEvent(new Event('cancel',{cancelable:true}))`),true)
    await evaluate(`document.querySelector('#combine-close').click()`)
    assert.equal(await evaluate(`document.querySelector('#combine-dialog').open`),false)
    await until(`document.querySelector('#combine-message').textContent.startsWith('Saved combined.pdf')`,'combined output saved')
    const milliseconds=performance.now()-started
    assert.match(await evaluate(`document.querySelector('#status').textContent`),/Saved combined.pdf/)
    await evaluate(`document.querySelector('#combine').click()`)
    assert.deepEqual(launched,[[destination]])
    const output=await fs.readFile(destination)
    const combined=await PDFDocument.load(output)
    assert.equal(combined.getPageCount(),5)
    assert.deepEqual([combined.getPage(4).getWidth(),combined.getPage(4).getHeight()],[40,80])
    assert.deepEqual(await Promise.all(protectedSources.map(async(file)=>hash(await fs.readFile(file)))),originalHashes)

    const pdfjs=await import(pathToFileURL(path.join(root,'..','simple_pdf_source','node_modules','pdfjs-dist','legacy','build','pdf.mjs')).href)
    const resultPdf=await pdfjs.getDocument({data:new Uint8Array(output),disableWorker:true,isEvalSupported:false}).promise
    const referenceBytes=await fs.readFile(path.join(root,'..','.codex-tmp','reliability-reference','test_doc-google-reference.pdf'))
    const reference=await pdfjs.getDocument({data:new Uint8Array(referenceBytes),disableWorker:true,isEvalSupported:false}).promise
    const sheetReferenceBytes=await convertOfficeBytes({bytes:prepareLegacySheetPreview(await fs.readFile(originalSheet)),inputExtension:'xls',outputExtension:'pdf',filter:'calc_pdf_Export'})
    const sheetReference=await pdfjs.getDocument({data:new Uint8Array(sheetReferenceBytes),disableWorker:true,isEvalSupported:false}).promise
    let spreadsheetHeaderColorTokenVisible=false
    try {
      assert.equal(reference.numPages,1)
      const text=async(pdf,page)=>((await(await pdf.getPage(page)).getTextContent()).items.map(item=>item.str||'').join(' ').replace(/\s+/g,' ').trim())
      assert.equal(await text(resultPdf,1),await text(reference,1),'combined DOC must preserve the layout-engine reference text')
      assert.match(await text(resultPdf,2),/Fixture page 3/)
      assert.match(await text(resultPdf,3),/Fixture page 1/)
      assert.equal(sheetReference.numPages,1,'native spreadsheet reference must have one printed page')
      const spreadsheetText=await text(resultPdf,4)
      assert.equal(spreadsheetText,await text(sheetReference,1),'combined XLS must preserve all native printed spreadsheet text')
      spreadsheetHeaderColorTokenVisible=spreadsheetText.includes('000000')
      assert.equal(spreadsheetHeaderColorTokenVisible,false,'black formatting tokens must not become visible header text')
      assert.deepEqual((await resultPdf.getPage(4)).view,(await sheetReference.getPage(1)).view,'combined XLS must preserve native paper dimensions')
      const sheetPage=await resultPdf.getPage(4),sheetViewport=sheetPage.getViewport({scale:1.5}),sheetCanvas=createCanvas(sheetViewport.width,sheetViewport.height)
      await sheetPage.render({canvasContext:sheetCanvas.getContext('2d'),viewport:sheetViewport}).promise
      await fs.mkdir(path.join(root,'tmp'),{recursive:true})
      await fs.writeFile(path.join(root,'tmp','combine-spreadsheet.png'),sheetCanvas.toBuffer('image/png'))
      const page=await resultPdf.getPage(5),viewport=page.getViewport({scale:1})
      const rendered=createCanvas(viewport.width,viewport.height),context=rendered.getContext('2d')
      await page.render({canvasContext:context,viewport}).promise
      const pixel=context.getImageData(10,20,1,1).data
      assert.ok(pixel[2]>150&&pixel[0]<100,'phone photo must rotate clockwise: blue in top-left')
    } finally {await resultPdf.destroy();await reference.destroy();await sheetReference.destroy()}
    assert.deepEqual(await Promise.all(protectedSources.map(async(file)=>hash(await fs.readFile(file)))),originalHashes,'reference conversion also leaves every original unchanged')
    await fs.mkdir(path.join(root,'tmp'),{recursive:true})
    await fs.writeFile(path.join(root,'tmp','combine-ui.png'),(await window.webContents.capturePage()).toPNG())
    const form=await PDFDocument.create();const formPage=form.addPage([300,300]);const field=form.getForm().createTextField('Name');field.setText('Keep this value');field.addToPage(formPage,{x:10,y:40,width:200,height:30})
    const formPath=path.join(directory,'fillable.pdf');await fs.writeFile(formPath,await form.save())
    selections.push([formPath])
    await evaluate(`document.querySelector('#combine-add').click()`)
    await until(`document.querySelectorAll('.combine-row').length===5`,'form source added')
    await evaluate(`(()=>{for(let index=0;index<4;index++)document.querySelector('[aria-label="Move up fillable.pdf"]').click();document.querySelector('#combine-save').click()})()`)
    await until(`document.querySelector('#combine-message').dataset.kind==='error' && !document.querySelector('#combine-save').disabled`,'form source refused')
    assert.match(await evaluate(`document.querySelector('#combine-message').textContent`),/fillable.pdf:.*interactive form fields/)
    assert.deepEqual(await fs.readFile(destination),output,'a failed job must keep the previous saved PDF intact')
    const result={passed:true,runtime:runtimeRoot.endsWith('.asar')?'packaged-asar':'source',files:4,pages:5,wordReferencePages:1,wordTextMatches:true,spreadsheetNativePages:1,spreadsheetTextAndPaperMatch:true,spreadsheetRangeValidated:true,spreadsheetPreviewRoutesToSource:true,spreadsheetHeaderColorTokenVisible,pageOrder:['Word original layout','PDF page 3','PDF page 1','spreadsheet printed page 1','oriented photo'],invalidRangeRejected:true,sourceOverwriteRefused:true,inputHashesUnchanged:true,workerSaveMilliseconds:milliseconds,dialogReorderRemoveAndEditingLock:true,dialogDismissedWhileWorking:true,completionToast:true,interactiveFormsRejected:true,outputBytes:output.length}
    await fs.mkdir(path.join(root,'tmp'),{recursive:true})
    await fs.writeFile(path.join(root,'tmp','combine-ui-result.json'),JSON.stringify(result,null,2))
    console.log(JSON.stringify(result))
  }catch(error){console.error(error);exitCode=1}
  finally {
    if(window&&!window.isDestroyed())window.destroy()
    const absolute=path.resolve(directory)
    assert.ok(absolute.startsWith(path.resolve(os.tmpdir())+path.sep)&&path.basename(absolute).startsWith('simple-combine-ui-'))
    await fs.rm(absolute,{recursive:true,force:true,maxRetries:5,retryDelay:120}).catch(()=>{})
  }
  app.exit(exitCode)
})
