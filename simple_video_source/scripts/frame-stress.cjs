'use strict'
// Real app/main/preload exports, checked against an independent FFmpeg decoder.
const {app,BrowserWindow,dialog,nativeImage}=require('electron')
const fs=require('node:fs/promises'),path=require('node:path'),os=require('node:os'),assert=require('node:assert/strict'),crypto=require('node:crypto')
const {spawnSync}=require('node:child_process')
const root=path.resolve(__dirname,'..'),directory=require('node:fs').mkdtempSync(path.join(os.tmpdir(),'simple-video-frames-'))
app.setPath('userData',path.join(directory,'profile'))
const saveChoices=[]
dialog.showSaveDialog=async(...args)=>{const name=path.basename(args.at(-1).defaultPath);saveChoices.push(name);return {canceled:false,filePath:path.join(directory,name)}}
require('../electron/main.cjs')
app.removeAllListeners('window-all-closed')
const pause=ms=>new Promise(r=>setTimeout(r,ms)),hash=b=>crypto.createHash('sha256').update(b).digest('hex')
function ffmpeg(args){const result=spawnSync('ffmpeg',['-hide_banner','-loglevel','error','-y',...args],{encoding:'utf8',windowsHide:true});if(result.error||result.status!==0)throw new Error(result.error?.message||result.stderr)}
app.whenReady().then(async()=>{
  let win,exitCode=0
  try{
    while(!(win=BrowserWindow.getAllWindows()[0]))await pause(10)
    const evaluate=code=>win.webContents.executeJavaScript(code).catch(error=>{throw new Error(`${error.message}\n${code.slice(0,250)}`)})
    const until=async(code,label,timeout=20_000)=>{const end=Date.now()+timeout;while(!(await evaluate(code))){if(Date.now()>end)throw new Error(`${label}: ${await evaluate(`document.querySelector('.toast')?.textContent`)}`);await pause(20)}}
    await until('Boolean(window.simpleVideo)','preload')
    const fixture=path.join(directory,'portrait odd.webm')
    ffmpeg(['-f','lavfi','-i','testsrc=size=321x481:rate=12:duration=65','-c:v','libvpx-vp9','-lossless','1','-pix_fmt','yuv444p','-an',fixture])
    const original=hash(await fs.readFile(fixture))
    win.webContents.send('file:open-external',fixture)
    await until(`document.querySelector('video')?.readyState>=2 && document.querySelector('video').videoWidth===321`,'odd portrait loaded')
    const dimensions=await evaluate(`(()=>{const v=document.querySelector('video');return [v.videoWidth,v.videoHeight,v.duration]})()`)
    assert.deepEqual(dimensions,[321,481,65])
    const fitGeometry=await evaluate(`(()=>{const v=document.querySelector('video'),r=v.getBoundingClientRect(),s=document.querySelector('.video-stage').getBoundingClientRect();return {mode:getComputedStyle(v).objectFit,width:r.width,height:r.height,stageWidth:s.width,stageHeight:s.height}})()`)
    assert.equal(fitGeometry.mode,'contain')
    assert.ok(fitGeometry.width<=fitGeometry.stageWidth+1&&fitGeometry.height<=fitGeometry.stageHeight+1,`Fit must show the entire portrait frame without overflowing its stage: ${JSON.stringify(fitGeometry)}`)
    const capture=async(seconds,format='PNG image')=>{
      await evaluate(`document.querySelector('[title="Export As (Ctrl+Shift+E)"]').click()`)
      await until(`Boolean(document.querySelector('.export-modal'))`,'export dialog')
      const began=performance.now(),before=saveChoices.length
      await evaluate(`(()=>{const v=document.querySelector('video');for(const time of [60.25,.25,33.75,${seconds}])v.currentTime=time;[...document.querySelectorAll('.export-options button')].find(b=>b.textContent.includes(${JSON.stringify(format)})).click()})()`)
      await until(`!document.querySelector('.export-modal')`,'frame export')
      const exportMilliseconds=performance.now()-began
      assert.equal(saveChoices.length,before+1)
      assert.ok(Math.abs(await evaluate(`document.querySelector('video').currentTime`)-seconds)<.001,'capture must not move the playhead')
      const output=await fs.readFile(path.join(directory,saveChoices.at(-1)))
      const expected=path.join(directory,'reference.png')
      ffmpeg(['-i',fixture,'-vf',`select=eq(n\\,${Math.round(seconds*12)})`,'-fps_mode','vfr','-frames:v','1',expected])
      const rendered=nativeImage.createFromBuffer(output),reference=nativeImage.createFromBuffer(await fs.readFile(expected))
      assert.deepEqual(rendered.getSize(),{width:321,height:481})
      const actualPixels=rendered.toBitmap(),expectedPixels=reference.toBitmap();let delta=0,bad=0,count=0
      for(let i=0;i<actualPixels.length;i++){if(i%4===3)continue;const diff=Math.abs(actualPixels[i]-expectedPixels[i]);delta+=diff;bad+=diff>24;count++}
      // JPEG's chroma subsampling affects saturated one-pixel test-pattern edges;
      // PNG retains the stricter pixel comparison against the decoder reference.
      assert.ok(delta/count<4&&bad/count<(format==='JPEG image'?.06:.015),`wrong decoded frame or color drift at ${seconds}: mean ${delta/count}, outliers ${bad/count}`)
      return {seconds,format,name:saveChoices.at(-1),exportMilliseconds,meanChannelError:delta/count,outlierFraction:bad/count}
    }
    const cases=[]
    for(const time of [1.25,1.75,64.25,.25,32.75])cases.push(await capture(time))
    assert.notEqual(cases[0].name,cases[1].name,'Two different frames within one second must get distinct suggested names')
    cases.push(await capture(15.25,'JPEG image'))

    // Changing the playhead during decoding must fail rather than label a wrong frame.
    await evaluate(`document.querySelector('[title="Export As (Ctrl+Shift+E)"]').click()`)
    await until(`Boolean(document.querySelector('.export-modal'))`,'race export dialog')
    const beforeRace=saveChoices.length
    await evaluate(`(()=>{const v=document.querySelector('video');v.currentTime=62.25;[...document.querySelectorAll('.export-options button')].find(b=>b.textContent.includes('PNG image')).click();setTimeout(()=>{v.currentTime=4.25},0)})()`)
    await until(`document.querySelector('.toast')?.textContent.includes('playhead moved')`,'changed playhead rejected')
    assert.equal(saveChoices.length,beforeRace,'stale frame must not reach a save dialog')
    await evaluate(`window.dispatchEvent(new KeyboardEvent('keydown',{key:'Escape',bubbles:true}))`)
    await until(`!document.querySelector('.export-modal')`,'close rejected export')
    cases.push(await capture(4.25))
    assert.equal(hash(await fs.readFile(fixture)),original)
    await fs.mkdir(path.join(root,'tmp'),{recursive:true})
    await fs.writeFile(path.join(root,'tmp','frame-stress.png'),(await win.webContents.capturePage()).toPNG())
    const result={passed:true,width:321,height:481,duration:65,fitGeometry,decoderReference:'FFmpeg',cases,rapidSeekCorrect:true,subsecondNamesDistinct:true,stalePlayheadRejected:true,recoveryExportPassed:true,sourceUnchanged:true}
    await fs.writeFile(path.join(root,'tmp','frame-stress-result.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result))
  }catch(error){console.error(error);exitCode=1}
  finally{for(const window of BrowserWindow.getAllWindows())window.destroy();const absolute=path.resolve(directory);assert.ok(absolute.startsWith(path.resolve(os.tmpdir())+path.sep)&&path.basename(absolute).startsWith('simple-video-frames-'));await Promise.race([fs.rm(absolute,{recursive:true,force:true,maxRetries:2,retryDelay:100}).catch(()=>{}),pause(2000)])}
  app.exit(exitCode)
})
