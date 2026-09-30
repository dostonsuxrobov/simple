// Render already-reviewed local print HTML through the real Electron PDF path.
// This writes a PDF only; it never sends a physical printer job.
const { app, BrowserWindow } = require('electron')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
let window, profile
;(async () => {
  const input = path.resolve(process.argv[2])
  const output = path.resolve(process.argv[3])
  profile = await fs.mkdtemp(path.join(os.tmpdir(), 'calc-pdf-qa-'))
  app.setPath('userData', profile)
  await app.whenReady()
  window = new BrowserWindow({show:false,webPreferences:{sandbox:true,contextIsolation:true,nodeIntegration:false}})
  await window.loadFile(input)
  await window.webContents.executeJavaScript('document.fonts.ready.then(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))))')
  const pdf = await window.webContents.printToPDF({preferCSSPageSize:true,printBackground:true,margins:{top:0,left:0,right:0,bottom:0}})
  await fs.writeFile(output,pdf)
  console.log(`PDF render verified: ${pdf.length} bytes`)
})().catch(error => { console.error(error); process.exitCode=1 }).finally(async () => {
  window?.destroy()
  if (profile) await fs.rm(profile,{recursive:true,force:true}).catch(()=>{})
  app.exit(process.exitCode || 0)
})
