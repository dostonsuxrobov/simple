'use strict'

// Exercise the actual Electron main process, isolated preload, renderer and
// installed office engine. Only native file choices are replaced. No source
// documents are saved, and no physical printer is contacted.
const { app, BrowserWindow, dialog } = require('electron')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const crypto = require('node:crypto')
const { pathToFileURL } = require('node:url')
const JSZip = require('jszip')
const root = path.resolve(__dirname, '..')
const directory = require('node:fs').mkdtempSync(path.join(os.tmpdir(), 'simple-pdf-word-'))
const selections = []
const openOptions = []
let saveAttempts = 0
app.setPath('userData', path.join(directory, 'profile'))
dialog.showOpenDialog = async (...args) => {
  openOptions.push(args.at(-1))
  return { canceled: false, filePaths: selections.shift() || [] }
}
dialog.showSaveDialog = async () => { saveAttempts += 1; return { canceled: true } }
require('../electron/main.cjs')

const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms))
const hash = (bytes) => crypto.createHash('sha256').update(bytes).digest('hex')
const bytesExpression = (bytes) => `Uint8Array.from(atob(${JSON.stringify(Buffer.from(bytes).toString('base64'))}), c => c.charCodeAt(0))`
const evaluate = (window, code) => window.webContents.executeJavaScript(code)
async function until(window, code, label, timeout = 75_000) {
  const deadline = Date.now() + timeout
  while (!(await evaluate(window, code))) {
    if (Date.now() > deadline) throw new Error(`Timed out: ${label}; ${await evaluate(window, `document.querySelector('.toast')?.textContent`)}`)
    await pause(30)
  }
}
async function text(pdf, number = 1) {
  return (await (await pdf.getPage(number)).getTextContent()).items.map((item) => item.str || '').join(' ').replace(/\s+/g, ' ').trim()
}
async function thumbnailPainted(window) {
  await until(window, `(()=>{const c=document.querySelector('.thumbnail-paper canvas');if(!c || c.width < 100)return false;const p=c.getContext('2d').getImageData(0,0,c.width,c.height).data;return p.some((v,i)=>i%4!==3 && v<180)})()`, 'page thumbnail painted')
}
async function createLandscapeDocx() {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>`)
  zip.file('_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`)
  zip.file('word/_rels/document.xml.rels', `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>`)
  zip.file('word/header1.xml', `<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Landscape header retained</w:t></w:r></w:p></w:hdr>`)
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:rPr><w:b/><w:color w:val="1254A1"/><w:sz w:val="48"/></w:rPr><w:t>Landscape import fixture</w:t></w:r></w:p><w:p><w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>Original page size, heading, header and table must survive.</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="12960" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/><w:insideV w:val="single" w:sz="8"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="6480"/><w:gridCol w:w="6480"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="6480" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Left cell: 123.45</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="6480" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Right cell: 678.90</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`)
  return zip.generateAsync({ type: 'nodebuffer' })
}

app.whenReady().then(async () => {
  let exitCode = 0
  const pdfs = []
  try {
    let window
    while (!(window = BrowserWindow.getAllWindows()[0])) await pause(10)
    await until(window, `Boolean(window.simple && [...document.querySelectorAll('button')].find(b => b.textContent === 'Choose file'))`, 'welcome page')
    const docx = path.join(directory, 'landscape.docx')
    const doc = path.join(root, '..', '.codex-tmp', 'reliability-reference', 'test_doc.doc')
    const docxBytes = await createLandscapeDocx()
    await fs.writeFile(docx, docxBytes)
    const originalHash = hash(await fs.readFile(doc))
    const originalDocxHash = hash(docxBytes)
    const pdfjs = await import(pathToFileURL(path.join(root, 'node_modules', 'pdfjs-dist', 'legacy', 'build', 'pdf.mjs')).href)
    const loadPdf = async (bytes) => {
      const pdf = await pdfjs.getDocument({ data: new Uint8Array(bytes), disableWorker: true, isEvalSupported: false }).promise
      pdfs.push(pdf)
      return pdf
    }

    // The visible Choose file path must use the real conversion IPC handler.
    selections.push([docx])
    const landscapeStarted = performance.now()
    await evaluate(window, `[...document.querySelectorAll('button')].find(b => b.textContent === 'Choose file').click()`)
    await until(window, `document.querySelector('.text-layer')?.textContent.includes('Landscape import fixture') && document.querySelector('.page-canvas')?.width > 0`, 'landscape document painted')
    const landscapeOpenMilliseconds = performance.now() - landscapeStarted
    const landscapePayload = await evaluate(window, `(async()=>{const p=await window.simple.openPath(${JSON.stringify(docx)}); return {...p,data:Array.from(p.data)}})()`)
    assert.equal(landscapePayload.converted, true)
    assert.equal(landscapePayload.path, null, 'the PDF must not save over the Word original')
    assert.equal(landscapePayload.sourcePath, docx)
    assert.equal(landscapePayload.name, 'landscape.pdf')
    const landscape = await loadPdf(landscapePayload.data)
    assert.equal(landscape.numPages, 1)
    assert.deepEqual((await landscape.getPage(1)).view, [0, 0, 792, 612], 'source landscape Letter dimensions must survive')
    const landscapeText = await text(landscape)
    for (const fragment of ['Landscape header retained', 'Landscape import fixture', 'Left cell: 123.45', 'Right cell: 678.90']) assert.ok(landscapeText.includes(fragment), fragment)
    const items = (await (await landscape.getPage(1)).getTextContent()).items
    const heading = items.find((item) => item.str === 'Landscape import fixture')
    assert.ok(heading.height >= 23.9 && heading.height <= 24.1, '24pt heading must retain its size')
    const left = items.find((item) => item.str.includes('Left cell:'))
    const right = items.find((item) => item.str.includes('Right cell:'))
    assert.ok(right.transform[4] - left.transform[4] > 300, 'table columns must remain in their original positions')
    assert.ok(Math.abs(right.transform[5] - left.transform[5]) < 1, 'table cells must stay on the same row')
    await fs.mkdir(path.join(root, 'tmp'), { recursive: true })
    await thumbnailPainted(window)
    await fs.writeFile(path.join(root, 'tmp', 'word-import-landscape.png'), (await window.webContents.capturePage()).toPNG())

    // Drag/drop uses open-bytes; the exact source gets the same cached layout.
    const dropped = await evaluate(window, `(async()=>{const p=await window.simple.openBytes('landscape.docx', ${bytesExpression(docxBytes)}.buffer);return {...p,data:Array.from(p.data)}})()`)
    assert.equal(dropped.sourcePath, null)
    assert.equal(dropped.path, null)
    assert.equal(dropped.converted, true)
    assert.deepEqual(Buffer.from(dropped.data), Buffer.from(landscapePayload.data))

    // Opening another source follows the app's own new-window behavior.
    const previousIds = new Set(BrowserWindow.getAllWindows().map((item) => item.id))
    const docStarted = performance.now()
    await evaluate(window, `window.simple.openInNewWindow(${JSON.stringify(doc)})`)
    let sampleWindow
    while (!(sampleWindow = BrowserWindow.getAllWindows().find((item) => !previousIds.has(item.id)))) await pause(10)
    await until(sampleWindow, `document.querySelector('.text-layer')?.textContent.includes('JUNE 10, 2026') && document.querySelector('.page-canvas')?.width > 0`, 'sample document painted with its cached date')
    const sampleOpenMilliseconds = performance.now() - docStarted
    const samplePayload = await evaluate(sampleWindow, `(async()=>{const p=await window.simple.openPath(${JSON.stringify(doc)});return {...p,data:Array.from(p.data)}})()`)
    assert.equal(samplePayload.converted, true)
    assert.equal(samplePayload.path, null)
    const sample = await loadPdf(samplePayload.data)
    const reference = await loadPdf(await fs.readFile(path.join(root, '..', '.codex-tmp', 'reliability-reference', 'test_doc-google-reference.pdf')))
    assert.equal(sample.numPages, 1)
    assert.equal(await text(sample), await text(reference), 'all sample text must equal the original Google render, including its cached date')
    const visibleSample = await evaluate(sampleWindow, `({pages:document.querySelectorAll('.continuous-page-slot').length,text:document.querySelector('.text-layer')?.textContent,canvasWidth:document.querySelector('.page-canvas')?.width,canvasHeight:document.querySelector('.page-canvas')?.height})`)
    assert.equal(visibleSample.pages, 1)
    assert.ok(visibleSample.canvasWidth > 500 && visibleSample.canvasHeight > 600)
    await thumbnailPainted(sampleWindow)
    await fs.writeFile(path.join(root, 'tmp', 'word-import-sample.png'), (await sampleWindow.webContents.capturePage()).toPNG())

    // Add pages has a separate IPC route: cover both its native picker and drop.
    selections.push([doc])
    const inserted = await evaluate(window, `(async()=>{const p=await window.simple.insertFiles(${bytesExpression(landscapePayload.data)},1);return {...p,data:Array.from(p.data)}})()`)
    assert.ok(openOptions.at(-1).filters[0].extensions.includes('doc'), 'Add pages must offer legacy Word files')
    assert.equal(inserted.added, 1)
    const insertedPdf = await loadPdf(inserted.data)
    assert.equal(insertedPdf.numPages, 2)
    assert.equal(await text(insertedPdf, 2), await text(reference))
    const droppedInsertion = await evaluate(window, `(async()=>{const p=await window.simple.insertDroppedFiles(${bytesExpression(samplePayload.data)},0,[{name:'landscape.docx',data:${bytesExpression(docxBytes)}}]);return {...p,data:Array.from(p.data)}})()`)
    assert.equal(droppedInsertion.added, 1)
    const droppedPdf = await loadPdf(droppedInsertion.data)
    assert.equal(droppedPdf.numPages, 2)
    assert.deepEqual((await droppedPdf.getPage(1)).view, [0, 0, 792, 612])
    assert.equal(await text(droppedPdf, 2), await text(reference))
    assert.equal(hash(await fs.readFile(doc)), originalHash)
    assert.equal(hash(await fs.readFile(docx)), originalDocxHash)
    assert.equal(saveAttempts, 0)
    const result = { passed: true, landscape: { pages: 1, width: 792, height: 612, headerAndTablePreserved: true, headingPoints: heading.height, openMilliseconds: landscapeOpenMilliseconds }, userDoc: { pages: 1, cachedDate: 'JUNE 10, 2026', allTextMatchesGoogleReference: true, openMilliseconds: sampleOpenMilliseconds }, pathAndBytesIpc: true, addPagesPickerAndDrop: true, originalsUnchanged: true, noSaveOrPrint: true }
    await fs.writeFile(path.join(root, 'tmp', 'word-import-result.json'), JSON.stringify(result, null, 2))
    console.log(JSON.stringify(result))
  } catch (error) { console.error(error); exitCode = 1 }
  finally {
    for (const pdf of pdfs) await pdf.destroy().catch(() => {})
    for (const window of BrowserWindow.getAllWindows()) window.destroy()
    const absolute = path.resolve(directory)
    assert.ok(absolute.startsWith(path.resolve(os.tmpdir()) + path.sep) && path.basename(absolute).startsWith('simple-pdf-word-'))
    await fs.rm(absolute, { recursive: true, force: true, maxRetries: 5, retryDelay: 150 }).catch(() => {})
  }
  app.exit(exitCode)
})
