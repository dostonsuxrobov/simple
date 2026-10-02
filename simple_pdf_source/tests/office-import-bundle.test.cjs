'use strict'

// The unified build bundles electron/main.cjs with esbuild (electron and
// mupdf stay external). This bundles it the same way and runs the bundle
// with a stand-in Electron whose "print" returns a fixed PDF, proving the
// converters (mammoth, SheetJS, markdown-it) are inside the bundle and run.
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawnSync } = require('node:child_process')
const test = require('node:test')
const JSZip = require('jszip')

const ROOT = path.join(__dirname, '..')

const STAND_IN = `
const Module = require('node:module')
const handlers = new Map()
const printed = []
const PDF = Buffer.from('%PDF-1.4\\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj 3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]>>endobj\\ntrailer<</Root 1 0 R>>\\n%%EOF\\n')
class WebContents {
  constructor() { this.id = Math.floor(Math.random() * 1e6) }
  on() {} once() {} send() {} setWindowOpenHandler() {} setAudioMuted() {}
  // The app window's own UI (dist/index.html) is not built here; only conversion pages are kept.
  async loadFile(file) { try { printed.push(require('node:fs').readFileSync(file, 'utf8')) } catch {} }
  async printToPDF() { return PDF }
}
class BrowserWindow {
  constructor() { this.webContents = new WebContents() }
  static fromWebContents() { return null }
  static getAllWindows() { return [] }
  static getFocusedWindow() { return null }
  loadFile(file) { return this.webContents.loadFile(file) }
  loadURL() { return Promise.resolve() }
  removeMenu() {} once() {} on() {} show() {} close() {} destroy() { this.destroyed = true }
  isDestroyed() { return Boolean(this.destroyed) }
}
const session = { fromPartition: () => ({ webRequest: { onBeforeRequest() {} }, setPermissionRequestHandler() {}, setPermissionCheckHandler() {}, on() {} }) }
const electron = {
  app: { isPackaged: true, requestSingleInstanceLock: () => true, on() {}, whenReady: () => Promise.resolve(), getVersion: () => 'test', quit() {}, getLocaleCountryCode: () => 'US' },
  BrowserWindow, session,
  dialog: { async showOpenDialog() { return { canceled: true, filePaths: [] } }, async showSaveDialog() { return { canceled: true } } },
  ipcMain: { handle: (channel, handler) => handlers.set(channel, handler), on() {} },
  nativeImage: { createFromPath: () => ({}) },
  shell: { showItemInFolder() {}, async openExternal() {} },
}
const load = Module._load
Module._load = function (request, parent, isMain) {
  if (request === 'electron') return electron
  if (request === 'mupdf') return {}
  return load.call(this, request, parent, isMain)
}
require(process.env.BUNDLE)
;(async () => {
  await new Promise((resolve) => setTimeout(resolve, 50))
  const results = {}
  for (const file of JSON.parse(process.env.INPUTS)) {
    try {
      const payload = await handlers.get('file:open-path')({ sender: new WebContents() }, file)
      results[require('node:path').basename(file)] = { converted: payload.converted, bytes: payload.data.length, html: printed.at(-1) || '' }
    } catch (error) {
      results[require('node:path').basename(file)] = { error: String(error) }
    }
  }
  process.stdout.write(JSON.stringify(results))
})()
`

test('the esbuild bundle of main.cjs carries the converters and runs them', { timeout: 180_000 }, async () => {
  const esbuild = require('esbuild')
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-pdf-bundle-'))
  try {
    const outfile = path.join(directory, 'electron', 'main.cjs')
    const result = await esbuild.build({
      entryPoints: [path.join(ROOT, 'electron', 'main.cjs')],
      outfile,
      bundle: true,
      platform: 'node',
      format: 'cjs',
      target: 'node22',
      external: ['electron', 'mupdf'],
      legalComments: 'none',
      logLevel: 'silent',
      metafile: true,
    })
    const inputs = Object.keys(result.metafile.inputs).map((input) => input.replaceAll('\\', '/'))
    for (const expected of ['electron/office-import.cjs', 'electron/image-to-pdf.cjs', 'node_modules/mammoth/', 'node_modules/xlsx/', 'node_modules/markdown-it/', 'node_modules/jszip/']) {
      assert.ok(inputs.some((input) => input.includes(expected)), `${expected} is bundled`)
    }
    const bundle = fs.readFileSync(outfile, 'utf8')
    for (const name of ['mammoth', 'xlsx', 'markdown-it', 'jszip', 'word-extractor']) assert.doesNotMatch(bundle, new RegExp(`require\\(["']${name}["']\\)`), `${name} is not left external`)

    const docx = new JSZip()
    docx.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
    docx.file('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    docx.file('word/document.xml', '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:color w:val="1254A1"/></w:rPr><w:t>BUNDLED_WORD_TEXT</w:t></w:r></w:p></w:body></w:document>')
    const files = {
      'letter.docx': await docx.generateAsync({ type: 'nodebuffer' }),
      'notes.md': Buffer.from('# BUNDLED_MARKDOWN\n\n**bold**\n'),
      'table.csv': Buffer.from('Item,Amount\nBUNDLED_CELL,"1,234.50"\n'),
    }
    for (const [name, data] of Object.entries(files)) fs.writeFileSync(path.join(directory, name), data)
    const run = spawnSync(process.execPath, ['-e', STAND_IN], {
      env: { ...process.env, BUNDLE: outfile, INPUTS: JSON.stringify(Object.keys(files).map((name) => path.join(directory, name))), SIMPLE_FORCE_NO_OFFICE: '1', TEMP: directory, TMP: directory },
      encoding: 'utf8',
      timeout: 120_000,
      windowsHide: true,
    })
    assert.equal(run.status, 0, run.stderr)
    const results = JSON.parse(run.stdout)
    assert.equal(results['letter.docx'].converted, true, JSON.stringify(results['letter.docx']))
    assert.match(results['letter.docx'].html, /color: #1254a1">BUNDLED_WORD_TEXT/)
    assert.match(results['notes.md'].html, /<h1>BUNDLED_MARKDOWN<\/h1>/)
    assert.match(results['notes.md'].html, /<strong>bold<\/strong>/)
    assert.match(results['table.csv'].html, /<td>BUNDLED_CELL<\/td><td[^>]*>1,234\.50<\/td>/)
  } finally {
    fs.rmSync(directory, { recursive: true, force: true, maxRetries: 3 })
  }
})
