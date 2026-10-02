'use strict'

// `simple --simple-self-test`: proves that the app as laid out on disk (the
// packaged EXE, or the unpackaged simple/ folder) has everything Save, Open,
// Import, Export and Combine need, so a missing or broken file is caught
// before a user meets it (the "converter absent from the bundle" class of bug):
// - every workspace bundle (modules/<mode>: main, preload, renderer, icon);
// - every shared I/O module in shared/ loads, and its data files parse;
// - the launcher's own modules load, launcher/main.cjs and the bootstrap
//   compile and every module they require is there, and the bundled Combine
//   worker compiles and carries every converter it needs (pdf-lib, mammoth,
//   SheetJS, JSZip, …);
// - MuPDF loads from this layout and unlocks a PDF that has only a
//   permissions password, as Combine does;
// - the format registry routes one sample per workspace;
// - safeWriteFile writes and verifies a file in a temporary folder;
// - the office engine probe answers without starting anything (an engine is
//   optional; Simple works without one);
// - inside Electron, the hardened HTML-to-PDF printer produces a PDF.
// Nothing outside a private temporary folder is written, and nothing leaves this PC.

const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const vm = require('node:vm')
const { createRequire } = require('node:module')
const { pathToFileURL } = require('node:url')

const DEFAULT_ROOT = path.resolve(__dirname, '..')
const MODES = Object.freeze(['docs', 'calc', 'pdf', 'image', 'video'])
const MODULE_FILES = Object.freeze(['dist/index.html', 'electron/main.cjs', 'electron/preload.cjs', 'build/icon.ico'])
/** Launcher and bootstrap files the unpackaged process requires directly. */
const LAUNCHER_MODULES = Object.freeze([
  'launcher/combine-host.cjs', 'launcher/combine-policy.cjs', 'launcher/combine-paths.cjs', 'launcher/associations.cjs', 'launcher/open-paths.cjs',
  'electron/routing.cjs', 'electron/launch.cjs', 'electron/open-list.cjs', 'electron/hand-off.cjs',
])
/**
 * Entry files that only run inside the app (they take the single-instance
 * lock when loaded), so they are compiled and their require() targets
 * resolved, never run.
 */
const ENTRY_FILES = Object.freeze(['launcher/main.cjs', 'electron/bootstrap.cjs'])
/**
 * Text each converter the Combine worker bundles leaves in it, so a bundle
 * built without one (the "converter absent from the bundle" class of bug)
 * fails here instead of in front of a user.
 */
const COMBINE_CONVERTERS = Object.freeze({
  'PDF assembly (pdf-lib)': 'PDFDocument',
  'Word layout (mammoth)': 'convertToHtml',
  'spreadsheet layout (SheetJS)': 'SheetJS',
  'Office packages (JSZip)': 'JSZip',
  'Simple Word pages': 'docxPrintJobs',
  'Simple spreadsheet pages': 'spreadsheetPrintJobs',
  'images to PDF pages': 'addImagePage',
})
/** One file name per workspace that the registry must route there. */
const ROUTE_SAMPLES = Object.freeze({ docs: 'a.docx', calc: 'a.xlsx', pdf: 'a.pdf', image: 'a.png', video: 'a.mp4' })

function existsReal(filePath) {
  try { return fs.statSync(filePath).isFile() } catch { return false }
}

/**
 * Compiles a CommonJS file without running it and resolves every require()
 * with a literal name in it, from its own folder.
 * @param {string} file absolute path
 * @returns {number} how many requires resolved
 * @throws {Error} naming a syntax error or the first require that does not resolve
 */
function checkEntryFile(file) {
  const source = fs.readFileSync(file, 'utf8')
  new vm.Script(`(function (exports, require, module, __filename, __dirname) {${source.replace(/^#!.*/, '')}\n})`, { filename: file })
  const resolveFrom = createRequire(file)
  const names = [...new Set([...source.matchAll(/\brequire\(\s*(['"])([^'"]+)\1\s*\)/g)].map((match) => match[2]))]
  for (const name of names) {
    // Electron's own module only exists inside Electron.
    if (name === 'electron' || name.startsWith('node:')) continue
    try { resolveFrom.resolve(name) } catch { throw new Error(`requires ${name}, which is missing`) }
  }
  return names.length
}

/** A one-page PDF with a correct cross-reference table, built in memory. */
function tinyPdf() {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 72 72] >>',
  ]
  let text = '%PDF-1.4\n'
  const offsets = objects.map((body, index) => {
    const offset = Buffer.byteLength(text, 'latin1')
    text += `${index + 1} 0 obj\n${body}\nendobj\n`
    return offset
  })
  const xref = Buffer.byteLength(text, 'latin1')
  text += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')}`
  text += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Buffer.from(text, 'latin1')
}

/**
 * Lists the shared modules and data files the app must carry: everything in
 * shared/electron, plus every electron/ entry shared/manifest.json vendors.
 * @param {string} root the simple/ folder (or the packaged app folder)
 * @returns {{modules: string[], data: string[]}} paths relative to root
 */
function sharedFiles(root) {
  const names = new Set()
  try {
    for (const entry of fs.readdirSync(path.join(root, 'shared', 'electron'))) names.add(`shared/electron/${entry}`)
  } catch {}
  try {
    const manifest = JSON.parse(fs.readFileSync(path.join(root, 'shared', 'manifest.json'), 'utf8').replace(/^﻿/, ''))
    for (const workspace of Object.values(manifest.workspaces || {})) {
      for (const source of workspace.electron || []) names.add(`shared/${source}`)
    }
  } catch {}
  const sorted = [...names].sort()
  return {
    modules: sorted.filter((name) => name.endsWith('.cjs')),
    data: sorted.filter((name) => name.endsWith('.json')),
  }
}

/**
 * Runs the self-test.
 * @param {object} [options]
 * @param {string} [options.root] the app folder (default: the folder above electron/)
 * @param {boolean} [options.print] also print a small PDF through html-to-pdf (Electron only, after app ready)
 * @param {number} [options.printTimeoutMs=20000]
 * @returns {Promise<{healthy: boolean, checks: Array<{name: string, ok: boolean, detail?: string}>, failed: string[]}>}
 */
async function runSelfTest(options = {}) {
  const root = path.resolve(options.root || DEFAULT_ROOT)
  const checks = []
  const check = async (name, work) => {
    try {
      const detail = await work()
      checks.push({ name, ok: true, ...(detail ? { detail: String(detail) } : {}) })
    } catch (error) {
      checks.push({ name, ok: false, detail: String((error && error.message) || error) })
    }
  }
  const load = (relative) => require(path.join(root, relative))

  for (const mode of MODES) {
    await check(`module ${mode}`, () => {
      const missing = MODULE_FILES.filter((file) => !fs.existsSync(path.join(root, 'modules', mode, file)))
      if (missing.length) throw new Error(`missing ${missing.join(', ')}`)
    })
  }

  const shared = sharedFiles(root)
  await check('shared I/O files are present', () => {
    if (!shared.modules.length) throw new Error('shared/electron has no modules')
    return `${shared.modules.length} modules, ${shared.data.length} data files`
  })
  for (const relative of shared.data) {
    await check(`parse ${relative}`, () => { JSON.parse(fs.readFileSync(path.join(root, relative), 'utf8').replace(/^﻿/, '')) })
  }
  for (const relative of shared.modules) {
    await check(`load ${relative}`, () => {
      const exported = load(relative)
      if (!exported || typeof exported !== 'object' || !Object.keys(exported).length) throw new Error('exports nothing')
    })
  }
  for (const relative of LAUNCHER_MODULES) await check(`load ${relative}`, () => { load(relative) })
  for (const relative of ENTRY_FILES) {
    await check(`compile ${relative} and resolve its requires`, () => `${checkEntryFile(path.join(root, relative))} requires`)
  }

  await check('Combine worker bundle compiles', () => {
    const { defaultWorkerPath } = load('launcher/combine-host.cjs')
    const workerPath = defaultWorkerPath()
    if (!existsReal(workerPath)) throw new Error(`missing ${workerPath}`)
    // Compile only: running it outside a worker thread is meaningless.
    const source = fs.readFileSync(workerPath, 'utf8')
    new vm.Script(source, { filename: workerPath })
    const missing = Object.entries(COMBINE_CONVERTERS).filter(([, marker]) => !source.includes(marker)).map(([name]) => name)
    if (missing.length) throw new Error(`the bundle has no ${missing.join(', ')}`)
    return `${Object.keys(COMBINE_CONVERTERS).length} converters bundled`
  })

  await check('Combine can unlock permissions-only PDFs (MuPDF)', async () => {
    const { mupdfPath, unlockPdfBytes } = load('launcher/combine-host.cjs')
    const file = mupdfPath()
    if (!file || !existsReal(file)) throw new Error('modules/pdf/vendor/mupdf is missing')
    for (const sibling of ['mupdf-wasm.js', 'mupdf-wasm.wasm']) {
      if (!existsReal(path.join(path.dirname(file), sibling))) throw new Error(`missing ${sibling} beside ${file}`)
    }
    // Load MuPDF (ES module and wasm) from this layout, protect a tiny PDF
    // with a permissions password only, and unlock it as Combine does.
    const mupdf = await import(pathToFileURL(file).href)
    const document = mupdf.Document.openDocument(tinyPdf(), 'application/pdf')
    let locked
    try {
      locked = Buffer.from(document.saveToBuffer('encrypt=aes-128,owner-password=simple-self-test,user-password=,permissions=-3904').asUint8Array())
    } finally {
      document.destroy?.()
    }
    if (!locked.includes('/Encrypt')) throw new Error('MuPDF did not protect the test PDF')
    const unlocked = Buffer.from(await unlockPdfBytes(locked))
    if (unlocked.subarray(0, 5).toString('latin1') !== '%PDF-' || unlocked.includes('/Encrypt')) throw new Error('the unlocked copy is not a plain PDF')
    return `unlocked a ${locked.length}-byte protected PDF`
  })

  await check('format registry routes every workspace', () => {
    const formats = load('shared/electron/formats.cjs')
    for (const [mode, name] of Object.entries(ROUTE_SAMPLES)) {
      const route = formats.routeForPath(path.join(os.tmpdir(), name))
      if (route.mode !== mode) throw new Error(`${name} routes to ${route.mode}, expected ${mode}`)
    }
  })

  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-self-test-'))
  try {
    await check('safeWriteFile writes and verifies', async () => {
      const { safeWriteFile } = load('shared/electron/safe-write.cjs')
      const target = path.join(scratch, 'self-test.txt')
      const first = await safeWriteFile(target, 'Simple self-test\n', { format: 'txt' })
      const second = await safeWriteFile(target, 'Simple self-test, again\n', { format: 'txt', expectedStamp: first.stamp })
      if (fs.readFileSync(target, 'utf8') !== 'Simple self-test, again\n') throw new Error('the file does not hold the new text')
      const leftovers = fs.readdirSync(scratch).filter((name) => name !== 'self-test.txt')
      if (leftovers.length) throw new Error(`left ${leftovers.join(', ')}`)
      return `${first.strategy}, then ${second.strategy}`
    })

    await check('office engine probe (optional engine)', async () => {
      const engine = load('shared/electron/office-engine.cjs')
      const status = await engine.getOfficeEngineStatus()
      if (!status || typeof status.available !== 'boolean') throw new Error('no status')
      return status.available ? `found (${status.source})` : `not on this PC (${status.reason}); Simple uses its own converters`
    })

    if (options.print) {
      await check('html-to-pdf prints a PDF', async () => {
        const { printHtmlToPdf } = load('shared/electron/html-to-pdf.cjs')
        const bytes = await printHtmlToPdf('<p>Simple self-test</p>', { title: 'Self-test', timeoutMs: options.printTimeoutMs || 20_000 })
        if (!bytes || Buffer.from(bytes).subarray(0, 5).toString('latin1') !== '%PDF-') throw new Error('the output is not a PDF')
        return `${bytes.length} bytes`
      })
    }
  } finally {
    fs.rmSync(scratch, { recursive: true, force: true })
  }

  const failed = checks.filter((item) => !item.ok).map((item) => `${item.name}: ${item.detail}`)
  return { healthy: !failed.length, checks, failed }
}

module.exports = { COMBINE_CONVERTERS, ENTRY_FILES, MODES, checkEntryFile, runSelfTest, sharedFiles }
