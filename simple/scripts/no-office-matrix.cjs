'use strict'

/**
 * Engine-absent matrix (design §5.3, §10.2): Simple must be fully usable on a
 * PC without LibreOffice. The optional office engine is only ever used when
 * it is already installed locally; this script proves nothing depends on it.
 * SIMPLE_FORCE_NO_OFFICE=1 is set for the whole run.
 *
 * Shared parts (always; Node only, no app is started):
 *   registry   For every workspace in simple/shared/electron/formats.json:
 *              every extension an Open dialog offers opens without the engine;
 *              Save never writes a legacy binary (.doc, .xls, any compound
 *              file) and never needs the engine (it goes to a modern sibling);
 *              every Export row that needs the engine is shown disabled, and a
 *              legacy binary is exported only as an explicit, lossy
 *              "values only" choice.
 *   engine     The probe reports no engine (forced-off). Every conversion
 *              request fails with a coded NEEDS_OFFICE_ENGINE that says what
 *              Simple does instead, never "download" or "install".
 *   save       For every format whose Save goes to a sibling file, the real
 *              performSave pipeline (document-guard.cjs, electron stubbed,
 *              prompts answered through SIMPLE_QA_DIALOGS) saves an edited
 *              copy of a sample: the edits land in the sibling (.xlsx, .docx,
 *              .pdf, .png) next to it, the sibling validates, has no legacy
 *              binary signature, the original's SHA-256 is unchanged, the
 *              second save goes straight to the sibling, and no prompt asks
 *              the user to download or install anything.
 *
 * Workspace parts (Windows; only workspaces ENABLED in simple/shared/manifest.json):
 *   Each enabled workspace's acceptance adapter (scripts/io-adapters/<mode>.cjs)
 *   lists `matrix` rows. Each row opens a copy of a sample in the real app
 *   with the engine forced off, and either saves an edit (expect 'sibling' or
 *   'in-place') or reads the notice the workspace shows (expect 'notice').
 *
 *   node scripts/no-office-matrix.cjs [--shared-only] [--module=calc] [--keep] [--json]
 * Exit code 0 when every row passed.
 */

process.env.SIMPLE_FORCE_NO_OFFICE = '1'

const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const Module = require('node:module')
const { EventEmitter } = require('node:events')
const { pathToFileURL } = require('node:url')

const ROOT = path.resolve(__dirname, '..')
const SHARED = path.join(ROOT, 'shared', 'electron')
const SAMPLES = path.resolve(ROOT, '..', 'Simple test examples')
const CFB_SIGNATURE = Buffer.from('d0cf11e0a1b11ae1', 'hex')
/** Words a message must never use: Simple never asks the user to fetch or set up software. */
const SETUP_WORDING = /\b(?:download|downloads|downloading|install|installs|installing|installer|reinstall)\b/i
const WORKSPACES = Object.freeze(['pdf', 'calc', 'docs', 'image', 'video'])

const formats = require(path.join(SHARED, 'formats.cjs'))

function sha256(bytes) {
  return crypto.createHash('sha256').update(bytes).digest('hex')
}

function isLegacyBinary(format) {
  return Boolean(format) && (format.validator === 'cfb' || (format.magic || []).some((magic) => String(magic.hex || '').toLowerCase() === 'd0cf11e0a1b11ae1'))
}

// ---------------------------------------------------------------------------
// Registry
// ---------------------------------------------------------------------------

/**
 * Checks the format registry's engine-absent policy for one workspace.
 * @param {string} workspace
 * @param {ReturnType<typeof formats.createRegistry>} [registry] default: the shared registry (tests pass a modified one)
 * @returns {string[]} problems
 */
function registryProblems(workspace, registry = formats) {
  const problems = []
  const options = { engine: false }
  const filters = registry.dialogFilters(workspace, options)
  const offered = new Set()
  for (const filter of filters) {
    for (const extension of filter.extensions || []) if (extension !== '*') offered.add(`.${String(extension).toLowerCase()}`)
  }
  for (const extension of offered) {
    const format = registry.formatForExtension(extension)
    if (!format) {
      problems.push(`${workspace}: the Open dialog offers ${extension}, which the registry does not know`)
      continue
    }
    const candidates = registry.formats.filter((item) => item.extensions.includes(extension))
    const opens = candidates.some((item) => {
      const policy = registry.openPolicy(item.id, workspace, options)
      return policy && policy.available
    })
    if (!opens) problems.push(`${workspace}: the Open dialog offers ${extension}, but it does not open without the office engine`)
  }
  for (const format of registry.openableFormats(workspace, options)) {
    const save = registry.savePolicy(format.id, workspace, options)
    if (!save) continue
    if (save.target === 'in-place') {
      if (save.usesEngine) problems.push(`${workspace}: saving ${format.id} in place needs the office engine; it must go to a sibling file`)
      if (isLegacyBinary(format)) problems.push(`${workspace}: Save writes the legacy binary ${format.id} without the office engine`)
    } else if (save.target === 'sibling') {
      const target = registry.formatById(save.formatId)
      if (!target) problems.push(`${workspace}: ${format.id} saves to the unknown sibling format ${save.formatId}`)
      else if (isLegacyBinary(target)) problems.push(`${workspace}: ${format.id} saves to the legacy binary ${target.id}`)
      else if (save.formatId === format.id) problems.push(`${workspace}: ${format.id} names itself as its sibling format`)
    }
  }
  for (const row of registry.exportFormats(workspace, options)) {
    const format = registry.formatById(row.id)
    if (row.available && row.usesEngine && registry.operation(row.id, workspace, 'export')?.engine === 'required') {
      problems.push(`${workspace}: Export offers ${row.id}, which needs the office engine`)
    }
    if (!row.available && row.reason !== 'needs-office-engine') problems.push(`${workspace}: Export lists ${row.id} as unavailable without saying why`)
    if (row.available && isLegacyBinary(format) && !(row.mode === 'values-only' && row.lossy && Array.isArray(row.lossy.loses) && row.lossy.loses.length)) {
      problems.push(`${workspace}: Export offers the legacy binary ${row.id} without the explicit "values only" loss list`)
    }
    for (const text of [row.label, row.description, row.reason].filter(Boolean)) {
      if (SETUP_WORDING.test(text)) problems.push(`${workspace}: Export row ${row.id} asks the user to download or install something: "${text}"`)
    }
  }
  return problems
}

// ---------------------------------------------------------------------------
// Engine
// ---------------------------------------------------------------------------

/**
 * Checks the probe and every conversion request with the engine forced off.
 * @returns {Promise<{problems: string[], pairs: number}>}
 */
async function engineProblems() {
  const engine = require(path.join(SHARED, 'office-engine.cjs'))
  const problems = []
  engine.clearOfficeEngineStatusCache()
  const status = await engine.getOfficeEngineStatus({ refresh: true })
  if (status.available !== false) problems.push('the office engine probe reports an engine while SIMPLE_FORCE_NO_OFFICE=1')
  if (status.reason !== 'forced-off') problems.push(`the office engine probe gives reason "${status.reason}", expected "forced-off"`)
  let pairs = 0
  for (const inputExtension of engine.INPUT_EXTENSIONS) {
    for (const outputExtension of engine.OUTPUT_EXTENSIONS) {
      if (inputExtension === outputExtension) continue
      pairs += 1
      const label = `${inputExtension} → ${outputExtension}`
      try {
        await engine.convertOfficeBytes({ bytes: Buffer.from('Simple engine-absent check'), inputExtension, outputExtension }, { cache: false })
        problems.push(`${label}: converted although no engine is available`)
      } catch (error) {
        if (!engine.isOfficeEngineError(error)) {
          problems.push(`${label}: threw an uncoded error (${error && error.message})`)
          continue
        }
        if (error.code === 'UNSUPPORTED_CONVERSION') continue
        if (error.code !== 'NEEDS_OFFICE_ENGINE') problems.push(`${label}: failed with ${error.code}, expected NEEDS_OFFICE_ENGINE`)
        if (SETUP_WORDING.test(String(error.message))) problems.push(`${label}: the message asks the user to download or install something: "${error.message}"`)
        if (['doc', 'xls', 'odt', 'ods'].includes(outputExtension)) {
          const expected = outputExtension === 'doc' || outputExtension === 'odt' ? 'docx' : 'xlsx'
          if (error.altFormat !== expected) problems.push(`${label}: offers ${error.altFormat || 'nothing'} instead, expected ${expected}`)
        }
      }
    }
  }
  return { problems, pairs }
}

// ---------------------------------------------------------------------------
// Save pipeline (electron stubbed)
// ---------------------------------------------------------------------------

/** A ZIP with stored entries (enough for the structural validators). */
function storedZip(files) {
  const locals = []
  const centrals = []
  let offset = 0
  for (const [name, content] of files) {
    const data = Buffer.from(content)
    const nameBytes = Buffer.from(name)
    const crc = zlib.crc32(data) >>> 0
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBytes.length, 26)
    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(20, 4)
    central.writeUInt16LE(20, 6)
    central.writeUInt32LE(crc, 16)
    central.writeUInt32LE(data.length, 20)
    central.writeUInt32LE(data.length, 24)
    central.writeUInt16LE(nameBytes.length, 28)
    central.writeUInt32LE(offset, 42)
    locals.push(local, nameBytes, data)
    centrals.push(central, nameBytes)
    offset += 30 + nameBytes.length + data.length
  }
  const directory = Buffer.concat(centrals)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(files.length, 8)
  end.writeUInt16LE(files.length, 10)
  end.writeUInt32LE(directory.length, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, directory, end])
}

const RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
const TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
const OFFICE_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'

function pngBytes() {
  const chunk = (type, data) => {
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data])
    const length = Buffer.alloc(4)
    length.writeUInt32BE(data.length)
    const crc = Buffer.alloc(4)
    crc.writeUInt32BE(zlib.crc32(body) >>> 0)
    return Buffer.concat([length, body, crc])
  }
  const header = Buffer.alloc(13)
  header.writeUInt32BE(1, 0)
  header.writeUInt32BE(1, 4)
  header[8] = 8
  header[9] = 2
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', header),
    chunk('IDAT', zlib.deflateSync(Buffer.from([0, 255, 255, 255]))),
    chunk('IEND', Buffer.alloc(0)),
  ])
}

function pdfBytes(marker) {
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>`,
  ]
  const text = `BT /F1 12 Tf 20 100 Td (${marker}) Tj ET`
  objects.push(`<< /Length ${text.length} >>\nstream\n${text}\nendstream`)
  let body = '%PDF-1.4\n'
  const offsets = []
  objects.forEach((object, index) => {
    offsets.push(body.length)
    body += `${index + 1} 0 obj\n${object}\nendobj\n`
  })
  const xref = body.length
  body += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n${offsets.map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('')}`
  body += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`
  return Buffer.from(body, 'latin1')
}

/** Minimal, structurally valid bytes of a format that carry `marker`, as a workspace serializer would produce. */
const FIXTURE_WRITERS = Object.freeze({
  xlsx: (marker) => storedZip([
    ['[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="${TYPES_NS}"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>`],
    ['_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${OFFICE_DOCUMENT}" Target="xl/workbook.xml"/></Relationships>`],
    ['xl/workbook.xml', `<?xml version="1.0"?><workbook><!-- ${marker} --></workbook>`],
  ]),
  docx: (marker) => storedZip([
    ['[Content_Types].xml', `<?xml version="1.0"?><Types xmlns="${TYPES_NS}"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`],
    ['_rels/.rels', `<?xml version="1.0"?><Relationships xmlns="${RELS_NS}"><Relationship Id="rId1" Type="${OFFICE_DOCUMENT}" Target="word/document.xml"/></Relationships>`],
    ['word/document.xml', `<?xml version="1.0"?><document><!-- ${marker} --></document>`],
  ]),
  pdf: (marker) => pdfBytes(marker),
  png: () => pngBytes(),
})

/** Installs a minimal electron stand-in (windows, message boxes and dialogs) for the save rows. */
function installElectronStub(appPath) {
  const windows = new Set()
  let nextId = 1
  class FakeWebContents extends EventEmitter {
    constructor() {
      super()
      this.id = nextId++
      this.destroyed = false
      this.mainFrame = { url: pathToFileURL(path.join(appPath, 'dist', 'index.html')).href }
    }
    send() {}
    isDestroyed() { return this.destroyed }
    destroy() { if (!this.destroyed) { this.destroyed = true; this.emit('destroyed') } }
  }
  class BrowserWindow extends EventEmitter {
    constructor() {
      super()
      this.webContents = new FakeWebContents()
      this.destroyed = false
      windows.add(this)
    }
    static fromWebContents(contents) { return [...windows].find((win) => win.webContents === contents) || null }
    static getAllWindows() { return [...windows] }
    isDestroyed() { return this.destroyed }
    destroy() { if (!this.destroyed) { this.destroyed = true; windows.delete(this); this.webContents.destroy(); this.emit('closed') } }
    focus() {}
    show() {}
    isMinimized() { return false }
    restore() {}
  }
  const unexpected = []
  const stub = {
    app: Object.assign(new EventEmitter(), { isPackaged: false, getPath: () => path.join(appPath, 'unused'), getAppPath: () => appPath, getVersion: () => '0.0.0' }),
    BrowserWindow,
    dialog: {
      // Scripted answers come from SIMPLE_QA_DIALOGS; reaching Electron's dialogs means a row asked something unscripted.
      async showSaveDialog(win, options) { unexpected.push(`Save dialog: ${(options || win).title || ''}`); return { canceled: true, filePath: '' } },
      async showOpenDialog() { unexpected.push('Open dialog'); return { canceled: true, filePaths: [] } },
      async showMessageBox(win, options) { const opts = options || win; unexpected.push(`message box: ${opts.message}`); return { response: opts.cancelId ?? 0, checkboxChecked: false } },
    },
    ipcMain: Object.assign(new EventEmitter(), { handle() {}, removeHandler() {} }),
    shell: { showItemInFolder() {}, openPath: async (target) => { unexpected.push(`shell.openPath ${target}`); return '' }, openExternal: async (target) => { unexpected.push(`shell.openExternal ${target}`) } },
  }
  const originalResolve = Module._resolveFilename
  Module._resolveFilename = function resolveElectronStub(request, ...rest) {
    if (request === 'electron') return 'electron-no-office-stub'
    return originalResolve.call(this, request, ...rest)
  }
  require.cache['electron-no-office-stub'] = { id: 'electron-no-office-stub', filename: 'electron-no-office-stub', loaded: true, exports: stub }
  return {
    BrowserWindow,
    unexpected,
    restore() {
      Module._resolveFilename = originalResolve
      delete require.cache['electron-no-office-stub']
    },
  }
}

/** The sample for an extension, if "Simple test examples" has one. */
function sampleFor(extension) {
  try {
    const name = fs.readdirSync(SAMPLES).find((file) => path.extname(file).toLowerCase() === extension)
    return name ? path.join(SAMPLES, name) : null
  } catch {
    return null
  }
}

/**
 * Every (workspace, format) whose Save goes to a sibling file without the engine.
 * @param {readonly string[]} [workspaces]
 * @returns {Array<{workspace: string, format: object, sibling: {formatId: string, extension: string}}>}
 */
function siblingRows(workspaces = WORKSPACES) {
  const rows = []
  for (const workspace of workspaces) {
    for (const format of formats.openableFormats(workspace, { engine: false })) {
      const save = formats.savePolicy(format.id, workspace, { engine: false })
      if (save && save.target === 'sibling') rows.push({ workspace, format, sibling: { formatId: save.formatId, extension: save.extension } })
    }
  }
  return rows
}

/**
 * Saves an edited copy of a sample for every sibling row through performSave.
 * @param {{root: string}} run scratch folder
 * @returns {Promise<Array<{name: string, ok: boolean, detail: string}>>}
 */
async function saveRows(run) {
  const outcomes = []
  const appPath = path.join(run.root, 'app')
  const electron = installElectronStub(appPath)
  const previousDialogs = process.env.SIMPLE_QA_DIALOGS
  try {
    const core = require(path.join(SHARED, 'io-core.cjs'))
    const stores = require(path.join(SHARED, 'stores.cjs'))
    const dialogs = require(path.join(SHARED, 'io-dialogs.cjs'))
    const guard = require(path.join(SHARED, 'document-guard.cjs'))
    const { validateFile } = require(path.join(SHARED, 'validators.cjs'))
    const quiet = { warn() {}, info() {} }
    core.configureIo({ journalDir: path.join(run.root, 'journal'), logger: quiet })
    stores.configureStores({ recoveryDir: path.join(run.root, 'recovery'), versionsDir: path.join(run.root, 'versions'), logger: quiet })
    const documents = path.join(run.root, 'Documents')
    fs.mkdirSync(documents, { recursive: true })

    let index = 0
    for (const row of siblingRows()) {
      index += 1
      const name = `save ${row.workspace} ${row.format.id} → sibling .${row.sibling.formatId}`
      const folder = path.join(run.root, `${String(index).padStart(2, '0')}-${row.workspace}-${row.format.id}`)
      fs.mkdirSync(folder, { recursive: true })
      const win = new electron.BrowserWindow()
      const dialogLog = path.join(folder, 'dialogs.log')
      const before = electron.unexpected.length
      try {
        const writer = FIXTURE_WRITERS[row.sibling.formatId]
        if (!writer) throw new Error(`no fixture writer for .${row.sibling.formatId}; add one to FIXTURE_WRITERS in scripts/no-office-matrix.cjs`)
        const extension = row.format.extensions[0]
        const sample = sampleFor(extension)
        const source = path.join(folder, sample ? path.basename(sample) : `Sample${extension}`)
        if (sample) fs.copyFileSync(sample, source)
        else fs.writeFileSync(source, isLegacyBinary(row.format) ? Buffer.concat([CFB_SIGNATURE, Buffer.alloc(1024)]) : Buffer.from(`Simple engine-absent sample (${row.format.id})\n`))
        const originalHash = sha256(fs.readFileSync(source))

        const script = path.join(folder, 'dialogs.json')
        fs.writeFileSync(script, JSON.stringify({ save: [], open: [], prompts: { 'prompts.sibling-save': ['save-sibling'] }, log: dialogLog }))
        process.env.SIMPLE_QA_DIALOGS = script
        dialogs.configureDialogs({ module: row.workspace, userData: path.join(run.root, 'userData', row.workspace), documentsDir: documents, desktopDir: documents, logger: quiet })
        guard.configureDocumentGuard({ module: row.workspace, logger: quiet })

        const opened = await guard.openDocument(source, { webContents: win.webContents, format: row.format.id })
        assert.equal(opened.ok, true, `open: ${JSON.stringify(opened)}`)
        const serialized = []
        const hooks = {
          module: row.workspace,
          engine: false,
          serialize: (format) => {
            serialized.push(format)
            const write = FIXTURE_WRITERS[format]
            if (!write) throw new Error(`asked to write ${format}`)
            return write(`edited ${row.format.id}`)
          },
        }
        const first = await guard.performSave({ sender: win.webContents }, { docId: opened.docId, mode: 'save', revision: 1 }, hooks)
        assert.equal(first.ok, true, `save: ${JSON.stringify(first)}`)
        const expected = path.join(folder, `${path.parse(source).name}${row.sibling.extension}`)
        assert.equal(path.resolve(first.path).toLowerCase(), expected.toLowerCase(), 'the edits go to the sibling next to the original')
        assert.equal(Boolean(first.sibling), true, 'the result says it is a sibling save')
        assert.deepEqual([...new Set(serialized)], [row.sibling.formatId], 'only the sibling format is serialized')
        const written = fs.readFileSync(first.path)
        assert.ok(!written.subarray(0, 8).equals(CFB_SIGNATURE), 'the sibling is not a legacy binary')
        const valid = await validateFile(row.sibling.formatId, first.path)
        assert.equal(valid.ok, true, `the sibling validates: ${valid.reason || ''}`)
        assert.equal(sha256(fs.readFileSync(source)), originalHash, 'the original is unchanged')

        const second = await guard.performSave({ sender: win.webContents }, { docId: opened.docId, mode: 'save', revision: 2 }, hooks)
        assert.equal(second.ok, true, `second save: ${JSON.stringify(second)}`)
        assert.equal(path.resolve(second.path).toLowerCase(), expected.toLowerCase(), 'later saves go straight to the sibling')
        assert.equal(sha256(fs.readFileSync(source)), originalHash, 'the original is still unchanged')

        const log = fs.existsSync(dialogLog) ? fs.readFileSync(dialogLog, 'utf8').split(/\r?\n/).filter(Boolean).map((line) => JSON.parse(line)) : []
        const prompts = log.filter((entry) => entry.type === 'prompt')
        assert.deepEqual(prompts.map((entry) => entry.key), ['prompts.sibling-save'], 'exactly one confirmation, on the first save')
        for (const entry of log) assert.ok(!SETUP_WORDING.test(String(entry.message || '')), `a prompt asks to download or install something: ${entry.message}`)
        assert.deepEqual(log.filter((entry) => entry.type !== 'prompt'), [], 'no Save or Open dialog')
        assert.deepEqual(electron.unexpected.slice(before), [], 'nothing unscripted was shown or handed to another program')
        outcomes.push({ name, ok: true, detail: `${path.basename(source)} unchanged; edits in ${path.basename(first.path)}` })
      } catch (error) {
        outcomes.push({ name, ok: false, detail: String((error && error.message) || error) })
      } finally {
        win.destroy()
      }
    }
  } finally {
    if (previousDialogs === undefined) delete process.env.SIMPLE_QA_DIALOGS
    else process.env.SIMPLE_QA_DIALOGS = previousDialogs
    electron.restore()
  }
  return outcomes
}

// ---------------------------------------------------------------------------
// Enabled workspaces (real app, through the acceptance harness)
// ---------------------------------------------------------------------------

/**
 * Runs an enabled workspace's adapter `matrix` rows in the real app with the engine forced off.
 * Row: {sample, expect: 'sibling'|'in-place'|'notice', marker?}. See the integration guide.
 * @param {string} mode
 * @param {{keep?: boolean}} options
 */
async function workspaceRows(mode, options) {
  const harness = require('./io-harness.cjs')
  const { validateFile } = require(path.join(SHARED, 'validators.cjs'))
  const { path: adapterPath, adapter } = harness.loadAdapter(mode)
  const rows = Array.isArray(adapter.matrix) ? adapter.matrix : []
  if (!rows.length) return [{ name: `${mode}: engine-absent rows`, ok: false, detail: `the ${mode} adapter lists no matrix rows (scripts/io-adapters/${mode}.cjs)` }]
  const outcomes = []
  for (const row of rows) {
    const name = `${mode}: ${row.sample} (${row.expect})`
    const run = harness.makeRunFolder(`no-office-${mode}`)
    try {
      const file = harness.copySample(run, row.sample)
      const before = harness.hashFile(file)
      const env = { SIMPLE_FORCE_NO_OFFICE: '1' }
      const dialogs = { prompts: { 'prompts.sibling-save': ['save-sibling'] } }
      if (row.expect === 'notice') {
        const result = await harness.launchSimple({ run, mode, adapterPath, files: [file], env, dialogs, timeoutMs: 240_000, step: { scenario: 'openNotice' } })
        assert.equal(result.result && result.result.ok, true, `${result.result && result.result.error} ${result.output.slice(-1000)}`)
        const notice = String(result.result.value.notice || '')
        assert.ok(notice, 'the workspace tells the user what it did instead')
        assert.ok(!SETUP_WORDING.test(notice), `the notice asks to download or install something: ${notice}`)
        assert.equal(harness.hashFile(file), before, 'the original is unchanged')
        outcomes.push({ name, ok: true, detail: notice })
        continue
      }
      const format = formats.formatForExtension(path.extname(file))
      const save = format && formats.savePolicy(format.id, mode, { engine: false })
      const target = row.expect === 'sibling' && save && save.target === 'sibling'
        ? path.join(path.dirname(file), `${path.parse(file).name}${save.extension}`)
        : file
      const result = await harness.launchSimple({
        run, mode, adapterPath, files: [file], env, dialogs, timeoutMs: 240_000,
        step: { scenario: 'editSave', marker: row.marker || 'QA-NO-OFFICE', watch: target, watchBefore: target === file ? before : null },
      })
      assert.equal(result.timedOut, false, `timed out. ${result.output.slice(-1000)}`)
      assert.equal(result.result && result.result.ok, true, `${result.result && result.result.error}`)
      for (const entry of result.dialogLog) assert.ok(!SETUP_WORDING.test(String(entry.message || '')), `a prompt asks to download or install something: ${entry.message}`)
      if (row.expect === 'sibling') {
        assert.notEqual(target, file, `${path.basename(file)} has no sibling policy without the engine`)
        assert.equal(harness.hashFile(file), before, 'the original is unchanged')
      }
      const written = fs.readFileSync(target)
      assert.ok(!written.subarray(0, 8).equals(CFB_SIGNATURE) || row.valuesOnly === true, 'no legacy binary was written')
      const targetFormat = formats.formatForExtension(path.extname(target))
      const valid = await validateFile(targetFormat ? targetFormat.id : path.extname(target), target)
      assert.equal(valid.ok, true, `the saved file validates: ${valid.reason || ''}`)
      outcomes.push({ name, ok: true, detail: `saved ${path.basename(target)}` })
    } catch (error) {
      outcomes.push({ name, ok: false, detail: String((error && error.message) || error) })
    } finally {
      if (!options.keep) await harness.removeRunFolder(run)
    }
  }
  return outcomes
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

/**
 * Runs the matrix.
 * @param {{sharedOnly?: boolean, modules?: string[], keep?: boolean, tempRoot?: string}} [options]
 * @returns {Promise<{ok: boolean, outcomes: Array<{name: string, ok: boolean, detail: string}>, enabled: string[]}>}
 */
async function runMatrix(options = {}) {
  const outcomes = []
  for (const workspace of WORKSPACES) {
    const problems = registryProblems(workspace)
    outcomes.push({ name: `registry ${workspace}: nothing offered needs the office engine`, ok: !problems.length, detail: problems.join('\n') })
  }
  const engine = await engineProblems()
  outcomes.push({ name: `engine: probe says forced-off; ${engine.pairs} conversion requests say what Simple does instead`, ok: !engine.problems.length, detail: engine.problems.join('\n') })

  const parent = options.tempRoot || os.tmpdir()
  fs.mkdirSync(parent, { recursive: true })
  const run = { root: fs.mkdtempSync(path.join(parent, 'simple-no-office-')) }
  try {
    outcomes.push(...await saveRows(run))
  } finally {
    if (!options.keep) fs.rmSync(run.root, { recursive: true, force: true })
  }

  let enabled = []
  if (!options.sharedOnly) {
    const harness = require('./io-harness.cjs')
    enabled = harness.enabledWorkspaces()
    const selected = options.modules && options.modules.length ? enabled.filter((name) => options.modules.includes(name)) : enabled
    for (const mode of selected) {
      if (process.platform !== 'win32') {
        outcomes.push({ name: `${mode}: engine-absent rows`, ok: true, detail: 'skipped (Windows only)' })
        continue
      }
      try {
        outcomes.push(...await workspaceRows(mode, options))
      } catch (error) {
        outcomes.push({ name: `${mode}: engine-absent rows`, ok: false, detail: String((error && error.message) || error) })
      }
    }
  }
  return { ok: outcomes.every((outcome) => outcome.ok), outcomes, enabled }
}

async function main(argv) {
  const options = { modules: [] }
  let json = false
  for (const argument of argv) {
    if (argument === '--shared-only') options.sharedOnly = true
    else if (argument === '--keep') options.keep = true
    else if (argument === '--json') json = true
    else if (argument.startsWith('--module=')) options.modules.push(...argument.slice(9).split(',').filter(Boolean))
    else {
      console.error(`Unknown option: ${argument}`)
      return 2
    }
  }
  const result = await runMatrix(options)
  if (json) {
    process.stdout.write(`${JSON.stringify(result)}\n`)
  } else {
    for (const outcome of result.outcomes) console.log(`${outcome.ok ? 'ok  ' : 'FAIL'} ${outcome.name}${outcome.detail && !outcome.ok ? `:\n     ${outcome.detail.split('\n').join('\n     ')}` : ''}`)
    if (options.sharedOnly) console.log('Shared rows only: no app was started.')
    else console.log(result.enabled.length ? `Workspaces checked in the app: ${result.enabled.join(', ')}.` : 'No workspace is enabled in simple/shared/manifest.json yet; only the shared rows ran.')
  }
  return result.ok ? 0 : 1
}

if (require.main === module) main(process.argv.slice(2)).then((code) => { process.exitCode = code })

module.exports = { FIXTURE_WRITERS, engineProblems, registryProblems, runMatrix, siblingRows }
