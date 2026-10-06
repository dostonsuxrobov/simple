// ExportModel (simple/shared/renderer/export-model.ts), the state of the Export
// As dialog (design §7), against a fake window.simpleIO and the real format
// registry: one-line descriptions, unavailable rows, remembered format, folder
// and options, strict extensions, failure prompts, warnings, cancel, and
// "open after exporting" going to Simple (never to another program).

import test from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { createRequire, registerHooks } from 'node:module'
import { fileURLToPath } from 'node:url'

// The shared renderer files are TypeScript ES modules written for Vite:
// extensionless relative imports and a plain JSON import, inside a package
// whose package.json says "type": "commonjs". These hooks let Node's built-in
// type stripping load them unchanged.
const RENDERER = new URL('../shared/renderer/', import.meta.url)
registerHooks({
  resolve(specifier, context, nextResolve) {
    if (/^\.\.?\//.test(specifier) && !/\.[a-z0-9]+$/i.test(specifier) && context.parentURL?.startsWith(RENDERER.href)) {
      return nextResolve(`${specifier}.ts`, context)
    }
    return nextResolve(specifier, context)
  },
  load(url, context, nextLoad) {
    if (url.startsWith(RENDERER.href) && url.endsWith('.json')) {
      return { format: 'module', source: `export default ${readFileSync(fileURLToPath(url), 'utf8')};`, shortCircuit: true }
    }
    if (url.startsWith(RENDERER.href) && url.endsWith('.ts')) {
      return { format: 'module-typescript', source: readFileSync(fileURLToPath(url), 'utf8'), shortCircuit: true }
    }
    return nextLoad(url, context)
  },
})

const {
  COMMON_EXTENSIONS,
  DEFAULT_EXPORT_FORMATS,
  EXPORT_PREF_KEYS,
  ExportModel,
  coerceOptionValue,
  defaultExportFormat,
  exportFileName,
  extensionMatches,
  formatTag,
  lossNote,
  multiFileFolderName,
  resolveTypedName,
  safeStem,
} = await import(new URL('export-model.ts', RENDERER).href)
const { DocumentSession } = await import(new URL('document-session.ts', RENDERER).href)
const { catalogText } = await import(new URL('io-client.ts', RENDERER).href)
const catalog = JSON.parse(readFileSync(new URL('io-catalog.json', RENDERER), 'utf8'))
const registry = createRequire(import.meta.url)('../shared/electron/formats.cjs')

const flush = () => new Promise((resolve) => setImmediate(resolve))

function deferred() {
  let resolve
  const promise = new Promise((done) => { resolve = done })
  return { promise, resolve }
}

function createIO(options = {}) {
  const prefs = new Map(Object.entries(options.prefs || {}))
  const calls = { chosen: [], prompts: [], shown: [], opened: [], openPath: [], prefsSet: [] }
  const io = {
    version: 1,
    module: options.module ?? 'calc',
    calls,
    prefs: {
      async get(key) {
        return prefs.has(key) ? structuredClone(prefs.get(key)) : undefined
      },
      async set(key, value) {
        calls.prefsSet.push([key, structuredClone(value)])
        prefs.set(key, structuredClone(value))
        return true
      },
    },
    stored: prefs,
    answers: [...(options.answers || [])],
    targets: [...(options.targets || [])],
    async chooseSavePath(request) {
      calls.chosen.push(structuredClone(request))
      const next = io.targets.length ? io.targets.shift() : null
      return typeof next === 'function' ? next(request) : next
    },
    async prompt(key, vars) {
      calls.prompts.push({ key, vars })
      return io.answers.length ? io.answers.shift() : 'cancel'
    },
    shell: {
      async showItem(path) {
        calls.shown.push(path)
        return { ok: true }
      },
      // Present only to prove nothing calls it: exports open inside Simple.
      async openPath(path) {
        calls.openPath.push(path)
      },
    },
    async openInSimple(path) {
      calls.opened.push(path)
      return { ok: true, action: 'launched', mode: 'pdf', appName: 'Simple PDF', path }
    },
    async capabilities() {
      return { officeEngine: { available: false }, platform: 'win32' }
    },
    onCapabilitiesChanged: () => () => {},
    pathForFile: () => null,
    async chooseOpenPaths() {
      return []
    },
    recovery: { async write() { return { ok: true } }, async list() { return [] }, async read() { return { ok: false, code: 'NOT_FOUND' } }, async discard() { return { ok: true } } },
    versions: { async list() { return [] }, async open() {} },
    clipboard: { async read() { return { files: [] } } },
    officeEngine: { async status() { return { available: false } } },
    onRequest: () => () => {},
  }
  return io
}

function exported(request, extra = {}) {
  const name = request.path.split('\\').pop()
  return { ok: true, path: request.path, name, format: request.format, warnings: [], ...extra }
}

function setup(options = {}) {
  const io = options.io || createIO(options)
  const toasts = []
  const exports = []
  const results = [...(options.results || [])]
  const model = new ExportModel({
    io,
    module: options.module,
    documentName: options.documentName ?? 'Budget.xlsx',
    documentPath: options.documentPath === undefined ? 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx' : options.documentPath,
    formats: options.formats || registry.exportFormats(options.module || 'calc', { engine: false }),
    options: options.optionDefs,
    order: options.order,
    session: options.session,
    exporter: async (request) => {
      exports.push({ ...request, options: { ...request.options } })
      const next = results.length ? results.shift() : null
      if (typeof next === 'function') return next(request)
      return next || exported(request)
    },
    notify: (kind, text, actions) => toasts.push({ kind, text, actions }),
    onShowWarnings: options.onShowWarnings,
    log: () => {},
  })
  return { io, model, toasts, exports }
}

const CSV_OPTIONS = [
  { id: 'scope', label: 'Include', type: 'choice', formats: ['csv', 'tsv', 'pdf'], defaultValue: 'sheet', choices: [{ id: 'sheet', label: 'Active sheet' }, { id: 'selection', label: 'Selection' }, { id: 'workbook', label: 'Whole workbook' }] },
  { id: 'delimiter', label: 'Delimiter', type: 'choice', formats: ['csv'], defaultValue: ',', choices: [{ id: ',', label: 'Comma' }, { id: ';', label: 'Semicolon' }] },
  { id: 'header', label: 'Header row', type: 'toggle', formats: ['csv', 'tsv'], defaultValue: true },
  { id: 'dpi', label: 'Resolution', type: 'number', formats: ['pdf'], defaultValue: 150, min: 72, max: 600, step: 1 },
  { id: 'range', label: 'Pages', type: 'text', formats: ['pdf'], defaultValue: '', visible: (values) => values.scope === 'selection', remember: false },
]

// ---------------------------------------------------------------------------
// Rows, defaults and options
// ---------------------------------------------------------------------------

test('rows come from the registry with one-line descriptions and honest tags', () => {
  const { model } = setup({ order: ['xlsx', 'pdf', 'csv'] })
  const rows = model.state.rows
  assert.deepEqual(rows.slice(0, 3).map((row) => row.id), ['xlsx', 'pdf', 'csv'], 'listed in the requested order first')
  for (const row of rows) {
    assert.ok(row.description, `${row.id} has a description`)
    assert.ok(!row.description.includes('\n'), 'one line')
    assert.ok(row.extension.startsWith('.'))
  }
  const byId = Object.fromEntries(rows.map((row) => [row.id, row]))
  assert.equal(byId.xls.tag, catalog.export.valuesOnly)
  assert.equal(byId.ods.tag, catalog.export.basicFormatting)
  assert.equal(byId.xlsx.tag, '')
  assert.equal(byId.csv.lossy, true)
  assert.equal(model.state.title, 'Export "Budget.xlsx"')
  assert.equal(model.state.openAfterLabel, catalog.export.openAfter)
})

test('a row this PC cannot produce stays visible, says why, and cannot be selected', () => {
  const formats = [
    { id: 'pdf', label: 'PDF document', extensions: ['.pdf'], description: 'Looks the same everywhere.', available: true },
    { id: 'ods', label: 'OpenDocument spreadsheet', extensions: ['.ods'], description: 'For LibreOffice.', available: false, reason: 'needs-office-engine' },
  ]
  const { model } = setup({ formats })
  const ods = model.state.rows.find((row) => row.id === 'ods')
  assert.equal(ods.selectable, false)
  assert.equal(ods.tag, catalog.export.notAvailable)
  assert.doesNotMatch(ods.tag, /install|download/i)
  assert.equal(model.select('ods'), false)
  assert.equal(model.state.selected.id, 'pdf')
  assert.equal(model.select('nope'), false)
})

test('first run preselects PDF (PNG in Images); a remembered format wins when it is still offered', async () => {
  assert.equal(setup({ module: 'calc' }).model.state.selected.id, 'pdf')
  assert.equal(setup({ module: 'docs' }).model.state.selected.id, 'pdf')
  assert.equal(setup({ module: 'pdf', formats: registry.exportFormats('pdf', { engine: false }) }).model.state.selected.id, 'pdf')
  assert.equal(setup({ module: 'image', formats: registry.exportFormats('image', { engine: false }), documentName: 'photo.jpg' }).model.state.selected.id, 'png')
  assert.equal(defaultExportFormat('video'), null)
  assert.deepEqual(Object.keys(DEFAULT_EXPORT_FORMATS).sort(), ['calc', 'docs', 'image', 'pdf', 'video'])

  const remembered = setup({ prefs: { exportFormat: 'csv', openAfterExport: true, lastExportFolder: 'D:\\Reports\\2026' } })
  await remembered.model.load()
  assert.equal(remembered.model.state.selected.id, 'csv')
  assert.equal(remembered.model.state.openAfter, true)
  assert.equal(remembered.model.state.folder, 'D:\\Reports\\2026')
  assert.equal(remembered.model.state.savesTo, 'Saves to: Reports › 2026')

  const gone = setup({ prefs: { exportFormat: 'numbers' } })
  await gone.model.load()
  assert.equal(gone.model.state.selected.id, 'pdf', 'a remembered format that is not offered falls back to the default')

  const created = await ExportModel.create({ io: createIO({ prefs: { exportFormat: 'tsv' } }), module: 'calc', documentName: 'B.xlsx', formats: registry.exportFormats('calc', { engine: false }), exporter: async () => ({ ok: false, code: 'UNKNOWN' }) })
  assert.equal(created.state.selected.id, 'tsv')
})

test('the first-run folder is the document folder; untitled documents have none yet', () => {
  assert.equal(setup().model.state.savesTo, 'Saves to: Documents › Finance')
  const untitled = setup({ documentPath: null, documentName: '' })
  assert.equal(untitled.model.state.folder, null)
  assert.equal(untitled.model.state.savesTo, '')
  assert.equal(untitled.model.fileName, 'Untitled spreadsheet.pdf')
})

test('options are contextual, validated and remembered per format', async () => {
  const { model, io } = setup({ optionDefs: CSV_OPTIONS, prefs: { exportOptions: { csv: { delimiter: ';', header: 'yes' }, pdf: { dpi: 9000 } } } })
  await model.load()
  assert.equal(model.state.selected.id, 'pdf')
  assert.deepEqual(model.state.options.map((option) => option.id), ['scope', 'dpi'], 'only the PDF options, and no page range unless Selection')
  assert.equal(model.state.options.find((option) => option.id === 'dpi').value, 600, 'a remembered value is clamped to the range')

  assert.equal(model.setOption('scope', 'selection'), true)
  assert.deepEqual(model.state.options.map((option) => option.id), ['scope', 'dpi', 'range'])
  assert.equal(model.setOption('scope', 'everything'), false, 'not one of the choices')
  assert.equal(model.setOption('delimiter', ';'), false, 'not an option of PDF')
  assert.equal(model.setOption('dpi', 299.6), true)
  assert.equal(model.state.options.find((option) => option.id === 'dpi').value, 300)
  assert.equal(model.setOption('dpi', Number.NaN), false)

  assert.equal(model.select('csv'), true)
  const csv = Object.fromEntries(model.state.options.map((option) => [option.id, option.value]))
  assert.deepEqual(csv, { scope: 'sheet', delimiter: ';', header: true }, 'remembered delimiter kept; invalid remembered header ignored')
  assert.equal(model.state.lossNote, 'Keeps: values from one sheet. Not kept: formulas, formatting, other sheets, charts and images.')

  model.setOption('header', false)
  io.targets.push({ path: 'C:\\Out\\Budget.csv', format: 'csv' })
  const outcome = await model.run()
  assert.equal(outcome.ok, true)
  const saved = io.stored.get(EXPORT_PREF_KEYS.options)
  assert.deepEqual(saved.csv, { scope: 'sheet', delimiter: ';', header: false })
  assert.equal(saved.pdf.dpi, 9000, 'other formats keep what was remembered before')
})

test('coerceOptionValue checks types, choices and ranges', () => {
  // Steps count from the minimum, as in <input type="number" min="1" step="5">: 1, 6, 11, … 96.
  const number = { id: 'q', label: 'Quality', type: 'number', defaultValue: 81, min: 1, max: 100, step: 5 }
  assert.equal(coerceOptionValue(number, 83), 81)
  assert.equal(coerceOptionValue(number, 84), 86)
  assert.equal(coerceOptionValue(number, 1000), 96, 'never above the maximum after snapping')
  assert.equal(coerceOptionValue(number, -4), 1)
  assert.equal(coerceOptionValue(number, '80'), undefined)
  assert.equal(coerceOptionValue({ id: 's', label: 'S', type: 'number', defaultValue: 1, min: 0, step: 0.1 }, 0.3000001), 0.3)
  assert.equal(coerceOptionValue({ id: 't', label: 'T', type: 'toggle', defaultValue: true }, 1), undefined)
  assert.equal(coerceOptionValue({ id: 'x', label: 'X', type: 'text', defaultValue: '' }, 'a'), 'a')
})

// ---------------------------------------------------------------------------
// File names and extensions
// ---------------------------------------------------------------------------

test('suggested names strip only known extensions and never double them', () => {
  const pdf = { extensions: ['.pdf'] }
  const png = { extensions: ['.png'] }
  assert.equal(exportFileName('Budget.xlsx', pdf), 'Budget.pdf')
  assert.equal(exportFileName('Q3 plan v2.1', pdf), 'Q3 plan v2.1.pdf')
  assert.equal(exportFileName('photo.jpg', png), 'photo.png')
  assert.equal(exportFileName('C:\\Users\\me\\Report.final.docx', pdf), 'Report.final.pdf')
  assert.equal(exportFileName('Notes: draft?', pdf), 'Notes- draft.pdf')
  assert.equal(exportFileName('.xlsx', pdf, 'Untitled'), 'Untitled.pdf')
  assert.equal(safeStem('archive.tar.gz'), 'archive.tar.gz', '.gz is not a document extension')
  assert.ok(COMMON_EXTENSIONS.includes('.xlsx'))
  assert.equal(multiFileFolderName('Budget.xlsx', { label: 'PNG image' }), 'Budget - PNG image pages')
  assert.equal(setup({ documentName: 'Q3 plan v2.1' }).model.fileName, 'Q3 plan v2.1.pdf')
})

test('typed names follow the strict extension rules', () => {
  const rows = registry.exportFormats('image', { engine: false })
  const png = rows.find((row) => row.id === 'png')
  assert.deepEqual(resolveTypedName('photo.png', png, rows), { name: 'photo.png', format: 'png', action: 'keep' })
  assert.deepEqual(resolveTypedName('photo.JPG', png, rows), { name: 'photo.JPG', format: 'jpeg', action: 'switch' })
  assert.deepEqual(resolveTypedName('photo.bmp', png, rows), { name: 'photo.png', format: 'png', action: 'replace' })
  assert.deepEqual(resolveTypedName('Q3 plan v2.1', png, rows), { name: 'Q3 plan v2.1.png', format: 'png', action: 'append' })
  assert.deepEqual(resolveTypedName('  scan  ', png, rows), { name: 'scan.png', format: 'png', action: 'append' })
  assert.equal(extensionMatches('C:\\Out\\photo.JPEG', rows.find((row) => row.id === 'jpeg')), true)
  assert.equal(extensionMatches('C:\\Out\\photo.jpg.png', rows.find((row) => row.id === 'jpeg')), false)
  assert.equal(formatTag({ available: true, mode: 'native' }), '')
  assert.equal(lossNote({ lossy: null }), '')
})

// ---------------------------------------------------------------------------
// Running an export
// ---------------------------------------------------------------------------

test('Export commits pending edits, asks main for the path, exports, remembers and reports', async () => {
  const prepared = []
  const tracked = []
  const session = {
    docId: 'doc-7',
    async prepare(action) {
      prepared.push(action)
      return true
    },
    async track(action, work) {
      tracked.push(action)
      return work()
    },
  }
  const { model, io, toasts, exports } = setup({ session, optionDefs: CSV_OPTIONS, targets: [{ path: 'C:\\Users\\me\\Documents\\Finance\\Budget.pdf', format: 'pdf' }] })
  const outcome = await model.run()
  assert.deepEqual(outcome, { ok: true, path: 'C:\\Users\\me\\Documents\\Finance\\Budget.pdf', name: 'Budget.pdf', format: 'pdf', warnings: [], opened: false })
  assert.deepEqual(prepared, ['export'])
  assert.deepEqual(tracked, ['export'])
  const request = io.calls.chosen[0]
  assert.equal(request.purpose, 'export')
  assert.equal(request.name, 'Budget.pdf')
  assert.equal(request.format, 'pdf')
  assert.equal(request.docId, 'doc-7')
  assert.equal(request.formats[0], 'pdf')
  assert.deepEqual([...request.formats].sort(), registry.exportFormats('calc', { engine: false }).filter((row) => row.available).map((row) => row.id).sort())
  assert.equal(exports.length, 1)
  assert.equal(exports[0].path, 'C:\\Users\\me\\Documents\\Finance\\Budget.pdf')
  assert.equal(exports[0].format, 'pdf')
  assert.equal(exports[0].docId, 'doc-7')
  assert.deepEqual(exports[0].options, { scope: 'sheet', dpi: 150 })
  assert.equal(typeof exports[0].onProgress, 'function')
  assert.equal(exports[0].signal.aborted, false)

  assert.equal(toasts.length, 1)
  assert.equal(toasts[0].kind, 'success')
  assert.equal(toasts[0].text, 'Exported "Budget.pdf" to Documents › Finance.')
  assert.deepEqual(toasts[0].actions.map((action) => action.label), [catalog.toast.actions.open, catalog.toast.actions.showInFolder])

  const remembered = Object.fromEntries(io.calls.prefsSet)
  assert.equal(remembered[EXPORT_PREF_KEYS.format], 'pdf')
  assert.equal(remembered[EXPORT_PREF_KEYS.folder], 'C:\\Users\\me\\Documents\\Finance')
  assert.equal(remembered[EXPORT_PREF_KEYS.openAfter], false)
  assert.equal(model.state.busy, false)
  assert.equal(io.calls.opened.length, 0, 'nothing opens unless asked')

  // [Open] opens the file in Simple; [Show in Folder] shows it in File Explorer.
  toasts[0].actions[0].run()
  await flush()
  assert.deepEqual(io.calls.opened, ['C:\\Users\\me\\Documents\\Finance\\Budget.pdf'])
  assert.equal(toasts[1].text, catalogText('toast.openedIn', { appName: 'Simple PDF' }))
  toasts[0].actions[1].run()
  await flush()
  assert.deepEqual(io.calls.shown, ['C:\\Users\\me\\Documents\\Finance\\Budget.pdf'])
  assert.equal(io.calls.openPath.length, 0)
})

test('"Open the file after exporting" opens it in Simple, never in another program', async () => {
  const { model, io } = setup({ targets: [{ path: 'C:\\Out\\Budget.pdf', format: 'pdf' }] })
  model.setOpenAfter(true)
  const outcome = await model.run()
  assert.equal(outcome.ok, true)
  assert.equal(outcome.opened, true)
  assert.deepEqual(io.calls.opened, ['C:\\Out\\Budget.pdf'])
  assert.equal(io.calls.openPath.length, 0)
  assert.equal(io.stored.get(EXPORT_PREF_KEYS.openAfter), true)

  // A format no Simple workspace opens is only shown in File Explorer by main, and the toast says why.
  const json = setup({ targets: [{ path: 'C:\\Out\\Budget.pdf', format: 'pdf' }] })
  json.io.openInSimple = async (path) => ({ ok: true, action: 'shown-in-folder', shownInFolder: true, mode: null, appName: null, path, reason: 'unsupported' })
  json.model.setOpenAfter(true)
  assert.equal((await json.model.run()).opened, false)
  assert.equal(json.toasts[json.toasts.length - 1].text, catalogText('toast.shownInFolder', { name: 'Budget.pdf' }))

  // An older bridge without openInSimple: File Explorer, never the default app, and no claim about the format.
  const old = setup({ targets: [{ path: 'C:\\Out\\Budget.pdf', format: 'pdf' }] })
  delete old.io.openInSimple
  old.model.setOpenAfter(true)
  assert.equal((await old.model.run()).opened, false)
  assert.deepEqual(old.io.calls.shown, ['C:\\Out\\Budget.pdf'])
  assert.equal(old.io.calls.openPath.length, 0)
  assert.equal(old.toasts.length, 1, 'only the "Exported" toast')
})

test('typing another offered extension in the dialog switches the format', async () => {
  const { model, exports, io } = setup({ optionDefs: CSV_OPTIONS, targets: [{ path: 'C:\\Out\\Budget.csv', format: 'csv' }] })
  assert.equal(model.state.selected.id, 'pdf')
  const outcome = await model.run()
  assert.equal(outcome.format, 'csv')
  assert.equal(exports[0].format, 'csv')
  assert.deepEqual(exports[0].options, { scope: 'sheet', delimiter: ',', header: true })
  assert.equal(model.state.selected.id, 'csv')
  assert.equal(io.stored.get(EXPORT_PREF_KEYS.format), 'csv')
})

test('a path whose extension does not match its format is never written', async () => {
  const { model, exports, io } = setup({
    answers: ['save-as'],
    targets: [{ path: 'C:\\Out\\photo.jpg.png', format: 'pdf' }, { path: 'C:\\Users\\me\\Documents\\Budget.pdf', format: 'pdf' }],
  })
  const outcome = await model.run()
  assert.equal(outcome.ok, true)
  assert.equal(io.calls.prompts[0].key, 'saveFailed.INVALID_NAME')
  assert.equal(io.calls.prompts[0].vars.name, 'photo.jpg.png')
  assert.equal(io.calls.chosen[1].purpose, 'export', 'a name problem reopens the export dialog, not the fallback folder')
  assert.equal(io.calls.chosen[1].reason, undefined)
  assert.deepEqual(exports.map((request) => request.path), ['C:\\Users\\me\\Documents\\Budget.pdf'])
})

test('export failures get the save prompts: Save As in the fallback folder, Replace, Try Again', async () => {
  const { model, exports, io, toasts } = setup({
    answers: ['save-as', 'replace', 'retry'],
    targets: [{ path: 'C:\\Shared\\Budget.pdf', format: 'pdf' }, { path: 'C:\\Users\\me\\Documents\\Budget.pdf', format: 'pdf' }],
    results: [
      (request) => ({ ok: false, code: 'LOCKED', path: request.path, name: 'Budget.pdf', technical: 'EBUSY rename after 7 attempts' }),
      (request) => ({ ok: false, code: 'CHANGED_ON_DISK', path: request.path, name: 'Budget.pdf' }),
      (request) => ({ ok: false, code: 'VERIFY_FAILED', path: request.path, name: 'Budget.pdf' }),
    ],
  })
  const outcome = await model.run()
  assert.equal(outcome.ok, true)
  assert.deepEqual(io.calls.prompts.map((prompt) => prompt.key), ['saveFailed.LOCKED', 'saveFailed.CHANGED_ON_DISK', 'saveFailed.VERIFY_FAILED'])
  assert.equal(io.calls.prompts[0].vars.technical, 'EBUSY rename after 7 attempts')
  assert.deepEqual(io.calls.chosen[1], { purpose: 'fallback', name: 'Budget.pdf', format: 'pdf', formats: io.calls.chosen[0].formats, reason: 'LOCKED' })
  assert.deepEqual(exports.map((request) => [request.path, request.force === true]), [
    ['C:\\Shared\\Budget.pdf', false],
    ['C:\\Users\\me\\Documents\\Budget.pdf', false],
    ['C:\\Users\\me\\Documents\\Budget.pdf', true],
    ['C:\\Users\\me\\Documents\\Budget.pdf', true],
  ])
  assert.equal(toasts.filter((toast) => toast.kind === 'success').length, 1, 'failures are prompts, never toasts')

  // Stopped partway: Show Original shows the kept original, then the same prompt returns.
  const restore = setup({
    answers: ['show-original', 'cancel'],
    targets: [{ path: 'C:\\Shared\\Budget.pdf', format: 'pdf' }],
    results: [{ ok: false, code: 'RESTORE_NEEDED', path: 'C:\\Shared\\Budget.pdf', backupPath: 'C:\\Shared\\~simple-0a1b2c3d.old' }],
  })
  assert.equal((await restore.model.run()).code, 'RESTORE_NEEDED')
  assert.deepEqual(restore.io.calls.shown, ['C:\\Shared\\~simple-0a1b2c3d.old'])
  assert.equal(restore.io.calls.prompts.length, 2)
  const stubborn = setup({ targets: [{ path: 'C:\\Shared\\Budget.pdf', format: 'pdf' }], results: [{ ok: false, code: 'RESTORE_NEEDED', path: 'C:\\Shared\\Budget.pdf' }] })
  stubborn.io.prompt = async () => 'show-original'
  assert.equal((await stubborn.model.run()).ok, false, 'a prompt that never decides cannot loop forever')

  const giveUp = setup({ answers: ['cancel'], targets: [{ path: 'C:\\Shared\\Budget.pdf', format: 'pdf' }], results: [{ ok: false, code: 'READ_ONLY', path: 'C:\\Shared\\Budget.pdf' }] })
  const failed = await giveUp.model.run()
  assert.equal(failed.ok, false)
  assert.equal(failed.code, 'READ_ONLY')
  assert.equal(giveUp.toasts.length, 0)
  assert.equal(giveUp.io.calls.prefsSet.length, 0, 'nothing is remembered from a failed export')
})

test('warnings are reported with a Details button', async () => {
  const shown = []
  const { model, toasts } = setup({
    onShowWarnings: (warnings, name) => shown.push({ warnings, name }),
    targets: [{ path: 'C:\\Out\\Budget.pdf', format: 'pdf' }],
    results: [(request) => exported(request, { warnings: ['Chart "Sales" (not supported in PDF)', 'Comment on B4'] })],
  })
  const outcome = await model.run()
  assert.deepEqual(outcome.warnings, ['Chart "Sales" (not supported in PDF)', 'Comment on B4'])
  assert.equal(toasts[0].kind, 'warning')
  assert.equal(toasts[0].text, catalogText('toast.exportedWithWarnings', { name: 'Budget.pdf', count: 2 }))
  assert.equal(toasts[0].actions[0].label, catalog.toast.actions.details)
  toasts[0].actions[0].run()
  assert.deepEqual(shown, [{ warnings: ['Chart "Sales" (not supported in PDF)', 'Comment on B4'], name: 'Budget.pdf' }])
  assert.deepEqual([...model.state.warnings], outcome.warnings)
})

test('progress and cancel: a canceled export leaves the model ready and says no file was created', async () => {
  const started = deferred()
  const { model, toasts, io } = setup({
    targets: [{ path: 'C:\\Out\\Budget.pdf', format: 'pdf' }],
    results: [(request) => new Promise((resolve) => {
      request.onProgress(40)
      started.resolve()
      request.signal.addEventListener('abort', () => resolve({ ok: false, code: 'CANCELED' }))
    })],
  })
  const running = model.run()
  await started.promise
  assert.equal(model.state.busy, true)
  assert.equal(model.state.progress, 40)
  assert.equal(model.state.statusText, 'Exporting… 40%')
  assert.equal(model.state.canExport, false)
  assert.equal((await model.run()).code, 'BLOCKED', 'one export at a time')
  assert.equal(model.select('csv'), false, 'the format is fixed while exporting')
  model.cancel()
  const outcome = await running
  assert.equal(outcome.code, 'CANCELED')
  assert.equal(toasts[0].text, catalog.toast.exportCanceled)
  assert.equal(model.state.busy, false)
  assert.equal(model.state.progress, null)
  assert.equal(io.calls.prefsSet.length, 0)

  // A canceled Save dialog is just a cancel: no toast, no export.
  const dialog = setup({ targets: [null] })
  const result = await dialog.model.run()
  assert.equal(result.code, 'CANCELED')
  assert.equal(dialog.exports.length, 0)
  assert.equal(dialog.toasts.length, 0)
})

test('a blocked commit stops the export before any dialog', async () => {
  const session = { docId: 'd', prepare: async () => false }
  const { model, io, exports } = setup({ session })
  const outcome = await model.run()
  assert.equal(outcome.code, 'BLOCKED')
  assert.equal(io.calls.chosen.length, 0)
  assert.equal(exports.length, 0)
})

test('with a DocumentSession: edits are committed first and a save waits for the export', async () => {
  const io = createIO({ targets: [{ path: 'C:\\Out\\Budget.pdf', format: 'pdf' }] })
  const adapterSaves = []
  let commits = 0
  const adapter = {
    async commitPendingEdits() { commits += 1 },
    async save(request) {
      adapterSaves.push(request)
      return { ok: true, path: 'C:\\Users\\me\\Documents\\Finance\\Budget.xlsx', name: 'Budget.xlsx', format: 'xlsx' }
    },
    async snapshotForRecovery() { return null },
    async restoreFromRecovery() {},
    title: () => 'Budget.xlsx',
    notify() {},
  }
  const session = new DocumentSession(adapter, io, { docId: 'doc-9', recovery: { enabled: false }, log: () => {} })
  const exportDone = deferred()
  const { model, exports } = setup({ io, session, results: [(request) => exportDone.promise.then(() => exported(request))] })
  session.noteChange()
  const exporting = model.run()
  await flush()
  assert.equal(commits, 1)
  assert.equal(exports[0].docId, 'doc-9')
  assert.equal(session.state.activity, 'export')
  const saving = session.save()
  await flush()
  assert.equal(adapterSaves.length, 0)
  assert.equal(session.state.status, 'waitingForExport')
  exportDone.resolve()
  assert.equal((await exporting).ok, true)
  assert.equal(await saving, true)
  assert.equal(adapterSaves.length, 1)
  assert.equal(session.dirty, false, 'export is a copy; only the save cleared dirty')
  session.dispose()
})

test('an export that replaced the document\'s own file marks the document changed, so Save writes it again', async () => {
  const own = 'C:\\Users\\me\\Documents\\Finance\\data.csv'
  const io = createIO({ targets: [{ path: own, format: 'csv' }] })
  const adapterSaves = []
  const adapter = {
    async commitPendingEdits() {},
    // Docs-style: the model equals the opened file, so the adapter calls it pristine.
    isPristine: () => true,
    async save(request) {
      adapterSaves.push(request)
      return { ok: true, path: own, name: 'data.csv', format: 'csv', strategy: 'rename' }
    },
    async snapshotForRecovery() { return null },
    async restoreFromRecovery() {},
    title: () => 'data.csv',
    notify() {},
  }
  const session = new DocumentSession(adapter, io, { docId: 'doc-own', recovery: { enabled: false }, log: () => {} })
  const { model } = setup({ io, session, documentName: 'data.csv', documentPath: own, results: [(request) => exported(request, { ownFileChanged: true })] })
  assert.equal(session.dirty, false)
  assert.equal((await model.run()).ok, true)
  assert.equal(session.dirty, true, 'the file holds the export now, not the document')
  assert.equal(await session.save(), true)
  assert.equal(adapterSaves[0].pristine, false, 'main is never asked to skip the write')
  assert.equal(session.dirty, false)
  session.dispose()
})

test('a failing preference store never fails an export', async () => {
  const { model, io } = setup({ targets: [{ path: 'C:\\Out\\Budget.pdf', format: 'pdf' }] })
  io.prefs.set = async () => { throw new TypeError('"lastExportFolder" must be an absolute folder path.') }
  io.prefs.get = async () => { throw new Error('unreadable') }
  await model.load()
  assert.equal((await model.run()).ok, true)
})

test('subscribers see every change; dispose stops notifications', () => {
  const { model } = setup()
  let calls = 0
  const unsubscribe = model.subscribe(() => { calls += 1 })
  model.select('csv')
  model.setOpenAfter(true)
  model.setOpenAfter(true)
  assert.equal(calls, 2)
  unsubscribe()
  model.select('pdf')
  assert.equal(calls, 2)
  model.dispose()
})

test('without the bridge the model explains nothing was exported', async () => {
  const model = new ExportModel({ io: null, module: 'calc', documentName: 'B.xlsx', formats: registry.exportFormats('calc', { engine: false }), exporter: async () => ({ ok: true, path: 'x', name: 'x', format: 'pdf' }) })
  const outcome = await model.run()
  assert.equal(outcome.ok, false)
  assert.equal(outcome.code, 'UNKNOWN')
})

// ---------------------------------------------------------------------------
// Local only
// ---------------------------------------------------------------------------

test('shared renderer code stays local: no network, no other programs, no installers', () => {
  for (const file of ['io-client.ts', 'document-session.ts', 'export-model.ts']) {
    // A Windows checkout (core.autocrlf) has CRLF line endings.
    const source = readFileSync(new URL(file, RENDERER), 'utf8').replace(/\r\n/g, '\n')
    assert.doesNotMatch(source, /\bopenPath\s*\(/, `${file} opens documents through openInSimple only`)
    assert.doesNotMatch(source, /openExternal/, file)
    assert.doesNotMatch(source, /\bfetch\s*\(|XMLHttpRequest|WebSocket|EventSource|sendBeacon/, `${file} makes no network requests`)
    assert.doesNotMatch(source, /https?:\/\//, `${file} has no remote URLs`)
    assert.doesNotMatch(source, /\.requestInstall\s*\(/, `${file} never asks to install anything`)
    assert.ok(source.startsWith(`// Vendored from simple/shared/renderer/${file} by simple/scripts/sync-shared.cjs. Do not edit here.\n`), `${file} carries the vendoring header`)
  }
})
