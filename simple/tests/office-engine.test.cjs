'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { fileURLToPath } = require('node:url')
const engine = require('../shared/electron/office-engine.cjs')

const KNOWN_HASH = Object.keys(engine.KNOWN_RUNTIME_HASHES)[0]
const KNOWN_VERSION = engine.KNOWN_RUNTIME_HASHES[KNOWN_HASH]
const SETUP_WORDING = /\b(install|installs|installed|installing|installer|download|downloads|downloading|libreoffice|soffice)\b/i

async function scratch(prefix) {
  return fs.mkdtemp(path.join(os.tmpdir(), `simple-office-engine-${prefix}-`))
}

async function remove(directory) {
  const resolved = path.resolve(directory)
  assert.ok(resolved.startsWith(path.resolve(os.tmpdir()) + path.sep) && path.basename(resolved).startsWith('simple-office-engine-'))
  await fs.rm(resolved, { recursive: true, force: true })
}

async function touch(filePath, contents = 'x') {
  await fs.mkdir(path.dirname(filePath), { recursive: true })
  await fs.writeFile(filePath, contents)
}

/** An isolated environment: no real %LOCALAPPDATA%, Program Files or companion folder is ever probed. */
function isolatedEnv(root, extra = {}) {
  return {
    LOCALAPPDATA: path.join(root, 'local'),
    ProgramFiles: path.join(root, 'pf'),
    'ProgramFiles(x86)': path.join(root, 'pf86'),
    ...extra,
  }
}

async function writeManifest(runtimeDirectory, manifest, encoding = 'utf8-bom') {
  const json = JSON.stringify(manifest, null, 2)
  let bytes
  if (encoding === 'utf8-bom') bytes = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(json, 'utf8')])
  else if (encoding === 'utf16le-bom') bytes = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(json, 'utf16le')])
  else bytes = Buffer.from(json, 'utf8')
  await touch(path.join(runtimeDirectory, 'simple-runtime.json'), bytes)
}

test('SIMPLE_FORCE_NO_OFFICE=1 reports no engine even when one exists, and conversions say what to do natively', async () => {
  const root = await scratch('forced')
  try {
    const executable = path.join(root, 'engine', 'soffice.exe')
    await touch(executable)
    const env = isolatedEnv(root, { SIMPLE_LIBREOFFICE_PATH: executable, SIMPLE_FORCE_NO_OFFICE: '1' })
    const status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.equal(status.available, false)
    assert.equal(status.reason, 'forced-off')
    assert.ok(Object.isFrozen(status))
    assert.equal(await engine.findOfficeConverter({ env, applicationRoots: [] }), null)
    assert.equal(engine.isOfficeEngineForcedOff({ SIMPLE_FORCE_NO_OFFICE: 'true' }), true)
    assert.equal(engine.isOfficeEngineForcedOff({ SIMPLE_FORCE_NO_OFFICE: '0' }), false)
    let calls = 0
    await assert.rejects(
      engine.convertOfficeBytes({ bytes: Buffer.from('legacy'), inputExtension: 'doc', outputExtension: 'pdf', filter: 'writer_pdf_Export' }, { env, executable, run: async () => { calls += 1 } }),
      (error) => {
        assert.equal(error.code, 'NEEDS_OFFICE_ENGINE')
        assert.equal(error.reason, 'forced-off')
        assert.equal(error.altExt, '.docx')
        assert.equal(error.workspace, 'docs')
        assert.match(error.message, /Open the file in Simple and save it as \.docx/)
        assert.doesNotMatch(error.message, SETUP_WORDING)
        assert.ok(engine.isOfficeEngineError(error))
        return true
      },
    )
    assert.equal(calls, 0, 'a forced-off engine must never be started')
  } finally { await remove(root) }
})

test('nothing on this PC: not-installed, and every conversion fails with a coded NEEDS_OFFICE_ENGINE error', async () => {
  const root = await scratch('absent')
  try {
    const env = isolatedEnv(root)
    const status = await engine.getOfficeEngineStatus({ env, applicationRoots: [path.join(root, 'app')], refresh: true })
    assert.deepEqual({ available: status.available, reason: status.reason, source: status.source, path: status.path }, { available: false, reason: 'not-installed', source: null, path: null })
    await assert.rejects(
      engine.convertOfficeBytes({ bytes: Buffer.from('xls bytes'), inputExtension: 'xls', outputExtension: 'pdf' }, { env, applicationRoots: [] }),
      (error) => error.code === 'NEEDS_OFFICE_ENGINE' && error.reason === 'not-installed' && error.altExt === '.xlsx' && /save it as \.xlsx/.test(error.message),
    )
  } finally { await remove(root) }
})

test('NEEDS_OFFICE_ENGINE wording names a native alternative and never asks to set up software', () => {
  for (const from of engine.INPUT_EXTENSIONS) {
    for (const to of engine.OUTPUT_EXTENSIONS) {
      if (from === to) continue
      const error = engine.needsOfficeEngineError({ inputExtension: from, outputExtension: to }, 'not-installed')
      assert.equal(error.code, 'NEEDS_OFFICE_ENGINE')
      assert.ok(error.message.length > 20, `${from}->${to} needs a message`)
      assert.doesNotMatch(error.message, SETUP_WORDING, `${from}->${to}: ${error.message}`)
      assert.ok(error.formatLabel, `${from}->${to} needs a format label`)
      assert.match(error.message, /Simple/, `${from}->${to} says what Simple does`)
    }
  }
  const save = engine.needsOfficeEngineError({ inputExtension: 'xlsx', outputExtension: 'xls' })
  assert.match(save.message, /Save it as an Excel workbook \(\.xlsx\) instead/)
  assert.deepEqual([save.altFormat, save.altExt, save.workspace, save.formatLabel], ['xlsx', '.xlsx', 'calc', 'older Excel workbook (.xls)'])
  const word = engine.needsOfficeEngineError({ inputExtension: 'docx', outputExtension: 'doc' })
  assert.match(word.message, /Save it as a Word document \(\.docx\) instead/)
  const slides = engine.needsOfficeEngineError({ inputExtension: 'pptx', outputExtension: 'pdf' })
  assert.match(slides.message, /Save the presentation as a PDF, then open the PDF in Simple/)
})

test('SIMPLE_LIBREOFFICE_PATH, a companion folder and a Program Files install are found, in that order', async () => {
  const root = await scratch('sources')
  try {
    const systemExecutable = path.join(root, 'pf', 'LibreOffice', 'program', 'soffice.exe')
    await touch(systemExecutable)
    await touch(path.join(root, 'pf', 'LibreOffice', 'program', 'soffice.bin'))
    await touch(path.join(root, 'pf', 'LibreOffice', 'program', 'version.ini'), '[Version]\r\nMsiProductVersion=26.2.6.2\r\n')
    const env = isolatedEnv(root)
    let status = await engine.getOfficeEngineStatus({ env, applicationRoots: [path.join(root, 'app')], refresh: true })
    assert.deepEqual([status.available, status.source, status.path, status.verified, status.version], [true, 'system', systemExecutable, true, '26.2.6.2'])

    const companion = path.join(root, 'app', 'tools', 'libreoffice', 'program', 'soffice.exe')
    await touch(companion)
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [path.join(root, 'app')], refresh: true })
    assert.deepEqual([status.source, status.path, status.verified], ['app-folder', companion, false], 'a companion copy without soffice.bin is used but reported as unverified')

    const explicit = path.join(root, 'managed', 'soffice.exe')
    await touch(explicit)
    status = await engine.getOfficeEngineStatus({ env: { ...env, SIMPLE_LIBREOFFICE_PATH: explicit }, applicationRoots: [path.join(root, 'app')], refresh: true })
    assert.deepEqual([status.source, status.path, status.verified], ['env', explicit, true])
    assert.equal(await engine.findOfficeConverter({ env: { ...env, SIMPLE_LIBREOFFICE_PATH: explicit }, applicationRoots: [path.join(root, 'app')] }), explicit)

    status = await engine.getOfficeEngineStatus({ env: { ...env, SIMPLE_LIBREOFFICE_PATH: path.join(root, 'missing.exe') }, applicationRoots: [path.join(root, 'app')], refresh: true })
    assert.equal(status.source, 'app-folder', 'a missing managed path falls back to the next local engine')
  } finally { await remove(root) }
})

test("Simple's per-user runtime counts only with a verified manifest; a UTF-8 or UTF-16 byte order mark is fine", async () => {
  const root = await scratch('runtime')
  try {
    const runtime = path.join(root, 'runtime')
    const env = isolatedEnv(root, { SIMPLE_OFFICE_RUNTIME_DIR: runtime })
    const executable = path.join(runtime, 'program', 'soffice.exe')
    await touch(executable)

    let status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.deepEqual([status.available, status.reason, status.source], [false, 'manifest-invalid', 'user-runtime'], 'no manifest: a half-prepared runtime is never used')

    await writeManifest(runtime, { version: KNOWN_VERSION, installerSha256: KNOWN_HASH.toUpperCase(), publisher: 'The Document Foundation' }, 'utf8-bom')
    const raw = await fs.readFile(path.join(runtime, 'simple-runtime.json'))
    assert.deepEqual([...raw.subarray(0, 3)], [0xef, 0xbb, 0xbf], 'fixture starts with the BOM Windows PowerShell 5.1 writes')
    assert.throws(() => JSON.parse(raw.toString('utf8')), SyntaxError, 'plain JSON.parse rejects it')
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.deepEqual([status.available, status.source, status.path, status.verified, status.version, status.reason], [true, 'user-runtime', executable, true, KNOWN_VERSION, null])

    await writeManifest(runtime, { version: KNOWN_VERSION, installerSha256: KNOWN_HASH }, 'utf16le-bom')
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.equal(status.available, true, 'UTF-16 manifests parse too')

    await writeManifest(runtime, { version: KNOWN_VERSION, installerSha256: 'ab'.repeat(32) })
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.deepEqual([status.available, status.reason], [false, 'manifest-invalid'], 'an unknown installer hash is rejected')
    assert.match(status.detail, /unknown installer hash/)

    await writeManifest(runtime, { version: '1.0', installerSha256: KNOWN_HASH })
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.equal(status.reason, 'manifest-invalid', 'a version that does not match the installer is rejected')

    await touch(path.join(runtime, 'simple-runtime.json'), '{ not json')
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.equal(status.reason, 'manifest-invalid')

    await fs.rm(executable)
    await writeManifest(runtime, { version: KNOWN_VERSION, installerSha256: KNOWN_HASH })
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.deepEqual([status.available, status.reason], [false, 'binary-missing'])

    const systemExecutable = path.join(root, 'pf', 'LibreOffice', 'program', 'soffice.exe')
    await touch(systemExecutable)
    status = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.deepEqual([status.available, status.source], [true, 'system'], 'an unusable runtime falls back to an engine installed on this PC')

    assert.deepEqual(engine.parseJsonWithBom(Buffer.from('﻿{"a":1}', 'utf8')), { a: 1 })
    const bigEndian = Buffer.from('{"b":2}', 'utf16le').swap16()
    assert.deepEqual(engine.parseJsonWithBom(Buffer.concat([Buffer.from([0xfe, 0xff]), bigEndian])), { b: 2 })
  } finally { await remove(root) }
})

test('the probe is cached, and a manifest change invalidates the cache at once', async () => {
  const root = await scratch('cache')
  try {
    const runtime = path.join(root, 'runtime')
    const env = isolatedEnv(root, { SIMPLE_OFFICE_RUNTIME_DIR: runtime })
    await touch(path.join(runtime, 'program', 'soffice.exe'))
    const manifestPath = path.join(runtime, 'simple-runtime.json')
    await writeManifest(runtime, { version: KNOWN_VERSION, installerSha256: 'cd'.repeat(32) })
    const first = await engine.getOfficeEngineStatus({ env, applicationRoots: [], refresh: true })
    assert.equal(first.reason, 'manifest-invalid')

    const companion = path.join(root, 'app', 'tools', 'libreoffice', 'program', 'soffice.exe')
    await touch(companion)
    const cached = await engine.getOfficeEngineStatus({ env, applicationRoots: [path.join(root, 'app')] })
    assert.equal(cached.source, 'app-folder', 'different application roots are a different cache entry')
    const same = await engine.getOfficeEngineStatus({ env, applicationRoots: [] })
    assert.equal(same, first, 'an unchanged environment returns the cached status object')

    await writeManifest(runtime, { version: KNOWN_VERSION, installerSha256: KNOWN_HASH })
    const later = new Date(Date.now() + 5000)
    await fs.utimes(manifestPath, later, later)
    const updated = await engine.getOfficeEngineStatus({ env, applicationRoots: [] })
    assert.notEqual(updated, first)
    assert.deepEqual([updated.available, updated.source], [true, 'user-runtime'], 'the new manifest is seen without waiting for the cache to expire')

    engine.clearOfficeEngineStatusCache()
    const fresh = await engine.getOfficeEngineStatus({ env, applicationRoots: [] })
    assert.notEqual(fresh, updated)
  } finally { await remove(root) }
})

test('conversion runs in a private job folder; prepareInput and policyId are part of the cache key', async () => {
  const root = await scratch('convert')
  try {
    const executable = path.join(root, 'engine', 'soffice.exe')
    await touch(executable)
    const cacheDirectory = path.join(root, 'cache')
    const runs = []
    const jobFolders = new Set()
    const run = async (file, args, timeoutMs) => {
      assert.equal(file, executable)
      assert.equal(timeoutMs, 60_000)
      const inputPath = args.at(-1)
      const outdir = args[args.indexOf('--outdir') + 1]
      jobFolders.add(path.dirname(inputPath))
      const profile = fileURLToPath(args[0].slice('-env:UserInstallation='.length))
      const registry = await fs.readFile(path.join(profile, 'user', 'registrymodifications.xcu'), 'utf8')
      assert.match(registry, /MacroSecurityLevel[\s\S]*<value>3<\/value>/, 'macros are disabled in the private profile')
      assert.ok(args.includes('--headless') && args.includes('--norestore') && args.includes('--nolockcheck'))
      assert.ok(path.basename(path.dirname(inputPath)).startsWith('simple-office-convert-'))
      const input = await fs.readFile(inputPath)
      runs.push({ input: input.toString(), target: args[args.indexOf('--convert-to') + 1] })
      await fs.writeFile(path.join(outdir, 'document.pdf'), Buffer.from(`%PDF-1.7 from ${input}`))
    }
    const request = { bytes: Buffer.from('legacy words'), inputExtension: '.DOC', outputExtension: 'pdf', filter: 'writer_pdf_Export' }
    const prepareInput = (bytes) => ({ bytes: Buffer.concat([bytes, Buffer.from(' (prepared)')]) })
    const base = { executable, run, cacheDirectory, env: {} }

    const first = await engine.convertOfficeBytes(request, { ...base, prepareInput, policyId: 'lock-dates-1' })
    assert.equal(first.toString(), '%PDF-1.7 from legacy words (prepared)')
    assert.deepEqual(runs.map((item) => item.target), ['pdf:writer_pdf_Export'])
    await engine.convertOfficeBytes(request, { ...base, prepareInput, policyId: 'lock-dates-1' })
    assert.equal(runs.length, 1, 'the same input and policy come from the cache')
    await engine.convertOfficeBytes(request, { ...base, prepareInput, policyId: 'lock-dates-2' })
    assert.equal(runs.length, 2, 'a new policyId converts again')
    const plain = await engine.convertOfficeBytes(request, base)
    assert.equal(runs.length, 3, 'no preparation is a different cache entry')
    assert.equal(plain.toString(), '%PDF-1.7 from legacy words')
    await engine.convertOfficeBytes({ ...request, filter: 'other_pdf_Export' }, base)
    assert.equal(runs.length, 4, 'the filter is part of the key')

    engine.clearConversionMemoryCache()
    const fromDisk = await engine.convertOfficeBytes(request, { ...base, prepareInput, policyId: 'lock-dates-1' })
    assert.equal(runs.length, 4, 'the disk cache serves a conversion after the memory cache is cleared')
    assert.equal(fromDisk.toString(), first.toString())

    await assert.rejects(engine.convertOfficeBytes(request, { ...base, prepareInput }), TypeError, 'prepareInput without a policyId is refused')
    await engine.convertOfficeBytes(request, { executable, run, env: {}, prepareInput, policyId: 'lock-dates-1' })
    assert.equal(runs.length, 5, 'an injected runner without a cache folder is never cached')
    assert.equal(jobFolders.size, 5, 'every conversion gets its own job folder')
    for (const folder of jobFolders) assert.equal(await fs.stat(folder).then(() => true, () => false), false, `${folder} was removed`)
  } finally { await remove(root) }
})

test('engine failures are coded and keep their detail out of the message', async () => {
  const root = await scratch('failures')
  try {
    const executable = path.join(root, 'soffice.exe')
    await touch(executable)
    const request = { bytes: Buffer.from('x'), inputExtension: 'docx', outputExtension: 'pdf' }
    const options = { executable, env: {}, cache: false }
    await assert.rejects(engine.convertOfficeBytes(request, { ...options, run: async () => {} }), (error) => error.code === 'CONVERSION_FAILED' && /no pdf file/.test(error.technical))
    await assert.rejects(
      engine.convertOfficeBytes(request, { ...options, run: async (_file, args) => fs.writeFile(path.join(args[args.indexOf('--outdir') + 1], 'document.pdf'), 'not a pdf at all') }),
      (error) => error.code === 'OUTPUT_INVALID' && !/not a pdf/.test(error.message),
    )
    await assert.rejects(engine.convertOfficeBytes(request, { ...options, run: async () => { throw new Error('boom') } }), /boom/)
    await assert.rejects(engine.convertOfficeBytes({ ...request, inputExtension: 'exe' }, options), (error) => error.code === 'UNSUPPORTED_CONVERSION')
    await assert.rejects(engine.convertOfficeBytes({ ...request, bytes: Buffer.alloc(0) }, options), (error) => error.code === 'EMPTY_INPUT')
    await assert.rejects(engine.convertOfficeBytes({ ...request, filter: 'a\nb' }, options), (error) => error.code === 'UNSUPPORTED_CONVERSION')
    await assert.rejects(
      engine.convertOfficeBytes(request, { executable: path.join(root, 'gone.exe'), env: {} }),
      (error) => error.code === 'NEEDS_OFFICE_ENGINE' && error.reason === 'binary-missing',
    )
    await assert.rejects(
      engine.convertOfficeBytes(request, { ...options, prepareInput: () => { throw new Error('bad container') }, policyId: 'p', run: async () => {} }),
      (error) => error.code === 'CONVERSION_FAILED' && /bad container/.test(error.technical),
    )

    const started = Date.now()
    await assert.rejects(engine.runOfficeConverter(process.execPath, ['-e', 'setTimeout(() => {}, 20000)'], 400), (error) => error.code === 'CONVERSION_TIMEOUT')
    assert.ok(Date.now() - started < 15_000, 'a stuck engine is stopped')
    await assert.rejects(engine.runOfficeConverter(process.execPath, ['-e', 'console.error("engine said no"); process.exit(3)']), (error) => (
      error.code === 'CONVERSION_FAILED' && /Exit code 3/.test(error.technical) && /engine said no/.test(error.technical) && !/engine said no/.test(error.message)
    ))
    await assert.rejects(engine.runOfficeConverter(path.join(root, 'missing.exe'), []), (error) => error.code === 'CONVERSION_FAILED')
    await engine.runOfficeConverter(process.execPath, ['-e', 'process.exit(0)'])
  } finally { await remove(root) }
})

test('shared office engine source: built-ins only, no setup wording, and the vendoring header', async () => {
  const source = await fs.readFile(path.join(__dirname, '..', 'shared', 'electron', 'office-engine.cjs'), 'utf8')
  assert.equal(source.split(/\r?\n/, 1)[0], '// Vendored from simple/shared/electron/office-engine.cjs by simple/scripts/sync-shared.cjs. Do not edit here.')
  for (const match of source.matchAll(/require\(\s*['"]([^'"]+)['"]\s*\)/g)) assert.match(match[1], /^node:/, `office-engine.cjs may only require Node built-ins, found ${match[1]}`)
  assert.equal(/Install LibreOffice/i.test(source), false, 'no "Install LibreOffice" wording')
  // XML namespace names in the private profile are identifiers, never fetched.
  const urls = source.replace(/xmlns:\w+="[^"]*"/g, '').match(/\b(?:https?|ftp|wss?):\/\/[^\s'"`)]+/gi) || []
  assert.deepEqual(urls, [], 'no remote addresses in the office engine')
  for (const label of Object.values(engine.FORMAT_LABELS)) assert.doesNotMatch(label, SETUP_WORDING)
})
