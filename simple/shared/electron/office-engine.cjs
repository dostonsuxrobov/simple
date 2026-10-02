// Vendored from simple/shared/electron/office-engine.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// The one office converter every workspace uses. It only ever uses a
// LibreOffice that is already on this PC: Simple never fetches, installs or
// updates one, and every format Simple offers works without it.
//
// - getOfficeEngineStatus() is a cheap, cached probe. Simple's per-user
//   runtime counts only when its simple-runtime.json manifest (read with or
//   without a byte order mark) names a known installer hash.
//   SIMPLE_FORCE_NO_OFFICE=1 makes every probe report "no engine".
// - convertOfficeBytes() converts in a private, macro-free profile with a hard
//   timeout and caches results by content. Without an engine it throws a coded
//   NEEDS_OFFICE_ENGINE error whose message says what Simple can do natively.
// - Workspace-specific preparation (for example locking legacy date fields in
//   a .doc) is injected through prepareInput and named by policyId, which is
//   part of the cache key.
// Only Node built-ins are required.

const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { execFile, spawn } = require('node:child_process')
const { pathToFileURL } = require('node:url')

/** Changes whenever conversion output for the same input could change. Part of every cache key. */
const CONVERSION_POLICY_VERSION = 'shared-engine-1'
const DEFAULT_TIMEOUT_MS = 60_000
const STATUS_TTL_MS = 30_000
const MAX_INPUT_BYTES = 256 * 1024 * 1024
const MAX_OUTPUT_BYTES = 512 * 1024 * 1024
const MAX_DISK_CACHE_BYTES = 256 * 1024 * 1024
const MAX_DISK_CACHE_FILES = 128
const MAX_MEMORY_CACHE_BYTES = 64 * 1024 * 1024
const MAX_MEMORY_CACHE_ENTRIES = 8
const MAX_DETAIL_CHARS = 4000
const JOB_PREFIX = 'simple-office-convert-'
const RUNTIME_MANIFEST = 'simple-runtime.json'

/**
 * Installer hashes of the runtime that simple/scripts/setup-office-runtime.ps1
 * prepares, mapped to the version it must report. A per-user runtime whose
 * manifest names any other hash is not used.
 */
const KNOWN_RUNTIME_HASHES = Object.freeze({
  f9877032fd908beb9c0ddf06df4af5c2e85f419c42e14876c4cce5aae5fb2660: '26.2.6',
})

/** Layouts an extracted per-user runtime can have, relative to its folder. */
const RUNTIME_LAYOUTS = Object.freeze([
  ['program', 'soffice.exe'],
  ['LibreOffice', 'program', 'soffice.exe'],
  ['Program Files', 'LibreOffice', 'program', 'soffice.exe'],
])

const INPUT_EXTENSIONS = Object.freeze(['doc', 'docx', 'xls', 'xlsx', 'odt', 'ods', 'rtf', 'ppt', 'pptx'])
const OUTPUT_EXTENSIONS = Object.freeze(['docx', 'doc', 'pdf', 'xls', 'xlsx', 'odt', 'ods'])

/** Plain names for every format the engine handles, used in user-facing messages. */
const FORMAT_LABELS = Object.freeze({
  doc: 'older Word document (.doc)',
  docx: 'Word document (.docx)',
  xls: 'older Excel workbook (.xls)',
  xlsx: 'Excel workbook (.xlsx)',
  odt: 'OpenDocument text (.odt)',
  ods: 'OpenDocument spreadsheet (.ods)',
  rtf: 'Rich Text document (.rtf)',
  ppt: 'older PowerPoint presentation (.ppt)',
  pptx: 'PowerPoint presentation (.pptx)',
  pdf: 'PDF',
})

/** Per-job LibreOffice profile: macros never run and external links are never updated. */
const PROFILE_REGISTRY = '<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry">'
  + '<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item>'
  + '<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>0</value></prop></item>'
  + '<item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item>'
  + '</oor:items>'

const statusCache = new Map()
const conversionCache = new Map()
let memoryCacheBytes = 0
let cachePrunePromise = Promise.resolve()

// ---------------------------------------------------------------------------
// Errors
// ---------------------------------------------------------------------------

const ERROR_FIELDS = Object.freeze([
  'technical', 'reason', 'inputExtension', 'outputExtension', 'formatLabel', 'altFormat', 'altLabel', 'altExt', 'workspace', 'source', 'path',
])

/**
 * A coded conversion error. `message` is plain, user-facing text that never
 * asks the user to fetch or set up software; `technical` holds the detail for
 * logs and the "Details" line.
 *
 * Codes: NEEDS_OFFICE_ENGINE (no usable engine; carries formatLabel, altFormat,
 * altLabel, altExt and workspace for the native alternative),
 * UNSUPPORTED_CONVERSION, EMPTY_INPUT, TOO_LARGE, CONVERSION_TIMEOUT,
 * CONVERSION_FAILED and OUTPUT_INVALID. NEEDS_OFFICE_ENGINE and TOO_LARGE use
 * the same code names as io-core's IoError.
 */
class OfficeEngineError extends Error {
  /**
   * @param {string} code
   * @param {string} message user-facing text
   * @param {object} [details] any of technical, reason, inputExtension, outputExtension,
   *   formatLabel, altFormat, altLabel, altExt, workspace, source, path
   */
  constructor(code, message, details = {}) {
    super(message)
    this.name = 'OfficeEngineError'
    this.code = code
    for (const key of ERROR_FIELDS) if (details[key] !== undefined && details[key] !== null) this[key] = details[key]
  }
}

/**
 * True for a coded error from this module, also across bundles and realms.
 * @param {unknown} value
 * @returns {boolean}
 */
function isOfficeEngineError(value) {
  return value instanceof OfficeEngineError || Boolean(value && value.name === 'OfficeEngineError' && typeof value.code === 'string')
}

function labelFor(extension) {
  return FORMAT_LABELS[extension] || (extension ? `.${extension}` : 'this kind of')
}

/**
 * What Simple can do natively when a conversion needs the engine.
 * @param {string} from input extension without the dot
 * @param {string} to output extension without the dot
 * @returns {{altFormat: string|null, workspace: 'docs'|'calc'|null, message: string}}
 */
function nativeAlternative(from, to) {
  if (to === 'doc' || to === 'odt') {
    return { altFormat: 'docx', workspace: 'docs', message: `Simple can't save ${labelFor(to)} files on this PC. Save it as a Word document (.docx) instead; the original file stays unchanged.` }
  }
  if (to === 'xls' || to === 'ods') {
    return { altFormat: 'xlsx', workspace: 'calc', message: `Simple can't save ${labelFor(to)} files on this PC. Save it as an Excel workbook (.xlsx) instead; the original file stays unchanged.` }
  }
  if (from === 'doc') {
    return { altFormat: 'docx', workspace: 'docs', message: "Simple can't fully convert older Word documents (.doc) on this PC. Open the file in Simple and save it as .docx first, then use the .docx." }
  }
  if (from === 'xls') {
    return { altFormat: 'xlsx', workspace: 'calc', message: "Simple can't fully convert older Excel workbooks (.xls) on this PC. Open the file in Simple and save it as .xlsx first, then use the .xlsx." }
  }
  if (from === 'ppt' || from === 'pptx') {
    return { altFormat: 'pdf', workspace: null, message: "Simple can't turn presentations into other formats on this PC. Save the presentation as a PDF, then open the PDF in Simple." }
  }
  if (from === 'rtf' || from === 'odt') {
    return { altFormat: 'docx', workspace: null, message: `Simple can't fully convert ${labelFor(from)} files on this PC. Save a Word (.docx) or PDF copy of it, then open that copy in Simple.` }
  }
  if (from === 'ods') {
    return { altFormat: 'xlsx', workspace: null, message: "Simple can't fully convert OpenDocument spreadsheets (.ods) on this PC. Save an Excel (.xlsx) copy of it, then open that copy in Simple." }
  }
  if (from === 'docx') {
    return { altFormat: 'pdf', workspace: 'docs', message: `Simple can't make a layout-exact ${labelFor(to)} of this Word document on this PC. Open it in Simple and export it from there instead.` }
  }
  if (from === 'xlsx') {
    return { altFormat: 'pdf', workspace: 'calc', message: `Simple can't make a layout-exact ${labelFor(to)} of this workbook on this PC. Open it in Simple and export it from there instead.` }
  }
  return { altFormat: null, workspace: null, message: "Simple can't convert this file on this PC. The original file is unchanged." }
}

/**
 * Builds the coded error for a conversion that needs an engine this PC does not have.
 * @param {{inputExtension: string, outputExtension: string}} request extensions without the dot
 * @param {string|null} [reason] the probe's reason ('forced-off', 'not-installed', ...)
 * @returns {OfficeEngineError}
 */
function needsOfficeEngineError(request, reason = null) {
  const from = normalizeExtension(request.inputExtension)
  const to = normalizeExtension(request.outputExtension)
  const alternative = nativeAlternative(from, to)
  const writing = ['doc', 'odt', 'xls', 'ods'].includes(to)
  return new OfficeEngineError('NEEDS_OFFICE_ENGINE', alternative.message, {
    reason,
    inputExtension: from,
    outputExtension: to,
    formatLabel: labelFor(writing ? to : from),
    altFormat: alternative.altFormat,
    altLabel: alternative.altFormat ? labelFor(alternative.altFormat) : null,
    altExt: alternative.altFormat ? `.${alternative.altFormat}` : null,
    workspace: alternative.workspace,
    technical: `No usable office engine (${reason || 'unavailable'}) for ${from} to ${to}.`,
  })
}

// ---------------------------------------------------------------------------
// Probe
// ---------------------------------------------------------------------------

/**
 * True when SIMPLE_FORCE_NO_OFFICE asks Simple to behave as if no engine exists.
 * @param {Record<string, string|undefined>} [env]
 * @returns {boolean}
 */
function isOfficeEngineForcedOff(env = process.env) {
  return /^(1|true|yes|on)$/i.test(String(env.SIMPLE_FORCE_NO_OFFICE || '').trim())
}

/**
 * Folder of Simple's per-user runtime: SIMPLE_OFFICE_RUNTIME_DIR, else
 * %LOCALAPPDATA%\simple\office-runtime, else null.
 * @param {Record<string, string|undefined>} [env]
 * @returns {string|null}
 */
function officeRuntimeDirectory(env = process.env) {
  if (env.SIMPLE_OFFICE_RUNTIME_DIR) return path.resolve(env.SIMPLE_OFFICE_RUNTIME_DIR)
  if (env.LOCALAPPDATA) return path.join(env.LOCALAPPDATA, 'simple', 'office-runtime')
  return null
}

/**
 * Folders whose tools\libreoffice or LibreOffice subfolder may hold a
 * companion engine: the portable EXE's folder, the app folder, its resources,
 * and the source tree two and three levels above this file (so the original
 * workspace layout, the vendored electron/simple-io layout and bundled
 * builds all find the same companion folder).
 * @returns {string[]}
 */
function defaultApplicationRoots() {
  const roots = [
    process.env.PORTABLE_EXECUTABLE_DIR,
    process.execPath ? path.dirname(process.execPath) : null,
    process.resourcesPath,
    path.resolve(__dirname, '..', '..'),
    path.resolve(__dirname, '..', '..', '..'),
  ]
  return [...new Set(roots.filter(Boolean).map((root) => path.resolve(root)))]
}

async function statOrNull(filePath) {
  try { return await fs.stat(filePath) } catch { return null }
}

async function isFile(filePath) {
  const stat = await statOrNull(filePath)
  return Boolean(stat && stat.isFile())
}

/**
 * Decodes JSON text that may start with a UTF-8 or UTF-16 byte order mark
 * (Windows PowerShell 5.1 writes one with Set-Content -Encoding utf8).
 * @param {Buffer} bytes
 * @returns {unknown}
 * @throws {SyntaxError} when the text is not JSON
 */
function parseJsonWithBom(bytes) {
  let text
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) text = bytes.subarray(2).toString('utf16le')
  else if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    const swapped = Buffer.from(bytes.subarray(2))
    swapped.swap16()
    text = swapped.toString('utf16le')
  } else text = bytes.toString('utf8')
  return JSON.parse(text.replace(/^﻿/, ''))
}

/**
 * Reads and checks a per-user runtime manifest.
 * @param {string} runtimeDirectory
 * @returns {Promise<{exists: boolean, ok: boolean, version: string|null, installerSha256: string|null, problem: string|null, mtimeMs: number|null}>}
 */
async function readRuntimeManifest(runtimeDirectory) {
  const manifestPath = path.join(runtimeDirectory, RUNTIME_MANIFEST)
  const stat = await statOrNull(manifestPath)
  if (!stat || !stat.isFile()) return { exists: false, ok: false, version: null, installerSha256: null, problem: 'missing', mtimeMs: null }
  const result = { exists: true, ok: false, version: null, installerSha256: null, problem: null, mtimeMs: stat.mtimeMs }
  if (stat.size > 64 * 1024) return { ...result, problem: 'too large' }
  let manifest
  try { manifest = parseJsonWithBom(await fs.readFile(manifestPath)) } catch (error) {
    return { ...result, problem: `unreadable (${error.message})` }
  }
  const hash = typeof manifest?.installerSha256 === 'string' ? manifest.installerSha256.trim().toLowerCase() : ''
  const version = typeof manifest?.version === 'string' ? manifest.version.trim() : ''
  result.installerSha256 = hash || null
  result.version = version || null
  if (!/^[a-f0-9]{64}$/.test(hash)) return { ...result, problem: 'no installer hash' }
  if (!Object.prototype.hasOwnProperty.call(KNOWN_RUNTIME_HASHES, hash)) return { ...result, problem: 'unknown installer hash' }
  if (version !== KNOWN_RUNTIME_HASHES[hash]) return { ...result, problem: `version ${version || 'missing'} does not match the installer` }
  return { ...result, ok: true }
}

/**
 * Best-effort LibreOffice version from program\version.ini; never executes anything.
 * @param {string} executable
 * @returns {Promise<string|null>}
 */
async function readProgramVersion(executable) {
  try {
    const iniPath = path.join(path.dirname(executable), 'version.ini')
    const stat = await statOrNull(iniPath)
    if (!stat || !stat.isFile() || stat.size > 64 * 1024) return null
    const text = await fs.readFile(iniPath, 'utf8')
    const match = /^\s*(?:MsiProductVersion|ProductVersion)\s*=\s*([0-9]+(?:\.[0-9]+)*)\s*$/mi.exec(text)
    return match ? match[1] : null
  } catch { return null }
}

/** True when the executable sits in a complete LibreOffice program folder. */
async function programFolderComplete(executable) {
  return isFile(path.join(path.dirname(executable), 'soffice.bin'))
}

function statusOf(fields) {
  return Object.freeze({
    available: false,
    source: null,
    path: null,
    version: null,
    verified: false,
    reason: null,
    detail: null,
    checkedAt: Date.now(),
    ...fields,
  })
}

async function probeOfficeEngine(env, applicationRoots) {
  const rejected = []
  const explicit = env.SIMPLE_LIBREOFFICE_PATH ? path.resolve(env.SIMPLE_LIBREOFFICE_PATH) : null
  if (explicit) {
    // A managed deployment chose this binary on purpose; it only has to exist.
    if (await isFile(explicit)) return statusOf({ available: true, source: 'env', path: explicit, version: await readProgramVersion(explicit), verified: true })
    rejected.push({ source: 'env', path: explicit, reason: 'binary-missing', detail: 'SIMPLE_LIBREOFFICE_PATH does not name a file.' })
  }

  const runtimeDirectory = officeRuntimeDirectory(env)
  if (runtimeDirectory) {
    let executable = null
    for (const layout of RUNTIME_LAYOUTS) {
      const candidate = path.join(runtimeDirectory, ...layout)
      if (await isFile(candidate)) { executable = candidate; break }
    }
    const manifest = await readRuntimeManifest(runtimeDirectory)
    if (executable && manifest.ok) {
      return statusOf({ available: true, source: 'user-runtime', path: executable, version: manifest.version, verified: true })
    }
    if (executable) rejected.push({ source: 'user-runtime', path: executable, reason: 'manifest-invalid', detail: `${RUNTIME_MANIFEST}: ${manifest.problem}` })
    else if (manifest.exists) rejected.push({ source: 'user-runtime', path: runtimeDirectory, reason: 'binary-missing', detail: 'The runtime folder has a manifest but no soffice.exe.' })
  }

  const candidates = []
  for (const root of applicationRoots) {
    candidates.push({ source: 'app-folder', path: path.join(root, 'tools', 'libreoffice', 'program', 'soffice.exe') })
    candidates.push({ source: 'app-folder', path: path.join(root, 'LibreOffice', 'program', 'soffice.exe') })
  }
  for (const root of [env.ProgramFiles, env['ProgramFiles(x86)']].filter(Boolean)) {
    candidates.push({ source: 'system', path: path.join(root, 'LibreOffice', 'program', 'soffice.exe') })
  }
  const seen = new Set()
  for (const candidate of candidates) {
    const key = candidate.path.toLowerCase()
    if (seen.has(key)) continue
    seen.add(key)
    if (!await isFile(candidate.path)) continue
    return statusOf({
      available: true,
      source: candidate.source,
      path: candidate.path,
      version: await readProgramVersion(candidate.path),
      verified: await programFolderComplete(candidate.path),
    })
  }

  const invalid = rejected.find((item) => item.reason === 'manifest-invalid')
  const missing = rejected.find((item) => item.reason === 'binary-missing')
  const relevant = invalid || missing || null
  return statusOf({
    reason: relevant ? relevant.reason : 'not-installed',
    source: relevant ? relevant.source : null,
    path: relevant ? relevant.path : null,
    detail: relevant ? relevant.detail : null,
  })
}

/**
 * @typedef {object} OfficeEngineStatus
 * @property {boolean} available true when a local engine can convert right now
 * @property {'env'|'user-runtime'|'app-folder'|'system'|null} source where the engine was found
 *   (for an unavailable status: the candidate that was rejected, if any)
 * @property {string|null} path the soffice executable (or the rejected candidate)
 * @property {string|null} version e.g. "26.2.6", when the manifest or version.ini says so
 * @property {boolean} verified per-user runtime: its manifest parsed and names a known installer hash;
 *   SIMPLE_LIBREOFFICE_PATH: the file exists; companion and system installs: the program folder is complete
 * @property {null|'forced-off'|'not-installed'|'manifest-invalid'|'binary-missing'} reason why no engine is available
 * @property {string|null} detail diagnostic text for logs (never shown as a user message)
 * @property {number} checkedAt Date.now() of the probe
 */

/**
 * Finds a local engine. The result is cached for 30 seconds, and the cache is
 * dropped as soon as the per-user runtime manifest changes. Never starts the
 * engine and never touches the network.
 *
 * @param {object} [options]
 * @param {boolean} [options.refresh=false] probe again even when a cached status is fresh
 * @param {Record<string, string|undefined>} [options.env=process.env] environment to read
 *   (SIMPLE_FORCE_NO_OFFICE, SIMPLE_LIBREOFFICE_PATH, SIMPLE_OFFICE_RUNTIME_DIR, LOCALAPPDATA, ProgramFiles)
 * @param {string[]} [options.applicationRoots] folders to search for a companion engine
 * @returns {Promise<OfficeEngineStatus>} a frozen status object
 */
async function getOfficeEngineStatus(options = {}) {
  const env = options.env || process.env
  if (isOfficeEngineForcedOff(env)) return statusOf({ reason: 'forced-off' })
  const applicationRoots = Array.isArray(options.applicationRoots)
    ? options.applicationRoots.filter((root) => typeof root === 'string' && root).map((root) => path.resolve(root))
    : defaultApplicationRoots()
  const runtimeDirectory = officeRuntimeDirectory(env)
  const key = JSON.stringify([
    env.SIMPLE_LIBREOFFICE_PATH || '', runtimeDirectory || '', env.ProgramFiles || '', env['ProgramFiles(x86)'] || '', applicationRoots,
  ])
  const manifestStat = runtimeDirectory ? await statOrNull(path.join(runtimeDirectory, RUNTIME_MANIFEST)) : null
  const manifestMtimeMs = manifestStat ? manifestStat.mtimeMs : null
  const cached = statusCache.get(key)
  if (!options.refresh && cached && cached.expires > Date.now() && cached.manifestMtimeMs === manifestMtimeMs) return cached.promise
  const entry = { expires: Date.now() + STATUS_TTL_MS, manifestMtimeMs, promise: probeOfficeEngine(env, applicationRoots) }
  statusCache.set(key, entry)
  try {
    return await entry.promise
  } catch (error) {
    if (statusCache.get(key) === entry) statusCache.delete(key)
    throw error
  }
}

/**
 * Compatibility helper for callers of the old per-workspace converters.
 * @param {object} [options] same as getOfficeEngineStatus
 * @returns {Promise<string|null>} the engine executable, or null when none is usable
 */
async function findOfficeConverter(options = {}) {
  const status = await getOfficeEngineStatus(options)
  return status.available ? status.path : null
}

// ---------------------------------------------------------------------------
// Conversion
// ---------------------------------------------------------------------------

function normalizeExtension(value) {
  return String(value || '').replace(/^\./, '').toLowerCase()
}

/**
 * Checks a conversion request.
 * @param {{bytes: Uint8Array|ArrayBuffer|Buffer, inputExtension: string, outputExtension: string, filter?: string}} input
 * @returns {{bytes: Buffer, inputExtension: string, outputExtension: string, filter: string}}
 * @throws {OfficeEngineError} UNSUPPORTED_CONVERSION, EMPTY_INPUT or TOO_LARGE
 */
function validateConversionInput(input) {
  const inputExtension = normalizeExtension(input?.inputExtension)
  const outputExtension = normalizeExtension(input?.outputExtension)
  if (!INPUT_EXTENSIONS.includes(inputExtension) || !OUTPUT_EXTENSIONS.includes(outputExtension) || inputExtension === outputExtension) {
    throw new OfficeEngineError('UNSUPPORTED_CONVERSION', `Simple can't convert ${labelFor(inputExtension)} files to ${labelFor(outputExtension)}.`, {
      inputExtension, outputExtension, technical: `Unsupported conversion ${inputExtension || '?'} to ${outputExtension || '?'}.`,
    })
  }
  const filter = String(input.filter || '')
  if (filter.length > 120 || /[\r\n\u0000]/.test(filter)) {
    throw new OfficeEngineError('UNSUPPORTED_CONVERSION', "Simple can't convert this file with the requested settings.", { inputExtension, outputExtension, technical: 'Invalid conversion filter.' })
  }
  const source = input.bytes
  if (!source || typeof source.byteLength !== 'number') {
    throw new OfficeEngineError('EMPTY_INPUT', 'The file is empty, so there is nothing to convert.', { inputExtension, outputExtension, technical: 'No input bytes.' })
  }
  const bytes = Buffer.isBuffer(source) ? source : Buffer.from(source instanceof ArrayBuffer ? new Uint8Array(source) : source)
  if (!bytes.length) throw new OfficeEngineError('EMPTY_INPUT', 'The file is empty, so there is nothing to convert.', { inputExtension, outputExtension, technical: 'Zero input bytes.' })
  if (bytes.length > MAX_INPUT_BYTES) {
    throw new OfficeEngineError('TOO_LARGE', 'The file is too large to convert. Simple converts files up to 256 MB.', { inputExtension, outputExtension, technical: `${bytes.length} input bytes.` })
  }
  return { bytes, inputExtension, outputExtension, filter }
}

/**
 * True when bytes start like a file of the given output type (CFB for .doc and
 * .xls, %PDF- for PDF, ZIP for the XML formats).
 * @param {Uint8Array|null|undefined} bytes
 * @param {string} extension without the dot
 * @returns {boolean}
 */
function hasOutputSignature(bytes, extension) {
  if (!bytes || bytes.length < 8) return false
  const buffer = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (extension === 'doc' || extension === 'xls') return buffer.subarray(0, 8).toString('hex') === 'd0cf11e0a1b11ae1'
  if (extension === 'pdf') return buffer.subarray(0, 5).toString('latin1') === '%PDF-'
  return buffer[0] === 0x50 && buffer[1] === 0x4b && buffer.length >= 30
}

function trimDetail(text) {
  const value = String(text || '').trim()
  return value.length > MAX_DETAIL_CHARS ? value.slice(-MAX_DETAIL_CHARS) : value
}

/**
 * A Windows system program by absolute path (System32), never by bare name:
 * Windows searches the current folder before PATH for a bare name, and the
 * current folder may be one where anyone could have placed a program.
 * @param {string} name e.g. "reg.exe"
 * @returns {string}
 */
function systemProgram(name) {
  const root = [process.env.SystemRoot, process.env.windir].find((value) => typeof value === 'string' && path.win32.isAbsolute(value)) || 'C:\\Windows'
  return path.win32.join(root, 'System32', name)
}

/**
 * Runs the engine once and settles when it exits. A run that exceeds the
 * timeout is stopped with its whole process tree.
 * @param {string} executable
 * @param {string[]} args
 * @param {number} [timeoutMs=60000]
 * @returns {Promise<void>}
 * @throws {OfficeEngineError} CONVERSION_TIMEOUT or CONVERSION_FAILED
 */
function runOfficeConverter(executable, args, timeoutMs = DEFAULT_TIMEOUT_MS) {
  return new Promise((resolve, reject) => {
    let child
    try {
      child = spawn(executable, args, { windowsHide: true, shell: false, stdio: ['ignore', 'pipe', 'pipe'] })
    } catch (error) {
      reject(new OfficeEngineError('CONVERSION_FAILED', "Simple couldn't convert this file. The original file is unchanged.", { technical: `${error.code || 'ERROR'} spawn ${executable}: ${error.message}` }))
      return
    }
    let detail = ''
    let timedOut = false
    let settled = false
    const collect = (data) => { detail = trimDetail(detail + data.toString()) }
    child.stdout.on('data', collect)
    child.stderr.on('data', collect)
    const timer = setTimeout(() => {
      timedOut = true
      if (process.platform === 'win32' && child.pid) execFile(systemProgram('taskkill.exe'), ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => {})
      else child.kill('SIGKILL')
    }, Math.max(1, Number(timeoutMs) || DEFAULT_TIMEOUT_MS))
    child.once('error', (error) => {
      clearTimeout(timer)
      if (settled) return
      settled = true
      reject(new OfficeEngineError('CONVERSION_FAILED', "Simple couldn't convert this file. The original file is unchanged.", { technical: `${error.code || 'ERROR'} spawn ${executable}: ${error.message}` }))
    })
    child.once('close', (code, signal) => {
      clearTimeout(timer)
      if (settled) return
      settled = true
      if (timedOut) {
        reject(new OfficeEngineError('CONVERSION_TIMEOUT', 'Converting this file took too long, so Simple stopped. The original file is unchanged.', { technical: `Stopped after ${timeoutMs} ms. ${detail}`.trim() }))
      } else if (code !== 0) {
        reject(new OfficeEngineError('CONVERSION_FAILED', "Simple couldn't convert this file. It may be damaged or protected. The original file is unchanged.", { technical: `Exit code ${code}${signal ? ` (${signal})` : ''}. ${detail}`.trim() }))
      } else resolve()
    })
  })
}

function preparedBytes(value) {
  const bytes = value && typeof value === 'object' && !ArrayBuffer.isView(value) && !(value instanceof ArrayBuffer) && value.bytes ? value.bytes : value
  if (!bytes || typeof bytes.byteLength !== 'number' || !bytes.byteLength) throw new TypeError('prepareInput must return non-empty bytes or { bytes }.')
  return Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes instanceof ArrayBuffer ? new Uint8Array(bytes) : bytes)
}

async function removeJobDirectory(directory) {
  // Only the uniquely allocated job folder under the temp folder is removed, never an input path.
  const resolved = path.resolve(directory)
  const relative = path.relative(path.resolve(os.tmpdir()), resolved)
  if (relative && !relative.startsWith('..') && !path.isAbsolute(relative) && path.basename(resolved).startsWith(JOB_PREFIX)) {
    await fs.rm(resolved, { recursive: true, force: true }).catch(() => {})
  }
}

async function executeOfficeConversion(request, executable, options) {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), JOB_PREFIX))
  const profile = path.join(directory, 'profile')
  const output = path.join(directory, 'output')
  const inputPath = path.join(directory, `document.${request.inputExtension}`)
  const run = options.run || runOfficeConverter
  try {
    await fs.mkdir(path.join(profile, 'user'), { recursive: true })
    await fs.mkdir(output)
    let bytes = request.bytes
    if (options.prepareInput) {
      try {
        bytes = preparedBytes(await options.prepareInput(Buffer.from(request.bytes), { inputExtension: request.inputExtension, outputExtension: request.outputExtension, filter: request.filter }))
      } catch (error) {
        throw new OfficeEngineError('CONVERSION_FAILED', "Simple couldn't prepare this file for conversion. The original file is unchanged.", { technical: `prepareInput (${options.policyId}) failed: ${error.message}` })
      }
    }
    await fs.writeFile(inputPath, bytes)
    // Every job has its own profile, so an open LibreOffice window cannot take
    // the request, macros never run and external links are never refreshed.
    await fs.writeFile(path.join(profile, 'user', 'registrymodifications.xcu'), PROFILE_REGISTRY)
    const common = [`-env:UserInstallation=${pathToFileURL(profile).href}`, '--headless', '--nologo', '--nodefault', '--nolockcheck', '--norestore']
    const target = `${request.outputExtension}${request.filter ? `:${request.filter}` : ''}`
    await run(executable, [...common, '--convert-to', target, '--outdir', output, inputPath], options.timeoutMs || DEFAULT_TIMEOUT_MS)
    let converted
    try { converted = await fs.readFile(path.join(output, `document.${request.outputExtension}`)) } catch (error) {
      throw new OfficeEngineError('CONVERSION_FAILED', "Simple couldn't convert this file. It may be damaged or protected. The original file is unchanged.", { technical: `The engine produced no ${request.outputExtension} file (${error.code || error.message}).` })
    }
    if (!converted.length || converted.length > MAX_OUTPUT_BYTES || !hasOutputSignature(converted, request.outputExtension)) {
      throw new OfficeEngineError('OUTPUT_INVALID', "The converted file was empty or damaged, so Simple didn't use it. The original file is unchanged.", { technical: `${converted.length} output bytes without a ${request.outputExtension} signature.` })
    }
    return converted
  } finally {
    await removeJobDirectory(directory)
  }
}

function defaultCacheDirectory(env = process.env) {
  return path.join(env.LOCALAPPDATA || os.tmpdir(), 'simple', 'conversion-cache', 'v1')
}

function cacheFiles(directory, key, extension) {
  const token = crypto.createHash('sha256').update(key).digest('hex')
  return { directory, data: path.join(directory, `${token}.${extension}`), metadata: path.join(directory, `${token}.json`) }
}

async function readDiskConversion(directory, key, extension) {
  const files = cacheFiles(directory, key, extension)
  try {
    const [metadataText, stat] = await Promise.all([fs.readFile(files.metadata, 'utf8'), fs.stat(files.data)])
    const metadata = JSON.parse(metadataText)
    if (!stat.isFile() || stat.size !== metadata.size || stat.size > MAX_DISK_CACHE_BYTES) return null
    const bytes = await fs.readFile(files.data)
    if (!hasOutputSignature(bytes, extension) || crypto.createHash('sha256').update(bytes).digest('hex') !== metadata.sha256) return null
    const now = new Date()
    void fs.utimes(files.data, now, now).catch(() => {})
    return bytes
  } catch { return null }
}

async function pruneDiskConversions(directory) {
  let names
  try { names = await fs.readdir(directory) } catch { return }
  const entries = []
  for (const name of names) {
    if (!/^[a-f0-9]{64}\.(docx?|xlsx?|pdf|odt|ods)$/.test(name)) continue
    try {
      const filePath = path.join(directory, name)
      const stat = await fs.stat(filePath)
      if (stat.isFile()) entries.push({ filePath, size: stat.size, touched: stat.mtimeMs })
    } catch {}
  }
  entries.sort((left, right) => right.touched - left.touched)
  let bytes = 0
  for (let index = 0; index < entries.length; index += 1) {
    bytes += entries[index].size
    if (index < MAX_DISK_CACHE_FILES && bytes <= MAX_DISK_CACHE_BYTES) continue
    // Remove only individually named cache files, never a computed directory tree.
    await fs.unlink(entries[index].filePath).catch(() => {})
    await fs.unlink(entries[index].filePath.replace(/\.[^.]+$/, '.json')).catch(() => {})
  }
}

async function storeDiskConversion(directory, key, extension, bytes) {
  if (bytes.length > MAX_DISK_CACHE_BYTES || !hasOutputSignature(bytes, extension)) return
  const files = cacheFiles(directory, key, extension)
  const suffix = `.${crypto.randomUUID()}.tmp`
  try {
    await fs.mkdir(files.directory, { recursive: true })
    await fs.writeFile(files.data + suffix, bytes)
    await fs.writeFile(files.metadata + suffix, JSON.stringify({ size: bytes.length, sha256: crypto.createHash('sha256').update(bytes).digest('hex') }))
    await fs.rename(files.data + suffix, files.data)
    await fs.rename(files.metadata + suffix, files.metadata)
    cachePrunePromise = cachePrunePromise.then(() => pruneDiskConversions(directory)).catch(() => {})
    await cachePrunePromise
  } catch {
    // A cache failure must never stop opening, saving or combining a document.
  } finally {
    await fs.unlink(files.data + suffix).catch(() => {})
    await fs.unlink(files.metadata + suffix).catch(() => {})
  }
}

function rememberConversion(key, entry, size) {
  entry.size = size
  memoryCacheBytes += size
  for (const [oldKey, old] of conversionCache) {
    if (memoryCacheBytes <= MAX_MEMORY_CACHE_BYTES && conversionCache.size <= MAX_MEMORY_CACHE_ENTRIES) break
    if (!old.size || oldKey === key) continue
    conversionCache.delete(oldKey)
    memoryCacheBytes -= old.size
  }
}

/**
 * Converts office bytes with the local engine.
 *
 * @param {object} input
 * @param {Uint8Array|ArrayBuffer|Buffer} input.bytes the source file's bytes (at most 256 MB)
 * @param {string} input.inputExtension doc, docx, xls, xlsx, odt, ods, rtf, ppt or pptx
 * @param {string} input.outputExtension docx, doc, pdf, xls, xlsx, odt or ods
 * @param {string} [input.filter] LibreOffice export filter, e.g. "writer_pdf_Export"
 * @param {object} [options]
 * @param {(bytes: Buffer, info: {inputExtension: string, outputExtension: string, filter: string}) => (Uint8Array|{bytes: Uint8Array}|Promise<Uint8Array|{bytes: Uint8Array}>)} [options.prepareInput]
 *   workspace-specific preparation of a private copy of the input (never the user's file)
 * @param {string} [options.policyId] names prepareInput and its version; required with prepareInput,
 *   part of the cache key so a changed preparation never reuses old output
 * @param {string} [options.executable] use this engine instead of probing
 * @param {(executable: string, args: string[], timeoutMs: number) => Promise<void>} [options.run] replaces the
 *   process runner (tests); without cacheDirectory such runs are never cached
 * @param {number} [options.timeoutMs=60000] stop a conversion that takes longer
 * @param {boolean} [options.cache=true] false skips the memory and disk caches
 * @param {string} [options.cacheDirectory] disk cache folder (default %LOCALAPPDATA%\simple\conversion-cache\v1)
 * @param {Record<string, string|undefined>} [options.env] environment for the probe (default process.env)
 * @param {string[]} [options.applicationRoots] companion-engine folders for the probe
 * @returns {Promise<Buffer>} the converted file, checked for the output format's signature
 * @throws {OfficeEngineError} NEEDS_OFFICE_ENGINE when no engine is usable (never mentions setup steps);
 *   UNSUPPORTED_CONVERSION, EMPTY_INPUT, TOO_LARGE, CONVERSION_TIMEOUT, CONVERSION_FAILED or OUTPUT_INVALID otherwise
 */
async function convertOfficeBytes(input, options = {}) {
  const request = validateConversionInput(input)
  if (options.prepareInput !== undefined && typeof options.prepareInput !== 'function') throw new TypeError('prepareInput must be a function.')
  if (options.prepareInput && (typeof options.policyId !== 'string' || !options.policyId.trim())) {
    throw new TypeError('convertOfficeBytes needs a policyId that names prepareInput, so cached conversions stay correct.')
  }
  const env = options.env || process.env
  if (isOfficeEngineForcedOff(env)) throw needsOfficeEngineError(request, 'forced-off')
  let executable = options.executable ? path.resolve(options.executable) : null
  if (!executable) {
    const status = await getOfficeEngineStatus({ env, applicationRoots: options.applicationRoots })
    if (!status.available) throw needsOfficeEngineError(request, status.reason)
    executable = status.path
  }
  const executableStat = await statOrNull(executable)
  if (!options.run && !(executableStat && executableStat.isFile())) throw needsOfficeEngineError(request, 'binary-missing')
  const useCache = options.cache !== false && (!options.run || Boolean(options.cacheDirectory))
  if (!useCache) return executeOfficeConversion(request, executable, options)

  const fingerprint = crypto.createHash('sha256').update(request.bytes).digest('hex')
  const policy = options.prepareInput ? `prepared:${options.policyId.trim()}` : (options.policyId ? `plain:${options.policyId.trim()}` : 'plain')
  const key = JSON.stringify([
    CONVERSION_POLICY_VERSION, policy, executable.toLowerCase(), executableStat ? executableStat.mtimeMs : 0,
    request.inputExtension, request.outputExtension, request.filter, fingerprint,
  ])
  const cacheDirectory = path.resolve(options.cacheDirectory || defaultCacheDirectory(env))
  const memoryKey = `${cacheDirectory.toLowerCase()}\n${key}`
  const existing = conversionCache.get(memoryKey)
  if (existing) return Buffer.from(await existing.promise)
  const entry = {
    size: 0,
    promise: (async () => {
      const disk = await readDiskConversion(cacheDirectory, key, request.outputExtension)
      if (disk) return disk
      const converted = await executeOfficeConversion(request, executable, options)
      await storeDiskConversion(cacheDirectory, key, request.outputExtension, converted)
      return converted
    })(),
  }
  conversionCache.set(memoryKey, entry)
  try {
    const result = await entry.promise
    if (conversionCache.get(memoryKey) === entry && !entry.size) rememberConversion(memoryKey, entry, result.length)
    return Buffer.from(result)
  } catch (error) {
    if (conversionCache.get(memoryKey) === entry) conversionCache.delete(memoryKey)
    throw error
  }
}

/**
 * Forgets conversions kept in memory (the disk cache stays). Useful after
 * memory pressure and in tests.
 */
function clearConversionMemoryCache() {
  conversionCache.clear()
  memoryCacheBytes = 0
}

/** Forgets cached probe results, so the next status call probes again. */
function clearOfficeEngineStatusCache() {
  statusCache.clear()
}

module.exports = {
  CONVERSION_POLICY_VERSION,
  FORMAT_LABELS,
  INPUT_EXTENSIONS,
  KNOWN_RUNTIME_HASHES,
  OUTPUT_EXTENSIONS,
  OfficeEngineError,
  clearConversionMemoryCache,
  clearOfficeEngineStatusCache,
  convertOfficeBytes,
  findOfficeConverter,
  getOfficeEngineStatus,
  hasOutputSignature,
  isOfficeEngineError,
  isOfficeEngineForcedOff,
  nativeAlternative,
  needsOfficeEngineError,
  officeRuntimeDirectory,
  parseJsonWithBom,
  readRuntimeManifest,
  runOfficeConverter,
  validateConversionInput,
}
