const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const { pathToFileURL } = require('node:url')
const { spawn, execFile } = require('node:child_process')
const crypto = require('node:crypto')
const { lockLegacyDateFields } = require('./legacy-fields.cjs')
const conversionCache = new Map()
let cachedBytes = 0
const CONVERSION_POLICY_VERSION = '4-locked-date-fields'
const MAX_DISK_CACHE_BYTES = 256 * 1024 * 1024
let cachePrunePromise = null

function conversionCacheDirectory() {
  return path.join(process.env.LOCALAPPDATA || os.tmpdir(), 'simple', 'conversion-cache', 'v1')
}

function cachePaths(key, extension) {
  const token = crypto.createHash('sha256').update(key).digest('hex')
  const directory = conversionCacheDirectory()
  return { directory, data: path.join(directory, `${token}.${extension}`), metadata: path.join(directory, `${token}.json`) }
}

function hasOutputSignature(bytes, extension) {
  if (!bytes || bytes.length < 8) return false
  if (['doc', 'xls'].includes(extension)) return bytes.subarray(0, 8).toString('hex') === 'd0cf11e0a1b11ae1'
  if (extension === 'pdf') return bytes.subarray(0, 5).toString('ascii') === '%PDF-'
  return bytes[0] === 0x50 && bytes[1] === 0x4b && bytes.length >= 30
}

async function readDiskConversion(key, extension) {
  const files = cachePaths(key, extension)
  try {
    const [metadata, stat] = await Promise.all([fs.readFile(files.metadata, 'utf8').then(JSON.parse), fs.stat(files.data)])
    if (!stat.isFile() || stat.size !== metadata.size || stat.size > MAX_DISK_CACHE_BYTES) return null
    const bytes = await fs.readFile(files.data)
    if (!hasOutputSignature(bytes, extension) || crypto.createHash('sha256').update(bytes).digest('hex') !== metadata.sha256) return null
    const now = new Date()
    void fs.utimes(files.data, now, now).catch(() => {})
    return bytes
  } catch { return null }
}

async function pruneDiskConversions() {
  const directory = conversionCacheDirectory()
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
    const entry = entries[index]
    bytes += entry.size
    if (index < 128 && bytes <= MAX_DISK_CACHE_BYTES) continue
    // Remove only individually named cache files, never a computed directory tree.
    await fs.unlink(entry.filePath).catch(() => {})
    await fs.unlink(entry.filePath.replace(/\.[^.]+$/, '.json')).catch(() => {})
  }
}

async function storeDiskConversion(key, extension, bytes) {
  if (bytes.length > MAX_DISK_CACHE_BYTES || !hasOutputSignature(bytes, extension)) return
  const files = cachePaths(key, extension)
  const suffix = `.${crypto.randomUUID()}.tmp`
  try {
    await fs.mkdir(files.directory, { recursive: true })
    await fs.writeFile(files.data + suffix, bytes)
    await fs.writeFile(files.metadata + suffix, JSON.stringify({ size: bytes.length, sha256: crypto.createHash('sha256').update(bytes).digest('hex') }))
    await fs.rename(files.data + suffix, files.data)
    await fs.rename(files.metadata + suffix, files.metadata)
    cachePrunePromise = (cachePrunePromise || Promise.resolve()).then(pruneDiskConversions).catch(() => {})
    await cachePrunePromise
  } catch { /* A cache failure must never stop opening or saving a document. */ }
  finally {
    await fs.unlink(files.data + suffix).catch(() => {})
    await fs.unlink(files.metadata + suffix).catch(() => {})
  }
}

async function findOfficeConverter() {
  const applicationRoots = [process.env.PORTABLE_EXECUTABLE_DIR, path.dirname(process.execPath), process.resourcesPath, path.resolve(__dirname, '..', '..')].filter(Boolean)
  const candidates = [
    process.env.SIMPLE_LIBREOFFICE_PATH,
    ...(process.env.LOCALAPPDATA ? [
      path.join(process.env.LOCALAPPDATA, 'simple', 'office-runtime', 'program', 'soffice.exe'),
      path.join(process.env.LOCALAPPDATA, 'simple', 'office-runtime', 'LibreOffice', 'program', 'soffice.exe'),
      path.join(process.env.LOCALAPPDATA, 'simple', 'office-runtime', 'Program Files', 'LibreOffice', 'program', 'soffice.exe'),
    ] : []),
    ...applicationRoots.flatMap((root) => [path.join(root, 'tools', 'libreoffice', 'program', 'soffice.exe'), path.join(root, 'LibreOffice', 'program', 'soffice.exe')]),
    ...[process.env.ProgramFiles, process.env['ProgramFiles(x86)']].filter(Boolean).map((root) => path.join(root, 'LibreOffice', 'program', 'soffice.exe')),
  ].filter(Boolean)
  for (const candidate of candidates) {
    try { if ((await fs.stat(candidate)).isFile()) return candidate } catch {}
  }
  return null
}

function runOfficeConverter(executable, args, timeoutMs = 60_000) {
  return new Promise((resolve, reject) => {
    const child = spawn(executable, args, { windowsHide: true, shell: false, stdio: ['ignore', 'pipe', 'pipe'] })
    let detail = ''
    let timedOut = false
    child.stdout.on('data', (data) => { detail = (detail + data.toString()).slice(-4000) })
    child.stderr.on('data', (data) => { detail = (detail + data.toString()).slice(-4000) })
    const timer = setTimeout(() => {
      timedOut = true
      if (process.platform === 'win32' && child.pid) execFile('taskkill.exe', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true }, () => {})
      else child.kill('SIGKILL')
    }, timeoutMs)
    child.once('error', (error) => { clearTimeout(timer); reject(error) })
    child.once('close', (code) => {
      clearTimeout(timer)
      if (timedOut) reject(new Error('Document conversion took too long. The original file is unchanged.'))
      else if (code !== 0) reject(new Error(`The document conversion engine could not open this file. ${detail.trim()}`))
      else resolve()
    })
  })
}

function validateConversionInput(input) {
  const inputExtension = String(input.inputExtension || '').replace(/^\./, '').toLowerCase()
  const outputExtension = String(input.outputExtension || '').replace(/^\./, '').toLowerCase()
  if (!['doc', 'docx', 'xls', 'xlsx', 'odt', 'ods', 'rtf', 'ppt', 'pptx'].includes(inputExtension)
      || !['docx', 'doc', 'pdf', 'xls', 'xlsx', 'odt', 'ods'].includes(outputExtension)) throw new Error('Unsupported document conversion.')
  if (inputExtension === outputExtension) throw new Error('Choose a different document output format.')
  const bytes = Buffer.from(input.bytes)
  if (!bytes.length || bytes.length > 256 * 1024 * 1024) throw new Error('The document is empty or too large to convert.')
  const filter = String(input.filter || '')
  if (filter.length > 120 || /[\r\n\u0000]/.test(filter)) throw new Error('Invalid document conversion filter.')
  return { bytes, inputExtension, outputExtension, filter }
}

async function executeOfficeConversion(input, options = {}) {
  const { bytes, inputExtension, outputExtension, filter } = validateConversionInput(input)
  const executable = options.executable || await findOfficeConverter()
  if (!executable) throw new Error('The document conversion engine is unavailable. Install LibreOffice and try again.')
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-convert-'))
  const profile = path.join(directory, 'profile')
  const output = path.join(directory, 'output')
  const inputPath = path.join(directory, `document.${inputExtension}`)
  const run = options.run || runOfficeConverter
  try {
    await fs.mkdir(path.join(profile, 'user'), { recursive: true })
    await fs.mkdir(output)
    await fs.writeFile(inputPath, inputExtension === 'doc' ? lockLegacyDateFields(bytes).bytes : bytes)
    // Every job has its own profile, so an open LibreOffice window cannot steal
    // the request. Disable macros and automatic external-document updates.
    await fs.writeFile(path.join(profile, 'user', 'registrymodifications.xcu'), `<?xml version="1.0" encoding="UTF-8"?><oor:items xmlns:oor="http://openoffice.org/2001/registry"><item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse"><value>3</value></prop></item><item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>0</value></prop></item><item oor:path="/org.openoffice.Office.Calc/Content/Update"><prop oor:name="Link" oor:op="fuse"><value>1</value></prop></item></oor:items>`)
    const common = [`-env:UserInstallation=${pathToFileURL(profile).href}`, '--headless', '--nologo', '--nodefault', '--nolockcheck', '--norestore']
    await run(executable, [...common, '--convert-to', `${outputExtension}${filter ? `:${filter}` : ''}`, '--outdir', output, inputPath])
    const converted = await fs.readFile(path.join(output, `document.${outputExtension}`))
    if (!converted.length || converted.length > 512 * 1024 * 1024) throw new Error('The converted document is empty or too large.')
    return converted
  } finally {
    // directory is the absolute, uniquely allocated job folder, never an input path.
    const relative = path.relative(path.resolve(os.tmpdir()), path.resolve(directory))
    if (relative && !relative.startsWith('..') && !path.isAbsolute(relative) && path.basename(directory).startsWith('simple-doc-convert-')) await fs.rm(directory, { recursive: true, force: true }).catch(() => {})
  }
}

async function convertOfficeBytes(input, options = {}) {
  input = validateConversionInput(input)
  if (options.run) return executeOfficeConversion(input, options)
  const executable = options.executable || await findOfficeConverter()
  if (!executable) return executeOfficeConversion(input, options)
  const fingerprint = crypto.createHash('sha256').update(Buffer.from(input.bytes)).digest('hex')
  const key = JSON.stringify([CONVERSION_POLICY_VERSION, executable, (await fs.stat(executable)).mtimeMs, input.inputExtension, input.outputExtension, input.filter, fingerprint])
  const existing = conversionCache.get(key)
  if (existing) return Buffer.from(await existing.promise)
  const entry = { size: 0, promise: (async () => {
    const disk = await readDiskConversion(key, input.outputExtension)
    if (disk) return disk
    const converted = await executeOfficeConversion(input, { ...options, executable })
    await storeDiskConversion(key, input.outputExtension, converted)
    return converted
  })() }
  conversionCache.set(key, entry)
  try {
    const result = await entry.promise
    entry.size = result.length
    cachedBytes += result.length
    for (const [oldKey, old] of conversionCache) {
      if (cachedBytes <= 64 * 1024 * 1024 && conversionCache.size <= 8) break
      if (!old.size) continue
      conversionCache.delete(oldKey)
      cachedBytes -= old.size
    }
    return Buffer.from(result)
  } catch (error) {
    conversionCache.delete(key)
    throw error
  }
}

module.exports = { findOfficeConverter, runOfficeConverter, convertOfficeBytes, hasOutputSignature, readDiskConversion, storeDiskConversion, CONVERSION_POLICY_VERSION }
