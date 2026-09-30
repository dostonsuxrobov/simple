const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const { pathToFileURL } = require('node:url')
const { spawn, execFile } = require('node:child_process')

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

async function convertOfficeBytes(input, options = {}) {
  const executable = options.executable || await findOfficeConverter()
  if (!executable) throw new Error('The document conversion engine is unavailable. Install LibreOffice and try again.')
  const inputExtension = String(input.inputExtension || '').replace(/^\./, '').toLowerCase()
  const outputExtension = String(input.outputExtension || '').replace(/^\./, '').toLowerCase()
  if (!['doc', 'docx', 'xls', 'xlsx', 'odt', 'ods', 'rtf', 'ppt', 'pptx'].includes(inputExtension)
      || !['docx', 'doc', 'pdf', 'xls', 'xlsx', 'odt', 'ods'].includes(outputExtension)) throw new Error('Unsupported document conversion.')
  if (inputExtension === outputExtension) throw new Error('Choose a different document output format.')
  const bytes = Buffer.from(input.bytes)
  if (!bytes.length || bytes.length > 256 * 1024 * 1024) throw new Error('The document is empty or too large to convert.')
  const filter = String(input.filter || '')
  if (filter.length > 120 || /[\r\n\u0000]/.test(filter)) throw new Error('Invalid document conversion filter.')
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-convert-'))
  const profile = path.join(directory, 'profile')
  const output = path.join(directory, 'output')
  const inputPath = path.join(directory, `document.${inputExtension}`)
  const run = options.run || runOfficeConverter
  try {
    await fs.mkdir(path.join(profile, 'user'), { recursive: true })
    await fs.mkdir(output)
    await fs.writeFile(inputPath, bytes)
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
    const resolvedDirectory = path.resolve(directory)
    const temporaryRoot = path.resolve(os.tmpdir()) + path.sep
    if (resolvedDirectory.startsWith(temporaryRoot) && path.basename(resolvedDirectory).startsWith('simple-doc-convert-')) {
      await fs.rm(resolvedDirectory, { recursive: true, force: true }).catch(() => {})
    }
  }
}

module.exports = { findOfficeConverter, runOfficeConverter, convertOfficeBytes }
