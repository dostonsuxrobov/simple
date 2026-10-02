'use strict'

/**
 * Static LOCAL-ONLY guard for the app code Simple ships from simple/:
 * simple/shared/** (the shared Save/Import/Export layer, vendored into every
 * workspace), simple/launcher/** and simple/electron/** (bootstrap and
 * routing). Tests, scripts and build output are not app code and are skipped.
 *
 * Everything Simple does stays on this PC, and a file handed from one file
 * type to another stays inside Simple. The guard fails when app code:
 *   - requires a network module (http, https, http2, net, tls, dgram, dns);
 *   - calls fetch() or x.fetch() with anything but a data:, blob: or file:
 *     literal, loads a page with loadURL() from anything but a file:/data:
 *     literal or pathToFileURL(), or uses net.request, downloadURL,
 *     autoUpdater, WebSocket, EventSource or sendBeacon;
 *   - starts a download program or command (curl, wget, bitsadmin, certutil,
 *     PowerShell, Invoke-WebRequest, …), or any program by a bare name such as
 *     'reg.exe' (Windows looks in the current folder first: use the absolute
 *     System32 path);
 *   - contains a remote http(s) address (XML namespace URIs, which are names
 *     and never fetched, and the local development server are allowed);
 *   - names a cloud storage service, in code, comments or text. The names are
 *     kept out of this source: scripts/local-only-names.json holds SHA-256
 *     hashes, compared with the words of each file (tests and scripts are
 *     checked for this rule too);
 *   - hands a document to another program: shell.openPath, shell.openItem, or
 *     shell.openExternal with anything but a Windows Settings page (ms-settings:).
 *     "Show in folder" (shell.showItemInFolder) is fine;
 *   - shows user-facing text (io-catalog.json, launcher pages, and every
 *     string literal in app .cjs/.js/.mjs/.ts files, such as error messages)
 *     that tells the user to download or install something.
 *
 *   node scripts/local-only-guard.cjs     prints problems; exit 1 when there are any
 *
 * verify.cjs (npm test) runs the same check.
 */

const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..')
/** Folders of simple/ that hold shipped app code. */
const APP_FOLDERS = Object.freeze(['shared', 'launcher', 'electron'])
/** Folders of simple/ checked only for service names (tests and tools are not shipped, but must not name them either). */
const NAME_ONLY_FOLDERS = Object.freeze(['tests', 'scripts'])
const TEXT_EXTENSIONS = new Set(['.cjs', '.js', '.mjs', '.ts', '.tsx', '.json', '.html', '.htm', '.css', '.md', '.ps1'])
const CODE_EXTENSIONS = new Set(['.cjs', '.js', '.mjs', '.ts', '.tsx', '.html', '.htm'])
const SKIPPED_FOLDERS = new Set(['node_modules', 'tests', 'test', '__tests__', 'fixtures', 'dist', 'release', '.stage'])

const NETWORK_MODULES = ['http', 'https', 'http2', 'net', 'tls', 'dgram', 'dns']
const NETWORK_MODULE_PATTERN = new RegExp(
  String.raw`(?:\brequire\s*\(\s*|\bimport\s*\(\s*|\bfrom\s+|\bimport\s+)(['"\`])(?:node:)?(${NETWORK_MODULES.join('|')})\1`,
  'g',
)
/** Remote-transfer APIs that have no local use in Simple. */
const NETWORK_CALLS = Object.freeze([
  { pattern: /\bnet\s*\.\s*(?:request|fetch)\s*\(/g, what: 'net.request/net.fetch (network request)' },
  { pattern: /\bdownloadURL\s*\(/g, what: 'downloadURL (download)' },
  { pattern: /\bautoUpdater\b/g, what: 'autoUpdater (update check)' },
  { pattern: /\bnew\s+WebSocket\s*\(/g, what: 'WebSocket (network connection)' },
  { pattern: /\bnew\s+EventSource\s*\(/g, what: 'EventSource (network connection)' },
  { pattern: /\bsendBeacon\s*\(/g, what: 'sendBeacon (telemetry)' },
  { pattern: /\bnew\s+XMLHttpRequest\s*\(/g, what: 'XMLHttpRequest (network request)' },
])
/** fetch( and x.fetch( whose first argument is a local data:, blob: or file: literal are fine. */
const FETCH_PATTERN = /(?<![\w$])fetch\s*\(\s*([^)]{0,200})/g
const LOCAL_FETCH_ARGUMENT = /^(['"`])(?:data|blob|file):/i
/** loadURL( is fine only for a file:/data: literal or a pathToFileURL() address; pages load with loadFile. */
const LOAD_URL_PATTERN = /\bloadURL\s*\(\s*([^)]{0,200})/g
const LOCAL_LOAD_ARGUMENT = /^(?:(['"`])(?:file|data):|(?:url\.)?pathToFileURL\s*\()/i
/** Starting a program: the first argument is checked. */
const SPAWN_PATTERN = /(?<![\w$])(?:spawn|spawnSync|execFile|execFileSync|execSync|fork)\s*\(\s*([^,)]{0,200})/g
/** Programs and commands that transfer files over the network. */
const NETWORK_PROGRAMS = /\b(?:curl|wget|bitsadmin|certutil|powershell|pwsh|ftp|tftp)(?:\.exe)?\b/i
const NETWORK_COMMANDS = /\b(?:Invoke-WebRequest|Invoke-RestMethod|Start-BitsTransfer|Net\.WebClient|DownloadFile|DownloadString)\b/gi
/** Hashes of service names that must never appear (see local-only-names.json). */
const SERVICE_NAME_HASHES = new Set(JSON.parse(fs.readFileSync(path.join(__dirname, 'local-only-names.json'), 'utf8')).hashes)
/** Host names allowed in http(s) literals: XML namespaces (never fetched) and the local development server. */
const ALLOWED_URL_HOSTS = new Set([
  'localhost', '127.0.0.1', '[::1]',
  'openoffice.org', 'www.openoffice.org', 'schemas.openxmlformats.org', 'schemas.microsoft.com', 'www.w3.org',
  'purl.org', 'ns.adobe.com', 'www.ecma-international.org', 'docs.oasis-open.org', 'urn.oasis-open.org',
  'xml.org', 'www.idpf.org', 'www.opengis.net', 'www.loc.gov', 'iptc.org', 'cipa.jp',
])
const URL_PATTERN = /\bhttps?:\/\/([^\s'"`<>()\\/]+)/gi
/** Opening a document in another program. */
const HANDOFF_CALLS = Object.freeze([
  { pattern: /\bshell\s*\.\s*openPath\s*\(/g, what: 'shell.openPath opens the file in another program; use openInSimple' },
  { pattern: /\bshell\s*\.\s*openItem\s*\(/g, what: 'shell.openItem opens the file in another program; use openInSimple' },
])
const OPEN_EXTERNAL_PATTERN = /\bopenExternal\s*\(\s*([^)]{0,200})/g
const SETTINGS_PAGE_ARGUMENT = /^(['"`])ms-settings:[a-z0-9-]*\1\s*$/i
/** User-facing words that ask the user to fetch or set up software. */
const SETUP_WORDING = /\b(?:download|downloads|downloading|install|installs|installing|reinstall|update now)\b/i

function toPosix(value) {
  return String(value).replaceAll('\\', '/')
}

/**
 * Replaces comments with spaces (keeping line numbers and string contents),
 * so code rules never match prose. Strings and template literals are kept.
 * @param {string} text JavaScript or TypeScript source
 * @returns {string}
 */
function stripComments(text) {
  let output = ''
  let index = 0
  let quote = null
  while (index < text.length) {
    const char = text[index]
    const next = text[index + 1]
    if (quote) {
      output += char
      if (char === '\\') {
        output += next ?? ''
        index += 2
        continue
      }
      if (char === quote) quote = null
      index += 1
      continue
    }
    if (char === '"' || char === "'" || char === '`') {
      quote = char
      output += char
      index += 1
      continue
    }
    if (char === '/' && next === '/') {
      while (index < text.length && text[index] !== '\n') { output += ' '; index += 1 }
      continue
    }
    if (char === '/' && next === '*') {
      output += '  '
      index += 2
      while (index < text.length && !(text[index] === '*' && text[index + 1] === '/')) {
        output += text[index] === '\n' ? '\n' : ' '
        index += 1
      }
      output += '  '
      index += 2
      continue
    }
    output += char
    index += 1
  }
  return output
}

/** The quoted string literals of a code text (after stripComments), with their offsets. */
function stringLiterals(code) {
  const literals = []
  const pattern = /(['"`])((?:\\.|(?!\1)[^\\])*)\1/g
  for (const match of code.matchAll(pattern)) literals.push({ index: match.index, value: match[2] })
  return literals
}

function hashName(value) {
  return crypto.createHash('sha256').update(value).digest('hex')
}

/**
 * Finds names of cloud storage services in any text: single words, camelCase
 * parts and pairs of words written together or with a space, hyphen,
 * underscore or dot between them are compared (lower-cased) with the hashes
 * in local-only-names.json.
 * @param {string} text
 * @returns {Array<{index: number, text: string}>} each match's offset and the text it covers
 */
function findServiceNames(text) {
  const words = []
  for (const match of String(text).matchAll(/[A-Za-z0-9]+/g)) {
    let offset = match.index
    for (const part of match[0].split(/(?<=[a-z0-9])(?=[A-Z])/)) {
      words.push({ word: part.toLowerCase(), start: offset, end: offset + part.length })
      offset += part.length
    }
  }
  const found = []
  for (let index = 0; index < words.length; index += 1) {
    const current = words[index]
    const candidates = [{ value: current.word, end: current.end }]
    const next = words[index + 1]
    if (next) {
      const gap = text.slice(current.end, next.start)
      if (/^[\s_-]{0,3}$/.test(gap)) candidates.push({ value: current.word + next.word, end: next.end })
      else if (gap === '.') candidates.push({ value: `${current.word}.${next.word}`, end: next.end })
    }
    for (const candidate of candidates) {
      if (SERVICE_NAME_HASHES.has(hashName(candidate.value))) {
        found.push({ index: current.start, text: text.slice(current.start, candidate.end) })
        break
      }
    }
  }
  return found
}

function lineOf(text, offset) {
  let line = 1
  for (let index = 0; index < offset && index < text.length; index += 1) if (text.charCodeAt(index) === 10) line += 1
  return line
}

function urlHostAllowed(host) {
  const name = host.toLowerCase().replace(/:\d+$/, '')
  if (ALLOWED_URL_HOSTS.has(name)) return true
  return [...ALLOWED_URL_HOSTS].some((allowed) => allowed.includes('.') && name.endsWith(`.${allowed}`))
}

/** Every user-facing leaf string of a catalog, with its dotted key. */
function catalogStrings(value, prefix = '') {
  if (typeof value === 'string') return [[prefix, value]]
  if (!value || typeof value !== 'object') return []
  return Object.entries(value).flatMap(([key, child]) => catalogStrings(child, prefix ? `${prefix}.${key}` : key))
}

/**
 * Checks one file's text against the local-only rules.
 * @param {string} relative path shown in messages, e.g. "shared/electron/io-ipc.cjs"
 * @param {string} text file contents
 * @returns {string[]} problems, each "<file>:<line>: <what>"
 */
function checkText(relative, text) {
  const problems = []
  const extension = path.extname(relative).toLowerCase()
  const report = (offset, what) => problems.push(`${relative}:${lineOf(text, offset)}: ${what}`)

  // Cloud service names are not allowed anywhere, comments included.
  for (const match of findServiceNames(text)) {
    report(match.index, `names a cloud storage service ("${match.text}"); describe it as "another program is using this file"`)
  }

  if (CODE_EXTENSIONS.has(extension)) {
    const code = stripComments(text)
    for (const match of code.matchAll(NETWORK_MODULE_PATTERN)) report(match.index, `loads the network module "${match[2]}"`)
    for (const { pattern, what } of NETWORK_CALLS) for (const match of code.matchAll(pattern)) report(match.index, `uses ${what}`)
    for (const match of code.matchAll(FETCH_PATTERN)) {
      if (!LOCAL_FETCH_ARGUMENT.test(match[1].trim())) report(match.index, 'calls fetch() with something other than a local data:, blob: or file: address')
    }
    for (const match of code.matchAll(LOAD_URL_PATTERN)) {
      if (!LOCAL_LOAD_ARGUMENT.test(match[1].trim())) report(match.index, 'loads a page with loadURL() from something other than a local file:/data: address; use loadFile')
    }
    for (const match of code.matchAll(SPAWN_PATTERN)) {
      const argument = match[1].trim()
      if (NETWORK_PROGRAMS.test(argument)) {
        report(match.index, `starts a program that can download ("${argument.slice(0, 60)}")`)
        continue
      }
      const literal = /^(['"`])([^'"`]*)\1$/.exec(argument)
      if (literal && !path.win32.isAbsolute(literal[2]) && !path.posix.isAbsolute(literal[2])) {
        report(match.index, `starts the program "${literal[2]}" by a bare name, which Windows looks up in the current folder first; use its absolute path`)
      }
    }
    for (const match of code.matchAll(NETWORK_COMMANDS)) report(match.index, `runs the download command "${match[0]}"`)
    for (const { pattern, what } of HANDOFF_CALLS) for (const match of code.matchAll(pattern)) report(match.index, what)
    for (const match of code.matchAll(OPEN_EXTERNAL_PATTERN)) {
      if (!SETTINGS_PAGE_ARGUMENT.test(match[1].trim())) report(match.index, 'shell.openExternal hands something to another program; documents open in Simple through openInSimple')
    }
    for (const match of code.matchAll(URL_PATTERN)) {
      if (!urlHostAllowed(match[1])) report(match.index, `contains the remote address "${match[0]}"`)
    }
  }

  if (/(^|\/)io-catalog\.json$/.test(relative)) {
    let catalog = null
    try { catalog = JSON.parse(text.replace(/^﻿/, '')) } catch (error) {
      problems.push(`${relative}: is not valid JSON (${error.message})`)
    }
    for (const [key, value] of catalogStrings(catalog)) {
      if (SETUP_WORDING.test(value)) problems.push(`${relative}: "${key}" asks the user to download or install something: "${value}"`)
    }
  }
  if (/^launcher\/.*\.html?$/.test(relative)) {
    const visible = text.replace(/<script[\s\S]*?<\/script>|<style[\s\S]*?<\/style>|<[^>]+>/g, ' ')
    const match = SETUP_WORDING.exec(visible)
    if (match) problems.push(`${relative}: user-facing text asks the user to download or install something ("${match[0]}")`)
  }
  if (CODE_EXTENSIONS.has(extension) && !/\.html?$/.test(extension)) {
    // Error messages and other strings may reach the user, wherever they are written.
    // A literal that is only a name (an event such as 'will-download', a key,
    // a path or channel) is not text shown to the user.
    for (const literal of stringLiterals(stripComments(text))) {
      if (/^[\w.:/\\-]*$/.test(literal.value.trim())) continue
      const match = SETUP_WORDING.exec(literal.value)
      if (match) report(literal.index, `a message asks the user to download or install something ("${match[0]}")`)
    }
  }
  return problems
}

/**
 * Checks one file's text only for service names (tests and tools).
 * @param {string} relative path shown in messages
 * @param {string} text file contents
 * @returns {string[]} problems, each "<file>:<line>: <what>"
 */
function checkNamesOnly(relative, text) {
  return findServiceNames(text).map((match) => `${relative}:${lineOf(text, match.index)}: names a cloud storage service ("${match.text}"); describe it as "another program is using this file"`)
}

function filesUnder(directory, relative = '', options = {}) {
  let entries = []
  try { entries = fs.readdirSync(path.join(directory, relative), { withFileTypes: true }) } catch (error) {
    if (error.code === 'ENOENT') return []
    throw error
  }
  const skipped = options.includeTests ? new Set(['node_modules', 'dist', 'release', '.stage']) : SKIPPED_FOLDERS
  const output = []
  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
    const child = path.join(relative, entry.name)
    if (entry.isDirectory()) {
      if (!skipped.has(entry.name)) output.push(...filesUnder(directory, child, options))
    } else if (entry.isFile() && TEXT_EXTENSIONS.has(path.extname(entry.name).toLowerCase()) && (options.includeTests || !/\.test\.[a-z]+$/i.test(entry.name))) {
      output.push(child)
    }
  }
  return output
}

/**
 * Runs the LOCAL-ONLY guard over the shipped app folders of simple/.
 * @param {object} [options]
 * @param {string} [options.root] the simple/ folder (tests point it at a fixture tree)
 * @param {readonly string[]} [options.folders] folders under root to check (default shared, launcher, electron)
 * @param {readonly string[]} [options.nameFolders] folders checked only for service names, test files
 *   included (default tests, scripts)
 * @returns {{ok: boolean, files: number, problems: string[], message: string}}
 */
function checkLocalOnly(options = {}) {
  const root = path.resolve(options.root || ROOT)
  const folders = options.folders || APP_FOLDERS
  const nameFolders = options.nameFolders || NAME_ONLY_FOLDERS
  const problems = []
  let files = 0
  for (const folder of folders) {
    for (const relative of filesUnder(path.join(root, folder))) {
      files += 1
      const shown = toPosix(path.join(folder, relative))
      problems.push(...checkText(shown, fs.readFileSync(path.join(root, folder, relative), 'utf8')))
    }
  }
  for (const folder of nameFolders) {
    for (const relative of filesUnder(path.join(root, folder), '', { includeTests: true })) {
      files += 1
      const shown = toPosix(path.join(folder, relative))
      problems.push(...checkNamesOnly(shown, fs.readFileSync(path.join(root, folder, relative), 'utf8')))
    }
  }
  const message = problems.length
    ? `Simple's app code must stay local and keep documents inside Simple (simple/scripts/local-only-guard.cjs):\n  ${problems.join('\n  ')}`
    : ''
  return { ok: !problems.length, files, problems, message }
}

if (require.main === module) {
  const result = checkLocalOnly()
  if (result.ok) console.log(`Local-only guard: ${result.files} app files checked, no problems.`)
  else console.error(result.message)
  process.exitCode = result.ok ? 0 : 1
}

module.exports = { APP_FOLDERS, NAME_ONLY_FOLDERS, checkLocalOnly, checkNamesOnly, checkText, findServiceNames, stripComments }
