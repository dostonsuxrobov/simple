'use strict'

/**
 * Vendors the shared Save/Import/Export layer from simple/shared/ into the
 * workspaces listed in simple/shared/manifest.json, or checks that the
 * vendored copies are current.
 *
 *   node scripts/sync-shared.cjs                 write mode: copy, rewrite preload blocks, remove stale vendored copies
 *   node scripts/sync-shared.cjs --check         write nothing; exit 1 and name every drifted file
 *   --module=pdf[,calc]                          limit to these workspaces (each must be enabled)
 *   --target-root <dir>                          folder that holds the simple_*_source workspaces (default: repo root)
 *   --shared-root <dir>                          folder that holds manifest.json and the shared sources (default: simple/shared)
 *
 * Only workspaces with "enabled": true are touched or checked. Shared sources
 * listed in the manifest but not created yet are reported as planned and skipped.
 * Destination folders must be named simple-io (electron/simple-io,
 * src/simple-io), so a typo in the manifest can never point the sync at a
 * workspace's own sources. Inside them, only a file that is plainly a
 * vendored copy (it starts with the vendored header, or it is a data file
 * named like one the manifest vendors) is ever removed, one file at a time;
 * anything else is reported and left in place.
 * Exit codes: 0 current, 1 drift or wiring problems, 2 usage or manifest error.
 */

const fs = require('node:fs')
const path = require('node:path')

const ROOT = path.resolve(__dirname, '..')
const DEFAULT_SHARED_ROOT = path.join(ROOT, 'shared')
const DEFAULT_TARGET_ROOT = path.resolve(ROOT, '..')
const FIX_COMMAND = 'npm run sync:shared'
const HEADER_EXTENSIONS = new Set(['.cjs', '.js', '.mjs', '.ts'])
const KINDS = Object.freeze(['electron', 'renderer'])
/** The last folder of every destination: the sync owns only folders with this name. */
const DESTINATION_FOLDER = 'simple-io'
const BEGIN_PATTERN = /^[ \t]*\/\/ <simple-io-bridge(?: [^>]*)?>[ \t]*$/
const END_PATTERN = /^[ \t]*\/\/ <\/simple-io-bridge>[ \t]*$/

/** Error for a broken manifest, unknown option or unusable tree (exit code 2). */
class SyncSetupError extends Error {}

function toPosix(value) {
  return String(value).replaceAll('\\', '/')
}

function isSafeRelative(value) {
  if (typeof value !== 'string' || !value || path.isAbsolute(value) || /^[a-z]:/i.test(value)) return false
  return !toPosix(value).split('/').some((part) => part === '' || part === '.' || part === '..')
}

function readOptional(filePath) {
  try { return fs.readFileSync(filePath) } catch (error) {
    if (error.code === 'ENOENT' || error.code === 'ENOTDIR') return null
    throw error
  }
}

function sleepSync(milliseconds) {
  Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, milliseconds)
}

/**
 * Replaces a file through a sibling temp file and rename, retrying briefly
 * while an editor, indexer or antivirus scanner holds the target.
 * @param {string} filePath
 * @param {Buffer|string} bytes
 */
function writeFileReplacing(filePath, bytes) {
  fs.mkdirSync(path.dirname(filePath), { recursive: true })
  const temporary = `${filePath}.~simple-sync-${process.pid}.tmp`
  fs.writeFileSync(temporary, bytes)
  for (let attempt = 0; ; attempt += 1) {
    try {
      fs.renameSync(temporary, filePath)
      return
    } catch (error) {
      if (attempt >= 20 || !['EPERM', 'EBUSY', 'EACCES'].includes(error.code)) {
        try { fs.rmSync(temporary, { force: true }) } catch {}
        throw error
      }
      sleepSync(100)
    }
  }
}

/**
 * Reads and validates manifest.json from a shared root.
 * @param {string} [sharedRoot] folder holding manifest.json (default simple/shared)
 * @returns {object} the parsed manifest
 * @throws {SyncSetupError} when the manifest is missing or malformed
 */
function loadManifest(sharedRoot = DEFAULT_SHARED_ROOT) {
  const manifestPath = path.join(sharedRoot, 'manifest.json')
  let manifest
  try {
    manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8').replace(/^﻿/, ''))
  } catch (error) {
    throw new SyncSetupError(`Cannot read ${manifestPath}: ${error.message}`)
  }
  const fail = (message) => { throw new SyncSetupError(`${manifestPath}: ${message}`) }
  if (manifest.schemaVersion !== 1) fail('schemaVersion must be 1.')
  if (typeof manifest.header !== 'string' || !manifest.header.includes('{source}')) fail('header must be a string containing {source}.')
  for (const kind of KINDS) {
    const destination = manifest.destinations?.[kind]
    if (!isSafeRelative(destination)) fail(`destinations.${kind} must be a relative folder.`)
    // The sync removes stale copies from these folders, so they must be its own.
    if (path.posix.basename(toPosix(destination)).toLowerCase() !== DESTINATION_FOLDER) {
      fail(`destinations.${kind} must be a folder named ${DESTINATION_FOLDER} (for example ${kind === 'electron' ? 'electron' : 'src'}/${DESTINATION_FOLDER}), never a workspace's own source folder.`)
    }
  }
  if (toPosix(manifest.destinations.electron).toLowerCase() === toPosix(manifest.destinations.renderer).toLowerCase()) {
    fail('destinations.electron and destinations.renderer must be different folders.')
  }
  const block = manifest.preloadBlock
  if (!block || !isSafeRelative(block.source) || !isSafeRelative(block.file)) fail('preloadBlock needs relative source and file paths.')
  if (!BEGIN_PATTERN.test(block.begin || '') || !END_PATTERN.test(block.end || '')) fail('preloadBlock begin/end must be // <simple-io-bridge …> and // </simple-io-bridge>.')
  if (!manifest.workspaces || typeof manifest.workspaces !== 'object') fail('workspaces must be an object.')
  for (const [name, workspace] of Object.entries(manifest.workspaces)) {
    if (!isSafeRelative(workspace.folder)) fail(`workspaces.${name}.folder must be a relative folder.`)
    if (typeof workspace.enabled !== 'boolean') fail(`workspaces.${name}.enabled must be true or false.`)
    if (typeof workspace.preload !== 'boolean') fail(`workspaces.${name}.preload must be true or false.`)
    for (const kind of KINDS) {
      const list = workspace[kind]
      if (!Array.isArray(list)) fail(`workspaces.${name}.${kind} must be an array.`)
      const names = new Set()
      for (const source of list) {
        if (!isSafeRelative(source) || !toPosix(source).startsWith(`${kind}/`)) fail(`workspaces.${name}.${kind} entry "${source}" must be a path under ${kind}/.`)
        const base = path.posix.basename(toPosix(source)).toLowerCase()
        if (names.has(base)) fail(`workspaces.${name}.${kind} vendors two files named ${base}.`)
        names.add(base)
      }
    }
    for (const rule of workspace.wiring || []) {
      if (!rule || typeof rule.call !== 'string' || !rule.call || (rule.when !== undefined && !isSafeRelative(rule.when))) {
        fail(`workspaces.${name}.wiring entries need a "call" string and an optional relative "when" source.`)
      }
    }
  }
  for (const pair of manifest.sameBytes || []) {
    if (!Array.isArray(pair) || pair.length !== 2 || !pair.every(isSafeRelative)) fail('sameBytes entries must be pairs of relative paths.')
  }
  return manifest
}

function headerFor(manifest, source) {
  return manifest.header.replace('{source}', toPosix(source))
}

/** True for the first line of any file the sync vendored (the header with any source). */
function isVendoredHeader(manifest, line) {
  const [before, after] = manifest.header.split('{source}')
  return line.length > before.length + after.length && line.startsWith(before) && line.endsWith(after)
}

/**
 * True when a file left in a destination folder is plainly a copy the sync
 * vendored: a source file that starts with the vendored header, or a data
 * file (no header) named like one the manifest vendors for this kind.
 * Anything else may be someone's work and is never removed.
 */
function isVendoredCopy(manifest, kind, filePath, name) {
  if (HEADER_EXTENSIONS.has(path.extname(name).toLowerCase())) {
    let head = ''
    try {
      const handle = fs.openSync(filePath, 'r')
      try {
        const buffer = Buffer.alloc(1024)
        head = buffer.subarray(0, fs.readSync(handle, buffer, 0, buffer.length, 0)).toString('utf8')
      } finally { fs.closeSync(handle) }
    } catch { return false }
    return isVendoredHeader(manifest, head.replace(/^\uFEFF/, '').split(/\r?\n/, 1)[0])
  }
  const lower = name.toLowerCase()
  return Object.values(manifest.workspaces).some((workspace) => workspace[kind].some((source) => path.posix.basename(toPosix(source)).toLowerCase() === lower))
}

function detectEol(text) {
  return text.includes('\r\n') ? '\r\n' : '\n'
}

/**
 * Finds every bridge block in a preload file.
 * @param {string} text preload file contents
 * @returns {{blocks: Array<{start: number, end: number}>, malformed: string|null}}
 *   start/end are character offsets; end includes the end marker line's line break.
 */
function findBridgeBlocks(text) {
  const blocks = []
  let open = null
  let offset = 0
  let lineNumber = 0
  while (offset < text.length) {
    lineNumber += 1
    const newline = text.indexOf('\n', offset)
    const lineEnd = newline === -1 ? text.length : newline + 1
    const line = text.slice(offset, lineEnd).replace(/\r?\n$/, '')
    if (BEGIN_PATTERN.test(line)) {
      if (open) return { blocks, malformed: `a second begin marker at line ${lineNumber} before the block from line ${open.line} was closed` }
      open = { start: offset, line: lineNumber }
    } else if (END_PATTERN.test(line)) {
      if (!open) return { blocks, malformed: `an end marker at line ${lineNumber} without a begin marker` }
      blocks.push({ start: open.start, end: lineEnd })
      open = null
    }
    offset = lineEnd
  }
  if (open) return { blocks, malformed: `a begin marker at line ${open.line} without an end marker` }
  return { blocks, malformed: null }
}

/**
 * Builds the exact bridge block text for a preload file.
 * @param {object} manifest
 * @param {string} body contents of the shared preload source
 * @param {string} eol line ending used by the preload file
 */
function renderBridgeBlock(manifest, body, eol) {
  const lines = body.replace(/^﻿/, '').replace(/(\r?\n)+$/, '').split(/\r?\n/)
  return [manifest.preloadBlock.begin, ...lines, manifest.preloadBlock.end].join(eol) + eol
}

/**
 * Returns the preload text with exactly one current bridge block (or none
 * when `blockText` is null). The first existing block keeps its position;
 * a missing block is appended at the end of the file.
 */
function rewritePreload(text, blocks, blockText) {
  const eol = detectEol(text)
  if (!blocks.length) {
    if (blockText === null) return text
    const base = text.length && !text.endsWith('\n') ? text + eol : text
    return `${base}${base.trim() ? eol : ''}${blockText}`
  }
  let output = ''
  let cursor = 0
  blocks.forEach((block, index) => {
    let before = text.slice(cursor, block.start)
    // Drop the blank separator line in front of a block that is removed.
    // `before` always starts at a line start, so a leading newline is a blank line.
    if (index > 0 || blockText === null) before = before.replace(/(^|\n)[ \t]*\r?\n$/, '$1')
    output += before
    if (index === 0 && blockText !== null) output += blockText
    cursor = block.end
  })
  return output + text.slice(cursor)
}

function selectWorkspaces(manifest, modules) {
  const names = Object.keys(manifest.workspaces)
  if (!modules || !modules.length) {
    return {
      selected: names.filter((name) => manifest.workspaces[name].enabled),
      skipped: names.filter((name) => !manifest.workspaces[name].enabled),
    }
  }
  for (const name of modules) {
    if (!manifest.workspaces[name]) throw new SyncSetupError(`Unknown module "${name}". Known modules: ${names.join(', ')}.`)
    if (!manifest.workspaces[name].enabled) {
      throw new SyncSetupError(`${name} is not enabled in simple/shared/manifest.json. Set "enabled": true for it once the workspace is ready to receive the shared I/O layer.`)
    }
  }
  return { selected: [...new Set(modules)], skipped: [] }
}

/**
 * Vendors shared files into enabled workspaces, or checks them.
 *
 * @param {object} [options]
 * @param {boolean} [options.check=false] write nothing; report drift instead
 * @param {string[]} [options.modules] workspace names to limit to (each must be enabled)
 * @param {string} [options.sharedRoot] folder holding manifest.json and shared sources
 * @param {string} [options.targetRoot] folder holding the workspace folders
 * @returns {{
 *   ok: boolean, check: boolean, workspaces: string[], skipped: string[],
 *   changes: string[], drift: Array<{folder: string, file: string, reason: string}>,
 *   planned: string[], problems: string[], fixCommand: string, message: string
 * }} `ok` is false when drift remains (check mode) or any problem was found.
 * @throws {SyncSetupError} for a malformed manifest, unknown module or missing workspace file
 */
function syncShared(options = {}) {
  const check = Boolean(options.check)
  const sharedRoot = path.resolve(options.sharedRoot || DEFAULT_SHARED_ROOT)
  const targetRoot = path.resolve(options.targetRoot || DEFAULT_TARGET_ROOT)
  const manifest = loadManifest(sharedRoot)
  const { selected, skipped } = selectWorkspaces(manifest, options.modules)
  const customRoots = path.resolve(sharedRoot) !== DEFAULT_SHARED_ROOT || targetRoot !== DEFAULT_TARGET_ROOT
  const fixCommand = customRoots
    ? `node "${path.join(ROOT, 'scripts', 'sync-shared.cjs')}" --shared-root "${sharedRoot}" --target-root "${targetRoot}"`
    : FIX_COMMAND
  const changes = []
  const drift = []
  const planned = new Set()
  const problems = []
  const sourceCache = new Map()

  // Reads one shared source once; null while it is only planned.
  function source(relative) {
    if (sourceCache.has(relative)) return sourceCache.get(relative)
    const bytes = readOptional(path.join(sharedRoot, relative))
    if (bytes === null) planned.add(relative)
    else if (HEADER_EXTENSIONS.has(path.extname(relative).toLowerCase())) {
      const firstLine = bytes.toString('utf8').replace(/^﻿/, '').split(/\r?\n/, 1)[0]
      const expected = headerFor(manifest, relative)
      if (firstLine !== expected) problems.push(`simple/shared/${relative} must begin with this line:\n  ${expected}`)
    }
    sourceCache.set(relative, bytes)
    return bytes
  }

  // Validate every listed source (headers, planned status) even while no
  // workspace is enabled, so a bad shared file fails before integration.
  for (const workspace of Object.values(manifest.workspaces)) {
    for (const kind of KINDS) for (const relative of workspace[kind]) source(relative)
  }
  source(manifest.preloadBlock.source)

  for (const [left, right] of manifest.sameBytes || []) {
    const a = source(left)
    const b = source(right)
    if (a && b && !a.equals(b)) problems.push(`simple/shared/${right} must be byte-identical to simple/shared/${left}.`)
  }

  for (const name of selected) {
    const workspace = manifest.workspaces[name]
    const workspaceRoot = path.join(targetRoot, workspace.folder)
    if (!fs.existsSync(workspaceRoot)) throw new SyncSetupError(`Workspace folder for ${name} is missing: ${workspaceRoot}`)

    for (const kind of KINDS) {
      const destination = manifest.destinations[kind]
      const destinationRoot = path.join(workspaceRoot, destination)
      const expected = new Map()
      for (const relative of workspace[kind]) {
        const bytes = source(relative)
        if (bytes !== null) expected.set(path.posix.basename(toPosix(relative)).toLowerCase(), { relative, bytes })
      }
      for (const { relative, bytes } of expected.values()) {
        const file = `${destination}/${path.posix.basename(toPosix(relative))}`
        const current = readOptional(path.join(workspaceRoot, file))
        if (current !== null && current.equals(bytes)) continue
        const reason = current === null ? `missing; vendor simple/shared/${relative}` : `differs from simple/shared/${relative}`
        drift.push({ folder: workspace.folder, file, reason })
        if (!check) {
          writeFileReplacing(path.join(workspaceRoot, file), bytes)
          changes.push(`${current === null ? 'Added' : 'Updated'} ${workspace.folder}/${file}`)
        }
      }
      let entries = []
      try { entries = fs.readdirSync(destinationRoot, { withFileTypes: true }) } catch (error) {
        if (error.code !== 'ENOENT') throw error
      }
      for (const entry of entries) {
        if (expected.has(entry.name.toLowerCase())) continue
        const file = `${destination}/${entry.name}`
        const entryPath = path.join(destinationRoot, entry.name)
        if (!entry.isFile() || !isVendoredCopy(manifest, kind, entryPath, entry.name)) {
          problems.push(`${workspace.folder}/${file} is not a copy vendored from simple/shared, so it was left in place. Move it out of ${workspace.folder}/${destination}, which holds only vendored files.`)
          continue
        }
        drift.push({ folder: workspace.folder, file, reason: 'not listed in simple/shared/manifest.json or its shared source is gone' })
        if (!check) {
          fs.unlinkSync(entryPath)
          changes.push(`Removed ${workspace.folder}/${file}`)
        }
      }
    }

    if (workspace.preload) {
      const preloadFile = manifest.preloadBlock.file
      const preloadPath = path.join(workspaceRoot, preloadFile)
      const preloadBytes = readOptional(preloadPath)
      if (preloadBytes === null) throw new SyncSetupError(`${workspace.folder}/${preloadFile} is missing.`)
      const text = preloadBytes.toString('utf8')
      const { blocks, malformed } = findBridgeBlocks(text)
      if (malformed) {
        // The extent of a half-marked block is unknown, so never rewrite it automatically.
        problems.push(`${workspace.folder}/${preloadFile} has ${malformed}. Fix the // <simple-io-bridge> markers by hand, then run ${fixCommand}.`)
      } else {
        const body = source(manifest.preloadBlock.source)
        const blockText = body === null ? null : renderBridgeBlock(manifest, body.toString('utf8'), detectEol(text))
        let reason = null
        if (blocks.length > 1) reason = `bridge block appears ${blocks.length} times; it must appear once`
        else if (blockText === null && blocks.length) reason = `bridge block present but simple/shared/${manifest.preloadBlock.source} does not exist`
        else if (blockText !== null && !blocks.length) reason = 'bridge block missing'
        else if (blockText !== null && text.slice(blocks[0].start, blocks[0].end) !== blockText) reason = `bridge block differs from simple/shared/${manifest.preloadBlock.source}`
        if (reason) {
          drift.push({ folder: workspace.folder, file: preloadFile, reason })
          if (!check) {
            writeFileReplacing(preloadPath, rewritePreload(text, blocks, blockText))
            changes.push(`Rewrote the bridge block in ${workspace.folder}/${preloadFile}`)
          }
        }
      }
    }
  }

  problems.push(...wiringProblems(manifest, selected, targetRoot, sharedRoot))
  const unresolved = check ? drift : []
  const lines = []
  for (const item of unresolved) {
    lines.push(`Shared I/O code is out of date in ${item.folder} (${item.file}: ${item.reason}). Run: ${fixCommand}`)
  }
  lines.push(...problems)
  return {
    ok: !unresolved.length && !problems.length,
    check,
    workspaces: selected,
    skipped,
    changes,
    drift,
    planned: [...planned].sort(),
    problems,
    fixCommand,
    message: lines.join('\n'),
  }
}

/**
 * Lists wiring problems for enabled workspaces: each `wiring[].call` must
 * appear in the workspace's electron/main.cjs once its `when` shared source exists.
 * @param {object} manifest
 * @param {string[]} workspaceNames
 * @param {string} targetRoot
 * @param {string} sharedRoot
 * @returns {string[]}
 */
function wiringProblems(manifest, workspaceNames, targetRoot, sharedRoot) {
  const problems = []
  for (const name of workspaceNames) {
    const workspace = manifest.workspaces[name]
    const rules = (workspace.wiring || []).filter((rule) => !rule.when || fs.existsSync(path.join(sharedRoot, rule.when)))
    if (!rules.length) continue
    const mainPath = path.join(targetRoot, workspace.folder, 'electron', 'main.cjs')
    const main = readOptional(mainPath)
    for (const rule of rules) {
      if (main === null || !main.toString('utf8').includes(rule.call)) {
        problems.push(`${workspace.folder}/electron/main.cjs must call ${rule.call} because the workspace is enabled in simple/shared/manifest.json${rule.when ? ` and simple/shared/${rule.when} exists` : ''}.`)
      }
    }
  }
  return problems
}

function parseArguments(argv) {
  const options = { check: false, modules: [] }
  for (let index = 0; index < argv.length; index += 1) {
    const argument = argv[index]
    const valued = (flag) => {
      if (argument === flag) {
        const value = argv[index + 1]
        if (!value || value.startsWith('--')) throw new SyncSetupError(`${flag} needs a value.`)
        index += 1
        return value
      }
      if (argument.startsWith(`${flag}=`)) return argument.slice(flag.length + 1)
      return undefined
    }
    let value
    if (argument === '--check') options.check = true
    else if (argument === '--help' || argument === '-h') options.help = true
    else if ((value = valued('--module')) !== undefined) options.modules.push(...value.split(',').map((item) => item.trim()).filter(Boolean))
    else if ((value = valued('--target-root')) !== undefined) options.targetRoot = value
    else if ((value = valued('--shared-root')) !== undefined) options.sharedRoot = value
    else throw new SyncSetupError(`Unknown option: ${argument}`)
  }
  return options
}

/**
 * Command-line entry point.
 * @param {string[]} argv arguments after the script name
 * @returns {number} process exit code
 */
function main(argv) {
  let options
  try {
    options = parseArguments(argv)
    if (options.help) {
      console.log('Usage: node scripts/sync-shared.cjs [--check] [--module=<name>[,<name>]] [--target-root <dir>] [--shared-root <dir>]')
      return 0
    }
    const result = syncShared(options)
    const scope = result.workspaces.length ? result.workspaces.join(', ') : 'no enabled workspaces'
    if (!result.ok) {
      console.error(result.message)
      return 1
    }
    for (const change of result.changes) console.log(change)
    const verb = result.check ? 'is current' : (result.changes.length ? `updated (${result.changes.length} change${result.changes.length === 1 ? '' : 's'})` : 'was already current')
    console.log(`Shared I/O code ${verb} for ${scope}.`)
    if (result.skipped.length) console.log(`Not enabled yet: ${result.skipped.join(', ')}.`)
    if (result.planned.length) console.log(`${result.planned.length} planned shared file${result.planned.length === 1 ? ' is' : 's are'} not created yet.`)
    return 0
  } catch (error) {
    if (error instanceof SyncSetupError) {
      console.error(error.message)
      return 2
    }
    console.error(error.stack || error)
    return 2
  }
}

if (require.main === module) process.exitCode = main(process.argv.slice(2))

module.exports = {
  DESTINATION_FOLDER,
  FIX_COMMAND,
  SyncSetupError,
  findBridgeBlocks,
  loadManifest,
  main,
  renderBridgeBlock,
  syncShared,
  wiringProblems,
}
