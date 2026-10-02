// Vendored from simple/shared/electron/formats.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// The format registry (formats.json) and everything derived from it:
// extension routing, Windows file associations, Open dialog filters, Export
// lists, the no-engine policy lookups, and content sniffing. Only Node
// built-ins and the sibling text codec are used, so this file runs unbundled
// in the launcher and bundled in every workspace main process.

const fs = require('node:fs')
const path = require('node:path')
const zlib = require('node:zlib')
const { decodeText, detectUtf16 } = require('./text-codec.cjs')

/** Bytes read from each end of a file for sniffing and content routing. */
const SNIFF_BYTES = 64 * 1024
const OPERATIONS = Object.freeze(['open', 'save', 'export'])
const ENGINE_MODES = Object.freeze(['none', 'upgrade', 'required'])
const ALL_FILES_FILTER = Object.freeze({ name: 'All files', extensions: Object.freeze(['*']) })

const CFB_SIGNATURE = Buffer.from('d0cf11e0a1b11ae1', 'hex')
const PNG_SIGNATURE = Buffer.from('89504e470d0a1a0a', 'hex')
const ZIP_LOCAL = 0x04034b50
const ZIP_CENTRAL = 0x02014b50
const ZIP_END = 0x06054b50
const ZIP64_LOCATOR = 0x07064b50
const ZIP64_END = 0x06064b50

// [Content_Types].xml main part → precise OOXML kind.
const OOXML_CONTENT_TYPES = Object.freeze([
  ['application/vnd.ms-word.template.macroEnabledTemplate.main+xml', 'dotm'],
  ['application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml', 'dotx'],
  ['application/vnd.ms-word.document.macroEnabled.main+xml', 'docm'],
  ['application/vnd.ms-excel.template.macroEnabled.main+xml', 'xltm'],
  ['application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml', 'xltx'],
  ['application/vnd.ms-excel.addin.macroEnabled.main+xml', 'xlam'],
  ['application/vnd.ms-excel.sheet.binary.macroEnabled.main', 'xlsb'],
  ['application/vnd.ms-excel.sheet.macroEnabled.main+xml', 'xlsm'],
])

const ODF_MIMETYPES = Object.freeze({
  'application/vnd.oasis.opendocument.text': 'odt',
  'application/vnd.oasis.opendocument.spreadsheet': 'ods',
  'application/vnd.oasis.opendocument.presentation': 'odp',
  'application/epub+zip': 'epub',
})

const HEVC_BRANDS = new Set(['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'hevm', 'hevs'])
const HEIF_BRANDS = new Set([...HEVC_BRANDS, 'mif1', 'msf1'])
const AVIF_BRANDS = new Set(['avif', 'avis'])
const HTML_ROOTS = new Set(['html', 'head', 'body', 'table', 'div', 'p', 'span', 'meta', 'title', 'style', 'script', 'link',
  'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'br', 'b', 'i', 'a', 'font', 'center', 'pre', 'form', 'img', 'section', 'article', 'main', 'header', 'nav'])
const DELIMITERS = Object.freeze([['\t', 'tab'], [';', 'semicolon'], [',', 'comma'], ['|', 'pipe']])

// ---------------------------------------------------------------------------
// Small helpers

function toBuffer(bytes) {
  if (!bytes) return Buffer.alloc(0)
  if (Buffer.isBuffer(bytes)) return bytes
  if (bytes instanceof Uint8Array) return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes instanceof ArrayBuffer) return Buffer.from(bytes)
  throw new TypeError('Expected a Buffer, Uint8Array or ArrayBuffer.')
}

/**
 * Lower-case extension with its dot, from a path, a file name, or a bare
 * dotted extension (".XLSX" and "C:\\a\\b.XLSX" both give ".xlsx").
 * A name without a dot ("README") has no extension.
 * @param {string} value
 * @returns {string} "" when there is no extension.
 */
function normalizedExtension(value) {
  const text = String(value || '')
  if (/^\.[a-z0-9]+$/i.test(text)) return text.toLowerCase()
  return path.extname(text).toLowerCase()
}

function plain(extension) {
  return extension.replace(/^\./, '')
}

function uniq(values) {
  return [...new Set(values)]
}

// ---------------------------------------------------------------------------
// Registry

function normalizeOperation(value) {
  if (value === undefined || value === null) return null
  const raw = typeof value === 'string' ? { mode: value } : { ...value }
  const engine = raw.engine || 'none'
  if (!ENGINE_MODES.includes(engine)) throw new Error(`Unknown engine policy "${engine}".`)
  return Object.freeze({ ...raw, mode: raw.mode || 'native', status: raw.status || 'active', engine })
}

function validateRegistry(data) {
  const problems = []
  if (!data || !Array.isArray(data.formats)) throw new Error('The format registry has no formats list.')
  const ids = new Set()
  const modes = new Set(data.routeModes || [])
  const workspaces = new Set(Object.keys(data.workspaces || {}))
  for (const format of data.formats) {
    if (!format.id || ids.has(format.id)) problems.push(`duplicate or missing format id "${format.id}"`)
    ids.add(format.id)
    if (!Array.isArray(format.extensions) || !format.extensions.every((ext) => /^\.[a-z0-9]+$/.test(ext))) {
      problems.push(`${format.id}: extensions must be lower-case and start with a dot`)
    }
    for (const mode of format.route || []) if (!modes.has(mode)) problems.push(`${format.id}: unknown route "${mode}"`)
    for (const [kind, targets] of Object.entries(format.contentRoutes || {})) {
      for (const mode of targets) if (!modes.has(mode)) problems.push(`${format.id}: content route ${kind} → unknown "${mode}"`)
    }
    for (const workspace of Object.keys(format.workspaces || {})) {
      if (!workspaces.has(workspace)) problems.push(`${format.id}: unknown workspace "${workspace}"`)
    }
    if (!['active', 'planned'].includes(format.status || 'active')) problems.push(`${format.id}: unknown status "${format.status}"`)
  }
  for (const kind of data.contentKinds || []) {
    if (ids.has(kind.id)) problems.push(`content kind "${kind.id}" shadows a format id`)
    if (kind.format && !ids.has(kind.format)) problems.push(`content kind "${kind.id}" points to unknown format "${kind.format}"`)
  }
  if (problems.length) throw new Error(`The format registry is invalid: ${problems.join('; ')}.`)
}

/**
 * Builds a registry API over registry data. The module exports one built from
 * formats.json; tests and tools build others (for example with a planned
 * format switched on) to see how routing and filters would change.
 *
 * @param {object} data Parsed formats.json.
 * @returns {object} The registry API (see the module exports for each function).
 */
function createRegistry(data) {
  validateRegistry(data)
  const formats = Object.freeze(data.formats.map((format) => Object.freeze({ ...format })))
  const byId = new Map(formats.map((format) => [format.id, format]))
  const kinds = new Map((data.contentKinds || []).map((kind) => [kind.id, kind]))
  const routeModes = Object.freeze([...(data.routeModes || [])])

  /**
   * The registry entry for one operation of a format in a workspace,
   * normalised to {mode, status, engine, …}; null when none is declared.
   * isOperationActive() tells whether it is live.
   * @param {object|string} formatOrId
   * @param {string} workspace e.g. "calc".
   * @param {'open'|'save'|'export'} op
   * @returns {object|null}
   */
  function operation(formatOrId, workspace, op) {
    const format = typeof formatOrId === 'string' ? byId.get(formatOrId) : formatOrId
    if (!format || !OPERATIONS.includes(op)) return null
    return normalizeOperation(format.workspaces?.[workspace]?.[op])
  }

  function isLive(format, workspace, op) {
    if ((format.status || 'active') !== 'active') return false
    const entry = operation(format, workspace, op)
    return Boolean(entry && entry.status === 'active')
  }

  function engineAllows(entry, engine) {
    return entry.engine !== 'required' || Boolean(engine)
  }

  // Default owner per extension: the first mode in `route` whose open
  // operation is live. Exactly one owner per extension.
  const modeByExtension = new Map()
  const routedFormatByExtension = new Map()
  const extensionsByMode = Object.fromEntries(routeModes.map((mode) => [mode, []]))
  for (const format of formats) {
    const mode = (format.route || []).find((candidate) => isLive(format, candidate, 'open'))
    if (!mode) continue
    for (const extension of format.extensions) {
      if (modeByExtension.has(extension)) {
        throw new Error(`The format registry routes ${extension} twice (${routedFormatByExtension.get(extension).id} and ${format.id}).`)
      }
      modeByExtension.set(extension, mode)
      routedFormatByExtension.set(extension, format)
      extensionsByMode[mode].push(extension)
    }
  }
  for (const mode of routeModes) Object.freeze(extensionsByMode[mode])
  Object.freeze(extensionsByMode)
  const routedExtensions = Object.freeze([...modeByExtension.keys()])

  /**
   * The format an extension names. Routed (active) formats win; otherwise the
   * first registry entry, including planned ones when `includePlanned`.
   * @param {string} extensionOrPath
   * @param {{includePlanned?: boolean}} [options]
   * @returns {object|null}
   */
  function formatForExtension(extensionOrPath, options = {}) {
    const raw = String(extensionOrPath || '')
    const extension = /^[a-z0-9]+$/i.test(raw) ? `.${raw.toLowerCase()}` : normalizedExtension(raw)
    if (!extension) return null
    if (routedFormatByExtension.has(extension)) return routedFormatByExtension.get(extension)
    const candidates = formats.filter((format) => format.extensions.includes(extension))
    const active = candidates.find((format) => (format.status || 'active') === 'active')
    if (active) return active
    return options.includePlanned ? candidates[0] || null : null
  }

  /**
   * Formats a workspace can open now.
   * @param {string} workspace
   * @param {{engine?: boolean}} [options] engine: the optional office engine is installed locally.
   * @returns {object[]}
   */
  function openableFormats(workspace, options = {}) {
    return formats.filter((format) => isLive(format, workspace, 'open') && engineAllows(operation(format, workspace, 'open'), options.engine))
  }

  /**
   * Extensions (with dots) a workspace can open now, in registry order.
   * @param {string} workspace
   * @param {{engine?: boolean}} [options]
   * @returns {string[]}
   */
  function openExtensions(workspace, options = {}) {
    if (workspace === 'launcher') return [...routedExtensions]
    return uniq(openableFormats(workspace, options).flatMap((format) => format.extensions))
  }

  function formatsForGroup(group, openable) {
    if (group.formats) return group.formats.map((id) => openable.find((format) => format.id === id)).filter(Boolean)
    if (group.families) return openable.filter((format) => group.families.includes(format.family))
    return []
  }

  /**
   * Native Open dialog filters for a workspace: "All supported …" first, then
   * one filter per family (or per format), then "All files" so mislabelled and
   * extensionless files can still be chosen and sniffed.
   * Engine-only formats are listed only when the engine is present.
   *
   * @param {string} workspace "docs", "pdf", "calc", "image", "video", "combine" or "launcher".
   * @param {{engine?: boolean, allFiles?: boolean}} [options] allFiles defaults to true.
   * @returns {{name: string, extensions: string[]}[]} Electron FileFilter objects.
   */
  function dialogFilters(workspace, options = {}) {
    const spec = data.dialogs?.[workspace]
    if (!spec) throw new Error(`No dialog filters are defined for "${workspace}".`)
    const filters = []
    if (workspace === 'launcher') {
      filters.push({ name: spec.all, extensions: routedExtensions.map(plain) })
      for (const group of spec.groups) {
        const extensions = group.routes.flatMap((mode) => extensionsByMode[mode] || [])
        if (extensions.length) filters.push({ name: group.name, extensions: extensions.map(plain) })
      }
    } else {
      const openable = openableFormats(workspace, options)
      const groups = spec.groups === 'per-format'
        ? openable.map((format) => ({ name: format.label, list: [format] }))
        : spec.groups.map((group) => ({ name: group.name, list: formatsForGroup(group, openable) }))
      // "All supported" lists grouped formats first, in group order, then any
      // openable format no group mentions, so nothing openable is left out.
      const all = uniq([...groups.flatMap((group) => group.list), ...openable].flatMap((format) => format.extensions))
      if (!all.length) return options.allFiles === false ? [] : [{ ...ALL_FILES_FILTER, extensions: ['*'] }]
      filters.push({ name: spec.all, extensions: all.map(plain) })
      for (const group of groups) {
        const extensions = uniq(group.list.flatMap((format) => format.extensions))
        if (extensions.length) filters.push({ name: group.name, extensions: extensions.map(plain) })
      }
    }
    if (options.allFiles !== false) filters.push({ ...ALL_FILES_FILTER, extensions: ['*'] })
    return filters
  }

  /**
   * `build.fileAssociations` for electron-builder: one entry per routed
   * workspace, extensions without dots, in registry order.
   * @returns {{ext: string[], name: string, description: string, role: string, icon: string}[]}
   */
  function fileAssociations() {
    const groups = data.associations?.groups || {}
    return routeModes
      .filter((mode) => extensionsByMode[mode].length)
      .map((mode) => ({
        ext: extensionsByMode[mode].map(plain),
        name: groups[mode]?.name || `simple ${mode}`,
        description: groups[mode]?.description || `File opened by simple ${mode}`,
        role: groups[mode]?.role || 'Editor',
        icon: data.associations?.icon || 'build/icon.ico',
      }))
  }

  /**
   * What Open does for a format in a workspace.
   * @param {string} formatId
   * @param {string} workspace
   * @param {{engine?: boolean}} [options]
   * @returns {{available: boolean, reason: null|'planned'|'needs-office-engine'|'not-supported',
   *   mode: string|null, engine: string, usesEngine: boolean, note?: string}}
   */
  function openPolicy(formatId, workspace, options = {}) {
    const format = byId.get(formatId)
    const entry = format && operation(format, workspace, 'open')
    if (!entry) return { available: false, reason: 'not-supported', mode: null, engine: 'none', usesEngine: false }
    if (!isLive(format, workspace, 'open')) return { available: false, reason: 'planned', mode: entry.mode, engine: entry.engine, usesEngine: false }
    if (!engineAllows(entry, options.engine)) return { available: false, reason: 'needs-office-engine', mode: entry.mode, engine: entry.engine, usesEngine: false }
    return {
      available: true,
      reason: null,
      mode: entry.mode,
      engine: entry.engine,
      usesEngine: entry.engine !== 'none' && Boolean(options.engine),
      ...(entry.note ? { note: entry.note } : {}),
    }
  }

  function siblingTarget(spec) {
    const match = /^sibling:([a-z0-9-]+)$/.exec(String(spec || ''))
    if (!match) return null
    const target = byId.get(match[1])
    if (!target) throw new Error(`Unknown sibling format "${match[1]}".`)
    return { target: 'sibling', formatId: target.id, extension: target.extensions[0] }
  }

  /**
   * Where Save writes an edited document of this format.
   * `in-place` keeps the file and format; `sibling` writes a new file with
   * `extension` next to the original (for example the edits to a legacy
   * .xls go to a modern .xlsx when the office engine is not installed).
   *
   * @param {string} formatId
   * @param {string} workspace
   * @param {{engine?: boolean}} [options]
   * @returns {{target: 'in-place', usesEngine: boolean}|{target: 'sibling', formatId: string, extension: string}|null}
   *   null when the workspace never saves this format.
   */
  function savePolicy(formatId, workspace, options = {}) {
    const format = byId.get(formatId)
    const entry = format && operation(format, workspace, 'save')
    if (!entry) return null
    const sibling = siblingTarget(entry.mode)
    if (sibling) return sibling
    const live = isLive(format, workspace, 'save')
    if (entry.mode === 'engine') {
      if (live && options.engine) return { target: 'in-place', usesEngine: true }
      return siblingTarget(entry.fallback) || null
    }
    if (entry.mode === 'in-place' && live) return { target: 'in-place', usesEngine: false }
    return siblingTarget(entry.fallback) || null
  }

  /**
   * Export As choices of a workspace, in registry order. Engine-only rows are
   * returned with `available:false` and reason "needs-office-engine" so the
   * dialog can show what Simple offers instead; planned rows are omitted.
   *
   * @param {string} workspace
   * @param {{engine?: boolean}} [options]
   * @returns {{id: string, label: string, extensions: string[], description: string, mode: string,
   *   available: boolean, reason: string|null, usesEngine: boolean, lossy: object|null}[]}
   */
  function exportFormats(workspace, options = {}) {
    const rows = []
    for (const format of formats) {
      if (!isLive(format, workspace, 'export')) continue
      const entry = operation(format, workspace, 'export')
      const available = engineAllows(entry, options.engine)
      rows.push({
        id: format.id,
        label: entry.label || format.label,
        extensions: [...format.extensions],
        description: format.description || '',
        mode: entry.mode,
        available,
        reason: available ? null : 'needs-office-engine',
        usesEngine: available && entry.engine !== 'none' && Boolean(options.engine),
        lossy: format.lossy || null,
      })
    }
    return rows
  }

  /**
   * Registry entry for a sniffed kind: a format id or a content kind.
   * @param {string} kind
   * @returns {{id: string, label: string, family: string, formatId: string|null}}
   */
  function describeKind(kind) {
    const format = byId.get(kind)
    if (format) return { id: kind, label: format.label, family: format.family, formatId: format.id }
    const content = kinds.get(kind)
    if (content) return { id: kind, label: content.label, family: content.family, formatId: content.format || null }
    return { id: kind, label: 'Unknown file', family: 'binary', formatId: null }
  }

  /**
   * Identifies a file by its content (§6.2 of the design): PDF, ZIP-based
   * (OOXML, ODF, EPUB, Numbers), compound files (doc, xls, ppt, encrypted
   * OOXML), images and video by magic number, then text by content (RTF,
   * Office 2003 XML, SVG, Excel web pages, HTML, delimited text, Markdown).
   * A name/content mismatch is reported, never treated as an error: the
   * content decides which reader runs.
   *
   * @param {Buffer|Uint8Array} head The first bytes of the file (up to 64 KB is enough).
   * @param {Buffer|Uint8Array} [tail] The last bytes of the file (needed for ZIP and CFB directories).
   * @param {number} [size] Full file size; defaults to head length when the head is the whole file.
   * @param {string} [name] File name or path, used only to report mismatches.
   * @returns {{kind: string, formatId: string|null, label: string, family: string,
   *   confidence: 'certain'|'high'|'medium'|'low', extension: string, extensionFormatId: string|null,
   *   matchesExtension: boolean, mislabeled: boolean, encrypted: boolean, details: object}}
   */
  function sniff(head, tail, size, name) {
    const headBuffer = toBuffer(head)
    const tailBuffer = tail ? toBuffer(tail) : headBuffer
    const total = Number.isFinite(size) ? size : headBuffer.length
    const found = detectKind(headBuffer, tailBuffer, total)
    const described = describeKind(found.kind)
    const extension = normalizedExtension(name || '')
    const extensionFormat = formatForExtension(extension, { includePlanned: true })
    const allowed = extensionFormat ? new Set([extensionFormat.id, ...(extensionFormat.matches || [])]) : null
    const matchesExtension = Boolean(allowed && allowed.has(found.kind))
    const identified = !['binary', 'empty'].includes(found.kind)
    return {
      kind: found.kind,
      formatId: described.formatId,
      label: described.label,
      family: described.family,
      confidence: found.confidence,
      extension,
      extensionFormatId: extensionFormat ? extensionFormat.id : null,
      matchesExtension,
      mislabeled: Boolean(extensionFormat && identified && !matchesExtension),
      encrypted: Boolean(found.encrypted),
      details: found.details || {},
    }
  }

  /**
   * Reads up to 64 KB from each end of a file and sniffs it.
   * @param {string} filePath
   * @returns {ReturnType<typeof sniff> & {size: number}}
   * @throws The fs error when the file cannot be read (ENOENT, EBUSY, EPERM …).
   */
  function sniffFile(filePath) {
    const sample = readSample(filePath)
    return { ...sniff(sample.head, sample.tail, sample.size, filePath), size: sample.size }
  }

  const routeCache = new Map()

  /**
   * The workspace that opens a path (§6.3). Synchronous and cheap: only
   * content-routed extensions (.txt, .html/.htm, .xml) read the first 64 KB,
   * and content routing only chooses between owners declared in the
   * registry. A file that cannot be read keeps its extension route.
   *
   * @param {string} filePath
   * @param {object} [options]
   * @param {string} [options.cwd] Folder that relative paths are resolved against (second-instance argv).
   * @param {boolean} [options.sniffUnknown=false] Sniff extensionless and unrouted files and route them
   *   to the owner of the detected format (for "All files" pickers and drops; argv routing leaves this off
   *   because workspaces still accept paths by extension).
   * @param {(filePath: string) => {head: Buffer, tail: Buffer, size: number}} [options.readSample] Test hook.
   * @returns {{mode: string|null, formatId: string|null, extension: string, by: 'extension'|'content'|null, kind?: string}}
   */
  function routeForPath(filePath, options = {}) {
    const extension = normalizedExtension(filePath)
    const format = routedFormatByExtension.get(extension) || null
    const fallback = format
      ? { mode: modeByExtension.get(extension), formatId: format.id, extension, by: 'extension' }
      : { mode: null, formatId: null, extension, by: null }
    const needsContent = format ? format.routeBy === 'content' : Boolean(options.sniffUnknown)
    if (!needsContent || typeof filePath !== 'string' || !filePath) return fallback
    const resolved = path.resolve(options.cwd || process.cwd(), filePath)
    let found
    try {
      const reader = options.readSample || readSample
      const stat = options.readSample ? null : fs.statSync(resolved)
      if (stat && !stat.isFile()) return fallback
      const key = stat ? `${resolved}\0${stat.size}\0${stat.mtimeMs}` : null
      if (key && routeCache.has(key)) {
        found = routeCache.get(key)
      } else {
        const sample = reader(resolved)
        found = sniff(sample.head, sample.tail, sample.size, resolved)
        if (key) {
          if (routeCache.size > 64) routeCache.clear()
          routeCache.set(key, found)
        }
      }
    } catch {
      return fallback
    }
    if (format) {
      const targets = format.contentRoutes?.[found.kind] || []
      const mode = targets.find((candidate) => isLive(format, candidate, 'open'))
      if (mode) return { mode, formatId: format.id, extension, by: 'content', kind: found.kind }
      return { ...fallback, kind: found.kind }
    }
    const target = found.formatId && byId.get(found.formatId)
    const mode = target && (target.route || []).find((candidate) => isLive(target, candidate, 'open'))
    if (!mode) return { ...fallback, kind: found.kind }
    return { mode, formatId: target.id, extension, by: 'content', kind: found.kind }
  }

  /**
   * routeForPath() without blocking the calling process: the content sample
   * of a content-routed file is read asynchronously. A file that cannot be
   * read keeps its extension route.
   * @param {string} filePath
   * @param {{cwd?: string, sniffUnknown?: boolean}} [options]
   * @returns {Promise<ReturnType<typeof routeForPath>>}
   */
  async function routeForPathAsync(filePath, options = {}) {
    const extension = normalizedExtension(filePath)
    const format = routedFormatByExtension.get(extension) || null
    const needsContent = format ? format.routeBy === 'content' : Boolean(options.sniffUnknown)
    if (!needsContent || typeof filePath !== 'string' || !filePath) return routeForPath(filePath, { ...options, readSample: undefined })
    const resolved = path.resolve(options.cwd || process.cwd(), filePath)
    let sample = null
    try {
      const stat = await fs.promises.stat(resolved)
      if (stat.isFile()) sample = await readSampleAsync(resolved)
    } catch {
      sample = null
    }
    if (!sample) return routeForPath(filePath, { ...options, readSample: () => { throw new Error('unreadable') } })
    return routeForPath(filePath, { ...options, readSample: () => sample })
  }

  /**
   * Every workspace a path with this extension can be routed to: its
   * extension route plus the owners its content can choose (only workspaces
   * that open the format now). Needs no file access.
   * @param {string} extensionOrPath
   * @returns {string[]}
   */
  function candidateModesForExtension(extensionOrPath) {
    const extension = normalizedExtension(extensionOrPath)
    const format = routedFormatByExtension.get(extension)
    if (!format) return []
    const modes = new Set([modeByExtension.get(extension)])
    if (format.routeBy === 'content') {
      for (const targets of Object.values(format.contentRoutes || {})) {
        for (const candidate of targets || []) if (isLive(format, candidate, 'open')) modes.add(candidate)
      }
    }
    return [...modes].filter(Boolean)
  }

  return Object.freeze({
    data,
    formats,
    routeModes,
    extensionsByMode,
    routedExtensions,
    formatById: (id) => byId.get(id) || null,
    formatForExtension,
    operation,
    isOperationActive: (formatOrId, workspace, op) => {
      const format = typeof formatOrId === 'string' ? byId.get(formatOrId) : formatOrId
      return Boolean(format && isLive(format, workspace, op))
    },
    defaultModeForExtension: (extensionOrPath) => modeByExtension.get(normalizedExtension(extensionOrPath)) || null,
    openableFormats,
    openExtensions,
    dialogFilters,
    fileAssociations,
    associationExtensions: () => [...routedExtensions],
    openPolicy,
    savePolicy,
    exportFormats,
    describeKind,
    sniff,
    sniffFile,
    routeForPath,
    routeForPathAsync,
    candidateModesForExtension,
  })
}

// ---------------------------------------------------------------------------
// Reading samples

/**
 * Reads the first and last 64 KB of a file (one read when the file is small).
 * @param {string} filePath
 * @param {number} [bytes=SNIFF_BYTES]
 * @returns {{head: Buffer, tail: Buffer, size: number}}
 */
function readSample(filePath, bytes = SNIFF_BYTES) {
  const handle = fs.openSync(filePath, 'r')
  try {
    const { size } = fs.fstatSync(handle)
    const head = Buffer.alloc(Math.min(size, bytes))
    const headRead = fs.readSync(handle, head, 0, head.length, 0)
    if (size <= bytes) {
      const whole = head.subarray(0, headRead)
      return { head: whole, tail: whole, size }
    }
    const tail = Buffer.alloc(Math.min(size, bytes))
    const tailRead = fs.readSync(handle, tail, 0, tail.length, size - tail.length)
    return { head: head.subarray(0, headRead), tail: tail.subarray(0, tailRead), size }
  } finally {
    fs.closeSync(handle)
  }
}

/**
 * readSample() without blocking: the first and last 64 KB of a file.
 * @param {string} filePath
 * @param {number} [bytes=SNIFF_BYTES]
 * @returns {Promise<{head: Buffer, tail: Buffer, size: number}>}
 */
async function readSampleAsync(filePath, bytes = SNIFF_BYTES) {
  const handle = await fs.promises.open(filePath, 'r')
  try {
    const { size } = await handle.stat()
    const head = Buffer.alloc(Math.min(size, bytes))
    const { bytesRead: headRead } = await handle.read(head, 0, head.length, 0)
    if (size <= bytes) {
      const whole = head.subarray(0, headRead)
      return { head: whole, tail: whole, size }
    }
    const tail = Buffer.alloc(Math.min(size, bytes))
    const { bytesRead: tailRead } = await handle.read(tail, 0, tail.length, size - tail.length)
    return { head: head.subarray(0, headRead), tail: tail.subarray(0, tailRead), size }
  } finally {
    await handle.close().catch(() => {})
  }
}

// ---------------------------------------------------------------------------
// Content detection

/** Reads `length` bytes at an absolute file offset from the head or tail sample. */
function makeReader(head, tail, size) {
  const tailStart = size - tail.length
  return (offset, length) => {
    if (offset < 0 || length < 0 || offset + length > size) return null
    if (offset + length <= head.length) return head.subarray(offset, offset + length)
    if (offset >= tailStart && offset + length <= size) return tail.subarray(offset - tailStart, offset - tailStart + length)
    return null
  }
}

function startsWithAscii(buffer, text, offset = 0) {
  return buffer.length >= offset + text.length && buffer.toString('latin1', offset, offset + text.length) === text
}

function detectKind(head, tail, size) {
  if (!size || !head.length) return { kind: 'empty', confidence: 'certain' }
  const pdfAt = head.subarray(0, 1024).indexOf('%PDF-', 0, 'latin1')
  if (pdfAt !== -1) return { kind: 'pdf', confidence: pdfAt === 0 ? 'certain' : 'high', details: pdfAt ? { offset: pdfAt } : {} }
  if (head.length >= 4 && head[0] === 0x50 && head[1] === 0x4b && [3, 5, 7].includes(head[2])) {
    return detectZip(head, tail, size)
  }
  if (head.length >= 8 && head.subarray(0, 8).equals(CFB_SIGNATURE)) return detectCfb(head, tail, size)
  const binary = detectBinaryMagic(head)
  if (binary) return binary
  return detectText(head, size)
}

// ZIP ----------------------------------------------------------------------

function zipNames(head, tail, size) {
  const read = makeReader(head, tail, size)
  const entries = []
  // End of central directory: scan the tail backwards (comment ≤ 64 KB).
  let endAt = -1
  for (let index = tail.length - 22; index >= 0; index -= 1) {
    if (tail.readUInt32LE(index) === ZIP_END) { endAt = index; break }
  }
  if (endAt !== -1) {
    const tailStart = size - tail.length
    let count = tail.readUInt16LE(endAt + 10)
    let cdSize = tail.readUInt32LE(endAt + 12)
    let cdOffset = tail.readUInt32LE(endAt + 16)
    if ((cdOffset === 0xffffffff || cdSize === 0xffffffff || count === 0xffff) && endAt >= 20 && tail.readUInt32LE(endAt - 20) === ZIP64_LOCATOR) {
      const recordOffset = Number(tail.readBigUInt64LE(endAt - 20 + 8))
      const record = read(recordOffset, 56)
      if (record && record.readUInt32LE(0) === ZIP64_END) {
        count = Number(record.readBigUInt64LE(32))
        cdSize = Number(record.readBigUInt64LE(40))
        cdOffset = Number(record.readBigUInt64LE(48))
      }
    } else if (tailStart + endAt - cdSize >= 0 && tailStart + endAt - cdSize !== cdOffset) {
      // Data prepended to the archive (self-extractors): trust the position.
      cdOffset = tailStart + endAt - cdSize
    }
    let offset = cdOffset
    for (let index = 0; index < count && index < 100000; index += 1) {
      const fixed = read(offset, 46)
      if (!fixed || fixed.readUInt32LE(0) !== ZIP_CENTRAL) break
      const nameLength = fixed.readUInt16LE(28)
      const extraLength = fixed.readUInt16LE(30)
      const commentLength = fixed.readUInt16LE(32)
      const nameBytes = read(offset + 46, nameLength)
      if (!nameBytes) break
      entries.push({
        name: nameBytes.toString('utf8'),
        method: fixed.readUInt16LE(10),
        compressedSize: fixed.readUInt32LE(20),
        localOffset: fixed.readUInt32LE(42),
      })
      offset += 46 + nameLength + extraLength + commentLength
    }
  }
  if (!entries.length) {
    // No readable central directory (huge or truncated archive): fall back to
    // the local headers visible in the head and central entries in the tail.
    for (const [buffer, base] of [[head, 0], [tail, size - tail.length]]) {
      for (let index = 0; index + 46 <= buffer.length; index += 1) {
        const signature = buffer.readUInt32LE(index)
        if (signature === ZIP_LOCAL && index + 30 <= buffer.length) {
          const nameLength = buffer.readUInt16LE(index + 26)
          if (index + 30 + nameLength > buffer.length) continue
          entries.push({ name: buffer.toString('utf8', index + 30, index + 30 + nameLength), method: buffer.readUInt16LE(index + 8), compressedSize: buffer.readUInt32LE(index + 18), localOffset: base + index })
        } else if (signature === ZIP_CENTRAL) {
          const nameLength = buffer.readUInt16LE(index + 28)
          if (index + 46 + nameLength > buffer.length) continue
          entries.push({ name: buffer.toString('utf8', index + 46, index + 46 + nameLength), method: buffer.readUInt16LE(index + 10), compressedSize: buffer.readUInt32LE(index + 20), localOffset: buffer.readUInt32LE(index + 42) })
        }
      }
    }
  }
  return { entries, read }
}

function readZipEntry(entry, read, limit = 256 * 1024) {
  if (!entry) return null
  const local = read(entry.localOffset, 30)
  if (!local || local.readUInt32LE(0) !== ZIP_LOCAL) return null
  const start = entry.localOffset + 30 + local.readUInt16LE(26) + local.readUInt16LE(28)
  const length = entry.compressedSize || local.readUInt32LE(18)
  const data = read(start, Math.min(length, limit))
  if (!data) return null
  if (entry.method === 0) return data
  if (entry.method !== 8) return null
  try {
    return zlib.inflateRawSync(data, { finishFlush: zlib.constants.Z_SYNC_FLUSH, maxOutputLength: limit })
  } catch {
    return null
  }
}

function detectZip(head, tail, size) {
  const { entries, read } = zipNames(head, tail, size)
  const names = new Set(entries.map((entry) => entry.name))
  const find = (name) => entries.find((entry) => entry.name === name)
  const has = (prefix) => entries.some((entry) => entry.name.startsWith(prefix))
  const details = { entries: entries.length }

  if (names.has('mimetype')) {
    const mimetype = readZipEntry(find('mimetype'), read)
    const value = mimetype ? mimetype.toString('latin1').trim() : ''
    if (ODF_MIMETYPES[value]) return { kind: ODF_MIMETYPES[value], confidence: 'certain', details: { ...details, mimetype: value } }
    if (value) return { kind: 'odf', confidence: 'high', details: { ...details, mimetype: value } }
  }
  if (names.has('[Content_Types].xml')) {
    const types = readZipEntry(find('[Content_Types].xml'), read)
    const typesText = types ? types.toString('utf8') : ''
    for (const [contentType, kind] of OOXML_CONTENT_TYPES) {
      if (typesText.includes(contentType)) return { kind, confidence: 'certain', details }
    }
    if (has('word/')) return { kind: names.has('word/vbaProject.bin') ? 'docm' : 'docx', confidence: typesText ? 'certain' : 'high', details }
    if (names.has('xl/workbook.bin')) return { kind: 'xlsb', confidence: 'high', details }
    if (has('xl/')) return { kind: names.has('xl/vbaProject.bin') ? 'xlsm' : 'xlsx', confidence: typesText ? 'certain' : 'high', details }
    if (has('ppt/')) return { kind: 'pptx', confidence: 'high', details }
  }
  if (names.has('Index/Document.iwa') || (names.has('Index.zip') && has('Metadata/'))) return { kind: 'numbers', confidence: 'high', details }
  if (names.has('content.xml') && names.has('META-INF/manifest.xml')) return { kind: 'odf', confidence: 'medium', details }
  return { kind: 'zip', confidence: entries.length ? 'high' : 'medium', details }
}

// Compound File Binary ------------------------------------------------------

function cfbNames(head, tail, size) {
  const names = new Set()
  // Directory entries are 128 bytes and sectors start at multiples of 512,
  // so every entry sits at an absolute offset divisible by 128.
  const scan = (buffer, base) => {
    const first = (128 - (base % 128)) % 128
    for (let index = first; index + 128 <= buffer.length; index += 128) {
      if (base + index < 512) continue // the header
      const nameLength = buffer.readUInt16LE(index + 64)
      const type = buffer[index + 66]
      if (nameLength < 4 || nameLength > 64 || nameLength % 2 || ![1, 2, 5].includes(type)) continue
      if (buffer.readUInt16LE(index + nameLength - 2) !== 0) continue
      const name = buffer.toString('utf16le', index, index + nameLength - 2)
      if (/^[\x01-\x06]?[\x20-\x7e]+$/.test(name)) names.add(name.replace(/^[\x01-\x06]/, ''))
    }
  }
  scan(head, 0)
  if (size > head.length) scan(tail, size - tail.length)
  return names
}

function detectCfb(head, tail, size) {
  const names = cfbNames(head, tail, size)
  const details = { streams: [...names].slice(0, 32) }
  if (names.has('EncryptedPackage') || names.has('EncryptionInfo')) {
    return { kind: 'ooxml-encrypted', confidence: 'certain', encrypted: true, details }
  }
  if (names.has('WordDocument')) return { kind: 'doc', confidence: 'certain', details }
  if (names.has('Workbook') || names.has('Book')) return { kind: 'xls', confidence: 'certain', details }
  if (names.has('PowerPoint Document')) return { kind: 'ppt', confidence: 'certain', details }
  return { kind: 'cfb', confidence: names.size ? 'high' : 'medium', details }
}

// Images, video and other binary magic numbers --------------------------------

function detectPng(head) {
  let animated = false
  for (let offset = 8; offset + 8 <= head.length;) {
    const length = head.readUInt32BE(offset)
    const type = head.toString('latin1', offset + 4, offset + 8)
    if (type === 'acTL') { animated = true; break }
    if (type === 'IDAT' || type === 'IEND') break
    offset += 12 + length
  }
  return animated
    ? { kind: 'apng', confidence: 'certain', details: { animated: true } }
    : { kind: 'png', confidence: 'certain', details: {} }
}

function detectIsoMedia(head) {
  const boxSize = head.readUInt32BE(0)
  const end = Math.min(head.length, boxSize >= 16 ? boxSize : 32)
  const major = head.toString('latin1', 8, 12)
  const brands = new Set([major])
  for (let offset = 16; offset + 4 <= end; offset += 4) brands.add(head.toString('latin1', offset, offset + 4))
  const details = { brands: [...brands] }
  const listed = [...brands]
  if (AVIF_BRANDS.has(major) || (listed.some((brand) => AVIF_BRANDS.has(brand)) && !HEVC_BRANDS.has(major))) {
    return { kind: 'avif', confidence: 'certain', details }
  }
  if (HEIF_BRANDS.has(major) || listed.some((brand) => HEVC_BRANDS.has(brand))) {
    return { kind: 'heic', confidence: 'certain', details }
  }
  if (major === 'qt  ') return { kind: 'mov', confidence: 'certain', details }
  if (major === 'crx ') return { kind: 'binary', confidence: 'medium', details }
  return { kind: 'mp4', confidence: 'high', details }
}

function detectEbml(head) {
  const limit = Math.min(head.length, 4096)
  for (let index = 4; index + 3 < limit; index += 1) {
    if (head[index] !== 0x42 || head[index + 1] !== 0x82) continue
    const marker = head[index + 2]
    let width = 1
    while (width <= 8 && !(marker & (0x80 >> (width - 1)))) width += 1
    if (width > 8) continue
    let length = marker & (0xff >> width)
    for (let extra = 1; extra < width; extra += 1) length = (length * 256) + head[index + 2 + extra]
    const start = index + 2 + width
    const docType = head.toString('latin1', start, Math.min(limit, start + length))
    if (docType === 'webm') return { kind: 'webm', confidence: 'certain', details: { docType } }
    if (docType === 'matroska') return { kind: 'mkv', confidence: 'certain', details: { docType } }
  }
  return { kind: 'mkv', confidence: 'medium', details: {} }
}

function detectIco(head) {
  if (head.length < 22) return null
  const count = head.readUInt16LE(4)
  if (count < 1 || count > 512) return null
  const reserved = head[9]
  const planes = head.readUInt16LE(10)
  const bits = head.readUInt16LE(12)
  const bytes = head.readUInt32LE(14)
  const offset = head.readUInt32LE(18)
  if (reserved !== 0 || planes > 1 || ![0, 1, 2, 4, 8, 16, 24, 32].includes(bits) || !bytes || offset < 6 + (16 * count)) return null
  return { kind: 'ico', confidence: 'high', details: { images: count } }
}

function detectGzip(head) {
  let inflated = null
  try {
    inflated = zlib.gunzipSync(head, { finishFlush: zlib.constants.Z_SYNC_FLUSH, maxOutputLength: 64 * 1024 })
  } catch (error) {
    // A buffer limit error still means the stream is valid gzip.
    if (error && error.code === 'ERR_BUFFER_TOO_LARGE') inflated = null
  }
  const text = inflated ? inflated.toString('utf8', 0, Math.min(inflated.length, 8192)) : ''
  if (/<svg[\s>]/i.test(text)) return { kind: 'svgz', confidence: 'certain', details: {} }
  return { kind: 'gzip', confidence: 'high', details: {} }
}

function detectBinaryMagic(head) {
  if (head.length >= 8 && head.subarray(0, 8).equals(PNG_SIGNATURE)) return detectPng(head)
  if (head.length >= 3 && head[0] === 0xff && head[1] === 0xd8 && head[2] === 0xff) return { kind: 'jpeg', confidence: 'certain', details: {} }
  if (startsWithAscii(head, 'GIF87a') || startsWithAscii(head, 'GIF89a')) {
    return { kind: 'gif', confidence: 'certain', details: { animated: head.indexOf('NETSCAPE2.0', 0, 'latin1') !== -1 } }
  }
  if (startsWithAscii(head, 'RIFF') && startsWithAscii(head, 'WEBP', 8)) {
    const vp8x = startsWithAscii(head, 'VP8X', 12)
    return { kind: 'webp', confidence: 'certain', details: { animated: Boolean(vp8x && head.length > 20 && (head[20] & 0x02)) } }
  }
  if (head.length >= 18 && startsWithAscii(head, 'BM') && [12, 16, 40, 52, 56, 64, 108, 124].includes(head.readUInt32LE(14))) {
    return { kind: 'bmp', confidence: 'high', details: {} }
  }
  if (head.length >= 4 && ((startsWithAscii(head, 'II') && head[2] === 0x2a && head[3] === 0) || (startsWithAscii(head, 'MM') && head[2] === 0 && head[3] === 0x2a))) {
    return { kind: 'tiff', confidence: 'certain', details: {} }
  }
  if (head.length >= 12 && startsWithAscii(head, 'ftyp', 4)) return detectIsoMedia(head)
  if (head.length >= 4 && head.readUInt32BE(0) === 0x1a45dfa3) return detectEbml(head)
  if (startsWithAscii(head, 'OggS')) return { kind: 'ogv', confidence: 'high', details: {} }
  if (head.length >= 6 && startsWithAscii(head, '8BPS')) {
    const version = head.readUInt16BE(4)
    if (version === 1) return { kind: 'psd', confidence: 'certain', details: { version } }
    if (version === 2) return { kind: 'psb', confidence: 'certain', details: { version } }
  }
  if (head.length >= 4 && head[0] === 0 && head[1] === 0 && head[2] === 1 && head[3] === 0) {
    const ico = detectIco(head)
    if (ico) return ico
  }
  if (head.length >= 3 && head[0] === 0x1f && head[1] === 0x8b && head[2] === 0x08) return detectGzip(head)
  return null
}

// Text ------------------------------------------------------------------------

function controlShare(text) {
  let control = 0
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    if (code === 0 || code < 0x09 || (code > 0x0d && code < 0x20 && code !== 0x1b) || code === 0xfffd) control += 1
  }
  return text.length ? control / text.length : 0
}

function looksBinary(head) {
  if (head.length >= 2 && ((head[0] === 0xff && head[1] === 0xfe) || (head[0] === 0xfe && head[1] === 0xff))) return false
  const utf16 = detectUtf16(head)
  if (utf16) {
    // NUL parity alone also matches binary headers; the text must read as text.
    const sample = Buffer.from(head.subarray(0, Math.min(head.length, 8192) & ~1))
    if (utf16 === 'utf-16be') sample.swap16()
    return controlShare(sample.toString('utf16le')) > 0.02
  }
  const sample = head.subarray(0, Math.min(head.length, 8192))
  let control = 0
  for (const byte of sample) {
    if (byte === 0) return true
    if (byte < 0x09 || (byte > 0x0d && byte < 0x20 && byte !== 0x1b)) control += 1
  }
  return sample.length > 0 && control / sample.length > 0.02
}

function stripXmlPreamble(text) {
  let rest = text
  for (let guard = 0; guard < 20; guard += 1) {
    const trimmed = rest.replace(/^\s+/, '')
    const match = /^(<\?[\s\S]*?\?>|<!--[\s\S]*?-->|<!DOCTYPE(?![^>]*\bhtml\b)[^>[]*(\[[\s\S]*?\])?\s*>)/i.exec(trimmed)
    if (!match) return trimmed
    rest = trimmed.slice(match[0].length)
  }
  return rest
}

function htmlBodyIsOnlyTable(text) {
  if (!/<table[\s>]/i.test(text)) return false
  const body = /<body[^>]*>([\s\S]*?)(<\/body>|$)/i.exec(text)
  const content = (body ? body[1] : text.replace(/<head[\s\S]*?<\/head>/i, ''))
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<(script|style)[\s\S]*?<\/\1>/gi, ' ')
  const visible = (markup) => markup.replace(/<[^>]*>/g, ' ').replace(/&nbsp;|&#160;/gi, ' ').replace(/\s+/g, '')
  const tables = content.match(/<table[\s\S]*?(<\/table>|$)/gi) || []
  const inside = visible(tables.join(' ')).length
  const outside = visible(content.replace(/<table[\s\S]*?(<\/table>|$)/gi, ' ')).length
  // A caption or title above the table is fine; paragraphs of prose are not.
  return outside <= 24 || outside <= inside * 0.05
}

function classifyMarkup(text) {
  const sample = text.slice(0, 64 * 1024)
  if (/<\?mso-application\s+progid\s*=\s*["']Excel\.Sheet["']/i.test(sample)
    || (/urn:schemas-microsoft-com:office:spreadsheet/i.test(sample) && /<(ss:)?Workbook[\s>]/i.test(sample))) {
    return { kind: 'spreadsheetml-2003', confidence: 'certain' }
  }
  if (/<\?mso-application\s+progid\s*=\s*["']Word\.Document["']/i.test(sample)
    || /schemas\.microsoft\.com\/office\/word\/2003\/wordml/i.test(sample)) {
    return { kind: 'wordml-2003', confidence: 'certain' }
  }
  const body = stripXmlPreamble(sample)
  const root = /^<([a-z][\w:.-]*)/i.exec(body)?.[1]?.toLowerCase() || ''
  if (root === 'office:document') {
    const mimetype = /office:mimetype\s*=\s*["']([^"']+)["']/i.exec(sample)?.[1] || ''
    if (mimetype === 'application/vnd.oasis.opendocument.spreadsheet') return { kind: 'fods', confidence: 'certain' }
    return { kind: 'xml', confidence: 'high', details: { root, mimetype } }
  }
  if (root === 'svg' || root === 'svg:svg') return { kind: 'svg', confidence: 'certain' }
  const html = /^<!doctype\s+html/i.test(body) || HTML_ROOTS.has(root) || /<(html|head|body)[\s>]/i.test(sample.slice(0, 4096))
  if (html) {
    if (/urn:schemas-microsoft-com:office:excel/i.test(sample) || /<meta[^>]+content\s*=\s*["']?Excel\.Sheet/i.test(sample)) {
      return { kind: 'excel-html', confidence: 'certain' }
    }
    if (htmlBodyIsOnlyTable(sample)) return { kind: 'html-table', confidence: 'high' }
    return { kind: 'html', confidence: /^<!doctype\s+html|^<html/i.test(body) ? 'certain' : 'high' }
  }
  if (root) return { kind: 'xml', confidence: /^\s*<\?xml/i.test(sample) ? 'certain' : 'medium', details: { root } }
  return null
}

function splitFields(line, delimiter) {
  if (delimiter === '|') return line.replace(/^\s*\|/, '').replace(/\|\s*$/, '').split('|').length
  let count = 1
  let quoted = false
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]
    if (char === '"') quoted = !quoted
    else if (char === delimiter && !quoted) count += 1
  }
  return count
}

/**
 * Detects consistently delimited text: the same delimiter splits at least
 * five lines into the same number (≥2) of fields.
 */
function detectDelimited(lines) {
  if (lines.length < 5) return null
  const rows = lines.slice(0, 200)
  let best = null
  for (const [delimiter, name] of DELIMITERS) {
    if (delimiter === '|' && rows.some((line) => /^\s*\|?\s*:?-{3,}/.test(line))) continue // a Markdown table
    const counts = rows.map((line) => (line.includes(delimiter) ? splitFields(line, delimiter) : 1))
    const tally = new Map()
    for (const count of counts) tally.set(count, (tally.get(count) || 0) + 1)
    const [columns, hits] = [...tally.entries()].sort((a, b) => b[1] - a[1])[0]
    if (columns < 2 || hits / rows.length < 0.9 || hits < 5) continue
    if (delimiter === ',' || delimiter === ' ') {
      const prose = rows.filter((line) => /[.!?]["')]?\s*$/.test(line) && line.length > 40).length
      if (prose / rows.length >= 0.5) continue
    }
    if (!best || hits > best.hits || (hits === best.hits && columns > best.columns)) best = { delimiter: name, columns, hits }
  }
  return best && { kind: 'delimited-text', confidence: best.hits === rows.length ? 'high' : 'medium', details: { delimiter: best.delimiter, columns: best.columns } }
}

function detectMarkdown(lines) {
  let headings = 0
  let signals = 0
  for (const line of lines.slice(0, 400)) {
    if (/^#{1,6}\s+\S/.test(line)) { headings += 1; signals += 1 } else if (/^\s*```/.test(line) || /^\s*\|?\s*:?-{3,}:?\s*\|/.test(line)) signals += 2
    else if (/^\s*([-*+]|\d+[.)])\s+\S/.test(line) || /^>\s/.test(line) || /\[[^\]]+\]\([^)\s]+\)/.test(line) || /\*\*[^*]+\*\*/.test(line)) signals += 1
  }
  if ((headings && signals >= 2) || signals >= 4) return { kind: 'md', confidence: headings ? 'high' : 'medium', details: { headings } }
  return null
}

function detectText(head, size) {
  if (looksBinary(head)) return { kind: 'binary', confidence: 'medium', details: {} }
  // Classification does not depend on the exact ANSI code page, so the
  // sniffer never spawns a registry query; readers decode the whole file.
  const decoded = decodeText(head, { truncated: head.length < size, ansiCodePage: 1252, cjk: false })
  const details = { encoding: decoded.encoding, bom: decoded.bom }
  const text = decoded.text
  const start = text.replace(/^\s+/, '')
  if (!start) return { kind: 'txt', confidence: 'low', details }
  if (start.startsWith('{\\rtf')) return { kind: 'rtf', confidence: 'certain', details }
  if (/^ID;[PN]/.test(start)) return { kind: 'sylk', confidence: 'high', details }
  if (/^TABLE\r?\n0,1\r?\n/.test(start)) return { kind: 'dif', confidence: 'high', details }
  if (start.startsWith('<')) {
    const markup = classifyMarkup(start)
    if (markup) return { ...markup, details: { ...details, ...(markup.details || {}) } }
  }
  if ((start[0] === '{' || start[0] === '[') && head.length >= size) {
    try {
      JSON.parse(start)
      return { kind: 'json', confidence: 'high', details }
    } catch {
      // Not JSON; keep classifying.
    }
  }
  const lines = text.split(/\r\n|\r|\n/)
  if (head.length < size && lines.length > 1) lines.pop() // likely cut mid-line
  const nonEmpty = lines.filter((line) => line.trim())
  const delimited = detectDelimited(nonEmpty)
  if (delimited) return { ...delimited, details: { ...details, ...delimited.details } }
  const markdown = detectMarkdown(nonEmpty)
  if (markdown) return { ...markdown, details: { ...details, ...markdown.details } }
  return { kind: 'txt', confidence: 'medium', details }
}

// ---------------------------------------------------------------------------

const registry = createRegistry(require('./formats.json'))

module.exports = {
  SNIFF_BYTES,
  createRegistry,
  normalizedExtension,
  readSample,
  readSampleAsync,
  ...registry,
}
