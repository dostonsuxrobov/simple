// Vendored from simple/shared/electron/io-dialogs.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// Native dialogs and decision prompts of the shared Save / Open / Export layer:
// - chooseSavePath(): Save As, Save a Copy, Export and the fallback Save As
//   after a failed save. It starts in the document's own folder (or the
//   remembered one), applies strict extension rules (photo.jpg never becomes
//   photo.jpg.png; "Q3 plan v2.1" becomes "Q3 plan v2.1.docx") and asks before
//   replacing a file the dialog did not confirm itself.
// - chooseOpenPaths(): the Open dialog with the format registry's filters.
// - showPrompt(): native message boxes worded by io-catalog.json, with the
//   catalog's default button and Esc on the cancel button.
// - fallback folders for files that cannot be saved where they are, and
//   sibling names for formats Simple saves next to the original.
// - remembered folders and other preferences in <userData>/io-prefs.json,
//   written through safeWriteFile.
// - SIMPLE_QA_DIALOGS: scripted answers for acceptance tests. Ignored when the
//   app is packaged.
// Only Node built-ins, electron (loaded lazily) and sibling files are required.

const fs = require('node:fs')
const fsp = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const core = require('./io-core.cjs')
const formats = require('./formats.cjs')
const { safeWriteFile } = require('./safe-write.cjs')

const PREFS_FILE = 'io-prefs.json'
const MAX_PREF_BYTES = 64 * 1024
const PREF_KEY_PATTERN = /^[A-Za-z][A-Za-z0-9._-]{0,63}$/
/** Preference keys that hold folders and must be absolute paths. */
const FOLDER_PREF_KEYS = Object.freeze(['lastSaveFolder', 'lastExportFolder', 'lastOpenFolder'])
const GRANT_TTL_MS = 15 * 60 * 1000
const MAX_DIALOG_ROUNDS = 6
const SAVE_PURPOSES = Object.freeze(['save', 'save-as', 'save-copy', 'export', 'fallback'])
const OPEN_PURPOSES = Object.freeze(['open', 'import', 'insert'])
const WARNING_PROMPTS = new Set(['prompts.not-responding', 'prompts.crashed'])

const settings = {
  module: null,
  userData: null,
  documentsDir: null,
  desktopDir: null,
}

let logger = {
  warn: (...args) => console.warn('[simple-io]', ...args),
}

// ---------------------------------------------------------------------------
// Electron access (lazy, so plain Node can load this file)
// ---------------------------------------------------------------------------

function electron() {
  try {
    const value = require('electron')
    return value && typeof value === 'object' ? value : null
  } catch {
    return null
  }
}

function electronApp() {
  const value = electron()
  return value && value.app && typeof value.app.getPath === 'function' ? value.app : null
}

function isPackaged() {
  const app = electronApp()
  return Boolean(app && app.isPackaged)
}

function usableWindow(window) {
  return window && typeof window.isDestroyed === 'function' && !window.isDestroyed() ? window : null
}

function windowForSender(sender) {
  const value = electron()
  if (!sender || !value || !value.BrowserWindow || typeof value.BrowserWindow.fromWebContents !== 'function') return null
  try {
    return usableWindow(value.BrowserWindow.fromWebContents(sender))
  } catch {
    return null
  }
}

/**
 * Configures the dialogs: the workspace (for filters and the app name) and the
 * folders used for preferences and fallbacks. Tests point these at temp folders.
 * @param {object} [options]
 * @param {string} [options.module] 'pdf' | 'calc' | 'docs' | 'image' | 'video' | 'launcher'
 * @param {string|null} [options.userData] folder of io-prefs.json (default: Electron's userData)
 * @param {string|null} [options.documentsDir] Documents folder for fallbacks (default: Electron's)
 * @param {string|null} [options.desktopDir] Desktop folder for fallbacks (default: Electron's)
 * @param {{warn: Function}} [options.logger]
 */
function configureDialogs(options = {}) {
  if (options.module !== undefined) settings.module = options.module ? String(options.module) : null
  for (const key of ['userData', 'documentsDir', 'desktopDir']) {
    if (Object.prototype.hasOwnProperty.call(options, key)) {
      settings[key] = options[key] ? path.resolve(options[key]) : null
      if (key === 'userData') resetPrefs()
    }
  }
  if (options.logger && typeof options.logger.warn === 'function') logger = { warn: options.logger.warn }
}

/** @returns {string} the workspace's display name, such as "Simple Spreadsheets" */
function appNameFor(module = settings.module) {
  return (module && core.catalog.appNames && core.catalog.appNames[module]) || 'Simple'
}

// ---------------------------------------------------------------------------
// Scripted answers for acceptance tests (SIMPLE_QA_DIALOGS)
// ---------------------------------------------------------------------------
// SIMPLE_QA_DIALOGS=<json file> holds
//   { "save": [answer, …], "open": [answer, …],
//     "prompts": { "<catalog key>": ["<button id>", …], "*": ["<button id>", …] },
//     "log": "<file that receives one JSON line per dialog>" }
// Save answers: "default" (accept the suggested path), "cancel" or null, an
// absolute path, or {"name": "x.docx"} for that name in the suggested folder.
// Open answers: an array of absolute paths, or "cancel"/null.
// Answers are used in order; when a list runs out, dialogs are canceled and
// prompts answer with their cancel button. Changing the file restarts the lists.

const qaState = { file: null, mtimeMs: null, script: null, cursors: new Map() }

function qaActive() {
  return Boolean(process.env.SIMPLE_QA_DIALOGS) && !isPackaged()
}

function qaScript() {
  if (!qaActive()) return null
  const file = path.resolve(process.env.SIMPLE_QA_DIALOGS)
  let stat = null
  try { stat = fs.statSync(file) } catch {}
  if (qaState.file !== file || qaState.mtimeMs !== (stat ? stat.mtimeMs : null)) {
    let script = {}
    if (stat) {
      try {
        script = JSON.parse(fs.readFileSync(file, 'utf8').replace(/^﻿/, '')) || {}
      } catch (error) {
        logger.warn(`SIMPLE_QA_DIALOGS file ${file} is not valid JSON: ${error.message}`)
      }
    }
    qaState.file = file
    qaState.mtimeMs = stat ? stat.mtimeMs : null
    qaState.script = script && typeof script === 'object' ? script : {}
    qaState.cursors = new Map()
  }
  return qaState.script
}

function qaTake(list, cursorKey) {
  if (!Array.isArray(list)) return { found: false }
  const index = qaState.cursors.get(cursorKey) || 0
  if (index >= list.length) return { found: false }
  qaState.cursors.set(cursorKey, index + 1)
  return { found: true, answer: list[index] }
}

function qaLog(record) {
  const script = qaState.script
  if (!script || typeof script.log !== 'string' || !script.log) return
  try {
    fs.appendFileSync(path.resolve(script.log), `${JSON.stringify({ at: new Date().toISOString(), ...record })}\n`)
  } catch {}
}

function qaSaveAnswer(defaultPath) {
  const script = qaScript()
  if (!script) return { scripted: false }
  const taken = qaTake(script.save, 'save')
  let answer = null
  if (taken.found) {
    const value = taken.answer
    if (value === 'default') answer = defaultPath
    else if (typeof value === 'string' && value !== 'cancel' && path.isAbsolute(value)) answer = path.resolve(value)
    else if (value && typeof value === 'object' && typeof value.name === 'string') answer = path.join(path.dirname(defaultPath), value.name)
  }
  qaLog({ type: 'save', defaultPath, answer })
  return { scripted: true, answer }
}

function qaOpenAnswer(defaultPath) {
  const script = qaScript()
  if (!script) return { scripted: false }
  const taken = qaTake(script.open, 'open')
  const answer = taken.found && Array.isArray(taken.answer)
    ? taken.answer.filter((item) => typeof item === 'string' && path.isAbsolute(item)).map((item) => path.resolve(item))
    : []
  qaLog({ type: 'open', defaultPath, answer })
  return { scripted: true, answer }
}

function qaPromptAnswer(key, prompt) {
  const script = qaScript()
  if (!script) return { scripted: false }
  const prompts = script.prompts && typeof script.prompts === 'object' ? script.prompts : {}
  const ids = prompt.buttons.map((button) => button.id)
  let taken = qaTake(prompts[key], `prompt:${key}`)
  if (!taken.found) taken = qaTake(prompts['*'], 'prompt:*')
  const answer = taken.found && ids.includes(taken.answer) ? taken.answer : prompt.cancelId
  qaLog({ type: 'prompt', key, message: prompt.message, answer })
  return { scripted: true, answer }
}

// ---------------------------------------------------------------------------
// Well-known folders
// ---------------------------------------------------------------------------

function userFolder(name) {
  const configuredFolder = name === 'documents' ? settings.documentsDir : settings.desktopDir
  if (configuredFolder) return configuredFolder
  const override = name === 'documents' ? process.env.SIMPLE_QA_DOCUMENTS_DIR : process.env.SIMPLE_QA_DESKTOP_DIR
  if (override && !isPackaged()) return path.resolve(override)
  const app = electronApp()
  if (app) {
    try { return app.getPath(name) } catch {}
  }
  return path.join(os.homedir(), name === 'documents' ? 'Documents' : 'Desktop')
}

/** @returns {string} the Documents folder (SIMPLE_QA_DOCUMENTS_DIR in unpackaged test runs) */
function documentsFolder() {
  return userFolder('documents')
}

/** @returns {string} the Desktop folder (SIMPLE_QA_DESKTOP_DIR in unpackaged test runs) */
function desktopFolder() {
  return userFolder('desktop')
}

async function isFolder(folder) {
  if (!folder) return false
  try {
    return (await fsp.stat(folder)).isDirectory()
  } catch {
    return false
  }
}

function sameFolder(a, b) {
  if (!a || !b) return false
  const left = path.resolve(a)
  const right = path.resolve(b)
  return process.platform === 'win32' ? left.toLowerCase() === right.toLowerCase() : left === right
}

/**
 * Whether two paths name the same file (case-insensitive on Windows).
 * @param {string} a
 * @param {string} b
 * @returns {boolean}
 */
function samePath(a, b) {
  return sameFolder(a, b)
}

// ---------------------------------------------------------------------------
// Preferences (<userData>/io-prefs.json)
// ---------------------------------------------------------------------------

const prefsState = { file: null, data: null, loading: null, queue: Promise.resolve() }

function resetPrefs() {
  prefsState.file = null
  prefsState.data = null
  prefsState.loading = null
}

function prefsFile() {
  if (settings.userData) return path.join(settings.userData, PREFS_FILE)
  const app = electronApp()
  if (app) {
    try { return path.join(app.getPath('userData'), PREFS_FILE) } catch {}
  }
  return path.join(os.tmpdir(), 'simple-io', PREFS_FILE)
}

async function loadPrefs() {
  const file = prefsFile()
  if (prefsState.file !== file) {
    prefsState.file = file
    prefsState.data = null
    prefsState.loading = null
  }
  if (prefsState.data) return prefsState.data
  if (!prefsState.loading) {
    prefsState.loading = (async () => {
      try {
        const value = JSON.parse((await fsp.readFile(file, 'utf8')).replace(/^﻿/, ''))
        return value && typeof value === 'object' && !Array.isArray(value) ? value : {}
      } catch (error) {
        if (error && error.code !== 'ENOENT') logger.warn(`Could not read ${file}; starting with default preferences: ${error.message}`)
        return {}
      }
    })()
  }
  const data = await prefsState.loading
  if (prefsState.file === file && !prefsState.data) prefsState.data = data
  return prefsState.data || data
}

function checkPrefKey(key) {
  if (typeof key !== 'string' || !PREF_KEY_PATTERN.test(key) || ['__proto__', 'constructor', 'prototype'].includes(key)) {
    throw new TypeError(`"${String(key).slice(0, 40)}" is not a valid preference name.`)
  }
  return key
}

/**
 * Reads one preference.
 * @param {string} key
 * @returns {Promise<unknown>} a copy of the stored value, or undefined
 */
async function getPref(key) {
  checkPrefKey(key)
  const data = await loadPrefs()
  if (!Object.prototype.hasOwnProperty.call(data, key)) return undefined
  return JSON.parse(JSON.stringify(data[key]))
}

/**
 * Stores one preference (undefined removes it) and writes io-prefs.json
 * through safeWriteFile. Folder preferences must be absolute paths. A failed
 * write is logged and reported, never thrown.
 * @param {string} key
 * @param {unknown} value JSON-serializable, at most 64 KB
 * @returns {Promise<boolean>} true when the preferences reached disk
 */
async function setPref(key, value) {
  checkPrefKey(key)
  let stored
  if (value !== undefined) {
    const encoded = JSON.stringify(value)
    if (encoded === undefined || Buffer.byteLength(encoded) > MAX_PREF_BYTES) throw new TypeError(`The value of "${key}" must be JSON of at most 64 KB.`)
    stored = JSON.parse(encoded)
    if (FOLDER_PREF_KEYS.includes(key) && stored !== null && (typeof stored !== 'string' || !path.isAbsolute(stored))) {
      throw new TypeError(`"${key}" must be an absolute folder path.`)
    }
  }
  const data = await loadPrefs()
  if (stored === undefined) delete data[key]
  else data[key] = stored
  const file = prefsState.file
  const run = prefsState.queue.then(async () => {
    await fsp.mkdir(path.dirname(file), { recursive: true })
    await safeWriteFile(file, `${JSON.stringify(data, null, 2)}\n`, { format: 'json' })
    return true
  })
  prefsState.queue = run.catch(() => {})
  try {
    return await run
  } catch (error) {
    logger.warn(`Could not save preferences to ${file}: ${error && (error.technical || error.message)}`)
    return false
  }
}

/**
 * Remembers the folder of the last save or export (errors are logged only).
 * @param {'save'|'export'|'open'} kind
 * @param {string} folder
 * @returns {Promise<boolean>}
 */
async function rememberFolder(kind, folder) {
  const key = kind === 'export' ? 'lastExportFolder' : kind === 'open' ? 'lastOpenFolder' : 'lastSaveFolder'
  if (!folder || !path.isAbsolute(folder)) return false
  try {
    const current = await getPref(key)
    if (current && sameFolder(current, folder)) return true
    return await setPref(key, path.resolve(folder))
  } catch (error) {
    logger.warn(`Could not remember the ${kind} folder: ${error && error.message}`)
    return false
  }
}

async function rememberedFolder(key) {
  try {
    const value = await getPref(key)
    return typeof value === 'string' && (await isFolder(value)) ? value : null
  } catch {
    return null
  }
}

// ---------------------------------------------------------------------------
// Names and extensions
// ---------------------------------------------------------------------------

function formatEntry(formatId) {
  return formatId ? formats.formatById(formatId) : null
}

function knownFormatFor(extension) {
  return extension ? formats.formatForExtension(extension, { includePlanned: true }) : null
}

/**
 * The first extension of a registry format, with its dot (".xlsx").
 * @param {string} formatId
 * @returns {string} "" for an unknown format
 */
function primaryExtension(formatId) {
  const entry = formatEntry(formatId)
  return entry ? entry.extensions[0] : ''
}

/**
 * A file name without its extension, but only when the extension is a known
 * format: "Q3 plan v2.1" keeps ".1", "photo.jpg" becomes "photo".
 * @param {string} name file name or path
 * @returns {string}
 */
function safeStem(name) {
  const base = path.basename(String(name || ''))
  const extension = path.extname(base)
  if (extension && extension.length < base.length && knownFormatFor(extension.toLowerCase())) return base.slice(0, -extension.length)
  return base
}

/**
 * The file name to suggest for a format: the document's own extension when it
 * belongs to the format (photo.jpeg stays .jpeg), else the format's first one.
 * @param {string} name
 * @param {string} formatId
 * @returns {string}
 */
function nameForFormat(name, formatId) {
  const base = path.basename(String(name || '')) || 'Untitled'
  const entry = formatEntry(formatId)
  if (!entry) return base
  const extension = path.extname(base).toLowerCase()
  if (extension && entry.extensions.includes(extension)) return base
  return `${safeStem(base)}${entry.extensions[0]}`
}

/**
 * Applies the extension rules (§3.4) to the path a Save dialog returned:
 * - a writable extension switches the format (photo.jpg in a PNG dialog is a JPEG);
 * - a known extension that is not writable here is replaced, never appended to;
 * - an unknown extension is kept and the format's extension appended
 *   ("Q3 plan v2.1" becomes "Q3 plan v2.1.docx");
 * - "photo.jpg.png", where the dialog appended its default extension to a name
 *   that already had a known one, is treated as "photo.jpg".
 *
 * @param {string} chosenPath what the dialog returned
 * @param {{format: string, formats: string[]}} rules the dialog's format and the formats writable for this purpose
 * @returns {{path: string, format: string, adjusted: boolean}}
 */
function applyExtensionRules(chosenPath, rules) {
  const writable = Array.isArray(rules.formats) ? rules.formats : []
  const current = formatEntry(rules.format)
  const currentExtension = current ? current.extensions[0] : ''
  const folder = path.dirname(chosenPath)
  const base = path.basename(chosenPath)
  const finish = (name, format) => {
    const finalPath = path.join(folder, name)
    return { path: finalPath, format, adjusted: name !== base }
  }
  const extension = path.extname(base).toLowerCase()
  const known = extension && extension.length < base.length ? knownFormatFor(extension) : null
  if (!known) return finish(`${base}${currentExtension}`, rules.format)
  const withoutExtension = base.slice(0, -extension.length)
  if (writable.includes(known.id)) {
    const inner = path.extname(withoutExtension).toLowerCase()
    const innerFormat = inner && inner.length < withoutExtension.length ? knownFormatFor(inner) : null
    if (innerFormat && innerFormat.id !== known.id && current && current.extensions.includes(extension)) {
      if (writable.includes(innerFormat.id)) return finish(withoutExtension, innerFormat.id)
      return finish(`${withoutExtension.slice(0, -inner.length)}${extension}`, known.id)
    }
    return finish(base, known.id)
  }
  return finish(`${withoutExtension}${currentExtension}`, rules.format)
}

/**
 * Where a format Simple cannot update in place gets its edits: "<stem>.<ext>"
 * next to the original, or "<stem> (edited).<ext>", "(edited 2)", … when that
 * name is already taken by a file Simple did not create for this document.
 *
 * @param {string} originalPath the document's file
 * @param {string} formatId registry id of the format to write ("xlsx")
 * @param {object} [options]
 * @param {boolean} [options.edited] always use the "(edited)" names (signed PDFs)
 * @param {string|null} [options.reuse] a sibling this document already wrote, which may be reused
 * @returns {Promise<string>}
 */
async function siblingPathFor(originalPath, formatId, options = {}) {
  const folder = path.dirname(path.resolve(originalPath))
  const stem = safeStem(path.basename(originalPath))
  const extension = primaryExtension(formatId) || path.extname(originalPath)
  const candidates = []
  if (!options.edited) candidates.push(`${stem}${extension}`)
  candidates.push(`${stem} (edited)${extension}`)
  for (let index = 2; index < 1000; index += 1) candidates.push(`${stem} (edited ${index})${extension}`)
  for (const name of candidates) {
    const candidate = path.join(folder, name)
    if (samePath(candidate, originalPath)) continue
    if (options.reuse && samePath(candidate, options.reuse)) return candidate
    if (!(await core.pathExists(candidate).catch(() => true))) return candidate
  }
  return path.join(folder, `${stem} (edited ${Date.now()})${extension}`)
}

/**
 * "<stem> (edited).<ext>" (then "(edited 2)", …) in the same folder: the Save As
 * suggestion after CHANGED_ON_DISK or SOURCE_MISSING, so both versions are kept.
 * @param {string} filePath
 * @returns {Promise<string>}
 */
async function editedCopyPath(filePath) {
  const base = path.basename(filePath)
  const extension = path.extname(base)
  const known = extension && knownFormatFor(extension.toLowerCase())
  const formatId = known ? known.id : null
  return siblingPathFor(filePath, formatId, { edited: true })
}

// ---------------------------------------------------------------------------
// Default and fallback locations
// ---------------------------------------------------------------------------

/**
 * The first folder Simple can really write to, of: the remembered save folder
 * (unless it is the failing one), Documents, Desktop. Used when a file cannot
 * be saved where it is (LOCKED, READ_ONLY, NO_PERMISSION, READ_ONLY_VOLUME,
 * FOLDER_MISSING, FILE_UNAVAILABLE, NAME_TOO_LONG).
 * @param {{failingFolder?: string|null}} [options]
 * @returns {Promise<string>}
 */
async function fallbackFolder(options = {}) {
  const candidates = []
  const remembered = await rememberedFolder('lastSaveFolder')
  if (remembered && !sameFolder(remembered, options.failingFolder)) candidates.push(remembered)
  candidates.push(documentsFolder(), desktopFolder())
  for (const folder of candidates) {
    if (!folder || sameFolder(folder, options.failingFolder)) continue
    const probe = await core.probeFolder(folder)
    if (probe.writable) return probe.folder
  }
  return documentsFolder()
}

/**
 * The path a Save-type dialog suggests, following §3.4 and §7.2.
 * @param {object} request
 * @param {string} request.purpose 'save' | 'save-as' | 'save-copy' | 'export' | 'fallback'
 * @param {string} request.format format id the dialog starts with
 * @param {string} [request.name] document name ("Budget.xlsx", "Untitled spreadsheet")
 * @param {string|null} [request.sourcePath] the document's file, when it has one
 * @param {string} [request.reason] the failure code that led to a fallback Save As
 * @returns {Promise<string>}
 */
async function suggestedSavePath(request) {
  const sourcePath = request.sourcePath ? path.resolve(request.sourcePath) : null
  const baseName = request.name || (sourcePath ? path.basename(sourcePath) : 'Untitled')
  const name = nameForFormat(baseName, request.format)
  const sourceFolder = sourcePath ? path.dirname(sourcePath) : null
  if (request.purpose === 'fallback') {
    const reason = String(request.reason || '')
    if (core.SAME_FOLDER_CODES.has(reason) && sourcePath && (await isFolder(sourceFolder))) {
      // Both versions are kept: "<name> (edited)" beside the file that changed or vanished.
      return editedCopyPath(path.join(sourceFolder, name))
    }
    if (core.FALLBACK_CODES.has(reason) || core.SAME_FOLDER_CODES.has(reason)) {
      // The file can't be written where it is: the same name in the first folder that works.
      const folder = await fallbackFolder({ failingFolder: sourceFolder })
      return path.join(folder, name)
    }
    // Any other failure (a format that could not be written, a damaged result) is not
    // about the folder, so Save As starts where it normally does.
  }
  if (request.purpose === 'export') {
    const folder = (await rememberedFolder('lastExportFolder')) || ((await isFolder(sourceFolder)) ? sourceFolder : null) || documentsFolder()
    return path.join(folder, name)
  }
  const folder = ((await isFolder(sourceFolder)) ? sourceFolder : null) || (await rememberedFolder('lastSaveFolder')) || documentsFolder()
  return path.join(folder, name)
}

// ---------------------------------------------------------------------------
// Grants: paths the user picked in a Save dialog, per window
// ---------------------------------------------------------------------------

const grants = new Map()

function grantKey(filePath) {
  const resolved = path.resolve(filePath)
  return process.platform === 'win32' ? resolved.toLowerCase() : resolved
}

function senderId(sender) {
  if (sender === null || sender === undefined) return null
  if (typeof sender === 'number') return sender
  return typeof sender.id === 'number' ? sender.id : null
}

/**
 * Records that the user chose `filePath` in a Save dialog of this window, so
 * a following save request may write there.
 * @param {object|number} sender webContents (or its id)
 * @param {string} filePath
 */
function grantSavePath(sender, filePath) {
  const id = senderId(sender)
  if (id === null || !filePath) return
  if (!grants.has(id)) grants.set(id, new Map())
  const now = Date.now()
  const map = grants.get(id)
  for (const [key, expires] of map) if (expires < now) map.delete(key)
  map.set(grantKey(filePath), now + GRANT_TTL_MS)
}

/**
 * Whether the user chose `filePath` in a Save dialog of this window recently.
 * @param {object|number} sender
 * @param {string} filePath
 * @returns {boolean}
 */
function isGranted(sender, filePath) {
  const id = senderId(sender)
  const map = id === null ? null : grants.get(id)
  if (!map || !filePath) return false
  const expires = map.get(grantKey(filePath))
  return Boolean(expires && expires >= Date.now())
}

/** Forgets every grant and recorded decision of a window (it closed). */
function revokeGrants(sender) {
  const id = senderId(sender)
  if (id !== null) {
    grants.delete(id)
    decisions.delete(id)
  }
}

// ---------------------------------------------------------------------------
// Decisions: answers the user gave in prompts main showed for a window
// ---------------------------------------------------------------------------
// A save request may carry flags that skip a safety check (force, recreate,
// confirmedLossy, confirmedSibling). Main honours them only when it showed
// the matching prompt to this window and the user chose that answer, so a
// page can never skip a check on its own.

const decisions = new Map()

function decisionKey(key, answer) {
  return `${key}\u0000${answer}`
}

/**
 * Records the answer the user chose in a prompt main showed for a window.
 * @param {object|number} sender webContents (or its id)
 * @param {string} key catalog key of the prompt, e.g. "saveFailed.CHANGED_ON_DISK"
 * @param {string} answer the chosen button id, e.g. "replace"
 * @param {string} [filePath] the file the prompt was about
 */
function recordDecision(sender, key, answer, filePath) {
  const id = senderId(sender)
  if (id === null || typeof key !== 'string' || typeof answer !== 'string') return
  if (!decisions.has(id)) decisions.set(id, new Map())
  const now = Date.now()
  const map = decisions.get(id)
  for (const [entryKey, entry] of map) if (entry.expires < now) map.delete(entryKey)
  const name = decisionKey(key, answer)
  const entry = map.get(name) || { expires: 0, paths: new Set(), any: false }
  entry.expires = now + GRANT_TTL_MS
  if (typeof filePath === 'string' && filePath && path.isAbsolute(filePath)) entry.paths.add(grantKey(filePath))
  else entry.any = true
  map.set(name, entry)
}

/**
 * Whether the user chose `answer` in prompt `key` for this window recently
 * (and, when `filePath` is given, for that file).
 * @param {object|number} sender
 * @param {string} key
 * @param {string} answer
 * @param {string} [filePath]
 * @returns {boolean}
 */
function hasDecision(sender, key, answer, filePath) {
  const id = senderId(sender)
  const map = id === null ? null : decisions.get(id)
  const entry = map ? map.get(decisionKey(key, answer)) : null
  if (!entry || entry.expires < Date.now()) return false
  if (!filePath) return entry.any || entry.paths.size > 0
  return entry.paths.has(grantKey(filePath))
}

/**
 * Forgets a window's recorded decisions about a file (it was saved), or all of them.
 * @param {object|number} sender
 * @param {string} [filePath]
 */
function clearDecisions(sender, filePath) {
  const id = senderId(sender)
  const map = id === null ? null : decisions.get(id)
  if (!map) return
  if (!filePath) {
    decisions.delete(id)
    return
  }
  const key = grantKey(filePath)
  for (const entry of map.values()) entry.paths.delete(key)
}

// ---------------------------------------------------------------------------
// Dialogs
// ---------------------------------------------------------------------------

function filterFor(formatId) {
  const entry = formatEntry(formatId)
  if (!entry) return null
  const patterns = entry.extensions.map((extension) => `*${extension}`).join(', ')
  return { name: `${entry.label} (${patterns})`, extensions: entry.extensions.map((extension) => extension.slice(1)) }
}

/**
 * Formats a workspace can write for a purpose now, from the registry:
 * in-place save formats for Save As, available export formats for Export.
 * @param {string} module
 * @param {'save'|'save-as'|'save-copy'|'export'|'fallback'} purpose
 * @param {{engine?: boolean}} [options]
 * @returns {string[]}
 */
function writableFormats(module, purpose, options = {}) {
  if (!module) return []
  if (purpose === 'export') {
    return formats.exportFormats(module, { engine: options.engine }).filter((row) => row.available).map((row) => row.id)
  }
  return formats.formats
    .filter((entry) => {
      const policy = formats.savePolicy(entry.id, module, { engine: options.engine })
      return Boolean(policy && policy.target === 'in-place')
    })
    .map((entry) => entry.id)
}

async function showSaveDialogRaw(window, options) {
  const qa = qaSaveAnswer(options.defaultPath)
  if (qa.scripted) return qa.answer
  const value = electron()
  if (!value || !value.dialog) throw new Error('The Save dialog needs Electron.')
  const owner = usableWindow(window)
  const result = owner ? await value.dialog.showSaveDialog(owner, options) : await value.dialog.showSaveDialog(options)
  if (!result || result.canceled || !result.filePath) return null
  return path.resolve(result.filePath)
}

async function showOpenDialogRaw(window, options) {
  const qa = qaOpenAnswer(options.defaultPath)
  if (qa.scripted) return qa.answer
  const value = electron()
  if (!value || !value.dialog) throw new Error('The Open dialog needs Electron.')
  const owner = usableWindow(window)
  const result = owner ? await value.dialog.showOpenDialog(owner, options) : await value.dialog.showOpenDialog(options)
  if (!result || result.canceled || !Array.isArray(result.filePaths)) return []
  return result.filePaths.map((item) => path.resolve(item))
}

/**
 * Shows the native Save dialog for Save As, Save a Copy, Export or the
 * fallback Save As after a failed save, and returns the final path and format
 * after the extension rules. Remembers the folder and grants the path to the
 * calling window. Nothing is written.
 *
 * @param {object} request usually straight from the renderer (`simpleIO.chooseSavePath`)
 * @param {'save'|'save-as'|'save-copy'|'export'|'fallback'} request.purpose
 * @param {string} [request.name] the document's name, used for the suggested file name
 * @param {string} [request.format] format id to start with (the document's format)
 * @param {string[]} [request.formats] format ids offered, current first (default: what the registry says is writable)
 * @param {string} [request.reason] failure code for purpose 'fallback' (LOCKED, CHANGED_ON_DISK, …)
 * @param {object} [context]
 * @param {object} [context.window] parent BrowserWindow
 * @param {object} [context.sender] calling webContents (receives the grant; the window is derived from it)
 * @param {string} [context.module] workspace (default: configureDialogs)
 * @param {string|null} [context.sourcePath] the document's file (main-side, from the registry)
 * @param {boolean} [context.engine] the optional office engine is available
 * @param {string} [context.title] dialog title
 * @returns {Promise<{path: string, format: string}|null>} null when canceled
 */
async function chooseSavePath(request = {}, context = {}) {
  const purpose = SAVE_PURPOSES.includes(request.purpose) ? request.purpose : 'save-as'
  const module = context.module || settings.module
  const requested = Array.isArray(request.formats) ? request.formats.filter((id) => typeof id === 'string' && formatEntry(id)) : []
  let writable = requested.length ? requested : writableFormats(module, purpose, { engine: context.engine })
  let format = typeof request.format === 'string' && formatEntry(request.format) ? request.format : writable[0]
  if (!format) return null
  if (!writable.includes(format)) writable = [format, ...writable]
  writable = [format, ...writable.filter((id) => id !== format)]
  const filters = writable.map(filterFor).filter(Boolean)
  const window = usableWindow(context.window) || windowForSender(context.sender)
  let defaultPath = await suggestedSavePath({
    purpose,
    format,
    name: typeof request.name === 'string' ? request.name : null,
    sourcePath: context.sourcePath || null,
    reason: request.reason,
  })
  for (let round = 0; round < MAX_DIALOG_ROUNDS; round += 1) {
    const options = { defaultPath, filters }
    if (context.title) options.title = String(context.title)
    const chosen = await showSaveDialogRaw(window, options)
    if (!chosen) return null
    const ruled = applyExtensionRules(chosen, { format, formats: writable })
    if (!samePath(ruled.path, chosen) && (await core.pathExists(ruled.path).catch(() => false))) {
      // The dialog only confirmed replacing the name it returned.
      const answer = await showPrompt('prompts.replace-existing', { finalName: path.basename(ruled.path) }, { window, module })
      if (answer !== 'replace') {
        defaultPath = ruled.path
        format = ruled.format
        continue
      }
    }
    grantSavePath(context.sender, ruled.path)
    await rememberFolder(purpose === 'export' ? 'export' : 'save', path.dirname(ruled.path))
    return { path: ruled.path, format: ruled.format }
  }
  return null
}

/**
 * Shows the native Open dialog: "All supported …" first, then the families,
 * then "All files", so mislabelled and extensionless files can be chosen too.
 *
 * @param {{multi?: boolean, purpose?: 'open'|'import'|'insert', filters?: object[]}} [request]
 * @param {{window?: object, sender?: object, module?: string, engine?: boolean, title?: string}} [context]
 * @returns {Promise<string[]>} absolute paths; empty when canceled
 */
async function chooseOpenPaths(request = {}, context = {}) {
  const purpose = OPEN_PURPOSES.includes(request.purpose) ? request.purpose : 'open'
  const module = context.module || settings.module
  let filters = Array.isArray(context.filters) ? context.filters : null
  if (!filters) {
    try {
      filters = module ? formats.dialogFilters(module, { engine: context.engine }) : null
    } catch {
      filters = null
    }
  }
  if (!filters || !filters.length) filters = [{ name: 'All files', extensions: ['*'] }]
  const window = usableWindow(context.window) || windowForSender(context.sender)
  const defaultPath = (await rememberedFolder('lastOpenFolder')) || documentsFolder()
  const properties = ['openFile']
  if (request.multi) properties.push('multiSelections')
  const options = { defaultPath, filters, properties }
  if (context.title) options.title = String(context.title)
  const paths = await showOpenDialogRaw(window, options)
  const chosen = request.multi ? paths : paths.slice(0, 1)
  if (chosen.length && purpose === 'open') await rememberFolder('open', path.dirname(chosen[0]))
  return chosen
}

/**
 * Shows a native decision box worded by io-catalog.json ("saveFailed.LOCKED",
 * "prompts.unsaved", …). The catalog's default button is focused (Enter) and
 * Esc chooses its cancel button. A saveFailed box ends its detail with
 * "Details: {technical}" when `vars.technical` is given.
 *
 * @param {string} key dotted catalog key of an entry with buttons
 * @param {Record<string, unknown>} [vars] placeholder values
 * @param {{window?: object, sender?: object, module?: string, signal?: AbortSignal, type?: string}} [context]
 *   signal closes the box as if it was canceled
 * @returns {Promise<string>} the id of the chosen button
 */
async function showPrompt(key, vars = {}, context = {}) {
  const entry = core.catalogEntry(key)
  if (!entry || typeof entry !== 'object' || !Array.isArray(entry.buttons) || !entry.buttons.length) {
    throw new TypeError(`"${key}" is not a prompt in io-catalog.json.`)
  }
  const module = context.module || settings.module
  const values = { appName: appNameFor(module), ...vars }
  const prompt = core.formatCatalogPrompt(key, values)
  const ids = prompt.buttons.map((button) => button.id)
  const cancelId = ids.includes(prompt.cancelId) ? prompt.cancelId : ids[ids.length - 1]
  const defaultId = ids.includes(prompt.defaultId) ? prompt.defaultId : ids[0]
  let detail = prompt.detail
  if (key.startsWith('saveFailed.') && values.technical) {
    detail = [detail, core.formatTemplate(core.catalog.detailsLine, { technical: values.technical })].filter(Boolean).join('\n\n')
  }
  const qa = qaPromptAnswer(key, { ...prompt, cancelId })
  if (qa.scripted) return qa.answer
  const value = electron()
  if (!value || !value.dialog) throw new Error('Prompts need Electron.')
  if (context.signal && context.signal.aborted) return cancelId
  const options = {
    type: context.type || (key.startsWith('saveFailed.') || WARNING_PROMPTS.has(key) ? 'warning' : 'question'),
    title: appNameFor(module),
    message: prompt.message,
    detail,
    buttons: prompt.buttons.map((button) => button.label),
    defaultId: ids.indexOf(defaultId),
    cancelId: ids.indexOf(cancelId),
    noLink: true,
  }
  if (context.signal) options.signal = context.signal
  const owner = usableWindow(context.window) || windowForSender(context.sender)
  const result = owner ? await value.dialog.showMessageBox(owner, options) : await value.dialog.showMessageBox(options)
  if (context.signal && context.signal.aborted) return cancelId
  const index = result && Number.isInteger(result.response) ? result.response : -1
  return ids[index] || cancelId
}

module.exports = {
  FOLDER_PREF_KEYS,
  OPEN_PURPOSES,
  SAVE_PURPOSES,
  appNameFor,
  applyExtensionRules,
  chooseOpenPaths,
  chooseSavePath,
  clearDecisions,
  configureDialogs,
  desktopFolder,
  documentsFolder,
  editedCopyPath,
  fallbackFolder,
  getPref,
  grantSavePath,
  hasDecision,
  isGranted,
  nameForFormat,
  primaryExtension,
  qaActive,
  recordDecision,
  rememberFolder,
  revokeGrants,
  safeStem,
  samePath,
  setPref,
  showPrompt,
  siblingPathFor,
  suggestedSavePath,
  writableFormats,
}
