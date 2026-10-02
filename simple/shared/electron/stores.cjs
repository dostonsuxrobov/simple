// Vendored from simple/shared/electron/stores.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// The two pruned stores the shared Save layer keeps in userData:
// - the recovery journal (design §4): one folder per open document,
//   recovery/<docId>/, holding generation-numbered parts and a manifest that
//   is written last. A crash at any point leaves the previous manifest
//   pointing at intact parts. Entries left by another process (a crash, a kill,
//   a power loss) or by a crashed window of this process are "orphaned" and
//   offered for recovery.
// - the versions store (design §3.8): before the first overwrite of a file in
//   each session, a copy of the file as it was on disk, so a save can always
//   be undone. Kept per file (10 versions), for 60 days and up to 500 MB.
// Both write only through safeWriteFile, and neither touches anything outside
// its own folder. Only Node built-ins, electron (loaded lazily) and sibling
// files are required.

const crypto = require('node:crypto')
const fs = require('node:fs')
const fsp = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const core = require('./io-core.cjs')
const { safeWriteFile } = require('./safe-write.cjs')

const MiB = 1024 * 1024
const DAY_MS = 24 * 60 * 60 * 1000

/** Manifest schema of a recovery entry. */
const RECOVERY_SCHEMA = 2
const MANIFEST_FILE = 'manifest.json'
const DAMAGED_FOLDER = '_damaged'
const ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9_-]{0,79}$/
const PART_PATTERN = /^[A-Za-z0-9][A-Za-z0-9._-]{0,79}$/
const PART_FILE_PATTERN = /^(.+)\.(\d+)$/
/** Snapshots above this size are refused; the renderer turns autosave off for such documents. */
const MAX_SNAPSHOT_BYTES = 512 * MiB
const MAX_EXTRA_BYTES = 256 * 1024
/** A folder without a manifest that is older than this is an abandoned first write. */
const ABANDONED_FOLDER_MS = 60 * 60 * 1000

const RECOVERY_DEFAULTS = Object.freeze({
  maxAgeMs: 30 * DAY_MS,
  maxTotalBytes: 2 * 1024 * MiB,
  damagedKeepMs: 7 * DAY_MS,
})

const VERSION_DEFAULTS = Object.freeze({
  maxPerFile: 10,
  maxTotalBytes: 500 * MiB,
  maxAgeMs: 60 * DAY_MS,
  maxFileBytes: 256 * MiB,
})

// "<ISO time>-<name>", or "<ISO time>_<n>-<name>" for a second copy in the same millisecond.
const VERSION_FILE_PATTERN = /^(\d{4}-\d{2}-\d{2}T\d{2}-\d{2}-\d{2}-\d{3}Z)(?:_(\d+))?-(.+)$/
const VERSION_ID_PATTERN = /^([0-9a-f]{40})\/([^\\/]+)$/
const SOURCE_INFO_FILE = 'source.json'
const CHUNK_BYTES = 4 * MiB

/** Identifies this process in recovery manifests; entries written under another token are orphaned. */
const PROCESS_TOKEN = crypto.randomBytes(8).toString('hex')

let logger = {
  warn: (...args) => console.warn('[simple-io]', ...args),
}

const configured = { recoveryDir: null, versionsDir: null }
let defaultRecovery = null
let defaultVersions = null

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function electronApp() {
  if (!process.versions.electron) return null
  try {
    const electron = require('electron')
    return electron && electron.app && typeof electron.app.getPath === 'function' ? electron.app : null
  } catch {
    return null
  }
}

function userDataFolder() {
  const app = electronApp()
  if (app) {
    try { return app.getPath('userData') } catch {}
  }
  return null
}

function appVersion() {
  const app = electronApp()
  try { return app && typeof app.getVersion === 'function' ? String(app.getVersion()) : null } catch { return null }
}

function toBuffer(data) {
  if (Buffer.isBuffer(data)) return data
  if (typeof data === 'string') return Buffer.from(data, 'utf8')
  if (ArrayBuffer.isView(data)) return Buffer.from(data.buffer, data.byteOffset, data.byteLength)
  if (data instanceof ArrayBuffer) return Buffer.from(data)
  throw new TypeError('A recovery part needs bytes (Buffer, Uint8Array, ArrayBuffer) or a string.')
}

function text(value, limit) {
  if (value === undefined || value === null) return null
  const result = String(value)
  return result.length > limit ? result.slice(0, limit) : result
}

/**
 * Normalises a file stamp, or returns null when it is not one.
 * @param {unknown} stamp
 * @returns {{size: number, mtimeMs: number, sha256?: string}|null}
 */
function cleanStamp(stamp) {
  if (!stamp || typeof stamp !== 'object') return null
  const size = Number(stamp.size)
  const mtimeMs = Number(stamp.mtimeMs)
  if (!Number.isFinite(size) || size < 0 || !Number.isFinite(mtimeMs)) return null
  const result = { size, mtimeMs }
  if (typeof stamp.sha256 === 'string' && /^[0-9a-f]{64}$/i.test(stamp.sha256)) result.sha256 = stamp.sha256.toLowerCase()
  return result
}

function checkRecoveryId(value) {
  const id = String(value ?? '')
  if (!ID_PATTERN.test(id)) throw new TypeError(`"${id.slice(0, 40)}" is not a valid recovery id.`)
  return id
}

/**
 * Lower-cased real path on Windows (the file system is case-insensitive), the
 * real path elsewhere; the resolved path when the file does not exist.
 * @param {string} filePath
 * @returns {string}
 */
function fileKey(filePath) {
  const resolved = path.resolve(String(filePath))
  let real = resolved
  try { real = fs.realpathSync.native(resolved) } catch {}
  return process.platform === 'win32' ? real.toLowerCase() : real
}

function isInside(child, parent) {
  const relative = path.relative(parent, child)
  return Boolean(relative) && !relative.startsWith('..') && !path.isAbsolute(relative)
}

async function readJsonFile(filePath) {
  const raw = await fsp.readFile(filePath, 'utf8')
  return JSON.parse(raw.replace(/^﻿/, ''))
}

async function folderSize(folder) {
  let total = 0
  let names = []
  try { names = await fsp.readdir(folder, { withFileTypes: true }) } catch { return 0 }
  for (const entry of names) {
    const full = path.join(folder, entry.name)
    try {
      if (entry.isDirectory()) total += await folderSize(full)
      else total += (await fsp.stat(full)).size
    } catch {}
  }
  return total
}

async function removeFolder(folder) {
  try {
    await fsp.rm(folder, { recursive: true, force: true, maxRetries: 3, retryDelay: 100 })
    return true
  } catch (error) {
    logger.warn(`Could not remove ${folder}: ${error && error.message}`)
    return false
  }
}

function jsonText(value) {
  return `${JSON.stringify(value, null, 2)}\n`
}

/**
 * Writes a small JSON file inside a store through safeWriteFile.
 * @param {string} filePath
 * @param {unknown} value
 */
async function writeJson(filePath, value) {
  await fsp.mkdir(path.dirname(filePath), { recursive: true })
  await safeWriteFile(filePath, jsonText(value), { format: 'json' })
}

// ---------------------------------------------------------------------------
// Recovery journal
// ---------------------------------------------------------------------------

function defaultRecoveryRoot() {
  if (configured.recoveryDir) return configured.recoveryDir
  if (process.env.SIMPLE_RECOVERY_DIR) return path.resolve(process.env.SIMPLE_RECOVERY_DIR)
  const userData = userDataFolder()
  if (userData) return path.join(userData, 'recovery')
  return path.join(os.tmpdir(), 'simple-recovery')
}

function normalizeSnapshot(snapshot) {
  if (!snapshot || typeof snapshot !== 'object') throw new TypeError('A recovery snapshot must be an object.')
  const docId = checkRecoveryId(snapshot.docId)
  const revision = Number(snapshot.revision)
  if (!Number.isSafeInteger(revision) || revision < 0) throw new TypeError('A recovery snapshot needs a whole revision number of 0 or more.')
  if (!Array.isArray(snapshot.parts) || !snapshot.parts.length) throw new TypeError('A recovery snapshot needs at least one part.')
  const names = new Set()
  const parts = snapshot.parts.map((part) => {
    const name = String((part && part.name) || '')
    if (!PART_PATTERN.test(name) || name.toLowerCase() === MANIFEST_FILE || names.has(name.toLowerCase())) {
      throw new TypeError(`"${name.slice(0, 40)}" is not a valid or unique recovery part name.`)
    }
    names.add(name.toLowerCase())
    if (part.ref) {
      const refPath = part.ref && typeof part.ref.path === 'string' ? part.ref.path : null
      if (!refPath || !path.isAbsolute(refPath)) throw new TypeError(`Recovery part "${name}" refers to a file without an absolute path.`)
      return { name, ref: { path: path.resolve(refPath), stamp: cleanStamp(part.ref.stamp) }, required: Boolean(part.required) }
    }
    if (part.data === undefined || part.data === null) throw new TypeError(`Recovery part "${name}" has no data.`)
    return { name, data: toBuffer(part.data), text: typeof part.data === 'string', compress: part.compress === true }
  })
  let extra = null
  const extraInput = snapshot.extra !== undefined && snapshot.extra !== null ? snapshot.extra : snapshot.meta
  if (extraInput !== undefined && extraInput !== null) {
    const encoded = JSON.stringify(extraInput)
    if (encoded === undefined || Buffer.byteLength(encoded) > MAX_EXTRA_BYTES) throw new TypeError('Recovery "meta" data must be small JSON.')
    extra = JSON.parse(encoded)
  }
  return {
    docId,
    revision,
    parts,
    module: text(snapshot.module, 32),
    title: text(snapshot.title, 260),
    kind: text(snapshot.kind, 40),
    format: text(snapshot.format, 40),
    sourcePath: snapshot.sourcePath ? path.resolve(String(snapshot.sourcePath)) : null,
    sourceStamp: cleanStamp(snapshot.sourceStamp),
    suggestedPath: snapshot.suggestedPath ? path.resolve(String(snapshot.suggestedPath)) : null,
    appVersion: text(snapshot.appVersion, 40) || appVersion(),
    extra,
  }
}

function isManifest(value) {
  return Boolean(value && typeof value === 'object' && value.schema === RECOVERY_SCHEMA && typeof value.docId === 'string'
    && Number.isSafeInteger(value.revision) && Number.isSafeInteger(value.generation) && Array.isArray(value.parts))
}

/**
 * Creates a recovery journal over one folder. Every method is safe to call
 * concurrently: work on one document runs in call order.
 *
 * @param {object} [options]
 * @param {string} [options.root] folder of the journal; default SIMPLE_RECOVERY_DIR, else <userData>/recovery
 * @param {string} [options.processToken] identity of the writing process (tests simulate other processes)
 * @param {{beforeManifestWrite?: (info: {docId: string, generation: number}) => Promise<void>|void}} [options.testHooks]
 * @returns {RecoveryStore}
 */
function createRecoveryStore(options = {}) {
  const token = options.processToken || PROCESS_TOKEN
  const hooks = options.testHooks || {}
  const queues = new Map()
  const discardedUpTo = new Map()
  const orphanedHere = new Set()
  const claimed = new Set()
  const lastWriteAt = new Map()

  function root() {
    return path.resolve(options.root || defaultRecoveryRoot())
  }

  function entryFolder(docId) {
    return path.join(root(), checkRecoveryId(docId))
  }

  function enqueue(docId, task) {
    const previous = queues.get(docId) || Promise.resolve()
    const run = previous.then(task)
    const settled = run.then(() => {}, () => {})
    queues.set(docId, settled)
    settled.then(() => { if (queues.get(docId) === settled) queues.delete(docId) })
    return run
  }

  /** @returns {Promise<null|{manifest: object}|{damaged: true, reason: string}>} */
  async function readManifest(folder) {
    let value
    try {
      value = await readJsonFile(path.join(folder, MANIFEST_FILE))
    } catch (error) {
      if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return null
      return { damaged: true, reason: error && error.message ? error.message : 'unreadable' }
    }
    if (!isManifest(value)) return { damaged: true, reason: 'not a recovery manifest' }
    return { manifest: value }
  }

  async function quarantine(folder, docId) {
    const target = path.join(root(), DAMAGED_FOLDER, `${docId}-${Date.now()}`)
    try {
      await fsp.mkdir(path.dirname(target), { recursive: true })
      await fsp.rename(folder, target)
      logger.warn(`Moved an unreadable recovery entry aside: ${target}`)
    } catch (error) {
      logger.warn(`Could not move the unreadable recovery entry ${folder} aside: ${error && error.message}`)
    }
  }

  async function partFileIntact(folder, part) {
    try {
      return (await fsp.stat(path.join(folder, part.file))).size === part.size
    } catch {
      return false
    }
  }

  async function removeOldGenerations(folder, manifest) {
    const keep = new Set(manifest.parts.filter((part) => part.file).map((part) => part.file.toLowerCase()))
    for (const part of (manifest.previous && manifest.previous.parts) || []) if (part.file) keep.add(part.file.toLowerCase())
    let names = []
    try { names = await fsp.readdir(folder) } catch { return }
    for (const name of names) {
      if (name === MANIFEST_FILE || keep.has(name.toLowerCase()) || !PART_FILE_PATTERN.test(name)) continue
      await core.removeFile(path.join(folder, name))
    }
  }

  function isOrphaned(manifest) {
    if (claimed.has(manifest.docId)) return false
    return orphanedHere.has(manifest.docId) || manifest.state === 'orphaned' || manifest.owner !== token
  }

  async function referenceState(ref) {
    if (!ref || !ref.path) return null
    if (!ref.stamp) return (await core.pathExists(ref.path).catch(() => false)) ? 'unknown' : 'missing'
    try {
      const comparison = await core.compareStamp(ref.path, ref.stamp)
      return comparison.state === 'restamped' ? 'same' : comparison.state
    } catch {
      return 'unknown'
    }
  }

  async function describe(manifest, folder) {
    const sourceState = manifest.sourcePath ? await referenceState({ path: manifest.sourcePath, stamp: manifest.sourceStamp }) : null
    let restorable = true
    let reason = null
    const refs = {}
    for (const part of manifest.parts) {
      if (!part.ref) continue
      const state = await referenceState(part.ref)
      refs[part.name] = { path: part.ref.path, stamp: part.ref.stamp || null, state, required: Boolean(part.required) }
      if (part.required && state !== 'same') {
        restorable = false
        reason = state === 'missing' ? 'base-missing' : 'base-changed'
      }
    }
    const orphaned = isOrphaned(manifest)
    return {
      id: manifest.docId,
      docId: manifest.docId,
      module: manifest.module || null,
      title: manifest.title || null,
      kind: manifest.kind || null,
      format: manifest.format || null,
      sourcePath: manifest.sourcePath || null,
      suggestedPath: manifest.suggestedPath || null,
      sourceStamp: manifest.sourceStamp || null,
      sourceState,
      sourceChanged: sourceState === 'changed',
      sourceMissing: sourceState === 'missing',
      revision: manifest.revision,
      generation: manifest.generation,
      createdAt: manifest.createdAt || null,
      updatedAt: manifest.updatedAt || null,
      offeredAt: manifest.offeredAt || null,
      orphaned,
      state: orphaned ? 'orphaned' : 'live',
      restorable,
      reason,
      refs,
      size: manifest.parts.reduce((sum, part) => sum + (part.size || 0), 0),
      folder,
      migratedFrom: manifest.migratedFrom || null,
    }
  }

  async function rewriteManifest(docId, patch) {
    const folder = entryFolder(docId)
    const current = await readManifest(folder)
    if (!current || !current.manifest) return false
    await writeJson(path.join(folder, MANIFEST_FILE), { ...current.manifest, ...patch })
    return true
  }

  /**
   * Writes one snapshot of a document. Parts whose bytes did not change since
   * the previous generation are kept, not rewritten. A snapshot whose revision
   * is not newer than a discard, or older than the stored one, is dropped.
   *
   * @param {RecoverySnapshot} snapshot
   * @returns {Promise<{ok: true, docId: string, generation: number, revision: number, written: string[], reused: string[]}
   *   | {ok: false, docId: string, dropped: 'discarded'|'stale'}
   *   | {ok: false, docId: string, code: string, message: string}>}
   */
  async function write(snapshot) {
    const input = normalizeSnapshot(snapshot)
    const total = input.parts.reduce((sum, part) => sum + (part.data ? part.data.length : 0), 0)
    if (total > MAX_SNAPSHOT_BYTES) {
      return { ok: false, docId: input.docId, code: 'TOO_LARGE', message: core.catalog.status.autosaveOff, technical: `${total} bytes is over the ${MAX_SNAPSHOT_BYTES}-byte recovery limit` }
    }
    return enqueue(input.docId, async () => {
      const floor = discardedUpTo.get(input.docId)
      if (floor !== undefined && input.revision <= floor) return { ok: false, docId: input.docId, dropped: 'discarded' }
      const folder = entryFolder(input.docId)
      const current = await readManifest(folder)
      const previous = current && current.manifest ? current.manifest : null
      if (previous && previous.revision > input.revision) return { ok: false, docId: input.docId, dropped: 'stale' }
      const generation = (previous ? previous.generation : 0) + 1
      try {
        await fsp.mkdir(folder, { recursive: true })
        const parts = []
        const written = []
        const reused = []
        for (const part of input.parts) {
          if (part.ref) {
            parts.push({ name: part.name, ref: part.ref, required: part.required })
            continue
          }
          // Reuse compares the content as given, so a recompressed part is not rewritten either.
          const rawSha256 = core.hashBytes(part.data)
          const encoding = part.compress ? 'gzip' : null
          const old = previous && previous.parts.find((item) => item.name === part.name && item.file)
          if (old && (old.rawSha256 || old.sha256) === rawSha256 && (old.encoding || null) === encoding && Boolean(old.text) === part.text
            && (await partFileIntact(folder, old))) {
            parts.push({ ...old })
            reused.push(part.name)
            continue
          }
          const stored = encoding === 'gzip' ? zlib.gzipSync(part.data) : part.data
          const file = `${part.name}.${generation}`
          await safeWriteFile(path.join(folder, file), stored)
          const entry = { name: part.name, file, size: stored.length, sha256: core.hashBytes(stored) }
          if (encoding) {
            entry.encoding = encoding
            entry.rawSha256 = rawSha256
            entry.rawSize = part.data.length
          }
          if (part.text) entry.text = true
          parts.push(entry)
          written.push(part.name)
        }
        if (typeof hooks.beforeManifestWrite === 'function') await hooks.beforeManifestWrite({ docId: input.docId, generation })
        const now = new Date().toISOString()
        const manifest = {
          schema: RECOVERY_SCHEMA,
          docId: input.docId,
          module: input.module,
          title: input.title,
          kind: input.kind,
          format: input.format,
          sourcePath: input.sourcePath,
          sourceStamp: input.sourceStamp,
          suggestedPath: input.suggestedPath,
          revision: input.revision,
          generation,
          createdAt: previous && previous.createdAt ? previous.createdAt : now,
          updatedAt: now,
          offeredAt: null,
          appVersion: input.appVersion,
          pid: process.pid,
          owner: token,
          state: 'live',
          extra: input.extra,
          parts,
          previous: previous ? { generation: previous.generation, revision: previous.revision, updatedAt: previous.updatedAt || null, parts: previous.parts } : null,
        }
        await writeJson(path.join(folder, MANIFEST_FILE), manifest)
        lastWriteAt.set(input.docId, Date.now())
        orphanedHere.delete(input.docId)
        claimed.add(input.docId)
        await removeOldGenerations(folder, manifest)
        return { ok: true, docId: input.docId, generation, revision: input.revision, written, reused }
      } catch (error) {
        const result = core.toIoResult(error, { path: folder })
        logger.warn(`Could not write a recovery copy of ${input.title || input.docId}: ${result.technical || result.message}`)
        return { ...result, docId: input.docId }
      }
    })
  }

  async function loadParts(folder, parts) {
    const loaded = []
    for (const part of parts) {
      if (part.ref) {
        const state = await referenceState(part.ref)
        loaded.push({ name: part.name, ref: { path: part.ref.path, stamp: part.ref.stamp || null, state, matches: state === 'same' }, required: Boolean(part.required) })
        continue
      }
      if (!part.file) continue
      if (!PART_FILE_PATTERN.test(part.file) || part.file.includes('/') || part.file.includes('\\')) return { ok: false, reason: `invalid part file ${part.file}` }
      let bytes
      try {
        bytes = await fsp.readFile(path.join(folder, part.file))
      } catch (error) {
        return { ok: false, reason: `${part.name}: ${error && error.code ? error.code : 'unreadable'}` }
      }
      if (bytes.length !== part.size || core.hashBytes(bytes) !== part.sha256) return { ok: false, reason: `${part.name}: content does not match its checksum` }
      if (part.encoding === 'gzip') {
        try {
          bytes = zlib.gunzipSync(bytes)
        } catch {
          return { ok: false, reason: `${part.name}: compressed content is damaged` }
        }
        if (part.rawSha256 && core.hashBytes(bytes) !== part.rawSha256) return { ok: false, reason: `${part.name}: content does not match its checksum` }
      }
      loaded.push({ name: part.name, data: part.text ? bytes.toString('utf8') : bytes })
    }
    return { ok: true, parts: loaded }
  }

  /**
   * Reads an entry for restoring. Every part is checked against its hash; when
   * the newest generation is damaged, the previous generation is returned.
   * Parts come back in snapshot order: bytes as Buffers, text written as a
   * string as a string, and file references with whether the file still matches.
   *
   * @param {string} docId
   * @returns {Promise<{ok: true, entry: RecoveryEntry, revision: number, generation: number, fellBack: boolean,
   *   parts: Array<{name: string, data?: Buffer|string, ref?: {path: string, stamp: object|null, state: string|null, matches: boolean}, required?: boolean}>,
   *   extra: unknown} | {ok: false, code: 'NOT_FOUND'|'DAMAGED', reason?: string}>}
   */
  async function read(docId) {
    const id = checkRecoveryId(docId)
    return enqueue(id, async () => {
      const folder = entryFolder(id)
      const current = await readManifest(folder)
      if (!current) return { ok: false, code: 'NOT_FOUND' }
      if (current.damaged) {
        await quarantine(folder, id)
        return { ok: false, code: 'DAMAGED', reason: current.damaged && current.reason }
      }
      const manifest = current.manifest
      let loaded = await loadParts(folder, manifest.parts)
      let revision = manifest.revision
      let generation = manifest.generation
      let fellBack = false
      if (!loaded.ok && manifest.previous && Array.isArray(manifest.previous.parts)) {
        const earlier = await loadParts(folder, manifest.previous.parts)
        if (earlier.ok) {
          logger.warn(`The newest recovery copy of ${manifest.title || id} was damaged (${loaded.reason}); using the one before it.`)
          loaded = earlier
          revision = manifest.previous.revision
          generation = manifest.previous.generation
          fellBack = true
        }
      }
      if (!loaded.ok) return { ok: false, code: 'DAMAGED', reason: loaded.reason }
      return { ok: true, entry: await describe(manifest, folder), revision, generation, fellBack, parts: loaded.parts, extra: manifest.extra ?? null }
    })
  }

  /**
   * Lists recovery entries, newest first. By default only orphaned entries
   * (left by another process, or by a crashed window of this one) are listed,
   * and each is marked as offered. Unreadable entries are moved to _damaged.
   *
   * @param {{includeLive?: boolean, markOffered?: boolean}} [options]
   * @returns {Promise<RecoveryEntry[]>}
   */
  async function list(options = {}) {
    const folder = root()
    let names = []
    try {
      names = await fsp.readdir(folder, { withFileTypes: true })
    } catch (error) {
      if (error && error.code === 'ENOENT') return []
      throw error
    }
    const entries = []
    for (const dirent of names) {
      if (!dirent.isDirectory() || dirent.name === DAMAGED_FOLDER || !ID_PATTERN.test(dirent.name)) continue
      const id = dirent.name
      if (queues.has(id)) await queues.get(id)
      const entryPath = path.join(folder, id)
      const current = await readManifest(entryPath)
      if (!current) {
        // A first write that never reached its manifest has nothing restorable.
        try {
          const stat = await fsp.stat(entryPath)
          if (Date.now() - stat.mtimeMs > ABANDONED_FOLDER_MS && !queues.has(id)) await quarantine(entryPath, id)
        } catch {}
        continue
      }
      if (current.damaged) {
        await quarantine(entryPath, id)
        continue
      }
      const described = await describe(current.manifest, entryPath)
      if (!described.orphaned && !options.includeLive) continue
      entries.push(described)
    }
    entries.sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)))
    if (options.markOffered !== false) {
      const now = new Date().toISOString()
      for (const entry of entries) {
        if (!entry.orphaned || entry.offeredAt) continue
        entry.offeredAt = now
        await enqueue(entry.id, () => rewriteManifest(entry.id, { offeredAt: now })).catch((error) => {
          logger.warn(`Could not mark the recovery entry ${entry.id} as offered: ${error && error.message}`)
        })
      }
    }
    return entries
  }

  /**
   * Drops an entry. With `upToRevision`, an entry holding newer edits is kept;
   * either way, later writes at or below the discarded revision are ignored,
   * so a snapshot that finishes after a save cannot bring back a stale entry.
   *
   * @param {string} docId
   * @param {number} [upToRevision]
   * @returns {Promise<{ok: true, removed: boolean, kept?: boolean}>}
   */
  function discard(docId, upToRevision) {
    const id = checkRecoveryId(docId)
    const bound = upToRevision === undefined || upToRevision === null ? null : Number(upToRevision)
    if (bound !== null && (!Number.isSafeInteger(bound) || bound < 0)) throw new TypeError('upToRevision must be a whole number of 0 or more.')
    return enqueue(id, async () => {
      const folder = entryFolder(id)
      const current = await readManifest(folder)
      const stored = current && current.manifest ? current.manifest.revision : null
      const floor = bound !== null ? bound : stored
      if (floor !== null && floor !== undefined) discardedUpTo.set(id, Math.max(discardedUpTo.get(id) ?? -1, floor))
      if (current && current.manifest && bound !== null && current.manifest.revision > bound) return { ok: true, removed: false, kept: true }
      const existed = Boolean(current)
      if (existed) await removeFolder(folder)
      orphanedHere.delete(id)
      claimed.delete(id)
      lastWriteAt.delete(id)
      return { ok: true, removed: existed }
    })
  }

  /**
   * Marks entries orphaned (their window crashed or was closed while hung) so
   * that the other windows of this process offer them.
   * @param {string[]} docIds
   * @returns {Promise<void>}
   */
  async function markOrphaned(docIds) {
    const ids = (Array.isArray(docIds) ? docIds : [docIds]).filter((id) => typeof id === 'string' && ID_PATTERN.test(id))
    for (const id of ids) {
      orphanedHere.add(id)
      claimed.delete(id)
    }
    await Promise.all(ids.map((id) => enqueue(id, () => rewriteManifest(id, { state: 'orphaned' })).catch((error) => {
      logger.warn(`Could not mark the recovery entry ${id} as orphaned: ${error && error.message}`)
    })))
  }

  /**
   * Takes an orphaned entry over for a window of this process (after it was
   * restored), so no other window offers it again.
   * @param {string} docId
   * @returns {Promise<boolean>} false when there is no such entry
   */
  async function claim(docId) {
    const id = checkRecoveryId(docId)
    claimed.add(id)
    orphanedHere.delete(id)
    return enqueue(id, () => rewriteManifest(id, { owner: token, pid: process.pid, state: 'live' }))
  }

  /**
   * Startup cleanup: entries older than 30 days that were offered at least
   * once, then the oldest offered entries while the store is over 2 GB, and
   * damaged entries after 7 days. Entries that were never offered are kept.
   *
   * @param {{now?: number, maxAgeMs?: number, maxTotalBytes?: number, damagedKeepMs?: number}} [options]
   * @returns {Promise<{removed: string[], damagedRemoved: number, totalBytes: number}>}
   */
  async function prune(options = {}) {
    const now = options.now ?? Date.now()
    const maxAgeMs = options.maxAgeMs ?? RECOVERY_DEFAULTS.maxAgeMs
    const maxTotalBytes = options.maxTotalBytes ?? RECOVERY_DEFAULTS.maxTotalBytes
    const damagedKeepMs = options.damagedKeepMs ?? RECOVERY_DEFAULTS.damagedKeepMs
    const folder = root()
    const report = { removed: [], damagedRemoved: 0, totalBytes: 0 }
    const damagedRoot = path.join(folder, DAMAGED_FOLDER)
    let damaged = []
    try { damaged = await fsp.readdir(damagedRoot, { withFileTypes: true }) } catch {}
    for (const dirent of damaged) {
      const full = path.join(damagedRoot, dirent.name)
      try {
        if (now - (await fsp.stat(full)).mtimeMs > damagedKeepMs && (await removeFolder(full))) report.damagedRemoved += 1
      } catch {}
    }
    const entries = await list({ includeLive: true, markOffered: false })
    const candidates = []
    for (const entry of entries) {
      const size = await folderSize(entry.folder)
      report.totalBytes += size
      if (!entry.orphaned || !entry.offeredAt) continue
      candidates.push({ entry, size, updated: Date.parse(entry.updatedAt) || 0 })
    }
    candidates.sort((a, b) => a.updated - b.updated)
    for (const candidate of candidates) {
      const tooOld = now - candidate.updated > maxAgeMs
      if (!tooOld && report.totalBytes <= maxTotalBytes) continue
      const removed = await enqueue(candidate.entry.id, () => removeFolder(candidate.entry.folder))
      if (removed) {
        report.removed.push(candidate.entry.id)
        report.totalBytes -= candidate.size
      }
    }
    return report
  }

  /**
   * Moves the Documents workspace's old recovery layout (recoveries.json plus
   * recovery/<id>.docx) into this store. Migrated entries have no source
   * stamp, so they restore as untitled documents, as before; their old path is
   * kept as `suggestedPath` for the Save As dialog. Safe to run again: entries
   * already migrated are skipped and the old files are removed last.
   *
   * @param {object} [options]
   * @param {string} [options.userData] folder holding recoveries.json (default: Electron's userData,
   *   else the parent of the store)
   * @param {string} [options.indexFile] default <userData>/recoveries.json
   * @param {string} [options.legacyFolder] default <userData>/recovery
   * @returns {Promise<{migrated: number, skipped: number, indexRemoved: boolean}>}
   */
  async function migrateDocsLegacy(options = {}) {
    const userData = options.userData ? path.resolve(options.userData) : (userDataFolder() || path.dirname(root()))
    const indexFile = options.indexFile ? path.resolve(options.indexFile) : path.join(userData, 'recoveries.json')
    const legacyFolder = options.legacyFolder ? path.resolve(options.legacyFolder) : path.join(userData, 'recovery')
    const report = { migrated: 0, skipped: 0, indexRemoved: false }
    let index
    try {
      index = await readJsonFile(indexFile)
    } catch (error) {
      if (error && error.code === 'ENOENT') return report
      logger.warn(`Could not read the old recovery index ${indexFile}: ${error && error.message}`)
      return report
    }
    if (!Array.isArray(index)) index = []
    let complete = true
    for (const item of index) {
      const id = String((item && item.id) || '').replace(/[^a-zA-Z0-9-]/g, '')
      if (!id || !ID_PATTERN.test(id)) {
        report.skipped += 1
        continue
      }
      const legacyFile = path.join(legacyFolder, `${id}.docx`)
      try {
        const done = await enqueue(id, async () => {
          const folder = entryFolder(id)
          const existing = await readManifest(folder)
          if (existing && existing.manifest) {
            await core.removeFile(legacyFile)
            return false
          }
          let bytes
          try {
            bytes = await fsp.readFile(legacyFile)
          } catch (error) {
            if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return false
            throw error
          }
          await fsp.mkdir(folder, { recursive: true })
          const file = 'document.docx.1'
          await safeWriteFile(path.join(folder, file), bytes)
          const updated = Number(item.updatedAt)
          const when = new Date(Number.isFinite(updated) && updated > 0 ? updated : Date.now()).toISOString()
          const stem = text(item.title, 200) || 'Recovered document'
          await writeJson(path.join(folder, MANIFEST_FILE), {
            schema: RECOVERY_SCHEMA,
            docId: id,
            module: 'docs',
            title: `${stem}.docx`,
            kind: 'document',
            format: 'docx',
            sourcePath: null,
            sourceStamp: null,
            suggestedPath: item.sourcePath ? path.resolve(String(item.sourcePath)) : null,
            revision: 0,
            generation: 1,
            createdAt: when,
            updatedAt: when,
            offeredAt: null,
            appVersion: null,
            pid: null,
            owner: null,
            state: 'orphaned',
            extra: null,
            migratedFrom: 'docs-recoveries-json',
            parts: [{ name: 'document.docx', file, size: bytes.length, sha256: core.hashBytes(bytes) }],
            previous: null,
          })
          await core.removeFile(legacyFile)
          return true
        })
        if (done) report.migrated += 1
        else report.skipped += 1
      } catch (error) {
        complete = false
        logger.warn(`Could not move the old recovery copy ${legacyFile}: ${error && error.message}`)
      }
    }
    if (complete) report.indexRemoved = await core.removeFile(indexFile)
    return report
  }

  /**
   * When the newest snapshot of any of these documents was written in this
   * process, for the "Changes up to {time}" wording of the guard's prompts.
   * @param {string[]} docIds
   * @returns {number|null} epoch milliseconds
   */
  function lastWriteFor(docIds) {
    let latest = null
    for (const id of docIds || []) {
      const at = lastWriteAt.get(id)
      if (at && (!latest || at > latest)) latest = at
    }
    return latest
  }

  return Object.freeze({
    root,
    entryFolder,
    write,
    read,
    list,
    discard,
    markOrphaned,
    claim,
    prune,
    migrateDocsLegacy,
    lastWriteFor,
    /** @returns {string} the token stamped into manifests written by this store */
    processToken: () => token,
  })
}

// ---------------------------------------------------------------------------
// Versions store
// ---------------------------------------------------------------------------

function defaultVersionsRoot() {
  if (configured.versionsDir) return configured.versionsDir
  const userData = userDataFolder()
  if (userData) return path.join(userData, 'versions')
  return path.join(os.tmpdir(), 'simple-versions')
}

function versionStamp(date = new Date()) {
  return date.toISOString().replace(/[:.]/g, '-')
}

function trimmedName(name) {
  const clean = String(name || 'file').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_')
  if (clean.length <= 120) return clean
  const extension = path.extname(clean).slice(0, 16)
  return `${clean.slice(0, 120 - extension.length)}${extension}`
}

function parseVersionName(name) {
  const match = VERSION_FILE_PATTERN.exec(name)
  if (!match) return null
  const iso = match[1].replace(/^(\d{4}-\d{2}-\d{2}T\d{2})-(\d{2})-(\d{2})-(\d{3})Z$/, '$1:$2:$3.$4Z')
  return { createdAt: iso, time: Date.parse(iso) || 0, sequence: match[2] ? Number(match[2]) : 0, name: match[3] }
}

function newestFirst(a, b) {
  return (b.time - a.time) || (b.sequence - a.sequence)
}

/**
 * Creates a versions store over one folder.
 *
 * @param {object} [options]
 * @param {string} [options.root] default <userData>/versions
 * @param {number} [options.maxPerFile=10] versions kept per file
 * @param {number} [options.maxTotalBytes=500 MB] size of the whole store
 * @param {number} [options.maxAgeMs=60 days] older versions are removed
 * @param {number} [options.maxFileBytes=256 MB] larger files are not copied
 * @returns {VersionsStore}
 */
function createVersionsStore(options = {}) {
  const limits = {
    maxPerFile: options.maxPerFile ?? VERSION_DEFAULTS.maxPerFile,
    maxTotalBytes: options.maxTotalBytes ?? VERSION_DEFAULTS.maxTotalBytes,
    maxAgeMs: options.maxAgeMs ?? VERSION_DEFAULTS.maxAgeMs,
    maxFileBytes: options.maxFileBytes ?? VERSION_DEFAULTS.maxFileBytes,
  }
  const backedUp = new Set()
  const inFlight = new Map()
  let pruneQueue = Promise.resolve()

  function root() {
    return path.resolve(options.root || defaultVersionsRoot())
  }

  /**
   * The folder name of a file's versions: SHA-1 of its lower-cased real path.
   * @param {string} filePath
   * @returns {string}
   */
  function keyFor(filePath) {
    return crypto.createHash('sha1').update(fileKey(filePath)).digest('hex')
  }

  async function versionsIn(folder) {
    let names = []
    try { names = await fsp.readdir(folder) } catch { return [] }
    const versions = []
    for (const name of names) {
      const parsed = parseVersionName(name)
      if (!parsed) continue
      try {
        const stat = await fsp.stat(path.join(folder, name))
        if (stat.isFile()) versions.push({ file: name, path: path.join(folder, name), size: stat.size, ...parsed })
      } catch {}
    }
    versions.sort(newestFirst)
    return versions
  }

  async function copyIntoStore(source, stat) {
    const folder = path.join(root(), keyFor(source))
    await fsp.mkdir(folder, { recursive: true })
    const infoPath = path.join(folder, SOURCE_INFO_FILE)
    try {
      const info = await readJsonFile(infoPath).catch(() => null)
      if (!info || info.path !== source) await writeJson(infoPath, { path: source, name: path.basename(source) })
    } catch (error) {
      logger.warn(`Could not record the source of ${folder}: ${error && error.message}`)
    }
    const stamp = versionStamp()
    let fileName = `${stamp}-${trimmedName(path.basename(source))}`
    for (let sequence = 2; await core.pathExists(path.join(folder, fileName)); sequence += 1) {
      fileName = `${stamp}_${sequence}-${trimmedName(path.basename(source))}`
    }
    const target = path.join(folder, fileName)
    const producer = async (sink) => {
      const handle = await fsp.open(source, 'r')
      try {
        const buffer = Buffer.allocUnsafe(CHUNK_BYTES)
        for (let position = 0; ;) {
          const { bytesRead } = await handle.read(buffer, 0, buffer.length, position)
          if (!bytesRead) break
          await sink.write(Buffer.from(buffer.subarray(0, bytesRead)))
          position += bytesRead
        }
      } finally {
        await handle.close().catch(() => {})
      }
    }
    const written = await safeWriteFile(target, producer, { expectedSize: stat.size })
    return { id: `${path.basename(folder)}/${fileName}`, path: target, size: written.stamp.size, folder, file: fileName }
  }

  async function pruneFolder(folder, keepFile, now = Date.now()) {
    const versions = await versionsIn(folder)
    let kept = 0
    for (const version of versions) {
      const age = now - (Date.parse(version.createdAt) || now)
      const keep = version.file === keepFile || (kept < limits.maxPerFile && age <= limits.maxAgeMs)
      if (keep) {
        kept += 1
        continue
      }
      await core.removeFile(version.path)
    }
    if (!kept) await removeFolder(folder)
  }

  /**
   * Copies a file into the store now, regardless of earlier copies this session.
   * Never throws: a failed copy is logged and reported.
   *
   * @param {string} filePath
   * @returns {Promise<{ok: true, id: string, path: string, size: number}
   *   | {ok: true, skipped: 'missing'|'not-a-file'}
   *   | {ok: false, skipped?: 'too-large', reason: string}>}
   */
  async function backup(filePath) {
    const source = path.resolve(String(filePath))
    let stat
    try {
      stat = await fsp.stat(source)
    } catch (error) {
      if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return { ok: true, skipped: 'missing' }
      const reason = core.toIoError(error, { path: source, stage: 'read', context: 'open' }).message
      logger.warn(`No backup copy of ${source}: ${reason}`)
      return { ok: false, reason }
    }
    if (!stat.isFile()) return { ok: true, skipped: 'not-a-file' }
    if (stat.size > limits.maxFileBytes) {
      const reason = `the file is larger than ${core.formatBytes(limits.maxFileBytes)}`
      logger.warn(`No backup copy of ${source}: ${reason}`)
      return { ok: false, skipped: 'too-large', reason }
    }
    try {
      const copy = await copyIntoStore(source, stat)
      await pruneFolder(copy.folder, copy.file)
      schedulePrune()
      return { ok: true, id: copy.id, path: copy.path, size: copy.size }
    } catch (error) {
      const ioError = core.toIoError(error, { path: source, stage: 'read', context: 'open' })
      const reason = ioError.code === 'LOCKED' || ioError.code === 'NO_PERMISSION' || ioError.code === 'FILE_UNAVAILABLE'
        ? 'another program is using this file'
        : (ioError.technical || ioError.message)
      logger.warn(`No backup copy of ${source}: ${ioError.technical || ioError.message}`)
      return { ok: false, reason }
    }
  }

  /**
   * Copies a file into the store before its first overwrite in this session;
   * later calls for the same file do nothing. Never throws and never blocks a
   * save: a failure is logged and reported.
   *
   * @param {string} filePath
   * @returns {ReturnType<typeof backup> | Promise<{ok: true, skipped: 'already-backed-up'}>}
   */
  async function backupOnce(filePath) {
    const key = fileKey(filePath)
    if (backedUp.has(key)) return { ok: true, skipped: 'already-backed-up' }
    if (inFlight.has(key)) {
      await inFlight.get(key)
      return { ok: true, skipped: 'already-backed-up' }
    }
    const run = backup(filePath)
    inFlight.set(key, run)
    try {
      const result = await run
      if (result.ok) backedUp.add(key)
      return result
    } finally {
      inFlight.delete(key)
    }
  }

  /**
   * @param {string} filePath
   * @returns {boolean} whether this session already copied the file
   */
  function hasBackup(filePath) {
    return backedUp.has(fileKey(filePath))
  }

  /**
   * The stored versions of a file, newest first.
   * @param {string} filePath
   * @returns {Promise<VersionEntry[]>}
   */
  async function list(filePath) {
    const source = path.resolve(String(filePath))
    const folder = path.join(root(), keyFor(source))
    return (await versionsIn(folder)).map((version) => ({
      id: `${path.basename(folder)}/${version.file}`,
      path: version.path,
      name: version.name,
      sourcePath: source,
      createdAt: version.createdAt,
      size: version.size,
    }))
  }

  /**
   * The file behind a version id from list(), or null for an unknown or
   * malformed id. Never resolves outside the store.
   * @param {string} id
   * @returns {Promise<{path: string, name: string, createdAt: string, size: number, sourcePath: string|null}|null>}
   */
  async function resolve(id) {
    const match = VERSION_ID_PATTERN.exec(String(id || ''))
    if (!match) return null
    const parsed = parseVersionName(match[2])
    if (!parsed) return null
    const folder = path.join(root(), match[1])
    const filePath = path.join(folder, match[2])
    if (!isInside(filePath, root())) return null
    let stat
    try { stat = await fsp.stat(filePath) } catch { return null }
    if (!stat.isFile()) return null
    const info = await readJsonFile(path.join(folder, SOURCE_INFO_FILE)).catch(() => null)
    return { path: filePath, name: parsed.name, createdAt: parsed.createdAt, size: stat.size, sourcePath: info && typeof info.path === 'string' ? info.path : null }
  }

  /**
   * Applies every retention rule: 10 versions per file, 60 days, and 500 MB in
   * total (oldest first). The newest version of each file goes last.
   * @param {{now?: number}} [options]
   * @returns {Promise<{removed: number, totalBytes: number}>}
   */
  async function prune(pruneOptions = {}) {
    const now = pruneOptions.now ?? Date.now()
    const report = { removed: 0, totalBytes: 0 }
    let folders = []
    try { folders = await fsp.readdir(root(), { withFileTypes: true }) } catch { return report }
    const all = []
    for (const dirent of folders) {
      if (!dirent.isDirectory() || !/^[0-9a-f]{40}$/.test(dirent.name)) continue
      const folder = path.join(root(), dirent.name)
      const before = await versionsIn(folder)
      await pruneFolder(folder, null, now)
      const after = await versionsIn(folder)
      report.removed += before.length - after.length
      after.forEach((version, index) => all.push({ ...version, folder, newest: index === 0 }))
    }
    report.totalBytes = all.reduce((sum, version) => sum + version.size, 0)
    if (report.totalBytes > limits.maxTotalBytes) {
      // Oldest first; the newest version of each file goes last.
      all.sort((a, b) => (a.newest === b.newest ? -newestFirst(a, b) : a.newest ? 1 : -1))
      for (const version of all) {
        if (report.totalBytes <= limits.maxTotalBytes) break
        if (await core.removeFile(version.path)) {
          report.totalBytes -= version.size
          report.removed += 1
        }
      }
      for (const folder of new Set(all.map((version) => version.folder))) {
        if (!(await versionsIn(folder)).length) await removeFolder(folder)
      }
    }
    return report
  }

  function schedulePrune() {
    pruneQueue = pruneQueue.then(() => prune()).catch((error) => {
      logger.warn(`Could not prune the versions store: ${error && error.message}`)
    })
    return pruneQueue
  }

  return Object.freeze({
    root,
    keyFor,
    backup,
    backupOnce,
    hasBackup,
    list,
    resolve,
    prune,
    /** @returns {Promise<void>} resolves when background pruning has finished */
    idle: () => pruneQueue.then(() => {}),
    /** Forgets which files were copied this session (a new session). */
    resetSession: () => backedUp.clear(),
  })
}

// ---------------------------------------------------------------------------
// Defaults
// ---------------------------------------------------------------------------

/**
 * Sets the store folders and the logger used by the default stores. Passing a
 * folder replaces the matching default store.
 * @param {object} [options]
 * @param {string|null} [options.recoveryDir]
 * @param {string|null} [options.versionsDir]
 * @param {{warn: Function}} [options.logger]
 */
function configureStores(options = {}) {
  if (Object.prototype.hasOwnProperty.call(options, 'recoveryDir')) {
    configured.recoveryDir = options.recoveryDir ? path.resolve(options.recoveryDir) : null
    defaultRecovery = null
  }
  if (Object.prototype.hasOwnProperty.call(options, 'versionsDir')) {
    configured.versionsDir = options.versionsDir ? path.resolve(options.versionsDir) : null
    defaultVersions = null
  }
  if (options.logger && typeof options.logger.warn === 'function') logger = { warn: options.logger.warn }
}

/** @returns {RecoveryStore} the process-wide recovery journal */
function recoveryStore() {
  if (!defaultRecovery) defaultRecovery = createRecoveryStore()
  return defaultRecovery
}

/** @returns {VersionsStore} the process-wide versions store */
function versionsStore() {
  if (!defaultVersions) defaultVersions = createVersionsStore()
  return defaultVersions
}

/**
 * @typedef {object} RecoverySnapshot
 * @property {string} docId letters, digits, "_" and "-" (the document id from the registry)
 * @property {number} revision the renderer's revision counter at snapshot time
 * @property {Array<{name: string, data: Buffer|Uint8Array|ArrayBuffer|string}
 *   | {name: string, ref: {path: string, stamp?: object}, required?: boolean}>} parts
 * @property {string} [module] workspace id
 * @property {string} [title] shown on the recovery card
 * @property {string} [kind] "document", "spreadsheet", …
 * @property {string} [format] registry format id of the document
 * @property {string|null} [sourcePath] the file the edits belong to (set by main, never by the renderer)
 * @property {object|null} [sourceStamp] its stamp when the document was opened or last saved (set by main)
 * @property {string|null} [suggestedPath] where Save As should start for an untitled entry
 * @property {unknown} [extra] small JSON (view state)
 */

/**
 * @typedef {object} RecoveryEntry
 * @property {string} id
 * @property {string} docId
 * @property {string|null} title
 * @property {string|null} sourcePath
 * @property {'same'|'changed'|'missing'|'unknown'|null} sourceState
 * @property {boolean} orphaned
 * @property {boolean} restorable false when a required base file changed or is gone
 * @property {string|null} reason 'base-changed' | 'base-missing'
 * @property {string|null} updatedAt ISO time of the snapshot
 * @property {string|null} offeredAt ISO time it was first offered
 * @property {number} revision
 */

/**
 * @typedef {ReturnType<typeof createRecoveryStore>} RecoveryStore
 * @typedef {ReturnType<typeof createVersionsStore>} VersionsStore
 * @typedef {{id: string, path: string, name: string, sourcePath: string, createdAt: string, size: number}} VersionEntry
 */

module.exports = {
  PROCESS_TOKEN,
  RECOVERY_DEFAULTS,
  RECOVERY_SCHEMA,
  VERSION_DEFAULTS,
  cleanStamp,
  configureStores,
  createRecoveryStore,
  createVersionsStore,
  fileKey,
  recoveryStore,
  versionsStore,
}
