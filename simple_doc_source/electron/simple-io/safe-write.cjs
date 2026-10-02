// Vendored from simple/shared/electron/safe-write.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// safeWriteFile(): the one way Simple writes a file the user can see.
//
// The new content goes to "~simple-<8 hex>.tmp" in the target's own folder,
// is flushed to disk, read back and hash-checked, and passes a structural
// validator before the target is touched. The target is then replaced by,
// in order of preference:
//   1. rename(temp → target): atomic when nobody holds the file;
//   2. two-step swap: move the original aside, rename the temp in, delete the
//      aside. Works when every holder allows FILE_SHARE_DELETE (virus scanners,
//      indexers, backup and sync programs), where a single rename fails;
//   3. verified in-place write: back up the original, overwrite, truncate,
//      flush and re-check the hash; restore from the backup on any failure.
//      Used for hard-linked files (to keep every link), for files with their
//      own (explicit or protected) permissions on Windows, because a rename
//      would replace them with the folder's inherited ones, and as the last
//      resort when holders allow writing but not deleting.
// A new file is created with a hard link, so a file another program creates
// under the same name meanwhile is never overwritten. The stamp the caller
// expects is checked again right before the replace and before an in-place
// write, not only at the start: a file another program wrote during a long
// save is reported as CHANGED_ON_DISK, never replaced.
// Steps 1 and 2 are retried for about 3.5 s on EPERM/EBUSY/EACCES. Every temp
// file, moved-aside original and backup is recorded in the io-core journal
// first, so io-core sweep() can finish or undo the save after a crash.
// The original is never left half-written: every outcome is "fully new and
// verified" or "untouched" (or RESTORE_NEEDED with the original's location).

const crypto = require('node:crypto')
const { execFile } = require('node:child_process')
const fsp = require('node:fs/promises')
const path = require('node:path')
const core = require('./io-core.cjs')
const { validateBytes } = require('./validators.cjs')

const { IoError, TRANSIENT_ERRNOS } = core
const CHUNK_BYTES = 4 * 1024 * 1024
const MAX_PRODUCER_VALIDATE_BYTES = 256 * 1024 * 1024

const targetQueues = new Map()

function queueKey(target) {
  return process.platform === 'win32' ? target.toLowerCase() : target
}

/** Runs writes to the same path one after another, in call order. */
function serializePerTarget(target, task) {
  const key = queueKey(target)
  const previous = targetQueues.get(key) || Promise.resolve()
  const run = previous.then(task)
  const settled = run.then(() => {}, () => {})
  targetQueues.set(key, settled)
  settled.then(() => {
    if (targetQueues.get(key) === settled) targetQueues.delete(key)
  })
  return run
}

function isProducer(data) {
  return typeof data === 'function'
}

function toBuffer(data) {
  if (Buffer.isBuffer(data)) return data
  if (typeof data === 'string') return Buffer.from(data, 'utf8')
  if (ArrayBuffer.isView(data)) return Buffer.from(data.buffer, data.byteOffset, data.byteLength)
  if (data instanceof ArrayBuffer) return Buffer.from(data)
  throw new TypeError('safeWriteFile needs a Buffer, Uint8Array, ArrayBuffer, string or producer function.')
}

function isTransient(error) {
  return Boolean(error && TRANSIENT_ERRNOS.has(error.code))
}

function throwIfAborted(signal, target) {
  if (signal && signal.aborted) throw new IoError('CANCELED', { path: target, reason: 'aborted' })
}

async function statOrNull(filePath, fsImpl) {
  try {
    return await fsImpl.stat(filePath)
  } catch (error) {
    if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return null
    throw error
  }
}

/** Follows a symbolic link so the link itself survives the save. */
async function resolveTarget(target, fsImpl) {
  let info
  try {
    info = await fsImpl.lstat(target)
  } catch (error) {
    if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return target
    throw error
  }
  if (!info.isSymbolicLink()) return target
  try {
    return path.resolve(await fsImpl.realpath(target))
  } catch (error) {
    if (!error || error.code !== 'ENOENT') throw error
    // A dangling link: create the file it points to.
    return path.resolve(path.dirname(target), await fsImpl.readlink(target))
  }
}

async function writeFully(handle, buffer, position) {
  let offset = 0
  while (offset < buffer.length) {
    const { bytesWritten } = await handle.write(buffer, offset, buffer.length - offset, position + offset)
    if (!bytesWritten) throw Object.assign(new Error('The disk accepted no more data.'), { code: 'EIO', syscall: 'write' })
    offset += bytesWritten
  }
}

async function writeChunks(handle, bytes, context) {
  for (let offset = 0; offset < bytes.length || offset === 0; offset += CHUNK_BYTES) {
    throwIfAborted(context.signal, context.target)
    const chunk = bytes.subarray(offset, Math.min(bytes.length, offset + CHUNK_BYTES))
    if (chunk.length) await writeFully(handle, chunk, offset)
    context.report('writing', { written: offset + chunk.length, total: bytes.length })
    if (!bytes.length) break
  }
}

async function runProducer(handle, producer, context) {
  const hash = crypto.createHash('sha256')
  let position = 0
  const sink = Object.freeze({
    signal: context.signal,
    /**
     * Appends bytes to the temp file.
     * @param {Buffer|Uint8Array|string} chunk
     */
    async write(chunk) {
      throwIfAborted(context.signal, context.target)
      const buffer = toBuffer(chunk)
      await writeFully(handle, buffer, position)
      hash.update(buffer)
      position += buffer.length
      context.report('writing', { written: position, total: context.expectedSize })
    },
  })
  await producer(sink)
  return { sha256: hash.digest('hex'), size: position }
}

async function copyFileToHandle(sourcePath, handle, fsImpl) {
  const source = await fsImpl.open(sourcePath, 'r')
  try {
    const buffer = Buffer.allocUnsafe(CHUNK_BYTES)
    let position = 0
    for (;;) {
      const { bytesRead } = await source.read(buffer, 0, buffer.length, position)
      if (!bytesRead) break
      await writeFully(handle, buffer.subarray(0, bytesRead), position)
      position += bytesRead
    }
    return position
  } finally {
    await source.close().catch(() => {})
  }
}

async function hashHandle(handle) {
  const hash = crypto.createHash('sha256')
  const buffer = Buffer.allocUnsafe(CHUNK_BYTES)
  let size = 0
  for (;;) {
    const { bytesRead } = await handle.read(buffer, 0, buffer.length, size)
    if (!bytesRead) break
    hash.update(buffer.subarray(0, bytesRead))
    size += bytesRead
  }
  return { sha256: hash.digest('hex'), size }
}

/** Copies the target's current bytes (through an open handle) into a new backup file. */
async function backupFromHandle(handle, backupPath) {
  // Errors on the backup's side (its drive is full, its folder is not
  // writable) are tagged, so they are not reported as a problem with the target.
  const onBackup = async (operation) => {
    try {
      return await operation()
    } catch (error) {
      if (error && typeof error === 'object') error.backupSide = true
      throw error
    }
  }
  await onBackup(() => fsp.mkdir(path.dirname(backupPath), { recursive: true }))
  const backup = await onBackup(() => fsp.open(backupPath, 'wx'))
  const hash = crypto.createHash('sha256')
  let size = 0
  try {
    const buffer = Buffer.allocUnsafe(CHUNK_BYTES)
    for (;;) {
      const { bytesRead } = await handle.read(buffer, 0, buffer.length, size)
      if (!bytesRead) break
      const chunk = buffer.subarray(0, bytesRead)
      await onBackup(() => writeFully(backup, chunk, size))
      hash.update(chunk)
      size += bytesRead
    }
    await onBackup(() => backup.sync())
  } finally {
    await backup.close().catch(() => {})
  }
  return { sha256: hash.digest('hex'), size }
}

/**
 * A Windows system program by absolute path (System32), never by bare name:
 * a bare name is searched in the current folder first, where anyone could
 * have placed a program with that name.
 * @param {string} name e.g. "icacls.exe"
 * @returns {string}
 */
function systemProgram(name) {
  const root = [process.env.SystemRoot, process.env.windir].find((value) => typeof value === 'string' && path.win32.isAbsolute(value)) || 'C:\\Windows'
  return path.win32.join(root, 'System32', name)
}

/**
 * Whether a file carries permissions of its own (an access rule that was not
 * inherited from its folder, or inheritance switched off). A rename-based
 * replace would swap those for the folder's inherited permissions, so such a
 * file is written in place. Windows only; any failure answers false.
 * @param {string} target
 * @param {object} options safeWriteFile options (permissionsProbe test hook, keepPermissions)
 * @returns {Promise<boolean>}
 */
function hasOwnPermissions(target, options) {
  if (options.keepPermissions === false) return Promise.resolve(false)
  if (typeof options.permissionsProbe === 'function') {
    return Promise.resolve().then(() => options.permissionsProbe(target)).then(Boolean, () => false)
  }
  if (process.platform !== 'win32') return Promise.resolve(false)
  return new Promise((resolve) => {
    execFile(systemProgram('icacls.exe'), [target], { windowsHide: true, timeout: 5000, encoding: 'latin1' }, (error, stdout) => {
      if (error || typeof stdout !== 'string') return resolve(false)
      resolve(parseOwnPermissions(stdout))
    })
  })
}

/**
 * Reads icacls output: true when at least one access rule lacks the (I)
 * "inherited" flag. The flags are the same in every Windows language.
 * @param {string} output
 * @returns {boolean}
 */
function parseOwnPermissions(output) {
  let rules = 0
  let own = 0
  for (const line of String(output).split(/\r?\n/)) {
    const match = /:((?:\([A-Za-z,]+\))+)\s*$/.exec(line)
    if (!match) continue
    rules += 1
    if (!/\(I\)/.test(match[1])) own += 1
  }
  return rules > 0 && own > 0
}

/** The stamp check of the preflight, repeated right before the file is replaced or written in place. */
async function recheckStamp(context) {
  const { options, target, fsImpl } = context
  if (!options.expectedStamp || options.force) return
  const comparison = await core.compareStamp(target, options.expectedStamp, { fs: fsImpl })
  if (comparison.state === 'changed') {
    throw new IoError('CHANGED_ON_DISK', {
      path: target,
      reason: 'changed-while-saving',
      technical: 'another program wrote the file while Simple was saving; the content differs from the version that was opened',
    })
  }
  if (comparison.state === 'missing' && !options.recreate) {
    throw new IoError('SOURCE_MISSING', { path: target, technical: 'the file disappeared while Simple was saving' })
  }
}

/** A failure to keep the save journal current: the risky step does not run without its record. */
function journalFailure(target) {
  const cause = core.journal.lastError()
  const technical = `the save journal in ${core.journalDir()} could not be written${cause ? `: ${core.technicalDetail(cause)}` : ''}`
  if (cause && (cause.code === 'ENOSPC' || cause.code === 'EDQUOT')) {
    return new IoError('DISK_FULL', { path: target, drive: core.describeDrive(core.journalDir()), reason: 'journal', technical, cause })
  }
  return new IoError('UNKNOWN', { path: target, reason: 'journal', technical, cause: cause || undefined })
}

/** A failure to write the safety copy of the original before an in-place write. */
function backupFailure(target, backupPath, error) {
  const technical = `the safety copy of the original could not be written to ${path.dirname(backupPath)}: ${core.technicalDetail(error)}`
  if (error && (error.code === 'ENOSPC' || error.code === 'EDQUOT')) {
    return new IoError('DISK_FULL', { path: target, drive: core.describeDrive(backupPath), reason: 'backup', technical, cause: error })
  }
  return new IoError('UNKNOWN', { path: target, reason: 'backup', technical, cause: error })
}

function defaultBackupDir() {
  return path.join(core.journalDir(), 'in-place-backups')
}

function backupNameFor(target) {
  const stamp = new Date().toISOString().replace(/[:.]/g, '-')
  const extension = path.extname(target).slice(0, 16)
  const stem = path.basename(target, path.extname(target)).slice(0, 80)
  return `${stamp}-${crypto.randomBytes(3).toString('hex')}-${stem}${extension}`
}

async function runValidators(context, readBackBytes) {
  const { options, target, tempPath } = context
  const notes = []
  if (options.format) {
    if (readBackBytes) {
      const result = validateBytes(options.format, readBackBytes, { ...(options.validatorOptions || {}), format: options.validatorOptions?.format })
      if (!result.ok) {
        throw new IoError('VALIDATION_FAILED', { path: target, reason: result.reason, technical: `${result.validator} check: the file ${result.reason}` })
      }
      if (result.skipped) notes.push(`no structural check for "${options.format}"`)
      if (result.details && Array.isArray(result.details.notes)) notes.push(...result.details.notes)
    } else {
      notes.push('structural check skipped for a very large file')
    }
  }
  if (typeof options.validate === 'function') {
    let verdict
    try {
      verdict = await options.validate(readBackBytes, { path: tempPath, target, format: options.format || null, size: context.size })
    } catch (error) {
      throw new IoError('VALIDATION_FAILED', { path: target, reason: error && error.message ? error.message : String(error), technical: core.technicalDetail(error), cause: error })
    }
    if (verdict === false || (verdict && typeof verdict === 'object' && verdict.ok === false)) {
      const reason = verdict && verdict.reason ? String(verdict.reason) : 'the workspace check rejected the file'
      throw new IoError('VALIDATION_FAILED', { path: target, reason, technical: `workspace check: ${reason}` })
    }
  }
  return notes
}

/** Rename-based replace with the two-step swap, retried on transient errors. */
async function replaceWithRetry(context) {
  const { fsImpl, target, tempPath, folder, delays, journalId } = context
  let lastError = null
  for (let attempt = 1; attempt <= delays.length + 1; attempt += 1) {
    context.attempts = attempt
    try {
      await fsImpl.rename(tempPath, target)
      return { strategy: 'rename' }
    } catch (error) {
      if (!isTransient(error)) throw error
      lastError = error
    }
    const asidePath = path.join(folder, core.tempFileName('old'))
    if (!(await core.journal.update(journalId, { phase: 'swapping', aside: asidePath }))) {
      // Without that record a crash could leave the original under a temp name
      // nobody knows about: no swap, only plain renames and then the in-place write.
      await core.journal.update(journalId, { phase: 'verified', aside: null })
      if (attempt <= delays.length) await core.sleep(delays[attempt - 1])
      continue
    }
    let movedAside = false
    try {
      await fsImpl.rename(target, asidePath)
      movedAside = true
    } catch (error) {
      await core.journal.update(journalId, { phase: 'verified', aside: null })
      if (error && error.code === 'ENOENT') {
        // The original disappeared meanwhile; the next rename creates the file.
        lastError = error
        continue
      }
      if (!isTransient(error)) throw error
      lastError = error
    }
    if (movedAside) {
      try {
        await fsImpl.rename(tempPath, target)
      } catch (error) {
        try {
          await core.retryTransient(() => fsImpl.rename(asidePath, target), { delays })
        } catch (restoreError) {
          context.keepForSweep = true
          throw new IoError('RESTORE_NEEDED', {
            path: target,
            asidePath,
            backupPath: asidePath,
            technical: `${core.technicalDetail(error)}; putting the original back failed: ${core.technicalDetail(restoreError)}`,
            cause: restoreError,
          })
        }
        await core.journal.update(journalId, { phase: 'verified', aside: null })
        if (!isTransient(error)) throw error
        lastError = error
        if (attempt <= delays.length) await core.sleep(delays[attempt - 1])
        continue
      }
      // The holder keeps its handle on the aside file, which disappears once it lets go.
      const removed = await core.removeFile(asidePath, fsImpl)
      return { strategy: 'swap', leftover: removed ? null : asidePath }
    }
    if (attempt <= delays.length) await core.sleep(delays[attempt - 1])
  }
  context.lastReplaceError = lastError
  return replaceInPlace(context, { afterRetries: true })
}

/**
 * Overwrites the target through its own handle, so hard links and the file's
 * identity survive. The original is copied to a backup first and written back
 * if anything fails.
 */
async function replaceInPlace(context, { afterRetries = false } = {}) {
  const { fsImpl, target, tempPath, bytes, sha256, size, delays, journalId, options } = context
  let handle
  try {
    if (afterRetries) {
      context.attempts += 1
      handle = await fsImpl.open(target, 'r+')
    } else {
      const opened = await core.retryTransient(() => fsImpl.open(target, 'r+'), { delays })
      context.attempts = opened.attempts
      handle = opened.value
    }
  } catch (error) {
    const attempts = context.attempts
    if (error && error.code === 'EBUSY') {
      throw new IoError('LOCKED', { path: target, technical: core.technicalDetail(error, attempts), attempts, errno: error.code, cause: error })
    }
    if (error && (error.code === 'EPERM' || error.code === 'EACCES')) {
      // Renames were refused too, so either a program holds the file without
      // sharing it or this account may not change it. Both need Save As.
      const code = afterRetries && isTransient(context.lastReplaceError) && context.lastReplaceError.code === 'EBUSY' ? 'LOCKED' : 'NO_PERMISSION'
      throw new IoError(code, { path: target, technical: core.technicalDetail(error, attempts), attempts, errno: error.code, cause: error })
    }
    throw error
  }
  const keepBackup = Boolean(options.backupDir)
  const backupPath = path.join(options.backupDir || defaultBackupDir(), backupNameFor(target))
  let original
  try {
    // The retries above can take seconds: never overwrite what another program wrote meanwhile.
    await recheckStamp(context)
    // 'backing-up' tells sweep() the target is untouched and the backup may be partial.
    if (!(await core.journal.update(journalId, { phase: 'backing-up', backup: backupPath, keepBackup }))) {
      await core.journal.update(journalId, { phase: 'verified', backup: null })
      throw journalFailure(target)
    }
    try {
      original = await backupFromHandle(handle, backupPath)
      if (!(await core.journal.update(journalId, { phase: 'in-place', originalSha256: original.sha256, originalSize: original.size }))) {
        throw journalFailure(target)
      }
    } catch (error) {
      await core.journal.update(journalId, { phase: 'verified', backup: null })
      await core.removeFile(backupPath)
      if (core.isIoError(error)) throw error
      // Without a backup the in-place write is not safe; report why the replace was refused.
      if (error && error.backupSide) throw backupFailure(target, backupPath, error)
      const blocker = context.lastReplaceError
      if (blocker && afterRetries) {
        throw new IoError('LOCKED', { path: target, technical: `${core.technicalDetail(blocker, context.attempts)}; backup failed: ${core.technicalDetail(error)}`, attempts: context.attempts, cause: error })
      }
      throw error
    }
    try {
      if (bytes) await writeFully(handle, bytes, 0)
      else await copyFileToHandle(tempPath, handle, fsImpl)
      await handle.truncate(size)
      await handle.sync()
      const check = await hashHandle(handle)
      if (check.sha256 !== sha256 || check.size !== size) {
        throw new IoError('VERIFY_FAILED', { path: target, technical: `in-place write read back ${check.size} of ${size} bytes with a different hash` })
      }
    } catch (error) {
      try {
        await copyFileToHandle(backupPath, handle, fsImpl)
        await handle.truncate(original.size)
        await handle.sync()
        const restored = await hashHandle(handle)
        if (restored.sha256 !== original.sha256) throw new Error('The restored file does not match the backup.')
      } catch (restoreError) {
        context.keepForSweep = true
        throw new IoError('RESTORE_NEEDED', {
          path: target,
          backupPath,
          technical: `${core.technicalDetail(error)}; restoring the original failed: ${core.technicalDetail(restoreError)}`,
          cause: restoreError,
        })
      }
      await core.journal.update(journalId, { phase: 'verified', backup: null })
      if (!keepBackup) await core.removeFile(backupPath)
      throw error
    }
  } finally {
    await handle.close().catch(() => {})
  }
  if (!keepBackup) await core.removeFile(backupPath)
  return { strategy: 'in-place', backupPath: keepBackup ? backupPath : undefined }
}

function createdMeanwhile(target) {
  return new IoError('CHANGED_ON_DISK', {
    path: target,
    reason: 'created-while-saving',
    technical: 'another program created a file with this name while Simple was saving',
  })
}

/**
 * Creates a file that did not exist at the preflight. A hard link never
 * replaces a file, so one another program created meanwhile is reported, not
 * overwritten. Where hard links are not available (FAT drives, some network
 * shares) the name is checked once more right before a rename.
 */
async function createNew(context) {
  const { fsImpl, target, tempPath, delays, options } = context
  const rename = async () => {
    const { attempts } = await core.retryTransient(() => fsImpl.rename(tempPath, target), { delays })
    context.attempts = attempts
    return { strategy: 'new' }
  }
  if (options.force) return rename()
  if (typeof fsImpl.link === 'function') {
    try {
      await fsImpl.link(tempPath, target)
      context.attempts = 1
    } catch (error) {
      if (error && error.code === 'EEXIST') throw createdMeanwhile(target)
      if (await core.pathExists(target, fsImpl)) throw createdMeanwhile(target)
      return rename()
    }
    // The target is complete; the temp name is only a second link to it.
    const removed = await core.removeFile(tempPath, fsImpl)
    return removed ? { strategy: 'new' } : { strategy: 'new', leftover: tempPath }
  }
  if (await core.pathExists(target, fsImpl)) throw createdMeanwhile(target)
  return rename()
}

async function syncFolder(folder) {
  if (process.platform === 'win32') return
  let handle
  try {
    handle = await fsp.open(folder, 'r')
    await handle.sync()
  } catch {
    // Not every file system supports flushing a folder.
  } finally {
    if (handle) await handle.close().catch(() => {})
  }
}

async function writeOnce(requestedPath, data, options, started) {
  const fsImpl = options.fs || fsp
  const delays = options.retryDelays || core.RETRY_DELAYS
  const signal = options.signal
  const onProgress = typeof options.onProgress === 'function' ? options.onProgress : null
  const producer = isProducer(data) ? data : null
  const bytes = producer ? null : toBuffer(data)
  const context = {
    fsImpl,
    options,
    delays,
    signal,
    bytes,
    target: requestedPath,
    folder: path.dirname(requestedPath),
    tempPath: null,
    journalId: null,
    attempts: 0,
    sha256: null,
    size: bytes ? bytes.length : null,
    expectedSize: bytes ? bytes.length : options.expectedSize,
    keepForSweep: false,
    lastReplaceError: null,
    report(phase, extra = {}) {
      if (!onProgress) return
      try { onProgress({ phase, path: context.target, ...extra }) } catch {}
    },
  }
  let stage = 'preflight'
  let tempCreated = false
  try {
    throwIfAborted(signal, requestedPath)
    context.report('preparing')

    // 1. Check the name, then resolve the real target (a symbolic link survives).
    const checkName = (candidate) => {
      const check = core.checkTargetPath(candidate)
      if (!check.ok) throw new IoError(check.code, { path: candidate, reason: check.reason, technical: `name check: ${check.reason}` })
    }
    checkName(requestedPath)
    const target = await resolveTarget(requestedPath, fsImpl)
    if (target !== requestedPath) checkName(target)
    context.target = target
    const folder = path.dirname(target)
    context.folder = folder

    // 2. Preflight: nothing is written yet.
    const folderStat = await statOrNull(folder, fsImpl)
    if (!folderStat || !folderStat.isDirectory()) {
      throw new IoError('FOLDER_MISSING', {
        path: target, folder, nearestFolder: await core.nearestExistingFolder(folder, { fs: fsImpl }), technical: 'the folder does not exist',
      })
    }
    let existing = await statOrNull(target, fsImpl)
    if (existing && existing.isDirectory()) throw new IoError('INVALID_NAME', { path: target, reason: 'folder-exists', technical: 'a folder has this name' })
    if (existing && (existing.mode & 0o200) === 0) throw new IoError('READ_ONLY', { path: target, technical: 'the file has the read-only attribute' })
    if (existing && options.expectedStamp && !options.force) {
      const comparison = await core.compareStamp(target, options.expectedStamp, { fs: fsImpl })
      if (comparison.state === 'changed') {
        throw new IoError('CHANGED_ON_DISK', {
          path: target,
          technical: `size ${options.expectedStamp.size} → ${comparison.stamp.size} bytes; the content differs from the version that was opened`,
        })
      }
      if (comparison.state === 'missing') existing = null
    }
    if (!existing && options.expectedStamp && !options.recreate) {
      throw new IoError('SOURCE_MISSING', { path: target, technical: 'the file is no longer at this path' })
    }
    if (typeof context.expectedSize === 'number') {
      const space = await core.freeSpace(folder, { fs: fsImpl })
      const required = core.requiredFreeBytes(context.expectedSize)
      if (space && space.total > 0 && space.free < required) {
        throw new IoError('DISK_FULL', {
          path: target,
          needed: required - space.free,
          drive: core.describeDrive(target),
          technical: `${core.formatBytes(space.free)} free, ${core.formatBytes(required)} needed`,
        })
      }
    }
    // Hard links and a file's own permissions survive only an in-place write.
    let identity = Boolean(existing) && (options.identity === 'always' || existing.nlink > 1)
    let ownPermissions = false
    if (existing && !identity) {
      ownPermissions = await hasOwnPermissions(target, options)
      identity = ownPermissions
    }

    // 3. Temp file in the same folder, recorded in the journal before it exists.
    stage = 'temp'
    let handle = null
    for (let tries = 0; !handle; tries += 1) {
      context.tempPath = path.join(folder, core.tempFileName('tmp'))
      if (context.journalId) await core.journal.update(context.journalId, { temp: context.tempPath })
      else context.journalId = await core.journal.add({ target, temp: context.tempPath, phase: 'writing', size: context.expectedSize ?? null })
      try {
        handle = await fsImpl.open(context.tempPath, 'wx')
      } catch (error) {
        if (error && error.code === 'EEXIST' && tries < 5) continue
        throw error
      }
    }
    tempCreated = true
    try {
      if (bytes) {
        context.sha256 = core.hashBytes(bytes)
        await writeChunks(handle, bytes, context)
      } else {
        const produced = await runProducer(handle, producer, context)
        context.sha256 = produced.sha256
        context.size = produced.size
      }
      await handle.sync()
    } finally {
      await handle.close().catch(() => {})
    }
    throwIfAborted(signal, target)

    // 4. Verify: read back from disk, then the structural and injected checks.
    stage = 'verify'
    context.report('verifying')
    // A virus scanner often opens a fresh file for a moment, so reads retry briefly.
    const { value: readBack } = await core.retryTransient(() => core.hashFile(context.tempPath, { fs: fsImpl, signal }), { delays })
    if (readBack.sha256 !== context.sha256 || readBack.size !== context.size) {
      throw new IoError('VERIFY_FAILED', { path: target, technical: `the temp file read back ${readBack.size} of ${context.size} bytes with a different hash` })
    }
    let validationBytes = bytes
    if (!validationBytes && (options.format || options.validate) && context.size <= MAX_PRODUCER_VALIDATE_BYTES) {
      validationBytes = (await core.retryTransient(() => fsImpl.readFile(context.tempPath), { delays })).value
    }
    const notes = await runValidators({ ...context, tempPath: context.tempPath }, validationBytes)
    await core.journal.update(context.journalId, { phase: 'verified', sha256: context.sha256, size: context.size })
    throwIfAborted(signal, target)
    const qaDelay = core.qaWriteDelayMs()
    if (qaDelay) await core.sleep(qaDelay)

    // 5. Replace. From here on the save is never abandoned halfway.
    stage = 'replace'
    context.report('replacing')
    if (existing) await recheckStamp(context)
    let outcome
    if (!existing) outcome = await createNew(context)
    else if (identity) outcome = await replaceInPlace(context)
    else outcome = await replaceWithRetry(context)
    if (outcome.strategy !== 'in-place') tempCreated = false

    // 6. Finish: stamp, journal, leftovers. Nothing here may turn the
    // completed replace into a reported failure.
    stage = 'done'
    let stat
    try {
      stat = await fsImpl.stat(target)
    } catch {
      stat = { size: context.size, mtimeMs: Date.now() }
    }
    if (ownPermissions) notes.push("written in place to keep the file's own permissions")
    if (outcome.leftover) {
      await core.journal.update(context.journalId, { phase: 'cleanup', temp: null, aside: outcome.leftover })
      notes.push(outcome.strategy === 'new' ? 'a leftover temporary name will be removed later' : 'the moved-aside original will be removed later')
    } else if (outcome.strategy === 'in-place') {
      if (await core.removeFile(context.tempPath, fsImpl)) await core.journal.remove(context.journalId)
      else await core.journal.update(context.journalId, { phase: 'cleanup', backup: null })
      tempCreated = false
    } else {
      await core.journal.remove(context.journalId)
    }
    // Kept entries of earlier failed saves of this file must never bring their original back over this one.
    await core.journal.supersede(target, context.journalId)
    await syncFolder(folder)
    context.report('done')
    const result = {
      path: target,
      stamp: { size: stat.size, mtimeMs: stat.mtimeMs, sha256: context.sha256 },
      strategy: outcome.strategy,
      attempts: Math.max(1, context.attempts),
      ms: Date.now() - started,
      notes,
    }
    if (outcome.backupPath) result.backupPath = outcome.backupPath
    return result
  } catch (error) {
    const ioError = core.toIoError(error, { path: context.target, stage, attempts: context.attempts || undefined })
    if (ioError.code === 'DISK_FULL' && !ioError.drive) ioError.drive = core.describeDrive(context.target)
    if (context.keepForSweep) {
      // The original sits in a backup or under a temp name; the entry stays for sweep(),
      // which puts it back unless a later save of the same file succeeds first.
      await core.journal.update(context.journalId, { restoreNeeded: true })
    } else {
      if (tempCreated && context.tempPath) {
        if (await core.removeFile(context.tempPath, fsImpl)) await core.journal.remove(context.journalId)
        else await core.journal.update(context.journalId, { phase: 'cleanup' })
      } else if (context.journalId) {
        await core.journal.remove(context.journalId)
      }
    }
    throw ioError
  }
}

/**
 * Writes a file the user can see, so that it ends up either fully replaced by
 * verified new content or exactly as it was. Never clears a read-only
 * attribute and never overwrites a file another program changed since
 * `expectedStamp` was taken (unless `force`).
 *
 * @param {string} targetPath file to create or replace
 * @param {Buffer|Uint8Array|ArrayBuffer|string|((sink: {write(chunk: Buffer|Uint8Array|string): Promise<void>, signal?: AbortSignal}) => Promise<void>)} data
 *   the bytes, or a producer that streams them through `sink.write()` (for very large files)
 * @param {object} [options]
 * @param {{size: number, mtimeMs: number, sha256?: string}} [options.expectedStamp] stamp captured in main when the file was opened or last saved; never from the renderer
 * @param {boolean} [options.force] the user chose Replace in the "changed by another program" prompt; skips the
 *   stamp checks and lets a new file replace one another program created meanwhile. Callers that know the
 *   stamp the user saw pass it as `expectedStamp` instead, so a later change is still caught
 * @param {boolean} [options.recreate] the user chose Save Here Again after the file was moved or deleted
 * @param {string} [options.format] registry format id, extension or validator id ('docx', '.xlsx', 'png').
 *   Pass it for bytes Simple produced; leave it out for an unchanged copy of the user's own bytes
 * @param {object} [options.validatorOptions] extra options for the structural validator (encoding, text, strictOdf, …)
 * @param {(bytes: Buffer|null, info: {path: string, target: string, format: string|null, size: number}) => (boolean|void|{ok: boolean, reason?: string}|Promise<boolean|void|{ok: boolean, reason?: string}>)} [options.validate]
 *   deeper check injected by the workspace (pdf-lib load, SheetJS reopen, validateDocxBytes); runs on the
 *   bytes read back from disk. `bytes` is null for producer output above 256 MB: read `info.path` instead
 * @param {'auto'|'always'} [options.identity='auto'] 'auto' writes in place for hard-linked files and, on Windows,
 *   for files with their own permissions (an explicit or protected access list); 'always' forces it
 * @param {boolean} [options.keepPermissions=true] false skips the permissions check (the file may then take its
 *   folder's inherited permissions)
 * @param {(target: string) => boolean|Promise<boolean>} [options.permissionsProbe] test hook replacing the icacls check
 * @param {string} [options.backupDir] keep the in-place backup there (versions store) instead of deleting it after success
 * @param {number} [options.expectedSize] size a producer will write, for the free-space check
 * @param {AbortSignal} [options.signal] cancels before the replace starts (CANCELED); never interrupts a replace
 * @param {(progress: {phase: 'preparing'|'writing'|'verifying'|'replacing'|'done', path: string, written?: number, total?: number}) => void} [options.onProgress]
 * @param {readonly number[]} [options.retryDelays] retry schedule (default io-core RETRY_DELAYS, about 3.5 s)
 * @param {object} [options.fs] fs/promises-compatible implementation, for tests that inject failures
 * @returns {Promise<{path: string, stamp: {size: number, mtimeMs: number, sha256: string},
 *   strategy: 'rename'|'swap'|'in-place'|'new', attempts: number, ms: number, notes: string[], backupPath?: string}>}
 * @throws {IoError} LOCKED, READ_ONLY, NO_PERMISSION, READ_ONLY_VOLUME, FOLDER_MISSING, FILE_UNAVAILABLE, DISK_FULL,
 *   NAME_TOO_LONG, INVALID_NAME, CHANGED_ON_DISK, SOURCE_MISSING, VERIFY_FAILED, VALIDATION_FAILED, RESTORE_NEEDED,
 *   CANCELED or UNKNOWN. On every code except RESTORE_NEEDED the target is unchanged and no temp file is left.
 */
function safeWriteFile(targetPath, data, options = {}) {
  const started = Date.now()
  let requested
  try {
    if (typeof targetPath !== 'string' || !targetPath.trim()) throw new TypeError('safeWriteFile needs a file path.')
    if (!isProducer(data)) toBuffer(data)
    requested = path.resolve(targetPath)
  } catch (error) {
    return Promise.reject(error)
  }
  const promise = serializePerTarget(requested, () => writeOnce(requested, data, options, started))
  return core.trackPendingWrite(promise, requested)
}

/**
 * Like safeWriteFile, but never throws for expected failures: resolves to the
 * IPC-ready IoResult shape instead.
 * @param {string} targetPath
 * @param {Parameters<typeof safeWriteFile>[1]} data
 * @param {Parameters<typeof safeWriteFile>[2]} [options]
 * @returns {Promise<({ok: true, name: string} & Awaited<ReturnType<typeof safeWriteFile>>) | ReturnType<typeof core.toIoResult>>}
 */
async function safeWriteResult(targetPath, data, options) {
  try {
    const result = await safeWriteFile(targetPath, data, options)
    return { ok: true, name: path.basename(result.path), ...result }
  } catch (error) {
    if (error instanceof TypeError) throw error
    return core.toIoResult(error, { path: targetPath ? path.resolve(String(targetPath)) : undefined })
  }
}

module.exports = {
  parseOwnPermissions,
  safeWriteFile,
  safeWriteResult,
}
