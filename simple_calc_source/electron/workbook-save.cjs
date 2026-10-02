'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const crypto = require('node:crypto')

const EDITABLE_SAVE_FORMATS = new Set(['xlsx', 'xls', 'ods', 'csv', 'tsv'])
// Formats that can only be written faithfully by the local document engine (LibreOffice).
const ENGINE_FORMATS = new Set(['xls', 'ods'])

function saveFormat(input, record) {
  const format = String(input.format || record.sourceFormat || 'xlsx').toLowerCase().replace(/^\./, '')
  if (EDITABLE_SAVE_FORMATS.has(format) || (input.sourceUnmodified === true && format === record.sourceFormat)) return format
  throw new Error(`Edited ${format.toUpperCase()} files cannot be saved in their original format yet. Use Export As to create an XLSX copy; the original file has not been changed.`)
}

/** Whether the document's original bytes can be copied (a file on disk or dropped bytes). */
function hasOriginalBytes(record) {
  const snapshot = record && record.sourceSnapshot
  if (snapshot && Buffer.isBuffer(snapshot.bytes)) return true
  return Boolean(record && record.path && snapshot && snapshot.path === record.path)
}

async function unchangedSourceBytes(record, required = false) {
  if (record.sourceSnapshot && Buffer.isBuffer(record.sourceSnapshot.bytes)) return Buffer.from(record.sourceSnapshot.bytes)
  const snapshot = record.sourceSnapshot
  if (!record.path || !snapshot || snapshot.path !== record.path) {
    if (required) throw new Error('The original file is unavailable. Reopen it before saving an unchanged copy.')
    return null
  }
  const handle = await fs.open(record.path, 'r')
  try {
    const info = await handle.stat()
    if (!info.isFile() || info.size !== snapshot.size || info.mtimeMs !== snapshot.modified) {
      throw new Error('This file changed outside Simple. Reopen it to load those changes, or use Save As to keep your edits in a separate copy.')
    }
    return await handle.readFile()
  } finally {
    await handle.close()
  }
}

async function assertSourceUnchanged(record) {
  const snapshot = record.sourceSnapshot
  if (!record.path || !snapshot || snapshot.path !== record.path) return
  const info = await fs.stat(record.path)
  if (!info.isFile() || info.size !== snapshot.size || info.mtimeMs !== snapshot.modified) {
    throw new Error('This file changed outside Simple. Reopen it to load those changes, or use Save As to keep your edits in a separate copy.')
  }
}

function extensionOf(filePath) {
  return path.extname(String(filePath || '')).toLowerCase()
}

function samePath(left, right) {
  return Boolean(left && right) && path.resolve(left).toLowerCase() === path.resolve(right).toLowerCase()
}

async function pathExists(filePath) {
  try {
    await fs.stat(filePath)
    return true
  } catch (error) {
    if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) return false
    throw error
  }
}

/**
 * A free file name next to `originalPath` with another extension: "Budget.xlsx", then
 * "Budget (edited).xlsx", "Budget (edited 2).xlsx"... Never the original file itself.
 */
async function siblingPath(originalPath, extension, exists = pathExists) {
  const directory = path.dirname(path.resolve(originalPath))
  const stem = path.basename(originalPath, path.extname(originalPath)) || 'Workbook'
  const names = [`${stem}.${extension}`, `${stem} (edited).${extension}`]
  for (let index = 2; index < 1000; index += 1) names.push(`${stem} (edited ${index}).${extension}`)
  for (const name of names) {
    const candidate = path.join(directory, name)
    if (samePath(candidate, originalPath)) continue
    if (!(await exists(candidate))) return candidate
  }
  throw new Error('No free file name was found next to the original. Use Save As to choose a name.')
}

/**
 * Save As filters. Without the document engine, XLS and ODS cannot be written faithfully, so
 * they are not offered (an edited .xls/.ods is saved as .xlsx next to the original instead).
 */
function saveFilters(preferred = 'xlsx', options = {}) {
  const officeEngine = options.officeEngine !== false
  const filters = [
    { name: 'Excel workbook', extensions: ['xlsx'] },
    ...(officeEngine ? [{ name: 'Excel 97–2003 workbook', extensions: ['xls'] }, { name: 'OpenDocument spreadsheet', extensions: ['ods'] }] : []),
    { name: 'Comma-separated values', extensions: ['csv'] },
    { name: 'Tab-separated values', extensions: ['tsv'] },
  ]
  if (!EDITABLE_SAVE_FORMATS.has(preferred)) filters.push({ name: `${preferred.toUpperCase()} original file`, extensions: [preferred] })
  return filters.sort((left, right) => Number(right.extensions.includes(preferred)) - Number(left.extensions.includes(preferred)))
}

function formatForPath(filePath, fallback = 'xlsx') {
  const ext = extensionOf(filePath).slice(1)
  return EDITABLE_SAVE_FORMATS.has(ext) || ext === fallback ? ext : fallback
}

function formatLabel(format) {
  return format === 'xls' ? 'Excel 97-2003' : format === 'ods' ? 'OpenDocument' : format.toUpperCase()
}

/**
 * Save a document. All decisions about where and how to write live here so they can be tested
 * without Electron; `deps` supplies the side effects:
 *   officeEngine      whether the local document engine (LibreOffice) is available
 *   chooseSavePath    async ({ defaultPath, format, saveAs }) => chosen path or null (cancel)
 *   serialize         async (workbook, format, options) => bytes (serializeWorkbook)
 *   writeFile         async (targetPath, bytes) => void (atomic write)
 *   dialectFor        (workbook, format, fallback) => delimited dialect or undefined
 *   backupDirectory   folder for the one-time original backup of an .xls/.ods overwrite
 *   exists            async (path) => boolean
 *
 * The merge base (record.mergeBase) is the package the model was imported from. It is used for
 * every XLSX save of the document, never the file written by the previous save: the model's
 * worksheet, chart and picture identities point into the imported package.
 */
async function saveDocument(record, input, deps) {
  const officeEngine = Boolean(deps.officeEngine)
  let requestedFormat = saveFormat(input, record)
  const originalCopyPossible = (format) => format === record.sourceFormat && input.sourceUnmodified === true && hasOriginalBytes(record)
  const needsEngine = (format) => ENGINE_FORMATS.has(format) && !officeEngine && !originalCopyPossible(format)
  const suggestedBase = input.suggestedName || record.originalName || 'Untitled'
  let targetPath = null
  let redirectedFrom = null

  const canOverwrite = !input.saveAs && record.path && extensionOf(record.path) === `.${requestedFormat}`
  if (canOverwrite && needsEngine(requestedFormat)) {
    // Without the document engine an edited .xls/.ods cannot be written back faithfully.
    // Write the edits to an .xlsx next to the original instead of failing; the original
    // file is not touched.
    redirectedFrom = record.path
    targetPath = await siblingPath(record.path, 'xlsx', deps.exists || pathExists)
    requestedFormat = 'xlsx'
  } else if (canOverwrite) {
    targetPath = record.path
  } else {
    const dialogFormat = needsEngine(requestedFormat) ? 'xlsx' : requestedFormat
    const chosen = await deps.chooseSavePath({ format: dialogFormat, suggestedName: suggestedBase, directory: record.path ? path.dirname(record.path) : null, saveAs: Boolean(input.saveAs) })
    if (!chosen) return null
    requestedFormat = formatForPath(chosen, dialogFormat)
    targetPath = chosen.toLowerCase().endsWith(`.${requestedFormat}`) ? chosen : `${chosen}.${requestedFormat}`
    if (needsEngine(requestedFormat)) {
      // The name was typed with .xls/.ods explicitly: keep the user's choice of name and
      // folder, but write an .xlsx file beside it (never over an existing file).
      redirectedFrom = targetPath
      targetPath = await siblingPath(targetPath, 'xlsx', deps.exists || pathExists)
      requestedFormat = 'xlsx'
    }
  }

  const originalCopy = originalCopyPossible(requestedFormat)
  const notes = []
  let bytes
  if (originalCopy) {
    bytes = await unchangedSourceBytes(record, true)
  } else {
    const options = {
      baseBytes: requestedFormat === 'xlsx' && Buffer.isBuffer(record.mergeBase) ? record.mergeBase : null,
      sourceFormat: record.sourceFormat,
      officeEngine,
      warnings: notes,
    }
    if (['csv', 'tsv'].includes(requestedFormat) && typeof deps.dialectFor === 'function') {
      const dialect = deps.dialectFor(input.workbook, requestedFormat, record.dialect)
      if (dialect) options.dialect = dialect
    }
    bytes = await deps.serialize(input.workbook, requestedFormat, options)
  }

  const overwritesSource = Boolean(record.path) && samePath(targetPath, record.path)
  if (overwritesSource) await assertSourceUnchanged(record)
  // Legacy edits pass through the Office engine. Keep the original binary once per open
  // document so unmodelled legacy features remain recoverable.
  if (overwritesSource && ENGINE_FORMATS.has(record.sourceFormat) && !originalCopy && !record.backupPath && deps.backupDirectory) {
    const originalBytes = await unchangedSourceBytes(record, true)
    await fs.mkdir(deps.backupDirectory, { recursive: true })
    const backupPath = path.join(deps.backupDirectory, `${path.basename(record.path, path.extname(record.path))}-${crypto.randomUUID()}.${record.sourceFormat}`)
    await fs.writeFile(backupPath, originalBytes, { flag: 'wx' })
    record.backupPath = backupPath
    await assertSourceUnchanged(record)
  }

  const resolvedTarget = path.resolve(targetPath)
  await deps.writeFile(resolvedTarget, bytes)
  record.path = resolvedTarget
  record.sourceFormat = requestedFormat
  record.originalName = path.basename(resolvedTarget)
  const info = await fs.stat(record.path)
  record.sourceSnapshot = { path: record.path, size: info.size, modified: info.mtimeMs }
  const result = { path: record.path, name: record.originalName, format: requestedFormat, backupPath: record.backupPath }
  if (redirectedFrom) {
    const originalName = path.basename(redirectedFrom)
    const originalExists = await (deps.exists || pathExists)(redirectedFrom)
    result.redirected = true
    result.originalPath = redirectedFrom
    result.message = originalExists
      ? `Saved your edits as "${record.originalName}" next to "${originalName}". ${formatLabel(extensionOf(redirectedFrom).slice(1))} files need the document engine to be written, so the original was not changed.`
      : `Saved as "${record.originalName}". ${formatLabel(extensionOf(redirectedFrom).slice(1))} files need the document engine to be written, so the workbook was saved as XLSX instead.`
  }
  if (notes.length) result.notes = [...new Set(notes)]
  return result
}

module.exports = {
  EDITABLE_SAVE_FORMATS,
  ENGINE_FORMATS,
  saveFormat,
  hasOriginalBytes,
  unchangedSourceBytes,
  assertSourceUnchanged,
  siblingPath,
  formatForPath,
  saveFilters,
  saveDocument,
}
