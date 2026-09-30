'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')

const EDITABLE_SAVE_FORMATS = new Set(['xlsx', 'xls', 'ods', 'csv', 'tsv'])

function saveFormat(input, record) {
  const format = String(input.format || record.sourceFormat || 'xlsx').toLowerCase().replace(/^\./, '')
  if (EDITABLE_SAVE_FORMATS.has(format) || (input.sourceUnmodified === true && format === record.sourceFormat)) return format
  throw new Error(`Edited ${format.toUpperCase()} files cannot be saved in their original format yet. Use Export As to create an XLSX copy; the original file has not been changed.`)
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

module.exports = { EDITABLE_SAVE_FORMATS, saveFormat, unchangedSourceBytes, assertSourceUnchanged }
