'use strict'

// Saving edits to a document whose original Simple cannot write back (a legacy
// .doc without a local Office engine, or RTF/HTML saved with a .doc name):
// the edits go to name.docx beside the original, then "name (edited).docx",
// "name (edited 2).docx" …, never over an existing file. The original is never
// opened for writing.

const fs = require('node:fs/promises')
const { sameFilePath, siblingCandidates } = require('./document-paths.cjs')

async function exists(filePath) {
  try { await fs.access(filePath); return true } catch { return false }
}

/** The first free sibling name, for the confirmation the renderer shows. */
async function proposedSiblingPath(sourcePath, extension = '.docx') {
  for (const candidate of siblingCandidates(sourcePath, extension)) {
    if (sameFilePath(candidate, sourcePath)) continue
    if (!await exists(candidate)) return candidate
  }
  throw new Error('Simple Docs could not find a free file name beside the original document.')
}

/**
 * Reserve the first free sibling name with an exclusive create, then replace
 * the empty placeholder through `write(target, data)` (the app's atomic write).
 * Returns the path written.
 */
async function writeBesideSource(sourcePath, data, write, extension = '.docx') {
  if (typeof write !== 'function') throw new Error('A file writer is required.')
  for (const candidate of siblingCandidates(sourcePath, extension)) {
    if (sameFilePath(candidate, sourcePath)) continue
    let handle
    try {
      handle = await fs.open(candidate, 'wx')
    } catch (error) {
      if (error.code === 'EEXIST') continue
      throw error
    }
    await handle.close()
    try {
      await write(candidate, data)
      return candidate
    } catch (error) {
      await fs.rm(candidate, { force: true }).catch(() => {})
      throw error
    }
  }
  throw new Error('Simple Docs could not find a free file name beside the original document.')
}

module.exports = { proposedSiblingPath, writeBesideSource }
