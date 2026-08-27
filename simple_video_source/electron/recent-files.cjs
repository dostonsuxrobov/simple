'use strict'

const path = require('node:path')

const RECENT_LIMIT = 12

function validRecent(item) {
  return item
    && typeof item.path === 'string'
    && item.path.length > 0
    && typeof item.name === 'string'
    && Number.isFinite(item.openedAt)
}

function sanitizeRecents(value, platform = process.platform) {
  if (!Array.isArray(value)) return []
  const seen = new Set()
  const result = []
  for (const item of value) {
    if (!validRecent(item)) continue
    const resolved = path.resolve(item.path)
    const identity = platform === 'win32' ? resolved.toLowerCase() : resolved
    if (seen.has(identity)) continue
    seen.add(identity)
    result.push({ path: resolved, name: path.basename(resolved), openedAt: item.openedAt })
    if (result.length === RECENT_LIMIT) break
  }
  return result
}

function mergeRecent(recents, filePath, openedAt = Date.now(), platform = process.platform) {
  const resolved = path.resolve(filePath)
  const identity = platform === 'win32' ? resolved.toLowerCase() : resolved
  const remaining = sanitizeRecents(recents, platform).filter((item) => {
    const itemIdentity = platform === 'win32' ? item.path.toLowerCase() : item.path
    return itemIdentity !== identity
  })
  return [{ path: resolved, name: path.basename(resolved), openedAt }, ...remaining].slice(0, RECENT_LIMIT)
}

module.exports = { RECENT_LIMIT, mergeRecent, sanitizeRecents }
