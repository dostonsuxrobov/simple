'use strict'

// Sheet-protection passwords use Excel's salted, spun hash (sheetProtection algorithmName /
// hashValue / saltValue / spinCount). Hashing runs in the main process with Node's crypto.

const crypto = require('node:crypto')
const Encryptor = require('exceljs/lib/utils/encryptor')

const ALGORITHMS = { 'SHA-512': 'sha512', 'SHA-384': 'sha384', 'SHA-256': 'sha256', 'SHA-1': 'sha1', MD5: 'md5' }
const DEFAULT_SPIN_COUNT = 100000
const MAX_SPIN_COUNT = 10_000_000

function hashPassword(password) {
  const saltValue = crypto.randomBytes(16).toString('base64')
  const hashValue = Encryptor.convertPasswordToHash(String(password), 'sha512', saltValue, DEFAULT_SPIN_COUNT)
  return { algorithmName: 'SHA-512', hashValue, saltValue, spinCount: DEFAULT_SPIN_COUNT }
}

/**
 * Excel's legacy 16-bit sheet password hash (the `password="CC3D"` attribute written by Excel
 * 2007/2010, XlsxWriter, openpyxl and LibreOffice). Weak by design, but a file protected this
 * way must still ask for its password and keep it on save.
 */
function legacyPasswordHash(password) {
  const text = String(password)
  let hash = 0
  for (let index = 0; index < text.length; index += 1) {
    const value = text.charCodeAt(index) << (index + 1)
    hash ^= (value & 0x7fff) | (value >> 15)
  }
  hash ^= text.length
  hash ^= 0xce4b
  return (hash & 0xffff).toString(16).toUpperCase().padStart(4, '0')
}

function verifyPassword(protection, password) {
  if (!protection || typeof protection !== 'object') return true
  if (!protection.hashValue) {
    const legacy = typeof protection.password === 'string' ? protection.password.trim() : ''
    if (!legacy) return true
    if (!/^[0-9a-f]{1,4}$/i.test(legacy)) return false
    return legacyPasswordHash(password) === legacy.toUpperCase().padStart(4, '0')
  }
  const algorithm = ALGORITHMS[String(protection.algorithmName || '').toUpperCase()] || ALGORITHMS[String(protection.algorithmName || '')]
  const spinCount = Math.trunc(Number(protection.spinCount) || 0)
  if (!algorithm || !protection.saltValue || spinCount < 0 || spinCount > MAX_SPIN_COUNT) return false
  let computed
  try {
    computed = Buffer.from(Encryptor.convertPasswordToHash(String(password), algorithm, String(protection.saltValue), spinCount), 'base64')
  } catch {
    return false
  }
  const expected = Buffer.from(String(protection.hashValue), 'base64')
  return computed.length === expected.length && crypto.timingSafeEqual(computed, expected)
}

function registerProtectionHandlers(ipcMain, options = {}) {
  const assertTrustedSender = typeof options.assertTrustedSender === 'function' ? options.assertTrustedSender : () => {}
  ipcMain.handle('protection:hash', (event, password) => {
    assertTrustedSender(event)
    if (typeof password !== 'string' || !password || password.length > 255) throw new Error('Passwords can be 1 to 255 characters.')
    return hashPassword(password)
  })
  ipcMain.handle('protection:verify', (event, input) => {
    assertTrustedSender(event)
    return verifyPassword(input && input.protection, input && typeof input.password === 'string' ? input.password : '')
  })
}

let protectionPatched = false

/**
 * ExcelJS 4.4 reads only the modern hash attributes of <sheetProtection>; the legacy
 * `password` attribute was dropped on open, so any password unprotected the sheet and the
 * saved file lost it. Keep it on read and write it back while no modern hash replaces it.
 */
function installProtectionPatches() {
  if (protectionPatched) return
  protectionPatched = true
  const SheetProtectionXform = require('exceljs/lib/xlsx/xform/sheet/sheet-protection-xform')
  const proto = SheetProtectionXform.prototype
  const originalParseOpen = proto.parseOpen
  proto.parseOpen = function parseOpen(node) {
    const handled = originalParseOpen.call(this, node)
    if (handled && node && node.name === this.tag && this.model && !this.model.hashValue) {
      const legacy = node.attributes && node.attributes.password
      if (typeof legacy === 'string' && /^[0-9a-f]{1,4}$/i.test(legacy.trim())) this.model.password = legacy.trim().toUpperCase()
    }
    return handled
  }
  const originalRender = proto.render
  proto.render = function render(xmlStream, model) {
    const legacy = model && model.sheet && !model.hashValue && typeof model.password === 'string' && /^[0-9a-f]{1,4}$/i.test(model.password)
      ? model.password.toUpperCase()
      : null
    if (!legacy) return originalRender.call(this, xmlStream, model)
    const stream = Object.create(xmlStream)
    stream.leafNode = (tag, attributes, text) => xmlStream.leafNode(tag, { password: legacy, ...(attributes || {}) }, text)
    return originalRender.call(this, stream, model)
  }
}

module.exports = { registerProtectionHandlers, hashPassword, verifyPassword, legacyPasswordHash, installProtectionPatches }
