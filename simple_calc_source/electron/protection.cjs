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

function verifyPassword(protection, password) {
  if (!protection || typeof protection !== 'object' || !protection.hashValue) return true
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

module.exports = { registerProtectionHandlers, hashPassword, verifyPassword }
