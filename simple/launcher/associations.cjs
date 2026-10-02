'use strict'

const { execFile } = require('node:child_process')
const path = require('node:path')
const { promisify } = require('node:util')
const { associationExtensions } = require('../shared/electron/formats.cjs')

const execFileAsync = promisify(execFile)
const ROOT = 'HKCU\\Software\\Classes'
const PROG_ID = 'Simple.Unified.File'
const REGISTERED_NAME = 'simple'
const APPLICATION_KEY = `${ROOT}\\Applications\\simple.exe`
const CAPABILITIES_KEY = `${APPLICATION_KEY}\\Capabilities`

/**
 * A Windows system program by absolute path (System32), never by bare name:
 * Windows searches the current folder before PATH for a bare name, and the
 * current folder may be one where anyone could have placed a program.
 * @param {string} name e.g. "reg.exe"
 * @returns {string}
 */
function systemProgram(name) {
  const root = [process.env.SystemRoot, process.env.windir].find((value) => typeof value === 'string' && path.win32.isAbsolute(value)) || 'C:\\Windows'
  return path.win32.join(root, 'System32', name)
}

async function reg(args, ignoreFailure = false) {
  try {
    await execFileAsync(systemProgram('reg.exe'), args, { windowsHide: true })
  } catch (error) {
    if (!ignoreFailure) throw error
  }
}

function addDefault(key, value) {
  return ['ADD', key, '/ve', '/t', 'REG_SZ', '/d', value, '/f']
}

function addValue(key, name, value) {
  return ['ADD', key, '/v', name, '/t', 'REG_SZ', '/d', value, '/f']
}

/**
 * The reg.exe ADD commands that offer Simple in "Open with" for every
 * extension the format registry routes (per user, HKCU only). Nothing here
 * takes over an extension's default app.
 * @param {string} executablePath The portable simple.exe.
 * @param {string[]} [extensions] Defaults to the registry's routed extensions.
 * @returns {string[][]} Argument lists for reg.exe, in execution order.
 */
function associationPlan(executablePath, extensions = associationExtensions()) {
  const command = `"${executablePath}" "%1"`
  const icon = `"${executablePath}",0`
  const progIdKey = `${ROOT}\\${PROG_ID}`
  const plan = [
    addDefault(progIdKey, 'simple supported file'),
    addDefault(`${progIdKey}\\DefaultIcon`, icon),
    addDefault(`${progIdKey}\\shell\\open\\command`, command),
    addDefault(APPLICATION_KEY, 'simple'),
    addValue(APPLICATION_KEY, 'FriendlyAppName', 'simple'),
    addDefault(`${APPLICATION_KEY}\\DefaultIcon`, icon),
    addDefault(`${APPLICATION_KEY}\\shell\\open\\command`, command),
    addValue(CAPABILITIES_KEY, 'ApplicationName', 'simple'),
    addValue(CAPABILITIES_KEY, 'ApplicationDescription', 'One app for documents, spreadsheets, PDFs, images, videos, and text files.'),
  ]
  for (const extension of extensions) {
    plan.push(
      addValue(`${ROOT}\\${extension}\\OpenWithProgids`, PROG_ID, ''),
      addValue(`${APPLICATION_KEY}\\SupportedTypes`, extension, ''),
      addValue(`${CAPABILITIES_KEY}\\FileAssociations`, extension, PROG_ID),
    )
  }
  plan.push(addValue('HKCU\\Software\\RegisteredApplications', REGISTERED_NAME, 'Software\\Classes\\Applications\\simple.exe\\Capabilities'))
  return plan
}

/**
 * Registers Simple for every routed extension under HKCU.
 * @param {string} executablePath
 * @returns {Promise<{executablePath: string, extensionCount: number}>}
 */
async function registerAssociations(executablePath) {
  if (process.platform !== 'win32') throw new Error('File-type registration is only available on Windows.')
  const extensions = associationExtensions()
  for (const args of associationPlan(executablePath, extensions)) await reg(args)
  return { executablePath, extensionCount: extensions.length }
}

/**
 * Removes everything registerAssociations added. Missing keys are ignored.
 * @returns {Promise<true>}
 */
async function unregisterAssociations() {
  if (process.platform !== 'win32') throw new Error('File-type registration is only available on Windows.')
  for (const extension of associationExtensions()) {
    await reg(['DELETE', `${ROOT}\\${extension}\\OpenWithProgids`, '/v', PROG_ID, '/f'], true)
  }
  await reg(['DELETE', `${ROOT}\\${PROG_ID}`, '/f'], true)
  await reg(['DELETE', APPLICATION_KEY, '/f'], true)
  await reg(['DELETE', 'HKCU\\Software\\RegisteredApplications', '/v', REGISTERED_NAME, '/f'], true)
  return true
}

module.exports = { PROG_ID, associationPlan, registerAssociations, unregisterAssociations }
