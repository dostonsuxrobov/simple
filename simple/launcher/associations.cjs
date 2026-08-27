'use strict'

const { execFile } = require('node:child_process')
const { promisify } = require('node:util')
const { SUPPORTED_EXTENSIONS } = require('../electron/routing.cjs')

const execFileAsync = promisify(execFile)
const ROOT = 'HKCU\\Software\\Classes'
const PROG_ID = 'Simple.Unified.File'
const REGISTERED_NAME = 'simple'

async function reg(args, ignoreFailure = false) {
  try {
    await execFileAsync('reg.exe', args, { windowsHide: true })
  } catch (error) {
    if (!ignoreFailure) throw error
  }
}

function addDefault(key, value) {
  return reg(['ADD', key, '/ve', '/t', 'REG_SZ', '/d', value, '/f'])
}

function addValue(key, name, value) {
  return reg(['ADD', key, '/v', name, '/t', 'REG_SZ', '/d', value, '/f'])
}

async function registerAssociations(executablePath) {
  if (process.platform !== 'win32') throw new Error('File-type registration is only available on Windows.')
  const command = `"${executablePath}" "%1"`
  const icon = `"${executablePath}",0`
  const progIdKey = `${ROOT}\\${PROG_ID}`
  const applicationKey = `${ROOT}\\Applications\\simple.exe`
  const capabilitiesKey = `${applicationKey}\\Capabilities`

  await addDefault(progIdKey, 'simple supported file')
  await addDefault(`${progIdKey}\\DefaultIcon`, icon)
  await addDefault(`${progIdKey}\\shell\\open\\command`, command)
  await addDefault(applicationKey, 'simple')
  await addValue(applicationKey, 'FriendlyAppName', 'simple')
  await addDefault(`${applicationKey}\\DefaultIcon`, icon)
  await addDefault(`${applicationKey}\\shell\\open\\command`, command)
  await addValue(capabilitiesKey, 'ApplicationName', 'simple')
  await addValue(capabilitiesKey, 'ApplicationDescription', 'One app for documents, spreadsheets, PDFs, images, videos, and text files.')

  for (const extension of SUPPORTED_EXTENSIONS) {
    await addValue(`${ROOT}\\${extension}\\OpenWithProgids`, PROG_ID, '')
    await addValue(`${applicationKey}\\SupportedTypes`, extension, '')
    await addValue(`${capabilitiesKey}\\FileAssociations`, extension, PROG_ID)
  }

  await addValue('HKCU\\Software\\RegisteredApplications', REGISTERED_NAME, 'Software\\Classes\\Applications\\simple.exe\\Capabilities')
  return { executablePath, extensionCount: SUPPORTED_EXTENSIONS.length }
}

async function unregisterAssociations() {
  if (process.platform !== 'win32') throw new Error('File-type registration is only available on Windows.')
  for (const extension of SUPPORTED_EXTENSIONS) {
    await reg(['DELETE', `${ROOT}\\${extension}\\OpenWithProgids`, '/v', PROG_ID, '/f'], true)
  }
  await reg(['DELETE', `${ROOT}\\${PROG_ID}`, '/f'], true)
  await reg(['DELETE', `${ROOT}\\Applications\\simple.exe`, '/f'], true)
  await reg(['DELETE', 'HKCU\\Software\\RegisteredApplications', '/v', REGISTERED_NAME, '/f'], true)
  return true
}

module.exports = { PROG_ID, registerAssociations, unregisterAssociations }
