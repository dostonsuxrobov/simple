'use strict'

const { spawn } = require('node:child_process')
const { app } = require('electron')

function portableExecutablePath() {
  return process.env.PORTABLE_EXECUTABLE_FILE || process.execPath
}

function launchArguments(args) {
  if (process.env.PORTABLE_EXECUTABLE_FILE || app.isPackaged) return args
  return [app.getAppPath(), ...args]
}

function launchDetached(args) {
  const child = spawn(portableExecutablePath(), launchArguments(args), {
    detached: true,
    stdio: 'ignore',
    windowsHide: true,
  })
  child.unref()
  return child.pid
}

module.exports = { launchDetached, portableExecutablePath }
