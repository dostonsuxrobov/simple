'use strict'

const path = require('node:path')
const { spawn } = require('node:child_process')
const { app } = require('electron')
const { OPEN_LIST_SWITCH, needsOpenList, removeOpenListSync, writeOpenListSync } = require('./open-list.cjs')

function portableExecutablePath() {
  return process.env.PORTABLE_EXECUTABLE_FILE || process.execPath
}

function launchArguments(args) {
  if (process.env.PORTABLE_EXECUTABLE_FILE || app.isPackaged) return args
  return [app.getAppPath(), ...args]
}

/** The profile root every Simple process shares (bootstrap sets SIMPLE_USER_DATA_DIR). */
function stateRoot() {
  return process.env.SIMPLE_USER_DATA_DIR || app.getPath('userData')
}

/**
 * Plain wording for a failure to start Simple again.
 * @param {NodeJS.ErrnoException} error
 * @returns {string}
 */
function launchFailureMessage(error) {
  const name = path.basename(portableExecutablePath())
  if (error && (error.code === 'ENOENT' || error.code === 'ENOTDIR')) {
    return `Simple couldn't start another window because ${name} was moved, renamed or deleted. Start Simple again from its new place.`
  }
  if (error && (error.code === 'EACCES' || error.code === 'EPERM')) {
    return `Simple couldn't start another window because Windows blocked ${name}. Start Simple again, or check that it is allowed to run.`
  }
  return `Simple couldn't start another window${error && error.code ? ` (${error.code})` : ''}. Try again.`
}

/**
 * Starts Simple again, detached, with these arguments. File paths that would
 * make the command line too long for the portable wrapper travel in a list
 * file instead (see open-list.cjs).
 *
 * The process is created before this function returns (the list file is
 * written synchronously and child_process.spawn creates the process in the
 * same call), so a caller that quits right afterwards still starts it. The
 * bootstrap relies on that: a workspace that is already running makes this
 * process quit while its module loads.
 * @param {string[]} args flags and absolute file paths
 * @param {{spawn?: typeof spawn}} [options] spawn replaces child_process.spawn (tests)
 * @returns {Promise<number>} the new process id, once it has started
 * @throws {Error} code LAUNCH_FAILED (cause: the spawn error) when it could not start
 */
function launchDetached(args, options = {}) {
  const command = portableExecutablePath()
  let launchArgs = launchArguments(args)
  let listFile = null
  if (needsOpenList(command, launchArgs)) {
    const flags = args.filter((argument) => String(argument).startsWith('-'))
    const paths = args.filter((argument) => !String(argument).startsWith('-')).map((argument) => path.resolve(argument))
    try {
      listFile = writeOpenListSync(stateRoot(), paths)
    } catch (error) {
      return Promise.reject(Object.assign(new Error(`Simple couldn't pass ${paths.length.toLocaleString('en-US')} files to another window${error && error.code ? ` (${error.code})` : ''}. Try again with fewer files.`), { code: 'LAUNCH_FAILED', cause: error }))
    }
    launchArgs = launchArguments([...flags, `${OPEN_LIST_SWITCH}${listFile}`])
  }
  return new Promise((resolve, reject) => {
    let child
    let started = false
    const fail = (error) => {
      // An error after the process started (nothing in practice) must not
      // delete the list the new process is reading; the listener stays so it
      // never becomes an uncaught exception.
      if (started) return
      if (listFile) removeOpenListSync(listFile)
      reject(Object.assign(new Error(launchFailureMessage(error)), { code: 'LAUNCH_FAILED', cause: error }))
    }
    try {
      child = (options.spawn || spawn)(command, launchArgs, {
        detached: true,
        stdio: 'ignore',
        windowsHide: true,
      })
    } catch (error) {
      fail(error)
      return
    }
    child.on('error', fail)
    child.once('spawn', () => {
      started = true
      child.unref()
      resolve(child.pid)
    })
  })
}

module.exports = { launchDetached, launchFailureMessage, portableExecutablePath }
