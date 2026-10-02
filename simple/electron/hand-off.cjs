'use strict'

// Hands the files of one invocation that belong to other workspaces to new
// Simple processes (the bootstrap keeps its own workspace's files), so no
// file of a mixed drop or Send to is lost without a word:
// - each process is started in the same tick (launchDetached creates it
//   before it returns), so this process may quit right after: its own
//   workspace may already be running, and the module then quits while it
//   loads;
// - a workspace that could not start is logged to stderr and reported in a
//   message box once the app is ready, naming its files; a quit waits until
//   every start has settled and that message was shown.

const path = require('node:path')

/** How many file names a failure message lists before "and N more". */
const MAX_NAMED_FILES = 10

/**
 * The message box for workspaces that could not start.
 * @param {Array<{label: string, paths: string[], message: string}>} failures
 * @returns {Electron.MessageBoxOptions}
 */
function failureMessageBox(failures) {
  const count = failures.reduce((total, failure) => total + failure.paths.length, 0)
  const where = [...new Set(failures.map((failure) => failure.label))].join(' and ')
  const names = failures.flatMap((failure) => failure.paths.map((filePath) => path.basename(filePath)))
  const listed = names.slice(0, MAX_NAMED_FILES)
  const more = names.length - listed.length
  const reasons = [...new Set(failures.map((failure) => failure.message).filter(Boolean))]
  return {
    type: 'warning',
    title: 'simple',
    message: `Simple couldn't open ${count === 1 ? '1 file' : `${count.toLocaleString('en-US')} files`} in ${where}.`,
    detail: [
      ...reasons,
      '',
      ...listed,
      ...(more > 0 ? [`and ${more.toLocaleString('en-US')} more`] : []),
    ].join('\n'),
    buttons: ['OK'],
    noLink: true,
  }
}

/**
 * Creates the hand-off for one bootstrap process.
 * @param {object} deps
 * @param {Electron.App} deps.app
 * @param {{showMessageBox: (options: Electron.MessageBoxOptions) => Promise<unknown>}} deps.dialog
 *   looked up when a message is shown
 * @param {(args: string[]) => Promise<number>} deps.launch launchDetached
 * @param {(mode: string) => string} [deps.label] the workspace's name for people
 * @param {(text: string) => void} [deps.log]
 * @returns {{start: (mode: string, paths: string[]) => Promise<boolean>, settled: () => Promise<void>}}
 *   start resolves true once the workspace started, false when it was reported as failed
 */
function createHandOff({ app, dialog, launch, label = (mode) => mode, log = (text) => process.stderr.write(text) }) {
  const pending = new Set()
  const failures = []
  let reporting = Promise.resolve()
  let showing = 0
  let holding = false

  function report() {
    reporting = reporting.then(async () => {
      await app.whenReady()
      while (failures.length) {
        const batch = failures.splice(0)
        showing += 1
        try {
          await dialog.showMessageBox(failureMessageBox(batch))
        } catch (error) {
          log(`Could not show the message about files that did not open: ${error && error.message}\n`)
        } finally {
          showing -= 1
        }
      }
    })
    return reporting
  }

  function busy() {
    return pending.size > 0 || failures.length > 0 || showing > 0
  }

  async function settled() {
    while (busy()) {
      await Promise.allSettled([...pending])
      await reporting
    }
  }

  app.on('will-quit', (event) => {
    if (!busy()) return
    event.preventDefault()
    if (holding) return
    holding = true
    settled().finally(() => {
      holding = false
      app.quit()
    })
  })

  function start(mode, paths) {
    let launched
    try {
      launched = launch([`--simple-mode=${mode}`, ...paths])
    } catch (error) {
      launched = Promise.reject(error)
    }
    const tracked = Promise.resolve(launched).then(() => true, (error) => {
      const message = (error && error.message) || "Simple couldn't start another window. Try again."
      log(`Could not open ${paths.length} file(s) in ${mode}: ${message}\n`)
      failures.push({ label: label(mode), paths, message })
      void report()
      return false
    }).finally(() => pending.delete(tracked))
    pending.add(tracked)
    return tracked
  }

  return { start, settled }
}

module.exports = { MAX_NAMED_FILES, createHandOff, failureMessageBox }
