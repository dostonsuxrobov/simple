'use strict'

// Opening files from the launcher (its Open button, a drop on its window, or
// a second instance): one new Simple process per workspace, however many
// files. Each path is routed once, here, without blocking the launcher's
// window (content samples are read asynchronously), and the decision travels
// with --simple-mode=<workspace> so the new process never reads a file to
// decide again. Electron is only needed for the real launch.

const path = require('node:path')
const { groupPathsByModeAsync, supportedPaths } = require('../electron/routing.cjs')

/**
 * Opens files in their workspaces and reports what really started.
 * @param {string[]} paths
 * @param {object} [options]
 * @param {string} [options.cwd] resolves relative paths (a second instance's working folder)
 * @param {(args: string[]) => Promise<number>} [options.launch] replaces launchDetached (tests)
 * @returns {Promise<{opened: number, unsupported: number, failed: number, message?: string}>}
 * @throws {Error} with a plain message when no workspace could be started
 */
async function launchValidatedPaths(paths, options = {}) {
  const launch = options.launch || require('../electron/launch.cjs').launchDetached
  const valid = supportedPaths(paths).map((filePath) => path.resolve(options.cwd || process.cwd(), filePath))
  const groups = [...(await groupPathsByModeAsync(valid)).entries()]
  const results = await Promise.allSettled(groups.map(([mode, pathsForMode]) => launch([`--simple-mode=${mode}`, ...pathsForMode])))
  let opened = 0
  let failed = 0
  let message
  results.forEach((result, index) => {
    const count = groups[index][1].length
    if (result.status === 'fulfilled') opened += count
    else {
      failed += count
      message ??= result.reason?.message
    }
  })
  if (failed && !opened) throw new Error(message || "Simple couldn't open these files. Try again.")
  return { opened, unsupported: Math.max(0, paths.length - valid.length), failed, ...(failed ? { message } : {}) }
}

module.exports = { launchValidatedPaths }
