'use strict'

const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')
const { spawn } = require('node:child_process')

const ROOT = path.resolve(__dirname, '..')
const SHARED_ROOT = path.join(ROOT, 'shared')
/** Unified-app folders the portable EXE packages directly; launcher/ also feeds the Combine worker bundle. */
const APP_ROOTS = Object.freeze({ launcher: path.join(ROOT, 'launcher'), electron: path.join(ROOT, 'electron') })
const SYNC_SHARED = path.join(ROOT, 'scripts', 'sync-shared.cjs')
const SHARED_MANIFEST = path.join(SHARED_ROOT, 'manifest.json')
const WORKSPACE = path.resolve(ROOT, '..')
const SOURCES = Object.freeze({
  docs: path.join(WORKSPACE, 'simple_doc_source'),
  calc: path.join(WORKSPACE, 'simple_calc_source'),
  pdf: path.join(WORKSPACE, 'simple_pdf_source'),
  image: path.join(WORKSPACE, 'simple_image_source'),
  video: path.join(WORKSPACE, 'simple_video_source'),
})
const ignored = /(^|[\\/])(node_modules|dist|release|tmp|\.git)([\\/]|$)/
const pending = new Set()
let timer = null
let sharedTimer = null
let sharedChanged = false
let appChanged = false
let sharedSyncing = false
let sharedSyncRequested = false
let building = false
let rebuildRequested = false
let manifestAtStart = null

/**
 * Runs a program without a shell, so a path with spaces (node.exe under
 * C:\Program Files) reaches Windows as one quoted argument.
 * @param {string} command an executable (process.execPath), never a .cmd file
 * @param {string[]} args
 * @param {{stdio?: import('node:child_process').StdioOptions}} [options]
 * @returns {Promise<void>} rejects when it exits with a nonzero code
 */
function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: ROOT,
      stdio: options.stdio || 'inherit',
      windowsHide: true,
      shell: false,
    })
    child.once('error', reject)
    child.once('exit', (code) => code === 0 ? resolve() : reject(new Error(`${path.basename(command)} ${args.map((argument) => path.basename(String(argument))).join(' ')} exited with ${code}`)))
  })
}

/** electron-builder's own command-line script, run with this Node (no npx.cmd, no shell). */
function electronBuilderCli() {
  return require.resolve('electron-builder/cli.js', { paths: [ROOT] })
}

function manifestHash() {
  try { return crypto.createHash('sha256').update(fs.readFileSync(SHARED_MANIFEST)).digest('hex') } catch { return null }
}

async function rebuild() {
  if (building) { rebuildRequested = true; return }
  building = true
  const modes = [...pending]
  pending.clear()
  const shared = sharedChanged
  sharedChanged = false
  const appFiles = appChanged
  appChanged = false
  try {
    const changed = [...modes, ...(shared ? ['shared I/O'] : []), ...(appFiles ? ['launcher'] : [])]
    console.log(`\nSource change detected in ${changed.join(', ')}. Rebuilding portable simple.exe...`)
    // The launcher requires simple/shared directly, so a shared-only change still repackages.
    // Staging any mode also rebuilds the Combine worker; a shared-only or
    // launcher-only change must rebuild it on its own, because the worker
    // bundles launcher and shared files (combineInputs in modules/manifest.json)
    // and verify.cjs rejects a stale one.
    if (modes.length) await run(process.execPath, [path.join(ROOT, 'scripts', 'sync-build.cjs'), `--modes=${modes.join(',')}`])
    else if (shared || appFiles) await run(process.execPath, [path.join(ROOT, 'scripts', 'sync-build.cjs'), '--combine-only'])
    await run(process.execPath, [path.join(ROOT, 'scripts', 'verify.cjs')])
    await run(process.execPath, [path.join(ROOT, 'scripts', 'prepare-portable-build.cjs')])
    await run(process.execPath, [electronBuilderCli(), '--win', 'portable', '--x64'])
    console.log(`Updated ${path.join(ROOT, 'release', 'simple.exe')}`)
  } catch (error) {
    console.error(error.stack || error)
  } finally {
    building = false
    if (rebuildRequested || pending.size) {
      rebuildRequested = false
      void rebuild()
    }
  }
}

function scheduleRebuild() {
  clearTimeout(timer)
  timer = setTimeout(() => void rebuild(), 1200)
}

// Vendors simple/shared into the enabled workspaces. The copies it writes are
// seen by the workspace watchers, which queue those modes for the rebuild.
// A changed manifest.json is never vendored automatically: it decides which
// workspaces and folders the sync writes to, so it is only checked, and the
// person who changed it runs npm run sync:shared once it is right.
async function syncShared() {
  if (sharedSyncing) { sharedSyncRequested = true; return }
  sharedSyncing = true
  try {
    if (manifestHash() !== manifestAtStart) {
      console.log('\nsimple/shared/manifest.json changed since npm run watch started, so nothing was vendored. Check the change, run npm run sync:shared, then restart npm run watch.')
      await run(process.execPath, [SYNC_SHARED, '--check']).catch(() => {})
      return
    }
    console.log('\nShared I/O change detected. Vendoring simple/shared...')
    await run(process.execPath, [SYNC_SHARED])
    sharedChanged = true
    scheduleRebuild()
  } catch (error) {
    console.error(error.stack || error)
  } finally {
    sharedSyncing = false
    if (sharedSyncRequested) {
      sharedSyncRequested = false
      void syncShared()
    }
  }
}

async function start() {
  console.log('Preparing all five workspaces before watching for updates...')
  manifestAtStart = manifestHash()
  await run(process.execPath, [SYNC_SHARED])
  await run(process.execPath, [path.join(ROOT, 'scripts', 'sync-build.cjs')])
  await run(process.execPath, [path.join(ROOT, 'scripts', 'verify.cjs')])

  for (const [mode, sourceRoot] of Object.entries(SOURCES)) {
    fs.watch(sourceRoot, { recursive: true }, (_event, fileName) => {
      if (!fileName || ignored.test(fileName)) return
      pending.add(mode)
      scheduleRebuild()
    })
  }
  fs.watch(SHARED_ROOT, { recursive: true }, () => {
    clearTimeout(sharedTimer)
    sharedTimer = setTimeout(() => void syncShared(), 400)
  })
  for (const folder of Object.values(APP_ROOTS)) {
    fs.watch(folder, { recursive: true }, (_event, fileName) => {
      if (!fileName || ignored.test(fileName)) return
      appChanged = true
      scheduleRebuild()
    })
  }

  console.log('Watching Docs, Calc, PDF, Image, Video, launcher and shared I/O sources. Press Ctrl+C to stop.')
}

if (require.main === module) {
  start().catch((error) => {
    console.error(error.stack || error)
    process.exitCode = 1
  })
}

module.exports = { electronBuilderCli, run }
