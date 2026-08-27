'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { spawn } = require('node:child_process')

const ROOT = path.resolve(__dirname, '..')
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
let building = false
let rebuildRequested = false

function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: ROOT,
      stdio: 'inherit',
      windowsHide: true,
      shell: process.platform === 'win32',
    })
    child.once('error', reject)
    child.once('exit', (code) => code === 0 ? resolve() : reject(new Error(`${command} exited with ${code}`)))
  })
}

async function rebuild() {
  if (building) { rebuildRequested = true; return }
  building = true
  const modes = [...pending]
  pending.clear()
  try {
    console.log(`\nSource change detected in ${modes.join(', ')}. Rebuilding portable simple.exe...`)
    await run(process.execPath, [path.join(ROOT, 'scripts', 'sync-build.cjs'), `--modes=${modes.join(',')}`])
    await run(process.execPath, [path.join(ROOT, 'scripts', 'verify.cjs')])
    await run(process.execPath, [path.join(ROOT, 'scripts', 'prepare-portable-build.cjs')])
    await run(process.platform === 'win32' ? 'npx.cmd' : 'npx', ['electron-builder', '--win', 'portable', '--x64'])
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

async function start() {
  console.log('Preparing all five workspaces before watching for updates...')
  await run(process.execPath, [path.join(ROOT, 'scripts', 'sync-build.cjs')])
  await run(process.execPath, [path.join(ROOT, 'scripts', 'verify.cjs')])

  for (const [mode, sourceRoot] of Object.entries(SOURCES)) {
    fs.watch(sourceRoot, { recursive: true }, (_event, fileName) => {
      if (!fileName || ignored.test(fileName)) return
      pending.add(mode)
      clearTimeout(timer)
      timer = setTimeout(() => void rebuild(), 1200)
    })
  }

  console.log('Watching Docs, Calc, PDF, Image, and Video sources. Press Ctrl+C to stop.')
}

start().catch((error) => {
  console.error(error.stack || error)
  process.exitCode = 1
})
