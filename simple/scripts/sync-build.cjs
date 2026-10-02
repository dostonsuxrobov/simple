'use strict'

const crypto = require('node:crypto')
const fs = require('node:fs/promises')
const path = require('node:path')
const { spawn } = require('node:child_process')
const esbuild = require('esbuild')
const { syncShared } = require('./sync-shared.cjs')

const ROOT = path.resolve(__dirname, '..')
const WORKSPACE = path.resolve(ROOT, '..')
const MODULES_ROOT = path.join(ROOT, 'modules')
const SHARED_PNG = path.join(ROOT, 'build', 'icon.png')
const SHARED_ICO = path.join(ROOT, 'build', 'icon.ico')
const CONFIG = Object.freeze({
  docs: { folder: 'simple_doc_source' },
  calc: { folder: 'simple_calc_source', replaceRendererIcon: true },
  pdf: { folder: 'simple_pdf_source' },
  image: { folder: 'simple_image_source' },
  video: { folder: 'simple_video_source' },
})
const IGNORED_FOLDERS = new Set(['node_modules', 'dist', 'release', 'tmp', '.git'])

function selectedModes() {
  const argument = process.argv.find((item) => item.startsWith('--modes='))
  if (!argument) return Object.keys(CONFIG)
  const modes = argument.slice('--modes='.length).split(',').map((item) => item.trim()).filter(Boolean)
  for (const mode of modes) if (!CONFIG[mode]) throw new Error(`Unknown mode: ${mode}`)
  return [...new Set(modes)]
}

function run(command, args, cwd) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd,
      stdio: 'inherit',
      windowsHide: true,
      shell: process.platform === 'win32',
    })
    child.once('error', reject)
    child.once('exit', (code) => code === 0 ? resolve() : reject(new Error(`${command} exited with code ${code}`)))
  })
}

async function sha256File(filePath) {
  return crypto.createHash('sha256').update(await fs.readFile(filePath)).digest('hex')
}

async function filesUnder(directory, relative = '') {
  const entries = await fs.readdir(path.join(directory, relative), { withFileTypes: true })
  const output = []
  for (const entry of entries.sort((a, b) => a.name.localeCompare(b.name))) {
    if (entry.name === '.simple-unified-lock-hash') continue
    const child = path.join(relative, entry.name)
    if (entry.isDirectory()) {
      if (!IGNORED_FOLDERS.has(entry.name)) output.push(...await filesUnder(directory, child))
    } else if (entry.isFile()) {
      output.push(child)
    }
  }
  return output
}

async function sourceFingerprint(sourceRoot) {
  const hash = crypto.createHash('sha256')
  for (const relative of await filesUnder(sourceRoot)) {
    hash.update(relative.replaceAll('\\', '/'))
    hash.update('\0')
    hash.update(await fs.readFile(path.join(sourceRoot, relative)))
    hash.update('\0')
  }
  return hash.digest('hex')
}

async function ensureDependencies(sourceRoot) {
  const lockPath = path.join(sourceRoot, 'package-lock.json')
  const lockHash = await sha256File(lockPath)
  const markerPath = path.join(sourceRoot, 'node_modules', '.simple-unified-lock-hash')
  let installedHash = ''
  try { installedHash = (await fs.readFile(markerPath, 'utf8')).trim() } catch {}
  if (installedHash === lockHash) return lockHash
  console.log(`\nInstalling locked dependencies for ${path.basename(sourceRoot)}...`)
  await run(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['ci', '--no-audit', '--no-fund'], sourceRoot)
  await fs.writeFile(markerPath, `${lockHash}\n`, 'utf8')
  return lockHash
}

async function replaceCalcRendererIcon(distRoot, sourceRoot) {
  const originalHash = await sha256File(path.join(sourceRoot, 'icon.png'))
  const candidates = []
  async function visit(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      const target = path.join(directory, entry.name)
      if (entry.isDirectory()) await visit(target)
      else if (entry.isFile() && path.extname(entry.name).toLowerCase() === '.png') candidates.push(target)
    }
  }
  await visit(distRoot)
  let replaced = 0
  for (const candidate of candidates) {
    if (await sha256File(candidate) !== originalHash) continue
    await fs.copyFile(SHARED_PNG, candidate)
    replaced += 1
  }
  if (!replaced) console.warn('Calc renderer icon asset was not present; window and file icons are still unified.')
  return replaced
}

async function stageMode(mode, lockHash) {
  const config = CONFIG[mode]
  const sourceRoot = path.join(WORKSPACE, config.folder)
  const temporaryRoot = path.join(ROOT, '.stage', `${mode}-${process.pid}`)
  const destinationRoot = path.join(MODULES_ROOT, mode)
  await fs.rm(temporaryRoot, { recursive: true, force: true })
  await fs.mkdir(path.join(temporaryRoot, 'electron'), { recursive: true })
  await fs.mkdir(path.join(temporaryRoot, 'build'), { recursive: true })

  console.log(`\nBuilding ${mode} renderer from ${config.folder}...`)
  await run(process.platform === 'win32' ? 'npm.cmd' : 'npm', ['run', 'build:web'], sourceRoot)
  await fs.cp(path.join(sourceRoot, 'dist'), path.join(temporaryRoot, 'dist'), { recursive: true })
  await fs.copyFile(path.join(sourceRoot, 'electron', 'preload.cjs'), path.join(temporaryRoot, 'electron', 'preload.cjs'))
  await fs.copyFile(SHARED_ICO, path.join(temporaryRoot, 'build', 'icon.ico'))

  console.log(`Bundling ${mode} backend...`)
  await esbuild.build({
    entryPoints: [path.join(sourceRoot, 'electron', 'main.cjs')],
    outfile: path.join(temporaryRoot, 'electron', 'main.cjs'),
    bundle: true,
    platform: 'node',
    format: 'cjs',
    target: 'node22',
    external: ['electron', 'mupdf'],
    legalComments: 'none',
    sourcemap: false,
  })

  if (mode === 'pdf') {
    // MuPDF's ES module loads its WASM beside import.meta.url. Preserve the
    // package layout instead of bundling it into a CommonJS backend.
    await fs.cp(path.join(sourceRoot, 'node_modules', 'mupdf'),
      path.join(temporaryRoot, 'vendor', 'mupdf'), { recursive: true })
  }

  const rendererIconsReplaced = config.replaceRendererIcon
    ? await replaceCalcRendererIcon(path.join(temporaryRoot, 'dist'), sourceRoot)
    : 0
  const packageJson = JSON.parse(await fs.readFile(path.join(sourceRoot, 'package.json'), 'utf8'))
  const fingerprint = await sourceFingerprint(sourceRoot)
  const backendHash = await sha256File(path.join(temporaryRoot, 'electron', 'main.cjs'))

  await fs.rm(destinationRoot, { recursive: true, force: true })
  await fs.rename(temporaryRoot, destinationRoot)
  return {
    source: config.folder,
    version: packageJson.version,
    sourceHash: fingerprint,
    lockHash,
    backendHash,
    rendererIconsReplaced,
  }
}

/**
 * Bundles the launcher's Combine worker into modules/shared/combine-worker.cjs
 * and records the hash of every file esbuild read for it (launcher, shared
 * I/O and workspace sources), so verify.cjs notices any change that would
 * make the bundled worker stale.
 * @returns {Promise<{combineInputs: Record<string, string>, combineWorkerHash: string}>}
 */
async function bundleCombineWorker() {
  const outfile = path.join(MODULES_ROOT, 'shared', 'combine-worker.cjs')
  const result = await esbuild.build({
    entryPoints: [path.join(ROOT, 'launcher', 'combine-worker.cjs')],
    outfile,
    bundle: true, platform: 'node', format: 'cjs', target: 'node22', legalComments: 'none',
    metafile: true,
  })
  const inputs = Object.keys(result.metafile.inputs)
    .map((input) => path.resolve(process.cwd(), input))
    .filter((input) => !input.split(path.sep).includes('node_modules'))
    .sort()
  const combineInputs = {}
  for (const input of inputs) combineInputs[path.relative(ROOT, input).replaceAll('\\', '/')] = await sha256File(input)
  return { combineInputs, combineWorkerHash: await sha256File(outfile) }
}

async function readManifest() {
  try { return JSON.parse(await fs.readFile(path.join(MODULES_ROOT, 'manifest.json'), 'utf8')) } catch { return { modules: {} } }
}

async function main() {
  // A workspace whose vendored shared I/O code drifted from simple/shared must not ship.
  const sharedIo = syncShared({ check: true })
  if (!sharedIo.ok) {
    console.error(sharedIo.message)
    process.exitCode = 1
    return
  }
  await fs.access(SHARED_PNG)
  await fs.access(SHARED_ICO)
  await fs.mkdir(MODULES_ROOT, { recursive: true })
  await fs.mkdir(path.join(ROOT, '.stage'), { recursive: true })
  const modes = selectedModes()
  const sourceRoots = Object.fromEntries(modes.map((mode) => [mode, path.join(WORKSPACE, CONFIG[mode].folder)]))
  for (const [mode, sourceRoot] of Object.entries(sourceRoots)) {
    try { await fs.access(path.join(sourceRoot, 'package.json')) } catch { throw new Error(`Missing ${mode} source at ${sourceRoot}`) }
  }

  const combineOnly = process.argv.includes('--combine-only')
  const previous = await readManifest()
  let modules = previous.modules || {}
  if (!combineOnly) {
    const lockEntries = await Promise.all(modes.map(async (mode) => [mode, await ensureDependencies(sourceRoots[mode])]))
    const lockHashes = Object.fromEntries(lockEntries)
    const results = await Promise.all(modes.map(async (mode) => [mode, await stageMode(mode, lockHashes[mode])]))
    modules = { ...modules, ...Object.fromEntries(results) }
  }
  const combine = await bundleCombineWorker()
  const manifest = {
    schemaVersion: 1,
    generatedAt: new Date().toISOString(),
    modules,
    shared: combine,
  }
  await fs.writeFile(path.join(MODULES_ROOT, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`, 'utf8')
  await fs.rm(path.join(ROOT, '.stage'), { recursive: true, force: true })
  console.log(combineOnly ? `\nRebuilt the Combine worker in ${MODULES_ROOT}` : `\nSynchronized ${modes.join(', ')} into ${MODULES_ROOT}`)
}

main().catch((error) => {
  console.error(error.stack || error)
  process.exitCode = 1
})
