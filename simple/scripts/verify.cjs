'use strict'

const assert = require('node:assert/strict')
const crypto = require('node:crypto')
const fs = require('node:fs')
const path = require('node:path')
const {
  EXTENSIONS_BY_MODE,
  MODES,
  SUPPORTED_EXTENSIONS,
  explicitMode,
  groupPathsByMode,
  modeForPath,
  supportedPaths,
} = require('../electron/routing.cjs')
const formats = require('../shared/electron/formats.cjs')
const { syncShared } = require('./sync-shared.cjs')
const { checkLocalOnly } = require('./local-only-guard.cjs')

const ROOT = path.resolve(__dirname, '..')
const WORKSPACE = path.resolve(ROOT, '..')
const SOURCE_BY_MODE = Object.freeze({
  docs: 'simple_doc_source',
  calc: 'simple_calc_source',
  pdf: 'simple_pdf_source',
  image: 'simple_image_source',
  video: 'simple_video_source',
})
const IGNORED_FOLDERS = new Set(['node_modules', 'dist', 'release', 'tmp', '.git'])

// Vendored shared I/O copies must match simple/shared before anything else is
// trusted; a shared change then shows up below as a stale module fingerprint.
const sharedIo = syncShared({ check: true })
assert.ok(sharedIo.ok, sharedIo.message)

// Everything stays on this PC and documents open inside Simple: no network
// modules or requests, no cloud service code, no hand-off to other programs.
const localOnly = checkLocalOnly()
assert.ok(localOnly.ok, localOnly.message)

function filesUnder(directory, relative = '') {
  const output = []
  const entries = fs.readdirSync(path.join(directory, relative), { withFileTypes: true })
    .sort((a, b) => a.name.localeCompare(b.name))
  for (const entry of entries) {
    if (entry.name === '.simple-unified-lock-hash') continue
    const child = path.join(relative, entry.name)
    if (entry.isDirectory()) {
      if (!IGNORED_FOLDERS.has(entry.name)) output.push(...filesUnder(directory, child))
    } else if (entry.isFile()) {
      output.push(child)
    }
  }
  return output
}

function sourceFingerprint(sourceRoot) {
  const hash = crypto.createHash('sha256')
  for (const relative of filesUnder(sourceRoot)) {
    hash.update(relative.replaceAll('\\', '/'))
    hash.update('\0')
    hash.update(fs.readFileSync(path.join(sourceRoot, relative)))
    hash.update('\0')
  }
  return hash.digest('hex')
}

function sourceExtensionSet(relativeSourcePath) {
  const source = fs.readFileSync(path.join(WORKSPACE, relativeSourcePath), 'utf8')
  const declaration = /SUPPORTED_EXTENSIONS\s*=\s*(?:\/\*[\s\S]*?\*\/\s*)?new Set\(\s*\[([\s\S]*?)\]\s*\)/.exec(source)
  assert.ok(declaration, `${relativeSourcePath} must declare SUPPORTED_EXTENSIONS.`)
  const body = declaration[1].replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/[^\n]*/g, '')
  return new Set(body.match(/\.[a-z0-9]+/g) || [])
}

assert.equal(modeForPath('Report.DOCX'), 'docs')
assert.equal(modeForPath('Legacy.DOC'), 'docs')
assert.equal(modeForPath('archive.PDF'), 'pdf')
assert.equal(modeForPath('photo.JPEG'), 'image')
assert.equal(modeForPath('animation.WEBP'), 'image')
assert.equal(modeForPath('movie.MP4'), 'video')
assert.equal(modeForPath('clip.WEBM'), 'video')
assert.equal(modeForPath('notes.txt'), 'pdf')
assert.equal(modeForPath('budget.XLSX'), 'calc')
assert.equal(modeForPath('lotus.123'), 'calc')
assert.equal(modeForPath('README'), null)
assert.equal(explicitMode(['--simple-mode=CALC']), 'calc')
assert.equal(explicitMode(['--anything']), null)
assert.deepEqual(supportedPaths(['simple.exe', 'a.docx', 'legacy.DOC', 'b.xlsx', 'a.docx', '--flag']), ['a.docx', 'legacy.DOC', 'b.xlsx'])
assert.deepEqual(supportedPaths(['--inspect=trap.pdf', '-trap.png', 'real.pdf']), ['real.pdf'])
assert.deepEqual([...groupPathsByMode(['a.doc', 'b.docx', 'c.xlsx', 'd.pdf']).entries()], [
  ['docs', ['a.doc', 'b.docx']],
  ['calc', ['c.xlsx']],
  ['pdf', ['d.pdf']],
])

const flattened = Object.values(EXTENSIONS_BY_MODE).flat()
assert.equal(new Set(flattened).size, flattened.length, 'Extension ownership must be exclusive.')
assert.deepEqual(new Set(flattened), new Set(SUPPORTED_EXTENSIONS))

/** What each workspace's own source accepts on its command line and in Open. */
const ACCEPTED_BY_MODE = Object.freeze({
  image: new Set(require(path.join(WORKSPACE, SOURCE_BY_MODE.image, 'electron', 'image-files.cjs')).SUPPORTED_EXTENSIONS),
  video: new Set(require(path.join(WORKSPACE, SOURCE_BY_MODE.video, 'electron', 'routing.cjs')).SUPPORTED_EXTENSIONS),
  docs: new Set(require(path.join(WORKSPACE, SOURCE_BY_MODE.docs, 'electron', 'document-files.cjs')).SUPPORTED_EXTENSIONS),
  pdf: sourceExtensionSet(path.join(SOURCE_BY_MODE.pdf, 'electron', 'main.cjs')),
  calc: sourceExtensionSet(path.join(SOURCE_BY_MODE.calc, 'electron', 'workbooks.cjs')),
})
assert.deepEqual(Object.keys(ACCEPTED_BY_MODE).sort(), [...MODES].sort())
const WORKSPACE_NAMES = Object.freeze({ docs: 'Docs', calc: 'Calc', pdf: 'PDF', image: 'Image', video: 'Video' })
// Routing may only send a file to a workspace that accepts it: its extension
// route, and every workspace its content can route it to (a delimited .txt
// goes to Calc), or that workspace drops the path and opens an empty window.
// A workspace may accept more than the unified app routes to it yet: those
// formats are switched on in simple/shared/electron/formats.json when the
// workspace integrates them.
for (const extension of SUPPORTED_EXTENSIONS) {
  const candidates = formats.candidateModesForExtension(extension)
  assert.ok(candidates.includes(modeForPath(`C:\\nowhere\\file${extension}`)), `${extension} routes outside its candidate workspaces.`)
  for (const mode of candidates) {
    const by = EXTENSIONS_BY_MODE[mode].includes(extension) ? 'routing' : 'content routing'
    assert.ok(ACCEPTED_BY_MODE[mode].has(extension), `${WORKSPACE_NAMES[mode]} ${by} sends ${extension} to a source that no longer accepts it.`)
  }
}
const notRoutedYet = []
for (const mode of ['image', 'video', 'docs']) {
  const extra = [...ACCEPTED_BY_MODE[mode]].filter((extension) => !EXTENSIONS_BY_MODE[mode].includes(extension))
  if (extra.length) notRoutedYet.push(`${mode}: ${extra.sort().join(' ')}`)
}

const packageJson = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))
const associated = packageJson.build.fileAssociations.flatMap((item) => Array.isArray(item.ext) ? item.ext : [item.ext]).map((ext) => `.${ext}`)
assert.deepEqual(new Set(associated), new Set(SUPPORTED_EXTENSIONS), 'Package associations must cover every routed extension.')
assert.ok(packageJson.build.fileAssociations.every((item) => item.icon === 'build/icon.ico'))

const sourcePng = fs.readFileSync(path.join(ROOT, 'assets', 'icon-source.png'))
assert.equal(crypto.createHash('sha256').update(sourcePng).digest('hex'), '9007a93c6cccd9e8b6abda5a8c889d54cdb61120d071aff63c9cc62483ec4ff9')

const sharedIco = fs.readFileSync(path.join(ROOT, 'build', 'icon.ico'))
const sharedIcoHash = crypto.createHash('sha256').update(sharedIco).digest('hex')
for (const mode of MODES) {
  const moduleRoot = path.join(ROOT, 'modules', mode)
  for (const required of ['dist/index.html', 'electron/main.cjs', 'electron/preload.cjs', 'build/icon.ico']) {
    assert.ok(fs.existsSync(path.join(moduleRoot, required)), `${mode} is missing ${required}`)
  }
  const moduleIconHash = crypto.createHash('sha256').update(fs.readFileSync(path.join(moduleRoot, 'build', 'icon.ico'))).digest('hex')
  assert.equal(moduleIconHash, sharedIcoHash, `${mode} does not use the shared icon.`)
}

const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'modules', 'manifest.json'), 'utf8'))
// The bundled Combine worker records every source esbuild read for it.
const combineInputs = manifest.shared?.combineInputs
assert.ok(combineInputs && Object.keys(combineInputs).length, 'The Combine worker was built by an older sync; run npm run sync.')
for (const [source, recorded] of Object.entries(combineInputs)) {
  let current = null
  try { current = crypto.createHash('sha256').update(fs.readFileSync(path.join(ROOT, source))).digest('hex') } catch {}
  assert.equal(current, recorded, `Combine is stale (${source}); run npm run sync.`)
}
assert.equal(manifest.shared?.combineWorkerHash, crypto.createHash('sha256').update(fs.readFileSync(path.join(ROOT, 'modules', 'shared', 'combine-worker.cjs'))).digest('hex'), 'The bundled Combine worker differs from the manifest.')
assert.deepEqual(Object.keys(manifest.modules).sort(), [...MODES].sort())
assert.deepEqual(Object.keys(SOURCE_BY_MODE).sort(), [...MODES].sort())
for (const mode of MODES) {
  const sourceFolder = SOURCE_BY_MODE[mode]
  assert.equal(manifest.modules[mode].source, sourceFolder, `${mode} points to the wrong source folder.`)
  assert.equal(
    manifest.modules[mode].sourceHash,
    sourceFingerprint(path.join(WORKSPACE, sourceFolder)),
    `${mode} is stale; run npm run sync.`,
  )
}
console.log(`Verified ${SUPPORTED_EXTENSIONS.length} routes, five current source builds, one shared icon, ${localOnly.files} local-only app files, and shared I/O code for ${sharedIo.workspaces.length ? sharedIo.workspaces.join(', ') : 'no enabled workspaces'}.`)
if (notRoutedYet.length) console.log(`Accepted by a workspace but not routed by the unified app yet (switch them on in formats.json when integrated): ${notRoutedYet.join('; ')}.`)
