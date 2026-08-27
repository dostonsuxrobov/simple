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

assert.equal(modeForPath('Report.DOCX'), 'docs')
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
assert.deepEqual(supportedPaths(['simple.exe', 'a.docx', 'b.xlsx', 'a.docx', '--flag']), ['a.docx', 'b.xlsx'])
assert.deepEqual([...groupPathsByMode(['a.docx', 'b.xlsx', 'c.pdf']).keys()], ['docs', 'calc', 'pdf'])

const flattened = Object.values(EXTENSIONS_BY_MODE).flat()
assert.equal(new Set(flattened).size, flattened.length, 'Extension ownership must be exclusive.')
assert.deepEqual(new Set(flattened), new Set(SUPPORTED_EXTENSIONS))

const imageSourceExtensions = require(path.join(WORKSPACE, SOURCE_BY_MODE.image, 'electron', 'image-files.cjs')).SUPPORTED_EXTENSIONS
const videoSourceExtensions = require(path.join(WORKSPACE, SOURCE_BY_MODE.video, 'electron', 'routing.cjs')).SUPPORTED_EXTENSIONS
assert.deepEqual(new Set(imageSourceExtensions), new Set(EXTENSIONS_BY_MODE.image), 'Image routing must match the image source.')
assert.deepEqual(new Set(videoSourceExtensions), new Set(EXTENSIONS_BY_MODE.video), 'Video routing must match the video source.')

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
console.log(`Verified ${SUPPORTED_EXTENSIONS.length} routes, five current source builds, and one shared icon.`)
