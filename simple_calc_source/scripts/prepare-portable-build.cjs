const fs = require('node:fs')
const path = require('node:path')

const projectRoot = path.resolve(__dirname, '..')
const sourcePath = path.join(projectRoot, 'build', 'portable-fast.nsi')
const targetPath = path.join(
  projectRoot,
  'node_modules',
  'app-builder-lib',
  'templates',
  'nsis',
  'portable.nsi',
)
const electronGetPath = path.join(
  projectRoot,
  'node_modules',
  'app-builder-lib',
  'out',
  'util',
  'electronGet.js',
)

const upstream = fs.readFileSync(targetPath, 'utf8')
if (!upstream.includes('!include "common.nsh"') || !upstream.includes('extractEmbeddedAppPackage')) {
  throw new Error('The installed Electron Builder portable template is incompatible with simple_calc’s fast launcher.')
}

const optimized = fs.readFileSync(sourcePath, 'utf8')
if (!optimized.includes('.simple-calc-runtime-ready') || !optimized.includes('useCachedRuntime:')) {
  throw new Error('The optimized portable launcher template is incomplete.')
}

fs.writeFileSync(targetPath, optimized)

// Antivirus and indexers can briefly hold Electron's freshly extracted files
// on Windows. The installed builder also protects extraction with a lock; that
// lock must be released before renaming its directory on Windows. Keep the
// upstream flow, then retry only transient filesystem errors.
let electronGetSource = fs.readFileSync(electronGetPath, 'utf8')
const retryMarker = 'simple_calc: retry transient Windows rename locks'
if (!electronGetSource.includes(retryMarker)) {
  const renameNeedle = '        await fs.rename(tmpDir, dir);'
  if (!electronGetSource.includes(renameNeedle)) {
    throw new Error('The installed Electron Builder extraction flow is incompatible with simple_calc’s Windows rename retry.')
  }
  const renameReplacement = `        // ${retryMarker}\n        for (let attempt = 0; ; attempt += 1) {\n            try {\n                await fs.rename(tmpDir, dir);\n                break;\n            }\n            catch (error) {\n                const transient = process.platform === "win32" && ["EPERM", "EACCES", "EBUSY", "ENOTEMPTY"].includes(error && error.code);\n                if (!transient || attempt >= 29) throw error;\n                await new Promise(resolve => setTimeout(resolve, Math.min(1000, 200 * (attempt + 1))));\n            }\n        }`
  electronGetSource = electronGetSource.replace(renameNeedle, renameReplacement)
}

const releaseMarker = 'simple_calc: release extraction lock before Windows rename'
if (!electronGetSource.includes(releaseMarker)) {
  const lockSetupNeedle = '        stale: 120000, // Increased from 60s to allow long-running extractions\n    });\n    try {'
  const renameMarkerNeedle = `        // ${retryMarker}`
  const finallyNeedle = '        await release().catch(err => builder_util_1.log.warn({ err }, "failed to release lockfile"));'
  if (!electronGetSource.includes(lockSetupNeedle) || !electronGetSource.includes(renameMarkerNeedle) || !electronGetSource.includes(finallyNeedle)) {
    throw new Error('The installed Electron Builder lock flow is incompatible with simple_calc’s Windows extraction fix.')
  }
  electronGetSource = electronGetSource
    .replace(lockSetupNeedle, `${lockSetupNeedle.slice(0, -6)}    let simpleCalcExtractionLockReleased = false;\n    try {`)
    .replace(renameMarkerNeedle, `        // ${releaseMarker}\n        await release();\n        simpleCalcExtractionLockReleased = true;\n${renameMarkerNeedle}`)
    .replace(finallyNeedle, `        if (!simpleCalcExtractionLockReleased) ${finallyNeedle.trim()}`)
}

const copyMarker = 'simple_calc: copy extracted runtime when Windows holds a directory handle'
if (!electronGetSource.includes(copyMarker)) {
  const retryRenameNeedle = '                await fs.rename(tmpDir, dir);'
  if (!electronGetSource.includes(retryRenameNeedle)) {
    throw new Error('The installed Electron Builder rename flow is incompatible with simple_calc’s Windows staging fix.')
  }
  electronGetSource = electronGetSource.replace(
    retryRenameNeedle,
    `                // ${copyMarker}\n                if (process.platform === "win32") await fs.cp(tmpDir, dir, { recursive: true, force: true });\n                else await fs.rename(tmpDir, dir);`,
  )
}
fs.writeFileSync(electronGetPath, electronGetSource)

console.log('Prepared the fast, cached portable launcher and resilient Windows packaging.')
