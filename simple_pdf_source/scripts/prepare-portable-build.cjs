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

const upstream = fs.readFileSync(targetPath, 'utf8')
if (!upstream.includes('!include "common.nsh"') || !upstream.includes('extractEmbeddedAppPackage')) {
  throw new Error('The installed Electron Builder portable template is incompatible with simple’s fast launcher.')
}

const optimized = fs.readFileSync(sourcePath, 'utf8')
if (!optimized.includes('.simple-runtime-ready') || !optimized.includes('useCachedRuntime:')) {
  throw new Error('The optimized portable launcher template is incomplete.')
}

fs.writeFileSync(targetPath, optimized)
console.log('Prepared the fast, cached portable launcher.')
