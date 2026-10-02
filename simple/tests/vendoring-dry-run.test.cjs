'use strict'

// Vendoring dry run (scripts/check-vendoring.cjs): every workspace is switched
// on in a temporary copy of the manifest and vendored into a temporary tree.
// The vendored electron/simple-io files load, each real preload keeps its own
// bridge next to window.simpleIO, the modules bundle with esbuild as the
// unified build bundles them, and src/simple-io type-checks with each
// workspace's own TypeScript settings. No workspace folder is written.

const test = require('node:test')
const assert = require('node:assert/strict')
const { checkVendoring, generatedMain } = require('../scripts/check-vendoring.cjs')
const { loadManifest } = require('../scripts/sync-shared.cjs')

test('the generated wiring contains every call the manifest requires', () => {
  const manifest = loadManifest()
  for (const [name, workspace] of Object.entries(manifest.workspaces)) {
    const main = generatedMain(name, workspace)
    for (const rule of workspace.wiring) assert.ok(main.includes(rule.call), `${name}: ${rule.call}`)
  }
})

test('every workspace vendors cleanly, and its preload and modules load (Node, TypeScript, Electron)', { timeout: 300_000 }, async (t) => {
  let result
  try {
    result = await checkVendoring({ types: true, electron: process.platform === 'win32' })
  } catch (error) {
    if (error && error.constructor && error.constructor.name === 'CheckSetupError') {
      t.skip(`a workspace can't be type-checked here: ${error.message}`)
      return
    }
    throw error
  }
  assert.equal(result.ok, true, `${result.problems.join('\n')}\n${result.lines.join('\n')}`)
  const text = result.lines.join('\n')
  for (const name of Object.keys(loadManifest().workspaces)) {
    assert.match(text, new RegExp(`^${name}: preload exposes .*simpleIO`, 'm'))
    assert.match(text, new RegExp(`^${name}: \\d+ electron/simple-io modules bundle with esbuild`, 'm'))
    if (process.platform === 'win32') assert.match(text, new RegExp(`^${name}: Electron loaded \\d+ modules, registered (\\d+) channels \\(bundled: \\1\\)`, 'm'))
  }
  assert.match(text, /type-checks in pdf, calc, docs, image/)
})
