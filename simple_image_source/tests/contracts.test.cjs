'use strict'
// WP1 contracts: the shared TypeScript contract files load under Node's built-in type stripping
// (no build step), their runtime constants are what every package relies on, and the tool registry
// seed covers every tool id.
const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping (not disabled by --no-experimental-strip-types), '
  + 'or run node with --experimental-strip-types.'

function canStripTypes() {
  return Boolean(process.features && process.features.typescript)
}

function load(relative) {
  if (!canStripTypes()) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

test('Node can load the TypeScript sources (built-in type stripping)', () => {
  assert.ok(canStripTypes(), STRIP_HELP)
})

test('the compiler enforces erasable TypeScript and explicit type-only imports', () => {
  const config = JSON.parse(fs.readFileSync(path.join(root, 'tsconfig.json'), 'utf8'))
  assert.equal(config.compilerOptions.erasableSyntaxOnly, true)
  assert.equal(config.compilerOptions.verbatimModuleSyntax, true)
  assert.equal(config.compilerOptions.allowImportingTsExtensions, true)
  const vite = fs.readFileSync(path.join(root, 'vite.config.ts'), 'utf8')
  assert.match(vite, /worker:\s*\{\s*format:\s*'es'/, 'module workers that code-split need worker.format "es"')
})

test('imaging and session contracts are type-only at runtime', () => {
  assert.deepEqual(Object.keys(load('imaging/types.ts')), [])
  assert.deepEqual(Object.keys(load('shared/session.ts')), [])
})

test('advanced contract constants', () => {
  const advanced = load('advanced/types.ts')
  assert.equal(advanced.TILE_SIZE, 256)
  assert.equal(advanced.TILE_SHIFT, 8)
  assert.equal(1 << advanced.TILE_SHIFT, advanced.TILE_SIZE)
  assert.equal(advanced.COALESCE_MS, 1200)
  assert.ok(Object.isFrozen(advanced.LIMITS))
  assert.deepEqual({ ...advanced.LIMITS }, {
    maxDimension: 20_000,
    maxPixels: 50_000_000,
    maxLayers: 200,
    historyMaxEntries: 50,
    historyBudgetBytes: 768 * 1024 * 1024,
    documentBudgetBytes: 3 * 1024 * 1024 * 1024,
    compositeTileCacheTiles: 512,
    psdMemoryLimitBytes: 1536 * 1024 * 1024,
    frameBudgetMs: 8,
  })
  assert.ok(Object.isFrozen(advanced.DEFAULT_LOCKS))
  assert.deepEqual({ ...advanced.DEFAULT_LOCKS }, { pixels: false, position: false, transparency: false })
  assert.ok(Object.isFrozen(advanced.BACKGROUND_LOCKS))
  assert.deepEqual({ ...advanced.BACKGROUND_LOCKS }, { pixels: false, position: true, transparency: true })
})

test('the blend mode menu lists all 27 Photoshop modes once, in menu order', () => {
  const { BLEND_MODE_MENU } = load('advanced/types.ts')
  assert.ok(Object.isFrozen(BLEND_MODE_MENU))
  const modes = BLEND_MODE_MENU.filter((entry) => entry !== '-')
  assert.equal(modes.length, 27)
  assert.equal(new Set(modes).size, 27)
  assert.equal(BLEND_MODE_MENU.filter((entry) => entry === '-').length, 5)
  assert.deepEqual(modes.slice(0, 2), ['normal', 'dissolve'])
  assert.deepEqual(modes.slice(-4), ['hue', 'saturation', 'color', 'luminosity'])
  assert.notEqual(BLEND_MODE_MENU[0], '-')
  assert.notEqual(BLEND_MODE_MENU.at(-1), '-')
})

test('the tool registry seed covers every tool id with an inert controller', () => {
  const { TOOL_FACTORIES, TOOL_GROUPS, TOOL_META } = load('advanced/tools/index.ts')
  const ids = Object.keys(TOOL_META).sort()
  assert.equal(ids.length, 18)
  assert.deepEqual(Object.keys(TOOL_FACTORIES).sort(), ids)
  const grouped = TOOL_GROUPS.flat()
  assert.deepEqual([...grouped].sort(), ids, 'every tool appears in exactly one toolbar group')
  assert.equal(new Set(grouped).size, grouped.length)
  for (const group of TOOL_GROUPS) {
    assert.equal(new Set(group.map((id) => TOOL_META[id].key)).size, 1, `${group.join('/')} share one shortcut key`)
  }
  for (const id of ids) {
    assert.match(TOOL_META[id].key, /^[A-Z]$/)
    assert.ok(TOOL_META[id].label.length > 0)
    const tool = TOOL_FACTORIES[id]()
    assert.equal(tool.id, id)
    assert.equal(tool.hasSession(), false)
    assert.equal(tool.keyDown({ key: 'Escape' }), false)
    assert.equal(tool.keyUp({ key: 'Escape' }), false)
    tool.activate({})
    tool.pointerDown({})
    tool.pointerMove({})
    tool.pointerUp({})
    tool.pointerCancel()
    tool.commitSession()
    tool.cancelSession()
    tool.deactivate()
  }
})
