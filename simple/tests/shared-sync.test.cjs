'use strict'

// Drift and wiring guard for the shared Save/Import/Export layer. Every
// vendoring case runs against a temporary copy: a temp shared root (the real
// manifest with stub sources) and temp workspace folders seeded with copies of
// the real preload files. Nothing is written into the real workspaces.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const { spawnSync } = require('node:child_process')
const { MODES } = require('../electron/routing.cjs')
const { SyncSetupError, findBridgeBlocks, loadManifest, syncShared } = require('../scripts/sync-shared.cjs')

const ROOT = path.resolve(__dirname, '..')
const WORKSPACE = path.resolve(ROOT, '..')
const SHARED_ROOT = path.join(ROOT, 'shared')
const SCRIPT = path.join(ROOT, 'scripts', 'sync-shared.cjs')
const realManifest = loadManifest(SHARED_ROOT)

function stubSource(manifest, relative) {
  if (relative.endsWith('.json')) return '{ "stub": true }\n'
  const header = manifest.header.replace('{source}', relative)
  return `${header}\r\n'use strict'\r\n// stub for ${relative}\r\n`
}

/**
 * Builds a disposable tree: <tmp>/shared (manifest + every listed source as a
 * stub) and <tmp>/repo/<workspace folders> holding a copy of each real
 * preload and a stub main.cjs that performs the manifest's wiring calls.
 */
function makeTree(t, { enabled = () => true, omit = [] } = {}) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-shared-sync-'))
  t.after(() => fs.rmSync(tmp, { recursive: true, force: true }))
  const sharedRoot = path.join(tmp, 'shared')
  const targetRoot = path.join(tmp, 'repo')
  const manifest = JSON.parse(JSON.stringify(realManifest))
  const sources = new Set([manifest.preloadBlock.source])
  for (const [name, workspace] of Object.entries(manifest.workspaces)) {
    workspace.enabled = enabled(name)
    for (const relative of [...workspace.electron, ...workspace.renderer]) sources.add(relative)
    const electronRoot = path.join(targetRoot, workspace.folder, 'electron')
    fs.mkdirSync(electronRoot, { recursive: true })
    fs.copyFileSync(path.join(WORKSPACE, workspace.folder, 'electron', 'preload.cjs'), path.join(electronRoot, 'preload.cjs'))
    const calls = (workspace.wiring || []).map((rule) => `${rule.call}{})`).join('\r\n')
    fs.writeFileSync(path.join(electronRoot, 'main.cjs'), `'use strict'\r\n${calls}\r\n`)
  }
  fs.mkdirSync(sharedRoot, { recursive: true })
  fs.writeFileSync(path.join(sharedRoot, 'manifest.json'), `${JSON.stringify(manifest, null, 2)}\n`)
  for (const relative of sources) {
    if (omit.includes(relative)) continue
    fs.mkdirSync(path.dirname(path.join(sharedRoot, relative)), { recursive: true })
    fs.writeFileSync(path.join(sharedRoot, relative), stubSource(manifest, relative))
  }
  return { tmp, sharedRoot, targetRoot, manifest, options: { sharedRoot, targetRoot } }
}

function vendoredPath(tree, name, kind, relative) {
  const workspace = tree.manifest.workspaces[name]
  return path.join(tree.targetRoot, workspace.folder, tree.manifest.destinations[kind], path.posix.basename(relative))
}

function preloadPath(tree, name) {
  return path.join(tree.targetRoot, tree.manifest.workspaces[name].folder, tree.manifest.preloadBlock.file)
}

function runCli(args) {
  return spawnSync(process.execPath, [SCRIPT, ...args], { encoding: 'utf8', windowsHide: true, timeout: 60_000 })
}

test('the real tree is current and the manifest follows the shared layout rules', () => {
  const result = syncShared({ check: true })
  assert.ok(result.ok, result.message)
  const cli = runCli(['--check'])
  assert.equal(cli.status, 0, cli.stderr)

  assert.deepEqual(Object.keys(realManifest.workspaces).sort(), [...MODES].sort())
  const appManifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'modules', 'manifest.json'), 'utf8'))
  for (const [name, workspace] of Object.entries(realManifest.workspaces)) {
    assert.equal(workspace.folder, appManifest.modules[name].source, `${name} points to the wrong workspace folder.`)
  }
  // Workspaces own unrelated src/shared and electron/shared folders, so vendored copies live in simple-io.
  assert.equal(realManifest.destinations.electron, 'electron/simple-io')
  assert.equal(realManifest.destinations.renderer, 'src/simple-io')
  assert.deepEqual(realManifest.workspaces.video.renderer, [])
  for (const name of ['image', 'video']) {
    assert.ok(!realManifest.workspaces[name].electron.includes('electron/office-engine.cjs'), `${name} must not receive the office engine.`)
  }

  const packageJson = JSON.parse(fs.readFileSync(path.join(ROOT, 'package.json'), 'utf8'))
  assert.equal(packageJson.scripts['sync:shared'], 'node scripts/sync-shared.cjs')
  assert.ok(packageJson.build.files.includes('shared/**/*'), 'The launcher requires simple/shared at runtime, so it must be packaged.')
})

test('every enabled real workspace is wired to the shared layer', () => {
  // Wiring is part of syncShared's check, which only looks at enabled workspaces.
  const enabled = Object.entries(realManifest.workspaces).filter(([, workspace]) => workspace.enabled).map(([name]) => name)
  const result = syncShared({ check: true })
  assert.deepEqual(result.workspaces, enabled)
  assert.deepEqual(result.problems, [])
})

test('write mode vendors byte-identical copies and one bridge block, then is idempotent', (t) => {
  const tree = makeTree(t)
  const originals = Object.fromEntries(Object.keys(tree.manifest.workspaces).map((name) => [name, fs.readFileSync(preloadPath(tree, name), 'utf8')]))
  assert.equal(syncShared({ ...tree.options, check: true }).ok, false)

  const first = syncShared(tree.options)
  assert.ok(first.ok, first.message)
  assert.ok(first.changes.length > 0)
  assert.deepEqual(first.planned, [])
  for (const [name, workspace] of Object.entries(tree.manifest.workspaces)) {
    for (const kind of ['electron', 'renderer']) {
      for (const relative of workspace[kind]) {
        assert.deepEqual(fs.readFileSync(vendoredPath(tree, name, kind, relative)), fs.readFileSync(path.join(tree.sharedRoot, relative)), `${name} ${relative}`)
      }
    }
    const text = fs.readFileSync(preloadPath(tree, name), 'utf8')
    const { blocks, malformed } = findBridgeBlocks(text)
    assert.equal(malformed, null)
    assert.equal(blocks.length, 1, `${name} must hold exactly one bridge block.`)
    // Integrated workspaces already hold a block (replaced in place); others get one appended.
    const outside = (source) => {
      const found = findBridgeBlocks(source).blocks
      return (found.length ? source.slice(0, found[0].start) + source.slice(found[found.length - 1].end) : source).trimEnd()
    }
    assert.equal(outside(text), outside(originals[name]), `${name} preload code outside the block must be kept.`)
    const block = text.slice(blocks[0].start, blocks[0].end)
    assert.ok(block.startsWith(`${tree.manifest.preloadBlock.begin}\r\n${tree.manifest.header.replace('{source}', tree.manifest.preloadBlock.source)}\r\n`))
    assert.ok(!/[^\r]\n/.test(block), 'The block must use the preload file\'s CRLF line endings.')
  }
  assert.ok(!fs.existsSync(vendoredPath(tree, 'video', 'electron', 'office-engine.cjs')))
  assert.ok(!fs.existsSync(path.join(tree.targetRoot, tree.manifest.workspaces.video.folder, 'src')))

  const second = syncShared(tree.options)
  assert.deepEqual(second.changes, [])
  assert.ok(syncShared({ ...tree.options, check: true }).ok)
})

test('a hand-edited vendored file fails the check by name, and write mode restores it', (t) => {
  const tree = makeTree(t)
  syncShared(tree.options)
  const edited = vendoredPath(tree, 'calc', 'electron', 'electron/safe-write.cjs')
  fs.appendFileSync(edited, '// quick local fix\r\n')

  const result = syncShared({ ...tree.options, check: true })
  assert.equal(result.ok, false)
  assert.deepEqual(result.drift.map((item) => `${item.folder}/${item.file}`), ['simple_calc_source/electron/simple-io/safe-write.cjs'])
  assert.match(result.message, /Shared I\/O code is out of date in simple_calc_source \(electron\/simple-io\/safe-write\.cjs: differs from simple\/shared\/electron\/safe-write\.cjs\)\. Run: .*sync-shared\.cjs/)

  const cli = runCli(['--check', '--shared-root', tree.sharedRoot, '--target-root', tree.targetRoot])
  assert.equal(cli.status, 1)
  assert.match(cli.stderr, /simple_calc_source \(electron\/simple-io\/safe-write\.cjs/)
  assert.match(cli.stderr, /Run: /)

  const repaired = runCli([`--target-root=${tree.targetRoot}`, `--shared-root=${tree.sharedRoot}`, '--module=calc'])
  assert.equal(repaired.status, 0, repaired.stderr)
  assert.match(repaired.stdout, /Updated simple_calc_source\/electron\/simple-io\/safe-write\.cjs/)
  assert.ok(syncShared({ ...tree.options, check: true }).ok)
})

test('a duplicated preload block fails the check, and write mode keeps exactly one', (t) => {
  const tree = makeTree(t)
  syncShared(tree.options)
  const preload = preloadPath(tree, 'docs')
  const text = fs.readFileSync(preload, 'utf8')
  const { blocks } = findBridgeBlocks(text)
  const block = text.slice(blocks[0].start, blocks[0].end)
  const original = text.slice(0, blocks[0].start)
  fs.writeFileSync(preload, `${text}\r\n${block}`)

  const result = syncShared({ ...tree.options, check: true })
  assert.equal(result.ok, false)
  assert.match(result.message, /simple_doc_source \(electron\/preload\.cjs: bridge block appears 2 times/)

  syncShared(tree.options)
  const repaired = fs.readFileSync(preload, 'utf8')
  assert.equal(findBridgeBlocks(repaired).blocks.length, 1)
  assert.equal(repaired, text, 'Removing the duplicate must restore the single-block file exactly.')
  assert.ok(repaired.startsWith(original))
})

test('an edited bridge block is detected and replaced in place', (t) => {
  const tree = makeTree(t)
  syncShared(tree.options)
  const preload = preloadPath(tree, 'image')
  const good = fs.readFileSync(preload, 'utf8')
  fs.writeFileSync(preload, good.replace("'use strict'\r\n// stub", "'use strict'\r\nconsole.log('patched')\r\n// stub"))
  const result = syncShared({ ...tree.options, check: true })
  assert.match(result.message, /simple_image_source \(electron\/preload\.cjs: bridge block differs/)
  syncShared(tree.options)
  assert.equal(fs.readFileSync(preload, 'utf8'), good)
})

test('unbalanced bridge markers are reported and never rewritten automatically', (t) => {
  const tree = makeTree(t)
  const preload = preloadPath(tree, 'pdf')
  const broken = `${fs.readFileSync(preload, 'utf8')}\r\n${tree.manifest.preloadBlock.begin}\r\nconst half = true\r\n`
  fs.writeFileSync(preload, broken)
  const result = syncShared(tree.options)
  assert.equal(result.ok, false)
  assert.match(result.message, /simple_pdf_source\/electron\/preload\.cjs has a begin marker at line \d+ without an end marker/)
  assert.equal(fs.readFileSync(preload, 'utf8'), broken)
})

test('stale vendored files are reported and removed', (t) => {
  const tree = makeTree(t)
  syncShared(tree.options)
  // A copy vendored from a shared source that was renamed since: it carries the header.
  const stray = path.join(path.dirname(vendoredPath(tree, 'pdf', 'renderer', 'renderer/io-client.ts')), 'old-helper.ts')
  fs.writeFileSync(stray, `${tree.manifest.header.replace('{source}', 'renderer/old-helper.ts')}\r\n// left behind\r\n`)
  const result = syncShared({ ...tree.options, check: true })
  assert.match(result.message, /simple_pdf_source \(src\/simple-io\/old-helper\.ts: not listed/)
  syncShared(tree.options)
  assert.ok(!fs.existsSync(stray))

  // A source that is deleted from simple/shared takes its vendored copies with it.
  fs.rmSync(path.join(tree.sharedRoot, 'electron/stores.cjs'))
  assert.equal(syncShared({ ...tree.options, check: true }).ok, false)
  syncShared(tree.options)
  assert.ok(!fs.existsSync(vendoredPath(tree, 'pdf', 'electron', 'electron/stores.cjs')))
  assert.ok(syncShared({ ...tree.options, check: true }).ok)
})

test('a file or folder in a destination that the sync did not vendor is reported and never removed', (t) => {
  const tree = makeTree(t)
  syncShared(tree.options)
  const folder = path.dirname(vendoredPath(tree, 'pdf', 'electron', 'electron/io-core.cjs'))
  const handWritten = path.join(folder, 'my-notes.cjs')
  fs.writeFileSync(handWritten, "'use strict'\n// someone's work\n")
  const data = path.join(folder, 'settings.json')
  fs.writeFileSync(data, '{ "mine": true }\n')
  const subfolder = path.join(folder, 'drafts')
  fs.mkdirSync(subfolder)
  fs.writeFileSync(path.join(subfolder, 'draft.cjs'), '// draft\n')
  for (const check of [true, false]) {
    const result = syncShared({ ...tree.options, check })
    assert.equal(result.ok, false)
    for (const name of ['my-notes.cjs', 'settings.json', 'drafts']) {
      assert.match(result.message, new RegExp(`simple_pdf_source/electron/simple-io/${name.replace('.', '\\.')} is not a copy vendored from simple/shared, so it was left in place`))
    }
    assert.ok(fs.existsSync(handWritten) && fs.existsSync(data) && fs.existsSync(path.join(subfolder, 'draft.cjs')), 'nothing that may be someone\'s work is removed')
  }
  fs.rmSync(handWritten); fs.rmSync(data); fs.rmSync(subfolder, { recursive: true })
  assert.ok(syncShared(tree.options).ok)

  // A data file (no header) named like one the manifest still vendors is a
  // vendored copy: dropped from one workspace, its copy there is removed.
  const manifestPath = path.join(tree.sharedRoot, 'manifest.json')
  const catalog = vendoredPath(tree, 'pdf', 'electron', 'electron/io-catalog.json')
  const dropped = JSON.parse(JSON.stringify(tree.manifest))
  dropped.workspaces.pdf.electron = dropped.workspaces.pdf.electron.filter((source) => source !== 'electron/io-catalog.json')
  fs.writeFileSync(manifestPath, JSON.stringify(dropped, null, 2))
  assert.ok(syncShared(tree.options).ok)
  assert.ok(!fs.existsSync(catalog))
  assert.ok(fs.existsSync(vendoredPath(tree, 'calc', 'electron', 'electron/io-catalog.json')))
  // Dropped everywhere, nothing proves the file was vendored: it is reported and kept.
  fs.writeFileSync(manifestPath, JSON.stringify(tree.manifest, null, 2))
  syncShared(tree.options)
  for (const workspace of Object.values(dropped.workspaces)) workspace.electron = workspace.electron.filter((source) => source !== 'electron/io-catalog.json')
  dropped.sameBytes = []
  fs.writeFileSync(manifestPath, JSON.stringify(dropped, null, 2))
  const kept = syncShared(tree.options)
  assert.equal(kept.ok, false)
  assert.match(kept.message, /simple_pdf_source\/electron\/simple-io\/io-catalog\.json is not a copy vendored from simple\/shared, so it was left in place/)
  assert.ok(fs.existsSync(catalog))
})

test('a destination that is not a simple-io folder is refused before anything is written or removed', (t) => {
  for (const [kind, destination] of [['electron', 'electron'], ['renderer', 'src'], ['electron', 'src/simple-io']]) {
    const tree = makeTree(t)
    const manifestPath = path.join(tree.sharedRoot, 'manifest.json')
    const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
    manifest.destinations[kind] = destination
    fs.writeFileSync(manifestPath, JSON.stringify(manifest, null, 2))
    const electronRoot = path.join(tree.targetRoot, 'simple_pdf_source', 'electron')
    const before = fs.readdirSync(electronRoot).sort()
    assert.throws(() => syncShared(tree.options), /must be a folder named simple-io|must be different folders/)
    assert.equal(runCli(['--shared-root', tree.sharedRoot, '--target-root', tree.targetRoot]).status, 2)
    assert.deepEqual(fs.readdirSync(electronRoot).sort(), before, `${kind} = ${destination}: the workspace's own files are untouched`)
    assert.ok(before.includes('main.cjs') && before.includes('preload.cjs'))
  }
})

test('npm run watch never vendors a changed manifest, and runs Node without a shell', async (t) => {
  const source = fs.readFileSync(path.join(ROOT, 'scripts', 'watch-and-build.cjs'), 'utf8')
  assert.match(source, /if \(manifestHash\(\) !== manifestAtStart\) \{[\s\S]*?\[SYNC_SHARED, '--check'\][\s\S]*?return\r?\n/)
  assert.doesNotMatch(source, /shell: (?!false)|'npx(\.cmd)?'/)
  const { electronBuilderCli, run } = require('../scripts/watch-and-build.cjs')
  assert.ok(fs.existsSync(electronBuilderCli()))
  // node.exe under C:\Program Files must start, path spaces and all.
  await run(process.execPath, ['-e', 'process.exit(0)'], { stdio: 'ignore' })
  await assert.rejects(run(process.execPath, ['-e', 'process.exit(3)'], { stdio: 'ignore' }), /exited with 3/)
  const spaced = fs.mkdtempSync(path.join(os.tmpdir(), 'simple watch run '))
  t.after(() => fs.rmSync(spaced, { recursive: true, force: true }))
  const script = path.join(spaced, 'print args.cjs')
  const out = path.join(spaced, 'out.json')
  fs.writeFileSync(script, `require('node:fs').writeFileSync(${JSON.stringify(out)}, JSON.stringify(process.argv.slice(2)))`)
  await run(process.execPath, [script, 'two words', 'C:\\Program Files\\x'], { stdio: 'ignore' })
  assert.deepEqual(JSON.parse(fs.readFileSync(out, 'utf8')), ['two words', 'C:\\Program Files\\x'])
})

test('planned shared files that do not exist yet are tolerated', (t) => {
  const tree = makeTree(t, { omit: ['electron/document-guard.cjs', 'preload/io-bridge.cjs'] })
  // Start from a preload without a bridge block (integrated workspaces already carry one).
  const withBlock = fs.readFileSync(preloadPath(tree, 'calc'), 'utf8')
  const existing = findBridgeBlocks(withBlock).blocks
  const original = existing.length ? (withBlock.slice(0, existing[0].start) + withBlock.slice(existing[existing.length - 1].end)).trimEnd() + '\r\n' : withBlock
  fs.writeFileSync(preloadPath(tree, 'calc'), original)
  const result = syncShared(tree.options)
  assert.ok(result.ok, result.message)
  assert.deepEqual(result.planned, ['electron/document-guard.cjs', 'preload/io-bridge.cjs'])
  assert.ok(!fs.existsSync(vendoredPath(tree, 'calc', 'electron', 'electron/document-guard.cjs')))
  assert.equal(fs.readFileSync(preloadPath(tree, 'calc'), 'utf8'), original, 'No markers are inserted before the bridge source exists.')
  assert.ok(syncShared({ ...tree.options, check: true }).ok)
})

test('disabled workspaces are neither written nor checked', (t) => {
  const tree = makeTree(t, { enabled: (name) => name !== 'pdf' })
  const pdfPreload = fs.readFileSync(preloadPath(tree, 'pdf'))
  const result = syncShared(tree.options)
  assert.ok(result.ok, result.message)
  assert.deepEqual(result.skipped, ['pdf'])
  assert.ok(!fs.existsSync(vendoredPath(tree, 'pdf', 'electron', 'electron/io-core.cjs')))
  assert.deepEqual(fs.readFileSync(preloadPath(tree, 'pdf')), pdfPreload)

  // Even a hand-made copy in a disabled workspace is ignored until it is enabled.
  fs.mkdirSync(path.dirname(vendoredPath(tree, 'pdf', 'electron', 'electron/io-core.cjs')), { recursive: true })
  fs.writeFileSync(vendoredPath(tree, 'pdf', 'electron', 'electron/io-core.cjs'), 'edited')
  fs.writeFileSync(path.join(tree.targetRoot, tree.manifest.workspaces.pdf.folder, 'electron', 'main.cjs'), '// not wired\n')
  assert.ok(syncShared({ ...tree.options, check: true }).ok)
  assert.throws(() => syncShared({ ...tree.options, modules: ['pdf'] }), SyncSetupError)
  assert.throws(() => syncShared({ ...tree.options, modules: ['nope'] }), /Unknown module/)
})

test('wiring is required only for enabled workspaces once the providing source exists', (t) => {
  const tree = makeTree(t)
  const main = path.join(tree.targetRoot, tree.manifest.workspaces.image.folder, 'electron', 'main.cjs')
  fs.writeFileSync(main, "'use strict'\r\nregisterSharedIo({})\r\n")
  const result = syncShared(tree.options)
  assert.equal(result.ok, false)
  assert.deepEqual(result.problems, [
    'simple_image_source/electron/main.cjs must call installWindowGuard( because the workspace is enabled in simple/shared/manifest.json and simple/shared/electron/document-guard.cjs exists.',
  ])

  fs.rmSync(path.join(tree.sharedRoot, 'electron/document-guard.cjs'))
  assert.ok(syncShared(tree.options).ok)
})

test('shared sources must carry the vendored header and paired catalogs must match', (t) => {
  const tree = makeTree(t, { enabled: () => false })
  fs.writeFileSync(path.join(tree.sharedRoot, 'electron/io-core.cjs'), "'use strict'\n")
  fs.writeFileSync(path.join(tree.sharedRoot, 'renderer/io-catalog.json'), '{ "different": true }\n')
  const result = syncShared({ ...tree.options, check: true })
  assert.equal(result.ok, false)
  assert.match(result.message, /simple\/shared\/electron\/io-core\.cjs must begin with this line:\n  \/\/ Vendored from simple\/shared\/electron\/io-core\.cjs by simple\/scripts\/sync-shared\.cjs\. Do not edit here\./)
  assert.match(result.message, /simple\/shared\/renderer\/io-catalog\.json must be byte-identical to simple\/shared\/electron\/io-catalog\.json/)
})

test('a malformed manifest or option is a setup error with exit code 2', (t) => {
  const tree = makeTree(t)
  const manifestPath = path.join(tree.sharedRoot, 'manifest.json')
  const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'))
  manifest.workspaces.pdf.electron.push('../escape.cjs')
  fs.writeFileSync(manifestPath, JSON.stringify(manifest))
  assert.throws(() => syncShared(tree.options), /must be a path under electron\//)
  assert.equal(runCli(['--shared-root', tree.sharedRoot, '--target-root', tree.targetRoot]).status, 2)
  assert.equal(runCli(['--unknown']).status, 2)
})
