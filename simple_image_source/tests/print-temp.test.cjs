const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const {
  LEGACY_STALE_AGE_MS,
  PRINT_DIRECTORY_PREFIX,
  cleanupStalePrintDirectories,
  createOwnedPrintDirectory,
  isDirectPrintDirectory,
} = require('../electron/print-temp.cjs')

test('stale cleanup preserves live owners and fresh unmarked directories', async () => {
  const root = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-print-temp-test-'))
  const now = Date.now()
  try {
    const live = await createOwnedPrintDirectory(root, { pid: 101, startedAt: now - 10_000 })
    const abandoned = await createOwnedPrintDirectory(root, { pid: 202, startedAt: now - 20_000 })
    const freshLegacy = await fs.mkdtemp(path.join(root, PRINT_DIRECTORY_PREFIX))
    const staleLegacy = await fs.mkdtemp(path.join(root, PRINT_DIRECTORY_PREFIX))
    const staleTime = new Date(now - LEGACY_STALE_AGE_MS - 1_000)
    await fs.utimes(staleLegacy, staleTime, staleTime)

    const result = await cleanupStalePrintDirectories({
      root,
      now,
      isProcessAlive: (pid) => pid === 101,
    })

    assert.equal(isDirectPrintDirectory(live, root), true)
    assert.equal(isDirectPrintDirectory(root, root), false)
    assert.equal(await fs.stat(live).then(() => true, () => false), true)
    assert.equal(await fs.stat(freshLegacy).then(() => true, () => false), true)
    assert.equal(await fs.stat(abandoned).then(() => true, () => false), false)
    assert.equal(await fs.stat(staleLegacy).then(() => true, () => false), false)
    assert.deepEqual(new Set(result.preserved), new Set([live, freshLegacy]))
    assert.deepEqual(new Set(result.removed), new Set([abandoned, staleLegacy]))
  } finally {
    await fs.rm(root, { recursive: true, force: true })
  }
})

test('cleanup never traverses or removes a similarly named directory outside its root', async () => {
  const parent = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-image-print-boundary-test-'))
  const root = path.join(parent, 'owned-root')
  const outside = path.join(parent, `${PRINT_DIRECTORY_PREFIX}outside`)
  try {
    await fs.mkdir(root)
    await fs.mkdir(outside)
    const result = await cleanupStalePrintDirectories({ root, now: Date.now() + LEGACY_STALE_AGE_MS * 2, isProcessAlive: () => false })
    assert.deepEqual(result, { removed: [], preserved: [] })
    assert.equal(await fs.stat(outside).then(() => true, () => false), true)
  } finally {
    await fs.rm(parent, { recursive: true, force: true })
  }
})
