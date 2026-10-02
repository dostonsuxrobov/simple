const test = require('node:test')
const assert = require('node:assert/strict')

test('document export is unavailable without an open document or while busy', async () => {
  const { documentExportDisabled } = await import('../src/ui-guards.js')
  assert.equal(documentExportDisabled(false, false), true)
  assert.equal(documentExportDisabled(true, true), true)
  assert.equal(documentExportDisabled(false, true), true)
  assert.equal(documentExportDisabled(true, false), false)
})

test('active modal lookup covers every visible modal backdrop', async () => {
  const { activeModal } = await import('../src/ui-guards.js')
  const expected = { id: 'future-modal' }
  const root = {
    querySelector(selector) {
      assert.equal(selector, '.modal-backdrop:not([hidden])')
      return expected
    },
  }
  assert.equal(activeModal(root), expected)
})

test('modal guard blocks all app file shortcuts but leaves unrelated keys alone', async () => {
  const { isModalBlockedShortcut } = await import('../src/ui-guards.js')
  for (const key of ['s', 'S', 'o', 'n', 'p', 'w', 'f']) {
    assert.equal(isModalBlockedShortcut({ key, ctrlKey: true }), true, `Ctrl+${key}`)
  }
  assert.equal(isModalBlockedShortcut({ key: 's', metaKey: true, shiftKey: true }), true)
  assert.equal(isModalBlockedShortcut({ key: 'E', ctrlKey: true, shiftKey: true }), true)
  assert.equal(isModalBlockedShortcut({ key: 'e', metaKey: true, shiftKey: true }), true)
  assert.equal(isModalBlockedShortcut({ key: 'F12' }), true, 'F12 (Save as)')
  assert.equal(isModalBlockedShortcut({ key: 'e', ctrlKey: true }), false)
  assert.equal(isModalBlockedShortcut({ key: 'p' }), false)
  // Text editing inside a dialog field keeps working.
  for (const key of ['a', 'c', 'v', 'x', 'z', 'y']) assert.equal(isModalBlockedShortcut({ key, ctrlKey: true }), false, `Ctrl+${key}`)
})

test('modal guard follows the physical key on non-Latin layouts and never blocks AltGr typing', async () => {
  const { isModalBlockedShortcut } = await import('../src/ui-guards.js')
  // Russian layout: Ctrl+ы is Ctrl+S, Ctrl+з is Ctrl+P, Ctrl+Shift+у is Ctrl+Shift+E.
  assert.equal(isModalBlockedShortcut({ key: 'ы', code: 'KeyS', ctrlKey: true }), true)
  assert.equal(isModalBlockedShortcut({ key: 'з', code: 'KeyP', ctrlKey: true }), true)
  assert.equal(isModalBlockedShortcut({ key: 'У', code: 'KeyE', ctrlKey: true, shiftKey: true }), true)
  assert.equal(isModalBlockedShortcut({ key: 'ф', code: 'KeyA', ctrlKey: true }), false)
  // Polish AltGr (Ctrl+Alt): ś and ó are typed characters, not Save or Open.
  assert.equal(isModalBlockedShortcut({ key: 'ś', code: 'KeyS', ctrlKey: true, altKey: true }), false)
  assert.equal(isModalBlockedShortcut({ key: 'ó', code: 'KeyO', ctrlKey: true, altKey: true }), false)
  // AZERTY keeps the typed letter: the key labelled A types "a" where QWERTY has Q.
  assert.equal(isModalBlockedShortcut({ key: 'a', code: 'KeyQ', ctrlKey: true }), false)
})

test('scheduling a debounced preview immediately invalidates an older in-flight result', async () => {
  const { canCommitPrintPreview, nextPrintPreviewGeneration } = await import('../src/ui-guards.js')
  let generation = nextPrintPreviewGeneration(0)
  const inFlightGeneration = generation
  let releaseStale
  const staleFinished = new Promise((resolve) => { releaseStale = resolve }).then(() => (
    canCommitPrintPreview(inFlightGeneration, generation, false)
  ))

  generation = nextPrintPreviewGeneration(generation)
  const scheduledGeneration = generation
  releaseStale()

  assert.equal(await staleFinished, false, 'the old result cannot commit during the debounce window')
  assert.equal(canCommitPrintPreview(scheduledGeneration, generation, false), true)
  assert.equal(canCommitPrintPreview(scheduledGeneration, generation, true), false)
})
