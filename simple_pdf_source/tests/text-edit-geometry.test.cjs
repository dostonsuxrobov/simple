'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const test = require('node:test')
const ts = require('typescript')
const { PDFDocument, StandardFonts } = require('pdf-lib')
const { loadMain } = require('./helpers/electron-harness.cjs')
const { textItems } = require('./helpers/pdf-test-utils.cjs')

// src/lib/editClipboard.ts is browser-side TypeScript without DOM needs;
// transpile it on the fly like tests/viewer-clipboard.test.cjs does.
const editClipboard = (async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-text-geometry-'))
  const sourcePath = path.resolve(__dirname, '..', 'src', 'lib', 'editClipboard.ts')
  const compiled = ts.transpileModule(await fs.readFile(sourcePath, 'utf8'), {
    compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 },
    fileName: sourcePath,
  })
  const target = path.join(directory, 'editClipboard.mjs')
  await fs.writeFile(target, compiled.outputText)
  process.once('exit', () => { try { require('node:fs').rmSync(directory, { recursive: true, force: true }) } catch { /* best effort */ } })
  return import(pathToFileURL(target).href)
})()

// "Hello world" in 12 pt Helvetica with its baseline at y = 700, and the text
// edit the renderer makes when that run is clicked: the box is the glyph box
// and the baseline sits 3 pt above its bottom.
const BASELINE = 700
const nativeEdit = (extra = {}) => ({
  pageIndex: 0,
  rect: { x: 72, y: 697, width: 64, height: 13 },
  originalRect: { x: 72, y: 697, width: 64, height: 13 },
  originalText: 'Hello world',
  text: 'Hello world',
  fontSize: 12,
  fontFamily: 'Arial',
  textFit: 'wrap',
  baselineOffset: BASELINE - 697,
  displayRotation: 0,
  angle: 0,
  align: 'left',
  color: [0, 0, 0],
  cover: true,
  modified: true,
  ...extra,
})

async function helloWorld() {
  const doc = await PDFDocument.create()
  const page = doc.addPage([500, 760])
  page.drawText('Hello world', { x: 72, y: BASELINE, size: 12, font: await doc.embedFont(StandardFonts.Helvetica) })
  return doc.save()
}

/** What the saver does with a box: the first baseline, from the box top. */
const savedBaselineFromTop = (edit) => edit.rect.height - edit.baselineOffset

test('resizing a native text box keeps its baseline a fixed distance from the top edge', async () => {
  const { resizeTextEditRect } = await editClipboard
  const edit = nativeEdit()
  // Bottom edge dragged 24 pt down: the top stays, the text must stay.
  const taller = resizeTextEditRect(edit, { x: 72, y: 673, width: 64, height: 37 })
  assert.equal(taller.baselineOffset, 27)
  assert.equal(savedBaselineFromTop(taller), savedBaselineFromTop(edit))
  // Top edge dragged 24 pt up (H field +24): the text follows the top edge.
  const higher = resizeTextEditRect(edit, { x: 72, y: 697, width: 64, height: 37 })
  assert.equal(higher.baselineOffset, 27)
  assert.equal(higher.rect.y + higher.rect.height - savedBaselineFromTop(higher), 710 + 24 - 10)
  // A move (X/Y or arrow keys) keeps the offset; the text moves with the box.
  assert.equal(resizeTextEditRect(edit, { x: 90, y: 650, width: 64, height: 13 }).baselineOffset, 3)
  // Narrower or wider does not change the height either.
  assert.equal(resizeTextEditRect(edit, { x: 72, y: 697, width: 200, height: 13 }).baselineOffset, 3)
  // Added text has no native baseline to keep.
  const added = resizeTextEditRect({ ...edit, baselineOffset: undefined, originalRect: undefined }, { x: 0, y: 0, width: 10, height: 50 })
  assert.equal(added.baselineOffset, undefined)
  assert.deepEqual(added.rect, { x: 0, y: 0, width: 10, height: 50 })
})

test('a box shown sideways grows along its reading height', async () => {
  const { resizeTextEditRect } = await editClipboard
  const sideways = nativeEdit({ displayRotation: 90, rect: { x: 72, y: 697, width: 13, height: 64 }, originalRect: { x: 72, y: 697, width: 13, height: 64 } })
  assert.equal(resizeTextEditRect(sideways, { x: 72, y: 697, width: 37, height: 64 }).baselineOffset, 27)
  assert.equal(resizeTextEditRect(sideways, { x: 72, y: 697, width: 13, height: 90 }).baselineOffset, 3)
})

test('a two-line edit in a box enlarged with the bottom handle saves at the shown position without shrinking', async () => {
  const { resizeTextEditRect } = await editClipboard
  const { invoke } = loadMain()
  const resized = resizeTextEditRect(nativeEdit(), { x: 72, y: 673, width: 140, height: 37 })
  const overlay = { id: 'native', type: 'text', ...resized, text: 'Hello world\nSecond line' }
  const result = await invoke('pdf:flatten-overlays', await helloWorld(), [overlay], {}, { report: true })
  assert.equal(result.ok, true, JSON.stringify(result.failures))
  assert.deepEqual(result.warnings.filter((warning) => warning.overlayId === 'native'), [], 'nothing was shrunk or overflowed')
  const [items] = await textItems(result.data)
  const first = items.find((item) => item.str.includes('Hello'))
  const second = items.find((item) => item.str.includes('Second'))
  assert.ok(Math.abs(first.transform[5] - BASELINE) < 0.05, `first baseline ${first.transform[5]}`)
  assert.ok(Math.abs(first.transform[5] - second.transform[5] - 12 * 1.18) < 0.05, 'second line one line height lower')
  assert.ok(Math.abs(first.height - 12) < 0.05, 'kept the 12 pt size')

  // The unadjusted box (the old behaviour) had to shrink the text to fit.
  const stale = { ...overlay, baselineOffset: 3 }
  const before = await invoke('pdf:flatten-overlays', await helloWorld(), [stale], {}, { report: true })
  assert.ok(before.warnings.some((warning) => warning.overlayId === 'native' && /^TEXT_(SHRUNK|OVERFLOW)$/.test(warning.code)))
})
