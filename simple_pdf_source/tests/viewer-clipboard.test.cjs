const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { pathToFileURL } = require('node:url')
const test = require('node:test')
const ts = require('typescript')

// Viewer clipboard helpers are browser-side TypeScript; transpile the two
// modules on the fly (no DOM needed for the pure functions tested here).
const modules = (async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-viewer-clipboard-'))
  const loaded = {}
  for (const name of ['textSelection', 'editClipboard']) {
    const sourcePath = path.resolve(__dirname, '..', 'src', 'lib', `${name}.ts`)
    const compiled = ts.transpileModule(await fs.readFile(sourcePath, 'utf8'), {
      compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 },
      fileName: sourcePath,
      reportDiagnostics: true,
    })
    const errors = (compiled.diagnostics || []).filter((item) => item.category === ts.DiagnosticCategory.Error)
    assert.equal(errors.length, 0, errors.map((item) => item.messageText).join('\n'))
    const target = path.join(directory, `${name}.mjs`)
    await fs.writeFile(target, compiled.outputText)
    loaded[name] = await import(pathToFileURL(target).href)
  }
  return loaded
})()

test('copied text keeps line breaks and word gaps between text-layer runs', async () => {
  const { joinPlainTextRuns } = (await modules).textSelection
  const text = joinPlainTextRuns([
    { text: 'The quick brown fox jumps over', line: 0, read: 10, advance: 200, fontHeight: 12 },
    { text: 'the lazy dog', line: 1, read: 10, advance: 80, fontHeight: 12 },
    { text: 'beside', line: 1, read: 94, advance: 40, fontHeight: 12 },
    { text: 'the river.', line: 1, read: 140, advance: 60, fontHeight: 12 },
  ])
  assert.equal(text, 'The quick brown fox jumps over\nthe lazy dog beside the river.')
})

test('runs that already carry whitespace or touch are not padded', async () => {
  const { joinPlainTextRuns } = (await modules).textSelection
  assert.equal(joinPlainTextRuns([
    { text: 'Hel', line: 0, read: 0, advance: 20, fontHeight: 12 },
    { text: 'lo ', line: 0, read: 20, advance: 18, fontHeight: 12 },
    { text: 'world', line: 0, read: 60, advance: 30, fontHeight: 12 },
  ]), 'Hello world')
  // An end-of-line marker inside one visual line still separates words.
  assert.equal(joinPlainTextRuns([
    { text: 'left', line: 0, read: 0, advance: 20, fontHeight: 12, eol: true },
    { text: 'right', line: 0, read: 20, advance: 20, fontHeight: 12 },
  ]), 'left right')
})

test('getTextContent items become lines at hasEOL markers and baseline jumps', async () => {
  const { plainTextFromTextContentItems } = (await modules).textSelection
  const item = (str, x, y, width, extra = {}) => ({ str, transform: [12, 0, 0, 12, x, y], width, height: 12, ...extra })
  const text = plainTextFromTextContentItems([
    item('Heading of page 1', 50, 700, 120),
    { str: '', hasEOL: true, transform: [12, 0, 0, 12, 170, 700], width: 0, height: 0 },
    item('The quick brown fox jumps over', 50, 670, 200),
    { str: '', hasEOL: true, transform: [12, 0, 0, 12, 250, 670], width: 0, height: 0 },
    item('the lazy dog', 50, 655, 80),
    item('today.', 140, 655, 40),
    item('New paragraph without marker', 50, 620, 180),
  ])
  assert.equal(text, 'Heading of page 1\nThe quick brown fox jumps over\nthe lazy dog today.\nNew paragraph without marker')
})

test('selected page texts join with one newline and skip empty pages', async () => {
  const { joinSelectedPageTexts } = (await modules).textSelection
  assert.equal(joinSelectedPageTexts(['Page one', '', null, 'Page three']), 'Page one\nPage three')
})

test('paste honours newer system clipboard content over the in-app payload', async () => {
  const { chooseEditPasteSource } = (await modules).editClipboard
  const base = { systemReadable: true, systemText: '', systemHasImage: false }
  // Clipboard still holds what Simple copied: paste the styled text box.
  assert.equal(chooseEditPasteSource({ ...base, internalText: 'Total due', systemText: 'Total due\r\n' }), 'internal')
  // The user copied something else afterwards: that wins.
  assert.equal(chooseEditPasteSource({ ...base, internalText: 'Total due', systemText: 'Invoice #4471' }), 'text')
  assert.equal(chooseEditPasteSource({ ...base, internalText: 'logo.png', systemHasImage: true }), 'image')
  // Nothing pasteable arrived later, or the clipboard cannot be read.
  assert.equal(chooseEditPasteSource({ ...base, internalText: 'Total due' }), 'internal')
  assert.equal(chooseEditPasteSource({ ...base, internalText: 'Total due', systemReadable: false, systemText: 'x' }), 'internal')
  assert.equal(chooseEditPasteSource({ ...base, internalText: null }), 'none')
  assert.equal(chooseEditPasteSource({ ...base, internalText: null, systemText: 'hello' }), 'text')
})

test('pasted text boxes fit their text and stay inside the page', async () => {
  const { pastedTextBoxSize, editClipboardPlainText } = (await modules).editClipboard
  const short = pastedTextBoxSize('Invoice #4471', 12, 480, 700)
  assert.ok(short.width > 60 && short.width < 200, JSON.stringify(short))
  assert.equal(short.wraps, false)
  const long = pastedTextBoxSize('word '.repeat(400), 12, 480, 700)
  assert.equal(long.width, 480)
  assert.equal(long.wraps, true)
  assert.ok(long.height <= 700 && long.height > 12 * 5)
  const lines = pastedTextBoxSize('one\ntwo\nthree', 12, 480, 700)
  assert.ok(lines.height >= 3 * 12 * 1.25)
  assert.equal(editClipboardPlainText({ kind: 'text', edit: { text: 'abc' } }), 'abc')
  assert.equal(editClipboardPlainText({ kind: 'object', edit: { label: 'logo.png' } }), 'logo.png')
})
