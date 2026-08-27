import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import ts from 'typescript'

const sourcePath = path.resolve('src/lib/editClipboard.ts')
const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-edit-clipboard-'))
const compiledPath = path.join(temporaryDirectory, 'editClipboard.mjs')

try {
  const source = await fs.readFile(sourcePath, 'utf8')
  const compiled = ts.transpileModule(source, {
    compilerOptions: {
      module: ts.ModuleKind.ES2022,
      target: ts.ScriptTarget.ES2022,
    },
    fileName: sourcePath,
    reportDiagnostics: true,
  })
  const errors = (compiled.diagnostics || []).filter((diagnostic) => diagnostic.category === ts.DiagnosticCategory.Error)
  assert.equal(errors.length, 0, errors.map((diagnostic) => diagnostic.messageText).join('\n'))
  await fs.writeFile(compiledPath, compiled.outputText)

  const clipboard = await import(`${new URL(`file:///${compiledPath.replaceAll('\\', '/')}`)}?t=${Date.now()}`)

  const sourceText = {
    overlayId: 'text-source',
    pageIndex: 0,
    rect: { x: 85, y: 4, width: 30, height: 20 },
    originalRect: { x: 80, y: 5, width: 31, height: 20 },
    originalText: 'old text',
    text: 'copied text',
    fontSize: 12,
    fontFamily: 'Arial',
    fontData: new Uint8Array([1, 2, 3]),
    sourceItemText: 'old text',
    sourceItemRect: { x: 80, y: 5, width: 31, height: 20 },
    sourceSelectionStart: 0,
    sourceSelectionEnd: 8,
    sourceSpaceWidth: 4,
    align: 'left',
    color: [0.1, 0.2, 0.3],
    backgroundColor: [0.9, 0.8, 0.7],
    cover: true,
    modified: true,
    caretOffset: 999,
  }
  const textPayload = clipboard.copyTextEdit(sourceText)
  sourceText.rect.x = 1
  sourceText.color[0] = 1
  sourceText.fontData[0] = 9
  assert.deepEqual(textPayload.edit.rect, { x: 85, y: 4, width: 30, height: 20 })
  assert.deepEqual(textPayload.edit.color, [0.1, 0.2, 0.3])
  assert.deepEqual([...textPayload.edit.fontData], [1, 2, 3])
  assert.equal(textPayload.edit.modified, false)
  assert.equal(textPayload.edit.caretOffset, 'copied text'.length)

  const pastedText = clipboard.pasteTextEditToPage(
    textPayload,
    3,
    { x: 0, y: 0, width: 100, height: 100 },
  )
  assert.deepEqual(pastedText.rect, { x: 70, y: 0, width: 30, height: 20 })
  assert.equal(pastedText.pageIndex, 3)
  assert.equal(pastedText.overlayId, undefined)
  assert.equal(pastedText.originalRect, undefined)
  assert.equal(pastedText.originalText, '')
  assert.equal(pastedText.sourceItemText, undefined)
  assert.equal(pastedText.sourceItemRect, undefined)
  assert.equal(pastedText.sourceSelectionStart, undefined)
  assert.equal(pastedText.sourceSelectionEnd, undefined)
  assert.equal(pastedText.sourceSpaceWidth, undefined)
  assert.equal(pastedText.cover, false)
  assert.equal(pastedText.modified, true)
  assert.equal(pastedText.caretOffset, pastedText.text.length)

  const sourceObject = {
    overlayId: 'object-source',
    candidateId: 'native-image-source',
    pageIndex: 0,
    kind: 'image',
    rect: { x: 10, y: 50, width: 20, height: 20 },
    originalRect: { x: 8, y: 52, width: 20, height: 20 },
    dataUrl: 'data:image/png;base64,AA==',
    opacity: 2,
    cover: true,
    label: 'Image',
    modified: true,
  }
  const objectPayload = clipboard.copyObjectEdit(sourceObject)
  sourceObject.rect.x = 70
  assert.deepEqual(objectPayload.edit.rect, { x: 10, y: 50, width: 20, height: 20 })
  assert.equal(objectPayload.edit.opacity, 1)
  assert.equal(objectPayload.edit.modified, false)

  const pasted = clipboard.pasteEditToPage(
    objectPayload,
    4,
    { x: 0, y: 0, width: 100, height: 80 },
  )
  assert.equal(pasted.kind, 'object')
  assert.deepEqual(pasted.edit.rect, { x: 22, y: 38, width: 20, height: 20 })
  assert.equal(pasted.edit.pageIndex, 4)
  assert.equal(pasted.edit.overlayId, undefined)
  assert.equal(pasted.edit.candidateId, undefined)
  assert.equal(pasted.edit.originalRect, undefined)
  assert.equal(pasted.edit.cover, false)
  assert.equal(pasted.edit.modified, true)

  assert.deepEqual(
    clipboard.pasteRectWithinPage(
      { x: -40, y: 200, width: 150, height: 120 },
      { x: 10, y: 20, width: 90, height: 70 },
    ),
    { x: 10, y: 20, width: 90, height: 70 },
  )

  const clonedPayload = clipboard.cloneEditClipboardPayload(textPayload)
  clonedPayload.edit.rect.x = 0
  clonedPayload.edit.fontData[0] = 8
  assert.equal(textPayload.edit.rect.x, 85)
  assert.equal(textPayload.edit.fontData[0], 1)
  assert.equal(clipboard.copyEditSelection(null, null), null)

  console.log('editClipboard tests passed')
} finally {
  await fs.rm(temporaryDirectory, { recursive: true, force: true })
}
