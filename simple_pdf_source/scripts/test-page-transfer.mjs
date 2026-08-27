import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import ts from 'typescript'

const sourcePath = path.resolve('src/lib/pageTransfer.ts')
const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-page-transfer-'))
const compiledPath = path.join(temporaryDirectory, 'pageTransfer.mjs')

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

  const transfer = await import(`${new URL(`file:///${compiledPath.replaceAll('\\', '/')}`)}?t=${Date.now()}`)

  assert.deepEqual(transfer.pageIndicesForTransfer(new Set([5, 1, 3]), 3), [1, 3, 5])
  assert.deepEqual(transfer.pageIndicesForTransfer(new Set([5, 1, 3]), 2), [2])
  assert.deepEqual(transfer.pageIndicesForTransfer(new Set([-1, 2, 2.5, 4]), 4), [2, 4])

  assert.equal(transfer.pageDropEdge(49, 0, 100), 'before')
  assert.equal(transfer.pageDropEdge(50, 0, 100), 'after')
  assert.equal(transfer.pageDropEdge(225, 200, 50), 'after')
  assert.equal(transfer.pageDropInsertIndex(4, 'before'), 4)
  assert.equal(transfer.pageDropInsertIndex(4, 'after'), 5)
  assert.equal(transfer.pageReorderDestination(0, 1), 0)
  assert.equal(transfer.pageReorderDestination(0, 2), 1)
  assert.equal(transfer.pageReorderDestination(4, 1), 1)
  assert.equal(transfer.pageReorderDestination(4, 4), 4)

  assert.equal(transfer.isPdfTransferFile({ name: 'selection.PDF', type: '' }), true)
  assert.equal(transfer.isPdfTransferFile({ name: 'selection', type: 'application/pdf' }), true)
  assert.equal(transfer.isPdfTransferFile({ name: 'notes.txt', type: 'text/plain' }), false)

  console.log('pageTransfer tests passed')
} finally {
  await fs.rm(temporaryDirectory, { recursive: true, force: true })
}
