import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { pathToFileURL } from 'node:url'
import ts from 'typescript'

const temporaryDirectory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-document-search-'))

try {
  for (const name of ['search', 'documentSearch']) {
    const sourcePath = path.resolve(`src/lib/${name}.ts`)
    const source = await fs.readFile(sourcePath, 'utf8')
    const compiled = ts.transpileModule(source, {
      compilerOptions: { module: ts.ModuleKind.ES2022, target: ts.ScriptTarget.ES2022 },
      fileName: sourcePath,
      reportDiagnostics: true,
    })
    const errors = (compiled.diagnostics || []).filter((item) => item.category === ts.DiagnosticCategory.Error)
    assert.equal(errors.length, 0, errors.map((item) => item.messageText).join('\n'))
    await fs.writeFile(path.join(temporaryDirectory, `${name}.mjs`), compiled.outputText.replace("from './search'", "from './search.mjs'"))
  }
  const { createDocumentSearch } = await import(pathToFileURL(path.join(temporaryDirectory, 'documentSearch.mjs')).href)

  const textPages = [
    '  Alpha\n beta ALPHA beta.  ',
    'The oﬃce “opens” — soon. office “opens” — soon.',
    'Banana banana',
    '',
    'No matching text',
  ]
  let reads = 0
  const search = createDocumentSearch(async (_pdf, index) => { reads += 1; return textPages[index] })
  const document = { numPages: textPages.length }
  assert.deepEqual(await search(document, 'ALPHA beta'), [
    { pageIndex: 0, excerpt: 'Alpha beta ALPHA beta.', count: 2 },
  ])
  assert.deepEqual(await search(document, 'office "opens" - soon.'), [
    { pageIndex: 1, excerpt: textPages[1], count: 2 },
  ])
  assert.deepEqual(await search(document, 'ana'), [
    { pageIndex: 2, excerpt: 'Banana banana', count: 2 },
  ])
  assert.deepEqual(await search(document, 'absent'), [])
  assert.deepEqual(await search(document, '  '), [])
  assert.equal(reads, 5, 'Changed queries should reuse the completed text index')
  await search({ numPages: textPages.length }, 'alpha')
  assert.equal(reads, 10, 'Another PDF identity must never reuse the first document index')

  const largeDocument = { numPages: 256 }
  let largeReads = 0
  let yields = 0
  let cancelled = false
  let cancelAtYield = false
  const largeSearch = createDocumentSearch(async (_pdf, index) => {
    largeReads += 1
    return `Page ${index + 1}: needle first. needle second.`
  }, {
    now: () => 0,
    yieldToUi: async () => { yields += 1; if (cancelAtYield) cancelled = true },
  })
  const firstMatches = await largeSearch(largeDocument, 'needle')
  assert.equal(firstMatches.length, 256)
  assert.equal(firstMatches.reduce((sum, item) => sum + item.count, 0), 512)
  assert.deepEqual(firstMatches.map((item) => item.pageIndex), Array.from({ length: 256 }, (_, index) => index))
  const coldReads = largeReads
  const firstYields = yields
  const secondMatches = await largeSearch(largeDocument, 'first')
  assert.equal(secondMatches.length, 256)
  assert.ok(secondMatches.every((item) => item.count === 1))
  assert.equal(largeReads, coldReads, 'A second query should perform no page reads or page normalization')
  assert.equal(yields - firstYields, 7, 'A fully cached scan still yields between bounded page batches')
  cancelAtYield = true
  const yieldsBeforeCancellation = yields
  assert.equal(await largeSearch(largeDocument, 'second', { isCancelled: () => cancelled }), null)
  assert.equal(yields - yieldsBeforeCancellation, 1, 'A cancelled cached scan stops at its first UI yield')
  assert.equal(largeReads, coldReads)

  let releasePage
  let pendingReads = 0
  let pendingCancelled = false
  const pendingDocument = { numPages: 100 }
  const pendingSearch = createDocumentSearch(async () => {
    pendingReads += 1
    if (pendingReads === 1) return new Promise((resolve) => { releasePage = resolve })
    return 'needle'
  })
  const pendingResult = pendingSearch(pendingDocument, 'needle', { isCancelled: () => pendingCancelled })
  assert.equal(pendingReads, 1)
  pendingCancelled = true
  releasePage('needle')
  assert.equal(await pendingResult, null)
  assert.equal(pendingReads, 1, 'Cancellation while extracting a page must skip the remaining 99 pages')
  assert.equal((await pendingSearch(pendingDocument, 'needle')).length, 100, 'Cancellation must not poison a later query')
  assert.equal(pendingReads, 101)
  const readsBeforeCancelledStart = pendingReads
  assert.equal(await pendingSearch(pendingDocument, 'needle', { isCancelled: () => true }), null)
  assert.equal(pendingReads, readsBeforeCancelledStart)

  let timerCancelled = false
  let timerReads = 0
  const timerSearch = createDocumentSearch(async () => { timerReads += 1; return 'needle' })
  const timer = setTimeout(() => { timerCancelled = true }, 0)
  assert.equal(await timerSearch({ numPages: 5000 }, 'needle', { isCancelled: () => timerCancelled }), null)
  clearTimeout(timer)
  assert.ok(timerReads <= 32, `A UI timer should cancel before the next batch, read ${timerReads} pages`)

  const retryReads = [0, 0, 0]
  const retrySearch = createDocumentSearch(async (_pdf, index) => {
    retryReads[index] += 1
    if (index === 1 && retryReads[index] === 1) throw new Error('Transient extraction failure')
    return 'needle'
  })
  const retryDocument = { numPages: 3 }
  await assert.rejects(retrySearch(retryDocument, 'needle'), /Transient extraction failure/)
  assert.equal((await retrySearch(retryDocument, 'needle')).length, 3)
  assert.deepEqual(retryReads, [1, 2, 1], 'A failed page must retry while earlier indexed pages remain reusable')

  console.log(JSON.stringify({
    passed: true,
    coldPageReads: coldReads,
    repeatQueryPageReads: largeReads - coldReads,
    cachedScanUiYields: 7,
    cancelledInFlightPageReads: 1,
    cancelledInFlightPagesSkipped: 99,
    timerCancellationPageReads: timerReads,
    checks: ['normalization', 'occurrence order/counts', 'PDF identity', 'cache reuse', 'cached scan yielding', 'in-flight cancellation', 'cancelled retry', 'real timer responsiveness', 'failed extraction retry'],
  }))
} finally {
  const resolvedTemporaryDirectory = path.resolve(temporaryDirectory)
  const resolvedTemporaryRoot = path.resolve(os.tmpdir())
  assert.ok(resolvedTemporaryDirectory.startsWith(`${resolvedTemporaryRoot}${path.sep}`))
  assert.ok(path.basename(resolvedTemporaryDirectory).startsWith('simple-document-search-'))
  await fs.rm(resolvedTemporaryDirectory, { recursive: true, force: true })
}
