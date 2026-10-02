// OCR engine smoke (WP1): recognise the clean300 fixture in Electron, offline,
// from file:// under the app's Content Security Policy, with the exact dist/ocr
// files produced by `npm run build:web`; then again with everything packed in an
// app.asar. Also checks: a missing asset fails fast, a broken language file or
// a worker that never starts fails instead of hanging, Stop cancels a running
// page, two pages run in parallel, print-intent rendering (minimised window)
// reads the same text, and a sideways scan is recognised through the
// orientation fallback.
//
//   npm run build:web && node scripts/smoke-ocr-engine.mjs [--no-asar] [--keep]
//
// Electron comes from node_modules/electron/dist or ELECTRON_PATH. The page
// under test is a small harness built with this project's vite.config.ts,
// because no app screen calls the OCR runtime yet.
import assert from 'node:assert/strict'
import { spawn, spawnSync } from 'node:child_process'
import { createHash } from 'node:crypto'
import { existsSync } from 'node:fs'
import fs from 'node:fs/promises'
import { createRequire } from 'node:module'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { gzipSync } from 'node:zlib'

const require = createRequire(import.meta.url)
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const fixture = require('../tests/fixtures/scan-fixture.cjs')
const args = new Set(process.argv.slice(2))
const RUN_TIMEOUT_MS = Number(process.env.SIMPLE_OCR_SMOKE_TIMEOUT || 150_000)
const FIRST_PAGE_BUDGET_MS = 5_000
const MISSING_ASSET_BUDGET_MS = 2_000
const MIN_ACCURACY = 0.98
const MAX_CENTRE_ERROR_PT = 1.5

function electronExecutable() {
  const candidates = [process.env.ELECTRON_PATH, path.join(root, 'node_modules', 'electron', 'dist', 'electron.exe')].filter(Boolean)
  const found = candidates.find((candidate) => existsSync(candidate))
  if (!found) throw new Error('Electron is not installed (node_modules/electron/dist/electron.exe); set ELECTRON_PATH to an Electron 43 executable.')
  return found
}

async function listFiles(directory, prefix = '') {
  const entries = await fs.readdir(path.join(directory, prefix), { withFileTypes: true })
  const files = []
  for (const entry of entries) {
    const relative = prefix ? `${prefix}/${entry.name}` : entry.name
    if (entry.isDirectory()) files.push(...await listFiles(directory, relative))
    else files.push(relative)
  }
  return files
}

const sha256 = (bytes) => createHash('sha256').update(bytes).digest('hex')

/** dist/ocr must hold exactly the manifest's files, each matching its sha256. */
async function verifyOcrAssets(directory) {
  const manifestPath = path.join(directory, 'manifest.json')
  assert.ok(existsSync(manifestPath), `${manifestPath} is missing; run npm run build:web first`)
  const manifest = JSON.parse(await fs.readFile(manifestPath, 'utf8'))
  assert.equal(manifest.schema, 1)
  assert.equal(manifest.core, 'tesseract-core-simd-lstm.wasm.js')
  const files = (await listFiles(directory)).filter((file) => file !== 'manifest.json').sort()
  assert.deepEqual(files, Object.keys(manifest.files).sort(), 'dist/ocr holds exactly the files listed in the manifest')
  for (const required of ['worker.min.js', 'core/tesseract-core-simd-lstm.wasm.js', 'tessdata/eng.traineddata.gz']) assert.ok(files.includes(required), `missing ${required}`)
  assert.ok(files.some((file) => file.startsWith('licenses/')), 'licence notices are missing')
  for (const file of files) {
    assert.ok(/^(worker\.min\.js|core\/[^/]+|tessdata\/[^/]+|licenses\/[^/]+)$/.test(file), `unexpected file ocr/${file}`)
    const bytes = await fs.readFile(path.join(directory, file))
    assert.equal(bytes.length, manifest.files[file].bytes, `${file} size`)
    assert.equal(sha256(bytes), manifest.files[file].sha256, `${file} sha256`)
  }
  for (const language of manifest.languages) assert.equal(manifest.files[language.file]?.sha256, language.sha256, `${language.code} language entry`)
  return { manifest, files }
}

const HARNESS = (imports) => `import { loadPdf } from '${imports.pdf}'
import { preflightOcrAssets } from '${imports.assets}'
import { disposeAllOcrEngines, getOcrEngine } from '${imports.engine}'
import { disposeOcrPrepWorker, preparePageInWorker, recognizePage } from '${imports.recognizePage}'
import { renderForOcr } from '${imports.renderForOcr}'

const violations: string[] = []
document.addEventListener('securitypolicyviolation', (event) => violations.push(event.violatedDirective + ' ' + event.blockedURI))
const send = (value: unknown) => console.log('OCR_SMOKE_RESULT ' + JSON.stringify(value))

async function settle<T>(promise: Promise<T>, limitMs: number) {
  const started = performance.now()
  try {
    const value = await Promise.race([promise, new Promise<never>((_, reject) => setTimeout(() => reject(new Error('no answer within ' + limitMs + ' ms')), limitMs))])
    return { ok: true, ms: Math.round(performance.now() - started), value }
  } catch (error) {
    const failure = error as { name?: string; code?: string; message?: string; detail?: string }
    return { ok: false, ms: Math.round(performance.now() - started), name: failure?.name, code: failure?.code, message: String(failure?.message ?? error), detail: failure?.detail }
  }
}

async function main() {
  const started = performance.now()
  const manifest = await preflightOcrAssets()
  const bytes = new Uint8Array(await (await fetch('./fixture.pdf')).arrayBuffer())
  const pdf = await loadPdf(bytes)
  const diagnostics: Record<string, unknown> = {}
  const phases: string[] = []
  const result = await recognizePage(await pdf.getPage(1), {
    pageIndex: 0,
    nativeDpi: 300,
    diagnostics,
    onProgress: (progress) => { if (phases[phases.length - 1] !== progress.phase) phases.push(progress.phase) },
  })
  const firstPageMs = Math.round(performance.now() - started)

  const warmStarted = performance.now()
  const warmDiagnostics: Record<string, unknown> = {}
  const warm = await recognizePage(await pdf.getPage(1), { pageIndex: 0, nativeDpi: 300, orientationFallback: true, diagnostics: warmDiagnostics })
  const warmMs = Math.round(performance.now() - warmStarted)
  const fallbackSkipped = warmDiagnostics.orientation === undefined

  const sidewaysPdf = await loadPdf(new Uint8Array(await (await fetch('./sideways.pdf')).arrayBuffer()))
  const sidewaysDiagnostics: Record<string, unknown> = {}
  const sidewaysStarted = performance.now()
  const sidewaysResult = await recognizePage(await sidewaysPdf.getPage(1), { pageIndex: 0, nativeDpi: 300, orientationFallback: true, diagnostics: sidewaysDiagnostics })
  const sideways = {
    ms: Math.round(performance.now() - sidewaysStarted),
    orientationCorrectedBy: sidewaysResult.orientationCorrectedBy,
    rotation: sidewaysResult.rotation,
    orientation: sidewaysDiagnostics.orientation,
    words: sidewaysResult.paragraphs.flatMap((paragraph) => paragraph.lines.flatMap((line) => line.words.map((word) => ({
      text: word.text,
      cx: (word.quad[0].x + word.quad[2].x) / 2,
      cy: (word.quad[0].y + word.quad[2].y) / 2,
    })))),
  }
  await sidewaysPdf.destroy()
  // A minimised window renders with print intent (no animation frames needed).
  const printDiagnostics: Record<string, unknown> = {}
  const printed = await recognizePage(await pdf.getPage(1), { pageIndex: 0, nativeDpi: 300, intent: 'print', diagnostics: printDiagnostics })
  const textOf = (page: typeof result) => page.paragraphs.flatMap((paragraph) => paragraph.lines.map((line) => line.text)).join('\\n')
  const printIntent = { intent: printDiagnostics.intent, sameText: textOf(printed) === textOf(result), wordCount: printed.wordCount }

  const render = await renderForOcr(await pdf.getPage(1), { nativeDpi: 300, cleanup: false })
  const prepared = await preparePageInWorker({ rgba: render.rgba, width: render.width, height: render.height, dpi: render.dpi })
  const pgm = prepared.pgm as Uint8Array
  const engine = getOcrEngine({ language: 'eng' })
  const controller = new AbortController()
  const running = engine.recognize({ pgm, dpi: render.dpi, signal: controller.signal })
  setTimeout(() => controller.abort(), 250)
  const cancel = await settle(running, 10_000)
  const afterCancel = await settle(engine.recognize({ pgm, dpi: render.dpi }).then((page) => (page.blocks ?? []).length), 30_000)
  const parallel = await settle(Promise.all([
    engine.recognize({ pgm, dpi: render.dpi }),
    engine.recognize({ pgm, dpi: render.dpi }),
  ]).then((pages) => pages.map((page) => Math.round(page.confidence))), 60_000)

  const missingPreflight = await settle(preflightOcrAssets({ base: './ocr-missing/' }), 10_000)
  const missingEngine = await settle(getOcrEngine({ language: 'eng', assetBase: './ocr-missing/' }).recognize({ pgm, dpi: render.dpi }), 10_000)
  const corrupt = await settle(getOcrEngine({ language: 'eng', assetBase: './ocr-corrupt/' }).recognize({ pgm, dpi: render.dpi }), 45_000)
  const hang = await settle(getOcrEngine({ language: 'eng', assetBase: './ocr-hang/', initTimeoutMs: 1_500 }).recognize({ pgm, dpi: render.dpi }), 15_000)

  await disposeAllOcrEngines()
  disposeOcrPrepWorker()
  await pdf.destroy()
  const words = result.paragraphs.flatMap((paragraph) => paragraph.lines.flatMap((line) => line.words.map((word) => ({
    text: word.text,
    cx: (word.quad[0].x + word.quad[2].x) / 2,
    cy: (word.quad[0].y + word.quad[2].y) / 2,
  }))))
  send({
    ok: true,
    location: location.href,
    visibility: document.visibilityState,
    hardwareConcurrency: navigator.hardwareConcurrency,
    manifest: { tesseractJs: manifest.tesseractJs, core: manifest.core, languages: manifest.languages.map((language) => language.code) },
    firstPageMs,
    warmMs,
    phases,
    diagnostics,
    meanConfidence: result.meanConfidence,
    wordCount: result.wordCount,
    contentKeyStable: result.contentKey === warm.contentKey,
    printIntent,
    fallbackSkipped,
    sideways,
    words,
    cancel,
    afterCancel,
    parallel,
    missingPreflight,
    missingEngine,
    corrupt,
    hang,
    violations,
  })
}

main().catch((error) => send({ ok: false, name: error?.name, code: error?.code, message: String(error?.message ?? error), detail: error?.detail, stack: String(error?.stack ?? ''), violations }))
`

const ELECTRON_MAIN = `const { app, BrowserWindow } = require('electron')
const path = require('node:path')
if (process.env.SIMPLE_OCR_SMOKE_USER_DATA) app.setPath('userData', process.env.SIMPLE_OCR_SMOKE_USER_DATA)
const emit = (record) => process.stdout.write(JSON.stringify(record) + '\\n')
app.whenReady().then(() => {
  // The app's renderer security settings; hidden, and not throttled while hidden.
  const window = new BrowserWindow({
    show: false,
    width: 1200,
    height: 900,
    webPreferences: { contextIsolation: true, nodeIntegration: false, sandbox: true, backgroundThrottling: false },
  })
  window.webContents.on('console-message', (...details) => {
    const event = details[0]
    const message = typeof event?.message === 'string' ? event.message : String(details[2] ?? '')
    emit({ level: event?.level ?? details[1], message })
    if (message.startsWith('OCR_SMOKE_RESULT ')) setTimeout(() => app.quit(), 50)
  })
  window.webContents.on('render-process-gone', (_event, reason) => { emit({ level: 'fatal', message: 'render-process-gone ' + JSON.stringify(reason) }); app.quit() })
  window.loadFile(path.join(__dirname, 'dist', 'index.html'))
})
app.on('window-all-closed', () => app.quit())
`

function alignWords(truth, recognised) {
  const columns = recognised.length + 1
  const cost = new Uint32Array((truth.length + 1) * columns)
  for (let i = 0; i <= truth.length; i += 1) cost[i * columns] = i
  for (let j = 0; j < columns; j += 1) cost[j] = j
  for (let i = 1; i <= truth.length; i += 1) {
    for (let j = 1; j < columns; j += 1) {
      cost[i * columns + j] = Math.min(cost[(i - 1) * columns + j] + 1, cost[i * columns + j - 1] + 1, cost[(i - 1) * columns + j - 1] + (truth[i - 1] === recognised[j - 1] ? 0 : 1))
    }
  }
  const matches = []
  let i = truth.length
  let j = recognised.length
  while (i > 0 && j > 0) {
    if (cost[i * columns + j] === cost[(i - 1) * columns + j - 1] + (truth[i - 1] === recognised[j - 1] ? 0 : 1)) {
      if (truth[i - 1] === recognised[j - 1]) matches.push([i - 1, j - 1])
      i -= 1
      j -= 1
    } else if (cost[i * columns + j] === cost[(i - 1) * columns + j] + 1) i -= 1
    else j -= 1
  }
  return { errors: cost[truth.length * columns + recognised.length], matches }
}

async function prepareApp(appDirectory, distOcr, fixturePdf, sidewaysPdf) {
  const harnessSource = path.join(appDirectory, '..', 'harness-src')
  await fs.mkdir(harnessSource, { recursive: true })
  const appIndex = await fs.readFile(path.join(root, 'index.html'), 'utf8')
  const csp = /<meta http-equiv="Content-Security-Policy"[^>]*>/.exec(appIndex)?.[0]
  assert.ok(csp, 'index.html has no Content-Security-Policy meta tag')
  await fs.writeFile(path.join(harnessSource, 'index.html'), `<!doctype html>\n<html lang="en">\n  <head>\n    <meta charset="UTF-8" />\n    ${csp}\n    <title>ocr smoke</title>\n  </head>\n  <body>\n    <script type="module" src="./harness.ts"></script>\n  </body>\n</html>\n`)
  const relative = (target) => {
    const value = path.relative(harnessSource, path.join(root, target)).split(path.sep).join('/')
    return value.startsWith('.') ? value : `./${value}`
  }
  await fs.writeFile(path.join(harnessSource, 'harness.ts'), HARNESS({
    pdf: relative('src/lib/pdf.ts'),
    assets: relative('src/lib/ocr/assets.ts'),
    engine: relative('src/lib/ocr/engine.ts'),
    recognizePage: relative('src/lib/ocr/recognizePage.ts'),
    renderForOcr: relative('src/lib/ocr/renderForOcr.ts'),
  }))
  const { build } = await import('vite')
  const dist = path.join(appDirectory, 'dist')
  await build({
    configFile: path.join(root, 'vite.config.ts'),
    root: harnessSource,
    base: './',
    logLevel: 'error',
    build: { outDir: dist, emptyOutDir: true, reportCompressedSize: false },
  })
  // Recognise with the exact files `npm run build:web` produced.
  await fs.rm(path.join(dist, 'ocr'), { recursive: true, force: true })
  await fs.cp(distOcr, path.join(dist, 'ocr'), { recursive: true })
  await fs.writeFile(path.join(dist, 'fixture.pdf'), fixturePdf)
  await fs.writeFile(path.join(dist, 'sideways.pdf'), sidewaysPdf)
  // Broken builds: no language file; a damaged language file (valid gzip, listed
  // in its manifest) so only Tesseract's start-up fails; a worker that never
  // answers, so only the start-up timeout can end the wait.
  const manifest = JSON.parse(await fs.readFile(path.join(distOcr, 'manifest.json'), 'utf8'))
  const brokenBuilds = [
    ['ocr-missing', {}],
    ['ocr-corrupt', { language: gzipSync(Buffer.from('not a traineddata file '.repeat(64))) }],
    ['ocr-hang', { worker: Buffer.from('/* never answers */ self.onmessage = () => {}\n') }],
  ]
  for (const [name, replace] of brokenBuilds) {
    const target = path.join(dist, name)
    const copy = structuredClone(manifest)
    const put = async (file, bytes) => {
      await fs.mkdir(path.dirname(path.join(target, file)), { recursive: true })
      await fs.writeFile(path.join(target, file), bytes)
      copy.files[file] = { bytes: bytes.length, sha256: sha256(bytes) }
    }
    await put('worker.min.js', replace.worker ?? await fs.readFile(path.join(distOcr, 'worker.min.js')))
    await put(`core/${manifest.core}`, await fs.readFile(path.join(distOcr, 'core', manifest.core)))
    if (name !== 'ocr-missing') {
      const language = replace.language ?? await fs.readFile(path.join(distOcr, 'tessdata', 'eng.traineddata.gz'))
      await put('tessdata/eng.traineddata.gz', language)
      copy.languages[0] = { ...copy.languages[0], bytes: language.length, sha256: sha256(language) }
    }
    await fs.writeFile(path.join(target, 'manifest.json'), JSON.stringify(copy))
  }
  await fs.writeFile(path.join(appDirectory, 'electron-main.cjs'), ELECTRON_MAIN)
  await fs.writeFile(path.join(appDirectory, 'package.json'), JSON.stringify({ name: 'simple-ocr-smoke', version: '1.0.0', main: 'electron-main.cjs' }))
}

function runElectron(executable, appPath, userData) {
  return new Promise((resolve, reject) => {
    const child = spawn(executable, [appPath, `--user-data-dir=${userData}`], {
      cwd: path.dirname(appPath),
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
      env: { ...process.env, SIMPLE_OCR_SMOKE_USER_DATA: userData, ELECTRON_ENABLE_LOGGING: '0' },
    })
    const consoleMessages = []
    let result = null
    let buffer = ''
    let stderr = ''
    const kill = () => { if (child.exitCode === null) spawnSync('taskkill.exe', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true }) }
    const timer = setTimeout(() => {
      kill()
      reject(new Error(`Electron did not report within ${RUN_TIMEOUT_MS / 1000} s\n${consoleMessages.slice(-20).map((item) => item.message).join('\n')}\n${stderr.slice(-2000)}`))
    }, RUN_TIMEOUT_MS)
    child.stdout.on('data', (chunk) => {
      buffer += chunk
      let newline
      while ((newline = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, newline).trim()
        buffer = buffer.slice(newline + 1)
        if (!line.startsWith('{')) continue
        try {
          const record = JSON.parse(line)
          if (typeof record.message === 'string' && record.message.startsWith('OCR_SMOKE_RESULT ')) result = JSON.parse(record.message.slice('OCR_SMOKE_RESULT '.length))
          else consoleMessages.push(record)
        } catch { /* not ours */ }
      }
    })
    child.stderr.on('data', (chunk) => { stderr += chunk })
    child.once('error', (error) => { clearTimeout(timer); reject(error) })
    child.once('exit', () => {
      clearTimeout(timer)
      if (result) resolve({ result, consoleMessages })
      else reject(new Error(`Electron exited without a result\n${consoleMessages.slice(-20).map((item) => item.message).join('\n')}\n${stderr.slice(-2000)}`))
    })
    // Never leave a hidden Electron behind if this script is interrupted.
    process.once('exit', kill)
  })
}

function compareWords(truth, words) {
  const alignment = alignWords(truth.words.map((word) => word.text), words.map((word) => word.text))
  let maxCentre = 0
  for (const [truthIndex, ocrIndex] of alignment.matches) {
    const expected = truth.words[truthIndex].centre
    const actual = words[ocrIndex]
    maxCentre = Math.max(maxCentre, Math.hypot(actual.cx - expected.x, actual.cy - expected.y))
  }
  return { errors: alignment.errors, accuracy: 1 - alignment.errors / truth.words.length, maxCentre }
}

function evaluate(label, { result, consoleMessages }, truth, sidewaysTruth) {
  const failures = []
  if (!result.ok) return { label, failures: [`renderer error: ${result.code ?? result.name}: ${result.message} ${result.detail ?? ''}\n${result.stack ?? ''}`] }
  const { errors, accuracy, maxCentre } = compareWords(truth, result.words)
  const alignment = { errors }
  const sideways = compareWords(sidewaysTruth, result.sideways.words)
  if (result.sideways.orientationCorrectedBy !== 90) failures.push(`sideways page: corrected by ${result.sideways.orientationCorrectedBy}, expected 90 (${JSON.stringify(result.sideways.orientation)})`)
  if (sideways.accuracy < MIN_ACCURACY) failures.push(`sideways page: word accuracy ${(sideways.accuracy * 100).toFixed(1)}%`)
  if (sideways.maxCentre > MAX_CENTRE_ERROR_PT) failures.push(`sideways page: word centre off by ${sideways.maxCentre.toFixed(2)} pt`)
  if (!result.fallbackSkipped) failures.push('the orientation fallback probed a page that was read well')
  const cspMessages = consoleMessages.filter((item) => /Content Security Policy|Refused to (load|execute|create|connect)/i.test(item.message))
  if (accuracy < MIN_ACCURACY) failures.push(`word accuracy ${(accuracy * 100).toFixed(1)}% < ${MIN_ACCURACY * 100}%`)
  if (maxCentre > MAX_CENTRE_ERROR_PT) failures.push(`word centre off by ${maxCentre.toFixed(2)} pt`)
  if (result.firstPageMs > FIRST_PAGE_BUDGET_MS) failures.push(`first page took ${result.firstPageMs} ms (budget ${FIRST_PAGE_BUDGET_MS} ms)`)
  if (result.violations.length || cspMessages.length) failures.push(`CSP violations: ${JSON.stringify([...result.violations, ...cspMessages.map((item) => item.message)])}`)
  if (!result.contentKeyStable) failures.push('the same page produced different content keys')
  if (result.printIntent.intent !== 'print' || !result.printIntent.sameText) failures.push(`print-intent render: ${JSON.stringify(result.printIntent)}`)
  if (result.cancel.ok || result.cancel.name !== 'AbortError' || result.cancel.ms > 2_000) failures.push(`cancel: ${JSON.stringify(result.cancel)}`)
  if (!result.afterCancel.ok || !(result.afterCancel.value > 0)) failures.push(`recognition after cancel: ${JSON.stringify(result.afterCancel)}`)
  if (!result.parallel.ok) failures.push(`parallel pages: ${JSON.stringify(result.parallel)}`)
  for (const [name, outcome] of [['missing asset (preflight)', result.missingPreflight], ['missing asset (engine)', result.missingEngine]]) {
    if (outcome.ok || outcome.code !== 'assets-missing' || outcome.ms > MISSING_ASSET_BUDGET_MS) failures.push(`${name}: ${JSON.stringify(outcome)}`)
  }
  if (result.corrupt.ok || result.corrupt.code !== 'engine-init') failures.push(`damaged language file: ${JSON.stringify(result.corrupt)}`)
  if (result.hang.ok || result.hang.code !== 'engine-timeout' || result.hang.ms > 5_000) failures.push(`worker that never starts: ${JSON.stringify(result.hang)}`)
  if (consoleMessages.some((item) => /could not be tracked/.test(item.message))) failures.push('the engine could not track its Web Worker (a failed start would leak it)')
  const summary = {
    label,
    location: result.location,
    words: `${truth.words.length - alignment.errors}/${truth.words.length}`,
    accuracy: Math.round(accuracy * 10000) / 100,
    maxCentrePt: Math.round(maxCentre * 100) / 100,
    meanConfidence: Math.round(result.meanConfidence),
    firstPageMs: result.firstPageMs,
    warmMs: result.warmMs,
    diagnostics: result.diagnostics,
    phases: result.phases,
    visibility: result.visibility,
    cancelMs: result.cancel.ms,
    afterCancelMs: result.afterCancel.ms,
    parallelMs: result.parallel.ms,
    missingPreflightMs: result.missingPreflight.ms,
    missingEngineMs: result.missingEngine.ms,
    corrupt: { code: result.corrupt.code, ms: result.corrupt.ms },
    hang: { code: result.hang.code, ms: result.hang.ms },
    printIntent: result.printIntent,
    sideways: {
      correctedBy: result.sideways.orientationCorrectedBy,
      words: `${result.sideways.words.length}`,
      accuracy: Math.round(sideways.accuracy * 10000) / 100,
      maxCentrePt: Math.round(sideways.maxCentre * 100) / 100,
      ms: result.sideways.ms,
      probes: result.sideways.orientation?.probes,
    },
    cspViolations: result.violations.length + cspMessages.length,
    consoleErrors: consoleMessages.filter((item) => item.level === 'error' || item.level === 3).map((item) => item.message).slice(0, 5),
  }
  return { label, failures, summary }
}

const executable = electronExecutable()
const workRoot = process.env.SIMPLE_OCR_SMOKE_DIR ? path.resolve(process.env.SIMPLE_OCR_SMOKE_DIR) : os.tmpdir()
await fs.mkdir(workRoot, { recursive: true })
const temporaryDirectory = await fs.mkdtemp(path.join(workRoot, 'simple-ocr-smoke-'))
const reports = []
try {
  const distOcr = path.join(root, 'dist', 'ocr')
  const { manifest, files } = await verifyOcrAssets(distOcr)
  console.log(`dist/ocr verified: ${files.length + 1} files, sha256 values match manifest (tesseract.js ${manifest.tesseractJs}, ${manifest.core})`)
  const variant = await fixture.buildScanVariant('clean300')
  const sidewaysVariant = await fixture.buildScanVariant('sideways300')
  const appDirectory = path.join(temporaryDirectory, 'app')
  await prepareApp(appDirectory, distOcr, variant.pdf, sidewaysVariant.pdf)

  const runs = [['file://', appDirectory]]
  if (!args.has('--no-asar')) {
    const asarPath = path.join(temporaryDirectory, 'app.asar')
    await require('@electron/asar').createPackage(appDirectory, asarPath)
    runs.push(['file:// inside app.asar', asarPath])
  }
  for (const [label, appPath] of runs) {
    const userData = path.join(temporaryDirectory, `user-data-${reports.length}`)
    let outcome
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      outcome = evaluate(label, await runElectron(executable, appPath, userData), variant.truth, sidewaysVariant.truth)
      // A first page over budget under heavy machine load gets one retry.
      const timingOnly = outcome.failures.length > 0 && outcome.failures.every((failure) => failure.startsWith('first page took'))
      if (!timingOnly || attempt === 2) break
      console.log(`${label}: ${outcome.failures[0]}; retrying once`)
    }
    reports.push(outcome)
    console.log(`${label}: ${outcome.failures.length ? 'FAIL' : 'ok'} ${JSON.stringify(outcome.summary ?? {})}`)
    for (const failure of outcome.failures) console.log(`  - ${failure}`)
  }
} finally {
  if (!args.has('--keep')) await fs.rm(temporaryDirectory, { recursive: true, force: true, maxRetries: 8, retryDelay: 250 })
  else console.log(`kept ${temporaryDirectory}`)
}
const failures = reports.flatMap((report) => report.failures.map((failure) => `${report.label}: ${failure}`))
console.log(JSON.stringify({ ok: failures.length === 0, failures, runs: reports.map((report) => report.summary) }))
if (failures.length) process.exitCode = 1
