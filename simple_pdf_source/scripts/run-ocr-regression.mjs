// Recognize text (OCR, design WP3) end to end in the built app:
//   main   scanned page: inspector state, export check, offer, recognise with
//          progress, edit at the clicked line, search highlight, copy of a
//          line, undo/redo, a cached re-run, save (the saved file is read back)
//   stop   six scanned pages: Stop keeps the finished pages only
//   mixed  typed text over a scan: recognising adds no duplicate words
//
//   npm run build:web && node scripts/run-ocr-regression.mjs [--keep] [--only=main,stop,mixed]
//
// SIMPLE_OCR_UI_DIR sets where fixtures and the Electron profile go (default:
// the system temp folder). Each scenario launches its own app instance with a
// private profile and remote debugging; every process started is killed.
import assert from 'node:assert/strict'
import { spawn, spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import fs from 'node:fs/promises'
import { createRequire } from 'node:module'
import net from 'node:net'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const require = createRequire(import.meta.url)
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const fixture = require('../tests/fixtures/scan-fixture.cjs')
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib')
const args = process.argv.slice(2)
const keep = args.includes('--keep')
const only = new Set((args.find((arg) => arg.startsWith('--only='))?.slice(7) || 'main,stop,mixed').split(',').filter(Boolean))
const SCENARIO_TIMEOUT_MS = Number(process.env.SIMPLE_OCR_UI_TIMEOUT || 240_000)
const MIN_SAVED_ACCURACY = 0.98

function electronExecutable() {
  const packaged = process.env.SIMPLE_TEST_EXECUTABLE ? path.resolve(process.env.SIMPLE_TEST_EXECUTABLE) : ''
  if (packaged) return { executable: packaged, appArgs: [] }
  const candidates = [process.env.ELECTRON_PATH, path.join(root, 'node_modules', 'electron', 'dist', 'electron.exe')].filter(Boolean)
  const found = candidates.find((candidate) => existsSync(candidate))
  if (!found) throw new Error('Electron is not installed (node_modules/electron/dist/electron.exe).')
  return { executable: found, appArgs: ['.'] }
}

function freePort() {
  return new Promise((resolve, reject) => {
    const server = net.createServer()
    server.once('error', reject)
    server.listen(0, '127.0.0.1', () => {
      const { port } = server.address()
      server.close(() => resolve(port))
    })
  })
}

function run(command, commandArgs, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, commandArgs, { cwd: root, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'], ...options })
    let stdout = ''
    let stderr = ''
    const timer = setTimeout(() => {
      spawnSync('taskkill.exe', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
      reject(new Error(`${path.basename(commandArgs[0] || command)} did not finish within ${SCENARIO_TIMEOUT_MS / 1000} s\n${stdout}${stderr}`))
    }, SCENARIO_TIMEOUT_MS)
    child.stdout.on('data', (chunk) => { stdout += chunk })
    child.stderr.on('data', (chunk) => { stderr += chunk })
    child.once('error', (error) => { clearTimeout(timer); reject(error) })
    child.once('exit', (code) => {
      clearTimeout(timer)
      if (code === 0) resolve(stdout.trim())
      else reject(new Error(`${path.basename(commandArgs[0] || command)} exited ${code}\n${stdout}${stderr}`))
    })
  })
}

/** Ground truth of a variant, with each word's ink box size (PDF points). */
async function truthOf(variant) {
  const layout = await fixture.groundTruth()
  const boxes = layout.lines.flatMap((line) => line.words.map((word) => word.inkBox))
  return {
    lines: variant.truth.lines.map((line) => line.words.map((word) => word.text).join(' ')),
    words: variant.truth.words.map((word, index) => ({
      text: word.text,
      line: word.lineIndex,
      centre: word.centre,
      width: boxes[index].width,
      height: boxes[index].height,
    })),
  }
}

/** One PDF holding each variant as a page. */
async function combine(variants) {
  const target = await PDFDocument.create()
  for (const variant of variants) {
    const source = await PDFDocument.load(variant.pdf)
    const [page] = await target.copyPages(source, [0])
    target.addPage(page)
  }
  return target.save({ useObjectStreams: false })
}

/** A scan with its title and first paragraph also typed as visible text over the scanned words. */
async function mixedPage(variant) {
  const layout = await fixture.groundTruth()
  const doc = await PDFDocument.load(variant.pdf)
  const page = doc.getPage(0)
  const fonts = new Map()
  const typed = layout.lines.slice(0, 4)
  for (const line of typed) {
    if (!fonts.has(line.font)) fonts.set(line.font, await doc.embedFont(StandardFonts[line.font]))
    page.drawText(line.text, { x: line.x, y: line.baseline, size: line.size, font: fonts.get(line.font), color: rgb(0.08, 0.08, 0.1) })
  }
  return { pdf: await doc.save({ useObjectStreams: false }), typedLines: typed.map((line) => line.text) }
}

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
  return cost[truth.length * columns + recognised.length]
}

async function savedWordAccuracy(file, truthWords) {
  const data = new Uint8Array(await fs.readFile(file))
  const pdf = await pdfjs.getDocument({ data, isEvalSupported: false, disableFontFace: true }).promise
  try {
    const page = await pdf.getPage(1)
    const content = await page.getTextContent()
    const words = content.items.map((item) => ('str' in item ? item.str : '')).join(' ').split(/\s+/).filter(Boolean)
    const errors = alignWords(truthWords, words)
    return { accuracy: 1 - errors / truthWords.length, words: words.length, errors }
  } finally {
    await pdf.destroy()
  }
}

async function launch(scenario, file, truthFile, workDirectory) {
  const { executable, appArgs } = electronExecutable()
  const port = await freePort()
  const profile = path.join(workDirectory, `profile-${scenario}`)
  const app = spawn(executable, [...appArgs, `--remote-debugging-port=${port}`, `--user-data-dir=${profile}`, file], {
    cwd: root,
    windowsHide: true,
    stdio: 'ignore',
  })
  const stop = () => {
    if (app.exitCode === null) spawnSync('taskkill.exe', ['/pid', String(app.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
  }
  process.once('exit', stop)
  try {
    const output = await run(process.execPath, ['scripts/smoke-ocr.mjs'], {
      env: { ...process.env, SIMPLE_OCR_UI_PORT: String(port), SIMPLE_OCR_UI_SCENARIO: scenario, SIMPLE_OCR_UI_TRUTH: truthFile },
    })
    const line = output.split('\n').reverse().find((item) => item.trim().startsWith('{'))
    return JSON.parse(line)
  } finally {
    app.kill()
    await Promise.race([new Promise((resolve) => app.once('exit', resolve)), new Promise((resolve) => setTimeout(resolve, 2_000))])
    stop()
    process.removeListener('exit', stop)
  }
}

const workRoot = process.env.SIMPLE_OCR_UI_DIR ? path.resolve(process.env.SIMPLE_OCR_UI_DIR) : os.tmpdir()
await fs.mkdir(workRoot, { recursive: true })
const workDirectory = await fs.mkdtemp(path.join(workRoot, 'simple-ocr-ui-'))
const reports = {}
const failures = []
try {
  assert.ok(existsSync(path.join(root, 'dist', 'index.html')), 'dist/index.html is missing; run npm run build:web first')
  assert.ok(existsSync(path.join(root, 'dist', 'ocr', 'manifest.json')), 'dist/ocr is missing; run npm run build:web first')

  const scenarios = []
  if (only.has('main')) {
    const noisy = await fixture.buildScanVariant('noisy300')
    const file = path.join(workDirectory, 'noisy300.pdf')
    await fs.writeFile(file, noisy.pdf)
    const truth = await truthOf(noisy)
    scenarios.push({ name: 'main', file, truth, verify: async (report) => {
      const saved = await savedWordAccuracy(file, truth.words.map((word) => word.text))
      report.saved = saved
      assert.ok(saved.accuracy >= MIN_SAVED_ACCURACY, `saved PDF: word accuracy ${(saved.accuracy * 100).toFixed(1)}% < ${MIN_SAVED_ACCURACY * 100}%`)
    } })
  }
  if (only.has('stop')) {
    const names = ['clean300', 'skew4', 'clean200', 'low150', 'hard300', 'color300']
    const variants = []
    for (const name of names) variants.push(await fixture.buildScanVariant(name))
    const file = path.join(workDirectory, 'six-scans.pdf')
    await fs.writeFile(file, await combine(variants))
    scenarios.push({ name: 'stop', file, truth: { ...(await truthOf(variants[0])), pages: names.length } })
  }
  if (only.has('mixed')) {
    const clean = await fixture.buildScanVariant('clean300')
    const mixed = await mixedPage(clean)
    const file = path.join(workDirectory, 'mixed.pdf')
    await fs.writeFile(file, mixed.pdf)
    scenarios.push({ name: 'mixed', file, truth: { ...(await truthOf(clean)), typedLines: mixed.typedLines } })
  }

  for (const scenario of scenarios) {
    const truthFile = path.join(workDirectory, `${scenario.name}-truth.json`)
    await fs.writeFile(truthFile, JSON.stringify(scenario.truth))
    let report
    for (let attempt = 1; attempt <= 2; attempt += 1) {
      try {
        report = await launch(scenario.name, scenario.file, truthFile, workDirectory)
        await scenario.verify?.(report)
        report.ok = true
        break
      } catch (error) {
        report = { ok: false, error: String(error?.message || error).slice(0, 4000) }
        // One retry: the machine may be busy (timeouts), never for a wrong result twice.
        if (attempt === 2) break
        console.log(`${scenario.name}: attempt ${attempt} failed, retrying once: ${report.error.split('\n')[0]}`)
        if (scenario.name === 'main') {
          // The first attempt saved the fixture; start again from the scan.
          const noisy = await fixture.buildScanVariant('noisy300')
          await fs.writeFile(scenario.file, noisy.pdf)
        }
      }
    }
    reports[scenario.name] = report
    if (!report.ok) failures.push(`${scenario.name}: ${report.error}`)
    console.log(`${scenario.name}: ${report.ok ? 'ok' : 'FAIL'} ${JSON.stringify(report)}`)
  }
} finally {
  if (!keep) await fs.rm(workDirectory, { recursive: true, force: true, maxRetries: 8, retryDelay: 250 })
  else console.log(`kept ${workDirectory}`)
}
console.log(JSON.stringify({ ok: failures.length === 0, failures, reports }))
if (failures.length) process.exitCode = 1
