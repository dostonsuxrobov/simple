// Launches the built app (dist/) in isolated Electron profiles and runs the
// viewer/navigation/search/selection smokes against generated fixtures:
//   1. scripts/smoke-search-selection.mjs on the interaction fixture
//   2. scripts/smoke-pointer-selection.mjs on the interaction fixture
//   3. scripts/smoke-long-selection.mjs and scripts/smoke-viewer-interaction.mjs
//      on a generated 40-page reading fixture
// Run `npm run build:web` first. Every launched process tree is killed.
import { spawn, spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { copyFile, mkdtemp, rm, writeFile } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib'

const root = path.resolve('.')
// This workspace's electron package may lack its downloaded binary; the
// unified app ships the same Electron version, so fall back to it (read-only).
const localElectron = path.join(root, 'node_modules', 'electron', 'dist', 'electron.exe')
const electron = process.env.SIMPLE_ELECTRON
  || (existsSync(localElectron) ? localElectron : path.resolve(root, '..', 'simple', 'node_modules', 'electron', 'dist', 'electron.exe'))
const packagedExecutable = process.env.SIMPLE_TEST_EXECUTABLE ? path.resolve(process.env.SIMPLE_TEST_EXECUTABLE) : ''
const executable = packagedExecutable || electron
const scratch = process.env.SIMPLE_TEST_SCRATCH ? path.resolve(process.env.SIMPLE_TEST_SCRATCH) : os.tmpdir()
// SIMPLE_VIEWER_SMOKES picks a subset: search, pointer, long, viewer.
const only = new Set((process.env.SIMPLE_VIEWER_SMOKES || 'search,pointer,long,viewer').split(',').map((item) => item.trim()))

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: root, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'], ...options })
    let stdout = ''
    let stderr = ''
    const timer = setTimeout(() => {
      spawnSync('taskkill.exe', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
    }, options.timeout || 180_000)
    child.stdout?.on('data', (chunk) => { stdout += chunk })
    child.stderr?.on('data', (chunk) => { stderr += chunk })
    child.once('error', reject)
    child.once('exit', (code) => {
      clearTimeout(timer)
      if (code === 0) resolve(stdout.trim())
      else reject(new Error(`${path.basename(args[0] || command)} exited ${code}\n${stdout}${stderr}`.trim()))
    })
  })
}

async function createReadingFixture(target) {
  const pdf = await PDFDocument.create()
  const regular = await pdf.embedFont(StandardFonts.Helvetica)
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold)
  for (let index = 0; index < 40; index += 1) {
    const page = pdf.addPage([500, 700])
    const { height } = page.getSize()
    page.drawRectangle({ x: 0, y: 0, width: 500, height, color: rgb(1, 1, 1) })
    // Search target hugging the very top of selected pages.
    if (index % 6 === 4) page.drawText(`Edge marker ${index + 1}`, { x: 52, y: height - 22, size: 10, font: regular })
    page.drawText(`Section ${index + 1} heading`, { x: 52, y: height - 64, size: 22, font: bold })
    page.drawText('The quick brown fox jumps over', { x: 52, y: height - 120, size: 14, font: regular })
    page.drawText('the lazy dog beside the riverbank today.', { x: 52, y: height - 138, size: 14, font: regular })
    for (let line = 0; line < 22; line += 1) {
      page.drawText(`Body line ${line + 1} of page ${index + 1} keeps the reader busy.`, { x: 52, y: height - 190 - line * 22, size: 12, font: regular })
    }
    if (index % 6 === 1) page.drawText(`Edge marker ${index + 1}`, { x: 52, y: 18, size: 10, font: regular })
  }
  pdf.setTitle('simple viewer interaction fixture')
  await writeFile(target, await pdf.save({ useObjectStreams: true }))
}

async function withApp(fixture, port, smoke, label) {
  const profile = await mkdtemp(path.join(scratch, `simple-viewer-${label}-profile-`))
  let app
  try {
    app = spawn(executable, [
      ...(!packagedExecutable ? ['.'] : []),
      `--remote-debugging-port=${port}`,
      `--user-data-dir=${profile}`,
      fixture,
    ], { cwd: root, windowsHide: true, stdio: 'ignore' })
    return await run(process.execPath, [smoke], { env: { ...process.env, SIMPLE_BENCH_PORT: String(port) }, timeout: 240_000 })
  } finally {
    if (app && app.exitCode === null) {
      spawnSync('taskkill.exe', ['/pid', String(app.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
      await new Promise((resolve) => setTimeout(resolve, 500))
    }
    await rm(profile, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }).catch(() => {})
  }
}

const temporaryDirectory = await mkdtemp(path.join(scratch, 'simple-viewer-pdf-'))
const results = {}
try {
  const interaction = path.join(temporaryDirectory, 'interaction.pdf')
  const reading = path.join(temporaryDirectory, 'reading.pdf')
  if (only.has('search') || only.has('pointer')) {
    await run(process.execPath, ['scripts/create-interaction-fixture.mjs'])
    await copyFile(path.join(root, 'tmp', 'pdfs', 'simple-interaction-fixture.pdf'), interaction)
  }
  if (only.has('viewer') || only.has('long')) await createReadingFixture(reading)
  if (only.has('search')) {
    const output = await withApp(interaction, Number(process.env.SIMPLE_BENCH_PORT || 9411), 'scripts/smoke-search-selection.mjs', 'search')
    const parsed = JSON.parse(output.split('\n').at(-1))
    results.searchSelection = { passed: true, navigation: parsed.search?.navigation }
  }
  if (only.has('pointer')) {
    await withApp(interaction, 9413, 'scripts/smoke-pointer-selection.mjs', 'pointer')
    results.pointerSelection = { passed: true }
  }
  if (only.has('long')) {
    const output = await withApp(reading, 9414, 'scripts/smoke-long-selection.mjs', 'long')
    const parsed = JSON.parse(output.split('\n').at(-1))
    results.longSelection = { passed: true, released: parsed.released }
  }
  if (only.has('viewer')) {
    const output = await withApp(reading, Number(process.env.SIMPLE_BENCH_PORT_2 || 9412), 'scripts/smoke-viewer-interaction.mjs', 'reading')
    results.viewerInteraction = JSON.parse(output.split('\n').at(-1))
  }
  console.log(JSON.stringify(results))
} finally {
  await rm(temporaryDirectory, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 }).catch(() => {})
}
