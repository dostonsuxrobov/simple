import assert from 'node:assert/strict'
import { spawn, spawnSync } from 'node:child_process'
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'

const root = path.resolve('.')
const port = Number(process.env.SIMPLE_BENCH_PORT || 9495)
const packagedExecutable = process.env.SIMPLE_TEST_EXECUTABLE ? path.resolve(process.env.SIMPLE_TEST_EXECUTABLE) : ''
const executable = packagedExecutable || path.join(root, 'node_modules', 'electron', 'dist', 'electron.exe')
const temporaryDirectory = await mkdtemp(path.join(os.tmpdir(), 'simple-edit-productivity-'))
const target = path.join(temporaryDirectory, 'edit-productivity.pdf')
const profile = path.join(temporaryDirectory, 'profile')

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: root, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'], ...options })
    let stdout = ''
    let stderr = ''
    child.stdout.on('data', (chunk) => { stdout += chunk })
    child.stderr.on('data', (chunk) => { stderr += chunk })
    child.once('error', reject)
    child.once('exit', (code) => code === 0 ? resolve(stdout.trim()) : reject(new Error(`${path.basename(command)} exited ${code}\n${stdout}${stderr}`)))
  })
}

let app
try {
  await run(process.execPath, ['scripts/create-interaction-fixture.mjs'])
  await copyFile(path.join(root, 'tmp', 'pdfs', 'simple-interaction-fixture.pdf'), target)
  app = spawn(executable, [...(!packagedExecutable ? ['.'] : []), `--remote-debugging-port=${port}`, `--user-data-dir=${profile}`, target], {
    cwd: root, windowsHide: true, stdio: 'ignore',
  })
  const output = await run(process.execPath, ['scripts/smoke-edit-productivity.mjs'], {
    env: { ...process.env, SIMPLE_BENCH_PORT: String(port) },
  })
  const ui = JSON.parse(output.split('\n').at(-1))
  const pdf = await pdfjs.getDocument({ data: new Uint8Array(await readFile(target)), disableWorker: true, isEvalSupported: false }).promise
  try {
    const page = await pdf.getPage(1)
    const items = (await page.getTextContent()).items.filter((item) => 'str' in item && item.str.trim())
    const joined = items.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim()
    assert.ok(joined.includes(ui.wrappedText), 'Saved PDF lost or clipped part of the wrapped replacement')
    assert.equal(joined.split(ui.wrappedText).length - 1, 1, 'Wrapped replacement was written more than once')
    assert.ok(joined.includes(ui.sizedText), 'Saved PDF did not retain the size-adjusted replacement')
    const sized = items.filter((item) => item.str.includes(ui.sizedText))
    assert.equal(sized.length, 1, 'Expected a single saved size-adjusted text run')
    assert.ok(Math.abs(sized[0].height - ui.fontSize) < 0.1, `Saved font size is ${sized[0].height}, expected ${ui.fontSize}`)
    const wrapStart = items.findIndex((item) => item.str.startsWith(ui.wrappedText.split(' ')[0]))
    assert.ok(wrapStart >= 0, 'Saved wrapped paragraph start is missing')
    const wrappedItems = []
    for (const item of items.slice(wrapStart)) {
      wrappedItems.push(item)
      if (wrappedItems.map((part) => part.str).join(' ').replace(/\s+/g, ' ').trim() === ui.wrappedText) break
    }
    assert.equal(wrappedItems.map((item) => item.str).join(' ').replace(/\s+/g, ' ').trim(), ui.wrappedText, 'Saved wrapped lines are incomplete or out of order')
    const baselines = new Set(wrappedItems.map((item) => Math.round(item.transform[5] * 10) / 10))
    assert.ok(baselines.size >= 3, 'Saved replacement did not occupy multiple lines')
    const descendingBaselines = [...baselines].sort((a, b) => b - a)
    for (let index = 1; index < descendingBaselines.length; index += 1) {
      assert.ok(Math.abs(descendingBaselines[index - 1] - descendingBaselines[index] - 22) < 0.2, 'Saved line spacing does not match the requested 22pt')
    }
    console.log(JSON.stringify({ ui, saved: { pages: pdf.numPages, fontSize: sized[0].height, wrappedBaselines: baselines.size, completeText: true } }))
  } finally {
    await pdf.destroy()
  }
} finally {
  if (app && app.exitCode === null) {
    app.kill()
    await Promise.race([new Promise((resolve) => app.once('exit', resolve)), new Promise((resolve) => setTimeout(resolve, 2_000))])
    if (app.exitCode === null) spawnSync('taskkill.exe', ['/pid', String(app.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
  }
  const resolved = path.resolve(temporaryDirectory)
  const temporaryRoot = path.resolve(os.tmpdir()) + path.sep
  assert.ok(resolved.startsWith(temporaryRoot) && path.basename(resolved).startsWith('simple-edit-productivity-'), 'Refusing to remove an unexpected temporary path')
  await rm(resolved, { recursive: true, force: true, maxRetries: 8, retryDelay: 250 })
}
