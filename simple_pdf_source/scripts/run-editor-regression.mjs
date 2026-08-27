import { spawn, spawnSync } from 'node:child_process'
import { copyFile, mkdtemp, rm } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'

const root = path.resolve('.')
const port = Number(process.env.SIMPLE_BENCH_PORT || 9334)
const electron = path.join(root, 'node_modules', 'electron', 'dist', 'electron.exe')
const source = path.join(root, 'tmp', 'pdfs', 'simple-interaction-fixture.pdf')
const target = path.join(root, 'tmp', 'pdfs', 'simple-editor-saved.pdf')
const profile = await mkdtemp(path.join(os.tmpdir(), 'simple-editor-regression-'))

function run(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, {
      cwd: root,
      windowsHide: true,
      stdio: ['ignore', 'pipe', 'pipe'],
      ...options,
    })
    let stdout = ''
    let stderr = ''
    child.stdout?.on('data', (chunk) => { stdout += chunk })
    child.stderr?.on('data', (chunk) => { stderr += chunk })
    child.once('error', reject)
    child.once('exit', (code) => {
      if (code === 0) resolve(stdout.trim())
      else reject(new Error(`${path.basename(command)} exited ${code}\n${stdout}${stderr}`.trim()))
    })
  })
}

let app
try {
  await run(process.execPath, ['scripts/create-interaction-fixture.mjs'])
  await copyFile(source, target)
  app = spawn(electron, [
    '.',
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${profile}`,
    target,
  ], {
    cwd: root,
    windowsHide: true,
    stdio: 'ignore',
  })
  const output = await run(process.execPath, ['scripts/smoke-edit-regressions.mjs'], {
    env: { ...process.env, SIMPLE_BENCH_PORT: String(port) },
  })
  console.log(output)
} finally {
  if (app && app.exitCode === null) {
    app.kill()
    await Promise.race([
      new Promise((resolve) => app.once('exit', resolve)),
      new Promise((resolve) => setTimeout(resolve, 2_000)),
    ])
    if (app.exitCode === null) {
      spawnSync('taskkill.exe', ['/pid', String(app.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
    }
  }
  await rm(profile, { recursive: true, force: true })
}
