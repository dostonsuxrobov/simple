'use strict'

// safeWriteFile() and io-core against the lock and target matrix of the shared
// Save design (§1.6, §3.2, §10.1). A PowerShell child holds the target open
// with each Windows FileShare mode, the way Word, Acrobat, virus scanners and
// sync programs do. Every case runs in a fresh temp folder with its own save
// journal; Windows-only cases skip elsewhere. Every child process is killed.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const fsp = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const { spawn, spawnSync } = require('node:child_process')

const SHARED = path.resolve(__dirname, '..', 'shared')
const core = require(path.join(SHARED, 'electron', 'io-core.cjs'))
const { safeWriteFile, safeWriteResult } = require(path.join(SHARED, 'electron', 'safe-write.cjs'))
const { findServiceNames } = require('../scripts/local-only-guard.cjs')

const isWindows = process.platform === 'win32'
const windowsOnly = isWindows ? {} : { skip: 'needs Windows share modes and attributes' }
const FAST = [5, 10, 20]
const POWERSHELL = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
const ROOT = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-safe-write-'))
const JOURNAL = path.join(ROOT, 'journal')
let scenarioCount = 0

core.configureIo({ journalDir: JOURNAL, logger: { warn() {}, info() {} } })
test.after(() => {
  for (const file of walk(ROOT)) {
    try { fs.chmodSync(file, 0o666) } catch {}
  }
  fs.rmSync(ROOT, { recursive: true, force: true })
})

function walk(folder) {
  const files = []
  for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
    const full = path.join(folder, entry.name)
    if (entry.isDirectory()) files.push(...walk(full))
    else files.push(full)
  }
  return files
}

function scenario(name) {
  scenarioCount += 1
  const folder = path.join(ROOT, `${String(scenarioCount).padStart(2, '0')}-${name}`)
  fs.mkdirSync(folder, { recursive: true })
  return folder
}

function leftovers(folder) {
  return fs.readdirSync(folder).filter((name) => name.startsWith('~simple-'))
}

function original(folder, name, content = 'ORIGINAL') {
  const filePath = path.join(folder, name)
  fs.writeFileSync(filePath, content)
  return filePath
}

async function rejectsWith(promise, code) {
  try {
    await promise
  } catch (error) {
    assert.equal(error.code, code, `expected ${code}, got ${error.code}: ${error.message} (${error.technical || ''})`)
    assert.ok(core.isIoError(error), 'failures must be IoErrors')
    return error
  }
  assert.fail(`expected ${code}, but the write succeeded`)
}

function backupsLeft() {
  try {
    return fs.readdirSync(path.join(JOURNAL, 'in-place-backups'))
  } catch (error) {
    if (error.code === 'ENOENT') return []
    throw error
  }
}

function journalIsClean() {
  assert.deepEqual(core.journal.entries(), [], 'a finished save must leave no journal entry')
}

function deadPid() {
  return spawnSync(process.execPath, ['-e', ''], { windowsHide: true }).pid
}

function killTree(child) {
  if (!child || child.exitCode !== null || child.signalCode !== null) return Promise.resolve()
  const exited = new Promise((resolve) => child.once('exit', resolve))
  if (isWindows) spawnSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { windowsHide: true, stdio: 'ignore' })
  else child.kill('SIGKILL')
  return Promise.race([exited, new Promise((resolve) => setTimeout(resolve, 5000))])
}

function collectOutput(child) {
  child.output = ''
  child.stdout.setEncoding('utf8')
  child.stdout.on('data', (chunk) => { child.output += chunk })
  child.stderr.setEncoding('utf8')
  child.stderr.on('data', (chunk) => { child.output += chunk })
}

function waitForOutput(child, marker, timeoutMs) {
  return new Promise((resolve, reject) => {
    const check = () => {
      if (child.output.includes(marker)) {
        cleanup()
        resolve()
      }
    }
    const onExit = (code) => {
      cleanup()
      reject(new Error(`child exited (${code}) before printing ${marker}: ${child.output}`))
    }
    const timer = setTimeout(() => {
      cleanup()
      reject(new Error(`timed out waiting for ${marker}: ${child.output}`))
    }, timeoutMs)
    function cleanup() {
      clearTimeout(timer)
      child.stdout.off('data', check)
      child.stderr.off('data', check)
      child.off('exit', onExit)
    }
    child.stdout.on('data', check)
    child.stderr.on('data', check)
    child.once('exit', onExit)
    check()
  })
}

/**
 * Opens `file` for reading in a PowerShell child with the given FileShare
 * mode. With `releaseOnGo`, the child waits for go(), keeps the handle 700 ms
 * longer, then closes it.
 */
async function holdFile(t, file, share, { releaseOnGo = false } = {}) {
  const script = [
    "$ErrorActionPreference = 'Stop'",
    `$stream = [IO.File]::Open('${file.replace(/'/g, "''")}', [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]'${share}')`,
    "[Console]::Out.WriteLine('READY'); [Console]::Out.Flush()",
    ...(releaseOnGo ? [
      '[void][Console]::In.ReadLine()',
      'Start-Sleep -Milliseconds 700',
      '$stream.Dispose()',
      "[Console]::Out.WriteLine('RELEASED'); [Console]::Out.Flush()",
    ] : []),
    '[void][Console]::In.ReadLine()',
    '$stream.Dispose()',
  ].join('\n')
  const child = spawn(POWERSHELL, ['-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64')], {
    windowsHide: true,
    stdio: ['pipe', 'pipe', 'pipe'],
  })
  collectOutput(child)
  t.after(() => killTree(child))
  await waitForOutput(child, 'READY', 60_000)
  return {
    go() { child.stdin.write('go\n') },
    async release() {
      if (child.exitCode === null) {
        child.stdin.write('done\n')
        child.stdin.write('done\n')
        await Promise.race([new Promise((resolve) => child.once('exit', resolve)), new Promise((resolve) => setTimeout(resolve, 5000))])
      }
      await killTree(child)
    },
  }
}

function withInjectedFs(overrides) {
  return { ...fsp, ...overrides }
}

function failing(code, syscall = 'write') {
  return Object.assign(new Error(`${code}: injected failure`), { code, syscall })
}

// ---------------------------------------------------------------------------
// The lock matrix (§1.6)
// ---------------------------------------------------------------------------

test('no other handle: replaced by rename, with nothing left behind', async () => {
  const folder = scenario('control')
  const target = original(folder, 'control.xlsx')
  const result = await safeWriteFile(target, 'NEW CONTENT')
  assert.equal(result.strategy, 'rename')
  assert.equal(result.attempts, 1)
  assert.equal(result.path, target)
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW CONTENT')
  assert.deepEqual(result.stamp, { size: 11, mtimeMs: fs.statSync(target).mtimeMs, sha256: core.hashBytes('NEW CONTENT') })
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('a new file is created with strategy "new"', async () => {
  const folder = scenario('new-file')
  const result = await safeWriteFile(path.join(folder, 'fresh.txt'), Buffer.from('hello'))
  assert.equal(result.strategy, 'new')
  assert.equal(fs.readFileSync(path.join(folder, 'fresh.txt'), 'utf8'), 'hello')
  assert.deepEqual(leftovers(folder), [])
})

for (const share of ['None', 'Read']) {
  test(`held with FileShare.${share}: LOCKED, the original is byte-identical, and nothing is left behind`, windowsOnly, async (t) => {
    const folder = scenario(`held-${share}`)
    const target = original(folder, `held-${share}.pdf`)
    const before = fs.readFileSync(target)
    const holder = await holdFile(t, target, share)
    const started = Date.now()
    const error = await rejectsWith(safeWriteFile(target, 'NEW CONTENT'), 'LOCKED')
    assert.ok(Date.now() - started >= 3000, 'the retry budget of about 3.5 s must be used before giving up')
    assert.match(error.technical, /EBUSY .*after \d+ attempts/)
    assert.equal(error.fileName, `held-${share}.pdf`)
    await holder.release()
    assert.deepEqual(fs.readFileSync(target), before)
    assert.deepEqual(leftovers(folder), [])
    assert.deepEqual(backupsLeft(), [])
    journalIsClean()
  })
}

test('held with FileShare.ReadWrite: saved by the verified in-place write', windowsOnly, async (t) => {
  const folder = scenario('held-readwrite')
  const target = original(folder, 'held-rw.docx')
  const holder = await holdFile(t, target, 'ReadWrite')
  const result = await safeWriteFile(target, 'NEW CONTENT FOR READWRITE')
  assert.equal(result.strategy, 'in-place')
  assert.ok(result.attempts > 1)
  await holder.release()
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW CONTENT FOR READWRITE')
  assert.deepEqual(leftovers(folder), [])
  assert.deepEqual(backupsLeft(), [], 'the in-place backup is removed after a verified write')
  journalIsClean()
})

test('held with FileShare.ReadWrite and Delete (scanner or sync program): saved by the two-step swap', windowsOnly, async (t) => {
  const folder = scenario('held-readwrite-delete')
  const target = original(folder, 'held-rwd.xlsx')
  const holder = await holdFile(t, target, 'ReadWrite, Delete')
  const result = await safeWriteFile(target, 'NEW CONTENT FOR SWAP')
  assert.equal(result.strategy, 'swap')
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW CONTENT FOR SWAP')
  await holder.release()
  assert.deepEqual(leftovers(folder), [], 'the moved-aside original disappears once the holder lets go')
  journalIsClean()
})

test('a lock released after 700 ms: saved by rename on a later attempt', windowsOnly, async (t) => {
  const folder = scenario('released')
  const target = original(folder, 'released.docx')
  const holder = await holdFile(t, target, 'Read', { releaseOnGo: true })
  holder.go()
  const result = await safeWriteFile(target, 'NEW CONTENT AFTER RELEASE')
  // The release can fall between the rename and the move-aside of one attempt; then the swap finishes it.
  assert.ok(['rename', 'swap'].includes(result.strategy), result.strategy)
  assert.ok(result.attempts > 1, `expected retries, got ${result.attempts} attempt(s)`)
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW CONTENT AFTER RELEASE')
  await holder.release()
  assert.deepEqual(leftovers(folder), [])
})

test('read-only attribute: READ_ONLY, and both the content and the attribute are kept', async () => {
  const folder = scenario('read-only')
  const target = original(folder, 'readonly.png')
  fs.chmodSync(target, 0o444)
  try {
    await rejectsWith(safeWriteFile(target, 'NEW CONTENT'), 'READ_ONLY')
    assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
    assert.equal(fs.statSync(target).mode & 0o200, 0, 'Simple must never clear the read-only attribute')
    if (isWindows) assert.match(spawnSync('attrib', [target], { encoding: 'utf8', windowsHide: true }).stdout, /^\s*A?\s*R/)
    assert.deepEqual(leftovers(folder), [])
  } finally {
    fs.chmodSync(target, 0o666)
  }
})

test('missing folder: FOLDER_MISSING with the nearest folder that still exists', async () => {
  const folder = scenario('missing-folder')
  const error = await rejectsWith(safeWriteFile(path.join(folder, 'gone', 'deeper', 'file.docx'), 'NEW'), 'FOLDER_MISSING')
  assert.equal(error.nearestFolder, folder)
  assert.equal(error.folder, path.join(folder, 'gone', 'deeper'))
  assert.match(error.message, /isn't available anymore/)
  assert.ok(!fs.existsSync(path.join(folder, 'gone')))
})

test('a 229-character file name saves; a 256-character one is NAME_TOO_LONG', async () => {
  const folder = scenario('long-name')
  const target = original(folder, `${'L'.repeat(224)}.xlsx`)
  const result = await safeWriteFile(target, 'NEW CONTENT')
  assert.equal(result.strategy, 'rename')
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW CONTENT')
  await rejectsWith(safeWriteFile(path.join(folder, `${'x'.repeat(251)}.xlsx`), 'NEW'), 'NAME_TOO_LONG')
  assert.deepEqual(leftovers(folder), [])
})

test('names Windows cannot store are INVALID_NAME before anything is written', async () => {
  const folder = scenario('invalid-name')
  for (const name of ['CON.txt', 'report.', 'trailing ', 'a?b.txt']) {
    const error = await rejectsWith(safeWriteFile(path.join(folder, name), 'NEW'), 'INVALID_NAME')
    assert.ok(error.reason)
  }
  fs.mkdirSync(path.join(folder, 'taken.docx'))
  const error = await rejectsWith(safeWriteFile(path.join(folder, 'taken.docx'), 'NEW'), 'INVALID_NAME')
  assert.equal(error.reason, 'folder-exists')
  assert.deepEqual(leftovers(folder), [])
})

test('a hard-linked file is written in place, so the other link sees the new content', async () => {
  const folder = scenario('hard-link')
  const first = original(folder, 'hl-a.pdf')
  const second = path.join(folder, 'hl-b.pdf')
  fs.linkSync(first, second)
  const result = await safeWriteFile(first, 'NEW CONTENT FOR BOTH LINKS')
  assert.equal(result.strategy, 'in-place')
  assert.equal(fs.readFileSync(second, 'utf8'), 'NEW CONTENT FOR BOTH LINKS')
  assert.equal(fs.statSync(first).nlink, 2)
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('identity "always" writes in place even without hard links', async () => {
  const folder = scenario('identity-always')
  const target = original(folder, 'keep-identity.txt')
  const before = fs.statSync(target)
  const result = await safeWriteFile(target, 'IN PLACE', { identity: 'always' })
  assert.equal(result.strategy, 'in-place')
  assert.equal(fs.readFileSync(target, 'utf8'), 'IN PLACE')
  if (isWindows) assert.equal(fs.statSync(target).birthtimeMs, before.birthtimeMs, 'the same file keeps its creation time')
})

test('stamps: a re-stamped file with the same content saves; an external edit is CHANGED_ON_DISK; force overwrites', async () => {
  const folder = scenario('stamps')
  const target = original(folder, 'stamp.csv', 'a,b\n1,2\n')
  const stamp = await core.stampFile(target)
  const past = new Date(Date.now() - 5 * 60 * 1000)
  fs.utimesSync(target, past, past)
  assert.equal((await core.compareStamp(target, stamp)).state, 'restamped')
  const saved = await safeWriteFile(target, 'a,b\n1,3\n', { expectedStamp: stamp })
  assert.equal(fs.readFileSync(target, 'utf8'), 'a,b\n1,3\n')

  fs.writeFileSync(target, 'EDITED ELSEWHERE')
  const error = await rejectsWith(safeWriteFile(target, 'MINE', { expectedStamp: saved.stamp }), 'CHANGED_ON_DISK')
  assert.match(error.message, /was changed by another program/)
  assert.equal(fs.readFileSync(target, 'utf8'), 'EDITED ELSEWHERE')
  assert.deepEqual(leftovers(folder), [])

  await safeWriteFile(target, 'MINE', { expectedStamp: saved.stamp, force: true })
  assert.equal(fs.readFileSync(target, 'utf8'), 'MINE')
})

test('a deleted source is SOURCE_MISSING until the user chooses Save Here Again', async () => {
  const folder = scenario('source-missing')
  const target = original(folder, 'moved.docx')
  const stamp = await core.stampFile(target)
  fs.unlinkSync(target)
  await rejectsWith(safeWriteFile(target, 'NEW', { expectedStamp: stamp }), 'SOURCE_MISSING')
  assert.ok(!fs.existsSync(target))
  const result = await safeWriteFile(target, 'NEW', { expectedStamp: stamp, recreate: true })
  assert.equal(result.strategy, 'new')
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW')
})

test('a hidden file is replaced (Windows drops the attribute on rename, as documented)', windowsOnly, async () => {
  const folder = scenario('hidden')
  const target = original(folder, 'hidden.txt')
  spawnSync('attrib', ['+H', target], { windowsHide: true })
  const result = await safeWriteFile(target, 'NEW HIDDEN CONTENT')
  assert.ok(['rename', 'swap'].includes(result.strategy))
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW HIDDEN CONTENT')
})

test('a symbolic link survives: the file it points to is written', async (t) => {
  const folder = scenario('symlink')
  const real = original(folder, 'real.txt')
  const link = path.join(folder, 'link.txt')
  try {
    fs.symlinkSync(real, link, 'file')
  } catch (error) {
    t.skip(`symbolic links are not available here (${error.code})`)
    return
  }
  const result = await safeWriteFile(link, 'THROUGH THE LINK')
  assert.equal(result.path, fs.realpathSync(real))
  assert.ok(fs.lstatSync(link).isSymbolicLink())
  assert.equal(fs.readFileSync(real, 'utf8'), 'THROUGH THE LINK')
})

// ---------------------------------------------------------------------------
// Crashes and the startup sweep (§3.2.1)
// ---------------------------------------------------------------------------

test('a save killed after its temp write leaves the original intact, and sweep() removes the temp', async (t) => {
  const folder = scenario('killed')
  const target = original(folder, 'crash.docx')
  const childJournal = path.join(folder, 'journal')
  const script = `
    const fs = require('node:fs')
    const core = require(${JSON.stringify(path.join(SHARED, 'electron', 'io-core.cjs'))})
    const { safeWriteFile } = require(${JSON.stringify(path.join(SHARED, 'electron', 'safe-write.cjs'))})
    core.configureIo({ journalDir: ${JSON.stringify(childJournal)} })
    safeWriteFile(${JSON.stringify(target)}, Buffer.from('NEW CONTENT'), {
      onProgress(progress) {
        if (progress.phase !== 'replacing') return
        fs.writeSync(1, 'PAUSED\\n')
        Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, 120000)
      },
    })
  `
  const child = spawn(process.execPath, ['-e', script], { windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] })
  collectOutput(child)
  t.after(() => killTree(child))
  await waitForOutput(child, 'PAUSED', 60_000)
  await killTree(child)
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  const temps = leftovers(folder)
  assert.equal(temps.length, 1)
  assert.match(temps[0], /^~simple-[0-9a-f]{8}\.tmp$/)

  const unrelated = path.join(folder, '~simple-0badc0de.tmp')
  fs.writeFileSync(unrelated, 'not recorded in any journal')
  const report = await core.sweep({ journalDir: childJournal })
  assert.equal(report.journals, 1)
  assert.deepEqual(report.actions.map((action) => action.outcome), ['removed-temp'])
  assert.deepEqual(leftovers(folder), ['~simple-0badc0de.tmp'], 'sweep() never deletes files by name pattern alone')
  assert.deepEqual(fs.readdirSync(childJournal), [])
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
})

function writeJournal(folder, pid, entries) {
  fs.mkdirSync(folder, { recursive: true })
  const file = path.join(folder, `journal-${pid}-0123abcd.json`)
  fs.writeFileSync(file, JSON.stringify({ schema: 1, pid, entries }))
  return file
}

test('sweep() completes a swap whose verified temp survived, or moves the original back', async () => {
  const folder = scenario('sweep-swap')
  const journalFolder = path.join(folder, 'journal')
  const pid = deadPid()

  const completed = path.join(folder, 'completed.xlsx')
  const completedTemp = original(folder, '~simple-11111111.tmp', 'NEW VERIFIED')
  const completedAside = original(folder, '~simple-22222222.old', 'ORIGINAL A')
  const rolledBack = path.join(folder, 'rolled-back.xlsx')
  const brokenTemp = original(folder, '~simple-33333333.tmp', 'PARTIAL')
  const rolledAside = original(folder, '~simple-44444444.old', 'ORIGINAL B')
  writeJournal(journalFolder, pid, [
    { id: 'a', target: completed, temp: completedTemp, aside: completedAside, phase: 'swapping', sha256: core.hashBytes('NEW VERIFIED') },
    { id: 'b', target: rolledBack, temp: brokenTemp, aside: rolledAside, phase: 'swapping', sha256: core.hashBytes('SOMETHING ELSE') },
  ])
  const report = await core.sweep({ journalDir: journalFolder })
  assert.deepEqual(report.actions.map((action) => action.outcome), ['completed', 'rolled-back'])
  assert.equal(fs.readFileSync(completed, 'utf8'), 'NEW VERIFIED')
  assert.equal(fs.readFileSync(rolledBack, 'utf8'), 'ORIGINAL B')
  assert.deepEqual(leftovers(folder), [])
})

test('sweep() undoes an interrupted in-place write from its backup and skips journals of live processes', async () => {
  const folder = scenario('sweep-in-place')
  const journalFolder = path.join(folder, 'journal')
  // The crash came after five bytes of the new content: new prefix, original rest.
  const target = original(folder, 'linked.pdf', 'NEW CNAL CONTENT')
  const temp = original(folder, '~simple-55555555.tmp', 'NEW CONTENT!!!!!')
  const backup = original(folder, 'backup.bin', 'ORIGINAL CONTENT')
  writeJournal(journalFolder, deadPid(), [
    {
      id: 'c', target, temp, backup, phase: 'in-place',
      sha256: core.hashBytes('NEW CONTENT!!!!!'), originalSha256: core.hashBytes('ORIGINAL CONTENT'),
    },
  ])
  const liveTemp = original(folder, '~simple-66666666.tmp', 'IN FLIGHT')
  const liveJournal = path.join(folder, 'live')
  writeJournal(liveJournal, process.ppid || 1, [{ id: 'd', target: path.join(folder, 'other.txt'), temp: liveTemp, phase: 'writing' }])

  const report = await core.sweep({ journalDir: journalFolder })
  assert.deepEqual(report.actions.map((action) => action.outcome), ['rolled-back'])
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL CONTENT')
  assert.ok(!fs.existsSync(backup))
  assert.ok(!fs.existsSync(temp))

  const liveReport = await core.sweep({ journalDir: liveJournal })
  assert.equal(liveReport.journals, 0, 'a journal another running process still updates is left alone')
  assert.ok(fs.existsSync(liveTemp))

  // While Simple holds its single-instance lock, no other live process owns a journal
  // in its folder: a reused process id does not hide a crashed save.
  const exclusiveReport = await core.sweep({ journalDir: liveJournal, exclusive: true })
  assert.equal(exclusiveReport.journals, 1)
  assert.ok(!fs.existsSync(liveTemp), 'the crashed save\'s temp is removed at once, not 30 minutes later')
})

test('sweep() never writes over a file whose content it cannot account for; the original is kept beside it', async () => {
  const folder = scenario('sweep-unknown')
  const journalFolder = path.join(folder, 'journal')
  // Newer work (saved later, by Simple or another program) is in the file.
  const target = original(folder, 'notes.txt', 'EDIT TWO')
  const temp = original(folder, '~simple-77777777.tmp', 'EDIT ONE')
  const backup = original(folder, 'backup.bin', 'ORIGINAL')
  writeJournal(journalFolder, deadPid(), [
    { id: 'u', target, temp, backup, phase: 'in-place', restoreNeeded: true, sha256: core.hashBytes('EDIT ONE'), originalSha256: core.hashBytes('ORIGINAL') },
  ])
  const report = await core.sweep({ journalDir: journalFolder })
  assert.deepEqual(report.actions.map((action) => action.outcome), ['kept-original'])
  assert.equal(fs.readFileSync(target, 'utf8'), 'EDIT TWO', 'the newest work is never replaced')
  const kept = path.join(folder, 'notes (original).txt')
  assert.equal(report.actions[0].detail, kept)
  assert.equal(fs.readFileSync(kept, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
})

test('RESTORE_NEEDED, then a later save of the same file, then sweep(): the later save stays', async () => {
  const folder = scenario('restore-needed-superseded')
  const target = original(folder, 'report.txt', 'ORIGINAL')
  // The in-place write fails, and so does writing the original back (the folder dropped).
  const broken = withInjectedFs({
    async open(filePath, flags, mode) {
      const handle = await fsp.open(filePath, flags, mode)
      if (flags !== 'r+') return handle
      return new Proxy(handle, {
        get(real, property) {
          if (property === 'write') return async () => { throw failing('EIO') }
          const value = real[property]
          return typeof value === 'function' ? value.bind(real) : value
        },
      })
    },
  })
  const error = await rejectsWith(safeWriteFile(target, 'EDIT ONE', { fs: broken, identity: 'always', retryDelays: FAST }), 'RESTORE_NEEDED')
  assert.ok(error.backupPath && fs.existsSync(error.backupPath))
  const kept = core.journal.entries().find((entry) => entry.target === target)
  assert.equal(kept.phase, 'in-place')
  assert.equal(kept.restoreNeeded, true, 'the entry is kept for sweep()')

  // The user keeps editing and saves again; this time it works.
  await safeWriteFile(target, 'EDIT TWO')
  const superseded = core.journal.entries().find((entry) => entry.id === kept.id)
  assert.equal(superseded.superseded, true, 'the later save marks the failed one superseded')

  // At the next start the failed save's leftovers go, and the later save stays.
  const deadJournal = path.join(folder, 'journal')
  writeJournal(deadJournal, deadPid(), [superseded])
  await core.journal.remove(superseded.id)
  const report = await core.sweep({ journalDir: deadJournal })
  assert.deepEqual(report.actions.map((action) => action.outcome), ['superseded'])
  assert.equal(fs.readFileSync(target, 'utf8'), 'EDIT TWO')
  assert.ok(!fs.existsSync(error.backupPath), 'the obsolete backup is removed')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('sweep() never restores a backup it cannot prove complete', async () => {
  const folder = scenario('sweep-partial-backup')
  const journalFolder = path.join(folder, 'journal')
  // Crash while the original was still being copied: the target is untouched.
  const untouched = original(folder, 'untouched.docx', 'ORIGINAL A')
  const partial = original(folder, 'partial-backup.bin', 'ORIG')
  // Crash mid-write with a backup that does not match the recorded original.
  const halfWritten = original(folder, 'half.docx', 'NEW B, HALF')
  const wrongBackup = original(folder, 'wrong-backup.bin', 'SOMETHING ELSE')
  writeJournal(journalFolder, deadPid(), [
    { id: 'e', target: untouched, backup: partial, phase: 'backing-up', sha256: core.hashBytes('NEW A') },
    { id: 'f', target: halfWritten, backup: wrongBackup, phase: 'in-place', sha256: core.hashBytes('NEW B'), originalSha256: core.hashBytes('ORIGINAL B') },
  ])
  const report = await core.sweep({ journalDir: journalFolder })
  assert.deepEqual(report.actions.map((action) => action.outcome), ['removed-temp', 'failed'])
  assert.equal(fs.readFileSync(untouched, 'utf8'), 'ORIGINAL A')
  assert.ok(!fs.existsSync(partial), 'a partial backup is removed')
  assert.equal(fs.readFileSync(halfWritten, 'utf8'), 'NEW B, HALF', 'nothing is overwritten without a proven backup')
  assert.equal(fs.readFileSync(wrongBackup, 'utf8'), 'SOMETHING ELSE', 'and nothing is deleted')
  assert.deepEqual(core.journal.entries(), [], 'a failure that cannot improve is reported, not retried forever')
})

test('sweep() moves an unreadable journal aside instead of guessing', async () => {
  const folder = scenario('sweep-damaged')
  fs.writeFileSync(path.join(folder, `journal-${deadPid()}-0123abcd.json`), '{ not json')
  const report = await core.sweep({ journalDir: folder })
  assert.equal(report.journals, 0)
  assert.ok(fs.readdirSync(folder).some((name) => name.startsWith('damaged-')))
})

// ---------------------------------------------------------------------------
// Injected failures
// ---------------------------------------------------------------------------

test('ENOSPC while writing the temp file: DISK_FULL, the target untouched, and no temp left', async () => {
  const folder = scenario('enospc')
  const target = original(folder, 'big.pdf')
  const injected = withInjectedFs({
    async open(filePath, flags, mode) {
      const handle = await fsp.open(filePath, flags, mode)
      if (flags !== 'wx') return handle
      return new Proxy(handle, {
        get(real, property) {
          if (property === 'write') return async () => { throw failing('ENOSPC') }
          const value = real[property]
          return typeof value === 'function' ? value.bind(real) : value
        },
      })
    },
  })
  const error = await rejectsWith(safeWriteFile(target, Buffer.alloc(1024, 1), { fs: injected }), 'DISK_FULL')
  assert.match(error.drive, /drive [A-Z]:|this drive|\\\\/)
  assert.match(error.message, /isn't enough space/)
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('too little free space is DISK_FULL before anything is written, with the amount to free', async () => {
  const folder = scenario('statfs')
  const target = original(folder, 'huge.xlsx')
  const injected = withInjectedFs({ async statfs() { return { bavail: 10, bsize: 4096, blocks: 1_000_000 } } })
  const error = await rejectsWith(safeWriteFile(target, Buffer.alloc(4 * 1024 * 1024), { fs: injected }), 'DISK_FULL')
  assert.ok(error.needed > 4 * 1024 * 1024)
  assert.match(core.toIoResult(error).needed, /MB$/)
  assert.deepEqual(leftovers(folder), [])
})

function swapFailure({ restoreFails }) {
  // The first rename over the target fails the way it does while a scanner
  // holds it; moving the original aside works; the second rename and,
  // optionally, the move back fail for good.
  let tempRenames = 0
  return withInjectedFs({
    async rename(from, to) {
      const base = path.basename(from)
      if (base.endsWith('.tmp')) {
        tempRenames += 1
        throw failing(tempRenames === 1 ? 'EPERM' : 'EIO', 'rename')
      }
      if (base.endsWith('.old') && restoreFails) throw failing('EIO', 'rename')
      return fsp.rename(from, to)
    },
  })
}

test('a swap whose second rename fails puts the original back', async () => {
  const folder = scenario('swap-rollback')
  const target = original(folder, 'swap.docx')
  const error = await rejectsWith(safeWriteFile(target, 'NEW', { fs: swapFailure({ restoreFails: false }), retryDelays: FAST }), 'FILE_UNAVAILABLE')
  assert.match(error.technical, /EIO rename/)
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('a swap that cannot put the original back reports RESTORE_NEEDED, and sweep() finishes the save later', async () => {
  const folder = scenario('swap-restore-needed')
  const target = original(folder, 'swap.docx')
  const error = await rejectsWith(safeWriteFile(target, 'NEW VERIFIED CONTENT', { fs: swapFailure({ restoreFails: true }), retryDelays: FAST }), 'RESTORE_NEEDED')
  assert.ok(error.asidePath && fs.existsSync(error.asidePath))
  assert.equal(fs.readFileSync(error.asidePath, 'utf8'), 'ORIGINAL', 'the original is intact beside the target')
  assert.equal(core.describeIoError(error).buttons.map((button) => button.id).join(), 'save-as,show-original,cancel')
  assert.ok(!fs.existsSync(target))

  // Hand this process's journal entry to a "dead" journal, as after a crash.
  const entry = core.journal.entries().find((item) => item.target === target)
  assert.equal(entry.phase, 'swapping')
  const deadJournal = path.join(folder, 'journal')
  writeJournal(deadJournal, deadPid(), [entry])
  await core.journal.remove(entry.id)
  assert.equal(entry.restoreNeeded, true)
  const report = await core.sweep({ journalDir: deadJournal })
  assert.deepEqual(report.actions.map((action) => action.outcome), ['completed'])
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW VERIFIED CONTENT')
  // The prompt told the user the original is safe: it is kept under a visible name, never deleted.
  const keptOriginal = path.join(folder, 'swap (original).docx')
  assert.equal(report.actions[0].detail, keptOriginal)
  assert.equal(fs.readFileSync(keptOriginal, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
})

test('bytes that read back differently are VERIFY_FAILED, and the target is untouched', async () => {
  const folder = scenario('verify-failed')
  const target = original(folder, 'flip.bin')
  const injected = withInjectedFs({
    async open(filePath, flags, mode) {
      const handle = await fsp.open(filePath, flags, mode)
      if (flags !== 'r' || !path.basename(filePath).startsWith('~simple-')) return handle
      return new Proxy(handle, {
        get(real, property) {
          if (property === 'read') {
            return async (buffer, offset, length, position) => {
              const result = await real.read(buffer, offset, length, position)
              if (result.bytesRead && position === 0) buffer[offset] ^= 0xff
              return result
            }
          }
          const value = real[property]
          return typeof value === 'function' ? value.bind(real) : value
        },
      })
    },
  })
  await rejectsWith(safeWriteFile(target, 'NEW CONTENT', { fs: injected }), 'VERIFY_FAILED')
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
})

test('a temp file that disappears while it is read back is VERIFY_FAILED, not a missing folder', async () => {
  const folder = scenario('verify-enoent')
  const target = original(folder, 'quarantined.txt')
  const injected = withInjectedFs({
    async open(filePath, flags, mode) {
      if (flags === 'r' && path.basename(filePath).startsWith('~simple-')) throw failing('ENOENT', 'open')
      return fsp.open(filePath, flags, mode)
    },
  })
  const error = await rejectsWith(safeWriteFile(target, 'NEW CONTENT', { fs: injected, retryDelays: FAST }), 'VERIFY_FAILED')
  assert.match(error.technical, /ENOENT/)
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('a file another program writes while Simple saves is never replaced: CHANGED_ON_DISK', async () => {
  const folder = scenario('changed-while-saving')
  const target = original(folder, 'big.pdf', 'VERSION ONE')
  const stamp = await core.stampFile(target)
  // The validator runs after the temp is written and read back, right before the replace.
  const error = await rejectsWith(safeWriteFile(target, 'MINE', {
    expectedStamp: stamp,
    validate: () => { fs.writeFileSync(target, 'NEWER VERSION FROM ANOTHER PROGRAM') },
  }), 'CHANGED_ON_DISK')
  assert.equal(error.reason, 'changed-while-saving')
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEWER VERSION FROM ANOTHER PROGRAM')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()

  // The same check runs before an in-place write.
  const linked = original(folder, 'linked.txt', 'VERSION ONE')
  const linkedStamp = await core.stampFile(linked)
  await rejectsWith(safeWriteFile(linked, 'MINE', {
    expectedStamp: linkedStamp,
    identity: 'always',
    validate: () => { fs.writeFileSync(linked, 'NEWER') },
  }), 'CHANGED_ON_DISK')
  assert.equal(fs.readFileSync(linked, 'utf8'), 'NEWER')
  assert.deepEqual(backupsLeft(), [])
})

test('a new file is never written over one another program created meanwhile', async () => {
  const folder = scenario('created-while-saving')
  const target = path.join(folder, 'export.csv')
  const error = await rejectsWith(safeWriteFile(target, 'MINE', {
    validate: () => { fs.writeFileSync(target, 'CREATED BY ANOTHER PROGRAM') },
  }), 'CHANGED_ON_DISK')
  assert.equal(error.reason, 'created-while-saving')
  assert.equal(fs.readFileSync(target, 'utf8'), 'CREATED BY ANOTHER PROGRAM')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
  // With the user's Replace, the new file replaces it.
  const replaced = await safeWriteFile(target, 'MINE', { force: true })
  assert.equal(replaced.strategy, 'rename')
  assert.equal(fs.readFileSync(target, 'utf8'), 'MINE')
})

test('a failing safety copy of the original is reported as such, never as "open in another program"', async () => {
  const folder = scenario('backup-fails')
  const target = original(folder, 'kept.txt')
  const notAFolder = original(folder, 'not-a-folder.bin', 'x')
  const error = await rejectsWith(safeWriteFile(target, 'NEW', { identity: 'always', backupDir: path.join(notAFolder, 'versions') }), 'UNKNOWN')
  assert.equal(error.reason, 'backup')
  assert.match(error.technical, /safety copy of the original/)
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('without a working save journal, no step that needs one runs: the original stays as it was', async (t) => {
  const folder = scenario('journal-broken')
  const target = original(folder, 'linked.docx')
  const blocker = original(folder, 'journal-is-a-file', 'x')
  core.configureIo({ journalDir: path.join(blocker, 'journal') })
  t.after(() => core.configureIo({ journalDir: JOURNAL }))
  const error = await rejectsWith(safeWriteFile(target, 'NEW', { identity: 'always' }), 'UNKNOWN')
  assert.equal(error.reason, 'journal')
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
  // A plain rename needs no journal record to stay safe.
  const renamed = await safeWriteFile(original(folder, 'plain.txt'), 'NEW')
  assert.equal(renamed.strategy, 'rename')
  core.configureIo({ journalDir: JOURNAL })
  for (const entry of core.journal.entries()) await core.journal.remove(entry.id)
})

test('icacls output: a file with permissions of its own is recognised in every language', () => {
  const { parseOwnPermissions } = require(path.join(SHARED, 'electron', 'safe-write.cjs'))
  const inherited = 'C:\\x\\a.csv NT AUTHORITY\\SYSTEM:(I)(F)\r\n           BUILTIN\\Administrators:(I)(F)\r\n\r\nSuccessfully processed 1 files; Failed processing 0 files\r\n'
  const explicit = 'C:\\x\\a.csv PC\\me:(F)\r\n\r\nErfolgreich verarbeitet: 1 Dateien\r\n'
  const mixed = 'C:\\x\\a.csv PC\\Accounting:(R)\r\n           NT AUTHORITY\\SYSTEM:(I)(F)\r\n'
  assert.equal(parseOwnPermissions(inherited), false)
  assert.equal(parseOwnPermissions(explicit), true)
  assert.equal(parseOwnPermissions(mixed), true)
  assert.equal(parseOwnPermissions(''), false)
})

test('a file with its own permissions keeps them: it is written in place, not replaced by a rename', windowsOnly, async () => {
  const folder = scenario('own-permissions')
  const target = original(folder, 'salaries.csv')
  const icacls = path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'icacls.exe')
  const user = process.env.USERNAME
  const set = spawnSync(icacls, [target, '/inheritance:r', '/grant:r', `${user}:F`], { windowsHide: true, encoding: 'latin1' })
  if (set.status !== 0) return
  const result = await safeWriteFile(target, 'NEW SALARIES')
  assert.equal(result.strategy, 'in-place')
  assert.ok(result.notes.some((note) => /own permissions/.test(note)))
  assert.equal(fs.readFileSync(target, 'utf8'), 'NEW SALARIES')
  const after = spawnSync(icacls, [target], { windowsHide: true, encoding: 'latin1' }).stdout
  assert.doesNotMatch(after, /\(I\)/, 'no inherited access rule came back')
  assert.deepEqual(leftovers(folder), [])
  // An ordinary file still uses the atomic rename.
  const plain = original(folder, 'plain.csv')
  assert.equal((await safeWriteFile(plain, 'NEW')).strategy, 'rename')
})

test('a structurally broken file is VALIDATION_FAILED before it can replace the original', async () => {
  const folder = scenario('validation')
  const sample = path.resolve(__dirname, '..', '..', 'Simple test examples', 'Complex document.docx')
  const target = path.join(folder, 'report.docx')
  fs.copyFileSync(sample, target)
  const good = fs.readFileSync(target)
  const truncated = good.subarray(0, Math.floor(good.length * 0.9))
  const error = await rejectsWith(safeWriteFile(target, truncated, { format: 'docx' }), 'VALIDATION_FAILED')
  assert.match(error.reason, /end-of-archive/)
  assert.deepEqual(fs.readFileSync(target), good)
  assert.deepEqual(leftovers(folder), [])

  const saved = await safeWriteFile(target, good, { format: 'docx' })
  assert.equal(saved.strategy, 'rename')

  const rejected = await rejectsWith(safeWriteFile(target, good, { format: 'docx', validate: () => ({ ok: false, reason: 'the review layer is missing' }) }), 'VALIDATION_FAILED')
  assert.equal(rejected.reason, 'the review layer is missing')
  const thrown = await rejectsWith(safeWriteFile(target, good, { validate: () => { throw new Error('pdf-lib could not load it') } }), 'VALIDATION_FAILED')
  assert.match(thrown.reason, /pdf-lib/)
  let seen = null
  await safeWriteFile(target, good, { validate: (bytes, info) => { seen = { length: bytes.length, info } } })
  assert.equal(seen.length, good.length)
  assert.equal(seen.info.target, target)

  const unknown = await safeWriteFile(path.join(folder, 'notes.unknownformat'), 'x', { format: 'unknownformat' })
  assert.match(unknown.notes.join(), /no structural check/)
  assert.deepEqual(leftovers(folder), [])
})

test('a producer streams large output through the same checks', async () => {
  const folder = scenario('producer')
  const target = original(folder, 'stream.json')
  const progress = []
  const result = await safeWriteFile(target, async (sink) => {
    await sink.write('{"rows":[')
    for (let index = 0; index < 1000; index += 1) await sink.write(`${index ? ',' : ''}${index}`)
    await sink.write(']}')
  }, { format: 'json', onProgress: (event) => progress.push(event.phase) })
  assert.equal(JSON.parse(fs.readFileSync(target, 'utf8')).rows.length, 1000)
  assert.equal(result.stamp.sha256, core.hashBytes(fs.readFileSync(target)))
  assert.deepEqual([...new Set(progress)], ['preparing', 'writing', 'verifying', 'replacing', 'done'])
  await rejectsWith(safeWriteFile(target, async (sink) => { await sink.write('{"rows":[1,2') }, { format: 'json' }), 'VALIDATION_FAILED')
  assert.equal(JSON.parse(fs.readFileSync(target, 'utf8')).rows.length, 1000)
  assert.deepEqual(leftovers(folder), [])
})

test('canceling before the replace leaves no file and no temp', async () => {
  const folder = scenario('cancel')
  const target = original(folder, 'cancel.png')
  const controller = new AbortController()
  const error = await rejectsWith(safeWriteFile(target, async (sink) => {
    await sink.write('part one')
    controller.abort()
    await sink.write('part two')
  }, { signal: controller.signal }), 'CANCELED')
  assert.equal(error.message, 'Saving "cancel.png" was canceled. Nothing was changed.')
  assert.equal(fs.readFileSync(target, 'utf8'), 'ORIGINAL')
  assert.deepEqual(leftovers(folder), [])
  journalIsClean()
})

test('writes to one file run in order, and quitting can wait for writes in flight', async () => {
  const folder = scenario('pending')
  const target = original(folder, 'queue.txt')
  const slow = withInjectedFs({
    async rename(from, to) {
      await core.sleep(150)
      return fsp.rename(from, to)
    },
  })
  const first = safeWriteFile(target, 'FIRST', { fs: slow })
  const second = safeWriteFile(target, 'SECOND', { fs: slow })
  assert.equal(core.pendingWriteCount(), 2)
  assert.equal(core.listPendingWrites()[0].label, target)
  const waited = await core.waitForPendingWrites({ timeoutMs: 10_000 })
  assert.deepEqual(waited, { settled: true, remaining: 0 })
  assert.equal((await first).strategy, 'rename')
  assert.equal((await second).strategy, 'rename')
  assert.equal(fs.readFileSync(target, 'utf8'), 'SECOND')
  assert.equal(core.pendingWriteCount(), 0)
  const timedOut = safeWriteFile(target, 'THIRD', { fs: slow })
  assert.deepEqual(await core.waitForPendingWrites({ timeoutMs: 10 }), { settled: false, remaining: 1 })
  await timedOut
})

test('safeWriteResult returns an IPC-ready result instead of throwing', async () => {
  const folder = scenario('result')
  const target = original(folder, 'result.txt')
  const ok = await safeWriteResult(target, 'NEW')
  assert.equal(ok.ok, true)
  assert.equal(ok.name, 'result.txt')
  fs.chmodSync(target, 0o444)
  try {
    const failed = await safeWriteResult(target, 'NEWER')
    assert.deepEqual(Object.keys(failed).sort(), ['code', 'message', 'name', 'ok', 'path', 'technical'])
    assert.equal(failed.technical, 'the file has the read-only attribute')
    assert.equal(failed.code, 'READ_ONLY')
    assert.equal(failed.message, '"result.txt" is read-only.')
    assert.doesNotThrow(() => JSON.stringify(failed))
  } finally {
    fs.chmodSync(target, 0o666)
  }
})

// ---------------------------------------------------------------------------
// io-core helpers
// ---------------------------------------------------------------------------

test('toIoError maps errno codes by stage', () => {
  const at = (code, stage, context) => core.toIoError(Object.assign(new Error(code), { code, syscall: 'open' }), { path: 'C:\\x\\y.docx', stage, context }).code
  assert.equal(at('ENOSPC', 'temp'), 'DISK_FULL')
  assert.equal(at('ENAMETOOLONG', 'temp'), 'NAME_TOO_LONG')
  assert.equal(at('EROFS', 'temp'), 'READ_ONLY_VOLUME')
  assert.equal(at('EPERM', 'temp'), 'NO_PERMISSION')
  assert.equal(at('EACCES', 'preflight'), 'NO_PERMISSION')
  assert.equal(at('EPERM', 'replace'), 'LOCKED')
  assert.equal(at('EBUSY', 'replace'), 'LOCKED')
  assert.equal(at('ENOENT', 'temp'), 'FOLDER_MISSING')
  // While the temp is read back, a temp a scanner took away or kept locked is not a folder problem.
  for (const code of ['ENOENT', 'EPERM', 'EACCES', 'EBUSY']) assert.equal(at(code, 'verify'), 'VERIFY_FAILED', code)
  assert.ok(core.FALLBACK_CODES.has('DISK_FULL'), 'Save As after DISK_FULL starts in another folder')
  for (const code of ['EIO', 'ETIMEDOUT', 'UNKNOWN', 'ENXIO', 'ECONNRESET']) assert.equal(at(code, 'replace'), 'FILE_UNAVAILABLE')
  assert.equal(at('ENOENT', 'read'), 'NOT_FOUND')
  assert.equal(at('EBUSY', 'read'), 'LOCKED')
  assert.equal(at('EPERM', 'read'), 'NO_PERMISSION')
  assert.equal(core.toIoError(new Error('serializer broke')).code, 'UNKNOWN')
  assert.match(core.toIoError(new Error('serializer broke')).technical, /serializer broke/)
  const busy = core.toIoError(Object.assign(new Error('busy'), { code: 'EBUSY', syscall: 'rename' }), { path: 'C:\\x\\y.docx', attempts: 8 })
  assert.equal(busy.technical, 'EBUSY rename after 8 attempts')

  // Coded errors from other shared modules keep their code and prompt details.
  const engine = core.toIoError(Object.assign(new Error('Saved as a modern copy instead.'), { code: 'NEEDS_OFFICE_ENGINE', formatLabel: 'Word 97–2003', altFormat: 'docx', altExt: '.docx' }), { path: 'C:\\x\\Old.doc' })
  assert.equal(engine.code, 'NEEDS_OFFICE_ENGINE')
  assert.equal(engine.message, 'Simple can\'t save Word 97–2003 files on this PC.')
  const result = core.toIoResult(engine)
  assert.equal(result.altFormat, 'docx')
  assert.equal(core.describeIoError(result).buttons[0].label, 'Save as .docx')
  const serialize = core.toIoError(Object.assign(new Error('The chart could not be written.'), { code: 'SERIALIZE_FAILED', formatLabel: 'Excel workbook' }))
  assert.equal(serialize.reason, 'The chart could not be written.')
  assert.match(core.describeIoError(serialize).detail, /^The chart could not be written\. Your changes are still open\./)
  const tooLarge = core.toIoResult(new core.IoError('TOO_LARGE', { context: 'open', path: 'C:\\x\\huge.pdf', limit: 2 * 1024 ** 3 }))
  assert.equal(tooLarge.limit, '2 GB')
  assert.equal(core.describeIoError(tooLarge).message, '"huge.pdf" is larger than 2 GB, the most Simple can open.')
  assert.equal(at('UNKNOWN', 'replace'), 'FILE_UNAVAILABLE', 'the libuv "UNKNOWN" errno is a storage problem, not a code')
})

test('stamps, reads and folder probes', async () => {
  const folder = scenario('core-helpers')
  const target = original(folder, 'read.txt', 'stamped content')
  const { bytes, stamp } = await core.readFileStamped(target)
  assert.equal(bytes.toString(), 'stamped content')
  assert.deepEqual(stamp, await core.stampFile(target))
  assert.equal((await core.compareStamp(target, stamp)).state, 'same')
  fs.writeFileSync(target, 'different content')
  assert.equal((await core.compareStamp(target, stamp)).state, 'changed')
  fs.unlinkSync(target)
  assert.equal((await core.compareStamp(target, stamp)).state, 'missing')

  await assert.rejects(core.readFileStamped(target), (error) => error.code === 'NOT_FOUND' && error.context === 'open')
  await assert.rejects(core.readFileStamped(path.join(folder, 'gone', 'x.txt')), (error) => error.code === 'FOLDER_MISSING' && error.nearestFolder === folder)
  original(folder, 'big.txt', 'x'.repeat(100))
  await assert.rejects(core.readFileStamped(path.join(folder, 'big.txt'), { maxBytes: 10 }), (error) => error.code === 'TOO_LARGE')

  assert.deepEqual(await core.probeFolder(folder), { folder, exists: true, writable: true })
  assert.deepEqual(leftovers(folder), [], 'the probe file is removed')
  const missing = await core.probeFolder(path.join(folder, 'nope', 'deeper'))
  assert.equal(missing.code, 'FOLDER_MISSING')
  assert.equal(missing.nearestFolder, folder)
  const space = await core.freeSpace(folder)
  assert.ok(space === null || space.free > 0)
  assert.equal(core.formatBytes(1536), '1.5 KB')
  assert.equal(core.formatBytes(5 * 1024 * 1024), '5 MB')
  if (isWindows) assert.equal(core.describeDrive('c:\\Users\\x.docx'), 'drive C:')
  if (isWindows) assert.equal(core.describeDrive('\\\\server\\share\\folder\\x.docx'), '\\\\server\\share')
})

// ---------------------------------------------------------------------------
// The catalog
// ---------------------------------------------------------------------------

const SAVE_CODES = ['LOCKED', 'READ_ONLY', 'NO_PERMISSION', 'READ_ONLY_VOLUME', 'FOLDER_MISSING', 'FILE_UNAVAILABLE', 'DISK_FULL',
  'NAME_TOO_LONG', 'INVALID_NAME', 'CHANGED_ON_DISK', 'SOURCE_MISSING', 'VERIFY_FAILED', 'VALIDATION_FAILED', 'SERIALIZE_FAILED',
  'NEEDS_OFFICE_ENGINE', 'RESTORE_NEEDED', 'UNKNOWN']
const OPEN_CODES = ['DAMAGED', 'ENCRYPTED', 'TOO_LARGE', 'UNSUPPORTED', 'NEEDS_OFFICE_ENGINE', 'NOT_FOUND', 'FOLDER_MISSING', 'LOCKED',
  'FILE_UNAVAILABLE', 'NO_PERMISSION', 'NAME_TOO_LONG', 'UNKNOWN']

function catalogStrings(node, trail = []) {
  if (typeof node === 'string') return [[trail.join('.'), node]]
  if (!node || typeof node !== 'object') return []
  return Object.entries(node).flatMap(([key, value]) => catalogStrings(value, [...trail, key]))
}

test('the catalog words every failure, prompt buttons are consistent, and the renderer copy is identical', () => {
  const electronCopy = fs.readFileSync(path.join(SHARED, 'electron', 'io-catalog.json'))
  assert.deepEqual(fs.readFileSync(path.join(SHARED, 'renderer', 'io-catalog.json')), electronCopy)
  const catalog = JSON.parse(electronCopy.toString('utf8'))
  for (const code of [...SAVE_CODES, ...OPEN_CODES]) assert.ok(core.IO_CODES[code], `${code} must be an IO code`)
  for (const code of SAVE_CODES) {
    const entry = catalog.saveFailed[code]
    assert.ok(entry && entry.message && typeof entry.detail === 'string', `saveFailed.${code} needs a message and detail`)
    const ids = entry.buttons.map((button) => button.id)
    assert.ok(ids.includes(entry.defaultId) && ids.includes(entry.cancelId), `saveFailed.${code} default and cancel must be buttons`)
    assert.equal(ids.at(-1), 'cancel')
  }
  for (const code of OPEN_CODES) assert.ok(catalog.openFailed[code]?.message, `openFailed.${code} needs a message`)
  for (const [key, prompt] of Object.entries(catalog.prompts)) {
    const ids = prompt.buttons.map((button) => button.id)
    assert.ok(ids.includes(prompt.defaultId) && ids.includes(prompt.cancelId), `prompts.${key}`)
  }
  for (const code of core.FALLBACK_CODES) assert.ok(catalog.saveFailed[code].buttons.some((button) => button.id === 'save-as'), `${code} must offer Save As`)
})

test('user-facing text never asks to install or download anything and names no cloud service', () => {
  const catalog = JSON.parse(fs.readFileSync(path.join(SHARED, 'electron', 'io-catalog.json'), 'utf8'))
  for (const [key, text] of catalogStrings(catalog)) {
    if (key === 'about') continue
    assert.doesNotMatch(text, /\b(install|download|update now|sign in)\w*/i, `${key}: ${text}`)
    assert.deepEqual(findServiceNames(text), [], `${key}: ${text}`)
    assert.doesNotMatch(text, /\bcloud\b/i, `${key}: ${text}`)
  }
  for (const file of ['io-core.cjs', 'safe-write.cjs', 'validators.cjs']) {
    const source = fs.readFileSync(path.join(SHARED, 'electron', file), 'utf8')
    assert.deepEqual(findServiceNames(source), [], file)
    assert.doesNotMatch(source, /\bcloud\b/i, file)
    assert.doesNotMatch(source, /Install LibreOffice|https?:\/\//i, file)
    assert.doesNotMatch(source, /require\((?!['"](?:node:|\.\/|electron['"]))/, `${file} may require only Node built-ins, electron and its siblings`)
  }
})

test('the manifest vendors every file safe-write needs into each workspace that receives it', () => {
  const manifest = JSON.parse(fs.readFileSync(path.join(SHARED, 'manifest.json'), 'utf8'))
  const requires = {
    'electron/safe-write.cjs': ['electron/io-core.cjs', 'electron/validators.cjs'],
    'electron/io-core.cjs': ['electron/io-catalog.json'],
    'electron/validators.cjs': ['electron/formats.json'],
  }
  for (const [file, needs] of Object.entries(requires)) {
    const source = fs.readFileSync(path.join(SHARED, file), 'utf8')
    const siblings = [...source.matchAll(/require\('\.\/([^']+)'\)/g)].map((match) => `electron/${match[1]}`)
    assert.deepEqual(siblings.sort(), [...needs].sort(), `${file} requires exactly these siblings`)
  }
  for (const [name, workspace] of Object.entries(manifest.workspaces)) {
    for (const file of Object.keys(requires)) assert.ok(workspace.electron.includes(file), `${name} must receive ${file}`)
    for (const [file, needs] of Object.entries(requires)) {
      for (const need of needs) assert.ok(workspace.electron.includes(need), `${name} vendors ${file} but not ${need}`)
    }
    if (workspace.renderer.length) assert.ok(workspace.renderer.includes('renderer/io-catalog.json'), `${name} renderer needs the catalog`)
  }
  assert.ok(manifest.sameBytes.some(([left, right]) => left === 'electron/io-catalog.json' && right === 'renderer/io-catalog.json'))
})

test('messages and prompts fill their placeholders, with neutral wording when a value is missing', () => {
  const locked = new core.IoError('LOCKED', { path: 'C:\\Work\\Budget.xlsx', technical: 'EBUSY open after 9 attempts' })
  assert.equal(locked.message, '"Budget.xlsx" is open in another program.')
  const prompt = core.describeIoError(locked)
  assert.deepEqual(prompt.buttons.map((button) => button.label), ['Try Again', 'Save As…', 'Cancel'])
  assert.equal(prompt.defaultId, 'retry')
  assert.match(prompt.detail, /Details: EBUSY open after 9 attempts$/)
  assert.equal(new core.IoError('LOCKED').message, '"this file" is open in another program.')
  const engine = core.formatCatalogPrompt('saveFailed.NEEDS_OFFICE_ENGINE', { formatLabel: 'Word 97–2003', altLabel: 'Word document (.docx)', altExt: '.docx', name: 'Old.doc' })
  assert.equal(engine.buttons[0].label, 'Save as .docx')
  assert.doesNotMatch(`${engine.message} ${engine.detail}`, /\{\w+\}/)
  assert.equal(core.formatTemplate('{appName} stopped', {}), 'Simple stopped')
  assert.equal(core.catalogEntry('status.saving'), 'Saving…')
})
