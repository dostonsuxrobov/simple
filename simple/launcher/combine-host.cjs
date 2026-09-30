'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const crypto = require('node:crypto')
const { Worker } = require('node:worker_threads')

async function atomicWrite(target, bytes) {
  const temporary = path.join(path.dirname(target), `.${path.basename(target)}.${crypto.randomUUID()}.tmp`)
  try {
    await fs.writeFile(temporary, bytes, { flag: 'wx' })
    await fs.rename(temporary, target)
  } finally { await fs.unlink(temporary).catch(() => {}) }
}

function runCombine(entries, onProgress) {
  const workerPath = path.join(__dirname, '..', 'modules', 'shared', 'combine-worker.cjs').replace(/app\.asar([\\/])/, 'app.asar.unpacked$1')
  const worker = new Worker(workerPath, { workerData: entries })
  return new Promise((resolve, reject) => {
    let settled = false
    const finish = (error, result) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      void worker.terminate()
      if (error) reject(error)
      else resolve(result)
    }
    const timer = setTimeout(() => finish(new Error('Combining took too long. Try fewer files. Your source files are unchanged.')), 10 * 60_000)
    worker.on('message', (message) => {
      if (message.progress) onProgress(message.progress)
      else if (message.error) finish(new Error(message.error))
      else if (message.result) finish(null, message.result)
    })
    worker.once('error', (error) => finish(error))
    worker.once('exit', (code) => { if (!settled) finish(new Error(`Combining stopped unexpectedly (${code}). Try again.`)) })
  })
}

module.exports = { atomicWrite, runCombine }
