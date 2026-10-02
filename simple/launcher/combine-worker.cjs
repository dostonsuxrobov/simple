'use strict'

// Runs Combine off the main thread. Printing Simple's own layouts needs a
// browser window, and unlocking a PDF protected by a permissions password
// needs MuPDF, which the launcher main process loads from the PDF workspace.
// So the worker asks the main process to print each HTML job ({print}) or
// unlock a PDF ({unlock}) and waits for the bytes ({printed}, {unlocked}).
const { parentPort, workerData } = require('node:worker_threads')
const { combineFiles } = require('./combine-service.cjs')
const { serializeError } = require('./combine-policy.cjs')

const pending = new Map()
let nextRequest = 0

parentPort.on('message', (message) => {
  const reply = message && (message.printed || message.unlocked)
  if (!reply || !pending.has(reply.id)) return
  const { resolve, reject, fallback } = pending.get(reply.id)
  pending.delete(reply.id)
  if (reply.error) reject(Object.assign(new Error(reply.error.message || fallback.message), { code: reply.error.code || fallback.code }))
  else resolve(reply.bytes)
})

function request(kind, payload, fallback) {
  return new Promise((resolve, reject) => {
    nextRequest += 1
    pending.set(nextRequest, { resolve, reject, fallback })
    parentPort.postMessage({ [kind]: { id: nextRequest, ...payload } })
  })
}

function printHtml(html, options) {
  return request('print', { html, options }, { message: 'Simple could not make PDF pages from this file.', code: 'PRINT_FAILED' })
}

function unlockPdf(bytes) {
  return request('unlock', { bytes }, { message: "Simple can't read this protected PDF.", code: 'ENCRYPTED' })
}

// workerData is {entries, title}; a plain entries array is still accepted.
const job = Array.isArray(workerData) ? { entries: workerData } : (workerData || {})

combineFiles(job.entries, (progress) => parentPort.postMessage({ progress }), { printHtml, unlockPdf, title: job.title }).then(
  (result) => parentPort.postMessage({ result }),
  (error) => parentPort.postMessage({ error: serializeError(error) }),
)
