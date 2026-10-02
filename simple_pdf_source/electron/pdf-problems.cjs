'use strict'

/**
 * Problems found while writing a PDF. A save never drops an edit silently:
 * `failures` are edits that could not be written at all, `warnings` are edits
 * that were written but changed (for example shrunk to fit their box). A
 * warning with `dataLoss: true` (a truncated form value) is treated like a
 * failure by callers that cannot show warnings to the user.
 */
function problem(code, message, details = {}) {
  const entry = { code, message }
  for (const [key, value] of Object.entries(details)) {
    if (value !== undefined) entry[key] = value
  }
  return entry
}

/**
 * Electron passes only `String(error)` ("<name>: <message>") across IPC, so
 * the machine-readable code travels as the error name. The renderer can read it
 * from "Error invoking remote method '…': CODE: message".
 */
function codedError(code, message, details) {
  const error = new Error(message)
  error.name = code
  error.code = code
  if (details !== undefined) error.details = details
  return error
}

function summarizeProblems(problems, limit = 3) {
  // `blockingMessage` words a warning for callers that abort instead of
  // saving a changed value (for example a value that would be truncated).
  const messages = [...new Set(problems.map((entry) => entry.blockingMessage || entry.message).filter(Boolean))]
  const shown = messages.slice(0, limit).join(' ')
  const hidden = messages.length - Math.min(limit, messages.length)
  return hidden > 0 ? `${shown} (+${hidden} more)` : shown
}

/** The single error thrown when a caller cannot receive a problem report. */
function unsavedChangesError(problems) {
  const codes = new Set(problems.map((entry) => entry.code))
  const code = codes.size === 1 ? [...codes][0] : 'SAVE_INCOMPLETE'
  return codedError(code, `Nothing was saved: ${summarizeProblems(problems)}`, problems)
}

module.exports = { codedError, problem, summarizeProblems, unsavedChangesError }
