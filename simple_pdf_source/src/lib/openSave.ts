import type { OpenDocument, PdfSaveProblem } from '../types'

/**
 * True when the file names an encryption dictionary anywhere. Linearized
 * ("Fast Web View") files keep /Encrypt in the first-page trailer near the
 * start, so the whole file is searched, as the main process does. Used only
 * for bytes the renderer opened itself; files opened from disk carry the
 * main process's answer in `payload.encrypted`. A false positive only costs
 * one unlock call, which then reports that nothing is encrypted.
 */
export function mayBeEncrypted(data: Uint8Array) {
  const chunk = 8 * 1024 * 1024
  const overlap = 16
  const decoder = new TextDecoder('latin1')
  for (let start = 0; start < data.length; start += chunk) {
    const text = decoder.decode(data.subarray(Math.max(0, start - overlap), Math.min(data.length, start + chunk)))
    let index = text.indexOf('/Encrypt')
    while (index >= 0) {
      // `/EncryptMetadata` and other longer names are different keys.
      if (!/[A-Za-z0-9_.-]/.test(text.charAt(index + 8))) return true
      index = text.indexOf('/Encrypt', index + 8)
    }
  }
  return false
}

/** "C:\\x\\contract.pdf" → "C:\\x\\contract (edited).pdf": the copy a signed original is saved as. */
export function signedCopyName(file: Pick<OpenDocument, 'path' | 'name'>) {
  const source = file.path || file.name
  return `${source.replace(/\.pdf$/i, '')} (edited).pdf`
}

/** Problems that mean a save would not write exactly what the user made. */
export function blockingProblems(output: { warnings: PdfSaveProblem[]; failures: PdfSaveProblem[] }) {
  return [...output.failures, ...output.warnings.filter((warning) => warning.dataLoss)]
}

export function problemText(problem: PdfSaveProblem) {
  return problem.blockingMessage || problem.message
}

/** One line for a toast: the first message, and how many more there are. */
export function summarizeProblems(problems: PdfSaveProblem[]) {
  const messages = [...new Set(problems.map(problemText))]
  return messages.length > 1 ? `${messages[0]} (+${messages.length - 1} more)` : messages[0] || ''
}
