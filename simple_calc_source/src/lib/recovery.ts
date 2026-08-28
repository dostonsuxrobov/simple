import type { SaveResult } from '../spreadsheet-types'

type RecoverySave = () => Promise<SaveResult | null>

let recoverySave: RecoverySave | null = null

export function registerRecoverySave(handler: RecoverySave) {
  recoverySave = handler
}

export function runRecoverySave(): Promise<SaveResult | null> {
  if (!recoverySave) return Promise.reject(new Error('No workbook is open to recover.'))
  return recoverySave()
}
