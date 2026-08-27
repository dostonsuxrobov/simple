import { SSF } from 'xlsx'
import type { CellScalar } from '../spreadsheet-types'

const CURRENCY_SYMBOL = /[$€£¥₹₩₽]/

export function formatScalar(value: CellScalar | undefined, numFmt?: string, fallbackDisplay?: string) {
  if (value === undefined || value === null) return ''
  if (typeof value !== 'number') return fallbackDisplay ?? String(value)

  const format = String(numFmt || '').trim()
  if (!format) return fallbackDisplay ?? String(value)
  try {
    const formatted = SSF.format(format, value)
    if (typeof formatted === 'string') return formatted
  } catch {
    // Keep the workbook usable when a vendor-specific format is not supported.
  }
  return fallbackDisplay ?? String(value)
}

export function accountingCurrencySymbol(numFmt?: string) {
  const format = String(numFmt || '')
  if (!isAccountingNumberFormat(format)) return ''
  return format.match(CURRENCY_SYMBOL)?.[0] || ''
}

export function isAccountingNumberFormat(numFmt?: string) {
  return String(numFmt || '').includes('*')
}

export function accountingDisplayParts(display: string, numFmt?: string) {
  const symbol = accountingCurrencySymbol(numFmt)
  if (!symbol) return null
  const text = String(display || '').trim()
  const symbolIndex = text.indexOf(symbol)
  if (symbolIndex < 0) return null
  const amount = `${text.slice(0, symbolIndex)}${text.slice(symbolIndex + symbol.length)}`.trim()
  return { symbol, amount }
}
