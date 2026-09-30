/**
 * Excel sheet protection. `sheet.sheetProtection` holds the OOXML/ExcelJS model: `sheet: true`
 * turns protection on; permission flags are `true` when the action is allowed (formatCells,
 * insertRows, sort, autoFilter, …); `objects === false` protects charts and pictures. Cells are
 * locked unless their style says `protection.locked === false`.
 */
import type { CellData, SheetData, WorkbookModel } from '../spreadsheet-types'

export type ProtectedAction =
  | 'formatCells' | 'formatColumns' | 'formatRows'
  | 'insertColumns' | 'insertRows' | 'insertHyperlinks'
  | 'deleteColumns' | 'deleteRows'
  | 'sort' | 'autoFilter' | 'pivotTables' | 'objects'

export const PROTECTION_OPTIONS: Array<{ id: ProtectedAction | 'selectLockedCells' | 'selectUnlockedCells'; label: string; defaultOn: boolean }> = [
  { id: 'selectLockedCells', label: 'Select locked cells', defaultOn: true },
  { id: 'selectUnlockedCells', label: 'Select unlocked cells', defaultOn: true },
  { id: 'formatCells', label: 'Format cells', defaultOn: false },
  { id: 'formatColumns', label: 'Format columns', defaultOn: false },
  { id: 'formatRows', label: 'Format rows', defaultOn: false },
  { id: 'insertColumns', label: 'Insert columns', defaultOn: false },
  { id: 'insertRows', label: 'Insert rows', defaultOn: false },
  { id: 'insertHyperlinks', label: 'Insert hyperlinks', defaultOn: false },
  { id: 'deleteColumns', label: 'Delete columns', defaultOn: false },
  { id: 'deleteRows', label: 'Delete rows', defaultOn: false },
  { id: 'sort', label: 'Sort', defaultOn: false },
  { id: 'autoFilter', label: 'Use AutoFilter', defaultOn: false },
  { id: 'pivotTables', label: 'Use PivotTable reports', defaultOn: false },
  { id: 'objects', label: 'Edit objects', defaultOn: false },
]

export const PROTECTED_MESSAGE = "The cell or chart you're trying to change is on a protected sheet. To make a change, unprotect the sheet (Tools › Unprotect sheet)."

type Protection = Record<string, unknown>

export function sheetProtection(sheet: Pick<SheetData, 'sheetProtection'> | null | undefined): Protection | null {
  const protection = sheet?.sheetProtection as Protection | null | undefined
  return protection && typeof protection === 'object' && protection.sheet ? protection : null
}

export function protectionAllows(protection: Protection | null, action: ProtectedAction): boolean {
  if (!protection) return true
  if (action === 'objects') return protection.objects !== false
  return protection[action] === true
}

function addressCoord(address: string) {
  const match = /^([A-Z]+)(\d+)$/.exec(address)
  if (!match) return null
  let col = 0
  for (const character of match[1]) col = col * 26 + character.charCodeAt(0) - 64
  return { row: Number(match[2]) - 1, col: col - 1 }
}

/** Whether a cell is locked (the default), honouring row and column default styles. */
export function cellLocked(sheet: SheetData, address: string): boolean {
  const own = sheet.cells[address]?.style?.protection as { locked?: boolean } | undefined
  if (own && own.locked !== undefined) return own.locked !== false
  const coord = addressCoord(address)
  if (coord) {
    const row = (sheet.rowProperties?.[String(coord.row + 1)]?.style as { protection?: { locked?: boolean } } | undefined)?.protection
    if (row && row.locked !== undefined) return row.locked !== false
    const column = (sheet.columnProperties?.[String(coord.col + 1)]?.style as { protection?: { locked?: boolean } } | undefined)?.protection
    if (column && column.locked !== undefined) return column.locked !== false
  }
  return true
}

const CONTENT_KEYS = ['value', 'formula', 'richText', 'type', 'arrayMember', 'formulaType', 'formulaRange', 'dynamicFormula'] as const
const FORMAT_KEYS = ['style', 'numFmt'] as const
const DERIVED_KEYS = new Set(['result', 'resultType', 'display'])

function contentOf(cell: CellData | undefined) {
  return JSON.stringify(CONTENT_KEYS.map((key) => cell?.[key] ?? null))
}

function formatOf(cell: CellData | undefined) {
  return JSON.stringify(FORMAT_KEYS.map((key) => cell?.[key] ?? null))
}

interface PatchLike { path: Array<string | number> }

/**
 * Why a change (given as immer patches from `before` to `after`) is not allowed by sheet
 * protection, or null. Sheets that were not protected before the change are never limited,
 * and turning protection on or off is always allowed (the unprotect command checks the
 * password itself).
 */
export function protectionViolation(before: WorkbookModel, after: WorkbookModel, patches: PatchLike[]): string | null {
  for (const patch of patches) {
    if (patch.path[0] !== 'sheets' || patch.path.length < 3) continue
    const index = Number(patch.path[1])
    const sheet = before.sheets[index]
    const protection = sheetProtection(sheet)
    if (!protection) continue
    const next = after.sheets.find((item) => item.id === sheet.id)
    const key = String(patch.path[2])
    switch (key) {
      case 'cells': {
        const addresses = patch.path.length >= 4
          ? [String(patch.path[3])]
          : [...new Set([...Object.keys(sheet.cells), ...Object.keys(next?.cells || {})])]
        for (const address of addresses) {
          const old = sheet.cells[address]
          const updated = next?.cells[address]
          if (patch.path.length >= 5 && DERIVED_KEYS.has(String(patch.path[4]))) continue
          if (contentOf(old) !== contentOf(updated) || (old?.hyperlink ?? null) !== (updated?.hyperlink ?? null)) {
            if (cellLocked(sheet, address)) return PROTECTED_MESSAGE
            if ((old?.hyperlink ?? null) !== (updated?.hyperlink ?? null) && !protectionAllows(protection, 'insertHyperlinks')) return PROTECTED_MESSAGE
          }
          if (formatOf(old) !== formatOf(updated) && !protectionAllows(protection, 'formatCells')) return PROTECTED_MESSAGE
          if ((old?.note ? JSON.stringify(old.note) : null) !== (updated?.note ? JSON.stringify(updated.note) : null) && !protectionAllows(protection, 'objects')) return PROTECTED_MESSAGE
        }
        break
      }
      case 'colWidths':
      case 'columnProperties':
      case 'hiddenCols':
        if (!protectionAllows(protection, 'formatColumns')) return PROTECTED_MESSAGE
        break
      case 'rowHeights':
      case 'rowProperties':
        if (!protectionAllows(protection, 'formatRows')) return PROTECTED_MESSAGE
        break
      case 'hiddenRows':
      case 'filteredRows':
        if (!protectionAllows(protection, 'formatRows') && !protectionAllows(protection, 'autoFilter')) return PROTECTED_MESSAGE
        break
      case 'filter':
      case 'autoFilter':
        if (!protectionAllows(protection, 'autoFilter')) return PROTECTED_MESSAGE
        break
      case 'charts':
      case 'images':
        if (!protectionAllows(protection, 'objects')) return PROTECTED_MESSAGE
        break
      case 'pivots':
        if (!protectionAllows(protection, 'pivotTables')) return PROTECTED_MESSAGE
        break
      case 'tables':
        // A table's own filter follows the AutoFilter permission; anything else about it is locked.
        if (patch.path[4] === 'filter' && protectionAllows(protection, 'autoFilter')) break
        return PROTECTED_MESSAGE
      case 'merges':
      case 'dataValidations':
      case 'conditionalFormattings':
        return PROTECTED_MESSAGE
      default:
        break
    }
  }
  return null
}
