export type CellScalar = string | number | boolean | null

export interface SpreadsheetColor {
  argb?: string
  rgb?: string
  indexed?: number
  theme?: number
  tint?: number
  auto?: boolean
}

export interface CellFont {
  name?: string
  family?: number | string
  scheme?: string
  charset?: number | string
  size?: number
  bold?: boolean
  italic?: boolean
  underline?: boolean | string
  strike?: boolean
  outline?: boolean
  shadow?: boolean
  vertAlign?: 'superscript' | 'subscript' | string
  condense?: boolean
  extend?: boolean
  color?: string | SpreadsheetColor
}

export interface CellFillStop {
  position?: number
  color?: string | SpreadsheetColor
}

export interface CellFill {
  type?: string
  pattern?: string
  color?: string
  fgColor?: string | SpreadsheetColor
  bgColor?: string | SpreadsheetColor
  degree?: number
  center?: { left?: number; top?: number; right?: number; bottom?: number; [key: string]: unknown }
  stops?: CellFillStop[]
}

export interface CellBorderSide {
  style?: string
  color?: string | SpreadsheetColor
}

export interface CellBorder {
  top?: CellBorderSide
  left?: CellBorderSide
  bottom?: CellBorderSide
  right?: CellBorderSide
  diagonal?: CellBorderSide
  vertical?: CellBorderSide
  horizontal?: CellBorderSide
  diagonalUp?: boolean
  diagonalDown?: boolean
  outline?: boolean
  [key: string]: CellBorderSide | boolean | undefined
}

export interface CellAlignment {
  horizontal?: string
  vertical?: string
  wrapText?: boolean
  clipText?: boolean
  shrinkToFit?: boolean
  indent?: number
  relativeIndent?: number
  textRotation?: number | 'vertical'
  readingOrder?: number | 'rtl' | 'ltr'
  justifyLastLine?: boolean
}

export interface CellStyle {
  font?: CellFont
  fill?: CellFill
  border?: CellBorder
  alignment?: CellAlignment
  numFmt?: string
  protection?: { locked?: boolean; hidden?: boolean }
  quotePrefix?: number
}

export interface CellNoteComment {
  author?: string
  text?: string
  hidden?: boolean
}

export interface CellNote {
  comments?: CellNoteComment[]
  texts?: Array<{ text?: string; font?: CellFont }>
  text?: string
  hidden?: boolean
}

export interface CellData {
  value?: CellScalar
  formula?: string
  formulaType?: 'array' | 'shared' | string
  formulaRange?: string
  dynamicFormula?: boolean
  result?: CellScalar
  resultType?: 'date' | 'error' | string
  display?: string
  style?: CellStyle
  numFmt?: string
  hyperlink?: string
  hyperlinkTooltip?: string
  note?: string | CellNote
  type?: string
}

export interface SheetData {
  id: string
  name: string
  sourceWorksheetId?: number
  sourceSheetName?: string
  sourceSheetIndex?: number
  state?: 'visible' | 'hidden' | 'veryHidden'
  rowCount: number
  colCount: number
  cells: Record<string, CellData>
  merges: string[]
  colWidths: Record<string, number>
  rowHeights: Record<string, number>
  hiddenRows?: number[]
  hiddenCols?: number[]
  frozen?: { rows?: number; columns?: number; topLeftCell?: string; activeCell?: string }
  views?: Array<Record<string, unknown>>
  properties?: Record<string, unknown>
  pageSetup?: Record<string, unknown>
  headerFooter?: Record<string, unknown>
  rowBreaks?: unknown[]
  rowProperties?: Record<string, Record<string, unknown>>
  columnProperties?: Record<string, Record<string, unknown>>
  sheetProtection?: Record<string, unknown> | null
  dataValidations?: Record<string, unknown>
  dataValidationsTruncated?: boolean
  conditionalFormattings?: unknown[]
  conditionalFormattingsTruncated?: boolean
  requiresSourcePackage?: string[]
  autoFilter?: unknown
}

export interface DefinedName {
  name: string
  ranges: string[]
  ref?: string
  localSheetIndex?: number
  localSheetId?: number
  hidden?: boolean
  comment?: string
  attributes?: Record<string, unknown>
}

export interface WorkbookModel {
  version: 1
  name: string
  activeSheetId: string
  sheets: SheetData[]
  definedNames?: DefinedName[]
  metadata?: {
    creator?: string
    created?: string
    modified?: string
    date1904?: boolean
    sourceDate1904?: boolean
    calcProperties?: Record<string, unknown>
    workbookProperties?: Record<string, unknown>
    workbookViews?: Array<Record<string, unknown>>
    definedNames?: Array<{ name: string; ranges?: string; formula?: string; localSheetId?: number }>
    [key: string]: unknown
  }
}

export interface WorkbookPayload {
  documentId: string
  path: string | null
  name: string
  sourceFormat: string
  requiresSaveAs: boolean
  warnings: string[]
  stats?: { sheets?: number; cells?: number; formulas?: number; [key: string]: number | undefined }
  workbook: WorkbookModel
}

export interface SaveResult {
  path: string
  name: string
  format: string
}

export interface RecentWorkbook {
  path: string
  name: string
  format: string
  openedAt: number
}

export interface SimpleCalcAPI {
  createWorkbook: () => Promise<{ documentId: string }>
  openWorkbook: () => Promise<WorkbookPayload | null>
  openPath: (filePath: string) => Promise<WorkbookPayload>
  openBytes: (name: string, data: ArrayBuffer) => Promise<WorkbookPayload>
  openInNewWindow: (filePath?: string) => Promise<boolean>
  saveWorkbook: (input: {
    documentId: string
    workbook: WorkbookModel
    saveAs: boolean
    format: 'xlsx' | 'ods' | 'csv' | 'tsv'
    suggestedName: string
    sourceUnmodified?: boolean
  }) => Promise<SaveResult | null>
  showItem: (filePath: string) => Promise<void>
  openExternal: (target: string) => Promise<void>
  getVersion: () => Promise<string>
  minimize: () => void
  toggleMaximize: () => void
  close: () => void
  onMaximized: (callback: (maximized: boolean) => void) => () => void
  onOpenExternal: (callback: (filePath: string) => void) => () => void
  onCloseRequested: (callback: () => void) => () => void
}

declare global {
  interface Window {
    simpleCalc: SimpleCalcAPI
  }
}
