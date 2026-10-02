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
  /** Run formatting for a rich-text cell; `value` holds its plain text. */
  richText?: Array<{ text: string; font?: CellFont }>
  style?: CellStyle
  numFmt?: string
  hyperlink?: string
  hyperlinkTooltip?: string
  note?: string | CellNote
  type?: string
  /**
   * Address of the dynamic-array anchor whose saved spill range covers this cell. Such a
   * cell only holds the file's cached value and yields to the live spill.
   */
  arrayMember?: string
}

/** An Excel table (ListObject). `ref` includes the header and totals rows. */
export interface SheetTable {
  id: string
  name: string
  displayName?: string
  ref: string
  headerRow: boolean
  totalsRow: boolean
  columns: Array<{ name: string; totalsRowFunction?: string; totalsRowLabel?: string; totalsRowFormula?: string }>
  style?: {
    theme?: string
    showRowStripes?: boolean
    showColumnStripes?: boolean
    showFirstColumn?: boolean
    showLastColumn?: boolean
  }
  /** True when the table was read from the source package (not created in the editor). */
  imported?: boolean
  /** Header filter buttons (the table's AutoFilter); on unless explicitly turned off. */
  showFilterButton?: boolean
  /** Criteria of the table's own AutoFilter, with `ref` covering the header and data rows. */
  filter?: SheetFilterState
}

/** One column's criteria in an AutoFilter / table filter. Offsets are 0-based within `ref`. */
export interface SheetFilterCriteria {
  /** Display values that remain visible; undefined means "all values". */
  values?: string[]
  /** Whether blank cells stay visible when `values` is set. */
  blanks?: boolean
  condition?: {
    operator: string
    value?: string | number
    operator2?: string
    value2?: string | number
    join?: 'and' | 'or'
  }
  /** Keep only cells whose fill (or font) colour matches, as #RRGGBB. */
  fillColor?: string
  fontColor?: string
}

export interface SheetFilterState {
  /** Range including the header row, e.g. "A1:F200". */
  ref: string
  columns: Record<number, SheetFilterCriteria>
  sort?: { column: number; descending: boolean }
}

// ---------------------------------------------------------------------------
// Charts (owned by src/lib/charts.ts, src/lib/chart-render.ts and
// electron/chart-xlsx.cjs). Row/column indices are 0-based like the grid;
// offsets are EMU (English Metric Units, 1px = 9525 EMU at 96 DPI) as in
// DrawingML's xdr:from / xdr:to.
// ---------------------------------------------------------------------------

export interface ChartAnchorPoint {
  row: number
  col: number
  rowOffsetEmu?: number
  colOffsetEmu?: number
}

export interface ChartAnchor {
  from: ChartAnchorPoint
  to: ChartAnchorPoint
  /** Excel's editAs: 'twoCell' (default) moves and sizes with cells, 'oneCell' moves only, 'absolute' neither. */
  editAs?: 'twoCell' | 'oneCell' | 'absolute'
}

export type PivotSummarize = 'sum' | 'count' | 'countNums' | 'countDistinct' | 'average' | 'max' | 'min' | 'product' | 'median' | 'stdDev' | 'stdDevp' | 'var' | 'varp'
export type PivotShowAs = 'normal' | 'percentOfGrandTotal' | 'percentOfRowTotal' | 'percentOfColumnTotal' | 'runningTotal'
export type PivotDateGroup = 'year' | 'quarter' | 'month' | 'yearQuarter' | 'yearMonth' | 'day'

export interface PivotAxisField {
  field: string
  order?: 'asc' | 'desc'
  dateGroup?: PivotDateGroup
  /** Subtotal rows/columns after each group of this (outer) field; on by default. */
  showTotals?: boolean
}

export interface PivotValueField {
  field: string
  summarize: PivotSummarize
  showAs?: PivotShowAs
  label?: string
}

export interface PivotFilterField {
  field: string
  /** Keys (see pivotFieldKeys) that are filtered out; new values show by default, as in Excel. */
  exclude?: string[]
}

/** A pivot table whose output block starts at `anchor` on the sheet holding it. */
export interface PivotTableModel {
  id: string
  name: string
  /** Source range with its header row: `Sheet1!A1:E200` or a table name. */
  source: string
  anchor: { row: number; col: number }
  rows: PivotAxisField[]
  columns: PivotAxisField[]
  values: PivotValueField[]
  filters: PivotFilterField[]
  /** Grand-total row at the bottom (on by default). */
  showRowGrandTotal?: boolean
  /** Grand-total column at the right (on by default). */
  showColumnGrandTotal?: boolean
  /** Size of the block written last (cleared before rewriting). */
  extent?: { rows: number; cols: number }
}

/** An Excel sparkline group (x14:sparklineGroup): shared look for a set of in-cell charts. */
export interface SparklineGroup {
  type: 'line' | 'column' | 'stacked'
  colors: { series?: string; negative?: string; axis?: string; markers?: string; first?: string; last?: string; high?: string; low?: string }
  markers?: boolean
  high?: boolean
  low?: boolean
  first?: boolean
  last?: boolean
  negative?: boolean
  displayXAxis?: boolean
  displayHidden?: boolean
  rightToLeft?: boolean
  dateAxis?: boolean
  displayEmptyCellsAs?: 'gap' | 'zero' | 'span'
  minAxisType?: 'individual' | 'group' | 'custom'
  maxAxisType?: 'individual' | 'group' | 'custom'
  manualMin?: number
  manualMax?: number
  /** Line width in points. */
  lineWeight?: number
  /** One sparkline per target cell: `source` like `Sheet1!A2:E2`, `cell` like `F2`. */
  sparklines: Array<{ source: string; cell: string }>
  /** Imported XML and its parsed signature: unedited groups are saved byte-for-byte. */
  sourceXml?: string
  signature?: string
}

/** A floating picture anchored to cells (xdr:pic in the sheet drawing). */
export interface SheetImage {
  id: string
  name?: string
  altText?: string
  /** A data: URL holding the picture bytes. */
  src: string
  anchor: ChartAnchor
  /** The source package's image id, reused on save while the picture is unchanged. */
  sourceImageId?: number
}

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter' | 'radar' | 'combo' | 'unsupported'
/** Plot type of one series inside a combo chart. */
export type ChartSeriesType = 'column' | 'bar' | 'line' | 'area' | 'scatter'
export type ChartGrouping = 'clustered' | 'stacked' | 'percentStacked'
export type ChartLegendPosition = 'right' | 'bottom' | 'top' | 'left' | 'none'
export type ChartMarkerSymbol = 'auto' | 'none' | 'circle' | 'square' | 'diamond' | 'triangle' | 'x' | 'star' | 'dash' | 'dot' | 'plus'
export type ChartDataLabelPosition = 'auto' | 'outEnd' | 'inEnd' | 'center' | 'inBase' | 'above' | 'below' | 'left' | 'right' | 'bestFit'

export interface ChartDataLabels {
  showValue?: boolean
  showCategory?: boolean
  showSeriesName?: boolean
  showPercent?: boolean
  position?: ChartDataLabelPosition
  numFmt?: string
}

export interface ChartSeries {
  id: string
  /** Literal series name, or the cached text of `nameRef`. */
  name?: string
  /** Cell holding the series name, e.g. "Sheet1!$B$1". */
  nameRef?: string
  /** Category labels, e.g. "Sheet1!$A$2:$A$13" (quoted sheet names allowed: "'My Sheet'!$A$2:$A$13"). */
  categoriesRef?: string
  valuesRef?: string
  /** Scatter X values. */
  xValuesRef?: string
  /** Cached/literal points, used when a reference is absent or cannot be resolved (external or deleted sheets). */
  categoriesCache?: string[]
  valuesCache?: Array<number | null>
  xValuesCache?: Array<number | null>
  /** Source number format of the values ("linked to source"). */
  valuesNumFmt?: string
  /** Series colour as #RRGGBB; undefined uses the palette. */
  color?: string
  /** Per-point fill overrides (#RRGGBB) keyed by 0-based point index (pie slices, varied bars). */
  pointColors?: Record<string, string>
  /** Plot type override inside a combo chart. */
  type?: ChartSeriesType
  secondaryAxis?: boolean
  marker?: ChartMarkerSymbol
  markerSize?: number
  smooth?: boolean
  /** Line/scatter connector; false shows markers only. */
  showLine?: boolean
  /** Line width in points. */
  lineWidth?: number
  dataLabels?: ChartDataLabels
  invertIfNegative?: boolean
}

export interface ChartAxis {
  title?: string
  numFmt?: string
  min?: number
  max?: number
  majorUnit?: number
  gridlines?: boolean
  minorGridlines?: boolean
  /** false hides the axis (DrawingML c:delete). */
  visible?: boolean
  /** Plot the axis in reverse order (c:orientation maxMin). */
  reverse?: boolean
  logBase?: number
  /** Tick label rotation in degrees (-90..90). */
  labelRotation?: number
}

export interface ChartStyle {
  /** Series colours (#RRGGBB) in order; defaults to the workbook theme accents. */
  palette?: string[]
  /** Chart area fill (#RRGGBB) or 'transparent'. */
  background?: string
  /** Chart area outline colour, or null for none. */
  border?: string | null
  fontFamily?: string
  textColor?: string
  /** Point sizes. */
  titleFontSize?: number
  fontSize?: number
  roundedCorners?: boolean
}

/**
 * A chart anchored over the grid. For cartesian charts `axes.x` is the category
 * axis (vertical for horizontal 'bar' charts), `axes.y` the primary value axis
 * and `axes.y2` the secondary value axis. Scatter charts use `axes.x` for the X
 * value axis.
 */
export interface SheetChart {
  id: string
  type: ChartType
  grouping?: ChartGrouping
  anchor: ChartAnchor
  /** Drawing object name, e.g. "Chart 1". */
  name?: string
  /** Alternative text. */
  description?: string
  title?: string
  /** Cell supplying the title text. */
  titleRef?: string
  /** true suppresses Excel's automatic title (the single series name). */
  autoTitleDeleted?: boolean
  series: ChartSeries[]
  /** Source range as typed by the user, e.g. "Sheet1!A1:C13". */
  dataRange?: string
  seriesIn?: 'columns' | 'rows'
  firstRowHeaders?: boolean
  firstColumnLabels?: boolean
  legend?: ChartLegendPosition
  axes?: { x?: ChartAxis; y?: ChartAxis; y2?: ChartAxis }
  style?: ChartStyle
  /** Colour each point differently (single-series charts; always on for pie/doughnut). */
  varyColors?: boolean
  /** Gap between bar clusters as a percentage of the bar width (0-500). */
  gapWidth?: number
  /** Bar overlap percentage (-100..100). */
  overlap?: number
  /** Doughnut hole size percentage (10-90). */
  holeSize?: number
  firstSliceAngle?: number
  displayBlanksAs?: 'gap' | 'zero' | 'span'
  /** Skip hidden rows/columns (Excel's default: true). */
  plotVisibleOnly?: boolean
  /** Imported from a 3-D chart; drawn flat but kept 3-D when copied unmodified. */
  threeD?: boolean
  /** Original plot type of an 'unsupported' chart (e.g. "bubble", "waterfall"). */
  unsupportedKind?: string
  /** Package part of the imported chart (e.g. "xl/charts/chart1.xml"); unmodified charts are copied byte-for-byte on save. */
  sourcePart?: string
  /** Import bookkeeping for the byte-copy path. */
  sourceInfo?: { drawingPart: string; anchorIndex: number; fingerprint: string; kind: 'chart' | 'chartex' }
  /** Set by editors on any user edit other than moving/resizing; forces the chart XML to be regenerated. */
  modified?: boolean
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
  /** Live filter criteria for the sheet's AutoFilter range (rows it hides are in hiddenRows / filteredRows). */
  filter?: SheetFilterState
  /** 1-based rows hidden by the active filter (a subset of hiddenRows). */
  filteredRows?: number[]
  tables?: SheetTable[]
  /** Floating pictures; undefined when the file's pictures were not loaded into the editor. */
  images?: SheetImage[]
  /** Pivot tables whose output lives on this sheet. */
  pivots?: PivotTableModel[]
  /** Excel sparkline groups drawn in this sheet's cells. */
  sparklineGroups?: SparklineGroup[]
  charts?: SheetChart[]
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
    /** The workbook's own clrScheme in OOXML theme-index order (lt1, dk1, lt2, dk2, accent1-6, hlink, folHlink). */
    themeColors?: string[]
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
  backupPath?: string
}

export type SpreadsheetExportFormat = 'xlsx' | 'xls' | 'ods' | 'csv' | 'tsv' | 'pdf' | 'html'

export interface SpreadsheetExportResult extends SaveResult {
  sheets?: number
  cells?: number
}

export type PrintScope = 'active-sheet' | 'selection' | 'workbook'
export type PrintOrientation = 'portrait' | 'landscape'
export type PrintScaling = 'actual' | 'fit-width' | 'fit-sheet'
export type PrintPaperSize = 'letter' | 'a4' | 'legal'
export type PrintMargins = 'normal' | 'narrow' | 'wide'

export interface SpreadsheetPrintOptions {
  useSavedLayout?: boolean
  scope: PrintScope
  orientation: PrintOrientation
  scaling: PrintScaling
  paperSize: PrintPaperSize
  margins: PrintMargins
  gridlines: boolean
  headings: boolean
}

export type SpreadsheetDisplayParts = Record<string, Record<string, {
  type: 'number' | 'boolean' | 'text'
  accounting?: { symbol: string; amount: string }
}>>

export interface SpreadsheetPrintPreview {
  options?: SpreadsheetPrintOptions
  warnings?: string[]
  html: string
  title: string
  sheets: number
  cells: number
  pages: number
  pageBreaks: number
  minimumScale: number
  oversizedDimensions: number
  paper: {
    label: string
    widthInches: number
    heightInches: number
  }
}

export interface SpreadsheetPrintResult {
  printed: boolean
  canceled: boolean
  sheets: number
  cells: number
  pages: number
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
    format: string
    suggestedName: string
    sourceUnmodified?: boolean
  }) => Promise<SaveResult | null>
  exportWorkbook: (input: {
    documentId: string
    name: string
    workbook: WorkbookModel
    displayValues: Record<string, Record<string, string>>
    displayParts?: SpreadsheetDisplayParts
    selection: { top: number; bottom: number; left: number; right: number }
    options: SpreadsheetPrintOptions
    format: SpreadsheetExportFormat
    suggestedName: string
    sourceUnmodified?: boolean
    /** The user agreed to a values-only XLS (no document engine). */
    acceptLoss?: boolean
  }) => Promise<SpreadsheetExportResult | null>
  /** What an export to `format` would lose, and whether it needs the user's consent first. */
  checkExport: (input: { workbook: WorkbookModel; format: SpreadsheetExportFormat }) => Promise<{ format: string; officeEngine: boolean; losses: string[]; confirmationRequired: boolean }>
  renderPrintPreview: (input: {
    documentId: string
    name: string
    workbook: WorkbookModel
    displayValues: Record<string, Record<string, string>>
    displayParts?: SpreadsheetDisplayParts
    selection: { top: number; bottom: number; left: number; right: number }
    options: SpreadsheetPrintOptions
  }) => Promise<SpreadsheetPrintPreview>
  printWorkbook: (input: {
    documentId: string
    name: string
    workbook: WorkbookModel
    displayValues: Record<string, Record<string, string>>
    displayParts?: SpreadsheetDisplayParts
    selection: { top: number; bottom: number; left: number; right: number }
    options: SpreadsheetPrintOptions
  }) => Promise<SpreadsheetPrintResult>
  showItem: (filePath: string) => Promise<void>
  openExternal: (target: string) => Promise<void>
  getVersion: () => Promise<string>
  /** Excel sheet-protection password hashing (main process). */
  protection?: {
    hash: (password: string) => Promise<{ algorithmName: string; hashValue: string; saltValue: string; spinCount: number }>
    verify: (protection: Record<string, unknown>, password: string) => Promise<boolean>
  }
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
