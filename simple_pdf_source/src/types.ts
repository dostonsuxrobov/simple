import type { PDFDocumentProxy } from 'pdfjs-dist'

export type ToolMode =
  | 'hand'
  | 'select'
  | 'edit'
  | 'highlight'
  | 'underline'
  | 'strikeout'
  | 'draw'
  | 'rectangle'
  | 'crop'
  | 'addText'
  | 'sign'

export interface DocumentPayload {
  data: Uint8Array | ArrayBuffer | { type: 'Buffer'; data: number[] }
  name: string
  path: string | null
  sourcePath: string | null
  converted: boolean
  signatureDetected: boolean
  /** Decided by the main process from the whole file; absent for bytes the renderer opened itself. */
  encrypted?: boolean
}

export interface OpenDocument {
  bytes: Uint8Array
  name: string
  path: string | null
  sourcePath: string | null
  converted: boolean
  signatureDetected: boolean
  /** An unprotected copy of a password- or permissions-protected PDF. */
  unlocked?: boolean
}

export interface PdfRect {
  x: number
  y: number
  width: number
  height: number
}

export interface HighlightOverlay {
  id: string
  type: 'highlight'
  pageIndex: number
  rect: PdfRect
  color: [number, number, number]
  opacity: number
}

export type DisplayRotation = 0 | 90 | 180 | 270

export interface MarkupOverlay {
  id: string
  type: 'markup'
  pageIndex: number
  style: 'underline' | 'strikeout' | 'rectangle'
  rect: PdfRect
  color: [number, number, number]
  opacity: number
  thickness: number
  /** Effective page rotation shown on screen when this markup was created. */
  displayRotation?: DisplayRotation
}

export interface InkPoint {
  x: number
  y: number
}

export interface InkOverlay {
  id: string
  type: 'ink'
  pageIndex: number
  points: InkPoint[]
  color: [number, number, number]
  opacity: number
  thickness: number
}

/** Font style measured from a scanned line's pixels (electron/scan-style.mjs). */
export interface ScanTextStyle {
  fontClass: 'serif' | 'sans' | 'mono'
  /** The Windows family that stands in for it. */
  fontFamily: string
  fontWeight: 400 | 700
  italic: boolean
  /** Points. */
  fontSize: number
  xHeight: number
  ascender?: number
  strokeRatio: number
  color: [number, number, number]
  background: [number, number, number]
  confidence: number
}

/** A recognised word of a scanned line. */
export interface ScanWord {
  text: string
  /** Ink box, PDF space. */
  rect: PdfRect
  /** Extent along the line's baseline from its origin, points. */
  start: number
  end: number
}

/** A run of recognised (invisible) text: where it starts, which way it reads, how long it is. */
export interface ScanTextRun {
  origin: { x: number; y: number }
  dir: { x: number; y: number }
  length: number
  /** Size of the invisible OCR text, points. */
  fontSize: number
}

/** A retouch patch: an RGBA PNG whose alpha covers only the old glyph pixels. */
export interface ScanPatch {
  rect: PdfRect
  dataUrl: string
  dpi: number
}

/**
 * What a minimal-footprint commit replaces when only some words of a scanned
 * line changed (electron/scan-replacement.mjs). The edit itself keeps the
 * whole line so it reopens as the line.
 */
export interface ScanReplacement {
  /** Changed words, inclusive (last = first - 1 for a pure insertion). */
  first: number
  last: number
  originalText: string
  text: string
  /** Box the new words are laid out in (reading frame, like PageTextEdit.rect). */
  rect: PdfRect
  /** Union of the replaced words' ink boxes. */
  originalRect: PdfRect
  baselineOffset: number
  /** The recognised text it replaces. */
  run: ScanTextRun
  patch?: ScanPatch
}

/**
 * An edit of scanned text (a line of an OCR text layer, Simple's or another
 * tool's): the old printed words are retouched out of the scan and the new
 * text is drawn in a matched font; or, in 'recognized-text' mode, only the
 * invisible recognised text is corrected and the page looks the same.
 */
export interface ScanEditInfo {
  /** Stable key for async preparation and patch composition. */
  key: string
  /** 'appearance' replaces the printed words; 'recognized-text' only fixes the invisible OCR text. */
  mode: 'appearance' | 'recognized-text'
  status: 'pending' | 'ready' | 'failed'
  /** Line geometry (PDF space) the patch and style were derived from. */
  lineRect: PdfRect
  baseline: { origin: { x: number; y: number }; dir: { x: number; y: number } }
  /** The line's recognised text, as found on the page. */
  lineText: string
  /** The recognised text of the whole line (removed on save in 'line' replacements). */
  run: ScanTextRun
  /** Word boxes aligned with the line's words (from the OCR cache or pixel segmentation). */
  words?: ScanWord[]
  /** RGBA PNG that covers only the old glyph pixels; drawn under the new text. */
  patch?: ScanPatch
  style?: ScanTextStyle
  /** True once the user changed font/size/style/colour; prepared styles must not overwrite it. */
  userStyled?: boolean
  /** Paper colour around the line (0..1 RGB): hides the old words until the patch is ready, or if it fails. */
  paper?: [number, number, number]
  /** Set when committing: only these words change (see ScanReplacement). */
  replace?: ScanReplacement
}

export interface TextOverlay {
  id: string
  type: 'text'
  pageIndex: number
  rect: PdfRect
  originalRect?: PdfRect
  /** Actual source glyph extents, including descenders outside the em box. */
  inkRect?: PdfRect
  originalText?: string
  text: string
  fontSize: number
  fontFamily: string
  fontWeight?: number
  fontStyle?: 'normal' | 'italic'
  /** Fit each line to the box, or wrap at its right edge. */
  textFit?: 'fit' | 'wrap'
  /** False after an intentional font/size/spacing change. */
  preserveSourceMetrics?: boolean
  lineHeight?: number
  letterSpacing?: number
  scaleX?: number
  angle?: number
  direction?: 'ltr' | 'rtl'
  fontKey?: string
  fontData?: Uint8Array
  /** Original PDF baseline, measured upward from rect.y in PDF points. */
  baselineOffset?: number
  /** Original PDF word-space advance in points, before replacement fitting. */
  sourceSpaceWidth?: number
  /** Effective page rotation shown on screen when this text box was created. */
  displayRotation?: DisplayRotation
  align: 'left' | 'center' | 'right'
  color: [number, number, number]
  backgroundColor?: [number, number, number]
  cover: boolean
  /** An edit of scanned (OCR) text. */
  scan?: ScanEditInfo
}

export interface ObjectOverlay {
  id: string
  type: 'object'
  pageIndex: number
  kind: 'image' | 'artwork'
  rect: PdfRect
  originalRect?: PdfRect
  dataUrl?: string
  opacity: number
  cover: boolean
  /** Effective page rotation shown on screen when this object was captured. */
  displayRotation?: DisplayRotation
}

export type PdfOverlay = HighlightOverlay | MarkupOverlay | InkOverlay | TextOverlay | ObjectOverlay

export interface Bookmark {
  id: string
  /** Destination page, or null for a heading or a link that does not go to a page. */
  pageIndex: number | null
  label: string
  depth?: number
  source?: 'document' | 'simple'
  expanded?: boolean
  bold?: boolean
  italic?: boolean
  color?: [number, number, number]
  url?: string | null
  /**
   * Where a document bookmark sat in the PDF's outline when the file was
   * opened ("2.1" is the first child of the second top-level item), and its
   * title then. Saving keeps that outline item, with its exact destination
   * and actions, instead of rebuilding it.
   */
  outlinePath?: string
  originalLabel?: string
}

/** A form field value: text, a check box state, or radio/choice export value(s). */
export type PdfFormValue = string | boolean | string[]

/** Something a save could not write as asked (see electron/pdf-problems.cjs). */
export interface PdfSaveProblem {
  code: string
  message: string
  overlayId?: string
  pageIndex?: number
  /** Form field name, as pdf.js reports it. */
  field?: string
  /** The saved value differs from what was typed (for example truncated). */
  dataLoss?: boolean
  blockingMessage?: string
}

export interface FlattenReport {
  ok: boolean
  data: Uint8Array
  /** Edits that were written, but changed (for example shrunk to fit). */
  warnings: PdfSaveProblem[]
  /** Edits that could not be written at all. */
  failures: PdfSaveProblem[]
  signatureDetected: boolean
}

export interface DocumentEdits {
  pageRotations?: Record<number, number>
  bookmarks?: Bookmark[]
}

export interface PageTextEdit {
  overlayId?: string
  pageIndex: number
  rect: PdfRect
  originalRect?: PdfRect
  inkRect?: PdfRect
  originalText: string
  text: string
  fontSize: number
  fontFamily: string
  fontWeight?: number
  fontStyle?: 'normal' | 'italic'
  /** Fit each line to the box, or wrap at its right edge. */
  textFit?: 'fit' | 'wrap'
  /** False after an intentional font/size/spacing change. */
  preserveSourceMetrics?: boolean
  lineHeight?: number
  letterSpacing?: number
  scaleX?: number
  angle?: number
  direction?: 'ltr' | 'rtl'
  fontKey?: string
  fontData?: Uint8Array
  /** Original PDF baseline, measured upward from rect.y in PDF points. */
  baselineOffset?: number
  /** Original PDF word-space advance in points, before replacement fitting. */
  sourceSpaceWidth?: number
  /** Complete PDF.js text item containing this directly edited selection. */
  sourceItemText?: string
  /** Original rectangle of the complete PDF.js text item. */
  sourceItemRect?: PdfRect
  /** UTF-16 offsets of the selection inside sourceItemText. */
  sourceSelectionStart?: number
  sourceSelectionEnd?: number
  /** Effective page rotation shown on screen when this text box was created. */
  displayRotation?: DisplayRotation
  align: 'left' | 'center' | 'right'
  color: [number, number, number]
  backgroundColor?: [number, number, number]
  cover: boolean
  /** An edit of scanned (OCR) text. */
  scan?: ScanEditInfo
  modified: boolean
  caretOffset?: number
  /** Selection to restore inside the full logical text run when editing starts. */
  selectionStart?: number
  selectionEnd?: number
}

export interface DetectedPageObject {
  id: string
  pageIndex: number
  kind: 'image' | 'artwork'
  rect: PdfRect
  /** Pixels decoded from the actual PDF image XObject, without page text. */
  dataUrl?: string
  label: string
}

export interface PageObjectEdit {
  overlayId?: string
  candidateId?: string
  pageIndex: number
  kind: 'image' | 'artwork'
  rect: PdfRect
  originalRect?: PdfRect
  dataUrl?: string
  opacity: number
  cover: boolean
  /** Effective page rotation shown on screen when this object was captured. */
  displayRotation?: DisplayRotation
  label: string
  modified: boolean
}

export interface RecentFile {
  path: string
  name: string
  openedAt: number
}

export interface SearchResult {
  pageIndex: number
  excerpt: string
  count: number
}

export interface ActiveSearchMatch {
  pageIndex: number
  /** Zero-based occurrence within this page. */
  occurrenceIndex: number
  /** Normalized query used by both the search index and text-layer matcher. */
  query: string
}

export interface PrinterSummary {
  name: string
  displayName: string
  supportsDuplex: boolean
  supportsColor: boolean
}

export type PrintDuplexMode = 'simplex' | 'longEdge' | 'shortEdge'
export type PrintPaperSize = 'Letter' | 'A4' | 'Legal'
export type PrintMarginMode = 'none' | 'minimum' | 'normal'
export type PrintScaleMode = 'fit' | 'actual' | 'shrink' | 'custom'

export interface PrintDirectOptions {
  deviceName?: string
  copies?: number
  /** Zero-based page indices, already resolved; omit to print every page. */
  pageIndices?: number[] | null
  landscape?: boolean
  color?: boolean
  duplexMode?: PrintDuplexMode
  collate?: boolean
  paperSize: PrintPaperSize
  marginMode: PrintMarginMode
  scaleMode: PrintScaleMode
  /** Decimal scale, where 1 is 100%. */
  customScale: number
}

export interface PrintDirectResult {
  success: boolean
  failureReason: string
}

export type PdfExportFormat = 'pdf' | 'png' | 'jpeg' | 'webp' | 'docx' | 'txt' | 'md' | 'html'

export interface ExportTextPage {
  pageNumber: number
  text: string
}

export interface ImageExportSession {
  id: string
  targetPath: string
}

export interface SimpleApi {
  openFile: () => Promise<DocumentPayload | null>
  openInNewWindow: (filePath?: string) => Promise<boolean>
  openPath: (filePath: string) => Promise<DocumentPayload>
  openBytes: (name: string, data: ArrayBuffer) => Promise<DocumentPayload>
  /** The path of a dropped file on disk, or '' when it has none (for example a browser download). */
  getPathForFile?: (file: File) => string
  unlockPdf: (data: Uint8Array, password: string) => Promise<{ status: 'none' | 'needs-password' | 'wrong-password' | 'unlocked'; data?: Uint8Array }>
  mutatePdf: (data: Uint8Array, operation: Record<string, unknown>) => Promise<Uint8Array>
  textBackground: (data: Uint8Array, pageIndex: number, edits: Array<{ type: 'text'; cover: boolean; originalRect: PdfRect; originalText?: string }>) => Promise<Uint8Array>
  /** With `report: true` the main process returns what it could not write instead of throwing. */
  flattenOverlays(
    data: Uint8Array,
    overlays: PdfOverlay[],
    formValues: Record<string, PdfFormValue> | undefined,
    documentEdits: DocumentEdits & { report: true },
  ): Promise<FlattenReport>
  flattenOverlays(
    data: Uint8Array,
    overlays: PdfOverlay[],
    formValues?: Record<string, PdfFormValue>,
    documentEdits?: DocumentEdits,
  ): Promise<Uint8Array>
  insertFiles: (data: Uint8Array, insertIndex: number) => Promise<{ data: Uint8Array; added: number } | null>
  insertDroppedFiles: (data: Uint8Array, insertIndex: number, files: Array<{ name: string; data: ArrayBuffer }>) => Promise<{ data: Uint8Array; added: number }>
  pickImage: () => Promise<{ dataUrl: string; name: string } | null>
  exportAsPdf: (input: { data: Uint8Array; indices: number[]; fullDocument: boolean; suggestedName: string }) => Promise<string | null>
  exportTextDocument: (input: { format: 'docx' | 'txt' | 'md' | 'html'; pages: ExportTextPage[]; title: string; baseName: string }) => Promise<string | null>
  beginImageExport: (input: { format: 'png' | 'jpeg' | 'webp'; pageCount: number; baseName: string }) => Promise<ImageExportSession | null>
  writeImageExportPage: (input: { id: string; pageNumber: number; data: Uint8Array }) => Promise<string>
  finishImageExport: (id: string) => Promise<string>
  cancelImageExport: (id: string) => Promise<boolean>
  exportPages: (data: Uint8Array, indices: number[], suggestedName: string) => Promise<string | null>
  startPageDrag: (data: Uint8Array, indices: number[], suggestedName: string) => Promise<string | null>
  listPrinters: () => Promise<PrinterSummary[]>
  printPdfDirect: (data: Uint8Array, name: string, options: PrintDirectOptions) => Promise<PrintDirectResult>
  /**
   * Write the PDF. Without a path, or with forceDialog, a Save dialog opens on
   * `name` (a full path opens it in that folder). A validly signed original is
   * never overwritten: the dialog offers a copy and `keptSignedOriginal` is set.
   */
  savePdf: (input: { data: Uint8Array; path: string | null; name: string; forceDialog: boolean }) => Promise<{ path: string; name: string; keptSignedOriginal?: boolean } | null>
  showItem: (filePath: string) => Promise<void>
  getVersion: () => Promise<string>
  minimize: () => void
  toggleMaximize: () => void
  close: () => void
  onMaximized: (callback: (maximized: boolean) => void) => () => void
  onOpenExternal: (callback: (filePath: string) => void) => () => void
  onCloseRequested: (callback: () => void) => () => void
  openExternal: (url: string) => Promise<void>
}

export interface ViewerDocument {
  file: OpenDocument
  pdf: PDFDocumentProxy
}

declare global {
  interface Window {
    simple: SimpleApi
  }
}
