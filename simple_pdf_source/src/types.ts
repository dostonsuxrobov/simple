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
}

export interface OpenDocument {
  bytes: Uint8Array
  name: string
  path: string | null
  sourcePath: string | null
  converted: boolean
  signatureDetected: boolean
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
  pageIndex: number
  label: string
  depth?: number
  source?: 'document' | 'simple'
  expanded?: boolean
  bold?: boolean
  italic?: boolean
  color?: [number, number, number]
  url?: string | null
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
  unlockPdf: (data: Uint8Array, password: string) => Promise<{ status: 'none' | 'needs-password' | 'wrong-password' | 'unlocked'; data?: Uint8Array }>
  mutatePdf: (data: Uint8Array, operation: Record<string, unknown>) => Promise<Uint8Array>
  textBackground: (data: Uint8Array, pageIndex: number, edits: Array<{ type: 'text'; cover: boolean; originalRect: PdfRect; originalText?: string }>) => Promise<Uint8Array>
  flattenOverlays: (
    data: Uint8Array,
    overlays: PdfOverlay[],
    formValues?: Record<string, string | boolean>,
    documentEdits?: { pageRotations?: Record<number, number>; bookmarks?: Bookmark[] },
  ) => Promise<Uint8Array>
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
  savePdf: (input: { data: Uint8Array; path: string | null; name: string; forceDialog: boolean }) => Promise<{ path: string; name: string } | null>
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
