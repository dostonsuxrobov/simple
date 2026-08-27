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

export interface MarkupOverlay {
  id: string
  type: 'markup'
  pageIndex: number
  style: 'underline' | 'strikeout' | 'rectangle'
  rect: PdfRect
  color: [number, number, number]
  opacity: number
  thickness: number
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
  originalText?: string
  text: string
  fontSize: number
  fontFamily: string
  fontWeight?: number
  fontStyle?: 'normal' | 'italic'
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
  originalText: string
  text: string
  fontSize: number
  fontFamily: string
  fontWeight?: number
  fontStyle?: 'normal' | 'italic'
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

export interface SimpleApi {
  openFile: () => Promise<DocumentPayload | null>
  openInNewWindow: (filePath?: string) => Promise<boolean>
  openPath: (filePath: string) => Promise<DocumentPayload>
  openBytes: (name: string, data: ArrayBuffer) => Promise<DocumentPayload>
  mutatePdf: (data: Uint8Array, operation: Record<string, unknown>) => Promise<Uint8Array>
  flattenOverlays: (
    data: Uint8Array,
    overlays: PdfOverlay[],
    formValues?: Record<string, string | boolean>,
    documentEdits?: { pageRotations?: Record<number, number>; bookmarks?: Bookmark[] },
  ) => Promise<Uint8Array>
  insertFiles: (data: Uint8Array, insertIndex: number) => Promise<{ data: Uint8Array; added: number } | null>
  insertDroppedFiles: (data: Uint8Array, insertIndex: number, files: Array<{ name: string; data: ArrayBuffer }>) => Promise<{ data: Uint8Array; added: number }>
  pickImage: () => Promise<{ dataUrl: string; name: string } | null>
  exportPages: (data: Uint8Array, indices: number[], suggestedName: string) => Promise<string | null>
  startPageDrag: (data: Uint8Array, indices: number[], suggestedName: string) => Promise<string | null>
  printPdf: (data: Uint8Array, name: string) => Promise<boolean>
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
