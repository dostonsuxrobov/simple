import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfRect } from '../types'
import { loadPdf, normalizeBytes } from './pdf'

const sourceBytes = new WeakMap<PDFDocumentProxy, Promise<Uint8Array>>()

/** The preview uses the same text-only removal as Save, on a disposable page. */
export async function textBackground(pdf: PDFDocumentProxy, pageIndex: number,
  edits: Array<{ type: 'text'; cover: boolean; originalRect: PdfRect; originalText?: string }>) {
  let bytes = sourceBytes.get(pdf)
  if (!bytes) {
    bytes = pdf.getData()
    sourceBytes.set(pdf, bytes)
  }
  const data = await window.simple.textBackground(await bytes, pageIndex, edits)
  return loadPdf(normalizeBytes(data))
}
