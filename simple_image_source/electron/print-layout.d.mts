export type PrintPaper = 'letter' | 'a4' | 'legal'
export type PrintOrientation = 'portrait' | 'landscape'
export type PrintScaleMode = 'fit' | 'fill' | 'actual' | 'custom'
export type PrintPosition = 'center' | 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right'

export interface ImagePrintSettings {
  paper: PrintPaper
  orientation: PrintOrientation
  marginMm: number
  scaleMode: PrintScaleMode
  scalePercent: number
  position: PrintPosition
  background: string
  grayscale: boolean
  copies: number
}

export interface PrintLayout {
  settings: ImagePrintSettings
  paper: { label: string; electronName: 'Letter' | 'A4' | 'Legal'; widthMm: number; heightMm: number }
  printable: { xMm: number; yMm: number; widthMm: number; heightMm: number }
  image: { xMm: number; yMm: number; widthMm: number; heightMm: number }
  clipped: boolean
  effectiveDpi: number
}

export const PRINT_PAPERS: Readonly<Record<PrintPaper, Readonly<{ label: string; electronName: 'Letter' | 'A4' | 'Legal'; widthMm: number; heightMm: number }>>>
export const DEFAULT_PRINT_SETTINGS: Readonly<ImagePrintSettings>
export function normalizePrintSettings(input?: Partial<ImagePrintSettings>): ImagePrintSettings
export function orientedPaper(settings?: Partial<ImagePrintSettings>): PrintLayout['paper']
export function computePrintLayout(imageWidth: number, imageHeight: number, settings?: Partial<ImagePrintSettings>): PrintLayout
export function electronPrintOptions(settings?: Partial<ImagePrintSettings>): {
  silent: true
  printBackground: true
  color: boolean
  landscape: boolean
  scaleFactor: 100
  copies: number
  collate: true
  margins: { marginType: 'none' }
  pageSize: 'Letter' | 'A4' | 'Legal'
}
export function buildPrintHtml(imageSource: string, imageWidth: number, imageHeight: number, settings?: Partial<ImagePrintSettings>, title?: string): string
