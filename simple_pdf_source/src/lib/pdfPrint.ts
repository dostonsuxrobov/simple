export type PrintPaperSize = 'Letter' | 'A4' | 'Legal'
export type PrintMarginMode = 'none' | 'minimum' | 'normal'
export type PrintScaleMode = 'fit' | 'actual' | 'shrink' | 'custom'

export interface PdfPrintLayout {
  paperSize: PrintPaperSize
  landscape: boolean
  marginMode: PrintMarginMode
  scaleMode: PrintScaleMode
  customScale: number
}

export const PRINT_PAPER_POINTS: Record<PrintPaperSize, { width: number; height: number }> = {
  Letter: { width: 612, height: 792 },
  A4: { width: 595.28, height: 841.89 },
  Legal: { width: 612, height: 1008 },
}

export const PRINT_MARGIN_POINTS: Record<PrintMarginMode, number> = {
  none: 0,
  minimum: 18,
  normal: 36,
}

export function calculatePdfPrintPlacement(sourceWidth: number, sourceHeight: number, rotation: number, layout: PdfPrintLayout) {
  const base = PRINT_PAPER_POINTS[layout.paperSize]
  const paperWidth = layout.landscape ? base.height : base.width
  const paperHeight = layout.landscape ? base.width : base.height
  const normalizedRotation = ((Math.round(rotation / 90) * 90) % 360 + 360) % 360
  const quarterTurn = normalizedRotation === 90 || normalizedRotation === 270
  const rawPaperWidth = quarterTurn ? paperHeight : paperWidth
  const rawPaperHeight = quarterTurn ? paperWidth : paperHeight
  const margin = PRINT_MARGIN_POINTS[layout.marginMode]
  const printableWidth = Math.max(1, rawPaperWidth - margin * 2)
  const printableHeight = Math.max(1, rawPaperHeight - margin * 2)
  const fitScale = Math.min(printableWidth / sourceWidth, printableHeight / sourceHeight)
  const scale = layout.scaleMode === 'actual' ? 1
    : layout.scaleMode === 'shrink' ? Math.min(1, fitScale)
      : layout.scaleMode === 'custom' ? Math.min(4, Math.max(0.25, layout.customScale))
        : fitScale
  const contentWidth = sourceWidth * scale
  const contentHeight = sourceHeight * scale
  const rawX = (rawPaperWidth - contentWidth) / 2
  const rawY = (rawPaperHeight - contentHeight) / 2

  // Convert the raw PDF coordinate placement to the orientation visible on paper.
  const visibleContentWidth = quarterTurn ? contentHeight : contentWidth
  const visibleContentHeight = quarterTurn ? contentWidth : contentHeight
  return {
    paperWidth,
    paperHeight,
    margin,
    printableWidth: paperWidth - margin * 2,
    printableHeight: paperHeight - margin * 2,
    scale,
    contentWidth: visibleContentWidth,
    contentHeight: visibleContentHeight,
    x: (paperWidth - visibleContentWidth) / 2,
    y: (paperHeight - visibleContentHeight) / 2,
    cropped: rawX < margin - 0.01 || rawY < margin - 0.01,
  }
}
