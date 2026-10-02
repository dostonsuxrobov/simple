export interface TextPaint {
  /** Original fill color (0..1 RGB) when it is known and the text is visibly painted. */
  color?: number[]
  /** Drawn in render mode 3 or 7 (paints nothing), e.g. an OCR text layer over a scan. */
  invisible: boolean
}
/** Resolvers consume text in content order: call once per pdf.js text item, in text-content order. */
export function textPaintResolver(list: { fnArray: number[]; argsArray: unknown[] }, OPS: Record<string, number>): (fontName: string, text: string) => TextPaint
export function textColorResolver(list: { fnArray: number[]; argsArray: unknown[] }, OPS: Record<string, number>): (fontName: string, text: string) => number[] | undefined
