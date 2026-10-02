export interface LayoutPdfSource {
  /** The page-view PDF request, for the Page view iframe. */
  readonly promise: Promise<Uint8Array>;
  /** True once the page-view conversion has failed. */
  readonly failed: boolean;
  readonly error: unknown;
  /** The page-view PDF, or the editor's own render when the page view failed. */
  pdfOr(render: () => Promise<Uint8Array>): Promise<Uint8Array>;
}

export function createLayoutPdfSource(load: () => Promise<Uint8Array>): LayoutPdfSource;
export const LAYOUT_FALLBACK_NOTE: string;
