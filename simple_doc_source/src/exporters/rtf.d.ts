import type { Document } from "@forevka/wordcanvas";

export interface NativeExportOptions {
  /** Document title written to the file's properties. */
  title?: string;
  /** Receives plain-language notes about content shown as a placeholder. */
  warnings?: string[];
}

/** Escape text for RTF (\uN escapes for non-ASCII). */
export function rtfText(value: string): string;
/** RTF text for a document prepared with prepareDocumentForExport(). */
export function documentToRtf(document: Document, options?: NativeExportOptions): string;
export function documentToRtfBytes(document: Document, options?: NativeExportOptions): Uint8Array;
