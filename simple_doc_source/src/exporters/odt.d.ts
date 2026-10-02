import type { Document } from "@forevka/wordcanvas";
import type { NativeExportOptions } from "./rtf.js";

export const ODT_MIME_TYPE: "application/vnd.oasis.opendocument.text";
export function xmlEscape(value: string): string;
export function documentToOdtParts(document: Document, options?: NativeExportOptions): {
  "content.xml": string;
  "styles.xml": string;
  "meta.xml": string;
  "META-INF/manifest.xml": string;
  media: Array<{ path: string; bytes: Uint8Array; type: string; extension: string }>;
};
/** ODT bytes for a document prepared with prepareDocumentForExport(). */
export function documentToOdt(document: Document, options?: NativeExportOptions): Promise<Uint8Array>;
