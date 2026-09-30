import type { Document } from "@forevka/wordcanvas";

export type StructuredDocumentExportFormat = "html" | "md" | "txt";

export function documentToHtml(document: Document, title?: string): string;
export function documentToMarkdown(document: Document, title?: string): string;
export function documentToText(document: Document): string;
export function serializeDocument(document: Document, format: StructuredDocumentExportFormat, title?: string): string;
export function prepareDocumentForExport(document: Document, resolveImage?: (source: string) => Promise<string>): Promise<Document>;
