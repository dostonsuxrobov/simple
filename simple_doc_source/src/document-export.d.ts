import type { Document } from "@forevka/wordcanvas";

type Block = Document["blocks"][number];
type Paragraph = Extract<Block, { kind: "paragraph" }>;
type Run = Paragraph["runs"][number];

export type StructuredDocumentExportFormat = "html" | "md" | "txt";

export interface StructuredExportOptions {
  /** BCP 47 language tag for the HTML `lang` attribute. */
  lang?: string;
  /** Author written as an HTML `<meta name="author">`. */
  author?: string;
}

export interface StructuredExportResult {
  text: string;
  /** Plain-language notes about content the format replaced with a placeholder. */
  warnings: string[];
}

/** A resolved image: a data: URI, or a Blob-like object with bytes, size and MIME type. */
export type ResolvedExportImage = string | { size: number; type: string; arrayBuffer(): Promise<ArrayBuffer> };

export function documentToHtml(document: Document, title?: string, options?: StructuredExportOptions): string;
export function documentToMarkdown(document: Document, title?: string, options?: StructuredExportOptions): string;
export function documentToText(document: Document): string;
export function serializeDocument(document: Document, format: StructuredDocumentExportFormat, title?: string): string;
export function serializeDocumentWithWarnings(document: Document, format: StructuredDocumentExportFormat, title?: string, options?: StructuredExportOptions): StructuredExportResult;
/**
 * Copy the document with editor object-URL images embedded as data: URIs. An
 * image that cannot be embedded becomes a placeholder; it never aborts the
 * export. Read the omissions with exportWarnings(copy).
 */
export function prepareDocumentForExport(document: Document, resolveImage?: (source: string) => Promise<ResolvedExportImage>): Promise<Document>;
export function exportWarnings(document: Document): string[];

export interface ListMarker { text: string; level: number; ordered: boolean; value: number }
export function createListState(document: Document): { marker(paragraph: Paragraph): ListMarker | null };
export function createNoteState(document: Document): {
  number(kind: "footnote" | "endnote", id: string): number;
  referenced(): Array<{ kind: "footnote" | "endnote"; id: string; number: number; blocks: Paragraph[] }>;
  unreferenced(): Array<{ kind: "footnote" | "endnote"; id: string; blocks: Paragraph[] }>;
};
export function noteReference(run: Run): { kind: "footnote" | "endnote"; id: string } | null;
export function bookmarkStarts(document: Document): Map<string, Array<{ name: string; offset: number }>>;
export function headingLevel(paragraph: Paragraph): number | null;
export function mathText(node: unknown): string;
export function runPlainText(run: Run): string;
export function safeImageSource(value: string): string | null;
export function safeLink(value: string): string | null;
export function sectionsOf(document: Document): Array<{ blocks: Document["blocks"]; props: Document["section"]; breakType?: string }>;
export function storyHasContent(blocks: Document["blocks"] | undefined): boolean;
