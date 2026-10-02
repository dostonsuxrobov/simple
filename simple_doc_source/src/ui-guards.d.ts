export interface ShortcutEventLike {
  key: string;
  /** The physical key ("KeyS"); decides on layouts that do not type Latin letters. */
  code?: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  shiftKey?: boolean;
  altKey?: boolean;
}

export function documentExportDisabled(documentOpen: boolean, busy: boolean): boolean;
export function activeModal(root: ParentNode): HTMLElement | null;
export function isModalBlockedShortcut(event: ShortcutEventLike): boolean;
export function nextPrintPreviewGeneration(current: number): number;
export function canCommitPrintPreview(requestGeneration: number, currentGeneration: number, modalHidden: boolean): boolean;

export interface ReviewLike {
  suggestions?: readonly unknown[];
  threads?: readonly unknown[];
}

export interface ReviewSummary {
  suggestions: number;
  /** Comment threads (a thread with replies counts once). */
  comments: number;
  total: number;
}

export function reviewContentKey(review: ReviewLike | null | undefined): string;
export function reviewSummary(review: ReviewLike | null | undefined): ReviewSummary;
export function describeReviewItems(summary: Pick<ReviewSummary, "suggestions" | "comments">): string;
export function documentContentKey(model: unknown): string;

export interface ZipEntry {
  name: string;
  method: number;
  compressedSize: number;
  size: number;
  localOffset: number;
}

export type FidelityKind =
  | "comments"
  | "trackedChanges"
  | "charts"
  | "smartArt"
  | "drawings"
  | "unsupportedPictures"
  | "linkedPictures"
  | "embeddedObjects"
  | "wordArt"
  | "fields";

export interface FidelityItem {
  kind: FidelityKind;
  count: number;
  /** Plain-language label, e.g. "2 comments". */
  label: string;
}

export interface FidelityReport {
  items: FidelityItem[];
}

export interface Relationship {
  type: string;
  target: string;
  external: boolean;
}

export interface DocxPart {
  kind: "body" | "header" | "footer" | "notes" | "comments";
  xml: string | null;
  relationships?: Map<string, Relationship>;
  name?: string;
}

export function listZipEntries(data: Uint8Array | ArrayBuffer): Map<string, ZipEntry>;
export function readZipText(data: Uint8Array | ArrayBuffer, entry: ZipEntry | undefined): Promise<string | null>;
export function parseRelationships(xml: string | null | undefined): Map<string, Relationship>;
export function scanDocxFidelity(data: Uint8Array | ArrayBuffer): Promise<FidelityReport>;
export function analyzeDocxParts(parts: readonly DocxPart[]): FidelityReport;
export function describeFidelityItems(items: readonly FidelityItem[]): string;

export function cleanErrorMessage(error: unknown, fallback?: string): string;
export function isMissingFileError(error: unknown): boolean;
export function openErrorMessage(error: unknown, fallback?: string): string;
export function recoveredDocumentName(name: string | null | undefined): string;
export function notificationDuration(message: string, tone?: "normal" | "error"): number;
