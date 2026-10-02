import type { Document } from "@forevka/wordcanvas";

export const PORTABLE_IMAGE_TYPES: readonly ["image/png", "image/jpeg"];
export type ImageCodec = (bytes: Uint8Array, type: string) => Promise<{ bytes: Uint8Array; type: string; width?: number; height?: number }>;

export function sniffImageType(bytes: Uint8Array | ArrayBuffer): string | null;
export function isPortableImageType(type: string): boolean;
/** Decode in the browser and re-encode as PNG (transparent) or JPEG (opaque). */
export const browserImageCodec: ImageCodec;
export function normalizeImageBytes(bytes: Uint8Array | ArrayBuffer, options?: { type?: string; codec?: ImageCodec }): Promise<{ bytes: Uint8Array; type: string; converted: boolean; from?: string }>;
export function normalizeDocumentImages(document: Document, options: {
  resolve: (source: string) => Promise<Blob | ArrayBuffer | Uint8Array>;
  register: (bytes: Uint8Array, type: string) => Promise<string> | string;
  codec?: ImageCodec;
}): Promise<{ document: Document; converted: number; failed: number; warnings: string[] }>;
