/**
 * Picture helpers for the importers: byte sniffing, natural size, data: URLs. Only
 * bytes that are already inside the imported file are ever used; nothing is fetched.
 */

export type ImportImageType = "image/png" | "image/jpeg" | "image/gif" | "image/webp" | "image/bmp" | "image/svg+xml";

export interface ImageBytes {
  bytes: Uint8Array;
  type: ImportImageType;
}

/** Largest picture the importers embed (bytes). */
export const MAX_IMAGE_BYTES = 40 * 1024 * 1024;

const startsWith = (bytes: Uint8Array, signature: number[], offset = 0) => signature.every((value, index) => bytes[offset + index] === value);

/** The picture type from its leading bytes, or null. */
export function sniffImageType(bytes: Uint8Array): ImportImageType | null {
  if (bytes.length < 4) return null;
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return "image/png";
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return "image/jpeg";
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return "image/gif";
  if (bytes.length >= 12 && startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) && startsWith(bytes, [0x57, 0x45, 0x42, 0x50], 8)) return "image/webp";
  if (startsWith(bytes, [0x42, 0x4d]) && bytes.length > 26) return "image/bmp";
  const head = new TextDecoder("utf-8").decode(bytes.subarray(0, Math.min(bytes.length, 512))).replace(/^\ufeff/, "").trimStart();
  if (/^(?:<\?xml[^>]*>\s*)?(?:<!--[\s\S]*?-->\s*)*(?:<!DOCTYPE svg[^>]*>\s*)?<svg[\s>]/i.test(head)) return "image/svg+xml";
  return null;
}

const u16be = (bytes: Uint8Array, offset: number) => (bytes[offset] << 8) | bytes[offset + 1];
const u16le = (bytes: Uint8Array, offset: number) => bytes[offset] | (bytes[offset + 1] << 8);
const u32be = (bytes: Uint8Array, offset: number) => ((bytes[offset] << 24) >>> 0) + (bytes[offset + 1] << 16) + (bytes[offset + 2] << 8) + bytes[offset + 3];
const i32le = (bytes: Uint8Array, offset: number) => bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24);

/** Natural pixel size of a picture, or null when it cannot be read. */
export function imageNaturalSize(bytes: Uint8Array, type: ImportImageType | null = sniffImageType(bytes)): { width: number; height: number } | null {
  try {
    if (type === "image/png" && bytes.length >= 24) return valid(u32be(bytes, 16), u32be(bytes, 20));
    if (type === "image/gif" && bytes.length >= 10) return valid(u16le(bytes, 6), u16le(bytes, 8));
    if (type === "image/bmp" && bytes.length >= 26) return valid(Math.abs(i32le(bytes, 18)), Math.abs(i32le(bytes, 22)));
    if (type === "image/jpeg") {
      let offset = 2;
      while (offset + 9 < bytes.length) {
        if (bytes[offset] !== 0xff) { offset += 1; continue; }
        const marker = bytes[offset + 1];
        if (marker === 0xd8 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0xff) { offset += marker === 0xff ? 1 : 2; continue; }
        const size = u16be(bytes, offset + 2);
        if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) return valid(u16be(bytes, offset + 7), u16be(bytes, offset + 5));
        offset += 2 + size;
      }
      return null;
    }
    if (type === "image/webp" && bytes.length >= 30) {
      const chunk = String.fromCharCode(bytes[12], bytes[13], bytes[14], bytes[15]);
      if (chunk === "VP8 ") return valid(u16le(bytes, 26) & 0x3fff, u16le(bytes, 28) & 0x3fff);
      if (chunk === "VP8L") {
        const bits = bytes[21] | (bytes[22] << 8) | (bytes[23] << 16) | (bytes[24] << 24);
        return valid((bits & 0x3fff) + 1, ((bits >>> 14) & 0x3fff) + 1);
      }
      if (chunk === "VP8X") return valid(1 + (bytes[24] | (bytes[25] << 8) | (bytes[26] << 16)), 1 + (bytes[27] | (bytes[28] << 8) | (bytes[29] << 16)));
    }
    if (type === "image/svg+xml") {
      const head = new TextDecoder("utf-8").decode(bytes.subarray(0, Math.min(bytes.length, 4096)));
      const tag = /<svg\b[^>]*>/i.exec(head)?.[0] ?? "";
      const width = Number.parseFloat(/\swidth\s*=\s*["']?\s*([\d.]+)(?:px)?\s*["']?/i.exec(tag)?.[1] ?? "");
      const height = Number.parseFloat(/\sheight\s*=\s*["']?\s*([\d.]+)(?:px)?\s*["']?/i.exec(tag)?.[1] ?? "");
      if (width > 0 && height > 0) return valid(width, height);
      const box = /\sviewBox\s*=\s*["']\s*[-\d.]+[\s,]+[-\d.]+[\s,]+([\d.]+)[\s,]+([\d.]+)/i.exec(tag);
      if (box) return valid(Number(box[1]), Number(box[2]));
    }
  } catch {
    return null;
  }
  return null;
}

function valid(width: number, height: number) {
  return width > 0 && height > 0 && width < 100_000 && height < 100_000 ? { width, height } : null;
}

const BASE64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

export function bytesToBase64(bytes: Uint8Array): string {
  let output = "";
  let i = 0;
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
    output += BASE64[n >> 18] + BASE64[(n >> 12) & 63] + BASE64[(n >> 6) & 63] + BASE64[n & 63];
  }
  if (i < bytes.length) {
    const n = (bytes[i] << 16) | ((bytes[i + 1] ?? 0) << 8);
    output += BASE64[n >> 18] + BASE64[(n >> 12) & 63] + (i + 1 < bytes.length ? BASE64[(n >> 6) & 63] : "=") + "=";
  }
  return output;
}

const BASE64_LOOKUP = (() => {
  const table = new Int16Array(128).fill(-1);
  for (let i = 0; i < BASE64.length; i += 1) table[BASE64.charCodeAt(i)] = i;
  table["-".charCodeAt(0)] = 62;
  table["_".charCodeAt(0)] = 63;
  return table;
})();

export function base64ToBytes(text: string): Uint8Array | null {
  const clean = text.replace(/[\s=]+/g, "");
  const output = new Uint8Array(Math.floor(clean.length * 3 / 4));
  let buffer = 0;
  let bits = 0;
  let position = 0;
  for (let i = 0; i < clean.length; i += 1) {
    const code = clean.charCodeAt(i);
    const value = code < 128 ? BASE64_LOOKUP[code] : -1;
    if (value < 0) return null;
    buffer = (buffer << 6) | value;
    bits += 6;
    if (bits >= 8) {
      bits -= 8;
      output[position++] = (buffer >> bits) & 0xff;
    }
  }
  return output.subarray(0, position);
}

export function hexToBytes(hex: string): Uint8Array {
  const clean = hex.replace(/[^0-9a-fA-F]/g, "");
  const output = new Uint8Array(clean.length >> 1);
  for (let i = 0; i < output.length; i += 1) output[i] = parseInt(clean.substr(i * 2, 2), 16);
  return output;
}

export function bytesToDataUrl(bytes: Uint8Array, type: string): string {
  return `data:${type};base64,${bytesToBase64(bytes)}`;
}

/** Decodes a data: URL to picture bytes; null when it is not a supported picture. */
export function decodeImageDataUrl(url: string): ImageBytes | null {
  const match = /^data:([^,]*?),([\s\S]*)$/i.exec(url.trim());
  if (!match) return null;
  const meta = match[1].toLowerCase();
  // Encoded size check before decoding (base64 grows 4/3).
  if (match[2].length > MAX_IMAGE_BYTES * 1.4) return null;
  let bytes: Uint8Array | null;
  if (/;base64\s*$/.test(meta)) bytes = base64ToBytes(match[2]);
  else {
    try {
      bytes = new TextEncoder().encode(decodeURIComponent(match[2]));
    } catch {
      return null;
    }
  }
  if (!bytes || !bytes.length || bytes.length > MAX_IMAGE_BYTES) return null;
  const type = sniffImageType(bytes);
  return type ? { bytes, type } : null;
}
