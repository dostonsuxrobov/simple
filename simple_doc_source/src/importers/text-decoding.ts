/**
 * Byte-to-text decoding for imported files: byte order marks, UTF-16 without a BOM,
 * strict UTF-8, then a Windows code page guess (1252 Western or 1251 Cyrillic).
 */

export type TextEncodingName = "utf-8" | "utf-16le" | "utf-16be" | "utf-32le" | "utf-32be" | "windows-1252" | "windows-1251" | (string & {});

export interface DecodedText {
  text: string;
  encoding: TextEncodingName;
  /** True when the encoding came from a byte order mark or an explicit declaration. */
  declared: boolean;
}

export function toBytes(input: Uint8Array | ArrayBuffer | ArrayBufferView): Uint8Array {
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
}

function decodeUtf32(bytes: Uint8Array, littleEndian: boolean): string {
  let text = "";
  for (let i = 0; i + 3 < bytes.length; i += 4) {
    const code = littleEndian
      ? (bytes[i] | (bytes[i + 1] << 8) | (bytes[i + 2] << 16) | (bytes[i + 3] << 24)) >>> 0
      : ((bytes[i] << 24) | (bytes[i + 1] << 16) | (bytes[i + 2] << 8) | bytes[i + 3]) >>> 0;
    text += code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : "�";
  }
  return text;
}

/** Decodes with a WHATWG label; unknown labels fall back to windows-1252. */
export function decodeWith(bytes: Uint8Array, label: string): string {
  const name = label.toLowerCase();
  if (name === "utf-32le" || name === "utf-32") return decodeUtf32(bytes, true);
  if (name === "utf-32be") return decodeUtf32(bytes, false);
  try {
    return new TextDecoder(name).decode(bytes);
  } catch {
    return new TextDecoder("windows-1252").decode(bytes);
  }
}

/** The encoding named by a byte order mark, with the BOM length. */
export function bomEncoding(bytes: Uint8Array): { encoding: TextEncodingName; length: number } | null {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { encoding: "utf-8", length: 3 };
  if (bytes[0] === 0xff && bytes[1] === 0xfe && bytes[2] === 0 && bytes[3] === 0 && bytes.length % 4 === 0) return { encoding: "utf-32le", length: 4 };
  if (bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 0xfe && bytes[3] === 0xff) return { encoding: "utf-32be", length: 4 };
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { encoding: "utf-16le", length: 2 };
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { encoding: "utf-16be", length: 2 };
  return null;
}

/** UTF-16 without a BOM: most ASCII text has a zero byte in every other position. */
function sniffUtf16(bytes: Uint8Array): TextEncodingName | null {
  const sample = Math.min(bytes.length - (bytes.length % 2), 8192);
  if (sample < 4) return null;
  let evenZero = 0;
  let oddZero = 0;
  for (let i = 0; i < sample; i += 2) {
    if (bytes[i] === 0) evenZero += 1;
    if (bytes[i + 1] === 0) oddZero += 1;
  }
  const pairs = sample / 2;
  if (oddZero / pairs > 0.3 && evenZero / pairs < 0.05) return "utf-16le";
  if (evenZero / pairs > 0.3 && oddZero / pairs < 0.05) return "utf-16be";
  return null;
}

const isAsciiLetter = (byte: number | undefined) => byte !== undefined && ((byte >= 0x41 && byte <= 0x5a) || (byte >= 0x61 && byte <= 0x7a));
const isHighLetter = (byte: number | undefined) => byte !== undefined && byte >= 0xc0;

/**
 * Western (1252) vs Cyrillic (1251) for 8-bit text. Accented Latin letters sit next to
 * ASCII letters ("café"); Cyrillic letters form whole words of high bytes ("привет").
 */
export function guessSingleByteEncoding(bytes: Uint8Array): "windows-1252" | "windows-1251" {
  let latin = 0;
  let cyrillic = 0;
  const limit = Math.min(bytes.length, 65536);
  for (let i = 0; i < limit; i += 1) {
    const byte = bytes[i];
    if (!isHighLetter(byte)) continue;
    const before = bytes[i - 1];
    const after = bytes[i + 1];
    if (isAsciiLetter(before) || isAsciiLetter(after)) latin += 1;
    else if (isHighLetter(before) || isHighLetter(after)) cyrillic += 1;
  }
  return cyrillic > latin * 2 ? "windows-1251" : "windows-1252";
}

function isValidUtf8(bytes: Uint8Array): boolean {
  try {
    new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return true;
  } catch {
    return false;
  }
}

/**
 * Decodes text bytes. `encoding` forces a label; `declared` is a label found inside the
 * file (HTML meta charset, RTF code page) that wins over guessing but not over a BOM.
 */
export function decodeTextBytes(input: Uint8Array | ArrayBuffer, options: { encoding?: string; declared?: string | null } = {}): DecodedText {
  const bytes = toBytes(input);
  if (options.encoding) return { text: decodeWith(bytes, options.encoding), encoding: options.encoding.toLowerCase(), declared: true };
  const bom = bomEncoding(bytes);
  if (bom) return { text: decodeWith(bytes.subarray(bom.length), bom.encoding), encoding: bom.encoding, declared: true };
  const utf16 = sniffUtf16(bytes);
  if (utf16) return { text: decodeWith(bytes, utf16).replace(/^\ufeff/, ""), encoding: utf16, declared: false };
  const declared = options.declared?.trim().toLowerCase();
  if (declared && declared !== "utf-16" && declared !== "unicode") {
    const label = declared === "utf8" ? "utf-8" : declared;
    if (label !== "utf-8" || isValidUtf8(bytes)) return { text: decodeWith(bytes, label), encoding: label, declared: true };
  }
  if (isValidUtf8(bytes)) return { text: new TextDecoder("utf-8").decode(bytes), encoding: "utf-8", declared: false };
  const guess = guessSingleByteEncoding(bytes);
  return { text: new TextDecoder(guess).decode(bytes), encoding: guess, declared: false };
}
