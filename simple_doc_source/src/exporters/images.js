// Picture normalization. The editor's PDF/print engine embeds only PNG and
// JPEG, and its DOCX writer labels unknown pictures as PNG. Converting GIF,
// WebP, BMP, SVG, ICO and AVIF pictures to PNG (transparent) or JPEG (opaque)
// before insertion or before a PDF/print export keeps them visible everywhere.
// Everything runs locally in the renderer; nothing is downloaded.

export const PORTABLE_IMAGE_TYPES = Object.freeze(["image/png", "image/jpeg"]);
const MAX_DIMENSION = 4096;
const SVG_SCALE = 2;

function ascii(bytes, start, end) {
  let text = "";
  for (let index = start; index < Math.min(end, bytes.length); index += 1) text += String.fromCharCode(bytes[index]);
  return text;
}

/** MIME type from the picture's own bytes, or null when unknown. */
export function sniffImageType(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input || []);
  if (bytes.length < 4) return null;
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) return "image/png";
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return "image/jpeg";
  if (ascii(bytes, 0, 4) === "GIF8") return "image/gif";
  if (ascii(bytes, 0, 4) === "RIFF" && ascii(bytes, 8, 12) === "WEBP") return "image/webp";
  if (bytes[0] === 0x42 && bytes[1] === 0x4d) return "image/bmp";
  if ((bytes[0] === 0x49 && bytes[1] === 0x49 && bytes[2] === 0x2a && bytes[3] === 0) || (bytes[0] === 0x4d && bytes[1] === 0x4d && bytes[2] === 0 && bytes[3] === 0x2a)) return "image/tiff";
  if (bytes[0] === 0 && bytes[1] === 0 && bytes[2] === 1 && bytes[3] === 0) return "image/x-icon";
  if (ascii(bytes, 4, 8) === "ftyp" && /^avi[fs]/.test(ascii(bytes, 8, 12))) return "image/avif";
  if (/^\s*(?:﻿)?(?:<\?xml[^>]*>\s*)?(?:<!--[\s\S]*?-->\s*)*(?:<!DOCTYPE svg[^>]*>\s*)?<svg[\s>]/i.test(new TextDecoder().decode(bytes.subarray(0, 1024)))) return "image/svg+xml";
  return null;
}

export function isPortableImageType(type) {
  return PORTABLE_IMAGE_TYPES.includes(String(type || "").toLowerCase());
}

function fittedSize(width, height) {
  const scale = Math.min(1, MAX_DIMENSION / Math.max(1, width), MAX_DIMENSION / Math.max(1, height));
  return { width: Math.max(1, Math.round(width * scale)), height: Math.max(1, Math.round(height * scale)) };
}

async function loadSvg(blob) {
  const url = URL.createObjectURL(blob);
  try {
    // An <img> never runs SVG scripts or fetches remote content for a blob URL.
    const image = new Image();
    image.decoding = "async";
    image.src = url;
    await image.decode();
    const natural = { width: image.naturalWidth || 300, height: image.naturalHeight || 150 };
    const size = fittedSize(natural.width * SVG_SCALE, natural.height * SVG_SCALE);
    const bitmap = await createImageBitmap(image, { resizeWidth: size.width, resizeHeight: size.height, resizeQuality: "high" });
    return bitmap;
  } finally {
    URL.revokeObjectURL(url);
  }
}

function hasTransparency(context, width, height) {
  const data = context.getImageData(0, 0, width, height).data;
  const step = Math.max(1, Math.floor((width * height) / 250_000)) * 4;
  for (let index = 3; index < data.length; index += step) if (data[index] < 255) return true;
  return false;
}

/**
 * Browser codec: decode with createImageBitmap (SVG through an <img>), then
 * encode PNG when the picture has transparency and JPEG (quality 0.92) when
 * it does not. Pictures larger than 4096 px on a side are scaled down.
 */
export async function browserImageCodec(bytes, type) {
  const blob = new Blob([bytes], { type });
  const bitmap = type === "image/svg+xml" ? await loadSvg(blob) : await createImageBitmap(blob);
  try {
    const size = fittedSize(bitmap.width, bitmap.height);
    const canvas = new OffscreenCanvas(size.width, size.height);
    const context = canvas.getContext("2d", { willReadFrequently: true });
    context.drawImage(bitmap, 0, 0, size.width, size.height);
    const transparent = hasTransparency(context, size.width, size.height);
    const encoded = await canvas.convertToBlob(transparent ? { type: "image/png" } : { type: "image/jpeg", quality: 0.92 });
    return { bytes: new Uint8Array(await encoded.arrayBuffer()), type: encoded.type, width: size.width, height: size.height };
  } finally {
    bitmap.close?.();
  }
}

/** PNG/JPEG bytes stay as they are; any other picture is converted. */
export async function normalizeImageBytes(input, options = {}) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const detected = sniffImageType(bytes) || String(options.type || "").toLowerCase() || null;
  if (detected && isPortableImageType(detected)) return { bytes, type: detected, converted: false };
  if (!detected) throw new Error("This picture format is not recognized.");
  const codec = options.codec || browserImageCodec;
  const result = await codec(bytes, detected);
  const type = sniffImageType(result?.bytes);
  if (!type || !isPortableImageType(type)) throw new Error("The picture could not be converted to PNG or JPEG.");
  return { bytes: result.bytes, type, converted: true, from: detected };
}

async function bytesFrom(resolved) {
  if (resolved instanceof Uint8Array) return { bytes: resolved, type: "" };
  if (resolved instanceof ArrayBuffer) return { bytes: new Uint8Array(resolved), type: "" };
  if (resolved && typeof resolved.arrayBuffer === "function") return { bytes: new Uint8Array(await resolved.arrayBuffer()), type: String(resolved.type || "") };
  throw new Error("A document picture could not be read.");
}

/**
 * Copy a document model with every non-PNG/JPEG picture converted. `resolve`
 * returns the bytes (Blob, ArrayBuffer or Uint8Array) for an image src, and
 * `register` returns a src for converted bytes (for example an object URL).
 * The live document is never changed.
 */
export async function normalizeDocumentImages(document, { resolve, register, codec } = {}) {
  if (typeof resolve !== "function" || typeof register !== "function") throw new Error("Picture normalization needs resolve and register functions.");
  const copy = structuredClone(document);
  const cache = new Map();
  const result = { document: copy, converted: 0, failed: 0, warnings: [] };
  const failures = new Map();
  async function normalize(source) {
    try {
      const { bytes, type } = await bytesFrom(await resolve(source));
      const normalized = await normalizeImageBytes(bytes, { type, codec });
      if (!normalized.converted) return { source };
      return { source: await register(normalized.bytes, normalized.type), converted: true };
    } catch (error) {
      return { source, failed: true, reason: error instanceof Error ? error.message : String(error) };
    }
  }
  async function visit(value) {
    if (!value || typeof value !== "object") return;
    if (value.kind === "image" && typeof value.src === "string" && value.src) {
      if (!cache.has(value.src)) cache.set(value.src, normalize(value.src));
      const outcome = await cache.get(value.src);
      if (outcome.converted) { value.src = outcome.source; result.converted += 1; }
      if (outcome.failed) { result.failed += 1; failures.set(outcome.reason, (failures.get(outcome.reason) || 0) + 1); }
    }
    for (const child of Object.values(value)) if (child && typeof child === "object") await visit(child);
  }
  await visit(copy);
  if (result.failed) result.warnings.push(`${result.failed} ${result.failed === 1 ? "picture" : "pictures"} could not be converted to PNG or JPEG and may be missing from PDF and print output.`);
  return result;
}
