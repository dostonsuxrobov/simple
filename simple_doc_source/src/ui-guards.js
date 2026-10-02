export function documentExportDisabled(documentOpen, busy) {
  return !documentOpen || busy;
}

export function activeModal(root) {
  return root.querySelector(".modal-backdrop:not([hidden])");
}

/**
 * The Latin letter of a Ctrl shortcut on any keyboard layout: the typed letter
 * on Latin layouts (AZERTY, Dvorak), the physical key on Cyrillic and others.
 */
function shortcutLetter(event) {
  const key = String(event?.key ?? "");
  if (/^[a-z]$/i.test(key)) return key.toLowerCase();
  const code = /^Key([A-Z])$/.exec(String(event?.code ?? ""));
  return code ? code[1].toLowerCase() : null;
}

/**
 * Window and document shortcuts that must not act behind an open dialog:
 * save, open, new, print, close, find and export, and F12 (Save as). Alt is
 * never blocked, because Ctrl+Alt is AltGr typing a character on many layouts.
 */
export function isModalBlockedShortcut(event) {
  if (event?.key === "F12" && !event.ctrlKey && !event.metaKey && !event.altKey) return true;
  const modifier = Boolean(event?.ctrlKey || event?.metaKey);
  if (!modifier || event.altKey) return false;
  const letter = shortcutLetter(event);
  return ["s", "o", "n", "p", "w", "f"].includes(letter) || (Boolean(event.shiftKey) && letter === "e");
}

export function nextPrintPreviewGeneration(current) {
  return Number.isSafeInteger(current) && current >= 0 && current < Number.MAX_SAFE_INTEGER
    ? current + 1
    : 1;
}

export function canCommitPrintPreview(requestGeneration, currentGeneration, modalHidden) {
  return Number.isSafeInteger(requestGeneration)
    && requestGeneration === currentGeneration
    && modalHidden === false;
}

// ---------------------------------------------------------------------------
// Unsaved-work guards. Pure helpers for src/main.ts, kept here so Node tests
// can check them without a browser.

const EMPTY_REVIEW_KEY = "[]";

/**
 * A stable key for the review content a save would have to carry (pending
 * suggestions and comment threads). Bookkeeping fields such as the rebase
 * version are ignored, so only real review changes change the key.
 */
export function reviewContentKey(review) {
  const suggestions = Array.isArray(review?.suggestions) ? review.suggestions : [];
  const threads = Array.isArray(review?.threads) ? review.threads : [];
  return suggestions.length || threads.length ? JSON.stringify([suggestions, threads]) : EMPTY_REVIEW_KEY;
}

/** How many review items exist; every thread counts as one comment. */
export function reviewSummary(review) {
  const suggestions = Array.isArray(review?.suggestions) ? review.suggestions.length : 0;
  const comments = Array.isArray(review?.threads) ? review.threads.length : 0;
  return { suggestions, comments, total: suggestions + comments };
}

function countLabel(count, one, many) {
  return `${count} ${count === 1 ? one : many}`;
}

function joinWords(parts) {
  if (parts.length <= 1) return parts.join("");
  return `${parts.slice(0, -1).join(", ")} and ${parts[parts.length - 1]}`;
}

/** "2 suggestions and 1 comment" (empty when there is nothing). */
export function describeReviewItems(summary) {
  const parts = [];
  if (summary?.suggestions) parts.push(countLabel(summary.suggestions, "suggestion", "suggestions"));
  if (summary?.comments) parts.push(countLabel(summary.comments, "comment", "comments"));
  return joinWords(parts);
}

function mergeRuns(runs) {
  if (runs.length < 2 || !runs.every((run) => run && typeof run === "object" && typeof run.text === "string")) return runs;
  const merged = [];
  let previousKey = null;
  for (const run of runs) {
    const { text, ...properties } = run;
    const key = JSON.stringify(properties, (name, value) => (name === "revision" ? undefined : value));
    if (merged.length && key === previousKey) {
      const last = merged[merged.length - 1];
      merged[merged.length - 1] = { ...last, text: last.text + text };
      continue;
    }
    merged.push(run);
    previousKey = key;
  }
  const visible = merged.filter((run) => run.text !== "");
  return visible.length ? visible : merged.slice(0, 1);
}

/**
 * Content identity of a WordCanvas document model. Undo creates new model
 * objects with bumped layout `revision` counters and may split runs, so two
 * models with the same content can differ in those details; this key ignores
 * them and nothing else.
 */
export function documentContentKey(model) {
  return JSON.stringify(model, (name, value) => {
    if (name === "revision") return undefined;
    if (name === "runs" && Array.isArray(value)) return mergeRuns(value);
    return value;
  });
}

// ---------------------------------------------------------------------------
// Open-time fidelity scan: content in a .docx that the editor drops when it
// imports the file, so saving over the original would remove it. Calibrated
// against WordCanvas 0.12.0's importer and exporter (see tests).

const ZIP_END_OF_DIRECTORY = 0x06054b50;
const ZIP_DIRECTORY_ENTRY = 0x02014b50;
const ZIP_LOCAL_ENTRY = 0x04034b50;
const MAX_SCANNED_PART_BYTES = 96 * 1024 * 1024;
// Picture formats the editor can show; anything else is skipped on import.
const SUPPORTED_PICTURE_EXTENSIONS = new Set(["png", "jpg", "jpeg", "gif", "bmp", "webp", "svg"]);
// Fields the importer keeps live in the main text; others become plain text.
const KEPT_BODY_FIELDS = new Set(["PAGE", "NUMPAGES", "DATE", "TIME", "IF"]);

function asBytes(data) {
  if (data instanceof Uint8Array) return data;
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  throw new TypeError("Expected document bytes.");
}

/** Central-directory entries of a ZIP package, keyed by name. ZIP64 is not scanned. */
export function listZipEntries(data) {
  const bytes = asBytes(data);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 65_557); offset -= 1) {
    if (view.getUint32(offset, true) === ZIP_END_OF_DIRECTORY) {
      end = offset;
      break;
    }
  }
  if (end < 0) throw new Error("This file is not a Word package.");
  const count = view.getUint16(end + 10, true);
  const directoryOffset = view.getUint32(end + 16, true);
  if (count === 0xffff || directoryOffset === 0xffffffff) throw new Error("Large ZIP64 packages are not scanned.");
  const decoder = new TextDecoder();
  const entries = new Map();
  let cursor = directoryOffset;
  for (let index = 0; index < count; index += 1) {
    if (cursor + 46 > bytes.length || view.getUint32(cursor, true) !== ZIP_DIRECTORY_ENTRY) throw new Error("This Word package has a damaged file list.");
    const nameLength = view.getUint16(cursor + 28, true);
    const name = decoder.decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength)).replace(/\\/g, "/");
    entries.set(name, {
      name,
      method: view.getUint16(cursor + 10, true),
      compressedSize: view.getUint32(cursor + 20, true),
      size: view.getUint32(cursor + 24, true),
      localOffset: view.getUint32(cursor + 42, true),
    });
    cursor += 46 + nameLength + view.getUint16(cursor + 30, true) + view.getUint16(cursor + 32, true);
  }
  return entries;
}

function decodeText(bytes) {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder("utf-16le").decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder("utf-16be").decode(bytes.subarray(2));
  return new TextDecoder().decode(bytes);
}

/** The text of one package part, or null when it is missing, too large or unreadable. */
export async function readZipText(data, entry) {
  if (!entry || entry.size > MAX_SCANNED_PART_BYTES || entry.compressedSize > MAX_SCANNED_PART_BYTES) return null;
  const bytes = asBytes(data);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const at = entry.localOffset;
  if (at + 30 > bytes.length || view.getUint32(at, true) !== ZIP_LOCAL_ENTRY) return null;
  const start = at + 30 + view.getUint16(at + 26, true) + view.getUint16(at + 28, true);
  if (start + entry.compressedSize > bytes.length) return null;
  const stored = bytes.subarray(start, start + entry.compressedSize);
  if (entry.method === 0) return decodeText(stored);
  if (entry.method !== 8) return null;
  try {
    const stream = new Blob([stored]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
    return decodeText(new Uint8Array(await new Response(stream).arrayBuffer()));
  } catch {
    return null;
  }
}

function attributeValue(source, name) {
  const match = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(["'])([\\s\\S]*?)\\1`).exec(source);
  return match ? match[2] : undefined;
}

function decodeEntities(value) {
  return String(value).replace(/&(amp|lt|gt|quot|apos);/g, (_entity, name) => ({ amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" })[name]);
}

/** Relationships of a part: id -> { type, target, external }. */
export function parseRelationships(xml) {
  const relationships = new Map();
  for (const match of String(xml || "").matchAll(/<(?:[\w.-]+:)?Relationship\b([^>]*)>/g)) {
    const id = attributeValue(match[1], "Id");
    if (!id) continue;
    relationships.set(id, {
      type: attributeValue(match[1], "Type") ?? "",
      target: decodeEntities(attributeValue(match[1], "Target") ?? ""),
      external: (attributeValue(match[1], "TargetMode") ?? "").toLowerCase() === "external",
    });
  }
  return relationships;
}

function resolvePartPath(baseDirectory, target) {
  const segments = (target.startsWith("/") ? target.slice(1) : `${baseDirectory}${target}`).split("/");
  const resolved = [];
  for (const segment of segments) {
    if (!segment || segment === ".") continue;
    if (segment === "..") resolved.pop();
    else resolved.push(segment);
  }
  return resolved.join("/");
}

function relationshipsPathFor(partName) {
  const slash = partName.lastIndexOf("/");
  return `${partName.slice(0, slash + 1)}_rels/${partName.slice(slash + 1)}.rels`;
}

const STORY_RELATIONSHIPS = [
  [/\/header$/i, "header"],
  [/\/footer$/i, "footer"],
  [/\/footnotes$/i, "notes"],
  [/\/endnotes$/i, "notes"],
  [/\/comments$/i, "comments"],
];

/**
 * Lists what in a .docx the editor cannot keep. Resolves to
 * `{ items: [{ kind, count, label }] }`; an unreadable package yields no items.
 */
export async function scanDocxFidelity(data) {
  const entries = listZipEntries(data);
  const read = (name) => readZipText(data, entries.get(name));
  const packageRelationships = parseRelationships(await read("_rels/.rels"));
  const office = [...packageRelationships.values()].find((relationship) => /\/officeDocument$/i.test(relationship.type) && !relationship.external);
  const mainPart = resolvePartPath("", office?.target || "word/document.xml");
  const mainDirectory = mainPart.slice(0, mainPart.lastIndexOf("/") + 1);
  const mainRelationships = parseRelationships(await read(relationshipsPathFor(mainPart)));
  const parts = [{ name: mainPart, kind: "body", xml: await read(mainPart), relationships: mainRelationships }];
  const seen = new Set([mainPart]);
  for (const relationship of mainRelationships.values()) {
    if (relationship.external) continue;
    const kind = STORY_RELATIONSHIPS.find(([pattern]) => pattern.test(relationship.type))?.[1];
    const name = resolvePartPath(mainDirectory, relationship.target);
    if (!kind || seen.has(name)) continue;
    seen.add(name);
    parts.push({ name, kind, xml: await read(name), relationships: parseRelationships(await read(relationshipsPathFor(name))) });
  }
  return analyzeDocxParts(parts);
}

const FALLBACK_BLOCK = /<((?:[\w.-]+:)?)Fallback\b[^>]*?(?:\/>|>[\s\S]*?<\/\1Fallback>)/g;
const OBJECT_BLOCK = /<((?:[\w.-]+:)?)object\b[^>]*?(?:\/>|>[\s\S]*?<\/\1object>)/g;
const DELETED_BLOCK = /<((?:[\w.-]+:)?)del\b[^>]*?(?:\/>|>[\s\S]*?<\/\1del>)/g;
const COMMENT_ELEMENT = /<(?:[\w.-]+:)?comment(?=[\s>/])/g;
const TRACKED_CHANGE = /<(?:[\w.-]+:)?(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|tblGridChange|tcPrChange|trPrChange|numberingChange)(?=[\s>/])/g;
const OBJECT_ELEMENT = /<(?:[\w.-]+:)?object(?=[\s>/])/g;
const EMBEDDED_CONTENT = /<(?:[\w.-]+:)?(?:altChunk|control)(?=[\s>/])/g;
const INK = /<(?:[\w.-]+:)?contentPart(?=[\s>/])/g;
const GRAPHIC_DATA = /<(?:[\w.-]+:)?graphicData\b[^>]*?\suri\s*=\s*["']([^"']+)["']/g;
const BLIP = /<(?:[\w.-]+:)?blip\b([^>]*)>/g;
const VML_IMAGE = /<(?:[\w.-]+:)?imagedata\b([^>]*)>/g;
const WORDART = /<(?:[\w.-]+:)?textpath\b[^>]*\sstring\s*=/g;
const FIELD_TOKEN = /<(?:[\w.-]+:)?fldChar\b([^>]*)>|<(?:[\w.-]+:)?instrText(?:\s[^>]*)?(?<!\/)>([\s\S]*?)<\/(?:[\w.-]+:)?instrText>|<(?:[\w.-]+:)?fldSimple\b([^>]*)>/g;

function countMatches(xml, pattern) {
  pattern.lastIndex = 0;
  let count = 0;
  while (pattern.exec(xml)) count += 1;
  return count;
}

function fieldIsKept(instruction, kind, simple) {
  if (kind === "header" || kind === "footer") return /\bNUMPAGES\b/.test(instruction) || /\bPAGE\b/.test(instruction);
  if (kind !== "body" || simple) return false;
  return KEPT_BODY_FIELDS.has(instruction.trim().split(/\s+/)[0]?.toUpperCase() ?? "");
}

/** Top-level fields the importer turns into plain text. */
function countDroppedFields(xml, kind) {
  let dropped = 0;
  let depth = 0;
  let instruction = "";
  let open = false;
  const settle = () => {
    if (!open) return;
    open = false;
    if (!fieldIsKept(decodeEntities(instruction), kind, false)) dropped += 1;
  };
  for (const match of xml.matchAll(FIELD_TOKEN)) {
    if (match[1] !== undefined) {
      const type = attributeValue(match[1], "(?:[\\w.-]+:)?fldCharType");
      if (type === "begin") {
        depth += 1;
        if (depth === 1) {
          instruction = "";
          open = true;
        }
      } else if (type === "separate") {
        if (depth === 1) settle();
      } else if (type === "end") {
        if (depth === 1) settle();
        depth = Math.max(0, depth - 1);
      }
    } else if (match[2] !== undefined) {
      if (depth === 1 && open) instruction += match[2];
    } else if (depth === 0 && !fieldIsKept(decodeEntities(attributeValue(match[3], "(?:[\\w.-]+:)?instr") ?? ""), kind, true)) {
      dropped += 1;
    }
  }
  return dropped;
}

function pictureExtension(target) {
  const clean = String(target).split(/[?#]/)[0];
  const dot = clean.lastIndexOf(".");
  return dot >= 0 && dot > clean.lastIndexOf("/") ? clean.slice(dot + 1).toLowerCase() : "";
}

const FIDELITY_ORDER = ["comments", "trackedChanges", "charts", "smartArt", "drawings", "unsupportedPictures", "linkedPictures", "embeddedObjects", "wordArt", "fields"];

function fidelityLabel(kind, count, formats) {
  switch (kind) {
    case "comments": return countLabel(count, "comment", "comments");
    case "trackedChanges": return "tracked changes (shown as accepted)";
    case "charts": return countLabel(count, "chart", "charts");
    case "smartArt": return countLabel(count, "SmartArt graphic", "SmartArt graphics");
    case "drawings": return countLabel(count, "drawing", "drawings");
    case "unsupportedPictures": return `${countLabel(count, "picture", "pictures")} in ${joinWords(formats)} format`;
    case "linkedPictures": return countLabel(count, "linked picture", "linked pictures");
    case "embeddedObjects": return countLabel(count, "embedded object", "embedded objects");
    case "wordArt": return countLabel(count, "WordArt object", "WordArt objects");
    case "fields": return `${countLabel(count, "field", "fields")} (kept as plain text)`;
    default: return countLabel(count, kind, kind);
  }
}

/**
 * Pure analysis of package parts: `[{ kind: "body" | "header" | "footer" |
 * "notes" | "comments", xml, relationships }]`.
 */
export function analyzeDocxParts(parts) {
  const counts = Object.fromEntries(FIDELITY_ORDER.map((kind) => [kind, 0]));
  const formats = new Set();
  for (const part of parts || []) {
    if (typeof part?.xml !== "string" || !part.xml) continue;
    if (part.kind === "comments") {
      counts.comments += countMatches(part.xml, COMMENT_ELEMENT);
      continue;
    }
    const relationships = part.relationships instanceof Map ? part.relationships : new Map();
    // The importer reads mc:Choice and never the mc:Fallback copy.
    let xml = part.xml.replace(FALLBACK_BLOCK, "");
    counts.trackedChanges += countMatches(xml, TRACKED_CHANGE);
    // Embedded (OLE) objects are dropped with their preview pictures.
    counts.embeddedObjects += countMatches(xml, OBJECT_ELEMENT);
    xml = xml.replace(OBJECT_BLOCK, "");
    counts.embeddedObjects += countMatches(xml, EMBEDDED_CONTENT);
    for (const match of xml.matchAll(GRAPHIC_DATA)) {
      const uri = match[1].trim().toLowerCase();
      if (/\/(?:drawingml\/2006\/picture|wordprocessingshape|wordprocessinggroup)$/.test(uri)) continue;
      if (/\/(?:chart|chartex)$/.test(uri)) counts.charts += 1;
      else if (/\/diagram$/.test(uri)) counts.smartArt += 1;
      else counts.drawings += 1;
    }
    counts.drawings += countMatches(xml, INK);
    const picture = (attributes, idAttribute) => {
      const embedded = attributeValue(attributes, `(?:[\\w.-]+:)?${idAttribute}`);
      const linked = attributeValue(attributes, "(?:[\\w.-]+:)?link");
      const relationship = relationships.get(embedded ?? linked ?? "");
      if (!relationship) return;
      if (relationship.external || (!embedded && linked)) {
        if (!/^https?:\/\//i.test(relationship.target)) counts.linkedPictures += 1;
        return;
      }
      const extension = pictureExtension(relationship.target);
      if (extension && !SUPPORTED_PICTURE_EXTENSIONS.has(extension)) {
        counts.unsupportedPictures += 1;
        formats.add(extension.toUpperCase());
      }
    };
    for (const match of xml.matchAll(BLIP)) picture(match[1], "embed");
    for (const match of xml.matchAll(VML_IMAGE)) picture(match[1], "id");
    counts.wordArt += countMatches(xml, WORDART);
    // Deleted revisions are already counted as tracked changes.
    counts.fields += countDroppedFields(xml.replace(DELETED_BLOCK, ""), part.kind);
  }
  const sortedFormats = [...formats].sort();
  const items = FIDELITY_ORDER
    .filter((kind) => counts[kind] > 0)
    .map((kind) => ({ kind, count: counts[kind], label: fidelityLabel(kind, counts[kind], sortedFormats) }));
  return { items };
}

/** "2 comments, 1 chart and tracked changes (shown as accepted)". */
export function describeFidelityItems(items) {
  return joinWords((items || []).map((item) => item.label));
}

// ---------------------------------------------------------------------------
// Messages and names.

const IPC_ERROR_PREFIX = /^Error invoking remote method '[^']*':\s*(?:[A-Za-z]*Error:\s*)?/;

/** The readable part of an error, without Electron's IPC wrapper. */
export function cleanErrorMessage(error, fallback = "Something went wrong.") {
  const raw = error instanceof Error ? error.message : typeof error === "string" ? error : "";
  return raw.replace(IPC_ERROR_PREFIX, "").trim() || fallback;
}

/** True only when the file itself is gone (moved, renamed or deleted). */
export function isMissingFileError(error) {
  const message = error instanceof Error ? error.message : String(error ?? "");
  return /\bENOENT\b|no such file or directory|cannot find the (?:file|path)/i.test(message);
}

/** A plain sentence for a failed open. */
export function openErrorMessage(error, fallback = "This document could not be opened.") {
  if (isMissingFileError(error)) return "That file is no longer available. It may have been moved, renamed or deleted.";
  const message = cleanErrorMessage(error, fallback);
  if (/\b(?:EBUSY|EPERM|EACCES)\b/.test(message)) return "Simple Docs can't read that file. Another program may be using it, or you may not have permission to open it.";
  return message;
}

/** The title of recovered work: an untitled copy, never "name.doc.docx". */
export function recoveredDocumentName(name) {
  const stem = String(name ?? "")
    .replace(/(?:\.(?:docx|docm|dotx|dotm|doc|dot))+$/i, "")
    .replace(/\s*\(recovered(?: \d+)?\)$/i, "")
    .trim();
  return `${stem || "Untitled document"} (recovered)`;
}

/** How long a toast stays: longer for long messages and errors. */
export function notificationDuration(message, tone = "normal") {
  const extra = Math.max(0, String(message ?? "").length - 60) * 45;
  return Math.min(12_000, Math.max(tone === "error" ? 6_000 : 3_200, 3_200 + extra));
}
