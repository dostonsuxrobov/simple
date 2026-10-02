/**
 * Markdown (.md) import: CommonMark plus GFM tables and strikethrough (markdown-it),
 * footnotes ([^1] and [^1]: ...), task lists (as ☐/☒) and YAML front matter (its title).
 * The rendered HTML goes through the same mapper as HTML files.
 */
import MarkdownIt from "markdown-it";
import { htmlToDocument } from "./html.ts";
import type { ImportOptions, ImportResult } from "./model.ts";
import { decodeTextBytes, toBytes } from "./text-decoding.ts";

export interface MarkdownImportOptions extends ImportOptions {
  /** Force a text encoding label. */
  encoding?: string;
}

const FENCE = /^ {0,3}(`{3,}|~{3,})/;
const DEFINITION = /^ {0,3}\[\^([^\]\s]+)\]:[ \t]?(.*)$/;
const TASK_ITEM = /^(\s*(?:[-*+]|\d{1,9}[.)])\s+)\[([ xX])\](?=\s)/;

let renderer: MarkdownIt | null = null;

const escapeAttribute = (value: string) => value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

function markdown(): MarkdownIt {
  if (renderer) return renderer;
  const md = new MarkdownIt("default", { html: true, linkify: true, typographer: false, breaks: false });
  md.inline.ruler.before("link", "simple_footnote_ref", (state, silent) => {
    if (state.src.charCodeAt(state.pos) !== 0x5b || state.src.charCodeAt(state.pos + 1) !== 0x5e) return false;
    const match = /^\[\^([^\]\s]+)\]/.exec(state.src.slice(state.pos, Math.min(state.posMax, state.pos + 200)));
    const notes: Map<string, string> | undefined = state.env?.simpleFootnotes;
    if (!match || !notes?.has(match[1])) return false;
    if (!silent) {
      const token = state.push("simple_footnote_ref", "sup", 0);
      token.meta = { label: match[1] };
    }
    state.pos += match[0].length;
    return true;
  });
  md.renderer.rules.simple_footnote_ref = (tokens, index) => `<sup data-simple-footnote="${escapeAttribute(String(tokens[index].meta?.label ?? ""))}">*</sup>`;
  renderer = md;
  return md;
}

interface Prepared {
  body: string;
  notes: Map<string, string>;
  title?: string;
}

/** Pulls out front matter and footnote definitions; turns task-list boxes into symbols. */
function prepare(source: string): Prepared {
  let text = source.replace(/^\ufeff/, "").replace(/\r\n?/g, "\n");
  let title: string | undefined;
  const front = /^---\n([\s\S]*?)\n(?:---|\.\.\.)\n/.exec(text);
  if (front && /^[\w-]+\s*:/m.test(front[1])) {
    const match = /^title\s*:\s*(.+)$/m.exec(front[1]);
    if (match) title = match[1].trim().replace(/^(["'])(.*)\1$/, "$2") || undefined;
    text = text.slice(front[0].length);
  }
  const lines = text.split("\n");
  const kept: string[] = [];
  const notes = new Map<string, string>();
  let fence: string | null = null;
  for (let i = 0; i < lines.length; i += 1) {
    const line = lines[i];
    const fenceMatch = FENCE.exec(line);
    if (fence) {
      if (fenceMatch && fenceMatch[1][0] === fence[0] && fenceMatch[1].length >= fence.length) fence = null;
      kept.push(line);
      continue;
    }
    if (fenceMatch) {
      fence = fenceMatch[1];
      kept.push(line);
      continue;
    }
    const definition = DEFINITION.exec(line);
    if (definition) {
      const body = [definition[2]];
      // Continuation: indented lines, possibly separated by blank lines.
      while (i + 1 < lines.length) {
        const next = lines[i + 1];
        if (/^(?: {4}|\t)/.test(next)) body.push(next.replace(/^(?: {4}|\t)/, ""));
        else if (!next.trim() && i + 2 < lines.length && /^(?: {4}|\t)/.test(lines[i + 2])) body.push("");
        else break;
        i += 1;
      }
      if (!notes.has(definition[1])) notes.set(definition[1], body.join("\n"));
      continue;
    }
    kept.push(line.replace(TASK_ITEM, (_all, prefix: string, mark: string) => `${prefix}${mark === " " ? "☐" : "☒"}`));
  }
  return { body: kept.join("\n"), notes, title };
}

/** Renders Markdown to the HTML the mapper reads (footnotes as data-simple-* hints). */
export function markdownToHtml(source: string): { html: string; title?: string } {
  const md = markdown();
  const { body, notes, title } = prepare(source);
  const env = { simpleFootnotes: notes };
  let html = md.render(body, env);
  for (const [label, text] of notes) html += `\n<aside data-simple-footnote-body="${escapeAttribute(label)}">${md.render(text, env)}</aside>`;
  return { html, title };
}

export function markdownToDocument(source: string, options: ImportOptions = {}): ImportResult {
  const { html, title } = markdownToHtml(source);
  const result = htmlToDocument(html, options, { trusted: true, authorStyles: false });
  result.format = "md";
  if (title) result.title = title;
  else delete result.title;
  return result;
}

export function importMarkdown(input: Uint8Array | ArrayBuffer | string, options: MarkdownImportOptions = {}): ImportResult {
  if (typeof input === "string") return markdownToDocument(input, options);
  const decoded = decodeTextBytes(toBytes(input), { encoding: options.encoding });
  const result = markdownToDocument(decoded.text, options);
  result.encoding = decoded.encoding;
  return result;
}
