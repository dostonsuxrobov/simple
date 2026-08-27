import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { inflateRawSync } from "node:zlib";
import { runImport } from "@forevka/wordcanvas/import";

const FIXTURE_NAME = "simple-docs-roundtrip-fixture.docx";
const DEFAULT_BASELINE = new URL(`./fixtures/${FIXTURE_NAME}`, import.meta.url);
const SENTINELS = [
  "P1-TYPOGRAPHY-7402",
  "P2-TABLE-IMAGE-1846",
  "P3-PAGINATION-9031",
];
const HEADING_FIXTURE = [
  ["FixtureTitle", "Simple Docs Round-trip Fixture"],
  ["Heading1", "Typography and structure"],
  ["Heading2", "Lists"],
  ["Heading1", "Tables and images"],
  ["Heading2", "Embedded image"],
  ["Heading1", "Pagination, review, and export"],
  ["Heading2", "Search targets"],
];
const TABLE_FIXTURE = [
  ["Item", "Owner", "Status", "Value"],
  ["DOCX import", "Simple Docs", "Ready", "100"],
  ["Page layout", "Canvas engine", "Verified", "23"],
  ["Round-trip", "QA corpus", "Pending edit", "1"],
];
const LIST_FIXTURE = [
  ["A first bullet with enough words to wrap naturally across a narrow reading column.", "bullet", 0],
  ["A nested bullet", "bullet", 1],
  ["A third-level bullet", "bullet", 2],
  ["First numbered step", "decimal", 0],
  ["Second numbered step", "decimal", 0],
  ["Third numbered step", "decimal", 0],
];
const PAGE_BREAK_HEADINGS = ["Tables and images", "Pagination, review, and export"];
const EXPECTED_TABLE_FRACTIONS = [0.32, 0.26, 0.25, 0.17];
const MAX_ZIP_ENTRY_BYTES = 64 * 1024 * 1024;
const MAX_ZIP_TOTAL_BYTES = 256 * 1024 * 1024;

function usage() {
  return [
    "Usage:",
    "  node qa/verify-roundtrip.mjs <baseline.docx> <output.docx> [--expect-token=cobalt|violet]",
    "",
    "The baseline argument may be omitted only with --baseline-default:",
    `  node qa/verify-roundtrip.mjs --baseline-default <output.docx> [--expect-token=cobalt|violet]`,
    "",
    "Use cobalt for a no-op open/save round trip (default), or violet after the",
    "fixture's controlled bold-run edit.",
  ].join("\n");
}

function parseArgs(argv) {
  let expectedToken = "cobalt";
  let useDefaultBaseline = false;
  const paths = [];
  for (const argument of argv) {
    if (argument === "--help" || argument === "-h") return { help: true };
    if (argument === "--baseline-default") {
      useDefaultBaseline = true;
      continue;
    }
    if (argument.startsWith("--expect-token=")) {
      expectedToken = argument.slice("--expect-token=".length).trim().toLowerCase();
      continue;
    }
    if (argument.startsWith("--")) throw new Error(`Unknown option: ${argument}`);
    paths.push(argument);
  }
  if (!new Set(["cobalt", "violet"]).has(expectedToken)) {
    throw new Error("--expect-token must be cobalt or violet.");
  }
  if (useDefaultBaseline) {
    if (paths.length !== 1) throw new Error("--baseline-default expects exactly one output path.");
    return { baseline: DEFAULT_BASELINE, output: paths[0], expectedToken };
  }
  if (paths.length !== 2) throw new Error("Expected baseline and output DOCX paths.");
  return { baseline: paths[0], output: paths[1], expectedToken };
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function findEndOfCentralDirectory(bytes) {
  const minimum = Math.max(0, bytes.length - 22 - 0xffff);
  for (let offset = bytes.length - 22; offset >= minimum; offset -= 1) {
    if (bytes.readUInt32LE(offset) !== 0x06054b50) continue;
    const commentLength = bytes.readUInt16LE(offset + 20);
    if (offset + 22 + commentLength <= bytes.length) return offset;
  }
  throw new Error("ZIP end-of-central-directory record was not found.");
}

/** Read a normal, single-disk DOCX ZIP without relying on a transitive ZIP package. */
function unzipDocx(source) {
  const bytes = Buffer.from(source);
  const eocd = findEndOfCentralDirectory(bytes);
  const disk = bytes.readUInt16LE(eocd + 4);
  const centralDisk = bytes.readUInt16LE(eocd + 6);
  const entriesOnDisk = bytes.readUInt16LE(eocd + 8);
  const entryCount = bytes.readUInt16LE(eocd + 10);
  const centralSize = bytes.readUInt32LE(eocd + 12);
  const centralOffset = bytes.readUInt32LE(eocd + 16);
  if (disk !== 0 || centralDisk !== 0 || entriesOnDisk !== entryCount) {
    throw new Error("Multi-disk ZIP files are not supported.");
  }
  if (entryCount === 0xffff || centralSize === 0xffffffff || centralOffset === 0xffffffff) {
    throw new Error("ZIP64 DOCX files are not supported by this focused verifier.");
  }
  if (entryCount > 10_000 || centralOffset + centralSize > bytes.length) {
    throw new Error("ZIP central directory is outside safe bounds.");
  }

  const entries = new Map();
  let offset = centralOffset;
  let totalInflated = 0;
  for (let index = 0; index < entryCount; index += 1) {
    if (offset + 46 > bytes.length || bytes.readUInt32LE(offset) !== 0x02014b50) {
      throw new Error(`Invalid ZIP central-directory entry ${index}.`);
    }
    const flags = bytes.readUInt16LE(offset + 8);
    const method = bytes.readUInt16LE(offset + 10);
    const compressedSize = bytes.readUInt32LE(offset + 20);
    const uncompressedSize = bytes.readUInt32LE(offset + 24);
    const nameLength = bytes.readUInt16LE(offset + 28);
    const extraLength = bytes.readUInt16LE(offset + 30);
    const commentLength = bytes.readUInt16LE(offset + 32);
    const localOffset = bytes.readUInt32LE(offset + 42);
    const nameStart = offset + 46;
    const nameEnd = nameStart + nameLength;
    if (nameEnd + extraLength + commentLength > bytes.length) throw new Error("Truncated ZIP entry name.");
    const name = bytes.subarray(nameStart, nameEnd).toString("utf8").replaceAll("\\", "/");
    if (!name || name.startsWith("/") || name.split("/").includes("..")) {
      throw new Error(`Unsafe ZIP entry name: ${JSON.stringify(name)}`);
    }
    if (flags & 0x1) throw new Error(`Encrypted ZIP entry is unsupported: ${name}`);
    if (compressedSize === 0xffffffff || uncompressedSize === 0xffffffff) {
      throw new Error(`ZIP64 entry is unsupported: ${name}`);
    }
    if (uncompressedSize > MAX_ZIP_ENTRY_BYTES || totalInflated + uncompressedSize > MAX_ZIP_TOTAL_BYTES) {
      throw new Error(`ZIP inflation limit exceeded at: ${name}`);
    }
    if (localOffset + 30 > bytes.length || bytes.readUInt32LE(localOffset) !== 0x04034b50) {
      throw new Error(`Invalid local ZIP header: ${name}`);
    }
    const localNameLength = bytes.readUInt16LE(localOffset + 26);
    const localExtraLength = bytes.readUInt16LE(localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const dataEnd = dataStart + compressedSize;
    if (dataEnd > bytes.length) throw new Error(`Truncated ZIP entry data: ${name}`);
    const compressed = bytes.subarray(dataStart, dataEnd);
    let inflated;
    if (method === 0) inflated = Buffer.from(compressed);
    else if (method === 8) inflated = inflateRawSync(compressed);
    else throw new Error(`Unsupported ZIP compression method ${method} for ${name}.`);
    if (inflated.length !== uncompressedSize) throw new Error(`ZIP size mismatch for ${name}.`);
    if (entries.has(name)) throw new Error(`Duplicate ZIP entry: ${name}`);
    entries.set(name, inflated);
    totalInflated += inflated.length;
    offset = nameEnd + extraLength + commentLength;
  }
  return entries;
}

function decodeXml(value) {
  return value.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (entity, key) => {
    if (key[0] === "#") {
      const radix = key[1]?.toLowerCase() === "x" ? 16 : 10;
      const raw = radix === 16 ? key.slice(2) : key.slice(1);
      const codePoint = Number.parseInt(raw, radix);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : entity;
    }
    return { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" }[key.toLowerCase()] ?? entity;
  });
}

function escapeRegex(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function attribute(tag, name) {
  const match = tag.match(new RegExp(`(?:^|\\s)${escapeRegex(name)}=(['\"])(.*?)\\1`, "i"));
  return match ? decodeXml(match[2]) : undefined;
}

function startTags(xml, name) {
  return xml.match(new RegExp(`<${escapeRegex(name)}\\b[^>]*>`, "gi")) ?? [];
}

function fullElements(xml, name) {
  return xml.match(new RegExp(`<${escapeRegex(name)}\\b[^>]*>[\\s\\S]*?<\\/${escapeRegex(name)}>`, "gi")) ?? [];
}

function textNodes(xml, tag = "w:t") {
  const values = [];
  const regex = new RegExp(`<${escapeRegex(tag)}\\b[^>]*>([\\s\\S]*?)<\\/${escapeRegex(tag)}>`, "gi");
  for (const match of xml.matchAll(regex)) values.push(decodeXml(match[1].replace(/<[^>]+>/g, "")));
  return values;
}

function xmlPart(entries, name) {
  const bytes = entries.get(name);
  if (!bytes) throw new Error(`Required DOCX part is missing: ${name}`);
  return bytes.toString("utf8").replace(/^\uFEFF/, "");
}

function relationshipRecords(xml) {
  return startTags(xml, "Relationship").map((tag) => ({
    id: attribute(tag, "Id"),
    type: attribute(tag, "Type"),
    target: attribute(tag, "Target"),
    targetMode: attribute(tag, "TargetMode"),
  }));
}

function resolveRelationshipPart(sourcePart, target) {
  if (!target) return undefined;
  const normalizedTarget = target.replaceAll("\\", "/");
  if (normalizedTarget.startsWith("/")) return path.posix.normalize(normalizedTarget.slice(1));
  return path.posix.normalize(path.posix.join(path.posix.dirname(sourcePart), normalizedTarget));
}

function paragraphText(paragraph) {
  return paragraph.runs.map((run) => run.text).join("");
}

function blockText(block) {
  if (block.kind === "paragraph") return paragraphText(block);
  if (block.kind === "table") {
    return block.rows.map((row) => row.cells.map((cell) => cell.blocks.map(blockText).join("\n")).join("\t")).join("\n");
  }
  if (block.kind === "image") return `[image ${block.widthPx}x${block.heightPx}]`;
  return `[${block.kind}]`;
}

function bodyText(doc) {
  return doc.blocks.map(blockText).join("\n");
}

function findParagraph(doc, text) {
  return doc.blocks.find((block) => block.kind === "paragraph" && paragraphText(block) === text);
}

function cellText(cell) {
  return cell.blocks.map(blockText).join("\n");
}

function color(value) {
  return String(value ?? "").replace(/^#/, "").toLowerCase();
}

function selectedStyle(style) {
  const fontSizePx = Number(style?.fontSizePx);
  return {
    fontFamily: style?.fontFamily,
    fontSizePx: Number.isFinite(fontSizePx) ? Number(fontSizePx.toFixed(4)) : undefined,
    bold: Boolean(style?.bold),
    italic: Boolean(style?.italic),
    underline: Boolean(style?.underline),
    strikethrough: Boolean(style?.strikethrough),
    color: color(style?.color),
    link: style?.link,
  };
}

function paragraphSignature(block, doc) {
  const list = block.style.list;
  const level = list ? doc.lists?.[list.listId]?.levels[list.level] : undefined;
  return {
    kind: "paragraph",
    text: paragraphText(block),
    namedStyle: block.style.namedStyle,
    pageBreakBefore: Boolean(block.style.pageBreakBefore),
    list: list ? { level: list.level, format: level?.format, text: level?.text, bulletChar: level?.bulletChar } : undefined,
  };
}

function documentSignature(doc) {
  return doc.blocks.map((block) => {
    if (block.kind === "paragraph") return paragraphSignature(block, doc);
    if (block.kind === "table") {
      return {
        kind: "table",
        cells: block.rows.map((row) => row.cells.map(cellText)),
        colFractions: block.colFractions?.map((value) => Number(value.toFixed(6))),
        widthMode: block.widthMode ?? "fixed",
      };
    }
    if (block.kind === "image") {
      return { kind: "image", widthPx: block.widthPx, heightPx: block.heightPx, align: block.align, wrap: block.wrap ?? "block" };
    }
    return { kind: block.kind };
  });
}

function controlledSignature(baselineSignature, expectedToken) {
  const signature = structuredClone(baselineSignature);
  if (expectedToken !== "violet") return signature;
  const paragraph = signature.find((block) => block.kind === "paragraph" && block.text.startsWith("This line mixes bold cobalt,"));
  if (!paragraph) throw new Error("Baseline controlled-edit paragraph was not found.");
  paragraph.text = paragraph.text.replace("This line mixes bold cobalt,", "This line mixes bold violet,");
  return signature;
}

function imageHashes(inspected) {
  return inspected.importResult.media.map((item) => sha256(item.bytes)).sort();
}

function inspectDocx(bytes, label) {
  const entries = unzipDocx(bytes);
  const documentXml = xmlPart(entries, "word/document.xml");
  const stylesXml = xmlPart(entries, "word/styles.xml");
  const numberingXml = xmlPart(entries, "word/numbering.xml");
  const documentRelsXml = xmlPart(entries, "word/_rels/document.xml.rels");
  const rootRelsXml = xmlPart(entries, "_rels/.rels");
  const contentTypesXml = xmlPart(entries, "[Content_Types].xml");
  const importResult = runImport(new Uint8Array(bytes), undefined, { collectMediaBytes: true });
  return {
    label,
    bytes,
    entries,
    documentXml,
    stylesXml,
    numberingXml,
    contentTypesXml,
    documentRelationships: relationshipRecords(documentRelsXml),
    rootRelationships: relationshipRecords(rootRelsXml),
    importResult,
  };
}

function createAudit() {
  const passes = [];
  const failures = [];
  return {
    passes,
    failures,
    expect(scope, name, condition, details = undefined) {
      if (condition) passes.push(`${scope}: ${name}`);
      else failures.push({ scope, name, ...(details === undefined ? {} : { details }) });
    },
  };
}

function arraysEqual(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function near(left, right, tolerance = 0.002) {
  return Number.isFinite(left) && Number.isFinite(right) && Math.abs(left - right) <= tolerance;
}

function normalizedUrl(value) {
  try {
    return new URL(value).href;
  } catch {
    return value;
  }
}

function auditPackage(audit, inspected) {
  const scope = inspected.label;
  const required = [
    "[Content_Types].xml",
    "_rels/.rels",
    "word/document.xml",
    "word/styles.xml",
    "word/numbering.xml",
    "word/_rels/document.xml.rels",
  ];
  for (const name of required) audit.expect(scope, `package part ${name}`, inspected.entries.has(name));

  const rootDocument = inspected.rootRelationships.find((rel) => rel.type?.endsWith("/officeDocument"));
  audit.expect(scope, "root relationship resolves word/document.xml",
    resolveRelationshipPart("_rels/.rels", rootDocument?.target)?.replace(/^_rels\//, "") === "word/document.xml",
    rootDocument);

  const contentOverrides = startTags(inspected.contentTypesXml, "Override");
  const mainType = contentOverrides.find((tag) => attribute(tag, "PartName") === "/word/document.xml");
  audit.expect(scope, "main document content type",
    attribute(mainType ?? "", "ContentType")?.endsWith("wordprocessingml.document.main+xml") === true,
    mainType);
  audit.expect(scope, "numbering relationship",
    inspected.documentRelationships.some((rel) => rel.type?.endsWith("/numbering") && inspected.entries.has(resolveRelationshipPart("word/document.xml", rel.target))));
  audit.expect(scope, "styles relationship",
    inspected.documentRelationships.some((rel) => rel.type?.endsWith("/styles") && inspected.entries.has(resolveRelationshipPart("word/document.xml", rel.target))));
}

function auditSentinelsAndHeadings(audit, inspected, expectedToken) {
  const scope = inspected.label;
  const doc = inspected.importResult.doc;
  const text = bodyText(doc);
  for (const sentinel of SENTINELS) audit.expect(scope, `sentinel ${sentinel}`, text.includes(sentinel));

  const expectedRun = `bold ${expectedToken}`;
  audit.expect(scope, `controlled run text ${expectedRun}`,
    doc.blocks.some((block) => block.kind === "paragraph" && block.runs.some((run) => run.text === expectedRun)));

  const headingPairs = doc.blocks
    .filter((block) => block.kind === "paragraph" && block.style.namedStyle)
    .map((block) => [block.style.namedStyle, paragraphText(block)]);
  for (const pair of HEADING_FIXTURE) {
    audit.expect(scope, `heading ${pair[0]} / ${pair[1]}`,
      headingPairs.some(([style, textValue]) => style === pair[0] && textValue === pair[1]));
  }

  const declaredStyles = startTags(inspected.stylesXml, "w:style").map((tag) => ({
    id: attribute(tag, "w:styleId"),
    type: attribute(tag, "w:type"),
  }));
  for (const id of ["Normal", "FixtureTitle", "Heading1", "Heading2"]) {
    audit.expect(scope, `declared paragraph style ${id}`,
      declaredStyles.some((style) => style.id === id && style.type === "paragraph"));
  }
}

function auditStyledRuns(audit, baseline, output, expectedToken) {
  const scope = output.label;
  const baselineParagraph = baseline.importResult.doc.blocks.find((block) => block.kind === "paragraph" && paragraphText(block).startsWith("This line mixes bold cobalt,"));
  const outputParagraph = output.importResult.doc.blocks.find((block) => block.kind === "paragraph" && paragraphText(block).startsWith(`This line mixes bold ${expectedToken},`));
  audit.expect(scope, "mixed-style paragraph preserved", Boolean(outputParagraph));
  if (!baselineParagraph || !outputParagraph) return;

  const expectedTexts = baselineParagraph.runs.map((run) => run.text);
  const controlledRunIndex = expectedTexts.indexOf("bold cobalt");
  audit.expect(scope, "baseline contains controlled bold run", controlledRunIndex >= 0, {
    expected: "bold cobalt",
    actual: expectedTexts,
  });
  if (controlledRunIndex >= 0) expectedTexts[controlledRunIndex] = `bold ${expectedToken}`;
  audit.expect(scope, "mixed-style run boundaries", arraysEqual(outputParagraph.runs.map((run) => run.text), expectedTexts), {
    expected: expectedTexts,
    actual: outputParagraph.runs.map((run) => run.text),
  });
  for (let index = 0; index < baselineParagraph.runs.length; index += 1) {
    audit.expect(scope, `mixed-style run ${index + 1} formatting`,
      arraysEqual(selectedStyle(outputParagraph.runs[index]?.style ?? {}), selectedStyle(baselineParagraph.runs[index].style)), {
        expected: selectedStyle(baselineParagraph.runs[index].style),
        actual: selectedStyle(outputParagraph.runs[index]?.style ?? {}),
      });
  }
  const amber = outputParagraph.runs.find((run) => run.text === "italic amber");
  const green = outputParagraph.runs.find((run) => run.text === "underlined green");
  const bold = outputParagraph.runs.find((run) => run.text === `bold ${expectedToken}`);
  audit.expect(scope, "bold run remains bold", bold?.style.bold === true && bold.style.italic === false);
  audit.expect(scope, "amber run remains italic and colored", amber?.style.italic === true && color(amber.style.color) === "b45309");
  audit.expect(scope, "green run remains underlined and colored", green?.style.underline === true && color(green.style.color) === "166534");
}

function auditLists(audit, inspected) {
  const scope = inspected.label;
  const doc = inspected.importResult.doc;
  const found = [];
  for (const [text, expectedFormat, expectedLevel] of LIST_FIXTURE) {
    const paragraph = findParagraph(doc, text);
    const membership = paragraph?.style.list;
    const definition = membership ? doc.lists?.[membership.listId] : undefined;
    const level = membership ? definition?.levels[membership.level] : undefined;
    audit.expect(scope, `list item ${text}`,
      Boolean(paragraph && membership && membership.level === expectedLevel && level?.format === expectedFormat), {
        list: membership,
        level,
      });
    found.push({ text, listId: membership?.listId, level: membership?.level, format: level?.format });
  }
  const bulletIds = new Set(found.slice(0, 3).map((item) => item.listId));
  const numberIds = new Set(found.slice(3).map((item) => item.listId));
  audit.expect(scope, "nested bullets share one definition", bulletIds.size === 1 && !bulletIds.has(undefined));
  audit.expect(scope, "numbered steps share one distinct definition",
    numberIds.size === 1 && !numberIds.has(undefined) && !bulletIds.has([...numberIds][0]));
  audit.expect(scope, "numbering XML contains bullet and decimal formats",
    /<w:numFmt\b[^>]*w:val=["']bullet["']/.test(inspected.numberingXml)
      && /<w:numFmt\b[^>]*w:val=["']decimal["']/.test(inspected.numberingXml));
}

function auditTable(audit, inspected) {
  const scope = inspected.label;
  const tables = inspected.importResult.doc.blocks.filter((block) => block.kind === "table");
  audit.expect(scope, "one top-level table", tables.length === 1, { count: tables.length });
  const table = tables[0];
  if (!table) return;
  const matrix = table.rows.map((row) => row.cells.map(cellText));
  audit.expect(scope, "4x4 table content", arraysEqual(matrix, TABLE_FIXTURE), { expected: TABLE_FIXTURE, actual: matrix });
  audit.expect(scope, "fixed table column fractions",
    table.colFractions?.length === EXPECTED_TABLE_FRACTIONS.length
      && table.colFractions.every((value, index) => near(value, EXPECTED_TABLE_FRACTIONS[index])),
    table.colFractions);
  const headerStyled = table.rows[0]?.cells.every((cell) => (
    color(cell.shading) === "111111"
    && cell.blocks.every((block) => block.kind !== "paragraph" || block.runs.every((run) => run.style.bold && color(run.style.color) === "ffffff"))
  ));
  audit.expect(scope, "table header shading and run style", headerStyled === true);
  audit.expect(scope, "numeric column remains right aligned",
    table.rows.every((row) => row.cells[3]?.blocks.every((block) => block.kind !== "paragraph" || block.style.align === "right")));
  audit.expect(scope, "OOXML table grid has four columns", startTags(inspected.documentXml, "w:gridCol").length === 4);
}

function auditImage(audit, baseline, output) {
  const scope = output.label;
  const images = output.importResult.doc.blocks.filter((block) => block.kind === "image");
  audit.expect(scope, "one image block", images.length === 1, { count: images.length });
  const image = images[0];
  audit.expect(scope, "image geometry and alignment",
    image?.widthPx === 240 && image?.heightPx === 96 && image.align === "center" && (image.wrap ?? "block") === "block",
    image);
  audit.expect(scope, "embedded image bytes unchanged",
    arraysEqual(imageHashes(output), imageHashes(baseline)), {
      expected: imageHashes(baseline),
      actual: imageHashes(output),
    });

  const imageRelationships = output.documentRelationships.filter((rel) => rel.type?.endsWith("/image"));
  audit.expect(scope, "one image relationship", imageRelationships.length === 1, imageRelationships);
  for (const relation of imageRelationships) {
    const target = resolveRelationshipPart("word/document.xml", relation.target);
    audit.expect(scope, `image relationship target ${target}`, Boolean(target && output.entries.has(target)), relation);
  }
  const embedIds = startTags(output.documentXml, "a:blip").map((tag) => attribute(tag, "r:embed")).filter(Boolean);
  audit.expect(scope, "drawing references the image relationship",
    embedIds.some((id) => imageRelationships.some((rel) => rel.id === id)));

  const baselineExtent = startTags(baseline.documentXml, "wp:extent")[0];
  const outputExtent = startTags(output.documentXml, "wp:extent")[0];
  audit.expect(scope, "image OOXML extent unchanged",
    attribute(outputExtent ?? "", "cx") === attribute(baselineExtent ?? "", "cx")
      && attribute(outputExtent ?? "", "cy") === attribute(baselineExtent ?? "", "cy"), {
        expected: { cx: attribute(baselineExtent ?? "", "cx"), cy: attribute(baselineExtent ?? "", "cy") },
        actual: { cx: attribute(outputExtent ?? "", "cx"), cy: attribute(outputExtent ?? "", "cy") },
      });
}

function auditHyperlink(audit, inspected) {
  const scope = inspected.label;
  const hyperlinkRelationships = inspected.documentRelationships.filter((rel) => rel.type?.endsWith("/hyperlink"));
  const openAiRelation = hyperlinkRelationships.find((rel) => normalizedUrl(rel.target) === "https://openai.com/");
  audit.expect(scope, "external OpenAI hyperlink relationship",
    Boolean(openAiRelation && openAiRelation.targetMode?.toLowerCase() === "external"), openAiRelation);

  const hyperlinkElement = fullElements(inspected.documentXml, "w:hyperlink")
    .find((element) => textNodes(element).join("") === "OpenAI documentation");
  audit.expect(scope, "hyperlink text is wrapped by w:hyperlink",
    Boolean(hyperlinkElement && attribute(hyperlinkElement.match(/^<w:hyperlink\b[^>]*>/i)?.[0] ?? "", "r:id") === openAiRelation?.id));

  const paragraph = findParagraph(inspected.importResult.doc, "OpenAI documentation");
  audit.expect(scope, "imported hyperlink remains actionable",
    paragraph?.runs.some((run) => normalizedUrl(run.style.link) === "https://openai.com/") === true);
}

function auditPageBreaksAndBands(audit, inspected) {
  const scope = inspected.label;
  const doc = inspected.importResult.doc;
  const breakParagraphs = doc.blocks
    .filter((block) => block.kind === "paragraph" && block.style.pageBreakBefore)
    .map(paragraphText);
  audit.expect(scope, "two semantic page breaks", arraysEqual(breakParagraphs, PAGE_BREAK_HEADINGS), {
    expected: PAGE_BREAK_HEADINGS,
    actual: breakParagraphs,
  });
  const pageBreakBeforeCount = (inspected.documentXml.match(/<w:pageBreakBefore\b[^>]*\/?\s*>/gi) ?? []).length;
  const explicitPageBreakCount = (inspected.documentXml.match(/<w:br\b(?=[^>]*w:type=["']page["'])[^>]*\/?\s*>/gi) ?? []).length;
  audit.expect(scope, "two OOXML page-break boundaries", pageBreakBeforeCount + explicitPageBreakCount === 2, {
    pageBreakBeforeCount,
    explicitPageBreakCount,
  });

  const headerRelation = inspected.documentRelationships.find((rel) => rel.type?.endsWith("/header"));
  const footerRelation = inspected.documentRelationships.find((rel) => rel.type?.endsWith("/footer"));
  const headerPart = resolveRelationshipPart("word/document.xml", headerRelation?.target);
  const footerPart = resolveRelationshipPart("word/document.xml", footerRelation?.target);
  audit.expect(scope, "header relationship and part", Boolean(headerPart && inspected.entries.has(headerPart)), headerRelation);
  audit.expect(scope, "footer relationship and part", Boolean(footerPart && inspected.entries.has(footerPart)), footerRelation);
  if (headerPart && inspected.entries.has(headerPart)) {
    audit.expect(scope, "header sentinel text",
      textNodes(xmlPart(inspected.entries, headerPart)).join("").includes("SIMPLE DOCS · DOCX COMPATIBILITY FIXTURE"));
  }
  if (footerPart && inspected.entries.has(footerPart)) {
    const footerXml = xmlPart(inspected.entries, footerPart);
    const simpleInstructions = startTags(footerXml, "w:fldSimple")
      .map((tag) => attribute(tag, "w:instr"))
      .filter(Boolean);
    const instructions = [...textNodes(footerXml, "w:instrText"), ...simpleInstructions].join(" ").toUpperCase();
    audit.expect(scope, "footer PAGE and NUMPAGES fields", /\bPAGE\b/.test(instructions) && /\bNUMPAGES\b/.test(instructions), instructions);
  }
}

function auditComparison(audit, baseline, output, expectedToken) {
  const expectedSignature = controlledSignature(documentSignature(baseline.importResult.doc), expectedToken);
  const actualSignature = documentSignature(output.importResult.doc);
  audit.expect("comparison", "top-level document sequence and text",
    arraysEqual(actualSignature, expectedSignature), { expected: expectedSignature, actual: actualSignature });

  const baselineWarningCodes = new Set(baseline.importResult.warnings.map((warning) => warning.code));
  const newWarnings = output.importResult.warnings.filter((warning) => !baselineWarningCodes.has(warning.code));
  audit.expect("comparison", "no new WordCanvas import warnings", newWarnings.length === 0, newWarnings);

  const baselineStyles = new Map((baseline.importResult.doc.stylesheet?.styles ?? []).map((style) => [style.id, style]));
  const outputStyles = new Map((output.importResult.doc.stylesheet?.styles ?? []).map((style) => [style.id, style]));
  for (const id of ["Normal", "FixtureTitle", "Heading1", "Heading2"]) {
    const before = baselineStyles.get(id);
    const after = outputStyles.get(id);
    audit.expect("comparison", `named style ${id} remains equivalent`,
      Boolean(before && after && arraysEqual(before, after)), { expected: before, actual: after });
  }
}

async function main() {
  let options;
  try {
    options = parseArgs(process.argv.slice(2));
  } catch (error) {
    console.error(`${error.message}\n\n${usage()}`);
    process.exitCode = 2;
    return;
  }
  if (options.help) {
    console.log(usage());
    return;
  }

  const baselinePath = options.baseline instanceof URL ? options.baseline : path.resolve(options.baseline);
  const outputPath = path.resolve(options.output);
  try {
    const [baselineBytes, outputBytes] = await Promise.all([readFile(baselinePath), readFile(outputPath)]);
    const baseline = inspectDocx(baselineBytes, "baseline");
    const output = inspectDocx(outputBytes, "output");
    const audit = createAudit();

    auditPackage(audit, baseline);
    auditPackage(audit, output);
    auditSentinelsAndHeadings(audit, baseline, "cobalt");
    auditSentinelsAndHeadings(audit, output, options.expectedToken);
    auditStyledRuns(audit, baseline, output, options.expectedToken);
    auditLists(audit, baseline);
    auditLists(audit, output);
    auditTable(audit, baseline);
    auditTable(audit, output);
    auditImage(audit, baseline, output);
    auditHyperlink(audit, baseline);
    auditHyperlink(audit, output);
    auditPageBreaksAndBands(audit, baseline);
    auditPageBreaksAndBands(audit, output);
    auditComparison(audit, baseline, output, options.expectedToken);

    const report = {
      ok: audit.failures.length === 0,
      expectedToken: options.expectedToken,
      baseline: {
        path: baselinePath instanceof URL ? baselinePath.pathname : baselinePath,
        bytes: baselineBytes.length,
        sha256: sha256(baselineBytes),
        importWarnings: baseline.importResult.warnings,
      },
      output: {
        path: outputPath,
        bytes: outputBytes.length,
        sha256: sha256(outputBytes),
        importWarnings: output.importResult.warnings,
      },
      passedChecks: audit.passes.length,
      failures: audit.failures,
    };
    console.log(JSON.stringify(report, null, 2));
    if (!report.ok) process.exitCode = 1;
  } catch (error) {
    console.error(JSON.stringify({
      ok: false,
      error: error instanceof Error ? error.message : String(error),
    }, null, 2));
    process.exitCode = 1;
  }
}

await main();
