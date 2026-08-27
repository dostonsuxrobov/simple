import { mkdir, writeFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { DocumentBuilder, inches, pt } from "@forevka/wordcanvas/builder";
import { installMeasureHost } from "@forevka/wordcanvas/export/measure";
import { runExport } from "@forevka/wordcanvas/export";

const outputDirectory = fileURLToPath(new URL("./fixtures/", import.meta.url));
const outputPath = path.join(outputDirectory, "simple-docs-roundtrip-fixture.docx");
const pixelBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAAG0lEQVR4nGP4z8DAwMDAxMDAwMDAwMDAwMAAAO0AAf6S/NsAAAAASUVORK5CYII=";
const pixel = Uint8Array.from(Buffer.from(
  pixelBase64,
  "base64",
));

const builder = DocumentBuilder.create({
  pageSize: "Letter",
  margins: { top: inches(0.8), right: inches(0.85), bottom: inches(0.8), left: inches(0.85) },
  idSeed: "simple-docs-qa",
});

builder
  .style({
    id: "Normal",
    name: "Normal",
    type: "paragraph",
    char: { fontFamily: "Arial", fontSizePx: pt(11), color: "#202124" },
    para: { align: "left", lineHeight: 1.2, spaceAfterPx: pt(7) },
  })
  .style({
    id: "FixtureTitle",
    name: "Fixture Title",
    type: "paragraph",
    basedOn: "Normal",
    char: { fontFamily: "Arial", fontSizePx: pt(26), bold: true, color: "#111111" },
    para: { align: "left", lineHeight: 1.05, spaceAfterPx: pt(16) },
  })
  .style({
    id: "Heading1",
    name: "Heading 1",
    type: "paragraph",
    basedOn: "Normal",
    char: { fontFamily: "Arial", fontSizePx: pt(18), bold: true, color: "#111111" },
    para: { align: "left", spaceBeforePx: pt(10), spaceAfterPx: pt(6), outlineLevel: 0 },
  })
  .style({
    id: "Heading2",
    name: "Heading 2",
    type: "paragraph",
    basedOn: "Normal",
    char: { fontFamily: "Arial", fontSizePx: pt(14), bold: true, color: "#333333" },
    para: { align: "left", spaceBeforePx: pt(8), spaceAfterPx: pt(4), outlineLevel: 1 },
  });

builder.header((story) => story.paragraph("SIMPLE DOCS · DOCX COMPATIBILITY FIXTURE", { fontFamily: "Arial", fontSizePx: pt(8), color: "#666666" }));
builder.footer((story) => {
  story.paragraph("", { fontFamily: "Arial", fontSizePx: pt(9), color: "#555555" })
    .align("center")
    .text("Page ")
    .pageField()
    .text(" of ")
    .numPagesField();
});

builder.paragraph("Simple Docs Round-trip Fixture").withStyle("FixtureTitle");
builder.paragraph("Typography and structure").withStyle("Heading1");
builder.paragraph("This line mixes ")
  .text("bold cobalt", { bold: true })
  .text(", ")
  .text("italic amber", { italic: true, color: "#b45309" })
  .text(", ")
  .text("underlined green", { underline: true, color: "#166534" })
  .text(", and ordinary body text.");
builder.paragraph("A document editor must preserve the neighboring runs while changing only a selected word. Sentinel: P1-TYPOGRAPHY-7402.");
builder.paragraph("Lists").withStyle("Heading2");
builder.bulletList([
  "A first bullet with enough words to wrap naturally across a narrow reading column.",
  { text: "A nested bullet", level: 1 },
  { text: "A third-level bullet", level: 2 },
]);
builder.numberedList(["First numbered step", "Second numbered step", "Third numbered step"]);
builder.paragraph("OpenAI documentation", { color: "#1d4ed8", underline: true }).link("https://openai.com/");

builder.pageBreak();
builder.paragraph("Tables and images").withStyle("Heading1");
builder.table([
  [
    { text: "Item", shading: "#111111", style: { bold: true, color: "#ffffff" } },
    { text: "Owner", shading: "#111111", style: { bold: true, color: "#ffffff" } },
    { text: "Status", shading: "#111111", style: { bold: true, color: "#ffffff" } },
    { text: "Value", shading: "#111111", style: { bold: true, color: "#ffffff" }, align: "right" },
  ],
  ["DOCX import", "Simple Docs", "Ready", { text: "100", align: "right" }],
  ["Page layout", "Canvas engine", "Verified", { text: "23", align: "right", shading: "#f3f4f6" }],
  ["Round-trip", "QA corpus", "Pending edit", { text: "1", align: "right" }],
], { colFractions: [0.32, 0.26, 0.25, 0.17], headerRow: true, widthMode: "fixed" });
builder.paragraph("Embedded image").withStyle("Heading2");
builder.image({ data: pixel, mime: "image/png" }, { widthPx: 240, heightPx: 96, align: "center", wrap: "block" });
builder.paragraph("Figure 1. Deterministic embedded PNG, centered and resized.", { italic: true, color: "#555555" }).align("center");
builder.paragraph("Sentinel: P2-TABLE-IMAGE-1846.");

builder.pageBreak();
builder.paragraph("Pagination, review, and export").withStyle("Heading1");
builder.paragraph("This third page proves explicit page breaks, header/footer bands, fields, and stable page ownership. Sentinel: P3-PAGINATION-9031.");
builder.paragraph("Search targets").withStyle("Heading2");
builder.paragraph("Orange highlight target: marmalade. The word marmalade appears twice on this page so next and previous search can be checked.");
builder.paragraph("Final paragraph. A controlled editing test should replace cobalt with violet and leave every surrounding style unchanged.");

await mkdir(outputDirectory, { recursive: true });
await installMeasureHost();
const { bytes, warnings } = await runExport(builder.build(), "docx", {
  [`data:image/png;base64,${pixelBase64}`]: pixel,
});
await writeFile(outputPath, bytes);
console.log(JSON.stringify({ outputPath, bytes: bytes.byteLength, warnings }, null, 2));
