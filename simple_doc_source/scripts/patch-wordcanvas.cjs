/* WordCanvas 0.12.0 adapters: anchored header images occupy no flow space, the
 * reviewed creation/fidelity fixes, and the generic SIMPLE_HOOKS layer.
 * The MIT dependency ships generated modules. Keep every patch explicit, pinned,
 * idempotent, and fail closed on upstream changes; never alter the source model.
 * See wordcanvas-patch.md for scope, sources, and remaining rendering limits.
 */
const fs = require('node:fs');
const path = require('node:path');
// Tests require() this file to review the SIMPLE_HOOKS source. Only running it
// (postinstall/pretest/prebuild:web) patches the dependency.
module.exports = { simpleHooks: simpleHooksDefinition() };
if (require.main !== module) return;
// SIMPLE_WORDCANVAS_PACKAGE_DIR points the patcher at a scratch copy of the
// package (the regression test uses it to prove the version guard).
const base = path.resolve(process.env.SIMPLE_WORDCANVAS_PACKAGE_DIR || path.join(__dirname, '../node_modules/@forevka/wordcanvas'));
const marker = '// SIMPLE_WORDCANVAS_HEADER_IMAGES_V1';
if (JSON.parse(fs.readFileSync(path.join(base, 'package.json'), 'utf8')).version !== '0.12.0') {
  throw new Error('WordCanvas patch requires reviewed version 0.12.0.');
}
const pending = new Map();
function patch(file, replacements, helpers = '') {
  const full = path.join(base, file);
  let source = fs.readFileSync(full, 'utf8').replace(/\r\n/g, '\n');
  if (source.includes(marker)) return;
  for (const [before, after, expected = 1] of replacements) {
    const count = source.split(before).length - 1;
    if (count !== expected) throw new Error(`${file}: expected ${expected} instances of ${JSON.stringify(before)}, found ${count}. No files written.`);
    source = source.split(before).join(after);
  }
  pending.set(full, `${marker}\n${helpers}\n${source}`);
}

const importHelpers = `
function simpleDrawingMetadata(container, blip) {
  const child = (n, name) => n?.children?.find(c => c?.tagName === name);
  const content = n => n?.children?.filter(c => typeof c === 'string').join('');
  const lum = child(blip, 'a:lum');
  const h = content(child(child(container, 'wp:positionH'), 'wp:align'));
  const v = content(child(child(container, 'wp:positionV'), 'wp:align'));
  const result = {};
  if (h || v) result.simplePosition = {h, v};
  if (lum) {
    const bright = Number(lum.attributes?.bright ?? 0), contrast = Number(lum.attributes?.contrast ?? 0);
    if (Number.isFinite(bright) && Number.isFinite(contrast)) result.simpleLuminance = {bright, contrast};
  }
  return result;
}
`;
patch('dist-node/chunk-S3H2FFFI.js', [
  ['{ kind: "image", relId, anchored: !!anchor }', '{ kind: "image", relId, anchored: !!anchor, ...simpleDrawingMetadata(container, blip) }'],
  ['align: inline.anchorAlign ?? (paraAlign === "justify" ? "left" : paraAlign)', '...simpleImageMetadata(inline),\n      align: inline.anchorAlign ?? (paraAlign === "justify" ? "left" : paraAlign)'],
], importHelpers + metadataHelper());
patch('dist-lib/pipeline-CsMT0pHL.js', [
  ['{ kind: "image", relId: r, anchored: !!n }', '{ kind: "image", relId: r, anchored: !!n, ...simpleDrawingMetadata(o, c) }'],
  ['align: a.anchorAlign ?? (H === "justify" ? "left" : H)', '...simpleImageMetadata(a),\n      align: a.anchorAlign ?? (H === "justify" ? "left" : H)'],
], importHelpers + metadataHelper());
patch('dist-lib/assets/worker-D0pm0kNa.js', [
  ['{ kind: "image", relId: s, anchored: !!n }', '{ kind: "image", relId: s, anchored: !!n, ...simpleDrawingMetadata(o, a) }'],
  ['align: c.anchorAlign ?? (L === "justify" ? "left" : L)', '...simpleImageMetadata(c),\n      align: c.anchorAlign ?? (L === "justify" ? "left" : L)'],
], importHelpers + metadataHelper());

function metadataHelper() {
  return `function simpleImageMetadata(image) {
  return {...image.simplePosition ? {simplePosition: {...image.simplePosition}} : {},
    ...image.simpleLuminance ? {simpleLuminance: {...image.simpleLuminance}} : {}};
}\n`;
}
const layoutHelpers = `
const simpleBandFlowCache = new WeakMap();
function simpleBandFlow(blocks) {
  if (!blocks.some(b => b.kind === 'image' && b.anchor)) return blocks;
  let flow = simpleBandFlowCache.get(blocks);
  if (!flow) { flow = blocks.filter(b => !(b.kind === 'image' && b.anchor)); simpleBandFlowCache.set(blocks, flow); }
  return flow;
}
function simplePlaceBandImages(pages, sections) {
  const coordinate = (relative, alignment, offset, extent, page, near, far) => {
    let start = near, span = page - near - far;
    if (relative === 'page') { start = 0; span = page; }
    else if (relative === 'leftMargin' || relative === 'topMargin') { start = 0; span = near; }
    else if (relative === 'rightMargin' || relative === 'bottomMargin') { start = page - far; span = far; }
    if (alignment === 'center') return start + (span - extent) / 2;
    if (alignment === 'right' || alignment === 'bottom') return start + span - extent;
    if (alignment === 'left' || alignment === 'top') return start;
    return start + (Number.isFinite(offset) ? offset : 0);
  };
  const layer = b => b.image?.behind || b.shape?.behind ? 0 : b.image?.front || b.shape?.front ? 2 : 1;
  for (const page of pages) {
    const section = sections[page.index], m = section.marginPx;
    for (const kind of ['header', 'footer']) {
      for (const block of section[page[kind + 'Source']] ?? []) {
        if (block.kind !== 'image' || !block.anchor) continue;
        const a = block.anchor, p = block.simplePosition;
        page.blocks.push({blockId: block.id, firstLineIndex: 0, lines: [], simpleBandImage: true,
          x: coordinate(a.relFromH, p?.h, a.offsetXPx, block.widthPx, section.pageWidthPx, m.left, m.right),
          y: coordinate(a.relFromV, p?.v, a.offsetYPx, block.heightPx, section.pageHeightPx, m.top, m.bottom),
          image: {src: block.src, width: block.widthPx, height: block.heightPx,
            z: a.z ?? 0, ...(a.behind ? {behind: true} : {front: true}),
            ...block.crop ? {crop: block.crop} : {}, ...block.rotation ? {rotation: block.rotation} : {},
            ...simpleImageMetadata(block)}});
      }
    }
    page.blocks.sort((a,b) => layer(a) - layer(b) || (a.image?.z ?? a.shape?.z ?? 0) - (b.image?.z ?? b.shape?.z ?? 0));
  }
}
` + metadataHelper();
const engineNode = [
  ['layoutBand(blocks, cw, cx, 999, 999, bandProbeCache, false)', 'layoutBand(simpleBandFlow(blocks), cw, cx, 999, 999, bandProbeCache, false)', 2],
  ['layoutBandCached(header.blocks, cw, cx, pageNum, pages.length, bandCache, raw)', 'layoutBandCached(simpleBandFlow(header.blocks), cw, cx, pageNum, pages.length, bandCache, raw)'],
  ['layoutBandCached(footer.blocks, cw, cx, pageNum, pages.length, bandCache, raw)', 'layoutBandCached(simpleBandFlow(footer.blocks), cw, cx, pageNum, pages.length, bandCache, raw)'],
  ['  return {\n    pages,\n    pageWidthPx: doc.section.pageWidthPx,', '  simplePlaceBandImages(pages, pageSections);\n  return {\n    pages,\n    pageWidthPx: doc.section.pageWidthPx,'],
];
patch('dist-node/chunk-CSJ442BN.js', engineNode, layoutHelpers);
patch('dist-lib/engine-BwNLlumM.js', [
  ['Oe(X, G, q, 999, 999, a, !1)', 'Oe(simpleBandFlow(X), G, q, 999, 999, a, !1)', 2],
  ['s(q.blocks, N, S, b, J.length, c, Y)', 's(simpleBandFlow(q.blocks), N, S, b, J.length, c, Y)'],
  ['s(K.blocks, N, S, b, J.length, c, Y)', 's(simpleBandFlow(K.blocks), N, S, b, J.length, c, Y)'],
  ['  return {\n    pages: J,\n    pageWidthPx: e.section.pageWidthPx,', '  simplePlaceBandImages(J, j);\n  return {\n    pages: J,\n    pageWidthPx: e.section.pageWidthPx,'],
], layoutHelpers);

// Preserve original DrawingML metadata in DOCX. The preview approximates the
// common Office 70/-70 washout with opacity; it does not change source pixels.
const exportHelpers = `
function simpleLuminanceXml(block, el) {
  return block.simpleLuminance ? el('a:lum', block.simpleLuminance) : '';
}
function simplePositionXml(block, axis, fallback, el) {
  const align = block.simplePosition?.[axis];
  return align ? el('wp:align', void 0, align) : fallback;
}
function simpleImageOpacity(image) {
  return image.simpleLuminance?.bright === 70000 && image.simpleLuminance?.contrast === -70000 ? 0.3 : 1;
}
`;
const workerEdits = [
  ['Rg($A, mA, fA, 999, 999, M, !1)', 'Rg(simpleBandFlow($A), mA, fA, 999, 999, M, !1)', 2],
  ['w(fA.blocks, HA, jA, yA, oA.length, l, _A)', 'w(simpleBandFlow(fA.blocks), HA, jA, yA, oA.length, l, _A)'],
  ['w(qA.blocks, HA, jA, yA, oA.length, l, _A)', 'w(simpleBandFlow(qA.blocks), HA, jA, yA, oA.length, l, _A)'],
  ['  return {\n    pages: oA,\n    pageWidthPx: B.section.pageWidthPx,', '  simplePlaceBandImages(oA, AA);\n  return {\n    pages: oA,\n    pageWidthPx: B.section.pageWidthPx,'],
  ['t("a:blip", E)', 't("a:blip", E, simpleLuminanceXml(B, t))'],
  ['t("wp:posOffset", void 0, String(bB(H.offsetXPx)))', 'simplePositionXml(B, "h", t("wp:posOffset", void 0, String(bB(H.offsetXPx))), t)'],
  ['t("wp:posOffset", void 0, String(bB(H.offsetYPx)))', 'simplePositionXml(B, "v", t("wp:posOffset", void 0, String(bB(H.offsetYPx))), t)'],
  ['const g = B.image(A.image.src), { width:', 'Q.save(); Q.opacity(simpleImageOpacity(A.image));\n    const g = B.image(A.image.src), { width:'],
  ['U && Q.restore(), C && Q.restore();\n    return;', 'U && Q.restore(), C && Q.restore();\n    Q.restore();\n    return;'],
];
patch('dist-lib/assets/worker-DueeItgB.js', workerEdits, layoutHelpers + exportHelpers);
patch('dist-lib/pipeline-BEYkJ78H.js', [
  ['T("a:blip", r)', 'T("a:blip", r, simpleLuminanceXml(e, T))'],
  ['T("wp:posOffset", void 0, String(ot(P.offsetXPx)))', 'simplePositionXml(e, "h", T("wp:posOffset", void 0, String(ot(P.offsetXPx))), T)'],
  ['T("wp:posOffset", void 0, String(ot(P.offsetYPx)))', 'simplePositionXml(e, "v", T("wp:posOffset", void 0, String(ot(P.offsetYPx))), T)'],
  ['const i = e.image(t.image.src), { width:', 'n.save(); n.opacity(simpleImageOpacity(t.image));\n    const i = e.image(t.image.src), { width:'],
  ['h && n.restore(), s && n.restore();\n    return;', 'h && n.restore(), s && n.restore();\n    n.restore();\n    return;'],
], exportHelpers);
patch('dist-node/export.js', [
  ['el("a:blip", blipAttr)', 'el("a:blip", blipAttr, simpleLuminanceXml(img, el))'],
  ['el("wp:posOffset", void 0, String(pxToEmu(a.offsetXPx)))', 'simplePositionXml(img, "h", el("wp:posOffset", void 0, String(pxToEmu(a.offsetXPx))), el)', 2],
  ['el("wp:posOffset", void 0, String(pxToEmu(a.offsetYPx)))', 'simplePositionXml(img, "v", el("wp:posOffset", void 0, String(pxToEmu(a.offsetYPx))), el)', 2],
  // Shape writer shares these variable names; undo its replacement explicitly.
  ['simplePositionXml(img, "h", el("wp:posOffset", void 0, String(pxToEmu(a.offsetXPx))), el)) + el("wp:positionV", { relativeFrom: a.relFromV }, simplePositionXml(img, "v", el("wp:posOffset", void 0, String(pxToEmu(a.offsetYPx))), el)) + el("wp:extent", { cx, cy }) + el("wp:wrapNone") + el("wp:docPr", { id: ctx.nextId(), name: "Shape" })', 'el("wp:posOffset", void 0, String(pxToEmu(a.offsetXPx)))) + el("wp:positionV", { relativeFrom: a.relFromV }, el("wp:posOffset", void 0, String(pxToEmu(a.offsetYPx)))) + el("wp:extent", { cx, cy }) + el("wp:wrapNone") + el("wp:docPr", { id: ctx.nextId(), name: "Shape" })'],
  ['const bytes = ctx.image(block.image.src);', 'doc.save(); doc.opacity(simpleImageOpacity(block.image));\n    const bytes = ctx.image(block.image.src);'],
  ['if (clip) doc.restore();\n    return;', 'if (clip) doc.restore();\n    doc.restore();\n    return;'],
], exportHelpers);
patch('dist-lib/editorApp-vN1g1Ew1.js', [
  ['n.pageBorders && Jn(z, n);\n    for (const H of n.blocks) yB(z, H, n.index);', 'n.pageBorders && Jn(z, n);\n    for (const H of n.blocks) if (H.simpleBandImage && H.image?.behind) HB(z, H, n.index);\n    for (const H of n.blocks) yB(z, H, n.index);'],
  ['for (const H of n.blocks) HB(z, H, n.index);', 'for (const H of n.blocks) if (!(H.simpleBandImage && H.image?.behind)) HB(z, H, n.index);'],
  ['const H = rB(n.image.src), BA = n.image.clip;', 'l.save(); l.globalAlpha *= simpleImageOpacity(n.image);\n      const H = rB(n.image.src), BA = n.image.clip;'],
  ['AA && l.restore(), BA && l.restore();\n      return;', 'AA && l.restore(), BA && l.restore();\n      l.restore();\n      return;'],
], exportHelpers);

// The generic object insertion command places the caret after an inserted
// object. A new table needs the first cell instead; retain the same transaction
// so undo/redo remains one operation. Keep this marker separate for upgrades
// from the already-installed header adapter.
{
  const file = path.join(base, 'dist-lib/editorApp-vN1g1Ew1.js');
  const tableMarker = '// SIMPLE_WORDCANVAS_TABLE_CARET_V1';
  let source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
  if (!source.includes(tableMarker)) {
    const before = `    return $Q(g, () => ({
      kind: "table",
      id: KB(),
      revision: 0,
      rows: Array.from({ length: A }, () => ({
        cells: Array.from({ length: B }, () => ({ id: KB(), blocks: [c()] }))
      }))
    }));`;
    const after = `    const table = {
      kind: "table", id: KB(), revision: 0,
      rows: Array.from({ length: A }, () => ({
        cells: Array.from({ length: B }, () => ({ id: KB(), blocks: [c()] }))
      }))
    };
    const result = $Q(g, () => table);
    return result ? {...result, selectionAfter: qA(table.rows[0].cells[0].blocks[0].id, 0)} : null;`;
    if (source.split(before).length !== 2) throw new Error('Table caret adapter: expected one reviewed insertion command. No files written.');
    source = `${tableMarker}\n${source.replace(before, after)}`;
    pending.set(file, source);
  }
}

// All expected patterns are validated before any file is written.
{
  const file = path.join(base, 'dist-lib/editorApp-vN1g1Ew1.js');
  const fontMarker = '// SIMPLE_WORDCANVAS_FONT_LIMIT_V1';
  let source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
  if (!source.includes(fontMarker)) {
    const before = 'const GU = 3e4, TU = 15e3, Mw = 10 * 1024 * 1024;';
    if (source.split(before).length !== 2) throw new Error('Font export adapter: reviewed bound not found. No files written.');
    pending.set(file, `${fontMarker}\n${source.replace(before, 'const GU = 3e4, TU = 15e3, Mw = 16 * 1024 * 1024;')}`);
  }
}
// Select a supplemental installed CJK face only for characters absent from the
// bundled SC subset. This leaves Chinese glyphs in SC and supplies Japanese and
// Korean glyphs from Malgun when that computer has an embeddable copy.
{
  const {fontCoverage} = require('../electron/font-coverage.cjs');
  const coverage = fontCoverage(fs.readFileSync(path.join(base, 'dist-node/fonts/NotoSansSC-Regular.ttf')));
  if (coverage.length < 100) throw new Error('Bundled CJK coverage could not be verified.');
  const helper = `
const simpleCjkCoverage = ${JSON.stringify(coverage)};
function simpleHasGlyph(ranges, cp) {
  let lo = 0, hi = (ranges?.length ?? 0) - 1;
  while (lo <= hi) { const m = (lo + hi) >>> 1, r = ranges[m]; if (cp < r[0]) hi = m - 1; else if (cp > r[1]) lo = m + 1; else return true; }
  return false;
}
function simpleCjkFamily(ch, family, lookup) {
  const cp = ch.codePointAt(0);
  if (family !== 'NotoSansSC' || simpleHasGlyph(simpleCjkCoverage, cp)) return family;
  const supplemental = lookup('Malgun Gothic');
  return simpleHasGlyph(supplemental?.simpleCoverage, cp) ? supplemental.family : family;
}
`;
  const patches = [
    ['dist-lib/engine-BwNLlumM.js', [
      ['import { B as Li,', 'import { b0 as simpleCustomFont, B as Li,'],
      ['t && de(e) ? "cjk" : "other"', 't && de(e) ? "cjk:" + simpleCjkFamily(e, t, simpleCustomFont) : "other"'],
      ['h === "cjk" && t && (A = t)', 'h?.startsWith("cjk:") && t && (A = h.slice(4))'],
    ]],
    ['dist-lib/assets/worker-DueeItgB.js', [
      ['A && CQ(B) ? "cjk" : "other"', 'A && CQ(B) ? "cjk:" + simpleCjkFamily(B, A, aJ) : "other"'],
      ['R === "cjk" && A && (H = A)', 'R?.startsWith("cjk:") && A && (H = R.slice(4))'],
    ]],
    ['dist-node/chunk-CSJ442BN.js', [
      ['  MATH_FONT_FAMILY,', '  customFontFor,\n  MATH_FONT_FAMILY,'],
      ['if (cjkFam && isCJK(ch)) return "cjk";', 'if (cjkFam && isCJK(ch)) return "cjk:" + simpleCjkFamily(ch, cjkFam, customFontFor);'],
      ['else if (bufScript === "cjk" && cjkFam) fontFamily = cjkFam;', 'else if (bufScript?.startsWith("cjk:") && cjkFam) fontFamily = bufScript.slice(4);'],
    ]],
  ];
  for (const [relative, edits] of patches) {
    const file = path.join(base, relative), coverageMarker = '// SIMPLE_WORDCANVAS_CJK_COVERAGE_V1';
    let source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
    if (source.includes(coverageMarker)) continue;
    for (const [before, after] of edits) {
      if (source.split(before).length !== 2) throw new Error(`${relative}: CJK coverage adapter changed upstream. No files written.`);
      source = source.replace(before, after);
    }
    pending.set(file, `${coverageMarker}\n${helper}\n${source}`);
  }
}
{
  const helper = String.raw`
function simpleRunContent(text, el, textEl) {
  return text.split(/(\t|\r\n|[\v\n\r\f])/).map(piece => {
    if (!piece) return '';
    if (piece === '\t') return el('w:tab');
    if (piece === '\f') return el('w:br', {'w:type':'page'});
    if (/^[\v\n\r]+$/.test(piece)) return el('w:br');
    return textEl(piece.replace(/[\u0000-\u0008\u000e-\u001f]/g, '\ufffd'));
  }).join('');
}
function simpleSquareImage(image) {
  if (image.anchor || image.wrap !== 'square') return image;
  return {...image, simplePosition: {...image.simplePosition, h: image.align === 'right' ? 'right' : image.align === 'center' ? 'center' : 'left'},
    anchor: {relFromH:'margin', relFromV:'paragraph', offsetXPx:0, offsetYPx:0, behind:false}};
}
function simpleImageWrap(image, el) {
  return image.wrap === 'square' ? el('wp:wrapSquare', {wrapText:'bothSides'}) : el('wp:wrapNone');
}
`;
  for (const [relative, runName, arg, elName, textName, imageName, context] of [
    ['dist-node/export.js', 'runContent', 'text', 'el', 'textEl', 'imageParagraphXml', 'ctx'],
    ['dist-lib/pipeline-BEYkJ78H.js', 'Br', 'e', 'T', 'Qf', 'qp', 't'],
    ['dist-lib/assets/worker-DueeItgB.js', 'HI', 'B', 't', 'NH', 'pK', 'A'],
  ]) {
    const file = path.join(base, relative), marker = '// SIMPLE_WORDCANVAS_DOCX_CREATION_V1';
    let source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
    if (source.includes(marker)) continue;
    const pattern = new RegExp(`function ${runName}\\(${arg}\\) \\{[\\s\\S]*?\\n\\}`, 'g');
    const matches = [...source.matchAll(pattern)];
    if (matches.length !== 1 || !matches[0][0].includes('w:tab')) throw new Error(`${relative}: text writer changed upstream. No files written.`);
    source = source.replace(pattern, `function ${runName}(${arg}) { return simpleRunContent(${arg}, ${elName}, ${textName}); }`);
    const imageArg = relative.startsWith('dist-node') ? 'img' : arg;
    const edits = [
      [`function ${imageName}(${imageArg}, ${context}) {`, `function ${imageName}(${imageArg}, ${context}) {\n  ${imageArg} = simpleSquareImage(${imageArg});`],
      [`${elName}("wp:wrapNone") + ${elName}("wp:docPr", { id: ${context}.nextId(), name: "image" })`, `simpleImageWrap(${imageArg}, ${elName}) + ${elName}("wp:docPr", { id: ${context}.nextId(), name: "image" })`],
    ];
    for (const [before, after] of edits) {
      if (source.split(before).length !== 2) throw new Error(`${relative}: image writer changed upstream. No files written.`);
      source = source.replace(before, after);
    }
    pending.set(file, `${marker}\n${helper}\n${source}`);
  }
}
{
  const edits = [
    ['dist-node/chunk-S3H2FFFI.js', '            warnings.add("soft-breaks", "Soft line breaks (Shift+Enter) became paragraph breaks.");\n            if (runs.length > 0) flushPara();\n            else blocks.push(paraOf([]));', '            runs.push({text:"\\v", style:markChar});\n            trailingBreak = false;\n            break;'],
    ['dist-lib/pipeline-CsMT0pHL.js', '!S.page && !S.column ? (t.add("soft-breaks", "Soft line breaks (Shift+Enter) became paragraph breaks."), W.length > 0 ? tt() : x.push(L([]))) : (W.length > 0 && tt(), S.page && (R = !0), S.column && (bt = !0)), _ = !0;', 'if (!S.page && !S.column) { W.push({text:"\\v", style:k}); _ = !1; break; }\n          W.length > 0 && tt(), S.page && (R = !0), S.column && (bt = !0), _ = !0;'],
    ['dist-lib/assets/worker-D0pm0kNa.js', '!x.page && !x.column ? (e.add("soft-breaks", "Soft line breaks (Shift+Enter) became paragraph breaks."), j.length > 0 ? lt() : E.push(V([]))) : (j.length > 0 && lt(), x.page && (U = !0), x.column && (bt = !0)), Y = !0;', 'if (!x.page && !x.column) { j.push({text:"\\v", style:I}); Y = !1; break; }\n          j.length > 0 && lt(), x.page && (U = !0), x.column && (bt = !0), Y = !0;'],
  ];
  for (const [relative, before, after] of edits) {
    const file = path.join(base, relative), marker = '// SIMPLE_WORDCANVAS_SOFT_BREAK_IMPORT_V1';
    const source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
    if (source.includes(marker)) continue;
    if (source.split(before).length !== 2) throw new Error(`${relative}: line break importer changed upstream. No files written.`);
    pending.set(file, `${marker}\n${source.replace(before, after)}`);
  }
}
{
  const patches = [
    ['dist-node/chunk-CSJ442BN.js', [
      ['        if (y + lineHeight > bottomY() && colHasContent()) {', '        const simpleRefs = refsInLines([{fragments:[...bl.frags, ...(right?.frags ?? [])].map(item=>item.frag)}], 0, 1);\n        const simpleNotes = simpleRefs.length ? measureNotes(simpleRefs) : {H:0, measures:[]};\n        if (y + lineHeight + simpleNotes.H > bottomY() && colHasContent()) {'],
      ['        if (levels !== null && bl.frags.length > 0) {', '        if (simpleNotes.measures.length) commitNotes(simpleNotes.measures, simpleNotes.H);\n        if (levels !== null && bl.frags.length > 0) {'],
    ]],
    ['dist-lib/engine-BwNLlumM.js', [
      ['        if (_ + Yt > re() && xe()) {', '        const simpleRefs = ui([{fragments:[...te.frags, ...(oe?.frags ?? [])].map(item=>item.frag)}], 0, 1);\n        const simpleNotes = simpleRefs.length ? _t(simpleRefs) : {H:0, measures:[]};\n        if (_ + Yt + simpleNotes.H > re() && xe()) {'],
      ['        if (F !== null && te.frags.length > 0) {', '        if (simpleNotes.measures.length) Gt(simpleNotes.measures, simpleNotes.H);\n        if (F !== null && te.frags.length > 0) {'],
    ]],
    ['dist-lib/assets/worker-DueeItgB.js', [
      ['        if (MA + Vg > IA() && QB()) {', '        const simpleRefs = tA([{fragments:[...NB.frags, ...(zB?.frags ?? [])].map(item=>item.frag)}], 0, 1);\n        const simpleNotes = simpleRefs.length ? XA(simpleRefs) : {H:0, measures:[]};\n        if (MA + Vg + simpleNotes.H > IA() && QB()) {'],
      ['        if (BA !== null && NB.frags.length > 0) {', '        if (simpleNotes.measures.length) AB(simpleNotes.measures, simpleNotes.H);\n        if (BA !== null && NB.frags.length > 0) {'],
    ]],
  ];
  for (const [relative, edits] of patches) {
    const file = path.join(base, relative), marker = '// SIMPLE_WORDCANVAS_FLOAT_NOTES_V1';
    let source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
    if (source.includes(marker)) continue;
    for (const [before, after] of edits) {
      if (source.split(before).length !== 2) throw new Error(`${relative}: wrapped paragraph note layout changed upstream. No files written.`);
      source = source.replace(before, after);
    }
    pending.set(file, `${marker}\n${source}`);
  }
}
{
  const file = path.join(base, 'dist-lib/editorApp-vN1g1Ew1.js'), marker = '// SIMPLE_WORDCANVAS_PASTE_COLORS_V1';
  let source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
  if (!source.includes(marker)) {
    const edits = [
      ['g.highlightColor = M.backgroundColor', 'g.highlightColor = simpleOfficeColor(M.backgroundColor, "#ffffff")'],
      ['g.color = M.color', 'g.color = simpleOfficeColor(M.color, g.color)'],
    ];
    for (const [before, after] of edits) {
      if (source.split(before).length !== 2) throw new Error('HTML color importer changed upstream. No files written.');
      source = source.replace(before, after);
    }
    const helper = `
let simpleColorContext;
function simpleOfficeColor(value, fallback) {
  const ctx = simpleColorContext ??= document.createElement('canvas').getContext('2d', {willReadFrequently:true});
  ctx.fillStyle = '#ffffff'; ctx.fillRect(0,0,1,1);
  ctx.fillStyle = fallback || '#202124'; ctx.fillStyle = value; ctx.fillRect(0,0,1,1);
  const pixel = ctx.getImageData(0,0,1,1).data;
  return '#' + [...pixel].slice(0,3).map(channel=>channel.toString(16).padStart(2,'0')).join('');
}
`;
    pending.set(file, `${marker}\n${helper}\n${source}`);
  }
}
// SIMPLE_HOOKS runs last: its anchors sit outside every region edited above.
// The marker line carries a fingerprint of the hook source, so an install that
// holds an older SIMPLE_HOOKS revision fails loudly instead of being skipped.
{
  const {marker: hooksMarker, markerLine, file: relative, helpers, edits} = module.exports.simpleHooks;
  const file = path.join(base, relative);
  let source = pending.get(file) ?? fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
  const applied = source.split('\n').find((line) => line.startsWith(hooksMarker));
  if (applied && applied !== markerLine) throw new Error(`${relative}: holds "${applied}" but this script expects "${markerLine}". Reinstall @forevka/wordcanvas 0.12.0 (npm ci) so the reviewed hooks are applied to pristine sources. No files written.`);
  if (!applied) {
    for (const [before, after] of edits) {
      const count = source.split(before).length - 1;
      if (count !== 1) throw new Error(`${relative}: SIMPLE_HOOKS expected one reviewed instance of ${JSON.stringify(before.slice(0, 120))}, found ${count}. No files written.`);
      source = source.split(before).join(after);
    }
    pending.set(file, `${markerLine}\n${helpers}\n${source}`);
  }
}
for (const [file, source] of pending) fs.writeFileSync(file, source);
console.log(`WordCanvas 0.12.0 adapters: ${pending.size ? `patched ${pending.size} modules` : 'already applied'}.`);

/* SIMPLE_HOOKS: one generic "Simple layer" over the browser editor
 * (dist-lib/editorApp) instead of one minified patch per feature. It adds
 * - a `simple:docchange` custom event after every committed model change
 *   (typing, commands, paste, undo/redo, remote ops, document loads), never
 *   for selection, view or transient preview changes;
 * - handle methods: insertImageBytes, insertBlocks, replaceBlock, replaceImage,
 *   undo/redo/canUndo/canRedo, setSelection, focus, seedReview,
 *   positionFromPoint, deleteWord and getModelRevision;
 * - a dialog bridge for every reachable prompt()/alert() (prompt() throws in
 *   Electron 43) and a rejecting openDocx instead of an alert;
 * - an after-insert text hook (AutoFormat) and a built-in autocorrect switch;
 * - Ctrl+Backspace / Ctrl+Delete word deletion (keydown and beforeinput).
 * src/engine-bridge.ts is the typed consumer; wordcanvas-patch.md documents
 * the contract. Every edit below must match exactly once or nothing is written.
 */
function simpleHooksDefinition() {
  const marker = '// SIMPLE_WORDCANVAS_SIMPLE_HOOKS_V1';
  // Module-level helpers. Engine names used here (xM, XA, xB, Qg, Jg, BI, UC,
  // hC) are imports or module functions of editorApp-vN1g1Ew1.js 0.12.0.
  const helpers = String.raw`
const simpleHookVersion = 1, simpleHookBands = ["header", "footer", "headerFirst", "headerEven", "footerFirst", "footerEven"];
function simpleHookApp(options) {
  const app = {
    version: simpleHookVersion,
    revision: 0,
    dialog: null,
    insertTextHook: null,
    autoCorrect: !0,
    docChanged(origin, canUndo, canRedo) {
      const payload = { revision: ++app.revision, origin, canUndo: !!canUndo, canRedo: !!canRedo };
      try {
        options?.onEvent?.({ type: "custom", name: "simple:docchange", payload });
      } catch (error) {
        console.error("[simple-hooks] a simple:docchange listener failed", error);
      }
      return payload;
    }
  };
  return app;
}
function simpleHookDialog(app, request) {
  const kind = request.kind, empty = kind === "prompt" ? null : kind === "confirm" ? !1 : void 0;
  const settle = (value) => kind === "prompt" ? typeof value == "string" ? value : null : kind === "confirm" ? value === !0 : void 0;
  const handler = app?.dialog;
  if (typeof handler == "function") {
    try {
      return Promise.resolve(handler({ ...request })).then(settle, (error) => (console.error("[simple-hooks] dialog handler failed", error), empty));
    } catch (error) {
      console.error("[simple-hooks] dialog handler failed", error);
      return Promise.resolve(empty);
    }
  }
  try {
    if (kind === "prompt") return Promise.resolve(settle(window.prompt(request.message, request.defaultValue ?? "")));
    if (kind === "confirm") return Promise.resolve(settle(window.confirm(request.message)));
    window.alert(request.message);
  } catch (error) {
    console.warn("[simple-hooks] window." + kind + "() is unavailable; register a dialog handler with setDialogHandler().", error);
  }
  return Promise.resolve(empty);
}
async function simpleHookPrepareImage(bytes, mime) {
  const data = ArrayBuffer.isView(bytes) ? new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength) : Object.prototype.toString.call(bytes) === "[object ArrayBuffer]" ? new Uint8Array(bytes) : null;
  if (!data || data.byteLength === 0) return null;
  const type = typeof mime == "string" && /^image\/[\w.+-]+$/i.test(mime) ? mime.toLowerCase() : "image/png";
  let widthPx = 0, heightPx = 0;
  try {
    const bitmap = await createImageBitmap(new Blob([data], { type }));
    widthPx = bitmap.width, heightPx = bitmap.height, bitmap.close();
  } catch {
    return null;
  }
  if (!(widthPx > 0 && heightPx > 0)) return null;
  const mediaId = await UC(data, type), src = hC(mediaId);
  return src ? { src, mediaId, mime: type, widthPx, heightPx } : null;
}
function simpleHookMapChildren(block, visit) {
  if (block?.kind === "table" && Array.isArray(block.rows)) {
    let changed = !1;
    const rows = block.rows.map((row) => {
      let rowChanged = !1;
      const cells = (row.cells ?? []).map((cell) => {
        const blocks = simpleHookMapBlocks(cell.blocks, visit);
        return blocks === cell.blocks ? cell : (rowChanged = !0, { ...cell, blocks });
      });
      return rowChanged ? (changed = !0, { ...row, cells }) : row;
    });
    return changed ? { ...block, rows, revision: (block.revision ?? 0) + 1 } : block;
  }
  if (block?.kind === "shape" && Array.isArray(block.text?.blocks)) {
    const blocks = simpleHookMapBlocks(block.text.blocks, visit);
    return blocks === block.text.blocks ? block : { ...block, text: { ...block.text, blocks }, revision: (block.revision ?? 0) + 1 };
  }
  return block;
}
function simpleHookMapBlocks(blocks, visit) {
  if (!Array.isArray(blocks)) return blocks;
  let out = null;
  for (let index = 0; index < blocks.length; index++) {
    const block = blocks[index];
    let next = visit(block);
    next === block && (next = simpleHookMapChildren(block, visit));
    next !== block && ((out ??= blocks.slice())[index] = next);
  }
  return out ?? blocks;
}
function simpleHookMapDocument(doc, visit) {
  let next = doc;
  const blocks = simpleHookMapBlocks(doc.blocks, visit);
  blocks !== doc.blocks && (next = { ...next, blocks });
  let section = doc.section;
  for (const band of simpleHookBands) {
    const list = doc.section?.[band];
    if (!list) continue;
    const mapped = simpleHookMapBlocks(list, visit);
    mapped !== list && (section = { ...section, [band]: mapped });
  }
  section !== doc.section && (next = { ...next, section });
  for (const key of ["footnotes", "endnotes"]) {
    const notes = doc[key];
    if (!notes) continue;
    let changed = null;
    for (const [id, list] of Object.entries(notes)) {
      const mapped = simpleHookMapBlocks(list, visit);
      mapped !== list && ((changed ??= { ...notes })[id] = mapped);
    }
    changed && (next = { ...next, [key]: changed });
  }
  return next;
}
function simpleHookRevisionMap(block, map = new Map()) {
  if (!block || typeof block != "object") return map;
  typeof block.id == "string" && Number.isFinite(block.revision) && map.set(block.id, block.revision);
  if (block.kind === "table") for (const row of block.rows ?? []) for (const cell of row.cells ?? []) for (const child of cell.blocks ?? []) simpleHookRevisionMap(child, map);
  else if (block.kind === "shape") for (const child of block.text?.blocks ?? []) simpleHookRevisionMap(child, map);
  return map;
}
// Layout caches measurements by block id + revision: a replacement must never
// reuse the revision of the block (or nested block) it replaces.
function simpleHookRevise(block, previous) {
  if (!block || typeof block != "object") return block;
  const own = Number.isFinite(block.revision) ? block.revision : 0, old = previous.get(block.id);
  const next = { ...block, revision: old === void 0 ? own : Math.max(own, old + 1) };
  if (next.kind === "table" && Array.isArray(next.rows)) next.rows = next.rows.map((row) => ({ ...row, cells: (row.cells ?? []).map((cell) => ({ ...cell, blocks: (cell.blocks ?? []).map((child) => simpleHookRevise(child, previous)) })) }));
  else if (next.kind === "shape" && Array.isArray(next.text?.blocks)) next.text = { ...next.text, blocks: next.text.blocks.map((child) => simpleHookRevise(child, previous)) };
  return next;
}
function simpleHookRemint(block) {
  if (!block || typeof block != "object") return block;
  const next = { ...block, id: Jg(), revision: Number.isFinite(block.revision) ? block.revision : 0 };
  if (next.kind === "table" && Array.isArray(next.rows)) next.rows = next.rows.map((row) => ({ ...row, cells: (row.cells ?? []).map((cell) => ({ ...cell, id: Jg(), blocks: (cell.blocks ?? []).map((child) => simpleHookRemint(child)) })) }));
  else if (next.kind === "shape" && Array.isArray(next.text?.blocks)) next.text = { ...next.text, blocks: next.text.blocks.map((child) => simpleHookRemint(child)) };
  return next;
}
function simpleHookFirstParagraph(block) {
  if (!block || typeof block != "object") return null;
  if (block.kind === "paragraph") return block.id;
  const children = block.kind === "table" ? (block.rows ?? []).flatMap((row) => (row.cells ?? []).flatMap((cell) => cell.blocks ?? [])) : block.kind === "shape" ? block.text?.blocks ?? [] : [];
  for (const child of children) {
    const id = simpleHookFirstParagraph(child);
    if (id) return id;
  }
  return null;
}
function simpleHookSelectionIn(doc, selection, replacement) {
  const valid = (position) => {
    const block = position && XA(doc, position.blockId);
    return !!block && Number.isInteger(position.offset) && position.offset >= 0 && position.offset <= xB(block.runs).length;
  };
  if (selection && valid(selection.anchor) && valid(selection.focus)) return selection;
  const first = simpleHookFirstParagraph(replacement);
  return first ? { anchor: { blockId: first, offset: 0 }, focus: { blockId: first, offset: 0 } } : null;
}
function simpleHookPatchRuns(runs, start, end, patch) {
  const out = [];
  let at = 0;
  for (const run of runs) {
    const from = at, to = at + run.text.length;
    if (at = to, to <= start || from >= end || run.text.length === 0) {
      out.push(run);
      continue;
    }
    const a = Math.max(start, from) - from, b = Math.min(end, to) - from;
    a > 0 && out.push({ text: run.text.slice(0, a), style: run.style });
    out.push({ text: run.text.slice(a, b), style: { ...run.style, ...patch } });
    b < run.text.length && out.push({ text: run.text.slice(b), style: run.style });
  }
  return out;
}
// One transaction for an insert-text hook result, so Ctrl+Z reverts only the
// correction. Text edits apply right to left; list edits apply last.
function simpleHookEditTransaction(state, edits, blockId) {
  if (!Array.isArray(edits) || edits.length === 0) return null;
  let doc = state.doc, selection = state.selection;
  const ops = [], apply = (op) => {
    const result = xM(doc, op);
    doc = result.doc, ops.push(op), selection && (selection = { anchor: result.mapPosition(selection.anchor), focus: result.mapPosition(selection.focus) });
  };
  const textEdits = edits.filter((edit) => edit && (edit.type === "replace" || edit.type === "format")).sort((x, y) => (Number(y.start) || 0) - (Number(x.start) || 0));
  for (const edit of textEdits) {
    const id = typeof edit.blockId == "string" ? edit.blockId : blockId, block = id ? XA(doc, id) : void 0;
    if (!block) continue;
    const length = xB(block.runs).length, clamp = (value) => Math.max(0, Math.min(length, Math.trunc(Number(value) || 0)));
    const start = clamp(edit.start), end = Math.max(start, clamp(edit.end ?? edit.start)), patch = edit.style && typeof edit.style == "object" ? edit.style : null;
    if (edit.type === "replace") {
      const text = typeof edit.text == "string" ? edit.text : "", style = { ...Qg(block.runs, end > start ? start + 1 : start) ?? {}, ...patch ?? {} };
      end > start && apply({ type: "deleteRange", blockId: id, start, end }), text && apply({ type: "insertText", at: { blockId: id, offset: start }, text, style });
    } else end > start && patch && apply({ type: "setRuns", blockId: id, runs: simpleHookPatchRuns(block.runs, start, end, patch) });
  }
  for (const edit of edits) {
    if (edit?.type !== "list") continue;
    const kind = edit.kind === "bullet" ? "bullet" : edit.kind === "number" || edit.kind === "decimal" ? "decimal" : null, id = typeof edit.blockId == "string" ? edit.blockId : blockId;
    if (!kind || !id || !XA(doc, id)) continue;
    const caret = { blockId: id, offset: 0 }, command = BI(kind)({ doc, selection: { anchor: caret, focus: caret }, cellSelection: null, pendingStyle: null });
    for (const op of command?.ops ?? []) apply(op);
  }
  return ops.length > 0 ? { ops, selectionAfter: selection, origin: "command" } : null;
}
`;
  // Editor-local helpers, inserted at the top of the editor factory (_M).
  const editorHelpers = String.raw`
  let simpleHookSeen = o;
  const simpleHookCommit = (origin) => {
    o !== simpleHookSeen && (simpleHookSeen = o, g.simple?.docChanged(origin, M !== "view" && t.canUndo, M !== "view" && t.canRedo));
  }, simpleHookAfterInsert = (kind, text, run) => {
    const before = o, previous = h?.focus;
    run();
    const hook = g.simple?.insertTextHook;
    if (typeof hook != "function" || o === before || M === "view" || !h || !zB(h)) return;
    const focus = h.focus, block = XA(o, focus.blockId);
    if (!block) return;
    const context = { kind, text, blockId: focus.blockId, offset: focus.offset, paragraphText: xB(block.runs), paragraphStyle: { ...block.style }, mode: M };
    if (kind === "paragraph" && previous && previous.blockId !== focus.blockId) {
      const prior = XA(o, previous.blockId);
      prior && (context.previousBlockId = prior.id, context.previousText = xB(prior.runs));
    }
    let edits;
    try {
      edits = hook(context);
    } catch (error) {
      console.error("[simple-hooks] insert-text hook failed", error);
      return;
    }
    Array.isArray(edits) && edits.length > 0 && _((state) => simpleHookEditTransaction(state, edits, focus.blockId));
  }, simpleHookDeleteWord = (direction) => {
    if (M === "view" || fI()) return !1;
    const before = o, plain = () => direction < 0 ? UM() : Qt();
    return _((state) => {
      const selection = state.selection;
      if (!selection) return null;
      if (!zB(selection)) return plain()(state);
      const target = ic?.simpleWordTarget?.(selection.focus, direction);
      if (!target || target.blockId !== selection.focus.blockId) return plain()(state);
      const { blockId, offset } = selection.focus, start = Math.min(offset, target.offset), end = Math.max(offset, target.offset);
      return start === end ? null : wA([{ type: "deleteRange", blockId, start, end }], qA(blockId, start), "command");
    }), o !== before;
  }, simpleHookTextWidth = () => {
    const caret = h ? og(Y, h.focus, J()) : null, page = caret ? Y.pages[caret.pageIndex] : null, section = o.section, columns = section.columns;
    const width = page ? page.widthPx - page.marginPx.left - page.marginPx.right : section.pageWidthPx - section.marginPx.left - section.marginPx.right;
    return columns?.count > 1 ? (width - (columns.gapPx ?? 0) * (columns.count - 1)) / columns.count : width;
  }, simpleHookPosition = (x, y) => {
    const point = Number.isFinite(x) && Number.isFinite(y) ? C.clientToPage(x, y) : null;
    return point ? UI(Y, point.pageIndex, point.x, point.y, J()) : null;
  }, simpleHookInsertImage = (image, options = {}) => {
    if (M === "view" || !image?.src) return !1;
    if (options.at) {
      const position = simpleHookPosition(options.at.clientX, options.at.clientY);
      position && l({ anchor: position, focus: position });
    }
    const naturalWidth = image.widthPx > 0 ? image.widthPx : 320, naturalHeight = image.heightPx > 0 ? image.heightPx : 200;
    let width = options.widthPx > 0 ? options.widthPx : 0, height = options.heightPx > 0 ? options.heightPx : 0;
    if (width && !height) height = width * naturalHeight / naturalWidth;
    else if (height && !width) width = height * naturalWidth / naturalHeight;
    else if (!width) width = naturalWidth, height = naturalHeight;
    const limit = options.maxWidthPx > 0 ? options.maxWidthPx : simpleHookTextWidth();
    limit > 0 && width > limit && (height = height * limit / width, width = limit), width = Math.max(1, Math.round(width)), height = Math.max(1, Math.round(height));
    const before = o;
    return _(jC(image.src, width, height, image.mediaId)), o === before && _(Eo(image.src, width, height, image.mediaId)), o !== before;
  }, simpleHookInsertBlocks = (blocks, options = {}) => {
    if (M === "view") return !1;
    const list = (Array.isArray(blocks) ? blocks : [blocks]).filter((block) => block && typeof block == "object" && typeof block.kind == "string");
    if (list.length === 0) return !1;
    const fresh = list.map((block) => options.keepIds ? simpleHookRevise(block, new Map()) : simpleHookRemint(block)), before = o;
    return _((state) => {
      if (Number.isInteger(options.index)) {
        const index = Math.max(0, Math.min(state.doc.blocks.length, options.index));
        return wA(fresh.map((block, offset) => ({ type: "insertBlock", index: index + offset, block })), state.selection, "command");
      }
      const removal = EI(state);
      if (!removal) return null;
      let doc = state.doc;
      for (const op of removal.ops) doc = xM(doc, op).doc;
      const placed = $Q({ ...state, doc, selection: { anchor: removal.at, focus: removal.at } }, () => fresh[0]);
      if (!placed) return null;
      const ops = [...removal.ops, ...placed.ops], first = placed.ops.find((op) => op.type === "insertBlock");
      for (let offset = 1; offset < fresh.length; offset++) ops.push({ type: "insertBlock", index: first.index + offset, block: fresh[offset] });
      return wA(ops, placed.selectionAfter, "command");
    }), o !== before;
  }, simpleHookReplaceBlock = (blockId, replacement) => {
    if (M !== "edit" || typeof blockId != "string" || !replacement) return !1;
    const before = o;
    return _((state) => {
      let found = !1, next = null;
      const doc = simpleHookMapDocument(state.doc, (block) => {
        if (found || block?.id !== blockId) return block;
        found = !0;
        const candidate = typeof replacement == "function" ? replacement(block) : replacement;
        return candidate && typeof candidate == "object" && typeof candidate.kind == "string" ? next = simpleHookRevise({ ...candidate, id: blockId }, simpleHookRevisionMap(block)) : block;
      });
      return next && doc !== state.doc ? wA([{ type: "setDocument", doc }], simpleHookSelectionIn(doc, state.selection, next), "command") : null;
    }), o !== before;
  }, simpleHookReplaceImage = (imageId, image, options = {}) => !!image?.src && simpleHookReplaceBlock(imageId, (old) => {
    if (old.kind !== "image") return null;
    const ratio = image.widthPx > 0 && image.heightPx > 0 ? image.widthPx / image.heightPx : old.widthPx / Math.max(1, old.heightPx);
    const width = options.widthPx > 0 ? options.widthPx : old.widthPx, height = options.heightPx > 0 ? options.heightPx : options.fit === "frame" ? old.heightPx : width / ratio;
    const next = { ...old, src: image.src, mediaId: image.mediaId, widthPx: Math.max(1, Math.round(width)), heightPx: Math.max(1, Math.round(height)) };
    if (options.crop === null) delete next.crop;
    else if (options.crop !== void 0) next.crop = options.crop;
    else if (!options.keepCrop) delete next.crop;
    return next;
  });`;
  const edits = [
    // Editor factory: model revision tracking and the editor-side hook bodies.
    ['  C.setTree(Y);\n  const v = () => ({ doc: o, selection: h, cellSelection: a, pendingStyle: d }), f = (E) => {',
      `  C.setTree(Y);${editorHelpers}\n  const v = () => ({ doc: o, selection: h, cellSelection: a, pendingStyle: d }), f = (E) => {`],
    // Committed transactions, undo, redo and remote ops emit simple:docchange.
    ['Vg(E.selectionAfter, E.origin === "transient");\n  }, Fg = (E) => {',
      'Vg(E.selectionAfter, E.origin === "transient"), E.origin !== "transient" && E.ops.length > 0 && simpleHookCommit(E.origin);\n  }, Fg = (E) => {'],
    ['F.record(E.inverseOps, "undo", E.selectionBefore, Date.now()), Vg(E.selectionBefore);',
      'F.record(E.inverseOps, "undo", E.selectionBefore, Date.now()), Vg(E.selectionBefore), simpleHookCommit("undo");'],
    ['F.record(E.ops, "redo", E.selectionAfter, Date.now()), Vg(E.selectionAfter);',
      'F.record(E.ops, "redo", E.selectionAfter, Date.now()), Vg(E.selectionAfter), simpleHookCommit("redo");'],
    ['    U = ci(U, o), E.length > 0 && SA(), Vg(i);\n  }, eI = (E) => {',
      '    U = ci(U, o), E.length > 0 && SA(), Vg(i), simpleHookCommit("remote");\n  }, eI = (E) => {'],
    // Typed text: built-in autocorrect switch and the after-insert hook.
    ['    if (E.length === 1 && h && zB(h)) {\n      const D = XA(o, h.focus.blockId), N = D ? xB(D.runs).slice(0, h.focus.offset) : "";',
      '    if (E.length === 1 && h && zB(h) && g.simple?.autoCorrect !== !1) {\n      const D = XA(o, h.focus.blockId), N = D ? xB(D.runs).slice(0, h.focus.offset) : "";'],
    ['    onInsertText: (E) => Qc(E),',
      '    onInsertText: (E) => simpleHookAfterInsert("text", E, () => Qc(E)),'],
    ['    onSplitParagraph: () => {\n      const E = h?.focus;\n      E && Nt(o, E) || _(Oc());\n    },',
      '    onSplitParagraph: () => simpleHookAfterInsert("paragraph", "\\n", () => {\n      const E = h?.focus;\n      E && Nt(o, E) || _(Oc());\n    }),'],
    // Word deletion: hidden-input beforeinput (IME/touch) and Ctrl+Backspace/Delete.
    ['    onDeleteForward: () => {\n      fI() || _(Qt());\n    },',
      '    onDeleteForward: () => {\n      fI() || _(Qt());\n    },\n    onDeleteWordBackward: () => simpleHookDeleteWord(-1),\n    onDeleteWordForward: () => simpleHookDeleteWord(1),'],
    ['      case "deleteContentForward":\n        F.preventDefault(), B.onDeleteForward();\n        return;',
      '      case "deleteContentForward":\n        F.preventDefault(), B.onDeleteForward();\n        return;\n      case "deleteWordBackward":\n      case "deleteWordForward":\n        if (Q) return;\n        F.preventDefault(), F.inputType === "deleteWordBackward" ? B.onDeleteWordBackward?.() : B.onDeleteWordForward?.();\n        return;'],
    ['    onDeleteSelection: () => _(UM()),\n    getStory: () => T,',
      '    onDeleteSelection: () => _(UM()),\n    simpleDeleteWord: (E) => simpleHookDeleteWord(E),\n    getStory: () => T,'],
    ['    if (!q) return;\n    const CA = (X) => {',
      '    if (!q) return;\n    if (jA && !V.altKey && !V.shiftKey && !V.isComposing && (V.key === "Backspace" || V.key === "Delete") && A.simpleDeleteWord) {\n      A.simpleDeleteWord(V.key === "Backspace" ? -1 : 1), V.preventDefault();\n      return;\n    }\n    const CA = (X) => {'],
    // Same word boundaries as Ctrl+Arrow, in logical (not visual) direction.
    ['B.addEventListener("cut", vA), {\n    destroy() {',
      'B.addEventListener("cut", vA), {\n    simpleWordTarget: (V, q) => c(V, q, !0),\n    destroy() {'],
    ['    undo: EQ,\n    redo: eQ,\n    destroy() {',
      '    simpleCanUndo: () => M !== "view" && t.canUndo,\n    simpleCanRedo: () => M !== "view" && t.canRedo,\n    simpleDeleteWord: (E) => simpleHookDeleteWord(E < 0 ? -1 : 1),\n    simpleInsertImage: (E, i) => simpleHookInsertImage(E, i),\n    simpleInsertBlocks: (E, i) => simpleHookInsertBlocks(E, i),\n    simpleReplaceBlock: (E, i) => simpleHookReplaceBlock(E, i),\n    simpleReplaceImage: (E, i, D) => simpleHookReplaceImage(E, i, D),\n    simplePositionFromPoint: (E, i) => simpleHookPosition(E, i),\n    undo: EQ,\n    redo: eQ,\n    destroy() {'],
    // Context-menu prompts (editor factory scope).
    ['      i("Edit Hyperlink…", () => {\n        const oA = prompt("Link URL:", FB);\n        oA !== null && _(HM(oA.trim() === "" ? null : oA.trim()));\n      }),',
      '      i("Edit Hyperlink…", () => {\n        simpleHookDialog(g.simple, { kind: "prompt", id: "hyperlink.edit", title: "Edit hyperlink", message: "Link URL:", defaultValue: FB }).then((oA) => {\n          oA !== null && _(HM(oA.trim() === "" ? null : oA.trim())), ig.focus();\n        });\n      }),'],
    ['      i("Insert Hyperlink…", () => {\n        const oA = prompt("Link URL:");\n        oA !== null && oA.trim() !== "" && _(HM(oA.trim()));\n      }, { icon: u.link }),',
      '      i("Insert Hyperlink…", () => {\n        simpleHookDialog(g.simple, { kind: "prompt", id: "hyperlink.insert", title: "Insert hyperlink", message: "Link URL:", defaultValue: "" }).then((oA) => {\n          oA !== null && oA.trim() !== "" && _(HM(oA.trim())), ig.focus();\n        });\n      }, { icon: u.link }),'],
    // Application scope: one hook state per mounted editor app.
    ['  const yB = {\n    engine: K,',
      '  const simpleApp = simpleHookApp(A);\n  const yB = {\n    simple: simpleApp,\n    engine: K,'],
    ['s = _M(Q, S, yB), vB(), IA(), MA(), DA(), aA(), window.__cw = { doc: S, tree: void 0, engine: K, editor: s, createLayoutEngine: fQ, sampleDoc: Jc, stressDoc: Iw, persist: DB };\n  }, lA = () => {',
      's = _M(Q, S, yB), vB(), IA(), MA(), DA(), aA(), window.__cw = { doc: S, tree: void 0, engine: K, editor: s, createLayoutEngine: fQ, sampleDoc: Jc, stressDoc: Iw, persist: DB, handle: mB }, simpleApp.docChanged("load", !1, !1);\n  }, lA = () => {'],
    // The engine's existing debug global also carries the (hooked) handle.
    ['      T?.destroy(), s.destroy(), M.root.remove();\n    }\n  };\n  if (mA = {\n    ...mB,',
      '      T?.destroy(), s.destroy(), M.root.remove();\n    }\n  };\n  window.__cw.handle = mB;\n  if (mA = {\n    ...mB,'],
    // handle.openDocx rejects on failure (the caller reports it) instead of an
    // alert followed by a resolved promise and an unchanged model.
    ['  }, ZA = async (m) => {',
      '  }, ZA = async (m, simpleRethrow = !1) => {'],
    ['      alert(`Could not open "${GA}": ${kA instanceof Error ? kA.message : String(kA)}`);',
      '      if (simpleRethrow) throw kA;\n      simpleHookDialog(simpleApp, { kind: "alert", id: "document.open-failed", title: "Could not open document", message: `Could not open "${GA}": ${kA instanceof Error ? kA.message : String(kA)}` });'],
    ['    openDocx: (m) => ZA(m),',
      '    openDocx: (m) => ZA(m, !0),'],
    ['        s.inspectContentControl() || alert("Place the caret inside a content control first.");',
      '        s.inspectContentControl() || simpleHookDialog(simpleApp, { kind: "alert", id: "content-control.none", title: "Content control", message: "Place the caret inside a content control first." });'],
    ['      const R = prompt("List items (comma-separated):", "Yes, No, N/A");\n      if (R === null) return;\n      const O = R.split(",").map((p) => p.trim()).filter((p) => p.length > 0).map((p) => ({ display: p, value: p }));\n      s.dispatch($M("dropDown", { alias: "Drop-Down List", listItems: O })), s.focus();',
      '      simpleHookDialog(simpleApp, { kind: "prompt", id: "content-control.dropdown-items", title: "Drop-down list", message: "List items (comma-separated):", defaultValue: "Yes, No, N/A" }).then((R) => {\n        if (R === null) return void s.focus();\n        const O = R.split(",").map((p) => p.trim()).filter((p) => p.length > 0).map((p) => ({ display: p, value: p }));\n        s.dispatch($M("dropDown", { alias: "Drop-Down List", listItems: O })), s.focus();\n      });'],
    ['            const uA = prompt("Rename bookmark:", gA);\n            uA && uA.trim() && uA !== gA && (s.dispatch(Mt(gA, uA.trim())), s.focus());',
      '            simpleHookDialog(simpleApp, { kind: "prompt", id: "bookmark.rename", title: "Rename bookmark", message: "Rename bookmark:", defaultValue: gA }).then((uA) => {\n              uA && uA.trim() && uA !== gA && s.dispatch(Mt(gA, uA.trim())), s.focus();\n            });'],
    ['        const b = prompt("Bookmark name:");\n        b && b.trim() && (s.dispatch(gt(b.trim())), s.focus(), N());',
      '        simpleHookDialog(simpleApp, { kind: "prompt", id: "bookmark.add", title: "Add bookmark", message: "Bookmark name:", defaultValue: "" }).then((b) => {\n          b && b.trim() ? (s.dispatch(gt(b.trim())), s.focus(), N()) : s.focus();\n        });'],
    // Handle surface (also spread into every custom ribbon action context).
    ['    invalidateDecorations: () => s.invalidateDecorations(),\n    destroy: () => {',
      `    invalidateDecorations: () => s.invalidateDecorations(),
    simpleHooks: simpleHookVersion,
    getModelRevision: () => simpleApp.revision,
    insertImageBytes: async (m, zA, k = {}) => {
      const kA = await simpleHookPrepareImage(m, zA);
      return !!kA && s.simpleInsertImage(kA, k ?? {});
    },
    insertBlocks: (m, zA = {}) => s.simpleInsertBlocks(m, zA ?? {}),
    replaceBlock: (m, zA) => s.simpleReplaceBlock(m, zA),
    replaceImage: async (m, zA, k, kA = {}) => {
      const GA = await simpleHookPrepareImage(zA, k);
      return !!GA && s.simpleReplaceImage(m, GA, kA ?? {});
    },
    setDialogHandler: (m) => {
      simpleApp.dialog = typeof m == "function" ? m : null;
    },
    setInsertTextHook: (m) => {
      simpleApp.insertTextHook = typeof m == "function" ? m : null;
    },
    setBuiltinAutoCorrect: (m) => {
      simpleApp.autoCorrect = m !== !1;
    },
    deleteWord: (m) => s.simpleDeleteWord(m),
    undo: () => s.undo(),
    redo: () => s.redo(),
    canUndo: () => s.simpleCanUndo(),
    canRedo: () => s.simpleCanRedo(),
    setSelection: (m) => s.setSelection(m),
    focus: () => s.focus(),
    seedReview: (m) => s.seedReview(m),
    positionFromPoint: (m, zA) => s.simplePositionFromPoint(m, zA),
    destroy: () => {`],
  ];
  const fingerprint = require('node:crypto').createHash('sha256').update(JSON.stringify([helpers, edits])).digest('hex').slice(0, 16);
  return { marker, markerLine: `${marker} ${fingerprint}`, fingerprint, file: 'dist-lib/editorApp-vN1g1Ew1.js', helpers, edits };
}
