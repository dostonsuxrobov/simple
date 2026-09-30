/* WordCanvas 0.12.0 adapter: anchored header images occupy no flow space.
 * The MIT dependency ships generated modules. Keep this patch explicit, pinned,
 * idempotent, and fail closed on upstream changes; never alter the source model.
 * See wordcanvas-patch.md for scope, sources, and remaining rendering limits.
 */
const fs = require('node:fs');
const path = require('node:path');
const base = path.resolve(__dirname, '../node_modules/@forevka/wordcanvas');
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
for (const [file, source] of pending) fs.writeFileSync(file, source);
console.log(`WordCanvas 0.12.0 header image adapter: ${pending.size ? `patched ${pending.size} modules` : 'already applied'}.`);
