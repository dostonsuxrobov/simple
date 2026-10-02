import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir, rm, copyFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import path from 'node:path';
import vm from 'node:vm';
import JSZip from 'jszip';
import { PDFDocument } from 'pdf-lib';
import { runImport } from '@forevka/wordcanvas/import';
import { runExport } from '@forevka/wordcanvas/export';
import { installMeasureHost } from '@forevka/wordcanvas/export/measure';
import { createLayoutEngine } from '../node_modules/@forevka/wordcanvas/dist-node/engine-N7X4PPLE.js';

// Entirely synthetic fixture. User documents are optional manual-QA inputs only.
const zip = new JSZip();
const namespaces = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>');
zip.file('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
zip.file('word/_rels/document.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>');
zip.file('word/_rels/header1.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/test.png"/></Relationships>');
zip.file('word/media/test.png', Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64'));
zip.file('word/document.xml', `<w:document ${namespaces}><w:body>${Array.from({length:20}, (_,i) => `<w:p><w:r><w:t>Synthetic body paragraph ${i + 1}: readable text.</w:t></w:r></w:p>`).join('')}<w:sectPr><w:headerReference w:type="default" r:id="rHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="776" w:right="720" w:bottom="864" w:left="720" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`);
zip.file('word/header1.xml', `<w:hdr ${namespaces}><w:p><w:r><w:drawing><wp:anchor behindDoc="1" relativeHeight="4" allowOverlap="1" simplePos="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="margin"><wp:align>center</wp:align></wp:positionV><wp:extent cx="6858000" cy="6858000"/><wp:wrapNone/><wp:docPr id="1" name="Synthetic background"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="test.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImage"><a:lum bright="70000" contrast="-70000"/></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6858000" cy="6858000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:hdr>`);
const bytes = await zip.generateAsync({type: 'uint8array'});
await installMeasureHost();
async function check(input, label) {
  const hash = createHash('sha256').update(input).digest('hex');
  const imported = runImport(input, undefined, {collectMediaBytes:true});
  const images = Object.fromEntries(imported.media.map(item => [item.src, item.bytes]));
  const doc = imported.doc;
  const before = JSON.stringify(doc);
  const header = doc.section.header.find(b => b.kind === 'image' && b.anchor?.behind);
  assert.ok(header, `${label}: original watermark belongs to header`);
  assert.deepEqual(header.simplePosition, {h:'center',v:'center'});
  assert.deepEqual(header.simpleLuminance, {bright:70000,contrast:-70000});
  const engine = createLayoutEngine();
  const layout = engine.layout(doc);
  assert.equal(layout.pages.length, 1, `${label}: background image must not create pages`);
  assert.equal(layout.pages[0].contentTopPx, doc.section.marginPx.top, `${label}: background must not reserve header height`);
  const placed = layout.pages[0].blocks.find(b => b.blockId === header.id);
  assert.ok(placed?.image.behind && placed.simpleBandImage);
  assert.equal(placed.x, doc.section.marginPx.left + (doc.section.pageWidthPx - doc.section.marginPx.left - doc.section.marginPx.right - 720) / 2);
  assert.equal(placed.y, doc.section.marginPx.top + (doc.section.pageHeightPx - doc.section.marginPx.top - doc.section.marginPx.bottom - 720) / 2);
  assert.deepEqual(placed.image.simpleLuminance, header.simpleLuminance);
  assert.equal(layout.pages[0].blocks[0].blockId, header.id, 'behind image sorts before body');
  assert.equal(JSON.stringify(doc), before, 'layout preserves source document model');
  assert.equal(createHash('sha256').update(input).digest('hex'), hash, 'source bytes unchanged');
  const exported = await runExport(doc, 'docx', images);
  const reopenedZip = await JSZip.loadAsync(exported.bytes);
  const headers = Object.keys(reopenedZip.files).filter(name => /^word\/header\d+\.xml$/.test(name));
  const headerXml = (await Promise.all(headers.map(name => reopenedZip.file(name).async('string')))).join('');
  assert.match(headerXml, /a:lum bright="70000" contrast="-70000"/);
  assert.equal((headerXml.match(/<wp:align>center<\/wp:align>/g) ?? []).length, 2);
  assert.match(headerXml, /behindDoc="1"/);
  const reimported = runImport(exported.bytes, undefined, {collectMediaBytes:true});
  assert.deepEqual(reimported.doc.section.header.find(b=>b.kind==='image').simpleLuminance, header.simpleLuminance);
  const pdf = await runExport(doc, 'pdf', images);
  assert.equal((await PDFDocument.load(pdf.bytes)).getPageCount(), 1, 'PDF uses same layout');
  assert.equal(JSON.stringify(doc), before, 'export preserves source model');
  console.log(`${label}: one page; centered behind-text image; original header, DrawingML settings and model preserved; DOCX/PDF exports verified.`);
  return {doc, layout, bytes: exported.bytes, pdf:pdf.bytes};
}
const synthetic = await check(bytes, 'Synthetic fixture');
// Deliberately in-flow header remains in-flow and reserves its original height.
const inline = structuredClone(synthetic.doc);
delete inline.section.header[0].anchor;
inline.section.header[0].heightPx = 120;
const inlineLayout = createLayoutEngine().layout(inline);
assert.ok(inlineLayout.pages[0].contentTopPx >= 120);
assert.ok(inlineLayout.pages[0].header.some(b=>b.image));
// Repeated header is placed on every page; first-page alternatives retain selection.
const repeated = structuredClone(synthetic.doc);
repeated.blocks = Array.from({length:100}, (_,i)=>({...structuredClone(repeated.blocks[i % repeated.blocks.length]),id:`long-${i}`}));
repeated.section.headerFirst = [];
const repeatedLayout = createLayoutEngine().layout(repeated);
assert.ok(repeatedLayout.pages.length > 1);
assert.equal(repeatedLayout.pages[0].blocks.filter(b=>b.simpleBandImage).length, 0);
for (const page of repeatedLayout.pages.slice(1)) assert.equal(page.blocks.filter(b=>b.simpleBandImage).length, 1);
console.log('Inline header reservation and repeated/first-page header selection verified.');
// Native editor creation uses U+000B for Shift+Enter and square wrapping without
// absolute anchor coordinates. Both must become standard, valid OOXML.
const created = structuredClone(synthetic.doc);
const picture = {...created.section.header[0], id:'square-created', widthPx:120, heightPx:80, wrap:'square', align:'left'};
delete picture.anchor;
delete picture.simplePosition;
delete picture.simpleLuminance;
const paragraph = {...structuredClone(created.blocks[0]), id:'soft-created'};
paragraph.runs[0].text = 'SOFT_BEFORE\vSOFT_AFTER\tTAB_AFTER';
paragraph.runs.push({text:'1', style:{...paragraph.runs[0].style, footnoteRef:'wrapped-note', verticalAlign:'super'}});
created.footnotes = {'wrapped-note': [{...structuredClone(paragraph), id:'wrapped-note-body', runs:[{text:'WRAPPED_FOOTNOTE_CONTENT', style:{...paragraph.runs[0].style, fontSizePx:12}}]}]};
created.blocks.unshift(picture, paragraph);
const noteLayout = createLayoutEngine().layout(created);
assert.ok(noteLayout.pages.some(page=>page.blocks.some(block=>block.blockId==='wrapped-note-body')), 'footnotes beside square images must appear on the page');
const createdBefore = JSON.stringify(created);
const sourceImport = runImport(bytes, undefined, {collectMediaBytes:true});
const creationExport = await runExport(created, 'docx', Object.fromEntries(sourceImport.media.map(item=>[item.src,item.bytes])));
const creationZip = await JSZip.loadAsync(creationExport.bytes);
for (const name of Object.keys(creationZip.files).filter(name=>name.endsWith('.xml'))) {
  assert.doesNotMatch(await creationZip.file(name).async('string'), /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/, `${name} is free of illegal XML controls`);
}
const creationXml = await creationZip.file('word/document.xml').async('string');
assert.match(creationXml, /<w:br\/>/);
assert.match(creationXml, /<wp:wrapSquare wrapText="bothSides"\/>/);
assert.equal(JSON.stringify(created), createdBefore, 'export never mutates the editor model');
const creationReopen = runImport(creationExport.bytes).doc;
assert.equal(creationReopen.blocks.find(block=>block.kind==='image').wrap, 'square');
assert.ok(creationReopen.blocks.some(block=>block.kind==='paragraph' && block.runs.map(run=>run.text).join('').includes('SOFT_BEFORE\vSOFT_AFTER')));
console.log('Created square wrapping and Shift+Enter survive valid DOCX export/reopen.');

// SIMPLE_HOOKS: the generic Simple layer over the browser editor. Browser behaviour
// (events, prompts, word deletion, pictures) is covered by an Electron smoke run;
// here the installed source is reviewed and the injected helpers run against the
// engine's real model functions.
const { simpleHooks } = createRequire(import.meta.url)('./patch-wordcanvas.cjs');
const packageDir = fileURLToPath(new URL('../node_modules/@forevka/wordcanvas/', import.meta.url));
const editorSource = (await readFile(path.join(packageDir, simpleHooks.file), 'utf8')).replace(/\r\n/g, '\n');
assert.equal(editorSource.split('\n').filter((line) => line.startsWith(simpleHooks.marker)).length, 1, 'SIMPLE_HOOKS applied exactly once');
assert.ok(editorSource.startsWith(`${simpleHooks.markerLine}\n`), 'installed editor carries the current SIMPLE_HOOKS revision');
for (const [, after] of simpleHooks.edits) assert.equal(editorSource.split(after).length - 1, 1, `SIMPLE_HOOKS edit present once: ${after.slice(0, 70)}`);
assert.equal((editorSource.match(/(^|[^.\w])prompt\(/gm) ?? []).length, 0, 'no engine prompt() remains (it throws in Electron)');
const remainingAlerts = [...editorSource.matchAll(/(?:^|[^.\w])alert\(`(Could not open the shared document|Share failed)/gm)].length;
assert.equal((editorSource.match(/(^|[^.\w])alert\(/gm) ?? []).length, remainingAlerts, 'only the collaboration/share alerts remain (unreachable offline)');
assert.equal(remainingAlerts, 2);
for (const event of ['"undo"', '"redo"', '"remote"', 'E.origin', '"load"']) assert.match(editorSource, new RegExp(`(simpleHookCommit\\(|docChanged\\()${event.replace(/[.]/g, '\\.')}`), `docchange emitted for ${event}`);

const importLine = editorSource.split('\n').find((line) => line.startsWith('import {') && line.includes('./paintStyle-'));
const aliases = Object.fromEntries([...importLine.matchAll(/(\w+) as (\w+)/g)].map(([, exported, local]) => [local, exported]));
const modelModule = await import(pathToFileURL(path.join(packageDir, 'dist-lib', importLine.match(/from "\.\/(paintStyle-[\w-]+\.js)"/)[1])).href);
const model = Object.fromEntries(['xM', 'XA', 'xB', 'Qg', 'Jg'].map((name) => [name, modelModule[aliases[name]]]));
for (const [name, value] of Object.entries(model)) assert.equal(typeof value, 'function', `editor import ${name} resolves`);
const logged = [];
const sandbox = {
  ...model,
  // Stand-ins for editor-module functions; the Electron smoke run covers the real ones.
  BI: (kind) => (state) => ({ ops: [{ type: 'setParaStyle', blockId: state.selection.focus.blockId, patch: { simpleTestList: kind } }], selectionAfter: state.selection, origin: 'command' }),
  UC: async (bytes, mime) => `media-${bytes.length}-${mime}`,
  hC: (mediaId) => `blob:${mediaId}`,
  createImageBitmap: async (blob) => {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    if (bytes[0] === 0) throw new Error('undecodable');
    return { width: 300, height: 100, close() {} };
  },
  Blob,
  console: { error: (...args) => logged.push(['error', ...args]), warn: (...args) => logged.push(['warn', ...args]) },
  window: { prompt: () => { throw new Error('prompt() is not supported.'); }, confirm: () => true, alert: (message) => logged.push(['alert', message]) },
};
const hooks = vm.runInNewContext(`${simpleHooks.helpers}\n;({ simpleHookApp, simpleHookDialog, simpleHookPrepareImage, simpleHookMapDocument, simpleHookRevise, simpleHookRevisionMap, simpleHookRemint, simpleHookSelectionIn, simpleHookPatchRuns, simpleHookEditTransaction })`, sandbox);
const plain = (value) => JSON.parse(JSON.stringify(value));

// docchange events: monotonic revision, typed custom event, isolated listeners.
const emitted = [];
const hookApp = hooks.simpleHookApp({ onEvent: (event) => emitted.push(event) });
hookApp.docChanged('typing', true, false);
hookApp.docChanged('undo', 1, 0);
assert.deepEqual(plain(emitted), [
  { type: 'custom', name: 'simple:docchange', payload: { revision: 1, origin: 'typing', canUndo: true, canRedo: false } },
  { type: 'custom', name: 'simple:docchange', payload: { revision: 2, origin: 'undo', canUndo: true, canRedo: false } },
]);
assert.equal(hooks.simpleHookApp({ onEvent: () => { throw new Error('listener'); } }).docChanged('command').revision, 1, 'a failing listener cannot break the edit');
assert.equal(logged.pop()[0], 'error');
assert.equal(hooks.simpleHookApp(undefined).docChanged('load').revision, 1);

// Dialog bridge: async Simple handler first; Electron's throwing prompt() is contained.
const dialogApp = hooks.simpleHookApp({});
const promptRequest = { kind: 'prompt', id: 'hyperlink.insert', message: 'Link URL:', defaultValue: '' };
assert.equal(await hooks.simpleHookDialog(dialogApp, promptRequest), null, 'no handler: Electron prompt() throws, the edit is cancelled');
assert.match(String(logged.pop()[1]), /window\.prompt\(\) is unavailable/);
assert.equal(await hooks.simpleHookDialog(dialogApp, { kind: 'confirm', id: 'x', message: 'm' }), true);
assert.equal(await hooks.simpleHookDialog(dialogApp, { kind: 'alert', id: 'x', message: 'Heads up' }), undefined);
assert.deepEqual(logged.pop(), ['alert', 'Heads up']);
const seenRequests = [];
dialogApp.dialog = async (request) => { seenRequests.push(request); request.message = 'mutated'; return 'https://example.org'; };
assert.equal(await hooks.simpleHookDialog(dialogApp, promptRequest), 'https://example.org');
assert.equal(promptRequest.message, 'Link URL:', 'handlers receive a copy of the request');
assert.deepEqual(plain(seenRequests[0]), { ...promptRequest, message: 'mutated' });
dialogApp.dialog = () => 42;
assert.equal(await hooks.simpleHookDialog(dialogApp, promptRequest), null, 'non-string prompt results cancel');
dialogApp.dialog = () => 'yes';
assert.equal(await hooks.simpleHookDialog(dialogApp, { kind: 'confirm', id: 'x', message: 'm' }), false, 'only true confirms');
dialogApp.dialog = () => { throw new Error('sync'); };
assert.equal(await hooks.simpleHookDialog(dialogApp, promptRequest), null);
dialogApp.dialog = async () => { throw new Error('async'); };
assert.equal(await hooks.simpleHookDialog(dialogApp, promptRequest), null);
assert.deepEqual(logged.splice(0).map(([level]) => level), ['error', 'error'], 'handler failures are reported, never thrown into the engine');

// Insert-text hook corrections: one transaction, right to left, caret mapped.
const runStyle = { fontFamily: 'Calibri', fontSizePx: 14.667, color: '#111111' };
const hookDoc = { section: { pageWidthPx: 816, pageHeightPx: 1056, marginPx: { top: 96, right: 96, bottom: 96, left: 96 } }, blocks: [{ kind: 'paragraph', id: 'p1', revision: 0, style: { align: 'left' }, runs: [{ text: 'Say ', style: runStyle }, { text: 'teh', style: { ...runStyle, italic: true } }, { text: ' 1/2 now', style: runStyle }] }] };
const caretAt = (offset) => ({ anchor: { blockId: 'p1', offset }, focus: { blockId: 'p1', offset } });
const transaction = hooks.simpleHookEditTransaction({ doc: hookDoc, selection: caretAt(15) }, [
  { type: 'format', start: 0, end: 3, style: { bold: true } },
  { type: 'replace', start: 4, end: 7, text: 'the' },
  { type: 'replace', start: 8, end: 11, text: '½' },
  { type: 'list', kind: 'bullet' },
], 'p1');
assert.deepEqual(plain(transaction.ops.map((op) => op.type)), ['deleteRange', 'insertText', 'deleteRange', 'insertText', 'setRuns', 'setParaStyle']);
assert.equal(transaction.origin, 'command');
let corrected = hookDoc;
for (const op of transaction.ops) corrected = model.xM(corrected, op).doc;
assert.equal(model.xB(corrected.blocks[0].runs), 'Say the ½ now');
assert.equal(corrected.blocks[0].runs.find((run) => run.text.includes('the')).style.italic, true, 'a replacement inherits the replaced text style');
assert.equal(corrected.blocks[0].runs.find((run) => run.text.startsWith('Say')).style.bold, true);
assert.equal(corrected.blocks[0].style.simpleTestList, 'bullet', 'list edits apply after text edits');
assert.deepEqual(plain(transaction.selectionAfter), caretAt(13), 'the caret follows the shorter replacement');
assert.equal(hooks.simpleHookEditTransaction({ doc: hookDoc, selection: caretAt(1) }, [], 'p1'), null);
assert.equal(hooks.simpleHookEditTransaction({ doc: hookDoc, selection: caretAt(1) }, [{ type: 'replace', blockId: 'missing', start: 0, end: 1, text: 'x' }], 'p1'), null);
const clamped = hooks.simpleHookEditTransaction({ doc: hookDoc, selection: caretAt(1) }, [{ type: 'replace', start: -5, end: 999, text: 'All new' }], 'p1');
assert.deepEqual(plain(clamped.ops.map((op) => op.type === 'deleteRange' ? [op.start, op.end] : op.text)), [[0, 15], 'All new']);
assert.deepEqual(plain(hooks.simpleHookPatchRuns([{ text: 'abcdef', style: { a: 1 } }], 2, 4, { b: 2 })), [{ text: 'ab', style: { a: 1 } }, { text: 'cd', style: { a: 1, b: 2 } }, { text: 'ef', style: { a: 1 } }]);

// Block replacement: nested, banded and note blocks; containers re-revised so layout
// caches (keyed by id + revision) never serve a stale measurement.
const cellParagraph = { kind: 'paragraph', id: 'cell-p', revision: 4, style: {}, runs: [{ text: 'CELL', style: runStyle }] };
const treeDoc = {
  section: { ...hookDoc.section, header: [{ kind: 'paragraph', id: 'head-p', revision: 1, style: {}, runs: [{ text: 'HEAD', style: runStyle }] }] },
  blocks: [hookDoc.blocks[0], { kind: 'table', id: 'table-1', revision: 2, rows: [{ cells: [{ id: 'cell-1', blocks: [cellParagraph] }] }] }],
  footnotes: { n1: [{ kind: 'paragraph', id: 'note-p', revision: 0, style: {}, runs: [{ text: 'NOTE', style: runStyle }] }] },
};
const treeBefore = JSON.stringify(treeDoc);
const replaceWith = (id, make) => hooks.simpleHookMapDocument(treeDoc, (block) => block?.id === id ? hooks.simpleHookRevise(make(block), hooks.simpleHookRevisionMap(block)) : block);
const cellReplaced = replaceWith('cell-p', (old) => ({ ...old, runs: [{ text: 'X', style: runStyle }] }));
assert.equal(cellReplaced.blocks[0], treeDoc.blocks[0], 'untouched blocks are shared');
assert.equal(cellReplaced.section, treeDoc.section);
assert.equal(cellReplaced.blocks[1].revision, 3, 'the containing table is re-revised');
assert.equal(cellReplaced.blocks[1].rows[0].cells[0].blocks[0].revision, 5);
assert.equal(model.xB(model.XA(cellReplaced, 'cell-p').runs), 'X');
const headReplaced = replaceWith('head-p', (old) => ({ ...old, revision: 0 }));
assert.equal(headReplaced.blocks, treeDoc.blocks);
assert.equal(headReplaced.section.header[0].revision, 2, 'a replacement never reuses the old revision');
assert.equal(replaceWith('note-p', (old) => ({ ...old })).footnotes.n1[0].revision, 1);
assert.equal(hooks.simpleHookMapDocument(treeDoc, (block) => block), treeDoc, 'no match keeps the document identity');
const nestedRevisions = plain([...hooks.simpleHookRevisionMap(treeDoc.blocks[1]).entries()]);
assert.deepEqual(nestedRevisions, [['table-1', 2], ['cell-p', 4]]);
const revisedTable = hooks.simpleHookRevise({ ...treeDoc.blocks[1], revision: 0, rows: [{ cells: [{ id: 'cell-1', blocks: [{ ...cellParagraph, revision: 0 }] }] }] }, hooks.simpleHookRevisionMap(treeDoc.blocks[1]));
assert.deepEqual([revisedTable.revision, revisedTable.rows[0].cells[0].blocks[0].revision], [3, 5], 'nested blocks are re-revised against the blocks they replace');
const reminted = hooks.simpleHookRemint(treeDoc.blocks[1]);
const remintedIds = [reminted.id, reminted.rows[0].cells[0].id, reminted.rows[0].cells[0].blocks[0].id];
assert.equal(new Set(remintedIds).size, 3);
assert.ok(remintedIds.every((id) => !['table-1', 'cell-1', 'cell-p'].includes(id)), 'inserted blocks get fresh ids');
assert.equal(JSON.stringify(treeDoc), treeBefore, 'helpers never mutate the source model');
const keptSelection = caretAt(2);
assert.equal(hooks.simpleHookSelectionIn(cellReplaced, keptSelection, cellReplaced.blocks[1]), keptSelection);
assert.deepEqual(plain(hooks.simpleHookSelectionIn(cellReplaced, { anchor: { blockId: 'cell-p', offset: 4 }, focus: { blockId: 'cell-p', offset: 4 } }, cellReplaced.blocks[1])), { anchor: { blockId: 'cell-p', offset: 0 }, focus: { blockId: 'cell-p', offset: 0 } }, 'a caret beyond the new text moves into the replacement');
assert.equal(hooks.simpleHookSelectionIn(cellReplaced, { anchor: { blockId: 'gone', offset: 0 }, focus: { blockId: 'gone', offset: 0 } }, { kind: 'image' }), null);

// Image bytes are decoded before they enter the media store.
assert.deepEqual(plain(await hooks.simpleHookPrepareImage(new Uint8Array([1, 2, 3]), 'IMAGE/PNG')), { src: 'blob:media-3-image/png', mediaId: 'media-3-image/png', mime: 'image/png', widthPx: 300, heightPx: 100 });
assert.equal((await hooks.simpleHookPrepareImage(new Uint8Array([7, 8]).buffer, 'text/html')).mime, 'image/png', 'non-image MIME types fall back to PNG');
assert.equal(await hooks.simpleHookPrepareImage(new Uint8Array([0, 1]), 'image/png'), null, 'undecodable bytes are refused');
assert.equal(await hooks.simpleHookPrepareImage(new Uint8Array(), 'image/png'), null);
assert.equal(await hooks.simpleHookPrepareImage('not bytes', 'image/png'), null);
console.log('SIMPLE_HOOKS source, docchange events, dialog bridge, correction transactions, block replacement and image preparation verified.');

// Idempotency, version guard and stale-revision guard of the patcher itself.
const patcher = fileURLToPath(new URL('./patch-wordcanvas.cjs', import.meta.url));
const patcherInputs = ['package.json', 'dist-node/chunk-S3H2FFFI.js', 'dist-lib/pipeline-CsMT0pHL.js', 'dist-lib/assets/worker-D0pm0kNa.js', 'dist-node/chunk-CSJ442BN.js', 'dist-lib/engine-BwNLlumM.js', 'dist-lib/assets/worker-DueeItgB.js', 'dist-lib/pipeline-BEYkJ78H.js', 'dist-node/export.js', 'dist-lib/editorApp-vN1g1Ew1.js', 'dist-node/fonts/NotoSansSC-Regular.ttf'];
const digests = async (dir) => Object.fromEntries(await Promise.all(patcherInputs.map(async (file) => [file, createHash('sha256').update(await readFile(path.join(dir, file))).digest('hex')])));
const runPatcher = (packageRoot) => spawnSync(process.execPath, [patcher], { encoding: 'utf8', env: { ...process.env, ...(packageRoot ? { SIMPLE_WORDCANVAS_PACKAGE_DIR: packageRoot } : {}) } });
const installedDigests = await digests(packageDir);
const rerun = runPatcher();
assert.equal(rerun.status, 0, rerun.stderr);
assert.match(rerun.stdout, /already applied/);
assert.deepEqual(await digests(packageDir), installedDigests, 'a second run changes nothing');
const scratchRoot = fileURLToPath(new URL('../tmp/wordcanvas-patch-test/', import.meta.url));
const stage = async (name, sourceDir) => {
  const dir = path.join(scratchRoot, name);
  await rm(dir, { recursive: true, force: true });
  for (const file of patcherInputs) {
    await mkdir(path.dirname(path.join(dir, file)), { recursive: true });
    await copyFile(path.join(sourceDir, file), path.join(dir, file));
  }
  return dir;
};
try {
  const future = path.join(scratchRoot, 'future-version');
  await rm(future, { recursive: true, force: true });
  await mkdir(future, { recursive: true });
  await writeFile(path.join(future, 'package.json'), JSON.stringify({ name: '@forevka/wordcanvas', version: '0.12.1' }));
  const refused = runPatcher(future);
  assert.notEqual(refused.status, 0, 'an unreviewed engine version is refused');
  assert.match(refused.stderr, /requires reviewed version 0\.12\.0/);
  const stale = await stage('stale-hooks', packageDir);
  const staleEditor = path.join(stale, simpleHooks.file);
  await writeFile(staleEditor, (await readFile(staleEditor, 'utf8')).replace(simpleHooks.markerLine, `${simpleHooks.marker} 0000000000000000`));
  const staleBefore = await digests(stale);
  const staleRun = runPatcher(stale);
  assert.notEqual(staleRun.status, 0, 'an older SIMPLE_HOOKS revision is not silently kept');
  assert.match(staleRun.stderr, /Reinstall @forevka\/wordcanvas 0\.12\.0/);
  assert.deepEqual(await digests(stale), staleBefore, 'nothing is written when the guard fails');
  // Upgrade review: SIMPLE_WORDCANVAS_PRISTINE_DIR=<extracted 0.12.0 package> applies every
  // adapter to pristine sources twice and compares the result with the installed modules.
  if (process.env.SIMPLE_WORDCANVAS_PRISTINE_DIR) {
    const pristine = await stage('pristine', path.resolve(process.env.SIMPLE_WORDCANVAS_PRISTINE_DIR));
    const first = runPatcher(pristine);
    assert.equal(first.status, 0, first.stderr);
    assert.match(first.stdout, /patched 9 modules/);
    const second = runPatcher(pristine);
    assert.equal(second.status, 0, second.stderr);
    assert.match(second.stdout, /already applied/);
    assert.deepEqual(await digests(pristine), installedDigests, 'pristine + patch equals the installed modules');
    console.log('Pristine WordCanvas 0.12.0 patched cleanly twice and matches the installed modules.');
  }
} finally {
  await rm(scratchRoot, { recursive: true, force: true });
}
console.log('Patcher is idempotent, refuses unreviewed engine versions and stale SIMPLE_HOOKS revisions.');
if (process.argv[2]) {
  const result = await check(await readFile(path.resolve(process.argv[2])), 'Optional supplied fixture');
  const output = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../.codex-tmp/watermark-qa');
  await mkdir(output, {recursive:true});
  await writeFile(path.join(output, 'patched-editor.pdf'), result.pdf);
  await writeFile(path.join(output, 'patched-roundtrip.docx'), result.bytes);
  await writeFile(path.join(output, 'layout.json'), JSON.stringify(result.layout, null, 2));
}
