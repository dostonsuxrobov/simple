import assert from 'node:assert/strict';
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
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
if (process.argv[2]) {
  const result = await check(await readFile(path.resolve(process.argv[2])), 'Optional supplied fixture');
  const output = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../../.codex-tmp/watermark-qa');
  await mkdir(output, {recursive:true});
  await writeFile(path.join(output, 'patched-editor.pdf'), result.pdf);
  await writeFile(path.join(output, 'patched-roundtrip.docx'), result.bytes);
  await writeFile(path.join(output, 'layout.json'), JSON.stringify(result.layout, null, 2));
}
