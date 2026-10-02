// Open, paste and drop smoke test for Simple Docs (Electron + CDP, built renderer).
// DOC-017 / DOC-SIE-15: opens copies of the sample .md and .html and generated .txt
// (UTF-8 with BOM, Windows-1252), .rtf, .odt, .dotx and .docm files; checks their
// structure, the one-line note, and that Save writes a separate .docx while the
// original stays byte-for-byte unchanged. DOC-008 / DOC-012 / DOC-SIE-25: pastes HTML
// with a heading, lists, a 2x2 table and a picture (one undo step, real numbering in
// the exported DOCX), pastes a screenshot, Ctrl+Shift+V, the paste options chip,
// a web address over selected text, the ribbon's Paste, and drops a picture (no error).
//
// Run `npm run build:web` first (or point SIMPLE_DOCS_DIST at another build), then
// `node tests/import-paste-smoke.mjs`. SIMPLE_SMOKE_OUT moves the work folder (default
// tmp/import-paste-smoke). Windows open off-screen, nothing touches the system
// clipboard, nothing is printed, and every process is killed at the end.
import { spawn, execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { copyFile, mkdir, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { deflateSync } from 'node:zlib';
import assert from 'node:assert/strict';
import JSZip from 'jszip';

const require = createRequire(import.meta.url);
const root = fileURLToPath(new URL('..', import.meta.url));
const output = process.env.SIMPLE_SMOKE_OUT ? path.resolve(process.env.SIMPLE_SMOKE_OUT) : path.join(root, 'tmp', 'import-paste-smoke');
const dist = process.env.SIMPLE_DOCS_DIST ? path.resolve(process.env.SIMPLE_DOCS_DIST) : null;
const examples = path.join(root, '..', 'Simple test examples');
const fixtures = path.join(output, 'fixtures');
await rm(output, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
await mkdir(path.join(output, 'profile'), { recursive: true });
await mkdir(fixtures, { recursive: true });

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
/** A valid RGBA PNG of one colour. */
function png(width, height, [r, g, b, a] = [200, 40, 40, 255]) {
  const row = Buffer.alloc(width * 4 + 1);
  for (let x = 0; x < width; x += 1) row.set([r, g, b, a], 1 + x * 4);
  const raw = Buffer.concat(Array.from({ length: height }, () => row));
  const chunk = (type, data) => {
    const length = Buffer.alloc(4);
    length.writeUInt32BE(data.length);
    const body = Buffer.concat([Buffer.from(type, 'latin1'), data]);
    const crc = Buffer.alloc(4);
    crc.writeUInt32BE(crc32(body));
    return Buffer.concat([length, body, crc]);
  };
  const header = Buffer.alloc(13);
  header.writeUInt32BE(width, 0);
  header.writeUInt32BE(height, 4);
  header.set([8, 6, 0, 0, 0], 8);
  return Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk('IHDR', header), chunk('IDAT', deflateSync(raw)), chunk('IEND', Buffer.alloc(0))]);
}
const PNG_4x2 = png(40, 20);

// ---- fixtures ---------------------------------------------------------------------------

const files = {};
for (const name of ['Complex document.md', 'Complex document.html']) {
  const source = path.join(examples, name);
  if (!existsSync(source)) continue;
  files[name] = path.join(fixtures, name);
  await copyFile(source, files[name]);
}
files.txt = path.join(fixtures, 'Notes UTF8.txt');
await writeFile(files.txt, Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('First line café\r\n\tIndented ünïcode line\r\nThird line\r\n', 'utf8')]));
files.cp1252 = path.join(fixtures, 'Legacy 1252.txt');
await writeFile(files.cp1252, Buffer.from('Caf\xe9 cr\xe8me \x93quoted\x94 \x80 5\r\nSecond line\r\n', 'latin1'));
files.rtf = path.join(fixtures, 'Report.rtf');
await writeFile(files.rtf, Buffer.from(String.raw`{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fswiss\fcharset0 Arial;}}
\pard\sa200\b\fs32 RTF Title\b0\fs22\par
\pard Body with {\b bold words} here.\par
\trowd\cellx3000\cellx6000\pard\intbl R1C1\cell R1C2\cell\row
\trowd\cellx3000\cellx6000\pard\intbl R2C1\cell R2C2\cell\row
\pard {\*\shppict{\pict\pngblip\picw4\pich2\picwgoal600\pichgoal300 ${PNG_4x2.toString('hex')}}}\par
\pard After the picture.\par
}`, 'latin1'));

const ODF_NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"';
{
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.text', { compression: 'STORE' });
  zip.file('content.xml', `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${ODF_NS} office:version="1.3"><office:automatic-styles><text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1"/></text:list-style></office:automatic-styles><office:body><office:text>
<text:h text:outline-level="1">ODT Heading</text:h><text:p>ODT body text.</text:p>
<text:list text:style-name="L1"><text:list-item><text:p>ODT first</text:p></text:list-item><text:list-item><text:p>ODT second</text:p></text:list-item></text:list>
<table:table table:name="T1"><table:table-column table:number-columns-repeated="2"/><table:table-row><table:table-cell><text:p>O1</text:p></table:table-cell><table:table-cell><text:p>O2</text:p></table:table-cell></table:table-row><table:table-row><table:table-cell><text:p>O3</text:p></table:table-cell><table:table-cell><text:p>O4</text:p></table:table-cell></table:table-row></table:table>
</office:text></office:body></office:document-content>`);
  zip.file('META-INF/manifest.xml', '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>');
  files.odt = path.join(fixtures, 'Minutes.odt');
  await writeFile(files.odt, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

async function wordPackage(file, mainContentType, paragraphs, extra = {}) {
  const zip = new JSZip();
  const overrides = Object.entries(extra).filter(([, value]) => value.contentType).map(([name, value]) => `<Override PartName="/${name}" ContentType="${value.contentType}"/>`).join('');
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/word/document.xml" ContentType="${mainContentType}"/>${overrides}</Types>`);
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs.map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  for (const [name, value] of Object.entries(extra)) zip.file(name, value.data);
  await writeFile(file, await zip.generateAsync({ type: 'nodebuffer' }));
}
files.dotx = path.join(fixtures, 'Invoice template.dotx');
await wordPackage(files.dotx, 'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml', ['TEMPLATE_BODY']);
files.docm = path.join(fixtures, 'Macro report.docm');
await wordPackage(files.docm, 'application/vnd.ms-word.document.macroEnabled.main+xml', ['MACRO_BODY'], { 'word/vbaProject.bin': { data: Buffer.from('not a real project') } });
// Like a document saved by Word: a stylesheet with Normal only (Heading 1 is added when pasted).
files.blank = path.join(fixtures, 'Paste target.docx');
await wordPackage(files.blank, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml', ['Hello world', 'Second paragraph'], {
  'word/styles.xml': {
    contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml',
    data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>',
  },
  'word/_rels/document.xml.rels': {
    data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
  },
});

const originals = new Map();
for (const file of Object.values(files)) originals.set(file, await readFile(file));

// ---- host -------------------------------------------------------------------------------

const hostPath = path.join(output, 'host.cjs');
await writeFile(hostPath, `const { app, dialog, shell, BrowserWindow } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
app.setPath('userData', path.join(__dirname, 'profile'));
const answers = () => { try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'answers.json'), 'utf8')); } catch { return {}; } };
const record = (name, options) => fs.appendFileSync(path.join(__dirname, 'dialogs.log'), JSON.stringify({ name, options }) + '\\n');
dialog.showOpenDialog = async (...args) => { record('open', args[args.length - 1]); const file = answers().open; return file ? { canceled: false, filePaths: [file] } : { canceled: true, filePaths: [] }; };
dialog.showSaveDialog = async (...args) => { record('save', args[args.length - 1]); const file = answers().save; return file ? { canceled: false, filePath: file } : { canceled: true }; };
dialog.showMessageBox = async () => ({ response: 0 });
shell.openExternal = async () => {};
const show = BrowserWindow.prototype.show;
BrowserWindow.prototype.show = function () { if (process.env.SIMPLE_SMOKE_SHOW) return show.call(this); try { this.setPosition(-3200, 40); this.showInactive(); } catch {} };
app.on('browser-window-created', (_event, window) => { try { window.webContents.setBackgroundThrottling(false); } catch {} });
require(${JSON.stringify(path.join(root, 'electron', 'main.cjs'))});
`);
const answer = (value) => writeFile(path.join(output, 'answers.json'), JSON.stringify(value));
const dialogLog = async () => {
  try {
    return (await readFile(path.join(output, 'dialogs.log'), 'utf8')).trim().split('\n').filter(Boolean).map((line) => JSON.parse(line));
  } catch {
    return [];
  }
};

const port = 24900 + Math.floor(Math.random() * 600);
const env = { ...process.env };
delete env.ELECTRON_RUN_AS_NODE;
if (dist) env.VITE_DEV_SERVER_URL = pathToFileURL(path.join(dist, 'index.html')).href;
const child = spawn(require('electron'), [`--remote-debugging-port=${port}`, `--user-data-dir=${path.join(output, 'profile')}`, hostPath], { cwd: root, env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
let stderr = '';
child.stderr.on('data', (data) => { stderr += String(data); });
const killAll = () => { try { execFileSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { try { child.kill(); } catch {} } };
const hardStop = setTimeout(() => { console.error('Import and paste smoke test timed out.'); killAll(); process.exit(124); }, Number(process.env.SIMPLE_SMOKE_TIMEOUT_MS || 300_000));

const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const seenTargets = new Set();
async function connect() {
  let target;
  const end = Date.now() + 45_000;
  while (!target && Date.now() < end) {
    try { target = (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find((candidate) => candidate.type === 'page' && !seenTargets.has(candidate.id)); } catch {}
    if (!target) await pause(150);
  }
  assert.ok(target, `A renderer window opens. ${stderr.slice(-1500)}`);
  seenTargets.add(target.id);
  const socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }); });
  let id = 0;
  const pending = new Map();
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data));
    const request = pending.get(message.id);
    if (!request) return;
    pending.delete(message.id);
    message.error ? request.reject(new Error(message.error.message)) : request.resolve(message.result);
  });
  // A window that closes never answers: settle what it still owes.
  socket.addEventListener('close', () => {
    for (const request of pending.values()) request.reject(new Error('The window closed.'));
    pending.clear();
  });
  const send = (method, params = {}) => new Promise((resolve, reject) => { const request = ++id; pending.set(request, { resolve, reject }); socket.send(JSON.stringify({ id: request, method, params })); });
  const evaluate = async (expression) => {
    const result = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true });
    if (result.exceptionDetails) throw new Error(`${result.exceptionDetails.text} ${result.exceptionDetails.exception?.description ?? ''}`);
    return result.result.value;
  };
  const until = async (expression, label, timeout = 30_000) => {
    const stop = Date.now() + timeout;
    let last;
    while (Date.now() < stop) {
      try { last = await evaluate(expression); if (last) return last; } catch (error) { last = error.message; }
      await pause(100);
    }
    throw new Error(`Timed out: ${label} (last: ${JSON.stringify(last)})`);
  };
  return { evaluate, until, close: () => socket.close() };
}

const results = [];
async function check(name, run) {
  const started = Date.now();
  try {
    await run();
    results.push({ name, ok: true, ms: Date.now() - started });
    console.log(`ok   ${name}`);
  } catch (error) {
    results.push({ name, ok: false, error: error.message });
    console.log(`FAIL ${name}\n     ${error.message}`);
  }
}

// Page-side helpers (WordCanvas exposes its handle as window.__cw.handle).
const H = 'window.__cw.handle';
const ready = `Boolean(window.__cw && window.__cw.handle && document.getElementById('loading-overlay').hidden)`;
const opened = `(document.getElementById('welcome').hidden && document.getElementById('loading-overlay').hidden)`;
const summary = `(() => {
  const doc = ${H}.getDocument();
  const text = (block) => (block.runs || []).map((run) => run.text).join('');
  const paragraphs = [];
  const visit = (blocks) => { for (const block of blocks) { if (block.kind === 'paragraph') paragraphs.push(block); if (block.kind === 'table') for (const row of block.rows) for (const cell of row.cells) visit(cell.blocks); } };
  visit(doc.blocks);
  const lists = paragraphs.filter((block) => block.style.list);
  return {
    texts: paragraphs.map(text),
    headings: paragraphs.filter((block) => /^Heading\\d$/.test(block.style.namedStyle || '')).map(text),
    tables: doc.blocks.filter((block) => block.kind === 'table').map((table) => [table.rows.length, table.rows[0].cells.length]),
    images: doc.blocks.filter((block) => block.kind === 'image').length,
    lists: lists.map((block) => [text(block), Boolean(doc.lists && doc.lists[block.style.list.listId])]),
    title: document.getElementById('document-title').textContent,
    notice: document.getElementById('document-compatibility').hidden ? '' : document.getElementById('document-notice-text').textContent,
    toast: document.getElementById('toast').hidden ? '' : document.getElementById('toast').textContent,
    toastTone: document.getElementById('toast').dataset.tone || '',
    state: document.getElementById('save-state').textContent,
  };
})()`;

async function openInWindow(first, file) {
  await first.evaluate(`window.simpleDocs.openInNewWindow(${JSON.stringify(file)}).then(() => true)`);
  const page = await connect();
  await page.until(`${ready} && ${opened}`, `${path.basename(file)} opens`, 45_000);
  return page;
}

async function closeWindow(page) {
  // The window may close before it answers.
  await Promise.race([page.evaluate('window.simpleDocs.confirmClose(), true').catch(() => {}), pause(1500)]);
  page.close();
}

let main;
try {
  main = await connect();
  await main.until(ready, 'the editor starts', 60_000);

  // ---- opening ------------------------------------------------------------------------

  await check('the Open dialog offers every format, "All supported documents" first', async () => {
    await answer({});
    await main.evaluate(`document.getElementById('open-document').click(), true`);
    await main.until(`true`, 'dialog');
    let entry;
    const stop = Date.now() + 10_000;
    while (!entry && Date.now() < stop) { entry = (await dialogLog()).find((item) => item.name === 'open'); if (!entry) await pause(100); }
    assert.ok(entry, 'the open dialog was shown');
    const [all] = entry.options.filters;
    assert.equal(all.name, 'All supported documents');
    for (const extension of ['docx', 'docm', 'dotx', 'dotm', 'doc', 'rtf', 'odt', 'txt', 'md', 'html', 'htm']) assert.ok(all.extensions.includes(extension), `${extension} is offered`);
    assert.equal(entry.options.filters.at(-1).name, 'All files');
  });

  if (files['Complex document.md']) {
    await check('a copy of the sample Markdown opens with its headings, tables, list and picture', async () => {
      const page = await openInWindow(main, files['Complex document.md']);
      const doc = await page.evaluate(summary);
      assert.equal(doc.title, 'Complex document');
      assert.deepEqual(doc.headings, ['created', 'Section 2 first-page header', 'Section 2 first-page footer', 'Footnote 1']);
      assert.deepEqual(doc.tables, [[3, 3], [2, 2]]);
      assert.deepEqual(doc.lists, [['BULLET_ONE', true], ['BULLET_TWO', true]]);
      assert.equal(doc.images, 1);
      assert.match(doc.notice, /^Markdown · Saving creates a \.docx file; “Complex document\.md” stays unchanged\./);
      assert.equal(doc.state, 'Not saved yet');
      await closeWindow(page);
    });
  }

  if (files['Complex document.html']) {
    await check('a copy of the sample web page opens with its headings, tables, list and picture', async () => {
      const page = await openInWindow(main, files['Complex document.html']);
      const doc = await page.evaluate(summary);
      assert.equal(doc.title, 'Complex document');
      assert.deepEqual(doc.headings, ['Section 2 first-page header', 'Section 2 first-page footer', 'Footnote 1']);
      assert.deepEqual(doc.tables, [[3, 3], [2, 2]]);
      assert.deepEqual(doc.lists, [['BULLET_ONE', true], ['BULLET_TWO', true]]);
      assert.equal(doc.images, 1);
      assert.match(doc.notice, /^Web page · Saving creates a \.docx file/);
      await closeWindow(page);
    });
  }

  await check('UTF-8 and Windows-1252 text files open line by line', async () => {
    const page = await openInWindow(main, files.txt);
    const doc = await page.evaluate(summary);
    assert.deepEqual(doc.texts, ['First line café', '\tIndented ünïcode line', 'Third line']);
    assert.equal(doc.title, 'Notes UTF8');
    assert.match(doc.notice, /^Plain text ·/);
    await closeWindow(page);
    const legacy = await openInWindow(main, files.cp1252);
    assert.deepEqual((await legacy.evaluate(summary)).texts, ['Café crème “quoted” € 5', 'Second line']);
    await closeWindow(legacy);
  });

  await check('an RTF file opens with its bold text, table and picture, and saves as a separate .docx', async () => {
    const page = await openInWindow(main, files.rtf);
    const doc = await page.evaluate(summary);
    assert.ok(doc.texts.includes('RTF Title') && doc.texts.includes('Body with bold words here.'));
    assert.deepEqual(doc.tables, [[2, 2]]);
    assert.equal(doc.images, 1);
    assert.ok(await page.evaluate(`${H}.getDocument().blocks.some((block) => block.kind === 'paragraph' && block.runs.some((run) => run.text === 'bold words' && run.style.bold))`));
    // Save: a question, then name.docx beside the original, which stays unchanged.
    await page.evaluate(`document.getElementById('save-button').click(), true`);
    await page.until(`!document.getElementById('choice-modal').hidden`, 'the save question');
    assert.match(await page.evaluate(`document.getElementById('choice-title').textContent`), /Save as “Report\.docx”\?/);
    await page.evaluate(`document.querySelector('#choice-actions [data-choice="save"]').click(), true`);
    await page.until(`document.getElementById('save-state').textContent === 'Saved locally'`, 'saved');
    const saved = path.join(fixtures, 'Report.docx');
    const zip = await JSZip.loadAsync(await readFile(saved));
    assert.ok(zip.file('word/document.xml'), 'a Word package was written');
    assert.ok((await readFile(files.rtf)).equals(originals.get(files.rtf)), 'the RTF file is unchanged');
    await closeWindow(page);
  });

  await check('an OpenDocument file opens with its heading, numbered list and table', async () => {
    const page = await openInWindow(main, files.odt);
    const doc = await page.evaluate(summary);
    assert.deepEqual(doc.headings, ['ODT Heading']);
    assert.deepEqual(doc.lists, [['ODT first', true], ['ODT second', true]]);
    assert.deepEqual(doc.tables, [[2, 2]]);
    assert.match(doc.notice, /^OpenDocument text ·/);
    await closeWindow(page);
  });

  await check('a template opens as a new untitled document and never saves over the template', async () => {
    const page = await openInWindow(main, files.dotx);
    const doc = await page.evaluate(summary);
    assert.equal(doc.title, 'Untitled document');
    assert.ok(doc.texts.includes('TEMPLATE_BODY'));
    assert.match(doc.notice, /New document from the template “Invoice template\.dotx”/);
    const target = path.join(fixtures, 'From template.docx');
    await answer({ save: target });
    await page.evaluate(`document.getElementById('save-button').click(), true`);
    await page.until(`document.getElementById('save-state').textContent === 'Saved locally'`, 'saved');
    const saveDialog = (await dialogLog()).filter((item) => item.name === 'save').at(-1);
    assert.equal(path.basename(saveDialog.options.defaultPath), 'Untitled document.docx');
    assert.ok((await readFile(files.dotx)).equals(originals.get(files.dotx)), 'the template is unchanged');
    const zip = await JSZip.loadAsync(await readFile(target));
    assert.match(await zip.file('[Content_Types].xml').async('string'), /wordprocessingml\.document\.main\+xml/);
    await answer({});
    await closeWindow(page);
  });

  await check('a macro-enabled document opens without its macros and saves as a .docx beside it', async () => {
    const page = await openInWindow(main, files.docm);
    const doc = await page.evaluate(summary);
    assert.equal(doc.title, 'Macro report');
    assert.match(doc.notice, /Macros aren’t run or kept/);
    await page.evaluate(`document.getElementById('save-button').click(), true`);
    await page.until(`!document.getElementById('choice-modal').hidden`, 'the save question');
    await page.evaluate(`document.querySelector('#choice-actions [data-choice="save"]').click(), true`);
    await page.until(`document.getElementById('save-state').textContent === 'Saved locally'`, 'saved');
    const zip = await JSZip.loadAsync(await readFile(path.join(fixtures, 'Macro report.docx')));
    assert.equal(zip.file('word/vbaProject.bin'), null, 'no macros in the saved file');
    assert.ok((await readFile(files.docm)).equals(originals.get(files.docm)), 'the .docm is unchanged');
    await closeWindow(page);
  });

  // ---- pasting --------------------------------------------------------------------------

  const paste = await openInWindow(main, files.blank);
  // Each paste check starts from the document as it opened.
  await paste.evaluate(`window.__smokeInitial = window.__cw.handle.getDocument(), true`);
  const reset = () => paste.evaluate(`window.__cw.handle.setDocument(window.__smokeInitial), true`);
  const input = `document.querySelector('#editor [contenteditable="true"][role="textbox"]')`;
  const caret = (blockIndex, offset) => `(() => { const block = ${H}.getDocument().blocks.filter((item) => item.kind === 'paragraph')[${blockIndex}]; ${H}.setSelection({ anchor: { blockId: block.id, offset: ${offset} }, focus: { blockId: block.id, offset: ${offset} } }); ${H}.focus(); return true; })()`;
  const select = (blockIndex, start, end) => `(() => { const block = ${H}.getDocument().blocks.filter((item) => item.kind === 'paragraph')[${blockIndex}]; ${H}.setSelection({ anchor: { blockId: block.id, offset: ${start} }, focus: { blockId: block.id, offset: ${end} } }); ${H}.focus(); return true; })()`;
  const firePaste = (data) => `(() => {
    const transfer = new DataTransfer();
    const data = ${JSON.stringify(data)};
    for (const [type, value] of Object.entries(data.strings || {})) transfer.setData(type, value);
    for (const file of data.files || []) transfer.items.add(new File([Uint8Array.from(atob(file.base64), (c) => c.charCodeAt(0))], file.name, { type: file.type }));
    const event = new ClipboardEvent('paste', { clipboardData: transfer, bubbles: true, cancelable: true });
    ${input}.dispatchEvent(event);
    return event.defaultPrevented;
  })()`;
  const revision = `${H}.getModelRevision()`;
  const HTML = '<html><body><!--StartFragment--><h1>Pasted heading</h1><ul><li>Bullet A</li><li>Bullet B</li></ul><ol><li>Number one</li><li>Number two</li></ol><table><tr><td>T1</td><td>T2</td></tr><tr><td>T3</td><td>T4</td></tr></table><p><img src="data:image/png;base64,' + PNG_4x2.toString('base64') + '" width="40" height="20"></p><!--EndFragment--></body></html>';

  await check('pasting HTML keeps the heading, real lists, the table and the picture as one undo step', async () => {
    await reset();
    await paste.evaluate(caret(1, 6));
    const before = await paste.evaluate(revision);
    assert.equal(await paste.evaluate(firePaste({ strings: { 'text/html': HTML, 'text/plain': 'Pasted heading\nBullet A' } })), true, 'the paste was handled');
    const doc = await paste.evaluate(summary);
    assert.deepEqual(doc.headings, ['Pasted heading']);
    assert.deepEqual(doc.lists, [['Bullet A', true], ['Bullet B', true], ['Number one', true], ['Number two', true]]);
    assert.deepEqual(doc.tables, [[2, 2]]);
    assert.equal(doc.images, 1);
    assert.equal(await paste.evaluate(revision), before + 1, 'one transaction');
    const listFormats = await paste.evaluate(`(() => { const doc = ${H}.getDocument(); return doc.blocks.filter((b) => b.kind === 'paragraph' && b.style.list).map((b) => doc.lists[b.style.list.listId].levels[0].format); })()`);
    assert.deepEqual(listFormats, ['bullet', 'bullet', 'decimal', 'decimal']);
    // The exported Word file has real numbering, the table, the heading style and the picture.
    const bytes = await paste.evaluate(`${H}.exportDocx().then((blob) => blob.arrayBuffer()).then((buffer) => Array.from(new Uint8Array(buffer)))`);
    const zip = await JSZip.loadAsync(Buffer.from(bytes));
    const xml = await zip.file('word/document.xml').async('string');
    assert.ok((xml.match(/<w:numPr>/g) || []).length >= 4, 'list paragraphs carry w:numPr');
    assert.equal((xml.match(/<w:tbl>/g) || []).length, 1);
    assert.equal((xml.match(/<w:tc>/g) || []).length, 4);
    assert.match(xml, /<w:pStyle w:val="Heading1"/);
    assert.ok(Object.keys(zip.files).some((name) => name.startsWith('word/media/')), 'the picture is in the package');
    // The paste options chip offers the three choices.
    assert.equal(await paste.evaluate(`!document.getElementById('paste-options').hidden`), true, 'the paste options chip is shown');
    // One Ctrl+Z removes the whole paste.
    await paste.evaluate(`${H}.undo(), true`);
    const undone = await paste.evaluate(summary);
    assert.deepEqual(undone.texts, ['Hello world', 'Second paragraph']);
    assert.equal(await paste.evaluate(`!document.getElementById('paste-options').hidden`), false, 'the chip goes away with the paste');
  });

  await check('paste options: Keep text only and Merge formatting re-paste the same content', async () => {
    await reset();
    await paste.evaluate(caret(0, 5));
    await paste.evaluate(firePaste({ strings: { 'text/html': '<p><span style="font-family:Georgia;font-size:24px;color:#c00000"><b>Big red</b> words</span></p>', 'text/plain': 'Big red words' } }));
    const pasted = await paste.evaluate(`${H}.getDocument().blocks[0].runs.map((run) => [run.text, run.style.fontFamily, run.style.bold === true])`);
    assert.deepEqual(pasted.find((run) => run[0] === 'Big red'), ['Big red', 'Georgia', true]);
    await paste.evaluate(`document.getElementById('paste-options-button').click(), true`);
    assert.equal(await paste.evaluate(`!document.getElementById('paste-options-menu').hidden`), true);
    await paste.evaluate(`document.querySelector('[data-paste-mode="text"]').click(), true`);
    const plain = await paste.evaluate(`${H}.getDocument().blocks[0].runs.map((run) => [run.text, run.style.fontFamily, run.style.bold === true])`);
    assert.equal(plain.map((run) => run[0]).join(''), 'HelloBig red words world');
    assert.ok(plain.every((run) => run[1] !== 'Georgia' && run[2] === false), 'text only takes the formatting at the caret');
    await paste.evaluate(`document.getElementById('paste-options-button').click(), true`);
    await paste.evaluate(`document.querySelector('[data-paste-mode="merge"]').click(), true`);
    const merged = await paste.evaluate(`${H}.getDocument().blocks[0].runs.map((run) => [run.text, run.style.fontFamily, run.style.bold === true])`);
    assert.deepEqual(merged.find((run) => run[0] === 'Big red').slice(1), [plain[0][1], true], 'merge keeps bold and takes the font at the caret');
    await paste.evaluate(`${H}.undo(), true`);
    assert.equal(await paste.evaluate(`${H}.getDocument().blocks[0].runs.map((run) => run.text).join('')`), 'Hello world', 'one undo removes the re-pasted content');
  });

  await check('Ctrl+Shift+V pastes text only, formatted like the caret, in one undo step', async () => {
    await reset();
    await paste.evaluate(caret(1, 0));
    const before = await paste.evaluate(revision);
    await paste.evaluate(`${input}.dispatchEvent(new KeyboardEvent('keydown', { key: 'V', code: 'KeyV', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true })), true`);
    await paste.evaluate(firePaste({ strings: { 'text/plain': 'Plain one\nPlain two\n', 'text/html': '<p><b>Plain one</b></p><p>Plain two</p>' } }));
    const doc = await paste.evaluate(summary);
    // Copied whole lines: the final line break keeps "Second paragraph" on its own line.
    assert.deepEqual(doc.texts.slice(0, 4), ['Hello world', 'Plain one', 'Plain two', 'Second paragraph']);
    assert.equal(await paste.evaluate(`${H}.getDocument().blocks[1].runs.some((run) => run.style.bold)`), false, 'no bold from the source');
    assert.equal(await paste.evaluate(revision), before + 1, 'one transaction');
    assert.equal(await paste.evaluate(`document.getElementById('paste-options').hidden`), true, 'no chip after text-only paste');
    await paste.evaluate(`${H}.undo(), true`);
    assert.deepEqual((await paste.evaluate(summary)).texts, ['Hello world', 'Second paragraph']);
  });

  await check('a web address pasted over selected text links the text', async () => {
    await reset();
    await paste.evaluate(select(0, 0, 5));
    await paste.evaluate(firePaste({ strings: { 'text/plain': 'https://example.com/page' } }));
    const runs = await paste.evaluate(`${H}.getDocument().blocks[0].runs.map((run) => [run.text, run.style.link || null])`);
    assert.deepEqual(runs[0], ['Hello', 'https://example.com/page']);
    assert.equal(runs.map((run) => run[0]).join(''), 'Hello world');
    await paste.evaluate(`${H}.undo(), true`);
  });

  await check('pasting a screenshot inserts the picture', async () => {
    await reset();
    await paste.evaluate(caret(1, 6));
    assert.equal(await paste.evaluate(firePaste({ files: [{ name: 'image.png', type: 'image/png', base64: PNG_4x2.toString('base64') }] })), true);
    await paste.until(`${H}.getDocument().blocks.some((block) => block.kind === 'image')`, 'the picture is inserted');
    await paste.evaluate(`${H}.undo(), true`);
  });

  await check('the ribbon Paste uses the same path (clipboard read stubbed, no system clipboard)', async () => {
    await reset();
    await paste.evaluate(caret(0, 11));
    await paste.evaluate(`(() => { navigator.clipboard.read = async () => [new ClipboardItem({ 'text/html': new Blob(['<ul><li>Ribbon item</li><li>Second item</li></ul>'], { type: 'text/html' }), 'text/plain': new Blob(['Ribbon item'], { type: 'text/plain' }) })]; return true; })()`);
    await paste.evaluate(`document.querySelector('[data-ribbon-item="home.clipboard.paste"]').click(), true`);
    await paste.until(`${H}.getDocument().blocks.some((block) => block.kind === 'paragraph' && block.style.list && block.runs.some((run) => run.text === 'Ribbon item'))`, 'the ribbon paste inserted a real list');
    await paste.evaluate(`${H}.undo(), true`);
  });

  await check('dropping a picture inserts it with no error', async () => {
    await reset();
    await paste.evaluate(caret(1, 0));
    await pause(200);
    // Drop where the caret is drawn: a point on the page.
    const point = await paste.evaluate(`(() => { const rect = ${input}.getBoundingClientRect(); return { x: Math.round(rect.left + 2), y: Math.round(rect.top + rect.height / 2) }; })()`);
    await paste.evaluate(caret(0, 0));
    await paste.evaluate(`(() => {
      const transfer = new DataTransfer();
      transfer.items.add(new File([Uint8Array.from(atob(${JSON.stringify(PNG_4x2.toString('base64'))}), (c) => c.charCodeAt(0))], 'photo.png', { type: 'image/png' }));
      const target = document.elementFromPoint(${point.x}, ${point.y});
      for (const type of ['dragenter', 'dragover', 'drop']) target.dispatchEvent(new DragEvent(type, { dataTransfer: transfer, bubbles: true, cancelable: true, clientX: ${point.x}, clientY: ${point.y} }));
      return true;
    })()`);
    await paste.until(`${H}.getDocument().blocks.some((block) => block.kind === 'image')`, 'the dropped picture is inserted');
    const order = await paste.evaluate(`${H}.getDocument().blocks.map((block) => block.kind === 'image' ? '[image]' : (block.runs || []).map((run) => run.text).join(''))`);
    assert.ok(order.indexOf('[image]') > order.indexOf('Hello world'), `the picture lands at the drop point, not at the caret (${JSON.stringify(order)})`);
    const doc = await paste.evaluate(summary);
    assert.notEqual(doc.toastTone, 'error', `no error toast (${doc.toast})`);
    assert.equal(await paste.evaluate(`document.getElementById('drop-overlay').hidden`), true);
  });
  await closeWindow(paste);
} catch (error) {
  results.push({ name: 'harness', ok: false, error: error.stack || error.message });
  console.log(`FAIL harness\n     ${error.stack || error.message}`);
} finally {
  clearTimeout(hardStop);
  try { main?.close(); } catch {}
  killAll();
}

const failed = results.filter((result) => !result.ok);
await writeFile(path.join(output, 'results.json'), JSON.stringify(results, null, 2));
console.log(`\n${results.length - failed.length}/${results.length} checks passed.`);
process.exit(failed.length ? 1 : 0);
