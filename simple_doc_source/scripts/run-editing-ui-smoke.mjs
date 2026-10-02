// Everyday-editing UI smoke test for Simple Docs (Electron + CDP, built renderer).
// Covers: focus after Blank document, Open, Edit document and dialogs; Tab;
// Word/Docs shortcuts on US and Russian layouts (physical key codes); in-app
// prompts for links, bookmarks and comments; Undo/Redo buttons; command search;
// Page view (find bar, Ctrl+F, zoom). Run `npm run build:web` first, or point
// SIMPLE_DOCS_DIST at another build. Windows open off-screen and every process
// is killed at the end; nothing is printed or sent anywhere.
import { spawn, execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { mkdir, writeFile, rm } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';
import JSZip from 'jszip';
import { PDFDocument, StandardFonts } from 'pdf-lib';

const require = createRequire(import.meta.url);
const root = fileURLToPath(new URL('..', import.meta.url));
const output = path.join(root, 'tmp', 'editing-ui-smoke');
const dist = process.env.SIMPLE_DOCS_DIST ? path.resolve(process.env.SIMPLE_DOCS_DIST) : null;
await rm(output, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
await mkdir(path.join(output, 'profile'), { recursive: true });

async function docx(file, paragraphs) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${paragraphs.map((text) => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`).join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`);
  await writeFile(file, await zip.generateAsync({ type: 'nodebuffer' }));
}
const plainDocx = path.join(output, 'opened-plain.docx');
const pageViewDocx = path.join(output, 'page-view-sample.docx');
await docx(plainDocx, ['Opened paragraph one', 'Opened paragraph two']);
await docx(pageViewDocx, ['Original page text', 'Second original paragraph']);
{
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([612, 792]);
  page.drawText('Original page text (Page view)', { x: 72, y: 700, size: 18, font: await pdf.embedFont(StandardFonts.Helvetica) });
  await writeFile(path.join(output, 'page-view.pdf'), await pdf.save());
}

const hostPath = path.join(output, 'host.cjs');
await writeFile(hostPath, `const { app, dialog, shell, ipcMain, BrowserWindow } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
app.setPath('userData', path.join(__dirname, 'profile'));
const answers = () => { try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'answers.json'), 'utf8')); } catch { return {}; } };
dialog.showOpenDialog = async () => { const file = answers().open; return file ? { canceled: false, filePaths: [file] } : { canceled: true, filePaths: [] }; };
dialog.showSaveDialog = async () => { fs.appendFileSync(path.join(__dirname, 'save-dialogs.log'), 'save\\n'); return { canceled: true }; };
dialog.showMessageBox = async () => ({ response: 0 });
shell.openExternal = async () => {};
// Page view needs a local Office engine; the test serves a fixed PDF instead.
const handle = ipcMain.handle.bind(ipcMain);
const withPageView = (payload) => { if (payload && /page-view/i.test(String(payload.name))) payload.originalLayout = { data: payload.data, extension: 'docx' }; return payload; };
ipcMain.handle = (channel, listener) => {
  if (channel === 'file:open-dialog' || channel === 'file:open-path') return handle(channel, async (...args) => withPageView(await listener(...args)));
  if (channel === 'document:original-pdf') return handle(channel, async () => new Uint8Array(fs.readFileSync(path.join(__dirname, 'page-view.pdf'))));
  return handle(channel, listener);
};
const show = BrowserWindow.prototype.show;
BrowserWindow.prototype.show = function () { if (process.env.SIMPLE_SMOKE_SHOW) return show.call(this); try { this.setPosition(-3200, 40); this.showInactive(); } catch {} };
app.on('browser-window-created', (_event, window) => { try { window.webContents.setBackgroundThrottling(false); } catch {} });
require(${JSON.stringify(path.join(root, 'electron', 'main.cjs'))});
`);
const answer = (value) => writeFile(path.join(output, 'answers.json'), JSON.stringify(value));

const port = 24100 + Math.floor(Math.random() * 700);
const env = { ...process.env };
delete env.ELECTRON_RUN_AS_NODE;
if (dist) env.VITE_DEV_SERVER_URL = pathToFileURL(path.join(dist, 'index.html')).href;
const child = spawn(require('electron'), [`--remote-debugging-port=${port}`, `--user-data-dir=${path.join(output, 'profile')}`, hostPath], { cwd: root, env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
let stderr = '';
child.stderr.on('data', (data) => { stderr += String(data); });
const killAll = () => { try { execFileSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { try { child.kill(); } catch {} } };
const hardStop = setTimeout(() => { console.error('Editing UI smoke test timed out.'); killAll(); process.exit(124); }, Number(process.env.SIMPLE_SMOKE_TIMEOUT_MS || 300_000));

const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function connect(exclude = []) {
  let target;
  const end = Date.now() + 45_000;
  while (!target && Date.now() < end) {
    try { target = (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find((candidate) => candidate.type === 'page' && !exclude.includes(candidate.id)); } catch {}
    if (!target) await pause(150);
  }
  assert.ok(target, `A renderer window opens. ${stderr.slice(-1500)}`);
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
  const send = (method, params = {}) => new Promise((resolve, reject) => { const request = ++id; pending.set(request, { resolve, reject }); socket.send(JSON.stringify({ id: request, method, params })); });
  const evaluate = async (expression) => {
    const result = await send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true, userGesture: true });
    if (result.exceptionDetails) throw new Error(`${result.exceptionDetails.text} ${result.exceptionDetails.exception?.description ?? ''}`);
    return result.result.value;
  };
  const until = async (expression, label, timeout = 20_000) => {
    const stop = Date.now() + timeout;
    let last;
    while (Date.now() < stop) {
      try { last = await evaluate(expression); if (last) return last; } catch (error) { last = error.message; }
      await pause(80);
    }
    throw new Error(`Timed out: ${label} (last: ${JSON.stringify(last)})`);
  };
  // `key` is what the layout types (a Cyrillic letter on a Russian keyboard), `code` the physical key.
  const MODIFIERS = { alt: 1, ctrl: 2, shift: 8 };
  const key = async (keyValue, code, vk, modifiers = [], text) => {
    const mask = modifiers.reduce((sum, name) => sum + MODIFIERS[name], 0);
    const down = { type: text ? 'keyDown' : 'rawKeyDown', key: keyValue, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mask };
    if (text) Object.assign(down, { text, unmodifiedText: text });
    await send('Input.dispatchKeyEvent', down);
    await send('Input.dispatchKeyEvent', { type: 'keyUp', key: keyValue, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mask });
    await pause(90);
  };
  const type = async (text) => { await send('Input.insertText', { text }); await pause(90); };
  const mouse = async (x, y, button = 'left') => { for (const kind of ['mousePressed', 'mouseReleased']) await send('Input.dispatchMouseEvent', { type: kind, button, clickCount: 1, x, y }); await pause(120); };
  await send('Page.enable');
  return { target, socket, evaluate, until, key, type, mouse };
}

// Page-side helpers (WordCanvas exposes its handle as window.__cw.handle).
const H = 'window.__cw.handle';
const PROXY = `document.querySelector('#editor [contenteditable="true"][role="textbox"]')`;
const focused = `document.activeElement === ${PROXY}`;
const texts = `${H}.getDocument().blocks.filter(b=>b.kind==='paragraph').map(b=>b.runs.map(r=>r.text).join(''))`;
const firstBlock = `${H}.getDocument().blocks.find(b=>b.kind==='paragraph')`;
const select = (start, end) => `(()=>{const b=${firstBlock};${H}.setSelection({anchor:{blockId:b.id,offset:${start}},focus:{blockId:b.id,offset:${end}}});${H}.focus();return true})()`;
const dialogOpen = `document.querySelector('.form-dialog-backdrop')!==null`;

const results = [];
async function check(name, run) {
  const started = Date.now();
  try {
    await run();
    results.push({ name, ok: true, ms: Date.now() - started });
    console.log(`ok   ${name}`);
  } catch (error) {
    results.push({ name, ok: false, error: String(error?.message ?? error) });
    console.log(`FAIL ${name}\n     ${String(error?.message ?? error).split('\n').join('\n     ')}`);
  }
}

try {
  // ---- Window 1: a blank document --------------------------------------------
  const w1 = await connect();
  await w1.until(`document.querySelector('#blank-document') && document.querySelector('#loading-overlay').hidden`, 'welcome screen', 60_000);

  await check('Blank document puts the caret in the document, so typing lands there', async () => {
    await w1.evaluate(`document.querySelector('#blank-document').click()`);
    await w1.until(focused, 'editor focus after Blank document', 5_000);
    await w1.type('Hello brave world');
    assert.deepEqual(await w1.evaluate(texts), ['Hello brave world']);
    assert.equal(await w1.evaluate(`document.querySelector('#save-state').textContent`), 'Unsaved');
  });

  await check('Undo and Redo buttons follow the history and keep the caret in the document', async () => {
    await w1.until(`!document.querySelector('#undo-button').disabled`, 'Undo enabled');
    assert.equal(await w1.evaluate(`document.querySelector('#redo-button').disabled`), true);
    await w1.evaluate(`document.querySelector('#undo-button').click()`);
    await w1.until(`JSON.stringify(${texts})==='[""]'`, 'undo by button');
    await w1.until(`!document.querySelector('#redo-button').disabled`, 'Redo enabled');
    await w1.evaluate(`document.querySelector('#redo-button').click()`);
    await w1.until(`${texts}[0]==='Hello brave world'`, 'redo by button');
    assert.equal(await w1.evaluate(focused), true);
  });

  await check('Tab inserts a tab and Shift+Tab keeps focus in the document', async () => {
    await w1.evaluate(select(17, 17));
    await w1.key('Tab', 'Tab', 9);
    assert.equal(await w1.evaluate(`${texts}[0]`), 'Hello brave world\t');
    assert.equal(await w1.evaluate(focused), true, 'focus stays in the document');
    await w1.key('Tab', 'Tab', 9, ['shift']);
    assert.equal(await w1.evaluate(focused), true);
    await w1.key('Backspace', 'Backspace', 8);
    assert.equal(await w1.evaluate(`${texts}[0]`), 'Hello brave world');
  });

  const align = `${firstBlock}.style.align`;
  await check('Ctrl+E/R/J/L align the paragraph (US layout)', async () => {
    for (const [letter, vk, expected] of [['e', 69, 'center'], ['r', 82, 'right'], ['j', 74, 'justify'], ['l', 76, 'left']]) {
      await w1.key(letter, `Key${letter.toUpperCase()}`, vk, ['ctrl']);
      await w1.until(`${align}==='${expected}'`, `Ctrl+${letter.toUpperCase()} → ${expected}`, 3_000);
    }
  });

  await check('Alignment, undo and bold shortcuts follow physical keys on a Russian layout', async () => {
    await w1.key('у', 'KeyE', 69, ['ctrl']);
    await w1.until(`${align}==='center'`, 'Ctrl+У → center', 3_000);
    await w1.key('д', 'KeyL', 76, ['ctrl']);
    await w1.until(`${align}==='left'`, 'Ctrl+Д → left', 3_000);
    await w1.key('я', 'KeyZ', 90, ['ctrl']);
    await w1.until(`${align}==='center'`, 'Ctrl+Я undoes', 3_000);
    await w1.key('н', 'KeyY', 89, ['ctrl']);
    await w1.until(`${align}==='left'`, 'Ctrl+Н redoes', 3_000);
    await w1.evaluate(select(6, 11));
    await w1.key('и', 'KeyB', 66, ['ctrl']);
    await w1.until(`${firstBlock}.runs.some(r=>r.text==='brave'&&r.style.bold)`, 'Ctrl+И → bold', 3_000);
  });

  await check('Ctrl+Space clears character formatting; Ctrl+] / Ctrl+[ change the size', async () => {
    await w1.evaluate(select(6, 11));
    await w1.key(' ', 'Space', 32, ['ctrl']);
    await w1.until(`!${firstBlock}.runs.some(r=>r.text.includes('brave')&&r.style.bold)`, 'bold cleared', 3_000);
    const size = () => w1.evaluate(`${firstBlock}.runs.find(r=>r.text.includes('brave')).style.fontSizePx`);
    const before = await size();
    await w1.evaluate(select(6, 11));
    await w1.key(']', 'BracketRight', 221, ['ctrl']);
    await w1.until(`${firstBlock}.runs.find(r=>r.text.includes('brave')).style.fontSizePx>${before}`, 'Ctrl+] grows', 3_000);
    await w1.key('х', 'BracketLeft', 219, ['ctrl']);
    await w1.until(`Math.abs(${firstBlock}.runs.find(r=>r.text.includes('brave')).style.fontSizePx-${before})<0.01`, 'Ctrl+[ (Russian х) shrinks', 3_000);
  });

  await check('Ctrl+= subscript and Ctrl+Shift+= superscript', async () => {
    await w1.evaluate(select(6, 11));
    await w1.key('=', 'Equal', 187, ['ctrl']);
    await w1.until(`${firstBlock}.runs.find(r=>r.text==='brave')?.style.verticalAlign==='sub'`, 'subscript', 3_000);
    await w1.key('+', 'Equal', 187, ['ctrl', 'shift']);
    await w1.until(`${firstBlock}.runs.find(r=>r.text==='brave')?.style.verticalAlign==='super'`, 'superscript', 3_000);
    await w1.key('+', 'Equal', 187, ['ctrl', 'shift']);
    await w1.until(`!${firstBlock}.runs.find(r=>r.text==='brave')?.style.verticalAlign`, 'superscript off', 3_000);
  });

  const lineHeight = `${firstBlock}.style.lineHeight`;
  await check('Ctrl+2 / Ctrl+5 / Ctrl+1 set double, 1.5 and single line spacing', async () => {
    await w1.evaluate(select(0, 0));
    for (const [digit, expected] of [['2', 2], ['5', 1.5], ['1', 1]]) {
      await w1.key(digit, `Digit${digit}`, 48 + Number(digit), ['ctrl']);
      await w1.until(`Math.abs(${lineHeight}-${expected})<0.001`, `Ctrl+${digit}`, 3_000);
    }
    assert.equal(await w1.evaluate(`document.querySelectorAll('.cw-pop').length`), 0, 'no menu left open');
  });

  await check('Ctrl+Alt+1/3 apply headings (Heading 3 included) and Ctrl+Shift+N returns to Normal', async () => {
    await w1.key('1', 'Digit1', 49, ['ctrl', 'alt']);
    await w1.until(`${firstBlock}.style.namedStyle==='Heading1'`, 'Heading 1', 3_000);
    await w1.key('3', 'Digit3', 51, ['ctrl', 'alt']);
    await w1.until(`${firstBlock}.style.namedStyle==='Heading3'`, 'Heading 3', 3_000);
    await w1.key('N', 'KeyN', 78, ['ctrl', 'shift']);
    await w1.until(`${firstBlock}.style.namedStyle==='Normal'`, 'Normal', 3_000);
  });

  await check('Shift+F3 cycles UPPERCASE, Capitalize Each Word and lowercase', async () => {
    await w1.evaluate(select(8, 8));
    for (const expected of ['Hello BRAVE world', 'Hello Brave world', 'Hello brave world']) {
      await w1.key('F3', 'F3', 114, ['shift']);
      await w1.until(`${texts}[0]===${JSON.stringify(expected)}`, expected, 3_000);
    }
  });

  await check('Ctrl+Shift+8 shows formatting marks', async () => {
    const active = `document.querySelector('[data-ribbon-item="home.paragraph.show-hide-formatting-marks"]').classList.contains('active')`;
    const before = await w1.evaluate(active);
    await w1.key('*', 'Digit8', 56, ['ctrl', 'shift']);
    await w1.until(`${active}!==${before}`, 'marks toggled', 3_000);
    await w1.key('*', 'Digit8', 56, ['ctrl', 'shift']);
    await w1.until(`${active}===${before}`, 'marks toggled back', 3_000);
  });

  await check('Ctrl+K opens the in-app link dialog and links the selection', async () => {
    await w1.evaluate(select(12, 17));
    await w1.key('k', 'KeyK', 75, ['ctrl']);
    await w1.until(dialogOpen, 'link dialog');
    assert.equal(await w1.evaluate(`document.activeElement.closest('.form-dialog')!==null`), true, 'the address field has focus');
    await w1.type('example.com/docs');
    await w1.key('Enter', 'Enter', 13, [], '\r');
    await w1.until(`${firstBlock}.runs.find(r=>r.text==='world')?.style.link==='https://example.com/docs'`, 'link applied', 5_000);
    await w1.until(focused, 'focus back in the document', 3_000);
  });

  await check('The context menu Insert Hyperlink uses the in-app dialog (engine prompt bridge)', async () => {
    try {
      await w1.evaluate(select(0, 5));
      // A point inside the selected "Hello" (the line of the caret, found by hit-testing).
      const point = await w1.evaluate(`(()=>{const r=${PROXY}.getBoundingClientRect();const y=r.top+r.height/2;for(let x=Math.max(0,r.left-900);x<r.left+40;x+=3){const p=${H}.positionFromPoint(x,y);const hit=document.elementFromPoint(x,y);if(p&&p.offset>=1&&p.offset<=4&&hit?.tagName==='CANVAS'&&document.querySelector('#editor').contains(hit))return{x,y}}return null})()`);
      assert.ok(point, 'a point inside the selected word');
      await w1.mouse(point.x, point.y, 'right');
      await w1.until(`[...document.querySelectorAll('.cw-menu .cw-menu-item')].some(b=>b.textContent.includes('Insert Hyperlink'))||JSON.stringify([...document.querySelectorAll('.cw-menu')].map(m=>m.textContent.slice(0,200)))`, 'context menu with Insert Hyperlink');
      assert.equal(await w1.evaluate(`[...document.querySelectorAll('.cw-menu .cw-menu-item')].some(b=>b.textContent.includes('Insert Hyperlink'))`), true, `menu: ${await w1.evaluate(`[...document.querySelectorAll('.cw-menu')].map(m=>m.textContent.slice(0,300)).join(' / ')`)}`);
      const item = await w1.evaluate(`(()=>{const r=[...document.querySelectorAll('.cw-menu .cw-menu-item')].find(b=>b.textContent.includes('Insert Hyperlink')).getBoundingClientRect();return{x:r.left+r.width/2,y:r.top+r.height/2}})()`);
      await w1.mouse(item.x, item.y);
      await w1.until(dialogOpen, 'hyperlink dialog from the context menu', 5_000);
      await w1.type('name@example.com');
      await w1.key('Enter', 'Enter', 13, [], '\r');
      await w1.until(`${firstBlock}.runs.some(r=>r.style.link==='mailto:name@example.com')`, 'mailto link', 5_000);
      await w1.until(focused, 'focus back in the document', 3_000);
    } finally {
      // Never leave an engine menu open (it captures Escape for the whole window).
      if (await w1.evaluate(`document.querySelector('.cw-menu')!==null`)) await w1.key('Escape', 'Escape', 27);
    }
  });

  await check('Ctrl+Alt+M asks for the author name once and adds the comment with it', async () => {
    await w1.evaluate(select(6, 11));
    await w1.key('m', 'KeyM', 77, ['ctrl', 'alt']);
    await w1.until(dialogOpen, 'comment dialog');
    const fields = await w1.evaluate(`[...document.querySelectorAll('.form-dialog .form-field-label')].map(n=>n.textContent)`);
    assert.deepEqual(fields, ['Your name', 'Comment']);
    await w1.evaluate(`(()=>{const name=document.querySelector('.form-dialog input');name.value='Test Author';name.dispatchEvent(new Event('input',{bubbles:true}));document.querySelector('.form-dialog textarea').focus()})()`);
    await w1.type('Please check this word');
    await w1.key('Enter', 'Enter', 13, ['ctrl']);
    await w1.until(`${H}.getReview().threads.length===1`, 'comment thread', 5_000);
    const thread = await w1.evaluate(`${H}.getReview().threads[0].comments[0]`);
    assert.deepEqual([thread.author.firstName, thread.author.lastName, thread.body[0].text], ['Test', 'Author', 'Please check this word']);
    await w1.until(focused, 'focus back after the comment', 3_000);
    await w1.key('m', 'KeyM', 77, ['ctrl', 'alt']);
    await w1.until(dialogOpen, 'second comment dialog');
    assert.deepEqual(await w1.evaluate(`[...document.querySelectorAll('.form-dialog .form-field-label')].map(n=>n.textContent)`), ['Comment'], 'the name is asked only once');
    await w1.key('Escape', 'Escape', 27);
    await w1.until(`!(${dialogOpen})`, 'Escape closes the dialog');
    await w1.until(focused, 'focus back after Escape', 3_000);
  });

  await check('Bookmarks: the add prompt validates Word names in the in-app dialog', async () => {
    await w1.evaluate(select(0, 5));
    await w1.evaluate(`document.querySelector('[data-ribbon-item="view.show.bookmarks-list-go-to-add-rename-delete"]').click()`);
    await w1.until(`[...document.querySelectorAll('.cw-float-drawer button')].some(b=>b.title.startsWith('Add a bookmark'))`, 'bookmarks pane');
    await w1.evaluate(`[...document.querySelectorAll('.cw-float-drawer button')].find(b=>b.title.startsWith('Add a bookmark')).click()`);
    await w1.until(dialogOpen, 'bookmark dialog');
    await w1.type('Chapter one');
    await w1.key('Enter', 'Enter', 13, [], '\r');
    await w1.until(`!document.querySelector('.form-dialog-error').hidden`, 'name error shown');
    await w1.evaluate(`(()=>{const i=document.querySelector('.form-dialog input');i.value='Chapter_one';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    await w1.key('Enter', 'Enter', 13, [], '\r');
    await w1.until(`Object.keys(${H}.getDocument().bookmarks??{}).includes('Chapter_one')`, 'bookmark added', 5_000);
    await w1.evaluate(`[...document.querySelectorAll('.cw-float-drawer button')].find(b=>b.title==='Close')?.click()`);
  });

  await check('Insert > Drop-down list asks for its choices in the in-app dialog', async () => {
    await w1.evaluate(`(()=>{const b=${firstBlock};const n=b.runs.map(r=>r.text).join('').length;${H}.setSelection({anchor:{blockId:b.id,offset:n},focus:{blockId:b.id,offset:n}});${H}.focus()})()`);
    await w1.evaluate(`document.querySelector('[data-ribbon-item="insert.controls.drop-down-list-content-control"]').click()`);
    await w1.until(dialogOpen, 'drop-down dialog');
    await w1.evaluate(`(()=>{const i=document.querySelector('.form-dialog input');i.value='Draft, Final';i.dispatchEvent(new Event('input',{bubbles:true}))})()`);
    await w1.key('Enter', 'Enter', 13, [], '\r');
    await w1.until(`Object.values(${H}.getDocument().sdts??{}).some(s=>s.type==='dropDown'&&s.listItems?.map(i=>i.value).join('|')==='Draft|Final')`, 'drop-down control with two choices', 5_000);
    await w1.until(focused, 'focus back in the document', 3_000);
    await w1.key('z', 'KeyZ', 90, ['ctrl']);
  });

  await check('Ctrl+H opens Replace with the replace box focused; Ctrl+F finds', async () => {
    await w1.evaluate(select(0, 0));
    await w1.key('h', 'KeyH', 72, ['ctrl']);
    await w1.until(`document.activeElement?.placeholder==='Replace'`, 'replace box focused', 3_000);
    await w1.key('Escape', 'Escape', 27);
    await w1.evaluate(select(6, 11));
    await w1.key('а', 'KeyF', 70, ['ctrl']);
    await w1.until(`document.activeElement?.placeholder==='Find'&&document.activeElement.value==='brave'`, 'Russian Ctrl+А opens Find with the selection', 3_000);
    await w1.key('Escape', 'Escape', 27);
  });

  await check('Alt+Q command search finds and runs a ribbon command, then returns to the document', async () => {
    await w1.evaluate(select(0, 0));
    await w1.key('q', 'KeyQ', 81, ['alt']);
    await w1.until(`document.activeElement?.id==='command-search-input'`, 'search focused', 3_000);
    await w1.type('justify');
    await w1.until(`document.querySelector('#command-search-list [aria-selected="true"] strong')?.textContent==='Justify'`, 'Justify listed first', 3_000);
    await w1.key('Enter', 'Enter', 13, [], '\r');
    await w1.until(`${align}==='justify'`, 'Justify ran', 3_000);
    await w1.until(focused, 'focus back in the document', 3_000);
    await w1.key('й', 'KeyQ', 81, ['alt']);
    await w1.until(`document.activeElement?.id==='command-search-input'`, 'Alt+Й (Russian) opens search', 3_000);
    await w1.until(`document.querySelector('#command-search-list [data-recent="true"] strong')?.textContent==='Justify'`, 'recent command listed first', 3_000);
    await w1.key('Escape', 'Escape', 27);
    await w1.until(focused, 'Escape returns to the document', 3_000);
  });

  await check('Ctrl+Backspace deletes the previous word as one step (DOC-006)', async () => {
    const before = await w1.evaluate(`${texts}[0]`);
    await w1.evaluate(`(()=>{const b=${firstBlock};const n=b.runs.map(r=>r.text).join('').length;${H}.setSelection({anchor:{blockId:b.id,offset:n},focus:{blockId:b.id,offset:n}});${H}.focus()})()`);
    await w1.key('Backspace', 'Backspace', 8, ['ctrl']);
    await w1.until(`${texts}[0]===${JSON.stringify(before.replace(/\S+$/, ''))}`, 'word deleted', 3_000);
    await w1.key('z', 'KeyZ', 90, ['ctrl']);
    await w1.until(`${texts}[0]===${JSON.stringify(before)}`, 'one undo restores it', 3_000);
    await w1.evaluate(select(0, 0));
    await w1.key('Delete', 'Delete', 46, ['ctrl']);
    await w1.until(`${texts}[0]===${JSON.stringify(before.replace(/^\S+/, ''))}`, 'Ctrl+Delete deletes the next word', 3_000);
    await w1.key('z', 'KeyZ', 90, ['ctrl']);
    await w1.until(`${texts}[0]===${JSON.stringify(before)}`, 'one undo restores it', 3_000);
  });

  await check('Cancelling the Envelope dialog returns to the document', async () => {
    await w1.evaluate(`document.querySelector('[data-ribbon-item="simple.envelopes.create"]').click()`);
    await w1.until(`!document.querySelector('#envelope-modal').hidden`, 'envelope dialog');
    await w1.key('Escape', 'Escape', 27);
    await w1.until(`document.querySelector('#envelope-modal').hidden`, 'envelope dialog closed');
    await w1.until(focused, 'focus back in the document', 3_000);
  });

  await check('Ctrl+Alt+F inserts a footnote', async () => {
    await w1.evaluate(select(5, 5));
    const notes = await w1.evaluate(`Object.keys(${H}.getDocument().footnotes??{}).length`);
    await w1.key('f', 'KeyF', 70, ['ctrl', 'alt']);
    await w1.until(`Object.keys(${H}.getDocument().footnotes??{}).length===${notes + 1}`, 'footnote added', 3_000);
    await w1.key('Escape', 'Escape', 27);
  });

  await check('F12 starts Save as; Ctrl+W asks to save and Cancel returns to the document', async () => {
    await w1.evaluate(select(0, 0));
    await w1.key('F12', 'F12', 123);
    // The document has a comment, so Save as first says comments can't be saved yet.
    await w1.until(`!document.querySelector('#choice-modal').hidden && /comment/i.test(document.querySelector('#choice-message').textContent)`, 'Save as review question', 5_000);
    await w1.evaluate(`document.querySelector('#choice-actions [data-choice="cancel"]').click()`);
    await w1.until(focused, 'focus back after the question', 3_000);
    await w1.key('w', 'KeyW', 87, ['ctrl']);
    await w1.until(`!document.querySelector('#close-modal').hidden`, 'close question', 5_000);
    await w1.evaluate(`document.querySelector('#cancel-close').click()`);
    await w1.until(focused, 'Cancel returns to the document', 3_000);
  });

  await check('Ribbon tooltips advertise the shortcuts', async () => {
    const titles = await w1.evaluate(`Object.fromEntries(['home.paragraph.center','home.font.grow-font','insert.links.insert-remove-hyperlink','home.editing.replace','insert.references.insert-footnote'].map(id=>[id,document.querySelector('[data-ribbon-item="'+id+'"]').title]))`);
    assert.match(titles['home.paragraph.center'], /\(Ctrl\+E\)$/);
    assert.match(titles['home.font.grow-font'], /\(Ctrl\+\]\)$/);
    assert.match(titles['insert.links.insert-remove-hyperlink'], /\(Ctrl\+K\)$/);
    assert.match(titles['home.editing.replace'], /\(Ctrl\+H\)$/);
    // The footnote tooltip stays verbatim (run-editor-stress.mjs finds the button by it); the key is announced instead.
    assert.equal(titles['insert.references.insert-footnote'], 'Insert footnote');
    assert.equal(await w1.evaluate(`document.querySelector('[data-ribbon-item="insert.references.insert-footnote"]').getAttribute('aria-keyshortcuts')`), 'Control+Alt+F');
    const missing = await w1.evaluate(`['home.font.grow-font','home.font.shrink-font','home.font.change-case','home.font.clear-all-formatting','home.font.superscript','home.font.subscript','home.paragraph.show-hide-formatting-marks','home.paragraph.align-left','home.paragraph.center','home.paragraph.align-right','home.paragraph.justify','home.paragraph.line-spacing','home.paragraph.decrease-indent','home.paragraph.increase-indent','home.styles.show-only-styles-in-use','home.editing.find-replace','home.editing.replace','home.editing.select-all','insert.links.insert-remove-hyperlink','insert.references.insert-footnote','insert.references.insert-endnote','home.font.bold','home.font.italic','home.font.underline'].filter(id=>!document.querySelector('[data-ribbon-item="'+id+'"]'))`);
    assert.deepEqual(missing, [], 'every ribbon control the shortcuts use exists');
  });

  // ---- Window 2: Page view (an Office rendering of the original) -------------
  await answer({ open: pageViewDocx });
  await w1.evaluate(`document.querySelector('#open-button').click()`);
  const w2 = await connect([w1.target.id]);
  await check('Page view hides the editor tools; Ctrl+F switches to Edit document without touching the model', async () => {
    await w2.until(`document.body.classList.contains('page-view') && !document.querySelector('#original-layout-pdf').hidden`, 'page view loaded', 45_000);
    assert.equal(await w2.evaluate(`document.querySelector('#editor').hidden`), true);
    const model = await w2.evaluate(`JSON.stringify(${texts})`);
    await w2.key('f', 'KeyF', 70, ['ctrl']);
    await w2.until(`!document.body.classList.contains('page-view') && document.activeElement?.placeholder==='Find'`, 'Edit document with Find open', 5_000);
    assert.equal(await w2.evaluate(`JSON.stringify(${texts})`), model, 'the model is unchanged');
    await w2.evaluate(`document.querySelector('#original-layout-button').click()`);
    await w2.until(`document.body.classList.contains('page-view')`, 'back to page view');
    const visibleTools = await w2.evaluate(`[...document.querySelectorAll('.cw-float-panel,.cw-fmtbar,.cw-ctxbar,.cw-linkbar,.cw-float-drawer')].filter(n=>getComputedStyle(n).display!=='none').length`);
    assert.equal(visibleTools, 0, 'no WordCanvas bar or pane over the page view');
    assert.equal(await w2.evaluate(`document.activeElement === document.body || !document.querySelector('#editor').contains(document.activeElement)`), true, 'the hidden editor has no focus');
  });

  await check('The Page view zoom choice reloads the PDF at that zoom', async () => {
    const before = await w2.evaluate(`(()=>{const f=document.querySelector('#original-layout-pdf');window.__smokeFrame=f;return f.src})()`);
    assert.match(before, /#view=Fit&toolbar=0$/);
    await w2.evaluate(`(()=>{const s=document.querySelector('#source-view-scale');s.value='150';s.dispatchEvent(new Event('change',{bubbles:true}))})()`);
    await w2.until(`document.querySelector('#original-layout-pdf').src.endsWith('#zoom=150&toolbar=0')`, 'zoom fragment', 3_000);
    assert.equal(await w2.evaluate(`document.querySelector('#original-layout-pdf')!==window.__smokeFrame`), true, 'a fresh frame loads the PDF');
    await w2.evaluate(`(()=>{const s=document.querySelector('#source-view-scale');s.value='FitH';s.dispatchEvent(new Event('change',{bubbles:true}))})()`);
    await w2.until(`document.querySelector('#original-layout-pdf').src.endsWith('#view=FitH&toolbar=0')`, 'fit width', 3_000);
  });

  await check('Edit document puts the caret in the document', async () => {
    await w2.evaluate(`document.querySelector('#edit-layout-button').click()`);
    await w2.until(focused, 'focus after Edit document', 5_000);
    await w2.type('X');
    assert.equal(await w2.evaluate(`${texts}[0]`), 'XOriginal page text');
  });

  // ---- Window 3: Open puts the caret in the opened document -----------------
  await answer({ open: plainDocx });
  await w2.evaluate(`document.querySelector('#open-button').click()`);
  const w3 = await connect([w1.target.id, w2.target.id]);
  await check('Open puts the caret in the opened document, so typing lands there', async () => {
    await w3.until(`document.querySelector('#document-title').textContent==='opened-plain' && document.querySelector('#loading-overlay').hidden`, 'document opened', 45_000);
    await w3.until(focused, 'focus after Open', 5_000);
    await w3.type('Typed ');
    assert.equal(await w3.evaluate(`${texts}[0]`), 'Typed Opened paragraph one');
  });
} catch (error) {
  results.push({ name: 'setup', ok: false, error: String(error?.stack ?? error) });
  console.log(`FAIL setup\n     ${String(error?.stack ?? error)}`);
} finally {
  clearTimeout(hardStop);
  killAll();
  await pause(300);
}

await writeFile(path.join(output, 'results.json'), JSON.stringify(results, null, 2));
const failed = results.filter((result) => !result.ok);
console.log(`\n${results.length - failed.length}/${results.length} editing UI checks passed.`);
if (failed.length) process.exit(1);
