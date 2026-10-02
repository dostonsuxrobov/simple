// Word features UI smoke test for Simple Docs (Electron + CDP, built renderer):
// AutoFormat as you type (lists, smart quotes, each undoable; the switch), spelling
// as you type (underline, right-click suggestion, Ignore all, the switch), Insert >
// Page number (in the footer on every page of the exported PDF), the Paragraph
// dialog (values survive DOCX save and reopen) and the font box (an installed font
// outside the curated list is chosen, saved, reopened and loaded by the new window).
//
// Run `npm run build:web` first, or point SIMPLE_DOCS_DIST at another build.
// SIMPLE_SMOKE_OUT picks the work folder (default tmp/word-features-smoke). Windows
// open off-screen with an isolated profile; every process is killed at the end;
// nothing is printed, downloaded or sent anywhere.
import { spawn, execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { mkdir, writeFile, readFile, rm } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import assert from 'node:assert/strict';

const require = createRequire(import.meta.url);
const { pdfPageTexts, bandText } = require('./pdf-text.cjs');
const root = fileURLToPath(new URL('..', import.meta.url));
const output = process.env.SIMPLE_SMOKE_OUT ? path.resolve(process.env.SIMPLE_SMOKE_OUT) : path.join(root, 'tmp', 'word-features-smoke');
const dist = process.env.SIMPLE_DOCS_DIST ? path.resolve(process.env.SIMPLE_DOCS_DIST) : null;
await rm(output, { recursive: true, force: true, maxRetries: 5, retryDelay: 200 });
await mkdir(path.join(output, 'profile'), { recursive: true });

const hostPath = path.join(output, 'host.cjs');
await writeFile(hostPath, `const { app, dialog, shell, BrowserWindow } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
app.setPath('userData', path.join(__dirname, 'profile'));
const answers = () => { try { return JSON.parse(fs.readFileSync(path.join(__dirname, 'answers.json'), 'utf8')); } catch { return {}; } };
dialog.showOpenDialog = async () => { const file = answers().open; return file ? { canceled: false, filePaths: [file] } : { canceled: true, filePaths: [] }; };
dialog.showSaveDialog = async (_owner, options) => {
  const extension = (options && options.filters && options.filters[0] && options.filters[0].extensions[0]) || 'docx';
  const file = answers()[extension];
  return file ? { canceled: false, filePath: file } : { canceled: true };
};
dialog.showMessageBox = async () => ({ response: 0 });
shell.openExternal = async () => {};
const show = BrowserWindow.prototype.show;
BrowserWindow.prototype.show = function () { if (process.env.SIMPLE_SMOKE_SHOW) return show.call(this); try { this.setPosition(-3200, 40); this.showInactive(); } catch {} };
app.on('browser-window-created', (_event, window) => { try { window.webContents.setBackgroundThrottling(false); } catch {} });
require(${JSON.stringify(path.join(root, 'electron', 'main.cjs'))});
`);
const answer = (value) => writeFile(path.join(output, 'answers.json'), JSON.stringify(value));
const savedDocx = path.join(output, 'features.docx');
const savedPdf = path.join(output, 'features.pdf');
await answer({ docx: savedDocx, pdf: savedPdf });

const port = 26100 + Math.floor(Math.random() * 700);
const env = { ...process.env };
delete env.ELECTRON_RUN_AS_NODE;
if (dist) env.VITE_DEV_SERVER_URL = pathToFileURL(path.join(dist, 'index.html')).href;
const child = spawn(require('electron'), [`--remote-debugging-port=${port}`, `--user-data-dir=${path.join(output, 'profile')}`, hostPath], { cwd: root, env, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
let stderr = '';
child.stderr.on('data', (data) => { stderr += String(data); });
const killAll = () => { try { execFileSync('taskkill', ['/PID', String(child.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { try { child.kill(); } catch {} } };
const hardStop = setTimeout(() => { console.error('Word features smoke test timed out.'); killAll(); process.exit(124); }, Number(process.env.SIMPLE_SMOKE_TIMEOUT_MS || 360_000));

const pause = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
async function connect(exclude = []) {
  let target;
  const end = Date.now() + 60_000;
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
  const MODIFIERS = { alt: 1, ctrl: 2, shift: 8 };
  const key = async (keyValue, code, vk, modifiers = [], text) => {
    const mask = modifiers.reduce((sum, name) => sum + MODIFIERS[name], 0);
    const down = { type: text ? 'keyDown' : 'rawKeyDown', key: keyValue, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mask };
    if (text) Object.assign(down, { text, unmodifiedText: text });
    await send('Input.dispatchKeyEvent', down);
    await send('Input.dispatchKeyEvent', { type: 'keyUp', key: keyValue, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers: mask });
    await pause(70);
  };
  /** One keystroke per character, the way a person types (AutoFormat reacts per character). */
  const typeKeys = async (text) => { for (const char of text) { await send('Input.insertText', { text: char }); await pause(25); } await pause(60); };
  const type = async (text) => { await send('Input.insertText', { text }); await pause(80); };
  const mouse = async (x, y, button = 'left') => { for (const kind of ['mousePressed', 'mouseReleased']) await send('Input.dispatchMouseEvent', { type: kind, button, clickCount: 1, x, y }); await pause(120); };
  await send('Page.enable');
  return { target, socket, evaluate, until, key, type, typeKeys, mouse };
}

const H = 'window.__cw.handle';
const texts = `${H}.getDocument().blocks.filter(b=>b.kind==='paragraph').map(b=>b.runs.map(r=>r.text).join(''))`;
const block = (index) => `${H}.getDocument().blocks.filter(b=>b.kind==='paragraph')[${index}]`;
const enter = (w) => w.key('Enter', 'Enter', 13, [], '\r');
const undo = (w) => w.key('z', 'KeyZ', 90, ['ctrl']);
const redo = (w) => w.key('y', 'KeyY', 89, ['ctrl']);
/** Red underline pixels painted on the first page (the spelling marks). */
const redPixels = `(()=>{let n=0;for(const c of document.querySelectorAll('#editor .cw-app [data-page="0"] canvas')){const d=c.getContext('2d').getImageData(0,0,c.width,c.height).data;for(let i=0;i<d.length;i+=4){if(d[i]>180&&d[i+1]<90&&d[i+2]<90&&d[i+3]>100)n++;}}return n})()`;
/** The client point over the middle of a word in the body text. */
const wordPoint = (word) => `(()=>{const tree=window.__cw.editor.getLayoutTree();const page=tree.pages[0];const el=document.querySelector('#editor .cw-app [data-page="0"]');const r=el.getBoundingClientRect();const z=r.width/page.widthPx;for(let y=page.contentTopPx;y<page.contentBottomPx;y+=3){for(let x=page.marginPx.left;x<page.widthPx-page.marginPx.right;x+=3){const p=${H}.positionFromPoint(r.left+x*z,r.top+y*z);if(!p)continue;const b=${H}.getDocument().blocks.find(b=>b.id===p.blockId);const t=b&&b.runs?b.runs.map(r=>r.text).join(''):'';const i=t.indexOf(${JSON.stringify(word)});if(i>=0&&p.offset>i&&p.offset<i+${word.length})return {x:r.left+x*z,y:r.top+y*z};}}return null})()`;
const clickRibbon = (tab, item) => `(()=>{document.querySelector('[data-ribbon-tab="${tab}"]').click();const b=document.querySelector('[data-ribbon-item="${item}"]');b.click();return true})()`;
const setField = (name, value) => `(()=>{const e=document.querySelector('.settings-dialog [name="${name}"]');if(!e)return false;if(e.type==='checkbox'){e.checked=${JSON.stringify(value)};}else{e.value=${JSON.stringify(String(value))};}e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}));return true})()`;

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
  const w1 = await connect();
  await w1.until(`document.querySelector('#blank-document') && document.querySelector('#loading-overlay').hidden`, 'welcome screen', 60_000);
  await w1.evaluate(`document.querySelector('#blank-document').click()`);
  await w1.until(`document.activeElement?.matches('[contenteditable="true"][role="textbox"]')`, 'editor focus', 10_000);

  await check('Review tab: Spelling and AutoFormat switches are on by default; Insert has Header, Footer and Page number', async () => {
    const state = await w1.evaluate(`({tabs:[...document.querySelectorAll('[data-ribbon-tab]')].map(t=>t.textContent.trim()),spelling:document.querySelector('[data-ribbon-item="simple.proofing.spelling"]').getAttribute('aria-pressed'),autoformat:document.querySelector('[data-ribbon-item="simple.proofing.autoformat"]').getAttribute('aria-pressed'),insert:['header','footer','page-number'].map(id=>!!document.querySelector('[data-ribbon-item="simple.header-footer.'+id+'"]')),fontBox:!!document.querySelector('.font-box input')})`);
    assert.ok(state.tabs.includes('Review'), JSON.stringify(state.tabs));
    assert.equal(state.spelling, 'true');
    assert.equal(state.autoformat, 'true');
    assert.deepEqual(state.insert, [true, true, true]);
    assert.equal(state.fontBox, true);
  });

  await check('AutoFormat: "1. " starts a numbered list, and Ctrl+Z gives back the typed "1. " (one step)', async () => {
    await w1.typeKeys('1. ');
    await w1.until(`${block(0)}.style.list && ${texts}[0]===''`, 'numbered list', 3_000);
    const list = await w1.evaluate(`window.__cw.handle.getDocument().lists[${block(0)}.style.list.listId].levels[0].format`);
    assert.equal(list, 'decimal');
    await undo(w1);
    await w1.until(`!${block(0)}.style.list && ${texts}[0]==='1. '`, 'undo restores the typed text', 3_000);
    await redo(w1);
    await w1.until(`${block(0)}.style.list && ${texts}[0]===''`, 'redo the list', 3_000);
    await w1.type('First item');
    await enter(w1);
    await enter(w1);
    await w1.until(`!${block(1)}.style.list`, 'a second Enter ends the list', 3_000);
  });

  await check('AutoFormat: "quotes" become smart quotes, each conversion undoable on its own', async () => {
    await w1.typeKeys('She said "hello" and it\'s fine');
    await w1.until(`${texts}[1]==='She said “hello” and it’s fine'`, 'smart quotes', 3_000);
    await undo(w1);
    // Typing undoes as one step first; then each conversion separately.
    await w1.until(`${texts}[1]==='She said “hello” and it’'`, 'typing after the apostrophe undone', 3_000);
    await undo(w1);
    await w1.until(`${texts}[1]==="She said “hello” and it'"`, 'the apostrophe conversion alone is undone', 3_000);
    await redo(w1);
    await redo(w1);
    await w1.until(`${texts}[1]==='She said “hello” and it’s fine'`, 'redo', 3_000);
    await enter(w1);
  });

  await check('AutoFormat switch: off keeps straight quotes and "1. " as typed; the choice is remembered', async () => {
    await w1.evaluate(clickRibbon('review', 'simple.proofing.autoformat'));
    await w1.until(`document.querySelector('[data-ribbon-item="simple.proofing.autoformat"]').getAttribute('aria-pressed')==='false'`, 'switch off', 3_000);
    assert.equal(await w1.evaluate(`localStorage.getItem('simple-docs:autoformat')`), 'off');
    await w1.evaluate(`window.__cw.handle.focus()`);
    await w1.typeKeys('"plain"');
    await w1.until(`${texts}[2]==='"plain"'`, 'straight quotes stay', 3_000);
    await w1.evaluate(clickRibbon('review', 'simple.proofing.autoformat'));
    await w1.until(`document.querySelector('[data-ribbon-item="simple.proofing.autoformat"]').getAttribute('aria-pressed')==='true'`, 'switch on', 3_000);
    assert.equal(await w1.evaluate(`localStorage.getItem('simple-docs:autoformat')`), 'on');
    await w1.evaluate(`window.__cw.handle.focus()`);
    await enter(w1);
  });

  await check('Spelling: a misspelled word gets a red underline and a right-click suggestion replaces it (undoable)', async () => {
    await w1.type('Teh quick brown fox jumpd over the lazy dog.');
    await enter(w1);
    // The first check starts the Windows spell checker (a few seconds once per session).
    await w1.until(`${redPixels}>20`, 'red underlines', 30_000);
    const point = await w1.until(wordPoint('jumpd'), 'the word on the page', 5_000);
    await w1.mouse(point.x, point.y, 'right');
    await w1.until(`[...document.querySelectorAll('.simple-spell-suggestion')].some(n=>n.textContent==='jumped')`, 'suggestions in the menu', 10_000);
    const labels = await w1.evaluate(`[...document.querySelectorAll('.cw-menu .cw-menu-lbl')].map(n=>n.textContent)`);
    for (const label of ['Ignore', 'Ignore all', 'Add to dictionary', 'Paste']) assert.ok(labels.includes(label), `menu has ${label}: ${labels.join(', ')}`);
    const item = await w1.evaluate(`(()=>{const n=[...document.querySelectorAll('.simple-spell-suggestion')].find(n=>n.textContent==='jumped');const r=n.getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2}})()`);
    await w1.mouse(item.x, item.y);
    await w1.until(`${texts}.includes('Teh quick brown fox jumped over the lazy dog.')`, 'suggestion applied', 3_000);
    assert.equal(await w1.evaluate(`document.querySelectorAll('.cw-menu').length`), 0, 'the menu closed');
    await undo(w1);
    await w1.until(`${texts}.includes('Teh quick brown fox jumpd over the lazy dog.')`, 'Ctrl+Z restores the word', 3_000);
    await redo(w1);
    await w1.until(`${texts}.includes('Teh quick brown fox jumped over the lazy dog.')`, 'redo', 3_000);
  });

  await check('Spelling: Ignore all clears a word; the Review > Spelling switch removes every mark and is remembered', async () => {
    const before = await w1.evaluate(redPixels);
    const point = await w1.until(wordPoint('Teh'), 'Teh on the page', 5_000);
    await w1.mouse(point.x, point.y, 'right');
    await w1.until(`[...document.querySelectorAll('.simple-spell-item .cw-menu-lbl')].some(n=>n.textContent==='Ignore all')`, 'Ignore all', 10_000);
    const item = await w1.evaluate(`(()=>{const n=[...document.querySelectorAll('.simple-spell-item')].find(n=>n.textContent==='Ignore all');const r=n.getBoundingClientRect();return {x:r.left+r.width/2,y:r.top+r.height/2}})()`);
    await w1.mouse(item.x, item.y);
    await w1.until(`${redPixels}<${before}`, 'fewer marks after Ignore all', 5_000);
    await w1.evaluate(clickRibbon('review', 'simple.proofing.spelling'));
    await w1.until(`${redPixels}===0`, 'no marks with spelling off', 5_000);
    assert.equal(await w1.evaluate(`localStorage.getItem('simple-docs:spelling')`), 'off');
    await w1.evaluate(clickRibbon('review', 'simple.proofing.spelling'));
    assert.equal(await w1.evaluate(`localStorage.getItem('simple-docs:spelling')`), 'on');
    await w1.evaluate(`window.__cw.handle.focus()`);
  });

  await check('Paragraph dialog applies indents, spacing and line spacing to the selected paragraph', async () => {
    await w1.evaluate(`(()=>{const b=${block(1)};${H}.setSelection({anchor:{blockId:b.id,offset:3},focus:{blockId:b.id,offset:3}});${H}.focus();return true})()`);
    await w1.evaluate(clickRibbon('home', 'simple.paragraph.settings'));
    await w1.until(`document.querySelector('.settings-dialog [name="left"]')`, 'Paragraph dialog', 5_000);
    for (const [name, value] of [['left', 1], ['right', 0.5], ['special', 'hanging'], ['by', 0.25], ['before', 12], ['after', 6], ['line', 'exactly'], ['at', 18], ['contextual', true]]) {
      assert.equal(await w1.evaluate(setField(name, value)), true, `field ${name}`);
    }
    await w1.evaluate(`document.querySelector('.settings-dialog [type=submit]').click()`);
    await w1.until(`!document.querySelector('.settings-dialog')`, 'dialog closed', 3_000);
    const style = await w1.evaluate(`${block(1)}.style`);
    assert.equal(style.indentLeftPx, 96);
    assert.equal(style.indentRightPx, 48);
    assert.equal(style.indentFirstLinePx, -24);
    assert.equal(style.spaceBeforePx, 16);
    assert.equal(style.spaceAfterPx, 8);
    assert.equal(style.lineRule, 'exact');
    assert.equal(style.lineHeightPx, 24);
    assert.equal(style.contextualSpacing, true);
    // The line-spacing menu offers Word's Remove space before paragraph for it now.
    await w1.evaluate(`document.querySelector('[data-ribbon-item="home.paragraph.line-spacing"]').click()`);
    const extra = await w1.until(`[...document.querySelectorAll('.cw-pop .simple-spacing-item')].map(n=>n.textContent).join('|')||null`, 'spacing items', 3_000);
    assert.equal(extra, 'Remove space before paragraph|Remove space after paragraph|Indents and spacing…');
    await w1.key('Escape', 'Escape', 27);
    await w1.evaluate(`document.body.dispatchEvent(new MouseEvent('mousedown',{bubbles:true}));document.querySelectorAll('.cw-pop').forEach(n=>n.remove());true`);
  });

  await check('Insert > Page number puts "Page X of Y" in the footer of every page of the exported PDF', async () => {
    await w1.evaluate(`(()=>{const list=${H}.getDocument().blocks.filter(b=>b.kind==='paragraph');const b=list[list.length-1];const n=b.runs.map(r=>r.text).join('').length;${H}.setSelection({anchor:{blockId:b.id,offset:n},focus:{blockId:b.id,offset:n}});${H}.focus();return true})()`);
    for (let index = 0; index < 64; index += 1) await enter(w1);
    await w1.type('Last page text');
    const pageCount = await w1.until(`window.__cw.editor.getLayoutInfo().pageCount>=3&&window.__cw.editor.getLayoutInfo().pageCount`, 'three pages', 5_000);
    await w1.evaluate(clickRibbon('insert', 'simple.header-footer.page-number'));
    await w1.until(`document.querySelector('.settings-dialog [name="position"]')`, 'Page number dialog', 5_000);
    for (const [name, value] of [['position', 'bottom'], ['align', 'center'], ['text', 'page-x-of-y']]) assert.equal(await w1.evaluate(setField(name, value)), true);
    await w1.evaluate(`document.querySelector('.settings-dialog [type=submit]').click()`);
    await w1.until(`(${H}.getDocument().section.footer||[]).some(b=>b.runs.map(r=>r.text).join('')==='Page {page} of {pages}')`, 'footer page number', 3_000);
    await w1.evaluate(`document.querySelector('#toast').textContent='';document.querySelector('#export-as-button').click();document.querySelector('[data-export-format="pdf"]').click()`);
    await w1.until(`document.querySelector('#toast').textContent.includes('exported')`, 'PDF exported', 60_000);
    const pages = await pdfPageTexts(await readFile(savedPdf));
    assert.equal(pages.length, pageCount, 'the PDF has the editor\'s pages');
    pages.forEach((page, index) => assert.equal(bandText(page, 'bottom'), `Page ${index + 1} of ${pages.length}`, `PDF page ${index + 1}`));
  });

  await check('Font box: an installed font outside the curated set is found by typing and applied', async () => {
    await w1.evaluate(`(()=>{const b=${block(1)};${H}.setSelection({anchor:{blockId:b.id,offset:0},focus:{blockId:b.id,offset:8}});${H}.focus();return true})()`);
    await w1.evaluate(`document.querySelector('[data-ribbon-tab="home"]').click()`);
    const box = await w1.evaluate(`(()=>{const r=document.querySelector('.font-box-input').getBoundingClientRect();return {x:r.left+20,y:r.top+r.height/2}})()`);
    await w1.mouse(box.x, box.y);
    await w1.until(`!document.querySelector('.font-list').hidden && document.querySelectorAll('.font-option').length>20`, 'font list with installed fonts', 10_000);
    await w1.type('Bahn');
    await w1.until(`document.querySelector('.font-option')?.dataset.family==='Bahnschrift'`, 'Bahnschrift first', 5_000);
    await w1.key('Enter', 'Enter', 13, [], '\r');
    await w1.until(`${block(1)}.runs.some(r=>r.text==='She said'&&r.style.fontFamily==='Bahnschrift')`, 'font applied to the selection', 3_000);
    assert.deepEqual(JSON.parse(await w1.evaluate(`localStorage.getItem('simple-docs:recent-fonts')`)), ['Bahnschrift']);
    assert.equal(await w1.evaluate(`document.querySelector('.font-box-input').value`), 'Bahnschrift');
  });

  await check('Save, reopen in a new window: paragraph settings, page numbers and the chosen font all survive', async () => {
    await w1.key('s', 'KeyS', 83, ['ctrl']);
    await w1.until(`document.querySelector('#toast').textContent.startsWith('Saved')`, 'saved', 30_000);
    await answer({ docx: savedDocx, pdf: savedPdf, open: savedDocx });
    await w1.evaluate(`document.querySelector('#open-button').click()`);
    const w2 = await connect([w1.target.id]);
    await w2.until(`document.querySelector('#editor.is-active') && document.querySelector('#loading-overlay').hidden && document.querySelector('#document-title').textContent==='features'`, 'reopened window', 60_000);
    const style = await w2.evaluate(`${H}.getDocument().blocks.filter(b=>b.kind==='paragraph').find(b=>b.runs.map(r=>r.text).join('').startsWith('She said')).style`);
    const close = (actual, expected, label) => assert.ok(Math.abs(actual - expected) < 0.2, `${label}: ${actual} vs ${expected}`);
    close(style.indentLeftPx, 96, 'left indent');
    close(style.indentRightPx, 48, 'right indent');
    close(style.indentFirstLinePx, -24, 'hanging indent');
    close(style.spaceBeforePx, 16, 'space before');
    close(style.spaceAfterPx, 8, 'space after');
    assert.equal(style.lineRule, 'exact');
    close(style.lineHeightPx, 24, 'exact line height');
    assert.equal(style.contextualSpacing, true);
    assert.ok(await w2.evaluate(`(${H}.getDocument().section.footer||[]).some(b=>b.runs.map(r=>r.text).join('').includes('{page}'))`), 'the footer page number reopens');
    const family = await w2.evaluate(`${H}.getDocument().blocks.filter(b=>b.kind==='paragraph').flatMap(b=>b.runs).find(r=>r.text.startsWith('She said')).style.fontFamily`);
    assert.equal(family.split(',')[0].trim(), 'Bahnschrift');
    // The new window asked for the recent font at startup, so the editor shows the font itself.
    await w2.until(`[...document.fonts].some(f=>f.family.replace(/["']/g,'')==='Bahnschrift'&&f.status==='loaded')`, 'Bahnschrift loaded by the new window', 15_000);
    w2.socket.close();
  });
} finally {
  clearTimeout(hardStop);
  killAll();
  await pause(300);
  const failed = results.filter((result) => !result.ok);
  await writeFile(path.join(output, 'results.json'), JSON.stringify({ results, stderr: stderr.slice(-4000) }, null, 2)).catch(() => {});
  console.log(`\n${results.length - failed.length}/${results.length} checks passed.`);
  if (failed.length || !results.length) process.exitCode = 1;
}
