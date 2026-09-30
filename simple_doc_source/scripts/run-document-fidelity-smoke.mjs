import { spawn } from 'node:child_process';
import { createRequire } from 'node:module';
import { mkdtemp, copyFile, readFile, writeFile, rm, mkdir } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import assert from 'node:assert/strict';
const require = createRequire(import.meta.url);
const root = fileURLToPath(new URL('..', import.meta.url));
const fixture = path.resolve(process.env.SIMPLE_DOCS_QA_INPUT || path.join(root, 'qa/fixtures/simple-docs-roundtrip-fixture.docx'));
const work = await mkdtemp(path.join(tmpdir(), 'simple-doc-fidelity-'));
const documentPath = path.join(work, `fidelity-source${path.extname(fixture)}`);
const output = path.resolve(process.env.SIMPLE_DOCS_QA_OUTPUT || path.join(root, '../.codex-tmp/document-fidelity'));
await mkdir(output, { recursive: true });
await copyFile(fixture, documentPath);
const original = await readFile(documentPath);
const port = 21000 + Math.floor(Math.random() * 900);
const child = spawn(require('electron'), [`--remote-debugging-port=${port}`, `--user-data-dir=${path.join(work, 'profile')}`, '.', documentPath], { cwd: root, windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'] });
let socket;
let id = 0;
const pending = new Map();
const send = (method, params = {}) => new Promise((resolve, reject) => {
  const requestId = ++id;
  pending.set(requestId, { resolve, reject });
  socket.send(JSON.stringify({ id: requestId, method, params }));
});
const evaluate = async (expression) => {
  const result = await send('Runtime.evaluate', { expression: `(async () => (${expression}))()`, awaitPromise: true, returnByValue: true });
  if (result.exceptionDetails) throw new Error(JSON.stringify(result.exceptionDetails));
  return result.result.value;
};
const waitFor = async (expression, timeout = 90000) => {
  const end = Date.now() + timeout;
  while (Date.now() < end) {
    if (await evaluate(expression)) return;
    await new Promise(resolve => setTimeout(resolve, 120));
  }
  throw new Error(`Timed out: ${expression}`);
};
const key = async (value, code, number, modifiers = 2) => {
  await send('Input.dispatchKeyEvent', { type: 'keyDown', key: value, code, windowsVirtualKeyCode: number, modifiers });
  await send('Input.dispatchKeyEvent', { type: 'keyUp', key: value, code, windowsVirtualKeyCode: number, modifiers });
};
try {
  let target;
  const started = Date.now();
  while (!target && Date.now() - started < 25000) {
    try { target = (await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()).find(item => item.type === 'page'); } catch {}
    if (!target) await new Promise(resolve => setTimeout(resolve, 100));
  }
  assert.ok(target, 'Electron renderer must start');
  socket = new WebSocket(target.webSocketDebuggerUrl);
  await new Promise(resolve => socket.addEventListener('open', resolve, { once: true }));
  socket.addEventListener('message', event => {
    const message = JSON.parse(String(event.data));
    const entry = pending.get(message.id);
    if (!entry) return;
    pending.delete(message.id);
    if (message.error) entry.reject(new Error(message.error.message));
    else entry.resolve(message.result);
  });
  await waitFor(`document.querySelector('#editor.is-active') && document.querySelector('#loading-overlay').hidden`);
  await waitFor(`document.querySelector('#document-layout-bar').hidden || document.querySelector('#original-layout-pdf').src.startsWith('blob:')`);
  const openedMs = Date.now() - started;
  const fontAudit = await evaluate(`({ registered: (await window.simpleDocs.getDocumentFonts()).map(f => f.family), loaded: [...new Set([...document.fonts].filter(f => f.status === 'loaded').map(f => f.family))], importNotice: document.querySelector('#document-compatibility').textContent, noticeVisible: !document.querySelector('#document-compatibility').hidden })`);
  await new Promise(resolve => setTimeout(resolve, 900));
  const screen = await send('Page.captureScreenshot', { format: 'png', fromSurface: true });
  await writeFile(path.join(output, 'editor.png'), Buffer.from(screen.data, 'base64'));
  if (await evaluate(`!document.querySelector('#document-layout-bar').hidden`)) {
    await evaluate(`document.querySelector('#edit-layout-button').click()`);
    await new Promise(resolve => setTimeout(resolve, 400));
    const editScreen = await send('Page.captureScreenshot', { format: 'png', fromSurface: true });
    await writeFile(path.join(output, 'editor-editview.png'), Buffer.from(editScreen.data, 'base64'));
    await evaluate(`document.querySelector('#original-layout-button').click()`);
  }
  if (['.docx', '.doc'].includes(path.extname(fixture).toLowerCase())) {
    await key('s', 'KeyS', 83);
    await waitFor(`document.querySelector('#toast').textContent === 'Saved.'`);
    assert.deepEqual(await readFile(documentPath), original, 'saving an unedited Word file must preserve every original byte');
  }
  const previewStarted = Date.now();
  await key('p', 'KeyP', 80);
  await waitFor(`!document.querySelector('#print-modal').hidden && document.querySelector('#print-preview-pdf').src.startsWith('blob:')`);
  const firstPrintMs = Date.now() - previewStarted;
  const bytes = await evaluate(`Array.from(new Uint8Array(await (await fetch(document.querySelector('#print-preview-pdf').src)).arrayBuffer()))`);
  await writeFile(path.join(output, 'simple-print.pdf'), Buffer.from(bytes));
  const firstStatus = await evaluate(`document.querySelector('#print-preview-status').textContent`);
  await new Promise(resolve => setTimeout(resolve, 300));
  const printScreen = await send('Page.captureScreenshot', { format: 'png', fromSurface: true });
  await writeFile(path.join(output, 'print.png'), Buffer.from(printScreen.data, 'base64'));
  await key('Escape', 'Escape', 27, 0);
  const repeatStarted = Date.now();
  await key('p', 'KeyP', 80);
  await waitFor(`!document.querySelector('#print-modal').hidden && document.querySelector('#print-preview-pdf').src.startsWith('blob:')`);
  const repeatPrintMs = Date.now() - repeatStarted;
  await key('Escape', 'Escape', 27, 0);
  let editSaveVerified = false;
  if (process.env.SIMPLE_DOCS_QA_EDIT === '1') {
    await evaluate(`document.querySelector('#edit-layout-button').click()`);
    await new Promise(resolve => setTimeout(resolve, 250));
    await evaluate(`[...document.querySelectorAll('button')].find(button => button.title === 'Select all (Ctrl+A)').click()`);
    await new Promise(resolve => setTimeout(resolve, 250));
    await evaluate(`document.querySelector('#editor [contenteditable="true"]').focus()`);
    await send('Input.insertText', { text: ' SIMPLE_QA_EDIT' });
    await waitFor(`document.querySelector('#save-state').classList.contains('is-dirty')`);
    await evaluate(`document.querySelector('#toast').textContent = ''`);
    await key('s', 'KeyS', 83);
    await waitFor(`document.querySelector('#toast').textContent === 'Saved.'`);
    const saved = await readFile(documentPath);
    assert.notDeepEqual(saved, original, 'edited document must be saved');
    if (path.extname(documentPath).toLowerCase() === '.doc') {
      assert.equal(saved.subarray(0, 8).toString('hex'), 'd0cf11e0a1b11ae1', 'edited DOC remains a real binary DOC');
      const extractor = new (require('word-extractor'))();
      const extracted = await extractor.extract(saved);
      assert.match(extracted.getBody(), /SIMPLE_QA_EDIT/, 'saved DOC contains the new edit');
      assert.deepEqual(await readFile(path.join(work, 'fidelity-source.before-simple-edit.doc')), original, 'complex original has an exact pre-edit backup');
    }
    await key('p', 'KeyP', 80);
    await waitFor(`!document.querySelector('#print-modal').hidden && document.querySelector('#print-preview-pdf').src.startsWith('blob:')`);
    const editedPdf = await evaluate(`Array.from(new Uint8Array(await (await fetch(document.querySelector('#print-preview-pdf').src)).arrayBuffer()))`);
    await writeFile(path.join(output, 'simple-print-edited.pdf'), Buffer.from(editedPdf));
    await key('Escape', 'Escape', 27, 0);
    const externalBytes = Buffer.concat([saved, Buffer.from('\nexternal-change')]);
    await writeFile(documentPath, externalBytes);
    await evaluate(`document.querySelector('#toast').textContent = ''`);
    await key('s', 'KeyS', 83);
    await waitFor(`document.querySelector('#toast').textContent.includes('changed outside Simple Docs')`);
    assert.deepEqual(await readFile(documentPath), externalBytes, 'external changes must never be overwritten');
    editSaveVerified = true;
  }
  const audit = { ok: true, openedMs, firstPrintMs, repeatPrintMs, firstStatus, fontAudit, preservedOriginalBytes: true, printBytes: bytes.length, editSaveVerified };
  await writeFile(path.join(output, 'audit.json'), JSON.stringify(audit, null, 2));
  console.log(JSON.stringify(audit, null, 2));
} finally {
  socket?.close();
  child.kill();
  await new Promise(resolve => setTimeout(resolve, 500));
  await rm(work, { recursive: true, force: true, maxRetries: 6, retryDelay: 200 }).catch(() => {});
}
