// UI regression for form filling, bookmark editing, edit commit/undo and the
// open paths that used to hang or lose the file path. Launches the built app
// (dist/) in isolated Electron profiles; run `npm run build:web` first.
//
//   A. A form + outline fixture: multi-line text keeps its line breaks, MaxLen
//      is enforced, a radio group, a dropdown and a multi-select list box are
//      fillable; renaming a bookmark in place (no window.prompt), cancelling a
//      rename, adding a toolbar bookmark; Escape keeps a typed text box (shown
//      top-anchored, as saved) and a placed signature; Ctrl+Z after the Bold
//      button undoes only the bold; Ctrl+S while typing applies the box and
//      saves in place, keeping Undo/Redo (undo past the save shows unsaved
//      changes, redo back to the saved state shows none); Ctrl+S while
//      renaming a bookmark keeps the new name and saves it. The saved file is
//      checked with pdf-lib and pdf.js: field values, every outline item with
//      its exact destination or link, the committed text and the signature.
//   B. A password-protected PDF shows the password prompt (wrong password,
//      then the right one) instead of hanging on "Opening document…".
//   C. A PDF dropped on the home screen from disk keeps its path (it is
//      remembered in Recent files, so Save writes back to it).
//   D. A signed PDF explains that Save keeps the original and writes a copy.
//   E. Ctrl+S pressed during a page operation saves once it has finished.
//
// SIMPLE_FORMS_SMOKES picks scenarios: forms,password,drop,signed,queue.
// Every launched process tree is killed; temporary files live under
// SIMPLE_TEST_SCRATCH (default: the system temp folder).
import assert from 'node:assert/strict'
import { spawn, spawnSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import os from 'node:os'
import path from 'node:path'
import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNumber, PDFRef, PDFString, StandardFonts } from 'pdf-lib'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)
const { encryptedPdf, linearizedLayout } = require('../tests/helpers/encrypted-fixtures.cjs')

const root = path.resolve('.')
const localElectron = path.join(root, 'node_modules', 'electron', 'dist', 'electron.exe')
const electron = process.env.SIMPLE_ELECTRON
  || (existsSync(localElectron) ? localElectron : path.resolve(root, '..', 'simple', 'node_modules', 'electron', 'dist', 'electron.exe'))
const packagedExecutable = process.env.SIMPLE_TEST_EXECUTABLE ? path.resolve(process.env.SIMPLE_TEST_EXECUTABLE) : ''
const executable = packagedExecutable || electron
const scratch = process.env.SIMPLE_TEST_SCRATCH ? path.resolve(process.env.SIMPLE_TEST_SCRATCH) : os.tmpdir()
const basePort = Number(process.env.SIMPLE_BENCH_PORT || 9497)
const only = new Set((process.env.SIMPLE_FORMS_SMOKES || 'forms,password,drop,signed,queue').split(',').map((item) => item.trim()))
const pause = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds))

// ---------------------------------------------------------------- fixtures

/**
 * Five pages. Page 1 holds the form; the outline has everything a rebuild
 * used to lose: an /XYZ destination with zoom, a web link, a closed heading
 * without a destination and its children, a GoTo action.
 */
async function formsAndOutlineFixture() {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  const pages = Array.from({ length: 5 }, (_, index) => {
    const page = doc.addPage([500, 760])
    page.drawText(`Fixture page ${index + 1}`, { x: 40, y: 720, size: 18, font })
    return page
  })
  const form = doc.getForm()
  const box = (y, height = 20, width = 200) => ({ x: 40, y, width, height })
  const zip = form.createTextField('zip')
  zip.setMaxLength(5)
  zip.addToPage(pages[0], box(660))
  const address = form.createTextField('Address')
  address.enableMultiline()
  address.addToPage(pages[0], box(570, 64))
  const color = form.createRadioGroup('color')
  ;['red', 'green', 'blue'].forEach((option, index) => color.addOptionToPage(option, pages[0], { x: 40 + index * 44, y: 530, width: 16, height: 16 }))
  const country = form.createDropdown('country')
  country.addOptions(['CA', 'US', 'UZ'])
  country.addToPage(pages[0], box(480))
  country.acroField.dict.set(PDFName.of('Opt'), doc.context.obj([
    [PDFString.of('CA'), PDFString.of('Canada')],
    [PDFString.of('US'), PDFString.of('United States')],
    [PDFString.of('UZ'), PDFString.of('Uzbekistan')],
  ]))
  const langs = form.createOptionList('langs')
  langs.addOptions(['en', 'ru', 'uz'])
  langs.enableMultiselect()
  langs.addToPage(pages[0], box(390, 64))

  const { context } = doc
  const [chapter1, website, part2, chapter2, chapter3, notes] = Array.from({ length: 6 }, () => context.nextRef())
  const outlineRoot = context.nextRef()
  const item = (title, extra) => context.obj({ Title: PDFHexString.fromText(title), ...extra })
  const dest = (...parts) => context.obj(parts)
  context.assign(chapter1, item('Chapter 1', { Parent: outlineRoot, Next: website, F: 2, Dest: dest(pages[1].ref, PDFName.of('XYZ'), 72, 700, 2) }))
  context.assign(website, item('Website', { Parent: outlineRoot, Prev: chapter1, Next: part2, A: { S: 'URI', URI: PDFString.of('https://example.com/guide') } }))
  context.assign(part2, item('Part II', { Parent: outlineRoot, Prev: website, Next: notes, First: chapter2, Last: chapter3, Count: -2 }))
  context.assign(chapter2, item('Chapter 2', { Parent: part2, Next: chapter3, Dest: dest(pages[2].ref, PDFName.of('XYZ'), 0, 500, null) }))
  context.assign(chapter3, item('Chapter 3', { Parent: part2, Prev: chapter2, Dest: dest(pages[3].ref, PDFName.of('FitH'), 300) }))
  context.assign(notes, item('Notes', { Parent: outlineRoot, Prev: part2, A: { S: 'GoTo', D: dest(pages[4].ref, PDFName.of('Fit')) } }))
  context.assign(outlineRoot, context.obj({ Type: 'Outlines', First: chapter1, Last: notes, Count: 4 }))
  doc.catalog.set(PDFName.of('Outlines'), outlineRoot)
  return doc.save({ updateFieldAppearances: false })
}

/** A plain document; `signed` adds an AcroForm with /SigFlags 3 (a signed form). */
async function plainFixture(label, pageCount, { signed = false } = {}) {
  const doc = await PDFDocument.create()
  const font = await doc.embedFont(StandardFonts.Helvetica)
  for (let index = 0; index < pageCount; index += 1) {
    doc.addPage([400, 520]).drawText(`${label} page ${index + 1}`, { x: 40, y: 460, size: 16, font })
  }
  if (signed) doc.catalog.set(PDFName.of('AcroForm'), doc.context.obj({ Fields: [], SigFlags: 3 }))
  return doc.save()
}

/**
 * A password-protected PDF in the layout of a linearized ("Fast Web View")
 * file: /Encrypt sits only in the first-page trailer near the start, and the
 * last 256 KB do not mention it (the case a tail scan missed).
 */
async function protectedFixture(password) {
  const plain = await encryptedPdf('Protected sample text', `encrypt=aes-256,user-password=${password},owner-password=${password}-owner`)
  return linearizedLayout(plain)
}

// ---------------------------------------------------------------- app + CDP

function killTree(child) {
  if (!child || child.exitCode !== null) return
  spawnSync('taskkill.exe', ['/pid', String(child.pid), '/t', '/f'], { stdio: 'ignore', windowsHide: true })
}

async function withApp(name, port, file, scenario) {
  const profile = await mkdtemp(path.join(work, `profile-${name}-`))
  const app = spawn(executable, [
    ...(!packagedExecutable ? ['.'] : []),
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${profile}`,
    ...(file ? [file] : []),
  ], { cwd: root, windowsHide: true, stdio: 'ignore' })
  const deadline = setTimeout(() => killTree(app), 240_000)
  let page
  try {
    page = await connect(port)
    return await scenario(page)
  } catch (error) {
    const state = page ? await page.evaluate(`({ toast: document.querySelector('.toast')?.textContent, busy: document.querySelector('.busy-overlay')?.textContent, dialog: document.querySelector('[role="dialog"]')?.textContent?.slice(0, 300), active: document.activeElement?.outerHTML?.slice(0, 300) })`).catch(() => null) : null
    throw new Error(`[${name}] ${error.message}\nUI state: ${JSON.stringify(state)}\nRuntime errors: ${JSON.stringify(page?.runtimeErrors.slice(0, 5) ?? [])}`, { cause: error })
  } finally {
    clearTimeout(deadline)
    page?.close()
    killTree(app)
    await Promise.race([new Promise((resolve) => app.once('exit', resolve)), pause(3_000)])
  }
}

async function connect(port) {
  const startupDeadline = Date.now() + 90_000
  let target
  while (!target && Date.now() < startupDeadline) {
    try {
      const targets = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json()
      target = targets.find((item) => item.type === 'page' && item.webSocketDebuggerUrl)
    } catch { /* Electron is still opening its debugging endpoint. */ }
    if (!target) await pause(60)
  }
  if (!target) throw new Error('The renderer did not start')
  const socket = new WebSocket(target.webSocketDebuggerUrl)
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true })
    socket.addEventListener('error', reject, { once: true })
  })
  let requestId = 0
  const pending = new Map()
  const runtimeErrors = []
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(String(event.data))
    if (message.method === 'Runtime.exceptionThrown') runtimeErrors.push(message.params.exceptionDetails?.exception?.description || message.params.exceptionDetails?.text)
    const request = pending.get(message.id)
    if (!request) return
    pending.delete(message.id)
    clearTimeout(request.timeout)
    if (message.error) request.reject(new Error(message.error.message))
    else if (message.result?.exceptionDetails) request.reject(new Error(message.result.exceptionDetails.exception?.description || message.result.exceptionDetails.text))
    else request.resolve(message.result?.result?.value ?? message.result)
  })
  const call = (method, params = {}) => new Promise((resolve, reject) => {
    const id = ++requestId
    const timeout = setTimeout(() => { pending.delete(id); reject(new Error(`Timed out: ${method}`)) }, 30_000)
    pending.set(id, { resolve, reject, timeout })
    socket.send(JSON.stringify({ id, method, params }))
  })
  const evaluate = (expression) => call('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true })
  await call('Runtime.enable')
  // Fail loudly instead of hanging if anything still calls a blocking dialog.
  await evaluate(`window.prompt = () => { throw new Error('window.prompt is not supported in Electron') }; true`)

  async function waitFor(expression, label, milliseconds = 20_000) {
    const deadline = Date.now() + milliseconds
    while (Date.now() < deadline) {
      // Strictly true: an `undefined` result must not count (evaluate()
      // hands back the whole CDP result object when there is no value).
      if (await evaluate(`Boolean(${expression})`).catch(() => false) === true) return
      await pause(50)
    }
    throw new Error(`Timed out waiting for ${label}`)
  }
  async function key(name, code, modifiers = 0, text) {
    const virtual = ({ Enter: 13, Escape: 27, Tab: 9, Backspace: 8 })[name] || name.toUpperCase().charCodeAt(0)
    const parameters = { key: name, code, modifiers, windowsVirtualKeyCode: virtual, nativeVirtualKeyCode: virtual }
    await call('Input.dispatchKeyEvent', { type: text ? 'keyDown' : 'rawKeyDown', ...parameters, ...(text ? { text, unmodifiedText: text } : {}) })
    await call('Input.dispatchKeyEvent', { type: 'keyUp', ...parameters })
  }
  async function clickAt(expression, xFraction = 0.5, yFraction = 0.5) {
    await waitFor(`Boolean(${expression})`, `click target ${expression}`)
    // Pages refit and re-render after opening; click only once the target
    // has stopped moving, and only if it is really the element under the point.
    let point = null
    const deadline = Date.now() + 30_000
    while (Date.now() < deadline) {
      const sample = await evaluate(`(() => {
        const element = ${expression}
        element.scrollIntoView({ block: 'nearest', inline: 'nearest' })
        const box = element.getBoundingClientRect()
        const x = box.left + box.width * ${xFraction}
        const y = box.top + box.height * ${yFraction}
        const hit = document.elementFromPoint(x, y)
        // A page that is still re-rendering hides its controls and ignores clicks.
        const settledPage = !element.closest('.is-page-transitioning') && !document.querySelector('.is-page-transitioning')
        return { x, y, hit: settledPage && Boolean(hit && (hit === element || element.contains(hit))), what: hit ? hit.outerHTML.slice(0, 160) : null }
      })()`)
      if (point && sample.hit && point.hit && Math.abs(sample.x - point.x) < 0.5 && Math.abs(sample.y - point.y) < 0.5) break
      point = sample
      await pause(150)
    }
    if (!point?.hit) throw new Error(`Click target is covered: ${expression} → ${point?.what}`)
    await call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...point })
    await call('Input.dispatchMouseEvent', { type: 'mousePressed', ...point, button: 'left', buttons: 1, clickCount: 1 })
    await call('Input.dispatchMouseEvent', { type: 'mouseReleased', ...point, button: 'left', buttons: 0, clickCount: 1 })
  }
  const insertText = (text) => call('Input.insertText', { text })
  // Real key presses, one per character (as a person types).
  async function typeText(text) {
    for (const character of text) {
      const upper = character.toUpperCase()
      const parameters = { key: character, code: /[a-z]/i.test(character) ? `Key${upper}` : '', windowsVirtualKeyCode: upper.charCodeAt(0), nativeVirtualKeyCode: upper.charCodeAt(0) }
      await call('Input.dispatchKeyEvent', { type: 'keyDown', ...parameters, text: character, unmodifiedText: character })
      await call('Input.dispatchKeyEvent', { type: 'keyUp', ...parameters })
    }
  }
  return {
    call, evaluate, waitFor, key, clickAt, insertText, typeText, runtimeErrors,
    close: () => {
      for (const request of pending.values()) clearTimeout(request.timeout)
      socket.close()
    },
  }
}

// ---------------------------------------------------------------- checks

/** Field values in a saved file, read with pdf-lib. */
async function savedFormValues(bytes) {
  const form = (await PDFDocument.load(bytes)).getForm()
  return {
    zip: form.getTextField('zip').getText() ?? '',
    Address: form.getTextField('Address').getText() ?? '',
    color: form.getRadioGroup('color').getSelected() ?? null,
    country: form.getDropdown('country').getSelected(),
    langs: [...form.getOptionList('langs').getSelected()].sort(),
  }
}

/** The saved outline: title, depth and what each item points at. */
async function savedOutline(bytes) {
  const doc = await PDFDocument.load(bytes, { updateMetadata: false })
  const { context } = doc
  const pageNumber = new Map(doc.getPages().map((page, index) => [page.ref.tag, index + 1]))
  const describe = (value) => {
    const resolved = context.lookup(value)
    if (!(resolved instanceof PDFArray)) return resolved === undefined ? null : String(resolved)
    return resolved.asArray().map((part, index) => {
      if (index === 0 && part instanceof PDFRef) return pageNumber.has(part.tag) ? `page${pageNumber.get(part.tag)}` : 'missing'
      if (part instanceof PDFName) return part.decodeText()
      if (part instanceof PDFNumber) return part.asNumber()
      return String(part)
    }).join(' ')
  }
  const items = []
  const walk = (parent, depth) => {
    let value = parent.get(PDFName.of('First'))
    while (value instanceof PDFRef) {
      const dict = context.lookup(value, PDFDict)
      const action = dict.lookupMaybe(PDFName.of('A'), PDFDict)
      const count = dict.lookupMaybe(PDFName.of('Count'), PDFNumber)
      items.push({
        title: dict.lookup(PDFName.of('Title')).decodeText(),
        depth,
        target: dict.get(PDFName.of('Dest')) !== undefined
          ? describe(dict.get(PDFName.of('Dest')))
          : action
            ? `${action.get(PDFName.of('S')).decodeText()}:${action.get(PDFName.of('URI')) ? action.lookup(PDFName.of('URI')).decodeText() : describe(action.get(PDFName.of('D')))}`
            : null,
        count: count ? count.asNumber() : null,
      })
      walk(dict, depth + 1)
      value = dict.get(PDFName.of('Next'))
    }
  }
  const outlineRoot = doc.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)
  if (outlineRoot) walk(outlineRoot, 0)
  return items
}

async function pageText(bytes, pageNumber) {
  const document = await pdfjs.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise
  try {
    const content = await (await document.getPage(pageNumber)).getTextContent()
    return content.items.map((item) => item.str || '').join(' ')
  } finally {
    await document.destroy()
  }
}

// ---------------------------------------------------------------- scenarios

const dirtyDot = `Boolean(document.querySelector('.dirty-dot'))`
const settled = `!document.querySelector('.busy-overlay')`
const field = (name) => `document.querySelector('[data-form-field="${name}"]')`
const bookmarkRows = `[...document.querySelectorAll('.bookmark-row')].map((row) => ({ label: row.querySelector('strong')?.textContent ?? row.querySelector('input')?.value, target: row.querySelector('small')?.textContent ?? null }))`
const pageSurface = (index) => `document.querySelector('.continuous-page-slot[data-page-index="${index}"] .page-surface')`
const editor = `document.querySelector('.inline-pdf-text-editor')`
const overlayWithText = (text) => `[...document.querySelectorAll('.text-overlay')].find((item) => item.textContent.includes(${JSON.stringify(text)}))`

async function formsAndBookmarks(page, target) {
  const results = {}
  await page.waitFor(`Boolean(${field('Address')}) && Boolean(${field('langs')})`, 'form controls', 60_000)

  // Multi-line text keeps its line breaks; MaxLen is enforced while typing.
  results.addressTag = await page.evaluate(`${field('Address')}.tagName`)
  assert.equal(results.addressTag, 'TEXTAREA', 'a multi-line field is edited in a textarea')
  await page.clickAt(field('Address'))
  await page.insertText('Line one')
  await page.key('Enter', 'Enter', 0, '\r')
  await page.insertText('Line two')
  await page.waitFor(`${field('Address')}.value === 'Line one\\nLine two'`, 'multi-line value')
  await page.clickAt(field('zip'))
  await page.insertText('1234567')
  await page.waitFor(`${field('zip')}.value === '12345'`, 'MaxLen enforced')
  // Radio group: the middle widget is "green".
  const radios = `[...document.querySelectorAll('[data-form-field="color"]')].sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left)`
  results.radioCount = await page.evaluate(`${radios}.length`)
  assert.equal(results.radioCount, 3, 'three radio buttons are shown')
  await page.clickAt(`${radios}[1]`)
  await page.waitFor(`${radios}.map((radio) => radio.checked).join() === 'false,true,false'`, 'green chosen')
  // Dropdown with export/display pairs and a multi-select list box.
  results.countryTag = await page.evaluate(`${field('country')}.tagName`)
  assert.equal(results.countryTag, 'SELECT')
  await page.evaluate(`(() => {
    const select = ${field('country')}
    select.focus()
    Object.getOwnPropertyDescriptor(HTMLSelectElement.prototype, 'value').set.call(select, 'CA')
    select.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await page.waitFor(`${field('country')}.value === 'CA'`, 'dropdown value')
  results.countryLabel = await page.evaluate(`${field('country')}.selectedOptions[0].textContent`)
  await page.evaluate(`(() => {
    const select = ${field('langs')}
    select.focus()
    for (const option of select.options) option.selected = option.value === 'ru' || option.value === 'uz'
    select.dispatchEvent(new Event('change', { bubbles: true }))
  })()`)
  await page.waitFor(`[...${field('langs')}.selectedOptions].map((option) => option.value).join() === 'ru,uz'`, 'list box values')
  await page.evaluate(`document.activeElement?.blur(); true`)

  // Bookmarks: every outline item is listed, headings and links included.
  await page.evaluate(`[...document.querySelectorAll('.sidebar [role="tab"]')].find((tab) => tab.textContent === 'Bookmarks').click()`)
  await page.waitFor(`document.querySelectorAll('.bookmark-row').length === 6`, 'six outline items')
  results.initialBookmarks = await page.evaluate(bookmarkRows)
  assert.deepEqual(results.initialBookmarks, [
    { label: 'Chapter 1', target: 'Page 2' }, { label: 'Website', target: 'Web link' }, { label: 'Part II', target: 'Heading' },
    { label: 'Chapter 2', target: 'Page 3' }, { label: 'Chapter 3', target: 'Page 4' }, { label: 'Notes', target: 'Page 5' },
  ])
  // Rename in place (Enter keeps it) …
  await page.evaluate(`document.querySelector('button[aria-label="Rename bookmark Chapter 2"]').click()`)
  await page.waitFor(`document.activeElement?.matches('input[aria-label="Bookmark name"]')`, 'rename field focused')
  await page.key('a', 'KeyA', 2)
  await page.insertText('Chapter Two')
  await page.key('Enter', 'Enter', 0, '\r')
  await page.waitFor(`!document.querySelector('input[aria-label="Bookmark name"]') && ${bookmarkRows}[3].label === 'Chapter Two'`, 'renamed bookmark')
  // … and Escape cancels a rename without touching the title.
  await page.evaluate(`document.querySelector('button[aria-label="Rename bookmark Notes"]').click()`)
  await page.waitFor(`document.activeElement?.matches('input[aria-label="Bookmark name"]')`, 'second rename field')
  await page.insertText('Changed')
  await page.key('Escape', 'Escape')
  await page.waitFor(`!document.querySelector('input[aria-label="Bookmark name"]') && ${bookmarkRows}[5].label === 'Notes'`, 'cancelled rename')
  // The toolbar button adds its own bookmark for the current page (page 1).
  await page.evaluate(`document.querySelector('button[aria-label="Bookmark this page"]').click()`)
  await page.waitFor(`document.querySelectorAll('.bookmark-row').length === 7`, 'toolbar bookmark')
  results.bookmarksAfterEdit = await page.evaluate(bookmarkRows)

  // Escape keeps a typed text box (it used to throw it away).
  await page.evaluate(`document.activeElement?.blur(); document.querySelector('button[aria-label="Add text (T)"]').click()`)
  await page.clickAt(pageSurface(0), 0.5, 0.86)
  await page.waitFor(`Boolean(${editor}) && document.activeElement === ${editor}`, 'new text box focused')
  await page.insertText('Kept by Escape')
  await page.key('Escape', 'Escape')
  await page.waitFor(`!${editor} && Boolean(${overlayWithText('Kept by Escape')})`, 'text kept after Escape')
  // The committed box starts its text at the top, as the editor and the
  // saved file do (a <button> used to centre it vertically).
  results.committedTextTopOffset = await page.evaluate(`(() => {
    const overlay = ${overlayWithText('Kept by Escape')}
    const lines = overlay.querySelector('.text-overlay-lines') || overlay
    return Math.round((lines.getBoundingClientRect().top - overlay.getBoundingClientRect().top) * 100) / 100
  })()`)
  assert.ok(Math.abs(results.committedTextTopOffset) < 1.5, `committed text is top-anchored (offset ${results.committedTextTopOffset}px)`)

  // A placed signature is a pending image edit: Escape keeps it too.
  await page.evaluate(`document.activeElement?.blur(); document.querySelector('button[aria-label="Sign document"]').click()`)
  const drawCanvas = `document.querySelector('canvas[aria-label="Draw your signature"]')`
  await page.waitFor(`Boolean(${drawCanvas})`, 'signature panel')
  const stroke = await page.evaluate(`(() => { const box = ${drawCanvas}.getBoundingClientRect(); return [[0.15, 0.6], [0.35, 0.3], [0.55, 0.7], [0.8, 0.35]].map(([x, y]) => ({ x: box.left + box.width * x, y: box.top + box.height * y })) })()`)
  await page.call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...stroke[0] })
  await page.call('Input.dispatchMouseEvent', { type: 'mousePressed', ...stroke[0], button: 'left', buttons: 1, clickCount: 1 })
  for (const point of stroke.slice(1)) await page.call('Input.dispatchMouseEvent', { type: 'mouseMoved', ...point, button: 'left', buttons: 1 })
  await page.call('Input.dispatchMouseEvent', { type: 'mouseReleased', ...stroke.at(-1), button: 'left', buttons: 0, clickCount: 1 })
  await page.evaluate(`[...document.querySelectorAll('.signature-dialog button')].find((button) => button.textContent.trim() === 'Use signature').click()`)
  await page.waitFor(`!${drawCanvas}`, 'signature chosen')
  await page.clickAt(pageSurface(0), 0.75, 0.62)
  await page.waitFor(`Boolean(document.querySelector('.object-edit-frame img'))`, 'signature placed')
  await page.evaluate(`document.activeElement?.blur(); true`)
  await page.key('Escape', 'Escape')
  await page.waitFor(`!document.querySelector('.object-edit-frame') && Boolean(document.querySelector('.object-overlay img'))`, 'signature kept after Escape')

  // Ctrl+Z after the Bold button undoes only the bold, not the whole box.
  await page.evaluate(`document.activeElement?.blur(); document.querySelector('button[aria-label="Add text (T)"]').click()`)
  await page.clickAt(pageSurface(0), 0.5, 0.93)
  await page.waitFor(`Boolean(${editor}) && document.activeElement === ${editor}`, 'second text box focused')
  await page.insertText('Bold then undo')
  await page.clickAt(`document.querySelector('.edit-inspector button[title="Bold"]')`)
  await page.waitFor(`document.querySelector('.edit-inspector button[title="Bold"]').getAttribute('aria-pressed') === 'true'`, 'bold on')
  await page.key('z', 'KeyZ', 2)
  await page.waitFor(`document.querySelector('.edit-inspector button[title="Bold"]')?.getAttribute('aria-pressed') === 'false' && ${editor}?.value === 'Bold then undo'`, 'only the bold undone')
  await page.key('Escape', 'Escape')
  await page.waitFor(`!${editor} && Boolean(${overlayWithText('Bold then undo')})`, 'second box kept')
  results.secondBoxWeight = await page.evaluate(`getComputedStyle(${overlayWithText('Bold then undo')}).fontWeight`)

  // Ctrl+S while still typing in a box applies that box and saves it, in
  // place (no dialog: the file has a path), keeping Undo and Redo.
  await page.evaluate(`document.activeElement?.blur(); document.querySelector('button[aria-label="Add text (T)"]').click()`)
  await page.clickAt(pageSurface(0), 0.5, 0.79)
  await page.waitFor(`Boolean(${editor}) && document.activeElement === ${editor}`, 'third text box focused')
  await page.insertText('Saved by Ctrl+S')
  await page.waitFor(dirtyDot, 'unsaved changes shown')
  await page.key('s', 'KeyS', 2)
  await page.waitFor(`!${editor} && Boolean(${overlayWithText('Saved by Ctrl+S')}) && !${dirtyDot} && ${settled} && !document.querySelector('.export-dialog')`, 'saved in place', 60_000)
  results.saveToast = await page.evaluate(`document.querySelector('.toast')?.textContent ?? ''`)
  results.undoAfterSave = await page.evaluate(`!document.querySelector('button[aria-label="Undo (Ctrl+Z)"]').disabled`)
  assert.ok(results.undoAfterSave, 'Undo stays available after saving')
  await page.key('z', 'KeyZ', 2)
  await page.waitFor(dirtyDot, 'undo past the save shows unsaved changes')
  await page.key('y', 'KeyY', 2)
  await page.waitFor(`!${dirtyDot}`, 'redo back to the saved state is clean')

  // Ctrl+S while renaming a bookmark keeps the new name and saves it.
  await page.evaluate(`document.querySelector('button[aria-label="Rename bookmark Chapter 3"]').click()`)
  await page.waitFor(`document.activeElement?.matches('input[aria-label="Bookmark name"]')`, 'third rename field')
  await page.insertText('Chapter Three')
  await page.key('s', 'KeyS', 2)
  await page.waitFor(`!document.querySelector('input[aria-label="Bookmark name"]') && ${bookmarkRows}[4].label === 'Chapter Three' && !${dirtyDot} && ${settled}`, 'rename saved with Ctrl+S', 60_000)

  // The saved file.
  const saved = await readFile(target)
  results.values = await savedFormValues(saved)
  assert.deepEqual(results.values, { zip: '12345', Address: 'Line one\nLine two', color: 'green', country: ['CA'], langs: ['ru', 'uz'] })
  results.outline = await savedOutline(saved)
  assert.deepEqual(results.outline.map(({ title, depth, target: destination }) => [title, depth, destination]), [
    ['Chapter 1', 0, 'page2 XYZ 72 700 2'],
    ['Website', 0, 'URI:https://example.com/guide'],
    ['Part II', 0, null],
    ['Chapter Two', 1, 'page3 XYZ 0 500 null'],
    ['Chapter Three', 1, 'page4 FitH 300'],
    ['Notes', 0, 'GoTo:page5 Fit'],
    ['Page 1', 0, 'page1 Fit'],
  ])
  assert.equal(results.outline[2].count, -2, 'the closed heading stays closed')
  results.page1Text = await pageText(saved, 1)
  assert.ok(results.page1Text.includes('Kept by Escape'), 'the box kept by Escape is in the saved file')
  assert.ok(results.page1Text.includes('Bold then undo'), 'the second box is in the saved file')
  assert.ok(results.page1Text.includes('Saved by Ctrl+S'), 'the box being typed in when Ctrl+S was pressed is in the saved file')
  results.page1Images = await imageDrawCount(saved, 1)
  assert.ok(results.page1Images >= 1, 'the signature kept by Escape is drawn in the saved file')
  return results
}

/** How many images page `pageNumber` paints (pdf.js operator list). */
async function imageDrawCount(bytes, pageNumber) {
  const document = await pdfjs.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise
  try {
    const operators = await (await document.getPage(pageNumber)).getOperatorList()
    const imageOps = new Set([pdfjs.OPS.paintImageXObject, pdfjs.OPS.paintInlineImageXObject, pdfjs.OPS.paintImageXObjectRepeat])
    return operators.fnArray.filter((op) => imageOps.has(op)).length
  } finally {
    await document.destroy()
  }
}

/** A signed document says how Save treats it: as a copy beside the original. */
async function signedNotice(page) {
  await page.waitFor(`[...document.querySelectorAll('[data-text-item="true"]')].some((item) => item.textContent.includes('Signed sample'))`, 'signed document', 60_000)
  const notice = await page.evaluate(`document.querySelector('.document-notice')?.textContent ?? ''`)
  assert.ok(notice.includes('Signed PDF') && notice.includes('copy'), `the signed document is explained: ${notice}`)
  return { notice }
}

/** Ctrl+S pressed while a page operation runs is saved afterwards, not swallowed. */
async function saveDuringOperation(page, target) {
  await page.waitFor(`[...document.querySelectorAll('[data-text-item="true"]')].some((item) => item.textContent.includes('Plain page 1'))`, 'plain document', 60_000)
  await page.evaluate(`document.querySelector('button[aria-label="More page actions"]').click()`)
  await page.waitFor(`Boolean([...document.querySelectorAll('.page-actions-popover button')].find((button) => button.textContent.includes('Insert blank page')))`, 'page actions')
  // Every toast shown from now on, in order.
  await page.evaluate(`(() => {
    window.__toasts = []
    new MutationObserver(() => {
      const text = document.querySelector('.toast > span')?.textContent
      if (text && window.__toasts.at(-1) !== text) window.__toasts.push(text)
    }).observe(document.body, { subtree: true, childList: true, characterData: true })
    return true
  })()`)
  // Press Ctrl+S from inside the page the moment the operation shows as
  // running: a MutationObserver sees the busy state in the same commit,
  // however quickly the operation then finishes.
  const pressedWhileBusy = await page.evaluate(`new Promise((resolve) => {
    const busyState = () => document.querySelector('.document-update-guard') || document.querySelector('.busy-overlay')
    const observer = new MutationObserver(() => {
      if (!busyState()) return
      observer.disconnect()
      clearTimeout(timer)
      const state = { busy: document.querySelector('.busy-overlay')?.textContent ?? '', guard: Boolean(document.querySelector('.document-update-guard')) }
      window.dispatchEvent(new KeyboardEvent('keydown', { key: 's', code: 'KeyS', ctrlKey: true, bubbles: true, cancelable: true }))
      resolve(state)
    })
    const timer = setTimeout(() => { observer.disconnect(); resolve(false) }, 15000)
    observer.observe(document.body, { subtree: true, childList: true })
    ;[...document.querySelectorAll('.page-actions-popover button')].find((button) => button.textContent.includes('Insert blank page')).click()
  })`)
  assert.ok(pressedWhileBusy && typeof pressedWhileBusy.busy === 'string', `the page operation was observed running: ${JSON.stringify(pressedWhileBusy)}`)
  try {
    await page.waitFor(`document.querySelector('.page-field span')?.textContent.trim() === '/ 5' && !${dirtyDot} && ${settled}`, 'saved after the operation', 60_000)
  } catch (error) {
    throw new Error(`${error.message}; pressed while: ${JSON.stringify(pressedWhileBusy)}; toasts: ${JSON.stringify(await page.evaluate('window.__toasts'))}; pages: ${await page.evaluate("document.querySelector('.page-field span')?.textContent")}`)
  }
  await page.waitFor(`window.__toasts.some((text) => text.startsWith('Saved'))`, 'save toast')
  const toasts = await page.evaluate(`window.__toasts`)
  const queued = toasts.indexOf('Will save when the current operation finishes')
  assert.ok(queued >= 0, `Ctrl+S during the operation is acknowledged: ${JSON.stringify(toasts)}`)
  assert.ok(toasts.findIndex((text) => text.startsWith('Saved')) > queued, `the save ran after the operation: ${JSON.stringify(toasts)}`)
  const saved = await PDFDocument.load(await readFile(target))
  assert.equal(saved.getPageCount(), 5, 'the queued save wrote the inserted page')
  return { toasts, pages: saved.getPageCount() }
}

async function passwordPrompt(page) {
  const dialog = `document.querySelector('form[aria-label="Password required"]')`
  const started = Date.now()
  await page.waitFor(`Boolean(${dialog})`, 'password prompt', 30_000)
  const shownAfter = Date.now() - started
  // What each submission sent, for the failure message.
  await page.evaluate(`(() => {
    window.__passwordSubmits = []
    window.__passwordEvents = []
    for (const type of ['keydown', 'input', 'focusin', 'focusout', 'submit']) {
      window.addEventListener(type, (event) => {
        window.__passwordEvents.push([type, event.key || event.data || '', event.defaultPrevented, event.target?.tagName, document.activeElement?.tagName, document.activeElement?.value, Boolean(document.querySelector('form[aria-label="Password required"]')?.textContent.includes('incorrect'))].join('|'))
      })
    }
    document.addEventListener('submit', (event) => {
      window.__passwordSubmits.push(event.target.querySelector('input[type="password"]')?.value ?? null)
    }, true)
    return true
  })()`)
  await page.waitFor(`document.activeElement?.matches('input[type="password"]')`, 'password field focused')
  await page.insertText('wrong')
  await page.waitFor(`document.activeElement?.value === 'wrong'`, 'wrong password typed')
  await page.key('Enter', 'Enter', 0, '\r')
  await page.waitFor(`${dialog}?.textContent.includes('That password is incorrect.') && document.activeElement?.matches('input[type="password"]')`, 'wrong password message')
  await page.typeText('secret')
  try {
    await page.waitFor(`document.activeElement?.value === 'secret'`, 'password typed')
  } catch (error) {
    throw new Error(`${error.message}; submitted: ${JSON.stringify(await page.evaluate('window.__passwordSubmits'))}; field: ${JSON.stringify(await page.evaluate('document.activeElement?.value'))}; events: ${JSON.stringify(await page.evaluate('window.__passwordEvents'))}`)
  }
  await page.key('Enter', 'Enter', 0, '\r')
  try {
    await page.waitFor(`!${dialog} && [...document.querySelectorAll('[data-text-item="true"]')].some((item) => item.textContent.includes('Protected sample'))`, 'unlocked document', 45_000)
  } catch (error) {
    throw new Error(`${error.message}; submitted: ${JSON.stringify(await page.evaluate('window.__passwordSubmits'))}`)
  }
  const notice = await page.evaluate(`document.querySelector('.document-notice')?.textContent ?? ''`)
  assert.ok(notice.includes('unprotected copy'), `the unlocked copy is explained: ${notice}`)
  const dirty = await page.evaluate(dirtyDot)
  assert.equal(dirty, false, 'unlocking alone is not an unsaved change')
  return { shownAfter, notice }
}

async function dropFromDisk(page, dropped) {
  await page.waitFor(`Boolean(document.querySelector('.drop-card'))`, 'home screen', 60_000)
  const point = await page.evaluate(`(() => { const box = document.querySelector('.drop-card').getBoundingClientRect(); return { x: box.left + box.width / 2, y: box.top + box.height / 2 } })()`)
  const data = { items: [], files: [dropped], dragOperationsMask: 1 }
  for (const type of ['dragEnter', 'dragOver', 'drop']) {
    await page.call('Input.dispatchDragEvent', { type, ...point, data })
  }
  await page.waitFor(`document.querySelector('.document-title')?.textContent === ${JSON.stringify(path.basename(dropped))}`, 'dropped document open', 45_000)
  const recent = await page.evaluate(`JSON.parse(localStorage.getItem('folio:recent-files:v1') || '[]').map((item) => item.path)`)
  assert.ok(recent.some((item) => item.toLowerCase() === dropped.toLowerCase()), `the dropped file keeps its path: ${JSON.stringify(recent)}`)
  return { recent }
}

// ---------------------------------------------------------------- run

const work = await mkdtemp(path.join(scratch, 'simple-forms-bookmarks-'))
const report = {}
try {
  if (only.has('forms')) {
    const target = path.join(work, 'forms-bookmarks.pdf')
    await writeFile(target, await formsAndOutlineFixture())
    report.forms = await withApp('forms', basePort, target, (page) => formsAndBookmarks(page, target))
  }
  if (only.has('password')) {
    const target = path.join(work, 'protected.pdf')
    await writeFile(target, await protectedFixture('secret'))
    report.password = await withApp('password', basePort + 1, target, passwordPrompt)
  }
  if (only.has('drop')) {
    const dropped = path.join(work, 'dropped-from-disk.pdf')
    await writeFile(dropped, await formsAndOutlineFixture())
    report.drop = await withApp('drop', basePort + 2, null, (page) => dropFromDisk(page, dropped))
  }
  if (only.has('signed')) {
    const target = path.join(work, 'signed.pdf')
    await writeFile(target, await plainFixture('Signed sample', 1, { signed: true }))
    report.signed = await withApp('signed', basePort + 3, target, signedNotice)
  }
  if (only.has('queue')) {
    const target = path.join(work, 'plain.pdf')
    await writeFile(target, await plainFixture('Plain', 4))
    report.queue = await withApp('queue', basePort + 4, target, (page) => saveDuringOperation(page, target))
  }
  console.log(JSON.stringify(report))
} finally {
  const resolved = path.resolve(work)
  assert.ok(path.basename(resolved).startsWith('simple-forms-bookmarks-'), 'Refusing to remove an unexpected temporary path')
  await rm(resolved, { recursive: true, force: true, maxRetries: 8, retryDelay: 250 })
}
