'use strict'

const test = require('node:test')
const assert = require('node:assert/strict')
const { EventEmitter } = require('node:events')
const {
  contextMenuModel,
  installContextMenus,
  nativeTemplate,
  workspaceMode,
} = require('../electron/context-menu.cjs')

const docsUrl = 'file:///C:/simple/resources/app.asar/modules/docs/dist/index.html'
const calcUrl = 'file:///C:/simple/resources/app.asar/modules/calc/dist/index.html'

test('workspace mode detection is bounded to known module URLs', () => {
  assert.equal(workspaceMode(docsUrl), 'docs')
  assert.equal(workspaceMode('file:///C:/simple/modules/PDF/dist/index.html'), 'pdf')
  assert.equal(workspaceMode('https://example.test/modules/docs-ish/index.html'), null)
})

test('editable fields, native selections, links, and decoded images receive compact native menus', () => {
  assert.equal(contextMenuModel({ isEditable: true }, calcUrl).kind, 'editable')
  assert.deepEqual(contextMenuModel({ selectionText: 'selected text' }, docsUrl).actions, ['copy-selection'])
  assert.deepEqual(contextMenuModel({ linkURL: 'https://example.test/' }, docsUrl).actions, ['copy-link'])
  assert.deepEqual(contextMenuModel({ mediaType: 'image', hasImageContents: true }, 'file:///preview.html').actions, ['copy-image'])
  assert.equal(contextMenuModel({ mediaType: 'image', hasImageContents: false }, 'file:///preview.html'), null)
})

test('selected linked images retain every applicable copy action', () => {
  assert.deepEqual(contextMenuModel({
    selectionText: 'linked image caption',
    mediaType: 'image',
    hasImageContents: true,
    linkURL: 'https://example.test/image',
  }, docsUrl).actions, ['copy-selection', 'copy-image', 'copy-link'])
})

test('canvas and blank workspace areas defer to renderer-owned semantic menus', () => {
  assert.equal(contextMenuModel({}, calcUrl), null)
  assert.equal(contextMenuModel({}, docsUrl), null)
  assert.equal(contextMenuModel({}, 'file:///C:/simple/modules/pdf/dist/index.html'), null)
  assert.equal(contextMenuModel({}, 'file:///C:/simple/modules/image/dist/index.html'), null)
  assert.equal(contextMenuModel({}, 'file:///C:/simple/modules/video/dist/index.html'), null)
  assert.equal(contextMenuModel({}, 'file:///C:/simple/launcher/index.html'), null)
})

test('editable command availability follows Chromium edit flags', () => {
  const items = nativeTemplate({}, {
    isEditable: true,
    editFlags: { canUndo: true, canCopy: true, canSelectAll: true },
  }, { kind: 'editable' }, { writeText() {} })
  const byRole = Object.fromEntries(items.filter((item) => item.role).map((item) => [item.role, item.enabled]))
  assert.deepEqual(byRole, {
    undo: true,
    redo: false,
    cut: false,
    copy: true,
    paste: false,
    delete: false,
    selectAll: true,
  })
})

test('composed copy callbacks preserve exact image coordinates and link address', () => {
  const copied = []
  const contents = { copyImageAt: (x, y) => copied.push(['image', x, y]) }
  const clipboard = { writeText: (value) => copied.push(['link', value]) }
  const params = { x: 41, y: 73, linkURL: 'https://example.test/exact' }
  const model = { kind: 'content', actions: ['copy-image', 'copy-link'] }
  const items = nativeTemplate(contents, params, model, clipboard)
  items.find((item) => item.label === 'Copy image').click()
  items.find((item) => item.label === 'Copy link address').click()
  assert.deepEqual(copied, [['image', 41, 73], ['link', 'https://example.test/exact']])
})

test('installation is idempotent and popup remains owned by the originating frame and window', () => {
  const app = new EventEmitter()
  const contents = new EventEmitter()
  contents.getURL = () => docsUrl
  contents.isDestroyed = () => false
  const owner = { isDestroyed: () => false }
  const frame = { id: 'frame-1' }
  let template = null
  let popupOptions = null
  const Menu = {
    buildFromTemplate(value) {
      template = value
      return { popup(options) { popupOptions = options } }
    },
  }
  const BrowserWindow = { fromWebContents: (value) => value === contents ? owner : null }
  const clipboard = { writeText() {} }

  installContextMenus({ app, Menu, BrowserWindow, clipboard })
  installContextMenus({ app, Menu, BrowserWindow, clipboard })
  assert.equal(app.listenerCount('web-contents-created'), 1)
  app.emit('web-contents-created', {}, contents)
  contents.emit('context-menu', {}, {
    selectionText: 'copy me', frameURL: docsUrl, frame, x: 120, y: 80, menuSourceType: 'mouse',
  })
  assert.equal(template[0].role, 'copy')
  assert.deepEqual(popupOptions, { window: owner, frame, x: 120, y: 80, sourceType: 'mouse' })
})
