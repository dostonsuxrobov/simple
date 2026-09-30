'use strict'

function workspaceMode(url) {
  const match = /[\\/]modules[\\/](docs|calc|pdf|image|video)[\\/]/i.exec(String(url || ''))
  return match?.[1].toLowerCase() || null
}

function contextMenuModel(params = {}, url = '') {
  const mode = workspaceMode(url)
  if (params.isEditable) return { kind: 'editable', mode, commands: [] }
  const actions = []
  if (String(params.selectionText || '').trim()) actions.push('copy-selection')
  if (params.mediaType === 'image' && params.hasImageContents) actions.push('copy-image')
  if (String(params.linkURL || '')) actions.push('copy-link')
  if (actions.length) return { kind: 'content', mode, actions }
  // Canvas workspaces own semantic menus in their renderers. Falling through is
  // intentional: it prevents blank chrome, modal backdrops, and print previews
  // from exposing enabled commands that cannot act on anything.
  return null
}

function nativeTemplate(contents, params, model, clipboard) {
  const flags = params.editFlags || {}
  if (model.kind === 'editable') {
    return [
      { role: 'undo', enabled: Boolean(flags.canUndo) },
      { role: 'redo', enabled: Boolean(flags.canRedo) },
      { type: 'separator' },
      { role: 'cut', enabled: Boolean(flags.canCut) },
      { role: 'copy', enabled: Boolean(flags.canCopy) },
      { role: 'paste', enabled: Boolean(flags.canPaste) },
      { role: 'delete', enabled: Boolean(flags.canDelete) },
      { type: 'separator' },
      { role: 'selectAll', enabled: Boolean(flags.canSelectAll) },
    ]
  }
  if (model.kind === 'content') {
    const items = []
    for (const action of model.actions) {
      if (action === 'copy-selection') items.push({ role: 'copy' })
      else if (action === 'copy-image') items.push({ label: 'Copy image', click: () => contents.copyImageAt(params.x, params.y) })
      else if (action === 'copy-link') items.push({ label: 'Copy link address', click: () => clipboard.writeText(params.linkURL) })
    }
    if (model.actions.includes('copy-selection')) items.push({ type: 'separator' }, { role: 'selectAll' })
    return items
  }
  return []
}

function installContextMenus({ app, Menu, BrowserWindow, clipboard }) {
  const installMarker = Symbol.for('simple.native-context-menus-installed')
  if (app[installMarker]) return
  Object.defineProperty(app, installMarker, { value: true })
  app.on('web-contents-created', (_event, contents) => {
    contents.on('context-menu', (_contextEvent, params) => {
      const model = contextMenuModel(params, params.frameURL || contents.getURL())
      if (!model) return
      const template = nativeTemplate(contents, params, model, clipboard)
      if (!template.length || contents.isDestroyed()) return
      const owner = BrowserWindow.fromWebContents(contents)
      if (!owner || owner.isDestroyed()) return
      Menu.buildFromTemplate(template).popup({
        window: owner,
        ...(params.frame ? { frame: params.frame } : {}),
        x: params.x,
        y: params.y,
        ...(params.menuSourceType ? { sourceType: params.menuSourceType } : {}),
      })
    })
  })
}

module.exports = {
  contextMenuModel,
  installContextMenus,
  nativeTemplate,
  workspaceMode,
}
