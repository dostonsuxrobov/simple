const test = require('node:test')
const assert = require('node:assert/strict')

const keys = () => import('../src/ui/keys.ts')
const shortcuts = () => import('../src/ui/shortcuts.ts')

// Ribbon item ids of WordCanvas 0.12.0 inside Simple Docs (File tab removed), as
// listed by the running app. The ids derive from button labels, so an engine
// upgrade must re-verify them; scripts/run-editing-ui-smoke.mjs checks them live.
const VERIFIED_RIBBON_ITEMS = new Set([
  'home.font.grow-font', 'home.font.shrink-font', 'home.font.change-case', 'home.font.clear-all-formatting',
  'home.font.bold', 'home.font.italic', 'home.font.underline', 'home.font.superscript', 'home.font.subscript',
  'home.paragraph.show-hide-formatting-marks', 'home.paragraph.align-left', 'home.paragraph.center',
  'home.paragraph.align-right', 'home.paragraph.justify', 'home.paragraph.line-spacing',
  'home.paragraph.decrease-indent', 'home.paragraph.increase-indent',
  'home.styles.show-only-styles-in-use', 'home.editing.find-replace', 'home.editing.replace', 'home.editing.select-all',
  'insert.links.insert-remove-hyperlink', 'insert.references.insert-footnote', 'insert.references.insert-endnote',
])

// Chromium on Windows: `key` is what the layout types, `code` the physical key.
const press = (key, code, modifiers = {}) => ({ key, code, ctrlKey: false, altKey: false, shiftKey: false, metaKey: false, ...modifiers })
const ctrl = (key, code, extra = {}) => press(key, code, { ctrlKey: true, ...extra })

test('shortcut letters follow Latin layouts by character and other layouts by physical key', async () => {
  const { shortcutLetter, keyToken, usesLatinKey } = await keys()
  assert.equal(shortcutLetter(ctrl('k', 'KeyK')), 'k')
  assert.equal(shortcutLetter(ctrl('K', 'KeyK', { shiftKey: true })), 'k')
  assert.equal(shortcutLetter(ctrl('л', 'KeyK')), 'k', 'Russian Ctrl+Л is Ctrl+K')
  assert.equal(shortcutLetter(ctrl('ק', 'KeyE')), 'e', 'Hebrew')
  assert.equal(shortcutLetter(ctrl('a', 'KeyQ')), 'a', 'AZERTY keeps the typed letter')
  assert.equal(shortcutLetter(ctrl('1', 'Digit1')), null)
  assert.equal(usesLatinKey({ key: 'д' }), false)
  assert.equal(keyToken(ctrl('&', 'Digit1')), '1', 'AZERTY digit row')
  assert.equal(keyToken(ctrl('ъ', 'BracketRight')), ']')
  assert.equal(keyToken(ctrl('+', 'Equal', { shiftKey: true })), '=')
  assert.equal(keyToken(ctrl('ю', 'Period', { shiftKey: true })), '.')
  assert.equal(keyToken(ctrl(' ', 'Space')), 'Space')
  assert.equal(keyToken(press('F3', 'F3', { shiftKey: true })), 'F3')
  assert.equal(keyToken(press('й', 'KeyQ', { altKey: true })), 'q')
  assert.equal(keyToken(press('Tab', 'Tab')), 'Tab')
  assert.equal(keyToken(press('Dead', 'Quote')), "'")
})

test('combos parse, match exactly and print for tooltips', async () => {
  const { parseCombo, comboMatches, formatCombo, withShortcutHint } = await keys()
  assert.deepEqual(parseCombo('Ctrl+Shift+='), { ctrl: true, alt: false, shift: true, key: '=' })
  assert.deepEqual(parseCombo('Shift+F3'), { ctrl: false, alt: false, shift: true, key: 'F3' })
  assert.deepEqual(parseCombo('Ctrl+Space'), { ctrl: true, alt: false, shift: false, key: 'Space' })
  assert.equal(comboMatches(parseCombo('Ctrl+L'), ctrl('l', 'KeyL')), true)
  assert.equal(comboMatches(parseCombo('Ctrl+L'), ctrl('L', 'KeyL', { shiftKey: true })), false, 'extra Shift never matches')
  assert.equal(comboMatches(parseCombo('Ctrl+L'), ctrl('l', 'KeyL', { altKey: true })), false, 'extra Alt never matches')
  assert.equal(comboMatches(parseCombo('Ctrl+L'), press('l', 'KeyL', { metaKey: true })), true)
  assert.equal(formatCombo('ctrl+shift+n'), 'Ctrl+Shift+N')
  assert.equal(formatCombo('Ctrl+]'), 'Ctrl+]')
  assert.equal(withShortcutHint('Align left', 'Ctrl+L'), 'Align left (Ctrl+L)')
  assert.equal(withShortcutHint('Align left (Ctrl+L)', 'Ctrl+L'), 'Align left (Ctrl+L)')
  assert.equal(withShortcutHint('Replace (Ctrl+F)', 'Ctrl+H'), 'Replace (Ctrl+H)')
  assert.equal(withShortcutHint('Wrap text around image (square)', 'Ctrl+K'), 'Wrap text around image (square) (Ctrl+K)')
})

test('every Word and Docs shortcut maps to its command on a US layout', async () => {
  const { findShortcut } = await shortcuts()
  const cases = [
    [ctrl('k', 'KeyK'), 'link'], [ctrl('h', 'KeyH'), 'replace'], [ctrl('f', 'KeyF'), 'find'],
    [ctrl('l', 'KeyL'), 'align-left'], [ctrl('e', 'KeyE'), 'align-center'], [ctrl('r', 'KeyR'), 'align-right'], [ctrl('j', 'KeyJ'), 'justify'],
    [ctrl(']', 'BracketRight'), 'grow-font'], [ctrl('[', 'BracketLeft'), 'shrink-font'],
    [ctrl('>', 'Period', { shiftKey: true }), 'grow-font'], [ctrl('<', 'Comma', { shiftKey: true }), 'shrink-font'],
    [ctrl('=', 'Equal'), 'subscript'], [ctrl('+', 'Equal', { shiftKey: true }), 'superscript'],
    [ctrl('1', 'Digit1'), 'line-spacing-1'], [ctrl('2', 'Digit2'), 'line-spacing-2'], [ctrl('5', 'Digit5'), 'line-spacing-1.5'],
    [ctrl('1', 'Digit1', { altKey: true }), 'heading-1'], [ctrl('2', 'Digit2', { altKey: true }), 'heading-2'], [ctrl('3', 'Digit3', { altKey: true }), 'heading-3'],
    [ctrl('N', 'KeyN', { shiftKey: true }), 'normal-style'], [ctrl(' ', 'Space'), 'clear-character-formatting'],
    [ctrl('m', 'KeyM', { altKey: true }), 'comment'], [ctrl('f', 'KeyF', { altKey: true }), 'footnote'], [ctrl('d', 'KeyD', { altKey: true }), 'endnote'],
    [press('F3', 'F3', { shiftKey: true }), 'change-case'], [ctrl('*', 'Digit8', { shiftKey: true }), 'formatting-marks'],
    [press('F12', 'F12'), 'save-as'], [ctrl('w', 'KeyW'), 'close'], [press('q', 'KeyQ', { altKey: true }), 'command-search'], [press('/', 'Slash', { altKey: true }), 'command-search'],
  ]
  for (const [event, command] of cases) {
    assert.equal(findShortcut(event)?.command, command, `${event.ctrlKey ? 'Ctrl+' : ''}${event.altKey ? 'Alt+' : ''}${event.shiftKey ? 'Shift+' : ''}${event.code}`)
  }
  // The engine handles these itself on Latin layouts, and the app handles Ctrl+N/O/S/P.
  for (const event of [ctrl('z', 'KeyZ'), ctrl('y', 'KeyY'), ctrl('b', 'KeyB'), ctrl('a', 'KeyA'), ctrl('n', 'KeyN'), ctrl('s', 'KeyS'), press('l', 'KeyL'), ctrl('Enter', 'Enter'), ctrl('Backspace', 'Backspace')]) {
    assert.equal(findShortcut(event), null, `${event.code} is left alone`)
  }
})

test('shortcuts work on a Russian layout through physical keys', async () => {
  const { findShortcut } = await shortcuts()
  const cases = [
    [ctrl('л', 'KeyK'), 'link'], [ctrl('р', 'KeyH'), 'replace'], [ctrl('а', 'KeyF'), 'find'],
    [ctrl('д', 'KeyL'), 'align-left'], [ctrl('у', 'KeyE'), 'align-center'], [ctrl('к', 'KeyR'), 'align-right'], [ctrl('о', 'KeyJ'), 'justify'],
    [ctrl('ъ', 'BracketRight'), 'grow-font'], [ctrl('х', 'BracketLeft'), 'shrink-font'],
    [ctrl('=', 'Equal'), 'subscript'], [ctrl('+', 'Equal', { shiftKey: true }), 'superscript'],
    [ctrl('Т', 'KeyN', { shiftKey: true }), 'normal-style'], [ctrl('ь', 'KeyM', { altKey: true }), 'comment'],
    [ctrl('ц', 'KeyW'), 'close'], [press('й', 'KeyQ', { altKey: true }), 'command-search'],
    // Engine shortcuts that only read `key` are repeated for non-Latin layouts.
    [ctrl('я', 'KeyZ'), 'undo'], [ctrl('н', 'KeyY'), 'redo'], [ctrl('Я', 'KeyZ', { shiftKey: true }), 'redo'],
    [ctrl('и', 'KeyB'), 'bold'], [ctrl('ш', 'KeyI'), 'italic'], [ctrl('г', 'KeyU'), 'underline'], [ctrl('ф', 'KeyA'), 'select-all'],
  ]
  for (const [event, command] of cases) assert.equal(findShortcut(event)?.command, command, `${event.key} (${event.code})`)
})

test('AltGr characters and IME composition never trigger shortcuts', async () => {
  const { findShortcut } = await shortcuts()
  const altGraph = (event) => ({ ...event, getModifierState: (name) => name === 'AltGraph' })
  // Polish AltGr+M types µ-like characters; German AltGr+Q types @.
  assert.equal(findShortcut(altGraph(ctrl('µ', 'KeyM', { altKey: true }))), null)
  assert.equal(findShortcut(altGraph(ctrl('²', 'Digit2', { altKey: true }))), null)
  assert.equal(findShortcut(altGraph(ctrl('@', 'KeyQ', { altKey: true }))), null)
  assert.equal(findShortcut({ ...ctrl('k', 'KeyK'), isComposing: true }), null)
  assert.equal(findShortcut({ ...ctrl('Process', 'KeyK'), keyCode: 229 }), null)
})

test('every ribbon control a shortcut clicks exists in the verified WordCanvas 0.12.0 ribbon', async () => {
  const { SHORTCUT_RIBBON_ITEMS, SHORTCUTS, shortcutFor } = await shortcuts()
  const { RIBBON_ITEM_IDS } = await import('../src/engine-bridge.ts')
  for (const [command, id] of Object.entries(SHORTCUT_RIBBON_ITEMS)) {
    assert.ok(VERIFIED_RIBBON_ITEMS.has(id), `${command} → ${id}`)
  }
  // Shared ids agree with the engine bridge's pinned list.
  assert.equal(SHORTCUT_RIBBON_ITEMS.find, RIBBON_ITEM_IDS.findReplace)
  assert.equal(SHORTCUT_RIBBON_ITEMS.link, RIBBON_ITEM_IDS.hyperlink)
  assert.equal(SHORTCUT_RIBBON_ITEMS['align-center'], RIBBON_ITEM_IDS.alignCenter)
  assert.equal(SHORTCUT_RIBBON_ITEMS['formatting-marks'], RIBBON_ITEM_IDS.formattingMarks)
  // No two bindings share a combo.
  const combos = SHORTCUTS.map((binding) => binding.combo)
  assert.equal(new Set(combos).size, combos.length)
  assert.equal(shortcutFor('link'), 'Ctrl+K')
  assert.equal(shortcutFor('undo'), undefined, 'layout fallbacks are not advertised')
})

test('command search shows the shortcut of the ribbon control a shortcut clicks', async () => {
  const { ribbonItemShortcut } = await shortcuts()
  assert.equal(ribbonItemShortcut('home.paragraph.center'), 'Ctrl+E')
  assert.equal(ribbonItemShortcut('insert.references.insert-footnote'), 'Ctrl+Alt+F')
  assert.equal(ribbonItemShortcut('home.font.grow-font'), 'Ctrl+]')
  assert.equal(ribbonItemShortcut('home.paragraph.line-spacing'), undefined, 'three shortcuts share the menu')
  assert.equal(ribbonItemShortcut('home.font.clear-all-formatting'), undefined, 'Ctrl+Space keeps the paragraph style, the button does not')
  assert.equal(ribbonItemShortcut('home.font.bold'), undefined, 'the engine advertises its own Ctrl+B')
  assert.equal(ribbonItemShortcut('view.zoom.zoom-in'), undefined)
})
