'use strict'
// WP6 shortcuts (src/advanced/shortcuts.ts): the Photoshop keyboard map of design 5.15.
//   - no chord is bound twice; every design shortcut is present with the right command;
//   - single-key shortcuts (no Ctrl, no Alt) are flagged requiresCanvasFocus, the others are not;
//   - tool switches ignore key repeat, [ / ] repeat; host-owned file shortcuts are never bound;
//   - chord normalisation is layout-aware (brackets and digits by physical key, letters by key with a
//     physical fallback for non-Latin layouts, '+' without Shift, Meta = Ctrl, AltGr ignored);
//   - spring-loaded keys (hold Space = Hand, hold Alt = Eyedropper in paint tools);
//   - every command a shortcut or menu item names exists in the command registry.
const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')

const root = path.join(__dirname, '..')
const STRIP_HELP = `Node ${process.versions.node} is not stripping TypeScript types, so the .ts sources these tests import cannot load. `
  + 'Use Node 22.18+, 23.6+ or 24+ with built-in type stripping (not disabled by --no-experimental-strip-types), '
  + 'or run node with --experimental-strip-types.'

function load(relative) {
  if (!(process.features && process.features.typescript)) throw new Error(STRIP_HELP)
  return require(path.join(root, 'src', relative))
}

const shortcuts = load('advanced/shortcuts.ts')
const commands = load('advanced/commands.ts')

const key = (name, extra = {}) => ({ key: name, code: '', ctrlKey: false, metaKey: false, altKey: false, shiftKey: false, ...extra })

// Design 5.15, chord -> command. Digits, F6 and F7 are editor-only commands (no CommandId in types.ts).
const DESIGN = [
  ['v', 'tool.move'], ['m', 'tool.marquee-rect'], ['l', 'tool.lasso'], ['w', 'tool.magic-wand'], ['c', 'tool.crop'],
  ['i', 'tool.eyedropper'], ['j', 'tool.spot-healing'], ['b', 'tool.brush'], ['s', 'tool.clone-stamp'], ['e', 'tool.eraser'],
  ['g', 'tool.gradient'], ['t', 'tool.text'], ['u', 'tool.shape'], ['h', 'tool.hand'], ['z', 'tool.zoom'],
  ['shift+m', 'tool.cycle-marquee'], ['shift+l', 'tool.cycle-lasso'], ['shift+g', 'tool.cycle-gradient'], ['shift+u', 'tool.cycle-shape'],
  ['x', 'edit.swap-colors'], ['d', 'edit.default-colors'],
  ['[', 'brush.smaller'], [']', 'brush.larger'], ['shift+[', 'brush.softer'], ['shift+]', 'brush.harder'],
  ...Array.from({ length: 10 }, (_, digit) => [String(digit), `tool.opacity-${digit}`]),
  ['ctrl+z', 'edit.undo'], ['ctrl+shift+z', 'edit.redo'], ['ctrl+y', 'edit.redo'], ['ctrl+alt+z', 'edit.toggle-last'],
  ['ctrl+j', 'layer.via-copy'], ['ctrl+shift+j', 'layer.via-cut'], ['ctrl+shift+n', 'layer.new'],
  ['ctrl+e', 'layer.merge-down'], ['ctrl+shift+e', 'layer.merge-visible'], ['ctrl+alt+shift+e', 'layer.stamp-visible'], ['ctrl+alt+g', 'layer.toggle-clipping'],
  ['ctrl+t', 'edit.free-transform'], ['ctrl+a', 'select.all'], ['ctrl+d', 'select.deselect'], ['ctrl+shift+d', 'select.reselect'],
  ['ctrl+shift+i', 'select.inverse'], ['shift+f6', 'select.feather'],
  ['alt+backspace', 'edit.fill-foreground'], ['ctrl+backspace', 'edit.fill-background'], ['alt+shift+backspace', 'edit.fill-foreground-preserve'],
  ['delete', 'edit.clear'], ['backspace', 'edit.clear'],
  ['ctrl+l', 'adjust.levels'], ['ctrl+m', 'adjust.curves'], ['ctrl+u', 'adjust.hue-saturation'], ['ctrl+b', 'adjust.color-balance'], ['ctrl+i', 'adjust.invert'],
  ['ctrl+shift+u', 'image.desaturate'], ['ctrl+alt+shift+b', 'adjust.black-white'],
  ['ctrl+shift+l', 'image.auto-tone'], ['ctrl+alt+shift+l', 'image.auto-contrast'], ['ctrl+shift+b', 'image.auto-color'],
  ['ctrl+alt+i', 'image.size'], ['ctrl+alt+c', 'image.canvas-size'], ['ctrl+alt+f', 'filter.repeat'],
  ['ctrl+c', 'edit.copy'], ['ctrl+shift+c', 'edit.copy-merged'], ['ctrl+x', 'edit.cut'], ['ctrl+v', 'edit.paste'], ['ctrl+shift+v', 'edit.paste-in-place'],
  ['ctrl+0', 'view.fit'], ['ctrl+1', 'view.actual-pixels'], ['ctrl++', 'view.zoom-in'], ['ctrl+=', 'view.zoom-in'], ['ctrl+-', 'view.zoom-out'],
  ['ctrl+[', 'layer.lower'], ['ctrl+]', 'layer.raise'], ['ctrl+shift+[', 'layer.to-back'], ['ctrl+shift+]', 'layer.to-front'],
  ['alt+[', 'layer.select-below'], ['alt+]', 'layer.select-above'],
  ['enter', 'session.commit'], ['ctrl+enter', 'session.commit'], ['escape', 'session.cancel'],
  ['tab', 'view.toggle-panels'], ['f7', 'view.panel-layers'], ['f6', 'view.panel-color'],
]

test('no chord is bound twice', () => {
  const seen = new Map()
  for (const binding of shortcuts.EDITOR_SHORTCUTS) {
    assert.ok(binding.chord, `a binding of ${binding.command} has no chord`)
    assert.ok(!seen.has(binding.chord), `"${binding.chord}" is bound to both ${seen.get(binding.chord)} and ${binding.command}`)
    seen.set(binding.chord, binding.command)
  }
  assert.equal(new Set(shortcuts.SHORTCUTS.map((binding) => binding.chord)).size, shortcuts.SHORTCUTS.length)
})

test('every shortcut of design 5.15 is present with its command', () => {
  for (const [chord, command] of DESIGN) {
    const binding = shortcuts.shortcutForChord(chord)
    assert.ok(binding, `missing shortcut ${chord} (${command})`)
    assert.equal(binding.command, command, `${chord} runs ${binding.command}, expected ${command}`)
  }
  // Space (hold) is the spring-loaded Hand, not a chord.
  assert.equal(shortcuts.SPRING_KEYS.space.tool, 'hand')
  assert.equal(shortcuts.SPRING_KEYS.space.from, 'all')
})

test('single-key shortcuts need canvas focus; Ctrl and Alt chords do not', () => {
  for (const binding of shortcuts.EDITOR_SHORTCUTS) {
    const parts = binding.chord.endsWith('++') ? [...binding.chord.slice(0, -2).split('+'), '+'] : binding.chord.split('+')
    const modifiers = parts.slice(0, -1)
    const single = !modifiers.includes('ctrl') && !modifiers.includes('alt')
    assert.equal(binding.requiresCanvasFocus, single, `${binding.chord}: requiresCanvasFocus should be ${single}`)
  }
  for (const chord of ['v', 'b', 'x', '[', 'shift+]', '5', 'delete', 'backspace', 'enter', 'escape', 'tab', 'f7', 'shift+f6', 'shift+m']) {
    assert.equal(shortcuts.shortcutForChord(chord).requiresCanvasFocus, true, chord)
  }
  for (const chord of ['ctrl+z', 'ctrl+j', 'alt+backspace', 'alt+[', 'ctrl+alt+shift+e', 'ctrl++']) {
    assert.equal(shortcuts.shortcutForChord(chord).requiresCanvasFocus, false, chord)
  }
})

test('tool switches ignore key repeat; brush size, undo and zoom repeat', () => {
  for (const binding of shortcuts.EDITOR_SHORTCUTS) {
    if (/^tool\.(?!opacity|flow)/.test(binding.command)) assert.equal(binding.repeat, false, `${binding.chord} (${binding.command}) must not repeat`)
  }
  for (const chord of ['[', ']', 'shift+[', 'shift+]', 'ctrl+z', 'ctrl+shift+z', 'ctrl+y', 'ctrl++', 'ctrl+-']) {
    assert.equal(shortcuts.shortcutForChord(chord).repeat, true, `${chord} repeats`)
  }
  for (const chord of ['ctrl+j', 'ctrl+e', 'ctrl+t', 'x', 'enter']) assert.equal(shortcuts.shortcutForChord(chord).repeat, false, chord)
})

test('host file shortcuts stay with the host', () => {
  for (const chord of ['ctrl+o', 'ctrl+s', 'ctrl+shift+s', 'ctrl+p', 'ctrl+alt+shift+w']) {
    assert.equal(shortcuts.shortcutForChord(chord), null, `${chord} belongs to the host`)
  }
})

test('every bound command exists; contract bindings use contract command ids only', () => {
  for (const binding of shortcuts.EDITOR_SHORTCUTS) assert.ok(commands.isEditorCommand(binding.command), `${binding.chord} -> unknown ${binding.command}`)
  const editorOnly = /^(tool\.(opacity|flow)-\d|view\.panel-|filter\.sharpen-more)/
  for (const binding of shortcuts.SHORTCUTS) assert.ok(!editorOnly.test(binding.command), `${binding.command} is editor-only and must not be in SHORTCUTS`)
  assert.ok(shortcuts.EDITOR_SHORTCUTS.length > shortcuts.SHORTCUTS.length)
  for (const binding of shortcuts.EDITOR_SHORTCUTS) {
    assert.ok(Object.isFrozen(binding), 'bindings are frozen')
    assert.equal(typeof binding.repeat, 'boolean')
  }
})

test('chordFromEvent normalises modifiers and keys', () => {
  const chord = shortcuts.chordFromEvent
  assert.equal(chord(key('b', { code: 'KeyB' })), 'b')
  assert.equal(chord(key('B', { code: 'KeyB', shiftKey: true })), 'shift+b')
  assert.equal(chord(key('z', { code: 'KeyZ', ctrlKey: true })), 'ctrl+z')
  assert.equal(chord(key('Z', { code: 'KeyZ', ctrlKey: true, shiftKey: true })), 'ctrl+shift+z')
  assert.equal(chord(key('z', { code: 'KeyZ', metaKey: true })), 'ctrl+z', 'Meta counts as Ctrl')
  assert.equal(chord(key('E', { code: 'KeyE', ctrlKey: true, altKey: true, shiftKey: true })), 'ctrl+alt+shift+e', 'modifier order')
  assert.equal(chord(key('Backspace', { code: 'Backspace', altKey: true })), 'alt+backspace')
  assert.equal(chord(key('Delete', { code: 'Delete' })), 'delete')
  assert.equal(chord(key('Escape', { code: 'Escape' })), 'escape')
  assert.equal(chord(key('Enter', { code: 'NumpadEnter' })), 'enter')
  assert.equal(chord(key('Tab', { code: 'Tab' })), 'tab')
  assert.equal(chord(key('F6', { code: 'F6', shiftKey: true })), 'shift+f6')
  assert.equal(chord(key(' ', { code: 'Space' })), 'space')
  // Brackets and digits by physical key.
  assert.equal(chord(key('[', { code: 'BracketLeft' })), '[')
  assert.equal(chord(key('{', { code: 'BracketLeft', shiftKey: true })), 'shift+[')
  assert.equal(chord(key('ü', { code: 'BracketLeft' })), '[', 'German layout: the [ key types ü')
  assert.equal(chord(key('1', { code: 'Digit1' })), '1')
  assert.equal(chord(key('!', { code: 'Digit1', shiftKey: true })), 'shift+1')
  assert.equal(chord(key('&', { code: 'Digit1' })), '1', 'French layout: digits need Shift but are still digits')
  assert.equal(chord(key('5', { code: 'Numpad5' })), '5')
  assert.equal(chord(key('0', { code: 'Digit0', ctrlKey: true })), 'ctrl+0')
  // Zoom keys.
  assert.equal(chord(key('=', { code: 'Equal', ctrlKey: true })), 'ctrl+=')
  assert.equal(chord(key('+', { code: 'Equal', ctrlKey: true, shiftKey: true })), 'ctrl++', '+ drops Shift')
  assert.equal(chord(key('+', { code: 'NumpadAdd', ctrlKey: true })), 'ctrl++')
  assert.equal(chord(key('-', { code: 'Minus', ctrlKey: true })), 'ctrl+-')
  assert.equal(chord(key('-', { code: 'NumpadSubtract', ctrlKey: true })), 'ctrl+-')
  // Non-Latin layouts fall back to the physical key.
  assert.equal(chord(key('и', { code: 'KeyB' })), 'b', 'Cyrillic: the B key is still the Brush')
  assert.equal(chord(key('я', { code: 'KeyZ', ctrlKey: true })), 'ctrl+z')
  // Nothing for lone modifiers, unknown keys and AltGr text.
  assert.equal(chord(key('Control', { code: 'ControlLeft', ctrlKey: true })), '')
  assert.equal(chord(key('Shift', { code: 'ShiftLeft', shiftKey: true })), '')
  assert.equal(chord(key('Alt', { code: 'AltLeft', altKey: true })), '')
  assert.equal(chord(key('Unidentified')), '')
  assert.equal(chord(key('Process', { code: 'KeyA' })), '', 'IME processing key')
  const altGr = { ...key('€', { code: 'KeyE', ctrlKey: true, altKey: true }), getModifierState: (name) => name === 'AltGraph' }
  assert.equal(chord(altGr), '', 'AltGr typing is text, not a shortcut')
})

test('shortcutForEvent finds the binding of a key event', () => {
  assert.equal(shortcuts.shortcutForEvent(key('j', { code: 'KeyJ', ctrlKey: true })).command, 'layer.via-copy')
  assert.equal(shortcuts.shortcutForEvent(key(']', { code: 'BracketRight' })).command, 'brush.larger')
  assert.equal(shortcuts.shortcutForEvent(key('q', { code: 'KeyQ' })), null)
})

test('spring-loaded keys: hold Space for the Hand, Alt for the Eyedropper in paint tools', () => {
  const spring = shortcuts.springForEvent
  assert.deepEqual(spring(key(' ', { code: 'Space' }), 'brush'), { key: 'space', tool: 'hand' })
  assert.deepEqual(spring(key(' ', { code: 'Space' }), 'marquee-rect'), { key: 'space', tool: 'hand' })
  assert.equal(spring(key(' ', { code: 'Space' }), 'hand'), null, 'already the Hand')
  assert.equal(spring(key(' ', { code: 'Space', ctrlKey: true }), 'brush'), null, 'Ctrl+Space is not the Hand')
  assert.deepEqual(spring(key('Alt', { code: 'AltLeft', altKey: true }), 'brush'), { key: 'alt', tool: 'eyedropper' })
  assert.deepEqual(spring(key('Alt', { code: 'AltLeft', altKey: true }), 'gradient'), { key: 'alt', tool: 'eyedropper' })
  assert.deepEqual(spring(key('Alt', { code: 'AltLeft', altKey: true }), 'paint-bucket'), { key: 'alt', tool: 'eyedropper' })
  for (const tool of ['eraser', 'clone-stamp', 'marquee-rect', 'move', 'spot-healing']) {
    assert.equal(spring(key('Alt', { code: 'AltLeft', altKey: true }), tool), null, `Alt is ${tool}'s own modifier`)
  }
  assert.equal(spring(key('Alt', { code: 'AltLeft', altKey: true, ctrlKey: true }), 'brush'), null)
  assert.equal(shortcuts.releasesSpring(key(' ', { code: 'Space' }), 'space'), true)
  assert.equal(shortcuts.releasesSpring(key('Alt', { code: 'AltLeft' }), 'alt'), true)
  assert.equal(shortcuts.releasesSpring(key('b'), 'space'), false)
})

test('labels use Photoshop notation', () => {
  assert.equal(shortcuts.chordLabel('ctrl+alt+shift+e'), 'Alt+Shift+Ctrl+E')
  assert.equal(shortcuts.chordLabel('ctrl+shift+n'), 'Shift+Ctrl+N')
  assert.equal(shortcuts.chordLabel('alt+backspace'), 'Alt+Backspace')
  assert.equal(shortcuts.chordLabel('shift+f6'), 'Shift+F6')
  assert.equal(shortcuts.chordLabel('ctrl++'), 'Ctrl++')
  assert.equal(shortcuts.chordLabel('ctrl+='), 'Ctrl++')
  assert.equal(shortcuts.chordLabel('['), '[')
  assert.equal(shortcuts.chordLabel('escape'), 'Esc')
  assert.equal(shortcuts.shortcutLabel('layer.new'), 'Shift+Ctrl+N')
  assert.equal(shortcuts.shortcutLabel('view.zoom-in'), 'Ctrl++')
  assert.equal(shortcuts.shortcutLabel('layer.flatten'), '', 'no shortcut')
  assert.equal(shortcuts.withShortcut('Brush Tool', 'tool.brush'), 'Brush Tool (B)')
  assert.equal(shortcuts.withShortcut('Flatten Image', 'layer.flatten'), 'Flatten Image')
})

test('every command a menu names is registered', () => {
  const source = fs.readFileSync(path.join(root, 'src', 'advanced', 'menus', 'MenuBar.tsx'), 'utf8')
  const named = [...source.matchAll(/\b(?:cmd|check)\('([a-z-]+\.[a-z0-9-]+)'/g)].map((match) => match[1])
  named.push(...[...source.matchAll(/\bfilter\('([a-z-]+)'\)/g)].map((match) => `filter.${match[1]}`))
  assert.ok(named.length > 60, `found ${named.length} menu commands`)
  for (const command of named) assert.ok(commands.isEditorCommand(command), `menu item ${command} is not a command`)
  // The adjustment submenus are generated from these types.
  const listed = /const ADJUSTMENTS[^=]*=\s*\[([\s\S]*?)\]/.exec(source)
  assert.ok(listed, 'adjustment list found')
  for (const match of listed[1].matchAll(/'([a-z-]+)'/g)) {
    if (match[1] === '-') continue
    assert.ok(commands.isEditorCommand(`adjust.${match[1]}`), `adjust.${match[1]}`)
    assert.ok(commands.isEditorCommand(`adjustment-layer.${match[1]}`), `adjustment-layer.${match[1]}`)
  }
})
