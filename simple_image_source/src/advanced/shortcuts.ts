// src/advanced/shortcuts.ts (WP6)
// The Photoshop keyboard map of the Advanced editor (design 5.15) and chord normalisation.
//   - SHORTCUTS: every binding whose command is in the shared CommandId set (types.ts ShortcutBinding).
//   - EDITOR_SHORTCUTS: SHORTCUTS plus the editor-only bindings (tool opacity / flow digits, F6 / F7 panels),
//     which have no CommandId in the frozen contract (commands.ts EditorCommandId).
//   - chordFromEvent(): 'ctrl+alt+shift+key'. Brackets and digits use event.code (the physical key, so they
//     work on every layout); letters use event.key, falling back to the physical key on non-Latin layouts
//     (Cyrillic, Greek...); '+' drops Shift because most layouts need Shift to type it.
//   - Single-key bindings (no Ctrl, no Alt) are flagged requiresCanvasFocus: they never fire while a text
//     field, the text tool or an IME composition has focus. Tool switches ignore key repeat; [ and ] repeat.
//   - Hold-to-use keys (Space = Hand, Alt in paint tools = Eyedropper) are not chords: see SPRING_KEYS.
// Host-owned in Advanced (never bound here): Ctrl+O, Ctrl+S, Ctrl+Shift+S, Ctrl+P and Alt+Shift+Ctrl+W.
// Pure and DOM-free (Node tests import it).
import type { ShortcutBinding, ToolId } from './types.ts'
import type { EditorCommandId } from './commands.ts'

/** A binding of the editor: like ShortcutBinding, but the command may be an editor-only command. */
export interface EditorShortcut {
  /** Normalised chord: modifiers in the order ctrl+alt+shift, then the key, e.g. 'ctrl+shift+e', '['. */
  readonly chord: string
  readonly command: EditorCommandId
  /** Active only when no text input or text-tool session has focus (single-key shortcuts). */
  readonly requiresCanvasFocus: boolean
  /** Fire on key repeat ([ and ] do; tool switches do not). */
  readonly repeat: boolean
}

/** The fields chordFromEvent reads (a KeyboardEvent, or a plain object in tests). */
export interface ChordSource {
  readonly key: string
  readonly code?: string
  readonly ctrlKey?: boolean
  readonly metaKey?: boolean
  readonly altKey?: boolean
  readonly shiftKey?: boolean
  getModifierState?(key: string): boolean
}

// ---------------------------------------------------------------------------------------------
// Chords
// ---------------------------------------------------------------------------------------------

const MODIFIER_KEYS = new Set(['control', 'shift', 'alt', 'altgraph', 'meta', 'os', 'capslock', 'numlock', 'scrolllock', 'fn', 'fnlock', 'hyper', 'super', 'symbol', 'symbollock'])

const CODE_KEYS: Readonly<Record<string, string>> = Object.freeze({
  BracketLeft: '[',
  BracketRight: ']',
  NumpadAdd: '+',
  NumpadSubtract: '-',
  NumpadEnter: 'enter',
  NumpadDecimal: '.',
  Space: 'space',
})

const NAMED_KEYS: Readonly<Record<string, string>> = Object.freeze({
  ' ': 'space',
  Spacebar: 'space',
  Esc: 'escape',
  Del: 'delete',
  Left: 'arrowleft',
  Right: 'arrowright',
  Up: 'arrowup',
  Down: 'arrowdown',
})

/** The key part of a chord, or '' for a lone modifier / unknown key. */
function keyName(event: ChordSource): string {
  const code = typeof event.code === 'string' ? event.code : ''
  if (CODE_KEYS[code]) return CODE_KEYS[code]
  // Digits by physical key: Shift+1 types '!' on US layouts and '1' needs Shift on French layouts.
  const digit = /^(?:Digit|Numpad)([0-9])$/.exec(code)
  if (digit) return digit[1]
  const raw = typeof event.key === 'string' ? event.key : ''
  if (!raw || raw === 'Unidentified' || raw === 'Process') return ''
  if (NAMED_KEYS[raw]) return NAMED_KEYS[raw]
  const lower = raw.toLowerCase()
  if (MODIFIER_KEYS.has(lower)) return ''
  if (raw.length === 1) {
    if (/^[a-z]$/.test(lower)) return lower
    // Non-Latin letters and dead keys: use the physical key so B is still the Brush on a Cyrillic layout.
    const physical = /^Key([A-Z])$/.exec(code)
    if (physical && /\p{L}/u.test(raw)) return physical[1].toLowerCase()
    if (code === 'Equal' && (raw === '+' || raw === '=')) return raw
    if (code === 'Minus' && (raw === '-' || raw === '_')) return '-'
    return raw
  }
  if (raw === 'Dead') {
    const physical = /^Key([A-Z])$/.exec(code)
    return physical ? physical[1].toLowerCase() : ''
  }
  return lower
}

/**
 * Normalised chord of a key event: 'ctrl+alt+shift+key' (only the modifiers that are down; Meta counts as
 * Ctrl). Returns '' for a lone modifier, an unknown key, or AltGr text input (Ctrl+Alt on European layouts).
 */
export function chordFromEvent(event: ChordSource): string {
  if (typeof event.getModifierState === 'function') {
    try {
      if (event.getModifierState('AltGraph')) return ''
    } catch {
      // Older engines throw for unknown modifier names.
    }
  }
  const key = keyName(event)
  if (!key) return ''
  const parts: string[] = []
  if (event.ctrlKey || event.metaKey) parts.push('ctrl')
  if (event.altKey) parts.push('alt')
  // '+' needs Shift on most layouts, so Ctrl++ is the same chord whether or not Shift was involved.
  if (event.shiftKey && key !== '+') parts.push('shift')
  parts.push(key)
  return parts.join('+')
}

/** True for a chord without Ctrl or Alt (letters, digits, [ ], Delete, Enter, Esc, Tab, F-keys, Shift+key). */
export function isSingleKeyChord(chord: string): boolean {
  const parts = chord.split('+')
  // A trailing empty part means the key itself was '+' ('ctrl++').
  const modifiers = chord.endsWith('++') ? parts.slice(0, -2) : parts.slice(0, -1)
  return !modifiers.includes('ctrl') && !modifiers.includes('alt')
}

// ---------------------------------------------------------------------------------------------
// The map
// ---------------------------------------------------------------------------------------------

type Spec = readonly [chord: string, command: EditorCommandId, repeat?: boolean]

function binding([chord, command, repeat]: Spec): EditorShortcut {
  return Object.freeze({ chord, command, requiresCanvasFocus: isSingleKeyChord(chord), repeat: Boolean(repeat) })
}

/** Bindings of shared CommandIds (design 5.15). */
const COMMAND_SPECS: readonly Spec[] = [
  // Tools (letters select the tool shown in that toolbox slot; Shift cycles the slot)
  ['v', 'tool.move'],
  ['m', 'tool.marquee-rect'],
  ['shift+m', 'tool.cycle-marquee'],
  ['l', 'tool.lasso'],
  ['shift+l', 'tool.cycle-lasso'],
  ['w', 'tool.magic-wand'],
  ['c', 'tool.crop'],
  ['i', 'tool.eyedropper'],
  ['j', 'tool.spot-healing'],
  ['b', 'tool.brush'],
  ['s', 'tool.clone-stamp'],
  ['e', 'tool.eraser'],
  ['g', 'tool.gradient'],
  ['shift+g', 'tool.cycle-gradient'],
  ['t', 'tool.text'],
  ['u', 'tool.shape'],
  ['shift+u', 'tool.cycle-shape'],
  ['h', 'tool.hand'],
  ['z', 'tool.zoom'],
  // Colours
  ['x', 'edit.swap-colors'],
  ['d', 'edit.default-colors'],
  // Brush size and hardness
  ['[', 'brush.smaller', true],
  [']', 'brush.larger', true],
  ['shift+[', 'brush.softer', true],
  ['shift+]', 'brush.harder', true],
  // History
  ['ctrl+z', 'edit.undo', true],
  ['ctrl+shift+z', 'edit.redo', true],
  ['ctrl+y', 'edit.redo', true],
  ['ctrl+alt+z', 'edit.toggle-last'],
  // Layers
  ['ctrl+j', 'layer.via-copy'],
  ['ctrl+shift+j', 'layer.via-cut'],
  ['ctrl+shift+n', 'layer.new'],
  ['ctrl+e', 'layer.merge-down'],
  ['ctrl+shift+e', 'layer.merge-visible'],
  ['ctrl+alt+shift+e', 'layer.stamp-visible'],
  ['ctrl+alt+g', 'layer.toggle-clipping'],
  ['ctrl+[', 'layer.lower', true],
  ['ctrl+]', 'layer.raise', true],
  ['ctrl+shift+[', 'layer.to-back'],
  ['ctrl+shift+]', 'layer.to-front'],
  ['alt+[', 'layer.select-below', true],
  ['alt+]', 'layer.select-above', true],
  // Edit
  ['ctrl+t', 'edit.free-transform'],
  ['alt+backspace', 'edit.fill-foreground'],
  ['alt+delete', 'edit.fill-foreground'],
  ['ctrl+backspace', 'edit.fill-background'],
  ['ctrl+delete', 'edit.fill-background'],
  ['alt+shift+backspace', 'edit.fill-foreground-preserve'],
  ['alt+shift+delete', 'edit.fill-foreground-preserve'],
  ['delete', 'edit.clear'],
  ['backspace', 'edit.clear'],
  ['ctrl+c', 'edit.copy'],
  ['ctrl+shift+c', 'edit.copy-merged'],
  ['ctrl+x', 'edit.cut'],
  ['ctrl+v', 'edit.paste'],
  ['ctrl+shift+v', 'edit.paste-in-place'],
  // Select
  ['ctrl+a', 'select.all'],
  ['ctrl+d', 'select.deselect'],
  ['ctrl+shift+d', 'select.reselect'],
  ['ctrl+shift+i', 'select.inverse'],
  ['shift+f6', 'select.feather'],
  // Image > Adjustments
  ['ctrl+l', 'adjust.levels'],
  ['ctrl+m', 'adjust.curves'],
  ['ctrl+u', 'adjust.hue-saturation'],
  ['ctrl+b', 'adjust.color-balance'],
  ['ctrl+i', 'adjust.invert'],
  ['ctrl+shift+u', 'image.desaturate'],
  ['ctrl+alt+shift+b', 'adjust.black-white'],
  ['ctrl+shift+l', 'image.auto-tone'],
  ['ctrl+alt+shift+l', 'image.auto-contrast'],
  ['ctrl+shift+b', 'image.auto-color'],
  // Image
  ['ctrl+alt+i', 'image.size'],
  ['ctrl+alt+c', 'image.canvas-size'],
  // Filter
  ['ctrl+alt+f', 'filter.repeat'],
  // View
  ['ctrl+0', 'view.fit'],
  ['ctrl+1', 'view.actual-pixels'],
  ['ctrl++', 'view.zoom-in', true],
  ['ctrl+=', 'view.zoom-in', true],
  ['ctrl+-', 'view.zoom-out', true],
  ['tab', 'view.toggle-panels'],
  // Sessions (crop, free transform, polygonal lasso, text)
  ['enter', 'session.commit'],
  ['ctrl+enter', 'session.commit'],
  ['escape', 'session.cancel'],
]

/** Bindings of editor-only commands: tool opacity (0-9) and flow (Shift+0-9) digits, F7 Layers, F6 Color. */
const EDITOR_SPECS: readonly Spec[] = [
  ...Array.from({ length: 10 }, (_, digit): Spec => [String(digit), `tool.opacity-${digit}` as EditorCommandId]),
  ...Array.from({ length: 10 }, (_, digit): Spec => [`shift+${digit}`, `tool.flow-${digit}` as EditorCommandId]),
  ['f7', 'view.panel-layers'],
  ['f6', 'view.panel-color'],
]

/** Every binding whose command is a shared CommandId. */
export const SHORTCUTS: readonly ShortcutBinding[] = Object.freeze(COMMAND_SPECS.map(binding) as ShortcutBinding[])

/** Every binding the editor dispatches (SHORTCUTS first, then the editor-only digits and panel keys). */
export const EDITOR_SHORTCUTS: readonly EditorShortcut[] = Object.freeze([...SHORTCUTS, ...EDITOR_SPECS.map(binding)])

const BY_CHORD: ReadonlyMap<string, EditorShortcut> = new Map(EDITOR_SHORTCUTS.map((entry) => [entry.chord, entry]))

/** The binding for a chord, or null. */
export function shortcutForChord(chord: string): EditorShortcut | null {
  return chord ? BY_CHORD.get(chord) ?? null : null
}

/** The binding a key event triggers, or null. */
export function shortcutForEvent(event: ChordSource): EditorShortcut | null {
  return shortcutForChord(chordFromEvent(event))
}

// ---------------------------------------------------------------------------------------------
// Hold-to-use keys
// ---------------------------------------------------------------------------------------------

/**
 * Spring-loaded tools: while the key is held the tool changes, and releasing it returns to the previous
 * tool. Alt only springs from tools that use Alt-click for nothing else (Photoshop: the brush, gradient and
 * paint bucket sample a colour; the eraser, clone stamp and selection tools use Alt themselves).
 */
export const SPRING_KEYS: Readonly<Record<'space' | 'alt', { readonly tool: ToolId; readonly from: readonly ToolId[] | 'all' }>> = Object.freeze({
  space: Object.freeze({ tool: 'hand' as ToolId, from: 'all' as const }),
  alt: Object.freeze({ tool: 'eyedropper' as ToolId, from: Object.freeze(['brush', 'gradient', 'paint-bucket'] as ToolId[]) }),
})

/** The spring for a key event in `tool`, or null. Only bare Space / bare Alt spring. */
export function springForEvent(event: ChordSource, tool: ToolId): { readonly key: 'space' | 'alt'; readonly tool: ToolId } | null {
  const ctrl = Boolean(event.ctrlKey || event.metaKey)
  if (!ctrl && !event.altKey && !event.shiftKey && (event.key === ' ' || event.code === 'Space')) {
    return tool === SPRING_KEYS.space.tool ? null : { key: 'space', tool: SPRING_KEYS.space.tool }
  }
  if (!ctrl && !event.shiftKey && event.key === 'Alt') {
    const from = SPRING_KEYS.alt.from
    if (from !== 'all' && !from.includes(tool)) return null
    return { key: 'alt', tool: SPRING_KEYS.alt.tool }
  }
  return null
}

/** True when the key event releases a spring of `key`. */
export function releasesSpring(event: ChordSource, key: 'space' | 'alt'): boolean {
  return key === 'space' ? event.key === ' ' || event.code === 'Space' : event.key === 'Alt'
}

// ---------------------------------------------------------------------------------------------
// Labels (menus and tooltips use Photoshop's Windows notation: Alt+Shift+Ctrl+Key)
// ---------------------------------------------------------------------------------------------

const KEY_LABELS: Readonly<Record<string, string>> = Object.freeze({
  backspace: 'Backspace',
  delete: 'Delete',
  enter: 'Enter',
  escape: 'Esc',
  tab: 'Tab',
  space: 'Space',
  '=': '+',
  arrowleft: 'Left',
  arrowright: 'Right',
  arrowup: 'Up',
  arrowdown: 'Down',
})

/** 'ctrl+alt+shift+e' -> 'Alt+Shift+Ctrl+E'; 'shift+f6' -> 'Shift+F6'; 'ctrl++' -> 'Ctrl++'. */
export function chordLabel(chord: string): string {
  if (!chord) return ''
  const plusKey = chord.endsWith('++') || chord === '+'
  const parts = plusKey ? [...chord.slice(0, -1).split('+').filter(Boolean), '+'] : chord.split('+')
  const key = parts[parts.length - 1]
  const modifiers = new Set(parts.slice(0, -1))
  const label: string[] = []
  if (modifiers.has('alt')) label.push('Alt')
  if (modifiers.has('shift')) label.push('Shift')
  if (modifiers.has('ctrl')) label.push('Ctrl')
  label.push(KEY_LABELS[key] ?? (/^f\d{1,2}$/.test(key) ? key.toUpperCase() : key.length === 1 ? key.toUpperCase() : key))
  return label.join('+')
}

/** Preferred label of a command's shortcut (the first binding listed), or ''. */
export function shortcutLabel(command: EditorCommandId): string {
  const entry = EDITOR_SHORTCUTS.find((item) => item.command === command)
  return entry ? chordLabel(entry.chord) : ''
}

/** "Brush Tool (B)" style tooltip text. */
export function withShortcut(label: string, command: EditorCommandId): string {
  const keys = shortcutLabel(command)
  return keys ? `${label} (${keys})` : label
}
