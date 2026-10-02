'use strict'

/**
 * The Latin letter of a Ctrl/Cmd shortcut, independent of the keyboard layout.
 * A layout that types an ASCII letter keeps it (AZERTY, Dvorak). Cyrillic,
 * Greek, Hebrew, Arabic and other layouts report their own character as `key`,
 * so the physical key `code` (KeyS → "s") decides instead, as in Word.
 */
function shortcutLetter(input) {
  const key = String(input?.key ?? '')
  if (/^[a-z]$/i.test(key)) return key.toLowerCase()
  const code = /^Key([A-Z])$/.exec(String(input?.code ?? ''))
  return code ? code[1].toLowerCase() : null
}

function usesLatinKey(input) {
  return /^[a-z]$/i.test(String(input?.key ?? ''))
}

/**
 * Document shortcuts handled by the main process. Ctrl+P, Ctrl+S,
 * Ctrl+Shift+S and Ctrl+Shift+E are always routed through main. Ctrl+O and
 * Ctrl+N stay with the renderer's own handler on Latin layouts and are routed
 * here only when the layout reports a non-Latin character.
 */
function documentShortcutAction(input) {
  if (!input || input.type !== 'keyDown' || !(input.control || input.meta) || input.alt) return null
  const letter = shortcutLetter(input)
  if (letter === 'p') return 'print'
  if (letter === 's') return input.shift ? 'save-as' : 'save'
  if (letter === 'e' && input.shift) return 'export'
  if (!usesLatinKey(input) && !input.shift) {
    if (letter === 'o') return 'open'
    if (letter === 'n') return 'new'
  }
  return null
}

module.exports = { documentShortcutAction, shortcutLetter }
