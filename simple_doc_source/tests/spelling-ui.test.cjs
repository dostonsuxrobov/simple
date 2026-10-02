'use strict'
// The spelling controller (src/ui/spelling.ts) with a fake engine bridge and dictionary:
// underlines for misspelled words, the word being typed left alone, Ignore of one
// occurrence, the switch (remembered) and Viewing mode. The right-click menu itself is
// covered by tests/word-features-smoke.mjs in Electron.
const test = require('node:test')
const assert = require('node:assert/strict')

const load = () => import('../src/ui/spelling.ts')

const style = { fontFamily: 'Calibri', fontSizePx: 14.667, bold: false, italic: false, underline: false, strikethrough: false, color: '#111111' }
const paragraph = (id, text) => ({ kind: 'paragraph', id, revision: 0, runs: [{ text, style }], style: { align: 'left' } })

function setup({ words = ['teh', 'jumpd', 'wrod'], data = new Map(), mode = 'edit' } = {}) {
  globalThis.window ??= globalThis
  const misspelled = new Set(words)
  const state = { doc: { blocks: [paragraph('a', 'Teh quick fox jumpd'), paragraph('b', 'A wrod and a wrod')] }, selection: null, mode, decorations: [], cleared: 0, lookups: 0 }
  const bridge = {
    getSelection: () => state.selection,
    setDecorations: (list) => { state.decorations = list; return true },
    clearDecorations: () => { state.decorations = []; state.cleared += 1; return true },
    positionFromPoint: () => null,
  }
  const handle = { getDocument: () => state.doc, getMode: () => state.mode }
  const api = {
    checkWords: async (list) => { state.lookups += list.length; return { source: 'windows', language: 'en-US', misspelled: list.map((word) => misspelled.has(word.toLowerCase())) } },
    getWordSuggestions: async () => [],
    getUserDictionary: async () => ({ words: [], ignored: [] }),
    getLanguages: async () => ({ available: true, languages: ['en-US'], preferred: 'en-US' }),
    onDictionaryChanged: () => () => {},
  }
  const host = { addEventListener() {}, removeEventListener() {} }
  const storage = { getItem: (key) => (data.has(key) ? data.get(key) : null), setItem: (key, value) => data.set(key, value) }
  return { state, bridge, handle, api, host, storage, data }
}

const marked = (state) => state.decorations.map((decoration) => `${decoration.range.anchor.blockId}:${decoration.range.anchor.offset}-${decoration.range.focus.offset}`)

test('misspelled words get red underlines; the word at the caret waits until the caret leaves it', async () => {
  const { createSpelling } = await load()
  const env = setup()
  const spelling = createSpelling({ ...env, notify() {}, focusEditor: () => true, replaceText: () => true })
  await spelling.checkNow()
  assert.deepEqual(marked(env.state), ['a:0-3', 'a:14-19', 'b:2-6', 'b:13-17'])
  assert.ok(env.state.decorations.every((decoration) => decoration.type === 'underline' && decoration.color === '#d93025' && !decoration.onClick), 'plain marks: a click places the caret as usual')
  const lookups = env.state.lookups
  env.state.selection = { anchor: { blockId: 'a', offset: 17 }, focus: { blockId: 'a', offset: 17 } }
  await spelling.checkNow()
  assert.deepEqual(marked(env.state), ['a:0-3', 'b:2-6', 'b:13-17'], 'jumpd is being typed')
  assert.equal(env.state.lookups, lookups, 'known words are not looked up again')
  // An edit re-checks only what changed.
  env.state.doc = { blocks: [env.state.doc.blocks[0], paragraph('b', 'A word and a wrod')] }
  env.state.selection = null
  await spelling.checkNow()
  assert.deepEqual(marked(env.state), ['a:0-3', 'a:14-19', 'b:13-17'])
  spelling.dispose()
})

test('Ignore hides one occurrence; the switch clears every mark, is remembered, and Viewing hides marks', async () => {
  const { createSpelling, occurrence, withoutIgnored } = await load()
  const env = setup()
  const changes = []
  const spelling = createSpelling({ ...env, notify() {}, focusEditor: () => true, replaceText: () => true, onStateChange: () => changes.push(spelling.enabled) })
  const issues = await spelling.checkNow()
  const second = issues.find((issue) => issue.blockId === 'b' && issue.start === 13)
  const ignored = new Set([occurrence(second, issues)])
  assert.deepEqual(withoutIgnored(issues, ignored).map((issue) => `${issue.blockId}:${issue.start}`), ['a:0', 'a:14', 'b:2'], 'only the second "wrod" is ignored')
  assert.equal(withoutIgnored(issues, new Set()), issues)

  spelling.setEnabled(false)
  assert.equal(spelling.enabled, false)
  assert.deepEqual(env.state.decorations, [])
  assert.equal(env.data.get('simple-docs:spelling'), 'off')
  assert.deepEqual(changes, [false])
  const again = createSpelling({ ...env, notify() {}, focusEditor: () => true, replaceText: () => true })
  assert.equal(again.enabled, false, 'a new window remembers the switch')
  again.dispose()
  spelling.setEnabled(true)
  assert.equal(env.data.get('simple-docs:spelling'), 'on')
  await spelling.checkNow()
  assert.equal(env.state.decorations.length, 4)
  env.state.mode = 'view'
  await spelling.checkNow()
  assert.deepEqual(env.state.decorations, [], 'no marks in Viewing mode')
  spelling.dispose()
})

test('without a dictionary on this computer nothing is marked and the switch says so', async () => {
  const { createSpelling } = await load()
  const env = setup()
  env.api.checkWords = async () => ({ source: null, misspelled: [] })
  const changes = []
  const spelling = createSpelling({ ...env, notify() {}, focusEditor: () => true, replaceText: () => true, onStateChange: () => changes.push(spelling.available) })
  await spelling.checkNow()
  assert.deepEqual(env.state.decorations, [])
  assert.equal(spelling.available, false)
  assert.deepEqual(changes, [false])
  const none = createSpelling({ ...env, api: null, notify() {}, focusEditor: () => true, replaceText: () => true })
  assert.equal(none.available, false, 'no preload bridge, no spelling')
  assert.deepEqual(await none.checkNow(), [])
  spelling.dispose()
  none.dispose()
})
