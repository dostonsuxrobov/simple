const test = require('node:test')
const assert = require('node:assert/strict')

const search = () => import('../src/ui/command-search.ts')
const author = () => import('../src/ui/author.ts')

const command = (id, label, extra = {}) => ({ id, label, enabled: true, run() {}, ...extra })

test('ribbon tooltips split into a command name, its shortcut and a hint', async () => {
  const { describeRibbonTitle } = await search()
  assert.deepEqual(describeRibbonTitle('Bold (Ctrl+B)'), { label: 'Bold', shortcut: 'Ctrl+B' })
  assert.deepEqual(describeRibbonTitle('Wrap text around image (square) — select an image first'), { label: 'Wrap text around image (square)', detail: 'select an image first' })
  assert.deepEqual(describeRibbonTitle('Grow font (Ctrl+])'), { label: 'Grow font', shortcut: 'Ctrl+]' })
  assert.deepEqual(describeRibbonTitle('Numbered list (Tab/Shift+Tab change level)'), { label: 'Numbered list (Tab/Shift+Tab change level)' })
  assert.deepEqual(describeRibbonTitle('  Insert   table '), { label: 'Insert table' })
})

test('command search ranks exact names, then prefixes, words, keywords and loose matches', async () => {
  const { rankCommands } = await search()
  const commands = [
    command('a', 'Center'),
    command('b', 'Insert table'),
    command('c', 'Insert image from your device'),
    command('d', 'Table of contents', { keywords: ['toc'] }),
    command('e', 'Insert link', { keywords: ['hyperlink', 'url'] }),
    command('f', 'Heading 1'),
    command('g', 'Suggesting (track changes)', { keywords: ['track changes', 'revisions'] }),
    command('h', 'Line spacing 1.5'),
  ]
  assert.equal(rankCommands('center', commands)[0].id, 'a')
  assert.equal(rankCommands('Cen', commands)[0].id, 'a')
  assert.deepEqual(rankCommands('table', commands).map((entry) => entry.id).slice(0, 2), ['d', 'b'], 'name prefix first, then word prefix')
  assert.equal(rankCommands('ins tab', commands)[0].id, 'b', 'several word prefixes')
  assert.equal(rankCommands('hyperlink', commands)[0].id, 'e', 'keywords')
  assert.equal(rankCommands('track changes', commands)[0].id, 'g')
  assert.equal(rankCommands('hdng', commands)[0].id, 'f', 'loose letter match')
  assert.equal(rankCommands('spacing 1.5', commands)[0].id, 'h')
  assert.deepEqual(rankCommands('zzzz', commands), [])
  assert.equal(rankCommands('é', [command('x', 'Résumé')])[0]?.id, 'x', 'accents are ignored')
})

test('recent commands come first, and disabled ones sink but stay findable', async () => {
  const { rankCommands, rememberRecent } = await search()
  const commands = [command('one', 'Insert table'), command('two', 'Insert link'), command('three', 'Insert footnote', { enabled: false, hint: 'select text first' })]
  assert.deepEqual(rankCommands('insert', commands, ['two']).map((entry) => entry.id), ['two', 'one', 'three'])
  assert.deepEqual(rankCommands('', commands, ['three', 'one'], { suggestedIds: ['two', 'one'] }).map((entry) => entry.id), ['three', 'one', 'two'], 'empty query: recent, then suggested')
  assert.equal(rankCommands('', commands, [], { limit: 1, suggestedIds: ['two', 'one'] }).length, 1)
  let recent = []
  for (const id of ['a', 'b', 'a', 'c', 'd', 'e', 'f', 'g', 'h', 'i']) recent = rememberRecent(recent, id)
  assert.deepEqual(recent, ['i', 'h', 'g', 'f', 'e', 'd', 'c', 'a'])
})

test('comment authors use the chosen name, else the Windows user name', async () => {
  const { authorIdentity, authorDisplayName, cleanAuthorName, windowsUserDisplayName, resolveAuthorName, readStoredAuthorName, storeAuthorName, AUTHOR_STORAGE_KEY } = await author()
  assert.deepEqual(authorIdentity('Doston Sanatovich'), { id: 'local-user', firstName: 'Doston', lastName: 'Sanatovich' })
  assert.deepEqual(authorIdentity('kms_b'), { id: 'local-user', firstName: 'kms_b', lastName: '' })
  assert.deepEqual(authorIdentity('  Ana   María López '), { id: 'local-user', firstName: 'Ana', lastName: 'María López' })
  assert.equal(authorDisplayName({ firstName: 'Ana', lastName: 'María López' }), 'Ana María López')
  assert.equal(cleanAuthorName('a\u0000b\n c'), 'a b c')
  assert.equal(cleanAuthorName('x'.repeat(200)).length, 80)
  assert.equal(windowsUserDisplayName('CONTOSO\\jdoe'), 'jdoe')
  assert.equal(windowsUserDisplayName('jdoe@contoso.com'), 'jdoe')
  assert.deepEqual(resolveAuthorName(null, 'kms_b'), { name: 'kms_b', chosen: false })
  assert.deepEqual(resolveAuthorName('Doston', 'kms_b'), { name: 'Doston', chosen: true })
  assert.deepEqual(resolveAuthorName('', ''), { name: 'Author', chosen: false })
  const values = new Map()
  const storage = { getItem: (key) => values.get(key) ?? null, setItem: (key, value) => values.set(key, value), removeItem: (key) => values.delete(key) }
  assert.equal(readStoredAuthorName(storage), null)
  assert.equal(storeAuthorName(storage, '  Doston  Sanatovich '), true)
  assert.equal(values.get(AUTHOR_STORAGE_KEY), 'Doston Sanatovich')
  assert.equal(readStoredAuthorName(storage), 'Doston Sanatovich')
  storeAuthorName(storage, '')
  assert.equal(readStoredAuthorName(storage), null)
  const broken = { getItem() { throw new Error('denied') }, setItem() { throw new Error('denied') }, removeItem() {} }
  assert.equal(readStoredAuthorName(broken), null)
  assert.equal(storeAuthorName(broken, 'x'), false)
})

test('the Windows user name is read from the user profile the app runs in, never a shared profile', async () => {
  const { windowsUserFromAppUrl } = await author()
  assert.equal(windowsUserFromAppUrl('file:///C:/Users/kms_b/Documents/simple/simple_doc_source/dist/index.html'), 'kms_b')
  assert.equal(windowsUserFromAppUrl('file:///C:/Users/Ana%20Mar%C3%ADa/AppData/Local/Temp/2x/resources/app.asar/dist/index.html'), 'Ana María')
  assert.equal(windowsUserFromAppUrl('file:///c:/users/doston/AppData/Local/Temp/simple/index.html'), 'doston')
  assert.equal(windowsUserFromAppUrl('file:///D:/Apps/Simple/resources/app.asar/dist/index.html'), null)
  assert.equal(windowsUserFromAppUrl('file:///C:/Users/Public/Simple/index.html'), null)
  assert.equal(windowsUserFromAppUrl('file:///C:/Users/Default/index.html'), null)
  assert.equal(windowsUserFromAppUrl('http://127.0.0.1:5173/'), null)
  assert.equal(windowsUserFromAppUrl('file:///C:/Users/%E0%A4%A/x.html'), null, 'a malformed path is ignored')
})
