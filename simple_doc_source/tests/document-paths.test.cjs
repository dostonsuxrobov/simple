const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const os = require('node:os')
const path = require('node:path')
const crypto = require('node:crypto')
const { dialogDefaultPath, safeStem, sameFilePath, siblingCandidates } = require('../electron/document-paths.cjs')
const { documentShortcutAction, shortcutLetter } = require('../electron/shortcuts.cjs')
const { proposedSiblingPath, writeBesideSource } = require('../electron/sibling-save.cjs')

test('default names keep every character after a dot and strip only document extensions (doc-io-fidelity-11)', () => {
  assert.equal(safeStem('J. Smith CV'), 'J. Smith CV')
  assert.equal(safeStem('Mr. Brown'), 'Mr. Brown')
  assert.equal(safeStem('Budget v1.2'), 'Budget v1.2')
  assert.equal(safeStem('Q3 plan v2.1.docx'), 'Q3 plan v2.1')
  assert.equal(safeStem('C:\\Clients\\J. Smith CV.DOCX'), 'J. Smith CV')
  assert.equal(safeStem('/tmp/Legacy.doc'), 'Legacy')
  assert.equal(safeStem('notes.txt.docx'), 'notes.txt')
  assert.equal(safeStem('a<b>:c?.pdf'), 'a b c')
  assert.equal(safeStem('Trailing dots...'), 'Trailing dots')
  assert.equal(safeStem(''), 'Untitled document')
  assert.equal(safeStem(null), 'Untitled document')
  assert.equal(safeStem('CON'), 'CON document')
})

test('Save as and Export dialogs open beside the document when its folder is known', () => {
  const folder = path.join(os.tmpdir(), 'Clients')
  assert.equal(dialogDefaultPath({ name: 'J. Smith CV', extension: '.pdf', folderOf: path.join(folder, 'J. Smith CV.docx') }), path.join(folder, 'J. Smith CV.pdf'))
  assert.equal(dialogDefaultPath({ name: 'Q3 plan v2.1', extension: '.pdf', folderOf: [null, undefined, path.join(folder, 'Q3 plan v2.1.docx')] }), path.join(folder, 'Q3 plan v2.1.pdf'))
  // Without a known folder Windows chooses; the name is still complete.
  assert.equal(dialogDefaultPath({ name: 'Mr. Brown letter', extension: '.md', folderOf: [null] }), 'Mr. Brown letter.md')
  assert.equal(dialogDefaultPath({ name: 'Relative', extension: '.html', folderOf: 'relative\\x.docx' }), 'Relative.html')
})

test('sibling names never reuse the original and continue with "(edited)" numbering', () => {
  const source = path.join(os.tmpdir(), 'Contract 1.5.doc')
  const [first, second, third] = siblingCandidates(source, '.docx')
  assert.equal(first, path.join(os.tmpdir(), 'Contract 1.5.docx'))
  assert.equal(second, path.join(os.tmpdir(), 'Contract 1.5 (edited).docx'))
  assert.equal(third, path.join(os.tmpdir(), 'Contract 1.5 (edited 2).docx'))
  assert.equal(sameFilePath('C:\\A\\b.doc', 'c:/a/B.DOC'), true)
  assert.equal(sameFilePath('C:\\A\\b.doc', 'C:\\A\\b.docx'), false)
})

test('an edited legacy .doc saves beside the original without touching it (DOC-SIE-14)', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-doc-sibling-'))
  try {
    const original = path.join(directory, 'Old report.doc')
    const originalBytes = Buffer.concat([Buffer.from('d0cf11e0a1b11ae1', 'hex'), crypto.randomBytes(1024)])
    await fs.writeFile(original, originalBytes)
    const writes = []
    const write = async (target, data) => { writes.push(target); await fs.writeFile(target, data) }

    assert.equal(await proposedSiblingPath(original), path.join(directory, 'Old report.docx'))
    const first = await writeBesideSource(original, Buffer.from('first edit'), write)
    assert.equal(first, path.join(directory, 'Old report.docx'))
    assert.equal(await proposedSiblingPath(original), path.join(directory, 'Old report (edited).docx'))
    const second = await writeBesideSource(original, Buffer.from('second edit'), write)
    assert.equal(second, path.join(directory, 'Old report (edited).docx'))
    assert.equal(await fs.readFile(first, 'utf8'), 'first edit', 'an existing sibling is never overwritten')
    assert.equal(await fs.readFile(second, 'utf8'), 'second edit')
    assert.deepEqual(await fs.readFile(original), originalBytes, 'the original stays byte-identical')
    assert.ok(!writes.some((target) => sameFilePath(target, original)))

    // A failed write leaves no placeholder behind.
    await assert.rejects(writeBesideSource(original, Buffer.from('x'), async () => { throw new Error('disk full') }), /disk full/)
    assert.equal(await fs.access(path.join(directory, 'Old report (edited 2).docx')).then(() => true, () => false), false)
    assert.deepEqual(await fs.readFile(original), originalBytes)
  } finally {
    await fs.rm(directory, { recursive: true, force: true })
  }
})

const key = (key, code, extra = {}) => ({ type: 'keyDown', control: true, meta: false, alt: false, shift: false, key, code, ...extra })

test('document shortcuts follow the physical key on non-Latin layouts (doc-io-fidelity-10)', () => {
  // Latin layouts.
  assert.equal(documentShortcutAction(key('s', 'KeyS')), 'save')
  assert.equal(documentShortcutAction(key('S', 'KeyS', { shift: true })), 'save-as')
  assert.equal(documentShortcutAction(key('p', 'KeyP')), 'print')
  assert.equal(documentShortcutAction(key('E', 'KeyE', { shift: true })), 'export')
  assert.equal(documentShortcutAction(key('e', 'KeyE')), null)
  // Ctrl+O and Ctrl+N stay with the renderer on Latin layouts.
  assert.equal(documentShortcutAction(key('o', 'KeyO')), null)
  assert.equal(documentShortcutAction(key('n', 'KeyN')), null)
  // Russian/Uzbek Cyrillic, Greek, Hebrew and Arabic report their own characters.
  assert.equal(documentShortcutAction(key('ы', 'KeyS')), 'save')
  assert.equal(documentShortcutAction(key('Ы', 'KeyS', { shift: true })), 'save-as')
  assert.equal(documentShortcutAction(key('π', 'KeyP')), 'print')
  assert.equal(documentShortcutAction(key('ק', 'KeyE', { shift: true })), 'export')
  assert.equal(documentShortcutAction(key('خ', 'KeyO')), 'open')
  assert.equal(documentShortcutAction(key('т', 'KeyN')), 'new')
  // AZERTY/Dvorak keep the typed letter even where the physical key differs.
  assert.equal(shortcutLetter({ key: 's', code: 'Semicolon' }), 's')
  assert.equal(shortcutLetter({ key: 'a', code: 'KeyQ' }), 'a')
  assert.equal(documentShortcutAction(key('a', 'KeyQ')), null)
  // Not shortcuts: AltGr, key release, no modifier, unknown keys.
  assert.equal(documentShortcutAction(key('ы', 'KeyS', { alt: true })), null)
  assert.equal(documentShortcutAction(key('s', 'KeyS', { type: 'keyUp' })), null)
  assert.equal(documentShortcutAction(key('s', 'KeyS', { control: false })), null)
  assert.equal(documentShortcutAction(key('F5', 'F5')), null)
  assert.equal(documentShortcutAction({ ...key('s', 'KeyS'), control: false, meta: true }), 'save')
})
