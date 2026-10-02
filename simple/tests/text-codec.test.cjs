'use strict'

// Text decoding and dialect-preserving encoding (design §6.8).

const test = require('node:test')
const assert = require('node:assert/strict')
const {
  WRITABLE_ENCODINGS,
  canEncode,
  decodeText,
  detectLineEnding,
  encodeText,
  labelForCodePage,
  normalizeEncoding,
  normalizeLineEndings,
  systemAnsiCodePage,
} = require('../shared/electron/text-codec.cjs')

const hex = (value) => Buffer.from(value.replace(/\s+/g, ''), 'hex')
const SAMPLE = 'Name;Amount\r\nCafé;3\r\nNaïve;5\r\n'

// Hand-encoded fixtures, so the decoder is not checked against itself.
const CP1252 = { bytes: hex('436166e9 2063 72e8 6d65 2096 2093 6f6b 94'), text: 'Café crème – “ok”' }
const CP1251 = { bytes: hex('cff0e8e2e5f2 2c20 ece8f0 2120 fdf2ee 20 f2e5f1f2 2e'), text: 'Привет, мир! это тест.' }
const SHIFT_JIS = { bytes: hex('93fa 967b 8cea 82cc 8365 834c 8358 8367 82c5 82b7 8142'), text: '日本語のテキストです。' }
const GBK = { bytes: hex('d6d0 cec4 cec4 b1be b2e2 cad4'), text: '中文文本测试' }
const EUC_KR = { bytes: hex('c7d1 b1b9 beee 20 c5d7 bdba c6ae'), text: '한국어 테스트' }

test('byte order marks decide the encoding and are reported', () => {
  const utf8 = decodeText(Buffer.concat([hex('efbbbf'), Buffer.from(SAMPLE)]))
  assert.deepEqual([utf8.text, utf8.encoding, utf8.bom, utf8.confidence], [SAMPLE, 'utf-8', true, 'certain'])

  const le = decodeText(Buffer.concat([hex('fffe'), Buffer.from(SAMPLE, 'utf16le')]))
  assert.deepEqual([le.text, le.encoding, le.bom], [SAMPLE, 'utf-16le', true])

  const beBody = Buffer.from(SAMPLE, 'utf16le').swap16()
  const be = decodeText(Buffer.concat([hex('feff'), beBody]))
  assert.deepEqual([be.text, be.encoding, be.bom], [SAMPLE, 'utf-16be', true])
})

test('UTF-16 without a byte order mark is found by NUL parity', () => {
  const le = decodeText(Buffer.from(SAMPLE, 'utf16le'))
  assert.deepEqual([le.text, le.encoding, le.bom, le.confidence], [SAMPLE, 'utf-16le', false, 'high'])
  const be = decodeText(Buffer.from(SAMPLE, 'utf16le').swap16())
  assert.deepEqual([be.text, be.encoding, be.bom], [SAMPLE, 'utf-16be', false])
  // Ordinary text never looks like UTF-16.
  assert.equal(decodeText(Buffer.from('plain ascii text')).encoding, 'utf-8')
})

test('strict UTF-8 wins over the ANSI code page', () => {
  const text = 'Ünïcödé — Привет — 日本 — 🙂'
  const result = decodeText(Buffer.from(text, 'utf8'), { ansiCodePage: 1252 })
  assert.deepEqual([result.text, result.encoding, result.bom, result.confidence, result.replacements], [text, 'utf-8', false, 'certain', 0])
  const ascii = decodeText(Buffer.from('a,b\n1,2\n'))
  assert.deepEqual([ascii.encoding, ascii.ascii, ascii.confidence], ['utf-8', true, 'high'])
})

test('invalid UTF-8 falls back to Windows-1252, or Windows-1251 for Cyrillic', () => {
  const western = decodeText(CP1252.bytes, { ansiCodePage: 1252 })
  assert.deepEqual([western.text, western.encoding], [CP1252.text, 'windows-1252'])

  const cyrillic = decodeText(CP1251.bytes, { ansiCodePage: 1252 })
  assert.deepEqual([cyrillic.text, cyrillic.encoding, cyrillic.confidence], [CP1251.text, 'windows-1251', 'medium'])

  // Western text with runs of accented letters stays Western.
  const german = decodeText(Buffer.from('Größe, Übergröße, Straße: schön', 'latin1'), { ansiCodePage: 1252 })
  assert.equal(german.encoding, 'windows-1252')
  // A Cyrillic system code page is used directly.
  assert.equal(decodeText(CP1251.bytes, { ansiCodePage: 1251 }).encoding, 'windows-1251')
  assert.equal(decodeText(Buffer.from([0x41, 0xe9, 0x42]), { ansiCodePage: 1250 }).encoding, 'windows-1250')
})

test('Shift-JIS and GBK are recognised when they are the only plausible reading', () => {
  const japanese = decodeText(SHIFT_JIS.bytes, { ansiCodePage: 1252 })
  assert.deepEqual([japanese.text, japanese.encoding], [SHIFT_JIS.text, 'shift_jis'])
  const chinese = decodeText(GBK.bytes, { ansiCodePage: 1252 })
  assert.deepEqual([chinese.text, chinese.encoding], [GBK.text, 'gbk'])
  const korean = decodeText(EUC_KR.bytes, { ansiCodePage: 1252 })
  assert.deepEqual([korean.text, korean.encoding], [EUC_KR.text, 'euc-kr'])
  // Western accented letters next to ASCII letters are valid GBK pairs, but not plausible ones.
  for (const western of ['Größe, Übergröße, Straße: schön', 'Le café est très bon. Où est la bibliothèque?']) {
    assert.equal(decodeText(Buffer.from(western, 'latin1'), { ansiCodePage: 1252 }).encoding, 'windows-1252', western)
  }
  // Turning the East Asian guess off keeps the ANSI reading.
  assert.equal(decodeText(SHIFT_JIS.bytes, { ansiCodePage: 1252, cjk: false }).encoding, 'windows-1252')
  // On a Japanese system the code page itself is Shift-JIS.
  assert.equal(decodeText(SHIFT_JIS.bytes, { ansiCodePage: 932, cjk: false }).encoding, 'shift_jis')
})

test('a forced encoding is honoured and a matching BOM is still stripped', () => {
  const forced = decodeText(CP1251.bytes, { encoding: 'cp1251' })
  assert.deepEqual([forced.text, forced.encoding, forced.confidence], [CP1251.text, 'windows-1251', 'certain'])
  const bom = decodeText(Buffer.concat([hex('efbbbf'), Buffer.from('x')]), { encoding: 'UTF8' })
  assert.deepEqual([bom.text, bom.bom], ['x', true])
})

test('a truncated sample ignores a cut multi-byte character at the end', () => {
  const bytes = Buffer.from('Привет', 'utf8')
  const cut = bytes.subarray(0, bytes.length - 1)
  assert.equal(decodeText(cut, { truncated: true }).encoding, 'utf-8')
  assert.equal(decodeText(cut, { truncated: true }).text, 'Приве')
  assert.notEqual(decodeText(cut, { ansiCodePage: 1252 }).encoding, 'utf-8', 'A complete file with a broken sequence is not UTF-8')
})

test('line endings are detected, including mixed files', () => {
  assert.deepEqual(detectLineEnding('a\r\nb\r\nc'), { lineEnding: 'crlf', mixedLineEndings: false })
  assert.deepEqual(detectLineEnding('a\nb\nc\r\n'), { lineEnding: 'lf', mixedLineEndings: true })
  assert.deepEqual(detectLineEnding('a\rb'), { lineEnding: 'cr', mixedLineEndings: false })
  assert.deepEqual(detectLineEnding('one line'), { lineEnding: null, mixedLineEndings: false })
  assert.equal(normalizeLineEndings('a\nb\r\nc\rd', 'crlf'), 'a\r\nb\r\nc\r\nd')
  assert.equal(normalizeLineEndings('a\r\nb', null), 'a\r\nb')
})

test('dialect round trip: decode then encode gives back the same bytes', () => {
  const fixtures = [
    Buffer.concat([hex('efbbbf'), Buffer.from(SAMPLE)]),
    Buffer.from(SAMPLE.replace(/\r\n/g, '\n')),
    Buffer.concat([hex('fffe'), Buffer.from(SAMPLE, 'utf16le')]),
    Buffer.concat([hex('feff'), Buffer.from(SAMPLE, 'utf16le').swap16()]),
    Buffer.from(SAMPLE, 'utf16le'),
    Buffer.from(SAMPLE, 'utf16le').swap16(),
    Buffer.concat([CP1252.bytes, Buffer.from('\r\n')]),
    Buffer.concat([CP1251.bytes, Buffer.from('\r\n'), CP1251.bytes]),
    SHIFT_JIS.bytes,
    GBK.bytes,
    EUC_KR.bytes,
  ]
  for (const bytes of fixtures) {
    const decoded = decodeText(bytes, { ansiCodePage: 1252 })
    const encoded = encodeText(decoded.text, decoded.dialect)
    assert.equal(encoded.unmappable, 0, decoded.encoding)
    assert.ok(encoded.bytes.equals(bytes), `${decoded.encoding}${decoded.bom ? ' with BOM' : ''} round trip`)
  }
})

test('edited text keeps the file dialect: encoding, BOM and line endings', () => {
  const original = decodeText(Buffer.concat([CP1251.bytes, Buffer.from('\r\n')]), { ansiCodePage: 1252 })
  const edited = `${original.text}Новая строка\n`
  const { bytes } = encodeText(edited, original.dialect)
  assert.equal(decodeText(bytes, { encoding: 'windows-1251' }).text, `${CP1251.text}\r\nНовая строка\r\n`)
  const utf16 = encodeText('a\nb', { encoding: 'utf-16le', bom: true, lineEnding: 'crlf' })
  assert.ok(utf16.bytes.equals(Buffer.concat([hex('fffe'), Buffer.from('a\r\nb', 'utf16le')])))
})

test('characters a code page cannot hold are counted, replaced, or refused', () => {
  const result = encodeText('Price: 5 € — 🙂 Ж', { encoding: 'windows-1252' })
  assert.equal(result.unmappable, 2)
  assert.deepEqual(result.unmappableSamples, ['🙂', 'Ж'])
  assert.equal(decodeText(result.bytes, { encoding: 'windows-1252' }).text, 'Price: 5 € — ? ?')
  assert.throws(() => encodeText('Ж', { encoding: 'windows-1252', onUnmappable: 'throw' }), (error) => error.code === 'UNMAPPABLE_CHARACTERS' && error.unmappable === 1)
  assert.deepEqual(canEncode('Ж', 'windows-1251'), { ok: true, unmappable: 0, unmappableSamples: [] })
  assert.equal(canEncode('Ж', 'windows-1252').ok, false)
  assert.equal(canEncode('🙂', 'utf-8').ok, true)
  assert.equal(encodeText(SHIFT_JIS.text, { encoding: 'shift_jis' }).bytes.equals(SHIFT_JIS.bytes), true)
})

test('encoding names are normalised and every writable encoding encodes', () => {
  assert.equal(normalizeEncoding('UTF8'), 'utf-8')
  assert.equal(normalizeEncoding('cp1251'), 'windows-1251')
  assert.equal(normalizeEncoding('Shift-JIS'), 'shift_jis')
  assert.equal(normalizeEncoding('latin1'), 'windows-1252')
  assert.equal(normalizeEncoding('GB2312'), 'gbk')
  assert.throws(() => normalizeEncoding('no-such-encoding'), RangeError)
  assert.equal(labelForCodePage(1251), 'windows-1251')
  assert.equal(labelForCodePage(65001), 'windows-1252')
  assert.equal(labelForCodePage(12345), 'windows-1252')
  for (const { id } of WRITABLE_ENCODINGS) {
    const { bytes, unmappable } = encodeText('abc 123\n', { encoding: id })
    assert.equal(unmappable, 0, id)
    assert.equal(decodeText(bytes, { encoding: id }).text, 'abc 123\n', id)
  }
})

test('the ANSI code page can be pinned for tests and is read once otherwise', () => {
  const previous = process.env.SIMPLE_ANSI_CODEPAGE
  try {
    process.env.SIMPLE_ANSI_CODEPAGE = '1251'
    assert.equal(systemAnsiCodePage(), 1251)
    assert.equal(decodeText(CP1251.bytes).encoding, 'windows-1251')
  } finally {
    if (previous === undefined) delete process.env.SIMPLE_ANSI_CODEPAGE
    else process.env.SIMPLE_ANSI_CODEPAGE = previous
  }
  const page = systemAnsiCodePage()
  assert.ok(Number.isInteger(page) && page > 0)
  assert.equal(systemAnsiCodePage(), page)
})
