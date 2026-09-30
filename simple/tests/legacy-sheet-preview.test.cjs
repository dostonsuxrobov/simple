'use strict'
const test = require('node:test')
const assert = require('node:assert/strict')
const XLSX = require('../../simple_calc_source/node_modules/xlsx')
const { prepareLegacySheetPreview } = require('../launcher/legacy-sheet-preview.cjs')

function record(type, data = Buffer.alloc(0)) {
  const head = Buffer.alloc(4); head.writeUInt16LE(type); head.writeUInt16LE(data.length, 2)
  return Buffer.concat([head, data])
}
function bof(type) { const data = Buffer.alloc(16); data.writeUInt16LE(0x0600); data.writeUInt16LE(type, 2); return record(0x0809, data) }
function header(text, wide = false, type = 0x0014) {
  const data = Buffer.alloc(3); data.writeUInt16LE(text.length); data[2] = Number(wide)
  return record(type, Buffer.concat([data, Buffer.from(text, wide ? 'utf16le' : 'latin1')]))
}
function fixture({ text = '&L&K000000Title 000000', wide = false, encrypted = false, malformed = false, flags, mini = false } = {}) {
  const parts = [bof(5), ...(encrypted ? [record(0x002f, Buffer.from([1, 0]))] : []), record(0x000a), bof(0x0010)]
  const offset = Buffer.concat(parts).length
  const h = header(text, wide)
  if (malformed) h.writeUInt16LE(0xffff, 4)
  if (flags != null) h[6] = flags
  parts.push(h, header('&R&K000000Page &P', wide, 0x0015), record(0x000a))
  let stream = Buffer.concat(parts)
  if (!mini) {
    // A well-bounded unknown record supplies inert bytes to use regular sectors.
    const filler = record(0x0867, Buffer.alloc(4096 - stream.length - 4))
    stream = Buffer.concat([stream, filler])
  }
  const cfb = XLSX.CFB.utils.cfb_new()
  XLSX.CFB.utils.cfb_add(cfb, 'Workbook', stream)
  XLSX.CFB.utils.cfb_add(cfb, 'Unrelated', Buffer.from('Unrelated contents must remain byte-identical.'))
  return { bytes: Buffer.from(XLSX.CFB.write(cfb, { type: 'buffer' })), offset, textLength: h.length, stream }
}
function stream(bytes) { return Buffer.from(XLSX.CFB.read(bytes, { type: 'buffer' }).FileIndex.find(e => e.name === 'Workbook').content) }
function records(bytes) {
  const result = []
  for (let offset = 0; offset < bytes.length;) {
    const length = bytes.readUInt16LE(offset + 2)
    result.push({ type: bytes.readUInt16LE(offset), offset, length })
    offset += 4 + length
  }
  return result
}
function textAt(bytes, offset) {
  const count = bytes.readUInt16LE(offset + 4), wide = bytes[offset + 6]
  return bytes.subarray(offset + 7, offset + 7 + count * (wide ? 2 : 1)).toString(wide ? 'utf16le' : 'latin1')
}

for (const wide of [false, true]) test(`black header/footer tokens disappear only in disposable ${wide ? 'Unicode' : 'compressed'} BIFF strings`, () => {
  const source = fixture({ wide, text: '&L&K000000Title 000000 &&K000000 &KFF0000 &"Font &K000000,Regular"' })
  const before = Buffer.from(source.bytes), result = prepareLegacySheetPreview(source.bytes)
  assert.notEqual(result, source.bytes)
  assert.deepEqual(source.bytes, before, 'caller bytes are never mutated')
  assert.equal(result.length, source.bytes.length)
  const old = stream(source.bytes), current = stream(result)
  assert.equal(current.length, old.length)
  assert.deepEqual(records(current), records(old), 'every BIFF record position/type/length is identical')
  assert.equal(textAt(current, source.offset), '&LTitle 000000 &&K000000 &KFF0000 &"Font &K000000,Regular"')
  const footer = records(current).find(r => r.type === 0x0015)
  assert.equal(textAt(current, footer.offset), '&RPage &P')
  const permitted = records(old).filter(r => [0x14, 0x15].includes(r.type))
  for (let i = 0; i < old.length; i++) if (old[i] !== current[i]) assert.ok(permitted.some(r => i >= r.offset + 4 && i < r.offset + 4 + r.length))
  const a = XLSX.CFB.read(source.bytes, { type: 'buffer' }), b = XLSX.CFB.read(result, { type: 'buffer' })
  assert.deepEqual(a.FullPaths, b.FullPaths)
  a.FileIndex.forEach((entry, index) => {
    assert.equal(entry.start, b.FileIndex[index].start); assert.equal(entry.size, b.FileIndex[index].size)
    if (entry.name !== 'Workbook') assert.deepEqual(entry.content, b.FileIndex[index].content)
  })
  assert.deepEqual(result.subarray(0, 512), source.bytes.subarray(0, 512), 'CFB header/allocation roots unchanged')
  assert.equal(prepareLegacySheetPreview(result), result, 'repeat preparation is a no-op')
})

test('literal escapes, font names and nonblack codes are retained', () => {
  const source = fixture({ text: '&&K000000 &"Font &K000000,Regular" &K123456 000000' })
  const result = prepareLegacySheetPreview(source.bytes)
  assert.equal(textAt(stream(result), source.offset), '&&K000000 &"Font &K000000,Regular" &K123456 000000')
})

test('malformed, encrypted, unsupported containers and strings remain untouched', () => {
  for (const options of [{ encrypted: true }, { malformed: true }, { flags: 2 }, { mini: true }]) {
    const { bytes } = fixture(options); assert.equal(prepareLegacySheetPreview(bytes), bytes)
  }
  const valid = fixture().bytes
  for (const mutate of [b => b.writeUInt16LE(4, 26), b => b.writeUInt32LE(1, 72), b => b.writeUInt32LE(999999, 76)]) {
    const input = Buffer.from(valid); mutate(input); assert.equal(prepareLegacySheetPreview(input), input)
  }
  for (const input of [Buffer.alloc(0), Buffer.from('not an XLS'), valid.subarray(0, valid.length - 1)]) assert.equal(prepareLegacySheetPreview(input), input)
})

test('cross-linked streams, root mini-streams and allocation cycles remain untouched', () => {
  const original = fixture().bytes
  const cfb = XLSX.CFB.read(original, { type: 'buffer' }), workbook = cfb.FileIndex.find(entry => entry.name === 'Workbook')
  const directory = (original.readUInt32LE(48) + 1) * 512
  for (const name of ['Unrelated', cfb.FileIndex.find(entry => entry.type === 5).name]) {
    const bytes = Buffer.from(original), index = cfb.FileIndex.findIndex(entry => entry.name === name)
    assert.ok(index < 4, 'small fixture directory fits one sector')
    bytes.writeUInt32LE(workbook.start, directory + index * 128 + 116)
    bytes.writeUInt32LE(4096, directory + index * 128 + 120)
    assert.equal(prepareLegacySheetPreview(bytes), bytes, `${name} cannot share Workbook sectors`)
  }
  const cycle = Buffer.from(original), fat = (cycle.readUInt32LE(76) + 1) * 512
  cycle.writeUInt32LE(workbook.start, fat + workbook.start * 4)
  assert.equal(prepareLegacySheetPreview(cycle), cycle)
})
