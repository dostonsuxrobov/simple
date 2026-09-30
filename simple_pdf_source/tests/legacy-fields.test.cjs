'use strict'
const test = require('node:test')
const assert = require('node:assert/strict')
const CFB = require('cfb')
const { lockLegacyDateFields, dateFieldLockOffsets } = require('../electron/legacy-fields.cjs')

function fieldTable(entries) {
  const bytes = Buffer.alloc(4 + entries.length * 6)
  entries.forEach(([cp, character, flags], index) => {
    bytes.writeUInt32LE(cp, index * 4)
    bytes[(entries.length + 1) * 4 + index * 2] = character
    bytes[(entries.length + 1) * 4 + index * 2 + 1] = flags
  })
  bytes.writeUInt32LE(entries.at(-1)[0] + 1, entries.length * 4)
  return bytes
}

function fixture(table, pair = 16, whichTable = 1) {
  const container = CFB.utils.cfb_new()
  const word = Buffer.alloc(1024)
  word.writeUInt16LE(0xa5ec, 0); word.writeUInt16LE(0xc1, 2)
  word.writeUInt16LE(whichTable ? 0x0200 : 0, 10)
  word.writeUInt16LE(14, 32); word.writeUInt16LE(22, 62); word.writeUInt16LE(60, 152)
  word.writeUInt32LE(0, 154 + pair * 8); word.writeUInt32LE(table.length, 158 + pair * 8)
  CFB.utils.cfb_add(container, 'WordDocument', word)
  CFB.utils.cfb_add(container, whichTable ? '1Table' : '0Table', table)
  CFB.utils.cfb_add(container, 'UntouchedMetadata', Buffer.from('Preserve arbitrary streams'))
  return Buffer.from(CFB.write(container, { type: 'buffer' }))
}

test('DATE and TIME cached results lock at the matching end; nested PAGE stays dynamic', () => {
  const table = fieldTable([[1,0x13,0x1f],[3,0x14,0],[5,0x13,0x21],[7,0x14,0],[9,0x15,0xc0],[12,0x95,0x80],[14,0x13,0x20],[16,0x14,0],[22,0x15,0x80]])
  const original = fixture(table)
  const before = Buffer.from(original)
  const result = lockLegacyDateFields(original)
  assert.equal(result.lockedFields, 2)
  assert.deepEqual(original, before)
  const container = CFB.read(result.bytes, { type:'buffer' })
  const output = CFB.find(container,'1Table').content
  const changed = [...output].flatMap((value,index) => value!==table[index] ? [index] : [])
  assert.equal(changed.length,2)
  changed.forEach(index=>assert.equal(output[index]^table[index],0x10,'must lock, never set private-result bit'))
  assert.deepEqual(CFB.find(container,'WordDocument').content,CFB.find(CFB.read(original,{type:'buffer'}),'WordDocument').content)
  assert.equal(CFB.find(container,'UntouchedMetadata').content.toString(),'Preserve arbitrary streams')
  assert.deepEqual(lockLegacyDateFields(result.bytes),{bytes:result.bytes,lockedFields:0},'already locked fields are unchanged')
})

test('all document-part field tables and both table streams are supported', () => {
  const table=fieldTable([[1,0x13,0x1f],[4,0x14,0],[8,0x15,0x80]])
  for(const pair of [16,17,18,19,48,57,59]) for(const which of [0,1]) {
    assert.equal(lockLegacyDateFields(fixture(table,pair,which)).lockedFields,1)
  }
})

test('missing cached values, other field types and malformed field structures are untouched', () => {
  const malformed = [
    [[1,0x13,0x1f],[4,0x15,0]], // no cached result
    [[1,0x13,0x1f],[4,0x14,0],[5,0x15,0x80]], // empty result
    [[1,0x13,0x21],[4,0x14,0],[8,0x15,0x80]], // PAGE
    [[1,0x13,0x1f],[4,0x14,0],[8,0x15,0]], // malformed separator flag
    [[1,0x13,0x1f],[4,0x14,0],[8,0x15,0x80],[9,0x13,0x1f]], // unfinished nesting
    [[1,0x13,0x1f],[1,0x14,0],[8,0x15,0x80]], // duplicate CP
    [[1,0x14,0],[4,0x15,0x80]], // no begin
  ]
  for(const entries of malformed) {
    const source=fixture(fieldTable(entries))
    assert.deepEqual(lockLegacyDateFields(source),{bytes:source,lockedFields:0})
  }
  const table=fieldTable([[1,0x13,0x1f],[4,0x14,0],[8,0x15,0x80]])
  for(const [start,length] of [[-1,22],[0,21],[1,22],[0,100000000],[NaN,22]]) assert.deepEqual(dateFieldLockOffsets(table,start,length),[])
  for(const bytes of [Buffer.alloc(0),Buffer.from('not DOC'),Buffer.from('d0cf11e0a1b11ae1','hex'),Buffer.alloc(512)]) assert.deepEqual(lockLegacyDateFields(bytes),{bytes,lockedFields:0})
})

test('encrypted DOCs and invalid FIB offsets are returned byte-for-byte unchanged', () => {
  const source=fixture(fieldTable([[1,0x13,0x1f],[4,0x14,0],[8,0x15,0x80]]))
  for(const mutate of [word=>word.writeUInt16LE(0x0300,10),word=>word.writeUInt16LE(65535,32),word=>word.writeUInt32LE(0xffffffff,154+16*8)]) {
    const container=CFB.read(source,{type:'buffer'});mutate(CFB.find(container,'WordDocument').content)
    const bytes=Buffer.from(CFB.write(container,{type:'buffer'}))
    assert.deepEqual(lockLegacyDateFields(bytes),{bytes,lockedFields:0})
  }
})
