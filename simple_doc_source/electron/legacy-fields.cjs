'use strict'

const CFB_SIGNATURE = Buffer.from('d0cf11e0a1b11ae1', 'hex')
// FibRgFcLcb97: main text, headers, footnotes, comments, endnotes, text boxes,
// and header text boxes. These pair offsets remain present in later DOC FIBs.
const FIELD_TABLE_PAIRS = [16, 17, 18, 19, 48, 57, 59]
const DATE_TIME_TYPES = new Set([0x1f, 0x20])
const LOCKED = 0x10

/** Find lock-byte offsets only after validating the complete nested FieldList. */
function dateFieldLockOffsets(table, start, length) {
  if (!Number.isInteger(start) || !Number.isInteger(length) || start < 0 || length < 4
    || start + length > table.length || (length - 4) % 6 !== 0) return []
  const count = (length - 4) / 6
  if (count > 100_000) return []
  const fieldsStart = start + (count + 1) * 4
  const stack = []
  const offsets = []
  let previousCp = -1
  for (let index = 0; index < count; index++) {
    const cp = table.readUInt32LE(start + index * 4)
    if (cp <= previousCp || cp > 0x7fffffff) return []
    previousCp = cp
    const entry = fieldsStart + index * 2
    const character = table[entry] & 0x1f // Upper three bits are reserved.
    if (character === 0x13) stack.push({ type: table[entry + 1], separator: null })
    else if (character === 0x14) {
      const field = stack.at(-1)
      if (!field || field.separator !== null) return []
      field.separator = cp
    } else if (character === 0x15) {
      const field = stack.pop()
      if (!field || Boolean(table[entry + 1] & 0x80) !== (field.separator !== null)) return []
      // A date with no cached result should remain calculable. PAGE, REF, TOC,
      // and all other fields retain their existing update behavior.
      if (DATE_TIME_TYPES.has(field.type) && field.separator !== null && cp > field.separator + 1
        && !(table[entry + 1] & LOCKED)) offsets.push(entry + 1)
    } else return []
  }
  if (stack.length || table.readUInt32LE(start + count * 4) <= previousCp) return []
  return offsets
}

/**
 * Preserve cached dates during import/print conversion without editing the file
 * supplied by the user. MS-DOC grffldEnd fLocked is bit 4, NOT bit 5 (private).
 * https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/28ab752b-055a-4725-8797-159bba0d125c
 */
function lockLegacyDateFields(value) {
  const source = Buffer.from(value)
  const unchanged = () => ({ bytes: source, lockedFields: 0 })
  if (source.length < 512 || !source.subarray(0, 8).equals(CFB_SIGNATURE)) return unchanged()
  try {
    const CFB = require('cfb')
    const container = CFB.read(Buffer.from(source), { type: 'buffer' })
    const word = CFB.find(container, 'WordDocument')?.content
    if (!word || word.length < 34 || word.readUInt16LE(0) !== 0xa5ec
      || ![0xc1, 0xd9, 0x101, 0x10c, 0x112].includes(word.readUInt16LE(2))) return unchanged()
    const flags = word.readUInt16LE(10)
    if (flags & 0x0100) return unchanged() // Never modify encrypted documents.
    const table = CFB.find(container, flags & 0x0200 ? '1Table' : '0Table')?.content
    if (!table) return unchanged()
    const csw = word.readUInt16LE(32)
    const clwPosition = 34 + csw * 2
    if (clwPosition + 2 > word.length) return unchanged()
    const clw = word.readUInt16LE(clwPosition)
    const pairCountPosition = clwPosition + 2 + clw * 4
    if (pairCountPosition + 2 > word.length) return unchanged()
    const pairCount = word.readUInt16LE(pairCountPosition)
    const pairsStart = pairCountPosition + 2
    if (pairsStart + pairCount * 8 > word.length) return unchanged()
    const offsets = new Set()
    for (const index of FIELD_TABLE_PAIRS) {
      if (index >= pairCount) continue
      const start = word.readUInt32LE(pairsStart + index * 8)
      const length = word.readUInt32LE(pairsStart + index * 8 + 4)
      if (!length) continue
      for (const offset of dateFieldLockOffsets(table, start, length)) offsets.add(offset)
    }
    if (!offsets.size) return unchanged()
    for (const offset of offsets) table[offset] |= LOCKED
    return { bytes: Buffer.from(CFB.write(container, { type: 'buffer' })), lockedFields: offsets.size }
  } catch {
    // The established converter remains responsible for unsupported or damaged
    // documents. An optional metadata repair must never make them less usable.
    return unchanged()
  }
}

module.exports = { lockLegacyDateFields, dateFieldLockOffsets }
