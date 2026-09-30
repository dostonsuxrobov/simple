'use strict'

const SIGNATURE = Buffer.from('d0cf11e0a1b11ae1', 'hex')
const FREE = 0xffffffff, END = 0xfffffffe

function withoutBlackColorCodes(text) {
  let output = ''
  for (let index = 0; index < text.length;) {
    if (text.slice(index, index + 2) === '&&') {
      output += '&&'; index += 2
    } else if (text.slice(index, index + 2) === '&"') {
      const end = text.indexOf('"', index + 2)
      if (end < 0) return text // Ambiguous font syntax: leave the whole header alone.
      output += text.slice(index, end + 1); index = end + 1
    } else if (text.slice(index, index + 8) === '&K000000') index += 8
    else output += text[index++]
  }
  return output
}

// LibreOffice 26.2's BIFF header parser ignores &K but prints its six RGB
// digits. Neutralize only black (the existing native default) in a disposable
// conversion copy. This is never used by Save or written over a source XLS.
// Other colors, escaped literals and font names retain their original bytes.
// A shorter string stays inside the same BIFF record with zeroed unused tail;
// the native reader consumes its declared character count. All record offsets,
// stream sizes, allocation tables and bytes outside that record remain intact.
function prepareLegacySheetPreview(bytes) {
  try {
    if (!Buffer.isBuffer(bytes) || bytes.length < 512 || bytes.length > 256 * 1024 * 1024 || !bytes.subarray(0, 8).equals(SIGNATURE)) return bytes
    // Deliberately bounded to BIFF8 in ordinary CFBv3 sectors. Mini-streams,
    // extended DIFAT chains and other containers use the unchanged native path.
    if (bytes.readUInt16LE(26) !== 3 || bytes.readUInt16LE(28) !== 0xfffe || bytes.readUInt16LE(30) !== 9 || bytes.readUInt16LE(32) !== 6 || bytes.length % 512) return bytes
    if (bytes.readUInt32LE(40) !== 0 || bytes.readUInt32LE(56) !== 4096 || bytes.readUInt32LE(72) !== 0 || bytes.readUInt32LE(68) !== END) return bytes
    const sectors = bytes.length / 512 - 1, fatCount = bytes.readUInt32LE(44)
    if (!fatCount || fatCount > 109) return bytes
    const fats = [], fatIds = new Set()
    for (let index = 0; index < 109; index++) {
      const sector = bytes.readUInt32LE(76 + index * 4)
      if (index >= fatCount) { if (sector !== FREE) return bytes; continue }
      if (sector >= sectors || fatIds.has(sector)) return bytes
      fatIds.add(sector); fats.push(bytes.subarray((sector + 1) * 512, (sector + 2) * 512))
    }
    const nextSector = sector => fats[Math.floor(sector / 128)]?.readUInt32LE((sector % 128) * 4)
    for (const sector of fatIds) if (nextSector(sector) !== 0xfffffffd) return bytes
    // A malformed CFB may cross-link its Workbook sectors with another stream,
    // the directory, or the root mini-stream. Validate exclusive allocation
    // before writing any byte, including structures we otherwise never touch.
    const allocated = new Set(fatIds)
    const claim = (start, count) => {
      const ids = []
      let current = start
      while (current !== END) {
        if (!Number.isInteger(current) || current < 0 || current >= sectors || allocated.has(current) || ids.length >= sectors || (count != null && ids.length >= count)) throw new Error('Invalid or overlapping CFB allocation')
        allocated.add(current); ids.push(current); current = nextSector(current)
      }
      if (count != null && ids.length !== count) throw new Error('Truncated CFB allocation')
      return ids
    }
    const directory = claim(bytes.readUInt32LE(48))
    // Read only the fixed directory fields needed here. Do not invoke a general
    // CFB parser before allocation validation: malformed chains can be cyclic.
    const entries = []
    for (const id of directory) for (let offset = 0; offset < 512; offset += 128) {
      const data = bytes.subarray((id + 1) * 512 + offset, (id + 1) * 512 + offset + 128)
      const type = data[66], nameLength = data.readUInt16LE(64), size = data.readUInt32LE(120)
      if (type && (![1, 2, 5].includes(type) || nameLength < 2 || nameLength > 64 || nameLength % 2 || data.readUInt16LE(nameLength - 2) !== 0 || data.readUInt32LE(124) !== 0 || size > bytes.length)) return bytes
      entries.push({ type, name: type ? data.subarray(0, nameLength - 2).toString('utf16le') : '', start: data.readUInt32LE(116), size, left: data.readUInt32LE(68), right: data.readUInt32LE(72), child: data.readUInt32LE(76) })
    }
    if (entries[0]?.type !== 5 || entries.filter(item => item.type === 5).length !== 1) return bytes
    const workbooks = entries.filter(item => item.type === 2 && /^(Workbook|Book)$/i.test(item.name))
    if (workbooks.length !== 1 || workbooks[0].size < 4096) return bytes
    const entry = workbooks[0]
    const children = new Set(), pending = [entries[0].child]
    while (pending.length) {
      const index = pending.pop()
      if (index === FREE) continue
      if (index >= entries.length || children.has(index) || ![1, 2].includes(entries[index].type)) return bytes
      children.add(index); pending.push(entries[index].left, entries[index].right)
    }
    if (!children.has(entries.indexOf(entry))) return bytes // Only a root workbook stream.
    const miniFatCount = bytes.readUInt32LE(64)
    if (miniFatCount > sectors) return bytes
    claim(bytes.readUInt32LE(60), miniFatCount)
    let chain
    for (const item of entries) {
      if (item.type !== 5 && !(item.type === 2 && item.size >= 4096)) continue
      if (!Number.isInteger(item.size) || item.size < 0 || item.size > bytes.length) return bytes
      if (!item.size) continue
      const ids = claim(item.start, Math.ceil(item.size / 512))
      if (item === entry) chain = ids
    }
    if (!chain) return bytes
    const stream = Buffer.alloc(entry.size)
    chain.forEach((id, index) => bytes.copy(stream, index * 512, (id + 1) * 512, (id + 1) * 512 + Math.min(512, entry.size - index * 512)))
    if (stream.readUInt16LE(0) !== 0x0809 || stream.readUInt16LE(4) !== 0x0600 || stream.readUInt16LE(6) !== 5) return bytes
    let changed = false, position = 0, worksheet = false
    while (position + 4 <= stream.length) {
      const type = stream.readUInt16LE(position), length = stream.readUInt16LE(position + 2)
      if (position + 4 + length > stream.length || type === 0x002f) return bytes // FILEPASS: encrypted workbook.
      const data = stream.subarray(position + 4, position + 4 + length)
      if (type === 0x0809) {
        if (length < 16 || data.readUInt16LE(0) !== 0x0600) return bytes
        worksheet = data.readUInt16LE(2) === 0x0010
      } else if (type === 0x000a) worksheet = false
      else if (type === 0x0014 || type === 0x0015) {
        if (length) {
          if (!worksheet || length < 3 || data[2] > 1) return bytes
          const count = data.readUInt16LE(0), width = data[2] ? 2 : 1, end = 3 + count * width
          if (end > length || data.subarray(end).some(value => value !== 0)) return bytes
          const encoding = width === 2 ? 'utf16le' : 'latin1'
          const text = data.subarray(3, end).toString(encoding), clean = withoutBlackColorCodes(text)
          if (clean !== text) {
            data.fill(0, 3, end); Buffer.from(clean, encoding).copy(data, 3); data.writeUInt16LE(clean.length, 0); changed = true
          }
        }
      }
      position += 4 + length
    }
    if (position !== stream.length || !changed) return bytes
    const output = Buffer.from(bytes)
    chain.forEach((id, index) => stream.copy(output, (id + 1) * 512, index * 512, Math.min((index + 1) * 512, stream.length)))
    return output
  } catch {
    return bytes // An unsupported or malformed workbook remains untouched.
  }
}

module.exports = { prepareLegacySheetPreview }
