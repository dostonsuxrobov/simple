'use strict'
const XLSX = require('xlsx')

// Office conversion rounds row heights to screen pixels and normalizes column
// widths/margins. Restore only fixed-size BIFF layout fields from our model;
// record positions remain untouched. Constant error tokens are also corrected
// when the converter substitutes #N/A while caching a different source error.
function retainLegacyLayout(bytes, model) {
  const cfb = XLSX.CFB.read(bytes, { type: 'buffer' })
  const entry = cfb.FileIndex.find(item => item.type === 2 && /^(Workbook|Book)$/i.test(item.name))
  if (!entry) throw new Error('Converted XLS has no Workbook stream.')
  const stream = Buffer.from(entry.content)
  if (stream.readUInt16LE(0) !== 0x0809 || stream.readUInt16LE(4) !== 0x0600) throw new Error('Converted XLS is not the expected BIFF8 format.')
  const records = []
  for (let offset = 0; offset + 4 <= stream.length;) {
    const length = stream.readUInt16LE(offset + 2)
    if (offset + 4 + length > stream.length) throw new Error('Converted XLS has a truncated layout record.')
    records.push({ type: stream.readUInt16LE(offset), offset, data: stream.subarray(offset + 4, offset + 4 + length) })
    offset += 4 + length
  }
  const bounds = records.filter(record => record.type === 0x0085 && record.data.length >= 8).map(record => {
    const wide = Boolean(record.data[7] & 1), length = record.data[6]
    return { offset: record.data.readUInt32LE(0), name: record.data.subarray(8, 8 + length * (wide ? 2 : 1)).toString(wide ? 'utf16le' : 'latin1') }
  }).sort((a,b) => a.offset - b.offset)
  for (let index = 0; index < bounds.length; index++) {
    const bound = bounds[index], sheet = model.sheets.find(item => item.name === bound.name)
    if (!sheet) throw new Error('Converted XLS sheet layout cannot be matched to the edited workbook.')
    const end = bounds[index + 1]?.offset ?? stream.length
    const setup = sheet.pageSetup || {}, margins = setup.margins || {}
    for (const record of records) {
      if (record.offset < bound.offset || record.offset >= end) continue
      const data = record.data
      if (record.type === 0x0006 && data.length === 24 && data.readUInt16LE(20) === 2 && data[22] === 0x1c) {
        const cell = sheet.cells?.[XLSX.utils.encode_cell({ r: data.readUInt16LE(0), c: data.readUInt16LE(2) })]
        const errors = { '#NULL!': 0, '#DIV/0!': 7, '#VALUE!': 15, '#REF!': 23, '#NAME?': 29, '#NUM!': 36, '#N/A': 42 }
        const sourceError = !cell?.formula && cell?.type === 'error' ? cell.value : /^#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A)$/.test(cell?.formula || '') ? cell.formula : null
        if (sourceError && Object.hasOwn(errors, sourceError) && data[6] === 2 && data.readUInt16LE(12) === 0xffff) {
          // MS-XLS PtgErr is exactly token0x1c + one BErr byte. Keep formula,
          // cached error and source constant consistent without resizing BIFF.
          data[23] = errors[sourceError]
          data[8] = errors[sourceError]
        }
      } else if (record.type === 0x0208 && data.length >= 16) {
        const height = sheet.rowHeights?.[String(data.readUInt16LE(0) + 1)]
        if (Number.isFinite(height) && height > 0 && height <= 408.95) {
          data.writeUInt16LE(Math.round(height * 20), 6)
          data[12] |= 0x40
        }
      } else if (record.type === 0x007d && data.length >= 12) {
        const first = data.readUInt16LE(0) + 1, last = Math.min(256, data.readUInt16LE(2) + 1)
        const width = sheet.colWidths?.[String(first)]
        if (Number.isFinite(width) && width > 0 && width <= 255 && first === last) data.writeUInt16LE(Math.round(width * 256), 4)
        const tail = sheet.properties?.legacyHiddenColumnTail
        if (tail && first >= tail.start && first <= tail.end && last <= tail.end) {
          data.writeUInt16LE(last - 1, 2)
          data.writeUInt16LE(0, 4)
          data.writeUInt16LE(data.readUInt16LE(8) | 1, 8)
        }
      } else if ([0x26, 0x27, 0x28, 0x29].includes(record.type) && data.length >= 8) {
        const key = {0x26:'left',0x27:'right',0x28:'top',0x29:'bottom'}[record.type]
        if (Number.isFinite(margins[key]) && margins[key] >= 0 && margins[key] <= 50) data.writeDoubleLE(margins[key], 0)
      } else if (record.type === 0x00a1 && data.length >= 34) {
        for (const [key, offset] of [['header',16],['footer',24]]) {
          if (Number.isFinite(margins[key]) && margins[key] >= 0 && margins[key] <= 50) data.writeDoubleLE(margins[key], offset)
        }
        for (const [key, offset] of [['horizontalDpi',12],['verticalDpi',14]]) {
          if (Number.isInteger(setup[key]) && setup[key] > 0 && setup[key] <= 65535) data.writeUInt16LE(setup[key], offset)
        }
      }
    }
  }
  entry.content = stream
  entry.size = stream.length
  return Buffer.from(XLSX.CFB.write(cfb, { type: 'buffer' }))
}
module.exports = { retainLegacyLayout }
