'use strict'

/**
 * Build a PDF from literal object bodies (object N is objects[N - 1]) with a
 * classic xref table, so tests control the exact bytes, spacing included.
 */
function buildRawPdf(objects, { root = 1, trailer = '' } = {}) {
  const header = Buffer.from('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n', 'latin1')
  const parts = [header]
  const offsets = []
  let length = header.length
  objects.forEach((body, index) => {
    offsets.push(length)
    const chunk = Buffer.concat([
      Buffer.from(`${index + 1} 0 obj\n`, 'latin1'),
      Buffer.isBuffer(body) ? body : Buffer.from(body, 'latin1'),
      Buffer.from('\nendobj\n', 'latin1'),
    ])
    parts.push(chunk)
    length += chunk.length
  })
  let xref = `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  for (const offset of offsets) xref += `${String(offset).padStart(10, '0')} 00000 n \n`
  xref += `trailer\n<< /Size ${objects.length + 1} /Root ${root} 0 R ${trailer}>>\nstartxref\n${length}\n%%EOF\n`
  parts.push(Buffer.from(xref, 'latin1'))
  return Buffer.concat(parts)
}

/**
 * Fill a `/ByteRange [0 0000000000 0000000000 0000000000]` placeholder with
 * the real offsets around the following /Contents hex string, as signers do.
 */
function patchByteRange(pdf) {
  const bytes = Buffer.from(pdf)
  const placeholder = bytes.indexOf('0000000000 0000000000 0000000000')
  const contents = bytes.indexOf('/Contents', placeholder)
  const gapStart = bytes.indexOf('<', contents)
  const gapEnd = bytes.indexOf('>', gapStart) + 1
  const values = [gapStart, gapEnd, bytes.length - gapEnd].map((value) => String(value).padStart(10, '0')).join(' ')
  bytes.write(values, placeholder, 'latin1')
  return bytes
}

module.exports = { buildRawPdf, patchByteRange }
