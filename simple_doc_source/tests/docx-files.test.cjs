const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const { validateDocxBytes, validateDocxPackage } = require('../electron/docx-files.cjs')

function crc32(bytes) {
  let crc = 0xffffffff
  for (const byte of bytes) {
    crc ^= byte
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0)
  }
  return (crc ^ 0xffffffff) >>> 0
}

function storedZip(files) {
  const localParts = []
  const centralParts = []
  let localOffset = 0

  for (const [name, value] of Object.entries(files)) {
    const nameBytes = Buffer.from(name, 'utf8')
    const data = Buffer.from(value)
    const checksum = crc32(data)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt16LE(0x0800, 6)
    local.writeUInt32LE(checksum, 14)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(nameBytes.length, 26)
    localParts.push(local, nameBytes, data)

    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(20, 4)
    central.writeUInt16LE(20, 6)
    central.writeUInt16LE(0x0800, 8)
    central.writeUInt32LE(checksum, 16)
    central.writeUInt32LE(data.length, 20)
    central.writeUInt32LE(data.length, 24)
    central.writeUInt16LE(nameBytes.length, 28)
    central.writeUInt32LE(localOffset, 42)
    centralParts.push(central, nameBytes)
    localOffset += local.length + nameBytes.length + data.length
  }

  const centralDirectory = Buffer.concat(centralParts)
  const eocd = Buffer.alloc(22)
  eocd.writeUInt32LE(0x06054b50, 0)
  eocd.writeUInt16LE(Object.keys(files).length, 8)
  eocd.writeUInt16LE(Object.keys(files).length, 10)
  eocd.writeUInt32LE(centralDirectory.length, 12)
  eocd.writeUInt32LE(localOffset, 16)
  return Buffer.concat([...localParts, centralDirectory, eocd])
}

function alternateMainPartDocx({ includeMain = true } = {}) {
  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Override PartName="/custom/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
    </Types>`
  return storedZip({
    '[Content_Types].xml': contentTypes,
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/main.xml"/></Relationships>',
    ...(includeMain ? { 'custom/main.xml': '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Alternate OPC part</w:t></w:r></w:p><w:sectPr/></w:body></w:document>' } : {}),
  })
}

test('the checked-in DOCX fixture still validates with its canonical main part', () => {
  const fixture = fs.readFileSync(path.join(__dirname, '..', 'qa', 'fixtures', 'simple-docs-roundtrip-fixture.docx'))
  assert.equal(validateDocxPackage(fixture).mainDocumentPart, 'word/document.xml')
  assert.equal(validateDocxBytes(fixture), fixture)
})

test('a standards-valid DOCX can declare its main document at a noncanonical package path', async () => {
  const docx = alternateMainPartDocx()
  assert.equal(validateDocxPackage(docx).mainDocumentPart, 'custom/main.xml')
  assert.equal(validateDocxBytes(docx), docx)
  const { runImport } = await import('@forevka/wordcanvas/import')
  const imported = runImport(new Uint8Array(docx), undefined, { collectMediaBytes: true })
  assert.match(JSON.stringify(imported.doc), /Alternate OPC part/)
})

test('a declared main document part must exist in the package', () => {
  assert.throws(() => validateDocxPackage(alternateMainPartDocx({ includeMain: false })), /declared main Word document part/)
})

test('OLE-wrapped encrypted Word files get an actionable password error', () => {
  const oleHeader = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
  assert.throws(() => validateDocxBytes(oleHeader), /Password-protected Word documents are not supported.*Remove the password/i)
})
