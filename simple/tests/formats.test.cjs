'use strict'

// Format registry, content sniffing and routing (design §6.1–§6.3).
// Samples are copied into a temporary folder before anything reads them.

const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const os = require('node:os')
const path = require('node:path')
const zlib = require('node:zlib')
const formats = require('../shared/electron/formats.cjs')
const routing = require('../electron/routing.cjs')
const { associationPlan } = require('../launcher/associations.cjs')
const { findServiceNames } = require('../scripts/local-only-guard.cjs')

const ROOT = path.resolve(__dirname, '..')
const SAMPLES = path.resolve(ROOT, '..', 'Simple test examples')
const REGISTRY_JSON = path.join(ROOT, 'shared', 'electron', 'formats.json')

// Routing as it shipped before the registry existed. Changing a route is a
// product decision: update this table on purpose, together with formats.json.
const SHIPPED_ROUTES = Object.freeze({
  docs: ['.docx', '.doc'],
  pdf: ['.pdf', '.txt', '.md'],
  image: ['.png', '.jpg', '.jpeg', '.jfif', '.jpe', '.jif', '.webp', '.gif', '.apng', '.bmp', '.svg', '.svgz', '.avif', '.ico', '.psd'],
  video: ['.mp4', '.m4v', '.webm', '.ogv', '.mov', '.mkv'],
  calc: [
    '.xlsx', '.xlsm', '.xlsb', '.xls', '.xltx', '.xltm', '.xlt', '.xlam', '.xla',
    '.xml', '.ods', '.fods', '.csv', '.tsv', '.tab', '.numbers', '.slk', '.sylk',
    '.dif', '.dbf', '.prn', '.wk1', '.wk2', '.wk3', '.wk4', '.wks', '.wq1',
    '.wq2', '.wb1', '.wb2', '.wb3', '.123', '.qpw', '.html', '.htm',
  ],
})

let scratch

test.before(() => {
  scratch = fs.mkdtempSync(path.join(os.tmpdir(), 'simple-formats-'))
})

test.after(() => {
  if (scratch) fs.rmSync(scratch, { recursive: true, force: true })
})

function write(name, bytes) {
  const target = path.join(scratch, name)
  fs.writeFileSync(target, bytes)
  return target
}

function copySample(sample, name = sample) {
  const target = path.join(scratch, name)
  fs.copyFileSync(path.join(SAMPLES, sample), target)
  return target
}

function freshRegistry(edit) {
  const data = JSON.parse(fs.readFileSync(REGISTRY_JSON, 'utf8'))
  edit(data)
  return formats.createRegistry(data)
}

function activate(data, formatId, workspace, operation = 'open') {
  const format = data.formats.find((entry) => entry.id === formatId)
  format.status = 'active'
  const entry = format.workspaces[workspace][operation]
  format.workspaces[workspace][operation] = typeof entry === 'string' ? entry : { ...entry, status: 'active' }
}

// --- fixture builders --------------------------------------------------------

function zip(entries) {
  const locals = []
  const centrals = []
  let offset = 0
  for (const entry of entries) {
    const name = Buffer.from(entry.name)
    const raw = Buffer.from(entry.data)
    const data = entry.deflate ? zlib.deflateRawSync(raw) : raw
    const crc = zlib.crc32(raw)
    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4)
    local.writeUInt16LE(entry.deflate ? 8 : 0, 8)
    local.writeUInt32LE(crc, 14)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(raw.length, 22)
    local.writeUInt16LE(name.length, 26)
    locals.push(local, name, data)
    const central = Buffer.alloc(46)
    central.writeUInt32LE(0x02014b50, 0)
    central.writeUInt16LE(20, 4)
    central.writeUInt16LE(20, 6)
    central.writeUInt16LE(entry.deflate ? 8 : 0, 10)
    central.writeUInt32LE(crc, 16)
    central.writeUInt32LE(data.length, 20)
    central.writeUInt32LE(raw.length, 24)
    central.writeUInt16LE(name.length, 28)
    central.writeUInt32LE(offset, 42)
    centrals.push(central, name)
    offset += 30 + name.length + data.length
  }
  const directory = Buffer.concat(centrals)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(entries.length, 8)
  end.writeUInt16LE(entries.length, 10)
  end.writeUInt32LE(directory.length, 12)
  end.writeUInt32LE(offset, 16)
  return Buffer.concat([...locals, directory, end])
}

function contentTypes(main) {
  return { name: '[Content_Types].xml', deflate: true, data: `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/main" ContentType="${main}"/></Types>` }
}

// A compound file whose directory sits in `sector` (0-based). Only the parts
// the sniffer reads are filled in.
function cfb(streams, { sector = 0, size = 4096 } = {}) {
  const buffer = Buffer.alloc(Math.max(size, 512 * (sector + 2)))
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(buffer, 0)
  buffer.writeUInt16LE(0x3e, 24)
  buffer.writeUInt16LE(3, 26)
  buffer.writeUInt16LE(0xfffe, 28)
  buffer.writeUInt16LE(9, 30)
  buffer.writeUInt32LE(sector, 48)
  const start = 512 * (sector + 1)
  ;['Root Entry', ...streams].forEach((name, index) => {
    const at = start + (index * 128)
    buffer.write(name, at, 'utf16le')
    buffer.writeUInt16LE((name.length + 1) * 2, at + 64)
    buffer[at + 66] = index ? 2 : 5
  })
  return buffer
}

function pngChunks(types) {
  const parts = [Buffer.from('89504e470d0a1a0a', 'hex')]
  for (const type of types) {
    const data = type === 'IHDR' ? Buffer.alloc(13) : type === 'acTL' ? Buffer.alloc(8) : Buffer.alloc(type === 'IEND' ? 0 : 4)
    const length = Buffer.alloc(4)
    length.writeUInt32BE(data.length)
    parts.push(length, Buffer.from(type, 'latin1'), data, Buffer.alloc(4))
  }
  return Buffer.concat(parts)
}

function ftyp(major, compatible) {
  const brands = Buffer.from([major, '\0\0\0\0', ...compatible].join(''), 'latin1')
  const box = Buffer.alloc(8)
  box.writeUInt32BE(8 + brands.length)
  box.write('ftyp', 4, 'latin1')
  return Buffer.concat([box, brands, Buffer.alloc(32)])
}

function ebml(docType) {
  const type = Buffer.from(docType, 'latin1')
  return Buffer.concat([Buffer.from('1a45dfa3a3428681', 'hex'), Buffer.from([0x42, 0x82, 0x80 | type.length]), type, Buffer.alloc(32)])
}

function utf16le(text, bom) {
  const body = Buffer.from(text, 'utf16le')
  return bom ? Buffer.concat([Buffer.from([0xff, 0xfe]), body]) : body
}

const DELIMITED = ['Name;Amount;Date', 'Apples;3;2024-01-02', 'Pears;5;2024-01-03', 'Plums;7;2024-01-04', 'Figs;11;2024-01-05', 'Kiwis;13;2024-01-06'].join('\r\n')
const PROSE = [
  'Dear team,',
  'Thank you for the update on the quarterly figures. The numbers look good, and the forecast is on track.',
  'Please send the final report before Friday, if possible.',
  'We will review it together next week.',
  'Kind regards,',
  'Alex',
].join('\r\n')
const WEB_PAGE = '<!DOCTYPE html><html><head><title>Notes</title></head><body><h1>Trip</h1><p>We met at noon and walked to the river.</p><table><tr><td>1</td></tr></table></body></html>'
const TABLE_PAGE = '<html><body><table><tr><th>Name</th><th>Qty</th></tr><tr><td>Apples</td><td>3</td></tr></table></body></html>'
const EXCEL_PAGE = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta name=ProgId content=Excel.Sheet></head><body><table><tr><td>1</td></tr></table></body></html>'
const WORD_2003 = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<?mso-application progid="Word.Document"?>\n<w:wordDocument xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"><w:body><w:p><w:r><w:t>Hi</w:t></w:r></w:p></w:body></w:wordDocument>'
const SHEET_2003 = '<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="A"/></Workbook>'
const SVG = '<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10"/></svg>'

// --- sniffing --------------------------------------------------------------------

test('sniff table: copies of the samples, mislabelled copies and generated fixtures', () => {
  const big = Buffer.alloc(200 * 1024, 7)
  const cases = [
    // [file name, bytes or sample to copy, expected kind, expected mislabeled]
    ['Complex document.docx', { sample: 'Complex document.docx' }, 'docx', false],
    ['Complex document.html', { sample: 'Complex document.html' }, 'html', false],
    ['Complex document.md', { sample: 'Complex document.md' }, 'md', false],
    ['Complex document.pdf', { sample: 'Complex document.pdf' }, 'pdf', false],
    ['Complex workbook.ods', { sample: 'Complex workbook.ods' }, 'ods', false],
    ['Complex workbook.xls', { sample: 'Complex workbook.xls' }, 'xls', false],
    ['Complex workbook.xlsx', { sample: 'Complex workbook.xlsx' }, 'xlsx', false],
    ['Complex workbook.pdf', { sample: 'Complex workbook.pdf' }, 'pdf', false],
    ['mislabeled-xls.xlsx', { sample: 'Complex workbook.xls' }, 'xls', true],
    ['mislabeled-ods.xlsx', { sample: 'Complex workbook.ods' }, 'ods', true],
    ['mislabeled-docx.doc', { sample: 'Complex document.docx' }, 'docx', true],
    ['html-named.xls', { sample: 'Complex document.html' }, 'html', true],
    ['pdf-named.docx', { sample: 'Complex document.pdf' }, 'pdf', true],
    ['xlsx-named.xlsm', { sample: 'Complex workbook.xlsx' }, 'xlsx', false],
    ['rtf-named.doc', '{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 Arial;}} Hello}', 'rtf', true],
    ['delimited.txt', DELIMITED, 'delimited-text', false],
    ['prose.txt', PROSE, 'txt', false],
    ['utf16-bom.csv', utf16le(DELIMITED, true), 'delimited-text', false],
    ['utf16-nobom.txt', utf16le(DELIMITED, false), 'delimited-text', false],
    ['word2003.xml', WORD_2003, 'wordml-2003', false],
    ['sheet2003.xml', SHEET_2003, 'spreadsheetml-2003', false],
    ['excel.htm', EXCEL_PAGE, 'excel-html', false],
    ['excel-web.xls', EXCEL_PAGE, 'excel-html', true],
    ['table.html', TABLE_PAGE, 'html-table', false],
    ['page.html', WEB_PAGE, 'html', false],
    ['drawing.svg', SVG, 'svg', false],
    ['flat.fods', '<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet"></office:document>', 'fods', false],
    ['data.slk', 'ID;PWXL;N;E\r\nC;Y1;X1;K"a"\r\nE\r\n', 'sylk', false],
    ['data.dif', 'TABLE\r\n0,1\r\n""\r\nVECTORS\r\n0,1\r\n""\r\n', 'dif', false],
    ['rows.json', '[{"a":1},{"a":2}]', 'json', false],
    ['generated.docx', zip([contentTypes('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'), { name: 'word/document.xml', data: '<w/>' }]), 'docx', false],
    ['generated.dotx', zip([contentTypes('application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml'), { name: 'word/document.xml', data: '<w/>' }]), 'dotx', false],
    ['generated.docm', zip([contentTypes('application/vnd.ms-word.document.macroEnabled.main+xml'), { name: 'word/document.xml', data: '<w/>' }, { name: 'word/vbaProject.bin', data: 'x' }]), 'docm', false],
    ['generated.xlsb', zip([{ name: '[Content_Types].xml', data: '<Types/>' }, { name: 'xl/workbook.bin', data: 'x' }]), 'xlsb', false],
    ['generated.xlsm', zip([{ name: '[Content_Types].xml', data: '<Types/>' }, { name: 'xl/workbook.xml', data: 'x' }, { name: 'xl/vbaProject.bin', data: 'x' }]), 'xlsm', false],
    ['generated.pptx', zip([{ name: '[Content_Types].xml', data: '<Types/>' }, { name: 'ppt/presentation.xml', data: 'x' }]), 'pptx', false],
    ['generated.odt', zip([{ name: 'mimetype', data: 'application/vnd.oasis.opendocument.text' }, { name: 'content.xml', data: '<x/>' }]), 'odt', false],
    ['generated.epub', zip([{ name: 'mimetype', data: 'application/epub+zip' }, { name: 'META-INF/container.xml', data: '<x/>' }]), 'epub', false],
    ['generated.numbers', zip([{ name: 'Index/Document.iwa', data: 'x' }]), 'numbers', false],
    ['archive.zip', zip([{ name: 'a.txt', data: 'hello' }]), 'zip', false],
    // The directory is only in the tail: a 200 KB first entry pushes it past the head.
    ['large.docx', zip([{ name: 'big.bin', data: big }, contentTypes('application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'), { name: 'word/document.xml', data: '<w/>' }]), 'docx', false],
    ['generated.doc', cfb(['WordDocument', '1Table']), 'doc', false],
    ['generated.xls', cfb(['Workbook']), 'xls', false],
    ['legacy-book.xls', cfb(['Book']), 'xls', false],
    ['generated.ppt', cfb(['PowerPoint Document', 'Current User']), 'ppt', false],
    ['encrypted.docx', cfb(['EncryptionInfo', 'EncryptedPackage']), 'ooxml-encrypted', true],
    ['tail-directory.doc', cfb(['WordDocument'], { sector: 400, size: 512 * 402 }), 'doc', false],
    ['image.png', pngChunks(['IHDR', 'IDAT', 'IEND']), 'png', false],
    ['animated.apng', pngChunks(['IHDR', 'acTL', 'IDAT', 'IEND']), 'apng', false],
    ['animated.png', pngChunks(['IHDR', 'acTL', 'IDAT', 'IEND']), 'apng', false],
    ['png-named.jpg', pngChunks(['IHDR', 'IDAT', 'IEND']), 'png', true],
    ['photo.jpg', Buffer.from('ffd8ffe000104a46494600', 'hex'), 'jpeg', false],
    ['photo.jfif', Buffer.from('ffd8ffe000104a46494600', 'hex'), 'jpeg', false],
    ['anim.gif', Buffer.concat([Buffer.from('GIF89a'), Buffer.alloc(20), Buffer.from('NETSCAPE2.0')]), 'gif', false],
    ['anim.webp', Buffer.concat([Buffer.from('RIFF'), Buffer.alloc(4), Buffer.from('WEBPVP8X'), Buffer.from([10, 0, 0, 0, 0x12, 0, 0, 0]), Buffer.alloc(16)]), 'webp', false],
    ['image.bmp', Buffer.concat([Buffer.from('BM'), Buffer.alloc(12), Buffer.from([40, 0, 0, 0]), Buffer.alloc(40)]), 'bmp', false],
    ['scan.tif', Buffer.from('49492a0008000000', 'hex'), 'tiff', false],
    ['scan.tiff', Buffer.from('4d4d002a00000008', 'hex'), 'tiff', false],
    ['icon.ico', Buffer.concat([Buffer.from('000001000100', 'hex'), Buffer.from([16, 16, 0, 0, 1, 0, 32, 0, 100, 0, 0, 0, 22, 0, 0, 0]), Buffer.alloc(100)]), 'ico', false],
    ['layers.psd', Buffer.concat([Buffer.from('8BPS'), Buffer.from([0, 1]), Buffer.alloc(32)]), 'psd', false],
    ['huge.psb', Buffer.concat([Buffer.from('8BPS'), Buffer.from([0, 2]), Buffer.alloc(32)]), 'psb', false],
    ['photo.heic', ftyp('heic', ['mif1', 'heic']), 'heic', false],
    ['photo.heif', ftyp('mif1', ['mif1', 'heic']), 'heic', false],
    ['photo.avif', ftyp('avif', ['avif', 'mif1']), 'avif', false],
    ['mif1-avif.avif', ftyp('mif1', ['mif1', 'avif']), 'avif', false],
    ['clip.mp4', ftyp('isom', ['isom', 'mp41']), 'mp4', false],
    ['clip.mov', ftyp('qt  ', ['qt  ']), 'mov', false],
    ['mov-named.mp4', ftyp('qt  ', ['qt  ']), 'mov', false],
    ['clip.webm', ebml('webm'), 'webm', false],
    ['clip.mkv', ebml('matroska'), 'mkv', false],
    ['clip.ogv', Buffer.concat([Buffer.from('OggS'), Buffer.alloc(32)]), 'ogv', false],
    ['drawing.svgz', zlib.gzipSync(Buffer.from(SVG)), 'svgz', false],
    ['notes.gz', zlib.gzipSync(Buffer.from('plain text')), 'gzip', false],
    ['late-header.pdf', Buffer.concat([Buffer.alloc(100, 0x20), Buffer.from('%PDF-1.7\n%%EOF\n')]), 'pdf', false],
    ['random.dbf', Buffer.from([0x03, 0x7c, 0x01, 0x01, 0, 0, 0, 0, 0x41, 0, 0x29, 0]), 'binary', false],
    ['empty.csv', Buffer.alloc(0), 'empty', false],
  ]
  const failures = []
  for (const [name, source, kind, mislabeled] of cases) {
    const filePath = source.sample ? copySample(source.sample, name) : write(name, source)
    const result = formats.sniffFile(filePath)
    if (result.kind !== kind || result.mislabeled !== mislabeled) {
      failures.push(`${name}: expected ${kind}${mislabeled ? ' (mislabeled)' : ''}, got ${result.kind}${result.mislabeled ? ' (mislabeled)' : ''}`)
    }
  }
  assert.deepEqual(failures, [])
})

test('sniff results carry registry labels, the format to read with, and details', () => {
  const delimited = formats.sniff(Buffer.from(DELIMITED), null, DELIMITED.length, 'data.txt')
  assert.equal(delimited.formatId, 'csv')
  assert.equal(delimited.label, 'Delimited text')
  assert.deepEqual([delimited.details.delimiter, delimited.details.columns], ['semicolon', 3])
  assert.equal(delimited.matchesExtension, true)

  const encrypted = formats.sniff(cfb(['EncryptionInfo', 'EncryptedPackage']), null, 4096, 'secret.xlsx')
  assert.equal(encrypted.encrypted, true)
  assert.equal(encrypted.label, 'Password-protected Office file')

  const xls = formats.sniffFile(copySample('Complex workbook.xls', 'labels.xlsx'))
  assert.equal(xls.formatId, 'xls')
  assert.equal(xls.extensionFormatId, 'xlsx')
  assert.equal(xls.label, 'Excel 97–2003 workbook')

  const utf16 = formats.sniff(utf16le(DELIMITED, true), null, undefined, 'x.csv')
  assert.equal(utf16.details.encoding, 'utf-16le')
  assert.equal(utf16.details.bom, true)

  // A head that is only a prefix of a longer file still sniffs.
  const prefix = Buffer.from(`${DELIMITED}\r\nCherries;1`)
  assert.equal(formats.sniff(prefix, null, prefix.length + 5000, 'long.txt').kind, 'delimited-text')
  assert.equal(formats.sniff(Buffer.alloc(0), null, 0, 'x.txt').kind, 'empty')
  assert.equal(formats.sniff(Buffer.from('hello'), null, 5, '').mislabeled, false)
})

test('prose with commas, Markdown tables and short lists are not delimited text', () => {
  const markdownTable = '# Prices\n\n| Item | Price |\n|------|-------|\n| A | 1 |\n| B | 2 |\n| C | 3 |\n| D | 4 |\n'
  assert.equal(formats.sniff(Buffer.from(markdownTable)).kind, 'md')
  const shortCsv = 'a,b\n1,2\n3,4\n'
  assert.equal(formats.sniff(Buffer.from(shortCsv)).kind, 'txt')
  const commaProse = Array.from({ length: 6 }, (_, index) => `Line ${index} talks about apples, and then it ends with a full stop.`).join('\n')
  assert.equal(formats.sniff(Buffer.from(commaProse)).kind, 'txt')
  const tabbed = Array.from({ length: 8 }, (_, index) => `${index}\tname ${index}\t${index * 2}`).join('\n')
  assert.equal(formats.sniff(Buffer.from(tabbed)).details.delimiter, 'tab')
  const quoted = ['"Name","Note"', '"A","x, y"', '"B","z"', '"C","w, v, u"', '"D","t"', '"E","s"'].join('\n')
  assert.equal(formats.sniff(Buffer.from(quoted)).details.columns, 2)
})

// --- routing ---------------------------------------------------------------------

test('extension routes are exactly the shipped ones and stay exclusive', () => {
  assert.deepEqual(JSON.parse(JSON.stringify(routing.EXTENSIONS_BY_MODE)), SHIPPED_ROUTES)
  assert.deepEqual(routing.MODES, ['docs', 'pdf', 'image', 'video', 'calc'])
  assert.deepEqual(routing.SUPPORTED_EXTENSIONS, Object.values(SHIPPED_ROUTES).flat())
  assert.deepEqual(formats.associationExtensions(), routing.SUPPORTED_EXTENSIONS)
  for (const [mode, extensions] of Object.entries(SHIPPED_ROUTES)) {
    for (const extension of extensions) {
      assert.equal(routing.modeForPath(`C:\\nowhere\\file${extension.toUpperCase()}`), mode, extension)
    }
  }
  assert.equal(routing.modeForPath('README'), null)
  assert.equal(routing.modeForPath('layers.psd'), 'image', 'the Image workspace opens .psd (layers in Advanced mode)')
})

test('content routing: delimited .txt opens in Spreadsheets, everything else keeps its route', () => {
  const delimited = write('route-delimited.txt', DELIMITED)
  const utf16 = write('route-utf16.txt', utf16le(DELIMITED, true))
  const prose = write('route-prose.txt', PROSE)
  const web = write('route-page.html', WEB_PAGE)
  const excel = write('route-excel.htm', EXCEL_PAGE)
  const word = write('route-word.xml', WORD_2003)
  const notes = write('route-notes.md', '# Title\n\n- one\n- two\n')

  assert.deepEqual(formats.routeForPath(delimited), { mode: 'calc', formatId: 'txt', extension: '.txt', by: 'content', kind: 'delimited-text' })
  assert.equal(routing.modeForPath(utf16), 'calc')
  assert.deepEqual(formats.routeForPath(prose), { mode: 'pdf', formatId: 'txt', extension: '.txt', by: 'extension', kind: 'txt' })
  assert.equal(routing.modeForPath(web), 'calc', 'Web pages stay in Spreadsheets until Documents imports HTML')
  assert.equal(routing.modeForPath(excel), 'calc')
  assert.equal(routing.modeForPath(word), 'calc', 'Word 2003 XML stays in Spreadsheets until Documents imports it')
  assert.equal(routing.modeForPath(notes), 'pdf')
  assert.equal(routing.modeForPath(path.join(scratch, 'missing.txt')), 'pdf', 'An unreadable file keeps its extension route')
  assert.equal(routing.modeForPath(path.basename(delimited), { cwd: scratch }), 'calc', 'Relative paths resolve against the given folder')

  assert.deepEqual(routing.supportedPaths(['simple.exe', delimited, prose, '--flag']), [delimited, prose])
  assert.deepEqual([...routing.groupPathsByMode([prose, delimited, web]).entries()], [['pdf', [prose]], ['calc', [delimited, web]]])
})

/**
 * Counts every file-system call that opens or stats a file under `folder`
 * while `action` runs (sync and async), restoring fs afterwards.
 */
async function countReads(folder, action) {
  const fsp = fs.promises
  const patched = [[fs, 'openSync'], [fs, 'statSync'], [fs, 'readFileSync'], [fsp, 'open'], [fsp, 'stat'], [fsp, 'readFile']]
  const originals = patched.map(([target, name]) => target[name])
  const reads = []
  const inside = (value) => typeof value === 'string' && path.resolve(value).toLowerCase().startsWith(path.resolve(folder).toLowerCase() + path.sep)
  patched.forEach(([target, name], index) => {
    target[name] = function counted(file, ...rest) {
      if (inside(file)) reads.push(`${name} ${path.basename(file)}`)
      return originals[index].call(this, file, ...rest)
    }
  })
  try {
    await action()
  } finally {
    patched.forEach(([target, name], index) => { target[name] = originals[index] })
  }
  return reads
}

test('a command line is routed once; the decision travels with the paths and is never re-sniffed', async () => {
  const folder = fs.mkdtempSync(path.join(scratch, 'once-'))
  const delimited = path.join(folder, 'once-delimited.txt')
  fs.writeFileSync(delimited, DELIMITED)
  const prose = path.join(folder, 'once-prose.txt')
  fs.writeFileSync(prose, PROSE)
  const page = path.join(folder, 'once-page.html')
  fs.writeFileSync(page, WEB_PAGE)
  // The new instance routes by content, relative to its own folder, reading each .txt once.
  let first
  const firstReads = await countReads(folder, () => {
    first = routing.routeCommandLine(['simple.exe', path.basename(delimited), path.basename(prose), 'b.xlsx', path.basename(page)], { cwd: folder })
  })
  assert.equal(first.mode, 'calc')
  assert.deepEqual([...first.groups.entries()], [['calc', ['once-delimited.txt', 'b.xlsx', 'once-page.html']], ['pdf', ['once-prose.txt']]])
  assert.deepEqual(firstReads.filter((read) => read.startsWith('openSync')).sort(), ['openSync once-delimited.txt', 'openSync once-prose.txt'],
    'each .txt is read once; a web page has one possible owner today and is never read')
  // A process started with --simple-mode keeps the forwarded decision and reads none of those files,
  // even when a sniff from here would differ.
  let forwarded
  const forwardedReads = await countReads(folder, () => {
    forwarded = routing.routeCommandLine(['simple.exe', '--simple-mode=calc', prose, delimited, page, path.join(folder, 'missing.txt'), 'c.pdf'])
  })
  assert.deepEqual(forwardedReads, [])
  assert.equal(forwarded.mode, 'calc')
  assert.deepEqual(forwarded.groups.get('calc'), [prose, delimited, page, path.join(folder, 'missing.txt')])
  assert.deepEqual(forwarded.groups.get('pdf'), ['c.pdf'], 'a mode never takes a file it cannot open')
  // Picking the routed arguments out of a command line reads nothing.
  assert.deepEqual(await countReads(folder, () => { routing.supportedPaths([prose, delimited, page]) }), [])

  // The launcher routes without blocking its window and forwards each decision,
  // so the process it starts reads nothing to decide again.
  const { launchValidatedPaths } = require('../launcher/open-paths.cjs')
  const launched = []
  let result
  const launcherReads = await countReads(folder, async () => {
    result = await launchValidatedPaths([delimited, prose, page, path.join(folder, 'Report.docx'), path.join(folder, 'notes.unknown')], {
      launch: async (args) => { launched.push(args); return 1 },
    })
  })
  assert.deepEqual(result, { opened: 4, unsupported: 1, failed: 0 })
  assert.deepEqual(launched, [
    ['--simple-mode=calc', delimited, page],
    ['--simple-mode=pdf', prose],
    ['--simple-mode=docs', path.join(folder, 'Report.docx')],
  ])
  assert.ok(launcherReads.every((read) => /^(open|stat) /.test(read)), `only asynchronous reads on the launcher's thread: ${launcherReads}`)
  assert.deepEqual(launcherReads.filter((read) => read.startsWith('open ')).sort(), ['open once-delimited.txt', 'open once-prose.txt'])
  for (const args of launched) {
    let child
    assert.deepEqual(await countReads(folder, () => { child = routing.routeCommandLine(['simple.exe', ...args]) }), [], `${args[0]}: the new process never re-reads`)
    assert.equal(child.mode, args[0].slice('--simple-mode='.length))
    assert.deepEqual([...child.groups.keys()], [child.mode])
  }
  // Relative paths from a second instance resolve against its folder.
  launched.length = 0
  await launchValidatedPaths(['once-delimited.txt'], { cwd: folder, launch: async (args) => { launched.push(args); return 1 } })
  assert.deepEqual(launched, [['--simple-mode=calc', delimited]])
  assert.deepEqual(routing.candidateModesForPath('x.TXT'), ['pdf', 'calc'])
  assert.deepEqual(routing.candidateModesForPath('x.docx'), ['docs'])
  assert.deepEqual(routing.keepOwnPaths(['simple.exe', 'a.txt', 'b.txt', '--flag', 'other'], ['b.txt']), ['simple.exe', 'b.txt', '--flag', 'other'])

  // The running workspace filters a second instance's command line without reading any file.
  const decided = ['simple.exe', 'locked.txt', 'other.txt', 'Report.docx', '--x']
  routing.filterSecondInstanceArgv(decided, { mode: 'calc', workingDirectory: scratch, additionalData: { simpleRouting: { mode: 'calc', paths: [path.join(scratch, 'locked.txt')] } } })
  assert.deepEqual(decided, ['simple.exe', path.join(scratch, 'locked.txt'), '--x'], 'only what the new instance routed here stays, made absolute')
  const undecided = ['simple.exe', 'notes.txt', 'Report.docx']
  routing.filterSecondInstanceArgv(undecided, { mode: 'calc', workingDirectory: scratch })
  assert.deepEqual(undecided, ['simple.exe', path.join(scratch, 'notes.txt')], 'without a decision only impossible routes are removed; a .txt is kept')

  // The asynchronous route matches the synchronous one and blocks nothing.
  assert.deepEqual(await formats.routeForPathAsync(delimited), formats.routeForPath(delimited))
  assert.deepEqual(await formats.routeForPathAsync(prose), formats.routeForPath(prose))
  assert.equal((await formats.routeForPathAsync(path.join(scratch, 'gone.txt'))).mode, 'pdf', 'an unreadable file keeps its extension route')
})

test('content routes switch on when the target workspace starts opening the format', () => {
  const registry = freshRegistry((data) => {
    activate(data, 'html', 'docs')
    activate(data, 'xml', 'docs')
    activate(data, 'txt', 'docs')
  })
  const route = (name, content) => registry.routeForPath(write(name, content)).mode
  assert.equal(route('later-page.html', WEB_PAGE), 'docs')
  assert.equal(route('later-table.html', TABLE_PAGE), 'calc')
  assert.equal(route('later-excel.htm', EXCEL_PAGE), 'calc')
  assert.equal(route('later-word.xml', WORD_2003), 'docs')
  assert.equal(route('later-sheet.xml', SHEET_2003), 'calc')
  assert.equal(route('later-prose.txt', PROSE), 'docs')
  assert.equal(route('later-delimited.txt', DELIMITED), 'calc')
  assert.equal(registry.defaultModeForExtension('.html'), 'docs')
  assert.equal(registry.defaultModeForExtension('.txt'), 'docs')
})

test('planned formats route, associate and appear in filters only once enabled', () => {
  // .psd, .jfif/.jpe/.jif, .ico, .apng and .svgz are active now that the Image workspace opens them.
  const planned = ['.tif', '.tiff', '.heic', '.heif']
  for (const extension of planned) {
    const format = formats.formatForExtension(extension, { includePlanned: true })
    assert.ok(format, `${extension} is in the registry`)
    assert.deepEqual(format.route, ['image'], extension)
    assert.ok(format.workspaces.image.open, `${extension} declares an Image open`)
    assert.equal(formats.defaultModeForExtension(extension), null, `${extension} is not routed while planned`)
  }
  assert.ok(!formats.openExtensions('image').includes('.tiff'))
  assert.ok(formats.openExtensions('image').includes('.psd'))

  const registry = freshRegistry((data) => activate(data, 'psd', 'image'))
  assert.equal(registry.defaultModeForExtension('LAYERS.PSD'), 'image')
  assert.ok(registry.extensionsByMode.image.includes('.psd'))
  assert.ok(registry.fileAssociations().find((group) => group.name === 'simple image').ext.includes('psd'))
  assert.ok(registry.dialogFilters('image')[0].extensions.includes('psd'))
  const psd = Buffer.concat([Buffer.from('8BPS'), Buffer.from([0, 1]), Buffer.alloc(32)])
  assert.equal(registry.routeForPath(write('no-extension-psd', psd), { sniffUnknown: true }).mode, 'image')

  for (const extension of ['.txt', '.md', '.html', '.htm', '.rtf', '.odt', '.docm', '.dotx', '.dotm', '.docx', '.doc']) {
    const format = formats.formatForExtension(extension, { includePlanned: true })
    assert.ok(format.workspaces.docs?.open, `Documents declares an open for ${extension}`)
  }
  for (const id of ['png', 'jpeg', 'docx', 'doc', 'rtf', 'odt', 'xlsx', 'xls', 'ods', 'csv', 'txt', 'md', 'html', 'pptx']) {
    assert.ok(formats.formatById(id).workspaces.pdf?.open, `PDF declares an import for ${id}`)
  }
})

test('extensionless files are sniffed only when asked', () => {
  const pdf = copySample('Complex document.pdf', 'no-extension-pdf')
  assert.equal(routing.modeForPath(pdf), null)
  assert.deepEqual(formats.routeForPath(pdf, { sniffUnknown: true }), { mode: 'pdf', formatId: 'pdf', extension: '', by: 'content', kind: 'pdf' })
  assert.equal(formats.routeForPath(write('no-extension-bin', Buffer.from([0, 1, 2, 3])), { sniffUnknown: true }).mode, null)
  assert.equal(formats.routeForPath(scratch, { sniffUnknown: true }).mode, null, 'A folder is never routed')
})

// --- associations, filters, policy ----------------------------------------------

test('package.json fileAssociations equal the registry and can be regenerated in place', () => {
  const packagePath = path.join(ROOT, 'package.json')
  const packageJson = JSON.parse(fs.readFileSync(packagePath, 'utf8'))
  assert.deepEqual(packageJson.build.fileAssociations, formats.fileAssociations(),
    'Run: node electron/routing.cjs --write-associations')

  const copy = path.join(scratch, 'package.json')
  const original = fs.readFileSync(packagePath)
  fs.writeFileSync(copy, original)
  assert.equal(routing.writeFileAssociations(copy), false)
  assert.ok(fs.readFileSync(copy).equals(original), 'A current file is left byte-identical')

  const stale = JSON.parse(original.toString('utf8'))
  stale.build.fileAssociations = [{ ext: ['docx'], name: 'old', description: 'old', role: 'Editor', icon: 'build/icon.ico' }]
  fs.writeFileSync(copy, JSON.stringify(stale, null, 2).replace(/\n/g, '\r\n'))
  assert.equal(routing.writeFileAssociations(copy), true)
  const rewritten = fs.readFileSync(copy, 'utf8')
  assert.deepEqual(JSON.parse(rewritten).build.fileAssociations, formats.fileAssociations())
  assert.ok(!/[^\r]\n/.test(rewritten), 'CRLF line endings are kept')
})

test('the launcher registers every routed extension, and only those', () => {
  const plan = associationPlan('C:\\Apps\\simple.exe')
  const openWith = plan.filter((args) => /\\OpenWithProgids$/.test(args[1])).map((args) => args[1].split('\\').at(-2))
  assert.deepEqual(openWith, routing.SUPPORTED_EXTENSIONS)
  assert.ok(plan.every((args) => args[0] === 'ADD' && /^HKCU\\/.test(args[1])))
  assert.ok(plan.some((args) => args.includes('"C:\\Apps\\simple.exe" "%1"')))
})

test('Open filters: "All supported" first, "All files" last, and nothing offered that cannot open now', () => {
  for (const workspace of ['docs', 'pdf', 'calc', 'image', 'video', 'combine', 'launcher']) {
    for (const engine of [false, true]) {
      const filters = formats.dialogFilters(workspace, { engine })
      assert.match(filters[0].name, /^All /, workspace)
      assert.deepEqual(filters.at(-1), { name: 'All files', extensions: ['*'] })
      const expected = formats.openExtensions(workspace, { engine }).map((extension) => extension.slice(1))
      assert.deepEqual(new Set(filters[0].extensions), new Set(expected), `${workspace} engine=${engine}`)
      for (const filter of filters.slice(1, -1)) {
        assert.ok(filter.extensions.length, `${workspace}: empty filter ${filter.name}`)
        for (const extension of filter.extensions) assert.ok(expected.includes(extension), `${workspace}: ${filter.name} offers .${extension}`)
      }
    }
  }
  const pdfWithout = formats.dialogFilters('pdf')[0].extensions
  const pdfWith = formats.dialogFilters('pdf', { engine: true })[0].extensions
  assert.ok(!pdfWithout.includes('docx') && pdfWith.includes('docx'), 'Engine-only imports are hidden without the engine')
  assert.deepEqual(new Set(formats.dialogFilters('launcher')[0].extensions), new Set(routing.SUPPORTED_EXTENSIONS.map((extension) => extension.slice(1))))
  assert.deepEqual(formats.dialogFilters('docs', { allFiles: false }).map((filter) => filter.name), ['All documents', 'Word documents'])
  assert.deepEqual(formats.dialogFilters('calc')[1], { name: 'Excel workbooks', extensions: ['xlsx', 'xlsm', 'xlsb', 'xls', 'xltx', 'xltm', 'xlt', 'xlam', 'xla'] })
})

test('every routed extension is opened by its owner workspace now', () => {
  for (const [mode, extensions] of Object.entries(routing.EXTENSIONS_BY_MODE)) {
    const openable = new Set(formats.openExtensions(mode, { engine: false }))
    for (const extension of extensions) assert.ok(openable.has(extension), `${mode} must open ${extension} without the office engine`)
  }
})

test('save and export policy without and with the local office engine', () => {
  assert.deepEqual(formats.savePolicy('xls', 'calc'), { target: 'sibling', formatId: 'xlsx', extension: '.xlsx' })
  assert.deepEqual(formats.savePolicy('xls', 'calc', { engine: true }), { target: 'in-place', usesEngine: true })
  assert.deepEqual(formats.savePolicy('doc', 'docs'), { target: 'sibling', formatId: 'docx', extension: '.docx' })
  assert.deepEqual(formats.savePolicy('ods', 'calc'), { target: 'sibling', formatId: 'xlsx', extension: '.xlsx' })
  assert.deepEqual(formats.savePolicy('docx', 'docs'), { target: 'in-place', usesEngine: false })
  assert.deepEqual(formats.savePolicy('csv', 'calc'), { target: 'in-place', usesEngine: false })
  assert.deepEqual(formats.savePolicy('docx', 'pdf'), { target: 'sibling', formatId: 'pdf', extension: '.pdf' })
  assert.equal(formats.savePolicy('mp4', 'video'), null)

  assert.deepEqual(formats.openPolicy('docx', 'pdf'), { available: false, reason: 'needs-office-engine', mode: 'convert', engine: 'required', usesEngine: false })
  assert.equal(formats.openPolicy('docx', 'pdf', { engine: true }).available, true)
  assert.equal(formats.openPolicy('rtf', 'docs').reason, 'planned')
  assert.equal(formats.openPolicy('png', 'calc').reason, 'not-supported')
  assert.equal(formats.openPolicy('doc', 'docs').mode, 'text-only')

  const calcExports = formats.exportFormats('calc')
  const xls = calcExports.find((row) => row.id === 'xls')
  assert.deepEqual([xls.available, xls.mode, xls.usesEngine], [true, 'values-only', false])
  assert.equal(formats.exportFormats('calc', { engine: true }).find((row) => row.id === 'xls').usesEngine, true)
  assert.ok(!calcExports.some((row) => row.id === 'json'), 'Planned exports are not offered')
  for (const workspace of ['docs', 'pdf', 'calc', 'image', 'video']) {
    for (const row of formats.exportFormats(workspace)) {
      if (!row.available) assert.equal(row.reason, 'needs-office-engine')
    }
  }
})

test('registry hygiene: valid shape, unique extensions per route, local-only wording', () => {
  const data = JSON.parse(fs.readFileSync(REGISTRY_JSON, 'utf8'))
  assert.throws(() => formats.createRegistry({ ...data, formats: [...data.formats, { ...data.formats[0] }] }), /invalid|twice/)
  assert.throws(() => formats.createRegistry({ ...data, formats: [{ ...data.formats[0], route: ['nowhere'] }] }), /unknown route/)
  for (const format of data.formats) {
    assert.match(format.mime, /^[a-z]+\/[\w.+-]+$/, `${format.id} has a MIME type`)
    assert.ok(format.label, `${format.id} has a label`)
    for (const [workspace, operations] of Object.entries(format.workspaces || {})) {
      for (const [operation, entry] of Object.entries(operations)) {
        assert.ok(['open', 'save', 'export'].includes(operation), `${format.id}.${workspace}.${operation}`)
        const engine = typeof entry === 'object' ? entry.engine : undefined
        assert.ok([undefined, 'upgrade', 'required'].includes(engine), `${format.id}.${workspace}.${operation} engine`)
      }
    }
  }
  const strings = JSON.stringify(data)
  assert.deepEqual(findServiceNames(strings), [], 'The registry is local-only')
  assert.doesNotMatch(strings, /cloud/i, 'The registry is local-only')
  assert.doesNotMatch(strings, /\b(install|download)/i, 'User-facing registry text never asks to install or download')
  assert.doesNotMatch(strings, /https?:\/\/(?!schemas\.|www\.w3\.org)/i, 'No remote URLs')
})
