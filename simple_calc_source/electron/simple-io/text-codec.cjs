// Vendored from simple/shared/electron/text-codec.cjs by simple/scripts/sync-shared.cjs. Do not edit here.
'use strict'

// Text decoding and encoding that round-trips the file's own dialect
// (encoding, byte order mark, line endings). Only Node built-ins: TextDecoder
// with full ICU covers every Windows code page Simple reads, and encoding back
// to a legacy code page uses a reverse table built from that same decoder, so
// decode(encode(text)) is exact for every character the code page can hold.

const { execFileSync } = require('node:child_process')
const path = require('node:path')

const UTF8_BOM = Buffer.from([0xef, 0xbb, 0xbf])
const UTF16LE_BOM = Buffer.from([0xff, 0xfe])
const UTF16BE_BOM = Buffer.from([0xfe, 0xff])
const UTF16_SAMPLE_BYTES = 4096
const UTF16_NUL_RATIO = 0.3

/** Windows code page number → WHATWG encoding label understood by TextDecoder. */
const CODE_PAGE_LABELS = Object.freeze({
  866: 'ibm866',
  874: 'windows-874',
  932: 'shift_jis',
  936: 'gbk',
  949: 'euc-kr',
  950: 'big5',
  1250: 'windows-1250',
  1251: 'windows-1251',
  1252: 'windows-1252',
  1253: 'windows-1253',
  1254: 'windows-1254',
  1255: 'windows-1255',
  1256: 'windows-1256',
  1257: 'windows-1257',
  1258: 'windows-1258',
  20866: 'koi8-r',
  21866: 'koi8-u',
  28591: 'windows-1252',
  65001: 'utf-8',
})

/** Encodings Simple can write; the export dialog lists these. */
const WRITABLE_ENCODINGS = Object.freeze([
  { id: 'utf-8', label: 'UTF-8' },
  { id: 'utf-16le', label: 'UTF-16 (little endian)' },
  { id: 'utf-16be', label: 'UTF-16 (big endian)' },
  { id: 'windows-1252', label: 'Western European (Windows-1252)' },
  { id: 'windows-1250', label: 'Central European (Windows-1250)' },
  { id: 'windows-1251', label: 'Cyrillic (Windows-1251)' },
  { id: 'windows-1253', label: 'Greek (Windows-1253)' },
  { id: 'windows-1254', label: 'Turkish (Windows-1254)' },
  { id: 'windows-1257', label: 'Baltic (Windows-1257)' },
  { id: 'shift_jis', label: 'Japanese (Shift-JIS)' },
  { id: 'gbk', label: 'Chinese Simplified (GBK)' },
])

const CJK_ENCODINGS = new Set(['shift_jis', 'gbk', 'euc-kr', 'big5'])
const ENCODING_ALIASES = Object.freeze({
  utf8: 'utf-8',
  'utf-8-bom': 'utf-8',
  'utf-16': 'utf-16le',
  utf16le: 'utf-16le',
  utf16be: 'utf-16be',
  ucs2: 'utf-16le',
  'ucs-2': 'utf-16le',
  latin1: 'windows-1252',
  'iso-8859-1': 'windows-1252',
  ascii: 'windows-1252',
  'us-ascii': 'windows-1252',
  cp1250: 'windows-1250',
  cp1251: 'windows-1251',
  cp1252: 'windows-1252',
  sjis: 'shift_jis',
  'shift-jis': 'shift_jis',
  cp932: 'shift_jis',
  'windows-31j': 'shift_jis',
  cp936: 'gbk',
  gb2312: 'gbk',
})

let cachedAnsiCodePage = null
const reverseTables = new Map()

/**
 * Normalises an encoding name to the label Simple uses everywhere.
 * @param {string} name e.g. "UTF8", "cp1251", "Shift-JIS".
 * @returns {string} A lower-case WHATWG label such as "utf-8" or "windows-1251".
 * @throws {RangeError} When TextDecoder does not know the encoding.
 */
function normalizeEncoding(name) {
  const raw = String(name || '').trim().toLowerCase()
  const label = ENCODING_ALIASES[raw] || raw
  if (label === 'utf-8' || label === 'utf-16le' || label === 'utf-16be') return label
  const canonical = new TextDecoder(label).encoding // throws RangeError for unknown labels
  return canonical === 'gb18030' ? 'gbk' : canonical
}

/**
 * A Windows system program by absolute path (System32), never by bare name:
 * Windows searches the current folder before PATH for a bare name, and the
 * current folder may be one where anyone could have placed a program.
 * @param {string} name e.g. "reg.exe"
 * @returns {string}
 */
function systemProgram(name) {
  const root = [process.env.SystemRoot, process.env.windir].find((value) => typeof value === 'string' && path.win32.isAbsolute(value)) || 'C:\\Windows'
  return path.win32.join(root, 'System32', name)
}

/**
 * The Windows ANSI code page, read once from
 * HKLM\SYSTEM\CurrentControlSet\Control\Nls\CodePage\ACP with System32's reg.exe (1252
 * elsewhere or when the registry cannot be read). SIMPLE_ANSI_CODEPAGE overrides it, which
 * keeps tests independent of the machine.
 * @returns {number}
 */
function systemAnsiCodePage() {
  const override = Number.parseInt(process.env.SIMPLE_ANSI_CODEPAGE || '', 10)
  if (Number.isInteger(override) && override > 0) return override
  if (cachedAnsiCodePage !== null) return cachedAnsiCodePage
  cachedAnsiCodePage = 1252
  if (process.platform === 'win32') {
    try {
      const output = execFileSync(systemProgram('reg.exe'), ['query', 'HKLM\\SYSTEM\\CurrentControlSet\\Control\\Nls\\CodePage', '/v', 'ACP'], {
        encoding: 'utf8', windowsHide: true, timeout: 3000, stdio: ['ignore', 'pipe', 'ignore'],
      })
      const match = /\bACP\s+REG_SZ\s+(\d+)/i.exec(output)
      if (match) cachedAnsiCodePage = Number(match[1])
    } catch {
      // Keep 1252: decoding still succeeds, only accents may differ.
    }
  }
  return cachedAnsiCodePage
}

/**
 * TextDecoder label for a Windows code page number; unknown pages map to
 * windows-1252 and UTF-8 (65001) maps to windows-1252 as well, because the
 * ANSI fallback is only consulted after strict UTF-8 already failed.
 * @param {number} codePage
 * @returns {string}
 */
function labelForCodePage(codePage) {
  const label = CODE_PAGE_LABELS[codePage]
  if (!label || label === 'utf-8') return 'windows-1252'
  return label
}

function toBuffer(bytes) {
  if (Buffer.isBuffer(bytes)) return bytes
  if (bytes instanceof Uint8Array) return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes instanceof ArrayBuffer) return Buffer.from(bytes)
  if (typeof bytes === 'string') return Buffer.from(bytes, 'utf8')
  throw new TypeError('decodeText expects a Buffer, Uint8Array or ArrayBuffer.')
}

function bomOf(buffer) {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) return 'utf-8'
  if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) return 'utf-16le'
  if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) return 'utf-16be'
  return null
}

/**
 * Detects BOM-less UTF-16 by NUL parity in the first 4 KB: ASCII text in
 * UTF-16LE has NULs at odd offsets, UTF-16BE at even offsets.
 * @param {Buffer} buffer
 * @returns {'utf-16le'|'utf-16be'|null}
 */
function detectUtf16(buffer) {
  const sample = Math.min(buffer.length, UTF16_SAMPLE_BYTES) & ~1
  if (sample < 4) return null
  let evenNul = 0
  let oddNul = 0
  for (let index = 0; index < sample; index += 1) {
    if (buffer[index] !== 0) continue
    if (index % 2) oddNul += 1
    else evenNul += 1
  }
  const half = sample / 2
  if (oddNul / half >= UTF16_NUL_RATIO && evenNul / half < 0.1) return 'utf-16le'
  if (evenNul / half >= UTF16_NUL_RATIO && oddNul / half < 0.1) return 'utf-16be'
  return null
}

/** Length of a trailing, incomplete UTF-8 sequence (0–3 bytes). */
function incompleteUtf8Tail(buffer) {
  const end = buffer.length
  for (let back = 1; back <= Math.min(3, end); back += 1) {
    const byte = buffer[end - back]
    if ((byte & 0xc0) === 0x80) continue // continuation byte; keep looking for the lead
    let need = 0
    if ((byte & 0xe0) === 0xc0) need = 2
    else if ((byte & 0xf0) === 0xe0) need = 3
    else if ((byte & 0xf8) === 0xf0) need = 4
    return need > back ? back : 0
  }
  return 0
}

function decodeUtf16(buffer, encoding) {
  const even = buffer.length & ~1
  if (encoding === 'utf-16le') return buffer.toString('utf16le', 0, even)
  const swapped = Buffer.from(buffer.subarray(0, even))
  swapped.swap16()
  return swapped.toString('utf16le')
}

function strictUtf8(buffer) {
  try {
    return new TextDecoder('utf-8', { fatal: true, ignoreBOM: true }).decode(buffer)
  } catch {
    return null
  }
}

function isAscii(buffer) {
  for (let index = 0; index < buffer.length; index += 1) if (buffer[index] > 0x7f) return false
  return true
}

// cp1251 letters outside 0xC0–0xFF: Ё Ў Є Ї І (upper) and ё ў є ї і (lower).
const CP1251_EXTRA_UPPER = new Set([0xa8, 0xa1, 0xaa, 0xaf, 0xb2])
const CP1251_EXTRA_LOWER = new Set([0xb8, 0xa2, 0xba, 0xbf, 0xb3])

/**
 * Scores windows-1251 against a Latin code page. Cyrillic text stored in
 * cp1251 puts nearly every high byte on a letter (А–я, Ё/ё, Ukrainian and
 * Belarusian letters), those bytes form whole words (long runs), and prose is
 * mostly lower case (0xE0–0xFF). Western text in cp1252 has isolated accented
 * letters inside ASCII words; double-byte East Asian text spreads evenly over
 * 0xA1–0xFE and is not "mostly lower case".
 * @returns {boolean}
 */
function looksCyrillic1251(buffer) {
  let high = 0
  let letters = 0
  let lower = 0
  let runs = 0
  let runBytes = 0
  let inRun = false
  for (let index = 0; index < buffer.length; index += 1) {
    const byte = buffer[index]
    if (byte < 0x80) { inRun = false; continue }
    high += 1
    const isLower = byte >= 0xe0 || CP1251_EXTRA_LOWER.has(byte)
    const isLetter = isLower || byte >= 0xc0 || CP1251_EXTRA_UPPER.has(byte)
    if (isLetter) {
      letters += 1
      if (isLower) lower += 1
      runBytes += 1
      if (!inRun) { runs += 1; inRun = true }
    } else {
      inRun = false
    }
  }
  if (high < 4 || !letters) return false
  const letterShare = letters / high
  const averageRun = runs ? runBytes / runs : 0
  const lowerShare = lower / letters
  return letterShare >= 0.8 && averageRun >= 2.5 && lowerShare >= 0.55
}

function decodeWith(label, buffer) {
  return new TextDecoder(label).decode(buffer)
}

function countReplacement(text) {
  let count = 0
  for (let index = text.indexOf('�'); index !== -1; index = text.indexOf('�', index + 1)) count += 1
  return count
}

/**
 * Share of double-byte pairs whose second byte is also ≥ 0xA1. Common GBK
 * (GB2312) and EUC-KR characters use high trail bytes; Western text misread
 * as a double-byte encoding pairs an accented letter with an ASCII letter.
 */
function highTrailShare(buffer) {
  let pairs = 0
  let high = 0
  for (let index = 0; index < buffer.length; index += 1) {
    if (buffer[index] < 0x81) continue
    if (index + 1 >= buffer.length) break
    pairs += 1
    if (buffer[index + 1] >= 0xa1) high += 1
    index += 1
  }
  return pairs ? high / pairs : 0
}

/**
 * Checks whether bytes read as a CJK double-byte encoding produce plausible
 * text: no decoding errors, almost no half-width katakana (a sign of
 * misreading), the script that encoding is used for and, for GBK and EUC-KR,
 * the byte pattern of their common characters.
 */
function plausibleCjk(label, buffer) {
  if ((label === 'gbk' || label === 'euc-kr') && highTrailShare(buffer) < 0.9) return null
  const text = decodeWith(label, buffer)
  if (countReplacement(text) > 0) return null
  let nonAscii = 0
  let kana = 0
  let han = 0
  let hangul = 0
  let halfWidth = 0
  for (const char of text) {
    const code = char.codePointAt(0)
    if (code < 0x80) continue
    nonAscii += 1
    if (code >= 0x3040 && code <= 0x30ff) kana += 1
    else if (code >= 0x4e00 && code <= 0x9fff) han += 1
    else if (code >= 0xac00 && code <= 0xd7af) hangul += 1
    else if (code >= 0xff61 && code <= 0xff9f) halfWidth += 1
  }
  if (nonAscii < 4 || halfWidth / nonAscii > 0.05) return null
  if (label === 'shift_jis' && (kana + han) / nonAscii >= 0.8 && kana / nonAscii >= 0.15) return text
  if ((label === 'gbk' || label === 'big5') && han / nonAscii >= 0.8 && kana === 0) return text
  if (label === 'euc-kr' && hangul / nonAscii >= 0.6) return text
  return null
}

/**
 * Describes the line endings of a string.
 * @param {string} text
 * @returns {{lineEnding: 'crlf'|'lf'|'cr'|null, mixedLineEndings: boolean}}
 *   lineEnding is the dominant style, null when the text has no line break.
 */
function detectLineEnding(text) {
  let crlf = 0
  let lf = 0
  let cr = 0
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    if (code === 13) {
      if (text.charCodeAt(index + 1) === 10) { crlf += 1; index += 1 } else cr += 1
    } else if (code === 10) lf += 1
  }
  const kinds = [['crlf', crlf], ['lf', lf], ['cr', cr]].filter(([, count]) => count > 0)
  if (!kinds.length) return { lineEnding: null, mixedLineEndings: false }
  kinds.sort((a, b) => b[1] - a[1])
  return { lineEnding: kinds[0][0], mixedLineEndings: kinds.length > 1 }
}

/**
 * Decodes file bytes to text and reports the dialect needed to write the file
 * back the same way. Order: BOM → BOM-less UTF-16 → strict UTF-8 → the
 * Windows ANSI code page, with windows-1251 chosen for Cyrillic content and,
 * unless `cjk:false`, Shift-JIS / GBK / EUC-KR when the bytes are only
 * plausible in that encoding (Big5 and the others are used when they are the
 * system code page).
 *
 * @param {Buffer|Uint8Array|ArrayBuffer} bytes
 * @param {object} [options]
 * @param {string} [options.encoding] Force an encoding (a "Reopen with encoding" choice). BOMs are still stripped.
 * @param {boolean} [options.truncated] The bytes are a prefix of a longer file: an incomplete trailing
 *   UTF-8 / UTF-16 sequence is ignored instead of disqualifying the encoding.
 * @param {number} [options.ansiCodePage] Code page to fall back to (defaults to the system ANSI code page).
 * @param {boolean} [options.cjk=true] Consider double-byte East Asian encodings.
 * @returns {{text: string, encoding: string, bom: boolean, lineEnding: 'crlf'|'lf'|'cr'|null,
 *   mixedLineEndings: boolean, confidence: 'certain'|'high'|'medium'|'low', ascii: boolean,
 *   replacements: number, dialect: {encoding: string, bom: boolean, lineEnding: 'crlf'|'lf'|'cr'|null}}}
 *   `replacements` counts U+FFFD produced by decoding (0 for a clean decode).
 */
function decodeText(bytes, options = {}) {
  let buffer = toBuffer(bytes)
  const truncated = Boolean(options.truncated)
  const bomEncoding = bomOf(buffer)
  let encoding
  let confidence
  let text
  let bom = false

  if (options.encoding) {
    encoding = normalizeEncoding(options.encoding)
    const bomMatches = bomEncoding && bomEncoding === encoding
    if (bomMatches) { bom = true; buffer = buffer.subarray(encoding === 'utf-8' ? 3 : 2) }
    confidence = 'certain'
  } else if (bomEncoding) {
    encoding = bomEncoding
    bom = true
    buffer = buffer.subarray(encoding === 'utf-8' ? 3 : 2)
    confidence = 'certain'
  }

  if (!encoding) {
    const utf16 = detectUtf16(buffer)
    if (utf16) { encoding = utf16; confidence = 'high' }
  }

  if (encoding === 'utf-16le' || encoding === 'utf-16be') {
    text = decodeUtf16(buffer, encoding)
    if (truncated && text.length) {
      const last = text.charCodeAt(text.length - 1)
      if (last >= 0xd800 && last <= 0xdbff) text = text.slice(0, -1)
    }
  } else if (encoding === 'utf-8') {
    const usable = truncated ? buffer.subarray(0, buffer.length - incompleteUtf8Tail(buffer)) : buffer
    text = new TextDecoder('utf-8', { ignoreBOM: true }).decode(usable)
  } else if (encoding) {
    text = decodeWith(encoding, buffer)
  } else {
    const usable = truncated ? buffer.subarray(0, buffer.length - incompleteUtf8Tail(buffer)) : buffer
    const utf8 = strictUtf8(usable)
    if (utf8 !== null) {
      encoding = 'utf-8'
      text = utf8
      confidence = isAscii(usable) ? 'high' : 'certain'
    } else {
      const codePage = Number.isInteger(options.ansiCodePage) ? options.ansiCodePage : systemAnsiCodePage()
      const ansi = labelForCodePage(codePage)
      const cjkAllowed = options.cjk !== false
      if (CJK_ENCODINGS.has(ansi)) {
        const native = decodeWith(ansi, buffer)
        if (countReplacement(native) === 0) { encoding = ansi; text = native; confidence = 'medium' }
      }
      if (!encoding && looksCyrillic1251(buffer)) {
        encoding = 'windows-1251'
        confidence = 'medium'
      }
      if (!encoding && cjkAllowed) {
        for (const label of ['shift_jis', 'euc-kr', 'gbk']) {
          const candidate = plausibleCjk(label, buffer)
          if (candidate !== null) { encoding = label; text = candidate; confidence = 'medium'; break }
        }
      }
      if (!encoding) {
        encoding = CJK_ENCODINGS.has(ansi) ? 'windows-1252' : ansi
        confidence = 'low'
      }
      if (text === undefined) text = decodeWith(encoding, buffer)
    }
  }

  const { lineEnding, mixedLineEndings } = detectLineEnding(text)
  const replacements = countReplacement(text)
  return {
    text,
    encoding,
    bom,
    lineEnding,
    mixedLineEndings,
    confidence,
    ascii: encoding === 'utf-8' && !bom && isAscii(buffer),
    replacements,
    dialect: { encoding, bom, lineEnding },
  }
}

/** Builds (once) the char → bytes table for a legacy encoding from TextDecoder. */
function reverseTable(encoding) {
  if (reverseTables.has(encoding)) return reverseTables.get(encoding)
  const table = new Map()
  const decoder = new TextDecoder(encoding)
  const single = Buffer.alloc(1)
  for (let byte = 0; byte < 0x100; byte += 1) {
    single[0] = byte
    const char = decoder.decode(single)
    if (char.length === 1 && char !== '�' && !table.has(char)) table.set(char, Buffer.from([byte]))
  }
  if (CJK_ENCODINGS.has(encoding)) {
    const pair = Buffer.alloc(2)
    for (let lead = 0x81; lead <= 0xfe; lead += 1) {
      for (let trail = 0x40; trail <= 0xfe; trail += 1) {
        pair[0] = lead
        pair[1] = trail
        const char = decoder.decode(pair)
        // Keep the first (canonical) byte sequence for each character.
        if (char && !char.includes('�') && [...char].length === 1 && !table.has(char)) table.set(char, Buffer.from(pair))
      }
    }
  }
  reverseTables.set(encoding, table)
  return table
}

/**
 * Converts every line break to one style.
 * @param {string} text
 * @param {'crlf'|'lf'|'cr'|null|undefined} lineEnding null/undefined keeps the text as is.
 * @returns {string}
 */
function normalizeLineEndings(text, lineEnding) {
  if (!lineEnding) return text
  const target = lineEnding === 'crlf' ? '\r\n' : lineEnding === 'cr' ? '\r' : '\n'
  return text.replace(/\r\n|\r|\n/g, target)
}

/**
 * Encodes text in a dialect, normally the `dialect` that decodeText returned.
 * For legacy code pages, characters the page cannot hold are replaced with "?"
 * and reported, so the caller can ask "{n} characters can't be saved in
 * {encoding}. Save as UTF-8 instead?" before writing.
 *
 * @param {string} text
 * @param {object} [dialect]
 * @param {string} [dialect.encoding='utf-8']
 * @param {boolean} [dialect.bom=false] Write a BOM (UTF-8 / UTF-16 only).
 * @param {'crlf'|'lf'|'cr'|null} [dialect.lineEnding] Normalise line breaks; null keeps them.
 * @param {'replace'|'throw'} [dialect.onUnmappable='replace'] 'throw' raises an Error with
 *   code 'UNMAPPABLE_CHARACTERS' instead of substituting.
 * @returns {{bytes: Buffer, encoding: string, unmappable: number, unmappableSamples: string[]}}
 */
function encodeText(text, dialect = {}) {
  const encoding = normalizeEncoding(dialect.encoding || 'utf-8')
  const source = normalizeLineEndings(String(text ?? ''), dialect.lineEnding)
  const bom = Boolean(dialect.bom)
  if (encoding === 'utf-8') {
    const body = Buffer.from(source, 'utf8')
    return { bytes: bom ? Buffer.concat([UTF8_BOM, body]) : body, encoding, unmappable: 0, unmappableSamples: [] }
  }
  if (encoding === 'utf-16le' || encoding === 'utf-16be') {
    const body = Buffer.from(source, 'utf16le')
    if (encoding === 'utf-16be') body.swap16()
    const prefix = encoding === 'utf-16le' ? UTF16LE_BOM : UTF16BE_BOM
    return { bytes: bom ? Buffer.concat([prefix, body]) : body, encoding, unmappable: 0, unmappableSamples: [] }
  }
  const table = reverseTable(encoding)
  const question = table.get('?') || Buffer.from('?')
  // Every character maps to at most two bytes and takes at least one UTF-16
  // unit, so twice the string length always fits.
  const output = Buffer.allocUnsafe(source.length * 2)
  let length = 0
  let unmappable = 0
  const samples = new Set()
  for (const char of source) {
    let mapped = table.get(char)
    if (!mapped) {
      unmappable += 1
      if (samples.size < 10) samples.add(char)
      mapped = question
    }
    for (let index = 0; index < mapped.length; index += 1) output[length++] = mapped[index]
  }
  if (unmappable && dialect.onUnmappable === 'throw') {
    const error = new Error(`${unmappable} character${unmappable === 1 ? '' : 's'} can't be saved in ${encoding}.`)
    error.code = 'UNMAPPABLE_CHARACTERS'
    error.unmappable = unmappable
    error.unmappableSamples = [...samples]
    throw error
  }
  return { bytes: Buffer.from(output.subarray(0, length)), encoding, unmappable, unmappableSamples: [...samples] }
}

/**
 * Counts characters an encoding cannot represent, without building output.
 * @param {string} text
 * @param {string} encoding
 * @returns {{ok: boolean, unmappable: number, unmappableSamples: string[]}}
 */
function canEncode(text, encoding) {
  const normalized = normalizeEncoding(encoding)
  if (normalized === 'utf-8' || normalized === 'utf-16le' || normalized === 'utf-16be') return { ok: true, unmappable: 0, unmappableSamples: [] }
  const { unmappable, unmappableSamples } = encodeText(text, { encoding: normalized })
  return { ok: unmappable === 0, unmappable, unmappableSamples }
}

module.exports = {
  CODE_PAGE_LABELS,
  WRITABLE_ENCODINGS,
  canEncode,
  decodeText,
  detectLineEnding,
  detectUtf16,
  encodeText,
  labelForCodePage,
  normalizeEncoding,
  normalizeLineEndings,
  systemAnsiCodePage,
}
