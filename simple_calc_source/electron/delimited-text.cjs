'use strict'

// Delimited text (CSV / TSV / TXT) reading and writing.
//
// Reading: encoding detection (BOMs, BOM-less UTF-16, legacy code pages), delimiter sniffing,
// RFC 4180 records with Excel's lenient quote rules, per-column number and date inference, and
// the source dialect, so that saving writes the file back the way it came.
//
// Writing: RFC 4180 quoting (embedded line breaks included), 4-digit years, values written as
// displayed, and the recorded dialect (delimiter, decimal comma, grouping, encoding, BOM, line
// endings).

const XLSX = require('xlsx')

const DELIMITER_CANDIDATES = [',', ';', '\t', '|']
const MAX_ROWS = 1_048_576
const MAX_COLS = 16_384
const INFERENCE_ROW_LIMIT = 200_000
const SPACE_GROUPS = new Set([' ', ' ', ' ', "'", '’'])
const CURRENCY_SYMBOLS = '$€£¥₹₽₩₺₴₸₦₱₫฿₪'
const ERROR_TOKENS = new Set(['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CONNECT!', '#BLOCKED!', '#UNKNOWN!', '#FIELD!', '#CALC!'])
const MULTI_BYTE_ENCODINGS = new Set(['shift_jis', 'euc-jp', 'euc-kr', 'gb18030', 'gbk', 'big5'])

// ---------------------------------------------------------------------------
// Locale hints (the system locale stands in for Excel's regional settings)
// ---------------------------------------------------------------------------

function defaultLocale() {
  try { return Intl.DateTimeFormat().resolvedOptions().locale || 'en-US' } catch { return 'en-US' }
}

function localeDecimalSeparator(locale) {
  try {
    const part = new Intl.NumberFormat(locale).formatToParts(1.5).find((item) => item.type === 'decimal')
    return part && part.value === ',' ? ',' : '.'
  } catch {
    return '.'
  }
}

function localeDayFirst(locale) {
  try {
    const parts = new Intl.DateTimeFormat(locale, { year: 'numeric', month: '2-digit', day: '2-digit' }).formatToParts(new Date(2024, 2, 15))
    const types = parts.map((part) => part.type).filter((type) => type === 'day' || type === 'month' || type === 'year')
    return types.indexOf('day') < types.indexOf('month')
  } catch {
    return false
  }
}

/** Windows' ANSI code page for a locale (what Excel uses for a CSV without a BOM). */
function ansiEncodingFor(locale) {
  const tag = String(locale || '').toLowerCase()
  const language = tag.split(/[-_]/)[0]
  if (language === 'ja') return 'shift_jis'
  if (language === 'ko') return 'euc-kr'
  if (language === 'zh') return /hant|-tw|-hk|-mo/.test(tag) ? 'big5' : 'gb18030'
  if (language === 'uz') return /cyrl/.test(tag) ? 'windows-1251' : 'windows-1254'
  if (language === 'sr') return /latn/.test(tag) ? 'windows-1250' : 'windows-1251'
  if (['ru', 'uk', 'be', 'bg', 'mk', 'kk', 'ky', 'mn', 'tg', 'tt', 'ba'].includes(language)) return 'windows-1251'
  if (['pl', 'cs', 'sk', 'hu', 'ro', 'hr', 'sl', 'bs', 'sq'].includes(language)) return 'windows-1250'
  if (language === 'el') return 'windows-1253'
  if (['tr', 'az'].includes(language)) return 'windows-1254'
  if (language === 'he') return 'windows-1255'
  if (['ar', 'fa', 'ur'].includes(language)) return 'windows-1256'
  if (['et', 'lv', 'lt'].includes(language)) return 'windows-1257'
  if (language === 'vi') return 'windows-1258'
  if (language === 'th') return 'windows-874'
  return 'windows-1252'
}

// ---------------------------------------------------------------------------
// Encoding detection
// ---------------------------------------------------------------------------

function tryDecode(bytes, label, fatal = true) {
  try {
    return new TextDecoder(label, { fatal }).decode(bytes)
  } catch {
    return null
  }
}

function plausibleCjk(text, label) {
  let nonAscii = 0
  let cjk = 0
  let hangul = 0
  let halfwidth = 0
  for (const character of text) {
    const code = character.codePointAt(0)
    if (code < 0x80) continue
    nonAscii += 1
    if (code >= 0xac00 && code <= 0xd7a3) hangul += 1
    else if ((code >= 0x3000 && code <= 0x30ff) || (code >= 0x4e00 && code <= 0x9fff) || (code >= 0x3400 && code <= 0x4dbf) || (code >= 0xff01 && code <= 0xff60)) cjk += 1
    else if (code >= 0xff61 && code <= 0xff9f) halfwidth += 1
  }
  if (!nonAscii) return false
  if (label === 'euc-kr') return hangul / nonAscii >= 0.5 && (hangul + cjk) / nonAscii >= 0.9
  if (label === 'shift_jis' || label === 'euc-jp') return (cjk + halfwidth) / nonAscii >= 0.9 && halfwidth / nonAscii <= 0.3
  return cjk / nonAscii >= 0.9
}

function guessLegacyEncoding(buffer, locale) {
  const sample = buffer.subarray(0, Math.min(buffer.length, 65_536))
  let high = 0
  let upper = 0
  let runs = 0
  let runTotal = 0
  let run = 0
  for (const byte of sample) {
    if (byte >= 0x80) {
      high += 1
      run += 1
      if (byte >= 0xc0) upper += 1
    } else if (run) {
      runs += 1
      runTotal += run
      run = 0
    }
  }
  if (run) { runs += 1; runTotal += run }
  const preferred = ansiEncodingFor(locale)
  if (!high) return preferred
  const averageRun = runTotal / Math.max(1, runs)
  // Multi-byte East Asian text: try the strict decoders. The permissive GB18030/Big5 tables
  // accept almost any byte pairs, so they are only tried on a Chinese system.
  const candidates = [...new Set([MULTI_BYTE_ENCODINGS.has(preferred) ? preferred : null, 'shift_jis', 'euc-kr', 'euc-jp'].filter(Boolean))]
  if (averageRun >= 1.6 || MULTI_BYTE_ENCODINGS.has(preferred)) {
    for (const label of candidates) {
      let text = tryDecode(sample, label)
      // A sample cut from a longer file may end inside a character.
      for (let cut = 1; text == null && cut <= 3 && sample.length < buffer.length; cut += 1) text = tryDecode(sample.subarray(0, sample.length - cut), label)
      if (text != null && plausibleCjk(text, label)) return label
    }
  }
  // Cyrillic text in windows-1251 is made of whole words of high bytes; accented Latin letters
  // in windows-1252/1250 are single high bytes inside ASCII words.
  if (averageRun >= 2.5 && upper / high >= 0.6) return 'windows-1251'
  if (MULTI_BYTE_ENCODINGS.has(preferred)) return 'windows-1252'
  if (preferred === 'windows-1251' && averageRun < 1.5) return 'windows-1252'
  return preferred
}

/** Decode delimited text. Returns { text, encoding, hadBom }. */
function decodeText(buffer, locale = defaultLocale()) {
  if (buffer.length >= 3 && buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf) {
    return { text: buffer.subarray(3).toString('utf8'), encoding: 'utf-8', hadBom: true }
  }
  if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
    return { text: buffer.subarray(2).toString('utf16le'), encoding: 'utf-16le', hadBom: true }
  }
  if (buffer.length >= 2 && buffer[0] === 0xfe && buffer[1] === 0xff) {
    return { text: Buffer.from(buffer.subarray(2)).swap16().toString('utf16le'), encoding: 'utf-16be', hadBom: true }
  }
  // BOM-less UTF-16: Latin text has a NUL in every other byte.
  const probe = buffer.subarray(0, Math.min(buffer.length, 4096))
  const pairs = Math.floor(probe.length / 2)
  if (pairs >= 2) {
    let evenZero = 0
    let oddZero = 0
    for (let index = 0; index < pairs; index += 1) {
      if (probe[index * 2] === 0) evenZero += 1
      if (probe[index * 2 + 1] === 0) oddZero += 1
    }
    const usable = buffer.length - (buffer.length % 2)
    if (oddZero / pairs >= 0.3 && evenZero / pairs <= 0.05) return { text: buffer.subarray(0, usable).toString('utf16le'), encoding: 'utf-16le', hadBom: false }
    if (evenZero / pairs >= 0.3 && oddZero / pairs <= 0.05) return { text: Buffer.from(buffer.subarray(0, usable)).swap16().toString('utf16le'), encoding: 'utf-16be', hadBom: false }
  }
  const utf8 = tryDecode(buffer, 'utf-8')
  if (utf8 != null) return { text: utf8, encoding: 'utf-8', hadBom: false }
  const encoding = guessLegacyEncoding(buffer, locale)
  const text = tryDecode(buffer, encoding, false)
  if (text != null) return { text, encoding, hadBom: false }
  return { text: buffer.toString('latin1'), encoding: 'windows-1252', hadBom: false }
}

// ---------------------------------------------------------------------------
// Records
// ---------------------------------------------------------------------------

/**
 * Index just past the closing quote of a quoted field that starts at `start` (the opening
 * quote), or -1 when the field is not properly closed before a delimiter / line break.
 */
function quotedFieldEnd(text, start, delimiter) {
  for (let index = start + 1; index < text.length; index += 1) {
    if (text[index] !== '"') continue
    if (text[index + 1] === '"') { index += 1; continue }
    let next = index + 1
    while (text[next] === ' ' || (text[next] === '\t' && delimiter !== '\t')) next += 1
    const after = text[next]
    return after === undefined || after === delimiter || after === '\n' || after === '\r' ? index + 1 : -1
  }
  return -1
}

/** Count delimiters outside quoted sections; a quote opens a section only at the start of a field. */
function countDelimitersOutsideQuotes(line, delimiter) {
  let count = 0
  let quoted = false
  let atFieldStart = true
  for (let index = 0; index < line.length; index += 1) {
    const character = line[index]
    if (quoted) {
      if (character === '"') {
        if (line[index + 1] === '"') index += 1
        else quoted = false
      }
      continue
    }
    if (character === '"' && atFieldStart) { quoted = true; atFieldStart = false; continue }
    if (character === delimiter) { count += 1; atFieldStart = true; continue }
    atFieldStart = false
  }
  return count
}

function sniffDelimiter(text, extensionDefault) {
  const lines = text.slice(0, 65_536).split(/\r\n|[\r\n]/).filter((line) => line.trim() !== '').slice(0, 50)
  if (!lines.length) return extensionDefault
  let best
  for (const delimiter of DELIMITER_CANDIDATES) {
    const counts = lines.map((line) => countDelimitersOutsideQuotes(line, delimiter))
    const tally = new Map()
    for (const count of counts) if (count > 0) tally.set(count, (tally.get(count) || 0) + 1)
    let mode = 0
    let modeLines = 0
    for (const [count, occurrences] of tally) {
      if (occurrences > modeLines || (occurrences === modeLines && count > mode)) {
        mode = count
        modeLines = occurrences
      }
    }
    if (!mode) continue
    const consistency = modeLines / counts.length
    if (consistency < 0.5) continue
    const score = consistency * 1_000_000 + (delimiter === extensionDefault ? 1_000 : 0) + Math.min(mode, 999)
    if (!best || score > best.score) best = { delimiter, score }
  }
  return best ? best.delimiter : extensionDefault
}

/**
 * Parse delimited records. A double quote starts a quoted field only at the start of a field
 * (optionally after spaces, when the quoted section closes cleanly); anywhere else it is a
 * literal character, as in Excel. Text after a closing quote is kept literally.
 */
function parseRecords(text, delimiter) {
  const records = []
  let record = []
  let field = ''
  let index = 0
  let atFieldStart = true
  const length = text.length
  while (index < length) {
    const character = text[index]
    if (atFieldStart) {
      let quoteAt = -1
      if (character === '"') quoteAt = index
      else if (character === ' ') {
        let probe = index
        while (text[probe] === ' ') probe += 1
        if (text[probe] === '"' && quotedFieldEnd(text, probe, delimiter) > 0) quoteAt = probe
      }
      if (quoteAt >= 0) {
        index = quoteAt + 1
        // Quoted section: "" is a literal quote; the section ends at the next lone quote.
        while (index < length) {
          const inner = text[index]
          if (inner === '"') {
            if (text[index + 1] === '"') { field += '"'; index += 2; continue }
            index += 1
            break
          }
          field += inner
          index += 1
        }
        atFieldStart = false
        // Excel tolerates spaces between a closing quote and the delimiter.
        let probe = index
        while (text[probe] === ' ') probe += 1
        if (probe >= length || text[probe] === delimiter || text[probe] === '\n' || text[probe] === '\r') index = probe
        continue
      }
    }
    atFieldStart = false
    if (character === delimiter) {
      record.push(field)
      field = ''
      atFieldStart = true
      index += 1
    } else if (character === '\n' || character === '\r') {
      if (character === '\r' && text[index + 1] === '\n') index += 1
      record.push(field)
      records.push(record)
      field = ''
      record = []
      atFieldStart = true
      index += 1
    } else {
      field += character
      index += 1
    }
  }
  if (field !== '' || record.length) {
    record.push(field)
    records.push(record)
  }
  return records
}

/** Line ending used outside quoted fields (\r\n, \n or \r) and whether the text ends with one. */
function detectLineEnding(text) {
  let crlf = 0
  let lf = 0
  let cr = 0
  let quoted = false
  let atFieldStart = true
  const limit = Math.min(text.length, 2_000_000)
  for (let index = 0; index < limit; index += 1) {
    const character = text[index]
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') index += 1
        else quoted = false
      }
      continue
    }
    if (character === '"' && atFieldStart) { quoted = true; atFieldStart = false; continue }
    if (character === '\r') {
      if (text[index + 1] === '\n') { crlf += 1; index += 1 } else cr += 1
      atFieldStart = true
      continue
    }
    if (character === '\n') { lf += 1; atFieldStart = true; continue }
    atFieldStart = character === ',' || character === ';' || character === '\t' || character === '|'
  }
  const lineEnding = lf > crlf && lf >= cr ? '\n' : cr > crlf && cr > lf ? '\r' : '\r\n'
  return { lineEnding, trailingNewline: /[\r\n]$/.test(text) }
}

// ---------------------------------------------------------------------------
// Numbers
// ---------------------------------------------------------------------------

/**
 * Parse a numeric token under one convention: 'dot' (1,234.5) or 'comma' (1.234,5). Group
 * separators may also be spaces (incl. no-break spaces) or apostrophes. Returns null when the
 * token is not a valid number under that convention.
 */
const SIMPLE_NUMBER = /^([+-]?)(\d+)(?:([.,])(\d+))?$/
const GROUPED_NUMBER = /^([+-]?)([1-9]\d{0,2})((?:[,.'’   ]\d{3})+)(?:([.,])(\d+))?$/

function parseNumber(rawText, convention) {
  let text = String(rawText).trim()
  if (!text) return null
  // Fast path for the common shapes: 42, -1.5, 2,25, 1.250 (no currency, percent or grouping runs).
  const simple = SIMPLE_NUMBER.exec(text)
  if (simple) {
    const integerDigits = simple[2]
    const separator = simple[3]
    const fractionDigits = simple[4]
    const sign = simple[1] === '-' ? -1 : 1
    if (!separator) return { value: sign * Number(integerDigits), decimals: 0, trailingZero: false, grouped: false, groupChar: null, percent: false, currency: null, paren: false, exponent: false }
    if (separator === (convention === 'comma' ? ',' : '.')) {
      return { value: sign * Number(`${integerDigits}.${fractionDigits}`), decimals: fractionDigits.length, trailingZero: fractionDigits.endsWith('0'), grouped: false, groupChar: null, percent: false, currency: null, paren: false, exponent: false }
    }
    if (fractionDigits.length === 3 && integerDigits.length <= 3 && integerDigits[0] !== '0') {
      return { value: sign * Number(`${integerDigits}${fractionDigits}`), decimals: 0, trailingZero: false, grouped: true, groupChar: separator, percent: false, currency: null, paren: false, exponent: false }
    }
    return null
  }
  // Fast path for plain grouped numbers: 1,234,567 / 1.234,56 / 12 500,5.
  const plainGrouped = GROUPED_NUMBER.exec(text)
  if (plainGrouped) {
    const groups = plainGrouped[3]
    const groupChar = groups[0]
    for (let index = 4; index < groups.length; index += 4) if (groups[index] !== groupChar) return null
    if (groupChar !== (convention === 'comma' ? '.' : ',') && !SPACE_GROUPS.has(groupChar)) return null
    const fraction = plainGrouped[5] || ''
    if (plainGrouped[4] && plainGrouped[4] !== (convention === 'comma' ? ',' : '.')) return null
    const digits = `${plainGrouped[2]}${groups.replace(/[^\d]/g, '')}`
    const groupedSign = plainGrouped[1] === '-' ? -1 : 1
    return { value: groupedSign * Number(`${digits}${fraction ? `.${fraction}` : ''}`), decimals: fraction.length, trailingZero: fraction.endsWith('0'), grouped: true, groupChar, percent: false, currency: null, paren: false, exponent: false }
  }
  let paren = false
  const parenthesized = /^\((.+)\)$/.exec(text)
  if (parenthesized) { paren = true; text = parenthesized[1].trim() }
  let percent = false
  const percentMatch = /^(.*?)(\s?)%$/.exec(text)
  if (percentMatch) { percent = true; text = percentMatch[1].trim() }
  let sign = 1
  let currency = null
  const signOf = (value) => (value === '-' || value === '−' ? -1 : 1)
  let match = /^([-+−])\s*(.*)$/.exec(text)
  if (match) { sign = signOf(match[1]); text = match[2] }
  if (text && CURRENCY_SYMBOLS.includes(text[0])) {
    const symbol = text[0]
    let rest = text.slice(1)
    const spaced = /^\s/.test(rest)
    rest = rest.trim()
    match = /^([-+−])\s*(.*)$/.exec(rest)
    if (match) {
      if (sign === -1) return null
      sign = signOf(match[1])
      rest = match[2]
    }
    currency = { symbol, prefix: true, spaced }
    text = rest
  } else if (text && CURRENCY_SYMBOLS.includes(text[text.length - 1])) {
    const symbol = text[text.length - 1]
    const rest = text.slice(0, -1)
    currency = { symbol, prefix: false, spaced: /\s$/.test(rest) }
    text = rest.trim()
  }
  if (!text) return null
  let exponent = ''
  const exponentMatch = /^(.*\d)[eE]([-+]?\d{1,3})$/.exec(text)
  if (exponentMatch) { text = exponentMatch[1]; exponent = exponentMatch[2] }
  const decimalSeparator = convention === 'comma' ? ',' : '.'
  const groupSeparator = convention === 'comma' ? '.' : ','
  const decimalAt = text.indexOf(decimalSeparator)
  if (decimalAt !== text.lastIndexOf(decimalSeparator)) return null
  const integerPart = decimalAt >= 0 ? text.slice(0, decimalAt) : text
  const fraction = decimalAt >= 0 ? text.slice(decimalAt + 1) : ''
  if (fraction && !/^\d+$/.test(fraction)) return null
  if (!integerPart && !fraction) return null
  let digits
  let grouped = false
  let groupChar = null
  if (/^\d*$/.test(integerPart)) {
    digits = integerPart || '0'
  } else {
    // Exactly one kind of group separator, first group 1-3 digits without a leading zero.
    const separators = new Set(integerPart.replace(/\d/g, ''))
    if (separators.size !== 1) return null
    const separator = [...separators][0]
    if (separator !== groupSeparator && !SPACE_GROUPS.has(separator)) return null
    const groups = integerPart.split(separator)
    if (!/^[1-9]\d{0,2}$/.test(groups[0]) || groups.slice(1).some((group) => !/^\d{3}$/.test(group))) return null
    digits = groups.join('')
    grouped = true
    groupChar = separator
  }
  if (exponent && grouped) return null
  const value = Number(`${digits}${fraction ? `.${fraction}` : ''}${exponent ? `e${exponent}` : ''}`)
  if (!Number.isFinite(value)) return null
  return {
    value: sign * (paren ? -1 : 1) * (percent ? value / 100 : value),
    decimals: fraction.length,
    trailingZero: /0$/.test(fraction),
    grouped,
    groupChar,
    percent,
    currency,
    paren,
    exponent: Boolean(exponent),
  }
}

/** Number format that shows the value the way the source text did. */
function numberFormatFor(parsed) {
  const decimals = Math.min(parsed.decimals, 15)
  if (parsed.exponent) return `0${decimals ? `.${'0'.repeat(decimals)}` : ''}E+00`
  let body = parsed.grouped ? '#,##0' : '0'
  if (decimals) body += `.${'0'.repeat(decimals)}`
  if (parsed.percent) return parsed.paren ? `${body}%;(${body}%)` : `${body}%`
  if (parsed.currency) {
    const symbol = `"${parsed.currency.symbol}"`
    const space = parsed.currency.spaced ? ' ' : ''
    const positive = parsed.currency.prefix ? `${symbol}${space}${body}` : `${body}${space}${symbol}`
    return parsed.paren ? `${positive};(${positive})` : positive
  }
  if (parsed.paren) return `${body};(${body})`
  if (parsed.grouped || parsed.trailingZero) return body
  return undefined
}

/** 'dot' | 'comma' | 'neutral' | 'ambiguous' | null for a numeric-looking token. */
function numberEvidence(text, delimiter) {
  // Fast path for 42 / 1.5 / 2,25 / 1.250 without building parse results.
  const simple = SIMPLE_NUMBER.exec(text)
  if (simple) {
    const separator = simple[3]
    if (!separator) return 'neutral'
    const integerDigits = simple[2]
    const groupable = simple[4].length === 3 && integerDigits.length <= 3 && integerDigits[0] !== '0'
    if (separator === '.') return groupable ? 'ambiguous' : 'dot'
    if (groupable) return 'ambiguous'
    return delimiter === ',' ? 'comma-weak' : 'comma'
  }
  // Words are not numbers (an exponent "E" and currency symbols are not letters here).
  if (!/\d/.test(text) || /[A-DF-Za-df-z]/.test(text)) return null
  const dot = parseNumber(text, 'dot')
  const comma = parseNumber(text, 'comma')
  if (!dot && !comma) return null
  if (dot && comma) return dot.value === comma.value ? 'neutral' : 'ambiguous'
  // In a comma-delimited file an "n,n" field can only come from a quoted field, where it is as
  // likely a text pair ("6,8") as a decimal; only the grouped form (1.234,56) is evidence.
  if (comma && delimiter === ',' && !/\.\d{3}/.test(text)) return 'comma-weak'
  return dot ? 'dot' : 'comma'
}

// ---------------------------------------------------------------------------
// Dates and times
// ---------------------------------------------------------------------------

const ISO_DATE_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T ](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
const YMD_SLASH_RE = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/
const DMY_RE = /^(\d{1,2})([/.-])(\d{1,2})\2(\d{4}|\d{2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([AaPp][Mm]))?)?$/
const TIME_RE = /^(\d{1,2}):(\d{2})(?::(\d{2}))?(?:\s*([AaPp][Mm]))?$/

function dateSerial(year, month, day, hours = 0, minutes = 0, seconds = 0) {
  if (year < 1900 || year > 9999 || month < 1 || month > 12 || day < 1 || day > 31) return null
  if (hours > 23 || minutes > 59 || seconds > 59) return null
  const utc = Date.UTC(year, month - 1, day, hours, minutes, seconds)
  const check = new Date(utc)
  if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null
  return (utc - Date.UTC(1899, 11, 30)) / 86_400_000
}

/** Excel's two-digit year window: 00-29 -> 2000-2029, 30-99 -> 1930-1999. */
function expandYear(text) {
  const year = Number(text)
  if (text.length > 2) return year
  return year < 30 ? 2000 + year : 1900 + year
}

function clockHours(hours, meridiem) {
  if (!meridiem) return hours
  if (hours < 1 || hours > 12) return NaN
  const pm = /^p/i.test(meridiem)
  return (hours % 12) + (pm ? 12 : 0)
}

function timeFormat(hoursText, secondsText, meridiem) {
  const hours = meridiem ? 'h' : hoursText.length === 2 ? 'hh' : 'h'
  return `${hours}:mm${secondsText != null ? ':ss' : ''}${meridiem ? ' AM/PM' : ''}`
}

/** Day/month evidence of a d/m/y-shaped token: 'dmy' | 'mdy' | 'ambiguous' | null (not a date). */
function dateEvidence(text) {
  if (text.length < 6 || !/^\d{1,2}[/.-]/.test(text)) return null
  const match = DMY_RE.exec(text)
  if (!match) return null
  if (match[2] === '-' && match[4].length !== 4) return null
  const first = Number(match[1])
  const second = Number(match[3])
  if (first > 12 && second > 12) return null
  if (first > 12) return 'dmy'
  if (second > 12) return 'mdy'
  return match[2] === '.' ? 'dot' : 'ambiguous'
}

function parseDayMonthYear(text, order) {
  const match = DMY_RE.exec(text)
  if (!match) return null
  const [, a, separator, b, yearText, hoursText, minutesText, secondsText, meridiem] = match
  if (separator === '-' && yearText.length !== 4) return null
  const day = Number(order === 'dmy' ? a : b)
  const month = Number(order === 'dmy' ? b : a)
  const hours = hoursText != null ? clockHours(Number(hoursText), meridiem) : 0
  const serial = dateSerial(expandYear(yearText), month, day, hours, Number(minutesText || 0), Number(secondsText || 0))
  if (serial == null) return null
  const dayText = order === 'dmy' ? a : b
  const monthText = order === 'dmy' ? b : a
  const dayToken = dayText.length === 2 ? 'dd' : 'd'
  const monthToken = monthText.length === 2 ? 'mm' : 'm'
  const yearToken = yearText.length === 4 ? 'yyyy' : 'yy'
  const join = separator === '.' ? '\\.' : separator === '-' ? '\\-' : '/'
  const date = order === 'dmy' ? `${dayToken}${join}${monthToken}${join}${yearToken}` : `${monthToken}${join}${dayToken}${join}${yearToken}`
  return { value: serial, numFmt: hoursText != null ? `${date} ${timeFormat(hoursText, secondsText, meridiem)}` : date }
}

// ---------------------------------------------------------------------------
// Inference
// ---------------------------------------------------------------------------

function columnName(index) {
  let value = index + 1
  let result = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    result = String.fromCharCode(65 + remainder) + result
    value = Math.floor((value - 1) / 26)
  }
  return result
}

function isPlainText(trimmed) {
  if (!trimmed) return true
  if (trimmed.startsWith('=') && trimmed.length > 1) return true
  if (trimmed === 'TRUE' || trimmed === 'FALSE' || ERROR_TOKENS.has(trimmed)) return true
  return /^0\d+$/.test(trimmed) || /^-?\d{16,}$/.test(trimmed)
}

function majority(dot, comma) {
  if (dot > comma) return 'dot'
  if (comma > dot) return 'comma'
  return null
}

/**
 * Decide, per column, how ambiguous numbers (1.250, 2,000) and dates (01/02/2024) are read.
 * Evidence comes from unambiguous tokens in the same column first, then from the rest of the
 * file, then from the delimiter and the system locale. A column whose only evidence is weak and
 * uncorroborated keeps its ambiguous values as text instead of guessing a 1000x scale.
 */
function inferColumns(records, delimiter, locale) {
  const columns = []
  const limit = Math.min(records.length, INFERENCE_ROW_LIMIT)
  for (let row = 0; row < limit; row += 1) {
    const record = records[row]
    for (let col = 0; col < Math.min(record.length, MAX_COLS); col += 1) {
      const trimmed = record[col].trim()
      if (isPlainText(trimmed)) continue
      if (!/\d/.test(trimmed)) continue
      const stats = columns[col] || (columns[col] = { dot: 0, comma: 0, commaWeak: 0, ambiguous: 0, dmy: 0, mdy: 0, dates: 0 })
      if ((trimmed[4] === '-' && ISO_DATE_RE.test(trimmed)) || (trimmed[4] === '/' && YMD_SLASH_RE.test(trimmed))) continue
      const date = dateEvidence(trimmed)
      if (date) {
        stats.dates += 1
        if (date === 'dmy') stats.dmy += 1
        else if (date === 'mdy') stats.mdy += 1
        continue
      }
      const evidence = numberEvidence(trimmed, delimiter)
      if (evidence === 'dot') stats.dot += 1
      else if (evidence === 'comma') stats.comma += 1
      else if (evidence === 'comma-weak') stats.commaWeak += 1
      else if (evidence === 'ambiguous') stats.ambiguous += 1
    }
  }
  let fileDot = 0
  let fileComma = 0
  let fileDmy = 0
  let fileMdy = 0
  for (const stats of columns) {
    if (!stats) continue
    fileDot += stats.dot
    fileComma += stats.comma
    fileDmy += stats.dmy
    fileMdy += stats.mdy
  }
  // Without any evidence: a comma-delimited file cannot use decimal commas, a semicolon is the
  // separator of decimal-comma locales, and tab/pipe files follow the system locale.
  const localeComma = localeDecimalSeparator(locale) === ','
  const tieBreak = delimiter === ',' ? 'dot' : delimiter === ';' ? 'comma' : localeComma ? 'comma' : 'dot'
  const fileDirection = majority(fileDot, fileComma)
  const decisions = columns.map((stats) => {
    if (!stats) return { numbers: fileDirection || tieBreak, ambiguousAsText: false, dates: null }
    const own = majority(stats.dot, stats.comma)
    let numbers
    let ambiguousAsText = false
    if (own) {
      const ownCount = own === 'dot' ? stats.dot : stats.comma
      const others = majority(fileDot - stats.dot, fileComma - stats.comma)
      if (ownCount >= stats.ambiguous || others === own) numbers = own
      else { numbers = own; ambiguousAsText = stats.ambiguous > 0 }
    } else if (stats.dot && stats.comma) {
      numbers = fileDirection || tieBreak
      ambiguousAsText = stats.ambiguous > 0
    } else {
      numbers = fileDirection || tieBreak
    }
    let dates = null
    if (stats.dmy && !stats.mdy) dates = 'dmy'
    else if (stats.mdy && !stats.dmy) dates = 'mdy'
    else if (stats.dmy || stats.mdy) dates = stats.dmy >= stats.mdy ? 'dmy' : 'mdy'
    return { numbers, ambiguousAsText, dates }
  })
  const fileDates = fileDmy && !fileMdy ? 'dmy' : fileMdy && !fileDmy ? 'mdy' : fileDmy || fileMdy ? (fileDmy >= fileMdy ? 'dmy' : 'mdy') : null
  const decimalComma = (fileDirection || tieBreak) === 'comma'
  const defaultDates = fileDates || (decimalComma || localeDayFirst(locale) ? 'dmy' : 'mdy')
  return { decisions, fileDirection, decimalComma, defaultDates, columns }
}

/** The model cell for one field, given its column's decisions. */
function fieldToCell(field, decision, defaultDates, delimiter) {
  if (field === '') return undefined
  const trimmed = field.trim()
  if (!trimmed) return { value: field, display: field }
  if (trimmed.startsWith('=') && trimmed.length > 1) return { formula: trimmed.slice(1) }
  if (trimmed === 'TRUE' || trimmed === 'FALSE') return { value: trimmed === 'TRUE', display: field }
  if (ERROR_TOKENS.has(trimmed)) return { value: trimmed, type: 'error', display: field }
  // Leading-zero identifiers and digit runs beyond double precision must stay text; coercing
  // them to numbers silently corrupts the source data.
  if (/^0\d+$/.test(trimmed) || /^-?\d{16,}$/.test(trimmed)) return { value: trimmed, display: field, numFmt: '@' }
  if (!/\d/.test(trimmed)) return { value: field, display: field }
  const iso = trimmed[4] === '-' ? ISO_DATE_RE.exec(trimmed) : null
  if (iso) {
    const serial = dateSerial(Number(iso[1]), Number(iso[2]), Number(iso[3]), Number(iso[4] || 0), Number(iso[5] || 0), Number(iso[6] || 0))
    if (serial != null) {
      const numFmt = iso[6] != null ? 'yyyy-mm-dd hh:mm:ss' : iso[4] != null ? 'yyyy-mm-dd hh:mm' : 'yyyy-mm-dd'
      return { value: serial, display: field, numFmt }
    }
  }
  const ymd = trimmed[4] === '/' ? YMD_SLASH_RE.exec(trimmed) : null
  if (ymd) {
    const serial = dateSerial(Number(ymd[1]), Number(ymd[2]), Number(ymd[3]))
    if (serial != null) return { value: serial, display: field, numFmt: `yyyy/${ymd[2].length === 2 ? 'mm' : 'm'}/${ymd[3].length === 2 ? 'dd' : 'd'}` }
  }
  const evidence = dateEvidence(trimmed)
  if (evidence) {
    const order = evidence === 'dmy' || evidence === 'mdy' ? evidence : evidence === 'dot' ? (decision && decision.dates === 'mdy' ? 'mdy' : 'dmy') : (decision && decision.dates) || defaultDates
    const parsed = parseDayMonthYear(trimmed, order)
    if (parsed) return { value: parsed.value, display: field, numFmt: parsed.numFmt }
    return { value: field, display: field }
  }
  const time = trimmed.includes(':') ? TIME_RE.exec(trimmed) : null
  if (time) {
    const hours = clockHours(Number(time[1]), time[4])
    const minutes = Number(time[2])
    const seconds = Number(time[3] || 0)
    if (Number.isFinite(hours) && hours <= 23 && minutes <= 59 && seconds <= 59) {
      return { value: (hours * 3600 + minutes * 60 + seconds) / 86_400, display: field, numFmt: timeFormat(time[1], time[3], time[4]) }
    }
  }
  const kind = numberEvidence(trimmed, delimiter)
  if (!kind) return { value: field, display: field }
  const convention = decision ? decision.numbers : 'dot'
  if (kind === 'ambiguous' && decision && decision.ambiguousAsText) return { value: field, display: field }
  if (kind === 'comma-weak' && convention !== 'comma') return { value: field, display: field }
  const parsed = parseNumber(trimmed, kind === 'dot' ? 'dot' : kind === 'comma' || kind === 'comma-weak' ? 'comma' : convention)
  if (!parsed) return { value: field, display: field }
  const cell = { value: parsed.value, display: field }
  const numFmt = numberFormatFor(parsed)
  if (numFmt) cell.numFmt = numFmt
  return cell
}

/**
 * Read delimited text into cells plus the source dialect. Returns undefined for content the
 * SheetJS readers handle better (HTML/XML tables, SYLK, DIF).
 */
function readDelimited(buffer, options = {}) {
  const locale = options.locale || defaultLocale()
  const { text, encoding, hadBom } = decodeText(buffer, locale)
  const head = text.slice(0, 2_048).replace(/^[\s﻿]+/, '')
  if (/^</.test(head) || /^ID;P/i.test(head) || /^TABLE\r?\n/.test(head)) return undefined
  const format = options.sourceFormat
  const extensionDefault = format === 'csv' ? ',' : format === 'prn' ? undefined : '\t'
  const delimiter = sniffDelimiter(text, extensionDefault)
  if (!delimiter) return undefined
  const records = parseRecords(text, delimiter)
  const { decisions, decimalComma, defaultDates } = inferColumns(records, delimiter, locale)
  const onCell = typeof options.onCell === 'function' ? options.onCell : null
  const columnLabels = []
  const cells = {}
  const notes = []
  const textColumns = new Set()
  const groupChars = new Map()
  let maxRow = 0
  let maxCol = 0
  let truncated = false
  const lengths = new Set()
  records.forEach((record, rowIndex) => {
    if (rowIndex >= MAX_ROWS) { truncated = true; return }
    // Blank lines do not make a file ragged.
    if (!(record.length === 1 && record[0] === '')) lengths.add(record.length)
    record.forEach((field, colIndex) => {
      if (colIndex >= MAX_COLS) { truncated = true; return }
      const decision = decisions[colIndex]
      const cell = fieldToCell(field, decision, defaultDates, delimiter)
      if (!cell) return
      if (decision && decision.ambiguousAsText && typeof cell.value === 'string' && numberEvidence(field.trim(), delimiter) === 'ambiguous') textColumns.add(colIndex)
      if (typeof cell.value === 'number' && cell.numFmt && /#,##0/.test(cell.numFmt)) {
        const preferred = decision ? decision.numbers : 'dot'
        const parsed = parseNumber(field, preferred) || parseNumber(field, preferred === 'dot' ? 'comma' : 'dot')
        if (parsed && parsed.groupChar) groupChars.set(parsed.groupChar, (groupChars.get(parsed.groupChar) || 0) + 1)
      }
      cells[`${columnLabels[colIndex] || (columnLabels[colIndex] = columnName(colIndex))}${rowIndex + 1}`] = cell
      if (onCell) onCell(cell, rowIndex + 1, colIndex + 1)
      maxRow = Math.max(maxRow, rowIndex + 1)
      maxCol = Math.max(maxCol, colIndex + 1)
    })
  })
  for (const col of [...textColumns].sort((a, b) => a - b).slice(0, 5)) {
    notes.push(`Column ${columnName(col)}: values such as 1.250 or 2,000 were kept as text because the file does not show which decimal separator it uses.`)
  }
  let groupSeparator = decimalComma ? '.' : ','
  let bestGroup = 0
  for (const [character, count] of groupChars) if (count > bestGroup) { bestGroup = count; groupSeparator = character }
  const { lineEnding, trailingNewline } = detectLineEnding(text)
  return {
    cells,
    maxRow,
    maxCol,
    truncated,
    notes,
    dialect: {
      delimiter,
      decimalComma,
      groupSeparator,
      dateOrder: defaultDates,
      encoding,
      hadBom,
      lineEnding,
      trailingNewline,
      ragged: lengths.size > 1,
    },
  }
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

const encoderCache = new Map()

/** Character -> bytes table for a legacy encoding, built from the platform decoder. */
function encoderFor(label) {
  if (encoderCache.has(label)) return encoderCache.get(label)
  let table = null
  try {
    const decoder = new TextDecoder(label)
    const multiByte = MULTI_BYTE_ENCODINGS.has(label)
    table = new Map()
    for (let byte = 0; byte < 0x80; byte += 1) {
      const text = decoder.decode(Uint8Array.of(byte))
      if (text.length === 1 && !table.has(text)) table.set(text, [byte])
    }
    for (let lead = 0x80; lead <= 0xff; lead += 1) {
      const single = decoder.decode(Uint8Array.of(lead))
      if (single && single !== '�' && [...single].length === 1 && !table.has(single)) table.set(single, [lead])
      if (!multiByte) continue
      for (let trail = 0x40; trail <= 0xfe; trail += 1) {
        const text = decoder.decode(Uint8Array.of(lead, trail))
        if (!text || text.includes('�') || [...text].length !== 1 || table.has(text)) continue
        table.set(text, [lead, trail])
      }
    }
  } catch {
    table = null
  }
  encoderCache.set(label, table)
  return table
}

/** Encode text in the dialect's encoding; null when a character cannot be represented. */
function encodeText(text, encoding, withBom) {
  const label = String(encoding || 'utf-8').toLowerCase()
  if (label === 'utf-8' || label === 'utf8') return Buffer.concat([withBom ? Buffer.from([0xef, 0xbb, 0xbf]) : Buffer.alloc(0), Buffer.from(text, 'utf8')])
  if (label === 'utf-16le') return Buffer.concat([withBom ? Buffer.from([0xff, 0xfe]) : Buffer.alloc(0), Buffer.from(text, 'utf16le')])
  if (label === 'utf-16be') return Buffer.concat([withBom ? Buffer.from([0xfe, 0xff]) : Buffer.alloc(0), Buffer.from(text, 'utf16le').swap16()])
  const table = encoderFor(label)
  if (!table) return null
  const bytes = []
  for (const character of text) {
    const encoded = table.get(character)
    if (!encoded) return null
    for (const byte of encoded) bytes.push(byte)
  }
  return Buffer.from(bytes)
}

function quoteField(text, delimiter) {
  if (text === '') return ''
  if (text.includes('"') || text.includes(delimiter) || text.includes('\n') || text.includes('\r') || /^\s|\s$/.test(text)) {
    return `"${text.replace(/"/g, '""')}"`
  }
  return text
}

/** Upgrade two-digit year tokens to four digits outside quoted literals and [sections]. */
function fourDigitYears(format) {
  let output = ''
  for (let index = 0; index < format.length; index += 1) {
    const character = format[index]
    if (character === '"') {
      const end = format.indexOf('"', index + 1)
      const stop = end < 0 ? format.length : end + 1
      output += format.slice(index, stop)
      index = stop - 1
      continue
    }
    if (character === '[') {
      const end = format.indexOf(']', index + 1)
      const stop = end < 0 ? format.length : end + 1
      output += format.slice(index, stop)
      index = stop - 1
      continue
    }
    if (character === '\\') {
      output += format.slice(index, index + 2)
      index += 1
      continue
    }
    if (character === 'y' || character === 'Y') {
      let end = index
      while (end < format.length && (format[end] === 'y' || format[end] === 'Y')) end += 1
      const run = end - index
      output += run <= 2 ? 'yyyy' : format.slice(index, end)
      index = end - 1
      continue
    }
    output += character
  }
  return output
}

function isDateFormat(format) {
  try { return Boolean(XLSX.SSF.is_date(format)) } catch { return false }
}

/** Swap US separators of formatted numbers for the dialect's (1,234.5 -> 1.234,5). */
function toDecimalComma(text, groupSeparator) {
  return text.replace(/\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d+/g, (number) => number.replace(/[.,]/g, (character) => (character === ',' ? groupSeparator : ',')))
}

function generalNumber(value) {
  let text = Number.isInteger(value) ? String(value) : String(Number(value.toPrecision(15)))
  text = text.replace(/e([+-]?)(\d+)$/i, (_all, sign, digits) => `E${sign || '+'}${digits.padStart(2, '0')}`)
  return text
}

function formatNumber(value, numFmt, dialect) {
  let text = null
  let date = false
  const format = typeof numFmt === 'string' ? numFmt.trim() : ''
  if (format && format !== 'General' && format !== '@') {
    date = isDateFormat(format)
    try {
      text = XLSX.SSF.format(date ? fourDigitYears(format) : format, value)
    } catch {
      text = null
    }
    if (text != null && !date) text = String(text).trim()
    if (text != null && (/^#+$/.test(text) || !text)) text = null
  }
  if (text == null) {
    date = false
    text = generalNumber(value)
  }
  if (!date && dialect.decimalComma) text = toDecimalComma(text, dialect.groupSeparator || '.')
  return text
}

function scalarText(cell, dialect) {
  if (!cell || typeof cell !== 'object') return ''
  const formula = typeof cell.formula === 'string' && cell.formula.trim()
  let value = formula ? cell.result : cell.value
  const type = formula ? cell.resultType : cell.type
  if (value == null) return formula && typeof cell.display === 'string' ? cell.display : ''
  if (value && typeof value === 'object') {
    if (value.type === 'richText' && Array.isArray(value.runs)) return value.runs.map((run) => String(run && run.text != null ? run.text : '')).join('')
    if (value.type === 'error') return String(value.value || '#VALUE!')
    if (value.type === 'date') {
      const date = new Date(value.value)
      if (Number.isNaN(date.getTime())) return String(value.value || '')
      value = (date.getTime() - Date.UTC(1899, 11, 30)) / 86_400_000
    } else if (value.type === 'binary') {
      return ''
    } else {
      return ''
    }
  }
  if (type === 'error') return String(value)
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return ''
    const numFmt = cell.numFmt || (cell.style && cell.style.numFmt)
    return formatNumber(value, type === 'date' && !numFmt ? 'yyyy-mm-dd' : numFmt, dialect)
  }
  return String(value)
}

function cellPosition(address) {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(String(address || '').toUpperCase())
  if (!match) return null
  let col = 0
  for (const character of match[1]) col = col * 26 + character.charCodeAt(0) - 64
  const row = Number(match[2])
  if (col < 1 || col > MAX_COLS || row < 1 || row > MAX_ROWS) return null
  return { row, col }
}

function hasValue(cell) {
  if (!cell || typeof cell !== 'object') return false
  if (cell.formula) return true
  return cell.value != null && cell.value !== ''
}

/**
 * Write one sheet as delimited text. `dialect` is the source file's dialect for a same-format
 * save; without one the output is Excel-friendly (UTF-8 with BOM, comma or tab, CRLF).
 * Notes about a fallback (for example an encoding that cannot represent a character) are
 * pushed to `warnings`.
 */
function writeDelimited(sheet, options = {}) {
  const format = options.format || 'csv'
  const source = options.dialect && typeof options.dialect === 'object' ? options.dialect : null
  const defaults = format === 'txt'
    ? { delimiter: '\t', encoding: 'utf-16le', hadBom: true }
    : { delimiter: format === 'tsv' ? '\t' : ',', encoding: 'utf-8', hadBom: true }
  const delimiter = source && DELIMITER_CANDIDATES.includes(source.delimiter) && (format !== 'tsv' || source.delimiter === '\t') ? source.delimiter : defaults.delimiter
  const dialect = {
    delimiter,
    decimalComma: Boolean(source && source.decimalComma && delimiter !== ','),
    groupSeparator: source && typeof source.groupSeparator === 'string' && source.groupSeparator.length === 1 ? source.groupSeparator : '.',
    encoding: (source && source.encoding) || defaults.encoding,
    hadBom: source ? Boolean(source.hadBom) : defaults.hadBom,
    lineEnding: source && ['\r\n', '\n', '\r'].includes(source.lineEnding) ? source.lineEnding : '\r\n',
    trailingNewline: source ? source.trailingNewline !== false : true,
    ragged: Boolean(source && source.ragged),
  }
  if (dialect.decimalComma && dialect.groupSeparator === ',') dialect.groupSeparator = '.'
  const rows = new Map()
  let maxRow = 0
  let maxCol = 0
  for (const [address, cell] of Object.entries((sheet && sheet.cells) || {})) {
    if (!hasValue(cell)) continue
    const position = cellPosition(address)
    if (!position) continue
    const text = scalarText(cell, dialect)
    if (text === '') continue
    let row = rows.get(position.row)
    if (!row) { row = new Map(); rows.set(position.row, row) }
    row.set(position.col, quoteField(text, delimiter))
    maxRow = Math.max(maxRow, position.row)
    maxCol = Math.max(maxCol, position.col)
  }
  const lines = []
  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const row = rows.get(rowNumber)
    const fields = []
    const width = dialect.ragged ? (row ? Math.max(...row.keys()) : 0) : maxCol
    for (let col = 1; col <= width; col += 1) fields.push((row && row.get(col)) || '')
    lines.push(fields.join(delimiter))
  }
  let text = lines.join(dialect.lineEnding)
  if (lines.length && dialect.trailingNewline) text += dialect.lineEnding
  let bytes = encodeText(text, dialect.encoding, dialect.hadBom)
  if (!bytes) {
    bytes = encodeText(text, 'utf-8', true)
    if (Array.isArray(options.warnings)) options.warnings.push(`Some characters cannot be written in the file's original ${String(dialect.encoding).toUpperCase()} encoding, so it was saved as UTF-8.`)
  }
  return bytes
}

module.exports = {
  readDelimited,
  writeDelimited,
  decodeText,
  parseRecords,
  sniffDelimiter,
  countDelimitersOutsideQuotes,
  parseNumber,
  numberFormatFor,
  fourDigitYears,
  encodeText,
  ansiEncodingFor,
  defaultLocale,
}
