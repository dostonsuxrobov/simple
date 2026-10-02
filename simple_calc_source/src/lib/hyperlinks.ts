/**
 * Hyperlink targets as Excel stores them: a place in this workbook ("#Sheet2!A1", "#MyName"),
 * a web or email address, or a file. Pure helpers shared by the link dialog, the grid (which
 * follows links) and the HYPERLINK() function's clickable cells.
 */
import { tokenizeFormulaText } from './formula-editing'

export type LinkTarget =
  | { kind: 'internal'; location: string }
  | { kind: 'external'; url: string }
  | { kind: 'file'; path: string; location?: string }
  | { kind: 'invalid'; reason: string }

const WEB_PROTOCOLS = new Set(['http:', 'https:', 'mailto:'])
const EMAIL = /^[^\s@/:]+@[^\s@]+\.[^\s@]+$/

function safeDecode(text: string) {
  try {
    return decodeURIComponent(text)
  } catch {
    return text
  }
}

/** A web address as typed into Excel's Address box: "name@host.com" is mail, "www.x.org" is https. */
export function normalizeWebAddress(text: string): string | null {
  const value = text.trim()
  // A file on this computer is never turned into a web address.
  if (!value || /^[a-z]:[\\/]|^\\\\|^file:/i.test(value)) return null
  if (/^mailto:/i.test(value)) return value
  if (EMAIL.test(value)) return `mailto:${value}`
  if (/^\/\//.test(value)) return `https:${value}`
  if (/^[a-z][a-z0-9+.-]*:/i.test(value) && !/^[a-z]:[\\/]/i.test(value)) return value
  return `https://${value}`
}

/** True when `target` is a web or email link the app may hand to the browser or mail client. */
export function isWebLink(target: string): boolean {
  try {
    return WEB_PROTOCOLS.has(new URL(target).protocol)
  } catch {
    return false
  }
}

/**
 * What a stored hyperlink (or HYPERLINK() location) points at. `workbookName` lets
 * "[Book1.xlsx]Sheet2!A1" count as a place in this workbook, as Excel reads it.
 */
export function classifyLinkTarget(raw: string, workbookName?: string): LinkTarget {
  const target = String(raw ?? '').trim()
  if (!target) return { kind: 'invalid', reason: 'The link is empty.' }
  if (target.startsWith('#')) {
    const location = safeDecode(target.slice(1)).trim()
    return location ? { kind: 'internal', location } : { kind: 'invalid', reason: 'The link has no destination.' }
  }
  const bracket = /^\[([^\]]+)\](.+)$/.exec(target)
  if (bracket) {
    const ownName = String(workbookName || '').trim().toLocaleLowerCase()
    if (ownName && bracket[1].trim().toLocaleLowerCase() === ownName) return { kind: 'internal', location: bracket[2].trim() }
    return { kind: 'file', path: bracket[1].trim(), location: bracket[2].trim() }
  }
  if (/^(?:https?|mailto):/i.test(target)) {
    return isWebLink(target) ? { kind: 'external', url: target } : { kind: 'invalid', reason: 'This web address isn’t valid.' }
  }
  if (/^file:/i.test(target)) {
    const [path, location] = target.split('#')
    // file://server/share/x.xlsx names a network share: kept as //server/... so it is refused.
    const host = /^file:\/\/([^/]+)\//i.exec(path)?.[1]
    if (host && host.toLowerCase() !== 'localhost' && !/^[a-z]:$/i.test(host)) {
      return { kind: 'file', path: `//${safeDecode(path.replace(/^file:\/\//i, ''))}`, ...(location ? { location: safeDecode(location) } : {}) }
    }
    const local = safeDecode(path.replace(/^file:(?:\/\/(?:localhost(?=\/))?)?/i, '')).replace(/^\/([a-z]:)/i, '$1')
    return local ? { kind: 'file', path: local, ...(location ? { location: safeDecode(location) } : {}) } : { kind: 'invalid', reason: 'The file link is empty.' }
  }
  if (/^www\./i.test(target)) return { kind: 'external', url: `https://${target}` }
  if (EMAIL.test(target)) return { kind: 'external', url: `mailto:${target}` }
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) && !/^[a-z]:[\\/]/i.test(target)) {
    return { kind: 'invalid', reason: 'Only web, email and in-workbook links can be opened.' }
  }
  if (/^[a-z]:[\\/]|^\\\\|[\\/]|\.[a-z0-9]{2,5}(?:#.*)?$/i.test(target)) {
    const [path, location] = target.split('#')
    return { kind: 'file', path, ...(location ? { location } : {}) }
  }
  return { kind: 'invalid', reason: 'This link isn’t a web address or a place in this workbook.' }
}

/** Excel quotes a sheet name in a reference unless it is a plain identifier that is not a cell address. */
export function quoteSheetName(name: string): string {
  const plain = /^[A-Za-z_\u00C0-\uFFFF][A-Za-z0-9_.\u00C0-\uFFFF]*$/.test(name)
    && !/^[A-Za-z]{1,3}\d+$/.test(name)
    && !/^R\d*C\d*$/i.test(name)
    && !/^(?:TRUE|FALSE)$/i.test(name)
  return plain ? name : `'${name.replace(/'/g, "''")}'`
}

/** "#Sheet2!A1" / "#'My sheet'!B2:C4" for a cell or range, or "#MyName" for a name (sheet null). */
export function internalLinkTarget(sheetName: string | null, reference: string): string {
  const ref = reference.trim().replace(/^=/, '')
  return sheetName ? `#${quoteSheetName(sheetName)}!${ref}` : `#${ref}`
}

export interface LinkLocation {
  /** Unquoted sheet name, when the location names one. */
  sheet?: string
  /** What follows the sheet ("A1", "B2:C9") or the whole location (a name, "A1"). */
  reference: string
}

/**
 * Split an in-workbook location: "'My Sheet'!A1", "Sheet2!B2:C4", "A1", "MyName", and the
 * LibreOffice form "Sheet2.A1" (when no "!" is present).
 */
export function splitLinkLocation(location: string): LinkLocation {
  const text = location.trim().replace(/^=/, '')
  const quoted = /^'((?:[^']|'')+)'[!.](.*)$/.exec(text)
  if (quoted) return { sheet: quoted[1].replace(/''/g, "'"), reference: quoted[2].trim() }
  const bang = text.lastIndexOf('!')
  if (bang > 0) return { sheet: text.slice(0, bang).trim(), reference: text.slice(bang + 1).trim() }
  const dotted = /^([^.]+)\.(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$/.exec(text)
  if (dotted) return { sheet: dotted[1].trim(), reference: dotted[2] }
  return { reference: text }
}

/** A1-style cell or range ("A1", "$B$2:C9"), the form a "Place in this workbook" link takes. */
export function isCellReference(text: string): boolean {
  return /^\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6}(?::\$?[A-Za-z]{1,3}\$?[1-9]\d{0,6})?$/.test(text.trim())
}

/** Readable destination for a link's tooltip: "Go to Sheet2!A1", "Email ann@x.org", or the URL. */
export function linkDescription(target: string, workbookName?: string): string {
  const parsed = classifyLinkTarget(target, workbookName)
  if (parsed.kind === 'internal') return `Go to ${parsed.location}`
  if (parsed.kind === 'external') return /^mailto:/i.test(parsed.url) ? `Email ${safeDecode(parsed.url.slice(7).split('?')[0])}` : parsed.url
  if (parsed.kind === 'file') return `Open ${parsed.path}${parsed.location ? ` (${parsed.location})` : ''}`
  return target
}

/**
 * The source text of the link_location argument of a formula that is a HYPERLINK() call
 * ("HYPERLINK(\"#Sheet2!A1\",\"Go\")" gives "\"#Sheet2!A1\""), or null for other formulas.
 */
export function hyperlinkFormulaArgument(formula: string): string | null {
  const text = `=${String(formula || '').replace(/^\s*=/, '')}`
  const tokens = tokenizeFormulaText(text).filter((token) => token.kind !== 'space')
  if (tokens.length < 4 || tokens[1].kind !== 'function' || tokens[2].kind !== 'open') return null
  if (tokens[1].text.replace(/^_xlfn\./i, '').toUpperCase() !== 'HYPERLINK') return null
  let depth = 0
  let argumentStart = -1
  let argumentEnd = -1
  for (let index = 2; index < tokens.length; index += 1) {
    const token = tokens[index]
    if (token.kind === 'open') {
      depth += 1
      if (depth === 1) argumentStart = token.end
      continue
    }
    if (token.kind === 'close') {
      depth -= 1
      if (depth === 0) {
        if (argumentEnd < 0) argumentEnd = token.start
        // The call must be the whole formula ("=HYPERLINK(...)&x" is not a link cell).
        if (index !== tokens.length - 1) return null
        break
      }
      continue
    }
    if (token.kind === 'separator' && depth === 1 && argumentEnd < 0) argumentEnd = token.start
  }
  if (depth !== 0 || argumentStart < 0 || argumentEnd < argumentStart) return null
  const argument = text.slice(argumentStart, argumentEnd).trim()
  return argument || null
}

/** Directory of a Windows or POSIX path ("C:\\a\\b.xlsx" gives "C:\\a"). */
export function pathDirectory(filePath: string): string {
  const index = Math.max(filePath.lastIndexOf('\\'), filePath.lastIndexOf('/'))
  return index > 0 ? filePath.slice(0, index) : ''
}

/**
 * A network location (\\server\share, //server/share, \\?\UNC\..., file://server/...). Links
 * never open these: reading one connects to another machine and can hand it the Windows sign-in.
 * Local device paths (\\?\C:\..., \\.\C:\...) are not network locations.
 */
export function isNetworkPath(path: string): boolean {
  const text = path.trim()
  if (/^(?:\\\\|\/\/)/.test(text)) return !/^[\\/]{2}[?.][\\/][a-z]:[\\/]/i.test(text)
  return /^file:\/\/(?!\/|localhost\/)[^/]/i.test(text)
}

/**
 * Resolve a link's file path against the folder of the workbook that holds it. Network
 * locations are never resolved (null), whether written in the link or reached through it.
 */
export function resolveLinkedPath(linkPath: string, workbookPath: string | null | undefined): string | null {
  const path = linkPath.trim()
  if (!path || isNetworkPath(path)) return null
  if (/^[a-z]:[\\/]/i.test(path) || /^\\\\/.test(path) || path.startsWith('/')) return path
  const base = workbookPath ? pathDirectory(workbookPath) : ''
  if (!base || isNetworkPath(base)) return null
  const separator = base.includes('\\') ? '\\' : '/'
  const parts = base.split(/[\\/]/)
  for (const segment of path.split(/[\\/]/)) {
    if (!segment || segment === '.') continue
    if (segment === '..') { if (parts.length > 1) parts.pop(); continue }
    parts.push(segment)
  }
  return parts.join(separator)
}

/** Spreadsheet files a link may open in a new simple_calc window. */
export function isSpreadsheetPath(path: string): boolean {
  return /\.(?:xlsx|xlsm|xlsb|xls|xltx|xltm|ods|fods|csv|tsv)$/i.test(path.trim())
}
