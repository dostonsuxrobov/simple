'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFNull, PDFNumber, PDFRef, PDFString } = require('../../simple_pdf_source/node_modules/pdf-lib')
const { convertOfficeBytes, getOfficeEngineStatus } = require('../shared/electron/office-engine.cjs')
const formats = require('../shared/electron/formats.cjs')
const { lockLegacyDateFields } = require('../../simple_doc_source/electron/legacy-fields.cjs')
const { addImagePage } = require('../../simple_pdf_source/electron/image-to-pdf.cjs')
const { CombineError, legacyFormatError, readError } = require('./combine-policy.cjs')
const { docxPrintJobs, sniffContainer, spreadsheetPrintJobs } = require('./combine-native.cjs')
const { prepareLegacySheetPreview } = require('./legacy-sheet-preview.cjs')

const WORD_EXTENSIONS = new Set(['.doc', '.docx'])
const SHEET_EXTENSIONS = new Set(['.xls', '.xlsx', '.ods'])
const EXTENSIONS = new Set(['.pdf', '.png', '.jpg', '.jpeg', '.csv', ...WORD_EXTENSIONS, ...SHEET_EXTENSIONS])
const MAX_BYTES = 512 * 1024 * 1024
/** Names the date-field locking applied to private .doc copies; part of the engine cache key. */
const LEGACY_DOC_POLICY = 'doc-locked-date-fields-1'

function pageIndices(value, count) {
  const input = String(value || '').trim()
  if (!input) return Array.from({ length: count }, (_, index) => index)
  const indices = []
  for (const part of input.split(',')) {
    const match = /^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/.exec(part)
    if (!match) throw new CombineError('INVALID_PAGES', 'Use page numbers such as 1–3, 5 (with a hyphen for ranges).')
    const first = Number(match[1])
    const last = Number(match[2] || match[1])
    if (first < 1 || last < first || last > count) throw new CombineError('INVALID_PAGES', `Choose pages between 1 and ${count}.`)
    for (let page = first; page <= last; page++) indices.push(page - 1)
    if (indices.length > 10_000) throw new CombineError('INVALID_PAGES', 'Select no more than 10,000 pages.')
  }
  return indices
}

async function imagePdf(bytes, extension) {
  const pdf = await PDFDocument.create()
  await addImagePage(pdf, bytes, extension)
  return pdf.save()
}

/** Sniffs only the ends of a file, as the registry reads them from disk. */
function sniffBytes(bytes) {
  const buffer = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  return formats.sniff(buffer.subarray(0, formats.SNIFF_BYTES), buffer.subarray(Math.max(0, buffer.length - formats.SNIFF_BYTES)), buffer.length)
}

/**
 * What a Word or spreadsheet file really is: the extension, corrected when the
 * bytes are plainly the other container (an old binary file named .docx, or a
 * modern file named .doc). CSV is always text.
 * @returns {'doc'|'docx'|'xls'|'xlsx'|'ods'|'csv'}
 * @throws {CombineError} ENCRYPTED for a password-protected .docx or .xlsx (an
 *   encrypted package inside a compound file, which is not an older .doc or .xls)
 */
function officeKind(extension, bytes) {
  if (extension === '.csv') return 'csv'
  const word = WORD_EXTENSIONS.has(extension)
  const container = sniffContainer(bytes)
  if (container === 'cfb') {
    if (sniffBytes(bytes).kind === 'ooxml-encrypted') {
      throw new CombineError('ENCRYPTED', "This file is password-protected, and Simple can't read password-protected Word or Excel files. Add a copy saved without a password.")
    }
    return word ? 'doc' : 'xls'
  }
  if (container === 'zip' && (extension === '.doc' || extension === '.xls')) return word ? 'docx' : 'xlsx'
  return extension.slice(1)
}

/**
 * Prints Simple's own print jobs and joins their pages into one document.
 * Each job is copied with its links: the printer writes links inside a job
 * (a bookmark, a footnote) as named destinations in that job's catalog, which
 * a plain page copy leaves behind; copyPagesKeepingLinks turns them into
 * links to the copied pages, so they survive into the combined PDF.
 */
async function printJobs(jobs, printHtml) {
  if (typeof printHtml !== 'function') throw new CombineError('PRINT_UNAVAILABLE', "Simple can't make PDF pages from this file here.")
  const output = await PDFDocument.create()
  for (const job of jobs) {
    const printed = await PDFDocument.load(await printHtml(job.html, job.options))
    const pages = await copyPagesKeepingLinks(output, printed, printed.getPageIndices())
    for (const page of pages) output.addPage(page)
  }
  return output
}

/**
 * The PDF pages of one source file.
 * @returns {Promise<{pdf: PDFDocument, builtIn: boolean}>}
 */
async function sourcePdf(entry, extension, bytes, engine, options) {
  if (extension === '.pdf') return { pdf: await PDFDocument.load(bytes, { ignoreEncryption: true }), builtIn: false }
  if (!WORD_EXTENSIONS.has(extension) && !SHEET_EXTENSIONS.has(extension) && extension !== '.csv') {
    return { pdf: await PDFDocument.load(await imagePdf(bytes, extension), { ignoreEncryption: true }), builtIn: false }
  }
  const kind = officeKind(extension, bytes)
  if (kind !== 'csv' && engine.available) {
    // Unchanged when a local engine exists: its layout-faithful export, with
    // the same private-copy preparation as before.
    const sheet = kind === 'xls' || kind === 'xlsx' || kind === 'ods'
    const converted = await engine.convert(
      { bytes: kind === 'xls' ? prepareLegacySheetPreview(bytes) : bytes, inputExtension: kind, outputExtension: 'pdf', filter: sheet ? 'calc_pdf_Export' : 'writer_pdf_Export' },
      kind === 'doc' ? { prepareInput: (input) => lockLegacyDateFields(input), policyId: LEGACY_DOC_POLICY } : {},
    )
    return { pdf: await PDFDocument.load(converted, { ignoreEncryption: true }), builtIn: false }
  }
  if (kind === 'doc' || kind === 'xls') throw legacyFormatError(`file.${kind}`)
  const name = path.basename(entry.path)
  const { jobs } = kind === 'docx'
    ? await docxPrintJobs(bytes, { name })
    : await spreadsheetPrintJobs(bytes, { kind, name, path: entry.path })
  return { pdf: await printJobs(jobs, options.printHtml), builtIn: true }
}

const KEY = Object.freeze({
  A: PDFName.of('A'), Annots: PDFName.of('Annots'), B: PDFName.of('B'), D: PDFName.of('D'), Dest: PDFName.of('Dest'),
  Dests: PDFName.of('Dests'), GoTo: PDFName.of('GoTo'), Kids: PDFName.of('Kids'), Names: PDFName.of('Names'), P: PDFName.of('P'), S: PDFName.of('S'),
})

/** Resolves a reference without throwing; anything unreadable is undefined. */
function resolved(context, value) {
  try { return value instanceof PDFRef ? context.lookup(value) : value } catch { return undefined }
}

function textOf(value) {
  try { return value instanceof PDFString || value instanceof PDFHexString ? value.decodeText() : null } catch { return null }
}

/** Looks a key up in a name tree (the catalog's /Names /Dests). */
function nameTreeValue(context, root, key) {
  const seen = new Set()
  const visit = (node, depth) => {
    node = resolved(context, node)
    if (!(node instanceof PDFDict) || depth > 32 || seen.has(node)) return undefined
    seen.add(node)
    const names = resolved(context, node.get(KEY.Names))
    if (names instanceof PDFArray) {
      for (let index = 0; index + 1 < names.size(); index += 2) {
        if (textOf(resolved(context, names.get(index))) === key) return names.get(index + 1)
      }
    }
    const kids = resolved(context, node.get(KEY.Kids))
    if (kids instanceof PDFArray) {
      for (let index = 0; index < kids.size(); index += 1) {
        const found = visit(kids.get(index), depth + 1)
        if (found !== undefined) return found
      }
    }
    return undefined
  }
  return visit(root, 0)
}

/**
 * The page a destination points at inside the same document: an explicit
 * [page /XYZ …] array, or a named destination looked up in the catalog.
 * @returns {{page: PDFRef|number, rest: object[]}|null} null when it can't be resolved
 */
function destinationTarget(source, destination) {
  const context = source.context
  let value = resolved(context, destination)
  if (value instanceof PDFName || value instanceof PDFString || value instanceof PDFHexString) {
    const catalog = source.catalog
    const text = value instanceof PDFName ? null : textOf(value)
    const dests = resolved(context, catalog.get(KEY.Dests))
    let named = dests instanceof PDFDict ? dests.get(value instanceof PDFName ? value : PDFName.of(text || '')) : undefined
    if (named === undefined && text !== null) {
      const names = resolved(context, catalog.get(KEY.Names))
      if (names instanceof PDFDict) named = nameTreeValue(context, names.get(KEY.Dests), text)
    }
    value = resolved(context, named)
  }
  if (value instanceof PDFDict) value = resolved(context, value.get(KEY.D))
  if (!(value instanceof PDFArray) || value.size() < 1) return null
  const first = value.get(0)
  const page = first instanceof PDFRef ? first : (first instanceof PDFNumber ? first.asNumber() : null)
  if (page === null) return null
  const rest = []
  for (let index = 1; index < value.size(); index += 1) {
    const item = value.get(index)
    rest.push(item instanceof PDFRef || item instanceof PDFArray || item instanceof PDFDict ? PDFNull : item)
  }
  return { page, rest }
}

/**
 * Copies pages so links inside the document keep working and nothing else
 * comes along.
 *
 * pdf-lib copies a page's objects one by one and never maps the source page
 * itself, so a link to another page (and every annotation's /P back to its
 * page) would copy that page again as a hidden extra page: links would lead
 * nowhere, and a page range would still carry every page it links to. So,
 * on the in-memory source only, each copied page loses its article beads and
 * its annotations' /P, and each link inside the document is set aside before
 * copying. Afterwards a link points at the copied target page when that page
 * is part of the output, and is removed when it is not. Links to web pages or
 * other files are kept as they are.
 *
 * @param {PDFDocument} output
 * @param {PDFDocument} source loaded in memory; changed, never saved
 * @param {number[]} indices source page indices, in output order (repeats allowed)
 * @returns {Promise<import('pdf-lib').PDFPage[]>} the copied pages, not yet added
 */
async function copyPagesKeepingLinks(output, source, indices) {
  const context = source.context
  const pages = source.getPages()
  const indexByRef = new Map(pages.map((page, index) => [page.ref.toString(), index]))
  const fixesByPage = new Map()
  const targetByAnnotation = new Map()
  for (const index of new Set(indices)) {
    const node = pages[index].node
    node.delete(KEY.B)
    const annots = resolved(context, node.get(KEY.Annots))
    if (!(annots instanceof PDFArray)) continue
    const fixes = []
    for (let position = 0; position < annots.size(); position += 1) {
      const annotation = resolved(context, annots.get(position))
      if (!(annotation instanceof PDFDict)) continue
      annotation.delete(KEY.P)
      if (!targetByAnnotation.has(annotation)) {
        let destination = annotation.get(KEY.Dest)
        let viaAction = false
        if (destination === undefined) {
          const action = resolved(context, annotation.get(KEY.A))
          if (action instanceof PDFDict && resolved(context, action.get(KEY.S)) === KEY.GoTo) {
            destination = action.get(KEY.D)
            viaAction = true
          }
        }
        if (destination === undefined) continue
        const target = destinationTarget(source, destination)
        const targetIndex = target ? (typeof target.page === 'number' ? target.page : indexByRef.get(target.page.toString())) : undefined
        targetByAnnotation.set(annotation, Number.isInteger(targetIndex) && targetIndex >= 0 && targetIndex < pages.length ? { index: targetIndex, rest: target.rest } : null)
        annotation.delete(KEY.Dest)
        if (viaAction) annotation.delete(KEY.A)
      }
      fixes.push({ position, target: targetByAnnotation.get(annotation) })
    }
    if (fixes.length) fixesByPage.set(index, fixes.sort((a, b) => b.position - a.position))
  }

  const copied = await output.copyPages(source, indices)
  const copiedRef = new Map()
  indices.forEach((index, position) => { if (!copiedRef.has(index)) copiedRef.set(index, copied[position].ref) })
  const fixedArrays = new Set()
  indices.forEach((index, position) => {
    const fixes = fixesByPage.get(index)
    if (!fixes) return
    const annots = resolved(output.context, copied[position].node.get(KEY.Annots))
    // Copies of one page can share one annotation array; fix it once.
    if (!(annots instanceof PDFArray) || fixedArrays.has(annots)) return
    fixedArrays.add(annots)
    for (const fix of fixes) {
      const ref = fix.target ? copiedRef.get(fix.target.index) : undefined
      const annotation = ref ? resolved(output.context, annots.get(fix.position)) : undefined
      if (ref && annotation instanceof PDFDict) annotation.set(KEY.Dest, output.context.obj([ref, ...fix.target.rest]))
      else annots.remove(fix.position)
    }
  })
  return copied
}

/**
 * Combines PDFs, Word documents, spreadsheets and photos into one PDF.
 *
 * Word and spreadsheet files use the local office engine when this PC has one
 * (unchanged behaviour). Without it, .docx, .xlsx and .ods are laid out by
 * Simple itself, .csv always is, and .doc or .xls fail with a coded
 * NEEDS_OFFICE_ENGINE error that says to save them as .docx or .xlsx in Simple
 * first. Source files are only read.
 *
 * @param {Array<{path: string, pages?: string}>} entries 2–100 absolute paths, in output order
 * @param {(progress: {index: number, total: number, name: string}) => void} [onProgress]
 * @param {object} [options]
 * @param {(html: string, options: object) => Promise<Uint8Array>} [options.printHtml] prints one of
 *   Simple's own print jobs (the Combine worker asks the launcher main process)
 * @param {Function} [options.convertOfficeBytes] replaces the local engine (tests); when given,
 *   the engine counts as available unless options.officeEngineStatus says otherwise
 * @param {() => Promise<{available: boolean}>} [options.officeEngineStatus] replaces the engine probe
 * @param {(bytes: Uint8Array) => Promise<Uint8Array>} [options.unlockPdf] removes a permissions-only
 *   encryption (the worker asks the launcher main process, which has MuPDF); without it, every
 *   encrypted PDF is refused
 * @param {string} [options.title] document title of the combined PDF (default "Combined document")
 * @returns {Promise<{bytes: Buffer, pageCount: number, builtIn: string[]}>} builtIn lists the
 *   files Simple laid out itself
 * @throws {CombineError|Error} "<file name>: <message>" with the cause's `code`
 */
async function combineFiles(entries, onProgress = () => {}, options = {}) {
  if (!Array.isArray(entries) || entries.length < 2 || entries.length > 100) throw new CombineError('INVALID', 'Choose between 2 and 100 files to combine.')
  const output = await PDFDocument.create()
  const builtIn = []
  let engine = null
  const engineFor = async () => {
    if (engine) return engine
    const status = options.officeEngineStatus
      ? await options.officeEngineStatus()
      : (options.convertOfficeBytes ? { available: true } : await getOfficeEngineStatus())
    engine = { available: Boolean(status && status.available), convert: options.convertOfficeBytes || convertOfficeBytes }
    return engine
  }
  let totalBytes = 0
  for (let index = 0; index < entries.length; index++) {
    const entry = entries[index]
    if (!entry || typeof entry.path !== 'string' || !path.isAbsolute(entry.path)) throw new CombineError('INVALID', 'Choose a local file to combine.')
    const name = path.basename(entry.path)
    const extension = path.extname(entry.path).toLowerCase()
    if (!EXTENSIONS.has(extension)) throw new CombineError('UNSUPPORTED', `${name}: choose PDF, Word, Excel, ODS, CSV, PNG, or JPEG files.`)
    onProgress({ index, total: entries.length, name })
    try {
      const stat = await fs.stat(entry.path)
      totalBytes += stat.size
      if (!stat.isFile() || !stat.size || stat.size > 256 * 1024 * 1024 || totalBytes > MAX_BYTES) throw new CombineError('TOO_LARGE', 'Files must be nonempty, under 256 MB each and 512 MB together.')
      const bytes = await fs.readFile(entry.path)
      const needsEngine = WORD_EXTENSIONS.has(extension) || SHEET_EXTENSIONS.has(extension)
      const converted = await sourcePdf(entry, extension, bytes, needsEngine ? await engineFor() : { available: false }, options)
      let source = converted.pdf
      if (source.isEncrypted) {
        // A PDF that opens without a password (a permissions password only)
        // is unlocked in memory, as the PDF workspace does; one that needs a
        // password is refused.
        if (extension !== '.pdf' || typeof options.unlockPdf !== 'function') {
          throw new CombineError('ENCRYPTED', 'This PDF is protected. Open it in Simple, save a copy, then add the copy.')
        }
        source = await PDFDocument.load(await options.unlockPdf(bytes))
        if (source.isEncrypted) throw new CombineError('ENCRYPTED', 'This PDF is protected. Open it in Simple, save a copy, then add the copy.')
      }
      const form = source.getForm()
      if (form.getFields().length || form.hasXFA()) {
        throw new CombineError('INTERACTIVE_FORM', 'This PDF contains interactive form fields. Fill and flatten it, or print it to a static PDF, before combining.')
      }
      const indices = pageIndices(entry.pages, source.getPageCount())
      if (output.getPageCount() + indices.length > 10_000) throw new CombineError('TOO_LARGE', 'A combined PDF can contain up to 10,000 pages.')
      const copied = await copyPagesKeepingLinks(output, source, indices)
      for (const page of copied) output.addPage(page)
      if (converted.builtIn) builtIn.push(name)
    } catch (caught) {
      const error = (caught && typeof caught.syscall === 'string' && readError(caught)) || caught
      const code = typeof error?.code === 'string' && /^[A-Z][A-Z0-9_]*$/.test(error.code) ? error.code : 'UNREADABLE'
      throw new CombineError(code, `${name}: ${error?.message || 'This file could not be combined.'}`, { technical: error?.technical })
    }
  }
  // Simple's own name only: never the PDF library's name or web address.
  output.setTitle(typeof options.title === 'string' && options.title.trim() ? options.title.trim().slice(0, 200) : 'Combined document')
  output.setCreator('simple')
  output.setProducer('simple')
  const bytes = Buffer.from(await output.save())
  if (bytes.length > MAX_BYTES) throw new CombineError('TOO_LARGE', 'The combined PDF exceeds 512 MB. Combine fewer files at a time.')
  return { bytes, pageCount: output.getPageCount(), builtIn }
}

module.exports = { EXTENSIONS, LEGACY_DOC_POLICY, combineFiles, copyPagesKeepingLinks, officeKind, pageIndices }
