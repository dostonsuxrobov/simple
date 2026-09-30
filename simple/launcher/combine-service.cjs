'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const { PDFDocument } = require('../../simple_pdf_source/node_modules/pdf-lib')
const { convertOfficeBytes } = require('../../simple_doc_source/electron/office-converter.cjs')
const { prepareLegacySheetPreview } = require('./legacy-sheet-preview.cjs')
const { addImagePage } = require('../../simple_pdf_source/electron/image-to-pdf.cjs')

const WORD_EXTENSIONS = new Set(['.doc', '.docx'])
const SHEET_EXTENSIONS = new Set(['.xls', '.xlsx', '.ods'])
const EXTENSIONS = new Set(['.pdf', '.png', '.jpg', '.jpeg', ...WORD_EXTENSIONS, ...SHEET_EXTENSIONS])
const MAX_BYTES = 512 * 1024 * 1024

function pageIndices(value, count) {
  const input = String(value || '').trim()
  if (!input) return Array.from({ length: count }, (_, index) => index)
  const indices = []
  for (const part of input.split(',')) {
    const match = /^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/.exec(part)
    if (!match) throw new Error('Use page numbers such as 1–3, 5 (with a hyphen for ranges).')
    const first = Number(match[1])
    const last = Number(match[2] || match[1])
    if (first < 1 || last < first || last > count) throw new Error(`Choose pages between 1 and ${count}.`)
    for (let page = first; page <= last; page++) indices.push(page - 1)
    if (indices.length > 10_000) throw new Error('Select no more than 10,000 pages.')
  }
  return indices
}

async function imagePdf(bytes, extension) {
  const pdf = await PDFDocument.create()
  await addImagePage(pdf, bytes, extension)
  return pdf.save()
}

async function combineFiles(entries, onProgress = () => {}, options = {}) {
  if (!Array.isArray(entries) || entries.length < 2 || entries.length > 100) throw new Error('Choose between 2 and 100 files to combine.')
  const output = await PDFDocument.create()
  let totalBytes = 0
  for (let index = 0; index < entries.length; index++) {
    const entry = entries[index]
    if (!entry || typeof entry.path !== 'string' || !path.isAbsolute(entry.path)) throw new Error('Choose a local file to combine.')
    const name = path.basename(entry.path)
    const extension = path.extname(entry.path).toLowerCase()
    if (!EXTENSIONS.has(extension)) throw new Error(`${name}: choose PDF, Word, Excel, ODS, PNG, or JPEG files.`)
    onProgress({ index, total: entries.length, name })
    try {
      const stat = await fs.stat(entry.path)
      totalBytes += stat.size
      if (!stat.isFile() || !stat.size || stat.size > 256 * 1024 * 1024 || totalBytes > MAX_BYTES) throw new Error('Files must be nonempty, under 256 MB each and 512 MB together.')
      let bytes = await fs.readFile(entry.path)
      if (WORD_EXTENSIONS.has(extension) || SHEET_EXTENSIONS.has(extension)) {
        if (extension === '.xls') bytes = prepareLegacySheetPreview(bytes)
        bytes = await (options.convertOfficeBytes || convertOfficeBytes)({ bytes, inputExtension: extension.slice(1), outputExtension: 'pdf', filter: SHEET_EXTENSIONS.has(extension) ? 'calc_pdf_Export' : 'writer_pdf_Export' })
      } else if (extension !== '.pdf') bytes = await imagePdf(bytes, extension)
      const source = await PDFDocument.load(bytes, { ignoreEncryption: true })
      if (source.isEncrypted) throw new Error('Unlock this PDF before combining it.')
      const form = source.getForm()
      if (form.getFields().length || form.hasXFA()) {
        throw new Error('This PDF contains interactive form fields. Fill and flatten it, or print it to a static PDF, before combining.')
      }
      const indices = pageIndices(entry.pages, source.getPageCount())
      if (output.getPageCount() + indices.length > 10_000) throw new Error('A combined PDF can contain up to 10,000 pages.')
      const copied = await output.copyPages(source, indices)
      for (const page of copied) output.addPage(page)
    } catch (error) {
      throw new Error(`${name}: ${error.message || 'This file could not be combined.'}`)
    }
  }
  output.setTitle('Combined document')
  output.setCreator('simple')
  const bytes = Buffer.from(await output.save())
  if (bytes.length > MAX_BYTES) throw new Error('The combined PDF exceeds 512 MB. Combine fewer files at a time.')
  return { bytes, pageCount: output.getPageCount() }
}

module.exports = { EXTENSIONS, pageIndices, combineFiles }
