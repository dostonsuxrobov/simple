const { createSpreadsheetPrintDocument } = require('./spreadsheet-print.cjs')
const { serializeWorkbook, nativeExportLosses } = require('./workbooks.cjs')

const EXPORT_FORMATS = new Set(['xlsx', 'xls', 'ods', 'csv', 'tsv', 'pdf', 'html'])
const TABULAR_FORMATS = new Set(['xlsx', 'xls', 'ods', 'csv', 'tsv'])
const PAGE_FORMATS = new Set(['pdf', 'html'])

function normalizeExportFormat(value) {
  const normalized = String(value || '').trim().toLowerCase().replace(/^\./, '')
  if (!EXPORT_FORMATS.has(normalized)) {
    throw new Error('Choose XLSX, XLS, ODS, CSV, TSV, PDF, or HTML for this export.')
  }
  return normalized
}

function exportFilter(format) {
  const normalized = normalizeExportFormat(format)
  const labels = {
    xlsx: 'Excel workbook',
    xls: 'Excel 97–2003 workbook',
    ods: 'OpenDocument spreadsheet',
    csv: 'Comma-separated values',
    tsv: 'Tab-separated values',
    pdf: 'PDF document',
    html: 'Web page',
  }
  return [{ name: labels[normalized], extensions: [normalized] }]
}

/**
 * What an export to `format` would lose. XLS and ODS are written by the basic SheetJS writers
 * when the document engine is absent (XLS keeps values only); with the engine nothing is listed.
 */
function exportLosses(workbook, format, options = {}) {
  const normalized = normalizeExportFormat(format)
  if (!['xls', 'ods'].includes(normalized) || options.officeEngine) return []
  return nativeExportLosses(workbook, normalized)
}

async function createSpreadsheetExport(input, format, serializationOptions = {}) {
  if (!input || typeof input !== 'object' || !input.workbook || typeof input.workbook !== 'object') {
    throw new Error('Invalid spreadsheet export request.')
  }
  const normalized = normalizeExportFormat(format)
  if (normalized === 'pdf') {
    return { format: normalized, printDocument: createSpreadsheetPrintDocument(input), bytes: null }
  }
  if (normalized === 'html') {
    const printDocument = createSpreadsheetPrintDocument(input)
    return { format: normalized, printDocument, bytes: Buffer.from(printDocument.html, 'utf8') }
  }
  const bytes = await serializeWorkbook(input.workbook, normalized, serializationOptions)
  return { format: normalized, printDocument: null, bytes }
}

module.exports = {
  EXPORT_FORMATS,
  PAGE_FORMATS,
  TABULAR_FORMATS,
  createSpreadsheetExport,
  exportFilter,
  exportLosses,
  normalizeExportFormat,
}
