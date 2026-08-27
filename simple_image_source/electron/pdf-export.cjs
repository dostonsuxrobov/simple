'use strict'

const path = require('node:path')
const { validateImageBytes, validateImageDimensions } = require('./image-files.cjs')

const MAX_PDF_PAGE_POINTS = 841.89
let pdfLibModule

function getPdfLib() {
  pdfLibModule ||= require('pdf-lib')
  return pdfLibModule
}

function ensurePdfExtension(filePath) {
  const value = String(filePath || '')
  if (!value) throw new Error('The PDF output path is invalid.')
  const extension = path.extname(value)
  if (extension.toLowerCase() === '.pdf') return value
  return `${extension ? value.slice(0, -extension.length) : value}.pdf`
}

function pdfPageDimensions(width, height) {
  const normalizedWidth = Number(width)
  const normalizedHeight = Number(height)
  if (!Number.isFinite(normalizedWidth) || !Number.isFinite(normalizedHeight) || normalizedWidth <= 0 || normalizedHeight <= 0) {
    throw new Error('The image dimensions are invalid for PDF conversion.')
  }
  const scale = Math.min(1, MAX_PDF_PAGE_POINTS / normalizedWidth, MAX_PDF_PAGE_POINTS / normalizedHeight)
  return {
    width: normalizedWidth * scale,
    height: normalizedHeight * scale,
    scale,
  }
}

async function imageToPdfBytes(value, title = 'Converted image') {
  const bytes = validateImageBytes(value, '.png')
  validateImageDimensions(bytes, '.png')
  const { PDFDocument } = getPdfLib()
  const pdf = await PDFDocument.create()
  const embedded = await pdf.embedPng(bytes)
  const pageSize = pdfPageDimensions(embedded.width, embedded.height)
  const page = pdf.addPage([pageSize.width, pageSize.height])
  page.drawImage(embedded, { x: 0, y: 0, width: pageSize.width, height: pageSize.height })
  pdf.setTitle(String(title || 'Converted image'))
  pdf.setCreator('simple')
  pdf.setProducer('simple')
  return Buffer.from(await pdf.save())
}

module.exports = {
  MAX_PDF_PAGE_POINTS,
  ensurePdfExtension,
  imageToPdfBytes,
  pdfPageDimensions,
}
