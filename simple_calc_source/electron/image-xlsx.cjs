'use strict'

// Floating pictures: read worksheet images into the editor model (as data URLs with cell
// anchors) and write the model's pictures back through ExcelJS. Pictures the editor did not
// replace reuse the source package's media entry, so they are not duplicated.

const EMU_PER_PIXEL = 9525
const MAX_MODEL_IMAGE_BYTES = 64 * 1024 * 1024
const MIME = {
  png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp',
  tif: 'image/tiff', tiff: 'image/tiff', svg: 'image/svg+xml', webp: 'image/webp', emf: 'image/x-emf', wmf: 'image/x-wmf',
}
const EXTENSION = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif', 'image/bmp': 'bmp', 'image/tiff': 'tiff', 'image/x-emf': 'emf', 'image/x-wmf': 'wmf' }

function columnPixels(sheet, index) {
  if ((sheet.hiddenCols || []).includes(index + 1)) return 0
  const raw = sheet.colWidths && sheet.colWidths[String(index + 1)]
  const width = Number(raw && typeof raw === 'object' ? raw.width : raw) || Number(sheet.properties && sheet.properties.defaultColWidth) || 8.43
  return Math.trunc(((256 * width + Math.trunc(128 / 7)) / 256) * 7)
}

function rowPixels(sheet, index) {
  if ((sheet.hiddenRows || []).includes(index + 1)) return 0
  const points = Number(sheet.rowHeights && sheet.rowHeights[String(index + 1)]) || Number(sheet.properties && sheet.properties.defaultRowHeight) || 15
  return points * (4 / 3)
}

/** The cell and offset a one-cell anchor's extent ends at. */
function extendPoint(start, offsetEmu, lengthPx, sizeOf) {
  let index = start
  let remaining = offsetEmu / EMU_PER_PIXEL + lengthPx
  for (let guard = 0; guard < 20000; guard += 1) {
    const size = sizeOf(index)
    if (remaining <= size || size === undefined) break
    remaining -= size
    index += 1
  }
  return { index, offsetEmu: Math.max(0, Math.round(remaining * EMU_PER_PIXEL)) }
}

function anchorPoint(anchor) {
  return {
    row: Math.max(0, Math.floor(Number(anchor.nativeRow) || 0)),
    col: Math.max(0, Math.floor(Number(anchor.nativeCol) || 0)),
    rowOffsetEmu: Math.max(0, Math.round(Number(anchor.nativeRowOff) || 0)),
    colOffsetEmu: Math.max(0, Math.round(Number(anchor.nativeColOff) || 0)),
  }
}

/**
 * Pictures of a worksheet for the model. Returns null (leave pictures to the package) when
 * the workbook's pictures exceed the in-editor budget or cannot be read.
 */
function imagesFromWorksheet(worksheet, sheet) {
  if (typeof worksheet.getImages !== 'function') return []
  const pictures = worksheet.getImages()
  if (!pictures.length) return []
  const workbook = worksheet.workbook || worksheet._workbook
  if (!workbook || typeof workbook.getImage !== 'function') return null
  const budget = workbook.__simpleCalcImageBudget || (workbook.__simpleCalcImageBudget = { used: 0, exceeded: false })
  if (budget.exceeded) return null
  const images = []
  pictures.forEach((picture, index) => {
    const media = workbook.getImage(Number(picture.imageId))
    if (!media || !media.buffer || !picture.range || !picture.range.tl) return
    budget.used += media.buffer.length
    if (budget.used > MAX_MODEL_IMAGE_BYTES) budget.exceeded = true
    const extension = String(media.extension || '').toLowerCase()
    const from = anchorPoint(picture.range.tl)
    let to
    if (picture.range.br) to = anchorPoint(picture.range.br)
    else {
      const ext = picture.range.ext || {}
      const width = Number(ext.width) || 96
      const height = Number(ext.height) || 96
      const column = extendPoint(from.col, from.colOffsetEmu, width, (col) => columnPixels(sheet, col))
      const row = extendPoint(from.row, from.rowOffsetEmu, height, (row) => rowPixels(sheet, row))
      to = { row: row.index, col: column.index, rowOffsetEmu: row.offsetEmu, colOffsetEmu: column.offsetEmu }
    }
    const editAs = picture.range.br ? (picture.range.editAs === 'absolute' || picture.range.editAs === 'twoCell' ? picture.range.editAs : 'oneCell') : 'oneCell'
    images.push({
      id: `image-${worksheet.id || 0}-${index + 1}-${picture.imageId}`,
      ...(media.name ? { name: String(media.name) } : {}),
      src: `data:${MIME[extension] || 'application/octet-stream'};base64,${media.buffer.toString('base64')}`,
      anchor: { from, to, ...(editAs !== 'twoCell' ? { editAs } : {}) },
      sourceImageId: Number(picture.imageId),
    })
  })
  return budget.exceeded ? null : images
}

function mediaBytes(medium) {
  if (!medium) return null
  if (Buffer.isBuffer(medium.buffer)) return medium.buffer
  if (typeof medium.base64 === 'string') {
    try { return Buffer.from(medium.base64.slice(medium.base64.indexOf(',') + 1), 'base64') } catch { return null }
  }
  return null
}

/** Media id with exactly these bytes, so a picture that is already in the package is not embedded twice. */
function existingMediaId(workbook, extension, bytes) {
  const crypto = require('node:crypto')
  if (!workbook.__simpleCalcMediaIndex) {
    const index = new Map()
    ;(workbook.media || []).forEach((medium, id) => {
      const data = medium && medium.type === 'image' ? mediaBytes(medium) : null
      if (data) index.set(`${String(medium.extension || '').toLowerCase()}:${crypto.createHash('sha256').update(data).digest('hex')}`, id)
    })
    workbook.__simpleCalcMediaIndex = index
  }
  const key = `${extension}:${crypto.createHash('sha256').update(bytes).digest('hex')}`
  return { id: workbook.__simpleCalcMediaIndex.get(key), key }
}

/** Replace a worksheet's pictures with the model's. Sheets without an `images` list are left alone. */
function applyImagesToWorksheet(worksheet, sheet, preserveBase, options = {}) {
  if (!Array.isArray(sheet.images)) {
    // A matched source worksheet whose pictures the editor modelled and then removed entirely
    // (the list was dropped instead of emptied) must not keep the source pictures.
    if (!options.clearMissing) return
  }
  const workbook = worksheet.workbook || worksheet._workbook
  if (!workbook || typeof worksheet.addImage !== 'function') return
  worksheet._media = (worksheet._media || []).filter((medium) => medium.type !== 'image')
  for (const image of Array.isArray(sheet.images) ? sheet.images : []) {
    if (!image || !image.anchor || !image.anchor.from || !image.anchor.to) continue
    let imageId = null
    if (preserveBase && Number.isInteger(image.sourceImageId) && workbook.getImage(image.sourceImageId)) imageId = image.sourceImageId
    if (imageId === null) {
      const match = /^data:([^;,]+);base64,(.+)$/i.exec(String(image.src || ''))
      const extension = match && EXTENSION[match[1].toLowerCase()]
      if (!extension) continue
      let bytes = null
      try { bytes = Buffer.from(match[2], 'base64') } catch { bytes = null }
      if (!bytes || !bytes.length) continue
      const existing = existingMediaId(workbook, extension, bytes)
      if (Number.isInteger(existing.id) && workbook.getImage(existing.id)) imageId = existing.id
      else {
        imageId = workbook.addImage({ buffer: bytes, extension })
        workbook.__simpleCalcMediaIndex.set(existing.key, imageId)
      }
    }
    const point = (value) => ({
      nativeCol: Math.max(0, Math.floor(Number(value.col) || 0)),
      nativeColOff: Math.max(0, Math.round(Number(value.colOffsetEmu) || 0)),
      nativeRow: Math.max(0, Math.floor(Number(value.row) || 0)),
      nativeRowOff: Math.max(0, Math.round(Number(value.rowOffsetEmu) || 0)),
    })
    worksheet.addImage(imageId, { tl: point(image.anchor.from), br: point(image.anchor.to), editAs: image.anchor.editAs || 'twoCell' })
  }
}

/**
 * Keep only the media some worksheet still shows. ExcelJS writes every media entry it loaded
 * from the source package, so a picture the user deleted (a signature, an ID scan) would
 * otherwise stay inside the saved file. Ids are remapped in place.
 */
function compactWorkbookMedia(workbook) {
  if (!workbook || !Array.isArray(workbook.media) || !workbook.media.length) return
  const worksheets = typeof workbook.eachSheet === 'function' ? workbook.worksheets : []
  const used = new Set()
  for (const worksheet of worksheets) {
    for (const medium of worksheet._media || []) {
      const id = medium ? Number(medium.imageId) : NaN
      if (Number.isInteger(id)) used.add(id)
    }
  }
  const remap = new Map()
  const media = []
  workbook.media.forEach((medium, id) => {
    // Non-image media (none today) are kept untouched.
    if (!medium || (medium.type === 'image' && !used.has(id))) return
    remap.set(id, media.length)
    media.push(medium)
  })
  if (media.length === workbook.media.length) return
  for (const worksheet of worksheets) {
    for (const medium of worksheet._media || []) {
      const id = medium ? Number(medium.imageId) : NaN
      if (remap.has(id)) medium.imageId = remap.get(id)
    }
  }
  workbook.media = media
  delete workbook.__simpleCalcMediaIndex
}

module.exports = { imagesFromWorksheet, applyImagesToWorksheet, compactWorkbookMedia, EMU_PER_PIXEL }
