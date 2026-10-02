'use strict'

const {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFStream,
  clip,
  concatTransformationMatrix,
  degrees,
  drawObject,
  endPath,
  popGraphicsState,
  pushGraphicsState,
  rectangle,
} = require('pdf-lib')
const { readableCopy } = require('./pdf-unlock.cjs')

const PAPER_SIZES = Object.freeze({
  Letter: Object.freeze({ width: 612, height: 792, label: 'Letter' }),
  A4: Object.freeze({ width: 595.28, height: 841.89, label: 'A4' }),
  Legal: Object.freeze({ width: 612, height: 1008, label: 'Legal' }),
})

const MARGIN_POINTS = Object.freeze({
  none: 0,
  minimum: 18,
  normal: 36,
})

const PDF_VIEWER_READY_TIMEOUT_MS = 30_000
const COORDINATE_ARRAY_NAMES = Object.freeze(['CL', 'Vertices', 'QuadPoints', 'L'])

function finiteNumber(value, fallback) {
  const number = Number(value)
  return Number.isFinite(number) ? number : fallback
}

function normalizedRotation(value) {
  const rotation = Math.round(finiteNumber(value, 0) / 90) * 90
  return ((rotation % 360) + 360) % 360
}

function normalizePrintLayout(input = {}) {
  const paperSize = Object.hasOwn(PAPER_SIZES, input.paperSize) ? input.paperSize : 'Letter'
  const marginMode = Object.hasOwn(MARGIN_POINTS, input.marginMode) ? input.marginMode : 'normal'
  const scaleMode = ['fit', 'actual', 'shrink', 'custom'].includes(input.scaleMode) ? input.scaleMode : 'fit'
  return {
    paperSize,
    landscape: Boolean(input.landscape),
    marginMode,
    scaleMode,
    customScale: Math.min(4, Math.max(0.25, finiteNumber(input.customScale, 1))),
  }
}

function paperDimensions(layout) {
  const normalized = normalizePrintLayout(layout)
  const paper = PAPER_SIZES[normalized.paperSize]
  return normalized.landscape
    ? { width: paper.height, height: paper.width }
    : { width: paper.width, height: paper.height }
}

function calculatePrintPlacement(sourceWidth, sourceHeight, rotation, input = {}) {
  const width = finiteNumber(sourceWidth, 0)
  const height = finiteNumber(sourceHeight, 0)
  if (width <= 0 || height <= 0) throw new Error('The PDF page has invalid dimensions.')

  const layout = normalizePrintLayout(input)
  const physicalPaper = paperDimensions(layout)
  const pageRotation = normalizedRotation(rotation)
  const quarterTurn = pageRotation === 90 || pageRotation === 270
  const rawPaper = quarterTurn
    ? { width: physicalPaper.height, height: physicalPaper.width }
    : physicalPaper
  const margin = MARGIN_POINTS[layout.marginMode]
  const printableWidth = Math.max(1, rawPaper.width - margin * 2)
  const printableHeight = Math.max(1, rawPaper.height - margin * 2)
  const fitScale = Math.min(printableWidth / width, printableHeight / height)
  let scale = fitScale
  if (layout.scaleMode === 'actual') scale = 1
  else if (layout.scaleMode === 'shrink') scale = Math.min(1, fitScale)
  else if (layout.scaleMode === 'custom') scale = layout.customScale

  const contentWidth = width * scale
  const contentHeight = height * scale
  const x = (rawPaper.width - contentWidth) / 2
  const y = (rawPaper.height - contentHeight) / 2
  return {
    ...layout,
    rotation: pageRotation,
    paperWidth: physicalPaper.width,
    paperHeight: physicalPaper.height,
    rawPaperWidth: rawPaper.width,
    rawPaperHeight: rawPaper.height,
    margin,
    printableWidth,
    printableHeight,
    fitScale,
    scale,
    x,
    y,
    contentWidth,
    contentHeight,
    cropped: x < margin - 0.01 || y < margin - 0.01,
  }
}

function normalizedIndices(input, count) {
  if (!Array.isArray(input) || input.length === 0) return Array.from({ length: count }, (_, index) => index)
  const output = []
  const seen = new Set()
  for (const value of input) {
    const index = Number(value)
    if (!Number.isInteger(index) || index < 0 || index >= count) throw new Error('The print page range is not valid.')
    if (!seen.has(index)) {
      seen.add(index)
      output.push(index)
    }
  }
  if (!output.length) throw new Error('The print page range is empty.')
  return output
}

function normalizedBox(box) {
  const left = Math.min(box.x, box.x + box.width)
  const right = Math.max(box.x, box.x + box.width)
  const bottom = Math.min(box.y, box.y + box.height)
  const top = Math.max(box.y, box.y + box.height)
  if (![left, right, bottom, top].every(Number.isFinite) || right <= left || top <= bottom) return null
  return { x: left, y: bottom, width: right - left, height: top - bottom }
}

function intersectBoxes(first, second) {
  const left = Math.max(first.x, second.x)
  const bottom = Math.max(first.y, second.y)
  const right = Math.min(first.x + first.width, second.x + second.width)
  const top = Math.min(first.y + first.height, second.y + second.height)
  if (right <= left || top <= bottom) return null
  return { x: left, y: bottom, width: right - left, height: top - bottom }
}

/**
 * Match PDF.js' page.view calculation: the visible page is the positive
 * intersection of CropBox and MediaBox, with MediaBox as the fallback for an
 * invalid/disjoint CropBox. This is the geometry the renderer preview shows.
 */
function visiblePageBox(page) {
  const mediaBox = normalizedBox(page.getMediaBox())
  if (!mediaBox) throw new Error('The PDF page has an invalid MediaBox.')
  const cropBox = normalizedBox(page.getCropBox())
  return cropBox ? intersectBoxes(mediaBox, cropBox) || mediaBox : mediaBox
}

function replaceNumberArray(context, dictionary, name, values) {
  dictionary.set(PDFName.of(name), context.obj(values))
}

function numberArray(dictionary, name) {
  const value = dictionary.lookup(PDFName.of(name))
  if (!(value instanceof PDFArray)) return null
  const numbers = []
  for (let index = 0; index < value.size(); index += 1) {
    const item = value.lookup(index)
    if (!(item instanceof PDFNumber)) return null
    numbers.push(item.asNumber())
  }
  return numbers
}

function transformPairs(values, scale, offsetX, offsetY, bounds = null) {
  const transformed = values.slice()
  for (let index = 0; index + 1 < transformed.length; index += 2) {
    const x = transformed[index] * scale + offsetX
    const y = transformed[index + 1] * scale + offsetY
    transformed[index] = bounds ? Math.min(bounds.right, Math.max(bounds.left, x)) : x
    transformed[index + 1] = bounds ? Math.min(bounds.top, Math.max(bounds.bottom, y)) : y
  }
  return transformed
}

function rectangleValues(box) {
  return [box.x, box.y, box.x + box.width, box.y + box.height]
}

function boxesEqual(first, second) {
  return Math.abs(first.x - second.x) < 0.0001
    && Math.abs(first.y - second.y) < 0.0001
    && Math.abs(first.width - second.width) < 0.0001
    && Math.abs(first.height - second.height) < 0.0001
}

function transformedBounds(values, matrix) {
  const [left, bottom, right, top] = values
  const [a, b, c, d, e, f] = matrix
  const points = [
    [left, bottom], [left, top], [right, bottom], [right, top],
  ].map(([x, y]) => [a * x + c * y + e, b * x + d * y + f])
  return {
    left: Math.min(...points.map(([x]) => x)),
    bottom: Math.min(...points.map(([, y]) => y)),
    right: Math.max(...points.map(([x]) => x)),
    top: Math.max(...points.map(([, y]) => y)),
  }
}

function wrapAppearanceStream(context, rawStream, stream, fullRect, visibleRect) {
  const boundingBox = numberArray(stream.dict, 'BBox')
  if (!boundingBox || boundingBox.length !== 4) return null
  const matrixValues = numberArray(stream.dict, 'Matrix')
  const matrix = matrixValues?.length === 6 ? matrixValues : [1, 0, 0, 1, 0, 0]
  const bounds = transformedBounds(boundingBox, matrix)
  const boundsWidth = bounds.right - bounds.left
  const boundsHeight = bounds.top - bounds.bottom
  const transform = boundsWidth > 0 && boundsHeight > 0
    ? [
        fullRect.width / boundsWidth,
        0,
        0,
        fullRect.height / boundsHeight,
        fullRect.x - bounds.left * (fullRect.width / boundsWidth),
        fullRect.y - bounds.bottom * (fullRect.height / boundsHeight),
      ]
    : [1, 0, 0, 1, fullRect.x, fullRect.y]

  const originalAppearance = rawStream instanceof PDFStream ? context.register(rawStream) : rawStream
  const appearanceName = PDFName.of('OriginalAppearance')
  const xObjects = context.obj({})
  xObjects.set(appearanceName, originalAppearance)
  const resources = context.obj({ XObject: xObjects })
  const wrapper = context.formXObject([
    pushGraphicsState(),
    rectangle(visibleRect.x, visibleRect.y, visibleRect.width, visibleRect.height),
    clip(),
    endPath(),
    concatTransformationMatrix(...transform),
    drawObject(appearanceName),
    popGraphicsState(),
  ], {
    BBox: rectangleValues(visibleRect),
    Matrix: [1, 0, 0, 1, 0, 0],
    Resources: resources,
  })
  return context.register(wrapper)
}

function clipAnnotationAppearances(context, annotation, fullRect, visibleRect) {
  const appearances = annotation.lookup(PDFName.of('AP'))
  if (!(appearances instanceof PDFDict)) return false

  const wrapEntry = (rawEntry) => {
    const entry = context.lookup(rawEntry)
    if (entry instanceof PDFStream) return wrapAppearanceStream(context, rawEntry, entry, fullRect, visibleRect)
    if (!(entry instanceof PDFDict)) return null

    let retained = 0
    for (const stateName of entry.keys()) {
      const rawState = entry.get(stateName)
      const state = context.lookup(rawState)
      const wrapper = state instanceof PDFStream
        ? wrapAppearanceStream(context, rawState, state, fullRect, visibleRect)
        : null
      if (wrapper) {
        entry.set(stateName, wrapper)
        retained += 1
      } else {
        entry.delete(stateName)
      }
    }
    return retained ? rawEntry : null
  }

  let clippedNormalAppearance = false
  for (const appearanceName of ['N', 'R', 'D']) {
    const key = PDFName.of(appearanceName)
    const rawEntry = appearances.get(key)
    if (!rawEntry) continue
    const replacement = wrapEntry(rawEntry)
    if (replacement) {
      appearances.set(key, replacement)
      if (appearanceName === 'N') clippedNormalAppearance = true
    }
    else appearances.delete(key)
  }
  if (!appearances.keys().length) annotation.delete(PDFName.of('AP'))
  return clippedNormalAppearance
}

function transformAnnotationGeometry(page, scale, offsetX, offsetY, clipBox) {
  const annotations = page.node.Annots()
  if (!annotations) return
  const context = page.doc.context
  const bounds = {
    left: clipBox.x,
    bottom: clipBox.y,
    right: clipBox.x + clipBox.width,
    top: clipBox.y + clipBox.height,
  }

  for (let index = annotations.size() - 1; index >= 0; index -= 1) {
    const annotation = annotations.lookup(index)
    if (!(annotation instanceof PDFDict)) continue
    const rect = numberArray(annotation, 'Rect')
    if (!rect || rect.length !== 4) {
      // A conforming annotation must have a Rect. Keeping an invalid one could
      // let a permissive viewer invent a placement outside the source CropBox.
      annotations.remove(index)
      continue
    }

    const transformedRect = transformPairs(rect, scale, offsetX, offsetY)
    const annotationBox = normalizedBox({
      x: transformedRect[0],
      y: transformedRect[1],
      width: transformedRect[2] - transformedRect[0],
      height: transformedRect[3] - transformedRect[1],
    })
    const visibleRect = annotationBox && intersectBoxes(annotationBox, clipBox)
    if (!visibleRect) {
      annotations.remove(index)
      continue
    }

    if (!boxesEqual(annotationBox, visibleRect)) {
      // A clipped Rect alone would squeeze the entire appearance into the
      // remaining area, visibly resurrecting the hidden portion. Wrap each
      // appearance in a Form XObject whose BBox/Rect are the intersection and
      // whose content retains the original full-rectangle transform.
      if (!clipAnnotationAppearances(context, annotation, annotationBox, visibleRect)) {
        // Without a normal appearance stream, viewers synthesize annotation
        // graphics differently. Shrinking its Rect/vertices would pull hidden
        // geometry back into view, so omit this rare partial annotation rather
        // than print content outside the author's visible CropBox.
        annotations.remove(index)
        continue
      }
    }

    replaceNumberArray(context, annotation, 'Rect', rectangleValues(visibleRect))

    // These entries contain absolute page-space coordinate pairs. Clamp them
    // to the visible rectangle as well, so viewer-generated appearances cannot
    // redraw geometry that the original CropBox hid.
    for (const name of COORDINATE_ARRAY_NAMES) {
      const values = numberArray(annotation, name)
      if (values) replaceNumberArray(context, annotation, name, transformPairs(values, scale, offsetX, offsetY, bounds))
    }

    const rectangleDifferences = numberArray(annotation, 'RD')
    if (rectangleDifferences) {
      replaceNumberArray(context, annotation, 'RD', rectangleDifferences.map((value) => value * scale))
    }

    const inkLists = annotation.lookup(PDFName.of('InkList'))
    if (inkLists instanceof PDFArray) {
      for (let strokeIndex = 0; strokeIndex < inkLists.size(); strokeIndex += 1) {
        const stroke = inkLists.lookup(strokeIndex)
        if (!(stroke instanceof PDFArray)) continue
        const values = []
        for (let valueIndex = 0; valueIndex < stroke.size(); valueIndex += 1) {
          const value = stroke.lookup(valueIndex)
          if (!(value instanceof PDFNumber)) {
            values.length = 0
            break
          }
          values.push(value.asNumber())
        }
        if (values.length) inkLists.set(strokeIndex, context.obj(transformPairs(values, scale, offsetX, offsetY, bounds)))
      }
    }
  }
}

function clipAndPlacePage(page, sourceBox, placement) {
  const scale = placement.scale
  const offsetX = placement.x - sourceBox.x * scale
  const offsetY = placement.y - sourceBox.y * scale
  const destinationClip = {
    x: placement.x,
    y: placement.y,
    width: sourceBox.width * scale,
    height: sourceBox.height * scale,
  }

  // normalize()/getContentStream() also make an empty-but-valid stream for a
  // page whose only visible objects are annotations.
  page.node.normalize()
  page.getContentStream()
  const context = page.doc.context
  const prefix = context.contentStream([
    pushGraphicsState(),
    rectangle(destinationClip.x, destinationClip.y, destinationClip.width, destinationClip.height),
    clip(),
    endPath(),
    concatTransformationMatrix(scale, 0, 0, scale, offsetX, offsetY),
  ])
  const suffix = context.contentStream([popGraphicsState()])
  page.node.wrapContentStreams(context.register(prefix), context.register(suffix))
  transformAnnotationGeometry(page, scale, offsetX, offsetY, destinationClip)
}

/**
 * Chromium's PDF viewer loads its extension shell before its out-of-process
 * plugin has parsed the PDF. The viewer emits an explicit page-title update
 * only after that second phase. Listening before loadURL avoids both the former
 * fixed delay and a fast-document race.
 */
function loadPdfViewerForPrint(browserWindow, url, options = {}) {
  const timeoutMs = Math.max(1, finiteNumber(options.timeoutMs, PDF_VIEWER_READY_TIMEOUT_MS))
  const webContents = browserWindow.webContents
  return new Promise((resolve, reject) => {
    let settled = false
    let loadComplete = false
    let viewerReady = false
    let timer = null

    const cleanup = () => {
      if (timer) clearTimeout(timer)
      webContents.removeListener('page-title-updated', onTitleUpdated)
      webContents.removeListener('did-fail-load', onFailedLoad)
      webContents.removeListener('render-process-gone', onRenderProcessGone)
      webContents.removeListener('destroyed', onDestroyed)
      browserWindow.removeListener('closed', onClosed)
    }
    const finish = (error) => {
      if (settled) return
      settled = true
      cleanup()
      if (error) reject(error)
      else resolve()
    }
    const maybeFinish = () => {
      if (loadComplete && viewerReady) finish()
    }
    const onTitleUpdated = (_event, _title, explicitSet) => {
      if (explicitSet === true) {
        viewerReady = true
        maybeFinish()
      }
    }
    const onFailedLoad = (_event, errorCode, errorDescription, _validatedUrl, isMainFrame) => {
      if (isMainFrame === false) return
      finish(new Error(`The PDF print viewer could not load (${errorCode}): ${errorDescription || 'unknown error'}`))
    }
    const onRenderProcessGone = (_event, details = {}) => {
      finish(new Error(`The PDF print viewer stopped before it was ready${details.reason ? `: ${details.reason}` : '.'}`))
    }
    const onDestroyed = () => finish(new Error('The PDF print viewer closed before it was ready.'))
    const onClosed = () => finish(new Error('The PDF print viewer closed before it was ready.'))

    webContents.on('page-title-updated', onTitleUpdated)
    webContents.on('did-fail-load', onFailedLoad)
    webContents.on('render-process-gone', onRenderProcessGone)
    webContents.on('destroyed', onDestroyed)
    browserWindow.on('closed', onClosed)
    timer = setTimeout(() => finish(new Error('The PDF print viewer did not become ready in time.')), timeoutMs)

    Promise.resolve()
      .then(() => browserWindow.loadURL(url))
      .then(() => {
        loadComplete = true
        maybeFinish()
      })
      .catch((error) => finish(error instanceof Error ? error : new Error(String(error))))
  })
}

function resolvePrinter(printers, requestedDeviceName = '') {
  const available = Array.isArray(printers)
    ? printers.filter((printer) => printer && typeof printer.name === 'string' && printer.name.trim())
    : []
  if (!available.length) throw new Error('No printer is available. Add or enable a printer in Windows, then try again.')
  const requested = String(requestedDeviceName || '')
  if (requested) {
    const selected = available.find((printer) => printer.name === requested)
    if (!selected) throw new Error('The selected printer is no longer available. Choose another printer and try again.')
    return selected
  }
  // Electron no longer marks PrinterInfo entries as the default. Omitting
  // deviceName on a silent job delegates selection to the Windows default.
  return { name: '', displayName: 'Default Windows printer' }
}

function nativePrintOptions(input = {}) {
  const layout = normalizePrintLayout(input)
  const options = {
    silent: true,
    copies: Math.min(999, Math.max(1, Math.trunc(Number(input.copies) || 1))),
    landscape: layout.landscape,
    pageSize: layout.paperSize,
    scaleFactor: 100,
    printBackground: true,
    color: input.color !== false,
    collate: input.collate !== false,
    // The generated PDF already contains the exact paper and margin placement.
    // Native margins would make the driver diverge from Simple's preview.
    margins: { marginType: 'none' },
  }
  if (input.deviceName) options.deviceName = String(input.deviceName)
  if (['simplex', 'shortEdge', 'longEdge'].includes(input.duplexMode)) options.duplexMode = input.duplexMode
  return options
}

function printWebContentsSilently(webContents, options = {}, lifecycle = {}) {
  if (!webContents || webContents.isDestroyed?.()) {
    return Promise.resolve({ success: false, failureReason: 'The print renderer is unavailable.' })
  }
  if (options.silent !== true) {
    return Promise.resolve({ success: false, failureReason: 'Direct printing requires silent mode.' })
  }
  return new Promise((resolve) => {
    let settled = false
    let timer = null
    const onDestroyed = () => finish({ success: false, failureReason: 'The print renderer closed before Windows confirmed the job. Check the print queue before trying again.' })
    const onCrashed = () => finish({ success: false, failureReason: 'The print renderer stopped before Windows confirmed the job. Check the print queue before trying again.' })
    const finish = (result) => {
      if (settled) return
      settled = true
      if (timer) clearTimeout(timer)
      webContents.removeListener?.('destroyed', onDestroyed)
      webContents.removeListener?.('render-process-gone', onCrashed)
      resolve(result)
    }
    webContents.once?.('destroyed', onDestroyed)
    webContents.once?.('render-process-gone', onCrashed)
    timer = setTimeout(() => finish({
      success: false,
      failureReason: 'Windows did not confirm the print job within two minutes. Check the print queue before trying again to avoid a duplicate copy.',
    }), Math.max(1, finiteNumber(lifecycle.timeoutMs, 120_000)))
    try {
      webContents.print(options, (success, failureReason) => {
        finish({ success: Boolean(success), failureReason: success ? '' : String(failureReason || 'The printer did not accept the job.') })
      })
    } catch (error) {
      finish({ success: false, failureReason: error instanceof Error ? error.message : String(error) })
    }
  })
}

async function preparePrintPdf(data, input = {}) {
  // Loading an encrypted file with ignoreEncryption copied still-encrypted
  // streams into an unencrypted print file, which printed blank pages. An
  // owner-locked file is decrypted to a temporary copy; a file that needs a
  // password raises PASSWORD_REQUIRED.
  const source = await PDFDocument.load(await readableCopy(data), { updateMetadata: false })
  const indices = normalizedIndices(input.pageIndices, source.getPageCount())
  const output = await PDFDocument.create()
  const copiedPages = await output.copyPages(source, indices)

  for (let offset = 0; offset < copiedPages.length; offset += 1) {
    const page = copiedPages[offset]
    const sourceBox = visiblePageBox(page)
    const rotation = normalizedRotation(page.getRotation().angle)
    const placement = calculatePrintPlacement(sourceBox.width, sourceBox.height, rotation, input)

    clipAndPlacePage(page, sourceBox, placement)
    page.setRotation(degrees(rotation))
    page.setMediaBox(0, 0, placement.rawPaperWidth, placement.rawPaperHeight)
    page.setCropBox(0, 0, placement.rawPaperWidth, placement.rawPaperHeight)
    page.setBleedBox(0, 0, placement.rawPaperWidth, placement.rawPaperHeight)
    page.setTrimBox(0, 0, placement.rawPaperWidth, placement.rawPaperHeight)
    page.setArtBox(0, 0, placement.rawPaperWidth, placement.rawPaperHeight)
    // UserUnit would scale the newly assigned point-based paper size a second
    // time. The renderer preview and placement model both operate in PDF units.
    page.node.delete(PDFName.of('UserUnit'))
    output.addPage(page)
  }

  output.setProducer('simple')
  output.setCreator('simple print')
  return output.save({ useObjectStreams: true })
}

module.exports = {
  MARGIN_POINTS,
  PAPER_SIZES,
  calculatePrintPlacement,
  loadPdfViewerForPrint,
  nativePrintOptions,
  normalizePrintLayout,
  paperDimensions,
  preparePrintPdf,
  printWebContentsSilently,
  resolvePrinter,
  visiblePageBox,
}
