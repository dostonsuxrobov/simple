'use strict'

const { matchesSignature, toBytes, validateImageBytes, validateImageDimensions } = require('./image-files.cjs')

// Keep untouched JPEGs compressed all the way to Chromium. Besides avoiding a
// large PNG encode, this preserves EXIF orientation and the source color profile.
function preparePrintableImage(input) {
  const data = toBytes(input?.data)
  const extension = matchesSignature(data, '.jpg') ? '.jpg' : '.png'
  const bytes = validateImageBytes(data, extension)
  const source = validateImageDimensions(bytes, extension)
  const width = Math.trunc(Number(input?.width))
  const height = Math.trunc(Number(input?.height))
  const sameSize = width === source.width && height === source.height
  // JPEG headers use sensor axes; the browser applies EXIF orientation. The
  // decoded natural dimensions are checked again before any print submission.
  const rotatedJpeg = extension === '.jpg' && width === source.height && height === source.width
  if (!sameSize && !rotatedJpeg) throw new Error('The print preview no longer matches the image dimensions.')
  return { bytes, extension, dimensions: { width, height } }
}

module.exports = { preparePrintableImage }
