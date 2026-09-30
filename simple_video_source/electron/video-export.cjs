'use strict'

const path = require('node:path')

const MAX_FRAME_BYTES = 256 * 1024 * 1024
const FRAME_FORMATS = Object.freeze({
  png: Object.freeze({ extension: '.png', label: 'PNG image' }),
  jpeg: Object.freeze({ extension: '.jpg', label: 'JPEG image' }),
})

function sourceStem(sourcePath) {
  const parsed = path.parse(String(sourcePath || 'video'))
  return parsed.name || 'video'
}

function frameTimestamp(seconds) {
  const milliseconds = Math.max(0, Math.round((Number.isFinite(seconds) ? seconds : 0) * 1000))
  const wholeSeconds = Math.floor(milliseconds / 1000)
  const hours = Math.floor(wholeSeconds / 3600)
  const minutes = Math.floor((wholeSeconds % 3600) / 60)
  const remainder = wholeSeconds % 60
  const clock = [hours, minutes, remainder].map((value) => String(value).padStart(2, '0')).join('-')
  // Different frames in the same second must not suggest the same output file.
  const fraction = milliseconds % 1000
  return fraction ? `${clock}.${String(fraction).padStart(3, '0')}` : clock
}

function suggestedCopyName(sourcePath) {
  const parsed = path.parse(String(sourcePath || 'video'))
  const stem = parsed.name || 'video'
  return `${stem} copy${parsed.ext || ''}`
}

function suggestedFrameName(sourcePath, format, seconds) {
  const definition = FRAME_FORMATS[format]
  if (!definition) throw new Error('Choose PNG or JPEG for a video frame.')
  return `${sourceStem(sourcePath)} frame ${frameTimestamp(seconds)}${definition.extension}`
}

function hasExpectedExtension(targetPath, expectedExtensions) {
  const actual = path.extname(String(targetPath || '')).toLowerCase()
  return expectedExtensions.some((extension) => actual === extension.toLowerCase())
}

function assertCopyTarget(sourcePath, targetPath) {
  const sourceExtension = path.extname(String(sourcePath || '')).toLowerCase()
  if (!sourceExtension || !hasExpectedExtension(targetPath, [sourceExtension])) {
    throw new Error(`A video copy must keep its original ${sourceExtension ? sourceExtension.toUpperCase() : 'file'} extension.`)
  }
  if (path.resolve(sourcePath).toLowerCase() === path.resolve(targetPath).toLowerCase()) {
    throw new Error('Choose a new filename when saving a copy of this video.')
  }
}

function assertFrameTarget(targetPath, format) {
  const definition = FRAME_FORMATS[format]
  if (!definition) throw new Error('Choose PNG or JPEG for a video frame.')
  const extensions = format === 'jpeg' ? ['.jpg', '.jpeg'] : ['.png']
  if (!hasExpectedExtension(targetPath, extensions)) {
    throw new Error(`The selected filename must end in ${extensions.join(' or ')}.`)
  }
}

function coerceFrameBytes(value) {
  let bytes
  if (Buffer.isBuffer(value)) bytes = value
  else if (value instanceof ArrayBuffer) bytes = Buffer.from(value)
  else if (ArrayBuffer.isView(value)) bytes = Buffer.from(value.buffer, value.byteOffset, value.byteLength)
  else throw new Error('The captured frame data is invalid.')

  if (!bytes.length) throw new Error('The captured frame is empty.')
  if (bytes.length > MAX_FRAME_BYTES) throw new Error('The captured frame is too large to export safely.')
  return bytes
}

function validateFrameBytes(value, format) {
  const bytes = coerceFrameBytes(value)
  if (format === 'png') {
    const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    if (bytes.length < signature.length || !bytes.subarray(0, signature.length).equals(signature)) {
      throw new Error('The captured frame is not a valid PNG image.')
    }
  } else if (format === 'jpeg') {
    const hasStart = bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8
    const hasEnd = bytes.length >= 4 && bytes[bytes.length - 2] === 0xff && bytes[bytes.length - 1] === 0xd9
    if (!hasStart || !hasEnd) throw new Error('The captured frame is not a valid JPEG image.')
  } else {
    throw new Error('Choose PNG or JPEG for a video frame.')
  }
  return bytes
}

module.exports = {
  FRAME_FORMATS,
  MAX_FRAME_BYTES,
  assertCopyTarget,
  assertFrameTarget,
  frameTimestamp,
  suggestedCopyName,
  suggestedFrameName,
  validateFrameBytes,
}
