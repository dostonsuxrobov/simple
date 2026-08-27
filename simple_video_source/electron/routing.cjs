'use strict'

const path = require('node:path')

// These are containers Chromium can commonly play without a plug-in. Whether a
// particular file plays still depends on the audio/video codecs inside it.
const SUPPORTED_EXTENSIONS = Object.freeze([
  '.mp4',
  '.m4v',
  '.webm',
  '.ogv',
  '.mov',
  '.mkv',
])

const SUPPORTED_SET = new Set(SUPPORTED_EXTENSIONS)

function isSupportedVideoPath(candidate) {
  return typeof candidate === 'string'
    && candidate.length > 0
    && SUPPORTED_SET.has(path.extname(candidate).toLowerCase())
}

function supportedPaths(argv, cwd = process.cwd(), platform = process.platform) {
  const seen = new Set()
  const result = []
  for (const argument of Array.isArray(argv) ? argv : []) {
    if (!isSupportedVideoPath(argument)) continue
    const resolved = path.resolve(cwd, argument)
    const identity = platform === 'win32' ? resolved.toLowerCase() : resolved
    if (seen.has(identity)) continue
    seen.add(identity)
    result.push(resolved)
  }
  return result
}

module.exports = {
  SUPPORTED_EXTENSIONS,
  isSupportedVideoPath,
  supportedPaths,
}
