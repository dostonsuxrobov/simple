'use strict'

const fs = require('node:fs')
const path = require('node:path')
const { pathToFileURL } = require('node:url')

let mupdfModule = null

/**
 * mupdf.js (ES module + WASM), imported on first use. The packaged app keeps
 * the package beside the bundled backend in vendor/mupdf, because its loader
 * finds the WASM file next to import.meta.url; development and tests use
 * node_modules. A failed import is not cached, so a later call can retry.
 */
function loadMupdf() {
  mupdfModule ||= (async () => {
    const bundled = path.join(__dirname, '../vendor/mupdf/dist/mupdf.js')
    return import(pathToFileURL(fs.existsSync(bundled) ? bundled : require.resolve('mupdf')).href)
  })().catch((error) => {
    mupdfModule = null
    throw error
  })
  return mupdfModule
}

module.exports = { loadMupdf }
