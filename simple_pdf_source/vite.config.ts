import { createHash } from 'node:crypto'
import { readFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { defineConfig, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'

const require = createRequire(import.meta.url)

// OCR runtime (tesseract.js) shipped for offline use. The files are read from
// node_modules and emitted as dist/ocr/** at build time (served from memory in
// dev), so no generated binaries enter the source tree. The engine wrapper
// relies on behaviour verified against these exact versions.
const OCR_PINNED_VERSIONS: Record<string, string> = {
  'tesseract.js': '7.0.0',
  'tesseract.js-core': '7.0.0',
  '@tesseract.js-data/eng': '1.0.0',
}
const OCR_CORE_FILE = 'tesseract-core-simd-lstm.wasm.js'
const OCR_LANGUAGES = [
  { code: 'eng', label: 'English', packageName: '@tesseract.js-data/eng', folder: '4.0.0_best_int' },
]

const OCR_NOTICE = `Text recognition (OCR) components shipped with Simple

worker.min.js
  tesseract.js 7.0.0, Apache License 2.0 (tesseract.js-LICENSE.md).
  Bundled third-party code: tesseract.js-worker-third-party.txt.

core/${OCR_CORE_FILE}
  tesseract.js-core 7.0.0, Apache License 2.0 (tesseract.js-core-LICENSE.txt):
  Tesseract OCR compiled to WebAssembly, statically linked with Leptonica,
  libjpeg, libpng, libtiff and zlib, each under its own permissive licence.

tessdata/*.traineddata.gz
  See tessdata-NOTICE.txt.

GlyphLessFont
  See GlyphLessFont-NOTICE.txt.
`

const TESSDATA_NOTICE = `Tesseract trained data

The language files in tessdata/ (for example eng.traineddata.gz, an integer
LSTM model, version 4.0.0) come from the Tesseract OCR project
(https://github.com/tesseract-ocr/tessdata) and are licensed under the Apache
License, Version 2.0; the full text is in tesseract.js-core-LICENSE.txt. They
are distributed through the npm packages @tesseract.js-data/<language> (MIT).
`

const GLYPHLESS_NOTICE = `GlyphLessFont

The invisible, searchable text layer that Simple writes into recognised pages
uses GlyphLessFont (pdf.ttf) from the Tesseract OCR project
(https://github.com/tesseract-ocr/tesseract), licensed under the Apache
License, Version 2.0; the full text is in tesseract.js-core-LICENSE.txt.
Simple adjusts the font's vertical metrics (ascent 0.8 em, descent -0.2 em).
`

const sha256 = (bytes: Uint8Array) => createHash('sha256').update(bytes).digest('hex')

function readPackageFile(specifier: string): Uint8Array {
  return new Uint8Array(readFileSync(require.resolve(specifier)))
}

/** Every file under ocr/, keyed by its path relative to ocr/, plus manifest.json. */
function collectOcrAssets(): Map<string, Uint8Array> {
  for (const [name, version] of Object.entries(OCR_PINNED_VERSIONS)) {
    const installed = (require(`${name}/package.json`) as { version: string }).version
    if (installed !== version) {
      throw new Error(`OCR assets: ${name} ${version} is required (found ${installed}). The OCR engine wrapper is verified against these exact versions.`)
    }
  }
  const files = new Map<string, Uint8Array>()
  files.set('worker.min.js', readPackageFile('tesseract.js/dist/worker.min.js'))
  files.set(`core/${OCR_CORE_FILE}`, readPackageFile(`tesseract.js-core/${OCR_CORE_FILE}`))
  const languages = OCR_LANGUAGES.map((language) => {
    const file = `tessdata/${language.code}.traineddata.gz`
    const bytes = readPackageFile(`${language.packageName}/${language.folder}/${language.code}.traineddata.gz`)
    files.set(file, bytes)
    return { code: language.code, label: language.label, model: 'best_int', file, bytes: bytes.length, sha256: sha256(bytes) }
  })
  const text = (value: string) => new TextEncoder().encode(value)
  files.set('licenses/NOTICE.txt', text(OCR_NOTICE))
  files.set('licenses/tesseract.js-LICENSE.md', readPackageFile('tesseract.js/LICENSE.md'))
  files.set('licenses/tesseract.js-worker-third-party.txt', readPackageFile('tesseract.js/dist/worker.min.js.LICENSE.txt'))
  files.set('licenses/tesseract.js-core-LICENSE.txt', readPackageFile('tesseract.js-core/LICENSE'))
  files.set('licenses/tessdata-NOTICE.txt', text(TESSDATA_NOTICE))
  files.set('licenses/GlyphLessFont-NOTICE.txt', text(GLYPHLESS_NOTICE))
  const manifest = {
    schema: 1,
    tesseractJs: OCR_PINNED_VERSIONS['tesseract.js'],
    tesseractCore: OCR_PINNED_VERSIONS['tesseract.js-core'],
    worker: 'worker.min.js',
    core: OCR_CORE_FILE,
    languages,
    files: Object.fromEntries([...files].map(([name, bytes]) => [name, { bytes: bytes.length, sha256: sha256(bytes) }])),
  }
  files.set('manifest.json', text(`${JSON.stringify(manifest, null, 2)}\n`))
  return files
}

const OCR_CONTENT_TYPES: Record<string, string> = {
  '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.gz': 'application/octet-stream',
  '.md': 'text/plain; charset=utf-8',
  '.txt': 'text/plain; charset=utf-8',
}

function simpleOcrAssets(): Plugin {
  let assets: Map<string, Uint8Array> | undefined
  const load = () => (assets ??= collectOcrAssets())
  return {
    name: 'simple-ocr-assets',
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        const pathname = new URL(request.url ?? '/', 'http://localhost').pathname
        if (!pathname.startsWith('/ocr/') || (request.method !== 'GET' && request.method !== 'HEAD')) return next()
        const bytes = load().get(decodeURIComponent(pathname.slice('/ocr/'.length)))
        if (!bytes) {
          response.statusCode = 404
          response.end('Not found')
          return
        }
        const extension = pathname.slice(pathname.lastIndexOf('.'))
        response.setHeader('Content-Type', OCR_CONTENT_TYPES[extension] ?? 'application/octet-stream')
        response.setHeader('Content-Length', String(bytes.length))
        response.setHeader('Cache-Control', 'no-store')
        response.end(request.method === 'HEAD' ? undefined : Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength))
      })
    },
    generateBundle() {
      for (const [name, bytes] of load()) this.emitFile({ type: 'asset', fileName: `ocr/${name}`, source: bytes })
    },
  }
}

export default defineConfig({
  plugins: [react(), simpleOcrAssets()],
  base: './',
  server: {
    port: 5173,
    strictPort: true,
  },
  build: {
    target: 'es2022',
  },
})
