import fs from 'node:fs'
import path from 'node:path'
import { defineConfig, type Plugin } from 'vite'
import react from '@vitejs/plugin-react'

/**
 * pdf.js draws some text only with its character maps and standard fonts at hand
 * (PDFs without embedded fonts). They are plain data files: copied next to the
 * built page, where src/pdf-pages.ts points pdf.js at them.
 */
function pdfjsData(): Plugin {
  let outDir = 'dist'
  let root = process.cwd()
  return {
    name: 'simple-pdfjs-data',
    apply: 'build',
    configResolved(config) {
      root = config.root
      outDir = path.resolve(config.root, config.build.outDir)
    },
    closeBundle() {
      for (const folder of ['cmaps', 'standard_fonts']) {
        const from = path.join(root, 'node_modules', 'pdfjs-dist', folder)
        if (!fs.existsSync(from)) throw new Error(`pdfjs-dist/${folder} is missing; run npm install.`)
        fs.cpSync(from, path.join(outDir, 'pdfjs', folder), { recursive: true })
      }
    },
  }
}

export default defineConfig({
  base: './',
  plugins: [react(), pdfjsData()],
  build: {
    target: 'es2022',
    sourcemap: false,
  },
  // The imaging worker is a module worker that code-splits; the default 'iife'
  // worker format cannot code-split and fails the build.
  worker: {
    format: 'es',
  },
})