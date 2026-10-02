import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  base: './',
  plugins: [react()],
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
