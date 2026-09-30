'use strict'

const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const test = require('node:test')

const root = path.resolve(__dirname, '..')

for (const mode of ['docs', 'calc', 'pdf', 'image', 'video']) {
  test(`the staged ${mode} backend has a silent direct-print path`, () => {
    const backend = fs.readFileSync(path.join(root, 'modules', mode, 'electron', 'main.cjs'), 'utf8')
    assert.match(backend, /silent:\s*true/)
    assert.doesNotMatch(backend, /silent:\s*false/)
    assert.match(backend, /getPrintersAsync/)
  })
}

test('default-only workspaces never replace the Windows default with a named queue', () => {
  for (const mode of ['calc', 'image', 'video']) {
    const backend = fs.readFileSync(path.join(root, 'modules', mode, 'electron', 'main.cjs'), 'utf8')
    assert.doesNotMatch(backend, /printer\??\.isDefault/)
  }
})

test('packaged preview smoke never submits a physical print job', () => {
  const smoke = fs.readFileSync(path.join(root, 'scripts', 'print-smoke.cjs'), 'utf8')
  const submissionPatterns = [
    /video-print-action'\)\.click\(\)/,
    /#print-submit'\)\.click\(\)/,
    /primary-action'\)\.click\(\)/,
    /System preview'\)\.click\(\)/,
    /image-print-dialog \.print-submit'\)\.click\(\)/,
  ]
  for (const pattern of submissionPatterns) assert.doesNotMatch(smoke, pattern)
  assert.match(smoke, /final submission intentionally skipped/i)
})
