const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const test = require('node:test')
const { importExtensions } = require('../electron/office-import.cjs')

// simple/scripts/verify.cjs reads this literal list to check the unified app's
// routing, so it must name exactly the formats the PDF workspace can open.
test('SUPPORTED_EXTENSIONS in main.cjs mirrors every importable format', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'electron', 'main.cjs'), 'utf8')
  const declaration = /SUPPORTED_EXTENSIONS\s*=\s*new Set\(\s*\[([\s\S]*?)\]\s*\)/.exec(source)
  assert.ok(declaration, 'main.cjs must declare SUPPORTED_EXTENSIONS')
  const declared = new Set(declaration[1].match(/\.[a-z0-9]+/g))
  const importable = new Set(importExtensions({ engine: true }).map((extension) => `.${extension}`))
  assert.deepEqual([...declared].sort(), [...importable].sort())
})
