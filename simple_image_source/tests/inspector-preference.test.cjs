const test = require('node:test')
const assert = require('node:assert/strict')

test('the image details inspector is open by default and remembers explicit choices', async () => {
  const { inspectorOpenFromStored, storedInspectorOpen } = await import('../src/inspector-preference.js')
  assert.equal(inspectorOpenFromStored(null), true)
  assert.equal(inspectorOpenFromStored('open'), true)
  assert.equal(inspectorOpenFromStored('closed'), false)
  assert.equal(inspectorOpenFromStored('unexpected-value'), true)
  assert.equal(storedInspectorOpen(true), 'open')
  assert.equal(storedInspectorOpen(false), 'closed')
})
