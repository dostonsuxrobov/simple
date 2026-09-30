import assert from 'node:assert/strict'
import { clearAfterSuccessfulCopy } from '../src/lib/clipboard'

{
  let clears = 0
  const result = await clearAfterSuccessfulCopy(async () => false, () => { clears += 1 })
  assert.equal(result, false)
  assert.equal(clears, 0, 'a reported clipboard failure must preserve every selected cell')
}

{
  let clears = 0
  const result = await clearAfterSuccessfulCopy(async () => { throw new Error('Clipboard permission denied') }, () => { clears += 1 })
  assert.equal(result, false)
  assert.equal(clears, 0, 'a rejected clipboard write must preserve every selected cell')
}

{
  let clears = 0
  const result = await clearAfterSuccessfulCopy(async () => true, () => { clears += 1 })
  assert.equal(result, true)
  assert.equal(clears, 1, 'a successful clipboard write should clear the selection exactly once')
}

console.log('Clipboard QA passed: Cut clears only after an affirmative clipboard write.')
