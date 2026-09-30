const test = require('node:test')
const assert = require('node:assert/strict')
const {
  directPrintOptions,
  ensurePrinterInstalled,
  printFailureMessage,
} = require('../electron/default-printer.cjs')

test('directPrintOptions suppresses native UI and leaves deviceName absent for the Windows default', () => {
  const original = { silent: false, deviceName: 'Wrong_Queue', printBackground: true, copies: 2 }
  assert.deepEqual(directPrintOptions(original), {
    silent: true,
    printBackground: true,
    copies: 2,
  })
  assert.equal(original.silent, false)
})

test('printer preflight rejects only a confirmed empty Windows printer list', async () => {
  await assert.rejects(() => ensurePrinterInstalled({ getPrintersAsync: async () => [] }), /No printers are installed/)
  await assert.doesNotReject(() => ensurePrinterInstalled({ getPrintersAsync: async () => [{ name: 'Printer_1' }] }))
  await assert.doesNotReject(() => ensurePrinterInstalled({ getPrintersAsync: async () => { throw new Error('enumeration unavailable') } }))
})

test('print callback failures are made actionable', () => {
  assert.match(printFailureMessage('Invalid printer settings'), /default printer/i)
  assert.match(printFailureMessage('printer offline'), /online/i)
  assert.match(printFailureMessage('driver rejected page size'), /default printer.*driver rejected page size/i)
})
