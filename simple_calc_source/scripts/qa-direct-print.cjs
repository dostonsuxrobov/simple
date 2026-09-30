const assert = require('node:assert/strict')
const {
  directPrintOptions,
  ensurePrinterInstalled,
  printFailureMessage,
} = require('../electron/default-printer.cjs')

async function main() {
  const original = { silent: false, deviceName: 'Wrong_Queue', printBackground: true, pageSize: 'A4' }
  assert.deepEqual(directPrintOptions(original), {
    silent: true,
    printBackground: true,
    pageSize: 'A4',
  })
  assert.equal(original.silent, false)

  await assert.rejects(() => ensurePrinterInstalled({ getPrintersAsync: async () => [] }), /No printers are installed/)
  await assert.doesNotReject(() => ensurePrinterInstalled({ getPrintersAsync: async () => [{ name: 'Printer_1' }] }))
  await assert.doesNotReject(() => ensurePrinterInstalled({ getPrintersAsync: async () => { throw new Error('enumeration unavailable') } }))
  assert.match(printFailureMessage('Invalid printer settings'), /default printer/i)
  assert.match(printFailureMessage('printer offline'), /online/i)
  assert.match(printFailureMessage('driver rejected page size'), /default printer.*driver rejected page size/i)

  console.log('Direct print QA passed: jobs use silent:true, omit deviceName for the Windows default, and never request a native dialog.')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
