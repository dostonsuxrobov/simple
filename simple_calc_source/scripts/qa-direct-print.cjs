const assert = require('node:assert/strict')
const {
  directPrintOptions,
  ensurePrinterInstalled,
  listPrinters,
  printFailureMessage,
  printOptionsForDocument,
  registerPrinterHandlers,
} = require('../electron/default-printer.cjs')
const { createSpreadsheetPrintDocument } = require('../electron/spreadsheet-print.cjs')

function printRequest(options) {
  const cells = {}
  for (let row = 1; row <= 150; row += 1) cells[`A${row}`] = { value: row }
  return {
    name: 'Job.xlsx',
    workbook: { version: 1, name: 'Job.xlsx', activeSheetId: 's1', sheets: [{ id: 's1', name: 'Job', rowCount: 150, colCount: 1, cells, merges: [], colWidths: {}, rowHeights: {} }] },
    options: { scope: 'active-sheet', orientation: 'landscape', paperSize: 'a4', gridlines: false, headings: false, ...options },
  }
}

/** The same steps main.cjs takes for workbook:print, against a webContents that records instead of printing. */
async function simulatePrint(input, printers) {
  const printDocument = createSpreadsheetPrintDocument(input)
  const calls = []
  const webContents = {
    getPrintersAsync: async () => printers,
    print: (options, callback) => { calls.push(options); callback(true, '') },
  }
  await ensurePrinterInstalled(webContents, printDocument.printJob)
  await new Promise((resolve, reject) => webContents.print(printOptionsForDocument(printDocument), (success, reason) => (success ? resolve() : reject(new Error(printFailureMessage(reason, printDocument.printJob))))))
  return { printDocument, calls }
}

async function main() {
  const original = { silent: false, deviceName: 'Wrong_Queue', printBackground: true, pageSize: 'A4' }
  assert.deepEqual(directPrintOptions(original), {
    silent: true,
    printBackground: true,
    pageSize: 'A4',
  })
  assert.equal(original.silent, false)
  assert.deepEqual(directPrintOptions(original, { copies: 1, collate: true }), { silent: true, printBackground: true, pageSize: 'A4', copies: 1, collate: true }, 'a job without a printer still goes to the Windows default')
  assert.deepEqual(directPrintOptions({ collate: true }, { deviceName: '  HP LaserJet  ', copies: 3, collate: false }), { silent: true, collate: false, deviceName: 'HP LaserJet', copies: 3 })
  assert.equal(directPrintOptions({}, { deviceName: 'Bad\nQueue', copies: 0 }).deviceName, undefined, 'an unsafe queue name is never sent')
  assert.equal(directPrintOptions({}, { copies: 0 }).copies, undefined)
  assert.equal(directPrintOptions({}, { copies: 50_000 }).copies, 999)

  await assert.rejects(() => ensurePrinterInstalled({ getPrintersAsync: async () => [] }), /No printers are installed/)
  await assert.doesNotReject(() => ensurePrinterInstalled({ getPrintersAsync: async () => [{ name: 'Printer_1' }] }))
  await assert.doesNotReject(() => ensurePrinterInstalled({ getPrintersAsync: async () => { throw new Error('enumeration unavailable') } }))
  await assert.rejects(() => ensurePrinterInstalled({ getPrintersAsync: async () => [{ name: 'Printer_1' }] }, { deviceName: 'Removed printer' }), /"Removed printer" is no longer available/, 'a removed printer is never swapped for another queue')
  await assert.doesNotReject(() => ensurePrinterInstalled({ getPrintersAsync: async () => [{ name: 'Printer_1' }] }, { deviceName: 'Printer_1' }))
  assert.match(printFailureMessage('Invalid printer settings'), /default printer/i)
  assert.match(printFailureMessage('printer offline'), /online/i)
  assert.match(printFailureMessage('driver rejected page size'), /default printer.*driver rejected page size/i)
  assert.match(printFailureMessage('printer offline', { deviceName: 'Office Laser' }), /"Office Laser" is unavailable/)
  assert.match(printFailureMessage('', { deviceName: 'Office Laser' }), /could not use the printer "Office Laser"/)

  // Printer list for the dialog: names as the OS knows them, display names, sorted, deduplicated.
  const listed = await listPrinters({ getPrintersAsync: async () => [
    { name: 'Zeta', displayName: 'Zeta Printer', description: 'Laser', options: { 'printer-location': 'Office' } },
    { name: 'Alpha', displayName: '', description: 42 },
    { name: 'Alpha', displayName: 'Duplicate' },
    { name: '' },
    null,
  ] })
  assert.deepEqual(listed, [{ name: 'Alpha', displayName: 'Alpha', description: '' }, { name: 'Zeta', displayName: 'Zeta Printer', description: 'Laser' }])
  assert.deepEqual(await listPrinters(null), [])

  const handlers = new Map()
  let checkedSender = false
  registerPrinterHandlers({ handle: (channel, handler) => handlers.set(channel, handler) }, { assertTrustedSender: () => { checkedSender = true } })
  assert.deepEqual([...handlers.keys()], ['workbook:list-printers'])
  const reply = await handlers.get('workbook:list-printers')({ sender: { getPrintersAsync: async () => [{ name: 'Office', displayName: 'Office Laser' }] } })
  assert.equal(checkedSender, true, 'the printer list answers trusted senders only')
  assert.deepEqual(reply, [{ name: 'Office', displayName: 'Office Laser', description: '' }])

  // The dialog's printer, copies, page range, scale and margins reach the print payload.
  const { printDocument, calls } = await simulatePrint(printRequest({
    scaling: 'custom', scalePercent: 80,
    margins: 'custom', customMargins: { top: 1, right: 0.4, bottom: 1, left: 0.4 },
    pageRange: { from: 2, to: 3 },
    printer: { deviceName: 'Office Laser', copies: 2, collate: false },
  }), [{ name: 'Office Laser' }, { name: 'Other' }])
  assert.equal(calls.length, 1)
  assert.deepEqual(calls[0], {
    printBackground: true, color: true, landscape: true, margins: { marginType: 'none' }, pageSize: 'A4',
    scaleFactor: 100, pagesPerSheet: 1, collate: false, silent: true, deviceName: 'Office Laser', copies: 2,
  })
  assert.equal(printDocument.pageCount, 2, 'only pages 2-3 are in the printed document')
  assert.ok(printDocument.totalPages > 3)
  assert.match(printDocument.html, /data-page-number="2"[\s\S]*data-page-number="3"/)
  assert.doesNotMatch(printDocument.html, /data-page-number="1"/)
  assert.match(printDocument.html, /data-scale="0\.8000"/, 'the custom scale is in the document')
  assert.match(printDocument.html, /--margin-top:96\.00px;--margin-right:38\.40px/, 'custom margins are in the document')
  await assert.rejects(() => simulatePrint(printRequest({ printer: { deviceName: 'Gone' } }), [{ name: 'Office Laser' }]), /"Gone" is no longer available/)
  const defaultJob = await simulatePrint(printRequest({}), [{ name: 'Office Laser' }])
  assert.equal(defaultJob.calls[0].deviceName, undefined, 'no printer choice means the Windows default')
  assert.equal(defaultJob.calls[0].copies, 1)

  console.log('Direct print QA passed: jobs use silent:true, omit deviceName for the Windows default, carry the chosen printer, copies and collation, keep page range, scale and margins in the document, refuse removed printers, list printers for the dialog, and never request a native dialog.')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
