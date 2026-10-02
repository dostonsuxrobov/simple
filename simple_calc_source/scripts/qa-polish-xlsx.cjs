'use strict'
/*
 * What the polish features store in the file survives an XLSX save and reopen:
 * a sheet's zoom (sheetView zoomScale, CALC-023), links to a place in the workbook
 * (CALC-005), and the calculation options (calcPr, CALC-027).
 *
 *   node scripts/qa-polish-xlsx.cjs
 */
const assert = require('node:assert/strict')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
const { workbookPayloadFromBytes, serializeWorkbook } = require('../electron/workbooks.cjs')

async function part(bytes, name) {
  return (await JSZip.loadAsync(bytes)).file(name).async('string')
}

async function main() {
  const book = new ExcelJS.Workbook()
  const first = book.addWorksheet('Summary', { views: [{ state: 'normal', zoomScale: 85 }] })
  first.getCell('A1').value = 'Go to data'
  first.getCell('A2').value = 12
  book.addWorksheet('Data detail').getCell('B3').value = 'Target'
  const base = Buffer.from(await book.xlsx.writeBuffer())

  const opened = (await workbookPayloadFromBytes('polish.xlsx', base)).workbook
  assert.equal(Number(opened.sheets[0].views?.[0]?.zoomScale), 85, 'a saved zoom is read back')

  // What the app writes: each sheet's zoom into its view, links as "#Sheet!A1" / "#Name",
  // a link on a formula cell, and calculation options in metadata.calcProperties.
  const model = structuredClone(opened)
  model.sheets[0].views = [{ ...(model.sheets[0].views?.[0] || {}), zoomScale: 150, zoomScaleNormal: 150 }]
  model.sheets[1].views = [{ zoomScale: 60, zoomScaleNormal: 60 }]
  model.sheets[0].cells.A1 = { ...model.sheets[0].cells.A1, hyperlink: "#'Data detail'!B3" }
  model.sheets[0].cells.A3 = { formula: 'A2*2', result: 24, hyperlink: '#Summary!A2' }
  model.sheets[0].cells.A4 = { value: 'Site', hyperlink: 'https://example.com/', hyperlinkTooltip: 'https://example.com/' }
  model.metadata = { ...(model.metadata || {}), calcProperties: { ...(model.metadata?.calcProperties || {}), calcMode: 'manual', iterate: true, iterateCount: 50, iterateDelta: 0.01 } }

  for (const baseBytes of [base, null]) {
    const label = baseBytes ? 'merged save' : 'fresh save'
    const saved = await serializeWorkbook(model, 'xlsx', { baseBytes })
    const sheetXml = await part(saved, 'xl/worksheets/sheet1.xml')
    assert.match(sheetXml, /zoomScale="150"/, `${label}: the zoom is written to the sheet view`)
    assert.match(sheetXml, /<hyperlink ref="A1" location="(?:'|&apos;)Data detail(?:'|&apos;)!B3"\/>/, `${label}: an in-workbook link is a location, not a relationship`)
    const workbookXml = await part(saved, 'xl/workbook.xml')
    assert.match(workbookXml, /calcMode="manual"/, `${label}: manual calculation is saved`)
    assert.match(workbookXml, /iterate="1"|iterate="true"/, `${label}: iterative calculation is saved`)

    const back = (await workbookPayloadFromBytes('back.xlsx', saved)).workbook
    assert.equal(Number(back.sheets[0].views?.[0]?.zoomScale), 150, `${label}: zoom reopens`)
    assert.equal(Number(back.sheets[1].views?.[0]?.zoomScale), 60, `${label}: each sheet keeps its own zoom`)
    assert.equal(back.sheets[0].cells.A1.hyperlink, "#'Data detail'!B3", `${label}: the link to a place reopens`)
    assert.equal(back.sheets[0].cells.A1.value, 'Go to data')
    assert.equal(back.sheets[0].cells.A3.formula, 'A2*2', `${label}: a linked formula keeps its formula`)
    assert.equal(back.sheets[0].cells.A3.hyperlink, '#Summary!A2', `${label}: and its link`)
    assert.equal(back.sheets[0].cells.A4.hyperlink, 'https://example.com/', `${label}: web links still work`)
    const calc = back.metadata?.calcProperties || {}
    assert.equal(String(calc.calcMode), 'manual', `${label}: calculation mode reopens`)
    assert.ok(calc.iterate === true || calc.iterate === 1 || calc.iterate === '1' || calc.iterate === 'true', `${label}: iteration reopens`)
    assert.equal(Number(calc.iterateCount), 50)
    assert.equal(Number(calc.iterateDelta), 0.01)
  }
  console.log('Polish XLSX QA passed: per-sheet zoom, in-workbook links (also on formulas), calculation options.')
}

main().catch((error) => {
  console.error(error)
  process.exit(1)
})
