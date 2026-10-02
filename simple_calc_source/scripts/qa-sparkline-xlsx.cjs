'use strict'

// Excel sparkline groups are written into the worksheet extLst, read back with their colours
// and flags, and unedited groups survive byte-for-byte.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')

async function sheetXml(bytes) {
  return (await JSZip.loadAsync(bytes)).file('xl/worksheets/sheet1.xml').async('string')
}

async function main() {
  const model = {
    version: 1, name: 'spark.xlsx', activeSheetId: 's1',
    sheets: [{
      id: 's1', name: 'Trend', state: 'visible', rowCount: 10, colCount: 8, merges: [], colWidths: {}, rowHeights: {},
      cells: { A1: { value: 1 }, B1: { value: 3 }, C1: { value: 2 }, A2: { value: -1 }, B2: { value: 4 }, C2: { value: 1 } },
      conditionalFormattings: [{ ref: 'A1:C2', rules: [{ type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: 'FF638EC6' } }] }],
      sparklineGroups: [
        { type: 'line', markers: true, high: true, colors: { series: '#1F4E79', high: '#00B050' }, displayEmptyCellsAs: 'gap', lineWeight: 1.25, sparklines: [{ source: 'Trend!A1:C1', cell: 'D1' }] },
        { type: 'column', negative: true, colors: { series: '#4472C4', negative: '#C00000' }, minAxisType: 'group', maxAxisType: 'group', sparklines: [{ source: 'Trend!A1:C1', cell: 'E1' }, { source: 'Trend!A2:C2', cell: 'E2' }] },
      ],
    }],
  }
  const first = await serializeWorkbook(model, 'xlsx')
  const xml = await sheetXml(first)
  assert.match(xml, /<ext uri="\{78C0D931[^"]*"[\s\S]*?<\/ext><ext uri="\{05C60535/, 'the sparklines share the worksheet extLst with the conditional-format extension')
  assert.match(xml, /<ext uri="\{05C60535-1F16-4fd2-B633-F4F36F0B64E0\}"/)
  assert.match(xml, /<x14:sparklineGroup [^>]*markers="1"[^>]*high="1"/)
  assert.match(xml, /<xm:f>Trend!A1:C1<\/xm:f><xm:sqref>D1<\/xm:sqref>/)
  assert.ok(xml.lastIndexOf('</extLst>') > xml.lastIndexOf('</conditionalFormatting>'), 'extLst stays the last element')

  const reopened = await workbookPayloadFromBytes('spark.xlsx', first)
  const groups = reopened.workbook.sheets[0].sparklineGroups
  assert.equal(groups.length, 2)
  assert.equal(groups[0].type, 'line')
  assert.equal(groups[0].markers, true)
  assert.equal(groups[0].colors.series, '#1F4E79')
  assert.equal(groups[0].colors.high, '#00B050')
  assert.equal(groups[0].lineWeight, 1.25)
  assert.equal(groups[1].type, 'column')
  assert.equal(groups[1].minAxisType, 'group')
  assert.deepEqual(groups[1].sparklines.map((item) => item.cell), ['E1', 'E2'])
  assert.ok(groups[0].sourceXml, 'the source XML is kept for byte-exact saving')

  const again = await serializeWorkbook(reopened.workbook, 'xlsx', { baseBytes: first })
  const xml2 = await sheetXml(again)
  assert.ok(xml2.includes(groups[0].sourceXml), 'an unedited group is copied byte-for-byte')
  const edited = structuredClone(reopened.workbook)
  edited.sheets[0].sparklineGroups[0].type = 'column'
  const xml3 = await sheetXml(await serializeWorkbook(edited, 'xlsx', { baseBytes: first }))
  assert.match(xml3, /<x14:sparklineGroup [^>]*type="column"[^>]*markers="1"/, 'an edited group is rewritten')
  const cleared = structuredClone(reopened.workbook)
  cleared.sheets[0].sparklineGroups = []
  assert.doesNotMatch(await sheetXml(await serializeWorkbook(cleared, 'xlsx', { baseBytes: first })), /sparkline/)
  // CALC-SIE-25: a sheet renamed in the editor; sources still naming the old sheet are written
  // with the new name (Excel would have rewritten them on rename).
  const renamed = structuredClone(reopened.workbook)
  renamed.sheets[0].name = 'Sales Q1'
  const xml4 = await sheetXml(await serializeWorkbook(renamed, 'xlsx', { baseBytes: first }))
  assert.match(xml4, /<xm:f>'Sales Q1'!A1:C1<\/xm:f><xm:sqref>D1<\/xm:sqref>/)
  assert.doesNotMatch(xml4, /<xm:f>Trend!/)
  const renamedBack = await workbookPayloadFromBytes('spark.xlsx', await serializeWorkbook(renamed, 'xlsx', { baseBytes: first }))
  assert.equal(renamedBack.workbook.sheets[0].sparklineGroups[1].sparklines[1].source, "'Sales Q1'!A2:C2")
  process.stdout.write('Sparkline XLSX QA passed: groups written into the extLst, read back with colours and flags, kept byte-exact, rewritten when edited, removed when cleared, sources follow a renamed sheet.\n')
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
