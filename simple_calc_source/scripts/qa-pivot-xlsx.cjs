'use strict'

// Pivot definitions are saved in a custom XML data part, reloaded onto their sheets, kept
// single across saves, and removed when the last pivot goes.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')

async function main() {
  const pivot = {
    id: 'pivot-1', name: 'PivotTable1', source: "Data!$A$1:$C$4", anchor: { row: 0, col: 0 },
    rows: [{ field: 'Region' }], columns: [], values: [{ field: 'Units', summarize: 'sum' }], filters: [], extent: { rows: 3, cols: 2 },
  }
  const model = {
    version: 1, name: 'pivot.xlsx', activeSheetId: 's2',
    sheets: [
      { id: 's1', name: 'Data', state: 'visible', rowCount: 10, colCount: 5, merges: [], colWidths: {}, rowHeights: {},
        cells: { A1: { value: 'Region' }, B1: { value: 'Units' }, A2: { value: 'East' }, B2: { value: 3 }, A3: { value: 'West' }, B3: { value: ']]> tricky' } } },
      { id: 's2', name: 'Pivot Table 1', state: 'visible', rowCount: 10, colCount: 5, merges: [], colWidths: {}, rowHeights: {},
        cells: { A1: { value: 'Region' }, B1: { value: 'Sum of Units' }, A2: { value: 'East' }, B2: { value: 3 } },
        pivots: [{ ...pivot, name: 'Pivot ]]> name' }] },
    ],
  }
  const first = await serializeWorkbook(model, 'xlsx')
  const zip = await JSZip.loadAsync(first)
  const items = Object.keys(zip.files).filter((name) => /^customXml\/item\d+\.xml$/.test(name))
  assert.equal(items.length, 1)
  const number = /item(\d+)/.exec(items[0])[1]
  assert.ok(zip.file(`customXml/itemProps${number}.xml`))
  assert.ok(zip.file(`customXml/_rels/item${number}.xml.rels`))
  assert.match(await zip.file('[Content_Types].xml').async('string'), new RegExp(`/customXml/itemProps${number}\.xml"[^>]*customXmlProperties`))
  assert.match(await zip.file('xl/_rels/workbook.xml.rels').async('string'), new RegExp(`relationships/customXml" Target="\.\./customXml/item${number}\.xml"`))

  const reopened = await workbookPayloadFromBytes('pivot.xlsx', first)
  const sheet = reopened.workbook.sheets.find((item) => item.name === 'Pivot Table 1')
  assert.equal(sheet.pivots.length, 1)
  assert.equal(sheet.pivots[0].name, 'Pivot ]]> name', 'CDATA terminators in the data survive')
  assert.deepEqual(sheet.pivots[0].values, [{ field: 'Units', summarize: 'sum' }])
  assert.deepEqual(sheet.pivots[0].extent, { rows: 3, cols: 2 })

  const second = await serializeWorkbook(reopened.workbook, 'xlsx', { baseBytes: first })
  const zip2 = await JSZip.loadAsync(second)
  assert.equal(Object.keys(zip2.files).filter((name) => /^customXml\/item\d+\.xml$/.test(name)).length, 1, 'saving again keeps one definition part')
  assert.equal((await zip2.file('xl/_rels/workbook.xml.rels').async('string')).match(/customXml\/item/g).length, 1)

  const without = structuredClone(reopened.workbook)
  for (const item of without.sheets) delete item.pivots
  const third = await serializeWorkbook(without, 'xlsx', { baseBytes: second })
  const zip3 = await JSZip.loadAsync(third)
  assert.equal(Object.keys(zip3.files).filter((name) => /^customXml\//.test(name)).length, 0, 'the part goes with the last pivot')
  assert.doesNotMatch(await zip3.file('xl/_rels/workbook.xml.rels').async('string'), /customXml/)
  assert.doesNotMatch(await zip3.file('[Content_Types].xml').async('string'), /customXml/)
  process.stdout.write('Pivot XLSX QA passed: definitions saved as custom XML, reloaded, kept single, and removed with the last pivot.\n')
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
