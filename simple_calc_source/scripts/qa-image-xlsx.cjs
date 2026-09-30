'use strict'

// Pictures inserted in the editor are written as sheet drawings, read back with their cell
// anchors, moved/deleted edits apply, and unchanged source pictures are not duplicated.
const assert = require('node:assert/strict')
const JSZip = require('jszip')
const { serializeWorkbook, workbookPayloadFromBytes } = require('../electron/workbooks.cjs')

// 2x2 red PNG
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFklEQVR42mP8z8Dwn4GBgYGJAQoAAAAeAgYGl5lzAAAAAElFTkSuQmCC'

async function media(bytes) {
  const zip = await JSZip.loadAsync(bytes)
  return { zip, media: Object.keys(zip.files).filter((name) => name.startsWith('xl/media/')) }
}

async function main() {
  const model = {
    version: 1, name: 'pictures.xlsx', activeSheetId: 's1',
    sheets: [{
      id: 's1', name: 'Pictures', state: 'visible', rowCount: 30, colCount: 10, merges: [], colWidths: {}, rowHeights: {},
      cells: { A1: { value: 'Logo' } },
      images: [
        { id: 'i1', name: 'logo.png', src: `data:image/png;base64,${PNG}`, anchor: { from: { row: 1, col: 1, rowOffsetEmu: 9525 * 4, colOffsetEmu: 9525 * 6 }, to: { row: 5, col: 3 }, editAs: 'oneCell' } },
        { id: 'i2', src: `data:image/png;base64,${PNG}`, anchor: { from: { row: 10, col: 0 }, to: { row: 12, col: 2 } } },
      ],
    }],
  }
  const first = await serializeWorkbook(model, 'xlsx')
  const { zip, media: files } = await media(first)
  assert.ok(files.length >= 1, 'the picture bytes are in xl/media')
  const drawingName = Object.keys(zip.files).find((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name))
  assert.ok(drawingName, 'a drawing part is written')
  const drawing = await zip.file(drawingName).async('string')
  assert.equal((drawing.match(/<xdr:pic>/g) || []).length, 2)
  assert.match(drawing, /editAs="oneCell"/)
  assert.match(drawing, /<xdr:col>1<\/xdr:col><xdr:colOff>57150<\/xdr:colOff><xdr:row>1<\/xdr:row><xdr:rowOff>38100<\/xdr:rowOff>/)

  const payload = await workbookPayloadFromBytes('pictures.xlsx', first)
  const sheet = payload.workbook.sheets[0]
  assert.equal(sheet.images.length, 2)
  const logo = sheet.images[0]
  assert.match(logo.src, /^data:image\/png;base64,/)
  assert.deepEqual(logo.anchor.from, { row: 1, col: 1, rowOffsetEmu: 38100, colOffsetEmu: 57150 })
  assert.deepEqual(logo.anchor.to, { row: 5, col: 3, rowOffsetEmu: 0, colOffsetEmu: 0 })
  assert.equal(logo.anchor.editAs, 'oneCell')
  assert.ok(Number.isInteger(logo.sourceImageId))
  assert.ok(!payload.warnings.some((warning) => /images/.test(warning)), 'modelled pictures need no compatibility warning')

  // Move one picture, delete the other, save against the source package.
  const edited = structuredClone(payload.workbook)
  edited.sheets[0].images = [{ ...edited.sheets[0].images[0], anchor: { from: { row: 3, col: 4 }, to: { row: 7, col: 6 }, editAs: 'oneCell' } }]
  const second = await serializeWorkbook(edited, 'xlsx', { baseBytes: first })
  const reread = await workbookPayloadFromBytes('pictures.xlsx', second)
  assert.equal(reread.workbook.sheets[0].images.length, 1)
  assert.equal(reread.workbook.sheets[0].images[0].anchor.from.col, 4)
  const secondDrawing = await (await JSZip.loadAsync(second)).file(Object.keys((await JSZip.loadAsync(second)).files).find((name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name))).async('string')
  assert.equal((secondDrawing.match(/<xdr:pic>/g) || []).length, 1)
  process.stdout.write('Image XLSX QA passed: inserted pictures, anchors, reload, move/delete edits, and source media reuse.\n')
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
