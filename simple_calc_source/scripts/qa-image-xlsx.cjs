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
  return { zip, media: Object.keys(zip.files).filter((name) => name.startsWith('xl/media/') && !zip.files[name].dir) }
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
  await deletedAndRepeatedPictures()
  process.stdout.write('Image XLSX QA passed: inserted pictures, anchors, reload, move/delete edits, source media reuse, deleted pictures removed from the package, and no growth on repeated saves.\n')
}

const RED = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg=='
const BLUE = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=='

/**
 * calc-file-io-objects-9: a deleted picture is not left inside the saved file (privacy), and an
 * inserted picture is embedded once, however often the workbook is saved.
 */
async function deletedAndRepeatedPictures() {
  const anchor = (row) => ({ from: { row, col: 1 }, to: { row: row + 2, col: 3 } })
  const model = {
    version: 1, name: 'two.xlsx', activeSheetId: 's1', metadata: {},
    sheets: [{ id: 's1', name: 'Sheet1', rowCount: 30, colCount: 10, merges: [], colWidths: {}, rowHeights: {}, cells: { A1: { value: 1 } },
      images: [
        { id: 'red', src: `data:image/png;base64,${RED}`, anchor: anchor(1) },
        { id: 'blue', src: `data:image/png;base64,${BLUE}`, anchor: anchor(5) },
      ] }],
  }
  const original = await serializeWorkbook(model, 'xlsx')
  assert.equal((await media(original)).media.length, 2)
  const opened = (await workbookPayloadFromBytes('two.xlsx', original)).workbook
  const blueBytes = Buffer.from(BLUE, 'base64')
  // The user deletes the blue picture and saves over the source package.
  opened.sheets[0].images = opened.sheets[0].images.filter((image) => !image.src.includes(BLUE))
  const withoutBlue = await serializeWorkbook(opened, 'xlsx', { baseBytes: original })
  const { zip, media: remaining } = await media(withoutBlue)
  assert.equal(remaining.length, 1, 'only the shown picture stays in xl/media')
  for (const name of remaining) assert.ok(!(await zip.file(name).async('nodebuffer')).equals(blueBytes), 'the deleted picture is gone from the package')
  // A picture inserted in the editor, saved three times (each time over the previous file).
  opened.sheets[0].images.push({ id: 'new', src: `data:image/png;base64,${BLUE}`, anchor: anchor(10) })
  let base = original
  const sizes = []
  for (let save = 0; save < 3; save += 1) {
    base = await serializeWorkbook(opened, 'xlsx', { baseBytes: base })
    sizes.push((await media(base)).media.length)
  }
  assert.deepEqual(sizes, [2, 2, 2], 'no orphaned copy is added per save')
  // The same picture inserted twice is stored once.
  opened.sheets[0].images.push({ id: 'again', src: `data:image/png;base64,${BLUE}`, anchor: anchor(15) })
  assert.equal((await media(await serializeWorkbook(opened, 'xlsx', { baseBytes: original }))).media.length, 2)
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`)
  process.exitCode = 1
})
