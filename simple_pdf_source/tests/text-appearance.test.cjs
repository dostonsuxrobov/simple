const test = require('node:test')
const assert = require('node:assert/strict')
const { normalizeCffSubsetHeader, embedPdfFont } = require('../electron/font-embedding.cjs')
test('CFF subset header repair retains the entire glyph payload and never mutates input', async () => {
  const input = Uint8Array.of(1, 0, 4, 13, 0, 1, 1, 1, 14, 66)
  assert.deepEqual([...normalizeCffSubsetHeader(input)], [1, 0, 4, 4, 0, 1, 1, 1, 14, 66])
  assert.equal(input[3], 13)
  const font = { embedder: { isCFF: () => true, serializeFont: async () => input } }
  assert.equal(await embedPdfFont({ embedFont: async () => font }, []), font)
  assert.equal((await font.embedder.serializeFont())[3], 4)
  assert.throws(() => normalizeCffSubsetHeader(Uint8Array.of(0, 1, 0, 0)), /invalid header/)
})
test('paint colors survive split glyphs, ligatures, repeated labels and graphics state', async () => {
  const { textColorResolver } = await import('../electron/text-appearance.mjs')
  const names = ['save','restore','setFont','setFillRGBColor','showText','setGState','setTextRenderingMode','paintFormXObjectBegin','paintFormXObjectEnd','setFillGray','setFillColorN']
  const OPS = Object.fromEntries(names.map((n,i)=>[n,i+1])), rows=[]
  const add=(name,args)=>rows.push([OPS[name],args])
  const glyphs = s => [...s].map(unicode=>({unicode}))
  add('setFont',['body',12]);add('setFillRGBColor',Uint8Array.of(35,31,32))
  add('showText',[glyphs('ofﬁce label')]);add('save',null)
  add('setFillRGBColor',Uint8Array.of(255,255,255));add('showText',[glyphs('label')]);add('restore',null)
  add('showText',[glyphs('label')]);add('setGState',[[['ca',0.5]]]);add('showText',[glyphs('faded')])
  add('setGState',[[['ca',1]]]);add('setTextRenderingMode',[3]);add('showText',[glyphs('scan')])
  const resolve=textColorResolver({fnArray:rows.map(r=>r[0]),argsArray:rows.map(r=>r[1])},OPS)
  assert.deepEqual(resolve('body','office'),[35/255,31/255,32/255])
  assert.equal(resolve('body','label'),undefined)
  assert.equal(resolve('body','label'),undefined)
  assert.equal(resolve('body','label'),undefined)
  assert.equal(resolve('body','faded'),undefined);assert.equal(resolve('body','scan'),undefined)
})

test('visible repeated text cannot inherit a differently colored off-page label', async () => {
  const { PDFDocument, StandardFonts, rgb } = require('pdf-lib')
  const pdfjs = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const { textColorResolver } = await import('../electron/text-appearance.mjs')
  const source = await PDFDocument.create(), font = await source.embedFont(StandardFonts.Helvetica)
  const page = source.addPage([300,300])
  page.drawText('Label', { font, x: -200, y: 250, size: 20, color: rgb(1,0,0) })
  page.drawText('Label', { font, x: 30, y: 220, size: 20, color: rgb(0,0,0) })
  page.drawText('Same', { font, x: 30, y: 190, size: 20, color: rgb(0,0,0) })
  page.drawText('Same', { font, x: 30, y: 160, size: 20, color: rgb(0,0,0) })
  const pdf = await pdfjs.getDocument({ data: await source.save(), disableWorker: true, isEvalSupported: false }).promise
  try {
    const page = await pdf.getPage(1), content = await page.getTextContent()
    const resolve = textColorResolver(await page.getOperatorList(), pdfjs.OPS)
    const items = content.items.filter(item => item.str?.trim())
    assert.deepEqual(items.map(item => item.str), ['Label','Same','Same'])
    assert.equal(resolve(items[0].fontName, items[0].str), undefined)
    assert.deepEqual(resolve(items[1].fontName, items[1].str), [0,0,0])
    assert.deepEqual(resolve(items[2].fontName, items[2].str), [0,0,0])
  } finally { await pdf.destroy() }
})

test('soft masks, blend modes, transfer functions and transparent groups retain sampled fallback', async () => {
  const { textColorResolver } = await import('../electron/text-appearance.mjs')
  const { OPS } = await import('pdfjs-dist/legacy/build/pdf.mjs')
  const rows = [], add = (name,args) => rows.push([OPS[name],args])
  const glyphs = value => [[...value].map(unicode => ({ unicode }))]
  add('setFont',['body',12]);add('setFillRGBColor',Uint8Array.of(64,128,192))
  for (const [name, state] of [['masked',[['SMask',true]]],['blended',[['BM','screen']]],['transferred',[['TR',[new Uint8Array(256)]]]]]) {
    add('save',null);add('setGState',[state]);add('showText',glyphs(name));add('restore',null)
  }
  add('setGState',[[['SMask',true]]]);add('setGState',[[['SMask',false]]]);add('showText',glyphs('maskReset'))
  add('setGState',[[['BM','multiply']]]);add('setGState',[[['BM','source-over']]]);add('showText',glyphs('blendReset'))
  add('save',null);add('setGState',[[['ca',.5]]]);add('beginGroup',[{isolated:true}]);add('setGState',[[['ca',1]]]);add('showText',glyphs('grouped'));add('endGroup',[{}]);add('restore',null)
  add('setGState',[[['Font',['other',12]]]]);add('showText',glyphs('fontFromState'))
  const resolve = textColorResolver({ fnArray: rows.map(row=>row[0]), argsArray: rows.map(row=>row[1]) },OPS)
  for (const label of ['masked','blended','transferred']) assert.equal(resolve('body',label),undefined)
  assert.deepEqual(resolve('body','maskReset'),[64/255,128/255,192/255])
  assert.deepEqual(resolve('body','blendReset'),[64/255,128/255,192/255])
  assert.equal(resolve('body','grouped'),undefined)
  assert.deepEqual(resolve('other','fontFromState'),[64/255,128/255,192/255])
})
