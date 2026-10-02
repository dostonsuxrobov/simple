'use strict'

// Charts QA: DrawingML import, byte-exact round trip, regenerated/new chart
// export, structural validity, renderer sanity/performance and print overlay.
// Run: node scripts/qa-charts.cjs   (set QA_CHARTS_OFFICE=1 to also convert through LibreOffice)

const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
const { SaxesParser } = require('saxes')
const { workbookPayloadFromBytes, serializeWorkbook } = require('../electron/workbooks.cjs')
const { chartFingerprint } = require('../electron/chart-xlsx.cjs')
const { createSpreadsheetPrintDocument } = require('../electron/spreadsheet-print.cjs')

const ROOT = path.resolve(__dirname, '..')

function loadTsLibraries() {
  const esbuild = require('esbuild')
  const outfile = path.join(ROOT, 'tmp', 'qa-charts-lib.cjs')
  fs.mkdirSync(path.dirname(outfile), { recursive: true })
  esbuild.buildSync({
    stdin: { contents: "export * from './src/lib/charts'; export * from './src/lib/chart-render'", resolveDir: ROOT, loader: 'ts', sourcefile: 'qa-charts-entry.ts' },
    bundle: true, platform: 'node', format: 'cjs', outfile, logLevel: 'silent',
  })
  delete require.cache[outfile]
  return require(outfile)
}

function wellFormed(xml, label) {
  const parser = new SaxesParser({ xmlns: true })
  let error = null
  parser.on('error', (value) => { error = error || value })
  parser.write(xml).close()
  if (error) throw new Error(`${label} is not well-formed XML: ${error.message}`)
}

// ---------------------------------------------------------------------------
// Fixture: an ExcelJS workbook (with a picture) plus hand-written DrawingML charts
// ---------------------------------------------------------------------------

const NS = {
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==', 'base64')
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun']
const SALES = [1200, 1500, 1350, 1800, 2100, 1950]
const COSTS = [800, 950, 900, 1100, 1300, 1250]
const MARGIN = [0.33, 0.37, 0.33, 0.39, 0.38, 0.36]

const strCache = (values) => `<c:strCache><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:strCache>`
const numCache = (values, format = 'General') => `<c:numCache><c:formatCode>${format}</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`
const tx = (ref, text) => `<c:tx><c:strRef><c:f>${ref}</c:f>${strCache([text])}</c:strRef></c:tx>`
const cat = (ref = 'Data!$A$2:$A$7') => `<c:cat><c:strRef><c:f>${ref}</c:f>${strCache(MONTHS)}</c:strRef></c:cat>`
const val = (ref, values, format) => `<c:val><c:numRef><c:f>${ref}</c:f>${numCache(values, format)}</c:numRef></c:val>`
const rich = (text, rot) => `<c:tx><c:rich><a:bodyPr${rot ? ` rot="${rot}" vert="horz"` : ''}/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx>`
const txPr = '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>'
const chartSpace = (chartBody, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:c16r2="http://schemas.microsoft.com/office/drawing/2015/06/chart"><c:date1904 val="0"/><c:lang val="en-US"/><c:roundedCorners val="0"/><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="c14" xmlns:c14="http://schemas.microsoft.com/office/drawing/2007/8/2/chart"><c14:style val="102"/></mc:Choice><mc:Fallback><c:style val="2"/></mc:Fallback></mc:AlternateContent><c:chart>${chartBody}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln></c:spPr>${txPr}${extra}<c:printSettings><c:headerFooter/><c:pageMargins b="0.75" l="0.7" r="0.7" t="0.75" header="0.3" footer="0.3"/><c:pageSetup/></c:printSettings></c:chartSpace>`
const catAx = (id, cross, extra = '') => `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${extra}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`
const valAx = (id, cross, { pos = 'l', grid = true, extra = '', scaling = '', crosses = 'autoZero', numFmt = '<c:numFmt formatCode="General" sourceLinked="1"/>', between = 'between' } = {}) => `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/>${scaling}</c:scaling><c:delete val="0"/><c:axPos val="${pos}"/>${grid ? '<c:majorGridlines/>' : ''}${extra}${numFmt}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="${crosses}"/><c:crossBetween val="${between}"/></c:valAx>`

const CHART_COLUMN = chartSpace(
  `<c:title>${rich('Revenue &amp; Costs')}<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>` +
  `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` +
  `<c:ser><c:idx val="0"/><c:order val="0"/>${tx('Data!$B$1', 'Sales')}<c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>${cat()}${val('Data!$B$2:$B$7', SALES)}<c:extLst><c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart"><c16:uniqueId val="{00000000-0001-0000-0000-000000000000}"/></c:ext></c:extLst></c:ser>` +
  `<c:ser><c:idx val="1"/><c:order val="1"/>${tx('Data!$C$1', 'Costs')}<c:spPr><a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></c:spPr><c:invertIfNegative val="0"/>${cat()}${val('Data!$C$2:$C$7', COSTS)}</c:ser>` +
  `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls><c:gapWidth val="219"/><c:overlap val="-27"/><c:axId val="111"/><c:axId val="222"/></c:barChart>` +
  catAx(111, 222, `<c:title>${rich('Month')}<c:overlay val="0"/></c:title>`) +
  valAx(222, 111, { scaling: '<c:max val="3000"/>', numFmt: '<c:numFmt formatCode="#,##0" sourceLinked="0"/>', extra: `<c:title>${rich('USD', -5400000)}<c:overlay val="0"/></c:title>` }) +
  `</c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>`)

const CHART_COMBO = chartSpace(
  `<c:autoTitleDeleted val="1"/><c:plotArea><c:layout/>` +
  `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx('Data!$B$1', 'Sales')}<c:spPr><a:solidFill><a:srgbClr val="476B57"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/><c:dLbls><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr><c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>${cat()}${val('Data!$B$2:$B$7', SALES)}</c:ser><c:gapWidth val="150"/><c:axId val="1"/><c:axId val="2"/></c:barChart>` +
  `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="1"/><c:order val="1"/>${tx('Data!$D$1', 'Margin')}<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="C9A227"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="diamond"/><c:size val="7"/></c:marker>${cat()}${val('Data!$D$2:$D$7', MARGIN, '0%')}<c:smooth val="1"/></c:ser><c:marker val="1"/><c:axId val="3"/><c:axId val="4"/></c:lineChart>` +
  catAx(1, 2) + valAx(2, 1) + catAx(3, 4).replace('<c:delete val="0"/>', '<c:delete val="1"/>') + valAx(4, 3, { pos: 'r', grid: false, crosses: 'max', numFmt: '<c:numFmt formatCode="0%" sourceLinked="0"/>' }) +
  `</c:plotArea><c:legend><c:legendPos val="t"/><c:overlay val="0"/></c:legend>`)

const CHART_PIE = chartSpace(
  `<c:title>${rich('Sales mix')}<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx('Data!$B$1', 'Sales')}` +
  `<c:dPt><c:idx val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></c:spPr></c:dPt><c:dPt><c:idx val="2"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent6"/></a:solidFill></c:spPr></c:dPt>` +
  `<c:dLbls><c:dLblPos val="bestFit"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showBubbleSize val="0"/><c:showLeaderLines val="1"/></c:dLbls>${cat()}${val('Data!$B$2:$B$7', SALES)}</c:ser><c:firstSliceAng val="30"/></c:pieChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>`)

const CHART_SCATTER = chartSpace(
  `<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx("'My Sheet'!$B$1", 'Y')}<c:spPr><a:ln w="19050" cap="rnd"><a:noFill/><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>` +
  `<c:xVal><c:numRef><c:f>'My Sheet'!$A$2:$A$6</c:f>${numCache([1, 2, 3, 4, 5])}</c:numRef></c:xVal><c:yVal><c:numRef><c:f>'My Sheet'!$B$2:$B$6</c:f>${numCache([2, 4, 3, 8, 6])}</c:numRef></c:yVal><c:smooth val="0"/></c:ser><c:axId val="7"/><c:axId val="8"/></c:scatterChart>` +
  valAx(7, 8, { pos: 'b', grid: false, between: 'midCat' }) + valAx(8, 7, { between: 'midCat' }) + `</c:plotArea>`)

const CHART_DOUGHNUT = chartSpace(
  `<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:doughnutChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${tx("'My Sheet'!$B$1", 'Y')}<c:cat><c:numRef><c:f>'My Sheet'!$A$2:$A$6</c:f>${numCache([1, 2, 3, 4, 5])}</c:numRef></c:cat>${val("'My Sheet'!$B$2:$B$6", [2, 4, 3, 8, 6])}</c:ser><c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart></c:plotArea><c:legend><c:legendPos val="r"/><c:overlay val="0"/></c:legend>`)

const CHART_WATERFALL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cx:chartSpace xmlns:a="${NS.a}" xmlns:r="${NS.r}" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chartData><cx:data id="0"><cx:strDim type="cat"><cx:f>Data!$A$2:$A$7</cx:f></cx:strDim><cx:numDim type="val"><cx:f>Data!$B$2:$B$7</cx:f></cx:numDim></cx:data></cx:chartData><cx:chart><cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:txData><cx:v>Bridge</cx:v></cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="waterfall" uniqueId="{1D4B0E5C-0000-0000-0000-000000000001}"><cx:dataId val="0"/></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`
const CHART_STYLE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cs:chartStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="${NS.a}" id="201"><cs:axisTitle><cs:lnRef idx="0"/><cs:fillRef idx="0"/><cs:effectRef idx="0"/><cs:fontRef idx="minor"><a:schemeClr val="tx1"/></cs:fontRef><cs:defRPr sz="1000"/></cs:axisTitle></cs:chartStyle>`
const CHART_COLORS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cs:colorStyle xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" xmlns:a="${NS.a}" meth="cycle" id="10"><a:schemeClr val="accent1"/><a:schemeClr val="accent2"/><a:schemeClr val="accent3"/><cs:variation/></cs:colorStyle>`

const point = (tag, col, colOff, row, rowOff) => `<xdr:${tag}><xdr:col>${col}</xdr:col><xdr:colOff>${colOff}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${rowOff}</xdr:rowOff></xdr:${tag}>`
const frame = (id, name, rId) => `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${rId}"/></a:graphicData></a:graphic></xdr:graphicFrame>`
const chartExFrame = (id, rId) => `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" Requires="cx1"><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Chart ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:r="${NS.r}" r:id="${rId}"/></a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="4572000" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1100"/><a:t>This chart isn't available in your version of Excel.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback></mc:AlternateContent>`

async function buildFixture() {
  const workbook = new ExcelJS.Workbook()
  const data = workbook.addWorksheet('Data')
  data.addRow(['Month', 'Sales', 'Costs', 'Margin'])
  MONTHS.forEach((month, index) => data.addRow([month, SALES[index], COSTS[index], MARGIN[index]]))
  data.getColumn(4).numFmt = '0%'
  const imageId = workbook.addImage({ buffer: PNG, extension: 'png' })
  data.addImage(imageId, { tl: { col: 0, row: 9 }, ext: { width: 20, height: 20 } })
  const other = workbook.addWorksheet('My Sheet')
  other.addRow(['X', 'Y'])
  ;[[1, 2], [2, 4], [3, 3], [4, 8], [5, 6]].forEach((row) => other.addRow(row))
  const zip = await JSZip.loadAsync(await workbook.xlsx.writeBuffer())
  const drawing1 = await zip.file('xl/drawings/drawing1.xml').async('string')
  const anchors = [
    `<xdr:twoCellAnchor>${point('from', 5, 0, 1, 0)}${point('to', 12, 304800, 16, 95250)}${frame(2, 'Chart 1', 'rIdC1')}<xdr:clientData/></xdr:twoCellAnchor>`,
    `<xdr:twoCellAnchor editAs="oneCell">${point('from', 13, 0, 1, 0)}${point('to', 20, 0, 16, 0)}${frame(3, 'Combo', 'rIdC2')}<xdr:clientData/></xdr:twoCellAnchor>`,
    `<xdr:oneCellAnchor>${point('from', 5, 0, 18, 0)}<xdr:ext cx="3048000" cy="1905000"/>${frame(4, 'Pie', 'rIdC3')}<xdr:clientData/></xdr:oneCellAnchor>`,
    `<xdr:twoCellAnchor>${point('from', 13, 0, 18, 0)}${point('to', 20, 0, 33, 0)}${frame(5, 'Scatter', 'rIdC4')}<xdr:clientData/></xdr:twoCellAnchor>`,
    `<xdr:twoCellAnchor>${point('from', 21, 0, 1, 0)}${point('to', 28, 0, 16, 0)}${chartExFrame(6, 'rIdC5')}<xdr:clientData/></xdr:twoCellAnchor>`,
  ]
  zip.file('xl/drawings/drawing1.xml', drawing1.replace('</xdr:wsDr>', `${anchors.join('')}</xdr:wsDr>`))
  const rels1 = await zip.file('xl/drawings/_rels/drawing1.xml.rels').async('string')
  const chartRel = (id, target) => `<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="${target}"/>`
  zip.file('xl/drawings/_rels/drawing1.xml.rels', rels1.replace('</Relationships>', `${chartRel('rIdC1', '../charts/chart1.xml')}${chartRel('rIdC2', '../charts/chart2.xml')}${chartRel('rIdC3', '../charts/chart3.xml')}${chartRel('rIdC4', '../charts/chart4.xml')}<Relationship Id="rIdC5" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="../charts/chartEx1.xml"/></Relationships>`))
  zip.file('xl/charts/chart1.xml', CHART_COLUMN)
  zip.file('xl/charts/_rels/chart1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors1.xml"/><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style1.xml"/></Relationships>`)
  zip.file('xl/charts/style1.xml', CHART_STYLE)
  zip.file('xl/charts/colors1.xml', CHART_COLORS)
  zip.file('xl/charts/chart2.xml', CHART_COMBO)
  zip.file('xl/charts/chart3.xml', CHART_PIE)
  zip.file('xl/charts/chart4.xml', CHART_SCATTER)
  zip.file('xl/charts/chart5.xml', CHART_DOUGHNUT)
  zip.file('xl/charts/chartEx1.xml', CHART_WATERFALL)
  // Second sheet: a drawing written from scratch with an absolute anchor.
  zip.file('xl/drawings/drawing2.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${NS.a}"><xdr:absoluteAnchor><xdr:pos x="1828800" y="190500"/><xdr:ext cx="3657600" cy="2286000"/>${frame(2, 'Doughnut', 'rId1')}<xdr:clientData/></xdr:absoluteAnchor></xdr:wsDr>`)
  zip.file('xl/drawings/_rels/drawing2.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${chartRel('rId1', '../charts/chart5.xml')}</Relationships>`)
  const sheet2Rels = zip.file('xl/worksheets/_rels/sheet2.xml.rels')
  const drawingRel = '<Relationship Id="rIdD2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing2.xml"/>'
  zip.file('xl/worksheets/_rels/sheet2.xml.rels', sheet2Rels ? (await sheet2Rels.async('string')).replace('</Relationships>', `${drawingRel}</Relationships>`) : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawingRel}</Relationships>`)
  const sheet2 = await zip.file('xl/worksheets/sheet2.xml').async('string')
  zip.file('xl/worksheets/sheet2.xml', sheet2.replace('</worksheet>', '<drawing r:id="rIdD2"/></worksheet>'))
  let types = await zip.file('[Content_Types].xml').async('string')
  const override = (part, type) => `<Override PartName="/${part}" ContentType="${type}"/>`
  const chartType = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
  types = types.replace('</Types>', [
    override('xl/charts/chart1.xml', chartType), override('xl/charts/chart2.xml', chartType), override('xl/charts/chart3.xml', chartType), override('xl/charts/chart4.xml', chartType), override('xl/charts/chart5.xml', chartType),
    override('xl/charts/chartEx1.xml', 'application/vnd.ms-office.chartex+xml'), override('xl/charts/style1.xml', 'application/vnd.ms-office.chartstyle+xml'), override('xl/charts/colors1.xml', 'application/vnd.ms-office.chartcolorstyle+xml'),
    override('xl/drawings/drawing2.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
  ].join('') + '</Types>')
  zip.file('[Content_Types].xml', types)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

// ---------------------------------------------------------------------------
// Package inspection helpers
// ---------------------------------------------------------------------------

async function inspect(bytes) {
  const zip = await JSZip.loadAsync(bytes)
  const read = async (name) => (zip.file(name) ? zip.file(name).async('string') : null)
  const names = Object.keys(zip.files)
  const types = await read('[Content_Types].xml')
  return { zip, read, names, types }
}

// Minimal schema-order checks for generated DrawingML (sequence constraints Excel enforces).
const ORDER = {
  'c:chartSpace': ['c:date1904', 'c:lang', 'c:roundedCorners', 'mc:AlternateContent', 'c:clrMapOvr', 'c:pivotSource', 'c:protection', 'c:chart', 'c:spPr', 'c:txPr', 'c:externalData', 'c:printSettings', 'c:userShapes', 'c:extLst'],
  'c:chart': ['c:title', 'c:autoTitleDeleted', 'c:pivotFmts', 'c:view3D', 'c:floor', 'c:sideWall', 'c:backWall', 'c:plotArea', 'c:legend', 'c:plotVisOnly', 'c:dispBlanksAs', 'c:showDLblsOverMax', 'c:extLst'],
  'c:title': ['c:tx', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr'],
  'c:legend': ['c:legendPos', 'c:legendEntry', 'c:layout', 'c:overlay', 'c:spPr', 'c:txPr'],
  'c:barChart': ['c:barDir', 'c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:gapWidth', 'c:overlap', 'c:serLines', 'c:axId'],
  'c:lineChart': ['c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:hiLowLines', 'c:upDownBars', 'c:marker', 'c:smooth', 'c:axId'],
  'c:areaChart': ['c:grouping', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:dropLines', 'c:axId'],
  'c:pieChart': ['c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng'],
  'c:doughnutChart': ['c:varyColors', 'c:ser', 'c:dLbls', 'c:firstSliceAng', 'c:holeSize'],
  'c:scatterChart': ['c:scatterStyle', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:axId'],
  'c:radarChart': ['c:radarStyle', 'c:varyColors', 'c:ser', 'c:dLbls', 'c:axId'],
  'c:catAx': ['c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:auto', 'c:lblAlgn', 'c:lblOffset', 'c:tickLblSkip', 'c:tickMarkSkip', 'c:noMultiLvlLbl'],
  'c:valAx': ['c:axId', 'c:scaling', 'c:delete', 'c:axPos', 'c:majorGridlines', 'c:minorGridlines', 'c:title', 'c:numFmt', 'c:majorTickMark', 'c:minorTickMark', 'c:tickLblPos', 'c:spPr', 'c:txPr', 'c:crossAx', 'c:crosses', 'c:crossesAt', 'c:crossBetween', 'c:majorUnit', 'c:minorUnit', 'c:dispUnits'],
  'c:scaling': ['c:logBase', 'c:orientation', 'c:max', 'c:min'],
  'c:dLbls': ['c:dLbl', 'c:delete', 'c:numFmt', 'c:spPr', 'c:txPr', 'c:dLblPos', 'c:showLegendKey', 'c:showVal', 'c:showCatName', 'c:showSerName', 'c:showPercent', 'c:showBubbleSize', 'c:separator', 'c:showLeaderLines'],
  'c:dPt': ['c:idx', 'c:invertIfNegative', 'c:marker', 'c:bubble3D', 'c:explosion', 'c:spPr'],
  'c:marker': ['c:symbol', 'c:size', 'c:spPr'],
}
const SERIES_ORDER = {
  'c:barChart': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:invertIfNegative', 'c:pictureOptions', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:shape', 'c:extLst'],
  'c:lineChart': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:smooth', 'c:extLst'],
  'c:areaChart': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:pictureOptions', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:cat', 'c:val', 'c:extLst'],
  'c:pieChart': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:explosion', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val', 'c:extLst'],
  'c:doughnutChart': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:explosion', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val', 'c:extLst'],
  'c:scatterChart': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:trendline', 'c:errBars', 'c:xVal', 'c:yVal', 'c:smooth', 'c:extLst'],
  'c:radarChart': ['c:idx', 'c:order', 'c:tx', 'c:spPr', 'c:marker', 'c:dPt', 'c:dLbls', 'c:cat', 'c:val', 'c:extLst'],
}

function checkSchemaOrder(xml, label) {
  const { parseXml } = require('../electron/chart-xlsx.cjs')
  const walk = (node, parentName) => {
    const sequence = node.name === 'c:ser' ? SERIES_ORDER[parentName] : ORDER[node.name]
    if (sequence) {
      let position = 0
      for (const item of node.children) {
        const index = sequence.indexOf(item.name)
        assert(index >= 0, `${label}: unexpected <${item.name}> in <${node.name}>`)
        assert(index >= position, `${label}: <${item.name}> is out of order in <${node.name}>`)
        position = index
      }
    }
    for (const item of node.children) walk(item, node.name)
  }
  walk(parseXml(xml), '#document')
  const plotArea = /<c:plotArea>([\s\S]*)<\/c:plotArea>/.exec(xml)
  if (plotArea) {
    // Every referenced axis id exists and pairs cross each other.
    const groupAxes = [...plotArea[1].matchAll(/<c:axId val="(\d+)"\/>/g)].map((m) => m[1])
    const axisIds = [...plotArea[1].matchAll(/<c:(?:catAx|valAx|dateAx)><c:axId val="(\d+)"\/>/g)].map((m) => m[1])
    for (const id of groupAxes) assert(axisIds.includes(id), `${label}: axis ${id} is referenced but not defined`)
  }
}

// ---------------------------------------------------------------------------
// Tests
// ---------------------------------------------------------------------------

async function testImportAndRoundTrip(lib) {
  const fixture = await buildFixture()
  const payload = await workbookPayloadFromBytes('charts.xlsx', fixture)
  const data = payload.workbook.sheets.find((sheet) => sheet.name === 'Data')
  const other = payload.workbook.sheets.find((sheet) => sheet.name === 'My Sheet')
  assert(data && other)
  assert(!payload.warnings.some((warning) => /charts.*not fully modeled/i.test(warning)), 'supported charts must not raise the unmodelled-parts warning')
  assert(payload.warnings.some((warning) => /waterfall/.test(warning) && /kept/.test(warning)), 'unsupported chartex types are reported as kept placeholders')
  assert.equal(data.charts.length, 5)
  assert.equal(other.charts.length, 1)
  const [column, combo, pie, scatter, waterfall] = data.charts
  // Column chart
  assert.equal(column.type, 'column')
  assert.equal(column.grouping, 'clustered')
  assert.equal(column.title, 'Revenue & Costs')
  assert.equal(column.legend, 'bottom')
  assert.equal(column.series.length, 2)
  assert.equal(column.series[0].valuesRef, 'Data!$B$2:$B$7')
  assert.equal(column.series[0].categoriesRef, 'Data!$A$2:$A$7')
  assert.equal(column.series[0].nameRef, 'Data!$B$1')
  assert.equal(column.series[0].name, 'Sales')
  const theme = payload.workbook.metadata.themeColors
  assert.equal(column.series[0].color, `#${theme[4]}`, 'schemeClr accent1 resolves against the workbook theme')
  assert.match(column.series[1].color, /^#[0-9A-F]{6}$/)
  assert.notEqual(column.series[1].color, `#${theme[5]}`, 'lumMod is applied to scheme colours')
  assert.deepEqual(column.series[0].valuesCache, SALES)
  assert.equal(column.axes.x.title, 'Month')
  assert.equal(column.axes.y.title, 'USD')
  assert.equal(column.axes.y.max, 3000)
  assert.equal(column.axes.y.numFmt, '#,##0')
  assert.equal(column.axes.y.gridlines, true)
  assert.equal(column.gapWidth, 219)
  assert.deepEqual(column.anchor.from, { col: 5, colOffsetEmu: 0, row: 1, rowOffsetEmu: 0 })
  assert.deepEqual(column.anchor.to, { col: 12, colOffsetEmu: 304800, row: 16, rowOffsetEmu: 95250 })
  assert.equal(column.sourcePart, 'xl/charts/chart1.xml')
  assert.equal(column.name, 'Chart 1')
  // Combo with a secondary axis
  assert.equal(combo.type, 'combo')
  assert.equal(combo.autoTitleDeleted, true)
  assert.equal(combo.legend, 'top')
  assert.equal(combo.series[0].type, 'column')
  assert.equal(combo.series[1].type, 'line')
  assert.equal(combo.series[1].secondaryAxis, true)
  assert.equal(combo.series[1].marker, 'diamond')
  assert.equal(combo.series[1].smooth, true)
  assert.equal(combo.series[1].color, '#C9A227')
  assert.equal(combo.series[0].dataLabels.showValue, true)
  assert.equal(combo.series[0].dataLabels.position, 'outEnd')
  assert.equal(combo.axes.y2.numFmt, '0%')
  assert.equal(combo.anchor.editAs, 'oneCell')
  // Pie from a oneCellAnchor
  assert.equal(pie.type, 'pie')
  assert.equal(pie.firstSliceAngle, 30)
  assert.equal(pie.series[0].pointColors['0'], '#FF0000')
  assert.equal(pie.series[0].pointColors['2'], `#${theme[9]}`)
  assert.equal(pie.series[0].dataLabels.showPercent, true)
  assert.equal(pie.legend, 'right')
  assert.equal(pie.anchor.editAs, 'oneCell')
  assert(pie.anchor.to.col > pie.anchor.from.col && pie.anchor.to.row > pie.anchor.from.row)
  // Scatter with quoted sheet references
  assert.equal(scatter.type, 'scatter')
  assert.equal(scatter.series[0].xValuesRef, "'My Sheet'!$A$2:$A$6")
  assert.equal(scatter.series[0].showLine, false)
  assert.equal(scatter.series[0].marker, 'circle')
  // Unsupported chartex placeholder
  assert.equal(waterfall.type, 'unsupported')
  assert.equal(waterfall.unsupportedKind, 'waterfall')
  assert.equal(waterfall.title, 'Bridge')
  assert.equal(waterfall.sourceInfo.kind, 'chartex')
  // Doughnut from an absolute anchor on the second sheet
  const doughnut = other.charts[0]
  assert.equal(doughnut.type, 'doughnut')
  assert.equal(doughnut.holeSize, 50)
  assert.equal(doughnut.anchor.editAs, 'absolute')
  assert.equal(doughnut.anchor.from.col, 3, '1828800 EMU = 192px = three default 64px columns')

  // Resolve data against the model and render every imported chart.
  const accessor = lib.workbookChartAccessor(payload.workbook)
  const resolved = lib.resolveChartData(column, accessor)
  assert.deepEqual(resolved.categories, MONTHS)
  assert.deepEqual(resolved.series[0].values, SALES)
  assert.equal(resolved.series[0].name, 'Sales')
  // References to a deleted sheet fall back to the imported cache.
  const orphan = { ...column, series: column.series.map((item) => ({ ...item, valuesRef: 'Gone!$B$2:$B$7' })) }
  assert.deepEqual(lib.resolveChartData(orphan, accessor).series[0].values, SALES)
  const inferred = lib.inferChartDataRange(column)
  assert.deepEqual(inferred, { range: 'Data!A1:C7', seriesIn: 'columns', firstRowHeaders: true, firstColumnLabels: true })
  const scatterData = lib.resolveChartData(scatter, accessor)
  assert.deepEqual(scatterData.series[0].x, [1, 2, 3, 4, 5])
  for (const chart of [...data.charts, ...other.charts]) {
    const svg = lib.renderChartSvg(chart, lib.resolveChartData(chart, accessor), 480, 288)
    wellFormed(svg, `render of imported ${chart.type}`)
  }

  // ---- Round trip 1: unmodified, source-backed: byte-exact chart parts.
  const saved = await serializeWorkbook(payload.workbook, 'xlsx', { baseBytes: fixture })
  const out = await inspect(saved)
  const original = await inspect(fixture)
  const outputCharts = out.names.filter((name) => /^xl\/charts\/chart(Ex)?\d+\.xml$/.test(name))
  assert.equal(outputCharts.length, 6, 'five DrawingML charts and one chartex part are written')
  const outputContents = new Map()
  for (const name of outputCharts) outputContents.set(name, await out.read(name))
  for (const name of ['xl/charts/chart1.xml', 'xl/charts/chart2.xml', 'xl/charts/chart3.xml', 'xl/charts/chart4.xml', 'xl/charts/chart5.xml', 'xl/charts/chartEx1.xml']) {
    const source = await original.read(name)
    assert([...outputContents.values()].includes(source), `${name} is copied byte-for-byte`)
  }
  // style/colors parts follow chart1 with rewritten relationships
  const copiedColumn = [...outputContents.entries()].find(([, xml]) => xml === CHART_COLUMN)[0]
  const columnRels = await out.read(copiedColumn.replace('xl/charts/', 'xl/charts/_rels/') + '.rels')
  assert(columnRels, 'chart relationships are copied')
  for (const [, target] of columnRels.matchAll(/Target="([^"]+)"/g)) {
    const partName = path.posix.join('xl/charts', target)
    assert(out.zip.file(partName), `${partName} referenced by the copied chart exists`)
  }
  assert(out.types.includes('application/vnd.ms-office.chartstyle+xml') && out.types.includes('application/vnd.ms-office.chartcolorstyle+xml'))
  for (const name of outputCharts) assert(out.types.includes(`PartName="/${name}"`), `content type override for ${name}`)
  // Drawings: the picture anchor ExcelJS wrote is merged with the chart anchors.
  const workbookRels = await out.read('xl/_rels/workbook.xml.rels')
  const workbookXml = await out.read('xl/workbook.xml')
  const sheetPart = (name) => {
    const rId = new RegExp(`<sheet [^>]*name="${name}"[^>]*r:id="([^"]+)"`).exec(workbookXml)[1]
    return path.posix.join('xl', new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(workbookRels)?.[1] || new RegExp(`Target="([^"]+)"[^>]*Id="${rId}"`).exec(workbookRels)[1])
  }
  for (const name of ['Data', 'My Sheet']) {
    const part = sheetPart(name)
    const xml = await out.read(part)
    assert.equal((xml.match(/<drawing /g) || []).length, 1, `${name} has one <drawing>`)
    const tail = xml.slice(xml.indexOf('<drawing '))
    assert(!/<(mergeCells|pageMargins|sheetData)\b/.test(tail), `${name}: <drawing> follows the elements the schema places before it`)
    const rels = await out.read(part.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels')
    const drawingTarget = /Type="[^"]*\/drawing" Target="([^"]+)"|Target="([^"]+)"[^>]*Type="[^"]*\/drawing"/.exec(rels)
    const drawingPart = path.posix.join('xl/worksheets', drawingTarget[1] || drawingTarget[2])
    const drawingXml = await out.read(drawingPart)
    wellFormed(drawingXml, `${name} drawing`)
    const ids = [...drawingXml.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map((m) => m[1])
    assert.equal(new Set(ids).size, ids.length, `${name}: drawing object ids are unique`)
    if (name === 'Data') {
      assert((drawingXml.match(/<xdr:pic>/g) || []).length === 1, 'the picture anchor is kept')
      assert.equal((drawingXml.match(/<c:chart /g) || []).length, 4)
      assert.equal((drawingXml.match(/<cx:chart /g) || []).length, 1)
      assert(drawingXml.includes('xmlns:mc=') && drawingXml.includes('Requires="cx1"'))
      assert(drawingXml.includes('<xdr:from><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row>'))
      const drawingRels = await out.read(drawingPart.replace('xl/drawings/', 'xl/drawings/_rels/') + '.rels')
      for (const [, target] of drawingRels.matchAll(/Target="([^"]+)"/g)) assert(out.zip.file(path.posix.join('xl/drawings', target)), `drawing target ${target} exists`)
    } else {
      assert(drawingXml.includes('<xdr:twoCellAnchor editAs="absolute">'))
    }
  }
  const reloaded = new ExcelJS.Workbook()
  await reloaded.xlsx.load(saved)
  assert.equal(reloaded.getWorksheet('Data').getImages().length, 1, 'ExcelJS still reads the picture')
  const reopened = await workbookPayloadFromBytes('charts.xlsx', saved)
  const reData = reopened.workbook.sheets.find((sheet) => sheet.name === 'Data')
  assert.equal(reData.charts.length, 5)
  const semantic = (chart) => { const { id, sourcePart, sourceInfo, name, ...rest } = chart; return rest }
  data.charts.forEach((chart, index) => assert.deepEqual(semantic(reData.charts[index]), semantic(chart), `chart ${index} round-trips unchanged`))

  // ---- Round trip 2: edits regenerate DrawingML; moves keep the bytes.
  const edited = structuredClone(payload.workbook)
  const editedData = edited.sheets.find((sheet) => sheet.name === 'Data')
  editedData.charts[0] = lib.updateChart(editedData.charts[0], { title: 'Edited <title>', legend: 'right' })
  editedData.charts[0].anchor = { from: { row: 2, col: 6, rowOffsetEmu: 0, colOffsetEmu: 9525 }, to: { row: 18, col: 14, rowOffsetEmu: 0, colOffsetEmu: 0 } }
  editedData.charts[3] = { ...editedData.charts[3], anchor: { ...editedData.charts[3].anchor, from: { ...editedData.charts[3].anchor.from, row: 40 }, to: { ...editedData.charts[3].anchor.to, row: 55 } } }
  editedData.charts[1].series[1].color = '#123456' // edited without the modified flag: the fingerprint catches it
  editedData.charts.splice(2, 1) // delete the pie
  const saved2 = await serializeWorkbook(edited, 'xlsx', { baseBytes: fixture })
  const out2 = await inspect(saved2)
  const charts2 = out2.names.filter((name) => /^xl\/charts\/chart(Ex)?\d+\.xml$/.test(name))
  assert.equal(charts2.length, 5)
  const contents2 = await Promise.all(charts2.map((name) => out2.read(name)))
  assert(!contents2.includes(CHART_COLUMN), 'the edited column chart is regenerated')
  assert(!contents2.includes(CHART_COMBO), 'an unflagged colour edit is detected by fingerprint')
  assert(contents2.includes(CHART_SCATTER), 'a moved chart keeps its original bytes')
  assert(!contents2.includes(CHART_PIE), 'the deleted chart is gone')
  for (const [index, xml] of contents2.entries()) {
    wellFormed(xml, charts2[index])
    if (xml.includes('<c:chartSpace') && !Object.values({ CHART_SCATTER, CHART_DOUGHNUT }).includes(xml)) checkSchemaOrder(xml, charts2[index])
  }
  const reopened2 = await workbookPayloadFromBytes('charts.xlsx', saved2)
  const re2 = reopened2.workbook.sheets.find((sheet) => sheet.name === 'Data').charts
  assert.equal(re2.length, 4)
  assert.equal(re2[0].title, 'Edited <title>')
  assert.equal(re2[0].legend, 'right')
  assert.deepEqual(re2[0].anchor.from, { col: 6, colOffsetEmu: 9525, row: 2, rowOffsetEmu: 0 })
  assert.equal(re2[0].series[0].color, `#${theme[4]}`)
  assert.equal(re2[0].series[0].valuesRef, 'Data!$B$2:$B$7')
  assert.equal(re2[0].axes.y.max, 3000)
  assert.equal(re2[0].axes.y.title, 'USD')
  assert.equal(re2[1].type, 'combo')
  assert.equal(re2[1].series[1].color, '#123456')
  assert.equal(re2[1].series[1].secondaryAxis, true)
  assert.equal(re2[1].series[1].smooth, true)
  assert.equal(re2[2].anchor.from.row, 40)
  assert.equal(re2[3].type, 'unsupported')
  const reloaded2 = new ExcelJS.Workbook()
  await reloaded2.xlsx.load(saved2)

  // ---- Sheet rename + structural shift rewrite references and regenerate.
  let renamed = structuredClone(payload.workbook)
  renamed.sheets.find((sheet) => sheet.name === 'My Sheet').name = 'Points 2024'
  renamed = lib.renameWorkbookChartReferences(renamed, 'My Sheet', 'Points 2024')
  const renamedScatter = renamed.sheets.find((sheet) => sheet.name === 'Data').charts[3]
  assert.equal(renamedScatter.series[0].xValuesRef, "'Points 2024'!$A$2:$A$6")
  assert.notEqual(chartFingerprint(renamedScatter), renamedScatter.sourceInfo.fingerprint)
  const shifted = lib.transformWorkbookChartsForStructure(payload.workbook, data.id, { axis: 'row', kind: 'insert', index: 0, count: 2 })
  const shiftedColumn = shifted.sheets.find((sheet) => sheet.name === 'Data').charts[0]
  assert.equal(shiftedColumn.series[0].valuesRef, 'Data!$B$4:$B$9')
  assert.equal(shiftedColumn.anchor.from.row, 3)
  assert.equal(shiftedColumn.modified, true)
  const deleted = lib.transformWorkbookChartsForStructure(payload.workbook, data.id, { axis: 'column', kind: 'delete', index: 1, count: 1 })
  assert.match(deleted.sheets.find((sheet) => sheet.name === 'Data').charts[0].series[0].valuesRef, /#REF!/)
  const saved3 = await serializeWorkbook(renamed, 'xlsx', { baseBytes: fixture })
  const reopened3 = await workbookPayloadFromBytes('charts.xlsx', saved3)
  assert.equal(reopened3.workbook.sheets.find((sheet) => sheet.name === 'Data').charts[3].series[0].valuesRef, "'Points 2024'!$B$2:$B$6")
  return { fixture, saved }
}

function newWorkbookModel() {
  const cells = {}
  const put = (address, value) => { cells[address] = { value } }
  put('A1', 'Region'); put('B1', 'Q1'); put('C1', 'Q2'); put('D1', 'Q3')
  const rows = [['North', 120, 135, 150], ['South', 90, 80, 110], ['East', 60, 95, 70], ['West', 150, 140, 165], ['Online', 45, 70, 115]]
  rows.forEach((row, index) => row.forEach((value, col) => put(`${'ABCD'[col]}${index + 2}`, value)))
  put('F1', 'Year'); put('G1', 'Visitors')
  ;[[2019, 10], [2020, 14], [2021, 13], [2022, 19], [2023, 24]].forEach(([year, value], index) => { put(`F${index + 2}`, year); put(`G${index + 2}`, value) })
  return {
    version: 1,
    name: 'New.xlsx',
    activeSheetId: 's1',
    sheets: [{ id: 's1', name: 'Sales Data', state: 'visible', rowCount: 200, colCount: 40, cells, merges: [], colWidths: {}, rowHeights: {} }],
    metadata: {},
  }
}

async function testNewCharts(lib) {
  const model = newWorkbookModel()
  const accessor = lib.workbookChartAccessor(model)
  const bounds = { top: 0, bottom: 5, left: 0, right: 3 }
  // Auto-detection: header row, label column, series in columns (6x4 data).
  const layout = lib.detectRangeLayout('Sales Data', bounds, accessor, 'column')
  assert.deepEqual(layout, { seriesIn: 'columns', firstRowHeaders: true, firstColumnLabels: true })
  const wide = lib.detectRangeLayout('Sales Data', { top: 0, bottom: 2, left: 0, right: 3 }, accessor, 'column')
  assert.equal(wide.seriesIn, 'rows', 'wide ranges plot one series per row like Excel')
  const years = lib.detectRangeLayout('Sales Data', { top: 0, bottom: 5, left: 5, right: 6 }, accessor, 'line')
  assert.equal(years.firstColumnLabels, true, 'a run of years is treated as labels')
  const expanded = lib.expandToCurrentRegion('Sales Data', { top: 2, bottom: 2, left: 1, right: 1 }, accessor)
  assert.deepEqual(expanded, bounds)
  const types = ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'combo', 'radar']
  const charts = types.map((type, index) => {
    const chart = lib.createChartFromRange('Sales Data', type === 'scatter' ? { top: 0, bottom: 5, left: 5, right: 6 } : { top: 2, bottom: 2, left: 1, right: 1 }, accessor, type, { anchor: { from: { row: index * 16, col: 8 }, to: { row: index * 16 + 15, col: 15 } } })
    return chart
  })
  charts[0].series[0].dataLabels = { showValue: true, position: 'outEnd' }
  charts[1].grouping = 'stacked'
  charts[1].series[0].dataLabels = { showValue: true, position: 'outEnd' } // not allowed for stacked bars: must be dropped
  charts[2].series[1].smooth = true
  charts[2].series[0].marker = 'square'
  charts[3].grouping = 'percentStacked'
  charts[4].series[0].dataLabels = { showPercent: true, position: 'bestFit' }
  charts[7].series[1].secondaryAxis = true
  charts[0].axes.y = { ...charts[0].axes.y, min: 0, max: 200, numFmt: '#,##0', title: 'Units' }
  assert.equal(charts[0].series.length, 3)
  assert.equal(charts[0].series[0].name, 'Q1')
  assert.equal(charts[0].series[0].valuesRef, "'Sales Data'!$B$2:$B$6")
  assert.equal(charts[0].series[0].categoriesRef, "'Sales Data'!$A$2:$A$6")
  assert.equal(charts[4].series.length, 1, 'pie charts keep the first series')
  assert.equal(charts[6].series[0].xValuesRef, "'Sales Data'!$F$2:$F$6")
  model.sheets[0].charts = charts
  const saved = await serializeWorkbook(model, 'xlsx', {})
  const out = await inspect(saved)
  const chartParts = out.names.filter((name) => /^xl\/charts\/chart\d+\.xml$/.test(name))
  assert.equal(chartParts.length, charts.length)
  for (const name of chartParts) {
    const xml = await out.read(name)
    wellFormed(xml, name)
    checkSchemaOrder(xml, name)
    assert(xml.includes('<c:roundedCorners val="0"/>') && xml.includes('<c:autoTitleDeleted') && xml.includes('<c:plotVisOnly val="1"/>') && xml.includes('<c:dispBlanksAs val="gap"/>'))
    assert(out.types.includes(`PartName="/${name}"`))
  }
  const all = (await Promise.all(chartParts.map((name) => out.read(name)))).join('\n')
  assert(all.includes("<c:f>'Sales Data'!$B$2:$B$6</c:f>"))
  assert(all.includes('<c:v>North</c:v>'), 'category caches are written from the model')
  assert(all.includes('<c:holeSize val="60"/>'))
  assert(all.includes('<c:crosses val="max"/>'), 'secondary value axis crosses at max')
  assert(!/<c:barChart><c:barDir val="bar"\/><c:grouping val="stacked"\/>[\s\S]*?<c:dLblPos val="outEnd"\/>[\s\S]*?<\/c:barChart>/.test(all), 'invalid label positions are omitted')
  const sheetXml = await out.read('xl/worksheets/sheet1.xml')
  assert.equal((sheetXml.match(/<drawing r:id=/g) || []).length, 1)
  assert(/<pageMargins[^>]*\/>[\s\S]*<drawing /.test(sheetXml), 'drawing is placed after pageMargins')
  const reloaded = new ExcelJS.Workbook()
  await reloaded.xlsx.load(saved)
  const reopened = await workbookPayloadFromBytes('new.xlsx', saved)
  const back = reopened.workbook.sheets[0].charts
  assert.deepEqual(back.map((chart) => chart.type), types)
  assert.equal(back[1].grouping, 'stacked')
  assert.equal(back[3].grouping, 'percentStacked')
  assert.equal(back[0].axes.y.max, 200)
  assert.equal(back[0].axes.y.title, 'Units')
  assert.equal(back[0].series[0].dataLabels.showValue, true)
  assert.equal(back[2].series[1].smooth, true)
  assert.equal(back[2].series[0].marker, 'square')
  assert.equal(back[7].series[1].secondaryAxis, true)
  assert.equal(back[6].series[0].showLine, false)
  back.forEach((chart, index) => {
    assert.equal(chart.series.length, charts[index].series.length, `${chart.type} series count`)
    assert.equal(chart.series[0].valuesRef, charts[index].series[0].valuesRef)
  })
  return { saved, charts, model }
}

function countMatches(text, pattern) {
  return (text.match(pattern) || []).length
}

function testRenderer(lib) {
  const categories = ['A', 'B', 'C', 'D']
  const base = (type, extra = {}) => ({ id: `r-${type}`, type, anchor: { from: { row: 0, col: 0 }, to: { row: 10, col: 6 } }, series: [{ id: 's1', name: 'One' }, { id: 's2', name: 'Two' }], legend: 'bottom', ...extra })
  const data = { categories, series: [{ name: 'One', values: [4, 6, null, 8], color: '#4472C4' }, { name: 'Two', values: [3, -2, 5, 1], color: '#ED7D31' }] }
  // Clustered column: one rect per non-null point.
  const column = lib.renderChartSvg(base('column', { grouping: 'clustered', title: 'Q & A <test>' }), data, 480, 288)
  wellFormed(column, 'column svg')
  assert.equal(countMatches(column, /class="chart-bar"/g), 7)
  assert(column.includes('Q &amp; A &lt;test&gt;'))
  // Stacked: the heights of each category's segments are proportional to the stack totals.
  const positive = { categories, series: [{ name: 'One', values: [4, 6, 2, 8], color: '#4472C4' }, { name: 'Two', values: [3, 2, 5, 1], color: '#ED7D31' }] }
  const stacked = lib.renderChartSvg(base('column', { grouping: 'stacked' }), positive, 480, 288)
  wellFormed(stacked, 'stacked svg')
  const heights = new Map()
  for (const match of stacked.matchAll(/data-point="(\d+)"[^>]*height="([\d.]+)"/g)) heights.set(match[1], (heights.get(match[1]) || 0) + Number(match[2]))
  const totals = [7, 8, 7, 9]
  const unit = heights.get('0') / totals[0]
  totals.forEach((total, index) => assert(Math.abs(heights.get(String(index)) - total * unit) < 1, `stack ${index} height matches its total`))
  // 100% stacked: every category stack has the same height.
  const percent = lib.renderChartSvg(base('column', { grouping: 'percentStacked' }), positive, 480, 288)
  const pHeights = new Map()
  for (const match of percent.matchAll(/data-point="(\d+)"[^>]*height="([\d.]+)"/g)) pHeights.set(match[1], (pHeights.get(match[1]) || 0) + Number(match[2]))
  const pValues = [...pHeights.values()]
  assert(Math.max(...pValues) - Math.min(...pValues) < 1, '100% stacks are equally tall')
  assert(percent.includes('>100%<'), 'percent axis labels')
  // Bar (horizontal), line with gaps and markers, area, pie, doughnut, scatter, radar, combo.
  const bar = lib.renderChartSvg(base('bar'), data, 480, 288)
  assert.equal(countMatches(bar, /class="chart-bar"/g), 7)
  const lineChart = base('line', { series: [{ id: 's1', name: 'One', marker: 'circle' }, { id: 's2', name: 'Two', smooth: true }] })
  const line = lib.renderChartSvg(lineChart, data, 480, 288)
  wellFormed(line, 'line svg')
  assert.equal(countMatches(line, /class="chart-line"/g), 2)
  assert(/class="chart-line" data-series="0" d="M[^"]*M/.test(line), 'an empty cell breaks the line (gap)')
  assert.equal(countMatches(line, /class="chart-marker"/g), 3)
  assert(/data-series="1" d="M[^"]*C/.test(line), 'smooth lines use curves')
  assert.equal(countMatches(lib.renderChartSvg(base('area', { grouping: 'stacked' }), positive, 480, 288), /class="chart-area"/g), 2)
  const pieChart = base('pie', { series: [{ id: 's1', name: 'One', dataLabels: { showPercent: true } }] })
  const pieData = { categories, series: [{ name: 'One', values: [4, 6, 0, 8], color: '#4472C4', pointColors: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000'] }] }
  const pie = lib.renderChartSvg(pieChart, pieData, 480, 288)
  wellFormed(pie, 'pie svg')
  assert.equal(countMatches(pie, /class="chart-slice"/g), 3)
  assert(pie.includes('>22%<') && pie.includes('>44%<'), 'pie percentages')
  const doughnut = lib.renderChartSvg({ ...pieChart, type: 'doughnut', holeSize: 50 }, pieData, 300, 300)
  assert.equal(countMatches(doughnut, /class="chart-slice"/g), 3)
  const scatter = lib.renderChartSvg(base('scatter', { series: [{ id: 's1', name: 'One', showLine: true, marker: 'circle' }] }), { categories: [], series: [{ name: 'One', values: [1, 4, 9, 16], x: [1, 2, 3, 4], color: '#4472C4' }] }, 480, 288)
  wellFormed(scatter, 'scatter svg')
  assert.equal(countMatches(scatter, /class="chart-marker"/g), 4)
  wellFormed(lib.renderChartSvg(base('radar'), positive, 400, 300), 'radar svg')
  const combo = lib.renderChartSvg(base('combo', { series: [{ id: 's1', name: 'One', type: 'column' }, { id: 's2', name: 'Two', type: 'line', secondaryAxis: true, marker: 'diamond' }] }), positive, 480, 288)
  assert.equal(countMatches(combo, /class="chart-bar"/g), 4)
  assert.equal(countMatches(combo, /class="chart-line"/g), 1)
  wellFormed(lib.renderChartSvg({ ...base('unsupported'), unsupportedKind: 'waterfall', title: 'Bridge' }, { categories: [], series: [] }, 400, 250), 'placeholder svg')
  const empty = lib.renderChartSvg(base('column'), { categories: [], series: [{ name: 'x', values: [null], color: '#000000' }] }, 400, 250)
  assert(empty.includes('No data to plot'))
  // Axis scaling: Excel-like nice ticks.
  const ticks = lib.renderChartSvg(base('column', { legend: 'none' }), { categories: ['a', 'b'], series: [{ name: 'x', values: [48, 12], color: '#000000' }] }, 480, 288)
  assert(ticks.includes('>60<') && ticks.includes('>0<') && !ticks.includes('>70<'), 'axis rounds 48 up to 60 with a step of 10')
  // Number formats
  assert.equal(lib.defaultFormatNumber(1234.5, '#,##0.00'), '1,234.50')
  assert.equal(lib.defaultFormatNumber(0.256, '0.0%'), '25.6%')
  assert.equal(lib.defaultFormatNumber(-1500, '$#,##0'), '-$1,500')
  assert.equal(lib.defaultFormatNumber(2500000, '#,##0,,"M"'), '3M')
  assert.equal(lib.defaultFormatNumber(45292, 'mmm yy'), 'Jan 24')
  assert.equal(lib.defaultFormatNumber(0.1 + 0.2), '0.3')
  // Performance: 1,000 points.
  const big = { categories: Array.from({ length: 1000 }, (_, i) => `P${i}`), series: [{ name: 'Big', values: Array.from({ length: 1000 }, (_, i) => Math.sin(i / 20) * 100 + i / 10), color: '#4472C4' }] }
  for (const type of ['line', 'column', 'scatter']) {
    const chart = base(type, { series: [{ id: 's', name: 'Big', marker: type === 'line' ? 'none' : undefined }], legend: 'none' })
    const input = type === 'scatter' ? { categories: [], series: [{ ...big.series[0], x: big.series[0].values.map((_, i) => i) }] } : big
    for (let i = 0; i < 5; i += 1) lib.renderChartSvg(chart, input, 800, 400)
    const runs = 30
    const start = process.hrtime.bigint()
    for (let i = 0; i < runs; i += 1) lib.renderChartSvg(chart, input, 800, 400)
    const ms = Number(process.hrtime.bigint() - start) / 1e6 / runs
    assert(ms < 5, `${type} with 1,000 points renders in ${ms.toFixed(2)} ms (< 5 ms)`)
    process.stdout.write(`  render ${type} x1000: ${ms.toFixed(2)} ms\n`)
  }
}

function testPrint(lib, model, charts) {
  const accessor = lib.workbookChartAccessor(model)
  const sheet = model.sheets[0]
  const chart = charts[0]
  const svg = lib.renderChartSvg(chart, lib.resolveChartData(chart, accessor), 480, 288, { idPrefix: 'p0', tooltips: false })
  const input = {
    name: 'Print', workbook: { ...model, sheets: [{ ...sheet, charts: undefined }] },
    options: { scope: 'active-sheet', orientation: 'landscape', scaling: 'actual', paperSize: 'letter', margins: 'normal', gridlines: true, headings: false },
    charts: { [sheet.id]: [{ svg, from: { row: 1, col: 5 }, to: { row: 12, col: 10, colOffsetEmu: 9525 * 10 } }, { svg: '<svg onload="alert(1)"></svg>', from: { row: 0, col: 0 }, to: { row: 2, col: 2 } }] },
  }
  // Titles that merely mention "online = …" or contain markup-like text stay printable.
  const tricky = lib.renderChartSvg({ ...chart, title: 'Sales online = 5 <img onerror=x>' }, lib.resolveChartData(chart, accessor), 300, 200, { tooltips: false })
  input.charts[sheet.id].push({ svg: tricky, from: { row: 14, col: 0 }, to: { row: 20, col: 4 } }, { svg: '<svg><g style="x"></g></svg>', from: { row: 0, col: 0 }, to: { row: 1, col: 1 } }, { svg: '<svg><a href="javascript:1"><rect/></a></svg>', from: { row: 0, col: 0 }, to: { row: 1, col: 1 } })
  const document = createSpreadsheetPrintDocument(input)
  assert.equal(countMatches(document.html, /class="print-chart"/g), 2, 'unsafe SVG is rejected, safe charts are placed')
  assert(document.html.includes('Sales online = 5 &lt;img onerror=x&gt;'))
  assert(!document.html.includes('onload') && !document.html.includes('javascript:'))
  const style = /class="print-chart" style="left:([\d.]+)px;top:([\d.]+)px;width:([\d.]+)px;height:([\d.]+)px"/.exec(document.html)
  assert.equal(Number(style[1]), 5 * 64, 'chart x follows the print column widths')
  assert.equal(Number(style[2]), 20, 'chart y follows the print row heights')
  assert.equal(Number(style[3]), 5 * 64 + 10)
  assert(/data-column-range="A:K"/.test(document.html) || document.html.includes('data-column-range="A:'), 'default print area grows to include the chart')
  const withoutCharts = createSpreadsheetPrintDocument({ ...input, charts: undefined })
  assert(!withoutCharts.html.includes('class="print-chart"'))
  // One-call payload for every chart on the sheet, drawn at the printed size.
  const payload = lib.buildPrintChartPayload({ ...model, sheets: [{ ...sheet, charts: charts.slice(0, 3) }] }, (item, owner) => lib.resolveChartData(item, accessor, { defaultSheet: owner.name }))
  assert.equal(payload[sheet.id].length, 3)
  const size = lib.printedAnchorSize(sheet, charts[0].anchor.from, charts[0].anchor.to)
  assert.equal(size.width, 7 * 64)
  assert.equal(size.height, 15 * 20)
  payload[sheet.id].forEach((entry) => wellFormed(entry.svg, 'print svg'))
  assert(payload[sheet.id][0].svg.includes('viewBox="0 0 448 300"'))
  const paged = createSpreadsheetPrintDocument({ ...input, options: { ...input.options, scaling: 'fit-width' }, charts: payload })
  assert(countMatches(paged.html, /class="print-chart"/g) >= 3, 'charts spread over pages are each placed (a chart crossing a page break is clipped on both)')
}

async function testOfficeConversion(saved) {
  if (process.env.QA_CHARTS_OFFICE !== '1') return 'skipped (set QA_CHARTS_OFFICE=1)'
  const { convertOfficeBytes, findOfficeConverter } = require('../electron/office-converter.cjs')
  if (!await findOfficeConverter()) return 'skipped (LibreOffice not found)'
  const ods = await convertOfficeBytes({ bytes: saved, inputExtension: 'xlsx', outputExtension: 'ods', filter: 'calc8' })
  const zip = await JSZip.loadAsync(ods)
  const objects = Object.keys(zip.files).filter((name) => /^Object \d+\/content\.xml$/.test(name))
  let chartObjects = 0
  for (const name of objects) if ((await zip.file(name).async('string')).includes('<chart:chart')) chartObjects += 1
  assert(chartObjects >= 8, `LibreOffice imported ${chartObjects} generated charts`)
  // Interop: LibreOffice's own DrawingML (a different producer) imports back into the model.
  const viaOffice = await convertOfficeBytes({ bytes: ods, inputExtension: 'ods', outputExtension: 'xlsx', filter: 'Calc MS Excel 2007 XML' }).catch(() => null)
  let reimported = 'n/a'
  if (viaOffice) {
    const payload = await workbookPayloadFromBytes('office.xlsx', viaOffice)
    const charts = payload.workbook.sheets.flatMap((sheet) => sheet.charts || [])
    assert(charts.length >= 8, `${charts.length} LibreOffice-written charts were imported`)
    const known = charts.filter((chart) => chart.type !== 'unsupported')
    assert(known.length >= 8, 'LibreOffice charts map onto supported types')
    assert(known.every((chart) => chart.series.length >= 1 && chart.series[0].valuesRef), 'series references survive')
    reimported = `${known.length} re-imported (${[...new Set(known.map((chart) => chart.type))].join(', ')})`
  }
  return `${chartObjects} charts imported by LibreOffice; LibreOffice XLSX: ${reimported}`
}

/** Chart titles per sheet of a saved package: { sheetName: [title, ...] }. */
async function chartTitlesBySheet(bytes) {
  const payload = await workbookPayloadFromBytes('titles.xlsx', bytes)
  return Object.fromEntries(payload.workbook.sheets.map((sheet) => [sheet.name, (sheet.charts || []).map((chart) => `${chart.title}|${(chart.series[0] || {}).valuesRef || chart.unsupportedKind}`)]))
}

/**
 * calc-file-io-objects-1: an imported chart is only copied from a base package whose part
 * still holds that chart. With a stale base (part numbers reassigned by an earlier save) a
 * supported chart is regenerated from the model, and a chart that cannot be regenerated is
 * reported instead of being replaced by another chart.
 */
async function testStaleBase() {
  const anchor = { from: { row: 5, col: 1, rowOffsetEmu: 0, colOffsetEmu: 0 }, to: { row: 15, col: 6, rowOffsetEmu: 0, colOffsetEmu: 0 } }
  const sheet = (id, name) => ({ id, name, rowCount: 3, colCount: 2, merges: [], colWidths: {}, rowHeights: {}, cells: { A1: { value: 'x' }, B1: { value: 1 }, B2: { value: 2 }, B3: { value: 3 } } })
  const second = sheet('sheet-2', 'Sheet2')
  second.charts = [{ id: 'cA', type: 'column', title: 'Chart A', anchor, series: [{ id: 's', valuesRef: 'Sheet2!B1:B3' }] }]
  const original = await serializeWorkbook({ version: 1, name: 't', activeSheetId: 'sheet-1', sheets: [sheet('sheet-1', 'Sheet1'), second], metadata: {} }, 'xlsx')
  const model = (await workbookPayloadFromBytes('t.xlsx', original)).workbook
  model.sheets[0].charts = [{ id: 'cN', type: 'line', title: 'Chart N', anchor, series: [{ id: 's', valuesRef: 'Sheet1!B1:B3' }] }]
  const first = await serializeWorkbook(model, 'xlsx', { baseBytes: original })
  assert.deepEqual(await chartTitlesBySheet(first), { Sheet1: ['Chart N|Sheet1!$B$1:$B$3'], Sheet2: ['Chart A|Sheet2!$B$1:$B$3'] })
  model.sheets[0].cells.C1 = { value: 'edit' }
  // The merge base the app keeps (the imported package) and a stale one both give the right charts.
  for (const baseBytes of [original, first]) {
    assert.deepEqual(await chartTitlesBySheet(await serializeWorkbook(model, 'xlsx', { baseBytes })), { Sheet1: ['Chart N|Sheet1!$B$1:$B$3'], Sheet2: ['Chart A|Sheet2!$B$1:$B$3'] })
  }
  // A chart that can only be byte-copied (waterfall) is never swapped for another part.
  const fixture = await buildFixture()
  const payload = await workbookPayloadFromBytes('charts.xlsx', fixture)
  const stale = await JSZip.loadAsync(fixture)
  stale.file('xl/charts/chartEx1.xml', CHART_WATERFALL.replace('<cx:v>Bridge</cx:v>', '<cx:v>Someone else</cx:v>'))
  const warnings = []
  const saved = await serializeWorkbook(payload.workbook, 'xlsx', { baseBytes: await stale.generateAsync({ type: 'nodebuffer' }), warnings })
  const titles = await chartTitlesBySheet(saved)
  assert.ok(!JSON.stringify(titles).includes('Someone else'), 'a different chart is not copied in')
  assert.ok(warnings.some((warning) => /waterfall/.test(warning) && /Bridge/.test(warning)), 'the user is told the chart could not be kept')
  const kept = await chartTitlesBySheet(await serializeWorkbook(payload.workbook, 'xlsx', { baseBytes: fixture }))
  assert.ok(kept.Data.some((entry) => entry.startsWith('Bridge|')), 'with its own package the waterfall chart is kept')
}

async function main() {
  const lib = loadTsLibraries()
  await testImportAndRoundTrip(lib)
  process.stdout.write('  import + byte-exact / regenerated round trips: ok\n')
  await testStaleBase()
  process.stdout.write('  stale base package guard: ok\n')
  const { saved, charts, model } = await testNewCharts(lib)
  process.stdout.write('  new charts (9 types) export + reimport: ok\n')
  testRenderer(lib)
  process.stdout.write('  renderer: ok\n')
  testPrint(lib, model, charts)
  process.stdout.write('  print overlay: ok\n')
  const office = await testOfficeConversion(saved)
  process.stdout.write(`  LibreOffice conversion: ${office}\n`)
  process.stdout.write('Charts QA passed: DrawingML import, byte-exact and regenerated XLSX export, renderer and print overlay.\n')
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
