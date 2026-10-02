'use strict'
const test=require('node:test')
const assert=require('node:assert/strict')
const {createCanvas}=require('@napi-rs/canvas')
const {PDFDocument,PDFName,PDFRawStream}=require('pdf-lib')
const {imageToPdfBytes,jpegOrientation}=require('../electron/image-to-pdf.cjs')

function withOrientation(bytes,orientation,little=true) {
  const metadata=Buffer.alloc(32)
  metadata.write('Exif\0\0',0,'ascii')
  metadata.write(little?'II':'MM',6,'ascii')
  const put16=(value,position)=>little?metadata.writeUInt16LE(value,position):metadata.writeUInt16BE(value,position)
  const put32=(value,position)=>little?metadata.writeUInt32LE(value,position):metadata.writeUInt32BE(value,position)
  put16(42,8);put32(8,10);put16(1,14);put16(0x112,16);put16(3,18);put32(1,20);put16(orientation,24)
  return Buffer.concat([bytes.subarray(0,2),Buffer.from([0xff,0xe1,0,34]),metadata,bytes.subarray(2)])
}

function fixture() {
  const canvas=createCanvas(80,40)
  const context=canvas.getContext('2d')
  for(const [color,x,y] of [['#ff0000',0,0],['#00ff00',40,0],['#0000ff',0,20],['#ffff00',40,20]]) {
    context.fillStyle=color;context.fillRect(x,y,40,20)
  }
  return canvas.toBuffer('image/jpeg',100)
}

test('EXIF orientation parser handles both byte orders and damaged metadata safely',()=>{
  for(const little of [true,false]) for(let orientation=1;orientation<=8;orientation++) {
    const bytes=withOrientation(fixture(),orientation,little)
    assert.equal(jpegOrientation(bytes),orientation)
    for(let length=0;length<38;length++) assert.equal(jpegOrientation(bytes.subarray(0,length)),1)
  }
  const corrupt=withOrientation(fixture(),6)
  corrupt.writeUInt32LE(0xffffffff,16)
  assert.equal(jpegOrientation(corrupt),1)
})

test('all eight camera orientations render correct corner colors without recompressing JPEGs',async()=>{
  const pdfjs=await import('pdfjs-dist/legacy/build/pdf.mjs')
  const expected={1:['r','g','b','y'],2:['g','r','y','b'],3:['y','b','g','r'],4:['b','y','r','g'],5:['r','b','g','y'],6:['b','r','y','g'],7:['y','g','b','r'],8:['g','y','r','b']}
  for(let orientation=1;orientation<=8;orientation++) {
    const jpeg=withOrientation(fixture(),orientation,orientation%2===0)
    const backing=Buffer.alloc(jpeg.length+63);jpeg.copy(backing,63)
    const output=await imageToPdfBytes(backing.subarray(63),'.jpg')
    const document=await PDFDocument.load(output)
    const resources=document.getPage(0).node.Resources().lookup(PDFName.of('XObject'))
    const image=resources.lookup(resources.keys()[0])
    assert.ok(image instanceof PDFRawStream)
    assert.deepEqual(Buffer.from(image.getContents()),jpeg,'compressed source bytes must be retained')
    const pdf=await pdfjs.getDocument({data:output.slice(),disableWorker:true,isEvalSupported:false}).promise
    try {
      const page=await pdf.getPage(1)
      const viewport=page.getViewport({scale:1})
      assert.deepEqual([viewport.width,viewport.height],orientation<5?[80,40]:[40,80])
      const canvas=createCanvas(viewport.width,viewport.height)
      const context=canvas.getContext('2d')
      await page.render({canvasContext:context,viewport}).promise
      const colors=[[.25,.25],[.75,.25],[.25,.75],[.75,.75]].map(([x,y])=>{
        const [r,g,b]=context.getImageData(Math.floor(x*canvas.width),Math.floor(y*canvas.height),1,1).data
        return r>150&&g>150?'y':r>150?'r':g>150?'g':b>150?'b':'?'
      })
      assert.deepEqual(colors,expected[orientation],`EXIF ${orientation}`)
    } finally {await pdf.destroy()}
  }
})

test('image types are recognised from their bytes, not their names', () => {
  const {sniffImageType}=require('../electron/image-to-pdf.cjs')
  const {bmp,gif,svg,tiff}=require('./helpers/import-fixtures.cjs')
  assert.equal(sniffImageType(fixture()),'jpeg')
  assert.equal(sniffImageType(createCanvas(4,4).toBuffer('image/png')),'png')
  assert.equal(sniffImageType(createCanvas(4,4).toBuffer('image/webp')),'webp')
  assert.equal(sniffImageType(gif()),'gif')
  assert.equal(sniffImageType(bmp()),'bmp')
  assert.equal(sniffImageType(tiff([4])),'tiff')
  assert.equal(sniffImageType(svg()),'svg')
  assert.equal(sniffImageType(Buffer.from('BMW owners club newsletter')),null)
  assert.equal(sniffImageType(Buffer.from('<html><body></body></html>')),null)
})

test('SVG pages take the drawing size from width, height or viewBox', () => {
  const {svgDimensions}=require('../electron/image-to-pdf.cjs')
  assert.deepEqual(svgDimensions('<svg width="200" height="100">'),{width:150,height:75})
  assert.deepEqual(svgDimensions('<svg width="2in" height="1in">'),{width:144,height:72})
  assert.deepEqual(svgDimensions('<svg viewBox="0 0 400 200">'),{width:300,height:150})
  assert.deepEqual(svgDimensions('<svg width="100pt" viewBox="0 0 400 200">'),{width:100,height:50})
  assert.deepEqual(svgDimensions('<svg>'),{width:225,height:112.5})
})

test('a PNG named .jpg is embedded as the PNG it is', async () => {
  const output=await imageToPdfBytes(createCanvas(30,10).toBuffer('image/png'),'.jpg')
  const document=await PDFDocument.load(output)
  assert.deepEqual([document.getPage(0).getWidth(),document.getPage(0).getHeight()],[30,10])
})

test('TIFF pages convert without a decoder dependency: one page per image, fax data passed through', async () => {
  const {tiffFirstImage}=require('../electron/image-to-pdf.cjs')
  const {tiff}=require('./helpers/import-fixtures.cjs')
  const pdfjs=await import('pdfjs-dist/legacy/build/pdf.mjs')
  const corners=async (bytes,pageNumber=1) => {
    const pdf=await pdfjs.getDocument({data:new Uint8Array(bytes),disableWorker:true,isEvalSupported:false}).promise
    try {
      const page=await pdf.getPage(pageNumber)
      const viewport=page.getViewport({scale:4})
      const canvas=createCanvas(Math.round(viewport.width),Math.round(viewport.height))
      const context=canvas.getContext('2d')
      context.fillStyle='#808080';context.fillRect(0,0,canvas.width,canvas.height)
      await page.render({canvasContext:context,viewport}).promise
      return [[.2,.5],[.8,.5]].map(([x,y])=>{
        const [r,g,b]=context.getImageData(Math.floor(x*canvas.width),Math.floor(y*canvas.height),1,1).data
        return r>150&&g<100?'red':b>150&&r<100?'blue':r<60&&g<60&&b<60?'black':r>200&&g>200&&b>200?'white':`${r},${g},${b}`
      })
    } finally {await pdf.destroy()}
  }
  const multi=await imageToPdfBytes(tiff([40,30,20]),'.tif')
  const document=await PDFDocument.load(multi)
  assert.equal(document.getPageCount(),3)
  for (let page=1;page<=3;page++) assert.deepEqual(await corners(multi,page),['red','blue'])
  // Group 4 fax strips stay CCITT-compressed in the PDF and keep their 72 dpi size.
  const fax=Buffer.from('SUkqABQAAAAzF/ABABAzF/ABABAMAAABAwABAAAAEAAAAAEBAwABAAAACAAAAAIBAwABAAAAAQAAAAMBAwABAAAABAAAAAYBAwABAAAAAQAAABEBBAACAAAAwgAAABYBAwABAAAABAAAABcBBAACAAAAugAAABoBBQABAAAAqgAAABsBBQABAAAAsgAAABwBAwABAAAAAQAAACgBAwABAAAAAgAAAAAAAABIAAAAAQAAAEgAAAABAAAABgAAAAYAAAAIAAAADgAAAA==','base64')
  const faxPdf=await imageToPdfBytes(fax,'.tif')
  const faxDocument=await PDFDocument.load(faxPdf)
  assert.deepEqual([faxDocument.getPage(0).getWidth(),faxDocument.getPage(0).getHeight()],[16,8])
  assert.match(Buffer.from(faxPdf).toString('latin1'),/\/CCITTFaxDecode/)
  assert.deepEqual(await corners(faxPdf),['black','white'])
  const lzw=Buffer.from('SUkqADAAAACANwKB6BQSBwUKChpQmFwqGQWIQeBwyKQ6FRGDRmKw2ORiMRuNwEAACgAAAQMAAQAAAAgAAAABAQMAAQAAAAQAAAACAQMAAwAAAK4AAAADAQMAAQAAAAUAAAAGAQMAAQAAAAIAAAARAQQAAQAAAAgAAAAVAQMAAQAAAAMAAAAWAQMAAQAAAAQAAAAXAQQAAQAAACcAAAAcAQMAAQAAAAEAAAAAAAAACAAIAAgA','base64')
  assert.deepEqual(await corners(await imageToPdfBytes(lzw,'.tif')),['red','blue'])
  const picture=tiffFirstImage(lzw)
  assert.equal(picture.mime,'image/png')
  const {loadImage}=require('@napi-rs/canvas')
  const image=await loadImage(picture.data)
  assert.deepEqual([image.width,image.height],[8,4])
  // BigTIFF is reported plainly instead of failing obscurely.
  await assert.rejects(imageToPdfBytes(Buffer.from('MM\u0000\u002b\u0000\u0008\u0000\u0000','latin1'),'.tif'),{code:'UNSUPPORTED_IMAGE'})
})
