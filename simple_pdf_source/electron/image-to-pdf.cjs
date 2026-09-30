'use strict'

/** Read the bounded EXIF IFD0 Orientation tag; malformed metadata is optional. */
function jpegOrientation(value) {
  const bytes = Buffer.isBuffer(value) ? value : Buffer.from(value)
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return 1
  let offset = 2
  while (offset + 4 <= bytes.length) {
    if (bytes[offset] !== 0xff) return 1
    while (bytes[offset + 1] === 0xff) offset++
    const marker = bytes[offset + 1]
    if (marker === 0xda || marker === 0xd9) return 1
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) { offset += 2; continue }
    if (offset + 4 > bytes.length) return 1
    const size = bytes.readUInt16BE(offset + 2)
    const end = offset + 2 + size
    if (size < 2 || end > bytes.length) return 1
    const start = offset + 4
    if (marker === 0xe1 && end - start >= 14 && bytes.toString('ascii',start,start+6) === 'Exif\0\0') {
      const tiff = start + 6
      const endian = bytes.toString('ascii',tiff,tiff+2)
      if (endian !== 'II' && endian !== 'MM') return 1
      const read16 = (position) => endian === 'II' ? bytes.readUInt16LE(position) : bytes.readUInt16BE(position)
      const read32 = (position) => endian === 'II' ? bytes.readUInt32LE(position) : bytes.readUInt32BE(position)
      if (read16(tiff+2) !== 42) return 1
      const directory = tiff + read32(tiff+4)
      if (directory < tiff+8 || directory+2 > end) return 1
      const count = read16(directory)
      if (count > Math.floor((end-directory-2)/12)) return 1
      for (let index=0;index<count;index++) {
        const entry=directory+2+index*12
        if (read16(entry) !== 0x0112) continue
        if (read16(entry+2) !== 3 || read32(entry+4) !== 1) return 1
        const orientation=read16(entry+8)
        return orientation>=1&&orientation<=8 ? orientation : 1
      }
    }
    offset=end
  }
  return 1
}

/** PDF uses a bottom-left origin; EXIF describes top-left display coordinates. */
function orientationPlacement(orientation,width,height) {
  const matrices = {
    1:[1,0,0,1,0,0], 2:[-1,0,0,1,width,0],
    3:[-1,0,0,-1,width,height], 4:[1,0,0,-1,0,height],
    5:[0,-1,-1,0,height,width], 6:[0,-1,1,0,0,width],
    7:[0,1,1,0,0,0], 8:[0,1,-1,0,height,0],
  }
  const swapped=orientation>=5&&orientation<=8
  return { width:swapped?height:width, height:swapped?width:height, matrix:matrices[orientation]||matrices[1] }
}

async function addImagePage(pdfDoc, value, extension, { maxDimension = 841.89 } = {}) {
  const { pushGraphicsState, popGraphicsState, concatTransformationMatrix } = require('pdf-lib')
  const lower=String(extension).toLowerCase()
  if (!['.png','.jpg','.jpeg'].includes(lower)) throw new Error('Choose a PNG or JPEG image to add to the PDF.')
  // pdf-lib's JPEG parser reads from buffer offset zero, so pooled Node buffers
  // and sliced typed arrays must be copied into an exact independent byte view.
  const bytes=new Uint8Array(value)
  const image=lower==='.png' ? await pdfDoc.embedPng(bytes) : await pdfDoc.embedJpg(bytes)
  if (!(image.width>0&&image.height>0)) throw new Error('The image dimensions are invalid.')
  const limit=Number.isFinite(maxDimension)&&maxDimension>0 ? maxDimension : 841.89
  const scale=Math.min(1,limit/image.width,limit/image.height)
  const width=image.width*scale
  const height=image.height*scale
  const placement=orientationPlacement(lower==='.png'?1:jpegOrientation(bytes),width,height)
  const page=pdfDoc.addPage([placement.width,placement.height])
  page.pushOperators(pushGraphicsState(),concatTransformationMatrix(...placement.matrix))
  page.drawImage(image,{x:0,y:0,width,height})
  page.pushOperators(popGraphicsState())
  return page
}

async function imageToPdfBytes(buffer,extension,title='Converted image') {
  const {PDFDocument}=require('pdf-lib')
  const pdfDoc=await PDFDocument.create()
  await addImagePage(pdfDoc,buffer,extension)
  pdfDoc.setTitle(title)
  pdfDoc.setCreator('simple')
  return pdfDoc.save()
}

module.exports={addImagePage,imageToPdfBytes,jpegOrientation,orientationPlacement}
