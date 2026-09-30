'use strict'
const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs')
const path = require('node:path')
const vm = require('node:vm')
const ts = require('typescript')

function harness() {
  const canvases = []
  let encodes = 0
  const pdfjs = {
    OPS: { save:1, restore:2, transform:3, paintImageXObject:4 },
    Util: {
      transform: (left,right) => [left[0]*right[0]+left[2]*right[1],left[1]*right[0]+left[3]*right[1],left[0]*right[2]+left[2]*right[3],left[1]*right[2]+left[3]*right[3],left[0]*right[4]+left[2]*right[5]+left[4],left[1]*right[4]+left[3]*right[5]+left[5]],
      getAxialAlignedBoundingBox: (_box,matrix) => [matrix[4],matrix[5],matrix[4]+matrix[0],matrix[5]+matrix[3]],
    },
  }
  const exports = {}
  const document = { createElement: () => {
    const canvas = { width:0,height:0,getContext:()=>({drawImage(){},setTransform(){}}),toDataURL:()=>{encodes++;return `data:image/png;base64,${encodes}`} }
    canvases.push(canvas)
    return canvas
  } }
  const source = fs.readFileSync(path.join(__dirname,'../src/lib/pageObjects.ts'),'utf8')
  const compiled = ts.transpileModule(source,{compilerOptions:{module:ts.ModuleKind.CommonJS,target:ts.ScriptTarget.ES2022}}).outputText
  vm.runInNewContext(compiled,{exports,require:()=>({pdfjs}),document})
  function page(count=30,size=80) {
    const fnArray=[],argsArray=[]
    for(let index=0;index<count;index++) {
      fnArray.push(1,3,4,2)
      argsArray.push([], [size,0,0,size,index*(size+5),0], ['photo'], [])
    }
    return {getOperatorList:async()=>({fnArray,argsArray}),objs:{get:()=>({width:100,height:100,bitmap:{}})}}
  }
  return {detect:exports.detectPageObjects,page,canvases,encodes:()=>encodes}
}

test('image-heavy Edit detects every bound without allocating or encoding a canvas', async () => {
  const h=harness()
  const page=h.page()
  const objects=await h.detect(page,0)
  assert.equal(objects.length,30)
  assert.equal(h.canvases.length,0)
  assert.equal(h.encodes(),0)
  const image=objects[12].dataUrl
  assert.match(image,/^data:image\/png/)
  assert.equal(h.encodes(),1)
  assert.equal(objects[12].dataUrl,image)
  assert.equal(h.encodes(),1,'reselecting the same image reuses its materialization')
  assert.equal(await h.detect(page,0),objects,'reentering Edit reuses detection')
  assert.equal(h.encodes(),1)
})

test('selected oversized PDF image bounds cannot allocate an unbounded output canvas', async () => {
  const h=harness()
  const [object]=await h.detect(h.page(1,20_000),0)
  assert.ok(object.dataUrl)
  const placed=h.canvases[1]
  assert.ok(placed.width<=8192&&placed.height<=8192)
  assert.ok(placed.width*placed.height<=16_000_000)
})
