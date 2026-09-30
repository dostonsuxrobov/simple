const test=require('node:test')
const assert=require('node:assert/strict')
test('page canvas limits apply below 1x backing scale as well as on extreme page edges',async()=>{
 const {boundedCanvasSize}=await import('../electron/canvas-size.mjs')
 for(const [width,height,dpr] of [[811.92*4,3157.92*4,1.5],[10,100000,2],[100000,10,2],[612,792,1.5]]){
  const size=boundedCanvasSize(width,height,dpr)
  assert.ok(size.width*size.height<=24000000)
  assert.ok(size.width<=16384 && size.height<=16384)
  assert.ok(size.width>=1 && size.height>=1)
 }
 assert.deepEqual(boundedCanvasSize(612,792,1.5),{width:918,height:1188})
 assert.throws(()=>boundedCanvasSize(Infinity,100,1),/Invalid/)
})
