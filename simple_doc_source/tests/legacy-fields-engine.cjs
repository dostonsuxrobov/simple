'use strict'
// Explicit integration check; omitted from the fast *.test.cjs suite because
// it starts the installed office engine. Input documents are read-only.
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const path = require('node:path')
const crypto = require('node:crypto')
const JSZip = require('jszip')
const { lockLegacyDateFields } = require('../electron/legacy-fields.cjs')
const { convertOfficeBytes, runOfficeConverter } = require('../electron/office-converter.cjs')

async function main() {
  const directory = path.resolve(__dirname, '../../.codex-tmp/reliability-reference')
  const source = path.join(directory, 'test_doc.doc')
  const original = await fs.readFile(source)
  const before = crypto.createHash('sha256').update(original).digest('hex')
  const prepared = lockLegacyDateFields(original)
  assert.equal(prepared.lockedFields, 1)
  assert.deepEqual(original, await fs.readFile(source))
  const pdfjs = await import('../../simple_pdf_source/node_modules/pdfjs-dist/legacy/build/pdf.mjs')
  async function textOf(bytes) {
    const pdf = await pdfjs.getDocument({ data: new Uint8Array(bytes), isEvalSupported: false }).promise
    try {
      const pages = []
      for(let page=1;page<=pdf.numPages;page++) pages.push((await(await pdf.getPage(page)).getTextContent()).items.map(item=>item.str||'').join(' '))
      return { pages:pdf.numPages,text:pages.join(' ').replace(/\s+/g,' ').trim() }
    } finally { await pdf.destroy() }
  }
  const google = await textOf(await fs.readFile(path.join(directory, 'test_doc-google-reference.pdf')))
  const expectedDate = google.text.match(/DATE:\s*([A-Z]+\s+\d{1,2},\s+\d{4})/)[1]
  const started = performance.now()
  const direct = await convertOfficeBytes({ bytes:prepared.bytes,inputExtension:'doc',outputExtension:'pdf',filter:'writer_pdf_Export' },{run:runOfficeConverter})
  const directText = await textOf(direct)
  assert.equal(directText.text.match(/DATE:\s*([A-Z]+\s+\d{1,2},\s+\d{4})/)[1],expectedDate)
  assert.equal(directText.pages,1)
  const docx = await convertOfficeBytes({ bytes:prepared.bytes,inputExtension:'doc',outputExtension:'docx' },{run:runOfficeConverter})
  const zip = await JSZip.loadAsync(docx)
  const xml = await zip.file('word/document.xml').async('string')
  // Word may store title case with an all-caps run property, and may split the
  // cached value across runs. The rendered round trip below verifies appearance.
  assert.ok(xml.replace(/<[^>]*>/g,'').toUpperCase().includes(expectedDate),'DOCX editor import must retain the cached date text')
  const reopened = await convertOfficeBytes({ bytes:docx,inputExtension:'docx',outputExtension:'pdf',filter:'writer_pdf_Export' },{run:runOfficeConverter})
  const reopenedText = await textOf(reopened)
  assert.equal(reopenedText.text.match(/DATE:\s*([A-Z]+\s+\d{1,2},\s+\d{4})/)[1],expectedDate,'a DOCX conversion round trip must not refresh the date')
  assert.equal(reopenedText.pages,1)
  assert.equal(crypto.createHash('sha256').update(await fs.readFile(source)).digest('hex'),before)
  const result = { passed:true,lockedFields:1,cachedDate:expectedDate,directPdfPages:1,docxRoundTripPdfPages:1,inputHashUnchanged:true,milliseconds:performance.now()-started }
  await fs.writeFile(path.join(directory,'legacy-field-result.json'),JSON.stringify(result,null,2))
  console.log(JSON.stringify(result))
}
main().catch(error=>{console.error(error);process.exitCode=1})
