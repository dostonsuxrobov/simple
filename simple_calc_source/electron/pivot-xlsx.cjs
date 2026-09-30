'use strict'

// Pivot table definitions travel with the workbook in a custom XML data part (the standard
// OOXML home for application data, which Excel keeps on save). The pivot's output is plain
// cells, so other spreadsheet apps show the last refreshed summary.

const crypto = require('node:crypto')

const ROOT = 'simpleCalcPivots'
const NAMESPACE = 'urn:simple-calc:pivots:v1'
const REL_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
const REL_CUSTOM_XML_PROPS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps'
const CONTENT_TYPE_PROPS = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml'
const MAX_DEFINITION_BYTES = 8 * 1024 * 1024

function escapeAttribute(value) {
  return String(value).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;')
}

async function ownItemPart(zip) {
  for (const name of Object.keys(zip.files)) {
    if (!/^customXml\/item\d+\.xml$/i.test(name)) continue
    const text = await zip.file(name).async('string')
    if (text.includes(`<${ROOT}`)) return { name, text }
  }
  return null
}

/** Read saved pivot definitions onto their sheets (matched by sheet name, then position). */
async function importPivotDefinitions(zipOrBuffer, sheets, warnings) {
  const JSZip = require('jszip')
  const zip = zipOrBuffer && typeof zipOrBuffer.file === 'function' ? zipOrBuffer : await JSZip.loadAsync(zipOrBuffer)
  const part = await ownItemPart(zip)
  if (!part) return
  const match = /<!\[CDATA\[([\s\S]*)\]\]>/.exec(part.text)
  if (!match) return
  let data
  try {
    data = JSON.parse(match[1].replace(/\]\]\]\]><!\[CDATA\[>/g, ']]>'))
  } catch {
    warnings.push('Pivot table definitions in this workbook could not be read; their last results are shown as plain cells.')
    return
  }
  for (const entry of Array.isArray(data && data.sheets) ? data.sheets : []) {
    const sheet = sheets.find((item) => item.name === entry.sheetName) || sheets[Number(entry.sheetIndex)]
    if (!sheet || !Array.isArray(entry.pivots)) continue
    sheet.pivots = entry.pivots.filter((pivot) => pivot && typeof pivot.id === 'string' && typeof pivot.source === 'string' && pivot.anchor)
  }
}

/** Write (or remove) the pivot definition part in a package built by ExcelJS. */
async function writePivotPackage(zip, model) {
  const sheets = (model.sheets || [])
    .map((sheet, index) => ({ sheetName: sheet.name, sheetIndex: index, pivots: Array.isArray(sheet.pivots) ? sheet.pivots : [] }))
    .filter((entry) => entry.pivots.length)
  const existing = await ownItemPart(zip)
  const relsName = 'xl/_rels/workbook.xml.rels'
  let rels = zip.file(relsName) ? await zip.file(relsName).async('string') : null
  let types = await zip.file('[Content_Types].xml').async('string')
  if (!sheets.length) {
    if (!existing) return
    // No pivots any more: drop the part and everything that points at it.
    const number = /item(\d+)\.xml$/i.exec(existing.name)[1]
    zip.remove(existing.name)
    zip.remove(`customXml/itemProps${number}.xml`)
    zip.remove(`customXml/_rels/item${number}.xml.rels`)
    if (rels) zip.file(relsName, rels.replace(new RegExp(`<Relationship\\b[^>]*Target="\\.\\./customXml/item${number}\\.xml"[^>]*/>`, 'g'), ''))
    zip.file('[Content_Types].xml', types.replace(new RegExp(`<Override\\b[^>]*PartName="/customXml/itemProps${number}\\.xml"[^>]*/>`, 'g'), ''))
    return
  }
  const json = JSON.stringify({ version: 1, sheets })
  if (json.length > MAX_DEFINITION_BYTES) return
  const payload = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${ROOT} xmlns="${NAMESPACE}"><![CDATA[${json.replace(/\]\]>/g, ']]]]><![CDATA[>')}]]></${ROOT}>`
  let number
  if (existing) number = /item(\d+)\.xml$/i.exec(existing.name)[1]
  else {
    number = 1
    while (zip.file(`customXml/item${number}.xml`) || zip.file(`customXml/itemProps${number}.xml`)) number += 1
  }
  zip.file(`customXml/item${number}.xml`, payload)
  if (!zip.file(`customXml/itemProps${number}.xml`)) {
    const id = `{${crypto.randomUUID().toUpperCase()}}`
    zip.file(`customXml/itemProps${number}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="${escapeAttribute(id)}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${NAMESPACE}"/></ds:schemaRefs></ds:datastoreItem>`)
  }
  if (!zip.file(`customXml/_rels/item${number}.xml.rels`)) {
    zip.file(`customXml/_rels/item${number}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL_CUSTOM_XML_PROPS}" Target="itemProps${number}.xml"/></Relationships>`)
  }
  if (!new RegExp(`PartName="/customXml/itemProps${number}\\.xml"`).test(types)) {
    types = types.replace('</Types>', `<Override PartName="/customXml/itemProps${number}.xml" ContentType="${CONTENT_TYPE_PROPS}"/></Types>`)
  }
  if (!/<Default\b[^>]*Extension="xml"/i.test(types)) types = types.replace('</Types>', '<Default Extension="xml" ContentType="application/xml"/></Types>')
  zip.file('[Content_Types].xml', types)
  if (rels && !new RegExp(`Target="\\.\\./customXml/item${number}\\.xml"`).test(rels)) {
    const ids = [...rels.matchAll(/\bId="rId(\d+)"/g)].map((item) => Number(item[1]))
    const id = `rId${Math.max(0, ...ids) + 1}`
    rels = rels.replace('</Relationships>', `<Relationship Id="${id}" Type="${REL_CUSTOM_XML}" Target="../customXml/item${number}.xml"/></Relationships>`)
    zip.file(relsName, rels)
  }
}

module.exports = { importPivotDefinitions, writePivotPackage }
