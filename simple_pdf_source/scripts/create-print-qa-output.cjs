'use strict'

const fs = require('node:fs/promises')
const path = require('node:path')
const { preparePrintPdf } = require('../electron/pdf-print.cjs')

async function main() {
  const root = path.resolve(__dirname, '..')
  const source = path.join(root, 'tmp', 'pdfs', 'simple-interaction-fixture.pdf')
  const target = path.join(root, 'tmp', 'pdfs', 'simple-print-qa-letter.pdf')
  const output = await preparePrintPdf(await fs.readFile(source), {
    paperSize: 'Letter',
    landscape: false,
    marginMode: 'normal',
    scaleMode: 'fit',
    customScale: 1,
  })
  await fs.writeFile(target, output)
  process.stdout.write(`${target}\n`)
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
})
