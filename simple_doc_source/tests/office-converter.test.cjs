const test = require('node:test')
const assert = require('node:assert/strict')
const fs = require('node:fs/promises')
const path = require('node:path')
const os = require('node:os')
const { fileURLToPath } = require('node:url')
const { convertOfficeBytes, readDiskConversion, storeDiskConversion } = require('../electron/office-converter.cjs')

test('Office conversion isolates source/profile and returns only a completed output', async () => {
  let folder
  const result = await convertOfficeBytes({ bytes: Buffer.from('original'), inputExtension: 'xls', outputExtension: 'xlsx', filter: 'Calc MS Excel 2007 XML' }, {
    executable: 'fixture-office',
    run: async (executable, args) => {
      assert.equal(executable, 'fixture-office')
      assert.ok(args.includes('--headless'))
      assert.ok(args.includes('xlsx:Calc MS Excel 2007 XML'))
      const inputPath = args.at(-1)
      folder = path.dirname(inputPath)
      const profilePath = fileURLToPath(args[0].replace('-env:UserInstallation=', ''))
      assert.equal(await fs.readFile(inputPath, 'utf8'), 'original')
      assert.match(await fs.readFile(path.join(profilePath, 'user', 'registrymodifications.xcu'), 'utf8'), /MacroSecurityLevel.*<value>3<\/value>/)
      const outputPath = args[args.indexOf('--outdir') + 1]
      await fs.writeFile(path.join(outputPath, 'document.xlsx'), 'converted')
    },
  })
  assert.equal(result.toString(), 'converted')
  await assert.rejects(fs.access(folder))
})

test('failed conversions remove temporary input and never claim a result', async () => {
  let folder
  await assert.rejects(convertOfficeBytes({ bytes: Buffer.from('original'), inputExtension: 'doc', outputExtension: 'docx' }, {
    executable: 'fixture-office',
    run: async (_executable, args) => { folder = path.dirname(args.at(-1)); throw new Error('conversion failed') },
  }), /conversion failed/)
  await assert.rejects(fs.access(folder))
})

test('conversion rejects unsafe extensions before launching the engine', async () => {
  await assert.rejects(convertOfficeBytes({ bytes: Buffer.from('original'), inputExtension: '../doc', outputExtension: 'docx' }, { executable: 'unused' }), /Unsupported/)
})

test('disk conversion cache checks content hash and signatures before reusing a result', async () => {
  const directory = await fs.mkdtemp(path.join(os.tmpdir(), 'simple-office-cache-test-'))
  const previous = process.env.LOCALAPPDATA
  process.env.LOCALAPPDATA = directory
  try {
    const bytes = Buffer.from('%PDF-1.7\nfixture content\n%%EOF')
    await storeDiskConversion('fixture-key', 'pdf', bytes)
    assert.deepEqual(await readDiskConversion('fixture-key', 'pdf'), bytes)
    assert.equal(await readDiskConversion('other-key', 'pdf'), null)
    const cacheDirectory = path.join(directory, 'simple', 'conversion-cache', 'v1')
    const file = (await fs.readdir(cacheDirectory)).find(name => name.endsWith('.pdf'))
    const corrupt = Buffer.from(bytes)
    corrupt[12] ^= 1
    await fs.writeFile(path.join(cacheDirectory, file), corrupt)
    assert.equal(await readDiskConversion('fixture-key', 'pdf'), null, 'same-size corruption must fail the hash check')
  } finally {
    if (previous === undefined) delete process.env.LOCALAPPDATA
    else process.env.LOCALAPPDATA = previous
    await fs.rm(directory, { recursive: true, force: true })
  }
})
