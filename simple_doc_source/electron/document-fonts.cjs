const fs = require('node:fs/promises')
const path = require('node:path')
const { fontCoverage } = require('./font-coverage.cjs')

// Use fonts already licensed and installed on this computer. Never bundle or
// copy Windows font files into a release. Missing families retain engine fallbacks.
const FONT_FAMILIES = [
  ['Arial', ['arial.ttf', 'arialbd.ttf', 'ariali.ttf', 'arialbi.ttf']],
  ['Calibri', ['calibri.ttf', 'calibrib.ttf', 'calibrii.ttf', 'calibriz.ttf']],
  ['Cambria', ['cambria.ttf', 'cambriab.ttf', 'cambriai.ttf', 'cambriaz.ttf']],
  ['Times New Roman', ['times.ttf', 'timesbd.ttf', 'timesi.ttf', 'timesbi.ttf']],
  ['Courier New', ['cour.ttf', 'courbd.ttf', 'couri.ttf', 'courbi.ttf']],
  ['Georgia', ['georgia.ttf', 'georgiab.ttf', 'georgiai.ttf', 'georgiaz.ttf']],
  ['Verdana', ['verdana.ttf', 'verdanab.ttf', 'verdanai.ttf', 'verdanaz.ttf']],
  ['Tahoma', ['tahoma.ttf', 'tahomabd.ttf']],
  ['Segoe UI', ['segoeui.ttf', 'segoeuib.ttf', 'segoeuii.ttf', 'segoeuiz.ttf']],
  ['Trebuchet MS', ['trebuc.ttf', 'trebucbd.ttf', 'trebucit.ttf', 'trebucbi.ttf']],
  ['Consolas', ['consola.ttf', 'consolab.ttf', 'consolai.ttf', 'consolaz.ttf']],
  ['Garamond', ['GARA.TTF', 'GARABD.TTF', 'GARAIT.TTF']],
  ['Arial Narrow', ['ARIALN.TTF', 'ARIALNB.TTF', 'ARIALNI.TTF', 'ARIALNBI.TTF']],
  ['Aptos', ['aptos.ttf', 'aptos-bold.ttf', 'aptos-italic.ttf', 'aptos-bolditalic.ttf']],
  ['Malgun Gothic', ['malgun.ttf', 'malgunbd.ttf']],
]

function fontMetrics(value) {
  const bytes = Buffer.from(value)
  if (bytes.length < 12 || ![0x00010000, 0x4f54544f, 0x74727565].includes(bytes.readUInt32BE(0))) return null
  const count = bytes.readUInt16BE(4)
  if (count > 256 || bytes.length < 12 + count * 16) return null
  const tables = new Map()
  for (let index = 0; index < count; index += 1) {
    const position = 12 + index * 16
    const offset = bytes.readUInt32BE(position + 8)
    const length = bytes.readUInt32BE(position + 12)
    if (offset + length > bytes.length) return null
    tables.set(bytes.toString('ascii', position, position + 4), bytes.subarray(offset, offset + length))
  }
  const head = tables.get('head')
  const hhea = tables.get('hhea')
  const os2 = tables.get('OS/2')
  if (!head || head.length < 20 || !hhea || hhea.length < 10) return null
  // Restricted or bitmap-only embedding cannot be used by the vector PDF exporter.
  if (os2?.length >= 10 && (os2.readUInt16BE(8) & 0x0202)) return null
  const units = head.readUInt16BE(18)
  if (!units) return null
  const ascent = hhea.readInt16BE(4) / units
  const descent = Math.abs(hhea.readInt16BE(6)) / units
  if (ascent <= 0 || ascent > 4 || descent > 4) return null
  return { ascent, descent }
}

async function installedDocumentFonts(options = {}) {
  const roots = options.roots || [
    path.join(process.env.WINDIR || 'C:\\Windows', 'Fonts'),
    ...(process.env.LOCALAPPDATA ? [path.join(process.env.LOCALAPPDATA, 'Microsoft', 'Windows', 'Fonts')] : []),
  ]
  const files = new Map()
  const fonts = []
  for (const [family, names] of FONT_FAMILIES) {
    const faces = {}
    let sizing = null
    let simpleCoverage
    for (let index = 0; index < names.length; index += 1) {
      for (const root of roots) {
        try {
          const filePath = path.join(root, names[index])
          const stat = await fs.stat(filePath)
          if (!stat.isFile() || stat.size > 16 * 1024 * 1024) continue
          const bytes = await fs.readFile(filePath)
          const metrics = fontMetrics(bytes)
          if (!metrics) continue
          const token = String(files.size)
          files.set(token, filePath)
          faces[['regular', 'bold', 'italic', 'boldItalic'][index]] = `simple-font://installed/${token}`
          if (index === 0) {
            sizing = metrics
            if (family === 'Malgun Gothic') simpleCoverage = fontCoverage(bytes)
          }
          break
        } catch { /* Uninstalled fonts use the bundled fallback. */ }
      }
    }
    if (faces.regular && sizing) fonts.push({ family, faces, sizing, ...simpleCoverage ? { simpleCoverage } : {} })
  }
  return { fonts, files }
}

module.exports = { fontMetrics, installedDocumentFonts }
