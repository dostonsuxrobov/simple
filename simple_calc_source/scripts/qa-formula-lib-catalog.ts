import type { Harness } from './qa-formula-library.ts'

export default async function run(h: Harness): Promise<void> {
  h.section('catalog')
  const catalog = await import('../src/lib/function-catalog.ts')
  const { FUNCTION_CATALOG, FUNCTION_CATEGORIES, functionSignature, searchFunctions, getFunctionInfo } = catalog

  // The catalog and the engine registry describe exactly the same set of functions.
  const registered = h.formulas.getFormulaFunctionNames()
  const registeredSet = new Set(registered)
  const catalogNames = Object.keys(FUNCTION_CATALOG)
  const missing = registered.filter((name) => !Object.prototype.hasOwnProperty.call(FUNCTION_CATALOG, name))
  const extra = catalogNames.filter((name) => !registeredSet.has(name))
  h.check('every registered function has a catalog entry', missing.length === 0, `missing ${missing.length}: ${missing.join(' ')}`)
  h.check('every catalog entry is a registered function', extra.length === 0, `not registered ${extra.length}: ${extra.join(' ')}`)

  // Entry shape.
  const categories = new Set<string>(FUNCTION_CATEGORIES)
  const malformed: string[] = []
  for (const [key, info] of Object.entries(FUNCTION_CATALOG)) {
    const problems: string[] = []
    if (info.name !== key) problems.push('name/key mismatch')
    if (key !== key.toUpperCase()) problems.push('key not upper-case')
    if (!info.description || !/[.]$/.test(info.description)) problems.push('description')
    if (!categories.has(info.category)) problems.push(`category ${info.category}`)
    const argNames = new Set<string>()
    for (const argument of info.args) {
      if (!argument.name || /[[\]:|]|\.\.\./.test(argument.name)) problems.push(`arg name "${argument.name}"`)
      if (!argument.description) problems.push(`arg "${argument.name}" description`)
      if (argNames.has(argument.name)) problems.push(`duplicate arg ${argument.name}`)
      argNames.add(argument.name)
    }
    if (problems.length) malformed.push(`${key} (${problems.join(', ')})`)
  }
  h.check('catalog entries are well formed', malformed.length === 0, malformed.join('; '))
  h.check('catalog is large', catalogNames.length >= 400, `${catalogNames.length} entries`)

  // Signatures.
  const signature = (name: string) => {
    const info = getFunctionInfo(name)
    return info ? functionSignature(info) : `<missing ${name}>`
  }
  const signatureCases: Array<[string, string]> = [
    ['SUMIF', 'SUMIF(range, criteria, [sum_range])'],
    ['SUM', 'SUM(number1, [number2], ...)'],
    ['PI', 'PI()'],
    ['VLOOKUP', 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])'],
    ['MAP', 'MAP(array1, [array2], ..., lambda)'],
    ['CHOOSE', 'CHOOSE(index_num, value1, [value2], ...)'],
    ['XLOOKUP', 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])'],
    ['IFS', 'IFS(logical_test1, value_if_true1, [logical_test2], [value_if_true2], ...)'],
    ['TEXTJOIN', 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)'],
    ['NORM.DIST', 'NORM.DIST(x, mean, standard_dev, cumulative)'],
    ['SLOPE', "SLOPE(known_y's, known_x's)"],
  ]
  for (const [name, expected] of signatureCases) {
    const actual = signature(name)
    h.check(`functionSignature(${name})`, actual === expected, `expected ${expected}, got ${actual}`)
  }
  const sumInfo = getFunctionInfo('sum')
  h.check('SUM repeating arg flags', Boolean(sumInfo && !sumInfo.args[0].optional && sumInfo.args[1].optional && sumInfo.args[1].repeating))
  h.check('getFunctionInfo strips prefixes', getFunctionInfo('_xlfn.xlookup')?.name === 'XLOOKUP' && getFunctionInfo('_xlfn._xlws.sort')?.name === 'SORT')
  h.check('getFunctionInfo unknown', getFunctionInfo('NOTAFUNCTION') === undefined)
  h.check('returnsArray flags', getFunctionInfo('FILTER')?.returnsArray === true && getFunctionInfo('SEQUENCE')?.returnsArray === true && !getFunctionInfo('SUM')?.returnsArray)
  h.check('categories', getFunctionInfo('VLOOKUP')?.category === 'Lookup' && getFunctionInfo('NORM.DIST')?.category === 'Statistical' && getFunctionInfo('PMT')?.category === 'Financial' && getFunctionInfo('DSUM')?.category === 'Database')

  // Search ranking.
  const names = (query: string, limit?: number) => searchFunctions(query, limit).map((info) => info.name)
  const first = (query: string) => names(query)[0]
  h.check('search "su" → SUM first', first('su') === 'SUM', names('su').join(' '))
  h.check('search "SU" is case-insensitive', first('SU') === 'SUM')
  h.check('search "vl" → VLOOKUP', first('vl') === 'VLOOKUP', names('vl').join(' '))
  h.check('search "xl" → XLOOKUP first', first('xl') === 'XLOOKUP', names('xl').join(' '))
  h.check('search "=_xlfn.xlo" → XLOOKUP', first('=_xlfn.xlo') === 'XLOOKUP', names('=_xlfn.xlo').join(' '))
  h.check('search "=su" ignores "="', first('=su') === 'SUM')
  h.check('search "dist" finds NORM.DIST', names('dist', 100).includes('NORM.DIST'), names('dist', 100).join(' '))
  h.check('search "norm.s" → NORM.S.*', names('norm.s').slice(0, 2).sort().join(' ') === 'NORM.S.DIST NORM.S.INV', names('norm.s').join(' '))
  h.check('search limit honored', names('s', 5).length === 5 && names('', 7).length === 7)
  h.check('search limit 0', names('s', 0).length === 0)
  h.check('search empty → SUM first', first('') === 'SUM')
  h.check('search no match', names('zzzz').length === 0)
  const ifResults = names('if', 200)
  const lastPrefix = ifResults.reduce((last, name, index) => (name.startsWith('IF') ? index : last), -1)
  const firstOther = ifResults.findIndex((name) => !name.startsWith('IF'))
  h.check('search "if": IF first, prefix matches before substring matches', ifResults[0] === 'IF' && ifResults.includes('COUNTIF') && lastPrefix < firstOther, ifResults.join(' '))
  h.check('search "if": IFERROR ranks before IFNA', ifResults.indexOf('IFERROR') < ifResults.indexOf('IFNA'))
  h.check('search "look" → popular lookups first', names('look').slice(0, 1)[0] === 'LOOKUP' && names('look').includes('VLOOKUP'), names('look').join(' '))
  h.check('search "norm" → modern before legacy', names('norm', 50).indexOf('NORM.DIST') < names('norm', 50).indexOf('NORMDIST'))
  h.check('search "count" → COUNTIF first', first('count') === 'COUNTIF', names('count').join(' '))
}
