// QA for the extended function library (src/lib/formula-library.ts) and the function catalog.
// Run: node --no-warnings --experimental-transform-types scripts/qa-formula-library.ts
//
// The src modules use extensionless relative imports (Vite/TS "Bundler" resolution). Node's
// ESM loader needs explicit extensions, so a synchronous resolve hook retries unresolved
// relative specifiers with ".ts" before the library is loaded dynamically. Domain checks live
// in scripts/qa-formula-lib-<part>.ts; each default-exports `(h: Harness) => void`.
import { registerHooks } from 'node:module'

registerHooks({
  resolve(specifier, context, nextResolve) {
    try {
      return nextResolve(specifier, context)
    } catch (error) {
      if (/^\.{1,2}\//.test(specifier) && !/\.[cm]?[jt]sx?$/.test(specifier)) {
        return nextResolve(`${specifier}.ts`, context)
      }
      throw error
    }
  },
})

const formulas = await import('../src/lib/formulas.ts')
const coreNames = new Set(formulas.getFormulaFunctionNames())
const library = await import('../src/lib/formula-library.ts')

type FormulaResult = number | string | boolean
export type Cells = Record<string, unknown>

export interface EvaluateOptions {
  row?: number
  column?: number
  hooks?: Record<string, unknown>
}

export interface Harness {
  /** Shared sample sheet (see DEFAULT_CELLS below). */
  cells: Cells
  /** Evaluate to the anchor (top-left) value. */
  value(formula: string, cells?: Cells, options?: EvaluateOptions): FormulaResult
  /** Evaluate to rows of values (a scalar result becomes [[value]]). */
  rows(formula: string, cells?: Cells, options?: EvaluateOptions): FormulaResult[][]
  /** Strict equality on the anchor value (use for text, booleans, error codes, exact numbers). */
  eq(formula: string, expected: FormulaResult, cells?: Cells, options?: EvaluateOptions): void
  /** Numeric closeness: |actual-expected| <= tolerance * max(1, |expected|). Default 1e-9. */
  near(formula: string, expected: number, tolerance?: number, cells?: Cells, options?: EvaluateOptions): void
  /** Exact array result (numbers compared with a 1e-12 relative tolerance). */
  rowsEq(formula: string, expected: FormulaResult[][], cells?: Cells, options?: EvaluateOptions): void
  /** Numeric array result within tolerance (non-numbers compared strictly). */
  rowsNear(formula: string, expected: FormulaResult[][], tolerance?: number, cells?: Cells, options?: EvaluateOptions): void
  /** Free-form assertion. */
  check(label: string, condition: boolean, detail?: string): void
  /** Label subsequent failures. */
  section(name: string): void
  /** The formulas.ts module (evaluateFormula, getFormulaFunctionNames, ...). */
  formulas: typeof formulas
}

/**
 * Shared sample sheet:
 *   A1:A10 = 1..10
 *   B1:B8  = 2,4,4,4,5,5,7,9            (population stdev 2, mean 5)
 *   C1:C5  = "text", TRUE, 3, (blank), "4"
 *   D1:D7  = 2,3,9,1,8,7,5              (known_y's from Excel's SLOPE/INTERCEPT docs)
 *   E1:E7  = 6,5,11,7,5,4,4             (known_x's)
 *   F1:F4  = "apple","Banana","cherry","apple pie"
 *   G1:H3  = {1,2;3,4;5,6}
 */
const DEFAULT_CELLS: Cells = {
  A1: 1, A2: 2, A3: 3, A4: 4, A5: 5, A6: 6, A7: 7, A8: 8, A9: 9, A10: 10,
  B1: 2, B2: 4, B3: 4, B4: 4, B5: 5, B6: 5, B7: 7, B8: 9,
  C1: 'text', C2: true, C3: 3, C5: '4',
  D1: 2, D2: 3, D3: 9, D4: 1, D5: 8, D6: 7, D7: 5,
  E1: 6, E2: 5, E3: 11, E4: 7, E5: 5, E6: 4, E7: 4,
  F1: 'apple', F2: 'Banana', F3: 'cherry', F4: 'apple pie',
  G1: 1, H1: 2, G2: 3, H2: 4, G3: 5, H3: 6,
}

function usedRange(cells: Cells) {
  let maxRow = 1
  let maxCol = 1
  for (const address of Object.keys(cells)) {
    const parsed = formulas.parseA1Address(address)
    if (!parsed) continue
    maxRow = Math.max(maxRow, parsed.row)
    maxCol = Math.max(maxCol, parsed.column)
  }
  return { maxRow, maxCol }
}

function hooksFor(cells: Cells, options: EvaluateOptions = {}) {
  return {
    currentCell: { row: options.row ?? 20, column: options.column ?? 26 },
    getUsedRange: () => usedRange(cells),
    getSheetNames: () => ['Sheet1', 'Data'],
    getSheetName: (id: string) => id,
    getCellFormula: (_sheet: string, address: string) => {
      const value = cells[address]
      return typeof value === 'string' && value.startsWith('=') ? value.slice(1) : null
    },
    ...(options.hooks ?? {}),
  }
}

let failures = 0
let passes = 0
let currentSection = ''
const failureLines: string[] = []

function fail(label: string, detail: string) {
  failures += 1
  if (failureLines.length < 400) failureLines.push(`[${currentSection}] ${label}: ${detail}`)
}

function show(value: unknown): string {
  return JSON.stringify(value)
}

function closeEnough(actual: unknown, expected: unknown, tolerance: number): boolean {
  if (typeof expected === 'number' && typeof actual === 'number') {
    if (Number.isNaN(expected)) return Number.isNaN(actual)
    return Math.abs(actual - expected) <= tolerance * Math.max(1, Math.abs(expected))
  }
  return actual === expected
}

const h: Harness = {
  cells: DEFAULT_CELLS,
  formulas,
  value(formula, cells = DEFAULT_CELLS, options) {
    const resolver = (_sheet: string, address: string) => cells[address] as never
    return formulas.evaluateFormula(formula, 'Sheet1', resolver, hooksFor(cells, options) as never)
  },
  rows(formula, cells = DEFAULT_CELLS, options) {
    const resolver = (_sheet: string, address: string) => cells[address] as never
    const result = formulas.evaluateFormulaDetailed(formula, 'Sheet1', resolver, hooksFor(cells, options) as never)
    if (!result.array) return [[result.value]]
    const rows: FormulaResult[][] = []
    for (let row = 0; row < result.array.rowCount; row += 1) {
      rows.push(result.array.values.slice(row * result.array.columnCount, (row + 1) * result.array.columnCount))
    }
    return rows
  },
  eq(formula, expected, cells, options) {
    let actual: unknown
    try {
      actual = h.value(formula, cells, options)
    } catch (error) {
      fail(formula, `threw ${(error as Error).stack}`)
      return
    }
    if (typeof expected === 'number' && typeof actual === 'number' ? closeEnough(actual, expected, 1e-12) : actual === expected) passes += 1
    else fail(formula, `expected ${show(expected)}, got ${show(actual)}`)
  },
  near(formula, expected, tolerance = 1e-9, cells, options) {
    let actual: unknown
    try {
      actual = h.value(formula, cells, options)
    } catch (error) {
      fail(formula, `threw ${(error as Error).stack}`)
      return
    }
    if (typeof actual === 'number' && closeEnough(actual, expected, tolerance)) passes += 1
    else fail(formula, `expected ≈${show(expected)} (±${tolerance}), got ${show(actual)}`)
  },
  rowsEq(formula, expected, cells, options) {
    h.rowsNear(formula, expected, 1e-12, cells, options)
  },
  rowsNear(formula, expected, tolerance = 1e-9, cells, options) {
    let actual: FormulaResult[][]
    try {
      actual = h.rows(formula, cells, options)
    } catch (error) {
      fail(formula, `threw ${(error as Error).stack}`)
      return
    }
    const sameShape = actual.length === expected.length && actual.every((row, index) => row.length === expected[index].length)
    if (sameShape && actual.every((row, r) => row.every((value, c) => closeEnough(value, expected[r][c], tolerance)))) passes += 1
    else fail(formula, `expected ${show(expected)}, got ${show(actual)}`)
  },
  check(label, condition, detail = '') {
    if (condition) passes += 1
    else fail(label, detail || 'check failed')
  },
  section(name) {
    currentSection = name
  },
}

const PARTS = ['array', 'text', 'info', 'math', 'engineering', 'stats', 'date', 'financial', 'catalog']
for (const part of PARTS) {
  h.section(part)
  let run: (harness: Harness) => void | Promise<void>
  try {
    run = (await import(`./qa-formula-lib-${part}.ts`)).default
  } catch (error) {
    fail(`load qa-formula-lib-${part}.ts`, (error as Error).stack ?? String(error))
    continue
  }
  try {
    await run(h)
  } catch (error) {
    fail(`run ${part}`, (error as Error).stack ?? String(error))
  }
}

// ---- Library integrity ---------------------------------------------------------------------
h.section('integrity')
{
  const owners = new Map<string, string[]>()
  for (const [module, specs] of Object.entries(library.LIBRARY_MODULES)) {
    for (const name of Object.keys(specs)) owners.set(name.toUpperCase(), [...(owners.get(name.toUpperCase()) ?? []), module])
  }
  const duplicates = [...owners].filter(([, modules]) => modules.length > 1).map(([name, modules]) => `${name} (${modules.join(', ')})`)
  h.check('no function is defined by two library modules', duplicates.length === 0, duplicates.join('; '))
  const overrides = new Set(['TEXTBEFORE', 'TEXTAFTER'])
  const shadowed = [...owners.keys()].filter((name) => coreNames.has(name) && !overrides.has(name))
  h.check('library does not silently re-define core functions', shadowed.length === 0, shadowed.join(' '))
  const registered = new Set(formulas.getFormulaFunctionNames())
  const unregistered = [...owners.keys()].filter((name) => !registered.has(name))
  h.check('every library function is registered', unregistered.length === 0, unregistered.join(' '))
  h.check('library adds several hundred functions', library.LIBRARY_FUNCTION_NAMES.length >= 300, `${library.LIBRARY_FUNCTION_NAMES.length}`)
  const registry = formulas.FUNCTION_REGISTRY
  const badArity = Object.entries(registry).filter(([, spec]) => !(spec.minArgs >= 0 && spec.minArgs <= spec.maxArgs)).map(([name]) => name)
  h.check('every spec has a valid arity', badArity.length === 0, badArity.join(' '))
  for (const name of ['RANDARRAY', 'CELL', 'INFO', 'RAND', 'NOW', 'TODAY', 'OFFSET', 'INDIRECT']) {
    h.check(`${name} is volatile`, registry[name]?.volatile === true)
  }
  const catalog = await import('../src/lib/function-catalog.ts')
  const arrayMismatch = Object.entries(registry)
    .filter(([name, spec]) => spec.returnsArray && !['LAMBDA', 'INDEX', 'ANCHORARRAY', 'TRIMRANGE'].includes(name) && !catalog.FUNCTION_CATALOG[name]?.returnsArray)
    .map(([name]) => name)
  h.check('array-returning functions are flagged in the catalog', arrayMismatch.length === 0, arrayMismatch.join(' '))
}

if (failures > 0) {
  console.error(failureLines.join('\n'))
  console.error(`\nFormula library QA: ${failures} failed, ${passes} passed.`)
  process.exit(1)
}
console.log(`Formula library QA passed: ${passes} checks across ${formulas.getFormulaFunctionNames().length} registered functions.`)
