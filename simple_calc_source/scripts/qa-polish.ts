/*
 * Excel/Sheets polish helpers: hyperlink targets (CALC-005), AutoComplete and the Alt+Down list
 * (CALC-017), error explanations (CALC-007/046), the keyboard reference (CALC-031), and the
 * formula check the editor runs before entering a typed formula (CALC-007).
 *
 *   npm run test:polish
 */
import assert from 'node:assert/strict'
import type { SheetData } from '../src/spreadsheet-types'
import {
  classifyLinkTarget,
  hyperlinkFormulaArgument,
  internalLinkTarget,
  isCellReference,
  isNetworkPath,
  isSpreadsheetPath,
  isWebLink,
  linkDescription,
  normalizeWebAddress,
  quoteSheetName,
  resolveLinkedPath,
  splitLinkLocation,
} from '../src/lib/hyperlinks'
import { acceptedAutoComplete, autoCompleteDraft, autoCompleteMatch, columnTextEntries, pickListEntries } from '../src/lib/autocomplete'
import { formulaErrorHelp, formulaErrorTooltip } from '../src/lib/error-help'
import { SHORTCUT_GROUPS, filterShortcuts } from '../src/lib/shortcuts'
import { diagnoseFormula } from '../src/lib/formulas'
import { usesExcelOnlySyntax } from '../src/lib/formula-check'
import '../src/lib/formula-library'

let groups = 0
function group(label: string, run: () => void) {
  try {
    run()
    groups += 1
    if (process.env.QA_VERBOSE) console.log(`ok - ${label}`)
  } catch (error) {
    console.error(`FAILED: ${label}`)
    throw error
  }
}

group('link targets: places in the workbook, web, email, files', () => {
  assert.deepEqual(classifyLinkTarget('#Sheet2!A1'), { kind: 'internal', location: 'Sheet2!A1' })
  assert.deepEqual(classifyLinkTarget("#'My Sheet'!B2:C4"), { kind: 'internal', location: "'My Sheet'!B2:C4" })
  assert.deepEqual(classifyLinkTarget('#TaxRate'), { kind: 'internal', location: 'TaxRate' })
  assert.deepEqual(classifyLinkTarget('#%27My%20Sheet%27!A1'), { kind: 'internal', location: "'My Sheet'!A1" })
  assert.equal(classifyLinkTarget('#').kind, 'invalid')
  assert.deepEqual(classifyLinkTarget('https://example.com/a?b=1'), { kind: 'external', url: 'https://example.com/a?b=1' })
  assert.deepEqual(classifyLinkTarget('www.example.com'), { kind: 'external', url: 'https://www.example.com' })
  assert.deepEqual(classifyLinkTarget('ann@example.com'), { kind: 'external', url: 'mailto:ann@example.com' })
  assert.deepEqual(classifyLinkTarget('mailto:ann@example.com'), { kind: 'external', url: 'mailto:ann@example.com' })
  assert.deepEqual(classifyLinkTarget('[Book1.xlsx]Sheet2!A1', 'Book1.xlsx'), { kind: 'internal', location: 'Sheet2!A1' })
  assert.deepEqual(classifyLinkTarget('[Other.xlsx]Sheet2!A1', 'Book1.xlsx'), { kind: 'file', path: 'Other.xlsx', location: 'Sheet2!A1' })
  assert.deepEqual(classifyLinkTarget('file:///C:/data/Book%202.xlsx'), { kind: 'file', path: 'C:/data/Book 2.xlsx' })
  assert.deepEqual(classifyLinkTarget('Budget.xlsx#Summary!B4'), { kind: 'file', path: 'Budget.xlsx', location: 'Summary!B4' })
  assert.equal(classifyLinkTarget('C:\\Tools\\run.exe').kind, 'file', 'a program is a file link (never opened)')
  assert.equal(classifyLinkTarget('javascript:alert(1)').kind, 'invalid')
  assert.equal(classifyLinkTarget('ftp://example.com/x').kind, 'invalid')
  assert.equal(classifyLinkTarget('').kind, 'invalid')
  // A reference without "#" is not a link Excel follows; the app tries it as a place itself.
  assert.equal(classifyLinkTarget('Sheet2!A1').kind, 'invalid')
})

group('typed web addresses', () => {
  assert.equal(normalizeWebAddress('example.com'), 'https://example.com')
  assert.equal(normalizeWebAddress('  ann@example.com '), 'mailto:ann@example.com')
  assert.equal(normalizeWebAddress('http://intranet/page'), 'http://intranet/page')
  assert.equal(normalizeWebAddress('//cdn.example.com/x'), 'https://cdn.example.com/x')
  assert.equal(normalizeWebAddress(''), null)
  assert.equal(isWebLink(normalizeWebAddress('example.com')!), true)
  assert.equal(normalizeWebAddress('C:\\secret\\file.txt'), null, 'a Windows path never becomes a web link')
  assert.equal(normalizeWebAddress('\\\\server\\share\\a.xlsx'), null)
  assert.equal(normalizeWebAddress('file:///C:/a.xlsx'), null)
  assert.equal(isWebLink('javascript:alert(1)'), false)
  assert.equal(isWebLink('file:///C:/x.xlsx'), false)
})

group('in-workbook link targets and locations', () => {
  assert.equal(internalLinkTarget('Sheet2', 'B3'), '#Sheet2!B3')
  assert.equal(internalLinkTarget('My Sheet', 'A1:C4'), "#'My Sheet'!A1:C4")
  assert.equal(internalLinkTarget("O'Neil", 'A1'), "#'O''Neil'!A1")
  assert.equal(internalLinkTarget('A1', 'B2'), "#'A1'!B2", 'a sheet named like a cell is quoted')
  assert.equal(internalLinkTarget(null, 'TaxRate'), '#TaxRate')
  assert.equal(quoteSheetName('Data_2024'), 'Data_2024')
  assert.equal(quoteSheetName('R1C1'), "'R1C1'")
  assert.deepEqual(splitLinkLocation("'My Sheet'!A1:B2"), { sheet: 'My Sheet', reference: 'A1:B2' })
  assert.deepEqual(splitLinkLocation("'It''s'!B2"), { sheet: "It's", reference: 'B2' })
  assert.deepEqual(splitLinkLocation('Sheet2!$C$3'), { sheet: 'Sheet2', reference: '$C$3' })
  assert.deepEqual(splitLinkLocation('Sheet2.A1'), { sheet: 'Sheet2', reference: 'A1' }, 'LibreOffice style')
  assert.deepEqual(splitLinkLocation('A1'), { reference: 'A1' })
  assert.deepEqual(splitLinkLocation('TaxRate'), { reference: 'TaxRate' })
  assert.equal(isCellReference('A1'), true)
  assert.equal(isCellReference('$B$2:C9'), true)
  assert.equal(isCellReference('A0'), false)
  assert.equal(isCellReference('ABCD1'), false)
  assert.equal(isCellReference('TaxRate'), false)
  assert.equal(linkDescription('#Sheet2!A1'), 'Go to Sheet2!A1')
  assert.equal(linkDescription('mailto:ann@example.com?subject=Hi'), 'Email ann@example.com')
  assert.equal(linkDescription('https://example.com'), 'https://example.com')
})

group('HYPERLINK() formulas: the location argument', () => {
  assert.equal(hyperlinkFormulaArgument('HYPERLINK("#Sheet2!A1","Go")'), '"#Sheet2!A1"')
  assert.equal(hyperlinkFormulaArgument('=HYPERLINK(A1)'), 'A1')
  assert.equal(hyperlinkFormulaArgument('HYPERLINK( "#"&ADDRESS(1,2) , "x")'), '"#"&ADDRESS(1,2)')
  assert.equal(hyperlinkFormulaArgument('_xlfn.HYPERLINK("https://example.com")'), '"https://example.com"')
  assert.equal(hyperlinkFormulaArgument('HYPERLINK("a")&"b"'), null, 'only a formula that is one HYPERLINK call is a link cell')
  assert.equal(hyperlinkFormulaArgument('SUM(1,2)'), null)
  assert.equal(hyperlinkFormulaArgument('HYPERLINK('), null)
})

group('linked files resolve next to the workbook; only spreadsheets open', () => {
  assert.equal(resolveLinkedPath('Book2.xlsx', 'C:\\data\\Book1.xlsx'), 'C:\\data\\Book2.xlsx')
  assert.equal(resolveLinkedPath('..\\other\\B.xlsx', 'C:\\data\\sub\\A.xlsx'), 'C:\\data\\other\\B.xlsx')
  assert.equal(resolveLinkedPath('D:\\x.xlsx', 'C:\\data\\A.xlsx'), 'D:\\x.xlsx')
  assert.equal(resolveLinkedPath('B.xlsx', null), null, 'an unsaved workbook has no folder')
  // Review F6: links never reach network shares (SMB connections can send the Windows sign-in).
  assert.equal(resolveLinkedPath('\\\\server\\share\\x.xlsx', 'C:\\data\\A.xlsx'), null)
  assert.equal(resolveLinkedPath('//server/share/x.xlsx', 'C:\\data\\A.xlsx'), null)
  assert.equal(resolveLinkedPath('B.xlsx', '\\\\server\\share\\A.xlsx'), null, 'nor through a relative link from a workbook on a share')
  assert.equal(resolveLinkedPath('\\\\?\\C:\\data\\B.xlsx', 'C:\\data\\A.xlsx'), '\\\\?\\C:\\data\\B.xlsx', 'a local device path is local')
  assert.equal(isNetworkPath('\\\\?\\UNC\\server\\share\\x.xlsx'), true)
  assert.equal(isNetworkPath('C:\\data\\x.xlsx'), false)
  const bracketed = classifyLinkTarget('[\\\\server\\share\\x.xlsx]Sheet1!A1', 'A.xlsx')
  assert.equal(bracketed.kind === 'file' && isNetworkPath(bracketed.path), true)
  const fileUrl = classifyLinkTarget('file://server/share/x.xlsx', 'A.xlsx')
  assert.equal(fileUrl.kind === 'file' && isNetworkPath(fileUrl.path), true, 'file://server/... is a network share')
  const localUrl = classifyLinkTarget('file:///C:/data/x.xlsx', 'A.xlsx')
  assert.deepEqual(localUrl, { kind: 'file', path: 'C:/data/x.xlsx' })
  assert.equal(isSpreadsheetPath('C:\\data\\B.xlsx'), true)
  assert.equal(isSpreadsheetPath('C:\\data\\b.ODS'), true)
  assert.equal(isSpreadsheetPath('C:\\Tools\\run.exe'), false)
  assert.equal(isSpreadsheetPath('C:\\Tools\\script.bat'), false)
})

const cells: SheetData['cells'] = {
  A1: { value: 'Fruit' },
  A2: { value: 'Apple' },
  A3: { value: 'Apricot' },
  A4: { value: 'Banana' },
  A5: { value: 42 },
  A6: { value: 'Blueberry' },
  A7: { formula: 'A2', result: 'Apple' },
  // A8 is the cell being typed in; A9 continues the block below it.
  A9: { value: 'Cherry' },
  A10: { value: 'apple' },
  // A11 blank: the block ends.
  A12: { value: 'Grape' },
  B1: { value: 'Other column' },
}

group('AutoComplete: the column block around the cell, text entries only', () => {
  const entries = columnTextEntries(cells, 7, 0)
  assert.deepEqual(entries, ['Blueberry', 'Banana', 'Apricot', 'Apple', 'Fruit', 'Cherry'], 'nearest first, numbers/formulas skipped, case-insensitive duplicates once, the block ends at a blank')
  assert.deepEqual(columnTextEntries(cells, 12, 0), ['Grape'], 'a blank cell separates blocks')
  assert.deepEqual(columnTextEntries(cells, 20, 0), [], 'nothing next to an isolated cell')
  assert.deepEqual(columnTextEntries(cells, 1, 1), ['Other column'], 'another column has its own entries')
})

group('AutoComplete: a unique, longer, case-insensitive match', () => {
  const entries = columnTextEntries(cells, 7, 0)
  assert.equal(autoCompleteMatch('Ap', entries), null, 'Apple and Apricot both match')
  assert.equal(autoCompleteMatch('App', entries), 'Apple')
  assert.equal(autoCompleteMatch('aPR', entries), 'Apricot')
  assert.equal(autoCompleteMatch('B', entries), null)
  assert.equal(autoCompleteMatch('Bl', entries), 'Blueberry')
  assert.equal(autoCompleteMatch('Banana', entries), null, 'nothing left to complete')
  assert.equal(autoCompleteMatch('Ch', entries), 'Cherry', 'entries below the cell count too')
  assert.equal(autoCompleteMatch('x', entries), null)
  assert.equal(autoCompleteMatch('', entries), null)
  assert.equal(autoCompleteMatch(' ', entries), null)
  assert.equal(autoCompleteMatch('App\nle', entries), null)
  assert.equal(autoCompleteMatch('ap', ['Apple', 'APPLE']), 'Apple', 'the same entry in two cases is one entry')
  assert.equal(autoCompleteMatch('App', ['App', 'Apple']), null, 'a complete entry that another extends is ambiguous')
  assert.equal(autoCompleteDraft('aPP', 'Apple'), 'aPPle')
  assert.equal(acceptedAutoComplete('aPPle', { typed: 'aPP', match: 'Apple' }), 'Apple', 'accepting takes the entry as written')
  assert.equal(acceptedAutoComplete('aPPx', { typed: 'aPP', match: 'Apple' }), 'aPPx', 'a changed entry is kept as typed')
  assert.equal(acceptedAutoComplete('Kiwi', null), 'Kiwi')
})

group('Alt+Down pick list: sorted distinct entries', () => {
  assert.deepEqual(pickListEntries(cells, 7, 0), ['Apple', 'Apricot', 'Banana', 'Blueberry', 'Cherry', 'Fruit'])
  assert.deepEqual(pickListEntries(cells, 30, 3), [])
  const many: SheetData['cells'] = {}
  for (let row = 1; row <= 3000; row += 1) many[`C${row}`] = { value: `Item ${row}` }
  assert.equal(pickListEntries(many, 1500, 2).length, 2000, 'the scan is capped at 1,000 cells each way')
})

group('error values explain themselves', () => {
  for (const code of ['#DIV/0!', '#N/A', '#NAME?', '#REF!', '#VALUE!', '#SPILL!', '#CIRC!', '#NUM!', '#NULL!', '#CALC!']) {
    const help = formulaErrorHelp(code)
    assert.ok(help, code)
    assert.equal(help!.code, code)
    assert.ok(help!.title.length > 3 && help!.text.length > 10, `${code} has a title and an explanation`)
  }
  assert.equal(formulaErrorHelp('#DIV/0!')!.title, 'Divide by zero')
  assert.match(formulaErrorHelp('#DIV/0!')!.text, /divides by zero/)
  assert.equal(formulaErrorHelp('#NAME?', 'SUMM isn’t a function this app knows.')!.text, 'SUMM isn’t a function this app knows.')
  assert.equal(formulaErrorHelp('hello'), null)
  assert.equal(formulaErrorHelp(5), null)
  assert.equal(formulaErrorHelp(null), null)
  assert.equal(formulaErrorTooltip(formulaErrorHelp('#N/A')!), '#N/A – Value not available. A value isn\'t available: usually a lookup found no match.')
})

group('the formula check before entering a typed formula', () => {
  const missing = diagnoseFormula('=SUM(A1:A3')
  assert.ok(missing && missing.suggestion === '=SUM(A1:A3)', 'a missing ")" is offered as a correction')
  const quote = diagnoseFormula('="abc')
  assert.ok(quote && quote.suggestion === '="abc"', 'a missing quote too')
  const broken = diagnoseFormula('=1+*2')
  assert.ok(broken && broken.kind === 'syntax' && !broken.suggestion, 'a real syntax problem has no correction')
  assert.ok(broken!.position >= 1 && broken!.position < 5, 'and points into the formula')
  const unknown = diagnoseFormula('=SUMM(1)')
  assert.ok(unknown && unknown.kind === 'unknown-function', 'an unknown function is reported but not a syntax problem (it enters as #NAME?)')
  assert.equal(diagnoseFormula('=SUM(A1:A3)'), null)
  // Valid Excel this app cannot calculate is entered as written, never refused.
  for (const formula of ['=A1:C3 B2:B9', '=Sales Q1', '=SUM((A1:A2,C1:C2))', '=(A1:A2,B1:B2)', '=[Budget.xlsx]Sheet1!A1', "='C:\\data\\[B.xlsx]Sheet 1'!A1", '=SUM([Book2.xlsx]Data!A1:A9)']) {
    assert.equal(usesExcelOnlySyntax(formula), true, formula)
  }
  for (const formula of ['=1+*2', '=SUM(A1 )', '=IF((A1>0),1,2)', '=(1,2)', '=Table1[Qty]+Sheet2!A1', '="[x]y!"&A1', '=A1', 'text']) {
    assert.equal(usesExcelOnlySyntax(formula), false, formula)
  }
})

group('keyboard reference', () => {
  const all = SHORTCUT_GROUPS.flatMap((entry) => entry.items.map((item) => item.keys))
  for (const keys of ['Ctrl+/', 'Alt+/', 'F9', 'Shift+F9', 'Ctrl+Alt+F9', 'Alt+Down', 'Ctrl+K', 'Ctrl+Mouse wheel']) assert.ok(all.includes(keys), `${keys} is listed`)
  for (const entry of SHORTCUT_GROUPS) assert.ok(entry.items.length > 0, entry.title)
  const zoom = filterShortcuts(SHORTCUT_GROUPS, 'zoom')
  assert.equal(zoom.length, 1)
  assert.equal(zoom[0].items[0].keys, 'Ctrl+Mouse wheel')
  assert.equal(filterShortcuts(SHORTCUT_GROUPS, 'no such command').length, 0)
  assert.equal(filterShortcuts(SHORTCUT_GROUPS, '').length, SHORTCUT_GROUPS.length)
})

console.log(`Polish QA passed: ${groups} groups (links, AutoComplete, pick list, error help, formula check, shortcuts).`)
