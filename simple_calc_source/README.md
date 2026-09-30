# simple_calc

`simple_calc` is a minimal, local-first spreadsheet editor for Windows. It opens modern and legacy workbook formats, keeps formulas as formulas, and keeps the original supported file format on Save. New workbooks use `.xlsx`. For modern Excel packages, saving is source-backed so workbook structures supported by the rich reader are retained instead of rebuilding every edited file from a blank workbook.

## What works

- Sparse, buffered virtualized grid for large worksheets, with indexed row/column geometry and responsive scrolling even far into imported sheets
- Cell values, formulas, cached results, broad live formula calculation, and cross-sheet references
- Multiple sheets, rename/duplicate/delete, merged cells, imported widths/heights, hidden rows/columns and sheets, and Google Sheets-style frozen panes: drag the two handles in the grid corner or drag a bold divider back to unfreeze; native sticky layers keep headers and frozen cells steady during fast scrolling
- Google Sheets-style View, Insert, Format, Data, and Tools menus with searchable commands and a compact formatting toolbar
- Formula-aware row and column insertion/deletion that shifts cells, formulas, ranges, merges, dimensions, validations, panes, and related sheet metadata together
- Excel-compatible row/column sizing: drag a header edge to resize or double-click it to auto-fit displayed content; dimensions are undoable and retained in `.xlsx`
- Fonts, theme/indexed/tinted colors, pattern and gradient fills, border presets, horizontal/vertical alignment, overflow/wrap/clip controls, indentation, rotation, merged cells, decimals, and common number formats
- Excel-compatible checkboxes and dropdown validation, safe web/email links, plain cell notes, alternating colors, gridline visibility, sheet hiding, zoom, and full screen
- Formula-aware internal copy/paste and fill with relative/absolute A1 references
- Pattern-aware fill handle for number/date/text series, repeated data, and shifted formulas
- Live `Ctrl+F` search with next/previous navigation and a selectable bottom-right Sum/Average/Minimum/Maximum/Count summary
- External TSV clipboard exchange, drag-and-drop opening, recent files, undo/redo, and zoom
- Two-pane live print preview for the active sheet, selected cells, or all visible sheets, with exact paper pages, automatic page breaks, paper, orientation, margin, scaling, gridline, and heading controls
- A top-level Export As dialog for independent XLSX, XLS, ODS, CSV, TSV, PDF, and standalone HTML copies; exporting never replaces the open workbook or clears its unsaved state
- Atomic local saves through an isolated Electron bridge

### Spreadsheet features

- **Calculation engine** — 532 Excel and Google Sheets functions, dynamic arrays that spill (`A1#`, `@`, `#SPILL!`), `LET`/`LAMBDA`/`MAP`/`REDUCE`/`SCAN`/`BYROW`/`BYCOL`, structured table references, and incremental recalculation that only re-evaluates what an edit affects
- **Formula editing** — colour-coded range finder, point-and-click references, function autocomplete with argument hints, `F4` absolute/relative cycling, and Excel-style entry parsing (currency, percentages, fractions, dates, times) with inferred number formats
- **Tables** (`Ctrl+T`) — 60 built-in styles, header/total/banded/first/last-column options, totals-row functions, calculated columns, auto-expansion when typing next to a table, header renames that rewrite `Table[Column]` references, per-table filter buttons, resize, and Convert to Range (which keeps the look and turns references into A1 references)
- **Pivot tables** — rows, columns, values and filters; Sum, Count, Average, Max, Min, Product, Count Numbers, Distinct Count, Median, StdDev(p) and Var(p); show values as % of grand/row/column total or running total; date grouping by year, quarter, month or day; subtotals and grand totals; autofit; `Alt+F5` / `Ctrl+Alt+F5` refresh. Definitions are saved with the workbook, and other apps see the last refreshed values
- **Charts** — column, bar, line, area, pie, doughnut, scatter, radar and combo charts with a docked editor; imported DrawingML charts are kept byte-for-byte until edited
- **Pictures** — imported pictures are shown; insert from a file or paste a screenshot, then move, resize (aspect-locked from corners), set alt text, print and save
- **Sparklines** — Excel sparkline groups (line, column, win/loss with markers, high/low/first/last/negative points, shared or custom axes) are shown, created from Insert › Sparklines, edited from the cell menu and saved back (unedited groups byte-for-byte); Google Sheets' `SPARKLINE()` function draws in-cell charts too
- **Sheet protection** — imported protected sheets are enforced (locked cells, formatting, row/column, object, filter and pivot permissions); Protect Sheet with an Excel-compatible SHA-512 password, Unprotect Sheet, and Format › Lock cells
- **Conditional formatting** — cell rules, top/bottom, averages, duplicates, text, dates, formulas, colour scales, data bars and icon sets, with rule priorities
- **Data tools** — AutoFilter and table filters (values, conditions, colours), multi-level sort (values, colours, custom lists), data validation (lists, numbers, dates, text length and custom formulas checked against the value being entered), circle invalid data, remove duplicates, text to columns, and cleanup commands
- **Structure** — insert/delete rows and columns, insert/delete cells with shift (`Ctrl+Shift+=` / `Ctrl+-`), group and outline rows or columns with collapse/expand buttons and level buttons (`Alt+Shift+→` / `Alt+Shift+←`), and Freeze Panes at the active cell
- **Analysis** — Goal Seek, formula auditing (trace precedents/dependents arrows, one level per press, red for error paths), named ranges manager, status-bar statistics
- **Formatting** — Format Cells dialog with custom number format codes (sections, colours, conditions, fractions, elapsed time, accounting), cell styles gallery, format painter, borders, merge variants, and Excel's `####` for numbers that do not fit

## Formula compatibility

The live engine registers 532 functions across every Excel category and the common Google Sheets additions:

- Lookup and dynamic arrays: `XLOOKUP`, `XMATCH`, `FILTER`, `SORT`, `SORTBY`, `UNIQUE`, `SEQUENCE`, `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `TOCOL`, `TOROW`, `WRAPROWS`, `EXPAND`, `GROUPBY`, `PIVOTBY`, `INDEX`, `MATCH`, `VLOOKUP`, `INDIRECT`, `OFFSET`
- Lambdas: `LET`, `LAMBDA`, `MAP`, `REDUCE`, `SCAN`, `BYROW`, `BYCOL`, `MAKEARRAY`, `ISOMITTED`
- Text: `TEXT`, `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTJOIN`, `REGEXTEST`, `REGEXEXTRACT`, `REGEXREPLACE`, `SUBSTITUTE`, `CONCAT` and the rest of the text family
- Statistics, math, dates, financial (`PMT`, `NPV`, `IRR`, `XIRR`, `XNPV`), engineering (`BIN2DEC`, `CONVERT`, `BESSELJ`), database (`DSUM` …), information (`ISFORMULA`, `CELL`) and `SUBTOTAL`/`AGGREGATE` that respect hidden and filtered rows

Operators broadcast over a range, so the mask idioms real workbooks are built on — `SUMPRODUCT((A2:A99="North")*(B2:B99))`, `SUM(B2:B99*C2:C99)` — calculate locally, as do whole-column references, named ranges, and omitted arguments (`IF(A1>5,1,)`). Formulas saved without a cached value are recalculated on open rather than left blank, which is how workbooks written by Google Sheets and other non-Excel generators arrive.

Structured table references (`Table1[Amount]`, `[@Units]`, `[#Totals]`) and dynamic-array spilling calculate locally, and Excel's `_xlfn.` / `_xlws.` file prefixes are handled on open and save. Imported formulas outside the live engine (for example web functions such as `GOOGLEFINANCE`) keep their formula text and cached result.

## File compatibility

The rich OOXML importer handles `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, and `.xlam`, with a compatibility-reader fallback. The compatibility importer also accepts `.xlsb`, `.xls`, `.xml`, `.ods`, `.fods`, `.numbers`, `.csv`, `.tsv`, `.txt`, `.slk`, `.sylk`, `.dif`, `.dbf`, `.prn`, and common Lotus / Works workbook extensions.

Save and export targets are:

- `.xlsx` — recommended; multiple sheets, formulas and cached values, styles, merges, dimensions, links, notes, and names. Untouched XLSX Save As operations copy the original package exactly. Edited rich OOXML workbooks overlay changes onto the source workbook so supported validations, conditional formatting, tables, images, views, and print settings remain attached.
- `.xls` — Save keeps the original file extension and location. Edited workbooks pass through the local LibreOffice engine to create a genuine BIFF workbook. The output is checked for expected sheets, values, and formulas before replacing the file. XLS limits (65,536 rows, 256 columns) are enforced. The first overwrite keeps an original backup in the app's `workbook-backups` folder; use **Original backup** in the compatibility bar to find it. Macros and unmodelled legacy objects may be omitted in the edited workbook; the original backup retains them.
- `.ods` — rich import and export through the local LibreOffice engine, retaining common styles, dimensions, merges, formulas, Unicode sheet names, and hidden-sheet state. The first edited overwrite creates an original backup, as with XLS. Conversion can still simplify advanced features.
- `.csv` / `.tsv` — active sheet only; these text formats cannot retain formatting or workbook architecture
- `.pdf` — selected cells, active sheet, or all visible sheets with print-layout controls
- `.html` — a standalone, escaped, formatted web page using the same scope and layout controls as PDF

No local editor can promise lossless conversion between every Excel, Google Sheets, ODF, legacy binary, and delimited-text feature because the formats have different object and formula models. `simple_calc` shows compatibility notes before destructive conversions involving macros, charts, pivots, slicers, external connections, signatures, and other native structures. Unsupported formulas keep their original formula text and cached result even when they cannot be recalculated locally. Macros are never executed, and saving a macro-enabled source as `.xlsx` removes VBA by design.

Unchanged original-format copies preserve the source bytes, including legacy and macro-enabled sources. Save checks for external file changes before replacing the source. Edits made while a save is running remain marked unsaved. Imported default row heights and column widths are honored; compact line boxes, zoom-scaled insets, and predictable fallback fonts keep ordinary numbers readable without changing the source font size or chosen top/middle/bottom alignment. Intentionally undersized rows remain undersized; double-click the row boundary to AutoFit.

Native XLS saving and rich ODS opening/saving need LibreOffice. Discovery checks `SIMPLE_LIBREOFFICE_PATH`, `%LOCALAPPDATA%\simple\office-runtime\program\soffice.exe`, adjacent portable runtime folders, and installed LibreOffice paths. Each conversion uses an isolated profile with macros and external updates disabled and a 60-second timeout. If the engine is unavailable or validation fails, the original file is left untouched and an actionable error is shown. ODS operations can take about 11 seconds per conversion on the tested machine; XLSX remains the faster working format.

Additional reliability checks: `npm run test:readability` runs the actual Electron grid at 50%, 100%, 150%, and 200% zoom and verifies source dimensions/alignments through an XLSX round trip. `npm run test:native-xls-save` needs the Office runtime and exercises editing plus the toolbar Save action on a disposable XLS copy, formulas, styles, exact unchanged bytes, and external-change protection.

### Imported form layout and printing

The legacy importer now reads explicit BIFF row heights even when Excel marks a row as automatically sized. It retains the workbook's Normal font, print area, saved margins, header/footer text, scale, and the separate fit-to-page flag. Dormant fit dimensions do not override an active percentage scale. After a genuine XLS conversion, fixed-size BIFF layout records restore exact row heights, column widths, margins, and hidden column tails that the converter would otherwise round or expand.

**Fit data width** and **Fit sheet** are available in View and the zoom selector. They change only the view. Pixel-aligned cell edges keep thin outside borders and merged entry boxes visible at fractional zoom; matching shared borders are painted once. Imported alignment and font sizes remain unchanged.

Print and PDF/HTML export default to **Use saved page layout**. Turn it off to choose manual page settings. Printing measures digits in the workbook's Normal font to resolve native column widths; the editor retains its familiar screen widths. Wrapped text stays within source row geometry, ordinary text can extend through blank neighboring cells, accounting amounts keep their currency symbol and right alignment, and merged cells collect their outside borders. Unsupported header/footer pictures produce a visible warning; native picture objects are not reconstructed in edited XLS copies. Mixed saved paper sizes remain supported in PDF; physical printing asks for separate sheets because a single printer job has one paper size.

The schedule fixture is checked in the real Electron app on a disposable copy. It opens at 100%, fits all 37 form rows at 73%, saves edited values and formulas back to true XLS, then exports a one-page Letter PDF at the source's 91% scale. Checks include the complete invoice label, unclipped contract paragraph, currency alignment, all three closed signature boxes, exact dimensions and margins, Normal font, original-file hashes, backups, and external-change rejection. A local run on 2026-09-05 measured 1.46 s to open, 0.28 s to fit, 0.25 s to preview, 0.37 s to export PDF, and 11.70 s to save edited XLS; these are sample timings, not performance guarantees. The Google export of this file paginates differently from its native saved layout, so it is not used as a print-fidelity oracle.

Format behavior follows Microsoft's [ROW record specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/4aab09eb-49ed-4d01-a3b1-1d726247d3c2), [WsBool fit-to-page flag](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/ccbd73f9-ff1d-4069-be31-13d16c074ec4), [Setup record specification](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/cfe70934-f513-4fc9-bd95-8e81e53a6072), and [column width definition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-3.0.1). Broad format decoding uses [SheetJS write options](https://docs.sheetjs.com/docs/api/write-options/) and [row properties](https://docs.sheetjs.com/docs/csf/features/rowprops/).

Further fidelity work should cover native header pictures/drawings, repeated heading columns, a broader corpus of multilingual fonts, unusual merged layouts, and unsupported formulas. Exact source-byte copies and original XLS/ODS backups remain the recovery path for features outside the editor's model.

### Complex workbook stress checks

`npm run test:complex-workbook` creates a synthetic workbook through the app's real creation/save API, then uses the live editor to change inputs and formatting. It exercises four sheets, 200 transaction rows, 609 linked formulas, accounting formats, Japanese/Arabic/accented text, explicit text identifiers, tight rows, merged paragraphs and boxes, a hidden audit sheet, saved print areas, repeated headings, and mixed A4/Letter pages. It saves/reopens native workbooks and exports PDF, HTML, CSV, ODS, and XLS into ignored `tmp/complex-stress`; it never reads or changes a user's workbook and never prints a physical job.

The stress pass fixed lost ODS styles/hidden sheets/XML-escaped sheet names, dropped modern Normal-font metadata, stale successful results surviving genuine formula errors, percentage input being stored as text, missing master-only merged borders, clipped outside borders on full-width fitted pages, ignored repeated heading rows, relative print-area references ignored by native readers, and unwanted page-number restarts after saving. Native XLS constant errors are checked against their source and their simple BIFF error token is kept consistent with the cached error. Imported unsupported formulas can still retain a clearly marked cached result; genuine calculation errors are saved and exported as errors. Explicit Text format and apostrophe-prefixed input remain text.

`npm run test:print-geometry` independently renders generated thin/thick merged borders in Chromium and checks that every table stays inside its printed page. An optional path to a synthetic saved workbook creates a complete PDF/HTML review copy in `tmp/complex-print`.

Print heading behavior follows Microsoft's [PrintTitleRows definition](https://learn.microsoft.com/en-us/office/vba/api/excel.pagesetup.printtitlerows). ODS uses LibreOffice's documented [conversion filters](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html). This stress case does not claim to validate every authoring feature or full Excel compatibility.

## Keyboard

- `Ctrl+O` open
- `Ctrl+S` save; `Ctrl+Shift+S` save as
- `Ctrl+P` opens the live page preview and sends the verified layout directly to the default Windows printer
- `Ctrl+Z` / `Ctrl+Y` undo / redo
- `Ctrl+C` / `Ctrl+X` / `Ctrl+V` copy / cut / paste
- `Ctrl+B` / `Ctrl+I` / `Ctrl+U` formatting
- `Ctrl+F` find
- `Ctrl+K` add or edit a link
- `Ctrl+~` show formulas; `Ctrl+\` clear direct formatting
- `Alt+Shift+5` strikethrough
- `Shift+F2` add or edit a note; `Ctrl+Alt+M` adds an offline comment saved as an Excel note
- `Shift+F11` add a sheet; `Space` toggles a selected checkbox
- `Ctrl+D` fill the selected range down; `Ctrl+R` fill it right
- `F2` or `Enter` edit a cell
- Arrow keys move; `Shift` + arrows extends the selection
- `Delete` clears cell contents without removing formatting
- `Ctrl+T` / `Ctrl+L` create a table; `Ctrl+Shift+L` toggles a filter
- `Ctrl+Shift+=` insert cells; `Ctrl+-` delete cells
- `Alt+Shift+→` / `Alt+Shift+←` group / ungroup rows or columns
- `Alt+F5` refresh the pivot table; `Ctrl+Alt+F5` refresh all
- `Alt+=` AutoSum; `Ctrl+1` Format Cells; `Ctrl+F3` named ranges

## Development

```powershell
npm install
npm run dev
```

Production builds:

```powershell
npm run build:dir       # unpacked Windows app
npm run build           # portable executable
npm run build:installer # NSIS installer with file associations
```

The app uses the official SheetJS CE 0.20.3 package for broad format decoding and ExcelJS for richer modern workbook handling. Workbook data remains on the local machine.
