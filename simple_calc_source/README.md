# simple_calc

`simple_calc` is a minimal, local-first spreadsheet editor for Windows. It opens modern and legacy workbook formats, keeps formulas as formulas, and saves a modern `.xlsx` by default for exchange with Excel, Google Sheets, Apple Numbers, and LibreOffice Calc. For modern Excel packages, saving is source-backed so workbook structures supported by the rich reader are retained instead of rebuilding every edited file from a blank workbook.

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
- Atomic local saves through an isolated Electron bridge

## Formula compatibility

Alongside arithmetic, comparisons, ranges, cross-sheet references, and core functions such as `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `IF`, `AND`, `OR`, `ROUND`, `TODAY`, and `NOW`, the live engine supports common Google Sheets / Excel families:

- Conditional totals: `SUMIF(S)`, `COUNTIF(S)`, `AVERAGEIF(S)`
- Math and statistics: `PRODUCT`, `MEDIAN`, `LARGE`, `SMALL`, `RANK`, `STDEV.S`, `VAR.S`, `ROUNDUP`, `ROUNDDOWN`, `INT`, `MOD`, `POWER`, `SQRT`
- Logical and error handling: `IFERROR`, `IFS`, `XOR`
- Text: `LEN`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `UPPER`, `LOWER`, `PROPER`, `JOIN`, `TEXTJOIN`, `SUBSTITUTE`, `FIND`, `SEARCH`, `EXACT`, `VALUE`, `CONCAT`
- Dates: `DATE`, `YEAR`, `MONTH`, `DAY`, `EDATE`, `EOMONTH`
- Lookups: `INDEX`, `MATCH`, `VLOOKUP`, `HLOOKUP`

Imported formulas outside the live engine are still preserved when the workbook is saved. Dynamic arrays, whole-column formula references, structured table references, and named-range evaluation are not yet calculated locally.

## File compatibility

The rich OOXML importer handles `.xlsx`, `.xlsm`, `.xltx`, `.xltm`, and `.xlam`, with a compatibility-reader fallback. The compatibility importer also accepts `.xlsb`, `.xls`, `.xml`, `.ods`, `.fods`, `.numbers`, `.csv`, `.tsv`, `.txt`, `.slk`, `.sylk`, `.dif`, `.dbf`, `.prn`, and common Lotus / Works workbook extensions.

Save and export targets are:

- `.xlsx` — recommended; multiple sheets, formulas and cached values, styles, merges, dimensions, links, notes, and names. Untouched XLSX Save As operations copy the original package exactly. Edited rich OOXML workbooks overlay changes onto the source workbook so supported validations, conditional formatting, tables, images, views, and print settings remain attached.
- `.ods` — broad open-format exchange
- `.csv` / `.tsv` — active sheet only; these text formats cannot retain formatting or workbook architecture

No local editor can promise lossless conversion between every Excel, Google Sheets, ODF, legacy binary, and delimited-text feature because the formats have different object and formula models. `simple_calc` shows compatibility notes before destructive conversions involving macros, charts, pivots, slicers, external connections, signatures, and other native structures. Unsupported formulas keep their original formula text and cached result even when they cannot be recalculated locally. Macros are never executed, and saving a macro-enabled source as `.xlsx` removes VBA by design.

## Keyboard

- `Ctrl+O` open
- `Ctrl+S` save; `Ctrl+Shift+S` save as
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
