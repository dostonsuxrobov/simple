# Complex files to try in Simple

These are synthetic test files, not real business records. They were created and edited using Simple's production document and workbook systems.

- [Complex workbook.xlsx](Complex%20workbook.xlsx): four sheets, 200 transaction rows, 609 formulas, a hidden audit sheet, merged notes, multilingual text and five printed pages. Some error cells are intentional. Change the tax-rate input to test recalculation.
- [Complex workbook.pdf](Complex%20workbook.pdf): the intended five-page spreadsheet output, including only the first 60 invoices in the transaction print range.
- [Complex workbook.xls](Complex%20workbook.xls) and [Complex workbook.ods](Complex%20workbook.ods): native format exports for save/reopen checks. Rich conversion needs the local Office runtime and takes longer than XLSX or PDF.
- [Complex document.docx](Complex%20document.docx): a six-page document with tables, lists, wrapped images, a footnote, manual line breaks, mixed page orientation and multilingual text.
- [Complex document.pdf](Complex%20document.pdf): the corresponding document output for visual comparison.

The [stress review](../simple/COMPLEX_FILE_REVIEW.md) explains the fixes, measurements and remaining limits. Your original Books PDFs were never edited; the local test copies and detailed inventory are in `.codex-tmp/books-stress-20260905` under the Simple workspace.
