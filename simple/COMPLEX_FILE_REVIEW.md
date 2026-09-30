# Complex-file stress review — September 5, 2026

Simple was exercised through its real editors, local file pickers, save handlers and exports. This pass fixed visible defects that simpler fixtures missed. The results below describe the tested files; they do not establish Word, Excel or Acrobat parity.

## Spreadsheet creation and external-reader checks

Created a four-sheet workbook with 200 transaction rows, 609 formulas, cross-sheet references, a hidden audit sheet, merged headings and notes, mixed fonts and row heights, Japanese/Arabic/accented text, accounting amounts, long identifiers, native error cells, print ranges and repeated headings. Edited inputs in the real grid, saved, exported and reopened XLSX, XLS and ODS, and checked PDF/HTML/CSV exports.

Fixed the following concrete failures:

- Rich ODS conversion dropped styles, hidden-sheet state and changed an ampersand in a sheet name, breaking references. It now uses the isolated Office converter in both directions. Normal-font metadata also survives XLSX reopen.
- Valid spreadsheet error cells were rejected on XLS Save; one converter error token disagreed with its cached value. Native errors now survive conversion and recalculation.
- Entering `8%` became text; formulas could then show stale cached totals. Numeric percentage entry and genuine error propagation now work in the grid, Save and HTML export; corrected inputs recalculate successfully.
- Merged boxes lost outer borders, the content edge clipped a full-width border, and repeated headings were ignored. Real Chromium print geometry and full PDF pages were checked after correction.
- Saved print-area references used relative row coordinates. Simple showed the intended range, but another engine printed all 200 invoices over 10 pages. Fully absolute references now produce the intended 60 invoices and five pages in both Simple and the independent engine.
- Merged headings and notes beginning in a frozen column were partly hidden by the frozen area. Each visible pane now draws its clipped part of the same merged cell, and selection outlines cover the full merge. Only one accessible master and one keyboard control remain for each cell.
- Imported font names and sizes outside the preset menus displayed the wrong choice: a 21-point heading showed 8. The menus now include the actual imported name and size without changing the file's formatting.
- The import library truncated fractional font sizes, reading 13.5 as 13. A per-workbook adapter restores exact OOXML sizes before applying styles. Tests retain 13.5/13.25, row 8.5, column 17.75 and Normal 9.5 through repeated save/reopen cycles.

The final frozen-pane review checked merged cells crossing rows, columns or both, far scrolling and return, 75%/100%/125% zoom, pointer selection and editing from either side, keyboard controls and the imported Cambria 13.5-point value. Viewing, selecting, scrolling and canceling an edit left every original workbook byte unchanged. [Pane verification](qa/complex-files/calc-frozen-merges-results.json), [complete heading and note](qa/complex-files/calc-frozen-merge-fixed.png), [125% view](qa/complex-files/calc-frozen-merge-125.png), [editing after scrolling](qa/complex-files/calc-frozen-merge-editor.png).

The original supplied XLS form was also rechecked: one page at its saved 91% print scale, native XLS Save, readable values and unchanged original SHA-256. [Complex results](qa/complex-files/calc-results.json), [summary page](qa/complex-files/calc-summary.png), [multilingual and tight-row page](qa/complex-files/calc-notes.png).

The complex workbook took 145 ms to create and save; PDF export 725 ms, HTML 185 ms and CSV 187 ms in the final run. Rich native conversion remains slower: XLS 12.236 seconds and ODS 13.384 seconds under concurrent test load (ODS was 10.851 seconds in an earlier quieter run). This is a remaining speed gap, traded for retaining the original format and layout.

## Document creation and reopen fidelity

Built a six-page document from a blank Simple editor with 13 table values, lists, pasted content, an inline picture, a square-wrapped picture, a footnote, manual line breaks, portrait and landscape sections, and Arabic, Chinese, Japanese and Korean text. Saved and independently inspected DOCX, PDF, HTML, Markdown and plain text, and converted DOCX to native DOC and back.

Fixed the failures this larger document exposed:

- Inserting a table left the caret after the table. It now enters the first cell, so typing fills the table.
- Local pictures disappeared from standalone HTML and Markdown exports. Their bytes now travel with the exported document. Header/footer variants, notes and manual line breaks also survive these exports.
- Pasted CSS colors were misread during PDF export, and the selected CJK font lacked some Japanese and Korean glyphs. Colors are normalized and available installed fonts are checked per glyph.
- Square-wrapped pictures were exported as inline pictures. DOCX now retains the supported floating anchor and wrap settings. Footnotes beside wrapped pictures are reserved and drawn correctly.
- Shift+Enter produced an invalid raw XML control character, breaking independent DOC conversion. It now saves and reopens as a Word line-break element.

The final fresh-window reopen retained all text and six pages. Pages two through six rendered pixel-identically. Page one differed in 0.218% of pixels, with a maximum text shift of 0.03125 point from Word's coordinate rounding; visual inspection found no meaningful layout change. Strict XML parsing, exported colors, multilingual text, the footnote beside the picture and native DOC conversion passed. [Document results](qa/complex-files/docs-results.json), [tables](qa/complex-files/docs-tables.png), [multilingual text](qa/complex-files/docs-multilingual.png), [landscape page](qa/complex-files/docs-landscape.png).

Creating headers/footers and page-number fields from blank documents, nested table creation, structured HTML paste, and complex long footnotes still need work. Font availability also varies between computers.

## PDF books: what changed

All 27 PDFs in the local Books folder were copied and inventoried. Eight text-edit scenarios across seven books and one annotation on a scanned book were saved and checked. Originals remained byte-identical by SHA-256. The inventory includes books up to 1,557 pages; the largest volume actually edited was 1,277 pages. Each text edit was reopened in Simple and independently rendered with MuPDF.

| Reproduced problem | Fix and evidence |
| --- | --- |
| Berling serif text became sans-serif in another PDF reader after Save; the Myriad cover face also changed. | Corrected a malformed CFF font-subset header generated by the font library. Saved glyphs now use the original embedded face. [Before the fix](qa/complex-files/pdf-font-before-fix.png), [fixed result](qa/complex-files/pdf-font-fixed.png). |
| Black EB Garamond was sampled as purple RGB 16/8/24; gray DejaVu 26/26/26 became 48/48/48. | Read the original PDF paint color for confidently matched text. Ambiguous repeated text, masks, transparency and special blending retain the sampled fallback. |
| Fragments of original “g” and “y” remained below the replacement. | Cover the source glyph ink extents, including descenders, while preserving the original text baseline and box calibration. [Saved comparison](qa/complex-files/pdf-descenders-fixed.png). |
| A long page at 400% allocated 41,028,736 screen pixels despite a nominal 24-million limit. | Enforce the area and edge limit at every scale, including below one backing pixel per screen pixel, and release the temporary bitmap after rendering. The same page now uses 23,997,924 pixels: about 96 MB instead of 164 MB per RGB-with-alpha bitmap. PDF vector output is unchanged. |
| Edit mode gave the usual text-editing instruction on an image-only scan. | Explain when a page has no editable text and show what can be added. A new note saved and reopened on a 319-page scan; its scan image and the other 318 pages’ content streams stayed intact. |

Additional inspected comparisons: [white text on the blue cover](qa/complex-files/pdf-cover-fixed.png) and [Cyrillic monospaced text](qa/complex-files/pdf-cyrillic-fixed.png).

| Book/font scenario | Open to selected text | Save |
| --- | ---: | ---: |
| Napoleon, 1,277 pages / Times New Roman | 939 ms | 1,995 ms |
| Elon Musk, 873 pages / EB Garamond | 789 ms | 1,922 ms |
| Colored cover, 260 pages / Myriad | 468 ms | 570 ms |
| Dense body, 260 pages / Berling | 683 ms | 473 ms |
| Mixed regular/bold insurance text / DejaVu | 360 ms | 164 ms |
| Russian volume, 783 pages / Liberation Mono | 704 ms | 3,608 ms |
| Long single-page article / Crimson Text | 309 ms | 55 ms |
| Single-page letter / Calibri | 275 ms | 103 ms |

Timings are individual local runs with warm OS caches, not product benchmarks. On all eight edited pages, the 144-DPI comparison found no pixel change outside the source line’s bounding box plus three points, using a channel-difference threshold of 8/255. All other pages’ decoded content streams matched byte-for-byte. This checks local disturbance, not perfect glyph identity within the edited line. High-zoom comparisons were also inspected for color, font shape, baseline, descender remnants and surrounding content. [Detailed PDF results](qa/complex-files/pdf-results.json).

## Images and video

Image creation/editing tests used a 12.6-megapixel composition with distinct quadrants, half transparency, rotation, crop, painting and undo, plus EXIF-oriented JPEGs. Fixed undo restoring stale transparency information and incorrect Fit sizing after crop/rotation. Temporary canvases release their backing pixels immediately; right-angle rotation avoids another full-image transparency scan.

Unchanged PNG/WebP saves retained exact source bytes. Rotated pixels and alpha values matched expectations; PDF output preserved geometry and transparent-edge appearance. The measured image open was 130 ms. [Image results](qa/complex-files/image-results.json).

A 65-second, odd-sized 321×481 portrait video revealed that Fit produced a 1280×1918 element inside a 1280×754 stage, visibly cutting off the video. The layout now fits within the stage. Frame export names include milliseconds so two captures within the same second get distinct suggested names. Rapid-seek PNG captures matched independently decoded FFmpeg frames with mean channel error around 0.04/255 and no channel error over 24; an interrupted capture correctly refused to save a stale frame, then recovered. [Video results](qa/complex-files/video-results.json).

Edited JPEG/WebP remain lossy. HDR, unusual video codecs and every camera profile were not covered.

## Combining complex files

The actual shared combining worker passed a 14-page ordered job: five spreadsheet pages, six Word pages, one edited book page, a transparent PNG and an EXIF-rotated JPEG. A reordered job using existing PDF exports produced six pages. The native spreadsheet conversion now honors the 60-invoice print range; formula totals were independently checked against the source quantities and prices. Selected spreadsheet, landscape Word and book pages rendered identically before and after combining. Image alpha/orientation, invalid-range recovery and all input hashes passed. [Combine results](qa/complex-files/combine-results.json).

The supplied legacy XLS exposed another failure: the native converter printed the black header-color code as literal `000000`. A conversion-only adapter now removes that recognized formatting code from a disposable copy. All other internal streams, workbook record positions and file length stay unchanged; Save never uses this adapter. Independent conversion retained one page, all 107 body words and their exact positions, and every retained header/footer text item. Malformed allocation chains, overlapping streams, encrypted files, unsupported containers, literal escapes and other colors are guarded. The adapter took 1.53 ms. This fixes the supplied black-color case; other legacy header colors and unsupported containers still need coverage. [Header verification](qa/complex-files/legacy-header-results.json), [corrected native output](qa/complex-files/legacy-header-fixed.png).

Cold native spreadsheet-plus-Word reference conversion took 22.23 seconds in the earlier full run. The final document's uncached reference conversion took 11.2 seconds; the final mixed worker job with conversions cached took 1.492 seconds, and the existing-PDF job took 1.466 seconds. Reusing a managed, isolated Office process is a concrete next performance project.

## Remaining priorities

1. **PDF text replacement and paragraph layout.** Current edits visually cover source text. Original wording may remain extractable/searchable underneath. A native content replacement engine is required before claiming true replacement or secure redaction; paragraph-aware reflow is needed for larger edits. Changes over textured backgrounds, missing glyphs and complex scripts need broader fixtures.
2. **Recognition for scanned books.** The new explanation and annotation flow do not supply OCR. Acrobat recognizes scans and reconstructs editable text; Simple still needs that separate capability and visual comparison against noisy, skewed and multilingual scans. [Adobe’s scanned-PDF workflow](https://helpx.adobe.com/acrobat/desktop/create-documents/scan-documents-to-pdfs/edit-scans.html).
3. **Move long saves off the window process, preserve undo across Save, and restore reading position.** Large-file saves still briefly occupy the process that owns windows. Recovery, cancellation and a saved-state checkpoint should be verified together.
4. **Expand a durable compatibility corpus.** Keep original/rendered/edited/reopened comparisons for nested tables, complex headers, drawings, formulas, font substitution, print areas and foreign-language documents. Add each reproduced defect to the corpus before changing the engine.

Even professional PDF editors can substitute unavailable fonts and alter the result. Matching embedded glyphs and checking saved output is more dependable than promising that every replacement will be indistinguishable. [Adobe’s font and text-editing guidance](https://helpx.adobe.com/acrobat/using/edit-text-pdfs1.html).

## Delivered files and validation

The combined app and the standalone spreadsheet, document and PDF executables in the workspace root were rebuilt and replaced. Previous versions are retained in `.cleanup-backups/complex-stress-upgrade-20260905`. [Delivered build hashes](qa/complex-files/delivered-builds.json) identify the exact checked copies. Separate image and video portable builds are also available in their source release folders; both are included in the combined app.

The final packaged combined app passed toolbar and keyboard print-preview checks for all five workspaces (ten checks), plus Export As toolbar and keyboard checks in all five. These checks opened and inspected previews and menus; they did not submit a physical print job. The packaged Combine dialog also passed native Word/XLS conversion, page ordering, range validation, image orientation, original-file protection and recovery after a failed job; the original XLS header no longer showed the stray color digits. [Final packaged verification](qa/complex-files/packaged-verification.json), [packaged Combine results](qa/complex-files/packaged-combine-results.json).

Source verification included 54 file routes, five current staged module builds, 24 launcher/shared tests, 41 PDF tests, 49 document tests with editor-library regressions, 368 spreadsheet formula assertions plus format/print checks, 22 image tests and 19 video tests. The larger save, reopen, independent-render and conversion results above supplement these checks.

The [example folder](../Simple%20test%20examples/README.md) contains the created document and workbook in editable and exported formats. All 27 Books originals and their source test copies were hash-checked again after testing and remained unchanged.
