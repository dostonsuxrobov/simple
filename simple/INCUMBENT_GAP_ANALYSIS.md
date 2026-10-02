# Simple vs. the incumbents: gap analysis and hardening pass

October 2, 2026. This review compares each Simple workspace with the product it should feel as strong as:

| Workspace | Compared with |
|---|---|
| PDF | Adobe Acrobat Pro |
| Spreadsheets | Microsoft Excel and Google Sheets |
| Documents | Microsoft Word and Google Docs |
| Images | Adobe Photoshop for Advanced mode, Windows Photos for Simple mode |

It records what was implemented in this pass and what remains.

**How the analysis was done:**
- An inventory of every workspace, built from its code.
- A feature checklist for each incumbent, covering 120–250 features with the small behaviours that make it feel polished.
- Bug hunters with 11 different focuses. A second agent independently checked every bug they reported; none were refuted.
- A Save / Import / Export audit with real round trips on copies of the sample files. LibreOffice was absent, as it is on the owner's PC.

Everything stays local: there are no cloud services and no downloads. A file that another program is holding is treated as an ordinary locked file.

## Summary

The "Gaps implemented" column counts against the complete incumbent checklist. Many open gaps are deliberately left at P2 or P3 to keep Simple minimal.

| Workspace | Gaps implemented | Verified bugs fixed | Save/Import/Export items done |
|---|---|---|---|
| PDF vs Adobe Acrobat Pro | 16/67 | 35/36 | 19/32 |
| Spreadsheets vs Microsoft Excel and Google Sheets | 25/52 | 36/36 | 18/30 |
| Documents vs Microsoft Word and Google Docs | 19/47 | 24/24 | 15/27 |
| Images vs Adobe Photoshop (Advanced) and Windows Photos (Simple) | 37/61 | 23/24 | 15/26 |

## What changed in this pass

### Save, Import and Export (the "edited files do not get saved" problem)

**Root causes found:**
- Some edits never marked the file as changed. In Documents this included Find & Replace, the right-click menu, the mini-toolbar, ruler and picture drags, and Ctrl+Enter.
- An entry still being typed, or a brush stroke still in progress, was not committed before saving.
- Saving claimed success while silently dropping content: PDF form values, comments and suggestions, and animation frames.
- Close prompts defaulted to discarding (PDF, Spreadsheets).
- A file held open by another program failed to save with a raw error.
- Edited `.xls`, `.ods` and `.doc` files could not be saved without LibreOffice.

**One shared save layer** now lives in `simple/shared` and is vendored into each workspace. A drift check makes sure the copies stay identical.
- **`safeWriteFile`**: writes a temporary file in the same folder, flushes it, reads it back to verify it, then replaces the original. It retries while another program holds the file, refuses read-only files, never leaves partial files behind, and repairs interrupted saves at startup.
- **Native Save / Don't Save / Cancel**, with Save as the default, on every close path: title bar, Alt+F4, taskbar, quitting and Windows sign-out. A window never closes during a save, and crashed or hung windows stay closable.
- **One format registry with content sniffing.** Mislabelled files still open, and routing and file pickers come from the registry.
- **A locally installed LibreOffice is used only when present.** Without it:
  - Spreadsheets saves an edited `.xls` or `.ods` as an `.xlsx` beside the original, and opens ODS natively.
  - Documents saves an edited `.doc` as a `.docx` beside the original.
  - Combine converts Word and Excel files natively.

**Per workspace:**
- **Spreadsheets:**
  - The entry being typed is committed before saving.
  - CSV saves follow RFC 4180, keep 4-digit years, and keep the original delimiter and encoding.
  - New sheets no longer inherit a deleted sheet's protection, tables or pictures.
  - Shared formulas and very large validation lists save correctly.
- **Documents:**
  - The changed state now comes from the document model.
  - Opening a file scans it for content the editor cannot keep, and the first save then defaults to "Save a copy".
  - A failed import can never overwrite the real file.
  - TXT, Markdown, HTML, RTF, ODT and DOCM/DOTX open natively, and RTF and ODT can be exported.
- **PDF:**
  - Form values are saved, including length limits, multi-line fields, radio buttons, dropdowns and lists.
  - Non-Latin and CJK text no longer fails or disappears.
  - Image edits no longer corrupt the page.
  - Deleted pages and images are really removed from the file.
  - Signed PDFs save as a copy.
  - Password-protected PDFs ask for the password instead of hanging.
  - Word, Markdown, HTML, TXT, XLSX, RTF, ODT and image import work without LibreOffice.
  - Text exports keep line breaks, and Word export keeps headings, formatting and images.
- **Images:**
  - EXIF, ICC, DPI and text metadata are kept.
  - Lossless WebP stays lossless.
  - Animations are no longer flattened in place.
  - JFIF, ICO, APNG, SVGZ and PSD open, as do mislabelled files and Illustrator or UTF-16 SVGs.

### Headline features

**PDF: Recognize text and edit scanned text, as in Acrobat**
- OCR is fully offline: Tesseract 5 LSTM with the English best model.
- Before recognition, pages are deskewed and their contrast and noise cleaned up.
- The result is an invisible text layer, so search, selection, copy and export work in Simple and in other readers.
- In Edit mode, click a recognized line and type:
  - the original glyphs are retouched out using matched background grain;
  - the new text uses a matching font class, weight, size and colour;
  - unedited words stay pixel-identical.
- Retouching is not secure redaction.

**Images: one Advanced button that opens a Photoshop-style editor**
- Layers: raster, text, shape and adjustment layers, plus groups, masks, clipping, all 27 blend modes, opacity and locks.
- Selections: marquee, lasso, polygonal lasso and magic wand, with feather, grow/contract and invert.
- Tools: 18, including clone stamp, spot healing and free transform.
- 13 non-destructive adjustment layers (Levels, Curves, Hue/Saturation and others) and 10 filters.
- A History panel and Photoshop shortcuts.
- PSD open and save through ag-psd.
- The Simple button returns to Simple mode. It asks first only when layers would be lost.

**Images: Simple mode**
- Resize, flip, and straighten with automatic cropping.
- Crop presets.
- Seven adjustment sliders, plus Auto, Compare and Looks.
- Markup: text, arrows and shapes.
- Paste.
- Zoom toward the cursor; Space-drag, middle-drag and pinch to pan; crisp pixels when zoomed in.
- Much deeper undo on large photos.

**Spreadsheets: Excel-exact semantics and polish**
- Formula semantics:
  - Criteria in COUNTIF and SUMIFS match Excel.
  - Operator precedence: `-2^2` gives 4.
  - Lookups accept wildcards.
  - Older formulas use implicit intersection.
  - Dates follow the 1900 date system.
  - Unformatted numbers display in General format.
- Calculation options: manual mode, F9 to recalculate, iterative calculation and a circular-reference finder.
- Clear messages for formula errors.
- Fixes:
  - Paste Values pastes the right cells.
  - Copy, Delete and fill skip rows hidden by a filter.
- Editing:
  - double-click on the fill handle;
  - Alt+Enter turns on wrapping;
  - a sort warning before sorting one column of a table;
  - AutoComplete from the column;
  - Ctrl+mouse-wheel zoom up to 400%;
  - internal hyperlinks;
  - Excel's keyboard staples.
- Print, PDF and HTML exports now show conditional formatting, table styles, sparklines and checkboxes.
- Printing lets you choose the printer, copies, page range, scale and margins.
- Dropdown chips, and drag-to-move cells.

**Documents: Word and Google Docs essentials**
- Shortcuts:
  - Word and Docs shortcuts, which also work on non-Latin keyboard layouts.
  - Undo and Redo buttons.
  - Ctrl+Backspace and Ctrl+Delete delete whole words.
  - Command search with Alt+Q.
- In-app dialogs replace the broken `window.prompt` ones.
- Spell check uses Windows' own local spellchecker: squiggles, suggestions and Add to dictionary.
- AutoFormat: smart quotes, dashes, automatic lists and automatic links.
- Header, Footer and Page number commands, and a Paragraph dialog.
- Every installed font, with search.
- Paste:
  - paste or drop images;
  - Ctrl+Shift+V pastes plain text;
  - pasted HTML keeps lists, tables and headings.

**Launcher**
- Combine works without LibreOffice.
- Opening many files at once no longer drops any.
- Combine keeps internal links, never copies pages you didn't select, and unlocks PDFs that only restrict permissions.

## Still open (recommended next steps)

**PDF:**
- Real comment annotations with a Comments pane (PDF-018, 019, 021).
- Secure redaction (PDF-032).
- Password protection (PDF-033).
- Compression (PDF-058).
- Stamp pages: page numbers, headers and footers, watermark, Bates (PDF-041).
- Organize and split pages (PDF-036, 037).
- Paragraph reflow (PDF-005).
- pdf.js CMaps for CJK text (PDF-046).
- Restoring the reading position (PDF-047).
- A single "More tools" panel (PDF-064).
- The same file opened in two windows (pdf-io-print-export-9).

**Spreadsheets:**
- AutoRecover through the shared journal (CALC-003).
- Locale-aware input (CALC-012).
- Formatting whole columns on very large sheets (CALC-016).
- Rich text inside cells (CALC-025).
- Stale tracked backup files to delete (CALC-048).

**Documents:**
- Real comment and tracked-change round trips (DOC-SIE-27).
- Keeping DOCX parts the editor does not own when saving (DOC-SIE-26).
- Table of contents and captions.

**Images:**
- An Export As dialog with quality, size estimate and resize (IMAGE-020).
- HEIC and TIFF (IMAGE-029).
- A redact tool (IMAGE-025).
- Copy text from image using the PDF OCR engine (IMAGE-026).
- Crop-edge clamping (image-editing-correctness-7).

**All workspaces:** the shared layer already contains:
- the autosave/recovery journal;
- one window per file;
- the Export As dialog model.

They still need to be wired into each workspace's page.

## How to verify

- `cd simple && npm test`: the format registry, routing, safe-write lock scenarios, guards, recovery and Combine.
- In each workspace folder (simple_*_source):
  - PDF: `npx tsc --noEmit` and `node --test tests/*.test.cjs`.
  - Images: `npm test`, `npm run smoke` and `npm run smoke:advanced`.
  - Documents: `npm test`.
  - Spreadsheets: each suite with `npm run test:<suite>`.
- `cd simple && npm run build` builds `simple/release/simple.exe`.

## Detailed status

## PDF vs Adobe Acrobat Pro

### Capability gaps (67; 16 implemented in this pass)

| ID | Priority | Area | Gap | What the incumbent does | Status |
|---|---|---|---|---|---|
| PDF-001 | P0 | OCR & scans | Recognize Text (OCR) as an invisible, searchable text layer | Scan & OCR > Recognize Text, with a language and resolution choice, on a page range or the whole document. The original image is kept and an invisible text layer is adde… | Implemented (OCR work packages 1-3) |
| PDF-002 | P0 | OCR & scans | Edit OCR'd text so it looks like the original (matched font, sampled background) | Edit PDF on a scan auto-OCRs the page and turns paragraphs into editable boxes. Replacement text uses a generated font that mimics the scanned typeface, in the matching … | Implemented (OCR work package 4) |
| PDF-005 | P0 | Edit text | Paragraph detection and reflow in Edit PDF | Each paragraph gets an editable box. Typing reflows the lines inside it, and the box grows downward without moving the rest of the page. Ctrl+A selects only that box's t… | Open |
| PDF-003 | P0 | OCR & scans | 'Make this page editable' entry point in Edit mode (auto-OCR prompt) | Choosing Edit PDF on a scan runs OCR automatically. A preference chooses between making all pages editable and the current page only, and pages become editable as they f… | Implemented (OCR work package 3) |
| PDF-006 | P0 | Edit text | Save is refused when text overflows its box or an original glyph can't be found | Save never refuses. Boxes grow to fit, and problems are shown as warnings. | Implemented (via PDF-SIE-5) |
| PDF-018 | P0 | Comments & markup | Save markup as real PDF annotations and load existing ones back as editable | Highlight, Underline, StrikeOut, Ink, Square, Circle, FreeText and Text are stored as annotations. QuadPoints cover each line of a multi-line selection. They appear in a… | Open |
| PDF-019 | P0 | Comments & markup | Select, recolour and delete markups; per-tool colour, thickness and opacity | A selected annotation shows handles, can be nudged with the arrow keys and removed with Delete. A properties bar offers colour, opacity and thickness, plus 'Make current… | Open |
| PDF-024 | P0 | Forms | Full AcroForm filling: radio, dropdown, list, multiline, comb/maxLen, zoom-scaled fonts, Tab order, field highlight | Every field type is fillable. Tab and Shift+Tab follow the tab order, and Space toggles checkboxes and radios. A 'Highlight existing fields' toggle is available, and req… | Implemented (via PDF-SIE-3, PDF-SIE-22) |
| PDF-029 | P0 | Signing | Signed-PDF detection is a byte heuristic: misses /Type/Sig and object streams, falsely triggered by unsigned fields | Detects real signatures and opens the Signatures pane. Saves are incremental so signatures stay valid. | Implemented (via PDF-SIE-7) |
| PDF-032 | P0 | Redaction | Secure redaction: mark text, areas, pages and search/pattern hits; apply irreversibly | Marks show as outlined boxes with a preview on hover. Apply removes text, image pixels and vectors underneath, then asks to remove hidden information. Default file name … | Open |
| PDF-046 | P0 | Viewing | pdf.js has no CMaps or standard font data, so CJK and non-embedded-font PDFs render or extract wrongly | Renders every conforming PDF, including Asian-language CID fonts and the non-embedded base-14 fonts. | Open |
| PDF-039 | P0 | Opening and files | Dropped PDFs lose their file path; dropping onto an open document replaces it | A dropped file opens as a real file: Save overwrites it, and it appears in the recent files list. Dropping elsewhere opens a new tab. | Implemented (via pdf-io-print-export-12) |
| PDF-043 | P0 | Bookmarks | Bookmark rename doesn't work (window.prompt is unsupported in Electron) | F2 or double-click renames in place. Ctrl+B names a new bookmark after the selected text. | Implemented (via pdf-io-print-export-8) |
| PDF-044 | P0 | Bookmarks | Outline entries with no page destination (headings, URL actions) are lost on save | Every outline item is preserved. | Implemented (via PDF-SIE-19) |
| PDF-047 | P0 | Viewing | Restore reading position (page, zoom, scroll) per file | 'Restore last view settings when reopening documents'. | Open |
| PDF-056 | P0 | History | Undo history is cleared on Save | Undo keeps working after Save, within the session. | Implemented (via pdf-io-print-export-11) |
| PDF-051 | P0 | Keyboard | Ctrl+Shift+Z does Undo instead of Redo | Ctrl+Shift+Z means Redo, the same as Ctrl+Y. | Implemented (via pdf-viewer-interaction-11) |
| PDF-026 | P1 | Organize pages | Every page operation is blocked on fillable PDFs | Insert, delete, extract, reorder and duplicate all work on forms. Copied fields are renamed or share values, and fields on deleted pages are removed. | Open |
| PDF-004 | P1 | OCR & scans | Scan cleanup: deskew, auto-orient, whiten background, despeckle | Enhance Scans / Camera Image. Options cover deskew, background removal, edge shadow removal, despeckle, and perspective correction with corner handles. | Partial (deskew and background normalisation inside OCR only) |
| PDF-007 | P1 | Search | Find & Replace in Edit mode; search sees edited text | In Edit mode, Find switches to Replace / Replace All, matches the surrounding font, and reflows the replacement. | Open |
| PDF-012 | P1 | Edit objects | Moving or resizing an image re-embeds it as PNG, losing JPEG compression and CMYK/ICC colour | Moving or resizing changes only the placement matrix. The image stream isn't touched. | Open |
| PDF-013 | P1 | Edit objects | Artwork regions are rasterized at screen resolution; the original stays under a white cover | Vector objects stay vector when moved, and the original content really moves. | Open |
| PDF-014 | P1 | Edit objects | Paste from the system clipboard (images and text), and paste in place on another page | Pasting a clipboard image or text creates an object or text box. Pasting on another page keeps the same coordinates. | Partial (newer system clipboard wins over the in-app clipboard) |
| PDF-020 | P1 | Comments & markup | Floating mini-toolbar and context-menu actions on selected text; fix the wrong Highlight hint | Selecting text pops up a small toolbar with Highlight, Underline, Strikethrough, Copy and Redact. | Open |
| PDF-021 | P1 | Comments & markup | Sticky notes and a Comments pane (list, jump, reply, status) | Ctrl+6 adds a sticky note with author and time filled in. The Comments list is grouped by page, with filter, search, replies and status. Clicking a comment selects its a… | Open |
| PDF-025 | P1 | Forms | Form save silently drops values and re-styles every filled field in Segoe UI | Values are kept, and the field's DA font, size and colour are respected. | Implemented (via PDF-SIE-3, PDF-SIE-4) |
| PDF-027 | P1 | Forms | Flatten form, Clear form, and Fill & Sign marks on flat (non-fillable) forms | Fill & Sign places text, check mark, cross, dot, line and date anywhere on a flat form, with A-/A+ sizing. Clear form asks for confirmation. Fields can be flattened on s… | Open |
| PDF-030 | P1 | Signing | Signature from an uploaded image, initials, date, background removal | Type, draw or upload an image for both signature and initials, saved for reuse. A date field. The aspect ratio is kept. | Open |
| PDF-033 | P1 | Security | Password-protect on save (open password, permissions); keep the password on unlocked files | 'Protect using password' for Viewing or Editing, AES-256. Permission options cover printing, changes and copying. | Open |
| PDF-034 | P1 | Security | Encrypted linearized files bigger than 256 KB skip the unlock path and open read-only | Prompts for the password and then allows full editing. | Implemented (via PDF-SIE-6) |
| PDF-036 | P1 | Organize pages | Thumbnail keyboard and context menu: arrows, Ctrl+A, rotate keys, right-click page commands | Arrow keys move through thumbnails and Ctrl+A selects all. Right-click offers Insert, Extract, Rotate, Delete, Crop and Print pages. | Open |
| PDF-037 | P1 | Organize pages | Extract and split (ranges, every N pages, by bookmark) and an Organize grid view | Extract with 'delete after' and 'separate files' options. Split by N pages, size or bookmarks. A full-window Organize grid with an insertion bar, and '+' between pages. | Open |
| PDF-038 | P1 | Create & combine | Dropping several files should combine them, not ignore all but the first | Dropping several files leads to Combine Files, which merges them in order with one bookmark per source file. | Open |
| PDF-041 | P1 | Headers/footers/watermarks | Page numbers, header/footer, text watermark and Bates in one 'Stamp pages' dialog | Header & Footer with 6 positions and tokens. Watermark with opacity, rotation and a behind/on-top choice. Bates numbering. All can be updated or removed later. | Open |
| PDF-048 | P1 | Saving | Autosave and crash recovery for unsaved edits | Autosaves changes to a temporary file every N minutes and offers recovery after a crash. | Open |
| PDF-049 | P1 | Viewing | Zoom: Ctrl+wheel anchored at the cursor, smooth pinch, range 10-1600% | Wheel zoom keeps the point under the cursor fixed. Range 1-6400%, with Marquee and Dynamic zoom. | Open |
| PDF-057 | P1 | Architecture | Move heavy PDF work off the Electron main process into a utilityProcess engine | Long operations show progress and can be cancelled. Other documents stay responsive. | Open |
| PDF-058 | P1 | Optimize | Reduce file size (compress) | Reduce File Size with High/Medium/Low levels, showing the size before and after. Saves to a new file by default. | Open |
| PDF-060 | P1 | Create & combine | Opening DOC/DOCX needs LibreOffice; mammoth is dead code and the README is wrong | Creates PDFs from Office files reliably (PDFMaker). | Implemented (via PDF-SIE-12) |
| PDF-064 | P1 | Chrome & UI | A single 'More tools' panel (the Advanced mode door), with tool search | An All tools pane with search, plus a customizable quick toolbar and contextual floating bars. | Open |
| PDF-008 | P1 | Search | Search options: match case, whole word, F3/Shift+F3, wrap feedback | Whole words only and Case-sensitive options. F3/Shift+F3 move between matches, and a notice appears when the search wraps. | Open |
| PDF-050 | P1 | Keyboard | Navigation shortcuts: Home/End, Alt+Left/Right view history, Space to pan, Ctrl+Shift+N | Home and End go to the first and last page. Alt+Left goes back after link and bookmark jumps. Holding Space switches to the Hand tool temporarily. | Open |
| PDF-009 | P2 | Edit text | Font menu: all installed fonts (including per-user), document fonts first, substitution warning | The font list shows document fonts first, then all system fonts. A warning appears when a font has to be substituted. | Open |
| PDF-010 | P2 | Edit text | Mixed formatting inside one text box (bold word, underline, super/subscript, colour runs) | Bold, italic, underline, superscript/subscript and colour can be applied to part of a box. A mixed selection shows blank values in the panel. | Open |
| PDF-031 | P2 | Signing | Certificate-based digital signatures (PAdES) and signature validation | Sign with a digital ID (PFX or Windows store) with a visible appearance and 'Lock after signing'. The Signatures pane shows Valid, Unknown or Invalid. The signature is s… | Open |
| PDF-059 | P2 | Export | Export to Word with structure (paragraphs, headings, bold/italic, images, simple tables, OCR when needed) | 'Retain flowing text' or 'retain page layout'. It rebuilds paragraphs, tables, lists and columns, and can run OCR if needed. | Open |
| PDF-011 | P2 | Edit text | Text size range limited to 4-96 pt; Add Text ignores the last-used style | Sizes from about 1 to 1296 pt. Add Text uses the last-used font, size and colour, and the box grows horizontally until you drag a width. | Open |
| PDF-015 | P2 | Edit objects | Image crop, free rotation, Shift aspect lock | Rotation handle above the box, Shift+drag keeps the aspect ratio, and an image crop that discards pixels on save. | Open |
| PDF-017 | P2 | Viewing | Snapshot tool (copy an area as an image) | Drag a rectangle to copy that area to the clipboard as an image. A click copies the whole page. | Open |
| PDF-022 | P2 | Comments & markup | More shapes (ellipse, line, arrow) with Shift constrain and fill; smoothed ink and eraser | Rectangle, oval, line, arrow, polygon and cloud, with Shift constraint. Freehand ink is smoothed, and an eraser removes strokes. | Open |
| PDF-023 | P2 | Comments & markup | Stamps (Approved, Draft, Confidential, dynamic date/name, custom image) | A stamp palette of business and dynamic stamps, plus custom stamps made from images. | Open |
| PDF-028 | P2 | Forms | Basic form authoring (text, checkbox, signature, date fields) | Prepare Form auto-detects fields and offers every field type, plus properties, format, validation, calculation, actions and tab order. | Open |
| PDF-035 | P2 | Redaction | Sanitize and remove hidden information | Remove Hidden Information shows a checklist with counts: metadata, attachments, comments, form fields, JavaScript, hidden layers, content outside the crop box. Sanitize … | Open |
| PDF-040 | P2 | Organize pages | Crop: apply to a range or all pages, numeric margins, remove white margins | Set Page Boxes dialog with margins, 'Remove white margins', page range and even/odd options. | Open |
| PDF-042 | P2 | Viewing | Page labels (i, ii, 1, A-1) in the page box and thumbnails | Logical labels appear in the page box, thumbnails and print dialog, and can be typed into the page box. | Open |
| PDF-045 | P2 | Bookmarks | Bookmark tree: expand/collapse, drag to nest and reorder, destination = current view | Hierarchical pane with expand and collapse, drag to nest, Set Destination, and new bookmarks at the current view. | Open |
| PDF-052 | P2 | Viewing | Two-page view with cover page; Rotate View (display only) | Single Page, Continuous, Two-Page and Two-Page Scrolling, with 'Show cover page'. Rotate View changes the display only. | Open |
| PDF-053 | P2 | Viewing | Dark mode for the interface (optional dimmed pages) | Light, dark and system themes. Pages can stay white or be darkened. | Open |
| PDF-054 | P2 | Viewing | Document properties and metadata editor; keep the original metadata honest | Ctrl+D Properties: Title, Author, Subject, Keywords, fonts list, PDF version, page size and file size. Option to show the title in the window. | Open |
| PDF-061 | P2 | Create & combine | TXT/MD to PDF replaces unsupported characters with '?' and doesn't render Markdown | Creates the PDF with full font fallback. | Open |
| PDF-062 | P2 | Create & combine | More image inputs: multi-page TIFF, BMP, GIF, WebP | Creates PDFs from every common image format. TIFF becomes one page per frame. | Open |
| PDF-063 | P2 | Print | Print: pages per sheet, booklet, odd/even, reverse, more paper sizes, comments on/off | Multiple pages per sheet, Booklet, Poster, odd/even pages, reverse, paper source by page size, and Document/Markups/Forms. | Open |
| PDF-016 | P3 | Edit objects | Multi-select, align/distribute, arrange order, smart guides | Shift-click and marquee selection, align and distribute, Bring Forward/Back, magenta smart guides while dragging. | Open |
| PDF-055 | P3 | Viewing | Attachments pane and layer (optional content) toggles | Attachments pane to open, save, add and remove embedded files. Layers pane with eye toggles. | Open |
| PDF-065 | P3 | Accessibility | Read Out Loud | Ctrl+Shift+Y activates it and Ctrl+Shift+V reads the page, using system voices. | Open |
| PDF-066 | P3 | Compare | Compare two PDFs (text changes) | Side-by-side view with synchronized scrolling, a colour-coded change list and a summary report. | Open |
| PDF-067 | P3 | Export | Export tables to the Spreadsheets workspace | Export to Excel detects tables and keeps numbers as numbers. 'Open table in spreadsheet'. | Open |

### Verified bugs (36; 35 fixed)

| ID | Severity | Bug | Status |
|---|---|---|---|
| pdf-viewer-interaction-1 | high | Leaving any text selection keeps every scrolled-past page in memory and freezes the current-page indicator | Fixed |
| pdf-viewer-interaction-2 | high | Deleted pages stay embedded in the saved PDF, and outline entries for them are left dangling | Fixed |
| pdf-viewer-interaction-3 | high | Editing text, adding text or signing on a page that isn't 'current' scrolls the edit box off-screen | Fixed |
| pdf-viewer-interaction-4 | high | Clicking a page never takes keyboard focus, so shortcuts stop working after using the search, page or zoom fields | Fixed |
| pdf-viewer-interaction-5 | high | Copied text glues lines together ('overthe') because the text layer drops PDF.js end-of-line markers | Fixed |
| pdf-viewer-interaction-6 | medium | Page navigation centres the target page, hiding its top whenever pages are taller than the window | Fixed |
| pdf-viewer-interaction-7 | medium | Search hits near the top or bottom of a page are scrolled back off-screen by the 220 ms page re-centre | Fixed |
| pdf-viewer-interaction-8 | medium | An old search hit pulls the view back on every zoom, rotate or refit, and the page indicator falls out of sync | Fixed |
| pdf-viewer-interaction-9 | medium | Ctrl+A / Select All silently selects only the two or three pages currently loaded | Fixed |
| pdf-viewer-interaction-10 | medium | Delete key on a focused bookmark or search result asks to delete pages | Fixed |
| pdf-viewer-interaction-11 | low | Ctrl+Shift+Z performs Undo instead of Redo | Fixed |
| pdf-viewer-interaction-12 | low | At low zoom the first and last pages can never become the current page | Fixed |
| pdf-edit-pipeline-1 | critical | Moving, deleting or replacing a native image corrupts text and inline-image bytes on that page in the saved PDF | Fixed |
| pdf-edit-pipeline-2 | high | Deleted pages, replaced text and deleted images remain recoverable inside the saved file | Fixed |
| pdf-edit-pipeline-3 | high | Editing rotated text (landscape /Rotate pages, vertical labels, inspector Rotation) squeezes it to a sliver and saves it in the wrong place | Fixed |
| pdf-edit-pipeline-4 | high | Inserted JPEGs: EXIF-rotated photos are saved sideways and stretched, and small JPEGs make Save fail with 'SOI not found in JPEG' | Fixed |
| pdf-edit-pipeline-5 | high | Resizing a native-text box with the handles or the H/Y fields does not update baselineOffset: text is saved at a different height, or Save fails permanently wi… | Fixed |
| pdf-edit-pipeline-6 | medium | One bookmark edit rewrites the whole outline: URL and heading-only entries are deleted, nesting is mangled, and every destination becomes /Fit | Fixed |
| pdf-edit-pipeline-7 | medium | Committed text boxes show their text vertically centered, but the editor and the saved PDF put it at the top | Fixed |
| pdf-edit-pipeline-8 | medium | Form text longer than the field's MaxLen is shown on screen but silently dropped on save | Fixed |
| pdf-edit-pipeline-9 | medium | Esc, or Ctrl+Z after using an inspector control, throws away the whole in-progress text edit with no way to redo it | Fixed |
| pdf-edit-pipeline-10 | medium | On OCR'd scans, editing text removes only the invisible OCR layer, so the scanned word stays visible under the new text | Fixed |
| pdf-edit-pipeline-11 | medium | Each text click, and each remount of an edited page, sends the whole PDF over IPC and re-parses it with pdf-lib | Fixed |
| pdf-edit-pipeline-12 | low | Ctrl+V pastes an old in-app text box or image even after the user copied something else | Fixed |
| pdf-io-print-export-1 | critical | The password prompt never appears, so opening any PDF with an open password hangs on "Opening document…" | Fixed |
| pdf-io-print-export-2 | high | Deleted pages stay embedded in the saved PDF: the content can be recovered and the file barely shrinks | Fixed |
| pdf-io-print-export-3 | high | Exporting some pages also embeds the content of non-selected pages that internal links point to | Fixed |
| pdf-io-print-export-4 | high | Encryption check reads only the last 256 KB, so linearized protected PDFs skip unlocking: endless loading, blank printouts, cryptic save errors | Fixed |
| pdf-io-print-export-5 | high | Signature detection only matches the exact text "/Type /Sig", so Acrobat- or iText-signed PDFs are overwritten in place without warning | Fixed |
| pdf-io-print-export-6 | high | Form text longer than a field's MaxLen is silently dropped on Save while the app reports success | Fixed |
| pdf-io-print-export-7 | medium | Saving after any bookmark edit rebuilds the whole outline: headings and URL bookmarks are deleted, children move under the wrong parent, exact positions become… | Fixed |
| pdf-io-print-export-8 | medium | "Rename bookmark" does nothing because window.prompt throws in Electron | Fixed |
| pdf-io-print-export-9 | medium | The same file can be open in several windows, and their saves silently overwrite each other | Open |
| pdf-io-print-export-10 | medium | After a renderer crash or hang, the frameless window can no longer be closed | Fixed |
| pdf-io-print-export-11 | medium | Saving wipes the undo/redo history | Fixed |
| pdf-io-print-export-12 | low | PDFs opened by drag-and-drop lose their file path: Ctrl+S asks where to save, and the file never reaches Recent files | Fixed |

### Save / Import / Export items (32; 19 done)

| ID | Priority | Area | Item | Status |
|---|---|---|---|---|
| PDF-SIE-1 | P0 | save | Moving, resizing, replacing or deleting a native image corrupts other text and inline images on the same page | Done |
| PDF-SIE-2 | P0 | save | Bulletproof atomic write: retry, fsync, short temp name, verified in-place fallback, plain errors and a persistent Retry / Save a copy banner for loc… | Done |
| PDF-SIE-3 | P0 | save | Typed form values are silently dropped (MaxLen overflow, name mismatch, multi-line newlines, per-field errors) while the app reports 'Saved successfu… | Done |
| PDF-SIE-4 | P0 | save | Every save runs pdf-lib form handling: WinAnsi failure blocks any save, XFA deleted, empty AcroForm added, giant multi-line appearances | Done |
| PDF-SIE-5 | P0 | save | Added and edited text: CJK and emoji saved as invisible .notdef, and one over-full box aborts the whole save | Done |
| PDF-SIE-6 | P0 | save | Encrypted PDFs: linearized files skip unlock so Save and page operations fail, user-password files hang on 'Opening…', and encrypted insert sources f… | Done |
| PDF-SIE-7 | P0 | save | Signed PDFs are overwritten in place because detection only matches '/Type /Sig' with a space | Done |
| PDF-SIE-8 | P0 | save | Commit in-progress edits before acting: Esc and Ctrl+Z discard pending text, image and signature edits, Ctrl+S is silently swallowed, and save() can … | Done |
| PDF-SIE-9 | P0 | recovery | Close and replace prompts become Save / Don't Save / Cancel, never close mid-save, and prompt on app quit | Done |
| PDF-SIE-10 | P0 | recovery | Crash recovery journal and handling for renderer crash, hang and Windows shutdown | Open |
| PDF-SIE-11 | P0 | save | Save overwrites changes made on disk, and the same file can be open in two windows | Open |
| PDF-SIE-12 | P0 | import | Built-in Word (DOCX/DOC) import without LibreOffice through every path, including launcher Combine | Done |
| PDF-SIE-13 | P1 | export | Text exports (TXT, MD, HTML, DOCX) lose every line break | Done |
| PDF-SIE-14 | P1 | save | Make the save target obvious: dropped files keep their path, visible Save As / Save a copy, dialogs open in the source folder | Open |
| PDF-SIE-15 | P1 | import | TXT, Markdown and HTML import through Chromium: encoding detection, rendered Markdown, script coverage, no main-process freeze | Done |
| PDF-SIE-16 | P1 | import | One import capability table, magic-byte sniffing, and every image format (including multi-page TIFF) on every import path | Done |
| PDF-SIE-17 | P1 | import | Predictable drag-and-drop with a document open (insert vs open), multi-file open and Combine into one PDF | Open |
| PDF-SIE-18 | P1 | import | Paste from the system clipboard and New PDF from clipboard | Open |
| PDF-SIE-19 | P1 | save | Editing any bookmark rewrites the whole outline: link and heading items are deleted, children re-parented, zoom destinations reset | Done |
| PDF-SIE-20 | P1 | save | Deleted pages and removed images remain recoverable inside the saved file | Done |
| PDF-SIE-21 | P1 | export | Unblock fillable PDFs: flatten option, flattened insert, and page-subset export that keeps document structure | Open |
| PDF-SIE-22 | P1 | save | Fill radio groups, dropdowns and list boxes | Done |
| PDF-SIE-23 | P1 | export | Complete Export As: PDF options (flatten, password, smaller, split), 300/600 DPI, multi-page TIFF, remembered defaults, single entry point | Open |
| PDF-SIE-24 | P1 | export | Export to Word with structure: headings, bold and italic, images, page breaks, then tables | Done |
| PDF-SIE-25 | P1 | import | Spreadsheet, RTF and ODT import without LibreOffice; PPTX is shown as needing an optional engine, never a dead end | Done |
| PDF-SIE-26 | P1 | save | Decide the mupdf AGPL licence question before adding more mupdf-based features | Done |
| PDF-SIE-27 | P2 | import | Repair damaged PDFs on open | Open |
| PDF-SIE-28 | P2 | import | Failed opens remove the file from Recent files | Open |
| PDF-SIE-29 | P2 | save | Large-file save: keep the base in main, progress and cancel, per-page text removal, bounded undo memory | Open |
| PDF-SIE-30 | P2 | save | Incremental (append-only) save for signed documents | Open |
| PDF-SIE-31 | P2 | export | Export to Excel/CSV and PowerPoint | Open |
| PDF-SIE-32 | P3 | export | Preserve file identity on save (ACL, hard links, ADS), OCR searchable PDF, PDF/A | Open |

## Spreadsheets vs Microsoft Excel and Google Sheets

### Capability gaps (52; 25 implemented in this pass)

| ID | Priority | Area | Gap | What the incumbent does | Status |
|---|---|---|---|---|---|
| CALC-001 | P0 | Data / Sort | Quick sort of a one-column selection scrambles rows (no 'Sort Warning: expand the selection?') | Excel: when the selection is one column (or a partial range) next to more data, A-Z / Z-A shows 'Sort Warning - Expand the selection / Continue with the current selectio… | Implemented |
| CALC-003 | P0 | File / Safety | No AutoRecover: a crash or power loss loses all unsaved work | Excel writes an AutoRecover snapshot every 10 minutes and shows a Document Recovery pane on the next start. Sheets saves continuously ('All changes saved'). | Open |
| CALC-002 | P0 | Data / Filter | Copy, clear, format and fill act on filtered-out (hidden) rows | Excel copies only the visible rows of a filtered range. Delete/clear, fill-down and formatting on a filtered list touch visible rows only. Alt+; (Select Visible Cells) d… | Implemented |
| CALC-004 | P0 | Editing / Undo | Undo polish: Ctrl+Shift+Z performs UNDO, selection is not restored, and the file stays dirty at the save point | Ctrl+Y (Excel/Sheets) and Ctrl+Shift+Z (Sheets, common elsewhere) redo. Undo jumps the selection and viewport to the change. Undoing back to the saved state clears the u… | Implemented |
| CALC-006 | P0 | Fill | Double-click on the fill handle does nothing | Excel and Sheets fill the selection down to the last row of the adjacent column's contiguous data (left neighbour first, then right). This is one of the most-used gestur… | Implemented |
| CALC-007 | P0 | Formulas / Authoring | Invalid formulas commit silently as a non-standard #PARSE! error; no hover explanation for error values | Excel blocks the commit with 'There's a problem with this formula' and offers to auto-close missing parentheses or quotes; typing =SUMM() gives #NAME? with a smart tag. … | Implemented |
| CALC-008 | P0 | Print / PDF / HTML export | Conditional formatting, table styles, sparklines and checkboxes are missing from print, PDF and HTML export | Excel and Sheets print and export exactly what the grid shows: CF fills, data bars, icon sets, banded table styles, sparklines and checkboxes. | Implemented |
| CALC-005 | P0 | Links | Internal hyperlinks (#Sheet2!A1, defined names) fail with 'Invalid URL'; Ctrl+K has no 'Place in this document' | Excel: Ctrl+K 'Place in This Document' links to a cell or defined name. Clicking it selects the target, and HYPERLINK("#Sheet2!A1","Go") works. Sheets links to a sheet, … | Implemented |
| CALC-009 | P0 | Charts | Chart editor 'Select range on the sheet' picker is never wired (onPickRange not passed) | Excel's Select Data dialog collapses so you can drag on the sheet, and the source range is outlined in colour and can be resized. Sheets' data range field has a grid ico… | Implemented |
| CALC-011 | P0 | Editing | Alt+Enter line breaks do not turn on Wrap Text, so multi-line entries collapse onto one line | Excel turns on Wrap Text automatically when a committed value contains a line break, and the row auto-grows. Sheets shows the break immediately. | Implemented |
| CALC-010 | P0 | Data / Filter | Sort by colour missing from the column filter menu (onSortByColor never passed) | Excel: Filter dropdown > Sort by Color > cell colour / font colour. Sheets: filter menu > Sort by color > Fill color / Text color. | Implemented |
| CALC-012 | P1 | Locale | Typed input, Ctrl+; and default formats are US-only (m/d/yyyy, '.' decimal, $) | Excel follows Windows regional settings for date order, decimal and list separators, and the currency symbol. Sheets has a per-file Locale (File > Settings) that changes… | Open |
| CALC-013 | P1 | Print | Printing goes silently to the default printer: no printer choice, copies, page range, custom scale % or custom margins | Excel Backstage print: printer picker with status, copies, collated, pages from-to, custom scaling %, Fit to N pages wide by M tall, custom margins and centre on page. S… | Implemented |
| CALC-016 | P1 | Formatting / Scale | Formatting or clearing large selections (full columns or select-all on >100k-cell sheets) fails with 'too large' | Excel formats whole columns, rows or the sheet instantly by storing column, row and sheet-default styles, not per-cell styles. Delete on a full column is instant. Sheets… | Open |
| CALC-017 | P1 | Editing | No AutoComplete from column entries and no Alt+Down 'Pick from drop-down list' | Excel suggests the rest of a text entry from the same contiguous column once the prefix is unique, shown as selected ghost text: Enter accepts, Backspace rejects. Alt+Do… | Implemented |
| CALC-014 | P1 | Page layout | No UI to set print area, print titles, headers/footers or page breaks (imports are honoured but cannot be authored) | Excel Page Layout tab: Set/Clear Print Area, Print Titles (rows and columns), Header/Footer with &[Page] of &[Pages], &[Tab], &[Date] codes, Insert/Reset Page Break, Pag… | Implemented |
| CALC-015 | P1 | Grid & navigation | Finite grid: a new sheet shows only 225 rows, the Name Box silently clamps A5000 to A225, and Ctrl+Down stops at the la… | Excel always exposes 1,048,576 x 16,384 cells: you can scroll or jump anywhere, and Ctrl+Down from an empty column goes to row 1,048,576. Sheets shows 1,000 rows with an… | Implemented |
| CALC-018 | P1 | Data validation | In-cell dropdowns render as an always-visible native <select> pill with 'Select…' placeholder; no chips, colours, searc… | Excel: the dropdown arrow appears only on the selected cell, the value displays with its number format, and the list is searchable (2024). Sheets: Chip, Arrow or Plain s… | Implemented |
| CALC-019 | P1 | Fill | AutoFill lacks named series (days, months, quarters), month/year date steps, the Ctrl-drag toggle and an Auto Fill Opti… | Dragging 'Mon' gives Tue, Wed…; 'January' gives February; 'Q1' gives Q2…Q4, Q1. Two dates a month apart step by month (1/31 then 2/28). Ctrl while dragging switches betw… | Implemented |
| CALC-020 | P1 | Grid & navigation | Enter/Tab collapse a multi-cell selection and do not return to the starting column after a Tab run | With a range selected, Enter, Tab, Shift+Enter and Shift+Tab move the active cell within the selection and wrap, which allows fast data entry into a block. Typing a row … | Implemented |
| CALC-021 | P1 | Keyboard | Excel keyboard staples missing or wrong: Shift+Backspace clears the whole selection; Ctrl+A is not progressive; no Ctrl… | Shift+Backspace collapses the selection to the active cell. Ctrl+A selects the current region first, then the sheet (Ctrl+Shift+8 is the region). Ctrl+Backspace scrolls … | Implemented |
| CALC-022 | P1 | Clipboard | Copy shows no marching-ants marquee (only cut does), so users can't see what is on the clipboard | Both Excel and Sheets show an animated dashed border around the copied range until Esc, a new edit or another copy. Excel's Enter pastes once and ends copy mode. | Implemented |
| CALC-023 | P1 | View | Zoom: no Ctrl+mouse-wheel, max 200%, no Zoom to Selection, and zoom is not saved per sheet | Excel: Ctrl+wheel zoom 10-400%, Zoom to Selection, zoom stored per sheet in the file (sheetView zoomScale) and restored on open. | Implemented |
| CALC-024 | P1 | Sheets | Switching sheets resets the selection to A1 and scroll to top; the saved active cell is ignored on open | Each sheet remembers its own active cell, selection and scroll position. Files store activeCell, topLeftCell and the selection per sheetView, and Excel restores them on … | Implemented |
| CALC-025 | P1 | Formatting | Rich text inside cells is not rendered, and editing a cell discards its runs | Excel and Sheets render mixed fonts, colours, bold and superscript within a cell. Selecting characters in edit mode lets you format part of the text, and edits keep the … | Open |
| CALC-027 | P1 | Formulas / Calculation | No calculation options: no F9 recalc, no manual mode, no iterative calculation, no circular-reference locator | Excel: F9/Shift+F9/Ctrl+Alt+F9; Automatic/Manual mode with 'Calculate' in the status bar; Enable iterative calculation (100 iterations, 0.001). The status bar shows 'Cir… | Implemented |
| CALC-029 | P1 | Clipboard / Editing | No drag-and-drop to move or copy cells by the selection border | Dragging the selection border moves the cells (references follow). Ctrl+drag copies; Shift+drag inserts between cells. 'There's already data here' asks before overwritin… | Implemented |
| CALC-031 | P1 | Menus / Discoverability | No Edit or Help menu, no keyboard-shortcut reference, and the command palette lacks arrow-key navigation and toolbar/co… | Sheets: Edit menu, Help > Keyboard shortcuts (Ctrl+/), Alt+/ menu search with keyboard navigation. Excel: Alt+Q search and KeyTips. | Implemented |
| CALC-046 | P1 | Cell hover / Feedback | Every formula or text cell shows a native tooltip with its formula or text on hover (noise), while errors and truncated… | Excel shows no hover tooltip for normal cells. It shows tips only for notes, hyperlinks (ScreenTip), ##### (the full value) and error smart tags. | Implemented |
| CALC-048 | P1 | Hygiene | Stale backups src/App.tsx.pre-edit-scroll-fix and src/styles.css.pre-edit-scroll-fix are tracked in git | n/a (repository quality) | Open |
| CALC-028 | P2 | Formula auditing | No background error indicators (number stored as text, inconsistent formula) or an error smart-tag menu | Excel's green triangle smart tag offers Convert to Number (for the whole selection), Copy formula from above, Ignore error and Show calculation steps. Sheets shows a red… | Open |
| CALC-030 | P2 | Grid & selection | Only one rectangular selection: no Ctrl+click or Shift+F8 multi-range | Ctrl+click adds ranges for formatting, clearing, charting and status-bar stats (without double-counting). Copy works when the ranges align. Go To Special results are mul… | Open |
| CALC-033 | P2 | Charts | No trendlines or error bars; missing bubble, histogram, box & whisker, waterfall and funnel types | Excel and Sheets trendlines: linear, exponential, logarithmic, polynomial (2-6), power and moving average, with forecast periods, equation and R². Error bars. Modern cha… | Open |
| CALC-035 | P2 | Pivot tables | Pivots lack drill-down (double-click for details), calculated fields, number-range grouping and GETPIVOTDATA | Double-clicking a value creates a new sheet with the source rows. Calculated fields (=Sales-Cost). Number grouping into bins (Excel Group; Sheets histogram grouping). GE… | Open |
| CALC-038 | P2 | Data tools | No Flash Fill (Ctrl+E) or pattern-by-example | Excel infers the transform from 1-2 examples (split names, reformat phones, initials) with a ghost preview; Sheets Smart Fill suggests a formula. | Open |
| CALC-026 | P2 | Comments | Comments are one-shot notes authored 'simple_calc'; notes show only as native title tooltips | Excel and Sheets: threaded comments with author name and time, Reply, Edit, Delete, Resolve and a Comments pane. Notes hover as an anchored yellow box that can be pinned… | Open |
| CALC-032 | P2 | Search | Find & Replace: no wildcards (* ? ~), regex, Find All list, search within selection or in notes | Excel treats * and ? as wildcards by default, has Find All with a clickable result list, and restricts the search to a multi-cell selection. Sheets has 'Search using reg… | Open |
| CALC-034 | P2 | Charts | Chart polish: Switch row/column, empty/hidden cell handling, log/reverse axis, display units, title linked to a cell, c… | Excel: Switch Row/Column; Hidden and Empty Cells (gaps/zero/connect, show hidden data); log scale; values in reverse order; display units (Thousands); =Sheet1!A1 titles;… | Open |
| CALC-036 | P2 | Formulas / Library | Missing functions and operators: IMAGE, FORECAST.ETS*, QUERY (Sheets), GETPIVOTDATA, trim-reference operators (A:.A, .:) | Excel 365 IMAGE(source, alt, sizing) for in-cell pictures; FORECAST.ETS/.CONFINT/.SEASONALITY/.STAT; TRIMRANGE trim refs A:.A; Sheets QUERY with SQL-like select/where/gr… | Open |
| CALC-037 | P2 | Grid & navigation | No Go To Special (blanks, constants, formulas, errors, visible cells, notes, CF, validation) | Ctrl+G > Special selects by type; 'Blanks, =, Up arrow, Ctrl+Enter' is a classic way to fill gaps. Reports 'No cells were found'. | Open |
| CALC-039 | P2 | Formatting | Alternating colours bake fixed green fills into cells; no removable or auto-extending banding | Sheets: Format > Alternating colors with presets, header and footer colours, and a 'Remove alternating colors' option; banding extends as rows are added and survives sor… | Open |
| CALC-040 | P2 | Platform | Spell check is disabled everywhere; no F7 spelling pass | Excel F7 spell check from the active cell, skipping formulas; Sheets Tools > Spelling. Both offer AutoCorrect. | Open |
| CALC-041 | P2 | Formulas / Authoring | Formula bar polish: no fx/✓/✗ buttons, no Insert Function browser (Shift+F3), no matching-parenthesis highlight, range … | Excel formula bar Cancel/Enter/fx; the Insert Function dialog is searchable by category and the argument dialog shows live values. Brackets are colour-matched and flash … | Open |
| CALC-042 | P2 | Status bar | Status bar shows no mode (Ready/Enter/Edit/Point), filter count ('12 of 300 records found') or calculation and circular… | Excel status bar: cell mode, 'X of Y records found' after filtering, 'Circular References: B5', 'Calculate', Caps Lock. | Open |
| CALC-043 | P2 | Sheets | Tab colour limited to 10 hex-labelled swatches; no 'All sheets' list, no Ctrl+drag duplicate, no multi-unhide | Excel: full theme and custom colour picker for tabs, right-click on the tab scroll arrows lists all sheets, Ctrl+drag duplicates as 'Sheet1 (2)', and the Unhide dialog h… | Open |
| CALC-045 | P2 | File | 'Open in new window' is exposed by the bridge but has no UI; only one workbook per window | Excel opens each workbook in its own window; Sheets uses browser tabs. Comparing two workbooks side by side is common. | Open |
| CALC-047 | P2 | Grid & navigation | PageUp/PageDown page by the default row height (constant importedDefaults=true), not the real geometry | PageDown scrolls exactly one visible screen, regardless of custom row heights, wrapped rows, hidden rows or zoom. | Open |
| CALC-049 | P2 | Protection | No workbook-structure protection or allow-edit ranges; Tab does not jump between unlocked cells on a protected sheet | Excel: Protect Workbook (structure) blocks add, delete, rename and move of sheets; Allow Edit Ranges; on a protected sheet Tab moves only between unlocked cells, which t… | Open |
| CALC-052 | P3 | Performance / Architecture | Recalculation runs on the UI thread; heavy workbooks freeze typing and scrolling | Excel uses multi-threaded recalculation; Sheets calculates server-side with progress. Both keep the UI responsive. | Open |
| CALC-044 | P3 | Keyboard / Editing | No F4 'repeat last action' outside formula editing | Excel F4 or Ctrl+Y repeats the last formatting or structural action (e.g. apply fill, insert row) on the new selection. | Open |
| CALC-050 | P3 | What-if | Only Goal Seek: no Data Tables, Scenario Manager or Solver | Excel What-If: Data Table (one and two variable, {=TABLE()}), Scenario Manager, Solver add-in (Simplex LP, GRG Nonlinear). | Open |
| CALC-051 | P3 | Objects | No shapes or text boxes; pictures cannot be cropped or rotated or placed in a cell | Excel and Sheets: shapes, text boxes, Place in Cell images that sort and filter with data, crop and rotate. | Open |

### Verified bugs (36; 36 fixed)

| ID | Severity | Bug | Status |
|---|---|---|---|
| calc-formula-engine-1 | high | COUNTIF/SUMIF(S)/AVERAGEIF(S)/MINIFS/MAXIFS compare date, %, $ and 1,000 criteria as text, and text criteria with < or > also match numbers and blanks | Fixed |
| calc-formula-engine-2 | high | Whole-column and whole-row references (A:A, 1:1) skip spilled dynamic-array cells below the last real cell | Fixed |
| calc-formula-engine-3 | high | Text values that start with "=" are re-run as formulas whenever another formula reads them | Fixed |
| calc-formula-engine-4 | high | SUBTOTAL/AGGREGATE on another sheet goes stale when rows are hidden or filtered, and ISFORMULA never tracks its target | Fixed |
| calc-formula-engine-5 | high | Formulas from older workbooks that rely on implicit intersection spill or show #SPILL!, and saving turns them into array formulas | Fixed |
| calc-formula-engine-6 | high | VLOOKUP/HLOOKUP/MATCH exact match ignore the * ? ~ wildcards, and INDEX/VLOOKUP/HLOOKUP reject fractional index arguments | Fixed |
| calc-formula-engine-7 | medium | Operators don't follow Excel: -2^2 gives -4, 2^3^2 gives 512, =+A1 turns text into #VALUE!, and typed +A1 entries stay text | Fixed |
| calc-formula-engine-8 | medium | DATE() rejects text and blank arguments, and VALUE() rejects date and time text | Fixed |
| calc-formula-engine-9 | medium | Unformatted numbers display with JavaScript float noise and booleans display as lowercase true/false; formatted negative .5 values and 1.005 round differently … | Fixed |
| calc-formula-engine-10 | medium | Text entered with a leading apostrophe turns into a number, date, boolean or formula the next time the cell is edited | Fixed |
| calc-formula-engine-11 | medium | Single-cell references passed to SUM/COUNT/AVERAGE/MAX and similar are treated like typed values, so text numbers and TRUE/FALSE are counted | Fixed |
| calc-formula-engine-12 | low | Date serials 0–60 don't follow Excel's 1900 leap-year rule, so DATE(1900,1,1) returns 2 and YEAR/MONTH/DAY of a blank cell return 1899/12/30 | Fixed |
| calc-file-io-objects-1 | critical | Imported charts get replaced by other charts, or dropped, on the second save because sourcePart is stale after the base is rebased | Fixed |
| calc-file-io-objects-2 | critical | A sheet added after deleting another sheet is saved with the deleted sheet's pictures, password protection, dropdowns, conditional formats, tables and header/f… | Fixed |
| calc-file-io-objects-3 | high | Semicolon CSV with dot decimals: values with exactly three decimals are multiplied by 1000 | Fixed |
| calc-file-io-objects-4 | high | CSV round trip rewrites data: 4-digit-year dates are saved as 2-digit (1949 becomes 2049 on reopen), %, $ and thousands separators are dropped, and the ';'/dec… | Fixed |
| calc-file-io-objects-5 | high | Workbooks with a whole-column or whole-row print area (e.g. $A:$F) cannot be saved or printed | Fixed |
| calc-file-io-objects-6 | high | A double quote inside an unquoted CSV field (e.g. an inch mark in TV 55") swallows delimiters and the following rows | Fixed |
| calc-file-io-objects-7 | high | Protection is quietly removed: legacy sheet passwords are dropped (any password unprotects) and <workbookProtection> is never written back | Fixed |
| calc-file-io-objects-8 | medium | Internal hyperlinks ('Place in This Document', e.g. table-of-contents links) are removed on open and lost on save | Fixed |
| calc-file-io-objects-9 | medium | Deleted pictures stay embedded in the saved .xlsx, and inserted pictures are embedded again on every save | Fixed |
| calc-file-io-objects-10 | medium | Duplicate sheet copies table names unchanged, so Excel has to repair the saved workbook | Fixed |
| calc-file-io-objects-11 | medium | Renaming a sheet breaks every sparkline that reads from it: blank in the editor, stale reference in the saved file | Fixed |
| calc-file-io-objects-12 | low | Saving a brand-new, still-empty workbook fails with 'The original file is unavailable' | Fixed |
| calc-grid-interaction-1 | critical | Paste Values (Ctrl+Shift+V) pastes the values of the wrong cells, counted from A1 | Fixed |
| calc-grid-interaction-2 | high | Delete, Ctrl+D/Ctrl+R, fill handle, Ctrl+Enter and Copy ignore active filters, so hidden filtered rows get overwritten or copied | Fixed |
| calc-grid-interaction-3 | high | Cut mode is never invalidated, so Ctrl+V moves the wrong cells after an insert, delete or sort, or after opening another workbook | Fixed |
| calc-grid-interaction-4 | high | Cutting and pasting to another sheet silently re-targets the formulas inside the moved block | Fixed |
| calc-grid-interaction-5 | high | Inserting or deleting rows/columns does not shift the filter state or filteredRows: rows stay hidden after unfiltering, SUBTOTAL is wrong, and the header row c… | Fixed |
| calc-grid-interaction-6 | high | Committing a cell with an arrow key (Enter mode) skips data validation and records two undo steps | Fixed |
| calc-grid-interaction-7 | medium | Escape (and Ctrl+Enter, Esc in the formula bar, prompt dialogs, menu commands) leaves keyboard focus on <body>, so the grid stops responding to keys | Fixed |
| calc-grid-interaction-8 | medium | Ctrl+D / Ctrl+R turn a date into a day-by-day series and refuse single-row or single-column selections (Excel copies) | Fixed |
| calc-grid-interaction-9 | medium | Picking a value from an in-cell validation dropdown stores the option text as a string, so numbers and dates become text | Fixed |
| calc-grid-interaction-10 | medium | While a cell is being edited the sheet cannot be scrolled, so a formula cannot point at cells outside the starting view | Fixed |
| calc-grid-interaction-11 | medium | With Find open, every workbook change pulls the selection back to a match, and 'All sheets' Next cannot reach other sheets | Fixed |
| calc-grid-interaction-12 | medium | The name box and Go To (F5, Ctrl+G) clamp targets to the rendered area (rowCount+25), so 'A1000' or 'A2:A5000' selects the wrong range | Fixed |

### Save / Import / Export items (30; 18 done)

| ID | Priority | Area | Item | Status |
|---|---|---|---|---|
| CALC-SIE-1 | P0 | save | Shared (filled-down) formulas: editing the master corrupts or blocks every XLSX save | Done |
| CALC-SIE-2 | P0 | save | Office-engine capability probe; edited .xls/.ods saves an .xlsx beside the original when LibreOffice is absent | Done |
| CALC-SIE-3 | P0 | save | Ctrl+S on a brand-new, unedited workbook fails after the file name is chosen | Done |
| CALC-SIE-4 | P0 | save | Unsaved-changes prompt: Save / Don't Save / Cancel, with Save as the default (today Enter discards) | Done |
| CALC-SIE-5 | P0 | save | Commit every pending input (cell, formula bar, side panels) with validation before save, close or export | Done |
| CALC-SIE-6 | P0 | save | Lossy-format check runs against the format actually chosen; a reduced-format save keeps a 'not fully saved' state | Open |
| CALC-SIE-7 | P0 | save | CSV/TSV writer: RFC 4180 quoting (embedded newlines corrupt rows) and 4-digit years | Done |
| CALC-SIE-8 | P0 | save | atomicWrite: retry on transient locks, fsync, refuse read-only targets, coded errors | Done |
| CALC-SIE-9 | P0 | save | Save/export failures shown in a persistent dialog with Retry / Save As XLSX / Show folder | Open |
| CALC-SIE-10 | P0 | recovery | AutoRecover journal, crash/hang handling, a close path that survives a dead renderer | Open |
| CALC-SIE-11 | P0 | save | Worksheet identity matching: new or duplicated sheets inherit a deleted sheet's protection, tables, validation and pictures; removed filters come back | Done |
| CALC-SIE-12 | P0 | save | Validation edits and row/column shifts on sheets with more than 10,000 validated cells are silently not saved | Done |
| CALC-SIE-13 | P1 | import | Open .ods (and fix .fods) without LibreOffice through SheetJS plus post-processing | Done |
| CALC-SIE-14 | P1 | export | XLS/ODS export without LibreOffice: native SheetJS writers with an explicit loss summary, or disabled | Done |
| CALC-SIE-15 | P1 | save | Closing during an in-flight save says 'discard' and can quit before the write finishes | Done |
| CALC-SIE-16 | P1 | save | 'Changed outside Simple': content-hash check, Overwrite/Save a copy/Reload, export no longer blocked | Open |
| CALC-SIE-17 | P1 | save | Source moved, renamed or deleted: raw ENOENT error; a clean document cannot be recreated | Open |
| CALC-SIE-18 | P1 | save | CSV Save preserves the source dialect (delimiter, decimal comma, encoding, BOM, line ending) | Done |
| CALC-SIE-19 | P1 | import | CSV import inference: per-column decimal and date order, EU dates and grouping, BOM-less UTF-16, encoding detection | Done |
| CALC-SIE-20 | P1 | import | Detect the real format from content; never show an empty workbook for a non-empty file | Done |
| CALC-SIE-21 | P1 | import | One File > Import… command (new sheet / replace / append / at cell / new workbook) with text options and preview | Open |
| CALC-SIE-22 | P1 | import | Drag and drop keeps the real path; multi-file drop; open in a new window | Open |
| CALC-SIE-23 | P1 | export | Complete Export As: scope, values mode, delimiter and encoding, every-sheet ZIP, JSON, Markdown, XML 2003, HTML data table | Open |
| CALC-SIE-24 | P1 | save | Duplicated sheets produce duplicate table names and ids; Excel repairs the file | Done |
| CALC-SIE-25 | P1 | save | Small save-fidelity losses: hyperlink on a formula cell, checkbox becomes a dropdown, sparkline ranges after a rename | Done |
| CALC-SIE-26 | P1 | save | QA suite aware of the office engine: test the LibreOffice-absent fallbacks and add save/import regression tests | Done |
| CALC-SIE-27 | P2 | import | Paste of comma/semicolon text offers 'Split into columns' | Open |
| CALC-SIE-28 | P2 | export | Set or clear the print area and repeat-title rows from the selection | Open |
| CALC-SIE-29 | P3 | import | Delimited .txt in the unified launcher reaches Spreadsheets | Open |
| CALC-SIE-30 | P3 | import | Open password-protected workbooks | Open |

## Documents vs Microsoft Word and Google Docs

### Capability gaps (47; 19 implemented in this pass)

| ID | Priority | Area | Gap | What the incumbent does | Status |
|---|---|---|---|---|---|
| DOC-001 | P0 | File & app shell / data safety | Dirty state is guessed from DOM events, so many edits never trigger the unsaved prompt or a recovery copy | Word and Docs mark the document as modified on every model change, whichever control made it: the mini toolbar, Replace All, ruler drags, object handles, Ctrl+Enter. The… | Implemented |
| DOC-002 | P0 | Review / Save | Save, Export, Print and recovery silently throw away pending suggestions and all comments while reporting 'Saved' | Word stores tracked changes (w:ins/w:del) and comments in the .docx. Docs keeps them and exports them to .docx. Neither app discards review content silently, and Word wa… | Implemented |
| DOC-003 | P0 | Import fidelity | Opening a DOCX with tracked changes or comments silently accepts insertions, drops deletions and ignores comments | Word and Docs show the existing markup (tracked changes in colour, comments in the margin) and let the user accept or reject it. Nothing is accepted on the user's behalf. | Implemented |
| DOC-004 | P0 | Crash recovery | Saving recovered work writes 'name.doc.docx' for .doc sources and overwrites the original without the external-change c… | Word's Document Recovery reopens the draft, keeps the original format, and asks where to save or warns if the original changed on disk. | Implemented |
| DOC-005 | P0 | Insert / Context menu / View | Five engine commands use window.prompt(), which throws in Electron 43: context-menu Insert/Edit Hyperlink, bookmark add… | Word: right-click > Link opens the hyperlink dialog, and Insert > Bookmark opens the bookmark dialog. Docs: Ctrl+K opens a link card. All of these work in place with no … | Implemented |
| DOC-006 | P0 | Text editing | Ctrl+Backspace / Ctrl+Delete (delete previous/next word) do nothing | Both apps delete the previous or next word as a single undo step and leave no doubled space. With a selection active, these keys act like plain Backspace/Delete. | Implemented |
| DOC-007 | P0 | Home / History | No visible Undo/Redo buttons: the engine's Undo group lives in the File tab that Simple removes | Word puts Undo/Redo at the left of Home (with a history dropdown). Docs shows them as the first toolbar buttons. Both are always visible. | Implemented |
| DOC-008 | P1 | Clipboard / Insert image | Images cannot be pasted (screenshots, browser copies) or dragged in from Explorer; dropping an image shows an error | Pressing Ctrl+V with a screenshot inserts it inline, scaled to the text width. Dragging an image file drops it at the pointer. EXIF orientation is respected. | Implemented |
| DOC-009 | P1 | Proofing | No spell check: no red squiggles, no suggestions, no Add to dictionary | Both apps underline misspellings as you type and offer suggestions, Ignore All and Add to Dictionary. They ignore URLs, ALL CAPS and words with digits, flag repeated wor… | Implemented |
| DOC-010 | P1 | Keyboard | Many Word/Docs-standard shortcuts are missing (Ctrl+K, Ctrl+H, Ctrl+L/E/R/J, Ctrl+]/[, Ctrl+=, Ctrl+1/2/5, Ctrl+Alt+1-3… | Muscle-memory shortcuts work everywhere in the document and are advertised in tooltips. Ctrl+Alt shortcuts do not fire for AltGr characters on European layouts. | Implemented |
| DOC-011 | P1 | Insert / Headers & footers | No Header, Footer or Page number commands; users must discover double-click on the band or the Insert Field context menu | Word: Insert > Header / Footer / Page Number (Top, Bottom, or Page X of Y, with format and start-at). Docs: Insert > Headers & footers and Page numbers, with a skip-firs… | Implemented |
| DOC-012 | P1 | Clipboard | No paste options: no Ctrl+Shift+V plain-text paste, no Keep text only, and pasting a URL over a selection replaces the … | Docs: Ctrl+Shift+V pastes without formatting. Word: the Paste Options button offers Keep Source, Merge or Keep Text Only. Both link the selected text when a URL is paste… | Implemented |
| DOC-013 | P1 | AutoCorrect / AutoFormat | No AutoFormat as you type: no smart quotes, dashes, auto-lists ('1. ', '- '), URL auto-link or sentence capitalization | Word and Docs convert quotes (locale-aware), dashes, (c), ordinals, list triggers and URLs as you type. Ctrl+Z immediately after a conversion reverts only the correction… | Implemented |
| DOC-014 | P1 | Paragraph formatting | No numeric per-paragraph indents (left, right, first-line, hanging) or Space Before/After fields; no 'Add space before/… | Word Paragraph dialog > Indents and Spacing (Left, Right, Special First line/Hanging, Before/After in pt), plus Add/Remove Space Before/After in the line-spacing menu. D… | Implemented |
| DOC-015 | P1 | UI / Discoverability | No command search ('Tell me' / Alt+Q / Docs 'Search the menus') | Word (Alt+Q) and Docs (Alt+/) find any command by name and run it, listing recently used actions first. | Implemented |
| DOC-017 | P1 | File formats (import) | Only .docx and .doc open; .dotx/.docm/.dotm, .rtf, .odt, .txt, .md and .html are rejected | Word opens docx/docm/dotx/dotm, doc, rtf, odt, txt and htm; opening a template creates an untitled copy. Docs imports docx, odt, rtf, txt, html and md. | Implemented |
| DOC-018 | P1 | Typography / Fonts | Only 15 curated installed fonts are offered; other installed fonts fall back to clones, and there is no font search or … | Word lists every installed font in its own face, with Recently Used and Theme fonts at the top and type-ahead. A missing font keeps its name and is substituted visibly. | Implemented |
| DOC-016 | P1 | Review | Author identity is hard-coded as 'Local Author' for comments and suggestions | Word uses the Office user name and initials and asks for a name on the first comment if none is set. Docs uses the account name. | Implemented |
| DOC-019 | P2 | Find & replace | Find/replace lacks Ctrl+H, selection prefill, regex/wildcards, a single-undo Replace All, and searching headers and not… | Word: Ctrl+H, selection prefill, Replace All with a count as one undo step, wildcards and ^p/^t, Results pane snippets, and headers/notes searched. Docs: regex with $1. | Open |
| DOC-020 | P2 | Review / DOCX fidelity | Comments are session-only: not read from DOCX and not written to DOCX | Word round-trips comments.xml, commentsExtended.xml (resolved state and replies) with author and date. Docs exports and imports comments in .docx. | Implemented |
| DOC-021 | P2 | Review / DOCX fidelity | Tracked changes are not round-tripped: suggestions are not saved as w:ins/w:del, and imported revisions are flattened | Word and Docs keep tracked insertions, deletions and formatting changes in the file across reviewers. | Open |
| DOC-022 | P2 | File & app shell | No version history and no optional autosave to the real file | Docs autosaves and offers version history with named versions and restore. Word AutoSave works for cloud files, AutoRecover for local files, and Version History. | Open |
| DOC-026 | P2 | Lists | No Restart / Continue numbering or Set numbering value; no checklist | Word: Restart at 1, Continue Numbering, Set Numbering Value. Docs: Restart / Continue, plus a checklist (Ctrl+Shift+9) with strikethrough when checked. | Open |
| DOC-029 | P2 | Tables | Table gaps: no Insert Table dialog beyond 10x8, no text-to-table conversion, no sort, no split cells, no distribute, no… | Word: an Insert Table dialog of any size, Convert Text to Table and back, Sort with header row, Split Cells, Distribute Rows/Columns, and double-click a border to autofi… | Open |
| DOC-031 | P2 | Images / Cross-workspace | No image adjustments and no 'Edit in Simple Images' handoff for advanced picture editing | Word: Corrections, Color, Transparency, Remove Background, Picture Styles, and Change Picture keeping size and position. Docs: Image options for recolour, transparency, … | Open |
| DOC-032 | P2 | File formats (import) / Cross-workspace | No 'Open PDF as editable document' (Word PDF Reflow / Docs OCR conversion) | Word converts a PDF into an editable document, warning that layout may change. Docs converts PDFs and images with OCR. Acrobat Pro exports to Word with high fidelity. | Open |
| DOC-036 | P2 | Text editing / Caret | Caret and selection polish: no Ctrl+Up/Down paragraph jumps, no PgUp/PgDn, no drag-to-move text, no smart spacing, no s… | Ctrl+Up/Down moves by paragraph and PgUp/PgDn by screen. Dragging a selection moves it (Ctrl copies). Word fixes spaces on cut and paste, and left-margin clicks select a… | Open |
| DOC-037 | P2 | UI / Minimalism | The Insert tab is cluttered with about 15 always-visible, mostly disabled shape and image controls, while Comment, Date… | Word shows object formatting only on contextual tabs. Insert keeps a short set (Table, Pictures, Shapes, Link, Comment, Header/Footer/Page Number, Symbol, Equation, Date… | Open |
| DOC-023 | P2 | File & app shell | No templates: the welcome screen offers only Blank document | Word and Docs start screens show resume, letter, report, meeting notes and memo templates. New from template opens an untitled copy. | Open |
| DOC-024 | P2 | File & app shell | Recent files are reachable only from the welcome card; no pin, no remove, no jump list | Word: File > Open > Recent at any time, with pin, remove, open file location and a taskbar jump list. | Open |
| DOC-025 | P2 | Status bar / Proofing | No selection word count and no Word Count dialog | Word's status bar shows 'N of M words' for a selection; clicking it opens full statistics (characters with and without spaces, paragraphs, lines, include notes). Docs: C… | Open |
| DOC-027 | P2 | Paragraph formatting / Ruler | No tab-stop UI (set, clear, leader dots); the ruler lacks hanging and right-indent markers | Word: a tab selector at the ruler corner (left, center, right, decimal, bar), click to set, drag off to clear, and a Tabs dialog with leaders. The ruler has first-line, … | Open |
| DOC-028 | P2 | Layout / Sections | Only 'Section break: next page' is available from the ribbon; continuous, even-page and odd-page breaks are missing | Word Layout > Breaks: Page, Column, Text Wrapping, and section breaks Next Page, Continuous, Even and Odd. Docs: next page or continuous. | Open |
| DOC-030 | P2 | Images / Accessibility | No alt text for images (not in the model, not exported) | Word and Docs: Alt text (Docs Ctrl+Alt+Y) with 'Mark as decorative', written to wp:docPr descr, HTML alt and PDF tags. | Open |
| DOC-033 | P2 | Units / Dialogs | Dialogs show px instead of pt/in/cm, and font size is capped at 72 pt | Word shows font sizes in pt (1-1638, half points) and measurements in inches or cm by locale. | Open |
| DOC-034 | P2 | Theme | No dark mode | Word's Black theme keeps a white page by default and offers a 'Switch Modes' toggle that darkens the page for display only. Apps follow the system theme. | Open |
| DOC-035 | P2 | Export | Export lacks ODT/RTF, a .doc choice in Save As, and PDF options (heading bookmarks, metadata, page range) | Word Save As covers docx, doc, dotx, pdf, odt, rtf, txt and htm, with PDF options for range, bookmarks from headings, properties, tagged PDF and PDF/A. Docs downloads do… | Open |
| DOC-038 | P3 | Design | No watermark command (DRAFT, CONFIDENTIAL, custom text or image) | Word Design > Watermark (presets, custom text or picture with washout). Docs Insert > Watermark. Both live in the header layer. | Open |
| DOC-039 | P3 | References | No captions, cross-references or table of figures | Word: Insert Caption with SEQ fields, Cross-reference with REF/PAGEREF, and Table of Figures. | Open |
| DOC-040 | P3 | Navigation pane | The outline pane cannot reorder sections by dragging headings, collapse levels, or promote/demote | Word Navigation pane: drag a heading to move its section, promote/demote, collapse. Docs outline: collapse. | Open |
| DOC-041 | P3 | Mailings / Cross-workspace | No mail merge or labels; envelopes only | Word Mailings: choose recipients from Excel/CSV, insert merge fields, preview, and finish to a document or print. Labels use Avery presets. | Open |
| DOC-042 | P3 | Printing | Print lacks current page, selection, odd/even only, pages per sheet and a print-markup option | Word Print: Current Page, Selection, custom ranges, odd or even only, 1-16 pages per sheet, Print Markup and manual duplex. | Open |
| DOC-043 | P3 | View / Zoom | Editing zoom lacks Page width / One page presets, and zoom is not remembered per document | Word: One Page, Multiple Pages and Page Width, with zoom remembered. Docs: Fit. | Open |
| DOC-044 | P3 | Mailings | Envelope replaces the whole document and wipes undo history; it cannot be added to the current letter | Word Envelopes > 'Add to Document' inserts the envelope as a first section in front of the letter, and the action can be undone. | Open |
| DOC-046 | P3 | Images / Layout | Picture layout lacks tight, through and top-and-bottom wrap, a rotate handle, a Layout Options button and alignment gui… | Word: a Layout Options button, seven wrap modes, a rotation handle with 15-degree snapping, and smart guides. Docs: Inline, Wrap, Break text, Behind and In front. | Open |
| DOC-045 | P3 | File / Properties | No document properties (title, author, subject, keywords) | Word File > Info edits docProps/core.xml, which carries into PDF metadata; Inspect Document removes hidden data. | Open |
| DOC-047 | P3 | References | No citations or bibliography manager | Word Source Manager (APA, MLA, Chicago) with a generated bibliography; Docs Citations sidebar. | Open |

### Verified bugs (24; 24 fixed)

| ID | Severity | Bug | Status |
|---|---|---|---|
| doc-io-fidelity-1 | critical | Saving an edited DOCX silently deletes charts, EMF/WMF/TIFF pictures and OLE objects that the editor skipped on import, with no warning and no backup | Fixed |
| doc-io-fidelity-2 | high | Find & Replace, Ctrl+Enter and mouse edits never mark the document dirty, so closing discards them without a prompt | Fixed |
| doc-io-fidelity-3 | high | Edits made from the right-click menu (Delete picture, Remove Hyperlink, Update Field, table and equation commands) are never detected as changes | Fixed |
| doc-io-fidelity-4 | medium | Restored recovery copies are never cleared and save over the source file without the external-change check or the original backup | Fixed |
| doc-io-fidelity-5 | medium | Save silently overwrites read-only documents and strips the read-only attribute | Fixed |
| doc-io-fidelity-6 | medium | If the Office page-view PDF fails, Print and PDF export of the unchanged document fail permanently | Fixed |
| doc-io-fidelity-7 | medium | Open failures are swallowed: double-clicked, Ctrl+O and dropped files show no error, and Recent removes valid files | Fixed |
| doc-io-fidelity-8 | medium | Installed-font metrics ignore the hhea line gap, making Calibri and Consolas lines about 18% tighter than in Word | Fixed |
| doc-io-fidelity-9 | medium | Markdown export puts whitespace inside ** and _ markers and emits **a****b**, so asterisks show literally | Fixed |
| doc-io-fidelity-10 | medium | Ctrl+S, Ctrl+P, Ctrl+O, Ctrl+N and Ctrl+Shift+S/E do nothing with Cyrillic, Greek, Hebrew or Arabic keyboard layouts | Fixed |
| doc-io-fidelity-11 | low | Export and Save-as default names are cut at a dot ('J. Smith CV' → 'J.pdf') and ignore the document's folder | Fixed |
| doc-io-fidelity-12 | low | Print 'Document title' and 'Page numbers' marks are drawn at the paper edge with the default None margins and get clipped | Fixed |
| doc-editor-integration-1 | critical | Edits made through WordCanvas's floating mini-toolbar, Find/Replace bar, context toolbars and ruler never mark the document dirty, so closing the window silent… | Fixed |
| doc-editor-integration-2 | high | Import failure is reported as "Document opened." and a later Save overwrites the real file with a blank document | Fixed |
| doc-editor-integration-3 | high | The right-click menu edit detector is dead code: WordCanvas runs menu commands on mouseup and removes the menu, so no click ever reaches it | Fixed |
| doc-editor-integration-4 | high | Saving recovered work writes straight to the original path with no external-change check, and a recovered .doc is saved silently as "name.doc.docx" | Fixed |
| doc-editor-integration-5 | medium | The editor never gets keyboard focus after Blank document, Open, or closing a dialog; typed characters are lost while the document is marked Unsaved | Fixed |
| doc-editor-integration-6 | medium | Tab in a normal paragraph moves keyboard focus out of the document (to the status-bar "Zoom out" button) instead of inserting a tab | Fixed |
| doc-editor-integration-7 | medium | Ctrl+Enter / Ctrl+Shift+Enter breaks and arrow-key nudges of a selected picture are not tracked as edits | Fixed |
| doc-editor-integration-8 | medium | Errors when opening a file are silently swallowed (Open button / Ctrl+O, drag-and-drop, Explorer double-click at startup) | Fixed |
| doc-editor-integration-9 | medium | Page view leaves WordCanvas's Find bar and toolbars active over the PDF; Ctrl+F there searches and replaces in the hidden editor model | Fixed |
| doc-editor-integration-10 | medium | The Page view zoom select (Fit page / Fit width / 100% / 150%) does nothing once the page view has loaded | Fixed |
| doc-editor-integration-11 | medium | Restored "Recovered work" entries are never cleared, so the same stale copy is offered on every launch | Fixed |
| doc-editor-integration-12 | low | Actions that do not edit the document mark it Unsaved (View-tab toggles, ribbon collapse, keys pressed outside the editor), causing false save prompts | Fixed |

### Save / Import / Export items (27; 15 done)

| ID | Priority | Area | Item | Status |
|---|---|---|---|---|
| DOC-SIE-1 | P0 | save | Derive dirty state from the model, not DOM-event guesses | Open |
| DOC-SIE-2 | P0 | save | Suggesting-mode edits and comments are discarded while the app reports 'Saved.' | Open |
| DOC-SIE-3 | P0 | save | Scan for lossy content at open, surface import warnings, and protect the first save | Done |
| DOC-SIE-4 | P0 | save | Keep a version backup before the first overwrite, independent of LibreOffice | Open |
| DOC-SIE-5 | P0 | import | A failed WordCanvas import binds the file path, and a later save overwrites it with the blank model | Done |
| DOC-SIE-6 | P0 | save | Robust file replace for locked and OneDrive files: retry, flush, verified rollback, coded plain-language errors | Done |
| DOC-SIE-7 | P0 | recovery | Restoring recovered work overwrites the source with no conflict check, and the entry never clears | Done |
| DOC-SIE-8 | P1 | save | Read-only files are silently overwritten and lose their attributes | Done |
| DOC-SIE-9 | P1 | save | Post-write bookkeeping failures report a completed save as failed and cause a false conflict next time | Open |
| DOC-SIE-10 | P1 | recovery | Recovery robustness: model-triggered snapshots, a safe index, visible failures, crash and shutdown hooks | Open |
| DOC-SIE-11 | P1 | save | Save silently does nothing during export or print, and exports never time out | Open |
| DOC-SIE-12 | P1 | import | One open helper with visible, specific errors on every open path | Open |
| DOC-SIE-13 | P1 | import | Dropping a path-less file replaces the open document and its recovery snapshot | Done |
| DOC-SIE-14 | P1 | import | Legacy .doc without LibreOffice: source-folder save, visible warnings, RTF/HTML sniffing, one-click full-formatting option | Done |
| DOC-SIE-15 | P1 | import | Accept .docm/.dotx/.dotm, add an 'All supported documents' filter, and use one shared format list for unified routing | Done |
| DOC-SIE-16 | P1 | import | Native importers for TXT, Markdown, HTML, RTF and ODT via one DOM-to-model mapper | Done |
| DOC-SIE-17 | P1 | export | Normalize inserted and imported images to PNG/JPEG so PDF and print include them and DOCX media is labeled correctly | Done |
| DOC-SIE-18 | P1 | export | Exports report their warnings, and HTML/MD skip what they cannot include instead of failing | Done |
| DOC-SIE-19 | P1 | export | Multi-section documents lose the first section's header and footer on an edited save | Done |
| DOC-SIE-20 | P2 | save | Save and Export dialogs: source folder, correct names, safe extension handling, show the saved path | Open |
| DOC-SIE-21 | P2 | save | Truthful save feedback: guard against open dialogs, persistent errors, status for untitled documents | Open |
| DOC-SIE-22 | P2 | export | One Export As dialog with PDF range and metadata options; remove the dead file:save-pdf IPC | Open |
| DOC-SIE-23 | P2 | export | Clean GFM Markdown export and working HTML bookmark anchors | Done |
| DOC-SIE-24 | P2 | export | Native ODT, RTF and EPUB export; offer .doc only when the office runtime is present | Done |
| DOC-SIE-25 | P2 | import | Rich paste and drop into the document: lists, tables, headings, images, plain-text paste | Done |
| DOC-SIE-26 | P2 | save | Merge-on-save: carry package parts and style definitions the model does not own | Open |
| DOC-SIE-27 | P3 | save | Round-trip comments and tracked changes for real | Open |

## Images vs Adobe Photoshop (Advanced) and Windows Photos (Simple)

### Capability gaps (61; 37 implemented in this pass)

| ID | Priority | Area | Gap | What the incumbent does | Status |
|---|---|---|---|---|---|
| IMAGE-016 | P0 | Simple: adjust | No light/colour adjustments (exposure, contrast, highlights, shadows, saturation, vibrance, warmth, tint, vignette, sha… | Windows Photos Adjustment pane and Preview Adjust Color: sliders centred at 0, live preview, per-section reset, double-click reset, histogram; edits re-editable until sa… | Implemented (Simple-mode upgrade) |
| IMAGE-050 | P0 | Advanced: shell | One-click Simple / Advanced mode switch that carries edits across | Photoshop Elements Quick/Guided/Expert tabs; Windows Photos Edit / Edit more. Simple edits arrive as editable state; returning to Simple with several layers warns and of… | Implemented (Advanced mode) |
| IMAGE-051 | P0 | Advanced: engine | Layered rendering engine (WebGL2 tiled compositor) to replace the single canvas | Photoshop/Photopea: GPU compositing, tile cache, mip levels, interactive brushes and adjustments on 50 MP+ documents. | Implemented (Advanced mode (CPU tiled compositor)) |
| IMAGE-052 | P0 | Advanced: layers | Layers panel: add, delete, duplicate, reorder, rename, visibility, opacity, lock, groups, merge/flatten/stamp | Photoshop Ctrl+Shift+N, Ctrl+J, Ctrl+E, Ctrl+Shift+E, Ctrl+Alt+Shift+E stamp, Ctrl+G, drag-reorder, Alt-click eye solo, '/' lock transparency, Background layer, number k… | Implemented (Advanced mode) |
| IMAGE-054 | P0 | Advanced: selections | Selections: marquee, lasso/polygonal, magic wand, add/subtract/intersect, select all/deselect/reselect/invert, feather/… | Photoshop M/L/W with Shift/Alt modifiers, Space to move mid-drag, marching ants, Ctrl+A/D/Shift+D/Shift+I, Modify menu, wand tolerance 32 + contiguous; Preview Instant A… | Implemented (Advanced mode) |
| IMAGE-055 | P0 | Advanced: transform | Move tool and Free Transform (scale, rotate, skew, perspective) with numeric fields and smart guides | Photoshop V move with nudges, Alt-drag duplicate, auto-select; Ctrl+T proportional by default, Shift 15 deg rotation snap, Ctrl-drag distort, Enter/Esc, Transform Again;… | Implemented (Advanced mode) |
| IMAGE-003 | P0 | Simple: view & zoom | Ctrl+wheel zoom is not cursor-anchored and its preventDefault is ignored (passive React wheel listener) | Photoshop, Photopea and Windows Photos zoom around the cursor with Ctrl/Alt+wheel and pinch; the app chrome never zooms. Plain wheel scrolls, Shift+wheel scrolls horizon… | Implemented (Simple-mode upgrade) |
| IMAGE-007 | P0 | Simple: history | Undo stores full bitmaps, budget is computed from current size, redo is unbounded | Photoshop keeps 50 states by default, spills to scratch disk, and undo cost is proportional to the changed area. Windows Photos/Preview keep the original plus parameters. | Implemented (Simple-mode upgrade) |
| IMAGE-009 | P0 | Simple: view & zoom | No Space-drag / middle-drag / hand-tool panning | Photoshop and Photopea: hold Space for the temporary Hand tool from any tool, middle-drag pans, double-click Hand fits. | Implemented (Simple-mode upgrade) |
| IMAGE-011 | P0 | Simple: open & import | No paste (Ctrl+V) and no New from clipboard | Photoshop: Ctrl+N pre-fills the clipboard size, Ctrl+V pastes as a new layer centred in view. Photopea/Paint: paste a screenshot straight into an empty app. | Implemented |
| IMAGE-013 | P0 | Simple: transform | No Resize image dialog (pixels / percent / presets, aspect lock, quality filtering, size estimate) | Windows Photos Resize: aspect chain, pixels or %, quality slider, estimated file size. Preview Adjust Size: Fit-into presets, before/after KB. Photoshop Image Size (Ctrl… | Implemented (Simple-mode upgrade) |
| IMAGE-014 | P0 | Simple: transform | Crop lacks aspect presets, Shift-constrain, orientation swap and numeric size | Windows Photos and Photoshop: Free/Original/1:1/4:3/3:2/16:9/9:16 presets, X swaps orientation, Shift locks ratio, Alt from centre, W x H crop to exact output (1080x1080… | Implemented (Simple-mode upgrade) |
| IMAGE-020 | P0 | Simple: export | Export has fixed 0.92 quality, no size estimate, no resize, and silently flattens transparency to white for JPEG | Photoshop Export As / Photopea: quality slider with live file size, scale, metadata None/All, Convert to sRGB; JPEG matte colour; Windows Photos Save options show qualit… | Open |
| IMAGE-001 | P0 | Simple: bugs / paint | Brush or eraser dragged off the canvas paints a stripe along the border | Photoshop, Photopea and Windows Photos Markup continue the stroke off-canvas invisibly; nothing is painted on the edge, and the stroke resumes where the pointer re-enter… | Implemented |
| IMAGE-002 | P0 | Simple: bugs / crop | Rotate, undo or redo during crop leaves the Crop tool active with no rectangle | Photoshop disables other commands (or asks Apply/Don't Apply) while a crop is pending; Windows Photos re-fits the crop frame to the rotated image. The crop UI never sile… | Implemented |
| IMAGE-004 | P0 | Simple: view & zoom | Pixels are blurred when zoomed in (no nearest-neighbour rendering) | Photoshop, Photopea and Windows Photos render nearest-neighbour above 100% so individual pixels are crisp; Photoshop adds a pixel grid above about 500%. | Implemented (Simple-mode upgrade) |
| IMAGE-010 | P0 | Simple: keyboard | No single-key tool shortcuts, [ ] brush size, Ctrl+1 actual size, rotate/flip keys | Photoshop: V/C/B/E/I/H/Z tools, [ ] size, Ctrl+1 100%, Ctrl+0 fit, D/X colours. Windows Photos: Ctrl+R rotate, F11 full screen, Del delete. Shortcuts never fire while ty… | Partial ([ ] brush size, Ctrl+1 and tool keys in Advanced; not every Simple-mode key) |
| IMAGE-012 | P0 | Simple: transform | No flip horizontal / vertical | Windows Photos, Preview Flip Horizontal/Vertical, Photoshop Flip Canvas; all undoable. | Implemented (Simple-mode upgrade) |
| IMAGE-017 | P1 | Simple: adjust | No auto-enhance / auto levels | Windows Photos Auto-enhance with intensity; Preview Auto Levels; Photoshop Auto Tone/Color (0.1% clip). Auto sets visible slider values. | Implemented (Simple-mode upgrade) |
| IMAGE-022 | P1 | Simple: export | EXIF, ICC profile and DPI are dropped on every re-encode | Photoshop and Windows Photos keep EXIF (date, camera, GPS) and colour profile when saving an edited JPEG unless the user strips metadata on export. | Implemented |
| IMAGE-024 | P1 | Simple: paint & markup | No markup objects: text, arrow, line, rectangle, ellipse, callout | Preview and Windows Photos Markup: shapes/arrows with handles, text boxes with font/size/colour, editable until saved, Shift constrains, last-used style remembered; Loup… | Implemented (Simple-mode upgrade (markup)) |
| IMAGE-026 | P1 | Simple: smart tools | No 'Copy text from image' (OCR) | Windows Photos 'Scan text' and Snipping Tool Text Actions; Acrobat Pro recognises text in images and makes it searchable/editable. | Open |
| IMAGE-029 | P1 | Simple: open & import | Missing common formats: HEIC/HEIF, TIFF, ICO, PSD composite; GIF edits drop animation silently; SVG rasterised tiny | Windows Photos opens HEIC, TIFF, ICO, RAW; Photopea opens PSD/TIFF/HEIC; Photoshop warns before flattening animation; SVG imported at a chosen size. | Partial (ICO, PSD, APNG, SVGZ, JFIF added; HEIC/TIFF still missing) |
| IMAGE-053 | P1 | Advanced: layers | Blend modes matching Photoshop formulas | 27 modes incl. Photoshop Soft Light, Vivid/Linear Light, Hard Mix, Divide, non-separable Hue/Saturation/Color/Luminosity; hover preview; Shift+Alt+letter shortcuts. | Implemented (Advanced mode) |
| IMAGE-056 | P1 | Advanced: masks | Layer masks and clipping masks | Add mask from selection, Alt hide-all, Shift-click disable, Alt-click view mask, '\' rubylith, paint target by thumbnail, Ctrl+I invert, Ctrl+Alt+G clip, Apply/Delete ma… | Implemented (Advanced mode) |
| IMAGE-057 | P1 | Advanced: adjustments | Adjustment layers with Properties: Levels, Curves, Hue/Saturation, Black & White, Color Balance, Exposure, Photo Filter… | Photoshop adjustment layers re-editable in Properties with clip, view-previous, reset; Ctrl+L/M/U; monotone Curves with targeted hand; Levels with Alt clipping preview a… | Implemented (Advanced mode) |
| IMAGE-058 | P1 | Advanced: filters | Core filters with preview conventions: Gaussian Blur, Unsharp Mask, Add/Reduce Noise, Median, Mosaic, Motion Blur, High… | Photoshop filter dialogs with live preview, hold for before, Alt Reset, Ctrl+F/Ctrl+Alt+F repeat, Edit > Fade, cancellable progress; Smart Filters. | Implemented (Advanced mode) |
| IMAGE-059 | P1 | Advanced: text | Editable text layers (point and paragraph) with a font picker | Photoshop T tool: click point text, drag paragraph box, Ctrl+Enter/Esc, layer named from text, font menu with previews and recent fonts, size/leading/tracking, alignment… | Implemented (Advanced mode) |
| IMAGE-061 | P1 | Advanced: paint | Full brush engine: hardness, opacity vs flow, spacing, pressure, smoothing, Shift-line, FG/BG colours, pencil, gradient… | Photoshop B/E/G: [ ] size, Shift+[ ] hardness, Alt-right-drag HUD, opacity vs flow, pressure, smoothing, Shift-click line, D/X, dithered gradients, bucket tolerance, Alt… | Implemented (Advanced mode) |
| IMAGE-064 | P1 | Advanced: files | Layered file format: open/save PSD (with composite fallback), Save a Copy for flat formats, Export layers | Photoshop/Photopea round-trip PSD with groups, masks, blend modes, text and adjustments; flat formats via Save a Copy; 'Some layers could not be read' warning; per-layer… | Implemented (Advanced mode (PSD)) |
| IMAGE-006 | P1 | Simple: performance | Full transparency rescan after every stroke and crop lags on large images | Photoshop and Photopea never block after a stroke; alpha is tracked per layer, not rediscovered. | Implemented |
| IMAGE-008 | P1 | Simple: paint | Fast strokes become polylines; cursor shows no brush size | Photoshop/Photopea use all coalesced input with smoothing and show the brush outline at true size with a crosshair inside. | Implemented |
| IMAGE-015 | P1 | Simple: transform | No straighten / free rotate with auto-crop | Windows Photos straighten slider (about +/-45 deg) with fine grid; Photoshop crop-tool rotate and Straighten line; auto-crop to largest inscribed rectangle; double-click… | Implemented (Simple-mode upgrade) |
| IMAGE-018 | P1 | Simple: adjust | No one-click looks/filters with live thumbnails and intensity | Windows Photos filter strip from the current photo, Original first, intensity slider; Photoshop 2025 Adjustment Presets with hover preview. | Implemented (Simple-mode upgrade) |
| IMAGE-019 | P1 | Simple: adjust | No before/after compare or Revert to original | Windows Photos hold-to-compare and Revert to original; Camera Raw '\' toggle and Y split view. | Implemented (Simple-mode upgrade) |
| IMAGE-021 | P1 | Simple: export | Saving after an edit uses fixed JPEG quality and there is no 'Save as copy' | Windows Photos Save split button with 'Save as copy' (name (1).jpg) as the safe option; Photoshop Save a Copy (Ctrl+Alt+S). | Partial (JPEG quality follows the source; no Save as copy) |
| IMAGE-023 | P1 | Simple: paint & markup | No eyedropper, swatches, brush opacity or highlighter | Preview/Windows Photos Markup: pen, pencil, highlighter with palette; Photoshop I eyedropper with sampling ring, Alt-click while painting, opacity/hardness. | Partial (eyedropper and brush polish; no swatches/highlighter) |
| IMAGE-025 | P1 | Simple: paint & markup | No redact (pixelate / blur / solid box) region tool | Screenshot editors and Photoshop Mosaic on a selection hide faces, plates and personal data quickly. | Open |
| IMAGE-063 | P1 | Advanced: history | History panel with named states, snapshots and Revert | Photoshop History: named steps, click to jump, snapshots, 50 states, F12 Revert, Ctrl+Alt+Z; undo covers selections and layer properties. | Implemented (Advanced mode) |
| IMAGE-065 | P1 | Advanced: shell | Command palette and contextual task bar to keep the Advanced UI minimal | Photoshop Search (Ctrl+F) finds tools/panels/commands; Contextual Task Bar suggests next actions. | Open |
| IMAGE-005 | P1 | Simple: browse | Album list goes stale after Save As or file changes; only the first dropped file opens | Windows Photos re-reads the folder and follows a renamed/saved-as file; dropping several files opens them all (Photoshop opens each as a document). | Open |
| IMAGE-030 | P1 | Simple: export | PDF export always embeds lossless PNG (huge photo PDFs) and has no page options | Microsoft Print to PDF, Preview and Acrobat embed JPEG for photos and offer page size/fit. | Open |
| IMAGE-036 | P1 | Simple: quality | Renderer has no tests; three copies of the format list | n/a (engineering quality); incumbents never disagree about which files they open. | Implemented |
| IMAGE-027 | P2 | Simple: smart tools | No background removal (on-device) | Windows Photos Background Remove/Blur/Replace on-device with refine brush; Preview Instant Alpha / Remove Background; Photoshop Remove Background creates a mask. | Open |
| IMAGE-028 | P2 | Simple: smart tools | No spot fix / object erase (inpainting) | Windows Photos Erase (auto-apply on release); Photoshop Remove tool and Spot Healing; Photopea Spot Healing via PatchMatch. | Open |
| IMAGE-031 | P2 | Simple: browse | Viewer polish missing: Delete to Recycle Bin, rename, Home/End, filmstrip, full screen, slideshow | Windows Photos: Del to Recycle Bin with confirm, F2 rename, filmstrip, F11, slideshow, Home/End, Ctrl+B set as background. | Open |
| IMAGE-033 | P2 | Simple: shell | No recent files, no crash recovery, no dirty marker in OS title | Photoshop recent files and auto-recovery every 10 min; Windows Photos recent; '*' in title for unsaved changes. | Open |
| IMAGE-060 | P2 | Advanced: vector | Shape layers (rectangle, rounded rect, ellipse, line/arrow, polygon/star) with fill and stroke | Photoshop U tools with live properties (W/H/X/Y, per-corner radius, stroke width/align/dash), Shift constrain, click for numeric size. | Open |
| IMAGE-062 | P2 | Advanced: retouch | Retouching: Clone Stamp, Healing Brush, Spot Healing, Content-Aware Fill, Red Eye, Dodge/Burn | Photoshop S clone with source overlay, J healing family, Shift+Backspace content-aware fill, O dodge/burn/sponge, 'Sample All Layers' for non-destructive retouch. | Open |
| IMAGE-068 | P2 | Advanced: layers | Smart Objects / Place Embedded (drop a file onto an open document) | Photoshop places a dropped file as a Smart Object with transform box; scaling is lossless; Smart Filters stay editable. | Open |
| IMAGE-069 | P2 | Advanced: transform | Perspective crop / perspective correction for photographed documents | Photoshop Perspective Crop tool turns four dragged corners into a rectangle; popular for receipts and documents. | Open |
| IMAGE-032 | P2 | Simple: shell | Context menu offers only Copy/Undo/Redo/Apply crop | Windows Photos right-click: rotate, copy, open with, set as, show in folder, print, info, delete; Photoshop menus change with context. | Open |
| IMAGE-034 | P2 | Simple: view & zoom | Zoom polish: editable % field, standard ladder, max 800%, pixel grid | Photoshop zoom ladder up to 3200%, editable zoom field, pixel grid above 500%. | Open |
| IMAGE-038 | P2 | Simple: print | Print always goes silently to the default printer; no photo paper sizes | Windows Photos print: printer picker, 4x6/5x7/wallet/full page; Photoshop printer selection. | Open |
| IMAGE-066 | P2 | Advanced: view | Rulers, guides, grid, snapping, Info readout, Navigator | Photoshop Ctrl+R rulers, drag guides, Ctrl+' grid, snapping in screen px, F8 Info, Navigator. | Open |
| IMAGE-067 | P2 | Advanced: layers | Layer styles: Drop Shadow, Stroke, Outer Glow, Color/Gradient Overlay | Photoshop Layer Style dialog, fx row, drag shadow on canvas, copy/paste style, Opacity vs Fill. | Open |
| IMAGE-070 | P3 | Advanced: AI | Select Subject / click-to-select object (on-device) | Photoshop Select Subject, Object Selection hover highlight, Select Sky/People. | Open |
| IMAGE-035 | P3 | Simple: shell | No dark theme | Photoshop defaults to a dark UI; Windows Photos follows the system theme. | Open |
| IMAGE-037 | P3 | Simple: smart tools | Batch resize / convert (PowerToys Image Resizer style) | PowerToys Image Resizer presets and format conversion; Photoshop Image Processor. | Open |
| IMAGE-071 | P3 | Advanced: vector | Pen tool, paths panel, warp/puppet/liquify | Photoshop P pen with anchor/handle editing and paths to selection; Warp, Puppet Warp, Liquify. | Open |
| IMAGE-072 | P3 | Advanced: colour | Colour management and high bit depth (P3/AdobeRGB sources, 16-bit, CMYK) | Photoshop converts embedded profiles to the working space, warns on mismatch, supports 16/32-bit and CMYK; Export converts to sRGB by default. | Open |

### Verified bugs (24; 23 fixed)

| ID | Severity | Bug | Status |
|---|---|---|---|
| image-editing-correctness-1 | critical | Ctrl+Z or Ctrl+S while the mouse is still down mid-stroke desyncs the revision: the app says Saved and Save writes the original bytes | Fixed |
| image-editing-correctness-2 | high | All Ctrl+letter shortcuts silently do nothing with a Cyrillic (Russian/Uzbek Cyrillic) or other non-Latin keyboard layout | Fixed |
| image-editing-correctness-3 | high | Brush and eraser strokes that leave the image paint (or erase) a stripe along the image border | Fixed |
| image-editing-correctness-4 | medium | The "Drop to open" overlay gets stuck over the image when a drag leaves the window without dropping | Fixed |
| image-editing-correctness-5 | medium | A single click outside the crop box replaces the selection with a 1x1 px crop; Enter or Apply then crops the image to one pixel | Fixed |
| image-editing-correctness-6 | medium | Enter in crop mode ignores the focused control: Enter on the crop Cancel button applies the crop, and Enter on the Crop button crops and immediately opens a ne… | Fixed |
| image-editing-correctness-7 | medium | Crop handles clamp the pointer instead of the edge, so dragging an edge to the image border leaves a sliver or moves the opposite edge | Open |
| image-editing-correctness-8 | medium | Album arrow-key browsing gets stuck on an unreadable sibling, and the counter never advances past it | Fixed |
| image-editing-correctness-9 | medium | Print dialog: a typed custom scale is clamped on every keystroke, so typing 50% gives 100% and typing 150% gives 400% | Fixed |
| image-editing-correctness-10 | low | Esc, Enter or Ctrl+Z during a crop drag leaves a stale drag, and the next crop box follows the mouse with no button pressed | Fixed |
| image-editing-correctness-11 | low | Ctrl+Z, Ctrl+Y or Rotate while cropping discards the crop box, reverts an earlier pixel edit, and leaves the Crop tool active with no box | Fixed |
| image-editing-correctness-12 | low | Ctrl+A selects the inspector text, after which Ctrl+C copies UI text instead of the image | Fixed |
| image-interaction-ux-1 | high | Print dialog Copies / Custom-scale fields reset on every keystroke (Backspace then '2' prints 12 copies straight to the printer) and the wrong values are saved… | Fixed |
| image-interaction-ux-2 | high | 'Save changes?' dialog never takes focus: Enter, Space and Tab work the toolbar behind it (e.g. rotate again), and Save then writes that unintended edit over t… | Fixed |
| image-interaction-ux-3 | high | 'Drop to open' overlay stays on screen permanently after a drag leaves the window or is cancelled | Fixed |
| image-interaction-ux-4 | medium | Brush and eraser strokes that go off the image paint (or erase) a stripe along the image edge | Fixed |
| image-interaction-ux-5 | medium | Print and save error messages appear underneath the print dialog and the Save prompt, so failures look like nothing happened | Fixed |
| image-interaction-ux-6 | medium | Alt+F4 toggles Image details and is swallowed instead of closing the window | Fixed |
| image-interaction-ux-7 | medium | Crop keys ignore focus and open menus: Enter on a focused 'Cancel' (or any control) applies the crop, and Escape to close a menu also throws away the crop | Fixed |
| image-interaction-ux-8 | medium | One click outside the crop box replaces the selection with a 1x1 px crop | Fixed |
| image-interaction-ux-9 | medium | Next/Previous cannot get past a file that fails to open; it retries the same file every time (and each step's toast covers the image counter) | Fixed |
| image-interaction-ux-10 | medium | Zoom ignores where you are looking: every zoom jumps to the top-left, and Ctrl+wheel / touchpad pinch zooms 12% per event regardless of how far you scroll | Fixed |
| image-interaction-ux-11 | low | '1:1' / 100% uses CSS pixels instead of screen pixels, and zoomed-in pixels are blurred: images look soft on scaled displays and pixel-level editing is hard to… | Fixed |
| image-interaction-ux-12 | low | Rotate, Undo or Redo while Crop is active removes the crop box but leaves the Crop tool selected and unresponsive | Fixed |

### Save / Import / Export items (26; 15 done)

| ID | Priority | Area | Item | Status |
|---|---|---|---|---|
| IMAGE-SIE-1 | P0 | save | Ctrl+Z or Ctrl+S during a brush/eraser stroke desyncs the revision: the app says Saved and Save writes the original bytes | Done |
| IMAGE-SIE-2 | P0 | save | A pending, unapplied crop is ignored by Save and by the close prompt; no dirty marker in the window title | Done |
| IMAGE-SIE-3 | P0 | save | Animated APNG and WebP files are overwritten in place with a single frame and no warning | Done |
| IMAGE-SIE-4 | P0 | save | Failed saves are invisible (the toast sits under the close prompt) and show raw IPC errors with no Retry or Save As | Done |
| IMAGE-SIE-5 | P0 | save | Save fails immediately with raw EBUSY/EPERM when another program, antivirus or OneDrive holds the file: no retry, no fallback | Done |
| IMAGE-SIE-6 | P0 | recovery | After a renderer crash or hang the frameless window cannot be closed and quit is blocked; a close during print is dropped | Done |
| IMAGE-SIE-7 | P0 | recovery | No autosave or crash recovery: a crash, update restart, sign-out or power loss loses all unsaved edits | Open |
| IMAGE-SIE-8 | P0 | save | Save, Save As and Export dialogs ignore the image's folder; GIF/BMP/SVG/AVIF edits land in an arbitrary folder as a new PNG | Open |
| IMAGE-SIE-9 | P1 | save | Rewrite atomicWrite: unsafe rollback, a backup cleanup that can revert a good save, leaked partial temps, no fsync, 210+ character names unsaveable | Done |
| IMAGE-SIE-10 | P1 | save | Read-only images are silently overwritten and lose the R flag; hidden attributes and hard links are broken | Done |
| IMAGE-SIE-11 | P1 | save | Quit or close during an in-flight save or export can cut the write short | Done |
| IMAGE-SIE-12 | P1 | save | The unsaved-changes prompt leaves focus on the toolbar, so Enter rotates the image instead of saving | Done |
| IMAGE-SIE-13 | P1 | save | No external-change or OneDrive-sync detection: Save overwrites files changed, replaced or renamed since open | Open |
| IMAGE-SIE-14 | P1 | save | An edited lossless WebP is saved lossy at q0.92; erased areas in a JPEG silently turn white | Done |
| IMAGE-SIE-15 | P1 | save | An edited Save or Export strips EXIF, XMP, IPTC, DPI and PNG text, and reduces 16-bit sources to 8-bit silently | Done |
| IMAGE-SIE-16 | P1 | import | Valid images are refused: mismatched content (WebP/PNG saved as .jpg) and the .jfif/.jpe/.jif/.ico/.apng/.svgz extensions | Done |
| IMAGE-SIE-17 | P1 | import | No clipboard paste, multi-file drop opens only the first file, browser drags do nothing, and there is no single obvious import entry | Open |
| IMAGE-SIE-18 | P1 | import | Illustrator SVGs (DOCTYPE with an internal subset) and UTF-16 SVGs are rejected; viewBox-only SVGs rasterize tiny | Done |
| IMAGE-SIE-19 | P1 | export | Save As cannot change format; typing photo.jpg writes photo.jpg.png; a rewritten extension can overwrite a file without confirmation | Open |
| IMAGE-SIE-20 | P1 | export | Single Export As dialog: format, quality or lossless, resize, metadata, PDF page, preview and size estimate | Open |
| IMAGE-SIE-21 | P1 | export | PDF export blocks the main process for seconds, embeds photos as PNG (5x bloat) and has no page options | Open |
| IMAGE-SIE-22 | P2 | import | Photos over 50 MP or 20,000 px cannot be opened even for viewing | Open |
| IMAGE-SIE-23 | P2 | import | TIFF, HEIC/HEIF, PSD, TGA and RAW/DNG are rejected with no conversion path | Open |
| IMAGE-SIE-24 | P2 | export | No BMP, GIF, TIFF or ICO writers; BMP and TIFF cannot be saved in place | Open |
| IMAGE-SIE-25 | P3 | import | Relative command-line paths from a second instance resolve against the wrong directory | Done |
| IMAGE-SIE-26 | P3 | import | Animated GIF, WebP and APNG are not played in view mode | Open |

## Launcher and cross-workspace consistency

| ID | Severity | Bug | Status |
|---|---|---|---|
| launcher-launcher-consistency-1 | high | Close/replace prompts in PDF and Calc only offer Discard, and Enter discards | Fixed |
| launcher-launcher-consistency-2 | high | Same file opens in two windows, and PDF/Image saves then silently overwrite newer content | Fixed |
| launcher-launcher-consistency-3 | medium | Windows restart/sign-out bypasses close protection; unsaved PDF, Calc and Image edits are lost silently | Fixed |
| launcher-launcher-consistency-4 | medium | Dropping a file onto PDF or Calc opens a path-less copy: Ctrl+S becomes Save As elsewhere, with no Recent entry | Fixed |
| launcher-launcher-consistency-5 | medium | Launcher multi-file open silently loses files beyond the portable wrapper's 8K command-line limit | Fixed |
| launcher-launcher-consistency-6 | medium | A crashed or hung editor renderer leaves a window that cannot be closed | Fixed |
| launcher-launcher-consistency-7 | medium | Combine breaks internal links and can copy unselected pages into the output | Fixed |
| launcher-launcher-consistency-8 | medium | Combine refuses owner-password PDFs that the PDF workspace opens without a prompt | Fixed |
| launcher-launcher-consistency-9 | low | launchDetached ignores spawn errors: false 'Opened' status plus an uncaught main-process exception | Fixed |
| launcher-launcher-consistency-10 | low | Save As and Export dialogs in PDF, Docs and Image start in an unrelated folder | Fixed |
| launcher-launcher-consistency-11 | low | Recent files lose entries when several PDF or Calc windows are open | Fixed |
| launcher-launcher-consistency-12 | low | Internal names and placeholder authors leak into window titles, prompts and saved files | Fixed |
