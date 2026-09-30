# Simple PDF and Acrobat Pro

Reviewed September 5, 2026 against the current Simple PDF source and official Adobe documentation. “This change” means this implementation pass; proposed work below is not implemented. No Acrobat-versus-Simple timing benchmark was run.

Simple already handles everyday reading, page organization and small visual edits. This change improves editing consistency and removes avoidable interaction costs. The largest remaining editing gap is the underlying content model: text replacement is drawn over the source, which can remain extractable. It is not a general native text rewrite engine or secure redaction.

## Existing capabilities and changes in this pass

| Area | Already in the app | This change |
| --- | --- | --- |
| Text editing | Direct caret/substring editing; add, move and resize text; font, size, weight, style, color, alignment and character spacing; source font and baseline preservation where available. | Shared fit/wrap rules for display and saved output; explicit text fitting and line spacing; the font field shows the source face; deliberate font/size/style changes bypass source-metric calibration that could override the requested formatting. |
| Edit workflow | Image selection, move/resize/rotate/flip/replace/duplicate/delete; artwork-region capture; copy/paste; properties panel. | Clear **Edit PDF** entry, duplication and selection nudge shortcuts. Arrow keys move a selection by 1 point, or 10 with Shift, when focus is outside a text field. Fixed the invisible text-layer background blocking real image clicks while retaining text priority over images. |
| Viewing | Continuous pages with nearby canvas rendering, virtualized thumbnails, selectable text, links, basic form fields and fit-to-window. | Fit width, actual size and explicit percentage zoom; toolbar layout adapts to narrower windows. |
| Search | Whole-document search, page results, occurrence highlighting, Enter/Shift+Enter navigation and shared PDF.js text extraction. | Reuses normalized page text between queries; abandons obsolete searches; yields between batches so input and scrolling can continue. |
| Document workflows | Native outline bookmarks, page reorder/duplicate/delete/crop/insert, dragging page selections out as PDFs, separate document windows, recent files, print preview and export. | These remain existing capabilities, rather than new features attributed to this pass. |

Implementation evidence: [editor](src/components/PdfPage.tsx), [properties](src/components/EditInspector.tsx), [app workflows](src/App.tsx), [toolbar](src/components/Toolbar.tsx), [page virtualization](src/components/ContinuousPdfViewer.tsx), [search](src/lib/documentSearch.ts) and [save logic](electron/main.cjs).

## Where Acrobat remains ahead

| Workflow | Adobe's documented behavior | Simple's remaining boundary |
| --- | --- | --- |
| Paragraph editing | Reflows text within the selected text box and offers paragraph/list formatting. Independent boxes do not push one another down or flow across pages. [Adobe: edit text](https://helpx.adobe.com/acrobat/using/edit-text-pdfs1.html) | Fit/wrap applies to the edited region. Robust paragraph/list reconstruction, mixed formatting and native text replacement remain future work. Full document reflow should not be promised as Acrobat parity. |
| Font handling | Provides formatting and fallback-font controls; availability of installed and embedded fonts affects editing. [Adobe: edit text](https://helpx.adobe.com/acrobat/using/edit-text-pdfs1.html) | Source-font retention and matching improve appearance, but unusual subsets, missing glyphs and complex scripts still require broader coverage and verification. |
| Image/object editing | Includes image crop, replacement, transforms, multi-object alignment and stacking order. [Adobe: modify images](https://helpx.adobe.com/acrobat/desktop/edit-documents/edit-images-or-objects/modify-image.html) | Basic raster transforms exist. Multi-object alignment/ordering and editing individual vector paths do not. Artwork regions become movable images; supported image paint removal is not a general vector-editing engine. |
| Scanned pages | Applies OCR to create an editable copy of a scan. [Adobe: edit scanned documents](https://helpx.adobe.com/acrobat/desktop/create-documents/scan-documents-to-pdfs/edit-scans.html) | No OCR. An image-only scan has no recognized text to select or edit. |
| Sensitive content | Offers applied redaction and optional removal of hidden information. [Adobe: redact and sanitize](https://helpx.adobe.com/acrobat/desktop/protect-documents/redact-pdfs/redact.html) | Covering or deleting visible text does not remove the underlying source. Simple has no secure redaction or sanitization workflow. |

Search currently reads the base PDF text rather than logical text edits. Saving keeps that base and the edits in memory, so a search in the same session can still match the original wording. Reopened saved output may contain both covered source text and replacement text. Text-based exports are best-effort conversions, not full document reconstruction. These limits follow from the current save/search model, not from the new search cache.

## Measured search work

Editing validation also covers real mouse and keyboard input, toolbar visibility at 1100px, fit/actual/manual zoom, 24pt text after saving, eight wrapped lines at 22pt spacing, overflow handling, and 1pt/10pt image movement. Existing editor regressions pass on normal and rotated pages. Live caret editing uses browser soft wrapping, so uncommon punctuation, complex scripts and fallback fonts can still differ slightly from saved rendering. The shared fitting rules reduce these differences without replacing the PDF engine.

`node scripts/test-document-search.mjs` passed. These are deterministic indexing call counts and responsiveness checks, not elapsed-time or Acrobat comparisons:

| Check | Result |
| --- | --- |
| First query across 256 synthetic pages | 256 page-text indexing reads; 512 occurrences in page order. |
| Different query on the same PDF identity | 0 additional page-text indexing reads; cached normalized text reused. |
| Fully cached 256-page scan | 7 UI yields with a fixed test clock, one between each 32-page batch. |
| Cancel while the first of 100 page reads is pending | 1 page read; the remaining 99 skipped. A later query completes successfully. |
| Real timer cancels a 5,000-page scan | Stops after 32 page reads in the recorded run. |

The tests also cover Unicode normalization, non-overlapping occurrence counts, distinct PDF identities and retry after extraction failure. The production scheduler yields after 32 pages or approximately 8 milliseconds between page operations; a single complex page can take longer. TypeScript checking also passed. The existing [search UI smoke](scripts/smoke-search-selection.mjs) covers visible highlight geometry and result navigation.

## Practical next priorities

The subsequent Books stress pass copied and inventoried 27 local PDFs, then exercised eight text edits and one scan annotation through the real editor and Save flow. It fixed malformed CFF subset headers that made saved Berling/Myriad faces fall back in other readers, used original PDF paint colors to avoid antialiasing tints, and measured glyph ink bounds so old descenders do not remain under replacements. Ambiguous or transparent paint runs retain the existing sampled-color fallback. All eight edited-page comparisons stayed within the source line area; every other page's decoded content streams stayed unchanged. Source Books files remained byte-identical. This is sampled visual evidence, not a claim of universal or Acrobat-equivalent editing. See the [complex-file review](../simple/COMPLEX_FILE_REVIEW.md).

The 400% long-page case also exposed a ineffective bitmap cap: the one-pixel backing-scale floor allowed 41,028,736 pixels. The cap now permits a backing scale below one when needed, limits both area and edge length, and releases the staging bitmap after rendering. That case now uses 23,997,924 pixels. Screen memory is bounded without changing saved vector content. Scan editing now explains when a page contains no editable text; added notes save correctly, but OCR and invisible retouching of scanned words remain unimplemented.

Follow-up reliability work completed image extraction on selection, including caching and a real 30-object test with zero encodes on entering Edit. It also corrected JPEG EXIF orientation during PDF conversion and bounded print-submission failures. The broader [reliability review](../simple/RELIABILITY_REVIEW.md) records those results.

| Rank | Change and reason | How to verify the benefit |
| --- | --- | --- |
| 1 | **Move PDF save/serialization into a worker.** Mutation and flattening handlers perform PDF processing in Electron's main process, which also owns application windows. [Source](electron/main.cjs) | Measure main-process event-loop delay and save completion on large PDFs; verify saved bytes, cancellation/failure handling and window responsiveness. |
| 2 | **Keep undo history across Save.** Saving currently clears undo/redo. Preserve history with a saved-state checkpoint so undo restores the correct dirty indicator. [Source](src/App.tsx) | Edit, save, undo, redo and save again; check content, dirty state and bounded memory. |
| 3 | **Restore reading position.** Recent files exist, but reopening resets to page 1 and fit-to-window. Persist page, zoom mode and offset per document identity. [Source](src/App.tsx) | Reopen a long and a mixed-size PDF; restore the same view and safely clamp changed page counts. |
| 4 | **Build paragraph understanding and a native content-editing engine.** This is the largest Acrobat-like editing gain and the largest engineering effort. Start with reliable paragraph selection, then replace source text while preserving font resources, layout and extraction semantics. [Current save model](electron/main.cjs) | Compare renderings and extracted text after editing/reopening; cover subsets, rotations, multiline paragraphs, RTL, tables and malformed input. |

Other useful follow-ups are search over current logical edits, incremental search results, multi-object alignment, recovery after a crash and interoperable comment annotations. OCR and secure redaction require separate implementations and dedicated validation; they are not completed by drawing replacements over a page.
