# Document reliability review

The supplied reference is a binary Word 97–2003 `.doc` file, not a native Google Doc. The previous fallback rebuilt its extracted text as DOCX; that removed all four tables, the watermark image and their page positions. Local Office conversion now retains these structures, while Page view supplies a compatible Office-rendered view and unchanged-document print/PDF output. The sample remains one page there, and its historical June 10, 2026 date stays unchanged. The date was independently checked against the actual Google-exported reference PDF.

A concrete editor defect was repaired: a 720 × 720 px image marked `behind: true` had reserved 720 px of header space. The pinned, reproducible adapter now places anchored header images in the correct page layer without reserving body space, preserves their original header ownership and centered positioning, and retains source luminance settings through DOCX saves. The sample's editing view now has one page instead of five. Inline headers and repeated/first-page header selection have regression coverage.

There are still material compatibility limits. Floating table positions (`w:tblpPr`) do not survive into the editable table model, and canvas/PDF washout is an opacity approximation. Against the actual Google-exported PDF, the local Office page view still differs in the company/customer block positions and header shading; some text overlaps in that local render. Page view is therefore explicitly labeled as a compatible rendering, not an exact Word or Google Docs representation.

## Implemented safeguards

- Original-format Save: unchanged DOC/DOCX bytes are reused; edited DOC is converted to a real binary Word file before any replacement.
- Original backups for edited complex files and a source hash check against changes made outside Simple.
- Installed Office font faces shared by canvas layout and PDF export, using the font's ascent/descent rather than generic substitute metrics.
- Print preview and physical submission use one composed PDF. Job callbacks are bounded, and crashed renderers and unconfirmed jobs yield actionable errors without automatic retries.
- Cached PDF exports check document model identity and review content; navigation alone does not change the exported representation. The Office cache is keyed by source content, engine path/version timestamp, format/filter and conversion-policy version, with signature and content-hash checks on disk reads.
- Page-view and edit-view switches retain the document and undo history. After a real edit, print/PDF export includes that edit.
- Legacy DATE/TIME fields with cached results are locked only in the temporary conversion input; PAGE, REF and other fields keep their behavior.

## Remaining priorities

1. Improve floating-table geometry, gradient/shading fidelity and exact image washout; expand original/edited/reference render comparisons covering sections, footnotes, RTL and CJK text.
2. Preserve unsupported OOXML parts through edited saves, with compatibility accounting instead of an importer that silently omits structures.
3. Reduce first-time Office startup with a managed, isolated conversion worker. The actual sample's first open took 13.970 seconds; a separate fresh app launch using the new disk cache took 1.161 seconds. Print setup took 155–156 ms and its repeat 140–146 ms. These are local sample measurements, not guarantees for all documents.
4. Test real device/driver combinations using a controlled printer corpus. Automated checks verify composition and readiness without consuming paper and cannot prove a physical printer's output.
5. Keep the signed Office runtime independently updateable and visible in diagnostics; do not silently require a machine-wide Office install or redistribute Windows fonts.

## Synthetic creation stress pass (September 2026)

The actual Electron editor created a 1,249-word, six-page document with two populated tables (13 cells, one long wrapped cell), two bullet items, mixed Latin fonts and accents, Arabic/Hebrew text, Chinese/Japanese/Korean text, a portrait-to-landscape section break, a picker-inserted picture changed to Square wrapping, a manual line break and a footnote. Generated files and screenshots live only in ignored `tmp/editor-stress`.

Fixed defects found by creation and independent reopening:

1. New table typing started outside the table. The same insertion transaction now selects its first cell, and Tab visits the subsequent cells.
2. Pasted CSS RGB colors were invalid in the Word/PDF writers. They now normalize to the model's RGB hex color format.
3. HTML and Markdown dropped object-URL images. Export resolves those local bytes into an isolated document copy; missing local images fail with an explanation instead of a successful incomplete export.
4. Markdown omitted notes and headers, and text/HTML omitted endnotes and section header variants. They now retain those stories with labels; HTML also carries section sizes and explicit page breaks. These formats preserve content but do not promise Word pagination or repeating page furniture.
5. Shift+Enter wrote invalid XML into DOCX and was expanded into separate paragraphs on import. It now writes `w:br` and remains a line break inside its paragraph. The corrected file also converts successfully through the independent local Office engine to binary DOC and back, retaining table, footnote and mixed-script markers.
6. Square-wrapped pictures reverted to inline on DOCX save. They now preserve standard square-wrap anchoring. Supported margin-aligned square pictures reopen directly in the editor; complex/free-positioned/header anchors retain compatible Office Page view.
7. Footnotes in text beside square images disappeared from PDF. That flow path now reserves and emits the notes on the correct page.
8. The bundled SC font lacked some Japanese and all tested Korean glyphs. Coverage-based fallback uses an installed, embeddable Malgun face for missing characters while keeping simplified Chinese in SC. PDF text extraction and visual inspection confirm the full test phrase. Systems without a suitable supplemental font still have a coverage limitation.

The confirmed fresh document window remains six pages, retains all compared text and saves unchanged DOCX bytes exactly. Pages 2–6 match pixel for pixel. Page 1 changes 1,056 of 484,704 pixels at 72 dpi (0.218%), with a maximum text-position shift of 0.03125 pt from OOXML unit rounding. Full-page and enlarged crop checks show readable, aligned cell text. Independent strict XML parsing passes; the blue pasted text, full mixed-script phrase and footnote are present in the reopened PDF. `python scripts/verify-editor-stress.py` reproduces this check with PyMuPDF and Pillow. Physical printers were not exercised.

Remaining authoring priorities: discoverable header/footer/page-number creation from a blank document, nested table insertion, HTML clipboard tables/images/lists (the current paste parser flattens complex structure), broad CJK font availability, and a larger corpus for long/splitting notes, arbitrary floating tables and complex multi-section layouts. Existing header/footer/page-number fixtures are covered by separate layout tests; this UI stress case does not claim full authoring coverage for those features.

## Primary references

- Microsoft documents how automatic line spacing uses 240ths of a line, while exact spacing uses point-based rules: [SpacingBetweenLines.Line](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.spacingbetweenlines.line).
- Microsoft describes paragraph indentation and spacing behavior: [Adjust indents and spacing](https://support.microsoft.com/en-US/Word/adjust-indents-and-spacing).
- LibreOffice documents its conversion filters, headless execution and profile selection: [Starting LibreOffice with parameters](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html).
- Electron documents native printer options, print callbacks and PDF behavior: [webContents](https://www.electronjs.org/docs/latest/api/web-contents).
- The editor's current public model, import and font contracts are available in [WordCanvas](https://github.com/Forevka/canvas-word).
- Microsoft defines the legacy field lock flag used by the conversion-copy safeguard: [MS-DOC grffldEnd](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/28ab752b-055a-4725-8797-159bba0d125c).
