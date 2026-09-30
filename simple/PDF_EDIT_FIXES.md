# PDF editing, saving, and export fixes — September 15, 2026

## Changes

- Native text edits remove the selected letters from the PDF content instead
  of covering them with a sampled background color. Photos, vector artwork,
  colors, neighboring text, hyperlinks, and annotations remain in place.
- The live editor uses the same removal engine as saved files. Selecting,
  moving, deleting, and committing text no longer adds a solid background patch.
- Valid native text whose baseline falls below its selection box can now be
  saved and exported. Actual overflowing multiline content is still reported.
- Embedded-font subsets are checked on a disposable document before they are
  registered in the output PDF, allowing font fallback before serialization.
- Very large pages are scaled within canvas limits for PNG, JPEG, and WebP.
- The unified build includes the text-removal engine and its WebAssembly runtime.

## Verification

- 50 PDF automated tests passed; TypeScript validation passed.
- 24 unified-app tests passed, including staged-source verification.
- Pixel comparisons passed for text over patterned vectors and images, cropped
  pages, all four page rotations, and shared nested PDF content.
- Link preservation and existing pending redaction annotations were checked.
- Nine book UI cases saved and reopened successfully: Napoleon, Garamond,
  a colored cover, Berling, mixed styles, Cyrillic, a tall page, Calibri, and
  an added note on a scanned page. All 27 source book files remained byte-identical.
- The real editor passed formatting, wrapping, cancellation, moving images,
  and save checks, including the packaged application loaded from its archive.
- Source and staged app integration passed repeated saves, saving while editing,
  re-editing after save, canceled Save As, and exports to PDF, PNG, JPEG, WebP,
  DOCX, TXT, Markdown, and HTML. Extracted text contains the replacement without
  the previous text.

Disposable test outputs and selected/committed screenshots are in
`../.codex-tmp/pdf-save-export-regression/`.

## Scope and dependency

Scanned text remains part of an image; this change does not add OCR. Text-based
Word/HTML/Markdown exports retain their existing reading-conversion behavior,
rather than reproducing the PDF's full layout. This is not a secure-redaction
feature.

The new pinned MuPDF 1.28.0 dependency is **AGPL-3.0-or-later**, not MIT. Its
license is included in the packaged app at `modules/pdf/vendor/mupdf/LICENSE`.
Distribution of an application using it must account for those license terms
or an appropriate commercial license.
