# simple

simple is a local-first, portable PDF reader and editor for Windows. The ready-to-run build is:

`release/simple-1.4.3-portable.exe`

No installation is required. Move the executable anywhere, launch it, and open a document from the home screen or by dragging a supported file onto the app. The first launch creates a versioned runtime cache under `%LOCALAPPDATA%\simple\cache`; later launches reuse it instead of extracting Chromium again. The cache can be deleted safely and is recreated automatically.

## Included workflows

- Read PDFs with selectable text, links, basic form fields, search, hand panning, automatic fit-to-window, manual zoom, and page navigation.
- Open every additional document in its own independent application window.
- Open or collapse a thumbnail sidebar with page numbers and multi-page selection.
- Delete or rotate left/right directly from a selected thumbnail, reorder/duplicate pages, insert blank pages, and crop.
- Select one page, Ctrl-click several pages, or Shift-click a range, then drag a selected thumbnail's grip or the page handle in the document view to export the selection as one PDF. Drag the thumbnail itself to reorder it.
- Drop one or several external PDFs on the top or bottom half of a page in the document view (or between thumbnails) to insert every dropped page at that exact position.
- Read and save native PDF outline bookmarks that remain available in Acrobat and other compatible readers.
- Scroll the complete document continuously while only nearby pages are rendered for speed.
- Select a word or substring and type directly on the page—there is no edit popup—while preserving the source font, baseline, spacing, color, and placement wherever the PDF exposes them.
- Select images on the page to move, resize, rotate, replace, duplicate, or delete them; capture artwork regions for the same direct controls.
- Add new text and images, then embed all content edits when saving.
- Add highlights, underlines, and freehand ink directly on pages, then embed their appearance when saving.
- Open PDF, JPG, PNG, DOC, DOCX, TXT, and Markdown files. DOCX conversion preserves common headings, lists, tables, images, and styling; non-PDF sources are never changed.
- Open a print preview containing all saved and unsaved edits, then use its printer control for the native Windows print dialog.
- Enter distraction-free fullscreen reading with the eye control and leave with Escape; undo and redo document operations, use keyboard shortcuts, and reopen recent local files.

## Shortcuts

| Action | Shortcut |
| --- | --- |
| Open / Save / Save As | `Ctrl+O` / `Ctrl+S` / `Ctrl+Shift+S` |
| Undo / Redo | `Ctrl+Z` / `Ctrl+Y` |
| Find a word or sentence | `Ctrl+F`; `Enter` / `Shift+Enter` for next / previous |
| Print preview | `Ctrl+P` |
| Previous / next page | `Page Up` / `Page Down` |
| First / last page | `Ctrl+Home` / `Ctrl+End` |
| Hand / Select / Edit / Add text / Crop | `H` / `V` / `E` / `T` / `C` |
| Toggle sidebar | `F4` |

## Development

```powershell
npm install
npm run dev
```

Create the portable Windows executable with:

```powershell
npm run build
```

## PDF editing model

PDF text is stored as positioned drawing instructions rather than editable paragraphs. simple reconstructs an edited text or artwork region and embeds the result when saving. Text wraps inside its edited box, but it does not reflow unrelated content elsewhere on the page, and original covered text can remain extractable. Artwork-region editing is a visual region workflow; individual paths inside complex vector illustrations are not separated. Crop uses the standard non-destructive PDF CropBox. Markup is embedded visually, not added to another reader's Comments list. Digitally signed and password-protected documents are treated cautiously because ordinary edits can invalidate signatures or cannot be re-encrypted safely.

simple does not claim secure redaction, OCR, certificate signing, password encryption, accessibility remediation, or Acrobat's cloud-review/AI services. Those need a deeper PDF engine or carefully reviewed security implementation. See `ACROBAT_COMPARISON.md` for the researched scope and priorities.
