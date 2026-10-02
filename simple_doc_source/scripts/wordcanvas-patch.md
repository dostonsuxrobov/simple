# WordCanvas 0.12.0 adapters

`patch-wordcanvas.cjs` holds every reviewed change to the pinned WordCanvas 0.12.0 dependency: the header image adapter, the creation and fidelity fixes listed below, and the generic SIMPLE_HOOKS layer (last section). Each adapter carries its own marker, validates every expected replacement before any file is written, and is idempotent.

## Header image adapter

The upstream band layout counts every header/footer image as in-flow content, including an anchored behind-text image. A 720 px watermark can reserve 726 px below the header and turn a one-page document into five pages. The same importer omits `wp:positionH/wp:align`, `wp:positionV/wp:align`, and `a:lum` metadata.

`patch-wordcanvas.cjs` applies an explicit, version-pinned adapter to nine emitted modules: browser/Node/import-worker importers, browser/Node/export-worker layout engines, browser/Node/export-worker exporters, and the browser canvas painter. It validates every expected replacement before writing any file and is idempotent. An unknown dependency version or changed unpatched pattern fails the build. The package's `postinstall`, `pretest`, and `prebuild:web` hooks apply it, so clean dependency installs reproduce the fix. Review and update the adapter before upgrading WordCanvas; do not remove the version guard just to make a build pass.

The adapter excludes anchored images from header/footer *flow measurement*, then places their computed image records in the page's behind/front layer. It honors page/margin/edge reference rectangles, explicit offsets, and left/center/right or top/center/bottom alignment. It respects the existing first/even/default header selection. Inline header images still reserve space. Background images paint before table/cell shading and text. Layout never moves or removes a block in the source document model: a header watermark remains a header watermark when DOCX is exported.

Both imported axes and the original DrawingML brightness/contrast values survive DOCX export. Browser/PDF previews approximate the common Office `bright="70000" contrast="-70000"` washout with 0.3 opacity; they do **not** implement a complete pixel-level Office luminance transform. The source image bytes and saved effect metadata remain intact. This approximation is darker than the supplied Google rendering. Other image effects, anchored shapes, nested floating images, inside/outside alignment and paragraph-relative positioning need further work.

The self-contained synthetic regression creates a DOCX in memory. It verifies one-page pagination, centered behind-text placement, unchanged source/model, header ownership and luminance/position DOCX round trips, PDF page count, inline-header spacing, and first/repeated header selection:

```text
node scripts/patch-wordcanvas.cjs
node scripts/test-wordcanvas-patch.mjs
```

An optional document path performs the same integration checks and emits local review artifacts; the default test never needs a private document. The supplied converted sample passes the one-page check, but a comparison against its original Google PDF still shows differences in floating table positioning, title width, gradients, paragraph alignment and table borders. The document app's Page view remains useful for these unsupported features. This patch does not imply Word/Google layout parity.

Primary references:

- [Microsoft: BehindDoc](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.anchor.behinddoc?view=openxml-3.0.1) describes layering relative to document text.
- [Microsoft: HorizontalPosition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.horizontalposition?view=openxml-3.0.1) and [VerticalPosition](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.verticalposition?view=openxml-3.0.1) describe alignment/offset within a reference rectangle.
- [Microsoft: LuminanceEffect](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.luminanceeffect?view=openxml-3.0.1) defines the preserved brightness/contrast attributes.
- [LibreOffice DrawingML import](https://github.com/LibreOffice/core/blob/master/oox/source/drawingml/fillproperties.cxx) recognizes the 70/-70 pair as Office watermark mode; other combined brightness/contrast effects use distinct color transforms.

The upstream dependency is MIT licensed; this adapter does not remove its copyright/license notices.

## Creation and fidelity adapters

The creation stress pass adds independent, idempotent adapters on the same pinned version:

- Table insertion keeps its existing undo transaction and selects the first cell.
- Pasted CSS colors normalize to solid RGB hex, which both Word and PDF writers understand.
- Shift+Enter exports as `w:br` instead of raw U+000B (invalid XML), and import keeps a soft line break in the same paragraph.
- Square-wrapped images export as aligned `wp:anchor` plus `wp:wrapSquare`; explicit free-position anchors retain their existing behavior. The source model is never changed by serialization.
- Floating text lines measure and reserve their footnotes before placement, using the same page-fit and note-commit routines as normal lines.
- The PDF font payload limit matches the host's 16 MB bound. The bundled SC cmap and an installed Malgun cmap choose a supplemental face only when SC lacks a character. The fallback stays within available font coverage, with no font download or redistribution.

Node adapter tests check valid XML, soft-break and wrapping round trips and note presence beside a float. `run-editor-stress.mjs` exercises the browser/import worker/export worker paths through native editor actions. `font-coverage.cjs` reads Unicode cmap formats 4 and 12; malformed tables yield no coverage. Future dependency upgrades must review every guarded replacement rather than bypassing the version check.

## SIMPLE_HOOKS: the Simple layer over the editor

Instead of one minified patch per feature, one adapter (`SIMPLE_WORDCANVAS_SIMPLE_HOOKS_V1`, applied last, browser editor module `dist-lib/editorApp-vN1g1Ew1.js` only) adds a small set of generic hooks. Features are then written as plain TypeScript in Simple through `src/engine-bridge.ts`, the single typed consumer. The bridge feature-detects every hook and degrades with one console warning per missing symbol; on an unpatched engine it detects document changes by polling the immutable document reference.

The marker line carries a fingerprint of the hook source (`// SIMPLE_WORDCANVAS_SIMPLE_HOOKS_V1 <sha256 prefix>`). When the script finds a different fingerprint, it stops with "Reinstall @forevka/wordcanvas 0.12.0 (npm ci)" and writes nothing, so an install never keeps running an older hook revision. After editing the hook source, reinstall the dependency (or restore `editorApp-vN1g1Ew1.js` from the 0.12.0 tarball) and run the script again.

### Document-change event

After every committed transaction that changes the document model, the editor emits the WordCanvas `custom` event `{ name: "simple:docchange", payload: { revision, origin, canUndo, canRedo } }` (`editor.on("custom", …)`, or the bridge's `onDocChange`).

- `revision` increases by one per event for the life of the mounted editor, across document loads. `handle.getModelRevision()` returns the current value.
- `origin` is `typing`, `paste` or `command` for edits (the engine's transaction origin), `undo`, `redo`, `remote` (collaboration ops, unused offline), or `load` after `openDocx`/`setDocument` replaced the document (not an edit).
- It fires for typing, IME commits, paste, ribbon and mini-toolbar commands, context-menu commands, Find & Replace, ruler and table-border drags, picture moves and resizes, dialogs (Page setup, Paragraph, Styles), header/footer and footnote edits, and the handle methods below.
- It does not fire for selection changes, view changes (zoom, ruler, grid, formatting marks, panes), decorations, pending formatting at a collapsed caret, or transient previews (drag ghosts, crop preview, IME composition text). Those transactions use the engine's `transient` origin and always end with one committed transaction.
- Review-only changes that leave the model untouched (comments, resolving threads, a deletion recorded as a suggestion in Suggesting mode) are reported by the existing `reviewChanged` event, not by `simple:docchange`. Dirty tracking needs both signals.
- Listeners run synchronously after the editor has re-rendered, so keep them cheap and debounce expensive work. A throwing listener is logged and never breaks the edit.

### Handle methods

All methods are on the handle returned by `whenReady()` and on every custom ribbon action context. `handle.simpleHooks === 1` identifies this revision. The engine's debug global also exposes the handle as `window.__cw.handle`.

| Method | Contract |
|---|---|
| `insertImageBytes(bytes, mime, opts?)` | Decodes the bytes first; undecodable data (for example SVG, HEIC or TIFF in Chromium) resolves `false`. Registers the bytes in the engine media store and inserts a picture at the caret (body paragraph or table cell) as one undo step. Default size: the natural size, capped at the text width of the caret's page (Word fits pictures to the column; the ribbon path keeps its 480 px cap). `opts`: `widthPx`/`heightPx` (one side keeps the aspect ratio), `maxWidthPx`, and `at: { clientX, clientY }` to place the caret at a drop point first. |
| `insertBlocks(blocks, opts?)` | Inserts model blocks in one transaction. At the caret, the selection is deleted first and the paragraph is split, using the engine's own object-insertion placement (which respects content controls). With `opts.index`, blocks go at that body index without a split. Every block and table cell gets a fresh id unless `opts.keepIds` is set. Blocks must only reference lists, styles, fields and notes that exist in the document. In Suggesting mode the insertion is recorded as a structural suggestion. |
| `replaceBlock(id, block \| (current) => block)` | Replaces one block anywhere (body, header/footer bands, table cells, shape text, footnotes, endnotes) as one undo step. The id is kept and containers are re-revised. The replacement never reuses the revision of what it replaces, because layout caches measurements by id and revision. The caret is kept when it is still valid; otherwise it moves to the first paragraph of the replacement. Refused (`false`) outside Editing mode, because the change could not be tracked as a suggestion. |
| `replaceImage(id, bytes, mime, opts?)` | Swaps a picture's bytes and keeps its id, wrap, alignment and anchor. By default the width is kept and the height follows the new aspect ratio; `fit: "frame"` keeps both sides. The old crop is dropped unless `keepCrop` or a new `crop` is given. Refused outside Editing mode. |
| `setDialogHandler(fn \| null)` | See the dialog bridge below. |
| `setInsertTextHook(fn \| null)`, `setBuiltinAutoCorrect(enabled)` | See the insert-text hook below. |
| `deleteWord(-1 \| 1)` | Ctrl+Backspace / Ctrl+Delete semantics for programmatic use. |
| `undo()`, `redo()`, `canUndo()`, `canRedo()` | The editor history. Simple removes the File tab, which holds the engine's Undo group. `canUndo`/`canRedo` are false in Viewing mode. |
| `setSelection(selection)`, `focus()`, `positionFromPoint(clientX, clientY)` | Sets the selection without scrolling, focuses the hidden input, and hit-tests a client point to a `DocPosition` (or `null`). |
| `seedReview(layer)` | Replaces the review layer (suggestions and comment threads) and emits `reviewChanged`; use it to restore review content. |
| `setDecorations(list)`, `clearDecorations()`, `invalidateDecorations()` | Already present at runtime in 0.12.0 but not declared in its types. Decorations are screen-only marks: `underline`, `highlight` or `box` over `range: { anchor, focus }`, or a badge `at` a position, each with an optional `onClick`. They are never exported or printed. |

### Dialog bridge

Electron 43 throws from `window.prompt()`, so the engine's hyperlink, bookmark and drop-down list commands failed. Every reachable engine dialog now calls the async handler registered with `setDialogHandler(request => …)`. A request is `{ kind: "prompt" | "confirm" | "alert", id, title, message, defaultValue? }`. Resolve a prompt with the entered text, or with `null` to cancel; resolve a confirm with `true` or `false`; an alert's result is ignored. The request ids are:

- `hyperlink.insert` and `hyperlink.edit` (context menu);
- `bookmark.add` and `bookmark.rename` (bookmarks pane);
- `content-control.dropdown-items` (Insert > Controls);
- `content-control.none` (alert);
- `document.open-failed` (alert, ribbon Open only).

The engine returns focus to the document after a prompt settles. Without a handler the engine falls back to the native dialog; when that throws, the command is cancelled with a console warning instead of an uncaught error. The handler receives a copy of the request, and a failing handler cancels the command.

`handle.openDocx()` now rejects when the file cannot be imported. Before, it showed an alert and resolved with the previous document still loaded; callers now report the error themselves. The developer inspector, collaboration and share alerts are unreachable offline and were left unchanged.

### Insert-text hook (AutoFormat)

`setInsertTextHook(hook)` calls `hook(context)` synchronously after the engine inserted typed text (one call per `beforeinput`; Shift+Enter inserts `"\v"`) and after Enter split a paragraph. It is not called for paste, IME composition, programmatic `insertText`, Viewing mode, a non-collapsed selection, or a keystroke that changed nothing. `context` is `{ kind: "text" | "paragraph", text, blockId, offset, paragraphText, paragraphStyle (a copy), mode }`, plus `previousBlockId` and `previousText` after Enter. The hook may return edits. They are applied as one separate transaction (origin `command`), so Ctrl+Z reverts only the correction:

- `{ type: "replace", blockId?, start, end, text, style? }`: the new text inherits the style of the first replaced character, patched with `style`;
- `{ type: "format", blockId?, start, end, style }`: patches the character style, for example `{ link: url }` or `{ verticalAlign: "super" }`;
- `{ type: "list", blockId?, kind: "bullet" | "number" }`: turns the paragraph into a list item after the text edits.

Offsets are clamped, edits must not overlap, and they apply right to left. In Suggesting mode they become suggestions. The engine's built-in typing conversions (smart quotes, `--` to an em dash, `(c)`/`(r)`/`(tm)`) run before the hook and stay on unless `setBuiltinAutoCorrect(false)` is called. A Simple AutoFormat that wants Word's rules (locale-aware quotes, `--` to an en dash between words) should turn them off.

### Word deletion

Ctrl+Backspace and Ctrl+Delete delete to the previous or next word boundary as one undo step, using the same boundaries as Ctrl+Left/Right in logical order. They are ignored with Alt (so AltGr layouts are unaffected) and during IME composition. With a selection they delete the selection. At a paragraph boundary they behave like Backspace/Delete: they merge with the neighbouring paragraph and respect table cells. Locked content controls are respected. The hidden input also handles `deleteWordBackward`/`deleteWordForward` input events for touch and IME keyboards.

### Verification

`node scripts/test-wordcanvas-patch.mjs` (part of `npm test`) checks that the installed module carries the current fingerprint and every hook edit exactly once, and that no engine `prompt()` call remains. It runs the injected helpers against the engine's real model functions, covering:

- event payloads and listener isolation;
- the dialog bridge, including Electron's throwing `prompt()`;
- correction transactions (ordering, style inheritance, caret mapping, clamping);
- block replacement in cells, bands and notes (structural sharing and revisions), id reminting and caret repair;
- image preparation.

It also runs the script again and expects no file changes, refuses a `0.12.1` package, and refuses a stale SIMPLE_HOOKS fingerprint without writing. With `SIMPLE_WORDCANVAS_PRISTINE_DIR=<extracted 0.12.0 package>` it additionally applies every adapter to pristine sources twice and compares the result with the installed modules byte for byte. `tests/engine-bridge.test.cjs` covers the bridge, including a contract check that every handle method the bridge uses is defined by this patch.

Browser behaviour was verified in Electron 43 against the built app through CDP:

- one event per typed character;
- no event for arrow keys, clicks, double-clicks, `setSelection`, search, opening the context menu or view toggles;
- one `command` event each for a ribbon Bold click, Find & Replace "All", context-menu Bullets, a picture drag (none during the six transient moves) and each handle method;
- `undo`, `redo`, `paste` and `load` origins;
- no event while an IME composition is open, and one on commit;
- the context-menu hyperlink prompt through an async handler;
- Ctrl+Backspace/Ctrl+Delete with one-step undo;
- AutoFormat corrections undone separately from typing.

Before upgrading WordCanvas, re-review every SIMPLE_HOOKS anchor (the script fails closed when one changes) and the ribbon ids in `src/engine-bridge.ts`, which derive from button labels. Then bump the marker version.
