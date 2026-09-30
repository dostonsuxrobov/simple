# WordCanvas 0.12.0 header image adapter

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

The creation stress pass adds independent, idempotent adapters on the same pinned version:

- Table insertion keeps its existing undo transaction and selects the first cell.
- Pasted CSS colors normalize to solid RGB hex, which both Word and PDF writers understand.
- Shift+Enter exports as `w:br` instead of raw U+000B (invalid XML), and import keeps a soft line break in the same paragraph.
- Square-wrapped images export as aligned `wp:anchor` plus `wp:wrapSquare`; explicit free-position anchors retain their existing behavior. The source model is never changed by serialization.
- Floating text lines measure and reserve their footnotes before placement, using the same page-fit and note-commit routines as normal lines.
- The PDF font payload limit matches the host's 16 MB bound. The bundled SC cmap and an installed Malgun cmap choose a supplemental face only when SC lacks a character. The fallback stays within available font coverage, with no font download or redistribution.

Node adapter tests check valid XML, soft-break and wrapping round trips and note presence beside a float. `run-editor-stress.mjs` exercises the browser/import worker/export worker paths through native editor actions. `font-coverage.cjs` reads Unicode cmap formats 4 and 12; malformed tables yield no coverage. Future dependency upgrades must review every guarded replacement rather than bypassing the version check.
