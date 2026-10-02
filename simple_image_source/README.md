# Simple Image

Simple Image is the image-viewing and editing component for the unified **simple** Windows app. It is local-first: images are decoded, edited, printed, and saved on the same computer, with no network service or account.

## Supported input formats

| Extension | Viewing | Editing and save behavior |
| --- | --- | --- |
| `.png`, `.apng` | Yes | Saves in place as PNG; an animated PNG opens its first frame and Save creates a still copy |
| `.jpg`, `.jpeg`, `.jfif`, `.jpe`, `.jif` | Yes | Saves in place as JPEG |
| `.webp` | Yes | Saves in place as WebP (animated files save a still copy) |
| `.gif` | Yes, editable first frame | Save opens a PNG dialog |
| `.bmp` | Yes | Save opens a PNG dialog |
| `.svg`, `.svgz` | Yes, rasterized for editing | Save opens a PNG dialog |
| `.avif` | Yes | Save opens a PNG dialog |
| `.ico` | Yes, largest icon | Save opens a PNG dialog |
| `.psd` | Yes, with layers in Advanced mode | Saves in place as a layered PSD; an unchanged document saves its original bytes |

The main process identifies every file by its content (magic bytes), not its name, and checks the pixel size before anything reaches the editor: a single image is limited to 512 MB, 50 megapixels and 20,000 pixels per side. PNG, JPEG, and WebP are the flat output formats; Advanced mode also saves and exports layered Photoshop documents (PSD).

## Features

- Open from the picker, Windows file association, command line, or drag and drop.
- Inspect format, exact pixel dimensions, megapixels, aspect ratio, transparency, file size, and zoom.
- Fit-to-window, 1:1, and 5–800% zoom.
- Adjustable crop rectangle with eight resize handles, free placement, and exact live dimensions.
- Rotate left or right.
- Basic painting with brush color and 1–120 px size, plus a transparency eraser.
- Memory-bounded undo and redo.
- Atomic Save and Save As to PNG, JPEG, or WebP.
- One-click conversion of the current edited image to a one-page PDF.
- Two-pane print setup for the current edited canvas: live Letter/A4/Legal paper preview, portrait or landscape orientation, preset/custom margins, fit/fill/actual/custom scaling, five placement anchors, transparency background color, whole-composition grayscale, and 1–99 copies. The preview and hidden Windows print renderer share the same physical layout calculations, and Print sends directly to the default printer without opening the native print dialog. Background editing, shortcuts, and drops are isolated while the dialog is open; submitted controls freeze until Windows returns; focus returns to the invoking control; and owned temporary print folders are never swept from another live process.
- Save / Don't save / Cancel protection before close or replacement.
- Paste support in the bridge: `readClipboardImage()` returns the system clipboard image as validated PNG bytes (or `null`), and `clipboardHasImage()` reports whether one is available. Clipboard images get the same signature and size checks as files.
- Sandboxed renderer, isolated preload bridge, blocked navigation, and no Node.js access in the page.

## Photoshop documents (PSD)

A `.psd` opens directly in Advanced mode, parsed in the renderer before anything else changes, so a file that cannot be read leaves the current image open. Photoshop documents are read and written with [ag-psd](https://github.com/Agamnentzar/ag-psd) 31.0.2 (MIT; it uses pako 2.1.0, MIT and Zlib, and base64-js 1.5.1, MIT). The library is loaded only when a PSD is opened or saved, as its own chunk, so Simple mode does not pay for it.

- **Kept editable:** pixel layers (including pixels outside the canvas), the Background layer, layer order and names, opacity, all 27 blend modes, visibility, clipping masks, lock flags, layer masks (default color, pixels, disabled state), text layers (text, font, size, color, alignment, leading, tracking, point or paragraph box, transform) and the 13 adjustment layers Simple supports (Brightness/Contrast, Levels, Curves, Exposure, Vibrance, Hue/Saturation including Colorize, Color Balance, Black & White, Photo Filter, Invert, Posterize, Threshold, Gradient Map), plus the document resolution.
- **Converted, and listed when the file opens:** layer styles are left out; smart objects become their pixels; vector masks become pixel masks; pass-through groups dissolve into their layers (named "Group / Layer"), other groups are merged into one layer with the group's blend mode, opacity and mask; channel mixer, selective color, color lookup and gradient or pattern fill layers are left out; fill opacity is combined with layer opacity; Blend If and knockout settings are dropped; 16- and 32-bit documents are edited with 8 bits per channel; a non-sRGB color profile is shown as sRGB. A document with more than 200 layers has its bottom layers merged into one (the look is exact). Text keeps Photoshop's own pixels until it is edited.
- **Refused with the Photoshop steps to fix it:** CMYK, Lab, Duotone and Multichannel documents ("Image > Mode > RGB Color, then save a copy"), large documents (`.psb`), more than 16 channels, and documents beyond the size limits. Opening a document whose layers need more than 1.5 GB is refused with a clear message.
- **Saving:** a PSD saves as an 8-bit RGB layered document with the merged image embedded, so every reader shows the exact look. Text layers are re-rendered by Photoshop when it opens the file, using the font's PostScript name (exact names come from the installed fonts when Chromium's local-font access is available, otherwise from a built-in table of Windows and Adobe fonts). Shape layers are saved as pixels. An adjustment that Photoshop cannot store is saved as a pixel layer that keeps the exact look. Saving over the original PSD when it used features Simple could not keep asks first.
- Code: `src/advanced/psd.ts` (import/export), `src/advanced/psdMapping.ts` (pure mapping, tested in Node), `src/advanced/fonts.ts` (PostScript and CSS font names).

## Development

```powershell
npm install
npm run dev
```

## Verification and builds

```powershell
npm test
npm run build:web
npm run smoke:print
npm run smoke:print-render
npm run smoke:psd
npm run build:dir
npm run smoke:packaged-print
npm run build
```

`npm test` runs TypeScript checking plus focused file-contract, atomic-save, PDF-export, print-layout, silent-default-printer, PSD import/export round-trip (`tests/psd.test.cjs`) and PSD mapping (`tests/psd-mapping.test.cjs`) tests. `npm run smoke:psd` (after `npm run build:web`) runs the real backend and renderer in a hidden window with an isolated profile: a generated layered PSD opens from the command line into Advanced mode, an unchanged Save writes the original bytes, and the renderer's PSD code edits, saves and re-reads a copy; set `SIMPLE_IMAGE_QA_DIR` to choose the folder for its files. `npm run smoke:print` drives the two-pane Electron interface, responsive layout, modal keyboard isolation, and the final IPC print payload without sending a real job. `npm run smoke:print-render` runs the same generated page through Chromium's PDF renderer and verifies a single, correctly oriented physical sheet. The portable build is emitted to `release/` and uses the same shared `simple` icon as the other components.

## Integration contract

The Electron entry point is `electron/main.cjs`. A file path supplied at startup is dispatched only when its lowercase extension is included in `SUPPORTED_EXTENSIONS` in `electron/image-files.cjs`. That exported list is the canonical routing contract for the unified launcher; `build.fileAssociations` in `package.json` lists the same extensions.
