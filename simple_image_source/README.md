# Simple Image

Simple Image is the image-viewing and editing component for the unified **simple** Windows app. It is local-first: images are decoded, edited, printed, and saved on the same computer, with no network service or account.

## Supported input formats

| Extension | Viewing | Editing and save behavior |
| --- | --- | --- |
| `.png` | Yes | Saves in place as PNG |
| `.jpg`, `.jpeg` | Yes | Saves in place as JPEG |
| `.webp` | Yes | Saves in place as WebP |
| `.gif` | Yes, editable first frame | Save opens a PNG dialog |
| `.bmp` | Yes | Save opens a PNG dialog |
| `.svg` | Yes, rasterized for editing | Save opens a PNG dialog |
| `.avif` | Yes | Save opens a PNG dialog |

The main process checks both the extension and the file signature before data reaches the editor. A single image is limited to 512 MB. PNG, JPEG, and WebP are the available output formats.

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
- Native Windows printing from the current edited canvas.
- Save / Don't save / Cancel protection before close or replacement.
- Sandboxed renderer, isolated preload bridge, blocked navigation, and no Node.js access in the page.

## Development

```powershell
npm install
npm run dev
```

## Verification and builds

```powershell
npm test
npm run build:web
npm run build
```

`npm test` runs TypeScript checking plus focused file-contract and atomic-save tests. The portable build is emitted to `release/` and uses the same shared `simple` icon as the other components.

## Integration contract

The Electron entry point is `electron/main.cjs`. A file path supplied at startup is dispatched only when its lowercase extension is included in `SUPPORTED_EXTENSIONS` in `electron/image-files.cjs`. That exported list is the canonical routing contract for the unified launcher.
