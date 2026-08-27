# simple

`simple` is one portable Windows executable containing five sibling applications:

- `simple_doc_source` for `.docx`
- `simple_calc_source` for spreadsheets and tabular formats
- `simple_pdf_source` for PDF, text/Markdown, and legacy `.doc`
- `simple_image_source` for viewing, cropping, rotating, painting, saving, printing, and converting images to PDF
- `simple_video_source` for local video playback

The five modes stay isolated internally, preserving their renderers, preload bridges, IPC handlers, save behavior, and tests. The bootstrap selects a mode from the file extension before that mode starts. Mixed command-line file lists are split and opened in the correct modes.

## Build

```powershell
npm install
npm run build
```

The result is `release/simple.exe`. Every build reads and rebuilds the current sibling sources, bundles each backend, stages each renderer, generates `modules/manifest.json` with source hashes, and packages everything into the standalone executable. The finished EXE does not need Node.js, the source folders, or any of the original EXEs.

`npm run watch` watches all five sibling source trees. When one changes, it rebuilds that module and then refreshes the portable EXE. `npm test` verifies routing coverage, staged modules, package associations, source icon integrity, and shared icons.

## Routing policy

- `.docx` always opens Documents, even though the PDF source can import DOCX.
- Images open the dedicated Images workspace; the PDF source can still import them when that workflow is chosen explicitly.
- Common video formats open the dedicated Video workspace.
- `.txt` and `.md` open Read & edit; spreadsheet text formats use `.csv`, `.tsv`, or `.tab`.
- All workbook formats claimed by Calc route to Spreadsheets, except the intentionally reassigned `.txt` collision.
- Matching is case-insensitive.

## Shared icon and Windows defaults

The supplied PNG is copied unchanged to `assets/icon-source.png`, padded transparently to a square (never stretched), and converted into a multi-resolution ICO. That ICO is used by the EXE, every internal window, drag icons, packaging associations, and optional Windows registration.

A portable application cannot silently take over Windows defaults. From the launcher, choose **Add file support**, then **Choose defaults** and select `simple` for the formats you want. Registration is per-user and does not require administrator rights. If the portable EXE is moved, open it and register again so Windows points to the new location. **Remove** deletes only `simple`'s own registration.

## Portable behavior

The single EXE contains Electron and all five built modules. Like the original apps, it stores recents/recovery state under the current Windows user's application-data folder. For fast, concurrent workspace launches, the portable wrapper keeps a build-specific extracted runtime under `%LOCALAPPDATA%\\simple\\cache`; a new build gets a new cache key, so updated code cannot reuse a stale runtime. No separate application installation is required.
