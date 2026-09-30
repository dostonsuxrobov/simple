# simple

`simple` is one portable Windows executable containing five sibling applications:

- `simple_doc_source` for editable `.docx` files and safe legacy `.doc` imports
- `simple_calc_source` for spreadsheets and tabular formats
- `simple_pdf_source` for PDF and text/Markdown
- `simple_image_source` for viewing, cropping, rotating, painting, saving, printing, and converting images to PDF
- `simple_video_source` for local video playback and current-frame printing

Choose **Combine files** on the home screen to arrange PDFs, Word documents, Excel/OpenDocument workbooks, PNGs, and JPEGs into one PDF. Reorder files, select page ranges, and save once; the result opens in PDF for review. Source files stay intact. See [the reliability review](RELIABILITY_REVIEW.md) for tested improvements and remaining priorities.

The five modes stay isolated internally, preserving their renderers, preload bridges, IPC handlers, save behavior, and tests. The bootstrap selects a mode from the file extension before that mode starts. Mixed command-line file lists are split and opened in the correct modes.

## Printing

Every workspace uses an in-app two-pane print setup: layout and output options stay on the left while a live physical-paper preview stays on the right. Documents and PDFs compose vector PDF output; spreadsheets paginate their actual cell layout; images and the current video frame use exact paper-aware image placement. Letter, A4, and Legal paper, portrait/landscape orientation, margins, and scaling are available across the applicable workspaces, with format-specific controls such as page ranges, gridlines, crop/fill positioning, color, and print metadata.

The preview and final print route share the same layout model in each workspace. Pressing **Print** submits that prepared layout directly to the Windows default printer without opening a second system dialog. Printer hardware and drivers can still impose a non-printable edge or reject paper settings they do not support; those failures are reported inside Simple.

## Build

```powershell
npm install
npm run build
```

The result is `release/simple.exe`. Every build reads and rebuilds the current sibling sources, bundles each backend and the Combine worker, stages each renderer, generates `modules/manifest.json` with source hashes, and packages everything into the standalone executable. The finished EXE does not need Node.js, the source folders, or any of the original EXEs. Accurate legacy Word conversion and native edited XLS output use the local office engine described below.

`npm run watch` watches all five sibling source trees. When one changes, it rebuilds that module and then refreshes the portable EXE. `npm test` verifies routing coverage, staged modules, package associations, source icon integrity, and shared icons.

## Routing policy

- `.docx` and legacy `.doc` always open Documents, even though the PDF source can import Word files for explicit PDF-conversion workflows.
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

### Office compatibility engine

Legacy DOC layout, Word-to-PDF combination, and native edited XLS saving use a separate local LibreOffice engine. On this computer it is prepared in `%LOCALAPPDATA%\simple\office-runtime`. For another computer, run `scripts/setup-office-runtime.ps1` or use an installed LibreOffice. A companion `tools/libreoffice/program/soffice.exe` next to the portable EXE is also detected. The setup script verifies the pinned official release checksum and publisher signature and extracts locally; it does not change Windows file associations. Keep the included license files with redistributed runtime copies.
