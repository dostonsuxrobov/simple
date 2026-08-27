# Simple Docs

Simple Docs is a local-first, portable Windows editor for `.docx` documents. It uses a canvas-based paginated layout engine, keeps document work offline, and provides direct editing, formatting, tables, images, comments, tracked suggestions, search, navigation, DOCX saving, PDF export, printing, and crash recovery.

## Development

- `npm run dev` starts the Vite renderer and Electron shell.
- `npm run typecheck` validates the renderer types.
- `npm run build` creates the single-file portable Windows executable.

The editor engine is `@forevka/wordcanvas`, used under its MIT license.
