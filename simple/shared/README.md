# simple/shared: the shared Save / Import / Export layer

This folder is the **single source of truth** for the code every workspace uses to save, open, import, export and recover files. Edit shared I/O code here and nowhere else.

## Layout

| Folder | Contents | Vendored to |
|---|---|---|
| `electron/` | CommonJS for the main process (`.cjs`) and its data (`.json`) | `<workspace>/electron/simple-io/` |
| `renderer/` | Strict TypeScript for the renderer (DOM types, no React) and its data | `<workspace>/src/simple-io/` |
| `preload/io-bridge.cjs` | The code that exposes `window.simpleIO` | the marked block in `<workspace>/electron/preload.cjs` |
| `manifest.json` | Which files go to which workspace, and whether that workspace is enabled | not vendored |

The launcher and bootstrap in `simple/` require `simple/shared/electron/*` directly. `simple/package.json` packages this folder through `"shared/**/*"` in `build.files`.

## Vendored copies are generated

The workspaces (`simple_pdf_source`, `simple_calc_source`, `simple_doc_source`, `simple_image_source`, `simple_video_source`) receive byte-identical copies, because standalone workspace builds cannot reach outside their folder and sandboxed preloads cannot `require` files. Those copies are **generated**:

- Never edit anything in `electron/simple-io/`, `src/simple-io/`, or between `// <simple-io-bridge v1>` and `// </simple-io-bridge>` in a preload. A stale vendored copy in those folders (it starts with the vendored header below, or it is a data file named like one the manifest vendors) is removed by the next sync, one file at a time; any other file or folder there is reported and left in place for you to move.
- The manifest's `destinations` must be folders named `simple-io` (`electron/simple-io`, `src/simple-io`); any other name is refused before anything is written, so a typo can never point the sync at a workspace's own sources.
- After changing a file here, run `npm run sync:shared` in `simple/`. It copies changed files, removes stale vendored copies, and replaces the preload block (inserting the markers at the end of the preload when they are missing).
- `npm run check:shared` (also run by `npm test` through `scripts/verify.cjs`, and by `npm run sync` before bundling) writes nothing and fails with the drifted workspace, the file, and the fix command. A duplicated or half-marked preload block also fails; markers that are not balanced must be fixed by hand.
- `npm run watch` watches this folder and vendors on every change, except a change to `manifest.json`: that is only checked (`--check`), and you run `npm run sync:shared` yourself once the manifest is right, then restart the watch.
- Options: `--module=<name>[,<name>]` limits the run to enabled workspaces; `--target-root <dir>` and `--shared-root <dir>` point the tool at a temporary copy (used by `tests/shared-sync.test.cjs`).

Every `.cjs`, `.js`, `.mjs` and `.ts` source here must begin with this exact line, so each vendored copy says where it comes from while staying byte-identical:

```
// Vendored from simple/shared/<path> by simple/scripts/sync-shared.cjs. Do not edit here.
```

JSON files cannot hold comments and carry no header. `renderer/io-catalog.json` must be byte-identical to `electron/io-catalog.json`.

## Enabling a workspace

Each workspace in `manifest.json` has `"enabled": false` until its integrator wires it up. Disabled workspaces are never written or checked. To enable one:

1. Set `"enabled": true` and run `npm run sync:shared`.
2. Wire it. For every `wiring` rule whose `when` source exists in this folder, the workspace's `electron/main.cjs` must contain the `call` (for example `registerSharedIo(`); the check fails until it does.
3. Run `npm test` here and in the workspace, then `npm run sync`. Vendored copies are part of the workspace fingerprint, so a shared change marks the module stale until it is rebuilt.

Files listed in the manifest that do not exist yet are reported as planned and skipped.

## Rules for shared code

- Require **only Node built-ins and `electron`**. Workspaces have different dependencies (Calc has no `cfb`, Image has no `jszip`), so third-party behaviour is injected by the caller: a deep validator, a `prepareInput` for the office converter, a module serializer.
- Renderer code must compile under each workspace's `tsc --noEmit` and must not import React.
- Every public function has JSDoc.
- Never require LibreOffice. It is used only when it is already on this PC (the probe in `office-engine.cjs` finds it); every format offered in a picker must work without it, and no message tells the user to fetch or set up anything. Nothing in the app runs or mentions `scripts/setup-office-runtime.ps1`, a developer tool for test PCs.
- Everything stays on this PC: no network modules or requests, no remote addresses, no cloud-service names (a file another program holds is LOCKED or FILE_UNAVAILABLE), and documents are handed only to Simple's own workspaces (`openInSimple`), never to another program. `scripts/local-only-guard.cjs` (run by `npm test`) enforces this for `shared/`, `launcher/` and `electron/`, and checks `tests/` and `scripts/` for service names too. The names themselves never appear in the source: `scripts/local-only-names.json` holds their SHA-256 hashes.
- Start Windows programs by absolute path (`systemProgram('reg.exe')`, System32), never by a bare name: Windows looks in the current folder first. The guard rejects bare-name `spawn`/`execFile` calls.
- A save request flag that skips a safety check (`force`, `recreate`, `confirmedLossy`, `confirmedSibling`) counts only when main showed that window the matching prompt (`simpleIO.prompt`) and the user chose that answer. `force` replaces only the version the user was shown.
- Restoring a recovery copy is `recovery.read(id)`, then rebuilding the model, then `recovery.adopt(id)`; `DocumentSession.restore` does all three. Until a save writes the file, a restored document (and one whose file an export replaced, `ownFileChanged`) is never saved through the pristine shortcut.

## Checks

- `npm test` runs `verify.cjs` (drift, wiring, local-only guard), `scripts/check-shared-renderer.cjs` (the renderer TypeScript type-checks with each workspace's own TypeScript and tsconfig) and every `tests/*.test.cjs` and `tests/*.test.mjs` (the renderer modules), including the fast parts of the I/O gates (`tests/io-gates.test.cjs`) and the vendoring dry run (`scripts/check-vendoring.cjs`, which vendors every workspace into a temporary copy, never into the workspace folders).
- `npm run test:io` runs `scripts/no-office-matrix.cjs` (nothing needs the office engine; legacy files save to a modern sibling) and `scripts/io-acceptance.cjs` (save, lock, crash recovery, close while saving, drop and export in the real app) for every enabled workspace. Each enabled workspace needs an adapter in `scripts/io-adapters/<workspace>.cjs`.
