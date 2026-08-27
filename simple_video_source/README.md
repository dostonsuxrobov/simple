# simple video

`simple_video_source` is the standalone video-viewer component used by the unified **simple** app. It is an Electron + Vite + React project and stays local-first: videos are read directly from disk and are never uploaded.

## Features

- Opens videos from the file picker, Windows command line/file association, recent list, or drag and drop.
- Custom play/pause, timeline seeking, volume, mute, playback speed, fullscreen, and Fit / Fill / Actual size controls.
- Shows duration, pixel dimensions, container, file size, modification time, and file location.
- Keeps a bounded 12-item recent-video list in Electron's per-user application data.
- Supports multi-file drops by opening additional videos in separate windows.
- Uses a sandboxed preload bridge with context isolation; the renderer has no Node.js or raw filesystem access.

## Supported routing extensions

`.mp4`, `.m4v`, `.webm`, `.ogv`, `.mov`, `.mkv`

An extension identifies the container, not the codecs inside it. Actual playback depends on the codecs supported by the Electron/Chromium build and the computer. MP4/M4V with H.264 + AAC and WebM with VP8/VP9/AV1 + Opus are the most portable choices. MOV and MKV files are routed here, but unsupported codecs produce a clear in-app message. `.ogg` is intentionally excluded because it is ambiguous with audio-only Ogg files; video Ogg uses `.ogv`.

## Keyboard shortcuts

| Action | Shortcut |
| --- | --- |
| Open | `Ctrl+O` |
| Play/pause | `Space` or `K` |
| Seek 5 seconds | `Left` / `Right` |
| Seek 10 seconds | `J` / `L` |
| Volume | `Up` / `Down` |
| Mute | `M` |
| Playback speed | `[` / `]` |
| Start/end | `Home` / `End` |
| Fullscreen | `F` |
| Shortcut reference | `?` |

## Development and verification

```powershell
npm install
npm run dev
npm test
npm run build:web
```

`npm run build` runs the focused tests, validates TypeScript, creates the production renderer, and packages a portable Windows executable. `npm run build:dir` creates an unpacked build for quick packaged-mode checks.

The pure routing contract lives in `electron/routing.cjs`; the unified launcher can consume the exact extension list without parsing the UI. Command-line arguments are resolved against the launching process's working directory, filtered by that list, de-duplicated case-insensitively on Windows, and opened in stable input order.
