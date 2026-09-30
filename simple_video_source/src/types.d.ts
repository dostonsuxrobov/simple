export {}

declare global {
  interface VideoFilePayload {
    path: string
    name: string
    url: string
    size: number
    modifiedAt: number
    extension: string
  }

  interface RecentVideo {
    path: string
    name: string
    openedAt: number
  }

  interface VideoExportResult {
    canceled: boolean
    path?: string
    name?: string
  }

  interface VideoFrameExportPayload {
    sourcePath: string
    format: 'png' | 'jpeg'
    seconds: number
    bytes: ArrayBuffer
  }

  interface VideoPrintOptions {
    paperSize: 'letter' | 'a4' | 'legal'
    orientation: 'portrait' | 'landscape'
    margins: 'normal' | 'narrow' | 'wide'
    scaleMode: 'fit' | 'fill' | 'actual' | 'custom'
    customScale: number
    colorMode: 'color' | 'grayscale'
    metadata: boolean
  }

  interface VideoPrintSession {
    sessionId: string
    width: number
    height: number
    seconds: number
    sourceName: string
  }

  interface VideoPrintPreview {
    html: string
    title: string
    page: { label: string; widthInches: number; heightInches: number; marginInches: number }
    frame: { width: number; height: number; timestamp: string }
    placement: {
      scale: number
      renderedWidth: number
      renderedHeight: number
      availableWidth: number
      availableHeight: number
      cropped: boolean
    }
    options: VideoPrintOptions
  }

  interface VideoPrintResult {
    printed: boolean
    canceled: boolean
    path?: string
  }

  interface Window {
    simpleVideo: {
      openFile: () => Promise<VideoFilePayload | null>
      openPath: (filePath: string) => Promise<VideoFilePayload>
      openInNewWindow: (filePath?: string | null) => Promise<boolean>
      newWindow: () => Promise<boolean>
      listRecents: () => Promise<RecentVideo[]>
      removeRecent: (filePath: string) => Promise<RecentVideo[]>
      clearRecents: () => Promise<RecentVideo[]>
      revealFile: (filePath: string) => Promise<boolean>
      exportOriginalCopy: (filePath: string) => Promise<VideoExportResult>
      exportFrame: (payload: VideoFrameExportPayload) => Promise<VideoExportResult>
      copyFrame: (bytes: ArrayBuffer) => Promise<{ width: number; height: number }>
      startPrint: (payload: { sourcePath: string; seconds: number; bytes: ArrayBuffer }) => Promise<VideoPrintSession>
      renderPrintPreview: (payload: { sessionId: string; options: VideoPrintOptions }) => Promise<VideoPrintPreview>
      printFrame: (payload: { sessionId: string; options: VideoPrintOptions }) => Promise<VideoPrintResult>
      endPrint: (sessionId: string) => Promise<boolean>
      pathForFile: (file: File) => string | null
      setTitle: (title: string) => void
      minimize: () => void
      toggleMaximize: () => void
      toggleFullscreen: () => void
      exitFullscreen: () => void
      close: () => void
      onOpenExternal: (callback: (filePath: string) => void) => () => void
      onMaximized: (callback: (value: boolean) => void) => () => void
      onFullscreen: (callback: (value: boolean) => void) => () => void
    }
  }
}
