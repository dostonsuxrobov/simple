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
