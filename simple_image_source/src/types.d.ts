/// <reference types="vite/client" />

interface ImagePayload {
  data: Uint8Array
  name: string
  path: string | null
  size: number
  /** Lower-case extension without the dot ('png', 'jpg', 'psd', ...). */
  format: string
  mime: string
  directSave: boolean
}

interface SavedImage {
  path: string
  name: string
  size: number
  format: string
}

interface ConvertedPdf {
  path: string
  name: string
  size: number
}

interface SaveImageInput {
  data: Uint8Array
  path: string | null
  name: string
  /** 'psd' is sent only from Advanced mode (layered documents). */
  format: 'png' | 'jpeg' | 'webp' | 'psd'
  forceDialog: boolean
  purpose?: 'save' | 'export'
}

interface ImagePrintSettings {
  paper: 'letter' | 'a4' | 'legal'
  orientation: 'portrait' | 'landscape'
  marginMm: number
  scaleMode: 'fit' | 'fill' | 'actual' | 'custom'
  scalePercent: number
  position: 'center' | 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right'
  background: string
  grayscale: boolean
  copies: number
}

interface SimpleImageBridge {
  openFile(): Promise<ImagePayload | null>
  openPath(path: string): Promise<ImagePayload>
  listSiblings(path: string): Promise<string[]>
  openBytes(name: string, data: ArrayBuffer): Promise<ImagePayload>
  openInNewWindow(path?: string): Promise<boolean>
  newWindow(): Promise<boolean>
  saveImage(input: SaveImageInput): Promise<SavedImage | null>
  convertToPdf(input: { data: Uint8Array; name: string }): Promise<ConvertedPdf | null>
  printImage(input: { data: Uint8Array; name: string; width: number; height: number; settings: ImagePrintSettings }): Promise<boolean>
  copyPng(data: Uint8Array): Promise<{ width: number; height: number }>
  /**
   * System clipboard image as validated PNG bytes, or null when the clipboard holds no image.
   * Optional until the preload provides it ('clipboard:read-image'); callers must feature-check.
   */
  readClipboardImage?(): Promise<Uint8Array | null>
  /** Whether the system clipboard currently offers an image. Optional like readClipboardImage. */
  clipboardHasImage?(): Promise<boolean>
  getVersion(): Promise<string>
  pathForFile(file: File): string | null
  setTitle(title: string): void
  minimize(): void
  toggleMaximize(): void
  confirmClose(): void
  onMaximized(callback: (maximized: boolean) => void): () => void
  onOpenExternal(callback: (path: string) => void): () => void
  onCloseRequested(callback: () => void): () => void
}

interface Window {
  simpleImage: SimpleImageBridge
}
