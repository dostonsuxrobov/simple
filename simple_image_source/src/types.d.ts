interface ImagePayload {
  data: Uint8Array
  name: string
  path: string | null
  size: number
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
  format: 'png' | 'jpeg' | 'webp'
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
