// src/shared/session.ts (WP1)
// Host-side session types shared by src/main.tsx (Simple shell) and src/advanced/useAdvancedMode.tsx.
import type { ReactNode, RefObject } from 'react'
import type { AdvancedEditorHandle, ImportedDocument, PsdFidelityIssue } from '../advanced/types.ts'

export type SaveFormat = 'png' | 'jpeg' | 'webp' | 'psd'
/** 'psd' is offered only in Advanced mode; Simple keeps exactly png, jpeg, webp, pdf (qa/smoke.cjs). */
export type ExportFormat = 'png' | 'jpeg' | 'webp' | 'pdf' | 'psd'
export type EditorMode = 'simple' | 'advanced'

export interface OpenImage {
  readonly name: string
  readonly path: string | null
  readonly size: number
  readonly sourceFormat: string
  readonly saveFormat: SaveFormat
  readonly mime: string
  readonly hasAlpha: boolean
  /** Source WebP was lossless (VP8L); re-encode at quality 1.0 so it stays lossless. */
  readonly webpLossless: boolean
  /** PSD features Simple cannot write back (empty for other formats). */
  readonly psdIssues: readonly PsdFidelityIssue[]
}

/** The single content-revision clock shared by both modes (drives Modified and close protection). */
export interface RevisionClock {
  next(): number
  current(): number
  set(revision: number): void
  saved(): number
}

export interface AdvancedModeDeps {
  readonly canvasRef: RefObject<HTMLCanvasElement | null>
  readonly getImage: () => OpenImage | null
  readonly updateImage: (patch: Partial<OpenImage>) => void
  readonly revisions: RevisionClock
  readonly notify: (message: string, tone?: 'normal' | 'error') => void
  /** Print dialog or unsaved-changes prompt open. */
  readonly isSuspended: () => boolean
  readonly save: (forceDialog: boolean) => Promise<boolean>
  readonly openExportMenu: () => void
  readonly print: () => void
  /** Applies pending Simple markup/adjust/looks sessions and cancels an open crop box. */
  readonly settleSimple: () => Promise<void>
  /** Writes flattened pixels into the Simple canvas as ONE Simple undo step restoring `beforeRevision`. */
  readonly adoptFlattened: (flattened: HTMLCanvasElement, beforeRevision: number, pushUndo: boolean) => void
}

export interface AdvancedModeApi {
  readonly mode: EditorMode
  readonly handle: AdvancedEditorHandle | null
  /** Simple -> Advanced with the current canvas as Background (or "Layer 0" when it has alpha). */
  enter(): Promise<void>
  /** Opens an imported PSD directly in Advanced (the Simple canvas is stale until exit). */
  enterWithDocument(document: ImportedDocument, name: string): void
  /** Advanced -> Simple: flattens silently when lossless, otherwise shows the flatten dialog. */
  requestExit(): void
  /** Every Save/Export/Print/Copy path obtains pixels here; Advanced canvases are released after `use`. */
  withOutputCanvas<T>(use: (canvas: HTMLCanvasElement) => Promise<T>): Promise<T>
  /** Format the title-bar Save should use right now (layered documents save as PSD). */
  effectiveSaveFormat(): SaveFormat
  /** PSD bytes for Save/Export when effectiveSaveFormat() === 'psd'. */
  encodePsd(): Promise<Uint8Array>
  /** Returns true when the key event was consumed by the Advanced editor. */
  handleHostKey(event: KeyboardEvent): boolean
  /** Rendered by main.tsx in place of the Simple toolbar/content/status bar. */
  readonly view: ReactNode
  /** Flatten / overwrite-PSD dialogs. */
  readonly dialog: ReactNode
}
