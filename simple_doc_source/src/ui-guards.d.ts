export interface ShortcutEventLike {
  key: string;
  ctrlKey?: boolean;
  metaKey?: boolean;
  shiftKey?: boolean;
}

export function documentExportDisabled(documentOpen: boolean, busy: boolean): boolean;
export function activeModal(root: ParentNode): HTMLElement | null;
export function isModalBlockedShortcut(event: ShortcutEventLike): boolean;
export function nextPrintPreviewGeneration(current: number): number;
export function canCommitPrintPreview(requestGeneration: number, currentGeneration: number, modalHidden: boolean): boolean;
