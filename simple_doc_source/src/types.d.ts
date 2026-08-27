declare global {
  interface DocumentPayload {
    data: Uint8Array;
    name: string;
    path: string | null;
    size: number;
  }

  interface SavedFile {
    path: string;
    name: string;
  }

  interface RecentFile {
    path: string;
    name: string;
    openedAt: number;
  }

  interface RecoveryFile {
    id: string;
    title: string;
    sourcePath: string | null;
    updatedAt: number;
  }

  interface SimpleDocsBridge {
    openFile(): Promise<DocumentPayload | null>;
    openPath(path: string): Promise<DocumentPayload>;
    openInNewWindow(path?: string): Promise<boolean>;
    newWindow(): Promise<boolean>;
    saveDocx(input: { data: Uint8Array; path: string | null; name: string; forceDialog: boolean }): Promise<SavedFile | null>;
    savePdf(input: { data: Uint8Array; name: string }): Promise<SavedFile | null>;
    printPdf(input: { data: Uint8Array; name: string }): Promise<boolean>;
    getRecents(): Promise<RecentFile[]>;
    removeRecent(path: string): Promise<void>;
    saveRecovery(input: { id: string; data: Uint8Array; title: string; sourcePath: string | null }): Promise<void>;
    getRecoveries(): Promise<RecoveryFile[]>;
    loadRecovery(id: string): Promise<DocumentPayload>;
    clearRecovery(id: string): Promise<void>;
    pathForFile(file: File): string | null;
    setTitle(title: string): void;
    minimize(): void;
    toggleMaximize(): void;
    toggleFullscreen(): void;
    confirmClose(): void;
    openExternal(url: string): Promise<void>;
    onMaximized(callback: (maximized: boolean) => void): () => void;
    onFullscreen(callback: (fullscreen: boolean) => void): () => void;
    onOpenExternal(callback: (path: string) => void): () => void;
    onCloseRequested(callback: () => void): () => void;
  }

  interface Window {
    simpleDocs: SimpleDocsBridge;
  }
}

export {};
