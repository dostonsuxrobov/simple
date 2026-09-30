declare global {
  interface DocumentPayload {
    data: Uint8Array;
    name: string;
    path: string | null;
    size: number;
    convertedFrom?: "doc" | "docx-renamed";
    sourcePath?: string | null;
    requiresSaveAs?: boolean;
    conversionWarnings?: string[];
    conversionMethod?: 'layout' | 'text';
    originalLayout?: { data: Uint8Array; extension: 'doc' | 'docx' };
    format?: 'doc' | 'docx';
    sourceData?: Uint8Array;
    sourceHash?: string;
  }

  interface SavedFile {
    path: string;
    name: string;
    format?: 'doc' | 'docx';
    sourceData?: Uint8Array;
    sourceHash?: string;
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
    getDocumentFonts(): Promise<import('@forevka/wordcanvas').CustomFontDef[]>;
    getOriginalLayoutPdf(input: { data: Uint8Array; extension: 'doc' | 'docx' }): Promise<Uint8Array>;
    onDocumentShortcut(callback: (action: 'print' | 'save' | 'save-as' | 'export') => void): () => void;
    openFile(): Promise<DocumentPayload | null>;
    openPath(path: string): Promise<DocumentPayload>;
    openBytes(input: { data: Uint8Array; name: string }): Promise<DocumentPayload>;
    openInNewWindow(path?: string): Promise<boolean>;
    newWindow(): Promise<boolean>;
    saveDocx(input: { data: Uint8Array; path: string | null; name: string; forceDialog: boolean; format?: 'doc' | 'docx'; sourceData?: Uint8Array; expectedHash?: string; protectOriginal?: boolean }): Promise<SavedFile | null>;
    savePdf(input: { data: Uint8Array; name: string }): Promise<SavedFile | null>;
    saveExport(input: { data: Uint8Array; name: string; format: "docx" | "pdf" | "html" | "md" | "txt" }): Promise<SavedFile | null>;
    composePrintPdf(input: { data: Uint8Array; name: string; settings: PrintSettings }): Promise<PrintComposition>;
    listPrinters(): Promise<PrinterSummary[]>;
    printPdf(input: {
      data: Uint8Array;
      name: string;
      deviceName: string;
      copies: number;
      color: boolean;
      collate: boolean;
      duplexMode?: "simplex" | "shortEdge" | "longEdge";
      paper: PrintSettings["paper"];
      orientation: PrintSettings["orientation"];
      paperWidth: number;
      paperHeight: number;
      mixedPaperSizes: boolean;
    }): Promise<PrintResult>;
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

  interface PrintSettings {
    paper: "Document" | "Letter" | "A4" | "Legal" | "A5" | "Tabloid";
    orientation: "portrait" | "landscape";
    margins: {
      preset: "none" | "narrow" | "normal" | "wide" | "custom";
      custom: { top: number; right: number; bottom: number; left: number };
    };
    scaling: "fit" | "actual" | "custom";
    scalePercent: number;
    pages: "all" | "custom";
    pageRange: string;
    printTitle: boolean;
    printPageNumbers: boolean;
    centerContent: boolean;
  }

  interface PrinterSummary {
    name: string;
    displayName: string;
    supportsDuplex: boolean;
    supportsColor: boolean;
  }

  interface PrintResult {
    success: boolean;
    failureReason: string;
    printerName?: string;
    printerLabel?: string;
  }

  interface PrintComposition {
    data: Uint8Array;
    sourcePageCount: number;
    outputPageCount: number;
    paperWidth: number;
    paperHeight: number;
    mixedPaperSizes: boolean;
    placements: Array<{ sourcePage: number; scale: number; clipped: boolean }>;
    settings: Omit<PrintSettings, "margins"> & {
      margins: { preset: PrintSettings["margins"]["preset"]; top: number; right: number; bottom: number; left: number };
    };
  }
}

export {};
