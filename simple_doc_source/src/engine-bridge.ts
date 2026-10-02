/**
 * The "Simple layer" over WordCanvas 0.12.0.
 *
 * One typed wrapper over the SIMPLE_HOOKS engine patch (scripts/patch-wordcanvas.cjs,
 * documented in scripts/wordcanvas-patch.md) and over the runtime-only engine surface
 * that the published types do not declare: decorations, review seeding, selection,
 * undo/redo and ribbon items. Every capability is feature-detected. A missing symbol
 * logs one console warning and degrades to a safe result (false, null or a no-op), so
 * the app keeps working on an unpatched engine; document-change tracking then falls
 * back to polling the immutable document reference.
 *
 * Model changes and review changes are separate signals: `onDocChange` fires for every
 * committed model change (typing, commands, paste, undo/redo, document loads) and never
 * for selection, view or transient preview changes; `onReviewChange` fires for comments
 * and for tracked-change records that leave the model untouched (for example a deletion
 * in Suggesting mode). Dirty tracking needs both.
 *
 * This module has no runtime imports, so Node tests can load it directly.
 */
import type { DocSelection, EditMode, EditorHandle, ReviewLayer, WordCanvasEventMap } from "@forevka/wordcanvas";
import type { Block, CharStyle, DocPosition, ParaStyle } from "@forevka/wordcanvas/query";

/** Hook API version this bridge was written against (the handle reports `simpleHooks`). */
export const SIMPLE_HOOKS_VERSION = 1;
/** Name of the WordCanvas `custom` event emitted after every committed model change. */
export const DOC_CHANGE_EVENT = "simple:docchange";

/** Where a committed model change came from. "unknown" only occurs in the polling fallback. */
export type DocChangeOrigin = "typing" | "paste" | "command" | "undo" | "redo" | "remote" | "load" | "unknown";

export interface DocChange {
  /** Increases by one per change for the life of the editor, across document loads. */
  revision: number;
  /** "load" means the document was replaced (openDocx/setDocument), not edited. */
  origin: DocChangeOrigin | (string & {});
  canUndo: boolean;
  canRedo: boolean;
}

/** Request ids WordCanvas 0.12.0 sends through the dialog bridge. */
export type EngineDialogId =
  | "hyperlink.insert"
  | "hyperlink.edit"
  | "bookmark.add"
  | "bookmark.rename"
  | "content-control.dropdown-items"
  | "content-control.none"
  | "document.open-failed";

interface DialogRequestBase {
  /** Stable call-site id, see EngineDialogId. */
  id: EngineDialogId | (string & {});
  /** Short heading for the dialog. */
  title?: string;
  /** The engine's label or message text. */
  message: string;
}

/** What the engine asks Simple to show instead of window.prompt/confirm/alert. */
export type EngineDialogRequest =
  | (DialogRequestBase & { kind: "prompt"; defaultValue?: string })
  | (DialogRequestBase & { kind: "confirm" })
  | (DialogRequestBase & { kind: "alert" });

export type EngineDialogResult = string | boolean | null | undefined | void;

/**
 * Shows an engine dialog. Resolve a prompt with the entered text, or null when the user
 * cancels; resolve a confirm with true or false; an alert's value is ignored. The engine
 * returns focus to the document after a prompt resolves.
 */
export type EngineDialogHandler = (request: EngineDialogRequest) => EngineDialogResult | Promise<EngineDialogResult>;

/** Passed to the insert-text hook after the engine inserted typed text or split a paragraph. */
export interface InsertTextContext {
  /** "paragraph" after Enter; "text" for typed characters (Shift+Enter inserts "\v"). */
  kind: "text" | "paragraph";
  text: string;
  /** Paragraph that holds the caret after the insertion. */
  blockId: string;
  /** Caret offset in that paragraph after the insertion. */
  offset: number;
  paragraphText: string;
  paragraphStyle: ParaStyle;
  mode: EditMode;
  /** For "paragraph": the paragraph that was split (the text before Enter). */
  previousBlockId?: string;
  previousText?: string;
}

/**
 * A correction returned by the insert-text hook. All edits from one call become one
 * transaction after the typed text, so Ctrl+Z reverts only the correction. Offsets are
 * UTF-16 offsets into the paragraph (`blockId`, default: the caret paragraph); edits must
 * not overlap. `style` patches the character style of the new or formatted text.
 */
export type InsertTextEdit =
  | { type: "replace"; blockId?: string; start: number; end: number; text: string; style?: Partial<CharStyle> }
  | { type: "format"; blockId?: string; start: number; end: number; style: Partial<CharStyle> }
  | { type: "list"; blockId?: string; kind: "bullet" | "number" };

/** Called synchronously on every typed insertion; keep it fast. Return edits or nothing. */
export type InsertTextHook = (context: InsertTextContext) => readonly InsertTextEdit[] | null | undefined | void;

export type DecorationClick = (event: { clientX: number; clientY: number }) => void;

/** Screen-only marks painted over the document (never exported or printed). */
export type Decoration =
  | { type: "underline" | "highlight" | "box"; range: DocSelection; color: string; thickness?: number; opacity?: number; onClick?: DecorationClick }
  | { type: "badge"; at: DocPosition; color: string; label?: string; onClick?: DecorationClick };

export interface InsertImageOptions {
  /** Display size in CSS px. One side keeps the picture's aspect ratio; default: natural size. */
  widthPx?: number;
  heightPx?: number;
  /** Width cap; default: the text width of the caret's page. */
  maxWidthPx?: number;
  /** Place the caret at this client point first (for drops). */
  at?: { clientX: number; clientY: number };
}

export interface ReplaceImageOptions {
  widthPx?: number;
  heightPx?: number;
  /** "width" (default) keeps the frame width and adopts the new aspect ratio; "frame" keeps both sides. */
  fit?: "width" | "frame";
  /** New crop, or null to clear it. By default the old crop is dropped unless keepCrop is set. */
  crop?: unknown;
  keepCrop?: boolean;
}

export interface InsertBlocksOptions {
  /** Insert at this body index instead of at the caret (no paragraph split). */
  index?: number;
  /** Keep the given block ids. By default every inserted block and cell gets a fresh id. */
  keepIds?: boolean;
}

/** The handle methods added by the SIMPLE_HOOKS patch (not in the published types). */
export interface SimpleHooksHandle {
  simpleHooks: number;
  getModelRevision(): number;
  insertImageBytes(bytes: Uint8Array | ArrayBuffer, mime: string, options?: InsertImageOptions): Promise<boolean>;
  insertBlocks(blocks: Block | readonly Block[], options?: InsertBlocksOptions): boolean;
  replaceBlock(blockId: string, block: Block | ((current: Block) => Block | null | undefined)): boolean;
  replaceImage(imageId: string, bytes: Uint8Array | ArrayBuffer, mime: string, options?: ReplaceImageOptions): Promise<boolean>;
  setDialogHandler(handler: EngineDialogHandler | null): void;
  setInsertTextHook(hook: InsertTextHook | null): void;
  setBuiltinAutoCorrect(enabled: boolean): void;
  deleteWord(direction: -1 | 1): boolean;
  undo(): void;
  redo(): void;
  canUndo(): boolean;
  canRedo(): boolean;
  setSelection(selection: DocSelection | null): void;
  focus(): void;
  seedReview(review: ReviewLayer): void;
  positionFromPoint(clientX: number, clientY: number): DocPosition | null;
  setDecorations(decorations: readonly Decoration[]): void;
  clearDecorations(): void;
  invalidateDecorations(): void;
}

/** The WordCanvas instance (only its event subscription is used). */
export interface EngineEvents {
  on<E extends keyof WordCanvasEventMap>(event: E, handler: (data: WordCanvasEventMap[E]) => void): () => void;
}

/** Methods the engine's internal editor (`window.__cw.editor`) also offers. */
type RuntimeMethods = Pick<SimpleHooksHandle, "setSelection" | "focus" | "undo" | "redo" | "seedReview" | "setDecorations" | "invalidateDecorations">;

/** The engine's debug global `window.__cw`; its editor is the fallback for runtime methods. */
interface EngineRuntime {
  editor?: Partial<RuntimeMethods>;
}

export interface EngineBridgeOptions {
  /** Where ribbon items live (the editor host or the document); default: the global document. */
  root?: ParentNode | null;
  /** Receives degradation warnings; default: console.warn. */
  warn?: (message: string) => void;
  /** Polling interval of the unpatched document-change fallback; default 500 ms. */
  pollMs?: number;
  /** Source of the engine debug global; default: globalThis.__cw. */
  runtime?: () => EngineRuntime | undefined;
}

export interface EngineBridge {
  readonly handle: EditorHandle;
  /** True when the SIMPLE_HOOKS patch is present. */
  readonly hooked: boolean;
  onDocChange(listener: (change: DocChange) => void): () => void;
  onReviewChange(listener: (review: ReviewLayer) => void): () => void;
  /** Current model revision (the revision of the last DocChange). */
  revision(): number;
  insertImageBytes(bytes: Uint8Array | ArrayBuffer, mime: string, options?: InsertImageOptions): Promise<boolean>;
  insertBlocks(blocks: Block | readonly Block[], options?: InsertBlocksOptions): boolean;
  /** Replaces a block anywhere in the document as one undo step (refused outside Editing mode). */
  replaceBlock(blockId: string, block: Block | ((current: Block) => Block | null | undefined)): boolean;
  /** Swaps a picture's bytes, keeping its id, width, wrap and anchor (refused outside Editing mode). */
  replaceImage(imageId: string, bytes: Uint8Array | ArrayBuffer, mime: string, options?: ReplaceImageOptions): Promise<boolean>;
  setDialogHandler(handler: EngineDialogHandler | null): boolean;
  setInsertTextHook(hook: InsertTextHook | null): boolean;
  /** Turns the engine's built-in smart quotes, -- to em dash and (c)/(r)/(tm) symbols on or off. */
  setBuiltinAutoCorrect(enabled: boolean): boolean;
  /** Deletes the previous (-1) or next (1) word, like Ctrl+Backspace / Ctrl+Delete. */
  deleteWord(direction: -1 | 1): boolean;
  setDecorations(decorations: readonly Decoration[]): boolean;
  clearDecorations(): boolean;
  invalidateDecorations(): boolean;
  seedReview(review: ReviewLayer): boolean;
  getSelection(): DocSelection | null;
  setSelection(selection: DocSelection | null): boolean;
  focus(): boolean;
  undo(): boolean;
  redo(): boolean;
  canUndo(): boolean;
  canRedo(): boolean;
  positionFromPoint(clientX: number, clientY: number): DocPosition | null;
  /** Clicks a ribbon command by its data-ribbon-item id (the primary button of a split button). */
  clickRibbonItem(id: string): boolean;
  hasRibbonItem(id: string): boolean;
  ribbonItemIds(): string[];
  /** Removes this bridge's listeners, pollers, dialog handler and insert-text hook. */
  dispose(): void;
}

/**
 * Ribbon item ids verified against WordCanvas 0.12.0 inside Simple Docs (the File tab is
 * removed by Simple, so its Undo/Redo items do not exist; use undo()/redo()). Ids derive
 * from labels, so an engine upgrade must re-verify them.
 */
export const RIBBON_ITEM_IDS = {
  paste: "home.clipboard.paste",
  cut: "home.clipboard.cut",
  copy: "home.clipboard.copy",
  formatPainter: "home.clipboard.format-painter",
  growFont: "home.font.grow-font",
  shrinkFont: "home.font.shrink-font",
  changeCase: "home.font.change-case",
  clearFormatting: "home.font.clear-all-formatting",
  bold: "home.font.bold",
  italic: "home.font.italic",
  underline: "home.font.underline",
  strikethrough: "home.font.strikethrough",
  superscript: "home.font.superscript",
  subscript: "home.font.subscript",
  highlight: "home.font.text-highlight-colour",
  fontColor: "home.font.font-colour",
  bulletedList: "home.paragraph.bulleted-list",
  numberedList: "home.paragraph.numbered-list",
  decreaseIndent: "home.paragraph.decrease-indent",
  increaseIndent: "home.paragraph.increase-indent",
  formattingMarks: "home.paragraph.show-hide-formatting-marks",
  alignLeft: "home.paragraph.align-left",
  alignCenter: "home.paragraph.center",
  alignRight: "home.paragraph.align-right",
  justify: "home.paragraph.justify",
  lineSpacing: "home.paragraph.line-spacing",
  findReplace: "home.editing.find-replace",
  selectAll: "home.editing.select-all",
  pageBreak: "insert.pages.page-break",
  sectionBreakNextPage: "insert.pages.section-break-next-page",
  insertTable: "insert.tables.insert-table",
  insertImage: "insert.illustrations.insert-image-from-your-device",
  insertEquation: "insert.equation.insert-equation",
  insertSymbol: "insert.symbols.insert-symbol-or-special-character",
  hyperlink: "insert.links.insert-remove-hyperlink",
  tableOfContents: "insert.references.insert-update-table-of-contents",
  footnote: "insert.references.insert-footnote",
  endnote: "insert.references.insert-endnote",
  pageLayout: "layout.page-setup.page-layout",
  outlinePane: "view.show.outline-navigation-pane",
  bookmarks: "view.show.bookmarks-list-go-to-add-rename-delete",
  ruler: "view.show.horizontal-ruler",
  zoomOut: "view.zoom.zoom-out",
  zoomIn: "view.zoom.zoom-in",
} as const;

export type RibbonItemName = keyof typeof RIBBON_ITEM_IDS;

export function isHookedHandle(handle: EditorHandle): handle is EditorHandle & SimpleHooksHandle {
  const version = (handle as Partial<SimpleHooksHandle>).simpleHooks;
  return typeof version === "number" && version >= SIMPLE_HOOKS_VERSION;
}

function isDocChange(value: unknown): value is DocChange {
  const change = value as Partial<DocChange> | null;
  return !!change && typeof change.revision === "number" && typeof change.origin === "string"
    && typeof change.canUndo === "boolean" && typeof change.canRedo === "boolean";
}

function attributeValue(value: string) {
  return `"${value.replace(/["\\]/g, "\\$&")}"`;
}

export function createEngineBridge(editor: EngineEvents, handle: EditorHandle, options: EngineBridgeOptions = {}): EngineBridge {
  const hooked = isHookedHandle(handle);
  const extended = handle as EditorHandle & Partial<SimpleHooksHandle>;
  const warned = new Set<string>();
  const warn = options.warn ?? ((message: string) => console.warn(message));
  const runtime = options.runtime ?? (() => (globalThis as { __cw?: EngineRuntime }).__cw);
  const pollMs = Math.max(50, options.pollMs ?? 500);
  const subscriptions = new Set<() => void>();
  let ownsDialogHandler = false;
  let ownsInsertTextHook = false;

  const degrade = (feature: string, consequence: string) => {
    if (warned.has(feature)) return;
    warned.add(feature);
    warn(`[engine-bridge] ${feature} is unavailable: the WordCanvas SIMPLE_HOOKS patch is missing or outdated (run npm install to re-apply scripts/patch-wordcanvas.cjs). ${consequence}`);
  };
  const safely = <T>(label: string, run: () => T, fallback: T): T => {
    try {
      return run();
    } catch (error) {
      console.error(`[engine-bridge] ${label} failed`, error);
      return fallback;
    }
  };
  const deliver = <T>(listener: (value: T) => void, value: T) => safely("listener", () => listener(value), undefined);
  // Prefer the handle; fall back to the engine's internal editor for methods it has.
  const runtimeMethod = <K extends keyof RuntimeMethods>(name: K): RuntimeMethods[K] | undefined => {
    const own = extended[name];
    if (typeof own === "function") return own.bind(handle) as RuntimeMethods[K];
    const fallback = safely("runtime lookup", () => runtime()?.editor, undefined);
    const method = fallback?.[name];
    return typeof method === "function" ? method.bind(fallback) as RuntimeMethods[K] : undefined;
  };
  const track = (unsubscribe: () => void) => {
    let active = true;
    const stop = () => {
      if (!active) return;
      active = false;
      subscriptions.delete(stop);
      unsubscribe();
    };
    subscriptions.add(stop);
    return stop;
  };

  // Unpatched fallback: the model is immutable, so a new document reference is a change.
  const pollListeners = new Set<(change: DocChange) => void>();
  let pollTimer: ReturnType<typeof setInterval> | null = null;
  let polledRevision = 0;
  let lastDocument: unknown = null;
  const readDocument = () => safely("getDocument", () => handle.getDocument() as unknown, lastDocument);
  const poll = () => {
    const current = readDocument();
    if (current === lastDocument) return;
    lastDocument = current;
    const change: DocChange = { revision: ++polledRevision, origin: "unknown", canUndo: false, canRedo: false };
    for (const listener of [...pollListeners]) deliver(listener, change);
  };

  const ribbonRoot = (): ParentNode | null => options.root ?? (globalThis as { document?: Document }).document ?? null;
  const ribbonElement = (id: string) => safely("ribbon lookup", () => ribbonRoot()?.querySelector<HTMLElement>(`[data-ribbon-item=${attributeValue(id)}]`) ?? null, null);

  return {
    handle,
    hooked,
    onDocChange(listener) {
      if (hooked) {
        return track(editor.on("custom", (event) => {
          if (event?.name === DOC_CHANGE_EVENT && isDocChange(event.payload)) deliver(listener, event.payload);
        }));
      }
      degrade("simple:docchange", `Polling the document every ${pollMs} ms instead; origins are reported as "unknown".`);
      pollListeners.add(listener);
      if (!pollTimer) {
        lastDocument = readDocument();
        pollTimer = setInterval(poll, pollMs);
      }
      return track(() => {
        pollListeners.delete(listener);
        if (pollListeners.size === 0 && pollTimer) {
          clearInterval(pollTimer);
          pollTimer = null;
        }
      });
    },
    onReviewChange(listener) {
      return track(editor.on("reviewChanged", (event) => deliver(listener, event.review)));
    },
    revision() {
      if (hooked && typeof extended.getModelRevision === "function") return safely("getModelRevision", () => extended.getModelRevision!(), 0);
      return polledRevision;
    },
    async insertImageBytes(bytes, mime, insertOptions) {
      if (!hooked || typeof extended.insertImageBytes !== "function") {
        degrade("insertImageBytes", "Pictures cannot be inserted from bytes.");
        return false;
      }
      try {
        return await extended.insertImageBytes(bytes, mime, insertOptions);
      } catch (error) {
        console.error("[engine-bridge] insertImageBytes failed", error);
        return false;
      }
    },
    insertBlocks(blocks, insertOptions) {
      if (!hooked || typeof extended.insertBlocks !== "function") {
        degrade("insertBlocks", "Blocks cannot be inserted.");
        return false;
      }
      return safely("insertBlocks", () => extended.insertBlocks!(blocks, insertOptions), false);
    },
    replaceBlock(blockId, block) {
      if (!hooked || typeof extended.replaceBlock !== "function") {
        degrade("replaceBlock", "Blocks cannot be replaced.");
        return false;
      }
      return safely("replaceBlock", () => extended.replaceBlock!(blockId, block), false);
    },
    async replaceImage(imageId, bytes, mime, replaceOptions) {
      if (!hooked || typeof extended.replaceImage !== "function") {
        degrade("replaceImage", "Pictures cannot be replaced.");
        return false;
      }
      try {
        return await extended.replaceImage(imageId, bytes, mime, replaceOptions);
      } catch (error) {
        console.error("[engine-bridge] replaceImage failed", error);
        return false;
      }
    },
    setDialogHandler(handler) {
      if (!hooked || typeof extended.setDialogHandler !== "function") {
        degrade("setDialogHandler", "Engine prompts fall back to window.prompt(), which Electron does not support.");
        return false;
      }
      extended.setDialogHandler(handler);
      ownsDialogHandler = handler !== null;
      return true;
    },
    setInsertTextHook(hook) {
      if (!hooked || typeof extended.setInsertTextHook !== "function") {
        degrade("setInsertTextHook", "AutoFormat as you type is disabled.");
        return false;
      }
      extended.setInsertTextHook(hook);
      ownsInsertTextHook = hook !== null;
      return true;
    },
    setBuiltinAutoCorrect(enabled) {
      if (!hooked || typeof extended.setBuiltinAutoCorrect !== "function") {
        degrade("setBuiltinAutoCorrect", "The engine's built-in smart quotes stay on.");
        return false;
      }
      extended.setBuiltinAutoCorrect(enabled);
      return true;
    },
    deleteWord(direction) {
      if (!hooked || typeof extended.deleteWord !== "function") {
        degrade("deleteWord", "Word deletion is unavailable.");
        return false;
      }
      return safely("deleteWord", () => extended.deleteWord!(direction < 0 ? -1 : 1), false);
    },
    setDecorations(decorations) {
      const method = runtimeMethod("setDecorations");
      if (!method) {
        degrade("setDecorations", "Spelling marks and other decorations are not shown.");
        return false;
      }
      return safely("setDecorations", () => (method(decorations), true), false);
    },
    clearDecorations() {
      if (typeof extended.clearDecorations === "function") return safely("clearDecorations", () => (extended.clearDecorations!(), true), false);
      const method = runtimeMethod("setDecorations");
      if (!method) {
        degrade("clearDecorations", "Decorations cannot be cleared.");
        return false;
      }
      return safely("clearDecorations", () => (method([]), true), false);
    },
    invalidateDecorations() {
      const method = runtimeMethod("invalidateDecorations");
      if (!method) {
        degrade("invalidateDecorations", "Decorations are not repainted on demand.");
        return false;
      }
      return safely("invalidateDecorations", () => (method(), true), false);
    },
    seedReview(review) {
      const method = runtimeMethod("seedReview");
      if (!method) {
        degrade("seedReview", "Review content cannot be restored.");
        return false;
      }
      return safely("seedReview", () => (method(review), true), false);
    },
    getSelection() {
      return safely("getSelection", () => handle.getSelection(), null);
    },
    setSelection(selection) {
      const method = runtimeMethod("setSelection");
      if (!method) {
        degrade("setSelection", "The selection cannot be set.");
        return false;
      }
      return safely("setSelection", () => (method(selection), true), false);
    },
    focus() {
      const method = runtimeMethod("focus");
      if (!method) {
        degrade("focus", "The editor cannot be focused programmatically.");
        return false;
      }
      return safely("focus", () => (method(), true), false);
    },
    undo() {
      const method = runtimeMethod("undo");
      if (!method) {
        degrade("undo", "Undo is only available through Ctrl+Z.");
        return false;
      }
      return safely("undo", () => (method(), true), false);
    },
    redo() {
      const method = runtimeMethod("redo");
      if (!method) {
        degrade("redo", "Redo is only available through Ctrl+Y.");
        return false;
      }
      return safely("redo", () => (method(), true), false);
    },
    canUndo() {
      if (!hooked || typeof extended.canUndo !== "function") {
        degrade("canUndo", "Undo availability is unknown and reported as false.");
        return false;
      }
      return safely("canUndo", () => extended.canUndo!(), false);
    },
    canRedo() {
      if (!hooked || typeof extended.canRedo !== "function") {
        degrade("canRedo", "Redo availability is unknown and reported as false.");
        return false;
      }
      return safely("canRedo", () => extended.canRedo!(), false);
    },
    positionFromPoint(clientX, clientY) {
      if (!hooked || typeof extended.positionFromPoint !== "function") {
        degrade("positionFromPoint", "Client points cannot be mapped to document positions.");
        return null;
      }
      return safely("positionFromPoint", () => extended.positionFromPoint!(clientX, clientY), null);
    },
    clickRibbonItem(id) {
      const element = ribbonElement(id);
      if (!element) {
        if (!warned.has(`ribbon:${id}`)) {
          warned.add(`ribbon:${id}`);
          warn(`[engine-bridge] Ribbon item "${id}" was not found; the WordCanvas ribbon changed or the item was removed.`);
        }
        return false;
      }
      const target = element.tagName.toUpperCase() === "BUTTON" ? element as HTMLButtonElement : element.querySelector<HTMLButtonElement>(":scope > button:first-child");
      if (!target || target.disabled || target.getAttribute("aria-disabled") === "true") return false;
      return safely("ribbon click", () => (target.click(), true), false);
    },
    hasRibbonItem(id) {
      return ribbonElement(id) !== null;
    },
    ribbonItemIds() {
      return safely("ribbon inventory", () => Array.from(ribbonRoot()?.querySelectorAll("[data-ribbon-item]") ?? [], (element) => element.getAttribute("data-ribbon-item") ?? "").filter(Boolean), [] as string[]);
    },
    dispose() {
      for (const stop of [...subscriptions]) stop();
      if (pollTimer) clearInterval(pollTimer);
      pollTimer = null;
      pollListeners.clear();
      if (ownsDialogHandler) extended.setDialogHandler?.(null);
      if (ownsInsertTextHook) extended.setInsertTextHook?.(null);
      ownsDialogHandler = ownsInsertTextHook = false;
    },
  };
}
