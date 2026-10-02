/**
 * Word/Docs editing commands for Simple Docs, run from keyboard shortcuts,
 * command search and the engine's dialog bridge. Wherever WordCanvas has a
 * ribbon control, the command clicks it (through the engine bridge), so a
 * shortcut behaves exactly like the button; popover menus (line spacing,
 * change case, the link box) are driven the same way. The few commands the
 * engine has no control for (adding a missing heading style, clearing only
 * character formatting) dispatch one ordinary undoable transaction through
 * the engine's editor object, feature-detected like the bridge does.
 */
import type { DocSelection, EditorHandle } from "@forevka/wordcanvas";
import type { EngineBridge, EngineDialogRequest } from "../engine-bridge.ts";
import {
  bookmarkNameProblem,
  clearCharacterRuns,
  findParagraph,
  headingStyleDefinition,
  isCollapsed,
  linkAddressProblem,
  linkRangeAt,
  nextCaseMode,
  normalizeLinkAddress,
  paragraphText,
  resolveStyleChar,
  runStyleAt,
  selectionRanges,
  wordRangeAt,
  type DocumentLike,
  type StylesheetLike,
} from "./editing.ts";
import { openFormDialog } from "./form-dialog.ts";
import { SHORTCUT_RIBBON_ITEMS, shortcutFor, type ShortcutCommand } from "./shortcuts.ts";
import { formatCombo } from "./keys.ts";

export interface EditorCommandContext {
  bridge: EngineBridge;
  handle: EditorHandle;
  notify(message: string, tone?: "normal" | "error"): void;
  /** Puts keyboard focus in the document (placing a caret first when there is none). */
  focusEditor(): boolean;
  /** A document is open and its editing view is shown. */
  editorShown(): boolean;
  author: { name(): string; chosen(): boolean; choose(name: string): void };
  icons: { link: string; comment: string; bookmark: string; list: string };
}

/**
 * Editing commands that command search lists next to the ribbon's own buttons
 * (which it reads from the ribbon): the ones with no button, or a better one.
 */
export const EDITOR_SEARCH_COMMANDS: ReadonlyArray<{ command: ShortcutCommand; label: string; group: string; keywords?: string[] }> = [
  { command: "link", label: "Insert link", group: "Insert", keywords: ["hyperlink", "url", "address", "web"] },
  { command: "comment", label: "New comment", group: "Review", keywords: ["comment", "note", "review", "annotate"] },
  { command: "find", label: "Find", group: "Home", keywords: ["search"] },
  { command: "replace", label: "Replace", group: "Home", keywords: ["find and replace", "substitute"] },
  { command: "heading-1", label: "Heading 1", group: "Home", keywords: ["title", "style"] },
  { command: "heading-2", label: "Heading 2", group: "Home", keywords: ["subtitle", "style"] },
  { command: "heading-3", label: "Heading 3", group: "Home", keywords: ["style"] },
  { command: "normal-style", label: "Normal text", group: "Home", keywords: ["body text", "paragraph style", "plain"] },
  { command: "clear-character-formatting", label: "Clear character formatting", group: "Home", keywords: ["reset", "remove formatting", "plain text"] },
  { command: "line-spacing-1", label: "Line spacing 1.0 (single)", group: "Home", keywords: ["single spacing"] },
  { command: "line-spacing-1.5", label: "Line spacing 1.5", group: "Home", keywords: ["one and a half"] },
  { command: "line-spacing-2", label: "Line spacing 2.0 (double)", group: "Home", keywords: ["double spacing"] },
  { command: "change-case", label: "Change case (cycle)", group: "Home", keywords: ["uppercase", "lowercase", "capitalize", "title case"] },
];

type RuntimeState = { doc: DocumentLike; selection: DocSelection | null };
type RuntimeTransaction = { ops: unknown[]; selectionAfter: DocSelection | null; origin: string };
type RuntimeEditor = { dispatch(command: (state: RuntimeState) => RuntimeTransaction | null): void };

/** The engine's live editor (window.__cw.editor, re-pointed on every document load), when it can dispatch. */
function runtimeEditor(): RuntimeEditor | null {
  const editor = (globalThis as { __cw?: { editor?: Partial<RuntimeEditor> } }).__cw?.editor;
  return editor && typeof editor.dispatch === "function" ? editor as RuntimeEditor : null;
}

const nextTask = () => new Promise<void>((resolve) => setTimeout(resolve, 0));
const sameName = (left: string | null | undefined, right: string) => String(left ?? "").trim().toLowerCase().replace(/\s+/g, " ") === right.trim().toLowerCase().replace(/\s+/g, " ");

export function createEditorCommands(context: EditorCommandContext) {
  const { bridge, handle, notify } = context;
  const doc = () => handle.getDocument() as unknown as DocumentLike;

  const editable = (what: string) => {
    if (!context.editorShown()) return false;
    if (handle.getMode() === "view") {
      notify(`Switch from Viewing to Editing to ${what}.`);
      return false;
    }
    return true;
  };

  /** The current selection, placing a caret at the document start when there is none. */
  const currentSelection = (): DocSelection | null => {
    let selection = bridge.getSelection();
    if (!selection) {
      context.focusEditor();
      selection = bridge.getSelection();
    }
    return selection;
  };

  const ribbonElement = (id: string) => document.querySelector<HTMLElement>(`[data-ribbon-item="${CSS.escape(id)}"]`);

  /** Shows the ribbon tab holding a control, so its popovers open next to it. */
  const showRibbonTab = (element: Element | null) => {
    const panel = element?.closest<HTMLElement>("[data-ribbon-panel]");
    const tab = panel?.dataset.ribbonPanel;
    if (!panel || !tab || panel.classList.contains("active")) return;
    document.querySelector<HTMLElement>(`[data-ribbon-tab="${CSS.escape(tab)}"]`)?.click();
  };

  const clickRibbon = (id: string | undefined) => Boolean(id && bridge.clickRibbonItem(id));

  /**
   * Opens a ribbon control's popover invisibly, waits one task (the engine
   * registers its outside-click and Escape listeners then, and its own close
   * removes them) and clicks the chosen element inside it.
   */
  const chooseFromPopover = async (ribbonId: string, pick: (popover: HTMLElement) => HTMLElement | null | undefined) => {
    const before = new Set(document.querySelectorAll(".cw-pop"));
    if (!bridge.clickRibbonItem(ribbonId)) return false;
    const popover = [...document.querySelectorAll<HTMLElement>(".cw-pop")].find((element) => !before.has(element));
    if (!popover) return false;
    popover.style.visibility = "hidden";
    await nextTask();
    if (!popover.isConnected) return false;
    const target = pick(popover);
    if (!target) {
      // An outside press closes the popover through the engine's own handler.
      document.body.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
      popover.remove();
      return false;
    }
    target.click();
    return true;
  };

  const menuItem = (popover: HTMLElement, label: string) => [...popover.querySelectorAll<HTMLButtonElement>(".cw-menu button")]
    .find((button) => (button.querySelector("span:last-child")?.textContent ?? button.textContent ?? "").trim() === label);

  const setLineSpacing = (label: "1.0" | "1.5" | "2.0") => {
    if (!editable("change line spacing")) return Promise.resolve(false);
    return chooseFromPopover(SHORTCUT_RIBBON_ITEMS["line-spacing-1"]!, (popover) => menuItem(popover, label));
  };

  /** Word's Shift+F3: lowercase → UPPERCASE → Capitalize Each Word, on the selection or the word at the caret. */
  const changeCase = async () => {
    if (!editable("change case")) return false;
    const selection = currentSelection();
    if (!selection) return false;
    let target = selection;
    if (isCollapsed(selection)) {
      const block = findParagraph(doc(), selection.focus.blockId);
      const word = block ? wordRangeAt(paragraphText(block), selection.focus.offset) : null;
      if (!block || !word) return false;
      target = { anchor: { blockId: block.id, offset: word.start }, focus: { blockId: block.id, offset: word.end } };
    }
    const text = selectionRanges(doc(), target).map((range) => paragraphText(range.block).slice(range.start, range.end)).join("\n");
    const mode = nextCaseMode(text);
    if (!mode) return false;
    if (target !== selection) bridge.setSelection(target);
    const label = mode === "upper" ? "UPPERCASE" : mode === "lower" ? "lowercase" : "Capitalize Each Word";
    const done = await chooseFromPopover(SHORTCUT_RIBBON_ITEMS["change-case"]!, (popover) => menuItem(popover, label));
    if (target !== selection) bridge.setSelection(selection);
    context.focusEditor();
    return done;
  };

  const styleCards = () => [...document.querySelectorAll<HTMLElement>("#editor .rib-gallery .style-card")];
  const styleCard = (name: string) => styleCards().find((card) => sameName(card.querySelector(".name")?.textContent?.replace(/\s*ⓐ$/, ""), name));

  /** Adds a paragraph style to the document's stylesheet as one undoable step. */
  const addStyle = (definition: ReturnType<typeof headingStyleDefinition>) => {
    const editor = runtimeEditor();
    if (!editor) return false;
    let added = false;
    try {
      editor.dispatch((state) => {
        const sheet = state.doc.stylesheet as StylesheetLike | undefined;
        if (!sheet || sheet.styles.some((style) => style.id === definition.id)) return null;
        added = true;
        return { ops: [{ type: "setStylesheet", stylesheet: { ...sheet, styles: [...sheet.styles, definition] } }], selectionAfter: state.selection, origin: "command" };
      });
    } catch (error) {
      console.warn("The heading style could not be added", error);
      return false;
    }
    return added;
  };

  /** Applies a paragraph style through the ribbon's style gallery (adding a missing heading first). */
  const applyStyle = async (target: { level: 1 | 2 | 3 } | "normal") => {
    if (!editable("apply a style")) return false;
    if (!currentSelection()) return false;
    const sheet = doc().stylesheet ?? null;
    const style = target === "normal"
      ? sheet?.styles.find((candidate) => candidate.id === sheet.defaultStyleId) ?? sheet?.styles.find((candidate) => candidate.id === "Normal")
      : sheet?.styles.find((candidate) => candidate.id === `Heading${target.level}` || sameName(candidate.name, `heading ${target.level}`));
    const name = style?.name ?? (target === "normal" ? "Normal" : `Heading ${target.level}`);
    let card = styleCard(name);
    if (!card && !style && target !== "normal") {
      addStyle(headingStyleDefinition(target.level, sheet));
      card = styleCard(name);
    }
    if (!card && style) {
      // "Show only styles in use" hides unused styles: show them for this one click.
      const filter = ribbonElement("home.styles.show-only-styles-in-use");
      if (filter?.classList.contains("active")) {
        filter.click();
        styleCard(name)?.click();
        filter.click();
        return true;
      }
    }
    if (!card) {
      notify(`This document has no “${name}” style.`, "error");
      return false;
    }
    card.click();
    return true;
  };

  /** Word's Ctrl+Space: removes manual character formatting and keeps the paragraph style, links and notes. */
  const clearCharacterFormatting = () => {
    if (!editable("clear formatting")) return false;
    const selection = currentSelection();
    if (!selection) return false;
    const editor = runtimeEditor();
    if (!editor) return clickRibbon(SHORTCUT_RIBBON_ITEMS["clear-character-formatting"]);
    let changed = false;
    editor.dispatch((state) => {
      let range = state.selection;
      if (!range) return null;
      if (isCollapsed(range)) {
        const block = findParagraph(state.doc, range.focus.blockId);
        const word = block ? wordRangeAt(paragraphText(block), range.focus.offset) : null;
        if (!block || !word) return null;
        range = { anchor: { blockId: block.id, offset: word.start }, focus: { blockId: block.id, offset: word.end } };
      }
      const sheet = state.doc.stylesheet ?? null;
      const ops = selectionRanges(state.doc, range)
        .filter((part) => part.end > part.start)
        .map((part) => ({ type: "setRuns", blockId: part.block.id, runs: clearCharacterRuns(part.block.runs, part.start, part.end, resolveStyleChar(sheet, part.block.style.namedStyle)) }));
      if (!ops.length) return null;
      changed = true;
      return { ops, selectionAfter: state.selection, origin: "command" };
    });
    return changed;
  };

  /** Find (Ctrl+F) or Replace (Ctrl+H) in the engine's find bar, seeded with a short selection. */
  const openFind = (replace: boolean) => {
    if (!context.editorShown()) return false;
    const selection = bridge.getSelection();
    let seed = "";
    if (selection && !isCollapsed(selection) && selection.anchor.blockId === selection.focus.blockId) {
      const block = findParagraph(doc(), selection.focus.blockId);
      const [start, end] = [selection.anchor.offset, selection.focus.offset].sort((a, b) => a - b);
      seed = block ? paragraphText(block).slice(start, end) : "";
      if (seed.length > 100 || /[\n\v\t]/.test(seed)) seed = "";
    }
    if (!clickRibbon(replace ? SHORTCUT_RIBBON_ITEMS.replace : SHORTCUT_RIBBON_ITEMS.find)) return false;
    const panel = [...document.querySelectorAll<HTMLElement>(".cw-float-panel")].find((element) => element.style.display !== "none" && element.querySelector("input"));
    const findInput = panel?.querySelector<HTMLInputElement>("input");
    if (findInput && seed) {
      findInput.value = seed;
      findInput.dispatchEvent(new Event("input", { bubbles: true }));
    }
    const replaceInput = panel ? [...panel.querySelectorAll<HTMLInputElement>("input")].find((input) => input.placeholder === "Replace") : null;
    const focus = replace && replaceInput ? replaceInput : findInput;
    focus?.focus();
    focus?.select();
    return true;
  };

  const linkDialog = async (options: { address: string; editing: boolean; askText: boolean }) => openFormDialog({
    title: options.editing ? "Edit link" : "Insert link",
    icon: context.icons.link,
    fields: [
      ...(options.askText ? [{ id: "text", label: "Text to display", placeholder: "Link text" }] : []),
      { id: "address", label: "Address", value: options.address, placeholder: "https://example.com", autofocus: true, hint: "A web page, an email address or a file." },
    ],
    submitLabel: options.editing ? "Apply" : "Insert",
    ...(options.editing ? { extraAction: { id: "remove", label: "Remove link" } } : {}),
    validate: (values) => {
      const problem = linkAddressProblem(values.address);
      return problem ? { field: "address", message: problem } : null;
    },
  });

  /** Sets (or with null removes) the link on the current selection through the ribbon's link box. */
  const applyLink = (address: string | null) => chooseFromPopover(SHORTCUT_RIBBON_ITEMS.link!, (popover) => {
    const input = popover.querySelector<HTMLInputElement>(".cw-dialog input");
    const buttons = [...popover.querySelectorAll<HTMLButtonElement>(".cw-dialog button")];
    if (!input) return null;
    if (address === null) return buttons.find((button) => button.classList.contains("danger"));
    input.value = address;
    return buttons.find((button) => button.classList.contains("primary"));
  });

  /** Ctrl+K: links the selection or the word at the caret, edits the link at the caret, or inserts a new linked text. */
  const insertLink = async () => {
    if (!editable("insert a link")) return false;
    const selection = currentSelection();
    if (!selection) return false;
    const block = findParagraph(doc(), selection.focus.blockId);
    const collapsed = isCollapsed(selection);
    const start = selection.anchor.blockId === selection.focus.blockId ? Math.min(selection.anchor.offset, selection.focus.offset) : selection.focus.offset;
    const existing = block ? linkRangeAt(block, collapsed ? selection.focus.offset : start) : null;
    const word = collapsed && block ? wordRangeAt(paragraphText(block), selection.focus.offset) : null;
    const askText = collapsed && !existing && !word;
    const result = await linkDialog({ address: existing?.link ?? "", editing: Boolean(existing), askText });
    if (!result) {
      context.focusEditor();
      return false;
    }
    let target: DocSelection = selection;
    if (existing && collapsed && block) target = { anchor: { blockId: block.id, offset: existing.start }, focus: { blockId: block.id, offset: existing.end } };
    if (result.action === "remove") {
      bridge.setSelection(target);
      const removed = await applyLink(null);
      bridge.setSelection(selection);
      context.focusEditor();
      return removed;
    }
    const address = normalizeLinkAddress(result.values.address);
    if (askText) {
      const text = result.values.text?.trim() || address.replace(/^mailto:/i, "");
      bridge.setSelection(selection);
      handle.insertText(text);
      const after = bridge.getSelection();
      if (!after) return false;
      target = { anchor: { blockId: after.focus.blockId, offset: after.focus.offset - text.length }, focus: after.focus };
    }
    bridge.setSelection(target);
    const applied = await applyLink(address);
    if (askText || (existing && collapsed)) bridge.setSelection({ anchor: target.focus, focus: target.focus });
    context.focusEditor();
    return applied;
  };

  /** Opens the Suggestions & comments pane on its comments list. */
  const showComments = () => {
    const toggle = [...document.querySelectorAll<HTMLButtonElement>("#editor button")].find((button) => button.title.startsWith("Suggestions & comments"));
    if (toggle && !toggle.classList.contains("active")) toggle.click();
    [...document.querySelectorAll<HTMLElement>("#editor .cw-review-tab")].find((tab) => /comment/i.test(tab.textContent ?? ""))?.click();
  };

  /** Ctrl+Alt+M: a comment on the selection or the word at the caret. The first comment asks for the author name. */
  const newComment = async () => {
    if (!editable("add a comment")) return false;
    let selection = currentSelection();
    if (!selection) return false;
    if (isCollapsed(selection)) {
      const block = findParagraph(doc(), selection.focus.blockId);
      const word = block ? wordRangeAt(paragraphText(block), selection.focus.offset) : null;
      if (block && word) selection = { anchor: { blockId: block.id, offset: word.start }, focus: { blockId: block.id, offset: word.end } };
    }
    const askName = !context.author.chosen();
    const result = await openFormDialog({
      title: "New comment",
      icon: context.icons.comment,
      ...(askName ? {} : { message: `Commenting as ${context.author.name()}.` }),
      fields: [
        ...(askName ? [{ id: "author", label: "Your name", value: context.author.name(), hint: "Shown on your comments and suggestions. To change it later, search for “Author name” (Alt+Q).", maxLength: 80 }] : []),
        { id: "comment", label: "Comment", multiline: true, placeholder: "Add a comment", autofocus: true, hint: "Ctrl+Enter adds the comment." },
      ],
      submitLabel: "Comment",
      validate: (values) => {
        if (askName && !values.author.trim()) return { field: "author", message: "Type the name to show on your comments." };
        if (!values.comment.trim()) return { field: "comment", message: "Write the comment first." };
        return null;
      },
    });
    if (!result) {
      context.focusEditor();
      return false;
    }
    if (askName) context.author.choose(result.values.author);
    bridge.setSelection(selection);
    const block = findParagraph(doc(), selection.focus.blockId);
    const end = selection.anchor.blockId === selection.focus.blockId ? Math.max(selection.anchor.offset, selection.focus.offset) : selection.focus.offset;
    const style = runStyleAt(block, end) ?? block?.runs[0]?.style;
    const id = style ? handle.addComment([{ text: result.values.comment.trim(), style }]) : null;
    if (!id) notify("A comment can't be added here.", "error");
    else showComments();
    context.focusEditor();
    return Boolean(id);
  };

  /** The in-app answer to an engine prompt (Electron has no window.prompt). */
  const answerEnginePrompt = async (request: Extract<EngineDialogRequest, { kind: "prompt" }>): Promise<string | null> => {
    if (request.id === "hyperlink.insert" || request.id === "hyperlink.edit") {
      const result = await linkDialog({ address: request.defaultValue ?? "", editing: request.id === "hyperlink.edit", askText: false });
      if (!result) return null;
      return result.action === "remove" ? "" : normalizeLinkAddress(result.values.address);
    }
    if (request.id === "bookmark.add" || request.id === "bookmark.rename") {
      const renaming = request.id === "bookmark.rename";
      const existing = Object.keys(doc().bookmarks ?? {});
      const current = renaming ? request.defaultValue : undefined;
      const result = await openFormDialog({
        title: renaming ? "Rename bookmark" : "Add bookmark",
        icon: context.icons.bookmark,
        ...(renaming ? {} : { message: "The bookmark marks the selected text, or the caret position." }),
        fields: [{ id: "name", label: "Bookmark name", value: request.defaultValue ?? "", placeholder: "Chapter_one", maxLength: 40, hint: "Start with a letter; use letters, numbers and underscores." }],
        submitLabel: renaming ? "Rename" : "Add",
        validate: (values) => {
          const problem = bookmarkNameProblem(values.name, existing, current);
          return problem ? { field: "name", message: problem } : null;
        },
      });
      return result ? result.values.name.trim() : null;
    }
    if (request.id === "content-control.dropdown-items") {
      const result = await openFormDialog({
        title: "Drop-down list",
        icon: context.icons.list,
        message: "The choices people can pick in this drop-down list.",
        fields: [{ id: "items", label: "Choices", value: request.defaultValue ?? "", hint: "Separate the choices with commas." }],
        submitLabel: "Insert",
        validate: (values) => values.items.split(",").some((item) => item.trim()) ? null : { field: "items", message: "Type at least one choice." },
      });
      return result ? result.values.items : null;
    }
    const result = await openFormDialog({
      title: request.title || "Simple Docs",
      fields: [{ id: "value", label: request.message.replace(/:\s*$/, ""), value: request.defaultValue ?? "" }],
    });
    return result ? result.values.value : null;
  };

  const run = (command: ShortcutCommand): boolean | Promise<boolean> => {
    switch (command) {
      case "find": return openFind(false);
      case "replace": return openFind(true);
      case "link": return insertLink();
      case "comment": return newComment();
      case "heading-1": return applyStyle({ level: 1 });
      case "heading-2": return applyStyle({ level: 2 });
      case "heading-3": return applyStyle({ level: 3 });
      case "normal-style": return applyStyle("normal");
      case "clear-character-formatting": return clearCharacterFormatting();
      case "line-spacing-1": return setLineSpacing("1.0");
      case "line-spacing-1.5": return setLineSpacing("1.5");
      case "line-spacing-2": return setLineSpacing("2.0");
      case "change-case": return changeCase();
      case "undo": return context.editorShown() && bridge.undo();
      case "redo": return context.editorShown() && bridge.redo();
      case "save-as":
      case "close":
      case "command-search":
        return false;
      default:
        return context.editorShown() && clickRibbon(SHORTCUT_RIBBON_ITEMS[command]);
    }
  };

  /**
   * Adds "(Ctrl+L)"-style shortcut notes to the ribbon controls the shortcuts
   * drive, and aria-keyshortcuts for assistive technology. The footnote and
   * endnote tooltips stay as WordCanvas words them (UI scripts find those
   * buttons by their exact title); command search still shows their keys.
   */
  const annotateRibbon = (withHint: (title: string, combo: string) => string) => {
    for (const [command, id] of Object.entries(SHORTCUT_RIBBON_ITEMS) as Array<[ShortcutCommand, string]>) {
      const combo = shortcutFor(command);
      // Ctrl+Space keeps the paragraph style; the ribbon's "Clear all formatting" does not.
      if (!combo || command.startsWith("line-spacing") || command === "clear-character-formatting") continue;
      const element = ribbonElement(id);
      const button = element?.tagName === "BUTTON" ? element : element?.querySelector<HTMLElement>(":scope > button:first-child");
      if (!button) continue;
      button.setAttribute("aria-keyshortcuts", formatCombo(combo).replace(/Ctrl/g, "Control"));
      if (button.title && command !== "footnote" && command !== "endnote") button.title = withHint(button.title, combo);
    }
    const lineSpacing = ribbonElement(SHORTCUT_RIBBON_ITEMS["line-spacing-1"]!);
    if (lineSpacing && !lineSpacing.title.includes("Ctrl+1")) lineSpacing.title = `${lineSpacing.title} (Ctrl+1, Ctrl+5, Ctrl+2)`;
    const styleHints: Array<[string, string]> = [["heading 1", "Ctrl+Alt+1"], ["heading 2", "Ctrl+Alt+2"], ["heading 3", "Ctrl+Alt+3"], ["normal", "Ctrl+Shift+N"]];
    for (const card of styleCards()) {
      const name = card.querySelector(".name")?.textContent ?? "";
      const hint = styleHints.find(([style]) => sameName(name, style));
      if (hint && !card.title.includes(formatCombo(hint[1]))) card.title = withHint(card.title, hint[1]);
    }
  };

  return { run, answerEnginePrompt, annotateRibbon, insertLink, newComment, showRibbonTab };
}

export type EditorCommands = ReturnType<typeof createEditorCommands>;
