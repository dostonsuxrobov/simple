/**
 * The font box on the Home tab: Word's font name box with type-to-search. It replaces
 * the editor's plain font list with every font installed on this computer, each shown
 * in its own face, with recently used fonts and the document's fonts on top. A font
 * the document uses but this computer lacks keeps its name and is marked as shown
 * with a similar font, like Word.
 */
import { familyKey, fontSections, normalizeFamily, type FontCatalogView, type FontEntry } from "./font-list.ts";

export interface FontPickerOptions {
  /** The editor's font family list; it stays in the page (hidden) for the editor's own bookkeeping. */
  select: HTMLSelectElement;
  /** What the list offers right now. */
  view(): FontCatalogView;
  /** Loads the installed font list (once); the list refreshes when it arrives. */
  loadInstalled(): Promise<void>;
  /** Fonts the open document uses. */
  documentFamilies(): string[];
  /** Recently used fonts, most recent first. */
  recentFamilies(): string[];
  /** Applies the font to the selection (or the caret); resolves false when it could not. */
  apply(family: string, entry: FontEntry): Promise<boolean>;
  /** Called after the list closed by Enter, Escape or a choice (return focus to the document). */
  onDone(): void;
}

export interface FontPicker {
  readonly element: HTMLElement;
  /** Shows the font at the caret (when the box is not being edited). */
  setCurrent(family: string | null | undefined): void;
  open(): void;
  close(): void;
  refresh(): void;
  destroy(): void;
}

let sequence = 0;

function cssFamily(family: string): string {
  return `"${family.replace(/["\\]/g, "\\$&")}", "Segoe UI", system-ui, sans-serif`;
}

export function createFontPicker(options: FontPickerOptions): FontPicker {
  const id = `font-box-${++sequence}`;
  const box = document.createElement("div");
  box.className = "font-box";
  const input = document.createElement("input");
  input.type = "text";
  input.className = "font-box-input";
  input.id = `${id}-input`;
  input.spellcheck = false;
  input.autocomplete = "off";
  input.setAttribute("role", "combobox");
  input.setAttribute("aria-label", "Font");
  input.setAttribute("aria-autocomplete", "list");
  input.setAttribute("aria-expanded", "false");
  input.setAttribute("aria-controls", `${id}-list`);
  input.title = "Font (type to search)";
  const toggle = document.createElement("button");
  toggle.type = "button";
  toggle.className = "font-box-toggle";
  toggle.tabIndex = -1;
  toggle.setAttribute("aria-label", "Show fonts");
  toggle.innerHTML = '<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="m8 10 4 4 4-4"/></svg>';
  box.append(input, toggle);
  const list = document.createElement("div");
  list.className = "font-list";
  list.id = `${id}-list`;
  list.setAttribute("role", "listbox");
  list.setAttribute("aria-label", "Fonts");
  list.hidden = true;

  const width = options.select.getBoundingClientRect().width || Number.parseFloat(options.select.style.width) || 150;
  box.style.width = `${Math.max(120, Math.round(width))}px`;
  options.select.insertAdjacentElement("beforebegin", box);
  options.select.hidden = true;
  options.select.style.display = "none";
  options.select.setAttribute("aria-hidden", "true");
  options.select.tabIndex = -1;
  document.body.append(list);

  let current = "";
  let typed = false;
  let active = -1;
  let items: Array<{ element: HTMLElement; entry: FontEntry }> = [];
  const faces = new IntersectionObserver((entries) => {
    for (const entry of entries) {
      if (!entry.isIntersecting) continue;
      const element = entry.target as HTMLElement;
      const family = element.dataset.family;
      const name = element.querySelector<HTMLElement>(".font-option-name");
      if (family && name && !name.style.fontFamily) name.style.fontFamily = cssFamily(family);
      faces.unobserve(element);
    }
  }, { root: list });

  const isOpen = () => !list.hidden;

  const position = () => {
    const rect = box.getBoundingClientRect();
    const top = rect.bottom + 2;
    const available = window.innerHeight - top - 12;
    list.style.left = `${Math.round(Math.max(6, Math.min(rect.left, window.innerWidth - 300 - 6)))}px`;
    list.style.top = `${Math.round(top)}px`;
    list.style.maxHeight = `${Math.max(160, Math.min(420, available))}px`;
  };

  const setActive = (index: number, scroll = true) => {
    if (!items.length) {
      active = -1;
      input.removeAttribute("aria-activedescendant");
      return;
    }
    active = Math.max(0, Math.min(items.length - 1, index));
    items.forEach((item, position) => item.element.setAttribute("aria-selected", String(position === active)));
    const element = items[active].element;
    input.setAttribute("aria-activedescendant", element.id);
    if (scroll) element.scrollIntoView({ block: "nearest" });
  };

  const render = () => {
    const query = typed ? input.value : "";
    const sections = fontSections(options.view(), { query, recent: options.recentFamilies(), document: options.documentFamilies() });
    faces.disconnect();
    list.replaceChildren();
    items = [];
    let index = 0;
    for (const section of sections) {
      if (!section.fonts.length && section.title !== "Matching fonts") continue;
      const heading = document.createElement("div");
      heading.className = "font-list-heading";
      heading.setAttribute("role", "presentation");
      heading.textContent = section.title;
      list.append(heading);
      if (!section.fonts.length) {
        const empty = document.createElement("div");
        empty.className = "font-list-empty";
        empty.textContent = "No installed font has that name.";
        list.append(empty);
      }
      for (const entry of section.fonts) {
        const element = document.createElement("div");
        element.className = `font-option is-${entry.state}`;
        element.id = `${id}-option-${index++}`;
        element.setAttribute("role", "option");
        element.dataset.family = entry.family;
        const name = document.createElement("span");
        name.className = "font-option-name";
        name.textContent = entry.label;
        element.append(name);
        if (entry.state === "missing") {
          const note = document.createElement("span");
          note.className = "font-option-note";
          note.textContent = "not installed";
          element.append(note);
          element.title = `${entry.family} isn't installed on this computer. Simple Docs shows it with a similar font.`;
        } else if (familyKey(entry.family) === familyKey(current)) {
          element.classList.add("is-current");
        }
        list.append(element);
        items.push({ element, entry });
        faces.observe(element);
      }
    }
    const currentIndex = items.findIndex((item) => familyKey(item.entry.family) === familyKey(current));
    setActive(typed ? 0 : currentIndex >= 0 ? currentIndex : 0, isOpen());
  };

  const open = () => {
    if (isOpen()) return;
    list.hidden = false;
    input.setAttribute("aria-expanded", "true");
    box.classList.add("is-open");
    position();
    render();
    void options.loadInstalled().then(() => {
      if (isOpen()) render();
    }).catch(() => {});
  };

  const close = (restore = true) => {
    if (!isOpen()) return;
    list.hidden = true;
    input.setAttribute("aria-expanded", "false");
    input.removeAttribute("aria-activedescendant");
    box.classList.remove("is-open");
    faces.disconnect();
    if (restore) {
      typed = false;
      show(current);
    }
  };

  const show = (family: string) => {
    const view = options.view();
    const key = familyKey(family);
    const missing = Boolean(key) && !view.ready.has(key) && !view.installed.some((name) => familyKey(name) === key) && !view.builtin.some((item) => familyKey(item.family) === key);
    input.value = family;
    box.classList.toggle("is-missing", missing);
    input.title = missing ? `${family} isn't installed on this computer. Simple Docs shows it with a similar font.` : "Font (type to search)";
  };

  const choose = async (index: number) => {
    const item = items[index];
    if (!item) return;
    close(false);
    typed = false;
    const applied = await options.apply(item.entry.family, item.entry).catch(() => false);
    if (applied) current = item.entry.family;
    show(current);
    options.onDone();
  };

  input.addEventListener("focus", () => {
    input.select();
  });
  input.addEventListener("mousedown", (event) => {
    event.stopPropagation();
    if (document.activeElement === input && !isOpen()) open();
  });
  input.addEventListener("click", () => {
    if (!isOpen()) open();
  });
  input.addEventListener("input", () => {
    typed = true;
    if (!isOpen()) open();
    else render();
  });
  input.addEventListener("keydown", (event) => {
    if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      if (!isOpen()) open();
      else setActive(active + (event.key === "ArrowDown" ? 1 : -1));
    } else if (event.key === "PageDown" || event.key === "PageUp") {
      if (!isOpen()) return;
      event.preventDefault();
      setActive(active + (event.key === "PageDown" ? 10 : -10));
    } else if (event.key === "Enter") {
      event.preventDefault();
      if (isOpen() && active >= 0) void choose(active);
      else if (!isOpen()) open();
    } else if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      close();
      options.onDone();
    } else if (event.key === "Tab") {
      close();
    }
  });
  input.addEventListener("blur", () => {
    // A click in the list keeps focus in the box (its mousedown is prevented).
    window.setTimeout(() => {
      if (document.activeElement !== input) close();
    }, 0);
  });
  toggle.addEventListener("mousedown", (event) => {
    event.preventDefault();
    event.stopPropagation();
  });
  toggle.addEventListener("click", () => {
    if (isOpen()) {
      close();
      return;
    }
    input.focus();
    open();
  });
  list.addEventListener("mousedown", (event) => event.preventDefault());
  list.addEventListener("click", (event) => {
    const option = event.target instanceof Element ? event.target.closest<HTMLElement>(".font-option") : null;
    if (!option) return;
    const index = items.findIndex((item) => item.element === option);
    if (index >= 0) void choose(index);
  });
  list.addEventListener("mousemove", (event) => {
    const option = event.target instanceof Element ? event.target.closest<HTMLElement>(".font-option") : null;
    const index = option ? items.findIndex((item) => item.element === option) : -1;
    if (index >= 0 && index !== active) setActive(index, false);
  });
  const outside = (event: Event) => {
    if (!isOpen()) return;
    const target = event.target as Node | null;
    if (target && (box.contains(target) || list.contains(target))) return;
    close();
  };
  const reposition = () => {
    if (isOpen()) position();
  };
  const windowBlur = () => close();
  document.addEventListener("pointerdown", outside, true);
  window.addEventListener("resize", reposition);
  window.addEventListener("blur", windowBlur);

  return {
    element: box,
    setCurrent(family) {
      const name = normalizeFamily(family);
      if (!name || name === current) return;
      current = name;
      if (document.activeElement !== input && !isOpen()) show(current);
    },
    open() {
      input.focus();
      open();
    },
    close: () => close(),
    refresh() {
      if (isOpen()) render();
      else show(current);
    },
    destroy() {
      faces.disconnect();
      document.removeEventListener("pointerdown", outside, true);
      window.removeEventListener("resize", reposition);
      window.removeEventListener("blur", windowBlur);
      list.remove();
      box.remove();
      options.select.hidden = false;
      options.select.style.display = "";
    },
  };
}
