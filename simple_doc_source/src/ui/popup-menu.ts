/**
 * A small drop-down menu under a ribbon button (Header, Footer): items, check items
 * and separators, with arrow-key navigation. One menu is open at a time; a press
 * outside, Escape or a choice closes it.
 */

export interface PopupMenuItem {
  label: string;
  /** A check item (Different first page): true shows the check mark. */
  checked?: boolean;
  disabled?: boolean;
  /** Shown as the item's tooltip. */
  hint?: string;
  run?: () => void;
}

export type PopupMenuEntry = PopupMenuItem | "separator";

let openMenu: { close(focusAnchor?: boolean): void } | null = null;

export function closePopupMenu(): void {
  openMenu?.close(false);
}

export function isPopupMenuOpen(): boolean {
  return openMenu !== null;
}

export function openPopupMenu(anchor: HTMLElement, entries: readonly PopupMenuEntry[], options: { label?: string; onClose?: (chosen: boolean) => void } = {}): { close(): void } {
  openMenu?.close(false);
  const menu = document.createElement("div");
  menu.className = "simple-menu";
  menu.setAttribute("role", "menu");
  if (options.label) menu.setAttribute("aria-label", options.label);
  const buttons: HTMLButtonElement[] = [];
  for (const entry of entries) {
    if (entry === "separator") {
      const separator = document.createElement("div");
      separator.className = "simple-menu-separator";
      separator.setAttribute("role", "separator");
      menu.append(separator);
      continue;
    }
    const button = document.createElement("button");
    button.type = "button";
    button.setAttribute("role", entry.checked === undefined ? "menuitem" : "menuitemcheckbox");
    if (entry.checked !== undefined) button.setAttribute("aria-checked", String(entry.checked));
    button.disabled = Boolean(entry.disabled || !entry.run);
    if (entry.hint) button.title = entry.hint;
    const check = document.createElement("span");
    check.className = "simple-menu-check";
    check.setAttribute("aria-hidden", "true");
    check.textContent = entry.checked ? "✓" : "";
    const label = document.createElement("span");
    label.textContent = entry.label;
    button.append(check, label);
    button.addEventListener("click", () => {
      close(false, true);
      entry.run?.();
    });
    menu.append(button);
    buttons.push(button);
  }
  document.body.append(menu);
  const rect = anchor.getBoundingClientRect();
  const width = menu.offsetWidth;
  const height = menu.offsetHeight;
  menu.style.left = `${Math.round(Math.max(6, Math.min(rect.left, window.innerWidth - width - 6)))}px`;
  menu.style.top = `${Math.round(rect.bottom + height + 8 > window.innerHeight ? Math.max(6, rect.top - height - 2) : rect.bottom + 2)}px`;
  anchor.setAttribute("aria-expanded", "true");

  const enabled = () => buttons.filter((button) => !button.disabled);
  const onKey = (event: KeyboardEvent) => {
    const items = enabled();
    const index = items.indexOf(document.activeElement as HTMLButtonElement);
    if (event.key === "Escape") {
      event.preventDefault();
      event.stopPropagation();
      close(true);
    } else if (event.key === "ArrowDown" || event.key === "ArrowUp") {
      event.preventDefault();
      const step = event.key === "ArrowDown" ? 1 : -1;
      items[(index + step + items.length) % items.length]?.focus();
    } else if (event.key === "Home" || event.key === "End") {
      event.preventDefault();
      (event.key === "Home" ? items[0] : items[items.length - 1])?.focus();
    } else if (event.key === "Tab") {
      close(false);
    }
  };
  const outside = (event: Event) => {
    const target = event.target as Node | null;
    if (target && (menu.contains(target) || anchor.contains(target))) return;
    close(false);
  };
  let closed = false;
  function close(focusAnchor = false, chosen = false) {
    if (closed) return;
    closed = true;
    if (openMenu === handle) openMenu = null;
    menu.remove();
    anchor.setAttribute("aria-expanded", "false");
    document.removeEventListener("pointerdown", outside, true);
    window.removeEventListener("blur", blur);
    window.removeEventListener("resize", blur);
    if (focusAnchor) anchor.focus();
    options.onClose?.(chosen);
  }
  const blur = () => close(false);
  const handle = { close: (focusAnchor?: boolean) => close(focusAnchor) };
  menu.addEventListener("keydown", onKey);
  menu.addEventListener("mousedown", (event) => event.preventDefault());
  document.addEventListener("pointerdown", outside, true);
  window.addEventListener("blur", blur);
  window.addEventListener("resize", blur);
  openMenu = handle;
  enabled()[0]?.focus();
  return { close: () => close(false) };
}
