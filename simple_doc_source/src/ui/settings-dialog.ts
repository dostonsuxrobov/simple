/**
 * A small in-app settings dialog (numbers with units, choices, check boxes), used for
 * the Paragraph and Page number dialogs. It shares the app's .modal-backdrop/.modal
 * look and its modal guards (focus trap, blocked shortcuts), like form-dialog.ts.
 *
 * A field can be left blank (null): a number field with no value, a choice showing
 * "—" or an indeterminate check box. That is how a selection with mixed formatting is
 * shown, and blank fields are left unchanged when the dialog is applied.
 */

export type SettingsValue = string | number | boolean | null;
export type SettingsValues = Record<string, SettingsValue>;

interface ControlBase {
  id: string;
  label: string;
  hint?: string;
  /** Hides the control (the dialog can show it again from onChange). */
  hidden?: boolean;
}

export type SettingsControl =
  | (ControlBase & { kind: "number"; value: number | null; unit?: string; min?: number; max?: number; step?: number })
  | (ControlBase & { kind: "select"; value: string | null; options: ReadonlyArray<{ value: string; label: string }> })
  | (ControlBase & { kind: "checkbox"; value: boolean | null });

export interface SettingsSection {
  title?: string;
  controls: SettingsControl[];
}

export interface SettingsDialogApi {
  values(): SettingsValues;
  set(id: string, value: SettingsValue): void;
  setUnit(id: string, unit: string): void;
  setLabel(id: string, label: string): void;
  setHidden(id: string, hidden: boolean): void;
  setNumberLimits(id: string, limits: { min?: number; max?: number; step?: number }): void;
}

export interface SettingsDialogOptions {
  title: string;
  /** Raw SVG markup shown above the title. */
  icon?: string;
  message?: string;
  sections: SettingsSection[];
  submitLabel?: string;
  /** An extra action left of Cancel ("Remove page numbers"). The dialog resolves with its id. */
  extraAction?: { id: string; label: string };
  /** Called after a field changed (by the person), to keep related fields consistent. */
  onChange?: (changed: string, dialog: SettingsDialogApi) => void;
  /** A problem to show (and the field to fix), or null when the values can be applied. */
  validate?: (values: SettingsValues) => { field?: string; message: string } | null;
  /** Runs after the dialog closed and before the promise settles (focus restoration). */
  onClosed?: (result: SettingsDialogResult) => void;
}

/** "submit" or the extra action id with the values; null when cancelled. */
export type SettingsDialogResult = { action: string; values: SettingsValues } | null;

let active: { close(result: SettingsDialogResult): void } | null = null;
let sequence = 0;

export function isSettingsDialogOpen(): boolean {
  return active !== null;
}

/** Cancels the open settings dialog, if any (for example when the window closes). */
export function cancelSettingsDialog(): void {
  active?.close(null);
}

/** A number the way the field shows it: up to two decimals, no trailing zeros. */
export function formatNumber(value: number): string {
  return String(Math.round(value * 100) / 100);
}

/** A typed number ("1.5", "1,5", " 12 "): null when blank, NaN when it is not a number. */
export function parseNumber(text: string): number | null {
  const value = String(text ?? "").trim().replace(/\s+/g, "").replace(",", ".");
  if (!value) return null;
  return /^[-+]?(?:\d+\.?\d*|\.\d+)$/.test(value) ? Number(value) : Number.NaN;
}

export function openSettingsDialog(options: SettingsDialogOptions, root: HTMLElement = document.body): Promise<SettingsDialogResult> {
  active?.close(null);
  const id = `settings-dialog-${++sequence}`;
  return new Promise((resolve) => {
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop form-dialog-backdrop settings-dialog-backdrop";
    const form = document.createElement("form");
    form.className = "modal form-dialog settings-dialog";
    form.noValidate = true;
    form.setAttribute("role", "dialog");
    form.setAttribute("aria-modal", "true");
    form.setAttribute("aria-labelledby", `${id}-title`);
    if (options.icon) {
      const icon = document.createElement("div");
      icon.className = "modal-icon";
      icon.innerHTML = options.icon;
      form.append(icon);
    }
    const title = document.createElement("h2");
    title.id = `${id}-title`;
    title.textContent = options.title;
    form.append(title);
    if (options.message) {
      const message = document.createElement("p");
      message.id = `${id}-message`;
      message.className = "form-dialog-message";
      message.textContent = options.message;
      form.setAttribute("aria-describedby", message.id);
      form.append(message);
    }

    type Entry = { control: SettingsControl; element: HTMLInputElement | HTMLSelectElement; row: HTMLElement; caption: HTMLElement; unit?: HTMLElement };
    const entries = new Map<string, Entry>();
    for (const section of options.sections) {
      const fieldset = document.createElement("fieldset");
      fieldset.className = "settings-section";
      if (section.title) {
        const legend = document.createElement("legend");
        legend.textContent = section.title;
        fieldset.append(legend);
      }
      const grid = document.createElement("div");
      grid.className = "settings-grid";
      for (const control of section.controls) {
        const row = document.createElement("label");
        row.className = `settings-field is-${control.kind}`;
        const caption = document.createElement("span");
        caption.className = "settings-label";
        caption.textContent = control.label;
        let element: HTMLInputElement | HTMLSelectElement;
        let unit: HTMLElement | undefined;
        if (control.kind === "select") {
          const select = document.createElement("select");
          if (control.value === null) {
            const blank = new Option("—", "");
            select.add(blank);
          }
          for (const option of control.options) select.add(new Option(option.label, option.value));
          select.value = control.value ?? "";
          element = select;
          row.append(caption, select);
        } else if (control.kind === "checkbox") {
          const input = document.createElement("input");
          input.type = "checkbox";
          input.checked = control.value === true;
          input.indeterminate = control.value === null;
          element = input;
          row.append(input, caption);
        } else {
          // A text field (not type=number), so "1,5" works where the comma is the decimal
          // separator; the arrow keys still step the value like a spin box.
          const input = document.createElement("input");
          input.type = "text";
          input.inputMode = "decimal";
          input.autocomplete = "off";
          input.spellcheck = false;
          input.setAttribute("role", "spinbutton");
          if (control.min !== undefined) input.dataset.min = String(control.min);
          if (control.max !== undefined) input.dataset.max = String(control.max);
          input.dataset.step = String(control.step ?? 1);
          input.value = control.value === null ? "" : formatNumber(control.value);
          if (control.value === null) input.placeholder = "—";
          input.addEventListener("keydown", (event) => {
            if (event.key !== "ArrowUp" && event.key !== "ArrowDown") return;
            event.preventDefault();
            const current = parseNumber(input.value);
            const step = Number(input.dataset.step) || 1;
            const min = input.dataset.min === undefined ? -Infinity : Number(input.dataset.min);
            const max = input.dataset.max === undefined ? Infinity : Number(input.dataset.max);
            const base = current === null || Number.isNaN(current) ? (Number.isFinite(min) ? Math.max(min, 0) : 0) : current;
            const next = Math.min(max, Math.max(min, Math.round((base + (event.key === "ArrowUp" ? step : -step)) / step) * step));
            input.value = formatNumber(next);
            input.dispatchEvent(new Event("input", { bubbles: true }));
            input.dispatchEvent(new Event("change", { bubbles: true }));
          });
          element = input;
          const box = document.createElement("span");
          box.className = "settings-input";
          box.append(input);
          unit = document.createElement("span");
          unit.className = "settings-unit";
          unit.textContent = control.unit ?? "";
          box.append(unit);
          row.append(caption, box);
        }
        element.id = `${id}-${control.id}`;
        element.name = control.id;
        if (control.hint) {
          const hint = document.createElement("small");
          hint.className = "form-field-hint";
          hint.id = `${element.id}-hint`;
          hint.textContent = control.hint;
          element.setAttribute("aria-describedby", hint.id);
          row.append(hint);
        }
        row.hidden = control.hidden === true;
        entries.set(control.id, { control, element, row, caption, unit });
        grid.append(row);
      }
      fieldset.append(grid);
      form.append(fieldset);
    }
    const error = document.createElement("p");
    error.className = "form-dialog-error";
    error.setAttribute("role", "alert");
    error.hidden = true;
    form.append(error);

    const actions = document.createElement("div");
    actions.className = "modal-actions";
    const button = (label: string, type: "button" | "submit", className = "") => {
      const element = document.createElement("button");
      element.type = type;
      element.textContent = label;
      if (className) element.className = className;
      return element;
    };
    let extra: HTMLButtonElement | null = null;
    if (options.extraAction) {
      extra = button(options.extraAction.label, "button", "form-dialog-extra");
      actions.append(extra);
    }
    const cancel = button("Cancel", "button");
    const submit = button(options.submitLabel ?? "OK", "submit", "modal-primary");
    actions.append(cancel, submit);
    form.append(actions);
    backdrop.append(form);

    const readValue = (entry: Entry): SettingsValue => {
      const { control, element } = entry;
      if (control.kind === "checkbox") return (element as HTMLInputElement).indeterminate ? null : (element as HTMLInputElement).checked;
      if (control.kind === "select") return element.value === "" ? null : element.value;
      return parseNumber(element.value);
    };
    const api: SettingsDialogApi = {
      values: () => Object.fromEntries([...entries].map(([name, entry]) => [name, readValue(entry)])),
      set(name, value) {
        const entry = entries.get(name);
        if (!entry) return;
        const { control, element } = entry;
        if (control.kind === "checkbox") {
          (element as HTMLInputElement).indeterminate = value === null;
          (element as HTMLInputElement).checked = value === true;
        } else if (control.kind === "select") {
          const select = element as HTMLSelectElement;
          if (value === null && ![...select.options].some((option) => option.value === "")) select.add(new Option("—", ""), 0);
          select.value = value === null ? "" : String(value);
        } else {
          element.value = typeof value === "number" && Number.isFinite(value) ? formatNumber(value) : "";
        }
      },
      setUnit(name, text) {
        const entry = entries.get(name);
        if (entry?.unit) entry.unit.textContent = text;
      },
      setLabel(name, text) {
        const entry = entries.get(name);
        if (entry) entry.caption.textContent = text;
      },
      setHidden(name, hidden) {
        const entry = entries.get(name);
        if (entry) entry.row.hidden = hidden;
      },
      setNumberLimits(name, limits) {
        const entry = entries.get(name);
        if (!entry || entry.control.kind !== "number") return;
        const element = entry.element as HTMLInputElement;
        if (limits.min !== undefined) element.dataset.min = String(limits.min);
        if (limits.max !== undefined) element.dataset.max = String(limits.max);
        if (limits.step !== undefined) element.dataset.step = String(limits.step);
      },
    };

    const close = (result: SettingsDialogResult) => {
      if (active !== handle) return;
      active = null;
      backdrop.remove();
      try {
        options.onClosed?.(result);
      } finally {
        resolve(result);
      }
    };
    const handle = { close };
    const showProblem = (problem: { field?: string; message: string }) => {
      error.textContent = problem.message;
      error.hidden = false;
      for (const entry of entries.values()) entry.element.removeAttribute("aria-invalid");
      const entry = (problem.field && entries.get(problem.field)) || null;
      if (entry) {
        entry.element.setAttribute("aria-invalid", "true");
        entry.element.focus();
      }
    };
    const trySubmit = () => {
      const values = api.values();
      const invalid = [...entries].find(([, entry]) => entry.control.kind === "number" && !entry.row.hidden && Number.isNaN(values[entry.control.id]));
      if (invalid) {
        showProblem({ field: invalid[0], message: `Type a number for ${invalid[1].control.label.toLowerCase()}.` });
        return;
      }
      const outside = [...entries].find(([, entry]) => {
        const value = values[entry.control.id];
        if (entry.control.kind !== "number" || entry.row.hidden || typeof value !== "number") return false;
        const element = entry.element as HTMLInputElement;
        return (element.dataset.min !== undefined && value < Number(element.dataset.min)) || (element.dataset.max !== undefined && value > Number(element.dataset.max));
      });
      if (outside) {
        const element = outside[1].element as HTMLInputElement;
        showProblem({ field: outside[0], message: `${outside[1].control.label} must be between ${element.dataset.min ?? "−∞"} and ${element.dataset.max ?? "∞"}.` });
        return;
      }
      const problem = options.validate?.(values) ?? null;
      if (problem) showProblem(problem);
      else close({ action: "submit", values });
    };
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      trySubmit();
    });
    const changed = (event: Event) => {
      const target = event.target as HTMLElement | null;
      const name = target && "name" in target ? String((target as HTMLInputElement).name) : "";
      if (target instanceof HTMLInputElement && target.type === "checkbox") target.indeterminate = false;
      if (!error.hidden) {
        error.hidden = true;
        for (const entry of entries.values()) entry.element.removeAttribute("aria-invalid");
      }
      if (name && entries.has(name)) options.onChange?.(name, api);
    };
    form.addEventListener("change", changed);
    form.addEventListener("input", (event) => {
      if (!error.hidden) changed(event);
    });
    cancel.addEventListener("click", () => close(null));
    extra?.addEventListener("click", () => close({ action: options.extraAction!.id, values: api.values() }));
    backdrop.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        close(null);
      }
    });
    // Clicking outside the dialog never dismisses typed values by accident.
    backdrop.addEventListener("pointerdown", (event) => {
      if (event.target === backdrop) event.preventDefault();
    });

    active = handle;
    root.append(backdrop);
    const first = [...entries.values()].find((entry) => !entry.row.hidden);
    if (first) {
      first.element.focus();
      if (first.element instanceof HTMLInputElement && first.control.kind === "number") first.element.select();
    } else {
      submit.focus();
    }
  });
}
