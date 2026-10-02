/**
 * A small in-app dialog with one or more text fields (never a native prompt).
 * It reuses the app's .modal-backdrop/.modal look, so main.ts's modal guards
 * (focus trap, blocked shortcuts) apply to it. One dialog is open at a time.
 */

export interface FormField {
  id: string;
  label: string;
  value?: string;
  placeholder?: string;
  /** A multi-line field: Enter adds a line, Ctrl+Enter submits. */
  multiline?: boolean;
  hint?: string;
  maxLength?: number;
  spellcheck?: boolean;
  /** Put the caret here first (default: the first field). */
  autofocus?: boolean;
}

export interface FormDialogOptions {
  title: string;
  message?: string;
  /** Raw SVG markup shown above the title. */
  icon?: string;
  fields: readonly FormField[];
  submitLabel?: string;
  cancelLabel?: string;
  /** An extra action left of Cancel, such as "Remove link". The dialog resolves with its id. */
  extraAction?: { id: string; label: string };
  /** A problem to show (and the field to fix), or null when the values can be submitted. */
  validate?: (values: Record<string, string>) => { field?: string; message: string } | null;
  /** Runs after the dialog closed and before the promise settles (focus restoration). */
  onClosed?: (result: FormDialogResult) => void;
}

/** "submit" or the extra action id with the field values; null when cancelled. */
export type FormDialogResult = { action: string; values: Record<string, string> } | null;

let active: { close(result: FormDialogResult): void } | null = null;
let sequence = 0;

export function isFormDialogOpen(): boolean {
  return active !== null;
}

/** Cancels the open dialog, if any (for example when the window closes). */
export function cancelFormDialog(): void {
  active?.close(null);
}

export function openFormDialog(options: FormDialogOptions, root: HTMLElement = document.body): Promise<FormDialogResult> {
  active?.close(null);
  const id = `form-dialog-${++sequence}`;
  return new Promise((resolve) => {
    const backdrop = document.createElement("div");
    backdrop.className = "modal-backdrop form-dialog-backdrop";
    const form = document.createElement("form");
    form.className = "modal form-dialog";
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

    const controls = new Map<string, HTMLInputElement | HTMLTextAreaElement>();
    const fields = document.createElement("div");
    fields.className = "form-dialog-fields";
    for (const field of options.fields) {
      const label = document.createElement("label");
      label.className = "form-field";
      const caption = document.createElement("span");
      caption.className = "form-field-label";
      caption.textContent = field.label;
      const control = field.multiline ? document.createElement("textarea") : document.createElement("input");
      control.id = `${id}-${field.id}`;
      control.name = field.id;
      control.value = field.value ?? "";
      control.spellcheck = field.spellcheck ?? Boolean(field.multiline);
      control.autocomplete = "off";
      if (control instanceof HTMLInputElement) control.type = "text";
      else control.rows = 4;
      if (field.placeholder) control.placeholder = field.placeholder;
      if (field.maxLength) control.maxLength = field.maxLength;
      label.append(caption, control);
      if (field.hint) {
        const hint = document.createElement("small");
        hint.className = "form-field-hint";
        hint.id = `${control.id}-hint`;
        hint.textContent = field.hint;
        control.setAttribute("aria-describedby", hint.id);
        label.append(hint);
      }
      controls.set(field.id, control);
      fields.append(label);
    }
    form.append(fields);
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
    const cancel = button(options.cancelLabel ?? "Cancel", "button");
    const submit = button(options.submitLabel ?? "OK", "submit", "modal-primary");
    actions.append(cancel, submit);
    form.append(actions);
    backdrop.append(form);

    const values = () => Object.fromEntries([...controls].map(([name, control]) => [name, control.value]));
    const close = (result: FormDialogResult) => {
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
      for (const control of controls.values()) control.removeAttribute("aria-invalid");
      const control = (problem.field && controls.get(problem.field)) || controls.values().next().value;
      if (control) {
        control.setAttribute("aria-invalid", "true");
        control.focus();
      }
    };
    const trySubmit = () => {
      const current = values();
      const problem = options.validate?.(current) ?? null;
      if (problem) showProblem(problem);
      else close({ action: "submit", values: current });
    };
    form.addEventListener("submit", (event) => {
      event.preventDefault();
      trySubmit();
    });
    form.addEventListener("input", () => {
      if (error.hidden) return;
      error.hidden = true;
      for (const control of controls.values()) control.removeAttribute("aria-invalid");
    });
    cancel.addEventListener("click", () => close(null));
    extra?.addEventListener("click", () => close({ action: options.extraAction!.id, values: values() }));
    backdrop.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        event.stopPropagation();
        close(null);
      } else if (event.key === "Enter" && (event.ctrlKey || event.metaKey) && event.target instanceof HTMLTextAreaElement) {
        event.preventDefault();
        trySubmit();
      }
    });
    // Clicking outside the dialog never dismisses typed text by accident.
    backdrop.addEventListener("pointerdown", (event) => {
      if (event.target === backdrop) event.preventDefault();
    });

    active = handle;
    root.append(backdrop);
    const first = options.fields.find((field) => field.autofocus) ?? options.fields[0];
    const control = first ? controls.get(first.id) : null;
    if (control) {
      control.focus();
      control.select();
    } else {
      submit.focus();
    }
  });
}
