import type { PageObjectEdit, PageTextEdit, PdfRect } from '../types'

export const EDIT_CLIPBOARD_VERSION = 1 as const
export const EDIT_PASTE_OFFSET = 12

export interface TextEditClipboardPayload {
  version: typeof EDIT_CLIPBOARD_VERSION
  kind: 'text'
  edit: PageTextEdit
}

export interface ObjectEditClipboardPayload {
  version: typeof EDIT_CLIPBOARD_VERSION
  kind: 'object'
  edit: PageObjectEdit
}

export type EditClipboardPayload = TextEditClipboardPayload | ObjectEditClipboardPayload

export type PastedPageEdit =
  | { kind: 'text'; edit: PageTextEdit }
  | { kind: 'object'; edit: PageObjectEdit }

function clamp(value: number, minimum: number, maximum: number) {
  return Math.min(maximum, Math.max(minimum, value))
}

function finite(value: number, fallback: number) {
  return Number.isFinite(value) ? value : fallback
}

function cloneRect(rect: PdfRect): PdfRect {
  return {
    x: finite(rect.x, 0),
    y: finite(rect.y, 0),
    width: Math.max(1, finite(rect.width, 1)),
    height: Math.max(1, finite(rect.height, 1)),
  }
}

function cloneTextEdit(edit: PageTextEdit): PageTextEdit {
  const text = String(edit.text ?? '')
  const caretOffset = Number.isFinite(edit.caretOffset)
    ? clamp(Number(edit.caretOffset), 0, text.length)
    : text.length
  return {
    ...edit,
    rect: cloneRect(edit.rect),
    originalRect: edit.originalRect ? cloneRect(edit.originalRect) : undefined,
    sourceItemRect: edit.sourceItemRect ? cloneRect(edit.sourceItemRect) : undefined,
    color: [...edit.color] as [number, number, number],
    backgroundColor: edit.backgroundColor
      ? [...edit.backgroundColor] as [number, number, number]
      : undefined,
    fontData: edit.fontData?.slice(),
    text,
    originalText: String(edit.originalText ?? ''),
    modified: false,
    caretOffset,
  }
}

function cloneObjectEdit(edit: PageObjectEdit): PageObjectEdit {
  return {
    ...edit,
    rect: cloneRect(edit.rect),
    originalRect: edit.originalRect ? cloneRect(edit.originalRect) : undefined,
    opacity: clamp(finite(edit.opacity, 1), 0, 1),
    modified: false,
  }
}

/** Create an independent, normalized clipboard snapshot of a text selection. */
export function copyTextEdit(edit: PageTextEdit): TextEditClipboardPayload {
  return {
    version: EDIT_CLIPBOARD_VERSION,
    kind: 'text',
    edit: cloneTextEdit(edit),
  }
}

/** Create an independent, normalized clipboard snapshot of an image/artwork selection. */
export function copyObjectEdit(edit: PageObjectEdit): ObjectEditClipboardPayload {
  return {
    version: EDIT_CLIPBOARD_VERSION,
    kind: 'object',
    edit: cloneObjectEdit(edit),
  }
}

/** Convenience helper for the App's mutually-exclusive edit selections. */
export function copyEditSelection(
  textEdit: PageTextEdit | null,
  objectEdit: PageObjectEdit | null,
): EditClipboardPayload | null {
  if (textEdit) return copyTextEdit(textEdit)
  if (objectEdit) return copyObjectEdit(objectEdit)
  return null
}

export function cloneEditClipboardPayload(payload: TextEditClipboardPayload): TextEditClipboardPayload
export function cloneEditClipboardPayload(payload: ObjectEditClipboardPayload): ObjectEditClipboardPayload
export function cloneEditClipboardPayload(payload: EditClipboardPayload): EditClipboardPayload
export function cloneEditClipboardPayload(payload: EditClipboardPayload): EditClipboardPayload {
  return payload.kind === 'text' ? copyTextEdit(payload.edit) : copyObjectEdit(payload.edit)
}

/**
 * Move a copied box right and down by `offset` PDF points, then keep the entire
 * box inside the supplied page bounds. PDF coordinates grow upward, so moving
 * visually downward subtracts from y.
 */
export function pasteRectWithinPage(
  source: PdfRect,
  pageBounds: PdfRect,
  offset = EDIT_PASTE_OFFSET,
): PdfRect {
  const bounds = cloneRect(pageBounds)
  const rect = cloneRect(source)
  const width = Math.min(rect.width, bounds.width)
  const height = Math.min(rect.height, bounds.height)
  const safeOffset = finite(offset, EDIT_PASTE_OFFSET)
  return {
    x: clamp(rect.x + safeOffset, bounds.x, bounds.x + bounds.width - width),
    y: clamp(rect.y - safeOffset, bounds.y, bounds.y + bounds.height - height),
    width,
    height,
  }
}

export function pasteTextEditToPage(
  payload: TextEditClipboardPayload,
  pageIndex: number,
  pageBounds: PdfRect,
  offset = EDIT_PASTE_OFFSET,
): PageTextEdit {
  const copied = cloneTextEdit(payload.edit)
  return {
    ...copied,
    overlayId: undefined,
    pageIndex,
    rect: pasteRectWithinPage(copied.rect, pageBounds, offset),
    originalRect: undefined,
    originalText: '',
    sourceItemText: undefined,
    sourceItemRect: undefined,
    sourceSelectionStart: undefined,
    sourceSelectionEnd: undefined,
    sourceSpaceWidth: undefined,
    cover: false,
    modified: true,
    caretOffset: copied.text.length,
    selectionStart: copied.text.length,
    selectionEnd: copied.text.length,
  }
}

export function pasteObjectEditToPage(
  payload: ObjectEditClipboardPayload,
  pageIndex: number,
  pageBounds: PdfRect,
  offset = EDIT_PASTE_OFFSET,
): PageObjectEdit {
  const copied = cloneObjectEdit(payload.edit)
  return {
    ...copied,
    overlayId: undefined,
    candidateId: undefined,
    pageIndex,
    rect: pasteRectWithinPage(copied.rect, pageBounds, offset),
    originalRect: undefined,
    cover: false,
    modified: true,
  }
}

export function pasteEditToPage(
  payload: EditClipboardPayload,
  pageIndex: number,
  pageBounds: PdfRect,
  offset = EDIT_PASTE_OFFSET,
): PastedPageEdit {
  return payload.kind === 'text'
    ? { kind: 'text', edit: pasteTextEditToPage(payload, pageIndex, pageBounds, offset) }
    : { kind: 'object', edit: pasteObjectEditToPage(payload, pageIndex, pageBounds, offset) }
}
