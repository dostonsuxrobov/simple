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
    // An edit of scanned text belongs to its line on its page (a copy is plain text).
    scan: edit.scan ? { ...edit.scan, lineRect: { ...edit.scan.lineRect } } : undefined,
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

/** A text box's height in its reading frame, the way the saver lays it out. */
function readingFrameHeight(edit: PageTextEdit, rect: PdfRect) {
  const angle = Number.isFinite(edit.angle) ? Number(edit.angle) : 0
  // The saver turns native runs (which keep their source rect) by their own
  // angle; boxes added in Simple are laid out in the rect as displayed.
  const rectAngle = edit.originalRect ? angle : 0
  const sideways = Math.abs(Math.sin(rectAngle + (Number(edit.displayRotation) || 0) * Math.PI / 180)) > 0.7
  return sideways ? rect.width : rect.height
}

/**
 * Give a text edit a new box (resize handles, the inspector's X/Y/W/H) without
 * moving its text. Native text keeps its PDF baseline as `baselineOffset`,
 * measured up from the box bottom, while the editor and the saver anchor the
 * first line to the box top; the offset therefore follows the box height so
 * the baseline keeps its distance from the top edge. A pure move changes
 * nothing.
 */
export function resizeTextEditRect(edit: PageTextEdit, rect: PdfRect): PageTextEdit {
  if (!Number.isFinite(edit.baselineOffset)) return { ...edit, rect }
  const growth = readingFrameHeight(edit, rect) - readingFrameHeight(edit, edit.rect)
  return { ...edit, rect, baselineOffset: Number(edit.baselineOffset) + growth }
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
    // Pasted or duplicated scanned text is a new text box: no retouch patch,
    // no recognised text to replace.
    scan: undefined,
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

/** The plain text written to the system clipboard alongside a payload. */
export function editClipboardPlainText(payload: EditClipboardPayload) {
  return payload.kind === 'text' ? payload.edit.text : payload.edit.label
}

function comparableClipboardText(text: string) {
  return text.replace(/\r\n?/g, '\n').trim()
}

export type EditPasteSource = 'internal' | 'text' | 'image' | 'none'

export interface EditPasteSourceInput {
  /** Plain text written to the system clipboard with the in-app payload, or null without one. */
  internalText: string | null
  /** False when the system clipboard could not be read at all. */
  systemReadable: boolean
  systemText: string
  systemHasImage: boolean
}

/**
 * Decide what Ctrl+V pastes. The in-app payload (a text box or image with its
 * styling) wins only while the system clipboard still holds the text Simple
 * wrote with it; anything copied afterwards elsewhere is newer and wins.
 */
export function chooseEditPasteSource(input: EditPasteSourceInput): EditPasteSource {
  const hasInternal = input.internalText !== null
  if (!input.systemReadable) return hasInternal ? 'internal' : 'none'
  const systemText = comparableClipboardText(input.systemText || '')
  if (hasInternal && systemText && systemText === comparableClipboardText(input.internalText || '')) return 'internal'
  if (input.systemHasImage) return 'image'
  if (systemText) return 'text'
  // Nothing pasteable arrived later (empty clipboard, files, unreadable
  // formats); keep the in-app payload useful.
  return hasInternal ? 'internal' : 'none'
}

/**
 * Size a new text box for pasted plain text, in displayed page units. Lines
 * longer than `maxWidth` wrap; the box never exceeds the page.
 */
export function pastedTextBoxSize(text: string, fontSize: number, maxWidth: number, maxHeight: number) {
  const lines = comparableClipboardText(text).split('\n')
  const averageGlyph = fontSize * 0.52
  const longest = Math.max(1, ...lines.map((line) => line.length))
  const naturalWidth = longest * averageGlyph + fontSize
  const width = clamp(naturalWidth, Math.min(maxWidth, fontSize * 6), Math.max(1, maxWidth))
  const charactersPerLine = Math.max(1, Math.floor((width - fontSize) / averageGlyph))
  const visualLines = lines.reduce((count, line) => count + Math.max(1, Math.ceil(line.length / charactersPerLine)), 0)
  const height = clamp(visualLines * fontSize * 1.25 + fontSize * 0.6, fontSize * 1.6, Math.max(1, maxHeight))
  return { width, height, wraps: naturalWidth > width }
}
