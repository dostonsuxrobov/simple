import { tokenizeFormulaText } from './formula-editing'

/**
 * A cut is destructive, so clearing must be gated on an affirmative clipboard
 * result. Rejections are treated as failed copies as an additional safety net.
 */
export async function clearAfterSuccessfulCopy(
  copy: () => Promise<boolean>,
  clear: () => void,
): Promise<boolean> {
  let copied = false
  try {
    copied = await copy()
  } catch {
    return false
  }
  if (!copied) return false
  clear()
  return true
}

/** Both flavours spreadsheet apps exchange; `html` is optional. */
export interface RichClipboardPayload {
  text: string
  html?: string
  /** Excel's "XML Spreadsheet" flavour (formulas); only the Electron bridge can read it. */
  excelXml?: string
  /** A copied picture as a PNG data URL (only when there is no text). */
  image?: string
}

/**
 * Optional Electron bridge (see electron/clipboard-bridge.cjs) exposed by the preload as
 * `window.simpleCalc.clipboard`. It reads/writes the OS clipboard directly, so HTML is never
 * sanitised by Chromium and Excel's <style> classes survive.
 */
interface ClipboardBridge {
  write?: (payload: RichClipboardPayload) => Promise<boolean> | boolean
  read?: () => Promise<Partial<RichClipboardPayload>> | Partial<RichClipboardPayload>
}

function clipboardBridge(): ClipboardBridge | null {
  const host = (globalThis as { simpleCalc?: { clipboard?: ClipboardBridge } }).simpleCalc
  return host?.clipboard ?? null
}

/**
 * Writes through a one-shot `copy` event (document.execCommand). This is synchronous, keeps
 * the HTML byte-for-byte (no sanitiser) and needs no clipboard permission.
 */
function writeViaCopyEvent(payload: RichClipboardPayload): boolean {
  if (typeof document === 'undefined' || typeof document.execCommand !== 'function') return false
  let handled = false
  const listener = (event: ClipboardEvent) => {
    if (!event.clipboardData) return
    event.clipboardData.setData('text/plain', payload.text)
    if (payload.html) event.clipboardData.setData('text/html', payload.html)
    event.preventDefault()
    event.stopImmediatePropagation()
    handled = true
  }
  document.addEventListener('copy', listener, true)
  try {
    return document.execCommand('copy') && handled
  } catch {
    return false
  } finally {
    document.removeEventListener('copy', listener, true)
  }
}

/**
 * Puts text/plain + text/html on the system clipboard. Tries, in order: the Electron bridge,
 * a synthetic copy event, the async Clipboard API with a ClipboardItem, and finally
 * writeText. Resolves true only when something was written.
 */
export async function writeRichClipboard(payload: RichClipboardPayload): Promise<boolean> {
  const bridge = clipboardBridge()
  if (bridge?.write) {
    try {
      if (await bridge.write({ text: payload.text, html: payload.html })) return true
    } catch {
      // Fall through to the web clipboard.
    }
  }
  if (payload.html && writeViaCopyEvent(payload)) return true
  const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard
  if (payload.html && clipboard?.write && typeof ClipboardItem !== 'undefined') {
    try {
      await clipboard.write([new ClipboardItem({
        'text/plain': new Blob([payload.text], { type: 'text/plain' }),
        'text/html': new Blob([payload.html], { type: 'text/html' }),
      })])
      return true
    } catch {
      // Some hosts reject text/html; plain text still beats nothing.
    }
  }
  if (clipboard?.writeText) {
    try {
      await clipboard.writeText(payload.text)
      return true
    } catch {
      return false
    }
  }
  return payload.html ? false : writeViaCopyEvent(payload)
}

async function blobText(item: ClipboardItem, type: string): Promise<string | undefined> {
  if (!item.types.includes(type)) return undefined
  try {
    return await (await item.getType(type)).text()
  } catch {
    return undefined
  }
}

/**
 * Reads text/plain and (when present) text/html. Prefers the Electron bridge, then the
 * unsanitised async Clipboard API (Chromium 122+), then the sanitised one, then readText.
 * Rejects only when no clipboard access is possible at all.
 */
export async function readRichClipboard(): Promise<RichClipboardPayload> {
  const bridge = clipboardBridge()
  if (bridge?.read) {
    try {
      const data = await bridge.read()
      if (data && (typeof data.text === 'string' || typeof data.html === 'string')) {
        const payload: RichClipboardPayload = { text: typeof data.text === 'string' ? data.text : '', html: data.html || undefined }
        if (typeof data.excelXml === 'string' && data.excelXml) payload.excelXml = data.excelXml
        if (typeof data.image === 'string' && data.image.startsWith('data:image/')) payload.image = data.image
        return payload
      }
    } catch {
      // Fall through to the web clipboard.
    }
  }
  const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard
  if (clipboard?.read) {
    const attempts: Array<() => Promise<ClipboardItems>> = [
      () => (clipboard.read as (options?: { unsanitized?: string[] }) => Promise<ClipboardItems>)({ unsanitized: ['text/html'] }),
      () => clipboard.read(),
    ]
    for (const attempt of attempts) {
      try {
        const items = await attempt()
        let text: string | undefined
        let html: string | undefined
        let image: string | undefined
        for (const item of items) {
          text ??= await blobText(item, 'text/plain')
          html ??= await blobText(item, 'text/html')
          if (image === undefined && item.types.includes('image/png')) {
            try {
              const blob = await item.getType('image/png')
              image = await new Promise<string>((resolve, reject) => {
                const reader = new FileReader()
                reader.onload = () => resolve(String(reader.result))
                reader.onerror = () => reject(reader.error)
                reader.readAsDataURL(blob)
              })
            } catch {
              // No picture flavour after all.
            }
          }
        }
        if (text !== undefined || html !== undefined) return { text: text ?? '', html: html || undefined }
        if (image) return { text: '', image }
      } catch {
        // Try the next strategy.
      }
    }
  }
  if (!clipboard?.readText) throw new Error('Clipboard access is not available.')
  return { text: await clipboard.readText() }
}

/** Raw (unsanitised) flavours from a native `paste` event — the highest-fidelity source. */
export function clipboardPayloadFromEvent(event: Pick<ClipboardEvent, 'clipboardData'>): RichClipboardPayload | null {
  const data = event.clipboardData
  if (!data) return null
  const html = data.getData('text/html')
  return { text: data.getData('text/plain'), html: html || undefined }
}

/** Fills a native `copy`/`cut` event; returns false when the event has no clipboardData. */
export function writeClipboardEvent(event: Pick<ClipboardEvent, 'clipboardData' | 'preventDefault'>, payload: RichClipboardPayload): boolean {
  if (!event.clipboardData) return false
  event.clipboardData.setData('text/plain', payload.text)
  if (payload.html) event.clipboardData.setData('text/html', payload.html)
  event.preventDefault()
  return true
}

/** What the internal clipboard remembers about where a copy or cut came from. */
export interface ClipboardSourceStamp {
  /** Document the cells were copied from (null when unknown). */
  documentId: string | null
  sheetId: string
  /** Identity of the source sheet's cell map at copy time. */
  cells: unknown
}

/**
 * Cut mode (and the copy marquee) only stays valid while the copied cells are exactly where
 * they were: same document, same sheet, and an untouched cell map. Any edit, insert, delete,
 * sort, undo or newly opened workbook replaces the cell map, so Ctrl+V can no longer move
 * the wrong rectangle, as Excel cancels cut mode on those changes.
 */
export function clipboardSourceIntact(
  stamp: ClipboardSourceStamp | null | undefined,
  current: { documentId: string | null; sheets: ReadonlyArray<{ id: string; cells: unknown }> } | null | undefined,
): boolean {
  if (!stamp || !current) return false
  if (stamp.documentId !== current.documentId) return false
  const sheet = current.sheets.find((item) => item.id === stamp.sheetId)
  return Boolean(sheet) && sheet!.cells === stamp.cells
}

function columnNumber(label: string) {
  let value = 0
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64
  return value - 1
}

function columnText(index: number) {
  let value = index + 1
  let label = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    label = String.fromCharCode(65 + remainder) + label
    value = Math.floor((value - 1) / 26)
  }
  return label
}

function quoteSheet(name: string) {
  if (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) && !/^R\d*C\d*$/i.test(name)) return name
  return `'${name.replace(/'/g, "''")}'`
}

function shiftCellText(text: string, rowDelta: number, colDelta: number) {
  const match = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(text)
  if (!match) return text
  const column = columnNumber(match[2]) + colDelta
  const row = Number(match[4]) + rowDelta
  if (column < 0 || row < 1 || column > 16_383 || row > 1_048_576) return '#REF!'
  return `${match[1]}${columnText(column)}${match[3]}${row}`
}

export interface MovedFormulaPlacement {
  /** Sheet the formula lived on (and the block was cut from). */
  sourceSheet: string
  /** Sheet the block was pasted on. */
  destinationSheet: string
  /** 0-based bounds of the cut block on the source sheet. */
  rect: { top: number; left: number; bottom: number; right: number }
  rowDelta: number
  colDelta: number
}

/**
 * Rewrites a formula that travels with a cut block to another sheet, the way Excel does:
 * references inside the block follow it (and stay unqualified, since the block now lives on
 * the formula's new sheet), while every other reference that implicitly meant the source
 * sheet gains an explicit `Source!` prefix so it keeps pointing at the same cells.
 * References qualified with another sheet are left alone.
 */
export function relocateMovedFormula(formula: string, placement: MovedFormulaPlacement): string {
  const text = `=${formula}`
  const source = placement.sourceSheet.toLocaleLowerCase()
  const { rect } = placement
  let output = ''
  let last = 0
  for (const token of tokenizeFormulaText(text)) {
    const reference = token.reference
    if (token.kind !== 'reference' || !reference) continue
    const qualified = reference.sheet !== undefined
    if (qualified && reference.sheet!.toLocaleLowerCase() !== source) continue
    const bang = reference.text.lastIndexOf('!')
    const body = reference.text.slice(bang + 1)
    const wholeLine = reference.bottom - reference.top >= 1_048_575 || reference.right - reference.left >= 16_383
    const inside = !wholeLine && reference.top >= rect.top && reference.bottom <= rect.bottom && reference.left >= rect.left && reference.right <= rect.right
    let replacement: string
    if (inside) {
      const spill = body.endsWith('#') ? '#' : ''
      const moved = body.replace(/#$/, '').split(':').map((part) => shiftCellText(part, placement.rowDelta, placement.colDelta)).join(':')
      replacement = `${qualified ? `${quoteSheet(placement.destinationSheet)}!` : ''}${moved}${spill}`
    } else {
      if (qualified) continue
      replacement = `${quoteSheet(placement.sourceSheet)}!${body}`
    }
    output += text.slice(last, token.start) + replacement
    last = token.end
  }
  if (!last) return formula
  return (output + text.slice(last)).slice(1)
}
