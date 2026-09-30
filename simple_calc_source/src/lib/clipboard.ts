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
