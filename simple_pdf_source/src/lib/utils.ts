export function cx(...values: Array<string | false | null | undefined>) {
  return values.filter(Boolean).join(' ')
}

export function clamp(value: number, min: number, max: number) {
  return Math.min(max, Math.max(min, value))
}

export function errorMessage(error: unknown) {
  // What main threw arrives wrapped ("Error invoking remote method 'x': CODE: message"): show the message.
  if (error instanceof Error) return error.message.replace(/^Error invoking remote method '[^']*': (?:[A-Za-z]*Error: )?/, '').replace(/^[A-Z][A-Z_]{2,}: /, '') || 'Something went wrong.'
  return String(error || 'Something went wrong.')
}

export function withoutExtension(name: string) {
  return name.replace(/\.pdf$/i, '')
}

export function pageRangeLabel(indices: number[]) {
  if (!indices.length) return 'No pages'
  if (indices.length === 1) return `Page ${indices[0] + 1}`
  return `${indices.length} pages`
}

export function isTypingTarget(target: EventTarget | null) {
  return target instanceof Element
    && Boolean(target.closest('input, textarea, select, [contenteditable="true"]'))
}
