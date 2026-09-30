/** The same fit and wrapping decisions are used in the editor and PDF writer. */
export function resolveTextFit(edit) {
  if (edit.textFit === 'fit' || edit.textFit === 'wrap') return edit.textFit
  return edit.originalText && !edit.originalText.includes('\n') && !String(edit.text || '').includes('\n') ? 'fit' : 'wrap'
}

export function layoutText(text, width, measure, mode = 'wrap') {
  const paragraphs = String(text || '').replace(/\r\n?/g, '\n').split('\n')
  const available = Math.max(1, Number(width) || 1)
  if (mode === 'fit') {
    const widest = Math.max(0, ...paragraphs.map(measure))
    return { lines: paragraphs, fitScale: widest > available ? available / widest : 1 }
  }
  const lines = []
  for (const paragraph of paragraphs) {
    if (!paragraph) { lines.push(''); continue }
    let line = ''
    for (const token of paragraph.match(/\s+|\S+/gu) || []) {
      if (measure(line + token) <= available) { line += token; continue }
      if (/^\s+$/u.test(token)) {
        // Whitespace at a soft line break does not create an empty line.
        if (line) { lines.push(line.trimEnd()); line = '' }
        continue
      }
      if (line) { lines.push(line.trimEnd()); line = '' }
      // Long URLs and identifiers must fit too. Iterate Unicode code points
      // so a line break cannot split a surrogate pair.
      for (const character of token) {
        if (line && measure(line + character) > available) {
          lines.push(line)
          line = ''
        }
        line += character
      }
    }
    if (line || !lines.length) lines.push(line)
  }
  return { lines, fitScale: 1 }
}
