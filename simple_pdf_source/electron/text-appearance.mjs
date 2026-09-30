const normalize = text => String(text || '').normalize('NFKC').replace(/\s/g, '')
const sameColor = (a, b) => a && b && a.length === 3 && b.length === 3 && a.every((n, i) => Number.isFinite(n) && n === b[i])

// Read the original paint colors, not antialiased screen pixels. PDF.js has
// already converted PDF color spaces to RGB in the display operator list.
// Match text in content order per font; ambiguous/multicolor/transparent
// runs deliberately retain the existing canvas sampling fallback.
export function textColorResolver(list, OPS) {
  const streams = new Map()
  const stack = []
  let state = { font: '', color: [0, 0, 0], alpha: 1, mode: 0, mask: false, blend: 'source-over', transfer: false, unsafeGroup: false }
  for (let index = 0; index < list.fnArray.length; index += 1) {
    const op = list.fnArray[index], args = list.argsArray[index]
    if (op === OPS.save || op === OPS.paintFormXObjectBegin) stack.push({ ...state })
    else if (op === OPS.beginGroup) {
      stack.push({ ...state })
      state = { ...state, alpha: 1, blend: 'source-over', mask: false,
        unsafeGroup: state.unsafeGroup || state.alpha !== 1 || state.mask || state.blend !== 'source-over' || Boolean(args?.[0]?.smask) }
    } else if (op === OPS.restore || op === OPS.paintFormXObjectEnd || op === OPS.endGroup) state = stack.pop() || state
    else if (op === OPS.setFont) state.font = args[0]
    else if (op === OPS.setTextRenderingMode) state.mode = args[0]
    else if (op === OPS.setFillRGBColor) state.color = Array.from(args).slice(0, 3).map(n => n / 255)
    else if (op === OPS.setFillGray) state.color = [args[0], args[0], args[0]]
    else if (op === OPS.setFillColorN || op === OPS.setFillColor || op === OPS.setFillCMYKColor) state.color = null
    else if (op === OPS.setGState) {
      for (const [key, value] of args[0] || []) {
        if (key === 'ca') state.alpha = value
        else if (key === 'SMask') state.mask = Boolean(value)
        else if (key === 'BM') state.blend = value
        else if (key === 'TR') state.transfer = Boolean(value)
        else if (key === 'Font') state.font = value?.[0] || ''
      }
    } else if (op === OPS.showText || op === OPS.showSpacedText || op === OPS.nextLineShowText || op === OPS.nextLineSetSpacingShowText) {
      const glyphs = args?.find(Array.isArray)
      if (!glyphs) continue
      let stream = streams.get(state.font)
      if (!stream) streams.set(state.font, stream = { text: '', colors: [], cursor: 0, matches: new Map() })
      const color = state.alpha === 1 && state.mode === 0 && !state.mask && !state.transfer && !state.unsafeGroup && state.blend === 'source-over' && state.color?.length === 3 ? state.color : null
      for (const glyph of glyphs) {
        if (typeof glyph !== 'object' || !glyph) continue
        const text = normalize(glyph.unicode)
        stream.text += text
        for (let i = 0; i < text.length; i++) stream.colors.push(color)
      }
    }
  }
  return (fontName, text) => {
    const stream = streams.get(fontName), needle = normalize(text)
    if (!stream || !needle) return undefined
    const index = stream.text.indexOf(needle, stream.cursor)
    if (index < 0) return undefined
    stream.cursor = index + needle.length
    if (!stream.matches.has(needle)) {
      // Text extraction can omit off-page/clipped glyphs present in the paint
      // list. A sequential match alone can therefore pick a hidden label's
      // color. Trust a repeated label only when every candidate agrees.
      let color = stream.colors[index]
      let position = stream.text.indexOf(needle)
      while (color && position >= 0) {
        for (let i = position; i < position + needle.length; i++) {
          if (!sameColor(color, stream.colors[i])) { color = null; break }
        }
        position = stream.text.indexOf(needle, position + 1)
      }
      stream.matches.set(needle, color)
    }
    const color = stream.matches.get(needle)
    return color ? [...color] : undefined
  }
}
