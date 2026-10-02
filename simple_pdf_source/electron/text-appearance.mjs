const normalize = text => String(text || '').normalize('NFKC').replace(/\s/g, '')
const sameColor = (a, b) => a && b && a.length === 3 && b.length === 3 && a.every((n, i) => Number.isFinite(n) && n === b[i])
// Render modes 3 (invisible) and 7 (clip only) paint nothing: OCR text layers.
const invisibleMode = mode => mode === 3 || mode === 7
// A font whose glyphs are this invisible is an OCR layer font as a whole.
const INVISIBLE_FONT_SHARE = 0.98

// Text paint per font, in content order: the normalised characters with their
// paint color and whether they were drawn invisibly.
function paintStreams(list, OPS) {
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
      if (!stream) streams.set(state.font, stream = { text: '', colors: [], hidden: [], hiddenCount: 0, cursor: 0, matches: new Map() })
      const color = state.alpha === 1 && state.mode === 0 && !state.mask && !state.transfer && !state.unsafeGroup && state.blend === 'source-over' && state.color?.length === 3 ? state.color : null
      const hidden = invisibleMode(state.mode)
      for (const glyph of glyphs) {
        if (typeof glyph !== 'object' || !glyph) continue
        const text = normalize(glyph.unicode)
        stream.text += text
        for (let i = 0; i < text.length; i++) {
          stream.colors.push(color)
          stream.hidden.push(hidden)
        }
        if (hidden) stream.hiddenCount += text.length
      }
    }
  }
  return streams
}

// The paint color of a matched run, or null. Text extraction can omit
// off-page/clipped glyphs present in the paint list. A sequential match alone
// can therefore pick a hidden label's color. Trust a repeated label only when
// every candidate agrees.
function matchedColor(stream, index, needle) {
  if (!stream.matches.has(needle)) {
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
  return stream.matches.get(needle)
}

// Each call consumes the item's characters in content order, so a lookup must
// be called once per text item, in text-content order.
function paintLookup(list, OPS) {
  const streams = paintStreams(list, OPS)
  return (fontName, text) => {
    const stream = streams.get(fontName), needle = normalize(text)
    if (!stream) return { color: undefined, invisible: false }
    const wholeFontInvisible = stream.text.length > 0 && stream.hiddenCount >= stream.text.length * INVISIBLE_FONT_SHARE
    if (!needle) return { color: undefined, invisible: wholeFontInvisible }
    const index = stream.text.indexOf(needle, stream.cursor)
    if (index < 0) return { color: undefined, invisible: wholeFontInvisible }
    stream.cursor = index + needle.length
    let invisible = true
    for (let i = index; i < index + needle.length && invisible; i++) invisible = stream.hidden[i]
    const color = matchedColor(stream, index, needle)
    return { color: color ? [...color] : undefined, invisible: wholeFontInvisible || invisible }
  }
}

// Read the original paint of each pdf.js text item, not antialiased screen
// pixels. PDF.js has already converted PDF color spaces to RGB in the display
// operator list. Text is matched in content order per font; ambiguous,
// multicolor and transparent runs return no color so callers keep their
// canvas-sampling fallback.
//
// `invisible` marks text that paints nothing (render mode 3 or 7), such as an
// OCR layer over a scan: every item of a font that is at least 98% invisible,
// otherwise an item whose matched characters were all drawn invisibly.
export function textPaintResolver(list, OPS) {
  const lookup = paintLookup(list, OPS)
  return (fontName, text) => {
    const { color, invisible } = lookup(fontName, text)
    return color && !invisible ? { color, invisible } : { invisible }
  }
}

export function textColorResolver(list, OPS) {
  const lookup = paintLookup(list, OPS)
  return (fontName, text) => lookup(fontName, text).color
}
