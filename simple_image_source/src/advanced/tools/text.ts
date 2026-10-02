// src/advanced/tools/text.ts (WP5)
// Horizontal Type tool (T), Photoshop behaviour (design 5.9, 5.13):
//   - click places point text (the click is the first baseline: left, centre or right per alignment);
//     drag draws a paragraph box (text wraps at its width); click on an existing text layer edits it;
//   - typing happens in an IME-safe <textarea> laid over the layer (TextEditOverlay.tsx) while the layer's
//     pixels are hidden; the options bar's text style applies live to the text being edited;
//   - Ctrl+Enter, numpad Enter, a click outside the text, another tool or commitSession() commit; Esc
//     cancels; committing empty text creates nothing (or deletes the edited layer);
//   - new text is one history step ("Type Tool"), an edit is one step ("Edit Type"); the layer stays
//     editable. Commits load the font first, so the pixels match what was typed.
// The overlay module (React) is loaded on first use, so this file stays loadable without a DOM.
import type { Affine, Point } from '../../imaging/types.ts'
import type { LayerId, TextLayer, TextSpec, TextStyle, ToolContext, ToolPointerEvent, ViewTransform } from '../types.ts'
import { createTextLayer } from '../document.ts'
import { cssFont, layoutText } from '../../shared/vector.ts'
import { alignmentFactor, anchoredTranslation, rectFromPoints } from '../toolGeometry.ts'
import type { TextEditOverlayHandle, TextEditState } from '../TextEditOverlay.tsx'
import type { AdvancedToolController } from './shared.ts'
import {
  createJobTracker,
  createSpacePan,
  findLayer,
  isToolContext,
  layerIndexOf,
  optionsOf,
  reportError,
  screenToDocDistance,
  strokeContrastPath,
  toScreen,
} from './shared.ts'

type OverlayModule = typeof import('../TextEditOverlay.tsx')

let overlayModule: Promise<OverlayModule> | null = null

function loadOverlay(): Promise<OverlayModule> {
  if (!overlayModule) {
    overlayModule = import('../TextEditOverlay.tsx')
    overlayModule.catch(() => { overlayModule = null })
  }
  return overlayModule
}

interface TextSession {
  /** The edited layer, or null for new text. */
  readonly layerId: LayerId | null
  readonly original: TextSpec | null
  spec: TextSpec
  /** Point text: document position of the alignment point on the first baseline. */
  readonly anchor: Point | null
  readonly boxHeight: number | null
  overlay: TextEditOverlayHandle | null
  hidden: boolean
  committing: boolean
  viewKey: string
}

/** Photoshop names a type layer after its text. */
export function textLayerName(text: string): string {
  const first = String(text ?? '').split(/\r\n|\r|\n/).map((line) => line.trim()).find(Boolean) ?? ''
  const name = first.length > 30 ? `${first.slice(0, 30).trimEnd()}…` : first
  return name || 'Text'
}

function applyAffine(m: Affine, p: Point): Point {
  return { x: m[0] * p.x + m[2] * p.y + m[4], y: m[1] * p.x + m[3] * p.y + m[5] }
}

function invertAffinePoint(m: Affine, p: Point): Point | null {
  const det = m[0] * m[3] - m[1] * m[2]
  if (!(Math.abs(det) > 1e-12)) return null
  const x = p.x - m[4]
  const y = p.y - m[5]
  return { x: (m[3] * x - m[2] * y) / det, y: (-m[1] * x + m[0] * y) / det }
}

function sameStyle(a: TextStyle, b: TextStyle): boolean {
  return a.fontFamily === b.fontFamily && a.fontSize === b.fontSize && a.fontWeight === b.fontWeight && a.italic === b.italic
    && a.underline === b.underline && a.color.r === b.color.r && a.color.g === b.color.g && a.color.b === b.color.b
    && a.align === b.align && a.lineHeight === b.lineHeight && a.letterSpacing === b.letterSpacing
}

function sameSpec(a: TextSpec, b: TextSpec): boolean {
  return a.text === b.text && a.boxWidth === b.boxWidth && sameStyle(a.style, b.style) && a.transform.every((value, index) => value === b.transform[index])
}

/** The spec with new text or style; point text keeps its alignment point on the first baseline fixed. */
function respec(session: TextSession, text: string, style: TextStyle): TextSpec {
  const base = { ...session.spec, text, style }
  if (!session.anchor || base.boxWidth !== null) return base
  const layout = layoutText({ text, style, boxWidth: null })
  const baseline = layout.lines[0]?.baseline ?? 0
  const m = base.transform
  const t = anchoredTranslation([m[0], m[1], m[2], m[3]], session.anchor, { x: alignmentFactor(style.align) * layout.width, y: baseline })
  return { ...base, transform: [m[0], m[1], m[2], m[3], t.x, t.y] }
}

/** Corners of the text box in document space (paragraph box or laid-out point text). */
function textBoxCorners(spec: TextSpec, boxHeight: number | null): Point[] {
  const layout = layoutText(spec)
  const width = spec.boxWidth !== null && spec.boxWidth > 0 ? spec.boxWidth : Math.max(layout.width, 1)
  const height = Math.max(layout.height, boxHeight ?? 0)
  return [{ x: 0, y: 0 }, { x: width, y: 0 }, { x: width, y: height }, { x: 0, y: height }].map((p) => applyAffine(spec.transform, p))
}

function insideText(spec: TextSpec, boxHeight: number | null, point: Point, pad: number): boolean {
  const local = invertAffinePoint(spec.transform, point)
  if (!local) return false
  const layout = layoutText(spec)
  const width = spec.boxWidth !== null && spec.boxWidth > 0 ? spec.boxWidth : Math.max(layout.width, 1)
  const height = Math.max(layout.height, boxHeight ?? 0)
  return local.x >= -pad && local.y >= -pad && local.x <= width + pad && local.y <= height + pad
}

async function loadFonts(spec: TextSpec): Promise<void> {
  const fonts = (globalThis as { document?: { fonts?: FontFaceSet } }).document?.fonts
  if (!fonts || typeof fonts.load !== 'function') return
  try {
    await fonts.load(cssFont(spec.style), spec.text || 'A')
  } catch {
    // A missing font falls back to the next family.
  }
}

export function createTextTool(): AdvancedToolController {
  let ctx: ToolContext | null = null
  let session: TextSession | null = null
  let press: { start: Point; current: Point; startScreen: Point; dragged: boolean } | null = null
  let unsubscribeEditor: (() => void) | null = null
  const jobs = createJobTracker()
  // While text is open, Space-drag pans the canvas without closing it (when the canvas, not the field, has focus).
  const pan = createSpacePan(() => ctx, () => ctx?.setCursor('text'))

  const overlayState = (s: TextSession, view: ViewTransform): Omit<TextEditState, 'text'> => ({
    style: s.spec.style,
    boxWidth: s.spec.boxWidth,
    boxHeight: s.boxHeight,
    transform: s.spec.transform,
    view,
  })

  const close = (s: TextSession, context: ToolContext | null) => {
    if (session === s) session = null
    s.overlay?.destroy()
    s.overlay = null
    if (context) {
      context.setOverlayElement(null)
      if (s.hidden) context.compositor.setPreview(null)
      context.view.requestOverlay()
      // Only while the Type tool is still the active tool (a commit may finish after a tool switch).
      if (ctx === context) context.setCursor('text')
    }
  }

  const commit = (): void => {
    const s = session
    const context = ctx
    if (!s || !context || s.committing) return
    if (s.overlay?.isComposing()) s.overlay.blur()
    const text = s.overlay ? s.overlay.getText() : s.spec.text
    const spec = respec(s, text, s.spec.style)
    s.spec = spec
    s.committing = true
    jobs.run(async () => {
      try {
        if (s.layerId === null) {
          if (!text.trim()) {
            close(s, context)
            return
          }
          const layer = await createTextLayer(textLayerName(text), spec)
          context.store.transact('Type Tool', 'text', (tx) => {
            const active = layerIndexOf(tx.state, tx.state.activeLayerId)
            tx.insertLayer(layer, active >= 0 ? active + 1 : tx.state.layers.length)
            tx.setActiveLayer(layer.id)
          })
        } else {
          const live = findLayer(context.store.getState(), s.layerId)
          if (live && live.kind === 'text') {
            if (!text.trim()) {
              context.store.transact('Edit Type', 'text', (tx) => tx.removeLayer(live.id))
            } else if (!sameSpec(live.text, spec)) {
              await loadFonts(spec)
              const rename = live.name === textLayerName(live.text.text) ? textLayerName(text) : undefined
              context.store.transact('Edit Type', 'text', (tx) => tx.updateLayer(live.id, rename ? { text: spec, name: rename } : { text: spec }))
            }
          }
        }
        close(s, context)
      } catch (error) {
        // Keep the text open so nothing typed is lost; the message says what went wrong.
        s.committing = false
        reportError(context, error, 'The text could not be placed.')
        s.overlay?.focus()
      }
    }).catch(() => {})
  }

  const cancel = () => {
    const s = session
    if (!s || s.committing) return
    close(s, ctx)
  }

  const open = (s: TextSession) => {
    const context = ctx
    if (!context) return
    session = s
    if (s.hidden && s.layerId) context.compositor.setPreview({ kind: 'layer-props', layerId: s.layerId, visible: false })
    context.setCursor('text')
    context.view.requestOverlay()
    loadOverlay().then((module) => {
      if (session !== s || !ctx) return
      const view = ctx.view.getView()
      s.viewKey = `${view.zoom}:${view.offsetX}:${view.offsetY}`
      s.overlay = module.mountTextEditOverlay({ text: s.spec.text, ...overlayState(s, view) }, {
        onInput: (text) => {
          if (session !== s || !ctx) return
          s.spec = respec(s, text, s.spec.style)
          s.overlay?.update({ transform: s.spec.transform })
          ctx.view.requestOverlay()
        },
        onCommit: () => commit(),
        onCancel: () => cancel(),
      })
      ctx.setOverlayElement(s.overlay.element)
      s.overlay.focus()
    }).catch((error) => {
      if (ctx) reportError(ctx, error, 'The text editor could not be opened.')
      if (session === s) close(s, ctx)
    })
  }

  const startNew = (start: Point, end: Point | null) => {
    if (!ctx) return
    const style: TextStyle = { ...optionsOf(ctx).text }
    if (end) {
      const rect = rectFromPoints(start, end)
      const width = Math.max(8, rect.width)
      open({
        layerId: null,
        original: null,
        spec: { text: '', style, boxWidth: width, transform: [1, 0, 0, 1, rect.x, rect.y] },
        anchor: null,
        boxHeight: Math.max(rect.height, style.fontSize * style.lineHeight),
        overlay: null,
        hidden: false,
        committing: false,
        viewKey: '',
      })
      return
    }
    const layout = layoutText({ text: '', style, boxWidth: null })
    const baseline = layout.lines[0]?.baseline ?? style.fontSize
    const spec: TextSpec = { text: '', style, boxWidth: null, transform: [1, 0, 0, 1, start.x, start.y - baseline] }
    open({ layerId: null, original: null, spec, anchor: { x: start.x, y: start.y }, boxHeight: null, overlay: null, hidden: false, committing: false, viewKey: '' })
  }

  const startEdit = (layer: TextLayer) => {
    if (!ctx) return
    if (layer.locks.pixels) {
      ctx.host.notify(`"${layer.name}" is locked. Unlock it in the Layers panel to edit its text.`, 'error')
      return
    }
    const spec = layer.text
    const layout = layoutText(spec)
    const anchor = spec.boxWidth === null
      ? applyAffine(spec.transform, { x: alignmentFactor(spec.style.align) * layout.width, y: layout.lines[0]?.baseline ?? 0 })
      : null
    if (ctx.store.getState().activeLayerId !== layer.id) ctx.store.transact('Select Layer', 'layer', (tx) => tx.setActiveLayer(layer.id))
    // The options bar shows (and then edits) the style of the text being edited.
    ctx.editor.updateOptions('text', spec.style)
    open({ layerId: layer.id, original: spec, spec, anchor, boxHeight: null, overlay: null, hidden: true, committing: false, viewKey: '' })
  }

  const textLayerAt = (point: Point): TextLayer | null => {
    if (!ctx) return null
    const state = ctx.store.getState()
    const pad = screenToDocDistance(ctx, 4)
    for (let index = state.layers.length - 1; index >= 0; index -= 1) {
      const layer = state.layers[index]
      if (layer.kind !== 'text' || !layer.visible) continue
      if (insideText(layer.text, null, point, pad)) return layer
    }
    return null
  }

  return {
    id: 'text',
    activate(context: ToolContext): void {
      if (!isToolContext(context)) return
      ctx = context
      press = null
      context.setCursor('text')
      context.setHint('Click to type, drag to make a text box, click text to edit it. Ctrl+Enter commits, Esc cancels.')
      unsubscribeEditor = context.editor.subscribe(() => {
        const s = session
        if (!s || s.committing) return
        const style = optionsOf(context).text
        if (sameStyle(style, s.spec.style)) return
        s.spec = respec(s, s.overlay ? s.overlay.getText() : s.spec.text, { ...style })
        s.overlay?.update({ style: s.spec.style, transform: s.spec.transform })
        context.view.requestOverlay()
      })
    },
    deactivate(): void {
      // Typed text is never dropped: leaving the tool commits it.
      commit()
      unsubscribeEditor?.()
      unsubscribeEditor = null
      press = null
      ctx?.setHint(null)
      // The commit (if any) keeps its own reference to the context.
      ctx = null
    },
    pointerDown(event: ToolPointerEvent): void {
      if (!ctx || jobs.busy) return
      if (session && pan.pointerDown(event)) return
      if (event.button !== 0) return
      const s = session
      if (s) {
        if (s.committing) return
        if (insideText(s.spec, s.boxHeight, event.doc, screenToDocDistance(ctx, 6))) {
          s.overlay?.focus()
          return
        }
        commit()
        return
      }
      const hit = textLayerAt(event.doc)
      if (hit) {
        startEdit(hit)
        return
      }
      press = { start: { x: event.doc.x, y: event.doc.y }, current: { x: event.doc.x, y: event.doc.y }, startScreen: { x: event.screen.x, y: event.screen.y }, dragged: false }
    },
    pointerMove(event: ToolPointerEvent): void {
      if (pan.pointerMove(event)) return
      if (!ctx || !press) return
      press.current = { x: event.doc.x, y: event.doc.y }
      if (!press.dragged && Math.hypot(event.screen.x - press.startScreen.x, event.screen.y - press.startScreen.y) >= 4) press.dragged = true
      ctx.view.requestOverlay()
    },
    pointerUp(event: ToolPointerEvent): void {
      if (pan.pointerUp(event)) return
      if (!ctx || !press) return
      const current = press
      press = null
      current.current = { x: event.doc.x, y: event.doc.y }
      try {
        if (current.dragged) startNew(current.start, current.current)
        else startNew(current.start, null)
      } catch (error) {
        reportError(ctx, error, 'The text could not be started.')
      }
    },
    pointerCancel(): void {
      press = null
      pan.reset()
      ctx?.view.requestOverlay()
    },
    keyDown(event: KeyboardEvent): boolean {
      const s = session
      if (!s || event.isComposing) return false
      if (pan.keyDown(event)) return true
      if (event.key === 'Escape') {
        cancel()
        return true
      }
      if (event.key === 'Enter' && (event.ctrlKey || event.metaKey || event.code === 'NumpadEnter')) {
        commit()
        return true
      }
      return false
    },
    keyUp(event: KeyboardEvent): boolean {
      return pan.held ? pan.keyUp(event) : false
    },
    drawOverlay(context: CanvasRenderingContext2D, view: ViewTransform): void {
      const s = session
      if (s) {
        const key = `${view.zoom}:${view.offsetX}:${view.offsetY}`
        if (s.overlay && key !== s.viewKey) {
          s.viewKey = key
          s.overlay.update({ view })
        }
        strokeContrastPath(context, textBoxCorners(s.spec, s.boxHeight).map((p) => toScreen(view, p)), true, 3)
        return
      }
      if (press && press.dragged) {
        const r = rectFromPoints(press.start, press.current)
        strokeContrastPath(context, [
          { x: r.x, y: r.y }, { x: r.x + r.width, y: r.y }, { x: r.x + r.width, y: r.y + r.height }, { x: r.x, y: r.y + r.height },
        ].map((p) => toScreen(view, p)), true, 3)
      }
    },
    hasSession: () => session !== null,
    commitSession(): void {
      commit()
    },
    cancelSession(): void {
      cancel()
    },
    isBusy: () => jobs.busy,
    whenIdle: () => jobs.whenIdle(),
  }
}
