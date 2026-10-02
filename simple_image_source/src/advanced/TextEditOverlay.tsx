// src/advanced/TextEditOverlay.tsx (WP5)
// The Type tool's in-place editor: a transparent <textarea> laid exactly over the text layer (same font,
// size, line height, letter spacing, alignment and colour; the layer's transform and the view zoom applied
// as one CSS matrix), so typing, selection, clipboard, native undo and IME composition all behave like any
// text field. The text tool hides the layer's pixels while editing and rasterizes the result on commit.
//   - IME-safe: the textarea is uncontrolled; keys pressed while composing (event.isComposing, key code
//     229) are left to the IME, so Enter confirms a composition instead of committing.
//   - Ctrl+Enter or numpad Enter commits, Esc cancels. Host file shortcuts (Ctrl+S, Ctrl+O, Ctrl+P,
//     Alt+Shift+Ctrl+W) keep working (Save settles the editor, which commits the text first); every other
//     key stays inside the field so single-key tool shortcuts never fire while typing.
// The text tool mounts it with mountTextEditOverlay() and hands the element to ToolContext.setOverlayElement.
import { useLayoutEffect, useRef, useState } from 'react'
import type { CSSProperties, KeyboardEvent as ReactKeyboardEvent, MutableRefObject } from 'react'
import { createRoot } from 'react-dom/client'
import { flushSync } from 'react-dom'
import type { Affine } from '../imaging/types.ts'
import type { TextStyle, ViewTransform } from './types.ts'
import { cssFont, layoutText } from '../shared/vector.ts'

export interface TextEditState {
  /** Text the field starts with (the field owns the value afterwards). */
  readonly text: string
  readonly style: TextStyle
  /** Paragraph width (wrapping), or null for point text. */
  readonly boxWidth: number | null
  /** Paragraph box height drawn while editing (display only), or null. */
  readonly boxHeight: number | null
  /** Text-local to document space. */
  readonly transform: Affine
  /** Document to screen (CSS px relative to the canvas). */
  readonly view: ViewTransform
}

export interface TextEditCallbacks {
  /** Every change of the field's value (composition included). */
  onInput(text: string): void
  onCommit(): void
  onCancel(): void
}

export interface TextEditOverlayHandle {
  /** The element to mount above the canvas (absolutely positioned, covers it, pointer-transparent). */
  readonly element: HTMLElement
  update(next: Partial<Omit<TextEditState, 'text'>>): void
  focus(): void
  /** Ends an IME composition (the IME commits what it shows) by moving focus away. */
  blur(): void
  getText(): string
  isComposing(): boolean
  destroy(): void
}

interface OverlayProps extends TextEditState, TextEditCallbacks {
  readonly fieldRef: MutableRefObject<HTMLTextAreaElement | null>
  readonly composingRef: MutableRefObject<boolean>
}

function rgb(color: TextStyle['color']): string {
  const byte = (value: number) => Math.max(0, Math.min(255, Math.round(Number(value) || 0)))
  return `rgb(${byte(color.r)}, ${byte(color.g)}, ${byte(color.b)})`
}

function fontSizeOf(style: TextStyle): number {
  return Math.min(5000, Math.max(1, Number.isFinite(style.fontSize) ? style.fontSize : 48))
}

/** Field geometry in text-local px: the laid-out box plus room for the caret. */
export function textFieldSize(text: string, style: TextStyle, boxWidth: number | null, boxHeight: number | null): { readonly width: number; readonly height: number; readonly lineHeight: number } {
  const size = fontSizeOf(style)
  const lineHeight = size * (style.lineHeight > 0 ? style.lineHeight : 1.2)
  const layout = layoutText({ text: text || ' ', style, boxWidth })
  const width = boxWidth !== null && boxWidth > 0 ? boxWidth : Math.max(1, layout.width) + Math.max(4, size * 0.6)
  const height = Math.max(layout.height, lineHeight, boxHeight ?? 0)
  return { width, height, lineHeight }
}

/** CSS matrix placing text-local px on screen. */
export function textFieldMatrix(transform: Affine, view: ViewTransform): string {
  const [a, b, c, d, e, f] = transform
  const z = view.zoom
  const values = [a * z, b * z, c * z, d * z, e * z + view.offsetX, f * z + view.offsetY].map((value) => (Number.isFinite(value) ? value : 0))
  return `matrix(${values.join(', ')})`
}

export function TextEditOverlay(props: OverlayProps) {
  const { style, boxWidth, boxHeight, transform, view, fieldRef, composingRef } = props
  const [value, setValue] = useState(props.text)
  const callbacks = useRef<TextEditCallbacks>(props)
  callbacks.current = props

  // Keep the caret visible at the end of existing text when the editor opens.
  useLayoutEffect(() => {
    const field = fieldRef.current
    if (!field) return
    const end = field.value.length
    field.setSelectionRange(end, end)
  }, [fieldRef])

  const size = textFieldSize(value, style, boxWidth, boxHeight)
  const paragraph = boxWidth !== null && boxWidth > 0
  const color = rgb(style.color)
  const css: CSSProperties = {
    position: 'absolute',
    left: 0,
    top: 0,
    margin: 0,
    padding: 0,
    border: 0,
    outline: 'none',
    boxSizing: 'content-box',
    display: 'block',
    width: `${size.width}px`,
    height: `${size.height}px`,
    transform: textFieldMatrix(transform, view),
    transformOrigin: '0 0',
    font: cssFont(style),
    lineHeight: `${size.lineHeight}px`,
    letterSpacing: `${Number.isFinite(style.letterSpacing) ? style.letterSpacing : 0}px`,
    textAlign: style.align,
    textDecoration: style.underline ? 'underline' : 'none',
    color,
    caretColor: color,
    background: 'transparent',
    resize: 'none',
    overflow: 'hidden',
    whiteSpace: paragraph ? 'pre-wrap' : 'pre',
    overflowWrap: paragraph ? 'anywhere' : 'normal',
    pointerEvents: 'auto',
    tabSize: 4,
  }

  const onKeyDown = (event: ReactKeyboardEvent<HTMLTextAreaElement>) => {
    const native = event.nativeEvent
    if (native.isComposing || native.keyCode === 229 || composingRef.current) {
      event.stopPropagation()
      return
    }
    const control = event.ctrlKey || event.metaKey
    if (event.key === 'Escape') {
      event.preventDefault()
      event.stopPropagation()
      callbacks.current.onCancel()
      return
    }
    if (event.key === 'Enter' && (control || native.code === 'NumpadEnter')) {
      event.preventDefault()
      event.stopPropagation()
      callbacks.current.onCommit()
      return
    }
    const key = event.key.toLowerCase()
    // Host file shortcuts continue to the window (Save settles the editor, which commits this text).
    if (control && !event.altKey && (key === 's' || key === 'o' || key === 'p')) return
    if (control && event.altKey && event.shiftKey && key === 'w') return
    event.stopPropagation()
  }

  const onInput = (event: { currentTarget: HTMLTextAreaElement }) => {
    const text = event.currentTarget.value
    setValue(text)
    callbacks.current.onInput(text)
  }

  return (
    <textarea
      ref={fieldRef}
      className="advanced-text-edit"
      aria-label="Text"
      defaultValue={props.text}
      spellCheck={false}
      autoCapitalize="off"
      autoComplete="off"
      autoCorrect="off"
      wrap={paragraph ? 'soft' : 'off'}
      style={css}
      onInput={onInput}
      onKeyDown={onKeyDown}
      onKeyUp={(event) => event.stopPropagation()}
      onCompositionStart={() => { composingRef.current = true }}
      onCompositionEnd={(event) => {
        composingRef.current = false
        onInput(event)
      }}
      onPointerDown={(event) => event.stopPropagation()}
      onMouseDown={(event) => event.stopPropagation()}
      onDoubleClick={(event) => event.stopPropagation()}
      onContextMenu={(event) => event.stopPropagation()}
    />
  )
}

/** Mounts the editor in its own React root inside a fresh container element. */
export function mountTextEditOverlay(initial: TextEditState, callbacks: TextEditCallbacks): TextEditOverlayHandle {
  const container = document.createElement('div')
  container.className = 'advanced-text-edit-layer'
  Object.assign(container.style, { position: 'absolute', left: '0', top: '0', right: '0', bottom: '0', overflow: 'hidden', pointerEvents: 'none', zIndex: '5' })
  const root = createRoot(container)
  const fieldRef: MutableRefObject<HTMLTextAreaElement | null> = { current: null }
  const composingRef: MutableRefObject<boolean> = { current: false }
  let state: TextEditState = initial
  let destroyed = false
  const render = () => {
    root.render(<TextEditOverlay {...state} {...callbacks} fieldRef={fieldRef} composingRef={composingRef} />)
  }
  flushSync(render)
  return {
    element: container,
    update(next) {
      if (destroyed) return
      state = { ...state, ...next }
      render()
    },
    focus() {
      fieldRef.current?.focus({ preventScroll: true })
    },
    blur() {
      fieldRef.current?.blur()
    },
    getText() {
      return fieldRef.current ? fieldRef.current.value : state.text
    },
    isComposing() {
      return composingRef.current
    },
    destroy() {
      if (destroyed) return
      destroyed = true
      container.remove()
      // Unmount outside any render or commit that may be running.
      setTimeout(() => root.unmount(), 0)
    },
  }
}
