// src/advanced/CanvasView.tsx (WP6)
// The Advanced canvas (design 5.6, 5.9, 5.15): the display the compositor draws into, an overlay canvas and
// a host for DOM overlays (the text editor), plus everything between the pointer and the tools.
//   - ViewModel (createViewModel): the view transform (zoom, offset) and the viewport size; implements the
//     ViewController tools use. "Fit on screen" keeps fitting while the window resizes until the user zooms
//     or pans; panning can never lose the image (constrainView); 100% and integer zooms snap to device pixels.
//   - Pointer input: one pointer at a time is routed to the active controller (the free-transform session
//     first) with document coordinates, pen pressure, modifier keys and the coalesced samples
//     (getCoalescedEvents) since the previous event, so fast strokes keep every sample. The middle button
//     pans in every tool; two touch points pinch-zoom and pan (the tool gesture is cancelled first).
//   - Wheel: Ctrl / Alt / trackpad pinch zoom at the cursor; plain wheel and Shift+wheel scroll.
//   - Overlay: a thin document outline, the marching ants of the selection (traced at the display level,
//     cached per selection version, animated at 10 fps only while a selection exists and the page is
//     visible) and the active controller's own overlay (crop and transform handles, brush ring, lasso path).
//     Brush-ring overlays disappear when the pointer leaves the canvas.
//   - CanvasSurface: cursor, DOM-overlay element and focus for the tool context, usable before mount.
// Keyboard input is not handled here: the host forwards keys to AdvancedEditor (one window listener).
import { useEffect, useLayoutEffect, useRef } from 'react'
import type { PointerEvent as ReactPointerEvent } from 'react'
import type { IntRect, MaskBuffer, Point, Size } from '../imaging/types.ts'
import type { DocumentStore, Selection, ToolController, ToolPointerEvent, ViewController, ViewTransform, ViewportSize } from './types.ts'
import type { AdvancedCompositor } from './compositor.ts'
import {
  actualPixelsView,
  centerView,
  clampZoom,
  constrainView,
  displayScale,
  docToScreen,
  fitView,
  panView,
  screenToDoc,
  snapView,
  viewsEqual,
  zoomForScale,
  zoomViewAt,
} from './viewport.ts'
import { isSelectAll } from './selection.ts'
import { cropMask, downsampleMask, traceOutline } from '../imaging/mask.ts'

// ---------------------------------------------------------------------------------------------
// View model
// ---------------------------------------------------------------------------------------------

/** Space around the document when it is fitted on screen (CSS px). */
export const FIT_PADDING = 24

export interface ViewModel extends ViewController {
  subscribe(listener: () => void): () => void
  /** Overlay redraw requests (ViewController.requestOverlay). */
  onOverlay(listener: () => void): () => void
  setViewport(size: ViewportSize): void
  /** The document changed size (crop, canvas size, rotate): refit or keep the image reachable. */
  documentResized(): void
  /** True while the view follows Fit on Screen. */
  readonly fitted: boolean
  /** True once the viewport has a measured size. */
  readonly measured: boolean
}

export function createViewModel(getDocSize: () => Size): ViewModel {
  let view: ViewTransform = { zoom: 1, offsetX: 0, offsetY: 0 }
  let viewport: ViewportSize = { width: 0, height: 0, dpr: 1 }
  let fitted = true
  // Photoshop opens a document fitted but never above 100%; View > Fit on Screen (Ctrl+0) also enlarges.
  let enlarge = false
  const listeners = new Set<() => void>()
  const overlayListeners = new Set<() => void>()

  const emit = (set: Set<() => void>) => {
    for (const listener of [...set]) {
      try {
        listener()
      } catch (error) {
        console.error(error)
      }
    }
  }

  const measured = () => viewport.width > 0 && viewport.height > 0

  const apply = (next: ViewTransform) => {
    let result = constrainView(next, getDocSize(), viewport)
    // Whole-number display scales (100%, 200%...) stay aligned with device pixels: crisp pixels.
    const scale = displayScale(result.zoom, viewport.dpr)
    if (scale >= 1 && Math.abs(scale - Math.round(scale)) < 1e-6) result = snapView(result, viewport.dpr)
    if (viewsEqual(result, view)) return
    view = result
    emit(listeners)
    emit(overlayListeners)
  }

  const refit = () => {
    if (!measured()) return
    let next = fitView(getDocSize(), viewport, FIT_PADDING)
    if (!enlarge && displayScale(next.zoom, viewport.dpr) > 1) next = snapView(centerView(getDocSize(), viewport, zoomForScale(1, viewport.dpr)), viewport.dpr)
    if (viewsEqual(next, view)) return
    view = next
    emit(listeners)
    emit(overlayListeners)
  }

  const model: ViewModel = {
    get fitted() {
      return fitted
    },
    get measured() {
      return measured()
    },
    getView: () => view,
    getViewportSize: () => viewport,
    docToScreen: (point: Point) => docToScreen(view, point),
    screenToDoc: (point: Point) => screenToDoc(view, point),
    zoomAt(zoom: number, anchor: Point): void {
      if (!measured()) return
      fitted = false
      apply(zoomViewAt(view, clampZoom(zoom, viewport.dpr), anchor))
    },
    panBy(dx: number, dy: number): void {
      if (!measured() || (!dx && !dy)) return
      fitted = false
      apply(panView(view, dx, dy))
    },
    fit(): void {
      fitted = true
      enlarge = true
      refit()
    },
    actualPixels(): void {
      if (!measured()) return
      fitted = false
      apply(actualPixelsView(view, viewport))
    },
    requestOverlay(): void {
      emit(overlayListeners)
    },
    subscribe(listener: () => void): () => void {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
    onOverlay(listener: () => void): () => void {
      overlayListeners.add(listener)
      return () => { overlayListeners.delete(listener) }
    },
    setViewport(size: ViewportSize): void {
      const next = {
        width: Math.max(0, Math.round(size.width)),
        height: Math.max(0, Math.round(size.height)),
        dpr: Number.isFinite(size.dpr) && size.dpr > 0 ? size.dpr : 1,
      }
      if (next.width === viewport.width && next.height === viewport.height && next.dpr === viewport.dpr) return
      viewport = next
      if (fitted) refit()
      else apply(view)
      emit(listeners)
      emit(overlayListeners)
    },
    documentResized(): void {
      if (fitted) refit()
      else apply(view)
    },
  }
  return model
}

// ---------------------------------------------------------------------------------------------
// Surface (cursor, DOM overlay, focus) shared with the tool context
// ---------------------------------------------------------------------------------------------

export interface CanvasSurface {
  setCursor(cursor: string): void
  setOverlayElement(element: HTMLElement | null): void
  focus(): void
  /** True while a pointer button is held on the canvas (spring-loaded keys wait for it). */
  isPointerDown(): boolean
  /** The element keyboard focus returns to. */
  element(): HTMLElement | null
}

interface SurfaceInternals extends CanvasSurface {
  cursor(): string
  attach(stage: HTMLDivElement, overlayHost: HTMLDivElement): () => void
  setPointerDown(down: boolean): void
  setBusy(busy: boolean): void
}

export function createCanvasSurface(): CanvasSurface {
  let stage: HTMLDivElement | null = null
  let host: HTMLDivElement | null = null
  let cursor = 'default'
  let busy = false
  let overlay: HTMLElement | null = null
  let pointerDown = false

  const applyCursor = () => {
    if (stage) stage.style.cursor = busy ? 'progress' : cursor || 'default'
  }

  const mountOverlay = () => {
    if (!host) return
    for (const child of [...host.children]) if (child !== overlay) child.remove()
    if (overlay && overlay.parentElement !== host) host.appendChild(overlay)
  }

  const surface: SurfaceInternals = {
    setCursor(next: string): void {
      cursor = next
      applyCursor()
    },
    setOverlayElement(element: HTMLElement | null): void {
      if (overlay && overlay !== element) overlay.remove()
      overlay = element
      mountOverlay()
    },
    focus(): void {
      // Focusing the canvas must not scroll anything or steal focus from the text editor overlay.
      if (stage && !(overlay && overlay.contains(document.activeElement))) stage.focus({ preventScroll: true })
    },
    isPointerDown: () => pointerDown,
    element: () => stage,
    cursor: () => cursor,
    attach(nextStage: HTMLDivElement, overlayHost: HTMLDivElement): () => void {
      stage = nextStage
      host = overlayHost
      applyCursor()
      mountOverlay()
      return () => {
        if (stage === nextStage) {
          overlay?.remove()
          stage = null
          host = null
        }
      }
    },
    setPointerDown(down: boolean): void {
      pointerDown = down
    },
    setBusy(next: boolean): void {
      if (busy === next) return
      busy = next
      applyCursor()
    },
  }
  return surface
}

function internals(surface: CanvasSurface): SurfaceInternals {
  return surface as SurfaceInternals
}

// ---------------------------------------------------------------------------------------------
// Marching ants
// ---------------------------------------------------------------------------------------------

/** Outlines with more segments than this are drawn as their bounds (design 5.8). */
const MAX_OUTLINE_SEGMENTS = 200_000
/** Pixels of the selection mask traced at most (a coarser trace is used beyond it). */
const MAX_TRACE_PIXELS = 4_000_000

interface Outline {
  readonly key: string
  /** Segments x0, y0, x1, y1 in document px relative to (originX, originY), or null for the bounds box. */
  readonly segments: Float32Array | null
  readonly originX: number
  readonly originY: number
  readonly bounds: IntRect
}

function traceSelection(selection: Selection, level: number): Outline {
  const bounds = selection.bounds
  let factor = 2 ** Math.max(0, level)
  while ((bounds.width / factor) * (bounds.height / factor) > MAX_TRACE_PIXELS) factor *= 2
  const key = `${selection.version}:${factor}`
  if (isSelectAll(selection)) return { key, segments: null, originX: 0, originY: 0, bounds }
  // One pixel of margin so edges on the bounds trace as closed outlines.
  const rect: IntRect = { x: bounds.x - 1, y: bounds.y - 1, width: bounds.width + 2, height: bounds.height + 2 }
  let mask: MaskBuffer = cropMask(selection.mask, rect)
  if (factor > 1) mask = downsampleMask(mask, factor)
  const segments = traceOutline(mask, 128, factor)
  if (segments.length / 4 > MAX_OUTLINE_SEGMENTS) return { key, segments: null, originX: 0, originY: 0, bounds }
  return { key, segments, originX: rect.x, originY: rect.y, bounds }
}

function drawAnts(context: CanvasRenderingContext2D, outline: Outline, view: ViewTransform, dpr: number, phase: number): void {
  // Device pixels, with the document origin snapped like the compositor does (edges sit on pixel edges).
  const ox = Math.round(view.offsetX * dpr)
  const oy = Math.round(view.offsetY * dpr)
  const scale = view.zoom * dpr
  const dash = Math.max(3, Math.round(4 * dpr))
  context.save()
  try {
    context.setTransform(1, 0, 0, 1, 0, 0)
    context.beginPath()
    if (outline.segments) {
      const s = outline.segments
      const px = (x: number) => Math.round(ox + (outline.originX + x) * scale) + 0.5
      const py = (y: number) => Math.round(oy + (outline.originY + y) * scale) + 0.5
      for (let i = 0; i < s.length; i += 4) {
        context.moveTo(px(s[i]), py(s[i + 1]))
        context.lineTo(px(s[i + 2]), py(s[i + 3]))
      }
    } else {
      const b = outline.bounds
      const x0 = Math.round(ox + b.x * scale) + 0.5
      const y0 = Math.round(oy + b.y * scale) + 0.5
      const x1 = Math.round(ox + (b.x + b.width) * scale) - 0.5
      const y1 = Math.round(oy + (b.y + b.height) * scale) - 0.5
      context.rect(x0, y0, Math.max(0, x1 - x0), Math.max(0, y1 - y0))
    }
    context.lineWidth = 1
    context.setLineDash([])
    context.strokeStyle = '#ffffff'
    context.stroke()
    context.setLineDash([dash, dash])
    context.lineDashOffset = -phase * dash / 2
    context.strokeStyle = '#000000'
    context.stroke()
  } finally {
    context.restore()
  }
}

function drawDocumentEdge(context: CanvasRenderingContext2D, doc: Size, view: ViewTransform, dpr: number): void {
  const ox = Math.round(view.offsetX * dpr)
  const oy = Math.round(view.offsetY * dpr)
  const scale = view.zoom * dpr
  context.save()
  try {
    context.setTransform(1, 0, 0, 1, 0, 0)
    context.lineWidth = 1
    context.strokeStyle = 'rgba(0, 0, 0, 0.22)'
    context.strokeRect(ox - 0.5, oy - 0.5, Math.round(doc.width * scale) + 1, Math.round(doc.height * scale) + 1)
  } finally {
    context.restore()
  }
}

// ---------------------------------------------------------------------------------------------
// The component
// ---------------------------------------------------------------------------------------------

export interface CanvasInput {
  /** The controller receiving pointer input and drawing the overlay (free-transform session first). */
  target(): ToolController | null
  /** A document job is running: pointer input is ignored (the cursor shows progress). */
  busy(): boolean
  /** The Alt key is springing to the eyedropper: Alt is not passed on (Alt-click would pick the background colour). */
  stripAlt(): boolean
  /** Tools whose overlay is only a hover cursor (hidden when the pointer leaves the canvas). */
  hoverOnly(controller: ToolController): boolean
  /** Pointer position for the status bar (document px), null when the pointer leaves. */
  hover?(point: Point | null): void
}

export interface CanvasViewProps {
  readonly store: DocumentStore
  readonly compositor: AdvancedCompositor
  readonly view: ViewModel
  readonly surface: CanvasSurface
  readonly input: CanvasInput
  readonly suspended: boolean
  /** Bumped by the editor when a document job starts or ends (cursor). */
  readonly busy: boolean
}

const WHEEL_LINE_PX = 16

interface PointerState {
  readonly id: number
  readonly type: ToolPointerEvent['pointerType']
  /** Who receives this pointer's events. */
  readonly mode: 'tool' | 'pan'
  /** Cursor to restore after a middle-button pan. */
  readonly cursor: string
  last: Point
  lastTime: number
}

interface PinchState {
  distance: number
  center: Point
}

export function CanvasView({ store, compositor, view, surface, input, suspended, busy }: CanvasViewProps) {
  const stageRef = useRef<HTMLDivElement>(null)
  const displayRef = useRef<HTMLCanvasElement>(null)
  const overlayRef = useRef<HTMLCanvasElement>(null)
  const overlayHostRef = useRef<HTMLDivElement>(null)
  const inputRef = useRef(input)
  inputRef.current = input
  const suspendedRef = useRef(suspended)
  suspendedRef.current = suspended
  const pointerRef = useRef<PointerState | null>(null)
  const touchesRef = useRef(new Map<number, Point>())
  const pinchRef = useRef<PinchState | null>(null)
  const insideRef = useRef(false)
  const rectRef = useRef<DOMRect | null>(null)
  const outlineRef = useRef<Outline | null>(null)
  const phaseRef = useRef(0)
  const drawRef = useRef<() => void>(() => {})

  useEffect(() => {
    internals(surface).setBusy(busy)
  }, [busy, surface])

  // DOM attachment: surface, compositor, viewport size.
  useLayoutEffect(() => {
    const stage = stageRef.current
    const display = displayRef.current
    const overlayHost = overlayHostRef.current
    if (!stage || !display || !overlayHost) return
    const detach = internals(surface).attach(stage, overlayHost)
    compositor.attach(display)
    const measure = () => {
      rectRef.current = stage.getBoundingClientRect()
      view.setViewport({ width: stage.clientWidth, height: stage.clientHeight, dpr: window.devicePixelRatio || 1 })
    }
    measure()
    const observer = new ResizeObserver(measure)
    observer.observe(stage)
    window.addEventListener('resize', measure)
    return () => {
      observer.disconnect()
      window.removeEventListener('resize', measure)
      detach()
    }
  }, [compositor, surface, view])

  // View changes reach the compositor; the editor's document changes reach the view.
  useEffect(() => {
    const push = () => {
      const current = view.getView()
      const size = view.getViewportSize()
      compositor.setView(current, size)
      // QA hook: "zoom offsetX offsetY dpr" maps document points to the stage (screen = doc * zoom + offset).
      const stage = stageRef.current
      if (stage) stage.dataset.view = `${current.zoom} ${current.offsetX} ${current.offsetY} ${size.dpr}`
    }
    push()
    const unsubscribeView = view.subscribe(push)
    let size = { width: store.getState().width, height: store.getState().height }
    const unsubscribeStore = store.subscribe((change) => {
      const state = store.getState()
      if (state.width !== size.width || state.height !== size.height) {
        size = { width: state.width, height: state.height }
        view.documentResized()
      }
      if (change.selection) drawRef.current()
    })
    return () => {
      unsubscribeView()
      unsubscribeStore()
    }
  }, [compositor, store, view])

  // Overlay drawing (coalesced to one draw per frame; a timer stands in where rAF does not fire).
  useEffect(() => {
    let frame: number | null = null
    let timer: ReturnType<typeof setTimeout> | null = null
    const draw = () => {
      if (frame !== null) cancelAnimationFrame(frame)
      if (timer !== null) clearTimeout(timer)
      frame = null
      timer = null
      const canvas = overlayRef.current
      if (!canvas) return
      const size = view.getViewportSize()
      const width = Math.max(1, Math.round(size.width * size.dpr))
      const height = Math.max(1, Math.round(size.height * size.dpr))
      if (canvas.width !== width || canvas.height !== height) {
        canvas.width = width
        canvas.height = height
      }
      const context = canvas.getContext('2d')
      if (!context) return
      context.setTransform(1, 0, 0, 1, 0, 0)
      context.clearRect(0, 0, width, height)
      if (!view.measured) return
      const current = view.getView()
      const state = store.getState()
      drawDocumentEdge(context, state, current, size.dpr)
      const selection = state.selection
      if (selection) {
        let outline = outlineRef.current
        const level = compositor.level
        let factor = 2 ** Math.max(0, level)
        while ((selection.bounds.width / factor) * (selection.bounds.height / factor) > MAX_TRACE_PIXELS) factor *= 2
        if (!outline || outline.key !== `${selection.version}:${factor}`) {
          try {
            outline = traceSelection(selection, level)
          } catch (error) {
            console.error(error)
            outline = { key: `${selection.version}:${factor}`, segments: null, originX: 0, originY: 0, bounds: selection.bounds }
          }
          outlineRef.current = outline
        }
        drawAnts(context, outline, current, size.dpr, phaseRef.current)
      } else {
        outlineRef.current = null
      }
      const target = inputRef.current.target()
      if (target && (insideRef.current || pointerRef.current || !inputRef.current.hoverOnly(target))) {
        context.save()
        try {
          context.setTransform(size.dpr, 0, 0, size.dpr, 0, 0)
          target.drawOverlay(context, current)
        } catch (error) {
          console.error(error)
        } finally {
          context.restore()
        }
      }
    }
    const schedule = () => {
      if (frame !== null || timer !== null) return
      frame = requestAnimationFrame(draw)
      timer = setTimeout(draw, 100)
    }
    drawRef.current = schedule
    const offOverlay = view.onOverlay(schedule)
    schedule()
    // Marching ants: 10 fps while a selection exists and the page is visible.
    const ants = setInterval(() => {
      if (!store.getState().selection || document.visibilityState === 'hidden') return
      phaseRef.current = (phaseRef.current + 1) % 1024
      schedule()
    }, 100)
    return () => {
      offOverlay()
      clearInterval(ants)
      if (frame !== null) cancelAnimationFrame(frame)
      if (timer !== null) clearTimeout(timer)
      drawRef.current = () => {}
    }
  }, [compositor, store, view])

  // Wheel (native, non-passive: React's onWheel is passive and cannot prevent the page zoom).
  useEffect(() => {
    const stage = stageRef.current
    if (!stage) return
    const onWheel = (event: WheelEvent) => {
      event.preventDefault()
      if (suspendedRef.current) return
      const unit = event.deltaMode === 1 ? WHEEL_LINE_PX : event.deltaMode === 2 ? stage.clientHeight : 1
      const rect = stage.getBoundingClientRect()
      const anchor = { x: event.clientX - rect.left, y: event.clientY - rect.top }
      if (event.ctrlKey || event.metaKey || event.altKey) {
        const factor = Math.min(2, Math.max(0.5, Math.exp(-event.deltaY * unit * 0.002)))
        view.zoomAt(view.getView().zoom * factor, anchor)
        return
      }
      let dx = -event.deltaX * unit
      let dy = -event.deltaY * unit
      if (event.shiftKey && !dx) {
        dx = dy
        dy = 0
      }
      view.panBy(dx, dy)
    }
    stage.addEventListener('wheel', onWheel, { passive: false })
    return () => stage.removeEventListener('wheel', onWheel)
  }, [view])

  // Never leave a gesture half done when the editor closes or input is suspended.
  useEffect(() => {
    if (!suspended) return
    const pointer = pointerRef.current
    if (pointer) {
      pointerRef.current = null
      internals(surface).setPointerDown(false)
      if (pointer.mode === 'tool') inputRef.current.target()?.pointerCancel()
    }
  }, [suspended, surface])

  useEffect(() => () => {
    if (pointerRef.current?.mode === 'tool') inputRef.current.target()?.pointerCancel()
    pointerRef.current = null
  }, [])

  // ----- pointer routing -----------------------------------------------------------------------

  const localPoint = (clientX: number, clientY: number): Point => {
    const rect = rectRef.current ?? stageRef.current?.getBoundingClientRect() ?? null
    return rect ? { x: clientX - rect.left, y: clientY - rect.top } : { x: clientX, y: clientY }
  }

  const toolEvent = (event: ReactPointerEvent<HTMLDivElement>, down: boolean): ToolPointerEvent => {
    const native = event.nativeEvent
    const screen = localPoint(event.clientX, event.clientY)
    const type: ToolPointerEvent['pointerType'] = event.pointerType === 'pen' ? 'pen' : event.pointerType === 'touch' ? 'touch' : 'mouse'
    const pressureOf = (value: number, buttons: number) => (type === 'pen'
      ? (Number.isFinite(value) ? Math.min(1, Math.max(0, value)) : 1)
      : buttons !== 0 ? 1 : 0)
    const coalesced: { doc: Point; pressure: number; time: number }[] = []
    if (!down && typeof native.getCoalescedEvents === 'function') {
      try {
        for (const sample of native.getCoalescedEvents()) {
          coalesced.push({
            doc: view.screenToDoc(localPoint(sample.clientX, sample.clientY)),
            pressure: pressureOf(sample.pressure, sample.buttons),
            time: sample.timeStamp,
          })
        }
      } catch {
        // Some synthetic events cannot report coalesced samples.
      }
    }
    return {
      doc: view.screenToDoc(screen),
      screen,
      pressure: pressureOf(event.pressure, event.buttons),
      button: event.button,
      buttons: event.buttons,
      shift: event.shiftKey,
      alt: event.altKey && !inputRef.current.stripAlt(),
      ctrl: event.ctrlKey || event.metaKey,
      pointerType: type,
      time: event.timeStamp,
      coalesced,
    }
  }

  const pinchFrom = (touches: Map<number, Point>): PinchState | null => {
    const points = [...touches.values()]
    if (points.length < 2) return null
    const [a, b] = points
    return { distance: Math.hypot(b.x - a.x, b.y - a.y), center: { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 } }
  }

  const releaseCapture = (event: ReactPointerEvent<HTMLDivElement>) => {
    try {
      if (event.currentTarget.hasPointerCapture(event.pointerId)) event.currentTarget.releasePointerCapture(event.pointerId)
    } catch {
      // Synthetic pointers have no capture.
    }
  }

  const onPointerDown = (event: ReactPointerEvent<HTMLDivElement>) => {
    if (suspendedRef.current) return
    const stage = event.currentTarget
    if (event.target !== stage && !(event.target instanceof HTMLCanvasElement)) return
    rectRef.current = stage.getBoundingClientRect()
    surface.focus()
    if (event.pointerType === 'touch') {
      touchesRef.current.set(event.pointerId, localPoint(event.clientX, event.clientY))
      if (touchesRef.current.size >= 2) {
        // A second finger: cancel the tool gesture and pinch instead.
        const pointer = pointerRef.current
        if (pointer && pointer.mode === 'tool') inputRef.current.target()?.pointerCancel()
        pointerRef.current = null
        internals(surface).setPointerDown(false)
        pinchRef.current = pinchFrom(touchesRef.current)
        try {
          stage.setPointerCapture(event.pointerId)
        } catch {
          // ignore
        }
        event.preventDefault()
        return
      }
    }
    if (pointerRef.current) return
    if (inputRef.current.busy()) return
    const mode: PointerState['mode'] = event.button === 1 ? 'pan' : 'tool'
    // Tools take the primary button (mouse, pen tip, touch); the middle button pans; others do nothing.
    if (mode === 'tool' && event.button !== 0) return
    pointerRef.current = {
      id: event.pointerId,
      type: event.pointerType as PointerState['type'],
      mode,
      cursor: internals(surface).cursor(),
      last: localPoint(event.clientX, event.clientY),
      lastTime: event.timeStamp,
    }
    internals(surface).setPointerDown(true)
    try {
      stage.setPointerCapture(event.pointerId)
    } catch {
      // Synthetic events (tests) have no active pointer to capture.
    }
    event.preventDefault()
    if (mode === 'pan') {
      surface.setCursor('grabbing')
      return
    }
    try {
      inputRef.current.target()?.pointerDown(toolEvent(event, true))
    } catch (error) {
      console.error(error)
    }
  }

  const onPointerMove = (event: ReactPointerEvent<HTMLDivElement>) => {
    const point = localPoint(event.clientX, event.clientY)
    if (event.pointerType === 'touch' && touchesRef.current.has(event.pointerId)) {
      touchesRef.current.set(event.pointerId, point)
      const pinch = pinchRef.current
      if (pinch) {
        const next = pinchFrom(touchesRef.current)
        if (next && pinch.distance > 0) {
          view.zoomAt(view.getView().zoom * (next.distance / pinch.distance), pinch.center)
          view.panBy(next.center.x - pinch.center.x, next.center.y - pinch.center.y)
          pinchRef.current = next
        }
        return
      }
    }
    const pointer = pointerRef.current
    if (pointer && pointer.id !== event.pointerId) return
    if (pointer && pointer.mode === 'pan') {
      view.panBy(point.x - pointer.last.x, point.y - pointer.last.y)
      pointer.last = point
      return
    }
    if (pointer) pointer.last = point
    insideRef.current = true
    inputRef.current.hover?.(view.screenToDoc(point))
    if (suspendedRef.current) return
    if (!pointer && inputRef.current.busy()) return
    try {
      inputRef.current.target()?.pointerMove(toolEvent(event, false))
    } catch (error) {
      console.error(error)
    }
  }

  const finishPointer = (event: ReactPointerEvent<HTMLDivElement>, cancelled: boolean) => {
    if (event.pointerType === 'touch') {
      touchesRef.current.delete(event.pointerId)
      if (pinchRef.current) {
        if (touchesRef.current.size < 2) pinchRef.current = null
        releaseCapture(event)
        return
      }
    }
    const pointer = pointerRef.current
    if (!pointer || pointer.id !== event.pointerId) return
    pointerRef.current = null
    internals(surface).setPointerDown(false)
    releaseCapture(event)
    if (pointer.mode === 'pan') {
      surface.setCursor(pointer.cursor)
      return
    }
    try {
      const target = inputRef.current.target()
      if (cancelled) target?.pointerCancel()
      else target?.pointerUp(toolEvent(event, false))
    } catch (error) {
      console.error(error)
    }
    drawRef.current()
  }

  const onPointerLeave = () => {
    insideRef.current = false
    inputRef.current.hover?.(null)
    drawRef.current()
  }

  return (
    <div
      ref={stageRef}
      className="ae-stage"
      tabIndex={0}
      role="application"
      aria-roledescription="image canvas"
      aria-label="Image canvas. Use the tool shortcuts to edit; hold Space to move around."
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={(event) => finishPointer(event, false)}
      onPointerCancel={(event) => finishPointer(event, true)}
      onLostPointerCapture={(event) => {
        if (pointerRef.current?.id === event.pointerId) finishPointer(event, false)
      }}
      onPointerLeave={onPointerLeave}
      onContextMenu={(event) => event.preventDefault()}
      onDragStart={(event) => event.preventDefault()}
    >
      <canvas ref={displayRef} className="ae-display" aria-hidden="true" />
      <canvas ref={overlayRef} className="ae-overlay" aria-hidden="true" />
      <div ref={overlayHostRef} className="ae-overlay-host" />
    </div>
  )
}

/** Releases the GPU memory of the canvases inside a stage element (called when the editor closes). */
export function releaseStageCanvases(stage: HTMLElement | null): void {
  if (!stage) return
  for (const canvas of stage.querySelectorAll('canvas')) {
    canvas.width = 1
    canvas.height = 1
  }
}
