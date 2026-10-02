// src/simple/useViewportNavigation.ts (WP8)
// Photoshop-grade navigation for the Simple viewport (design 4.1):
//   - a NATIVE, non-passive wheel listener (React's onWheel is passive, so its preventDefault was a no-op):
//     Ctrl+wheel and trackpad pinch zoom by exp(-deltaY_px * 0.002), clamped to [0.5, 2] per event;
//     plain wheel keeps native scrolling (Shift+wheel scrolls sideways);
//   - every zoom keeps the document point under the anchor (the cursor, or the viewport centre for keys
//     and buttons) under it: the point is measured before the zoom state changes and the scroll position is
//     corrected in a layout effect after the commit, independent of the stage padding and centring;
//   - Ctrl+= / Ctrl+- step through Photoshop's zoom ladder in DEVICE percent; 1:1 (Ctrl+1) is one image pixel
//     per device pixel (zoom = 1 / devicePixelRatio), so the label shows device percent;
//   - Space+drag, middle-button drag, View-tool drag and two-finger touch pan (touch also pinches);
//     double-click with the View tool toggles Fit and 1:1 at the click point;
//   - crisp pixels (image-rendering: pixelated) from 200% device zoom; at most 3200% CSS and a canvas box of
//     at most 200,000 px per side.
// usePointerReadout() feeds the status bar: x, y and the pixel's #RRGGBB (alpha when < 255) at 30 Hz.
import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react'
import type { RefObject } from 'react'
import type { Rgba8 } from '../imaging/types.ts'
import { pixelAt } from './ops.ts'

export interface ClientPoint {
  readonly x: number
  readonly y: number
}

/** Photoshop's zoom presets, device percent. */
export const ZOOM_LADDER: readonly number[] = Object.freeze([6.25, 8.33, 12.5, 16.67, 25, 33.33, 50, 66.67, 100, 150, 200, 300, 400, 500, 600, 700, 800, 1200, 1600, 2400, 3200])
export const MIN_ZOOM = 0.01
export const MAX_ZOOM = 32
/** Largest CSS size of the canvas box per side. */
export const MAX_CSS_EXTENT = 200_000
/** Device zoom from which pixels are drawn as crisp squares. */
export const PIXELATED_FROM = 2

export function devicePixelRatioNow(): number {
  return typeof window !== 'undefined' && window.devicePixelRatio > 0 ? window.devicePixelRatio : 1
}

/** Largest CSS zoom for content of `size`. */
export function maxZoomFor(size: { width: number; height: number }): number {
  const longest = Math.max(1, size.width, size.height)
  return Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, MAX_CSS_EXTENT / longest))
}

export function clampZoom(zoom: number, size: { width: number; height: number }): number {
  if (!Number.isFinite(zoom)) return 1
  return Math.max(MIN_ZOOM, Math.min(maxZoomFor(size), zoom))
}

/** Next CSS zoom on the device-percent ladder in `direction` (+1 in, -1 out). */
export function ladderZoom(zoom: number, dpr: number, direction: 1 | -1): number {
  const percent = zoom * dpr * 100
  if (direction > 0) {
    const next = ZOOM_LADDER.find((step) => step > percent * 1.001)
    return (next ?? ZOOM_LADDER[ZOOM_LADDER.length - 1]) / 100 / dpr
  }
  for (let index = ZOOM_LADDER.length - 1; index >= 0; index -= 1) {
    if (ZOOM_LADDER[index] < percent * 0.999) return ZOOM_LADDER[index] / 100 / dpr
  }
  return Math.min(zoom, ZOOM_LADDER[0] / 100 / dpr)
}

/** Zoom factor of one wheel event (pixel, line or page deltas). */
export function wheelZoomFactor(deltaY: number, deltaMode: number, pageHeight: number): number {
  const pixels = deltaY * (deltaMode === 1 ? 16 : deltaMode === 2 ? Math.max(1, pageHeight) : 1)
  return Math.max(0.5, Math.min(2, Math.exp(-pixels * 0.002)))
}

/**
 * Pointer events that started a pan (Space, middle button, View-tool drag, second touch). The viewport
 * sees them first (capture phase); the canvas, crop, markup and level layers ignore them. They are not
 * stopped, so document-level listeners (closing menus) still run.
 */
const panEvents = new WeakSet<Event>()
export function isPanEvent(event: Event): boolean {
  return panEvents.has(event)
}

/** Movement (CSS px) before a View-tool drag becomes a pan, so clicks and double-clicks stay clicks. */
const PAN_THRESHOLD = 3

function isTextTarget(target: EventTarget | null): boolean {
  return target instanceof Element && Boolean(target.closest('input, textarea, select, [contenteditable="true"]'))
}

export interface ViewportNavigationOptions {
  readonly viewportRef: RefObject<HTMLDivElement | null>
  /** The element whose box is the document (the canvas stack). */
  readonly stackRef: RefObject<HTMLDivElement | null>
  /** False while no image is open or Advanced owns the screen. */
  readonly enabled: boolean
  /** Current CSS zoom (state). */
  readonly zoom: number
  /** Size of the document box content at zoom 1 (image, or the crop frame while straightening). */
  readonly contentSize: { width: number; height: number }
  /** Commits a new CSS zoom (and leaves Fit mode). */
  readonly applyZoom: (zoom: number) => void
  readonly fit: () => void
  readonly fitMode: boolean
  /** Primary-button drags pan (the View tool). */
  readonly panWithPrimary: boolean
}

export interface ViewportNavigation {
  zoomTo(zoom: number, anchor?: ClientPoint | null): void
  step(direction: 1 | -1, anchor?: ClientPoint | null): void
  actualPixels(anchor?: ClientPoint | null): void
  /** Space pressed (main.tsx forwards keydown): returns true when it was consumed for panning. */
  handleSpaceDown(event: KeyboardEvent): boolean
  readonly dpr: number
  readonly devicePercent: number
  readonly pixelated: boolean
  readonly spaceHeld: boolean
  readonly panning: boolean
}

interface PanState {
  readonly pointerId: number
  readonly startX: number
  readonly startY: number
  readonly scrollLeft: number
  readonly scrollTop: number
  /** False until the pointer moved past PAN_THRESHOLD (View-tool drags only). */
  active: boolean
}

interface PinchState {
  readonly startDistance: number
  readonly startZoom: number
  lastMid: ClientPoint
}

export function useViewportNavigation(options: ViewportNavigationOptions): ViewportNavigation {
  // fit, fitMode and panWithPrimary are read through optionsRef by the native listeners.
  const { viewportRef, stackRef, enabled, zoom } = options
  const [dpr, setDpr] = useState(devicePixelRatioNow)
  const [spaceHeld, setSpaceHeld] = useState(false)
  const [panning, setPanning] = useState(false)
  const renderedZoomRef = useRef(zoom)
  const targetZoomRef = useRef(zoom)
  const pendingAnchorRef = useRef<{ clientX: number; clientY: number; docX: number; docY: number } | null>(null)
  const optionsRef = useRef(options)
  optionsRef.current = options
  const spaceRef = useRef(false)
  const pointerOverRef = useRef(false)
  const panRef = useRef<PanState | null>(null)
  const touchesRef = useRef(new Map<number, ClientPoint>())
  const pinchRef = useRef<PinchState | null>(null)

  // A zoom that did not come through zoomTo (Fit, a new image) resets the target.
  if (renderedZoomRef.current !== zoom) {
    renderedZoomRef.current = zoom
    if (!pendingAnchorRef.current) targetZoomRef.current = zoom
  }

  useEffect(() => {
    const update = () => setDpr(devicePixelRatioNow())
    window.addEventListener('resize', update)
    const query = window.matchMedia?.(`(resolution: ${devicePixelRatioNow()}dppx)`)
    query?.addEventListener?.('change', update)
    return () => {
      window.removeEventListener('resize', update)
      query?.removeEventListener?.('change', update)
    }
  }, [dpr])

  const zoomTo = useCallback((next: number, anchor?: ClientPoint | null) => {
    const { contentSize: size, applyZoom: commit } = optionsRef.current
    const clamped = clampZoom(next, size)
    targetZoomRef.current = clamped
    if (Math.abs(clamped - renderedZoomRef.current) < 1e-9) return
    const stack = stackRef.current
    const viewport = viewportRef.current
    if (stack && viewport) {
      const rect = stack.getBoundingClientRect()
      const box = viewport.getBoundingClientRect()
      const point = anchor ?? { x: box.left + box.width / 2, y: box.top + box.height / 2 }
      const current = renderedZoomRef.current || 1
      pendingAnchorRef.current = { clientX: point.x, clientY: point.y, docX: (point.x - rect.left) / current, docY: (point.y - rect.top) / current }
    }
    commit(clamped)
  }, [stackRef, viewportRef])

  useLayoutEffect(() => {
    const anchor = pendingAnchorRef.current
    if (!anchor) return
    pendingAnchorRef.current = null
    const stack = stackRef.current
    const viewport = viewportRef.current
    if (!stack || !viewport) return
    const rect = stack.getBoundingClientRect()
    viewport.scrollLeft += rect.left + anchor.docX * zoom - anchor.clientX
    viewport.scrollTop += rect.top + anchor.docY * zoom - anchor.clientY
  }, [stackRef, viewportRef, zoom])

  const step = useCallback((direction: 1 | -1, anchor?: ClientPoint | null) => {
    zoomTo(ladderZoom(targetZoomRef.current, devicePixelRatioNow(), direction), anchor)
  }, [zoomTo])

  const actualPixels = useCallback((anchor?: ClientPoint | null) => {
    zoomTo(1 / devicePixelRatioNow(), anchor)
  }, [zoomTo])

  const handleSpaceDown = useCallback((event: KeyboardEvent) => {
    if (!optionsRef.current.enabled) return false
    if (event.key !== ' ' && event.code !== 'Space') return false
    if (event.ctrlKey || event.metaKey || event.altKey || event.isComposing || isTextTarget(event.target)) return false
    // A focused button keeps Space for itself unless the pointer is over the image (the user wants to pan).
    const onControl = event.target instanceof Element && Boolean(event.target.closest('button, a, [role="menuitem"], [role="slider"]'))
    if (onControl && !pointerOverRef.current) return false
    event.preventDefault()
    if (!event.repeat && !spaceRef.current) {
      spaceRef.current = true
      setSpaceHeld(true)
    }
    return true
  }, [])

  useEffect(() => {
    const release = () => {
      if (!spaceRef.current) return
      spaceRef.current = false
      setSpaceHeld(false)
    }
    const onKeyUp = (event: KeyboardEvent) => {
      if (event.key === ' ' || event.code === 'Space') release()
    }
    window.addEventListener('keyup', onKeyUp)
    window.addEventListener('blur', release)
    return () => {
      window.removeEventListener('keyup', onKeyUp)
      window.removeEventListener('blur', release)
    }
  }, [])

  useEffect(() => {
    const viewport = viewportRef.current
    if (!viewport || !enabled) return
    const onWheel = (event: WheelEvent) => {
      if (!optionsRef.current.enabled || !(event.ctrlKey || event.metaKey)) return
      event.preventDefault()
      const factor = wheelZoomFactor(event.deltaY, event.deltaMode, viewport.clientHeight)
      zoomTo(targetZoomRef.current * factor, { x: event.clientX, y: event.clientY })
    }
    const endPan = (event: PointerEvent) => {
      const pan = panRef.current
      if (!pan || pan.pointerId !== event.pointerId) return
      panRef.current = null
      if (pan.active) {
        setPanning(false)
        try { viewport.releasePointerCapture(event.pointerId) } catch { /* already released */ }
      }
    }
    const activate = (pan: PanState) => {
      pan.active = true
      setPanning(true)
      try { viewport.setPointerCapture(pan.pointerId) } catch { /* the pointer is gone */ }
    }
    const onPointerDown = (event: PointerEvent) => {
      if (!optionsRef.current.enabled) return
      const target = event.target instanceof Element ? event.target : null
      if (!target || target.closest('button, input, textarea, select, .album-nav, .image-context-menu')) return
      if (event.pointerType === 'touch') {
        touchesRef.current.set(event.pointerId, { x: event.clientX, y: event.clientY })
        if (touchesRef.current.size === 2) {
          const [a, b] = [...touchesRef.current.values()]
          pinchRef.current = { startDistance: Math.max(1, Math.hypot(a.x - b.x, a.y - b.y)), startZoom: targetZoomRef.current, lastMid: { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 } }
          if (panRef.current?.active) setPanning(false)
          panRef.current = null
          panEvents.add(event)
          event.preventDefault()
          return
        }
      }
      const onStage = Boolean(target.closest('.canvas-stage'))
      const overlay = Boolean(target.closest('.crop-layer, .markup-layer, .markup-text-editor, .level-line-layer'))
      const explicit = event.button === 1 || (event.button === 0 && spaceRef.current)
      const viewDrag = event.button === 0 && optionsRef.current.panWithPrimary && onStage && !overlay
      if (!explicit && !viewDrag) return
      panEvents.add(event)
      // Middle button: no autoscroll; Space: no focus change. A View-tool press stays a normal click until it moves.
      if (explicit) event.preventDefault()
      const pan: PanState = { pointerId: event.pointerId, startX: event.clientX, startY: event.clientY, scrollLeft: viewport.scrollLeft, scrollTop: viewport.scrollTop, active: false }
      panRef.current = pan
      if (explicit) activate(pan)
    }
    const onPointerMove = (event: PointerEvent) => {
      if (touchesRef.current.has(event.pointerId)) {
        touchesRef.current.set(event.pointerId, { x: event.clientX, y: event.clientY })
        const pinch = pinchRef.current
        if (pinch && touchesRef.current.size >= 2) {
          const [a, b] = [...touchesRef.current.values()]
          const mid = { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 }
          viewport.scrollLeft -= mid.x - pinch.lastMid.x
          viewport.scrollTop -= mid.y - pinch.lastMid.y
          pinch.lastMid = mid
          zoomTo(pinch.startZoom * Math.hypot(a.x - b.x, a.y - b.y) / pinch.startDistance, mid)
          event.preventDefault()
          return
        }
      }
      const pan = panRef.current
      if (!pan || pan.pointerId !== event.pointerId) return
      if (!pan.active) {
        if (Math.hypot(event.clientX - pan.startX, event.clientY - pan.startY) < PAN_THRESHOLD) return
        activate(pan)
      }
      viewport.scrollLeft = pan.scrollLeft - (event.clientX - pan.startX)
      viewport.scrollTop = pan.scrollTop - (event.clientY - pan.startY)
    }
    const onPointerEnd = (event: PointerEvent) => {
      if (touchesRef.current.delete(event.pointerId) && touchesRef.current.size < 2) pinchRef.current = null
      endPan(event)
    }
    const onDoubleClick = (event: MouseEvent) => {
      if (!optionsRef.current.enabled || !optionsRef.current.panWithPrimary) return
      const target = event.target instanceof Element ? event.target : null
      if (!target?.closest('.canvas-stage') || target.closest('button')) return
      event.preventDefault()
      const anchor = { x: event.clientX, y: event.clientY }
      const actual = 1 / devicePixelRatioNow()
      if (optionsRef.current.fitMode || Math.abs(targetZoomRef.current - actual) > 1e-6) zoomTo(actual, anchor)
      else optionsRef.current.fit()
    }
    const onAuxClick = (event: MouseEvent) => { if (event.button === 1) event.preventDefault() }
    const onEnter = () => { pointerOverRef.current = true }
    const onLeave = () => { pointerOverRef.current = false }
    viewport.addEventListener('wheel', onWheel, { passive: false })
    viewport.addEventListener('pointerdown', onPointerDown, { capture: true })
    viewport.addEventListener('pointermove', onPointerMove)
    viewport.addEventListener('pointerup', onPointerEnd)
    viewport.addEventListener('pointercancel', onPointerEnd)
    viewport.addEventListener('dblclick', onDoubleClick)
    viewport.addEventListener('auxclick', onAuxClick)
    viewport.addEventListener('pointerenter', onEnter)
    viewport.addEventListener('pointerleave', onLeave)
    return () => {
      viewport.removeEventListener('wheel', onWheel)
      viewport.removeEventListener('pointerdown', onPointerDown, { capture: true })
      viewport.removeEventListener('pointermove', onPointerMove)
      viewport.removeEventListener('pointerup', onPointerEnd)
      viewport.removeEventListener('pointercancel', onPointerEnd)
      viewport.removeEventListener('dblclick', onDoubleClick)
      viewport.removeEventListener('auxclick', onAuxClick)
      viewport.removeEventListener('pointerenter', onEnter)
      viewport.removeEventListener('pointerleave', onLeave)
      panRef.current = null
      pinchRef.current = null
      touchesRef.current.clear()
    }
  }, [enabled, viewportRef, zoomTo])

  useEffect(() => {
    if (enabled) return
    spaceRef.current = false
    setSpaceHeld(false)
    setPanning(false)
  }, [enabled])

  const devicePercent = zoom * dpr * 100
  return {
    zoomTo,
    step,
    actualPixels,
    handleSpaceDown,
    dpr,
    devicePercent,
    pixelated: zoom * dpr >= PIXELATED_FROM - 1e-9,
    spaceHeld,
    panning,
  }
}

export interface PointerReadout {
  readonly x: number
  readonly y: number
  readonly color: Rgba8 | null
}

/** Status-bar readout of the image pixel under the pointer, throttled to 30 Hz (one 1 x 1 readback). */
export function usePointerReadout(canvasRef: RefObject<HTMLCanvasElement | null>, enabled: boolean) {
  const [readout, setReadout] = useState<PointerReadout | null>(null)
  const pendingRef = useRef<ClientPoint | null>(null)
  const frameRef = useRef(0)
  const lastRef = useRef(0)
  const enabledRef = useRef(enabled)
  enabledRef.current = enabled

  const flush = useCallback(() => {
    frameRef.current = 0
    const point = pendingRef.current
    const canvas = canvasRef.current
    if (!point || !canvas || !enabledRef.current) return
    const now = performance.now()
    if (now - lastRef.current < 33) {
      frameRef.current = window.requestAnimationFrame(flush)
      return
    }
    lastRef.current = now
    const rect = canvas.getBoundingClientRect()
    if (!rect.width || !rect.height) return
    const x = Math.floor((point.x - rect.left) * canvas.width / rect.width)
    const y = Math.floor((point.y - rect.top) * canvas.height / rect.height)
    if (x < 0 || y < 0 || x >= canvas.width || y >= canvas.height) {
      setReadout(null)
      return
    }
    let color: Rgba8 | null = null
    try { color = pixelAt(canvas, x, y) } catch { color = null }
    setReadout({ x, y, color })
  }, [canvasRef])

  const track = useCallback((clientX: number, clientY: number) => {
    if (!enabledRef.current) return
    pendingRef.current = { x: clientX, y: clientY }
    if (!frameRef.current) frameRef.current = window.requestAnimationFrame(flush)
  }, [flush])

  const clear = useCallback(() => {
    pendingRef.current = null
    if (frameRef.current) window.cancelAnimationFrame(frameRef.current)
    frameRef.current = 0
    setReadout(null)
  }, [])

  useEffect(() => {
    if (!enabled) clear()
  }, [clear, enabled])

  useEffect(() => () => { if (frameRef.current) window.cancelAnimationFrame(frameRef.current) }, [])

  return { readout, track, clear }
}
