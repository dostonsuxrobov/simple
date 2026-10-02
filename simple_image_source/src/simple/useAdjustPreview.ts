// src/simple/useAdjustPreview.ts (WP8)
// Live preview for the Adjust panel and Looks (design 4.4 / 4.5). When the panel opens, the image is drawn
// once into a proxy at the displayed device size (at most 2.5 MP). Every slider change applies the compiled
// quick-adjust kernel to a copy of the proxy on the main thread (a few ms: point operations commute with
// downscaling, so the preview matches the final result) and paints it into an overlay canvas that exactly
// covers the image; a Look is applied to that result in the imaging worker (latest request wins). The
// full-resolution commit happens in main.tsx through the worker (one undo step).
import { useEffect, useRef, useState } from 'react'
import type { RefObject } from 'react'
import type { AdjustmentSpec, LookId, PixelBuffer, QuickAdjust } from '../imaging/types.ts'
import { compileAdjustment } from '../imaging/adjustments.ts'
import { LOOK_IDS, applyLook } from '../imaging/looks.ts'
import { getImagingClient } from '../shared/workerClient.ts'
import { putBuffer, releaseCanvas } from '../shared/canvas.ts'
import { canvasProxy, pixelsToDataUrl, proxySize } from './ops.ts'

export interface PendingAdjust {
  readonly quick: QuickAdjust
  readonly look: LookId
  /** 0..100 */
  readonly intensity: number
}

export const NEUTRAL_QUICK: QuickAdjust = Object.freeze({ exposure: 0, brightness: 0, contrast: 0, highlights: 0, shadows: 0, saturation: 0, warmth: 0, auto: null })
export const NEUTRAL_ADJUST: PendingAdjust = Object.freeze({ quick: NEUTRAL_QUICK, look: 'none', intensity: 100 })
export const QUICK_SLIDERS: readonly { readonly key: Exclude<keyof QuickAdjust, 'auto'>; readonly label: string }[] = Object.freeze([
  { key: 'exposure', label: 'Exposure' },
  { key: 'brightness', label: 'Brightness' },
  { key: 'contrast', label: 'Contrast' },
  { key: 'highlights', label: 'Highlights' },
  { key: 'shadows', label: 'Shadows' },
  { key: 'saturation', label: 'Saturation' },
  { key: 'warmth', label: 'Warmth' },
])
/** Largest preview proxy (design 4.4). */
export const PREVIEW_MAX_PIXELS = 2_500_000
const THUMBNAIL_EDGE = 128

export function isNeutralQuick(quick: QuickAdjust): boolean {
  return !quick.auto && QUICK_SLIDERS.every(({ key }) => !quick[key])
}

export function hasLook(adjust: PendingAdjust): boolean {
  return adjust.look !== 'none' && adjust.intensity > 0
}

export function isNeutralAdjust(adjust: PendingAdjust): boolean {
  return isNeutralQuick(adjust.quick) && !hasLook(adjust)
}

export function quickSpec(quick: QuickAdjust): AdjustmentSpec {
  return { type: 'quick', ...quick }
}

export interface AdjustPreviewOptions {
  readonly active: boolean
  readonly canvasRef: RefObject<HTMLCanvasElement | null>
  readonly overlayRef: RefObject<HTMLCanvasElement | null>
  readonly pending: PendingAdjust
  /** Device pixels per image pixel on screen (zoom * devicePixelRatio). */
  readonly displayScale: number
  /** Changes whenever the image content changes (document id and revision). */
  readonly sourceKey: string
}

export interface AdjustPreview {
  /** The overlay shows the adjusted preview (false while neutral or not ready). */
  readonly showing: boolean
  /** The unadjusted proxy (Auto analyses it). */
  readonly proxy: PixelBuffer | null
  readonly thumbnails: Readonly<Partial<Record<LookId, string>>>
}

interface ProxyState {
  readonly key: string
  readonly scale: number
  readonly buffer: PixelBuffer
}

export function useAdjustPreview(options: AdjustPreviewOptions): AdjustPreview {
  const { active, canvasRef, overlayRef, pending, displayScale, sourceKey } = options
  const [proxy, setProxy] = useState<ProxyState | null>(null)
  const [showing, setShowing] = useState(false)
  const [thumbnails, setThumbnails] = useState<Partial<Record<LookId, string>>>({})
  const lookAbortRef = useRef<AbortController | null>(null)
  const requestRef = useRef(0)

  // Build (or rebuild) the proxy: on open, when the image changes, or after a large zoom change.
  useEffect(() => {
    if (!active) return
    const canvas = canvasRef.current
    if (!canvas || !canvas.width || !canvas.height) return
    const scale = Math.max(0.01, Math.min(1, displayScale))
    if (proxy && proxy.key === sourceKey && scale <= proxy.scale * 1.5 && scale >= proxy.scale / 3) return
    const timer = window.setTimeout(() => {
      try {
        const size = proxySize(canvas.width, canvas.height, scale, PREVIEW_MAX_PIXELS)
        setProxy({ key: sourceKey, scale, buffer: canvasProxy(canvas, size) })
      } catch {
        setProxy(null)
      }
    }, proxy ? 150 : 0)
    return () => window.clearTimeout(timer)
  }, [active, canvasRef, displayScale, proxy, sourceKey])

  // Look thumbnails from a 128 px proxy (shown at 64 CSS px).
  useEffect(() => {
    if (!active) return
    const canvas = canvasRef.current
    if (!canvas || !canvas.width || !canvas.height) return
    let cancelled = false
    const timer = window.setTimeout(() => {
      try {
        const size = proxySize(canvas.width, canvas.height, THUMBNAIL_EDGE / Math.max(canvas.width, canvas.height), THUMBNAIL_EDGE * THUMBNAIL_EDGE)
        const base = canvasProxy(canvas, size)
        const next: Partial<Record<LookId, string>> = {}
        for (const look of LOOK_IDS) {
          if (cancelled) return
          next[look] = pixelsToDataUrl(look === 'none' ? base : applyLook(base, look, 100))
        }
        if (!cancelled) setThumbnails(next)
      } catch {
        if (!cancelled) setThumbnails({})
      }
    }, 60)
    return () => {
      cancelled = true
      window.clearTimeout(timer)
    }
  }, [active, canvasRef, sourceKey])

  // Paint the preview for the pending settings.
  useEffect(() => {
    const overlay = overlayRef.current
    const request = ++requestRef.current
    lookAbortRef.current?.abort()
    lookAbortRef.current = null
    if (!active || !proxy || !overlay || isNeutralAdjust(pending)) {
      setShowing(false)
      return
    }
    const { width, height } = proxy.buffer
    const data = new Uint8ClampedArray(proxy.buffer.data)
    if (!isNeutralQuick(pending.quick)) compileAdjustment(quickSpec(pending.quick)).apply(data)
    const paint = (pixels: PixelBuffer) => {
      if (request !== requestRef.current) return
      if (overlay.width !== pixels.width || overlay.height !== pixels.height) {
        overlay.width = pixels.width
        overlay.height = pixels.height
      }
      putBuffer(overlay, pixels, 0, 0)
      setShowing(true)
    }
    if (!hasLook(pending)) {
      paint({ width, height, data })
      return
    }
    const controller = new AbortController()
    lookAbortRef.current = controller
    getImagingClient()
      .run('look', { src: { width, height, data }, look: pending.look, intensity: pending.intensity }, { signal: controller.signal })
      .then(paint)
      .catch((error: unknown) => {
        if ((error as Error)?.name === 'AbortError' || request !== requestRef.current) return
        // The worker failed: preview on this thread instead.
        try { paint(applyLook({ width, height, data }, pending.look, pending.intensity)) } catch { setShowing(false) }
      })
    return () => controller.abort()
  }, [active, overlayRef, pending, proxy])

  // Closing the panel frees the proxy and the overlay's backing store. The overlay element is captured while
  // the panel is open: by the time the cleanup runs, React has already cleared the ref of the unmounted canvas.
  useEffect(() => {
    if (!active) {
      lookAbortRef.current?.abort()
      lookAbortRef.current = null
      setProxy(null)
      setShowing(false)
      setThumbnails({})
      return
    }
    const overlay = overlayRef.current
    return () => releaseCanvas(overlay)
  }, [active, overlayRef])

  return { showing, proxy: proxy?.buffer ?? null, thumbnails }
}
