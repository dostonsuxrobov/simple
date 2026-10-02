// src/advanced/panels/LevelsEditor.tsx (WP6)
// Photoshop's Levels editor (design 5.10, 5.14): a channel menu (RGB, Red, Green, Blue), the histogram of
// the pixels the adjustment receives, the input slider with black, midtone (gamma) and white handles, the
// output slider, and the five numbers. Handles drag with the mouse and move with the arrow keys
// (Shift = 10); the midtone handle sits where input maps to 50% grey (ib + (iw - ib) * 0.5^gamma), so it
// follows the black and white handles like in Photoshop.
// HistogramView (also used by Curves and Threshold) draws 256 bins, clipping rare spikes so the shape of
// the rest stays readable.
import { useEffect, useRef, useState } from 'react'
import type { KeyboardEvent as ReactKeyboardEvent, PointerEvent as ReactPointerEvent } from 'react'
import type { AdjustmentSpec, Histogram, LevelsChannel } from '../../imaging/types.ts'
import type { ChangePhase } from './AdjustmentControls.tsx'
import { NumberInput, SelectField, clampNumber } from './AdjustmentControls.tsx'

export type HistogramChannel = 'rgb' | 'red' | 'green' | 'blue' | 'luma'

const CHANNEL_COLORS: Readonly<Record<HistogramChannel, string>> = Object.freeze({
  rgb: '#4d4d4d',
  luma: '#4d4d4d',
  red: '#d33a3a',
  green: '#2f9a3f',
  blue: '#3460d6',
})

function binsOf(histogram: Histogram, channel: HistogramChannel): Uint32Array {
  if (channel === 'red') return histogram.red
  if (channel === 'green') return histogram.green
  if (channel === 'blue') return histogram.blue
  return histogram.luma
}

export interface HistogramViewProps {
  readonly histogram: Histogram | null
  readonly channel: HistogramChannel
  /** Canvas height in bins' units (the width is always 256). */
  readonly height?: number
  /** Draws a vertical line at this value (Threshold). */
  readonly marker?: number
  readonly className?: string
}

/** 256 vertical bars; the tallest few bins are clipped so one spike does not flatten the rest. */
export function HistogramView({ histogram, channel, height = 72, marker, className }: HistogramViewProps) {
  const ref = useRef<HTMLCanvasElement>(null)
  useEffect(() => {
    const canvas = ref.current
    const context = canvas?.getContext('2d')
    if (!canvas || !context) return
    context.clearRect(0, 0, canvas.width, canvas.height)
    if (histogram && histogram.count > 0) {
      const bins = binsOf(histogram, channel)
      const sorted = Array.from(bins).sort((a, b) => a - b)
      const cap = Math.max(1, Math.min(sorted[255], sorted[251] * 1.35))
      context.fillStyle = CHANNEL_COLORS[channel]
      for (let x = 0; x < 256; x += 1) {
        const value = Math.min(1, bins[x] / cap)
        if (value <= 0) continue
        const bar = Math.max(1, Math.round(value * height))
        context.fillRect(x, height - bar, 1, bar)
      }
    }
    if (marker !== undefined && Number.isFinite(marker)) {
      context.fillStyle = '#e0362c'
      context.fillRect(Math.round(clampNumber(marker, 0, 255)), 0, 1, height)
    }
  }, [channel, height, histogram, marker])
  useEffect(() => {
    const canvas = ref.current
    return () => {
      if (canvas) {
        canvas.width = 1
        canvas.height = 1
      }
    }
  }, [])
  return (
    <div className={`ae-histogram ${className ?? ''}`}>
      <canvas ref={ref} width={256} height={height} aria-label={histogram && histogram.count > 0 ? 'Histogram' : 'No pixels to measure'} role="img" />
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// Slider with triangle handles
// ---------------------------------------------------------------------------------------------

interface HandleSpec {
  readonly id: string
  readonly label: string
  /** 0..255 position. */
  readonly position: number
  readonly tone: 'black' | 'gray' | 'white'
  /** Text for assistive technology. */
  readonly valueText: string
  /** New model value for a pointer position 0..255. */
  readonly fromPosition: (position: number) => void
  /** Keyboard step (direction -1 / 1, Shift = big). */
  readonly step: (direction: 1 | -1, big: boolean) => void
}

function HandleTrack({ handles, gradient, disabled, onInteractive, onCommit }: {
  readonly handles: readonly HandleSpec[]
  readonly gradient: string
  readonly disabled?: boolean
  readonly onInteractive?: (active: boolean) => void
  /** A drag ended (commit the last value). */
  readonly onCommit: () => void
}) {
  const trackRef = useRef<HTMLDivElement>(null)
  const startDrag = (event: ReactPointerEvent<HTMLButtonElement>, handle: HandleSpec) => {
    if (disabled || event.button !== 0) return
    event.preventDefault()
    const element = event.currentTarget
    element.focus({ preventScroll: true })
    try {
      element.setPointerCapture(event.pointerId)
    } catch {
      // ignore
    }
    onInteractive?.(true)
    const report = (clientX: number) => {
      const rect = trackRef.current?.getBoundingClientRect()
      if (!rect || rect.width <= 0) return
      handle.fromPosition(clampNumber(((clientX - rect.left) / rect.width) * 255, 0, 255))
    }
    const move = (moveEvent: PointerEvent) => report(moveEvent.clientX)
    const end = () => {
      element.removeEventListener('pointermove', move)
      element.removeEventListener('pointerup', end)
      element.removeEventListener('pointercancel', end)
      onInteractive?.(false)
      onCommit()
    }
    element.addEventListener('pointermove', move)
    element.addEventListener('pointerup', end)
    element.addEventListener('pointercancel', end)
  }
  const onKey = (event: ReactKeyboardEvent<HTMLButtonElement>, handle: HandleSpec) => {
    if (event.key === 'ArrowLeft' || event.key === 'ArrowDown') {
      event.preventDefault()
      handle.step(-1, event.shiftKey)
    } else if (event.key === 'ArrowRight' || event.key === 'ArrowUp') {
      event.preventDefault()
      handle.step(1, event.shiftKey)
    }
  }
  return (
    <div ref={trackRef} className="ae-lv-track">
      <div className="ae-lv-bar" style={{ background: gradient }} />
      {handles.map((handle) => (
        <button
          key={handle.id}
          type="button"
          role="slider"
          aria-label={handle.label}
          aria-valuetext={handle.valueText}
          aria-valuemin={0}
          aria-valuemax={255}
          aria-valuenow={Math.round(handle.position)}
          className={`ae-lv-handle is-${handle.tone}`}
          style={{ left: `${(handle.position / 255) * 100}%` }}
          disabled={disabled}
          onPointerDown={(event) => startDrag(event, handle)}
          onKeyDown={(event) => onKey(event, handle)}
        />
      ))}
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// The editor
// ---------------------------------------------------------------------------------------------

type LevelsKey = 'rgb' | 'red' | 'green' | 'blue'

const CHANNELS: readonly (readonly [LevelsKey, string])[] = [['rgb', 'RGB'], ['red', 'Red'], ['green', 'Green'], ['blue', 'Blue']]

const MIN_GAMMA = 0.1
const MAX_GAMMA = 9.99

/** Input value whose output is 50% grey: where Photoshop draws the midtone handle. */
export function midtonePosition(channel: LevelsChannel): number {
  return channel.inBlack + (channel.inWhite - channel.inBlack) * Math.pow(0.5, channel.gamma)
}

/** Gamma that puts the midtone handle at `position`. */
export function gammaForPosition(channel: LevelsChannel, position: number): number {
  const span = channel.inWhite - channel.inBlack
  if (span <= 0) return channel.gamma
  const t = clampNumber((position - channel.inBlack) / span, 0.001, 0.999)
  return Number(clampNumber(Math.log(t) / Math.log(0.5), MIN_GAMMA, MAX_GAMMA).toFixed(2))
}

export interface LevelsEditorProps {
  readonly spec: Extract<AdjustmentSpec, { type: 'levels' }>
  readonly onChange: (next: AdjustmentSpec, phase: ChangePhase) => void
  readonly histogram?: Histogram | null
  readonly disabled?: boolean
  readonly onInteractive?: (active: boolean) => void
}

export function LevelsEditor({ spec, onChange, histogram, disabled, onInteractive }: LevelsEditorProps) {
  const [key, setKey] = useState<LevelsKey>('rgb')
  const channel = spec[key]
  // Drags report from closures that outlive one render: read the newest values from refs.
  const specRef = useRef(spec)
  specRef.current = spec
  const set = (patch: Partial<LevelsChannel>, phase: ChangePhase) => {
    const current = specRef.current
    const next = { ...current, [key]: { ...current[key], ...patch } }
    specRef.current = next
    onChange(next, phase)
  }
  const live = (patch: Partial<LevelsChannel>) => set(patch, 'live')
  const commitLast = () => onChange(specRef.current, 'commit')
  const now = () => specRef.current[key]

  const inputHandles: HandleSpec[] = [
    {
      id: 'black',
      label: 'Input black point',
      position: channel.inBlack,
      tone: 'black',
      valueText: String(channel.inBlack),
      fromPosition: (position) => live({ inBlack: Math.round(clampNumber(position, 0, now().inWhite - 2)) }),
      step: (direction, big) => set({ inBlack: clampNumber(now().inBlack + direction * (big ? 10 : 1), 0, now().inWhite - 2) }, 'commit'),
    },
    {
      id: 'gray',
      label: 'Input midtones (gamma)',
      position: midtonePosition(channel),
      tone: 'gray',
      valueText: channel.gamma.toFixed(2),
      fromPosition: (position) => live({ gamma: gammaForPosition(now(), position) }),
      step: (direction, big) => set({ gamma: Number(clampNumber(now().gamma - direction * (big ? 0.1 : 0.01), MIN_GAMMA, MAX_GAMMA).toFixed(2)) }, 'commit'),
    },
    {
      id: 'white',
      label: 'Input white point',
      position: channel.inWhite,
      tone: 'white',
      valueText: String(channel.inWhite),
      fromPosition: (position) => live({ inWhite: Math.round(clampNumber(position, now().inBlack + 2, 255)) }),
      step: (direction, big) => set({ inWhite: clampNumber(now().inWhite + direction * (big ? 10 : 1), now().inBlack + 2, 255) }, 'commit'),
    },
  ]
  const outputHandles: HandleSpec[] = [
    {
      id: 'out-black',
      label: 'Output black',
      position: channel.outBlack,
      tone: 'black',
      valueText: String(channel.outBlack),
      fromPosition: (position) => live({ outBlack: Math.round(position) }),
      step: (direction, big) => set({ outBlack: clampNumber(now().outBlack + direction * (big ? 10 : 1), 0, 255) }, 'commit'),
    },
    {
      id: 'out-white',
      label: 'Output white',
      position: channel.outWhite,
      tone: 'white',
      valueText: String(channel.outWhite),
      fromPosition: (position) => live({ outWhite: Math.round(position) }),
      step: (direction, big) => set({ outWhite: clampNumber(now().outWhite + direction * (big ? 10 : 1), 0, 255) }, 'commit'),
    },
  ]
  const histogramChannel: HistogramChannel = key === 'rgb' ? 'luma' : key
  const tint = key === 'rgb' ? '#fff' : key === 'red' ? '#ff4a4a' : key === 'green' ? '#45d45a' : '#4a78ff'
  return (
    <div className="ae-levels">
      <SelectField label="Channel" inline value={key} options={CHANNELS} disabled={disabled} onChange={setKey} />
      <HistogramView histogram={histogram ?? null} channel={histogramChannel} height={80} className="is-levels" />
      <HandleTrack handles={inputHandles} gradient={`linear-gradient(to right, #000, ${tint})`} disabled={disabled} onInteractive={onInteractive} onCommit={commitLast} />
      <div className="ae-lv-fields">
        <NumberInput label="Input black point" value={channel.inBlack} min={0} max={channel.inWhite - 2} disabled={disabled} onChange={(inBlack) => set({ inBlack }, 'commit')} />
        <NumberInput label="Input midtones (gamma)" value={channel.gamma} min={MIN_GAMMA} max={MAX_GAMMA} step={0.01} digits={2} fixed disabled={disabled} onChange={(gamma) => set({ gamma }, 'commit')} />
        <NumberInput label="Input white point" value={channel.inWhite} min={channel.inBlack + 2} max={255} disabled={disabled} onChange={(inWhite) => set({ inWhite }, 'commit')} />
      </div>
      <div className="ae-lv-caption">Output Levels</div>
      <HandleTrack handles={outputHandles} gradient={`linear-gradient(to right, #000, ${tint})`} disabled={disabled} onInteractive={onInteractive} onCommit={commitLast} />
      <div className="ae-lv-fields is-output">
        <NumberInput label="Output black" value={channel.outBlack} min={0} max={255} disabled={disabled} onChange={(outBlack) => set({ outBlack }, 'commit')} />
        <NumberInput label="Output white" value={channel.outWhite} min={0} max={255} disabled={disabled} onChange={(outWhite) => set({ outWhite }, 'commit')} />
      </div>
    </div>
  )
}
