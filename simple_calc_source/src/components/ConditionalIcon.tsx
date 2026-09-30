import { memo } from 'react'
import type { ReactNode } from 'react'
import { iconSetSize, normalizeIconSetName } from '../lib/conditional-format'

// Inline-SVG icons for Excel's conditional-format icon sets (16 x 16 design grid). Glyphs are
// built once per (set, index) and reused, so painting thousands of grid cells stays cheap.

const GREEN = '#2F9E44'
const YELLOW = '#F0AE1B'
const RED = '#D6392B'
const GRAY = '#80868F'
const BLACK = '#2A2A2A'
const BLUE = '#3E6FB0'
const EMPTY = '#CCD3DD'
const GOLD = '#F2B50F'
const GOLD_EDGE = '#C08D00'

const DARKER: Record<string, string> = {
  [GREEN]: '#23803A',
  [YELLOW]: '#C98D0B',
  [RED]: '#AE2A1F',
  [BLACK]: '#111111',
  [GRAY]: '#646A72',
  '#F2A3A3': '#D98080',
}

const ARROW = 'M8 1.4 L14.6 8 H10.35 V14.6 H5.65 V8 H1.4 Z'
const STAR = 'M8 1.6 L9.65 6.33 L14.66 6.44 L10.66 9.47 L12.11 14.26 L8 11.4 L3.89 14.26 L5.34 9.47 L1.34 6.44 L6.35 6.33 Z'
const HALF_STAR = 'M8 1.6 L6.35 6.33 L1.34 6.44 L5.34 9.47 L3.89 14.26 L8 11.4 Z'

function arrow(color: string, rotate: number) {
  return <path d={ARROW} fill={color} transform={rotate ? `rotate(${rotate} 8 8)` : undefined} />
}

function circle(color: string) {
  return <circle cx="8" cy="8" r="6.2" fill={color} stroke={DARKER[color] || color} strokeWidth="0.9" />
}

function rimmed(color: string) {
  return (
    <>
      <rect x="1.2" y="1.2" width="13.6" height="13.6" rx="3.2" fill="#3B3B3B" />
      <circle cx="8" cy="8" r="4.7" fill={color} />
    </>
  )
}

function flag(color: string) {
  return (
    <>
      <path d="M3.6 1.4 V14.6" stroke="#5A5A5A" strokeWidth="1.5" strokeLinecap="round" />
      <path d="M4.3 2 H13.6 L11.2 5.3 L13.6 8.6 H4.3 Z" fill={color} />
    </>
  )
}

function symbol(kind: 'x' | '!' | 'check', circled: boolean) {
  const color = kind === 'x' ? RED : kind === '!' ? YELLOW : GREEN
  const ink = circled ? '#FFFFFF' : color
  const width = circled ? 1.9 : 2.6
  const glyph = kind === 'x'
    ? <path d={circled ? 'M5.4 5.4 L10.6 10.6 M10.6 5.4 L5.4 10.6' : 'M3.6 3.6 L12.4 12.4 M12.4 3.6 L3.6 12.4'} stroke={ink} strokeWidth={width} strokeLinecap="round" />
    : kind === '!'
      ? (
        <>
          <path d={circled ? 'M8 4.1 V8.9' : 'M8 2 V10'} stroke={ink} strokeWidth={width} strokeLinecap="round" />
          <circle cx="8" cy={circled ? 11.5 : 13.6} r={circled ? 1.05 : 1.45} fill={ink} />
        </>
      )
      : <path d={circled ? 'M4.7 8.3 L7.1 10.6 L11.3 5.6' : 'M2.6 8.6 L6.4 12.3 L13.4 3.8'} stroke={ink} strokeWidth={width} strokeLinecap="round" strokeLinejoin="round" fill="none" />
  if (!circled) return glyph
  return (
    <>
      <circle cx="8" cy="8" r="6.6" fill={color} />
      {glyph}
    </>
  )
}

function star(fill: 0 | 1 | 2) {
  return (
    <>
      <path d={STAR} fill={fill === 2 ? GOLD : '#FFFFFF'} stroke={GOLD_EDGE} strokeWidth="0.9" strokeLinejoin="round" />
      {fill === 1 && <path d={HALF_STAR} fill={GOLD} />}
    </>
  )
}

function triangle(index: number) {
  if (index === 0) return <path d="M8 13.4 L1.8 3.4 H14.2 Z" fill={RED} />
  if (index === 1) return <rect x="2.4" y="6.5" width="11.2" height="3" rx="0.6" fill={YELLOW} />
  return <path d="M8 2.6 L14.2 12.6 H1.8 Z" fill={GREEN} />
}

function sign(index: number) {
  if (index === 0) return <path d="M8 1.6 L14.4 8 L8 14.4 L1.6 8 Z" fill={RED} />
  if (index === 1) return <path d="M8 1.8 L14.8 13.9 H1.2 Z" fill={YELLOW} strokeLinejoin="round" />
  return circle(GREEN)
}

function rating(filled: number) {
  const bars = [
    { x: 1.4, h: 4 },
    { x: 5.1, h: 7 },
    { x: 8.8, h: 10 },
    { x: 12.5, h: 13 },
  ]
  return (
    <>
      {bars.map((bar, index) => (
        <rect key={index} x={bar.x} y={14.6 - bar.h} width="2.6" height={bar.h} rx="0.5" fill={index < filled ? BLUE : EMPTY} />
      ))}
    </>
  )
}

function quarters(index: number) {
  const radius = 6.1
  const fraction = Math.min(4, Math.max(0, index)) / 4
  let pie: ReactNode = null
  if (fraction >= 1) pie = <circle cx="8" cy="8" r={radius} fill="#3F3F3F" />
  else if (fraction > 0) {
    const angle = fraction * Math.PI * 2
    const x = 8 + radius * Math.sin(angle)
    const y = 8 - radius * Math.cos(angle)
    pie = <path d={`M8 8 L8 ${8 - radius} A${radius} ${radius} 0 ${fraction > 0.5 ? 1 : 0} 1 ${x.toFixed(3)} ${y.toFixed(3)} Z`} fill="#3F3F3F" />
  }
  return (
    <>
      <circle cx="8" cy="8" r={radius} fill="#FFFFFF" stroke="#3F3F3F" strokeWidth="1.2" />
      {pie}
    </>
  )
}

function boxes(filled: number) {
  const cells = [
    { x: 1.6, y: 8.6 },
    { x: 8.6, y: 8.6 },
    { x: 1.6, y: 1.6 },
    { x: 8.6, y: 1.6 },
  ]
  return (
    <>
      {cells.map((cell, index) => (
        <rect key={index} x={cell.x} y={cell.y} width="5.8" height="5.8" rx="1" fill={index < filled ? BLUE : EMPTY} />
      ))}
    </>
  )
}

function glyphFor(set: string, index: number): ReactNode {
  switch (set) {
    case '3Arrows': return [arrow(RED, 180), arrow(YELLOW, 90), arrow(GREEN, 0)][index]
    case '3ArrowsGray': return [arrow(GRAY, 180), arrow(GRAY, 90), arrow(GRAY, 0)][index]
    case '4Arrows': return [arrow(RED, 180), arrow(YELLOW, 135), arrow(YELLOW, 45), arrow(GREEN, 0)][index]
    case '4ArrowsGray': return [arrow(GRAY, 180), arrow(GRAY, 135), arrow(GRAY, 45), arrow(GRAY, 0)][index]
    case '5Arrows': return [arrow(RED, 180), arrow(YELLOW, 135), arrow(YELLOW, 90), arrow(YELLOW, 45), arrow(GREEN, 0)][index]
    case '5ArrowsGray': return [arrow(GRAY, 180), arrow(GRAY, 135), arrow(GRAY, 90), arrow(GRAY, 45), arrow(GRAY, 0)][index]
    case '3Flags': return flag([RED, YELLOW, GREEN][index] ?? GREEN)
    case '3TrafficLights1': return circle([RED, YELLOW, GREEN][index] ?? GREEN)
    case '3TrafficLights2': return rimmed([RED, YELLOW, GREEN][index] ?? GREEN)
    case '4TrafficLights': return circle([BLACK, RED, YELLOW, GREEN][index] ?? GREEN)
    case '4RedToBlack': return circle([BLACK, GRAY, '#F2A3A3', RED][index] ?? RED)
    case '3Signs': return sign(index)
    case '3Symbols': return symbol((['x', '!', 'check'] as const)[index] ?? 'check', true)
    case '3Symbols2': return symbol((['x', '!', 'check'] as const)[index] ?? 'check', false)
    case '3Stars': return star((Math.min(2, index) as 0 | 1 | 2))
    case '3Triangles': return triangle(index)
    case '4Rating': return rating(index + 1)
    case '5Rating': return rating(index)
    case '5Quarters': return quarters(index)
    case '5Boxes': return boxes(index)
    default: return null
  }
}

const GLYPHS = new Map<string, ReactNode>()

function cachedGlyph(set: string, index: number): ReactNode {
  const key = `${set}:${index}`
  if (GLYPHS.has(key)) return GLYPHS.get(key)
  const size = iconSetSize(set)
  const glyph = index >= 0 && index < size ? glyphFor(set, index) ?? null : null
  GLYPHS.set(key, glyph)
  return glyph
}

export interface ConditionalIconProps {
  /** Icon set name ("3Arrows", "5Quarters", ...). */
  set: string
  /** Icon index within the set, 0 = lowest. */
  index: number
  size?: number
  /** Accessible label; decorative (aria-hidden) when omitted. */
  title?: string
  className?: string
}

/** One icon of an Excel icon set, drawn with inline SVG. */
export const ConditionalIcon = memo(function ConditionalIcon({ set, index, size = 14, title, className }: ConditionalIconProps) {
  const glyph = cachedGlyph(normalizeIconSetName(set), Math.floor(index))
  if (!glyph) return null
  return (
    <svg
      className={className}
      width={size}
      height={size}
      viewBox="0 0 16 16"
      focusable="false"
      role={title ? 'img' : undefined}
      aria-label={title}
      aria-hidden={title ? undefined : true}
    >
      {glyph}
    </svg>
  )
})
