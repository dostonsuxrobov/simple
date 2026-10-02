// src/simple/AdjustPanel.tsx (WP8)
// The Adjust panel (design 4.4 / 4.5), shown in the inspector column in place of the image details:
//   - "Light & color": Exposure, Brightness, Contrast, Highlights, Shadows, Saturation and Warmth (-100..100,
//     double-click resets one), Auto (hidden Levels plus visible Saturation/Warmth), Reset, Compare (hold the
//     button or press \);
//   - "Looks": nine one-click looks with an intensity slider.
// Cancel discards; Done applies everything at full resolution as ONE undo step. Closing the panel any other
// way (another tool, Save, Export, Print, Advanced) applies too, so a pending adjustment is never lost.
import { Check, Eye, SlidersHorizontal, Sparkles, Wand2, X } from 'lucide-react'
import type { LookId, QuickAdjust } from '../imaging/types.ts'
import type { PendingAdjust } from './useAdjustPreview.ts'
import { QUICK_SLIDERS, isNeutralAdjust } from './useAdjustPreview.ts'
import { LooksStrip } from './LooksStrip.tsx'

export type AdjustTab = 'light' | 'looks'

export interface AdjustPanelProps {
  readonly pending: PendingAdjust
  readonly tab: AdjustTab
  readonly thumbnails: Readonly<Partial<Record<LookId, string>>>
  readonly busy: boolean
  readonly progress: number | null
  readonly comparing: boolean
  readonly autoAvailable: boolean
  readonly onTab: (tab: AdjustTab) => void
  readonly onQuick: (quick: QuickAdjust) => void
  readonly onLook: (look: LookId) => void
  readonly onIntensity: (value: number) => void
  readonly onAuto: () => void
  readonly onReset: () => void
  readonly onCompare: (on: boolean) => void
  readonly onCancel: () => void
  readonly onDone: () => void
}

export function AdjustPanel(props: AdjustPanelProps) {
  const { pending, tab, busy, progress, comparing } = props
  const neutral = isNeutralAdjust(pending)
  const autoActive = Boolean(pending.quick.auto)
  return (
    <div className="adjust-panel" aria-label="Adjust">
      <div className="inspector-heading">
        <SlidersHorizontal />
        <div><strong>Adjust</strong><span>{busy ? `Applying… ${Math.round((progress ?? 0) * 100)}%` : 'Preview updates as you drag'}</span></div>
        <button className="adjust-close" type="button" aria-label="Apply and close" title="Apply and close" disabled={busy} onClick={props.onDone}><X /></button>
      </div>
      <div className="adjust-tabs" role="tablist" aria-label="Adjust">
        <button type="button" role="tab" aria-selected={tab === 'light'} className={tab === 'light' ? 'active' : ''} onClick={() => props.onTab('light')}>Light &amp; color</button>
        <button type="button" role="tab" aria-selected={tab === 'looks'} className={tab === 'looks' ? 'active' : ''} onClick={() => props.onTab('looks')}><Sparkles />Looks</button>
      </div>
      {tab === 'light' ? (
        <div className="adjust-sliders" role="tabpanel">
          <div className="adjust-row">
            <button type="button" className={`adjust-auto ${autoActive ? 'active' : ''}`} disabled={busy || !props.autoAvailable} onClick={props.onAuto} title="Automatic levels, colour and contrast"><Wand2 />Auto</button>
            <button type="button" disabled={busy || neutral} onClick={props.onReset}>Reset</button>
          </div>
          {QUICK_SLIDERS.map(({ key, label }) => (
            <label key={key} className="adjust-slider" data-adjust={key}>
              <span>{label}</span>
              <input
                type="range"
                min={-100}
                max={100}
                step={1}
                value={pending.quick[key]}
                disabled={busy}
                onChange={(event) => props.onQuick({ ...pending.quick, [key]: Number(event.target.value) })}
                onDoubleClick={() => props.onQuick({ ...pending.quick, [key]: 0 })}
              />
              <output>{pending.quick[key] > 0 ? `+${pending.quick[key]}` : pending.quick[key]}</output>
            </label>
          ))}
        </div>
      ) : (
        <div role="tabpanel">
          <LooksStrip look={pending.look} intensity={pending.intensity} thumbnails={props.thumbnails} disabled={busy} onLook={props.onLook} onIntensity={props.onIntensity} />
        </div>
      )}
      <button
        type="button"
        className={`adjust-compare ${comparing ? 'active' : ''}`}
        disabled={busy || neutral}
        title={'Hold to see the original (or press \\)'}
        onPointerDown={(event) => { event.currentTarget.setPointerCapture(event.pointerId); props.onCompare(true) }}
        onPointerUp={() => props.onCompare(false)}
        onPointerCancel={() => props.onCompare(false)}
        onKeyDown={(event) => { if (event.key === ' ' || event.key === 'Enter') { event.preventDefault(); props.onCompare(true) } }}
        onKeyUp={(event) => { if (event.key === ' ' || event.key === 'Enter') props.onCompare(false) }}
        onBlur={() => props.onCompare(false)}
      ><Eye />Compare</button>
      <div className="adjust-actions">
        <button type="button" disabled={busy} onClick={props.onCancel}>Cancel</button>
        <button type="button" className="adjust-done" disabled={busy} onClick={props.onDone}><Check />Done</button>
      </div>
    </div>
  )
}
