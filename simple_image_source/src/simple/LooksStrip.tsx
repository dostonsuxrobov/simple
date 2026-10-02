// src/simple/LooksStrip.tsx (WP8)
// One-click Looks (design 4.5): nine looks from src/imaging/looks.ts with thumbnails rendered from a 128 px
// proxy, and an Intensity slider (0..100, default 100). The pick is previewed live and applied with the
// Adjust panel's Done (same pipeline, one undo step).
import type { LookId } from '../imaging/types.ts'
import { LOOKS, LOOK_IDS } from '../imaging/looks.ts'

export interface LooksStripProps {
  readonly look: LookId
  readonly intensity: number
  readonly thumbnails: Readonly<Partial<Record<LookId, string>>>
  readonly disabled: boolean
  readonly onLook: (look: LookId) => void
  readonly onIntensity: (value: number) => void
}

export function LooksStrip(props: LooksStripProps) {
  const { look, intensity, thumbnails, disabled } = props
  return (
    <div className="looks-strip">
      <div className="looks-grid" role="radiogroup" aria-label="Looks">
        {LOOK_IDS.map((id) => (
          <button
            key={id}
            type="button"
            role="radio"
            aria-checked={look === id}
            className={look === id ? 'active' : ''}
            disabled={disabled}
            data-look={id}
            onClick={() => props.onLook(id)}
          >
            <span className="look-thumb">{thumbnails[id] ? <img src={thumbnails[id]} alt="" draggable={false} /> : null}</span>
            <span className="look-label">{LOOKS[id].label}</span>
          </button>
        ))}
      </div>
      <label className="adjust-slider look-intensity">
        <span>Intensity</span>
        <input
          type="range"
          min={0}
          max={100}
          step={1}
          value={intensity}
          disabled={disabled || look === 'none'}
          onChange={(event) => props.onIntensity(Number(event.target.value))}
          onDoubleClick={() => props.onIntensity(100)}
        />
        <output>{intensity}</output>
      </label>
    </div>
  )
}
