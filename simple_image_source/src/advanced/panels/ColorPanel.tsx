// src/advanced/panels/ColorPanel.tsx (WP6)
// The Color / Swatches panel (design 5.14):
//   - foreground and background chips: click one to edit it here (it opens the colour picker); the arrow
//     swaps them (X) and the small pair resets them to black and white (D);
//   - H S B or R G B sliders whose tracks show what each one does with the current colour, a hex field,
//     and "pick from screen" (EyeDropper API: samples anywhere, also outside the window);
//   - recent colours, and swatches kept per viewer (editor settings in localStorage): click = foreground,
//     Ctrl+click = background, Alt+click = delete; + adds the colour being edited.
import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import { ArrowLeftRight, Pipette, Plus } from 'lucide-react'
import type { Rgb8 } from '../../imaging/types.ts'
import { hsvToRgb, parseHex, rgbToHsv, toHex } from '../../imaging/color.ts'
import { MAX_SWATCHES, rememberColor, sameColor } from '../editorState.ts'
import type { PanelDefinition, PanelProps } from './PanelHost.tsx'
import { useEditorState } from './PanelHost.tsx'
import type { ChangePhase } from './AdjustmentControls.tsx'
import { SliderField } from './AdjustmentControls.tsx'
import { ColorPickerPopover, cssColor, eyeDropperAvailable, pickScreenColor } from './ColorPicker.tsx'

type Which = 'foreground' | 'background'
type Mode = 'hsb' | 'rgb'

const MODE_KEY = 'simple-image:advanced-color-mode'

function readMode(): Mode {
  try {
    return localStorage.getItem(MODE_KEY) === 'rgb' ? 'rgb' : 'hsb'
  } catch {
    return 'hsb'
  }
}

function writeMode(mode: Mode): void {
  try {
    localStorage.setItem(MODE_KEY, mode)
  } catch {
    // A per-viewer convenience only.
  }
}

function rgb(h: number, s: number, v: number): Rgb8 {
  const [r, g, b] = hsvToRgb(h, s, v)
  return { r: Math.round(r * 255), g: Math.round(g * 255), b: Math.round(b * 255) }
}

function hexOf(color: Rgb8): string {
  return toHex(color).toUpperCase()
}

interface Hsv {
  readonly h: number
  readonly s: number
  readonly v: number
}

/** HSB of a colour; greys keep the previous hue and black keeps hue and saturation (as in Photoshop). */
function hsvFrom(color: Rgb8, previous: Hsv | null): Hsv {
  const [h, s, v] = rgbToHsv(color.r / 255, color.g / 255, color.b / 255)
  if (!previous) return { h, s, v }
  if (v === 0) return { h: previous.h, s: previous.s, v }
  if (s === 0) return { h: previous.h, s, v }
  return { h, s, v }
}

function ColorPanel({ context, suspended }: PanelProps): ReactNode {
  const editor = useEditorState(context.editor)
  const [which, setWhich] = useState<Which>('foreground')
  const [mode, setModeState] = useState<Mode>(readMode)
  const [pickerFor, setPickerFor] = useState<HTMLElement | null>(null)
  const color = editor[which]
  const [hex, setHex] = useState(hexOf(color).slice(1))
  const hexEditing = useRef(false)
  const [hsv, setHsv] = useState<Hsv>(() => hsvFrom(color, null))

  useEffect(() => {
    if (!hexEditing.current) setHex(hexOf(color).slice(1))
    setHsv((current) => (sameColor(rgb(current.h, current.s, current.v), color) ? current : hsvFrom(color, current)))
  }, [color])

  const applyHsv = (next: Hsv, phase: ChangePhase) => {
    setHsv(next)
    apply(rgb(next.h, next.s, next.v), phase)
  }

  const apply = (next: Rgb8, phase: ChangePhase | 'commit') => {
    const state = context.editor.getState()
    if (sameColor(state[which], next) && phase !== 'commit') return
    const patch: { foreground?: Rgb8; background?: Rgb8; recentColors?: readonly Rgb8[] } = { [which]: next }
    if (phase === 'commit') patch.recentColors = rememberColor(state.recentColors, next)
    context.editor.update(patch)
  }

  const setMode = (next: Mode) => {
    setModeState(next)
    writeMode(next)
  }

  const commitHex = () => {
    hexEditing.current = false
    const parsed = parseHex(hex)
    if (parsed) apply(parsed, 'commit')
    else setHex(hexOf(color).slice(1))
  }

  const pick = async () => {
    const picked = await pickScreenColor()
    if (picked) apply(picked, 'commit')
  }

  const chip = (target: Which) => (
    <button
      type="button"
      className={`ae-cpanel-chip is-${target} ${which === target ? 'is-active' : ''}`}
      style={{ background: cssColor(editor[target]) }}
      title={`${target === 'foreground' ? 'Foreground' : 'Background'} colour ${hexOf(editor[target])}${which === target ? ' (click to open the picker)' : ' (click to edit)'}`}
      aria-label={`${target === 'foreground' ? 'Foreground' : 'Background'} colour ${hexOf(editor[target])}`}
      aria-pressed={which === target}
      disabled={suspended}
      onClick={(event) => {
        const element = event.currentTarget
        setWhich(target)
        setPickerFor((current) => (current === element ? null : element))
      }}
    />
  )

  const swatchClick = (swatch: Rgb8, index: number, event: { altKey: boolean; ctrlKey: boolean; metaKey: boolean }) => {
    if (event.altKey) {
      context.editor.update({ swatches: editor.swatches.filter((_, i) => i !== index) })
      return
    }
    const target: Which = event.ctrlKey || event.metaKey ? 'background' : 'foreground'
    const state = context.editor.getState()
    context.editor.update({ [target]: swatch, recentColors: rememberColor(state.recentColors, swatch) })
  }

  const hsbSliders = (
    <>
      <SliderField
        label="H"
        inline
        value={Math.round(hsv.h)}
        min={0}
        max={360}
        unit="°"
        disabled={suspended}
        track="linear-gradient(to right, #f00, #ff0, #0f0, #0ff, #00f, #f0f, #f00)"
        onChange={(value, phase) => applyHsv({ ...hsv, h: value % 360 }, phase)}
      />
      <SliderField
        label="S"
        inline
        value={Math.round(hsv.s * 100)}
        min={0}
        max={100}
        unit="%"
        disabled={suspended}
        track={`linear-gradient(to right, ${cssColor(rgb(hsv.h, 0, hsv.v))}, ${cssColor(rgb(hsv.h, 1, hsv.v))})`}
        onChange={(value, phase) => applyHsv({ ...hsv, s: value / 100 }, phase)}
      />
      <SliderField
        label="B"
        inline
        value={Math.round(hsv.v * 100)}
        min={0}
        max={100}
        unit="%"
        disabled={suspended}
        track={`linear-gradient(to right, #000, ${cssColor(rgb(hsv.h, hsv.s, 1))})`}
        onChange={(value, phase) => applyHsv({ ...hsv, v: value / 100 }, phase)}
      />
    </>
  )
  const rgbSliders = (
    <>
      <SliderField label="R" inline value={color.r} min={0} max={255} disabled={suspended} track={`linear-gradient(to right, ${cssColor({ ...color, r: 0 })}, ${cssColor({ ...color, r: 255 })})`} onChange={(r, phase) => apply({ ...color, r }, phase)} />
      <SliderField label="G" inline value={color.g} min={0} max={255} disabled={suspended} track={`linear-gradient(to right, ${cssColor({ ...color, g: 0 })}, ${cssColor({ ...color, g: 255 })})`} onChange={(g, phase) => apply({ ...color, g }, phase)} />
      <SliderField label="B" inline value={color.b} min={0} max={255} disabled={suspended} track={`linear-gradient(to right, ${cssColor({ ...color, b: 0 })}, ${cssColor({ ...color, b: 255 })})`} onChange={(b, phase) => apply({ ...color, b }, phase)} />
    </>
  )

  return (
    <div className="ae-cpanel">
      <div className="ae-cpanel-top">
        <div className="ae-cpanel-chips">
          {chip('background')}
          {chip('foreground')}
          <button type="button" className="ae-cpanel-swap" title={`Swap colours (${context.shortcut('edit.swap-colors') || 'X'})`} aria-label="Swap colours" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('edit.swap-colors')}>
            <ArrowLeftRight aria-hidden="true" />
          </button>
          <button type="button" className="ae-cpanel-default" title={`Default colours, black and white (${context.shortcut('edit.default-colors') || 'D'})`} aria-label="Default colours" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('edit.default-colors')}>
            <span /><span />
          </button>
        </div>
        <label className="ae-cpanel-hex" title={`${which === 'foreground' ? 'Foreground' : 'Background'} colour as hex`}>
          <span>#</span>
          <input
            type="text"
            aria-label={`${which === 'foreground' ? 'Foreground' : 'Background'} colour (hex)`}
            maxLength={7}
            spellCheck={false}
            value={hex}
            disabled={suspended}
            onFocus={(event) => {
              hexEditing.current = true
              event.currentTarget.select()
            }}
            onChange={(event) => setHex(event.currentTarget.value.replace(/^#/, '').toUpperCase())}
            onBlur={commitHex}
            onKeyDown={(event) => {
              if (event.key === 'Enter') {
                commitHex()
                context.focusCanvas()
              } else if (event.key === 'Escape') {
                hexEditing.current = false
                setHex(hexOf(color).slice(1))
              }
            }}
          />
        </label>
        {eyeDropperAvailable() && (
          <button type="button" className="ae-cpanel-pick" title="Pick a colour from anywhere on the screen" aria-label="Pick from screen" disabled={suspended} onClick={() => void pick()}>
            <Pipette aria-hidden="true" />
          </button>
        )}
        <button
          type="button"
          className="ae-cpanel-mode"
          title={mode === 'hsb' ? 'Showing hue, saturation and brightness (click for red, green and blue)' : 'Showing red, green and blue (click for hue, saturation and brightness)'}
          aria-label={mode === 'hsb' ? 'Show RGB sliders' : 'Show HSB sliders'}
          disabled={suspended}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => setMode(mode === 'hsb' ? 'rgb' : 'hsb')}
        >{mode === 'hsb' ? 'HSB' : 'RGB'}</button>
      </div>
      <div className="ae-cpanel-sliders">{mode === 'hsb' ? hsbSliders : rgbSliders}</div>
      {editor.recentColors.length > 0 && (
        <div className="ae-cpanel-strip">
          <span className="ae-cpanel-tag" aria-hidden="true">Recent</span>
          <div className="ae-swatches is-recent" role="group" aria-label="Recent colours">
            {editor.recentColors.slice(0, 9).map((recent, index) => (
              <button
                key={`${toHex(recent)}-${index}`}
                type="button"
                className="ae-swatch"
                style={{ background: cssColor(recent) }}
                title={`${hexOf(recent)} (click: ${which})`}
                aria-label={`Recent colour ${hexOf(recent)}`}
                disabled={suspended}
                onMouseDown={(event) => event.preventDefault()}
                onClick={() => apply(recent, 'commit')}
              />
            ))}
          </div>
        </div>
      )}
      <div className="ae-cpanel-strip">
        <div className="ae-swatches" role="group" aria-label="Swatches">
          {editor.swatches.map((swatch, index) => (
            <button
              key={`${toHex(swatch)}-${index}`}
              type="button"
              className="ae-swatch"
              style={{ background: cssColor(swatch) }}
              title={`${hexOf(swatch)}: click for foreground, Ctrl+click for background, Alt+click to delete`}
              aria-label={`Swatch ${hexOf(swatch)}`}
              disabled={suspended}
              onMouseDown={(event) => event.preventDefault()}
              onClick={(event) => swatchClick(swatch, index, event)}
            />
          ))}
          <button
            type="button"
            className="ae-swatch ae-swatch-add"
            title={`Add the ${which} colour to the swatches`}
            aria-label="Add swatch"
            disabled={suspended || editor.swatches.length >= MAX_SWATCHES || editor.swatches.some((swatch) => sameColor(swatch, color))}
            onMouseDown={(event) => event.preventDefault()}
            onClick={() => context.editor.update({ swatches: [...editor.swatches, color].slice(-MAX_SWATCHES) })}
          ><Plus aria-hidden="true" /></button>
        </div>
      </div>
      {pickerFor && !suspended && (
        <ColorPickerPopover
          anchor={pickerFor}
          value={color}
          title={which === 'foreground' ? 'Foreground colour' : 'Background colour'}
          onChange={(next, phase) => apply(next, phase)}
          onClose={() => setPickerFor(null)}
        />
      )}
    </div>
  )
}

export const panel: PanelDefinition = { id: 'color', title: 'Color', component: ColorPanel, grow: 0 }
