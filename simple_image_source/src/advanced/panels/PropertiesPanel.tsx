// src/advanced/panels/PropertiesPanel.tsx (WP6)
// The Properties panel (design 5.14) shows what the active layer is:
//   - adjustment layer: its settings (AdjustmentControls: Levels and Curves editors with the histogram of
//     the image below the layer, sliders, gradient stops ...), Auto for Levels, reset, visibility and
//     "clip to the layer below". Changes apply live; one slider drag is one history step;
//   - type layer: font, style, size, leading, tracking, colour and alignment (fonts load before the
//     layer re-renders, so the pixels match);
//   - shape layer: fill, stroke, stroke width, corner radius, arrow heads;
//   - pixel layer: position (X / Y, editable unless locked) and size of its content, quick actions;
//   - a layer mask section when the layer has one (enable, link, invert, apply, delete);
//   - the document (canvas size, resolution, layers, memory) when no layer is active and under pixel layers.
import { useEffect, useRef, useState } from 'react'
import type { ReactNode } from 'react'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  Bold,
  CornerLeftDown,
  Eye,
  EyeOff,
  Italic,
  Link2,
  RotateCcw,
  Shapes,
  SlidersHorizontal,
  SquareDot,
  Type as TypeIcon,
  Underline,
  Image as ImageIcon,
  FileImage,
} from 'lucide-react'
import type { AdjustmentSpec, Histogram, Rgba8 } from '../../imaging/types.ts'
import type { AdjustmentLayer, DocumentState, Layer, RasterLayer, ShapeLayer, ShapeSpec, TextLayer, TextStyle } from '../types.ts'
import { adjustmentLabel, defaultAdjustment } from '../../imaging/adjustments.ts'
import { autoTone } from '../../imaging/autoEnhance.ts'
import { blendModeLabel } from '../../imaging/blend.ts'
import { compositeRect } from '../composite.ts'
import { formatBytes } from '../memory.ts'
import { availableFontFamilies } from '../fonts.ts'
import { cssFont } from '../../shared/vector.ts'
import { isBackgroundLayer, layerPixelBounds, setInteractive } from '../tools/shared.ts'
import type { PanelContext, PanelDefinition, PanelProps } from './PanelHost.tsx'
import { useDocumentSelector, usePixelVersion } from './PanelHost.tsx'
import type { ChangePhase } from './AdjustmentControls.tsx'
import { AdjustmentControls, CheckField, NumberInput, SelectField, histogramBelow, proxyLevel } from './AdjustmentControls.tsx'
import { ColorButton } from './ColorPicker.tsx'

// ---------------------------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------------------------

function report(context: PanelContext, error: unknown, fallback: string): void {
  context.host.notify(error instanceof Error && error.message ? error.message : fallback, 'error')
}

function useInteractive(context: PanelContext): (active: boolean) => void {
  const activeRef = useRef(false)
  useEffect(() => () => {
    if (activeRef.current) setInteractive(context.commands, false)
  }, [context])
  return (active: boolean) => {
    if (activeRef.current === active) return
    activeRef.current = active
    setInteractive(context.commands, active)
  }
}

/** A stable small id per layer object (dependency keys for "the layers below changed"). */
const objectIds = new WeakMap<object, number>()
let objectCounter = 0
function objectId(value: object): number {
  let id = objectIds.get(value)
  if (id === undefined) {
    objectCounter += 1
    id = objectCounter
    objectIds.set(value, id)
  }
  return id
}

function Section({ title, icon, actions, children }: { readonly title: string; readonly icon?: ReactNode; readonly actions?: ReactNode; readonly children: ReactNode }) {
  return (
    <section className="ae-prop-section">
      <div className="ae-prop-head">
        {icon}
        <span className="ae-prop-name">{title}</span>
        {actions && <span className="ae-prop-actions">{actions}</span>}
      </div>
      <div className="ae-prop-body">{children}</div>
    </section>
  )
}

function IconToggle({ label, pressed, disabled, onClick, children }: { readonly label: string; readonly pressed?: boolean; readonly disabled?: boolean; readonly onClick: () => void; readonly children: ReactNode }) {
  return (
    <button
      type="button"
      className={`ae-icon-mini ${pressed ? 'is-on' : ''}`}
      title={label}
      aria-label={label}
      aria-pressed={pressed}
      disabled={disabled}
      onMouseDown={(event) => event.preventDefault()}
      onClick={onClick}
    >{children}</button>
  )
}

// ---------------------------------------------------------------------------------------------
// Adjustment layers
// ---------------------------------------------------------------------------------------------

const HISTOGRAM_TYPES = new Set<AdjustmentSpec['type']>(['levels', 'curves', 'threshold'])

function AdjustmentProperties({ context, layer, state, suspended }: { readonly context: PanelContext; readonly layer: AdjustmentLayer; readonly state: DocumentState; readonly suspended: boolean }) {
  const { store } = context
  const interactive = useInteractive(context)
  const version = usePixelVersion(store, 400)
  const [histogram, setHistogram] = useState<Histogram | null>(null)
  const index = state.layers.findIndex((entry) => entry.id === layer.id)
  const belowKey = state.layers.slice(0, Math.max(0, index)).map(objectId).join(',')
  const wantsHistogram = HISTOGRAM_TYPES.has(layer.adjustment.type)

  useEffect(() => {
    if (!wantsHistogram) return
    // A moment after the image below changes (not during every slider step of this layer).
    const timer = setTimeout(() => setHistogram(histogramBelow(store.getState(), layer.id)), 80)
    return () => clearTimeout(timer)
  }, [belowKey, layer.id, store, version, wantsHistogram, state.width, state.height])

  const change = (next: AdjustmentSpec, _phase: ChangePhase) => {
    try {
      store.transact(`Modify ${adjustmentLabel(next.type)} Layer`, 'adjustment', (tx) => tx.updateLayer(layer.id, { adjustment: next }), { coalesceKey: `adjustment:${layer.id}` })
    } catch (error) {
      report(context, error, 'The adjustment could not be changed.')
    }
  }

  const auto = () => {
    try {
      const current = store.getState()
      const level = proxyLevel(current.width, current.height)
      const width = Math.max(1, Math.ceil(current.width / 2 ** level))
      const height = Math.max(1, Math.ceil(current.height / 2 ** level))
      const below = compositeRect(current, { x: 0, y: 0, width, height }, { level, belowLayerId: layer.id })
      change(autoTone(below), 'commit')
    } catch (error) {
      report(context, error, 'Auto could not analyse the image.')
    }
  }

  const reset = () => change(defaultAdjustment(layer.adjustment.type), 'commit')
  const type = layer.adjustment.type
  return (
    <Section
      title={adjustmentLabel(type)}
      icon={<SlidersHorizontal className="ae-prop-icon" aria-hidden="true" />}
      actions={(
        <>
          {type === 'levels' && <button type="button" className="ae-text-button" title="Stretch each channel to the full range (Auto Tone)" disabled={suspended} onClick={auto}>Auto</button>}
          <IconToggle label="Reset to the default settings" disabled={suspended} onClick={reset}><RotateCcw aria-hidden="true" /></IconToggle>
          <IconToggle label={layer.visible ? 'Hide this adjustment' : 'Show this adjustment'} pressed={!layer.visible} disabled={suspended} onClick={() => {
            try {
              store.transact(layer.visible ? 'Hide Layer' : 'Show Layer', 'layer', (tx) => tx.updateLayer(layer.id, { visible: !layer.visible }))
            } catch (error) {
              report(context, error, 'That could not be changed.')
            }
          }}>{layer.visible ? <Eye aria-hidden="true" /> : <EyeOff aria-hidden="true" />}</IconToggle>
          <IconToggle label={layer.clipped ? 'Release from the layer below (affects all layers below again)' : 'Clip to the layer below (affects only that layer)'} pressed={layer.clipped} disabled={suspended || !context.isEnabled('layer.toggle-clipping')} onClick={() => context.run('layer.toggle-clipping')}><CornerLeftDown aria-hidden="true" /></IconToggle>
        </>
      )}
    >
      <AdjustmentControls spec={layer.adjustment} histogram={histogram} disabled={suspended} onInteractive={interactive} onChange={change} />
      {!layer.visible && <p className="ae-note">This adjustment layer is hidden, so its settings show no effect.</p>}
    </Section>
  )
}

// ---------------------------------------------------------------------------------------------
// Type layers
// ---------------------------------------------------------------------------------------------

/** True when the font can be drawn right now (local fonts always can). */
function fontReady(style: TextStyle): boolean {
  const fonts = (typeof document !== 'undefined' ? document.fonts : null) as FontFaceSet | null
  if (!fonts || typeof fonts.check !== 'function') return true
  try {
    return fonts.check(cssFont(style))
  } catch {
    return true
  }
}

function loadFont(style: TextStyle): Promise<void> {
  const fonts = (typeof document !== 'undefined' ? document.fonts : null) as FontFaceSet | null
  if (!fonts || typeof fonts.load !== 'function') return Promise.resolve()
  return Promise.race([
    fonts.load(cssFont(style)).then(() => undefined, () => undefined),
    new Promise<void>((resolve) => setTimeout(resolve, 1500)),
  ])
}

function TextProperties({ context, layer, suspended }: { readonly context: PanelContext; readonly layer: TextLayer; readonly suspended: boolean }) {
  const { store } = context
  const sequence = useRef(0)
  const session = context.commands.services?.activeTool?.()
  const editing = Boolean(session && session.id === 'text' && session.hasSession())
  const style = layer.text.style
  const families = availableFontFamilies()
  const familyOptions = (families.includes(style.fontFamily) ? families : [style.fontFamily, ...families]).map((family) => [family, family] as const)
  const disabled = suspended || editing

  const update = (patch: Partial<TextStyle>) => {
    const ticket = sequence.current + 1
    sequence.current = ticket
    const next = { ...style, ...patch }
    // A font that is not drawable yet loads first, so the re-rendered pixels use it (otherwise apply now).
    const needsFont = (patch.fontFamily !== undefined || patch.fontWeight !== undefined || patch.italic !== undefined) && !fontReady(next)
    const apply = () => {
      // A newer change (or another layer) supersedes this one.
      if (ticket !== sequence.current) return
      const current = store.getState().layers.find((entry) => entry.id === layer.id)
      if (!current || current.kind !== 'text') return
      try {
        store.transact('Character Style', 'text', (tx) => tx.updateLayer(layer.id, { text: { ...current.text, style: { ...current.text.style, ...patch } } }), { coalesceKey: `text-style:${layer.id}` })
      } catch (error) {
        report(context, error, 'The text could not be changed.')
      }
    }
    if (needsFont) void loadFont(next).then(apply)
    else apply()
  }

  return (
    <Section title={layer.name} icon={<TypeIcon className="ae-prop-icon" aria-hidden="true" />}>
      {editing && <p className="ae-note">Finish typing first (Ctrl+Enter); the Type tool's options bar styles the text being edited.</p>}
      <SelectField label="Font" value={style.fontFamily} options={familyOptions} disabled={disabled} onChange={(fontFamily) => update({ fontFamily })} />
      <div className="ae-prop-grid">
        <label className="ae-mini-field" title="Font size">
          <span>Size</span>
          <NumberInput label="Font size in pixels" value={style.fontSize} min={1} max={5000} step={1} digits={1} disabled={disabled} onChange={(fontSize) => update({ fontSize })} />
          <i>px</i>
        </label>
        <label className="ae-mini-field" title="Leading (line height, a multiple of the size)">
          <span>Leading</span>
          <NumberInput label="Line height" value={style.lineHeight} min={0.5} max={5} step={0.05} digits={2} disabled={disabled} onChange={(lineHeight) => update({ lineHeight })} />
          <i>×</i>
        </label>
        <label className="ae-mini-field" title="Tracking (letter spacing)">
          <span>Tracking</span>
          <NumberInput label="Letter spacing in pixels" value={style.letterSpacing} min={-200} max={1000} step={0.5} digits={1} disabled={disabled} onChange={(letterSpacing) => update({ letterSpacing })} />
          <i>px</i>
        </label>
        <div className="ae-mini-field">
          <span>Colour</span>
          <ColorButton label="Text colour" value={style.color} disabled={disabled} onChange={(color) => update({ color })} />
        </div>
      </div>
      <div className="ae-inline-row">
        <div className="ae-segmented" role="group" aria-label="Style">
          <button type="button" title="Bold" aria-label="Bold" aria-pressed={style.fontWeight === 700} className={style.fontWeight === 700 ? 'is-active' : ''} disabled={disabled} onClick={() => update({ fontWeight: style.fontWeight === 700 ? 400 : 700 })}><Bold aria-hidden="true" /></button>
          <button type="button" title="Italic" aria-label="Italic" aria-pressed={style.italic} className={style.italic ? 'is-active' : ''} disabled={disabled} onClick={() => update({ italic: !style.italic })}><Italic aria-hidden="true" /></button>
          <button type="button" title="Underline" aria-label="Underline" aria-pressed={style.underline} className={style.underline ? 'is-active' : ''} disabled={disabled} onClick={() => update({ underline: !style.underline })}><Underline aria-hidden="true" /></button>
        </div>
        <div className="ae-segmented" role="group" aria-label="Alignment">
          {([['left', AlignLeft], ['center', AlignCenter], ['right', AlignRight]] as const).map(([align, Icon]) => (
            <button key={align} type="button" title={`Align ${align}`} aria-label={`Align ${align}`} aria-pressed={style.align === align} className={style.align === align ? 'is-active' : ''} disabled={disabled} onClick={() => update({ align })}><Icon aria-hidden="true" /></button>
          ))}
        </div>
      </div>
      <p className="ae-note">{layer.text.boxWidth === null ? 'Point text' : `Paragraph text, ${Math.round(layer.text.boxWidth)} px wide`}. Edit the words with the Type tool (T).</p>
    </Section>
  )
}

// ---------------------------------------------------------------------------------------------
// Shape layers
// ---------------------------------------------------------------------------------------------

const SHAPE_NAMES: Readonly<Record<ShapeSpec['kind'], string>> = Object.freeze({ rectangle: 'Rectangle', ellipse: 'Ellipse', line: 'Line', arrow: 'Arrow' })

function ShapeProperties({ context, layer, suspended }: { readonly context: PanelContext; readonly layer: ShapeLayer; readonly suspended: boolean }) {
  const { store } = context
  const shape = layer.shape
  const lastFill = useRef<Rgba8>(shape.fill ?? { r: 64, g: 128, b: 230, a: 255 })
  const lastStroke = useRef<Rgba8>(shape.stroke ?? { r: 0, g: 0, b: 0, a: 255 })
  if (shape.fill) lastFill.current = shape.fill
  if (shape.stroke) lastStroke.current = shape.stroke
  const update = (patch: Partial<ShapeSpec>) => {
    const current = store.getState().layers.find((entry) => entry.id === layer.id)
    if (!current || current.kind !== 'shape') return
    try {
      store.transact('Edit Shape', 'shape', (tx) => tx.updateLayer(layer.id, { shape: { ...current.shape, ...patch } }), { coalesceKey: `shape:${layer.id}` })
    } catch (error) {
      report(context, error, 'The shape could not be changed.')
    }
  }
  const lineLike = shape.kind === 'line' || shape.kind === 'arrow'
  const width = Math.round(Math.abs(shape.x2 - shape.x1))
  const height = Math.round(Math.abs(shape.y2 - shape.y1))
  return (
    <Section title={layer.name} icon={<Shapes className="ae-prop-icon" aria-hidden="true" />}>
      <p className="ae-note">{SHAPE_NAMES[shape.kind]} · {lineLike ? `${Math.round(Math.hypot(shape.x2 - shape.x1, shape.y2 - shape.y1))} px long` : `${width} × ${height} px`}</p>
      {!lineLike && (
        <div className="ae-inline-row">
          <CheckField label="Fill" checked={shape.fill !== null} disabled={suspended} onChange={(on) => update({ fill: on ? lastFill.current : null })} />
          <ColorButton label="Fill colour" value={shape.fill ?? lastFill.current} disabled={suspended || !shape.fill} onChange={(color) => update({ fill: { ...color, a: shape.fill?.a ?? 255 } })} />
        </div>
      )}
      <div className="ae-inline-row">
        <CheckField label="Stroke" checked={shape.stroke !== null} disabled={suspended || (lineLike && shape.fill === null && shape.stroke !== null)} onChange={(on) => update({ stroke: on ? lastStroke.current : null })} />
        <ColorButton label="Stroke colour" value={shape.stroke ?? lastStroke.current} disabled={suspended || !shape.stroke} onChange={(color) => update({ stroke: { ...color, a: shape.stroke?.a ?? 255 } })} />
        <label className="ae-mini-field" title="Stroke width">
          <span>Width</span>
          <NumberInput label="Stroke width in pixels" value={shape.strokeWidth} min={0} max={1000} step={1} digits={1} disabled={suspended} onChange={(strokeWidth) => update({ strokeWidth })} />
          <i>px</i>
        </label>
      </div>
      {shape.kind === 'rectangle' && (
        <label className="ae-mini-field" title="Corner radius">
          <span>Radius</span>
          <NumberInput label="Corner radius in pixels" value={shape.cornerRadius} min={0} max={10000} disabled={suspended} onChange={(cornerRadius) => update({ cornerRadius })} />
          <i>px</i>
        </label>
      )}
      {lineLike && (
        <SelectField label="Arrow heads" inline value={shape.arrowHeads} options={[['none', 'None'], ['end', 'At the end'], ['both', 'Both ends']]} disabled={suspended} onChange={(arrowHeads) => update({ arrowHeads })} />
      )}
    </Section>
  )
}

// ---------------------------------------------------------------------------------------------
// Pixel layers and the document
// ---------------------------------------------------------------------------------------------

function PixelProperties({ context, layer, suspended }: { readonly context: PanelContext; readonly layer: RasterLayer; readonly suspended: boolean }) {
  const { store } = context
  const bounds = layerPixelBounds(layer)
  const background = isBackgroundLayer(layer)
  const positionLocked = background || layer.locks.position
  const moveTo = (x: number, y: number) => {
    if (!bounds) return
    const dx = Math.round(x) - bounds.x
    const dy = Math.round(y) - bounds.y
    if (!dx && !dy) return
    try {
      store.transact('Move', 'move', (tx) => tx.updateLayer(layer.id, { offset: { x: layer.offsetX + dx, y: layer.offsetY + dy } }), { coalesceKey: `move:${layer.id}` })
    } catch (error) {
      report(context, error, 'The layer could not be moved.')
    }
  }
  return (
    <Section title={layer.name} icon={<ImageIcon className="ae-prop-icon" aria-hidden="true" />}>
      <p className="ae-note">{background ? 'Background: opaque, locked in place. Make it a layer to move it or give it transparency.' : `Pixel layer · ${blendModeLabel(layer.blendMode)} · ${Math.round(layer.opacity * 100)}%`}</p>
      {bounds ? (
        <div className="ae-prop-grid">
          <label className="ae-mini-field" title={positionLocked ? 'The position is locked' : 'Left edge of the layer content'}>
            <span>X</span>
            <NumberInput label="Layer X position" value={bounds.x} min={-1_000_000} max={1_000_000} disabled={suspended || positionLocked} onChange={(x) => moveTo(x, bounds.y)} />
            <i>px</i>
          </label>
          <label className="ae-mini-field" title={positionLocked ? 'The position is locked' : 'Top edge of the layer content'}>
            <span>Y</span>
            <NumberInput label="Layer Y position" value={bounds.y} min={-1_000_000} max={1_000_000} disabled={suspended || positionLocked} onChange={(y) => moveTo(bounds.x, y)} />
            <i>px</i>
          </label>
          <div className="ae-mini-field"><span>W</span><output>{bounds.width.toLocaleString()}</output><i>px</i></div>
          <div className="ae-mini-field"><span>H</span><output>{bounds.height.toLocaleString()}</output><i>px</i></div>
        </div>
      ) : <p className="ae-note">This layer is empty.</p>}
      <div className="ae-prop-buttons">
        {background
          ? <button type="button" className="ae-text-button" disabled={suspended} onClick={() => context.run('layer.from-background')}>Make it a layer</button>
          : <button type="button" className="ae-text-button" disabled={suspended || !context.isEnabled('edit.free-transform')} title={`Free Transform (${context.shortcut('edit.free-transform')})`} onClick={() => context.run('edit.free-transform')}>Transform…</button>}
        <button type="button" className="ae-text-button" disabled={suspended || !bounds} title="Select the pixels of this layer" onClick={() => context.run('select.load-layer-alpha')}>Select pixels</button>
      </div>
    </Section>
  )
}

function DocumentProperties({ context, state, suspended, compact }: { readonly context: PanelContext; readonly state: DocumentState; readonly suspended: boolean; readonly compact?: boolean }) {
  const [memory, setMemory] = useState(() => context.store.memoryUsage().total)
  useEffect(() => {
    const timer = setInterval(() => setMemory(context.store.memoryUsage().total), 2000)
    return () => clearInterval(timer)
  }, [context])
  const megapixels = (state.width * state.height) / 1_000_000
  const rows: [string, string][] = [
    ['Canvas', `${state.width.toLocaleString()} × ${state.height.toLocaleString()} px`],
    ['Resolution', `${state.ppi} ppi · ${(state.width / state.ppi).toFixed(2)} × ${(state.height / state.ppi).toFixed(2)} in`],
  ]
  if (!compact) {
    rows.push(['Size', `${megapixels.toFixed(megapixels < 10 ? 2 : 1)} MP · RGB, 8 bits`], ['Layers', String(state.layers.length)], ['Memory', formatBytes(memory)])
  }
  return (
    <Section title="Document" icon={<FileImage className="ae-prop-icon" aria-hidden="true" />}>
      <dl className="ae-prop-list">
        {rows.map(([term, value]) => <div key={term}><dt>{term}</dt><dd title={value}>{value}</dd></div>)}
      </dl>
      <div className="ae-prop-buttons">
        <button type="button" className="ae-text-button" disabled={suspended || !context.isEnabled('image.size')} title={context.shortcut('image.size')} onClick={() => context.run('image.size')}>Image Size…</button>
        <button type="button" className="ae-text-button" disabled={suspended || !context.isEnabled('image.canvas-size')} title={context.shortcut('image.canvas-size')} onClick={() => context.run('image.canvas-size')}>Canvas Size…</button>
      </div>
    </Section>
  )
}

// ---------------------------------------------------------------------------------------------
// Layer mask
// ---------------------------------------------------------------------------------------------

function MaskProperties({ context, layer, state, suspended }: { readonly context: PanelContext; readonly layer: Layer; readonly state: DocumentState; readonly suspended: boolean }) {
  const { store } = context
  const mask = layer.mask
  if (!mask) return null
  const targeted = state.activeLayerId === layer.id && state.editTarget === 'mask'
  const set = (label: string, patch: { maskEnabled?: boolean; maskLinked?: boolean }) => {
    try {
      store.transact(label, 'layer', (tx) => tx.updateLayer(layer.id, patch))
    } catch (error) {
      report(context, error, 'The mask could not be changed.')
    }
  }
  const invertMask = () => {
    if (!targeted) {
      try {
        store.transact('Select Layer Mask', 'layer', (tx) => tx.setActiveLayer(layer.id, 'mask'), { affectsOutput: false })
      } catch (error) {
        report(context, error, 'The mask could not be selected.')
        return
      }
    }
    context.run('adjust.invert')
  }
  return (
    <Section
      title="Layer Mask"
      icon={<SquareDot className="ae-prop-icon" aria-hidden="true" />}
      actions={(
        <>
          <IconToggle label={mask.linked ? 'Unlink the mask (move the layer without it)' : 'Link the mask to the layer'} pressed={mask.linked} disabled={suspended} onClick={() => set(mask.linked ? 'Unlink Layer Mask' : 'Link Layer Mask', { maskLinked: !mask.linked })}><Link2 aria-hidden="true" /></IconToggle>
        </>
      )}
    >
      <p className="ae-note">{targeted ? 'Painting and adjustments now change the mask: black hides, white shows.' : 'Click the mask thumbnail in Layers to paint on it.'}</p>
      <div className="ae-prop-buttons">
        <button type="button" className="ae-text-button" disabled={suspended} onClick={() => set(mask.enabled ? 'Disable Layer Mask' : 'Enable Layer Mask', { maskEnabled: !mask.enabled })}>{mask.enabled ? 'Disable' : 'Enable'}</button>
        <button type="button" className="ae-text-button" disabled={suspended} title="Swap what the mask shows and hides" onClick={invertMask}>Invert</button>
        {layer.kind === 'raster' && <button type="button" className="ae-text-button" disabled={suspended || !context.isEnabled('layer.apply-mask')} onClick={() => context.run('layer.apply-mask')}>Apply</button>}
        <button type="button" className="ae-text-button" disabled={suspended} onClick={() => context.run('layer.delete-mask')}>Delete</button>
      </div>
    </Section>
  )
}

// ---------------------------------------------------------------------------------------------
// The panel
// ---------------------------------------------------------------------------------------------

function PropertiesPanel({ context, suspended }: PanelProps): ReactNode {
  const state = useDocumentSelector(context.store, (current) => current, (a, b) => a.layers === b.layers && a.activeLayerId === b.activeLayerId
    && a.editTarget === b.editTarget && a.width === b.width && a.height === b.height && a.ppi === b.ppi)
  const active = state.layers.find((layer) => layer.id === state.activeLayerId) ?? null
  let body: ReactNode
  if (!active) {
    body = <DocumentProperties context={context} state={state} suspended={suspended} />
  } else if (active.kind === 'adjustment') {
    body = <AdjustmentProperties key={active.id} context={context} layer={active} state={state} suspended={suspended} />
  } else if (active.kind === 'text') {
    body = <TextProperties key={active.id} context={context} layer={active} suspended={suspended} />
  } else if (active.kind === 'shape') {
    body = <ShapeProperties key={active.id} context={context} layer={active} suspended={suspended} />
  } else {
    body = (
      <>
        <PixelProperties key={active.id} context={context} layer={active} suspended={suspended} />
        <DocumentProperties context={context} state={state} suspended={suspended} compact />
      </>
    )
  }
  const maskFirst = Boolean(active?.mask && state.editTarget === 'mask')
  return (
    <div className="ae-properties-panel">
      {maskFirst && active && <MaskProperties context={context} layer={active} state={state} suspended={suspended} />}
      {body}
      {!maskFirst && active?.mask && <MaskProperties context={context} layer={active} state={state} suspended={suspended} />}
    </div>
  )
}

export const panel: PanelDefinition = { id: 'properties', title: 'Properties', component: PropertiesPanel, grow: 2.4 }
