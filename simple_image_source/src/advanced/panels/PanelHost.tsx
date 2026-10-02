// src/advanced/panels/PanelHost.tsx (WP6)
// The right-hand dock of the Advanced editor and its dialogs (design 3.2, 5.14), with a stable, typed API
// that the full panels and dialogs plug into without editing this file:
//   - A panel module is src/advanced/panels/<Name>Panel.tsx exporting `panel: PanelDefinition`; a dialog
//     module is src/advanced/dialogs/<Name>Dialog.tsx exporting `dialog: DialogDefinition`. Both are found at
//     build time (import.meta.glob) and replace the built-in placeholder for their id / requests.
//     Such modules may import types and the hook functions (useDocumentSelector, useEditorState,
//     useHistoryState) from this file, but must not read its values at module top level (import cycle).
//   - PanelContext is created once per editor session and never changes identity: the document store,
//     editor store, compositor, imaging client, host, view, the command context (for the operations
//     commands.ts exports) and run / isEnabled / label / shortcut for commands.
//   - Dialogs: commands ask for parameters through DialogController.open(request); a registered dialog
//     shows it, otherwise the built-in parameter dialog does when it can (numeric settings with a live
//     preview for adjustments). Dialog keys (Enter, Escape) are handled and stopped inside the dialog.
//   - Built-in placeholders keep the editor usable until the full panels land: Layers (visibility, blend
//     mode, opacity, thumbnails, rename, new / mask / adjustment / delete), Properties (document, layer and
//     adjustment settings), History (click to jump; future states dimmed) and Color (colours, swatches).
// Panels collapse to an overlay below 1080 px (advanced.css).
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import type { ComponentType, KeyboardEvent as ReactKeyboardEvent, ReactNode } from 'react'
import {
  ChevronDown,
  ChevronRight,
  Eye,
  EyeOff,
  Layers as LayersIcon,
  Link2,
  Lock,
  Plus,
  SlidersHorizontal,
  SquareDashed,
  Trash2,
  Type as TypeIcon,
  Shapes,
} from 'lucide-react'
import type {
  AdjustmentSpec,
  AdjustmentType,
  BlendMode,
  FilterSpec,
  FilterType,
  ImagingClient,
  ResampleMethod,
  Rgb8,
} from '../../imaging/types.ts'
import type {
  AdvancedHost,
  Compositor,
  DocumentState,
  DocumentStore,
  EditorState,
  EditorStore,
  HistoryState,
  Layer,
  LayerId,
  ViewController,
} from '../types.ts'
import { BLEND_MODE_MENU, LIMITS } from '../types.ts'
import type { CanvasAnchor, CommandContext, DialogRequest, EditorCommandId } from '../commands.ts'
import {
  CANVAS_ANCHORS,
  applyAdjustmentToLayer,
  applyFilterToLayer,
  modifySelection,
  resizeCanvas,
  resizeImage,
  rotateCanvas,
} from '../commands.ts'
import { ADJUSTMENT_TYPES, adjustmentLabel, defaultAdjustment } from '../../imaging/adjustments.ts'
import { defaultFilter, filterLabel } from '../../imaging/filters.ts'
import { blendModeLabel } from '../../imaging/blend.ts'
import { parseHex, toHex } from '../../imaging/color.ts'
import { compositeRect } from '../composite.ts'
import { maxPyramidLevel } from '../pyramid.ts'
import { formatBytes } from '../memory.ts'
import { MAX_SWATCHES, rememberColor, sameColor } from '../editorState.ts'

// ---------------------------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------------------------

export type PanelId = 'layers' | 'properties' | 'history' | 'color'

export const PANEL_ORDER: readonly PanelId[] = Object.freeze(['layers', 'properties', 'history', 'color'] as const)

/** Everything a panel or dialog needs. Created once per editor session; its identity never changes. */
export interface PanelContext {
  readonly store: DocumentStore
  readonly editor: EditorStore
  readonly compositor: Compositor
  readonly imaging: ImagingClient
  readonly host: AdvancedHost
  readonly view: ViewController
  /** Pass to the operations commands.ts exports (applyAdjustmentToLayer, resizeImage, ...). */
  readonly commands: CommandContext
  /** Runs a command exactly like its menu item and shortcut. */
  run(command: EditorCommandId): void
  isEnabled(command: EditorCommandId): boolean
  label(command: EditorCommandId): string
  /** "Shift+Ctrl+N" style text, or ''. */
  shortcut(command: EditorCommandId): string
  /** Opens a parameter dialog; false when nothing can show the request. */
  openDialog(request: DialogRequest): boolean
  /** Returns keyboard focus to the canvas (after a click in a panel, so shortcuts keep working). */
  focusCanvas(): void
}

export interface PanelProps {
  readonly context: PanelContext
  /** A host modal owns input: render inert. */
  readonly suspended: boolean
}

export interface PanelDefinition {
  readonly id: PanelId
  readonly title: string
  readonly component: ComponentType<PanelProps>
  /** Share of the dock height while expanded (default 1). */
  readonly grow?: number
}

export interface DialogProps {
  readonly context: PanelContext
  readonly request: DialogRequest
  /** Closes the dialog; call it after applying or when the user cancels. */
  close(): void
}

export interface DialogDefinition {
  /** True for the requests this dialog shows. */
  handles(request: DialogRequest): boolean
  readonly component: ComponentType<DialogProps>
}

// ---------------------------------------------------------------------------------------------
// Hooks for panels (function declarations: safe to import from panel modules despite the glob cycle)
// ---------------------------------------------------------------------------------------------

/** Re-renders only when the selected part of the document state changes (pixel edits keep `layers` identical). */
export function useDocumentSelector<T>(store: DocumentStore, select: (state: DocumentState) => T, isEqual: (a: T, b: T) => boolean = Object.is): T {
  const cache = useRef<{ state: DocumentState; value: T } | null>(null)
  const selectRef = useRef(select)
  selectRef.current = select
  const equalRef = useRef(isEqual)
  equalRef.current = isEqual
  const getSnapshot = useCallback(() => {
    const state = store.getState()
    const previous = cache.current
    if (previous && previous.state === state) return previous.value
    const value = selectRef.current(state)
    if (previous && equalRef.current(previous.value, value)) {
      cache.current = { state, value: previous.value }
      return previous.value
    }
    cache.current = { state, value }
    return value
  }, [store])
  return useSyncExternalStore(useCallback((listener: () => void) => store.subscribe(() => listener()), [store]), getSnapshot)
}

export function useEditorState(editor: EditorStore): EditorState {
  return useSyncExternalStore(useCallback((listener: () => void) => editor.subscribe(listener), [editor]), () => editor.getState())
}

export function useHistoryState(store: DocumentStore): HistoryState {
  return useSyncExternalStore(useCallback((listener: () => void) => store.subscribe(() => listener()), [store]), () => store.history.getState())
}

/** The document's pixelVersion, updated at most every `ms` (thumbnails stay cheap during brush strokes). */
export function usePixelVersion(store: DocumentStore, ms = 300): number {
  const [version, setVersion] = useState(() => store.getState().pixelVersion)
  useEffect(() => {
    let timer: ReturnType<typeof setTimeout> | null = null
    let last = 0
    const update = () => {
      timer = null
      last = Date.now()
      setVersion(store.getState().pixelVersion)
    }
    const unsubscribe = store.subscribe(() => {
      if (timer !== null) return
      timer = setTimeout(update, Math.max(0, last + ms - Date.now()))
    })
    return () => {
      unsubscribe()
      if (timer !== null) clearTimeout(timer)
    }
  }, [ms, store])
  return version
}

/** Sets the `inert` property (React 18 has no prop for it): a host modal makes the element unreachable. */
export function useInert<T extends HTMLElement>(inert: boolean) {
  const ref = useRef<T>(null)
  useLayoutEffect(() => {
    const element = ref.current as (T & { inert?: boolean }) | null
    if (element) element.inert = inert
  })
  return ref
}

// ---------------------------------------------------------------------------------------------
// Registry (modules added by later work packages)
// ---------------------------------------------------------------------------------------------

interface PanelModule { readonly panel?: PanelDefinition }
interface DialogModule { readonly dialog?: DialogDefinition }

const PANEL_MODULES = import.meta.glob<PanelModule>('./*Panel.tsx', { eager: true })
const DIALOG_MODULES = import.meta.glob<DialogModule>('../dialogs/*Dialog.tsx', { eager: true })

function registeredPanels(): Map<PanelId, PanelDefinition> {
  const out = new Map<PanelId, PanelDefinition>()
  for (const module of Object.values(PANEL_MODULES)) {
    const definition = module?.panel
    // memo() / forwardRef() components are objects, plain components functions.
    if (definition && PANEL_ORDER.includes(definition.id) && definition.component) out.set(definition.id, definition)
  }
  return out
}

function registeredDialogs(): DialogDefinition[] {
  const out: DialogDefinition[] = []
  for (const module of Object.values(DIALOG_MODULES)) {
    const definition = module?.dialog
    if (definition && typeof definition.handles === 'function' && definition.component) out.push(definition)
  }
  return out
}

const PANELS = registeredPanels()
const DIALOGS = registeredDialogs()

// ---------------------------------------------------------------------------------------------
// Dialog controller
// ---------------------------------------------------------------------------------------------

export interface DialogController {
  open(request: DialogRequest): boolean
  canOpen(request: DialogRequest): boolean
  close(): void
  current(): DialogRequest | null
  subscribe(listener: () => void): () => void
}

export function createDialogController(): DialogController {
  let current: DialogRequest | null = null
  const listeners = new Set<() => void>()
  const emit = () => {
    for (const listener of [...listeners]) listener()
  }
  const controller: DialogController = {
    canOpen: (request) => DIALOGS.some((definition) => definition.handles(request)) || genericFields(request) !== null,
    open(request: DialogRequest): boolean {
      if (!controller.canOpen(request)) return false
      current = request
      emit()
      return true
    },
    close(): void {
      if (!current) return
      current = null
      emit()
    },
    current: () => current,
    subscribe(listener: () => void): () => void {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
  }
  return controller
}

// ---------------------------------------------------------------------------------------------
// Settings forms (built-in dialog and the Properties placeholder)
// ---------------------------------------------------------------------------------------------

type Field =
  | { readonly kind: 'number'; readonly path: string; readonly label: string; readonly min: number; readonly max: number; readonly step?: number; readonly unit?: string; readonly slider?: boolean }
  | { readonly kind: 'boolean'; readonly path: string; readonly label: string }
  | { readonly kind: 'select'; readonly path: string; readonly label: string; readonly options: readonly (readonly [string, string])[] }
  | { readonly kind: 'color'; readonly path: string; readonly label: string }

const n = (path: string, label: string, min: number, max: number, extra: { step?: number; unit?: string; slider?: boolean } = {}): Field => ({ kind: 'number', path, label, min, max, slider: true, ...extra })
const b = (path: string, label: string): Field => ({ kind: 'boolean', path, label })
const s = (path: string, label: string, options: readonly (readonly [string, string])[]): Field => ({ kind: 'select', path, label, options })

const BALANCE = (tone: 'shadows' | 'midtones' | 'highlights', title: string): Field[] => [
  n(`${tone}.cyanRed`, `${title}: Cyan – Red`, -100, 100),
  n(`${tone}.magentaGreen`, `${title}: Magenta – Green`, -100, 100),
  n(`${tone}.yellowBlue`, `${title}: Yellow – Blue`, -100, 100),
]

const ADJUSTMENT_FIELDS: Readonly<Partial<Record<AdjustmentType, readonly Field[]>>> = Object.freeze({
  'brightness-contrast': [n('brightness', 'Brightness', -150, 150), n('contrast', 'Contrast', -50, 100), b('legacy', 'Use Legacy')],
  levels: [
    n('rgb.inBlack', 'Input black', 0, 253), n('rgb.gamma', 'Midtones (gamma)', 0.1, 9.99, { step: 0.01 }), n('rgb.inWhite', 'Input white', 2, 255),
    n('rgb.outBlack', 'Output black', 0, 255), n('rgb.outWhite', 'Output white', 0, 255),
  ],
  exposure: [n('exposure', 'Exposure', -20, 20, { step: 0.01 }), n('offset', 'Offset', -0.5, 0.5, { step: 0.001 }), n('gamma', 'Gamma correction', 0.01, 9.99, { step: 0.01 })],
  vibrance: [n('vibrance', 'Vibrance', -100, 100), n('saturation', 'Saturation', -100, 100)],
  'hue-saturation': [n('master.hue', 'Hue', -180, 180, { unit: '°' }), n('master.saturation', 'Saturation', -100, 100), n('master.lightness', 'Lightness', -100, 100), b('colorize', 'Colorize')],
  'color-balance': [...BALANCE('midtones', 'Midtones'), ...BALANCE('shadows', 'Shadows'), ...BALANCE('highlights', 'Highlights'), b('preserveLuminosity', 'Preserve Luminosity')],
  'black-white': [
    n('reds', 'Reds', -200, 300, { unit: '%' }), n('yellows', 'Yellows', -200, 300, { unit: '%' }), n('greens', 'Greens', -200, 300, { unit: '%' }),
    n('cyans', 'Cyans', -200, 300, { unit: '%' }), n('blues', 'Blues', -200, 300, { unit: '%' }), n('magentas', 'Magentas', -200, 300, { unit: '%' }),
  ],
  'photo-filter': [{ kind: 'color', path: 'color', label: 'Filter colour' }, n('density', 'Density', 0, 100, { unit: '%' }), b('preserveLuminosity', 'Preserve Luminosity')],
  posterize: [n('levels', 'Levels', 2, 255)],
  threshold: [n('level', 'Threshold level', 1, 255)],
  'gradient-map': [b('reverse', 'Reverse'), b('dither', 'Dither')],
})

const FILTER_FIELDS: Readonly<Record<FilterType, readonly Field[]>> = Object.freeze({
  'gaussian-blur': [n('radius', 'Radius', 0.1, 250, { step: 0.1, unit: 'px' })],
  'motion-blur': [n('angle', 'Angle', -90, 90, { unit: '°' }), n('distance', 'Distance', 1, 999, { unit: 'px' })],
  'unsharp-mask': [n('amount', 'Amount', 1, 500, { unit: '%' }), n('radius', 'Radius', 0.1, 250, { step: 0.1, unit: 'px' }), n('threshold', 'Threshold', 0, 255, { unit: 'levels' })],
  sharpen: [s('strength', 'Strength', [['normal', 'Sharpen'], ['more', 'Sharpen More']])],
  'add-noise': [n('amount', 'Amount', 0.1, 400, { step: 0.1, unit: '%' }), s('distribution', 'Distribution', [['uniform', 'Uniform'], ['gaussian', 'Gaussian']]), b('monochromatic', 'Monochromatic')],
  median: [n('radius', 'Radius', 1, 100, { unit: 'px' })],
  'reduce-noise': [n('strength', 'Strength', 0, 10), n('preserveDetails', 'Preserve Details', 0, 100, { unit: '%' })],
  pixelate: [n('cellSize', 'Cell Size', 2, 200, { unit: 'square' })],
  emboss: [n('angle', 'Angle', -180, 180, { unit: '°' }), n('height', 'Height', 1, 10, { unit: 'px' }), n('amount', 'Amount', 1, 500, { unit: '%' })],
  'find-edges': [],
})

const RESAMPLE_OPTIONS: readonly (readonly [ResampleMethod, string])[] = [
  ['auto', 'Automatic'], ['bicubic', 'Bicubic (smooth gradients)'], ['lanczos3', 'Lanczos (sharper)'], ['bilinear', 'Bilinear'], ['area', 'Area average (reduction)'], ['nearest', 'Nearest Neighbor (hard edges)'],
]

const ANCHOR_LABELS: Readonly<Record<CanvasAnchor, string>> = Object.freeze({
  'top-left': 'Top left', top: 'Top', 'top-right': 'Top right', left: 'Left', center: 'Center', right: 'Right',
  'bottom-left': 'Bottom left', bottom: 'Bottom', 'bottom-right': 'Bottom right',
})

function genericFields(request: DialogRequest): readonly Field[] | null {
  switch (request.kind) {
    case 'adjustment': return ADJUSTMENT_FIELDS[request.type] ?? null
    case 'filter': return FILTER_FIELDS[request.type]?.length ? FILTER_FIELDS[request.type] : null
    case 'modify-selection': return [n('amount', request.operation === 'feather' ? 'Feather Radius' : request.operation === 'expand' ? 'Expand By' : 'Contract By', request.operation === 'feather' ? 0.1 : 1, request.operation === 'feather' ? 1000 : 500, { step: request.operation === 'feather' ? 0.1 : 1, unit: 'px', slider: false })]
    case 'rotate-canvas': return [n('angle', 'Angle (clockwise)', -359.99, 359.99, { step: 0.01, unit: '°', slider: false })]
    case 'canvas-size': return [
      n('width', 'Width', 1, LIMITS.maxDimension, { unit: 'px', slider: false }),
      n('height', 'Height', 1, LIMITS.maxDimension, { unit: 'px', slider: false }),
      s('anchor', 'Anchor', CANVAS_ANCHORS.map((anchor) => [anchor, ANCHOR_LABELS[anchor]] as const)),
    ]
    case 'image-size': return [
      n('width', 'Width', 1, LIMITS.maxDimension, { unit: 'px', slider: false }),
      n('height', 'Height', 1, LIMITS.maxDimension, { unit: 'px', slider: false }),
      b('constrain', 'Constrain proportions'),
      n('ppi', 'Resolution', 1, 10000, { unit: 'ppi', slider: false }),
      s('method', 'Resample', RESAMPLE_OPTIONS),
    ]
    default: return null
  }
}

function getPath(value: unknown, path: string): unknown {
  let current: unknown = value
  for (const key of path.split('.')) {
    if (!current || typeof current !== 'object') return undefined
    current = (current as Record<string, unknown>)[key]
  }
  return current
}

function setPath<T>(value: T, path: string, next: unknown): T {
  const [head, ...rest] = path.split('.')
  const source = (value && typeof value === 'object' ? value : {}) as Record<string, unknown>
  return { ...source, [head]: rest.length ? setPath(source[head], rest.join('.'), next) : next } as T
}

function clampNumber(value: number, field: Extract<Field, { kind: 'number' }>): number {
  const clamped = Math.min(field.max, Math.max(field.min, value))
  const step = field.step ?? 1
  return step >= 1 ? Math.round(clamped) : Math.round(clamped / step) * step
}

function NumberField({ field, value, disabled, onChange, onCommit }: {
  readonly field: Extract<Field, { kind: 'number' }>
  readonly value: number
  readonly disabled?: boolean
  readonly onChange: (value: number) => void
  readonly onCommit?: () => void
}) {
  const [text, setText] = useState(String(value))
  const editing = useRef(false)
  useEffect(() => {
    if (!editing.current) setText(String(Number.isFinite(value) ? Number(value.toFixed(4)) : field.min))
  }, [field.min, value])
  const id = `ae-field-${field.path.replace(/\W/g, '-')}`
  const commitText = () => {
    editing.current = false
    const parsed = Number(text.replace(',', '.'))
    if (Number.isFinite(parsed)) onChange(clampNumber(parsed, field))
    else setText(String(value))
    onCommit?.()
  }
  return (
    <div className="ae-field">
      <label htmlFor={id}>{field.label}</label>
      <div className="ae-field-control">
        {field.slider !== false && (
          <input
            type="range"
            aria-label={field.label}
            min={field.min}
            max={field.max}
            step={field.step ?? 1}
            value={Number.isFinite(value) ? value : field.min}
            disabled={disabled}
            onChange={(event) => onChange(clampNumber(Number(event.currentTarget.value), field))}
            onPointerUp={() => onCommit?.()}
            onKeyUp={() => onCommit?.()}
          />
        )}
        <input
          id={id}
          className="ae-number"
          type="text"
          inputMode="decimal"
          value={text}
          disabled={disabled}
          onFocus={() => { editing.current = true }}
          onChange={(event) => setText(event.currentTarget.value)}
          onBlur={commitText}
          onKeyDown={(event) => {
            if (event.key === 'Enter') commitText()
            if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
              event.preventDefault()
              const step = (field.step ?? 1) * (event.shiftKey ? 10 : 1)
              const next = clampNumber(value + (event.key === 'ArrowUp' ? step : -step), field)
              setText(String(Number(next.toFixed(4))))
              onChange(next)
            }
          }}
        />
        {field.unit && <span className="ae-unit">{field.unit}</span>}
      </div>
    </div>
  )
}

function SettingsForm<T>({ fields, value, disabled, onChange, onCommit }: {
  readonly fields: readonly Field[]
  readonly value: T
  readonly disabled?: boolean
  readonly onChange: (next: T, path: string) => void
  readonly onCommit?: () => void
}) {
  return (
    <div className="ae-form">
      {fields.map((field) => {
        const current = getPath(value, field.path)
        if (field.kind === 'number') {
          return <NumberField key={field.path} field={field} value={Number(current)} disabled={disabled} onChange={(next) => onChange(setPath(value, field.path, next), field.path)} onCommit={onCommit} />
        }
        if (field.kind === 'boolean') {
          return (
            <label key={field.path} className="ae-check">
              <input type="checkbox" checked={Boolean(current)} disabled={disabled} onChange={(event) => { onChange(setPath(value, field.path, event.currentTarget.checked), field.path); onCommit?.() }} />
              <span>{field.label}</span>
            </label>
          )
        }
        if (field.kind === 'select') {
          return (
            <div key={field.path} className="ae-field">
              <label htmlFor={`ae-field-${field.path}`}>{field.label}</label>
              <select id={`ae-field-${field.path}`} value={String(current)} disabled={disabled} onChange={(event) => { onChange(setPath(value, field.path, event.currentTarget.value), field.path); onCommit?.() }}>
                {field.options.map(([option, text]) => <option key={option} value={option}>{text}</option>)}
              </select>
            </div>
          )
        }
        const color = current as Rgb8 | undefined
        return (
          <div key={field.path} className="ae-field">
            <label htmlFor={`ae-field-${field.path}`}>{field.label}</label>
            <input
              id={`ae-field-${field.path}`}
              type="color"
              value={color ? toHex(color) : '#000000'}
              disabled={disabled}
              onChange={(event) => {
                const parsed = parseHex(event.currentTarget.value)
                if (parsed) onChange(setPath(value, field.path, parsed), field.path)
              }}
              onBlur={() => onCommit?.()}
            />
          </div>
        )
      })}
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// Built-in parameter dialog
// ---------------------------------------------------------------------------------------------

function dialogTitle(request: DialogRequest): string {
  switch (request.kind) {
    case 'adjustment': return adjustmentLabel(request.type)
    case 'filter': return filterLabel(request.type)
    case 'modify-selection': return request.operation === 'feather' ? 'Feather Selection' : request.operation === 'expand' ? 'Expand Selection' : 'Contract Selection'
    case 'rotate-canvas': return 'Rotate Canvas'
    case 'canvas-size': return 'Canvas Size'
    case 'image-size': return 'Image Size'
  }
}

function initialValues(request: DialogRequest, context: PanelContext): Record<string, unknown> {
  const state = context.store.getState()
  switch (request.kind) {
    case 'adjustment': return { ...defaultAdjustment(request.type) } as unknown as Record<string, unknown>
    case 'filter': {
      const last = context.editor.getState().lastFilter
      return { ...(last && last.type === request.type ? last : defaultFilter(request.type)) } as unknown as Record<string, unknown>
    }
    case 'modify-selection': return { amount: request.operation === 'feather' ? 2 : 1 }
    case 'rotate-canvas': return { angle: 0 }
    case 'canvas-size': return { width: state.width, height: state.height, anchor: 'center' }
    case 'image-size': return { width: state.width, height: state.height, constrain: true, ppi: state.ppi, method: 'auto' }
  }
}

function GenericDialog({ context, request, close }: DialogProps) {
  const fields = genericFields(request) ?? []
  const [values, setValues] = useState<Record<string, unknown>>(() => initialValues(request, context))
  // Enter in a number field commits its text and applies in the same key event: read the newest values here.
  const valuesRef = useRef(values)
  const [working, setWorking] = useState(false)
  const rootRef = useRef<HTMLDivElement>(null)
  const previewing = request.kind === 'adjustment'
  const startSize = useMemo(() => {
    const state = context.store.getState()
    return { width: state.width, height: state.height }
  }, [context])

  // Live preview of destructive adjustments on the active layer (pixels only, not masks).
  useEffect(() => {
    if (!previewing || request.kind !== 'adjustment') return
    const state = context.store.getState()
    const layerId = state.activeLayerId
    if (!layerId || state.editTarget === 'mask') return
    context.compositor.setPreview({ kind: 'adjustment', layerId, spec: values as unknown as AdjustmentSpec, selection: state.selection })
  }, [context, previewing, request, values])
  useEffect(() => () => {
    if (previewing) context.compositor.setPreview(null)
  }, [context, previewing])

  useLayoutEffect(() => {
    rootRef.current?.querySelector<HTMLElement>('input:not([type="range"]), select, button')?.focus()
  }, [])

  const change = (next: Record<string, unknown>, path: string) => {
    if (request.kind === 'image-size' && next.constrain && (path === 'width' || path === 'height')) {
      const ratio = startSize.width / startSize.height
      if (path === 'width') next = { ...next, height: Math.max(1, Math.round(Number(next.width) / ratio)) }
      else next = { ...next, width: Math.max(1, Math.round(Number(next.height) * ratio)) }
    }
    valuesRef.current = next
    setValues(next)
  }

  const apply = async () => {
    if (working) return
    setWorking(true)
    const ctx = context.commands
    const values = valuesRef.current
    let ok = false
    try {
      if (previewing) context.compositor.setPreview(null)
      switch (request.kind) {
        case 'adjustment': ok = await applyAdjustmentToLayer(ctx, values as unknown as AdjustmentSpec); break
        case 'filter': ok = await applyFilterToLayer(ctx, values as unknown as FilterSpec); break
        case 'modify-selection': ok = await modifySelection(ctx, request.operation, Number(values.amount)); break
        case 'rotate-canvas': ok = await rotateCanvas(ctx, Number(values.angle)); break
        case 'canvas-size': ok = resizeCanvas(ctx, Number(values.width), Number(values.height), values.anchor as CanvasAnchor); break
        case 'image-size': ok = await resizeImage(ctx, Number(values.width), Number(values.height), { method: values.method as ResampleMethod, ppi: Number(values.ppi) }); break
      }
    } finally {
      setWorking(false)
    }
    if (ok) {
      close()
      context.focusCanvas()
    }
  }

  const cancel = () => {
    if (working) return
    close()
    context.focusCanvas()
  }

  const onKeyDown = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (event.key === 'Escape') {
      event.preventDefault()
      cancel()
    } else if (event.key === 'Enter' && !(event.target instanceof HTMLButtonElement)) {
      event.preventDefault()
      void apply()
    }
    // The dialog owns the keyboard: nothing typed here becomes an editor shortcut.
    event.stopPropagation()
  }

  const megapixels = request.kind === 'image-size' ? (Number(values.width) * Number(values.height)) / 1_000_000 : 0
  return (
    <div ref={rootRef} className="ae-dialog" role="dialog" aria-modal="true" aria-labelledby="ae-dialog-title" onKeyDown={onKeyDown}>
      <h2 id="ae-dialog-title">{dialogTitle(request)}</h2>
      <SettingsForm fields={fields} value={values} disabled={working} onChange={change} />
      {request.kind === 'image-size' && (
        <p className="ae-dialog-note">{Number.isFinite(megapixels) ? `${megapixels.toFixed(1)} megapixels` : ''}{megapixels * 1_000_000 > LIMITS.maxPixels ? ' · larger than the 50 megapixel limit' : ''}</p>
      )}
      {request.kind === 'adjustment' && context.store.getState().selection && <p className="ae-dialog-note">Only the selected area changes.</p>}
      <div className="ae-dialog-actions">
        <button type="button" onClick={cancel} disabled={working}>Cancel</button>
        <button type="button" className="ae-primary" onClick={() => void apply()} disabled={working}>{working ? 'Working…' : 'OK'}</button>
      </div>
    </div>
  )
}

export function DialogHost({ controller, context, suspended }: { readonly controller: DialogController; readonly context: PanelContext; readonly suspended: boolean }) {
  const request = useSyncExternalStore(useCallback((listener: () => void) => controller.subscribe(listener), [controller]), () => controller.current())
  const ref = useInert<HTMLDivElement>(suspended)
  if (!request) return null
  const definition = DIALOGS.find((entry) => entry.handles(request))
  const Component = definition ? definition.component : GenericDialog
  return (
    <div ref={ref} className="ae-dialog-layer">
      <div className="ae-dialog-backdrop" onPointerDown={(event) => event.preventDefault()} />
      <Component key={JSON.stringify(request)} context={context} request={request} close={() => controller.close()} />
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// Placeholder panels
// ---------------------------------------------------------------------------------------------

const THUMB = 36

function LayerThumb({ layer, store, version }: { readonly layer: Layer; readonly store: DocumentStore; readonly version: number }) {
  const ref = useRef<HTMLCanvasElement>(null)
  useEffect(() => {
    const canvas = ref.current
    if (!canvas || layer.kind === 'adjustment') return
    const state = store.getState()
    if (!state.layers.some((entry) => entry.id === layer.id)) return
    const level = Math.min(maxPyramidLevel(state.width, state.height), Math.max(0, Math.ceil(Math.log2(Math.max(state.width, state.height) / THUMB))))
    const width = Math.max(1, Math.ceil(state.width / 2 ** level))
    const height = Math.max(1, Math.ceil(state.height / 2 ** level))
    const context = canvas.getContext('2d')
    if (!context) return
    try {
      const pixels = compositeRect(state, { x: 0, y: 0, width, height }, { level, onlyLayerId: layer.id })
      context.clearRect(0, 0, canvas.width, canvas.height)
      const image = new ImageData(pixels.data as Uint8ClampedArray<ArrayBuffer>, pixels.width, pixels.height)
      context.putImageData(image, Math.floor((canvas.width - width) / 2), Math.floor((canvas.height - height) / 2))
    } catch (error) {
      console.error(error)
    }
  }, [layer, store, version])
  useEffect(() => {
    const canvas = ref.current
    return () => {
      if (canvas) {
        canvas.width = 1
        canvas.height = 1
      }
    }
  }, [])
  if (layer.kind === 'adjustment') return <span className="ae-thumb is-icon"><SlidersHorizontal aria-hidden="true" /></span>
  return <canvas ref={ref} className="ae-thumb" width={THUMB} height={THUMB} aria-hidden="true" />
}

function layerKindLabel(layer: Layer): string {
  if (layer.kind === 'raster') return layer.isBackground ? 'Background' : 'Pixel layer'
  if (layer.kind === 'text') return 'Type layer'
  if (layer.kind === 'shape') return 'Shape layer'
  return `${adjustmentLabel(layer.adjustment.type)} adjustment`
}

const BLEND_OPTIONS = BLEND_MODE_MENU

function LayersPlaceholder({ context, suspended }: PanelProps) {
  const { store } = context
  const layers = useDocumentSelector(store, (state) => state.layers)
  const activeId = useDocumentSelector(store, (state) => state.activeLayerId)
  const editTarget = useDocumentSelector(store, (state) => state.editTarget)
  const thumbVersion = usePixelVersion(store, 300)
  const [renaming, setRenaming] = useState<LayerId | null>(null)
  const active = layers.find((layer) => layer.id === activeId) ?? null
  const isBackground = Boolean(active && active.kind === 'raster' && active.isBackground)

  const transact = (label: string, run: Parameters<DocumentStore['transact']>[2], coalesceKey?: string) => {
    try {
      store.transact(label, 'layer', run, coalesceKey ? { coalesceKey } : undefined)
    } catch (error) {
      context.host.notify(error instanceof Error ? error.message : 'That could not be changed.', 'error')
    }
  }

  const select = (layer: Layer, target: 'pixels' | 'mask' = 'pixels') => {
    transact('Select Layer', (tx) => tx.setActiveLayer(layer.id, target))
  }

  const toggleVisible = (layer: Layer, solo: boolean) => {
    if (solo) {
      // Alt+click: show only this layer, or show everything again when it already is the only one.
      const others = layers.filter((entry) => entry.id !== layer.id)
      const soloed = layer.visible && others.every((entry) => !entry.visible)
      transact(soloed ? 'Show All Layers' : 'Hide Other Layers', (tx) => {
        for (const entry of layers) tx.updateLayer(entry.id, { visible: soloed ? true : entry.id === layer.id })
      })
      return
    }
    transact(layer.visible ? 'Hide Layer' : 'Show Layer', (tx) => tx.updateLayer(layer.id, { visible: !layer.visible }))
  }

  const rename = (layer: Layer, name: string) => {
    setRenaming(null)
    const trimmed = name.trim()
    if (trimmed && trimmed !== layer.name) transact('Rename Layer', (tx) => tx.updateLayer(layer.id, { name: trimmed }))
    context.focusCanvas()
  }

  const onListKey = (event: ReactKeyboardEvent<HTMLDivElement>) => {
    if (renaming) return
    const index = layers.findIndex((layer) => layer.id === activeId)
    if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
      event.preventDefault()
      event.stopPropagation()
      const next = layers[index + (event.key === 'ArrowUp' ? 1 : -1)]
      if (next) select(next)
    } else if (event.key === 'F2' && active) {
      event.preventDefault()
      event.stopPropagation()
      setRenaming(active.id)
    }
  }

  return (
    <div className="ae-layers">
      <div className="ae-layer-props">
        <select
          aria-label="Blend mode"
          value={active?.blendMode ?? 'normal'}
          disabled={!active || isBackground || suspended}
          onChange={(event) => active && transact('Blending Change', (tx) => tx.updateLayer(active.id, { blendMode: event.currentTarget.value as BlendMode }))}
        >
          {BLEND_OPTIONS.map((mode, index) => mode === '-'
            ? <option key={`sep-${index}`} disabled>──────────</option>
            : <option key={mode} value={mode}>{blendModeLabel(mode)}</option>)}
        </select>
        <label className="ae-opacity">
          <span>Opacity</span>
          <input
            type="number"
            min={0}
            max={100}
            step={1}
            aria-label="Layer opacity in percent"
            value={active ? Math.round(active.opacity * 100) : 100}
            disabled={!active || isBackground || suspended}
            onChange={(event) => {
              const value = Number(event.currentTarget.value)
              if (active && Number.isFinite(value)) transact('Opacity Change', (tx) => tx.updateLayer(active.id, { opacity: Math.min(100, Math.max(0, value)) / 100 }), `opacity:${active.id}`)
            }}
          />
          <span aria-hidden="true">%</span>
        </label>
      </div>
      <div className="ae-layer-list" role="listbox" aria-label="Layers" tabIndex={0} onKeyDown={onListKey}>
        {[...layers].reverse().map((layer) => {
          const selected = layer.id === activeId
          const background = layer.kind === 'raster' && layer.isBackground
          const locked = layer.locks.pixels || layer.locks.position || layer.locks.transparency
          return (
            <div
              key={layer.id}
              role="option"
              aria-selected={selected}
              className={`ae-layer-row ${selected ? 'is-active' : ''} ${layer.visible ? '' : 'is-hidden'} ${layer.clipped ? 'is-clipped' : ''}`}
              onClick={() => select(layer)}
              onDoubleClick={() => setRenaming(layer.id)}
              title={layerKindLabel(layer)}
            >
              <button
                type="button"
                className="ae-eye"
                aria-label={layer.visible ? `Hide ${layer.name}` : `Show ${layer.name}`}
                aria-pressed={layer.visible}
                title={layer.visible ? 'Hide (Alt-click: show only this layer)' : 'Show'}
                disabled={suspended}
                onMouseDown={(event) => event.preventDefault()}
                onClick={(event) => {
                  event.stopPropagation()
                  toggleVisible(layer, event.altKey)
                }}
              >{layer.visible ? <Eye /> : <EyeOff />}</button>
              {layer.clipped && <span className="ae-clip-mark" aria-label="Clipped to the layer below">↳</span>}
              <LayerThumb layer={layer} store={store} version={thumbVersion} />
              {layer.mask && (
                <button
                  type="button"
                  className={`ae-mask-chip ${selected && editTarget === 'mask' ? 'is-target' : ''} ${layer.mask.enabled ? '' : 'is-disabled'}`}
                  title={layer.mask.enabled ? 'Layer mask (click to paint on it)' : 'Layer mask (disabled)'}
                  aria-label={`Layer mask of ${layer.name}`}
                  onMouseDown={(event) => event.preventDefault()}
                  onClick={(event) => {
                    event.stopPropagation()
                    select(layer, 'mask')
                  }}
                ><Link2 aria-hidden="true" /></button>
              )}
              {renaming === layer.id ? (
                <input
                  className="ae-rename"
                  aria-label="Layer name"
                  defaultValue={layer.name}
                  autoFocus
                  onClick={(event) => event.stopPropagation()}
                  onBlur={(event) => rename(layer, event.currentTarget.value)}
                  onKeyDown={(event) => {
                    event.stopPropagation()
                    if (event.key === 'Enter') rename(layer, event.currentTarget.value)
                    if (event.key === 'Escape') {
                      setRenaming(null)
                      context.focusCanvas()
                    }
                  }}
                />
              ) : (
                <span className="ae-layer-name">
                  {layer.kind === 'text' && <TypeIcon className="ae-kind" aria-hidden="true" />}
                  {layer.kind === 'shape' && <Shapes className="ae-kind" aria-hidden="true" />}
                  <span className={background ? 'is-italic' : undefined}>{layer.name}</span>
                </span>
              )}
              {locked ? <Lock className="ae-lock" aria-label="Locked" /> : <span />}
            </div>
          )
        })}
      </div>
      <div className="ae-panel-footer">
        <button type="button" title={`New Layer (${context.shortcut('layer.new')})`} aria-label="New layer" disabled={suspended || !context.isEnabled('layer.new')} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('layer.new')}><Plus /></button>
        <select
          className="ae-adjust-menu"
          aria-label="New adjustment layer"
          title="New adjustment layer"
          value=""
          disabled={suspended}
          onChange={(event) => {
            const type = event.currentTarget.value as AdjustmentType
            if (type) context.run(`adjustment-layer.${type}` as EditorCommandId)
          }}
        >
          <option value="">Adjust…</option>
          {ADJUSTMENT_TYPES.map((type) => <option key={type} value={type}>{adjustmentLabel(type)}</option>)}
        </select>
        <button type="button" title="Add layer mask" aria-label="Add layer mask" disabled={suspended || !context.isEnabled('layer.add-mask')} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('layer.add-mask')}><SquareDashed /></button>
        <button type="button" title="Delete layer" aria-label="Delete layer" disabled={suspended || !context.isEnabled('layer.delete')} onMouseDown={(event) => event.preventDefault()} onClick={() => context.run('layer.delete')}><Trash2 /></button>
      </div>
    </div>
  )
}

function PropertiesPlaceholder({ context, suspended }: PanelProps) {
  const { store } = context
  const state = useDocumentSelector(store, (current) => current, (a, c) => a.layers === c.layers && a.activeLayerId === c.activeLayerId && a.width === c.width && a.height === c.height && a.ppi === c.ppi)
  const active = state.layers.find((layer) => layer.id === state.activeLayerId) ?? null
  if (active && active.kind === 'adjustment') {
    const fields = ADJUSTMENT_FIELDS[active.adjustment.type]
    const layerId = active.id
    return (
      <div className="ae-properties">
        <div className="ae-prop-title"><SlidersHorizontal aria-hidden="true" /><span>{adjustmentLabel(active.adjustment.type)}</span></div>
        {fields && fields.length ? (
          <SettingsForm
            fields={fields}
            value={active.adjustment}
            disabled={suspended}
            onChange={(next) => {
              try {
                store.transact(`Modify ${adjustmentLabel(active.adjustment.type)} Layer`, 'adjustment', (tx) => tx.updateLayer(layerId, { adjustment: next }), { coalesceKey: `adjustment:${layerId}` })
              } catch (error) {
                context.host.notify(error instanceof Error ? error.message : 'The adjustment could not be changed.', 'error')
              }
            }}
          />
        ) : <p className="ae-muted">These settings are edited in the full Properties panel, which this version does not include yet.</p>}
      </div>
    )
  }
  const rows: [string, string][] = [['Document', `${state.width.toLocaleString()} × ${state.height.toLocaleString()} px`], ['Resolution', `${state.ppi} ppi`]]
  if (active) {
    rows.unshift(['Layer', active.name], ['Kind', layerKindLabel(active)])
    const bounds = active.kind === 'raster'
      ? (() => {
        const local = active.surface.contentBounds()
        return local ? { x: local.x + active.offsetX, y: local.y + active.offsetY, width: local.width, height: local.height } : null
      })()
      : active.kind === 'text' || active.kind === 'shape' ? (() => {
        const local = active.raster.surface.contentBounds()
        return local ? { x: local.x + active.raster.offsetX, y: local.y + active.raster.offsetY, width: local.width, height: local.height } : null
      })() : null
    if (bounds) rows.push(['Position', `X ${bounds.x}  Y ${bounds.y}`], ['Size', `${bounds.width} × ${bounds.height} px`])
    rows.push(['Opacity', `${Math.round(active.opacity * 100)}%`], ['Blend', blendModeLabel(active.blendMode)])
  }
  const memory = store.memoryUsage()
  rows.push(['Memory', formatBytes(memory.total)])
  return (
    <dl className="ae-properties ae-prop-list">
      {rows.map(([term, value]) => <div key={term}><dt>{term}</dt><dd title={value}>{value}</dd></div>)}
    </dl>
  )
}

function HistoryPlaceholder({ context, suspended }: PanelProps) {
  const history = useHistoryState(context.store)
  const listRef = useRef<HTMLDivElement>(null)
  useEffect(() => {
    listRef.current?.querySelector('[aria-selected="true"]')?.scrollIntoView({ block: 'nearest' })
  }, [history.cursor, history.entries.length])
  const jump = (index: number) => {
    try {
      context.store.history.jumpTo(index)
    } catch (error) {
      context.host.notify(error instanceof Error ? error.message : 'History could not move to that step.', 'error')
    }
    context.focusCanvas()
  }
  return (
    <div ref={listRef} className="ae-history" role="listbox" aria-label="History">
      {history.trimmed && <p className="ae-muted ae-history-note">Older steps were removed to save memory.</p>}
      {history.entries.map((entry, index) => (
        <div
          key={entry.id}
          role="option"
          aria-selected={index === history.cursor}
          aria-disabled={suspended || undefined}
          className={`ae-history-row ${index === history.cursor ? 'is-current' : ''} ${index > history.cursor ? 'is-future' : ''}`}
          onClick={() => !suspended && jump(index)}
        >
          {index === 0 ? <LayersIcon aria-hidden="true" /> : <span className="ae-history-dot" aria-hidden="true" />}
          <span>{entry.label}</span>
        </div>
      ))}
    </div>
  )
}

function ColorPlaceholder({ context, suspended }: PanelProps) {
  const editor = useEditorState(context.editor)
  const [hex, setHex] = useState(toHex(editor.foreground))
  useEffect(() => setHex(toHex(editor.foreground)), [editor.foreground])
  const setForeground = (color: Rgb8) => {
    context.editor.update({ foreground: color, recentColors: rememberColor(editor.recentColors, color) })
  }
  return (
    <div className="ae-color">
      <div className="ae-color-row">
        <input type="color" aria-label="Foreground colour" value={toHex(editor.foreground)} disabled={suspended} onChange={(event) => { const color = parseHex(event.currentTarget.value); if (color) setForeground(color) }} />
        <input
          className="ae-hex"
          aria-label="Foreground colour (hex)"
          value={hex}
          disabled={suspended}
          onChange={(event) => setHex(event.currentTarget.value)}
          onBlur={() => { const color = parseHex(hex); if (color) setForeground(color); else setHex(toHex(editor.foreground)) }}
          onKeyDown={(event) => {
            event.stopPropagation()
            if (event.key === 'Enter') {
              const color = parseHex(hex)
              if (color) setForeground(color)
            }
          }}
        />
        <button
          type="button"
          className="ae-add-swatch"
          title="Add the foreground colour to the swatches"
          aria-label="Add swatch"
          disabled={suspended || editor.swatches.some((swatch) => sameColor(swatch, editor.foreground))}
          onClick={() => context.editor.update({ swatches: [...editor.swatches, editor.foreground].slice(-MAX_SWATCHES) })}
        ><Plus /></button>
      </div>
      <div className="ae-swatches" role="group" aria-label="Swatches">
        {editor.swatches.map((swatch, index) => (
          <button
            key={`${toHex(swatch)}-${index}`}
            type="button"
            className="ae-swatch"
            style={{ background: toHex(swatch) }}
            title={`${toHex(swatch)} (click: foreground, Alt-click: background)`}
            aria-label={`Swatch ${toHex(swatch)}`}
            disabled={suspended}
            onMouseDown={(event) => event.preventDefault()}
            onClick={(event) => (event.altKey ? context.editor.update({ background: swatch }) : setForeground(swatch))}
          />
        ))}
      </div>
      {editor.recentColors.length > 0 && (
        <div className="ae-swatches is-recent" role="group" aria-label="Recent colours">
          {editor.recentColors.map((color, index) => (
            <button key={`${toHex(color)}-${index}`} type="button" className="ae-swatch" style={{ background: toHex(color) }} aria-label={`Recent ${toHex(color)}`} disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => setForeground(color)} />
          ))}
        </div>
      )}
    </div>
  )
}

const PLACEHOLDERS: Readonly<Record<PanelId, PanelDefinition>> = Object.freeze({
  layers: { id: 'layers', title: 'Layers', component: LayersPlaceholder, grow: 2 },
  properties: { id: 'properties', title: 'Properties', component: PropertiesPlaceholder, grow: 1 },
  history: { id: 'history', title: 'History', component: HistoryPlaceholder, grow: 1 },
  color: { id: 'color', title: 'Color', component: ColorPlaceholder, grow: 0 },
})

/** The definition that renders a panel: a registered module, otherwise the built-in placeholder. */
export function panelDefinition(id: PanelId): PanelDefinition {
  return PANELS.get(id) ?? PLACEHOLDERS[id]
}

// ---------------------------------------------------------------------------------------------
// The dock
// ---------------------------------------------------------------------------------------------

const COLLAPSE_KEY = 'simple-image:advanced-collapsed-panels'

function readCollapsed(): Set<PanelId> {
  try {
    const value = JSON.parse(localStorage.getItem(COLLAPSE_KEY) ?? '[]') as unknown
    return new Set(Array.isArray(value) ? value.filter((id): id is PanelId => PANEL_ORDER.includes(id as PanelId)) : [])
  } catch {
    return new Set()
  }
}

export interface PanelHostProps {
  readonly context: PanelContext
  readonly suspended: boolean
  /** Shown as an overlay (narrow windows) instead of a docked column. */
  readonly overlay?: boolean
  readonly onCloseOverlay?: () => void
}

export function PanelHost({ context, suspended, overlay, onCloseOverlay }: PanelHostProps): ReactNode {
  const editor = useEditorState(context.editor)
  const [collapsed, setCollapsed] = useState<Set<PanelId>>(readCollapsed)
  const ref = useInert<HTMLElement>(suspended)
  const visible = PANEL_ORDER.filter((id) => editor.panels[id])
  if (!visible.length) return null

  const toggleCollapsed = (id: PanelId) => {
    setCollapsed((current) => {
      const next = new Set(current)
      if (next.has(id)) next.delete(id)
      else next.add(id)
      try {
        localStorage.setItem(COLLAPSE_KEY, JSON.stringify([...next]))
      } catch {
        // A per-viewer convenience only.
      }
      return next
    })
  }

  return (
    <aside ref={ref} className={`ae-dock ${overlay ? 'is-overlay' : ''}`} aria-label="Panels">
      {overlay && <button type="button" className="ae-dock-close" onClick={onCloseOverlay} aria-label="Close panels">×</button>}
      {visible.map((id) => {
        const definition = panelDefinition(id)
        const Component = definition.component
        const open = !collapsed.has(id)
        const grow = definition.grow ?? 1
        return (
          <section key={id} className={`ae-panel ${open ? 'is-open' : 'is-collapsed'}`} style={open && grow > 0 ? { flexGrow: grow } : undefined} aria-labelledby={`ae-panel-${id}`}>
            <h3 id={`ae-panel-${id}`}>
              <button type="button" aria-expanded={open} onMouseDown={(event) => event.preventDefault()} onClick={() => toggleCollapsed(id)}>
                {open ? <ChevronDown aria-hidden="true" /> : <ChevronRight aria-hidden="true" />}
                <span>{definition.title}</span>
              </button>
            </h3>
            {open && <div className="ae-panel-body"><Component context={context} suspended={suspended} /></div>}
          </section>
        )
      })}
    </aside>
  )
}
