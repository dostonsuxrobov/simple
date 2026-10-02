// src/advanced/AdvancedEditor.tsx (WP6)
// The Advanced (Photoshop-style) editor shell (design 3.2, 5.6, 5.9, 5.14, 5.15), loaded lazily by
// useAdvancedMode.tsx. It keeps the WP1 contract:
//   - export function AdvancedEditor(props: AdvancedEditorProps); `initial` is read once; a 'canvas' is the
//     live Simple canvas, read in 256-row stripes and never modified. It becomes the Background (or
//     "Layer 0" when it has transparency); a PSD opens with its layers.
//   - props.onReady(handle) once the document is ready; the handle implements AdvancedEditorHandle plus the
//     optional AdvancedEditorHandleExtras (handleKeyDown, handleKeyUp, canEncodePsd).
//   - Input is ignored while props.suspended; the host keeps Ctrl+O / S / P and Alt+Shift+Ctrl+W.
//   - The document store starts at the host's current revision (AdvancedHostExtras.currentRevision), so
//     undoing back to the opening state restores the host's Modified state exactly.
// Layout: options bar (menus, tool options, Undo / Redo, the "Simple" button) / toolbox with foreground and
// background colours / canvas / panel dock (PanelHost) / status bar. Below 1080 px the panels become an
// overlay opened from the options bar.
// Edits are never dropped silently: settle() (Save, Export, Print, Close, leaving Advanced) commits open
// text, crop, transform and lasso sessions and waits for running tool and command jobs; tools commit their
// sessions when another tool is chosen (WP5). Keys reach the editor only through the host (one window
// listener); single-key shortcuts never fire while a text field or the text tool has focus.
import { useCallback, useEffect, useLayoutEffect, useMemo, useRef, useState, useSyncExternalStore } from 'react'
import type { ComponentType, PointerEvent as ReactPointerEvent, ReactNode } from 'react'
import {
  AlignCenter,
  AlignLeft,
  AlignRight,
  ArrowLeftRight,
  ArrowLeftToLine,
  Bandage,
  Blend,
  Bold,
  Brush,
  Check,
  CircleDashed,
  Crop,
  Eraser,
  Hand,
  Italic,
  Lasso,
  LassoSelect,
  LoaderCircle,
  Move,
  PaintBucket,
  PanelRight,
  Pipette,
  Redo2,
  Shapes,
  SquareDashed,
  Stamp,
  TriangleAlert,
  Type,
  Underline,
  Undo2,
  WandSparkles,
  X,
  ZoomIn,
} from 'lucide-react'
import type { LucideIcon } from 'lucide-react'
import type { BlendMode, PixelBuffer, Point, Rgb8, Rgba8, SelectionOp } from '../imaging/types.ts'
import type {
  AdvancedEditorHandle,
  AdvancedEditorProps,
  AdvancedHost,
  AdvancedInitialContent,
  EditorState,
  Layer,
  LayerId,
  PsdFidelityIssue,
  SampleSource,
  ShapeKind,
  ToolId,
  ToolOptions,
} from './types.ts'
import { BLEND_MODE_MENU, LIMITS } from './types.ts'
import type { AdvancedEditorHandleExtras, AdvancedHostExtras } from './useAdvancedMode.tsx'
import type { AdvancedDocumentStore } from './document.ts'
import { createDocumentStore, createRasterLayer, isFlatEquivalent } from './document.ts'
import { createSurface } from './tiles.ts'
import type { AdvancedCompositor } from './compositor.ts'
import { createCompositor } from './compositor.ts'
import { formatBytes, setExternalBytes } from './memory.ts'
import { zoomLabel } from './viewport.ts'
import type { AdvancedToolController, FreeTransformController } from './tools/index.ts'
import { TOOL_FACTORIES, TOOL_GROUPS, TOOL_META, createFreeTransform } from './tools/index.ts'
import type { CanvasInput, CanvasSurface, ViewModel } from './CanvasView.tsx'
import { CanvasView, createCanvasSurface, createViewModel } from './CanvasView.tsx'
import { MenuBar } from './menus/MenuBar.tsx'
import type { DialogController, PanelContext } from './panels/PanelHost.tsx'
import { DialogHost, PanelHost, createDialogController, useEditorState, useHistoryState } from './panels/PanelHost.tsx'
import { releaseHistorySnapshots } from './panels/HistoryPanel.tsx'
import { ColorPickerPopover } from './panels/ColorPicker.tsx'
import type { CommandContext, CommandServices, EditorCommandId } from './commands.ts'
import {
  attachCommands,
  commandLabel,
  commandsBusy,
  copyPixels,
  isCommandEnabled,
  placePixelsAsLayer,
  runCommandAsync,
  whenCommandsIdle,
} from './commands.ts'
import type { AdvancedEditorStore } from './editorState.ts'
import { createEditorStore, EDITOR_STORAGE_KEY, isToolId, rememberColor } from './editorState.ts'
import { releasesSpring, shortcutForEvent, shortcutLabel, springForEvent, withShortcut } from './shortcuts.ts'
import { getImagingClient } from '../shared/workerClient.ts'
import { bufferToCanvas, readCanvasStripes } from '../shared/canvas.ts'
import { availableFontFamilies } from './fonts.ts'
import { blendModeLabel } from '../imaging/blend.ts'
import { parseHex, toHex } from '../imaging/color.ts'
import './advanced.css'

// ---------------------------------------------------------------------------------------------
// Small external stores (status bar, busy state) so hot updates re-render only their subscribers
// ---------------------------------------------------------------------------------------------

interface Cell<T> {
  get(): T
  set(value: T): void
  subscribe(listener: () => void): () => void
}

function cell<T>(initial: T, equal: (a: T, b: T) => boolean = Object.is): Cell<T> {
  let value = initial
  const listeners = new Set<() => void>()
  return {
    get: () => value,
    set(next: T): void {
      if (equal(value, next)) return
      value = next
      for (const listener of [...listeners]) listener()
    },
    subscribe(listener: () => void): () => void {
      listeners.add(listener)
      return () => { listeners.delete(listener) }
    },
  }
}

function useCell<T>(source: Cell<T>): T {
  return useSyncExternalStore(useCallback((listener: () => void) => source.subscribe(listener), [source]), source.get)
}

// ---------------------------------------------------------------------------------------------
// The editor session (everything that lives as long as the document)
// ---------------------------------------------------------------------------------------------

interface InitialDocument {
  readonly width: number
  readonly height: number
  readonly ppi: number
  /** Bottom to top. */
  readonly layers: readonly Layer[]
  readonly baseLabel: string
  readonly issues: readonly PsdFidelityIssue[]
  /** Bytes the host keeps outside the document (the Simple canvas). */
  readonly externalBytes: number
}

interface HoverInfo {
  readonly x: number
  readonly y: number
  readonly color: Rgba8 | null
}

type EditorHandle = AdvancedEditorHandle & AdvancedEditorHandleExtras

interface EditorSession {
  readonly store: AdvancedDocumentStore
  readonly compositor: AdvancedCompositor
  readonly editor: AdvancedEditorStore
  readonly view: ViewModel
  readonly surface: CanvasSurface
  readonly dialogs: DialogController
  readonly context: CommandContext
  readonly panelContext: PanelContext
  readonly input: CanvasInput
  readonly hint: Cell<string | null>
  readonly hover: Cell<HoverInfo | null>
  readonly busy: Cell<boolean>
  /** Bumps when the options bar must refresh outside the editor state (free-transform fields, crop session). */
  readonly optionsTick: Cell<number>
  readonly handle: EditorHandle
  transform(): FreeTransformController | null
  activeTool(): AdvancedToolController | null
  run(command: EditorCommandId): void
  isEnabled(command: EditorCommandId): boolean
  dispose(): void
}

/** Tools whose overlay is only a hover cursor (hidden when the pointer leaves the canvas). */
const HOVER_ONLY_TOOLS: ReadonlySet<ToolId> = new Set<ToolId>(['brush', 'eraser', 'clone-stamp', 'spot-healing'])

/** Keys a focused button, slider or list uses itself (never turned into editor shortcuts there). */
const CONTROL_KEYS: ReadonlySet<string> = new Set(['Enter', ' ', 'Tab', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight', 'Home', 'End', 'PageUp', 'PageDown'])

type FocusKind = 'canvas' | 'control' | 'text'

/** Where a key event happened: a text field (all keys belong to it), another control, or the canvas / page. */
function focusKind(target: EventTarget | null): FocusKind {
  if (!(target instanceof Element)) return 'canvas'
  const field = target.closest('input, textarea, select, [contenteditable=""], [contenteditable="true"]')
  if (field) {
    if (field instanceof HTMLInputElement && ['range', 'checkbox', 'radio', 'color', 'button', 'submit', 'reset'].includes(field.type)) return 'control'
    return 'text'
  }
  if (target.closest('button, [role="menuitem"], [role="option"], [role="listbox"], [role="slider"], a[href]')) return 'control'
  return 'canvas'
}

function messageOf(error: unknown, fallback: string): string {
  return error instanceof Error && error.message ? error.message : fallback
}

async function encodePng(pixels: PixelBuffer): Promise<Uint8Array> {
  const canvas = new OffscreenCanvas(pixels.width, pixels.height)
  try {
    const context = canvas.getContext('2d')
    if (!context) throw new Error('The image could not be prepared for the clipboard.')
    context.putImageData(new ImageData(pixels.data as Uint8ClampedArray<ArrayBuffer>, pixels.width, pixels.height), 0, 0)
    const blob = await canvas.convertToBlob({ type: 'image/png' })
    return new Uint8Array(await blob.arrayBuffer())
  } finally {
    canvas.width = 1
    canvas.height = 1
  }
}

async function decodePng(png: Uint8Array): Promise<PixelBuffer> {
  const bitmap = await createImageBitmap(new Blob([png as BlobPart], { type: 'image/png' }), { premultiplyAlpha: 'none', colorSpaceConversion: 'none' })
  const canvas = new OffscreenCanvas(Math.max(1, bitmap.width), Math.max(1, bitmap.height))
  try {
    if (bitmap.width * bitmap.height > LIMITS.maxPixels || bitmap.width > LIMITS.maxDimension || bitmap.height > LIMITS.maxDimension) {
      throw new Error('The image on the clipboard is larger than 50 megapixels.')
    }
    const context = canvas.getContext('2d', { willReadFrequently: true })
    if (!context) throw new Error('The clipboard image could not be read.')
    context.drawImage(bitmap, 0, 0)
    const data = context.getImageData(0, 0, bitmap.width, bitmap.height)
    return { width: data.width, height: data.height, data: data.data }
  } finally {
    bitmap.close()
    canvas.width = 1
    canvas.height = 1
  }
}

function readInitial(initial: AdvancedInitialContent): InitialDocument {
  if (initial.kind === 'document') {
    const document = initial.document
    return { width: document.width, height: document.height, ppi: document.ppi, layers: document.layers, baseLabel: 'Open', issues: document.issues, externalBytes: 0 }
  }
  const canvas = initial.canvas
  if (!canvas.width || !canvas.height) throw new Error('The image is empty.')
  const surface = createSurface()
  let transparent = false
  readCanvasStripes(canvas, 256, (top, stripe) => {
    if (!transparent) {
      const data = stripe.data
      for (let p = 3; p < data.length; p += 4) {
        if (data[p] < 255) {
          transparent = true
          break
        }
      }
    }
    surface.write(0, top, stripe)
  })
  // Photoshop: an opaque image opens as the locked Background; one with transparency as "Layer 0".
  const layer = transparent
    ? createRasterLayer({ name: 'Layer 0', surface })
    : createRasterLayer({ name: 'Background', surface, isBackground: true })
  return {
    width: canvas.width,
    height: canvas.height,
    ppi: initial.ppi > 0 ? initial.ppi : 72,
    layers: [layer],
    baseLabel: 'Advanced editor',
    issues: [],
    externalBytes: canvas.width * canvas.height * 4,
  }
}

/** Present only while a QA script drives the editor (see createEditorSession). */
interface QaWindow {
  __SIMPLE_IMAGE_QA__?: boolean
  __simpleAdvanced?: { readonly store: AdvancedDocumentStore; readonly compositor: AdvancedCompositor; readonly editor: AdvancedEditorStore; readonly view: ViewModel; run(command: EditorCommandId): void }
}
const qaWindow: QaWindow = typeof window !== 'undefined' ? window as unknown as QaWindow : {}

function createEditorSession(host: AdvancedHost, initial: InitialDocument, isSuspended: () => boolean): EditorSession {
  let disposed = false
  const extras = host as AdvancedHost & Partial<AdvancedHostExtras>
  // After the editor closes nothing may reach the host any more (a late commit must not move its revision).
  const liveHost: AdvancedHost & AdvancedHostExtras = {
    nextRevision: () => host.nextRevision(),
    setRevision: (revision) => { if (!disposed) host.setRevision(revision) },
    currentRevision: () => (typeof extras.currentRevision === 'function' ? extras.currentRevision() : 0),
    notify: (message, tone) => { if (!disposed) host.notify(message, tone) },
    isSuspended: () => host.isSuspended(),
    requestExit: () => { if (!disposed) host.requestExit() },
    save: (forceDialog) => host.save(forceDialog),
    openExportMenu: () => host.openExportMenu(),
    print: () => host.print(),
    copyPng: (png) => host.copyPng(png),
    readClipboardImage: () => host.readClipboardImage(),
  }

  const store = createDocumentStore({
    width: initial.width,
    height: initial.height,
    ppi: initial.ppi,
    layers: initial.layers,
    host: liveHost,
    baseLabel: initial.baseLabel,
    revision: liveHost.currentRevision(),
  })
  setExternalBytes(initial.externalBytes)
  const detachCommands = attachCommands(store)
  const compositor = createCompositor(store)
  const editor = createEditorStore(EDITOR_STORAGE_KEY)
  const view = createViewModel(() => {
    const state = store.getState()
    return { width: state.width, height: state.height }
  })
  const surface = createCanvasSurface()
  const dialogs = createDialogController()
  const imaging = getImagingClient()
  const hint = cell<string | null>(null)
  const hover = cell<HoverInfo | null>(null, (a, b) => a === b || Boolean(a && b && a.x === b.x && a.y === b.y && a.color?.r === b.color?.r && a.color?.g === b.color?.g && a.color?.b === b.color?.b && a.color?.a === b.color?.a))
  const busy = cell(false)
  const optionsTick = cell(0)

  // ----- tools ------------------------------------------------------------------------------
  const controllers = new Map<ToolId, AdvancedToolController>()
  let active: AdvancedToolController | null = null
  let activeId: ToolId | null = null
  let transform: FreeTransformController | null = null
  let springKey: 'space' | 'alt' | null = null
  const cursors = { tool: 'default', transform: 'default' }
  const hints: { tool: string | null; transform: string | null } = { tool: null, transform: null }

  const transformOpen = () => Boolean(transform && transform.hasSession())
  const applyCursor = () => surface.setCursor(transformOpen() ? cursors.transform : cursors.tool)
  const applyHint = () => hint.set(transformOpen() ? hints.transform ?? hints.tool : hints.tool)

  const services: CommandServices = {
    activeTool: () => active,
    transformSession: () => (transformOpen() ? transform : null),
    startTransform: () => startTransform(),
    openDialog: (request) => dialogs.open(request),
    canOpenDialog: (request) => dialogs.canOpen(request),
    clipboard: {
      async write(pixels: PixelBuffer): Promise<void> {
        await liveHost.copyPng(await encodePng(pixels))
      },
      async read(): Promise<PixelBuffer | null> {
        const png = await liveHost.readClipboardImage()
        return png && png.byteLength ? decodePng(png) : null
      },
    },
    busy: () => Boolean(active?.isBusy?.() || transform?.isBusy?.()),
  }

  const makeContext = (kind: 'tool' | 'transform'): CommandContext => ({
    store,
    editor,
    view,
    compositor,
    imaging,
    host: liveHost,
    setCursor: (cursor: string) => {
      cursors[kind] = cursor || 'default'
      applyCursor()
    },
    setHint: (text: string | null) => {
      hints[kind] = text
      applyHint()
    },
    setOverlayElement: (element: HTMLElement | null) => surface.setOverlayElement(element),
    services,
  })
  const context = makeContext('tool')
  const transformContext = makeContext('transform')

  const controllerFor = (id: ToolId): AdvancedToolController => {
    let controller = controllers.get(id)
    if (!controller) {
      controller = TOOL_FACTORIES[id]() as AdvancedToolController
      controllers.set(id, controller)
    }
    return controller
  }

  let switching = false
  const syncTool = () => {
    if (disposed || switching) return
    const wanted = editor.getState().tool
    if (wanted === activeId) return
    // A tool finishing an asynchronous commit keeps the input until it is done (WP5 contract).
    if (active?.isBusy?.()) {
      switching = true
      void (active.whenIdle?.() ?? Promise.resolve()).finally(() => {
        switching = false
        syncTool()
      })
      return
    }
    try {
      active?.deactivate()
    } catch (error) {
      console.error(error)
    }
    hints.tool = null
    cursors.tool = 'default'
    activeId = wanted
    active = controllerFor(wanted)
    try {
      active.activate(context)
    } catch (error) {
      console.error(error)
      liveHost.notify(messageOf(error, 'That tool could not be started.'), 'error')
    }
    applyCursor()
    applyHint()
    view.requestOverlay()
  }

  const startTransform = () => {
    if (disposed || transformOpen()) return
    if (isSuspended()) return
    // Photoshop applies an open crop or text edit before transforming.
    if (active?.hasSession()) active.commitSession()
    const controller = createFreeTransform({
      onEnd: () => {
        if (transform !== controller) return
        transform = null
        hints.transform = null
        applyCursor()
        applyHint()
        optionsTick.set(optionsTick.get() + 1)
        view.requestOverlay()
      },
    })
    transform = controller
    try {
      controller.activate(transformContext)
    } catch (error) {
      console.error(error)
    }
    if (!controller.hasSession()) {
      if (transform === controller) transform = null
      applyCursor()
      applyHint()
      return
    }
    applyCursor()
    applyHint()
    optionsTick.set(optionsTick.get() + 1)
    view.requestOverlay()
  }

  const endSpring = () => {
    if (!springKey) return
    springKey = null
    const state = editor.getState()
    if (state.springFrom) editor.update({ tool: state.springFrom, springFrom: null })
  }

  // ----- status: hover position and colour -------------------------------------------------
  let hoverPoint: Point | null = null
  let hoverTimer: ReturnType<typeof setTimeout> | null = null
  const sampleHover = () => {
    hoverTimer = null
    if (disposed) return
    const point = hoverPoint
    if (!point) {
      hover.set(null)
      return
    }
    const state = store.getState()
    const x = Math.floor(point.x)
    const y = Math.floor(point.y)
    let color: Rgba8 | null = null
    if (x >= 0 && y >= 0 && x < state.width && y < state.height) {
      try {
        color = compositor.sample(x, y, 1, 'all')
      } catch {
        color = null
      }
    }
    hover.set({ x, y, color })
  }

  const busyNow = () => commandsBusy(store) || Boolean(services.busy?.())
  const busyTimer = setInterval(() => busy.set(busyNow()), 120)

  const input: CanvasInput = {
    target: () => (transformOpen() ? transform : active),
    busy: () => busyNow(),
    stripAlt: () => springKey === 'alt',
    hoverOnly: (controller) => HOVER_ONLY_TOOLS.has(controller.id),
    hover: (point) => {
      hoverPoint = point
      if (hoverTimer === null) hoverTimer = setTimeout(sampleHover, 30)
    },
  }

  // ----- commands ---------------------------------------------------------------------------
  const isEnabled = (command: EditorCommandId) => !isSuspended() && isCommandEnabled(command, context)

  const run = (command: EditorCommandId) => {
    if (disposed || isSuspended()) return
    const pending = runCommandAsync(command, context)
    busy.set(busyNow())
    void pending.finally(() => busy.set(busyNow()))
    // Keep the keyboard on the canvas after menus and buttons (dialogs take focus themselves).
    if (!dialogs.current()) surface.focus()
  }

  const panelContext: PanelContext = {
    store,
    editor,
    compositor,
    imaging,
    host: liveHost,
    view,
    commands: context,
    run,
    isEnabled,
    label: (command) => commandLabel(command, context),
    shortcut: (command) => shortcutLabel(command),
    openDialog: (request) => dialogs.open(request),
    focusCanvas: () => surface.focus(),
  }

  // ----- keyboard ---------------------------------------------------------------------------
  const textSessionOpen = () => activeId === 'text' && Boolean(active?.hasSession())

  const handleKeyDown = (event: KeyboardEvent): boolean => {
    if (disposed || isSuspended()) return false
    if (event.isComposing || event.keyCode === 229) return false
    // An open dialog owns the keyboard (its own handlers stop the keys they use).
    if (dialogs.current()) return false
    const focus = focusKind(event.target)
    if (focus === 'text') return false
    const controlKey = focus === 'control' && CONTROL_KEYS.has(event.key)
    // Sessions first: free transform, then the active tool (Enter / Esc / arrows / Backspace / Space-pan).
    if (!controlKey) {
      try {
        const sessionTarget = transformOpen() ? transform : active
        if (sessionTarget && sessionTarget.keyDown(event)) {
          event.preventDefault()
          return true
        }
      } catch (error) {
        console.error(error)
      }
    }
    // Spring-loaded tools: hold Space for the Hand, Alt for the Eyedropper in paint tools.
    const state = editor.getState()
    if (springKey) {
      if ((springKey === 'space' && (event.key === ' ' || event.code === 'Space')) || (springKey === 'alt' && event.key === 'Alt')) {
        event.preventDefault()
        return true
      }
    } else if (!controlKey && !event.repeat && !surface.isPointerDown() && !transformOpen()) {
      const spring = springForEvent(event, state.tool)
      if (spring && !(spring.key === 'space' && focus === 'control')) {
        springKey = spring.key
        editor.update({ springFrom: state.tool, tool: spring.tool })
        event.preventDefault()
        return true
      }
    }
    const binding = shortcutForEvent(event)
    if (!binding) {
      // Alt alone must not move focus to a (hidden) window menu.
      if (event.key === 'Alt') event.preventDefault()
      return false
    }
    if (controlKey) return false
    if (binding.requiresCanvasFocus && textSessionOpen()) return false
    event.preventDefault()
    if (event.repeat && !binding.repeat) return true
    let command = binding.command
    // A tool letter selects the tool its toolbox slot shows (the one last used in that slot).
    if (command.startsWith('tool.')) {
      const id = command.slice(5)
      if (isToolId(id)) command = `tool.${editor.slotTool(id)}` as EditorCommandId
    }
    if (springKey && command.startsWith('tool.') && isToolId(command.slice(5))) endSpring()
    run(command)
    return true
  }

  const handleKeyUp = (event: KeyboardEvent): boolean => {
    if (disposed) return false
    if (springKey && releasesSpring(event, springKey)) {
      endSpring()
      if (event.key === 'Alt') event.preventDefault()
      return true
    }
    if (isSuspended()) return false
    try {
      const target = transformOpen() ? transform : active
      return Boolean(target?.keyUp(event))
    } catch (error) {
      console.error(error)
      return false
    }
  }

  // A key released while the window was in the background never arrives: end springs on blur.
  const onBlur = () => endSpring()
  const onVisibility = () => {
    if (document.visibilityState === 'hidden') endSpring()
  }
  window.addEventListener('blur', onBlur)
  document.addEventListener('visibilitychange', onVisibility)

  // ----- handle -----------------------------------------------------------------------------
  const settle = async (): Promise<void> => {
    if (disposed) return
    // A value typed into a panel, the options bar or the colour picker but not confirmed yet is part of the
    // work: leaving the field commits it (synchronously, through its blur handler) before Save reads.
    const focused = typeof document !== 'undefined' ? document.activeElement : null
    if (focused instanceof HTMLInputElement && focused.closest('.ae-dock, .ae-options, .ae-cp-popover')) {
      try {
        focused.blur()
      } catch (error) {
        console.error(error)
      }
    }
    // Open work is the user's work: commit it (Photoshop asks; the WP5 tools commit on switch, too).
    try {
      if (transformOpen()) transform?.commitSession()
    } catch (error) {
      console.error(error)
    }
    try {
      if (active?.hasSession()) active.commitSession()
    } catch (error) {
      console.error(error)
    }
    await Promise.all([
      transform?.whenIdle?.() ?? Promise.resolve(),
      active?.whenIdle?.() ?? Promise.resolve(),
      whenCommandsIdle(store),
    ])
    store.commitStroke()
  }

  const handle: EditorHandle = {
    settle,
    renderFlattened: () => compositor.renderToCanvas(),
    encodePsd: async () => {
      const composite = await compositor.flatten()
      const module = await import('./psd.ts')
      return module.exportPsd(store.getState(), composite)
    },
    isFlatEquivalent: () => isFlatEquivalent(store.getState()),
    layerCount: () => store.getState().layers.length,
    size: () => ({ width: store.getState().width, height: store.getState().height }),
    hasTransparency: async () => {
      const state = store.getState()
      const bottom = state.layers[0]
      // A visible Background is opaque, and nothing above it can make the composite transparent.
      if (bottom && bottom.kind === 'raster' && bottom.isBackground && bottom.visible) return false
      const flat = await compositor.flatten()
      for (let p = 3; p < flat.data.length; p += 4) if (flat.data[p] < 255) return true
      return false
    },
    copySelection: async (merged: boolean) => {
      const entry = copyPixels(store.getState(), merged)
      return entry ? bufferToCanvas(entry.pixels) : null
    },
    placeImage: (pixels: PixelBuffer, name: string) => {
      placePixelsAsLayer(context, pixels, name, 'center', 'Place')
    },
    fidelityIssues: () => initial.issues,
    focus: () => surface.focus(),
    canEncodePsd: () => true,
    handleKeyDown,
    handleKeyUp,
  }

  const unsubscribeEditor = editor.subscribe(syncTool)
  // Keep EditorState.view in step with the canvas view (throttled: panning must not re-render panels).
  let viewTimer: ReturnType<typeof setTimeout> | null = null
  const unsubscribeView = view.subscribe(() => {
    if (viewTimer !== null) return
    viewTimer = setTimeout(() => {
      viewTimer = null
      if (!disposed) editor.update({ view: view.getView() })
    }, 200)
  })
  // The free-transform fields follow the frame while it is dragged; the crop buttons follow its session.
  let cropSession = false
  const unsubscribeOverlay = view.onOverlay(() => {
    if (transformOpen()) {
      optionsTick.set(optionsTick.get() + 1)
    } else if (activeId === 'crop' || cropSession) {
      const open = activeId === 'crop' && Boolean(active?.hasSession())
      if (open !== cropSession) {
        cropSession = open
        optionsTick.set(optionsTick.get() + 1)
      }
    }
  })
  syncTool()

  // QA hook: qa/advanced-smoke.cjs and qa/advanced-perf.cjs set window.__SIMPLE_IMAGE_QA__ before entering
  // Advanced to read the document and compositor directly. Nothing is exposed otherwise.
  if (qaWindow.__SIMPLE_IMAGE_QA__ === true) qaWindow.__simpleAdvanced = { store, compositor, editor, view, run: (command: EditorCommandId) => run(command) }

  return {
    store,
    compositor,
    editor,
    view,
    surface,
    dialogs,
    context,
    panelContext,
    input,
    hint,
    hover,
    busy,
    optionsTick,
    handle,
    transform: () => (transformOpen() ? transform : null),
    activeTool: () => active,
    run,
    isEnabled,
    dispose(): void {
      if (disposed) return
      // Nothing reaches the host from here on: it settled before closing, and a replaced document is discarded.
      disposed = true
      try {
        if (transformOpen()) transform?.cancelSession()
        if (active?.hasSession()) active.cancelSession()
        active?.deactivate()
      } catch (error) {
        console.error(error)
      }
      active = null
      transform = null
      controllers.clear()
      clearInterval(busyTimer)
      if (hoverTimer !== null) clearTimeout(hoverTimer)
      if (viewTimer !== null) clearTimeout(viewTimer)
      window.removeEventListener('blur', onBlur)
      document.removeEventListener('visibilitychange', onVisibility)
      unsubscribeEditor()
      unsubscribeView()
      unsubscribeOverlay()
      detachCommands()
      dialogs.close()
      editor.dispose()
      compositor.dispose()
      releaseHistorySnapshots(store)
      store.dispose()
      setExternalBytes(0)
      if (qaWindow.__simpleAdvanced?.store === store) delete qaWindow.__simpleAdvanced
    },
  }
}

// ---------------------------------------------------------------------------------------------
// Toolbox
// ---------------------------------------------------------------------------------------------

const TOOL_ICONS: Readonly<Record<string, LucideIcon>> = Object.freeze({
  Move, SquareDashed, CircleDashed, Lasso, LassoSelect, WandSparkles, Crop, Pipette, Bandage, Brush, Stamp, Eraser, Blend, PaintBucket, Type, Shapes, Hand, ZoomIn,
})

function ToolIcon({ tool }: { readonly tool: ToolId }) {
  const Icon = TOOL_ICONS[TOOL_META[tool].icon] ?? Move
  return <Icon aria-hidden="true" />
}

function toolTitle(tool: ToolId): string {
  return withShortcut(TOOL_META[tool].label, `tool.${tool}` as EditorCommandId)
}

function css(color: Rgb8): string {
  return `rgb(${color.r}, ${color.g}, ${color.b})`
}

function Toolbox({ session, state, suspended }: { readonly session: EditorSession; readonly state: EditorState; readonly suspended: boolean }) {
  const [flyout, setFlyout] = useState<number | null>(null)
  const holdRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const rootRef = useRef<HTMLDivElement>(null)
  const current = state.tool

  useEffect(() => {
    if (flyout === null) return
    const close = (event: PointerEvent) => {
      if (rootRef.current && event.target instanceof Node && rootRef.current.contains(event.target)) return
      setFlyout(null)
    }
    document.addEventListener('pointerdown', close, true)
    return () => document.removeEventListener('pointerdown', close, true)
  }, [flyout])

  const choose = (tool: ToolId) => {
    setFlyout(null)
    session.run(`tool.${tool}` as EditorCommandId)
  }

  // The chips open the same colour picker as the Color panel (live changes, Escape restores).
  const [picker, setPicker] = useState<{ readonly which: 'foreground' | 'background'; readonly anchor: HTMLElement } | null>(null)
  useEffect(() => {
    if (suspended) setPicker(null)
  }, [suspended])
  const chip = (which: 'foreground' | 'background') => {
    const label = which === 'foreground' ? 'Foreground colour' : 'Background colour'
    return (
      <button
        type="button"
        className={`ae-chip is-${which}`}
        title={`${label} ${toHex(state[which]).toUpperCase()} (click to choose)`}
        aria-label={`${label} ${toHex(state[which]).toUpperCase()}`}
        aria-haspopup="dialog"
        aria-expanded={picker?.which === which}
        data-color={toHex(state[which])}
        style={{ background: css(state[which]) }}
        disabled={suspended}
        onClick={(event) => {
          const anchor = event.currentTarget
          setPicker((open) => (open?.which === which ? null : { which, anchor }))
        }}
      />
    )
  }

  return (
    <nav ref={rootRef} className="ae-toolbox" aria-label="Tools">
      <div className="ae-tools" role="toolbar" aria-orientation="vertical" aria-label="Tools">
        {TOOL_GROUPS.map((group, index) => {
          const shown = group.length > 1 ? (group.includes(current) ? current : session.editor.slotTool(group[0])) : group[0]
          const activeSlot = group.includes(current)
          const multi = group.length > 1
          return (
            <div key={group[0]} className="ae-tool-slot">
              <button
                type="button"
                className={`ae-tool ${activeSlot ? 'is-active' : ''} ${multi ? 'has-more' : ''}`}
                title={multi ? `${toolTitle(shown)}. Right-click or hold for more tools; Shift+${TOOL_META[shown].key} switches.` : toolTitle(shown)}
                aria-label={TOOL_META[shown].label}
                aria-pressed={activeSlot}
                aria-haspopup={multi ? 'menu' : undefined}
                aria-keyshortcuts={TOOL_META[shown].key}
                disabled={suspended}
                onMouseDown={(event) => event.preventDefault()}
                onPointerDown={() => {
                  if (!multi) return
                  holdRef.current = setTimeout(() => setFlyout(index), 450)
                }}
                onPointerUp={() => {
                  if (holdRef.current) clearTimeout(holdRef.current)
                  holdRef.current = null
                }}
                onPointerLeave={() => {
                  if (holdRef.current) clearTimeout(holdRef.current)
                  holdRef.current = null
                }}
                onContextMenu={(event) => {
                  event.preventDefault()
                  if (multi) setFlyout(index)
                }}
                onClick={() => {
                  if (flyout === index) return
                  choose(shown)
                }}
              ><ToolIcon tool={shown} /></button>
              {flyout === index && (
                <div className="ae-flyout" role="menu" aria-label={`${TOOL_META[group[0]].label} group`}>
                  {group.map((tool) => (
                    <button key={tool} type="button" role="menuitemradio" aria-checked={tool === shown} className="ae-flyout-item" onClick={() => choose(tool)}>
                      <ToolIcon tool={tool} />
                      <span>{TOOL_META[tool].label}</span>
                      <kbd>{TOOL_META[tool].key}</kbd>
                    </button>
                  ))}
                </div>
              )}
            </div>
          )
        })}
      </div>
      <div className="ae-colors" aria-label="Colours">
        {chip('foreground')}
        {chip('background')}
        <button type="button" className="ae-swap" title={withShortcut('Swap colours', 'edit.swap-colors')} aria-label="Swap colours" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('edit.swap-colors')}><ArrowLeftRight aria-hidden="true" /></button>
        <button type="button" className="ae-default-colors" title={withShortcut('Default colours (black and white)', 'edit.default-colors')} aria-label="Default colours" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('edit.default-colors')}><span /><span /></button>
      </div>
      {picker && !suspended && (
        <ColorPickerPopover
          anchor={picker.anchor}
          value={state[picker.which]}
          title={picker.which === 'foreground' ? 'Foreground colour' : 'Background colour'}
          onChange={(color, phase) => {
            const which = picker.which
            session.editor.update(phase === 'commit'
              ? { [which]: color, recentColors: rememberColor(session.editor.getState().recentColors, color) }
              : { [which]: color })
          }}
          onClose={() => {
            setPicker(null)
            session.surface.focus()
          }}
        />
      )}
    </nav>
  )
}

// ---------------------------------------------------------------------------------------------
// Options bar controls
// ---------------------------------------------------------------------------------------------

interface NumberOptionProps {
  readonly label: string
  readonly title?: string
  /** Model value (e.g. 0..1 for opacity). */
  readonly value: number
  readonly min: number
  readonly max: number
  /** Display = model * scale (100 for percentages). */
  readonly scale?: number
  readonly step?: number
  readonly unit?: string
  readonly digits?: number
  readonly disabled?: boolean
  readonly onChange: (value: number) => void
  readonly onDone?: () => void
}

function NumberOption({ label, title, value, min, max, scale = 1, step = 1, unit, digits = 0, disabled, onChange, onDone }: NumberOptionProps) {
  const format = (model: number) => (Number.isFinite(model) ? (model * scale).toFixed(digits).replace(/\.0+$/, '') : '')
  const [text, setText] = useState(format(value))
  const editing = useRef(false)
  useEffect(() => {
    if (!editing.current) setText(format(value))
  })
  const clampDisplay = (display: number) => Math.min(max * scale, Math.max(min * scale, display))
  const commit = (display: number) => {
    const clamped = clampDisplay(display)
    onChange(clamped / scale)
  }
  const fromText = () => {
    editing.current = false
    const parsed = Number(text.replace(',', '.').replace(/[^\d.+-]/g, ''))
    if (Number.isFinite(parsed) && text.trim() !== '') commit(parsed)
    setText(format(value))
  }
  // Photoshop's scrubby label: drag the label left / right to change the value.
  const scrub = (event: ReactPointerEvent<HTMLSpanElement>) => {
    if (disabled || event.button !== 0) return
    event.preventDefault()
    const target = event.currentTarget
    const startX = event.clientX
    const start = value * scale
    try {
      target.setPointerCapture(event.pointerId)
    } catch {
      // ignore
    }
    const move = (moveEvent: PointerEvent) => {
      const dx = moveEvent.clientX - startX
      const factor = moveEvent.shiftKey ? 10 : 1
      const next = Math.round((start + dx * step * factor) / step) * step
      commit(next)
    }
    const up = () => {
      target.removeEventListener('pointermove', move)
      target.removeEventListener('pointerup', up)
      target.removeEventListener('pointercancel', up)
      onDone?.()
    }
    target.addEventListener('pointermove', move)
    target.addEventListener('pointerup', up)
    target.addEventListener('pointercancel', up)
  }
  return (
    <label className="ae-opt" title={title ?? label}>
      <span className="ae-opt-label" onPointerDown={scrub}>{label}</span>
      <input
        className="ae-opt-num"
        type="text"
        inputMode="decimal"
        value={text}
        disabled={disabled}
        onFocus={(event) => {
          editing.current = true
          event.currentTarget.select()
        }}
        onChange={(event) => setText(event.currentTarget.value)}
        onBlur={fromText}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === 'Escape') {
            if (event.key === 'Escape') setText(format(value))
            else fromText()
            event.currentTarget.blur()
            onDone?.()
          } else if (event.key === 'ArrowUp' || event.key === 'ArrowDown') {
            event.preventDefault()
            const delta = (event.key === 'ArrowUp' ? step : -step) * (event.shiftKey ? 10 : 1)
            const next = clampDisplay(value * scale + delta)
            setText(next.toFixed(digits).replace(/\.0+$/, ''))
            commit(next)
          }
        }}
      />
      {unit && <span className="ae-opt-unit">{unit}</span>}
    </label>
  )
}

function CheckOption({ label, title, checked, disabled, onChange }: { readonly label: string; readonly title?: string; readonly checked: boolean; readonly disabled?: boolean; readonly onChange: (checked: boolean) => void }) {
  return (
    <label className="ae-opt ae-opt-check" title={title ?? label}>
      <input type="checkbox" checked={checked} disabled={disabled} onChange={(event) => onChange(event.currentTarget.checked)} />
      <span>{label}</span>
    </label>
  )
}

function SelectOption<T extends string | number>({ label, title, value, options, disabled, onChange }: {
  readonly label?: string
  readonly title?: string
  readonly value: T
  readonly options: readonly (readonly [T, string] | '-')[]
  readonly disabled?: boolean
  readonly onChange: (value: T) => void
}) {
  return (
    <label className="ae-opt" title={title ?? label}>
      {label && <span className="ae-opt-label is-static">{label}</span>}
      <select
        className="ae-opt-select"
        value={String(value)}
        disabled={disabled}
        aria-label={label ?? title}
        onChange={(event) => {
          const raw = event.currentTarget.value
          const match = options.find((option) => option !== '-' && String(option[0]) === raw)
          if (match && match !== '-') onChange(match[0])
        }}
      >
        {options.map((option, index) => option === '-'
          ? <option key={`sep-${index}`} disabled>──────────</option>
          : <option key={String(option[0])} value={String(option[0])}>{option[1]}</option>)}
      </select>
    </label>
  )
}

function ColorOption({ label, value, compact, disabled, onChange }: { readonly label: string; readonly value: Rgb8; readonly compact?: boolean; readonly disabled?: boolean; readonly onChange: (value: Rgb8) => void }) {
  return (
    <label className="ae-opt ae-opt-color" title={label}>
      {!compact && <span className="ae-opt-label is-static">{label}</span>}
      <span className="ae-opt-swatch" style={{ background: css(value) }}>
        <input type="color" aria-label={label} value={toHex(value)} disabled={disabled} onChange={(event) => { const color = parseHex(event.currentTarget.value); if (color) onChange(color) }} />
      </span>
    </label>
  )
}

const BLEND_OPTIONS: readonly (readonly [BlendMode, string] | '-')[] = BLEND_MODE_MENU.map((mode) => (mode === '-' ? '-' : [mode, blendModeLabel(mode)] as const))
const SAMPLE_OPTIONS: readonly (readonly [SampleSource, string])[] = [['current', 'Current Layer'], ['current-below', 'Current & Below'], ['all', 'All Layers']]
const SELECTION_OPS: readonly (readonly [SelectionOp, string, string])[] = [
  ['replace', 'New', 'New selection'],
  ['add', 'Add', 'Add to selection (Shift)'],
  ['subtract', 'Sub', 'Subtract from selection (Alt)'],
  ['intersect', 'Int', 'Intersect with selection (Shift+Alt)'],
]

function SelectionOpButtons({ value, disabled, onChange }: { readonly value: SelectionOp; readonly disabled?: boolean; readonly onChange: (op: SelectionOp) => void }) {
  return (
    <div className="ae-segmented" role="group" aria-label="Selection mode">
      {SELECTION_OPS.map(([op, short, title]) => (
        <button key={op} type="button" title={title} aria-pressed={value === op} className={value === op ? 'is-active' : ''} disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => onChange(op)}>{short}</button>
      ))}
    </div>
  )
}

function useTransformInfo(session: EditorSession) {
  useCell(session.optionsTick)
  const controller = session.transform()
  return controller ? { controller, info: controller.info() } : null
}

function TransformOptions({ session, suspended }: { readonly session: EditorSession; readonly suspended: boolean }) {
  const state = useTransformInfo(session)
  if (!state || !state.info) return null
  const { controller, info } = state
  const set = (patch: Partial<typeof info>) => {
    controller.setInfo(patch)
    session.view.requestOverlay()
  }
  return (
    <div className="ae-options-group" aria-label="Free transform">
      <span className="ae-opt-title">Free Transform</span>
      <NumberOption label="X" value={info.x} min={-1e6} max={1e6} step={1} digits={1} unit="px" disabled={suspended} onChange={(x) => set({ x })} />
      <NumberOption label="Y" value={info.y} min={-1e6} max={1e6} step={1} digits={1} unit="px" disabled={suspended} onChange={(y) => set({ y })} />
      <NumberOption label="W" value={info.width} min={0.01} max={1e6} step={1} digits={1} unit="px" disabled={suspended} onChange={(width) => set({ width })} />
      <NumberOption label="H" value={info.height} min={0.01} max={1e6} step={1} digits={1} unit="px" disabled={suspended} onChange={(height) => set({ height })} />
      <NumberOption label="∠" title="Angle" value={info.angle} min={-360} max={360} step={0.1} digits={1} unit="°" disabled={suspended} onChange={(angle) => set({ angle })} />
      <button type="button" className="ae-opt-button" title="Cancel transform (Esc)" aria-label="Cancel transform" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('session.cancel')}><X aria-hidden="true" /></button>
      <button type="button" className="ae-opt-button is-primary" title="Commit transform (Enter)" aria-label="Commit transform" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('session.commit')}><Check aria-hidden="true" /></button>
    </div>
  )
}

function BrushOptions({ session, tool, options, disabled }: {
  readonly session: EditorSession
  readonly tool: 'brush' | 'eraser' | 'clone'
  readonly options: ToolOptions['brush']
  readonly disabled: boolean
}) {
  const update = (patch: Partial<ToolOptions['brush']>) => session.editor.updateOptions(tool, patch)
  const done = () => session.surface.focus()
  return (
    <>
      <NumberOption label="Size" title="Brush size ([ and ])" value={options.size} min={1} max={5000} unit="px" disabled={disabled} onChange={(size) => update({ size })} onDone={done} />
      <NumberOption label="Hardness" title="Hardness (Shift+[ and Shift+])" value={options.hardness} min={0} max={1} scale={100} unit="%" disabled={disabled} onChange={(hardness) => update({ hardness })} onDone={done} />
      {tool === 'brush' && <SelectOption title="Blend mode" value={options.blendMode} options={BLEND_OPTIONS} disabled={disabled} onChange={(blendMode) => update({ blendMode })} />}
      <NumberOption label="Opacity" title="Opacity (number keys)" value={options.opacity} min={0} max={1} scale={100} unit="%" disabled={disabled} onChange={(opacity) => update({ opacity })} onDone={done} />
      <NumberOption label="Flow" title="Flow (Shift+number keys)" value={options.flow} min={0.01} max={1} scale={100} unit="%" disabled={disabled} onChange={(flow) => update({ flow })} onDone={done} />
      <NumberOption label="Smoothing" value={options.smoothing} min={0} max={1} scale={100} unit="%" disabled={disabled} onChange={(smoothing) => update({ smoothing })} onDone={done} />
      <CheckOption label="Pen size" title="Pen pressure controls the size" checked={options.pressureSize} disabled={disabled} onChange={(pressureSize) => update({ pressureSize })} />
      <CheckOption label="Pen opacity" title="Pen pressure controls the flow" checked={options.pressureOpacity} disabled={disabled} onChange={(pressureOpacity) => update({ pressureOpacity })} />
    </>
  )
}

const SHAPE_OPTIONS: readonly (readonly [ShapeKind, string])[] = [['rectangle', 'Rectangle'], ['ellipse', 'Ellipse'], ['line', 'Line'], ['arrow', 'Arrow']]

function ToolOptionsBar({ session, state, suspended }: { readonly session: EditorSession; readonly state: EditorState; readonly suspended: boolean }) {
  const transform = useTransformInfo(session)
  const disabled = suspended
  const options = state.options
  const done = () => session.surface.focus()
  const tool = state.tool
  if (transform) return <TransformOptions session={session} suspended={suspended} />
  const label = <span className="ae-opt-title"><ToolIcon tool={tool} /><span>{TOOL_META[tool].label.replace(/ Tool$/, '')}</span></span>
  let body: ReactNode = null
  switch (tool) {
    case 'brush':
    case 'eraser':
      body = <BrushOptions session={session} tool={tool} options={options[tool]} disabled={disabled} />
      break
    case 'clone-stamp':
      body = (
        <>
          <BrushOptions session={session} tool="clone" options={options.clone} disabled={disabled} />
          <CheckOption label="Aligned" checked={options.clone.aligned} disabled={disabled} onChange={(aligned) => session.editor.updateOptions('clone', { aligned })} />
          <SelectOption label="Sample" value={options.clone.sample} options={SAMPLE_OPTIONS} disabled={disabled} onChange={(sample) => session.editor.updateOptions('clone', { sample })} />
        </>
      )
      break
    case 'spot-healing':
      body = (
        <>
          <NumberOption label="Size" value={options.heal.size} min={1} max={5000} unit="px" disabled={disabled} onChange={(size) => session.editor.updateOptions('heal', { size })} onDone={done} />
          <NumberOption label="Hardness" value={options.heal.hardness} min={0} max={1} scale={100} unit="%" disabled={disabled} onChange={(hardness) => session.editor.updateOptions('heal', { hardness })} onDone={done} />
          <SelectOption label="Sample" value={options.heal.sample} options={SAMPLE_OPTIONS} disabled={disabled} onChange={(sample) => session.editor.updateOptions('heal', { sample })} />
        </>
      )
      break
    case 'marquee-rect':
    case 'marquee-ellipse':
      body = (
        <>
          <SelectionOpButtons value={options.marquee.op} disabled={disabled} onChange={(op) => session.editor.updateOptions('marquee', { op })} />
          <NumberOption label="Feather" value={options.marquee.feather} min={0} max={1000} step={1} digits={1} unit="px" disabled={disabled} onChange={(feather) => session.editor.updateOptions('marquee', { feather })} onDone={done} />
          <CheckOption label="Anti-alias" checked={options.marquee.antiAlias} disabled={disabled} onChange={(antiAlias) => session.editor.updateOptions('marquee', { antiAlias })} />
          <SelectOption label="Style" value={options.marquee.style} options={[['normal', 'Normal'], ['fixed-ratio', 'Fixed Ratio'], ['fixed-size', 'Fixed Size']]} disabled={disabled} onChange={(style) => session.editor.updateOptions('marquee', { style })} />
          {options.marquee.style === 'fixed-ratio' && (
            <>
              <NumberOption label="W" value={options.marquee.ratio.width} min={0.001} max={1000} step={0.1} digits={2} disabled={disabled} onChange={(width) => session.editor.updateOptions('marquee', { ratio: { ...options.marquee.ratio, width } })} onDone={done} />
              <NumberOption label="H" value={options.marquee.ratio.height} min={0.001} max={1000} step={0.1} digits={2} disabled={disabled} onChange={(height) => session.editor.updateOptions('marquee', { ratio: { ...options.marquee.ratio, height } })} onDone={done} />
            </>
          )}
          {options.marquee.style === 'fixed-size' && (
            <>
              <NumberOption label="W" value={options.marquee.fixedSize.width} min={1} max={LIMITS.maxDimension} unit="px" disabled={disabled} onChange={(width) => session.editor.updateOptions('marquee', { fixedSize: { ...options.marquee.fixedSize, width } })} onDone={done} />
              <NumberOption label="H" value={options.marquee.fixedSize.height} min={1} max={LIMITS.maxDimension} unit="px" disabled={disabled} onChange={(height) => session.editor.updateOptions('marquee', { fixedSize: { ...options.marquee.fixedSize, height } })} onDone={done} />
            </>
          )}
        </>
      )
      break
    case 'lasso':
    case 'lasso-polygon':
      body = (
        <>
          <SelectionOpButtons value={options.lasso.op} disabled={disabled} onChange={(op) => session.editor.updateOptions('lasso', { op })} />
          <NumberOption label="Feather" value={options.lasso.feather} min={0} max={1000} step={1} digits={1} unit="px" disabled={disabled} onChange={(feather) => session.editor.updateOptions('lasso', { feather })} onDone={done} />
          <CheckOption label="Anti-alias" checked={options.lasso.antiAlias} disabled={disabled} onChange={(antiAlias) => session.editor.updateOptions('lasso', { antiAlias })} />
        </>
      )
      break
    case 'magic-wand':
      body = (
        <>
          <SelectionOpButtons value={options.wand.op} disabled={disabled} onChange={(op) => session.editor.updateOptions('wand', { op })} />
          <NumberOption label="Tolerance" value={options.wand.tolerance} min={0} max={255} disabled={disabled} onChange={(tolerance) => session.editor.updateOptions('wand', { tolerance })} onDone={done} />
          <CheckOption label="Anti-alias" checked={options.wand.antiAlias} disabled={disabled} onChange={(antiAlias) => session.editor.updateOptions('wand', { antiAlias })} />
          <CheckOption label="Contiguous" checked={options.wand.contiguous} disabled={disabled} onChange={(contiguous) => session.editor.updateOptions('wand', { contiguous })} />
          <SelectOption label="Sample" value={options.wand.sample} options={SAMPLE_OPTIONS} disabled={disabled} onChange={(sample) => session.editor.updateOptions('wand', { sample })} />
        </>
      )
      break
    case 'crop': {
      const sessionOpen = Boolean(session.activeTool()?.hasSession())
      body = (
        <>
          <SelectOption label="Ratio" value={options.crop.aspect} options={[['free', 'Free'], ['original', 'Original'], ['1:1', '1 : 1'], ['4:3', '4 : 3'], ['3:2', '3 : 2'], ['16:9', '16 : 9'], ['5:4', '5 : 4'], ['7:5', '7 : 5']]} disabled={disabled} onChange={(aspect) => session.editor.updateOptions('crop', { aspect })} />
          <button type="button" className="ae-opt-button" title="Swap width and height" aria-pressed={options.crop.portrait} disabled={disabled || options.crop.aspect === 'free'} onMouseDown={(event) => event.preventDefault()} onClick={() => session.editor.updateOptions('crop', { portrait: !options.crop.portrait })}><ArrowLeftRight aria-hidden="true" /></button>
          <SelectOption label="Overlay" value={options.crop.overlay} options={[['thirds', 'Rule of Thirds'], ['grid', 'Grid'], ['none', 'None']]} disabled={disabled} onChange={(overlay) => session.editor.updateOptions('crop', { overlay })} />
          <CheckOption label="Delete Cropped Pixels" checked={options.crop.deleteCroppedPixels} disabled={disabled} onChange={(deleteCroppedPixels) => session.editor.updateOptions('crop', { deleteCroppedPixels })} />
          {sessionOpen && (
            <>
              <button type="button" className="ae-opt-button" title="Cancel crop (Esc)" aria-label="Cancel crop" disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('session.cancel')}><X aria-hidden="true" /></button>
              <button type="button" className="ae-opt-button is-primary" title="Crop (Enter)" aria-label="Commit crop" disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('session.commit')}><Check aria-hidden="true" /></button>
            </>
          )}
        </>
      )
      break
    }
    case 'eyedropper':
      body = (
        <>
          <SelectOption label="Sample Size" value={options.eyedropper.size} options={[[1, 'Point Sample'], [3, '3 by 3 Average'], [5, '5 by 5 Average'], [11, '11 by 11 Average']]} disabled={disabled} onChange={(size) => session.editor.updateOptions('eyedropper', { size })} />
          <SelectOption label="Sample" value={options.eyedropper.sample} options={SAMPLE_OPTIONS} disabled={disabled} onChange={(sample) => session.editor.updateOptions('eyedropper', { sample })} />
        </>
      )
      break
    case 'gradient':
      body = (
        <>
          <SelectOption label="Gradient" value={options.gradient.preset} options={[['foreground-background', 'Foreground to Background'], ['foreground-transparent', 'Foreground to Transparent'], ['black-white', 'Black, White']]} disabled={disabled} onChange={(preset) => session.editor.updateOptions('gradient', { preset })} />
          <SelectOption label="Type" value={options.gradient.kind} options={[['linear', 'Linear'], ['radial', 'Radial'], ['angle', 'Angle'], ['reflected', 'Reflected'], ['diamond', 'Diamond']]} disabled={disabled} onChange={(kind) => session.editor.updateOptions('gradient', { kind })} />
          <SelectOption title="Blend mode" value={options.gradient.blendMode} options={BLEND_OPTIONS} disabled={disabled} onChange={(blendMode) => session.editor.updateOptions('gradient', { blendMode })} />
          <NumberOption label="Opacity" value={options.gradient.opacity} min={0} max={1} scale={100} unit="%" disabled={disabled} onChange={(opacity) => session.editor.updateOptions('gradient', { opacity })} onDone={done} />
          <CheckOption label="Reverse" checked={options.gradient.reverse} disabled={disabled} onChange={(reverse) => session.editor.updateOptions('gradient', { reverse })} />
          <CheckOption label="Dither" checked={options.gradient.dither} disabled={disabled} onChange={(dither) => session.editor.updateOptions('gradient', { dither })} />
        </>
      )
      break
    case 'paint-bucket':
      body = (
        <>
          <SelectOption title="Blend mode" value={options.bucket.blendMode} options={BLEND_OPTIONS} disabled={disabled} onChange={(blendMode) => session.editor.updateOptions('bucket', { blendMode })} />
          <NumberOption label="Opacity" value={options.bucket.opacity} min={0} max={1} scale={100} unit="%" disabled={disabled} onChange={(opacity) => session.editor.updateOptions('bucket', { opacity })} onDone={done} />
          <NumberOption label="Tolerance" value={options.bucket.tolerance} min={0} max={255} disabled={disabled} onChange={(tolerance) => session.editor.updateOptions('bucket', { tolerance })} onDone={done} />
          <CheckOption label="Anti-alias" checked={options.bucket.antiAlias} disabled={disabled} onChange={(antiAlias) => session.editor.updateOptions('bucket', { antiAlias })} />
          <CheckOption label="Contiguous" checked={options.bucket.contiguous} disabled={disabled} onChange={(contiguous) => session.editor.updateOptions('bucket', { contiguous })} />
          <SelectOption label="Sample" value={options.bucket.sample} options={SAMPLE_OPTIONS} disabled={disabled} onChange={(sample) => session.editor.updateOptions('bucket', { sample })} />
        </>
      )
      break
    case 'text': {
      const text = options.text
      const families = availableFontFamilies()
      const familyOptions = (families.includes(text.fontFamily) ? families : [text.fontFamily, ...families]).map((family) => [family, family] as const)
      body = (
        <>
          <SelectOption title="Font" value={text.fontFamily} options={familyOptions} disabled={disabled} onChange={(fontFamily) => session.editor.updateOptions('text', { fontFamily })} />
          <NumberOption label="Size" value={text.fontSize} min={1} max={5000} digits={1} unit="px" disabled={disabled} onChange={(fontSize) => session.editor.updateOptions('text', { fontSize })} onDone={done} />
          <div className="ae-segmented" role="group" aria-label="Style">
            <button type="button" title="Bold" aria-label="Bold" aria-pressed={text.fontWeight === 700} className={text.fontWeight === 700 ? 'is-active' : ''} disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.editor.updateOptions('text', { fontWeight: text.fontWeight === 700 ? 400 : 700 })}><Bold aria-hidden="true" /></button>
            <button type="button" title="Italic" aria-label="Italic" aria-pressed={text.italic} className={text.italic ? 'is-active' : ''} disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.editor.updateOptions('text', { italic: !text.italic })}><Italic aria-hidden="true" /></button>
            <button type="button" title="Underline" aria-label="Underline" aria-pressed={text.underline} className={text.underline ? 'is-active' : ''} disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.editor.updateOptions('text', { underline: !text.underline })}><Underline aria-hidden="true" /></button>
          </div>
          <div className="ae-segmented" role="group" aria-label="Alignment">
            {([['left', AlignLeft], ['center', AlignCenter], ['right', AlignRight]] as const).map(([align, Icon]) => (
              <button key={align} type="button" title={`Align ${align}`} aria-label={`Align ${align}`} aria-pressed={text.align === align} className={text.align === align ? 'is-active' : ''} disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.editor.updateOptions('text', { align })}><Icon aria-hidden="true" /></button>
            ))}
          </div>
          <ColorOption label="Color" value={text.color} disabled={disabled} onChange={(color) => session.editor.updateOptions('text', { color })} />
        </>
      )
      break
    }
    case 'shape': {
      const shape = options.shape
      const fill: Rgb8 = shape.fill ?? state.foreground
      const stroke: Rgb8 = shape.stroke ?? state.foreground
      body = (
        <>
          <SelectOption title="Shape" value={shape.kind} options={SHAPE_OPTIONS} disabled={disabled} onChange={(kind) => session.editor.updateOptions('shape', { kind })} />
          <CheckOption label="Fill" checked={shape.fill !== null} disabled={disabled} onChange={(on) => session.editor.updateOptions('shape', { fill: on ? { ...fill, a: 255 } : null })} />
          {shape.fill && <ColorOption label="Fill colour" compact value={shape.fill} disabled={disabled} onChange={(color) => session.editor.updateOptions('shape', { fill: { ...color, a: 255 } })} />}
          <CheckOption label="Stroke" checked={shape.stroke !== null} disabled={disabled} onChange={(on) => session.editor.updateOptions('shape', { stroke: on ? { ...stroke, a: 255 } : null })} />
          {shape.stroke && <ColorOption label="Stroke colour" compact value={shape.stroke} disabled={disabled} onChange={(color) => session.editor.updateOptions('shape', { stroke: { ...color, a: 255 } })} />}
          <NumberOption label="Width" value={shape.strokeWidth} min={0} max={1000} unit="px" disabled={disabled} onChange={(strokeWidth) => session.editor.updateOptions('shape', { strokeWidth })} onDone={done} />
          {shape.kind === 'rectangle' && <NumberOption label="Radius" value={shape.cornerRadius} min={0} max={10000} unit="px" disabled={disabled} onChange={(cornerRadius) => session.editor.updateOptions('shape', { cornerRadius })} onDone={done} />}
        </>
      )
      break
    }
    case 'move':
      body = (
        <>
          <CheckOption label="Auto-Select" title="Click selects the topmost layer under the pointer (Ctrl-click does it once)" checked={options.move.autoSelect} disabled={disabled} onChange={(autoSelect) => session.editor.updateOptions('move', { autoSelect })} />
          <CheckOption label="Show Transform Controls" checked={options.move.showTransformControls} disabled={disabled} onChange={(showTransformControls) => session.editor.updateOptions('move', { showTransformControls })} />
        </>
      )
      break
    case 'hand':
    case 'zoom':
      body = (
        <>
          <button type="button" className="ae-opt-text-button" disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('view.actual-pixels')} title={withShortcut('100%', 'view.actual-pixels')}>100%</button>
          <button type="button" className="ae-opt-text-button" disabled={disabled} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('view.fit')} title={withShortcut('Fit on Screen', 'view.fit')}>Fit Screen</button>
        </>
      )
      break
  }
  return <div className="ae-options-group" aria-label={`${TOOL_META[tool].label} options`}>{label}{body}</div>
}

// ---------------------------------------------------------------------------------------------
// Status bar
// ---------------------------------------------------------------------------------------------

function StatusBar({ session }: { readonly session: EditorSession }) {
  const hint = useCell(session.hint)
  const hover = useCell(session.hover)
  const busy = useCell(session.busy)
  useSyncExternalStore(useCallback((listener: () => void) => session.view.subscribe(listener), [session]), () => session.view.getView())
  const size = useSyncExternalStore(
    useCallback((listener: () => void) => session.store.subscribe(() => listener()), [session]),
    () => session.store.getState().width * 100_000 + session.store.getState().height,
  )
  const state = session.store.getState()
  const view = session.view.getView()
  const dpr = session.view.getViewportSize().dpr
  const [memory, setMemory] = useState(() => session.store.memoryUsage().total)
  useEffect(() => {
    const timer = setInterval(() => setMemory(session.store.memoryUsage().total), 1500)
    return () => clearInterval(timer)
  }, [session])
  void size
  return (
    <footer className="ae-status" aria-label="Status">
      <span className="ae-status-zoom" title="Zoom (Ctrl++ / Ctrl+- / Ctrl+0 / Ctrl+1)">{zoomLabel(view.zoom, dpr)}</span>
      <span>{state.width.toLocaleString()} × {state.height.toLocaleString()} px · {state.ppi} ppi</span>
      <span className="ae-status-cursor">{hover ? `x ${hover.x}  y ${hover.y}` : ''}</span>
      <span className="ae-status-color">
        {hover?.color && hover.color.a > 0 ? (
          <>
            <i style={{ background: css(hover.color) }} aria-hidden="true" />
            {`R ${hover.color.r}  G ${hover.color.g}  B ${hover.color.b}${hover.color.a < 255 ? `  A ${Math.round((hover.color.a / 255) * 100)}%` : ''}`}
          </>
        ) : ''}
      </span>
      <span title="Memory used by the document and its history">{formatBytes(memory)}</span>
      <span className="ae-status-hint" aria-live="polite">{busy ? <><LoaderCircle className="ae-spin" aria-hidden="true" /> Working…</> : hint ?? ''}</span>
    </footer>
  )
}

// ---------------------------------------------------------------------------------------------
// Shell
// ---------------------------------------------------------------------------------------------

const NARROW = '(max-width: 1079px)'

function useNarrow(): boolean {
  const [narrow, setNarrow] = useState(() => typeof window !== 'undefined' && window.matchMedia(NARROW).matches)
  useEffect(() => {
    const query = window.matchMedia(NARROW)
    const update = () => setNarrow(query.matches)
    update()
    query.addEventListener('change', update)
    return () => query.removeEventListener('change', update)
  }, [])
  return narrow
}

function HistoryButtons({ session, suspended }: { readonly session: EditorSession; readonly suspended: boolean }) {
  const history = useHistoryState(session.store)
  const canUndo = history.cursor > 0 || Boolean(session.transform())
  const canRedo = history.cursor < history.entries.length - 1
  return (
    <div className="ae-history-buttons">
      <button type="button" className="ae-icon-button" title={withShortcut('Undo', 'edit.undo')} aria-label="Undo" disabled={suspended || !canUndo} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('edit.undo')}><Undo2 aria-hidden="true" /></button>
      <button type="button" className="ae-icon-button" title={withShortcut('Redo', 'edit.redo')} aria-label="Redo" disabled={suspended || !canRedo} onMouseDown={(event) => event.preventDefault()} onClick={() => session.run('edit.redo')}><Redo2 aria-hidden="true" /></button>
    </div>
  )
}

function EditorShell({ session, host, suspended }: { readonly session: EditorSession; readonly host: AdvancedHost; readonly suspended: boolean }) {
  const state = useEditorState(session.editor)
  const busy = useCell(session.busy)
  const narrow = useNarrow()
  const [dockOpen, setDockOpen] = useState(false)
  const anyPanel = state.panels.layers || state.panels.properties || state.panels.history || state.panels.color

  useEffect(() => {
    if (!narrow) setDockOpen(false)
  }, [narrow])

  return (
    <div className={`ae-shell ${anyPanel && !narrow ? 'has-dock' : ''}`} aria-busy={busy || undefined}>
      <header className="ae-options" role="toolbar" aria-label="Options">
        <MenuBar
          run={session.run}
          isEnabled={session.isEnabled}
          label={(command) => commandLabel(command, session.context)}
          shortcut={(command) => shortcutLabel(command)}
          isChecked={(command) => {
            const panels = session.editor.getState().panels
            if (command === 'view.panel-layers') return panels.layers
            if (command === 'view.panel-properties') return panels.properties
            if (command === 'view.panel-history') return panels.history
            if (command === 'view.panel-color') return panels.color
            return false
          }}
          disabled={suspended}
          onDismiss={() => session.surface.focus()}
        />
        <span className="ae-rule" aria-hidden="true" />
        <ToolOptionsBar session={session} state={state} suspended={suspended} />
        <span className="ae-spacer" />
        <HistoryButtons session={session} suspended={suspended} />
        {narrow && anyPanel && (
          <button type="button" className={`ae-icon-button ${dockOpen ? 'is-active' : ''}`} title="Panels" aria-label="Show panels" aria-pressed={dockOpen} disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => setDockOpen((open) => !open)}><PanelRight aria-hidden="true" /></button>
        )}
        <span className="ae-rule" aria-hidden="true" />
        <button type="button" className="ae-simple" title="Back to Simple mode (layers are flattened; you are asked first when that would lose anything)" disabled={suspended} onMouseDown={(event) => event.preventDefault()} onClick={() => host.requestExit()}>
          <ArrowLeftToLine aria-hidden="true" /><span>Simple</span>
        </button>
      </header>
      <Toolbox session={session} state={state} suspended={suspended} />
      <main className="ae-canvas-area">
        <CanvasView store={session.store} compositor={session.compositor} view={session.view} surface={session.surface} input={session.input} suspended={suspended} busy={busy} />
      </main>
      {anyPanel && (!narrow || dockOpen) && (
        <PanelHost context={session.panelContext} suspended={suspended} overlay={narrow} onCloseOverlay={() => setDockOpen(false)} />
      )}
      <StatusBar session={session} />
      {/* Parameter dialogs cover the whole editor (the canvas stays visible for live previews). */}
      <DialogHost controller={session.dialogs} context={session.panelContext} suspended={suspended} />
    </div>
  )
}

// ---------------------------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------------------------

export function AdvancedEditor({ host, initial, suspended, onReady }: AdvancedEditorProps) {
  const [session, setSession] = useState<EditorSession | null>(null)
  const [error, setError] = useState<string | null>(null)
  const suspendedRef = useRef(suspended)
  suspendedRef.current = suspended

  // Read the document once (after a first paint, so the "Reading" message shows on big images).
  useEffect(() => {
    let cancelled = false
    let created: EditorSession | null = null
    const timer = setTimeout(() => {
      if (cancelled) return
      try {
        const document = readInitial(initial)
        created = createEditorSession(host, document, () => suspendedRef.current || host.isSuspended())
        setSession(created)
      } catch (failure) {
        console.error(failure)
        setError(messageOf(failure, 'The image could not be opened in the Advanced editor.'))
      }
    }, 0)
    return () => {
      cancelled = true
      clearTimeout(timer)
      created?.dispose()
      setSession(null)
    }
  }, [host, initial])

  useEffect(() => {
    if (session) onReady?.(session.handle)
  }, [onReady, session])

  if (error) {
    return (
      <div className="ae-shell ae-message-shell" role="alert">
        <div className="ae-message">
          <TriangleAlert aria-hidden="true" />
          <strong>The image could not be opened here</strong>
          <span>{error}</span>
          <button type="button" onClick={() => host.requestExit()}>Back to Simple</button>
        </div>
      </div>
    )
  }
  if (!session) {
    return (
      <div className="ae-shell ae-message-shell" aria-busy="true">
        <div className="ae-message" aria-live="polite"><LoaderCircle className="ae-spin" aria-hidden="true" /><span>Reading the image…</span></div>
      </div>
    )
  }
  return <EditorShell session={session} host={host} suspended={suspended} />
}

export default AdvancedEditor

