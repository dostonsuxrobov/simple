// src/advanced/tools/index.ts
// Owner: WP5 (seeded by WP1). Tool registry of the Advanced editor (design 5.9, 8.4).
//   - TOOL_FACTORIES: one controller factory per ToolId (a fresh controller per call).
//   - TOOL_GROUPS: toolbar order; tools sharing a slot cycle with Shift + their key.
//   - TOOL_META: Photoshop label, shortcut letter and lucide-react icon name per tool.
//   - createFreeTransform(): Free Transform (Ctrl+T) is not a ToolId; the editor creates it on demand and
//     routes pointer, key and overlay calls to it while hasSession() is true (see tools/transform.ts).
// Every controller follows the ToolController contract plus the optional ToolControllerExtras
// (isBusy / whenIdle for asynchronous commits; the editor's settle() awaits whenIdle()).
// DOM-free at module level so Node tests can load it.
import type { ToolController, ToolFactory, ToolId } from '../types.ts'
import { createBrushTool } from './brush.ts'
import { createBucketTool } from './bucket.ts'
import { createCloneTool } from './clone.ts'
import { createCropTool } from './crop.ts'
import { createEyedropperTool } from './eyedropper.ts'
import { createGradientTool } from './gradient.ts'
import { createHandTool } from './hand.ts'
import { createHealTool } from './heal.ts'
import { createLassoTool } from './lasso.ts'
import { createMarqueeTool } from './marquee.ts'
import { createMoveTool } from './move.ts'
import { createShapeTool } from './shape.ts'
import { createTextTool } from './text.ts'
import { createWandTool } from './wand.ts'
import { createZoomTool } from './zoom.ts'

import { createFreeTransform as createTransformController } from './transform.ts'
import type { FreeTransformController, FreeTransformOptions } from './transform.ts'
import { withCssPixelOverlay } from './shared.ts'

export type { FreeTransformController, FreeTransformOptions } from './transform.ts'
export type { AdvancedToolController, ToolControllerExtras } from './shared.ts'
export { nextSelectionVersion } from './shared.ts'

/**
 * Free Transform (Ctrl+T): activate(ctx) starts a session on the active layer or its selected pixels
 * (check hasSession() afterwards: false means the user was told why not); route pointer, key and overlay
 * calls to it while hasSession() is true; onEnd fires once when the session closes.
 */
export function createFreeTransform(options?: FreeTransformOptions): FreeTransformController {
  return withCssPixelOverlay(createTransformController(options))
}

/** label = Photoshop tool name, key = shortcut letter, icon = lucide-react component name. */
export const TOOL_META: Readonly<Record<ToolId, { readonly label: string; readonly key: string; readonly icon: string }>> = Object.freeze({
  'move': { label: 'Move Tool', key: 'V', icon: 'Move' },
  'marquee-rect': { label: 'Rectangular Marquee Tool', key: 'M', icon: 'SquareDashed' },
  'marquee-ellipse': { label: 'Elliptical Marquee Tool', key: 'M', icon: 'CircleDashed' },
  'lasso': { label: 'Lasso Tool', key: 'L', icon: 'Lasso' },
  'lasso-polygon': { label: 'Polygonal Lasso Tool', key: 'L', icon: 'LassoSelect' },
  'magic-wand': { label: 'Magic Wand Tool', key: 'W', icon: 'WandSparkles' },
  'crop': { label: 'Crop Tool', key: 'C', icon: 'Crop' },
  'eyedropper': { label: 'Eyedropper Tool', key: 'I', icon: 'Pipette' },
  'spot-healing': { label: 'Spot Healing Brush Tool', key: 'J', icon: 'Bandage' },
  'brush': { label: 'Brush Tool', key: 'B', icon: 'Brush' },
  'clone-stamp': { label: 'Clone Stamp Tool', key: 'S', icon: 'Stamp' },
  'eraser': { label: 'Eraser Tool', key: 'E', icon: 'Eraser' },
  'gradient': { label: 'Gradient Tool', key: 'G', icon: 'Blend' },
  'paint-bucket': { label: 'Paint Bucket Tool', key: 'G', icon: 'PaintBucket' },
  'text': { label: 'Horizontal Type Tool', key: 'T', icon: 'Type' },
  'shape': { label: 'Shape Tool', key: 'U', icon: 'Shapes' },
  'hand': { label: 'Hand Tool', key: 'H', icon: 'Hand' },
  'zoom': { label: 'Zoom Tool', key: 'Z', icon: 'ZoomIn' },
})

/** Toolbar order; tools that share a slot cycle with Shift + their key (Shift+M, Shift+L, Shift+G). */
export const TOOL_GROUPS: readonly (readonly ToolId[])[] = Object.freeze([
  Object.freeze(['move'] as const),
  Object.freeze(['marquee-rect', 'marquee-ellipse'] as const),
  Object.freeze(['lasso', 'lasso-polygon'] as const),
  Object.freeze(['magic-wand'] as const),
  Object.freeze(['crop'] as const),
  Object.freeze(['eyedropper'] as const),
  Object.freeze(['spot-healing'] as const),
  Object.freeze(['brush'] as const),
  Object.freeze(['clone-stamp'] as const),
  Object.freeze(['eraser'] as const),
  Object.freeze(['gradient', 'paint-bucket'] as const),
  Object.freeze(['text'] as const),
  Object.freeze(['shape'] as const),
  Object.freeze(['hand'] as const),
  Object.freeze(['zoom'] as const),
])

/** A controller that accepts every call and changes nothing (no session, no history). */
export function createInertTool(id: ToolId): ToolController {
  return {
    id,
    activate() {},
    deactivate() {},
    pointerDown() {},
    pointerMove() {},
    pointerUp() {},
    pointerCancel() {},
    keyDown: () => false,
    keyUp: () => false,
    drawOverlay() {},
    hasSession: () => false,
    commitSession() {},
    cancelSession() {},
  }
}

const tool = (create: () => ToolController): ToolFactory => () => withCssPixelOverlay(create())

export const TOOL_FACTORIES: Readonly<Record<ToolId, ToolFactory>> = Object.freeze({
  'move': tool(() => createMoveTool()),
  'marquee-rect': tool(() => createMarqueeTool('rect')),
  'marquee-ellipse': tool(() => createMarqueeTool('ellipse')),
  'lasso': tool(() => createLassoTool('free')),
  'lasso-polygon': tool(() => createLassoTool('polygon')),
  'magic-wand': tool(() => createWandTool()),
  'crop': tool(() => createCropTool()),
  'eyedropper': tool(() => createEyedropperTool()),
  'spot-healing': tool(() => createHealTool()),
  'brush': tool(() => createBrushTool('brush')),
  'clone-stamp': tool(() => createCloneTool()),
  'eraser': tool(() => createBrushTool('eraser')),
  'gradient': tool(() => createGradientTool()),
  'paint-bucket': tool(() => createBucketTool()),
  'text': tool(() => createTextTool()),
  'shape': tool(() => createShapeTool()),
  'hand': tool(() => createHandTool()),
  'zoom': tool(() => createZoomTool()),
})
