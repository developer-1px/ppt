import {
  createCanvasSceneAdapter,
  type Bounds,
  type CanvasSceneEntry,
} from 'canvas/foundation'
import type { Pointer } from 'zod-crud'
import type { Rect } from './retouchModel'
import type { RetouchSurfaceItem } from './retouchObjectSurface'

export type RetouchCanvasSceneEntry = CanvasSceneEntry & {
  blockId: string
  pointer: Pointer
}

export function retouchCanvasSceneEntries(
  items: readonly RetouchSurfaceItem[],
): RetouchCanvasSceneEntry[] {
  return items.map((item, index) => ({
    blockId: item.block.id,
    bounds: rectToCanvasBounds(item.rect),
    canResize: true,
    id: item.block.id,
    isGroup: false,
    parentId: null,
    path: [index],
    pointer: item.pointer,
  }))
}

export function retouchCanvasSelectionBounds(
  entries: readonly RetouchCanvasSceneEntry[],
  pointers: readonly Pointer[],
): Rect | null {
  const scene = createCanvasSceneAdapter([...entries])
  const selectedIds = canvasSelectionIdsFromPointers(entries, pointers)
  const bounds = scene.getBounds(selectedIds)

  return bounds ? canvasBoundsToRect(bounds) : null
}

export function canvasSelectionIdsFromPointers(
  entries: readonly RetouchCanvasSceneEntry[],
  pointers: readonly Pointer[],
) {
  const selectedPointers = new Set(pointers)

  return entries
    .filter((entry) => selectedPointers.has(entry.pointer))
    .map((entry) => entry.id)
}

function rectToCanvasBounds(rect: Rect): Bounds {
  return {
    h: rect.height,
    w: rect.width,
    x: rect.x,
    y: rect.y,
  }
}

function canvasBoundsToRect(bounds: Bounds): Rect {
  return {
    height: bounds.h,
    width: bounds.w,
    x: bounds.x,
    y: bounds.y,
  }
}
