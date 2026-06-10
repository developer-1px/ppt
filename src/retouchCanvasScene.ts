import {
  createCanvasSceneAdapter,
  type Bounds,
  type CanvasSceneEntry,
} from 'canvas/foundation'
import type { Pointer } from 'zod-crud'
import type { Rect } from './retouchModel'

export function retouchCanvasSceneEntries(
  items: readonly { pointer: Pointer; rect: Rect }[],
): CanvasSceneEntry[] {
  return items.map((item, index) => ({
    bounds: rectToCanvasBounds(item.rect),
    canResize: true,
    id: item.pointer,
    isGroup: false,
    parentId: null,
    path: [index],
  }))
}

export function retouchCanvasSelectionBounds(
  entries: readonly CanvasSceneEntry[],
  pointers: readonly Pointer[],
): Rect | null {
  if (pointers.length === 0) {
    return null
  }

  const scene = createCanvasSceneAdapter([...entries])
  const bounds = scene.getBounds([...pointers])

  return bounds === null ? null : canvasBoundsToRect(bounds)
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
