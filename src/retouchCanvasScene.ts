import {
  createCanvasSceneAdapter,
  type Bounds,
  type CanvasSceneEntry,
  type CanvasTransformAdapter,
  type CanvasTransformItem,
} from 'canvas/foundation'
import type { Pointer } from 'zod-crud'
import type { Rect } from './retouchModel'
import type { RetouchSurfaceItem } from './retouchObjectSurface'

export type RetouchCanvasSceneEntry = CanvasSceneEntry & {
  blockId: string
  pointer: Pointer
}

export type RetouchCanvasTransformItem = CanvasTransformItem & {
  pointer: Pointer
  rect: Rect
}

export const retouchCanvasTransformAdapter: CanvasTransformAdapter<RetouchCanvasTransformItem> = {
  resizeSelection: ({ from, items, selection, to }) => {
    const selected = new Set(selection)
    const scaleX = from.w === 0 ? 1 : to.w / from.w
    const scaleY = from.h === 0 ? 1 : to.h / from.h

    return items.map((item) => {
      if (!selected.has(item.id)) {
        return item
      }

      return {
        ...item,
        rect: {
          x: to.x + (item.rect.x - from.x) * scaleX,
          y: to.y + (item.rect.y - from.y) * scaleY,
          width: item.rect.width * scaleX,
          height: item.rect.height * scaleY,
        },
      }
    })
  },

  translateSelection: ({ dx, dy, items, selection }) => {
    const selected = new Set(selection)

    return items.map((item) =>
      selected.has(item.id)
        ? {
            ...item,
            rect: {
              ...item.rect,
              x: item.rect.x + dx,
              y: item.rect.y + dy,
            },
          }
        : item,
    )
  },
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

export function retouchCanvasTransformItems(
  items: readonly RetouchSurfaceItem[],
): RetouchCanvasTransformItem[] {
  return items.map((item) => ({
    id: item.block.id,
    pointer: item.pointer,
    rect: item.rect,
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
