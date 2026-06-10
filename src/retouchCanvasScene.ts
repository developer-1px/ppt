import {
  createCanvasSceneAdapter,
} from 'canvas/foundation'
import type { Pointer } from 'zod-crud'
import type { Rect } from './retouchModel'

export function retouchCanvasSelectionBounds(
  items: readonly { pointer: Pointer; rect: Rect }[],
  pointers: readonly Pointer[],
): Rect | null {
  if (pointers.length === 0) {
    return null
  }

  const scene = createCanvasSceneAdapter(
    items.map((item, index) => ({
      bounds: {
        h: item.rect.height,
        w: item.rect.width,
        x: item.rect.x,
        y: item.rect.y,
      },
      canResize: true,
      id: item.pointer,
      isGroup: false,
      parentId: null,
      path: [index],
    })),
  )
  const bounds = scene.getBounds([...pointers])

  return bounds === null
    ? null
    : {
        height: bounds.h,
        width: bounds.w,
        x: bounds.x,
        y: bounds.y,
      }
}
