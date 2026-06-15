import {
  clamp,
  scaleItemBounds,
  type Bounds,
} from 'canvas/core'
import {
  createCanvasSceneAdapter,
  type CanvasSceneEntry,
  type CanvasTransformAdapter,
} from 'canvas/foundation'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  updatePPTElementGeometry,
  type PPTElement,
  type PPTGeometry,
  type PPTSlide,
} from './pptModel'

export function createPPTCanvasScene(slide: PPTSlide) {
  const entries: CanvasSceneEntry[] = slide.elements
    .filter((element) => element.visible !== false)
    .map((element, index) => ({
      bounds: pptGeometryToBounds(element.geometry),
      canResize: element.locked !== true,
      id: element.id,
      isGroup: false,
      parentId: null,
      path: [index],
    }))

  return createCanvasSceneAdapter(entries)
}

export const pptCanvasTransformAdapter: CanvasTransformAdapter<PPTElement> = {
  resizeSelection({ from, items, selection, to }) {
    const selected = new Set(selection)

    return items.map((item) => {
      if (!selected.has(item.id)) {
        return item
      }

      const bounds = scaleItemBounds(pptGeometryToBounds(item.geometry), from, to)

      return updatePPTElementGeometry(item, boundsToPPTGeometry(bounds))
    })
  },
  translateSelection({ dx, dy, items, selection }) {
    const selected = new Set(selection)

    return items.map((item) => {
      if (!selected.has(item.id)) {
        return item
      }

      return updatePPTElementGeometry(item, {
        ...item.geometry,
        x: clamp(item.geometry.x + dx, 0, PPT_SLIDE_WIDTH - item.geometry.w),
        y: clamp(item.geometry.y + dy, 0, PPT_SLIDE_HEIGHT - item.geometry.h),
      })
    })
  },
}

export function pptGeometryToBounds(geometry: PPTGeometry): Bounds {
  return {
    h: geometry.h,
    w: geometry.w,
    x: geometry.x,
    y: geometry.y,
  }
}

export function boundsToPPTGeometry(bounds: Bounds): PPTGeometry {
  const w = Math.min(PPT_SLIDE_WIDTH, Math.max(24, bounds.w))
  const h = Math.min(PPT_SLIDE_HEIGHT, Math.max(24, bounds.h))

  return {
    h,
    w,
    x: clamp(bounds.x, 0, PPT_SLIDE_WIDTH - w),
    y: clamp(bounds.y, 0, PPT_SLIDE_HEIGHT - h),
  }
}
