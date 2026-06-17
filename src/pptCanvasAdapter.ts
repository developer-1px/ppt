import {
  clampCanvasBoundsToFrame,
  type Bounds,
} from 'canvas/core'
import {
  createCanvasSceneAdapter,
  type CanvasSceneEntry,
  type CanvasTransformAdapter,
  resizeCanvasSelectionItems,
  translateCanvasSelectionItems,
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
    return resizeCanvasSelectionItems({
      from,
      getItemBounds: getPPTElementBounds,
      getItemId: getPPTElementId,
      items,
      selection,
      to,
      updateItemBounds: updatePPTElementBounds,
    })
  },
  translateSelection({ dx, dy, items, selection }) {
    return translateCanvasSelectionItems({
      dx,
      dy,
      getItemBounds: getPPTElementBounds,
      getItemId: getPPTElementId,
      items,
      selection,
      updateItemBounds: updatePPTElementBounds,
    })
  },
}

function getPPTElementId(element: PPTElement) {
  return element.id
}

function getPPTElementBounds(element: PPTElement) {
  return pptGeometryToBounds(element.geometry)
}

function updatePPTElementBounds(
  element: PPTElement,
  bounds: Bounds,
) {
  const geometry = boundsToPPTGeometry(bounds)

  return updatePPTElementGeometry(
    element,
    element.geometry.rotation === undefined
      ? geometry
      : { ...geometry, rotation: element.geometry.rotation },
  )
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
  return clampCanvasBoundsToFrame({
    bounds,
    frame: {
      h: PPT_SLIDE_HEIGHT,
      w: PPT_SLIDE_WIDTH,
      x: 0,
      y: 0,
    },
    minHeight: 24,
    minWidth: 24,
  })
}
