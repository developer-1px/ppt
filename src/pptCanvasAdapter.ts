import {
  clampCanvasBoundsToFrame,
  type Bounds,
} from 'canvas/core'
import {
  canSelectSameTypeCanvasItems,
  createCanvasSceneAdapter,
  getCanvasItemGroupIndexRange,
  getCanvasItemGroupMemberIdsForGroup,
  getCanvasItemPointerSelection,
  insertCanvasItemAtTargetPlacement,
  moveCanvasItemToTargetPlacement,
  selectSameTypeCanvasItems,
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

type PPTTargetPlacement = 'after' | 'before'

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

export function getPPTElementPointerSelection({
  additive,
  elementId,
  scene,
  selection,
}: {
  additive: boolean
  elementId: string
  scene: ReturnType<typeof createPPTCanvasScene>
  selection: string[]
}) {
  return getCanvasItemPointerSelection({
    additive,
    itemId: elementId,
    scene,
    selection,
  })
}

export function insertPPTSlideAtTargetPlacement({
  placement,
  slide,
  slides,
  targetSlideId,
}: {
  placement: PPTTargetPlacement
  slide: PPTSlide
  slides: readonly PPTSlide[]
  targetSlideId: string
}) {
  return insertCanvasItemAtTargetPlacement({
    getItemId: getPPTSlideId,
    item: slide,
    items: slides,
    placement,
    targetItemId: targetSlideId,
  })
}

export function movePPTSlideToTargetPlacement({
  placement,
  slideId,
  slides,
  targetSlideId,
}: {
  placement: PPTTargetPlacement
  slideId: string
  slides: readonly PPTSlide[]
  targetSlideId: string
}) {
  return moveCanvasItemToTargetPlacement({
    getItemId: getPPTSlideId,
    itemId: slideId,
    items: slides,
    placement,
    targetItemId: targetSlideId,
  })
}

export function getPPTElementGroupIndexRange({
  elements,
  groupId,
}: {
  elements: readonly PPTElement[]
  groupId: string
}) {
  return getCanvasItemGroupIndexRange({
    getItemGroupId: getPPTElementGroupId,
    groupId,
    items: elements,
  })
}

export function getPPTElementGroupMemberIds({
  elements,
  groupId,
}: {
  elements: readonly PPTElement[]
  groupId: string
}) {
  return getCanvasItemGroupMemberIdsForGroup({
    getItemGroupId: getPPTElementGroupId,
    getItemId: getPPTElementId,
    groupId,
    items: elements,
  })
}

export function selectSameTypePPTElements({
  elements,
  selection,
}: {
  elements: readonly PPTElement[]
  selection: readonly string[]
}) {
  return selectSameTypeCanvasItems({
    getItemId: getPPTElementId,
    getItemType: getPPTElementTypeKey,
    isItemSelectable: isPPTSelectableElement,
    items: elements,
    selection,
  })
}

export function canSelectSameTypePPTElements({
  elements,
  selection,
}: {
  elements: readonly PPTElement[]
  selection: readonly string[]
}) {
  return canSelectSameTypeCanvasItems({
    getItemId: getPPTElementId,
    getItemType: getPPTElementTypeKey,
    isItemSelectable: isPPTSelectableElement,
    items: elements,
    selection,
  })
}

function getPPTSlideId(slide: PPTSlide) {
  return slide.id
}

function getPPTElementId(element: PPTElement) {
  return element.id
}

function getPPTElementGroupId(element: PPTElement) {
  return element.groupId
}

function getPPTElementTypeKey(element: PPTElement) {
  if (element.kind === 'shape') {
    return `shape:${element.shape}`
  }

  return 'Object'
}

function isPPTSelectableElement(element: PPTElement) {
  return element.visible !== false
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
