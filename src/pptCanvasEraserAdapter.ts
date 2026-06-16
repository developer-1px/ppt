import {
  getCanvasEraserHitItemIds,
  getCanvasMergedEraserHitIds,
} from 'canvas/app/eraser-hit-testing'
import type {
  Bounds,
  Point,
} from 'canvas/core'
import { unionCanvasRectList } from 'canvas/foundation'
import { pptGeometryToBounds } from './pptCanvasAdapter'
import type {
  PPTElement,
  PPTFreeform,
  PPTSlide,
} from './pptModel'

const PPT_CANVAS_ERASER_RADIUS = 8

type CanvasEraserHitInput = Parameters<typeof getCanvasEraserHitItemIds>[0]
type CanvasEraserItemReadModel = CanvasEraserHitInput['itemReadModel']
type CanvasEraserItem = ReturnType<CanvasEraserItemReadModel['getAllItems']>[number]
type CanvasEraserStrokeItem = Extract<
  CanvasEraserItem,
  { type: 'highlight' | 'marker' }
>

export const mergePPTEraserHitElementIds = getCanvasMergedEraserHitIds

export function getPPTEraserHitElementIds({
  points,
  scene,
  slide,
}: {
  points: Point[]
  scene: CanvasEraserHitInput['scene']
  slide: PPTSlide
}) {
  return getCanvasEraserHitItemIds({
    itemReadModel: createPPTEraserItemReadModel(slide),
    points,
    radius: PPT_CANVAS_ERASER_RADIUS,
    scene,
  })
}

function createPPTEraserItemReadModel(slide: PPTSlide): CanvasEraserItemReadModel {
  const items = slide.elements.flatMap(toCanvasEraserStrokeItem)
  const itemById = new Map(items.map((item) => [item.id, item]))

  return {
    findEditableTextItem: () => null,
    findItem: (id) => itemById.get(id),
    getAllIds: () => items.map((item) => item.id),
    getAllItems: () => items,
    getItemBounds: getCanvasEraserItemBounds,
    getSelection: (ids) => ids.filter((id) => itemById.has(id)),
    getSelectionBounds: (ids) =>
      unionCanvasRectList(Array.from(ids).flatMap((id) => {
        const item = itemById.get(id)

        return item ? [getCanvasEraserItemBounds(item)] : []
      })),
    getSelectedItems: (ids) =>
      ids.flatMap((id) => {
        const item = itemById.get(id)

        return item ? [item] : []
      }),
  }
}

function toCanvasEraserStrokeItem(element: PPTElement): CanvasEraserStrokeItem[] {
  if (
    element.kind !== 'freeform' ||
    element.visible === false ||
    element.locked === true
  ) {
    return []
  }

  return [createCanvasEraserStrokeItem(element)]
}

function createCanvasEraserStrokeItem(
  element: PPTFreeform,
): CanvasEraserStrokeItem {
  const bounds = pptGeometryToBounds(element.geometry)

  return {
    ...bounds,
    id: element.id,
    opacity: element.opacity ?? 1,
    points: getPPTFreeformWorldPoints(element),
    stroke: element.stroke.color,
    strokeWidth: element.stroke.width,
    type: (element.opacity ?? 1) < 1 ? 'highlight' : 'marker',
  }
}

function getCanvasEraserItemBounds(item: CanvasEraserItem): Bounds {
  return {
    h: item.h,
    w: item.w,
    x: item.x,
    y: item.y,
  }
}

function getPPTFreeformWorldPoints(element: PPTFreeform) {
  return element.points.map((point) => ({
    x: element.geometry.x + point.x,
    y: element.geometry.y + point.y,
  }))
}
