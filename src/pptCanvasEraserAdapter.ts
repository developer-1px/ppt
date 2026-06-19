import {
  getPPTCanvasEraserHitItemIds,
  getPPTCanvasMergedEraserHitIds,
  type PPTCanvasEraserHitInput,
} from './pptCanvasAppAffordanceAdapter'
import type {
  Bounds,
  Point,
} from './pptCanvasCoreAdapter'
import { unionPPTCanvasRectList } from './pptCanvasFoundationAdapter'
import { pptGeometryToBounds } from './pptCanvasAdapter'
import type {
  PPTElement,
  PPTFreeform,
  PPTSlide,
} from './pptModel'

const PPT_CANVAS_ERASER_RADIUS = 8

type PPTEraserItemReadModel = PPTCanvasEraserHitInput['itemReadModel']
type PPTEraserItem = ReturnType<PPTEraserItemReadModel['getAllItems']>[number]
type PPTEraserStrokeItem = Extract<
  PPTEraserItem,
  { type: 'highlight' | 'marker' }
>

export const mergePPTEraserHitElementIds = getPPTCanvasMergedEraserHitIds

export function getPPTEraserHitElementIds({
  points,
  scene,
  slide,
}: {
  points: Point[]
  scene: PPTCanvasEraserHitInput['scene']
  slide: PPTSlide
}) {
  return getPPTCanvasEraserHitItemIds({
    itemReadModel: createPPTEraserItemReadModel(slide),
    points,
    radius: PPT_CANVAS_ERASER_RADIUS,
    scene,
  })
}

function createPPTEraserItemReadModel(slide: PPTSlide): PPTEraserItemReadModel {
  const items = slide.elements.flatMap(toPPTEraserStrokeItem)
  const itemById = new Map(items.map((item) => [item.id, item]))

  return {
    findEditableTextItem: () => null,
    findItem: (id: string) => itemById.get(id),
    getAllIds: () => items.map((item) => item.id),
    getAllItems: () => items,
    getItemBounds: getPPTEraserItemBounds,
    getSelection: (ids: string[]) => ids.filter((id) => itemById.has(id)),
    getSelectionBounds: (ids: Iterable<string>) =>
      unionPPTCanvasRectList(Array.from(ids).flatMap((id) => {
        const item = itemById.get(id)

        return item ? [getPPTEraserItemBounds(item)] : []
      })),
    getSelectedItems: (ids: string[]) =>
      ids.flatMap((id) => {
        const item = itemById.get(id)

        return item ? [item] : []
      }),
  }
}

function toPPTEraserStrokeItem(element: PPTElement): PPTEraserStrokeItem[] {
  if (
    element.kind !== 'freeform' ||
    element.visible === false ||
    element.locked === true
  ) {
    return []
  }

  return [createPPTEraserStrokeItem(element)]
}

function createPPTEraserStrokeItem(
  element: PPTFreeform,
): PPTEraserStrokeItem {
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

function getPPTEraserItemBounds(item: PPTEraserItem): Bounds {
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
