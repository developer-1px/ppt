import {
  RESIZE_HANDLES,
  clamp,
  clampCanvasBoundsToFrame,
  clampCanvasPointToBounds,
  createCanvasSequentialIdFactory,
  getCanvasBoundsAnchorPoints,
  getCanvasBoundsCenter,
  handlePoint,
  normalizeBounds,
  normalizeCanvasPointsToLocalBounds,
  pointDistance,
  unique,
  type Bounds,
  type Point,
  type ResizeHandle,
  type Tool,
  type Viewport,
} from 'canvas/core'

export const PPT_RESIZE_HANDLES = RESIZE_HANDLES
export const clampPPTCanvasValue = clamp
export const clampPPTCanvasBoundsToFrame = clampCanvasBoundsToFrame
export const clampPPTCanvasPointToBounds = clampCanvasPointToBounds
export const createPPTCanvasSequentialIdFactory =
  createCanvasSequentialIdFactory
export const getPPTCanvasBoundsAnchorPoints = getCanvasBoundsAnchorPoints
export const getPPTCanvasBoundsCenter = getCanvasBoundsCenter
export const getPPTCanvasHandlePoint = handlePoint
export const normalizePPTCanvasBounds = normalizeBounds
export const normalizePPTCanvasPointsToLocalBounds =
  normalizeCanvasPointsToLocalBounds
export const getPPTCanvasPointDistance = pointDistance
export const uniquePPTCanvasValues = unique

export type { Bounds, Point, ResizeHandle, Tool, Viewport }
