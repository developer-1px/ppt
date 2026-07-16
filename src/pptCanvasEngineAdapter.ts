import {
  CANVAS_COMMAND_AFFORDANCES,
  CANVAS_TOOL_AFFORDANCES,
  alignCanvasCommand,
  createCanvasAffordanceConfig,
  createCanvasShape,
  createCanvasText,
  deleteCanvasCommand,
  distributeCanvasCommand,
  duplicateCanvasCommand,
  getCanvasAngleConstrainedPoint,
  getCanvasAspectRatioLockedPoint,
  getCanvasWheelViewport,
  groupCanvasCommand,
  lockCanvasCommand,
  nudgeCanvasCommand,
  reorderCanvasCommand,
  selectAllCanvasCommand,
  ungroupCanvasCommand,
  unlockAllCanvasCommand,
  type CanvasAlignMode,
  type CanvasCommandItem,
  type CanvasCommandItemsResult,
  type CanvasCreatedShapeKind,
  type CanvasCreatedRectBoundsInput,
  type CanvasCreationAdapter,
  type CanvasCreationItem,
  type CanvasDistributeMode,
  type CanvasReorderMode,
} from 'canvas/engine'
import type {
  Bounds,
  Point,
} from 'canvas/core'

export const PPT_COMMAND_AFFORDANCES = CANVAS_COMMAND_AFFORDANCES
export const PPT_CENTER_OUT_CREATION_POINTS_MODEL =
  'canvas-center-out-creation-points'
export const PPT_ANGLE_CONSTRAINED_LINE_ENDPOINT_MODEL =
  'canvas-angle-constrained-line-endpoint'
export const PPT_CREATED_RECT_BOUNDS_MODEL =
  'canvas-created-rect-bounds'
export const PPT_TOOL_AFFORDANCES = CANVAS_TOOL_AFFORDANCES

type PPTCanvasCreatedRectBoundsInput = CanvasCreatedRectBoundsInput & {
  defaultSize?: Readonly<{
    h: number
    w: number
  }>
  dragThreshold?: number
}

type PPTCanvasCenterOutCreationPointsInput = Readonly<{
  currentWorld: Point
  startWorld: Point
}>

const PPT_CANVAS_DEFAULT_RECT_SIZE = {
  h: 112,
  w: 168,
}

const PPT_CANVAS_DEFAULT_DRAG_THRESHOLD = 6

export const alignPPTCanvasCommand = alignCanvasCommand
export const createPPTCanvasAffordanceConfig = createCanvasAffordanceConfig
export const createPPTCanvasShape = createCanvasShape
export const createPPTCanvasText = createCanvasText
export const deletePPTCanvasCommand = deleteCanvasCommand
export const distributePPTCanvasCommand = distributeCanvasCommand
export const duplicatePPTCanvasCommand = duplicateCanvasCommand
export const getPPTCanvasAspectLockedCreationPoint =
  getCanvasAspectRatioLockedPoint
export const getPPTCanvasWheelViewport = getCanvasWheelViewport
export const groupPPTCanvasCommand = groupCanvasCommand
export const getPPTCanvasAngleConstrainedLineEndPoint =
  getCanvasAngleConstrainedPoint
export const lockPPTCanvasCommand = lockCanvasCommand
export const nudgePPTCanvasCommand = nudgeCanvasCommand
export const reorderPPTCanvasCommand = reorderCanvasCommand
export const selectAllPPTCanvasCommand = selectAllCanvasCommand
export const ungroupPPTCanvasCommand = ungroupCanvasCommand
export const unlockAllPPTCanvasCommand = unlockAllCanvasCommand

export function getPPTCanvasCenterOutCreationPoints({
  currentWorld,
  startWorld,
}: PPTCanvasCenterOutCreationPointsInput): {
  current: Point
  start: Point
} {
  return {
    current: currentWorld,
    start: getPPTCanvasMirrorPoint(startWorld, currentWorld),
  }
}

export function getPPTCanvasCreatedRectBounds({
  currentWorld,
  defaultSize = PPT_CANVAS_DEFAULT_RECT_SIZE,
  dragThreshold = PPT_CANVAS_DEFAULT_DRAG_THRESHOLD,
  preserveAspectRatio = false,
  resizeFromCenter = false,
  startWorld,
}: PPTCanvasCreatedRectBoundsInput): Bounds {
  const endWorld = preserveAspectRatio
    ? getPPTCanvasAspectLockedCreationPoint({ currentWorld, startWorld })
    : currentWorld
  const rawBounds = resizeFromCenter
    ? normalizePPTCanvasCreationBounds(
      getPPTCanvasMirrorPoint(startWorld, endWorld),
      endWorld,
    )
    : normalizePPTCanvasCreationBounds(startWorld, endWorld)

  if (rawBounds.w > dragThreshold && rawBounds.h > dragThreshold) {
    return rawBounds
  }

  return resizeFromCenter
    ? {
        h: defaultSize.h,
        w: defaultSize.w,
        x: startWorld.x - defaultSize.w / 2,
        y: startWorld.y - defaultSize.h / 2,
      }
    : {
        h: defaultSize.h,
        w: defaultSize.w,
        x: startWorld.x,
        y: startWorld.y,
      }
}

function getPPTCanvasMirrorPoint(origin: Point, point: Point): Point {
  return {
    x: origin.x - (point.x - origin.x),
    y: origin.y - (point.y - origin.y),
  }
}

function normalizePPTCanvasCreationBounds(start: Point, end: Point): Bounds {
  const x = Math.min(start.x, end.x)
  const y = Math.min(start.y, end.y)

  return {
    h: Math.abs(end.y - start.y),
    w: Math.abs(end.x - start.x),
    x,
    y,
  }
}

export type PPTCanvasAlignMode = CanvasAlignMode
export type PPTCanvasCommandItemsResult<TItem extends CanvasCommandItem> =
  CanvasCommandItemsResult<TItem>
export type PPTCanvasCreatedShapeKind = CanvasCreatedShapeKind
export type PPTCanvasCreationAdapter<TItem extends CanvasCreationItem> =
  CanvasCreationAdapter<TItem>
export type PPTCanvasDistributeMode = CanvasDistributeMode
export type PPTCanvasReorderMode = CanvasReorderMode
