import {
  CANVAS_ANGLE_CONSTRAINED_LINE_ENDPOINT_MODEL,
  CANVAS_CENTER_OUT_CREATION_POINTS_MODEL,
  CANVAS_CREATED_RECT_BOUNDS_MODEL,
  CANVAS_COMMAND_AFFORDANCES,
  CANVAS_TOOL_AFFORDANCES,
  alignCanvasCommand,
  createCanvasAffordanceConfig,
  createCanvasShape,
  createCanvasText,
  deleteCanvasCommand,
  distributeCanvasCommand,
  duplicateCanvasCommand,
  getCanvasAspectLockedCreationPoint,
  getCanvasCenterOutCreationPoints,
  getCanvasCreatedRectBounds,
  getCanvasWheelViewport,
  groupCanvasCommand,
  getCanvasAngleConstrainedLineEndPoint,
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
  type CanvasCreationAdapter,
  type CanvasCreationItem,
  type CanvasDistributeMode,
  type CanvasReorderMode,
} from 'canvas/engine'

export const PPT_COMMAND_AFFORDANCES = CANVAS_COMMAND_AFFORDANCES
export const PPT_CENTER_OUT_CREATION_POINTS_MODEL =
  CANVAS_CENTER_OUT_CREATION_POINTS_MODEL
export const PPT_ANGLE_CONSTRAINED_LINE_ENDPOINT_MODEL =
  CANVAS_ANGLE_CONSTRAINED_LINE_ENDPOINT_MODEL
export const PPT_CREATED_RECT_BOUNDS_MODEL =
  CANVAS_CREATED_RECT_BOUNDS_MODEL
export const PPT_TOOL_AFFORDANCES = CANVAS_TOOL_AFFORDANCES

export const alignPPTCanvasCommand = alignCanvasCommand
export const createPPTCanvasAffordanceConfig = createCanvasAffordanceConfig
export const createPPTCanvasShape = createCanvasShape
export const createPPTCanvasText = createCanvasText
export const deletePPTCanvasCommand = deleteCanvasCommand
export const distributePPTCanvasCommand = distributeCanvasCommand
export const duplicatePPTCanvasCommand = duplicateCanvasCommand
export const getPPTCanvasAspectLockedCreationPoint =
  getCanvasAspectLockedCreationPoint
export const getPPTCanvasCenterOutCreationPoints =
  getCanvasCenterOutCreationPoints
export const getPPTCanvasCreatedRectBounds = getCanvasCreatedRectBounds
export const getPPTCanvasWheelViewport = getCanvasWheelViewport
export const groupPPTCanvasCommand = groupCanvasCommand
export const getPPTCanvasAngleConstrainedLineEndPoint =
  getCanvasAngleConstrainedLineEndPoint
export const lockPPTCanvasCommand = lockCanvasCommand
export const nudgePPTCanvasCommand = nudgeCanvasCommand
export const reorderPPTCanvasCommand = reorderCanvasCommand
export const selectAllPPTCanvasCommand = selectAllCanvasCommand
export const ungroupPPTCanvasCommand = ungroupCanvasCommand
export const unlockAllPPTCanvasCommand = unlockAllCanvasCommand

export type PPTCanvasAlignMode = CanvasAlignMode
export type PPTCanvasCommandItemsResult<TItem extends CanvasCommandItem> =
  CanvasCommandItemsResult<TItem>
export type PPTCanvasCreatedShapeKind = CanvasCreatedShapeKind
export type PPTCanvasCreationAdapter<TItem extends CanvasCreationItem> =
  CanvasCreationAdapter<TItem>
export type PPTCanvasDistributeMode = CanvasDistributeMode
export type PPTCanvasReorderMode = CanvasReorderMode
