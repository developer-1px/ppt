import {
  CANVAS_MARQUEE_SELECTION_MODEL,
  EMPTY_CANVAS_SNAP_GUIDES,
  getCanvasFullySelectedItemGroupIds,
  getCanvasGroupExpandedSelectionIds,
  getCanvasGroupedItemPointerSelection,
  getCanvasGroupedItemSelection,
  getCanvasMarqueeSelection,
  getCanvasMoveSnap,
  getCanvasSelectedItems,
  getCanvasSingleItemSelection,
  deleteCanvasSelectionItems,
  isAdditivePointerInput,
  mapCanvasSelectionItems,
  moveCanvasSelection,
  normalizeCanvasRotationDegrees,
  removeCanvasSelectionIds,
  resizeCanvasSelection,
  type CanvasSnapGuides,
} from 'canvas/foundation'

export const PPT_MARQUEE_SELECTION_MODEL = CANVAS_MARQUEE_SELECTION_MODEL
export const EMPTY_PPT_CANVAS_SNAP_GUIDES = EMPTY_CANVAS_SNAP_GUIDES

export const deletePPTCanvasSelectionItems = deleteCanvasSelectionItems
export const getPPTCanvasFullySelectedItemGroupIds =
  getCanvasFullySelectedItemGroupIds
export const getPPTCanvasGroupExpandedSelectionIds =
  getCanvasGroupExpandedSelectionIds
export const getPPTCanvasGroupedItemPointerSelection =
  getCanvasGroupedItemPointerSelection
export const getPPTCanvasGroupedItemSelection =
  getCanvasGroupedItemSelection
export const getPPTCanvasMarqueeSelection = getCanvasMarqueeSelection
export const getPPTCanvasMoveSnap = getCanvasMoveSnap
export const getPPTCanvasSelectedItems = getCanvasSelectedItems
export const getPPTCanvasSingleItemSelection = getCanvasSingleItemSelection
export const isAdditivePPTPointerInput = isAdditivePointerInput
export const mapPPTCanvasSelectionItems = mapCanvasSelectionItems
export const movePPTCanvasSelection = moveCanvasSelection
export const normalizePPTCanvasRotationDegrees =
  normalizeCanvasRotationDegrees
export const removePPTCanvasSelectionIds = removeCanvasSelectionIds
export const resizePPTCanvasSelection = resizeCanvasSelection

export type PPTCanvasSnapGuides = CanvasSnapGuides
