import {
  CANVAS_MARQUEE_SELECTION_MODEL,
  EMPTY_CANVAS_SNAP_GUIDES,
  alignCanvasSelectionItems,
  cloneCanvasSelectionItems,
  getCanvasFullySelectedItemGroupIds,
  getCanvasGroupExpandedSelectionIds,
  getCanvasGroupedItemPointerSelection,
  getCanvasGroupedItemSelection,
  getCanvasCommandAvailability,
  getCanvasMarqueeSelection,
  getCanvasMoveSnap,
  getCanvasSelectableItemIds,
  getCanvasSelectedItemIds,
  getCanvasSelectedItems,
  getCanvasSingleItemSelection,
  deleteCanvasSelectionItems,
  distributeCanvasSelectionItems,
  groupCanvasSelectionItems,
  isAdditivePointerInput,
  mapCanvasSelectionItems,
  moveCanvasSelection,
  normalizeCanvasRotationDegrees,
  removeCanvasSelectionIds,
  reorderCanvasSelectionItems,
  resizeCanvasSelection,
  ungroupCanvasSelectionItems,
  unionCanvasRectList,
  type CanvasAlignMode,
  type CanvasCommandAdapter,
  type CanvasCommandAvailability,
  type CanvasCommandAvailabilityConfig,
  type CanvasCommandItem,
  type CanvasDistributeMode,
  type CanvasReorderMode,
  type CanvasSnapGuides,
} from 'canvas/foundation'

export const PPT_MARQUEE_SELECTION_MODEL = CANVAS_MARQUEE_SELECTION_MODEL
export const EMPTY_PPT_CANVAS_SNAP_GUIDES = EMPTY_CANVAS_SNAP_GUIDES

export const alignPPTCanvasSelectionItems = alignCanvasSelectionItems
export const clonePPTCanvasSelectionItems = cloneCanvasSelectionItems
export const deletePPTCanvasSelectionItems = deleteCanvasSelectionItems
export const distributePPTCanvasSelectionItems =
  distributeCanvasSelectionItems
export const getPPTCanvasCommandBaseAvailability =
  getCanvasCommandAvailability
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
export const getPPTCanvasSelectableItemIds = getCanvasSelectableItemIds
export const getPPTCanvasSelectedItemIds = getCanvasSelectedItemIds
export const getPPTCanvasSelectedItems = getCanvasSelectedItems
export const getPPTCanvasSingleItemSelection = getCanvasSingleItemSelection
export const groupPPTCanvasSelectionItems = groupCanvasSelectionItems
export const isAdditivePPTPointerInput = isAdditivePointerInput
export const mapPPTCanvasSelectionItems = mapCanvasSelectionItems
export const movePPTCanvasSelection = moveCanvasSelection
export const normalizePPTCanvasRotationDegrees =
  normalizeCanvasRotationDegrees
export const removePPTCanvasSelectionIds = removeCanvasSelectionIds
export const reorderPPTCanvasSelectionItems =
  reorderCanvasSelectionItems
export const resizePPTCanvasSelection = resizeCanvasSelection
export const ungroupPPTCanvasSelectionItems =
  ungroupCanvasSelectionItems
export const unionPPTCanvasRectList = unionCanvasRectList

export type PPTCanvasFoundationAlignMode = CanvasAlignMode
export type PPTCanvasCommandAdapterBase<TItem extends CanvasCommandItem> =
  CanvasCommandAdapter<TItem>
export type PPTCanvasCommandBaseAvailability = CanvasCommandAvailability
export type PPTCanvasCommandBaseAvailabilityConfig =
  CanvasCommandAvailabilityConfig
export type PPTCanvasFoundationDistributeMode = CanvasDistributeMode
export type PPTCanvasFoundationReorderMode = CanvasReorderMode
export type PPTCanvasSnapGuides = CanvasSnapGuides
