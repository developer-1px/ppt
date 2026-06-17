import {
  CANVAS_MARQUEE_SELECTION_MODEL,
  EMPTY_CANVAS_SNAP_GUIDES,
  alignCanvasSelectionItems,
  canFlipCanvasSelectionItems,
  canSelectSameTypeCanvasItems,
  canTidyCanvasSelectionItems,
  cloneCanvasSelectionItems,
  createCanvasSceneAdapter,
  flipCanvasSelectionItems,
  getCanvasFullySelectedItemGroupIds,
  getCanvasGroupExpandedSelectionIds,
  getCanvasGroupedItemPointerSelection,
  getCanvasGroupedItemSelection,
  getCanvasCommandAvailability,
  getCanvasItemGroupIndexRange,
  getCanvasItemGroupMemberIdsForGroup,
  getCanvasItemPointerSelection,
  getCanvasMarqueeSelection,
  getCanvasMoveSnap,
  getCanvasSelectableItemIds,
  getCanvasSelectedItemIds,
  getCanvasSelectedItems,
  getCanvasSingleItemSelection,
  deleteCanvasSelectionItems,
  distributeCanvasSelectionItems,
  groupCanvasSelectionItems,
  insertCanvasItemAtTargetPlacement,
  isAdditivePointerInput,
  mapCanvasSelectionItems,
  moveCanvasItemToTargetPlacement,
  moveCanvasSelection,
  moveCanvasSelectionItemsToIndex,
  normalizeCanvasRotationDegrees,
  removeCanvasSelectionIds,
  reorderCanvasSelectionItems,
  resizeCanvasSelection,
  resizeCanvasSelectionItems,
  selectSameTypeCanvasItems,
  tidyCanvasSelectionItems,
  translateCanvasSelectionItems,
  ungroupCanvasSelectionItems,
  unionCanvasRectList,
  type CanvasAlignMode,
  type CanvasCommandAdapter,
  type CanvasCommandAvailability,
  type CanvasCommandAvailabilityConfig,
  type CanvasCommandItem,
  type CanvasDistributeMode,
  type CanvasReorderMode,
  type CanvasSceneEntry,
  type CanvasSnapGuides,
  type CanvasTransformAdapter,
  type CanvasTransformItem,
} from 'canvas/foundation'

export const PPT_MARQUEE_SELECTION_MODEL = CANVAS_MARQUEE_SELECTION_MODEL
export const EMPTY_PPT_CANVAS_SNAP_GUIDES = EMPTY_CANVAS_SNAP_GUIDES

export const alignPPTCanvasSelectionItems = alignCanvasSelectionItems
export const canFlipPPTCanvasSelectionItems =
  canFlipCanvasSelectionItems
export const canSelectSameTypePPTCanvasItems =
  canSelectSameTypeCanvasItems
export const canTidyPPTCanvasSelectionItems =
  canTidyCanvasSelectionItems
export const clonePPTCanvasSelectionItems = cloneCanvasSelectionItems
export const createPPTCanvasSceneAdapter = createCanvasSceneAdapter
export const deletePPTCanvasSelectionItems = deleteCanvasSelectionItems
export const distributePPTCanvasSelectionItems =
  distributeCanvasSelectionItems
export const flipPPTCanvasSelectionItems = flipCanvasSelectionItems
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
export const getPPTCanvasItemGroupIndexRange =
  getCanvasItemGroupIndexRange
export const getPPTCanvasItemGroupMemberIdsForGroup =
  getCanvasItemGroupMemberIdsForGroup
export const getPPTCanvasItemPointerSelection =
  getCanvasItemPointerSelection
export const getPPTCanvasMarqueeSelection = getCanvasMarqueeSelection
export const getPPTCanvasMoveSnap = getCanvasMoveSnap
export const getPPTCanvasSelectableItemIds = getCanvasSelectableItemIds
export const getPPTCanvasSelectedItemIds = getCanvasSelectedItemIds
export const getPPTCanvasSelectedItems = getCanvasSelectedItems
export const getPPTCanvasSingleItemSelection = getCanvasSingleItemSelection
export const groupPPTCanvasSelectionItems = groupCanvasSelectionItems
export const insertPPTCanvasItemAtTargetPlacement =
  insertCanvasItemAtTargetPlacement
export const isAdditivePPTPointerInput = isAdditivePointerInput
export const mapPPTCanvasSelectionItems = mapCanvasSelectionItems
export const movePPTCanvasItemToTargetPlacement =
  moveCanvasItemToTargetPlacement
export const movePPTCanvasSelection = moveCanvasSelection
export const movePPTCanvasSelectionItemsToIndex =
  moveCanvasSelectionItemsToIndex
export const normalizePPTCanvasRotationDegrees =
  normalizeCanvasRotationDegrees
export const removePPTCanvasSelectionIds = removeCanvasSelectionIds
export const reorderPPTCanvasSelectionItems =
  reorderCanvasSelectionItems
export const resizePPTCanvasSelection = resizeCanvasSelection
export const resizePPTCanvasSelectionItems =
  resizeCanvasSelectionItems
export const selectSameTypePPTCanvasItems = selectSameTypeCanvasItems
export const tidyPPTCanvasSelectionItems = tidyCanvasSelectionItems
export const translatePPTCanvasSelectionItems =
  translateCanvasSelectionItems
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
export type PPTCanvasSceneEntryBase = CanvasSceneEntry
export type PPTCanvasSnapGuides = CanvasSnapGuides
export type PPTCanvasTransformAdapterBase<TItem extends CanvasTransformItem> =
  CanvasTransformAdapter<TItem>
