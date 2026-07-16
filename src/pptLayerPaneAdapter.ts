import {
  createSlideEditLayerPaneDescriptor,
  getSlideEditObjectVisibilityCommandAvailability,
  getSlideEditObjectVisibilityState,
  type SlideEditLayerPaneDescriptor,
  type SlideEditLayerPaneDropPlacement,
  type SlideEditLayerPaneHostCommandEffect,
  type SlideEditLayerPaneIntent,
  type SlideEditLayerPaneKeyboardIntent,
  type SlideEditLayerPaneRowDescriptor,
  type SlideEditObjectSelectionPolicy,
  type SlideEditObjectVisibilityCommandAvailability,
  type SlideEditObjectVisibilityCommandId,
  type SlideEditObjectVisibilityDescriptor,
  type SlideEditObjectVisibilityHostCommandEffect,
  type SlideEditObjectVisibilityState,
} from './pptSlideEditAffordanceAdapter'
import {
  getPPTElementGroupIndexRange,
  getPPTElementGroupMemberIds,
  movePPTElementsToIndex,
} from './pptCanvasAdapter'
import { uniquePPTCanvasValues } from './pptCanvasCoreAdapter'
import {
  getPPTCanvasFullySelectedItemGroupIds,
  getPPTCanvasGroupExpandedSelectionIds,
  getPPTCanvasGroupedItemPointerSelection,
  getPPTCanvasGroupedItemSelection,
  getPPTCanvasSingleItemSelection,
} from './pptCanvasFoundationAdapter'
import {
  findPPTElement,
  type PPTElement,
  type PPTShapeKind,
  type PPTSlide,
} from './pptModel'

export type PPTLayerPaneRowDescriptor =
  SlideEditLayerPaneRowDescriptor<string, string, string>
export type PPTLayerPaneDescriptor =
  SlideEditLayerPaneDescriptor<string, string, string>
export type PPTLayerPaneHostCommandEffect =
  SlideEditLayerPaneHostCommandEffect<string, string>
export type PPTLayerPaneIntent = SlideEditLayerPaneIntent<string>
export type PPTLayerPaneKeyboardIntent = SlideEditLayerPaneKeyboardIntent<string>
export type PPTLayerPaneDropPlacement = Exclude<
  SlideEditLayerPaneDropPlacement,
  'none'
>
export type PPTObjectVisibilityDescriptor =
  SlideEditObjectVisibilityDescriptor<string, string, string>
export type PPTObjectVisibilityHostCommandEffect =
  SlideEditObjectVisibilityHostCommandEffect<string, string>

const PPT_LAYER_PANE_GROUP_ROW_PREFIX = 'ppt-layer-group:'

export function createPPTLayerPaneDescriptor({
  activeObjectId = null,
  collapsedGroupIds,
  selectedObjectIds,
  slide,
}: {
  activeObjectId?: string | null
  collapsedGroupIds: ReadonlySet<string>
  selectedObjectIds: readonly string[]
  slide: PPTSlide
}): PPTLayerPaneDescriptor {
  return createSlideEditLayerPaneDescriptor({
    activeObjectId,
    objects: getPPTLayerPaneObjectInputs({ collapsedGroupIds, slide }),
    selectedObjectIds: getPPTLayerPaneSelectedRowIds(slide, selectedObjectIds),
    slideId: slide.id,
  })
}

export function getPPTLayerPaneGroupIdFromRowId(objectId: string) {
  return objectId.startsWith(PPT_LAYER_PANE_GROUP_ROW_PREFIX)
    ? objectId.slice(PPT_LAYER_PANE_GROUP_ROW_PREFIX.length)
    : null
}

export function getPPTLayerPaneDefaultFocusObjectId(
  slide: PPTSlide,
  selection: readonly string[],
) {
  const firstSelectedObjectId = selection[0]
  const firstSelectedElement = findPPTElement(slide, firstSelectedObjectId ?? null)

  if (
    firstSelectedElement?.groupId &&
    getPPTLayerPaneFullySelectedGroupIds(slide, selection)
      .includes(firstSelectedElement.groupId)
  ) {
    return toPPTLayerPaneGroupRowId(firstSelectedElement.groupId)
  }

  return firstSelectedObjectId ?? null
}

export function getPPTLayerPaneActualObjectIds(
  slide: PPTSlide,
  objectIds: readonly string[],
) {
  return uniquePPTCanvasValues(getPPTCanvasGroupExpandedSelectionIds({
    getItemGroupId: (element) => element.groupId,
    getItemId: (element) => element.id,
    getSelectionGroupId: getPPTLayerPaneGroupIdFromRowId,
    items: slide.elements,
    selection: objectIds,
  }))
}

export function getPPTObjectVisibilityDescriptors(
  slideId: string,
  rows: readonly PPTLayerPaneRowDescriptor[],
): PPTObjectVisibilityDescriptor[] {
  return rows.map((row) => ({
    isHidden: row.isHidden,
    isLocked: row.isLocked,
    isSelectable: row.isSelectable,
    objectId: row.objectId,
    placeholderId: null,
    slideId,
  }))
}

export function getPPTObjectVisibilityState(
  row: PPTLayerPaneRowDescriptor,
  selectionPolicy: SlideEditObjectSelectionPolicy = 'allow-hidden-selection',
): SlideEditObjectVisibilityState {
  return getSlideEditObjectVisibilityState({
    isHidden: row.isHidden,
    isLocked: row.isLocked,
    selectionPolicy,
  })
}

export function getPPTObjectVisibilityCommandId(
  row: PPTLayerPaneRowDescriptor,
): SlideEditObjectVisibilityCommandId {
  return row.isHidden ? 'show-objects' : 'hide-objects'
}

export function getPPTObjectVisibilityAvailability({
  descriptors,
  row,
}: {
  descriptors: readonly PPTObjectVisibilityDescriptor[]
  row: PPTLayerPaneRowDescriptor
}): SlideEditObjectVisibilityCommandAvailability<string> {
  return getSlideEditObjectVisibilityCommandAvailability({
    commandId: getPPTObjectVisibilityCommandId(row),
    objects: descriptors,
    selectedObjectIds: [row.objectId],
  })
}

export function getPPTLayerPaneDropIndex(
  slide: PPTSlide,
  targetObjectId: string,
  placement: PPTLayerPaneDropPlacement,
) {
  const targetGroupId = getPPTLayerPaneGroupIdFromRowId(targetObjectId)

  if (targetGroupId) {
    const groupRange = getPPTElementGroupIndexRange({
      elements: slide.elements,
      groupId: targetGroupId,
    })

    if (!groupRange) {
      return null
    }

    return placement === 'before'
      ? groupRange.firstIndex
      : groupRange.lastIndex + 1
  }

  const targetIndex = slide.elements.findIndex((element) =>
    element.id === targetObjectId)

  if (targetIndex < 0) {
    return null
  }

  return placement === 'before' ? targetIndex : targetIndex + 1
}

export function getPPTLayerPaneKeyboardDropIndex(
  descriptor: PPTLayerPaneDescriptor,
  slide: PPTSlide,
  objectId: string,
  toIndex: number,
) {
  const fromIndex = descriptor.rows.findIndex((row) => row.objectId === objectId)

  if (fromIndex < 0) {
    return null
  }

  if (fromIndex < toIndex) {
    const targetRow = descriptor.rows[toIndex - 1]

    return targetRow
      ? getPPTLayerPaneDropIndex(slide, targetRow.objectId, 'after')
      : slide.elements.length
  }

  const targetRow = descriptor.rows[toIndex]

  return targetRow
    ? getPPTLayerPaneDropIndex(slide, targetRow.objectId, 'before')
    : slide.elements.length
}

export function reorderPPTLayerPaneElement(
  elements: readonly PPTElement[],
  objectId: string,
  toIndex: number,
) {
  const groupId = getPPTLayerPaneGroupIdFromRowId(objectId)
  const selection = groupId
    ? getPPTElementGroupMemberIds({ elements, groupId })
    : [objectId]
  const result = movePPTElementsToIndex({ elements, selection, toIndex })

  return result.changed ? result.items : null
}

export function getPPTLayerPaneSelection({
  currentSelection,
  mode,
  objectIds,
  slide,
}: {
  currentSelection: string[]
  mode: 'additive' | 'range' | 'replace'
  objectIds: readonly string[]
  slide: PPTSlide
}) {
  const targetObjectId = objectIds.at(-1)

  if (!targetObjectId) {
    return currentSelection
  }

  const targetGroupId = getPPTLayerPaneGroupIdFromRowId(targetObjectId)

  if (targetGroupId) {
    return getPPTCanvasGroupedItemSelection({
      additive: mode === 'additive',
      fallbackSelection: mode === 'additive' ? currentSelection : [],
      getItemGroupId: (element) => element.groupId,
      getItemId: (element) => element.id,
      groupId: targetGroupId,
      items: slide.elements,
      selection: currentSelection,
    })
  }

  if (mode === 'range') {
    return getPPTLayerPaneActualObjectIds(slide, objectIds)
  }

  const fallbackSelection = getPPTCanvasSingleItemSelection({
    additive: mode === 'additive',
    itemId: targetObjectId,
    selection: currentSelection,
  })

  return getPPTCanvasGroupedItemPointerSelection({
    additive: mode === 'additive',
    fallbackSelection,
    getItemGroupId: (element) => element.groupId,
    getItemId: (element) => element.id,
    isItemSelectable: () => true,
    itemId: targetObjectId,
    items: slide.elements,
    selection: currentSelection,
  })
}

function getPPTLayerPaneObjectInputs({
  collapsedGroupIds,
  slide,
}: {
  collapsedGroupIds: ReadonlySet<string>
  slide: PPTSlide
}) {
  const groupedElements = new Map<string, PPTElement[]>()

  for (const element of slide.elements) {
    if (!element.groupId) continue
    const group = groupedElements.get(element.groupId) ?? []
    group.push(element)
    groupedElements.set(element.groupId, group)
  }

  const addedGroupIds = new Set<string>()
  const objects: Array<{
    displayName: string
    groupId?: string | null
    isExpanded?: boolean
    isGroup?: boolean
    isHidden?: boolean
    isLocked?: boolean
    isRenamable?: boolean
    isReorderable?: boolean
    isSelectable?: boolean
    kindLabel: string
    objectId: string
    order?: number
    parentObjectId?: string | null
  }> = []
  let order = 0

  for (const element of slide.elements) {
    if (!element.groupId) {
      objects.push(getPPTLayerPaneElementInput(element, order++))
      continue
    }
    if (addedGroupIds.has(element.groupId)) continue

    addedGroupIds.add(element.groupId)
    const groupElements = groupedElements.get(element.groupId) ?? []
    const groupRowId = toPPTLayerPaneGroupRowId(element.groupId)
    const isExpanded = !collapsedGroupIds.has(element.groupId)

    objects.push({
      displayName: `Group ${addedGroupIds.size}`,
      groupId: element.groupId,
      isExpanded,
      isGroup: true,
      isHidden: groupElements.every((member) => member.visible === false),
      isLocked: groupElements.length > 0 &&
        groupElements.every((member) => member.locked === true),
      isRenamable: false,
      isReorderable: groupElements.length > 0 &&
        groupElements.every((member) => member.locked !== true),
      isSelectable: groupElements.length > 0,
      kindLabel: 'Group',
      objectId: groupRowId,
      order: order++,
      parentObjectId: null,
    })

    if (!isExpanded) continue

    for (const groupElement of groupElements) {
      objects.push(getPPTLayerPaneElementInput(groupElement, order++, groupRowId))
    }
  }

  return objects
}

function getPPTLayerPaneElementInput(
  element: PPTElement,
  order: number,
  parentObjectId: string | null = null,
) {
  return {
    displayName: element.name,
    groupId: element.groupId ?? null,
    isGroup: false,
    isHidden: element.visible === false,
    isLocked: element.locked === true,
    isRenamable: true,
    isReorderable: element.locked !== true,
    isSelectable: true,
    kindLabel: getPPTElementKindLabel(element),
    objectId: element.id,
    order,
    parentObjectId,
  }
}

function toPPTLayerPaneGroupRowId(groupId: string) {
  return `${PPT_LAYER_PANE_GROUP_ROW_PREFIX}${groupId}`
}

function getPPTLayerPaneFullySelectedGroupIds(
  slide: PPTSlide,
  selectedObjectIds: readonly string[],
) {
  return getPPTCanvasFullySelectedItemGroupIds({
    getItemGroupId: (element) => element.groupId,
    getItemId: (element) => element.id,
    items: slide.elements,
    selection: selectedObjectIds,
  })
}

function getPPTLayerPaneSelectedRowIds(
  slide: PPTSlide,
  selectedObjectIds: readonly string[],
) {
  const selectedRowIds = new Set(selectedObjectIds)

  for (const groupId of getPPTLayerPaneFullySelectedGroupIds(slide, selectedObjectIds)) {
    selectedRowIds.add(toPPTLayerPaneGroupRowId(groupId))
  }

  return [...selectedRowIds]
}

function getPPTElementKindLabel(element: PPTElement) {
  switch (element.kind) {
    case 'shape':
      return getPPTShapeLabel(element.shape)
    case 'textBox':
      return 'Text'
    case 'image':
      return 'Image'
    case 'line':
      return 'Line'
    case 'freeform':
      return 'Freeform'
    case 'table':
      return 'Table'
    case 'comment':
      return 'Comment'
  }
}

function getPPTShapeLabel(shape: PPTShapeKind) {
  return shape === 'ellipse'
    ? 'Oval'
    : shape === 'diamond'
      ? 'Diamond'
      : 'Rectangle'
}
