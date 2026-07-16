import {
  clampPPTCanvasBoundsToFrame,
  createPPTCanvasSequentialIdFactory,
  type Bounds,
} from './pptCanvasCoreAdapter'
import {
  alignPPTCanvasSelectionItems,
  clonePPTCanvasSelectionItems,
  deletePPTCanvasSelectionItems,
  distributePPTCanvasSelectionItems,
  getPPTCanvasCommandBaseAvailability,
  getPPTCanvasSelectableItemIds,
  getPPTCanvasSelectedItemIds,
  groupPPTCanvasSelectionItems,
  mapPPTCanvasSelectionItems,
  movePPTCanvasSelection,
  reorderPPTCanvasSelectionItems,
  ungroupPPTCanvasSelectionItems,
  unionPPTCanvasRectList,
  type PPTCanvasCommandAdapterBase,
  type PPTCanvasCommandBaseAvailability,
  type PPTCanvasCommandBaseAvailabilityConfig,
  type PPTCanvasFoundationAlignMode,
  type PPTCanvasFoundationDistributeMode,
  type PPTCanvasFoundationReorderMode,
} from './pptCanvasFoundationAdapter'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTElement,
  type PPTSlide,
} from './pptModel'
import {
  boundsToPPTGeometry,
  pptCanvasTransformAdapter,
  pptGeometryToBounds,
} from './pptCanvasAdapter'

export type PPTCanvasCommandAvailability = PPTCanvasCommandBaseAvailability & {
  cut: boolean
  nudge: boolean
  paste: boolean
}

type PPTCanvasCommandAvailabilityConfig = PPTCanvasCommandBaseAvailabilityConfig & {
  commands: PPTCanvasCommandBaseAvailabilityConfig['commands'] & Readonly<{
    cut?: boolean
    nudge?: boolean
    paste?: boolean
  }>
}

export function createPPTCanvasCommandAdapter({
  frame = getPPTSlideBounds(),
}: {
  frame?: Bounds
} = {}): PPTCanvasCommandAdapterBase<PPTElement> {
  return {
    alignSelection({ items, mode, selection }) {
      return alignPPTElements(items, selection, mode, frame)
    },
    cloneSelection({ createId, ids, items, offset }) {
      return clonePPTElements(items, ids, createId, offset)
    },
    deleteSelection({ items, selection }) {
      return deletePPTCanvasSelectionItems({
        getItemId: getPPTCommandElementId,
        isItemSelectable: isPPTCommandElementEditable,
        items,
        selection,
      })
    },
    distributeSelection({ items, mode, selection }) {
      return distributePPTElements(items, selection, mode)
    },
    groupSelection({ groupId, items, selection }) {
      return groupPPTCanvasSelectionItems({
        getItemId: getPPTCommandElementId,
        groupId,
        groupItem: ({ groupId, item }) => ({ ...item, groupId }),
        isItemSelectable: isPPTCommandElementGroupable,
        items,
        selection,
      })
    },
    lockSelection({ items, selection }) {
      return {
        items: mapPPTCanvasSelectionItems({
          getItemId: getPPTCommandElementId,
          items,
          mapItem: (item) => ({ ...item, locked: true }),
          selection,
        }),
        selection,
      }
    },
    nudgeSelection({ dx, dy, items, selection }) {
      return movePPTCanvasSelection({
        adapter: pptCanvasTransformAdapter,
        dx,
        dy,
        items,
        selection: getEditablePPTSelection(items, selection),
      })
    },
    pasteItems({ clipboard, createId, offset }) {
      return clonePPTElementsFromSource(clipboard, createId, offset)
    },
    reorderSelection({ items, mode, selection }) {
      return reorderPPTElements(items, selection, mode)
    },
    selectAll({ items }) {
      return getPPTCanvasSelectableItemIds({
        getItemId: getPPTCommandElementId,
        isItemSelectable: isPPTCommandElementVisible,
        items,
      })
    },
    ungroupSelection({ items, selection }) {
      return ungroupPPTCanvasSelectionItems({
        getItemGroupId: (item) => item.groupId,
        getItemId: getPPTCommandElementId,
        items,
        selection,
        ungroupItem: ({ item }) => ungroupPPTElement(item),
      })
    },
    unlockAll({ items, selection }) {
      return {
        items: items.map(unlockPPTElement),
        selection,
      }
    },
  }
}

export function getPPTCanvasCommandAvailability({
  canPaste,
  canRedo,
  canUndo,
  config,
  hasHiddenSelection = false,
  hasGroupedSelection = false,
  hasLockedItems = false,
  hasLockedSelection = false,
  selection,
}: {
  canPaste: boolean
  canRedo: boolean
  canUndo: boolean
  config: PPTCanvasCommandAvailabilityConfig
  hasGroupedSelection?: boolean
  hasHiddenSelection?: boolean
  hasLockedItems?: boolean
  hasLockedSelection?: boolean
  selection: readonly string[]
}): PPTCanvasCommandAvailability {
  const baseAvailability = getPPTCanvasCommandBaseAvailability({
    canRedo,
    canUndo,
    config,
    hasSelectedGroup: hasGroupedSelection,
    selection,
  })
  const hasSelection = selection.length > 0
  const canEditSelection = hasSelection && !hasLockedSelection
  const canTransformSelection = canEditSelection && !hasHiddenSelection

  return {
    ...baseAvailability,
    alignBottom: config.commands.alignBottom && canTransformSelection,
    alignCenter: config.commands.alignCenter && canTransformSelection,
    alignLeft: config.commands.alignLeft && canTransformSelection,
    alignMiddle: config.commands.alignMiddle && canTransformSelection,
    alignRight: config.commands.alignRight && canTransformSelection,
    alignTop: config.commands.alignTop && canTransformSelection,
    bringForward: baseAvailability.bringForward && canTransformSelection,
    bringToFront: baseAvailability.bringToFront && canTransformSelection,
    cut: config.commands.cut !== false && canEditSelection,
    delete: baseAvailability.delete && !hasLockedSelection,
    duplicate: baseAvailability.duplicate && canEditSelection,
    distributeHorizontal:
      baseAvailability.distributeHorizontal && !hasLockedSelection,
    distributeVertical:
      baseAvailability.distributeVertical && !hasLockedSelection,
    group: baseAvailability.group && canTransformSelection,
    lockSelection: baseAvailability.lockSelection && !hasLockedSelection,
    nudge: config.commands.nudge !== false && canTransformSelection,
    paste: config.commands.paste !== false && canPaste,
    sendBackward: baseAvailability.sendBackward && canTransformSelection,
    sendToBack: baseAvailability.sendToBack && canTransformSelection,
    ungroup: baseAvailability.ungroup && canEditSelection,
    unlockAll: baseAvailability.unlockAll && hasLockedItems,
  }
}

export function createPPTElementIdFactory(slide: PPTSlide) {
  return createPPTCanvasSequentialIdFactory({
    existingIds: slide.elements.map((element) => element.id),
    formatId: ({ index, prefix }) => `${slide.id}-${prefix}-${index}`,
    startIndex: slide.elements.length + 1,
  })
}

export function getPPTSlideBounds(): Bounds {
  return {
    h: PPT_SLIDE_HEIGHT,
    w: PPT_SLIDE_WIDTH,
    x: 0,
    y: 0,
  }
}

export function getPPTElementsBounds(elements: PPTElement[]): Bounds | null {
  return unionPPTCanvasRectList(elements.map((element) =>
    pptGeometryToBounds(element.geometry)))
}

export function updatePPTElementBounds(
  element: PPTElement,
  bounds: Bounds,
): PPTElement {
  const geometry = boundsToPPTGeometry(bounds)

  return {
    ...element,
    geometry: element.geometry.rotation === undefined
      ? geometry
      : { ...geometry, rotation: element.geometry.rotation },
  }
}

export function getPPTElementIdPrefix(element: PPTElement) {
  if (element.kind === 'textBox') {
    return 'text-copy'
  }

  return `${element.kind}-copy`
}

function alignPPTElements(
  items: PPTElement[],
  selection: string[],
  mode: PPTCanvasFoundationAlignMode,
  slideFrame: Bounds,
) {
  return alignPPTCanvasSelectionItems({
    frame: selection.length === 1 ? slideFrame : undefined,
    getItemBounds: (item) => pptGeometryToBounds(item.geometry),
    getItemId: (item) => item.id,
    isItemSelectable: isPPTCommandElementEditable,
    items,
    mode,
    selection,
    updateItemBounds: updatePPTElementBounds,
  })
}

function distributePPTElements(
  items: PPTElement[],
  selection: string[],
  mode: PPTCanvasFoundationDistributeMode,
) {
  return distributePPTCanvasSelectionItems({
    getItemBounds: (item) => pptGeometryToBounds(item.geometry),
    getItemId: (item) => item.id,
    isItemSelectable: isPPTCommandElementEditable,
    items,
    mode,
    selection,
    updateItemBounds: updatePPTElementBounds,
  })
}

function isPPTCommandElementEditable(element: PPTElement) {
  return element.locked !== true
}

function isPPTCommandElementVisible(element: PPTElement) {
  return element.visible !== false
}

function isPPTCommandElementGroupable(element: PPTElement) {
  return isPPTCommandElementEditable(element) &&
    isPPTCommandElementVisible(element)
}

function getPPTCommandElementId(element: PPTElement) {
  return element.id
}

function getEditablePPTSelection(
  items: PPTElement[],
  selection: readonly string[],
) {
  return getPPTCanvasSelectedItemIds({
    getItemId: getPPTCommandElementId,
    isItemSelectable: isPPTCommandElementEditable,
    items,
    selection,
  })
}

function clonePPTElements(
  items: PPTElement[],
  ids: string[],
  createId: (prefix: string) => string,
  offset: { x: number; y: number },
) {
  return clonePPTElementsFromSource(items, createId, offset, {
    isItemSelectable: isPPTCommandElementEditable,
    selection: ids,
  })
}

function clonePPTElementsFromSource(
  items: PPTElement[],
  createId: (prefix: string) => string,
  offset: { x: number; y: number },
  options: {
    isItemSelectable?: (item: PPTElement) => boolean
    selection?: readonly string[]
  } = {},
) {
  const selection = options.selection ?? getPPTCanvasSelectableItemIds({
    getItemId: getPPTCommandElementId,
    items,
  })

  return clonePPTCanvasSelectionItems({
    cloneItem: ({ item, targetGroupId }) =>
      clonePPTElement(
        item,
        createId(getPPTElementIdPrefix(item)),
        offset,
        targetGroupId,
      ),
    createGroupId: () => createId('group-copy'),
    getItemGroupId: (item) => item.groupId,
    getItemId: getPPTCommandElementId,
    isItemSelectable: options.isItemSelectable ?? (() => true),
    items,
    selection,
  })
}

function clonePPTElement(
  item: PPTElement,
  id: string,
  offset: { x: number; y: number },
  groupId: string | undefined,
): PPTElement {
  const bounds = clampPPTCanvasBoundsToFrame({
    bounds: {
      ...pptGeometryToBounds(item.geometry),
      x: item.geometry.x + offset.x,
      y: item.geometry.y + offset.y,
    },
    frame: getPPTSlideBounds(),
    minHeight: 24,
    minWidth: 24,
  })
  const geometry = {
    ...item.geometry,
    ...bounds,
  }

  return {
    ...item,
    geometry,
    groupId,
    id,
    name: `${item.name} Copy`,
  }
}

function reorderPPTElements(
  items: PPTElement[],
  selection: string[],
  mode: PPTCanvasFoundationReorderMode,
) {
  return reorderPPTCanvasSelectionItems({
    getItemId: (item) => item.id,
    isItemSelectable: isPPTCommandElementEditable,
    items,
    mode,
    selection,
  })
}

function unlockPPTElement(item: PPTElement): PPTElement {
  if (item.locked !== true) {
    return item
  }

  const next = { ...item }
  delete next.locked

  return next
}

function ungroupPPTElement(item: PPTElement): PPTElement {
  if (!item.groupId) {
    return item
  }

  const next = { ...item }
  delete next.groupId

  return next
}
