import {
  clamp,
  type Bounds,
} from 'canvas/core'
import {
  alignCanvasSelectionItems,
  type CanvasAlignMode,
  type CanvasCommandAdapter,
  type CanvasCommandAvailability,
  type CanvasCommandAvailabilityConfig,
  type CanvasDistributeMode,
  type CanvasReorderMode,
  cloneCanvasSelectionItems,
  deleteCanvasSelectionItems,
  distributeCanvasSelectionItems,
  getCanvasSelectableItemIds,
  getCanvasSelectedItemIds,
  getCanvasCommandAvailability,
  mapCanvasSelectionItems,
  moveCanvasSelection,
  reorderCanvasSelectionItems,
  unionCanvasRectList,
} from 'canvas/foundation'
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

export type PPTCanvasCommandAvailability = CanvasCommandAvailability & {
  cut: boolean
  nudge: boolean
  paste: boolean
}

type PPTCanvasCommandAvailabilityConfig = CanvasCommandAvailabilityConfig & {
  commands: CanvasCommandAvailabilityConfig['commands'] & Readonly<{
    cut?: boolean
    nudge?: boolean
    paste?: boolean
  }>
}

export function createPPTCanvasCommandAdapter({
  frame = getPPTSlideBounds(),
}: {
  frame?: Bounds
} = {}): CanvasCommandAdapter<PPTElement> {
  return {
    alignSelection({ items, mode, selection }) {
      return alignPPTElements(items, selection, mode, frame)
    },
    cloneSelection({ createId, ids, items, offset }) {
      return clonePPTElements(items, ids, createId, offset)
    },
    deleteSelection({ items, selection }) {
      return deleteCanvasSelectionItems({
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
      const selectedItemIds = getCanvasSelectedItemIds({
        getItemId: getPPTCommandElementId,
        isItemSelectable: isPPTCommandElementGroupable,
        items,
        selection,
      })

      if (selectedItemIds.length < 2) {
        return { items, selection }
      }

      return {
        items: mapCanvasSelectionItems({
          getItemId: getPPTCommandElementId,
          items,
          mapItem: (item) => ({ ...item, groupId }),
          selection: selectedItemIds,
        }),
        selection: selectedItemIds,
      }
    },
    lockSelection({ items, selection }) {
      return {
        items: mapCanvasSelectionItems({
          getItemId: getPPTCommandElementId,
          items,
          mapItem: (item) => ({ ...item, locked: true }),
          selection,
        }),
        selection,
      }
    },
    nudgeSelection({ dx, dy, items, selection }) {
      return moveCanvasSelection({
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
      return getCanvasSelectableItemIds({
        getItemId: getPPTCommandElementId,
        isItemSelectable: isPPTCommandElementVisible,
        items,
      })
    },
    ungroupSelection({ items, selection }) {
      const selectedGroupIds = getSelectedPPTGroupIds(items, selection)

      if (selectedGroupIds.size === 0) {
        return { items, selection: [] }
      }

      return {
        items: items.map((item) =>
          item.groupId && selectedGroupIds.has(item.groupId)
            ? ungroupPPTElement(item)
            : item),
        selection: items
          .filter((item) => item.groupId && selectedGroupIds.has(item.groupId))
          .map((item) => item.id),
      }
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
  const baseAvailability = getCanvasCommandAvailability({
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
  const ids = new Set(slide.elements.map((element) => element.id))
  let next = slide.elements.length + 1

  return (prefix: string) => {
    let id = `${slide.id}-${prefix}-${next}`

    while (ids.has(id)) {
      next += 1
      id = `${slide.id}-${prefix}-${next}`
    }

    ids.add(id)
    next += 1

    return id
  }
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
  return unionCanvasRectList(elements.map((element) =>
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
  mode: CanvasAlignMode,
  slideFrame: Bounds,
) {
  return alignCanvasSelectionItems({
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
  mode: CanvasDistributeMode,
) {
  return distributeCanvasSelectionItems({
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
  return getCanvasSelectedItemIds({
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
  const selection = options.selection ?? getCanvasSelectableItemIds({
    getItemId: getPPTCommandElementId,
    items,
  })

  return cloneCanvasSelectionItems({
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
  const geometry = {
    ...item.geometry,
    x: clamp(item.geometry.x + offset.x, 0, PPT_SLIDE_WIDTH - item.geometry.w),
    y: clamp(item.geometry.y + offset.y, 0, PPT_SLIDE_HEIGHT - item.geometry.h),
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
  mode: CanvasReorderMode,
) {
  return reorderCanvasSelectionItems({
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

function getSelectedPPTGroupIds(
  items: PPTElement[],
  selection: string[],
) {
  const selected = new Set(selection)
  const groupIds = new Set<string>()

  for (const item of items) {
    if (selected.has(item.id) && item.groupId) {
      groupIds.add(item.groupId)
    }
  }

  return groupIds
}
