import {
  clamp,
  type Bounds,
} from 'canvas/core'
import {
  type CanvasAlignMode,
  type CanvasCommandAdapter,
  type CanvasCommandAvailability,
  type CanvasDistributeMode,
  type CanvasReorderMode,
} from 'canvas/foundation'
import {
  PPT_SLIDE_HEIGHT,
  PPT_SLIDE_WIDTH,
  type PPTElement,
  type PPTSlide,
} from './pptModel'
import {
  boundsToPPTGeometry,
  pptGeometryToBounds,
} from './pptCanvasAdapter'

export type PPTCanvasCommandAvailability = CanvasCommandAvailability & {
  cut: boolean
  nudge: boolean
  paste: boolean
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
      const selected = new Set(selection)

      return items.filter((item) => !selected.has(item.id) || item.locked === true)
    },
    distributeSelection({ items, mode, selection }) {
      return distributePPTElements(items, selection, mode)
    },
    groupSelection({ groupId, items, selection }) {
      const selected = new Set(selection)
      const selectedItems = items.filter((item) =>
        selected.has(item.id) && item.locked !== true && item.visible !== false)
      const groupable = new Set(selectedItems.map((item) => item.id))

      if (selectedItems.length < 2) {
        return { items, selection }
      }

      return {
        items: items.map((item) =>
          groupable.has(item.id) ? { ...item, groupId } : item),
        selection: selectedItems.map((item) => item.id),
      }
    },
    lockSelection({ items, selection }) {
      const selected = new Set(selection)

      return {
        items: items.map((item) =>
          selected.has(item.id) ? { ...item, locked: true } : item),
        selection,
      }
    },
    nudgeSelection({ dx, dy, items, selection }) {
      const selected = new Set(selection)

      return items.map((item) =>
        selected.has(item.id) && item.locked !== true
          ? updatePPTElementBounds(item, {
              ...pptGeometryToBounds(item.geometry),
              x: item.geometry.x + dx,
              y: item.geometry.y + dy,
            })
          : item)
    },
    pasteItems({ clipboard, createId, offset }) {
      return clonePPTElementsFromSource(clipboard, createId, offset)
    },
    reorderSelection({ items, mode, selection }) {
      return reorderPPTElements(items, selection, mode)
    },
    selectAll({ items }) {
      return items
        .filter((item) => item.visible !== false)
        .map((item) => item.id)
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
  hasHiddenSelection = false,
  hasGroupedSelection = false,
  hasLockedItems = false,
  hasLockedSelection = false,
  selection,
}: {
  canPaste: boolean
  canRedo: boolean
  canUndo: boolean
  hasGroupedSelection?: boolean
  hasHiddenSelection?: boolean
  hasLockedItems?: boolean
  hasLockedSelection?: boolean
  selection: readonly string[]
}): PPTCanvasCommandAvailability {
  const hasSelection = selection.length > 0
  const canDistribute = selection.length >= 3
  const canEditSelection = hasSelection && !hasLockedSelection
  const canTransformSelection = canEditSelection && !hasHiddenSelection

  return {
    alignBottom: canTransformSelection,
    alignCenter: canTransformSelection,
    alignLeft: canTransformSelection,
    alignMiddle: canTransformSelection,
    alignRight: canTransformSelection,
    alignTop: canTransformSelection,
    bringForward: canTransformSelection,
    bringToFront: canTransformSelection,
    cut: canEditSelection,
    delete: canEditSelection,
    duplicate: canEditSelection,
    distributeHorizontal: canDistribute && !hasLockedSelection,
    distributeVertical: canDistribute && !hasLockedSelection,
    group: canTransformSelection && selection.length >= 2,
    lockSelection: hasSelection && !hasLockedSelection,
    nudge: canTransformSelection,
    paste: canPaste,
    redo: canRedo,
    selectAll: true,
    sendBackward: canTransformSelection,
    sendToBack: canTransformSelection,
    undo: canUndo,
    ungroup: canEditSelection && hasGroupedSelection,
    unlockAll: hasLockedItems,
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
  if (elements.length === 0) {
    return null
  }

  const left = Math.min(...elements.map((element) => element.geometry.x))
  const top = Math.min(...elements.map((element) => element.geometry.y))
  const right = Math.max(...elements.map((element) => element.geometry.x + element.geometry.w))
  const bottom = Math.max(...elements.map((element) => element.geometry.y + element.geometry.h))

  return {
    h: bottom - top,
    w: right - left,
    x: left,
    y: top,
  }
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
  const selected = new Set(selection)
  const selectedItems = items.filter((item) => selected.has(item.id))
  const frame = selectedItems.length === 1
    ? slideFrame
    : getPPTElementsBounds(selectedItems)

  if (!frame) {
    return items
  }

  return items.map((item) => {
    if (!selected.has(item.id) || item.locked === true) {
      return item
    }

    const bounds = pptGeometryToBounds(item.geometry)
    const next = { ...bounds }

    if (mode === 'alignLeft') {
      next.x = frame.x
    } else if (mode === 'alignCenter') {
      next.x = frame.x + frame.w / 2 - bounds.w / 2
    } else if (mode === 'alignRight') {
      next.x = frame.x + frame.w - bounds.w
    } else if (mode === 'alignTop') {
      next.y = frame.y
    } else if (mode === 'alignMiddle') {
      next.y = frame.y + frame.h / 2 - bounds.h / 2
    } else if (mode === 'alignBottom') {
      next.y = frame.y + frame.h - bounds.h
    }

    return updatePPTElementBounds(item, next)
  })
}

function distributePPTElements(
  items: PPTElement[],
  selection: string[],
  mode: CanvasDistributeMode,
) {
  const selected = new Set(selection)
  const selectedItems = items.filter((item) =>
    selected.has(item.id) && item.locked !== true)
  const distributed = mode === 'distributeHorizontal'
    ? distributePPTElementsHorizontally(selectedItems)
    : distributePPTElementsVertically(selectedItems)
  const byId = new Map(distributed.map((item) => [item.id, item]))

  return items.map((item) => byId.get(item.id) ?? item)
}

function distributePPTElementsHorizontally(elements: PPTElement[]) {
  const sorted = [...elements].sort((a, b) => a.geometry.x - b.geometry.x)
  const first = sorted[0]
  const last = sorted.at(-1)

  if (!first || !last || sorted.length < 3) {
    return elements
  }

  const totalWidth = sorted.reduce((sum, element) => sum + element.geometry.w, 0)
  const span = last.geometry.x + last.geometry.w - first.geometry.x
  const gap = (span - totalWidth) / (sorted.length - 1)
  let nextX = first.geometry.x

  return sorted.map((element, index) => {
    const x = index === 0
      ? first.geometry.x
      : index === sorted.length - 1
        ? last.geometry.x
        : nextX
    nextX = x + element.geometry.w + gap

    return updatePPTElementBounds(element, {
      ...pptGeometryToBounds(element.geometry),
      x,
    })
  })
}

function distributePPTElementsVertically(elements: PPTElement[]) {
  const sorted = [...elements].sort((a, b) => a.geometry.y - b.geometry.y)
  const first = sorted[0]
  const last = sorted.at(-1)

  if (!first || !last || sorted.length < 3) {
    return elements
  }

  const totalHeight = sorted.reduce((sum, element) => sum + element.geometry.h, 0)
  const span = last.geometry.y + last.geometry.h - first.geometry.y
  const gap = (span - totalHeight) / (sorted.length - 1)
  let nextY = first.geometry.y

  return sorted.map((element, index) => {
    const y = index === 0
      ? first.geometry.y
      : index === sorted.length - 1
        ? last.geometry.y
        : nextY
    nextY = y + element.geometry.h + gap

    return updatePPTElementBounds(element, {
      ...pptGeometryToBounds(element.geometry),
      y,
    })
  })
}

function clonePPTElements(
  items: PPTElement[],
  ids: string[],
  createId: (prefix: string) => string,
  offset: { x: number; y: number },
) {
  const sourceIds = new Set(ids)
  const sourceItems = items
    .filter((item) => sourceIds.has(item.id) && item.locked !== true)

  return clonePPTElementsFromSource(sourceItems, createId, offset)
}

function clonePPTElementsFromSource(
  items: PPTElement[],
  createId: (prefix: string) => string,
  offset: { x: number; y: number },
) {
  const groupIdBySource = new Map<string, string>()

  return items.map((item) => {
    const groupId = item.groupId
      ? getClonedPPTGroupId(groupIdBySource, item.groupId, createId)
      : undefined

    return clonePPTElement(
      item,
      createId(getPPTElementIdPrefix(item)),
      offset,
      groupId,
    )
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

function getClonedPPTGroupId(
  groupIdBySource: Map<string, string>,
  sourceGroupId: string,
  createId: (prefix: string) => string,
) {
  const existing = groupIdBySource.get(sourceGroupId)

  if (existing) {
    return existing
  }

  const next = createId('group-copy')
  groupIdBySource.set(sourceGroupId, next)

  return next
}

function reorderPPTElements(
  items: PPTElement[],
  selection: string[],
  mode: CanvasReorderMode,
) {
  const selected = new Set(
    items
      .filter((item) => selection.includes(item.id) && item.locked !== true)
      .map((item) => item.id),
  )

  if (selected.size === 0) {
    return items
  }

  if (mode === 'bringToFront') {
    return [
      ...items.filter((item) => !selected.has(item.id)),
      ...items.filter((item) => selected.has(item.id)),
    ]
  }

  if (mode === 'sendToBack') {
    return [
      ...items.filter((item) => selected.has(item.id)),
      ...items.filter((item) => !selected.has(item.id)),
    ]
  }

  return mode === 'bringForward'
    ? movePPTSelectionForward(items, selected)
    : movePPTSelectionBackward(items, selected)
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

function movePPTSelectionForward(
  items: PPTElement[],
  selected: ReadonlySet<string>,
) {
  const next = [...items]

  for (let index = next.length - 2; index >= 0; index -= 1) {
    if (selected.has(next[index].id) && !selected.has(next[index + 1].id)) {
      const item = next[index]
      next[index] = next[index + 1]
      next[index + 1] = item
    }
  }

  return next
}

function movePPTSelectionBackward(
  items: PPTElement[],
  selected: ReadonlySet<string>,
) {
  const next = [...items]

  for (let index = 1; index < next.length; index += 1) {
    if (selected.has(next[index].id) && !selected.has(next[index - 1].id)) {
      const item = next[index]
      next[index] = next[index - 1]
      next[index - 1] = item
    }
  }

  return next
}
