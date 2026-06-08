import type { JSONPatchOperation, Pointer, SelectionSnap } from 'zod-crud'
import {
  blockLocationFromPointer,
  getRect,
  moveRect,
  rectEquals,
  setArrangePatch,
  type RetouchDeck,
} from './retouchModel'
import { selectionSnapForPointers } from './retouchSelectionSnap'

type KeyboardNudgeDelta = {
  x: number
  y: number
}

type LayoutKeyboardNudgePatch = {
  label: 'nudge layout'
  mergeKey: string
  operations: JSONPatchOperation[]
  pointer: Pointer
  selection: SelectionSnap
}

export function createLayoutKeyboardNudgePatch({
  activeSlideId,
  deckValue,
  delta,
  selectedPointer,
  selectedPointers,
}: {
  activeSlideId: string
  deckValue: RetouchDeck
  delta: KeyboardNudgeDelta
  selectedPointer: Pointer | null
  selectedPointers: Pointer[]
}): LayoutKeyboardNudgePatch | null {
  const targets = selectedPointers
    .map((pointer) => {
      const location = blockLocationFromPointer(deckValue, pointer)

      if (!location || location.slide.id !== activeSlideId) {
        return null
      }

      const startRect = getRect(location.block)

      return {
        pointer,
        rect: moveRect(startRect, delta.x, delta.y),
        startRect,
      }
    })
    .filter((target): target is NonNullable<typeof target> => target !== null)

  if (
    targets.length === 0 ||
    targets.every((target) => rectEquals(target.rect, target.startRect))
  ) {
    return null
  }

  const targetPointers = targets.map((target) => target.pointer)
  const focusPointer = targetPointers.at(-1)
  const primaryPointer = focusPointer ?? selectedPointer ?? targets[0].pointer

  return {
    label: 'nudge layout',
    mergeKey: `layout:nudge:${targetPointers.join('|')}`,
    operations: targets.flatMap((target) =>
      setArrangePatch(target.pointer, target.rect),
    ),
    pointer: primaryPointer,
    selection: selectionSnapForPointers(targetPointers, focusPointer),
  }
}
