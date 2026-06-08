import type { Pointer, SelectionSnap } from 'zod-crud'

export function selectionSnapForPointers(
  pointers: readonly Pointer[],
  primaryPointer = pointers.at(-1),
): SelectionSnap {
  const selectedPointers = [...new Set(pointers)]
  const fallbackPrimary = selectedPointers.at(-1) ?? null
  const primary =
    primaryPointer !== undefined && selectedPointers.includes(primaryPointer)
      ? primaryPointer
      : fallbackPrimary
  const primaryIndex = primary !== null ? selectedPointers.indexOf(primary) : -1

  return {
    anchor: primary,
    focus: primary,
    primaryIndex,
    selectedPointers,
    selectionRanges: selectedPointers.map((pointer) => ({
      anchor: pointer,
      focus: pointer,
    })),
  }
}
