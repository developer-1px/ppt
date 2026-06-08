import { useMemo } from 'react'
import type { Pointer } from 'zod-crud'
import {
  blockLocationFromPointer,
  findBlockLocation,
  type RetouchDeck,
  type RetouchSlide,
} from './retouchModel'
import { SAMPLE_DECK } from './sampleDeck'
import {
  arrangeResetEquals,
  textResetEquals,
} from './layoutInteraction'
import type {
  EditingState,
  ResetScope,
  RetouchMode,
} from './retouchViewState'

export function useRetouchSelectionState({
  activeSlide,
  deckValue,
  editing,
  focusPointer,
  hasDeckChanges,
  mode,
  selectedPointersFromDocument,
}: {
  activeSlide: RetouchSlide
  deckValue: RetouchDeck
  editing: EditingState | null
  focusPointer: Pointer | null
  hasDeckChanges: boolean
  mode: RetouchMode
  selectedPointersFromDocument: readonly Pointer[]
}) {
  const selectedPointers = selectedPointersFromDocument.filter((pointer) => {
    const location = blockLocationFromPointer(deckValue, pointer)

    return location?.slide.id === activeSlide.id
  })
  const selectedPointerSet = new Set(selectedPointers)
  const primaryPointer =
    focusPointer !== null && selectedPointerSet.has(focusPointer)
      ? focusPointer
      : (selectedPointers.at(-1) ?? null)
  const selectedLocation = useMemo(
    () => blockLocationFromPointer(deckValue, primaryPointer),
    [deckValue, primaryPointer],
  )
  const activeSelectedLocation =
    selectedLocation?.slide.id === activeSlide.id ? selectedLocation : null
  const selectedPointer = activeSelectedLocation?.pointer ?? null
  const selectedBlock = activeSelectedLocation?.block ?? null
  const baseSelectedLocation =
    selectedBlock
      ? findBlockLocation(SAMPLE_DECK, activeSlide.id, selectedBlock.id)
      : null
  const baseSelectedBlock = baseSelectedLocation?.block ?? null
  const canResetSelectedLayout =
    selectedPointer !== null &&
    selectedBlock !== null &&
    baseSelectedBlock !== null &&
    !arrangeResetEquals(selectedBlock, baseSelectedBlock)
  const canResetSelectedText =
    selectedPointer !== null &&
    selectedBlock !== null &&
    baseSelectedBlock !== null &&
      (!textResetEquals(selectedBlock, baseSelectedBlock) ||
        editing?.pointer === selectedPointer)
  const canResetDeck = !selectedPointer && hasDeckChanges
  const canReset =
    mode === 'layout'
      ? canResetSelectedLayout
      : canResetSelectedText || canResetDeck
  const resetScope: ResetScope =
    mode === 'layout' ? 'layout' : selectedPointer ? 'text' : 'deck'
  const resetTitle =
    resetScope === 'layout'
      ? 'Reset layout'
      : resetScope === 'text'
        ? 'Reset text'
        : 'Reset deck'

  return {
    baseSelectedLocation,
    canReset,
    canResetSelectedLayout,
    resetScope,
    resetTitle,
    selectedBlock,
    selectedPointer,
    selectedPointerSet,
    selectedPointers,
  }
}
