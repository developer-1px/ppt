import { useMemo } from 'react'
import type { Pointer } from 'zod-crud'
import {
  SAMPLE_DECK,
  blockLocationFromPointer,
  findBlockLocation,
  type RetouchDeck,
  type RetouchSlide,
} from './retouchModel'
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
    focusPointer && selectedPointerSet.has(focusPointer)
      ? focusPointer
      : (selectedPointers.at(-1) ?? null)
  const selectedLocation = useMemo(
    () => blockLocationFromPointer(deckValue, primaryPointer),
    [deckValue, primaryPointer],
  )
  const selectedPointer =
    selectedLocation?.slide.id === activeSlide.id ? selectedLocation.pointer : null
  const selectedBlock =
    selectedLocation?.slide.id === activeSlide.id ? selectedLocation.block : null
  const baseSelectedLocation =
    selectedBlock === null
      ? null
      : findBlockLocation(SAMPLE_DECK, activeSlide.id, selectedBlock.id)
  const canResetSelectedLayout = Boolean(
    selectedPointer &&
      selectedBlock &&
      baseSelectedLocation &&
      !arrangeResetEquals(selectedBlock, baseSelectedLocation.block),
  )
  const canResetSelectedText = Boolean(
    selectedPointer &&
      selectedBlock &&
      baseSelectedLocation &&
      (!textResetEquals(selectedBlock, baseSelectedLocation.block) ||
        editing?.pointer === selectedPointer),
  )
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
