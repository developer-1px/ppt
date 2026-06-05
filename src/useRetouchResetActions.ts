import type { RefObject } from 'react'
import type { JSONPatchOperation, Pointer, SelectionSnap } from 'zod-crud'
import {
  EMPTY_TEXT_BOX_HEIGHT,
  SAMPLE_DECK,
  SAMPLE_SLIDES,
  blockLocationFromPointer,
  findBlockLocation,
  getRect,
  setArrangePatch,
  setLayoutPatch,
  setTextPatch,
  type BlockLocation,
  type Rect,
  type RetouchDeck,
  type SlideBlock,
} from './retouchModel'
import { autoHeightRect } from './editableTextDom'
import { minimumHeightForBlock, type Point } from './layoutInteraction'
import {
  PLAIN_TEXT_BLOCK_EDITOR_SELECTOR,
  normalizePlainTextBlockEditorText,
} from './plainTextBlockEditor'

type EditingState = {
  clientPoint?: Point
  pointer: Pointer
}

type CommitPatch = (
  patch: JSONPatchOperation[],
  pointer: Pointer,
  label: string,
  mergeKey?: string,
  selection?: SelectionSnap,
) => void

type HistoryApi = {
  redo: () => void
  undo: () => void
}

type SelectionApi = {
  empty?: () => void
}

export function useRetouchResetActions({
  baseSelectedLocation,
  canResetSelectedLayout,
  commitDeckReset,
  clearTransientState,
  commitActiveTextEdit,
  commitPatch,
  commitTextPatch,
  deckValue,
  editing,
  hasDeckChanges,
  history,
  mode,
  resetExportFeedback,
  selectedBlock,
  selectedPointer,
  selection,
  setActiveSlideId,
  setEditing,
  slideRef,
  stageRef,
}: {
  baseSelectedLocation: BlockLocation | null
  canResetSelectedLayout: boolean
  commitDeckReset: () => void
  clearTransientState: () => void
  commitActiveTextEdit: () => RetouchDeck
  commitPatch: CommitPatch
  commitTextPatch: (pointer: Pointer, text: string, rect: Rect) => RetouchDeck
  deckValue: RetouchDeck
  editing: EditingState | null
  hasDeckChanges: boolean
  history: HistoryApi
  mode: 'text' | 'layout'
  resetExportFeedback: () => void
  selectedBlock: SlideBlock | null
  selectedPointer: Pointer | null
  selection: SelectionApi | undefined
  setActiveSlideId: (slideId: string) => void
  setEditing: (editing: EditingState | null) => void
  slideRef: RefObject<HTMLDivElement | null>
  stageRef: RefObject<HTMLDivElement | null>
}) {
  function resetSelectedLayout() {
    if (
      mode !== 'layout' ||
      !selectedPointer ||
      !selectedBlock ||
      !baseSelectedLocation ||
      !canResetSelectedLayout
    ) {
      return
    }

    commitPatch(
      setArrangePatch(selectedPointer, getRect(baseSelectedLocation.block), {
        includeHeight:
          selectedBlock.text === baseSelectedLocation.block.text &&
          selectedBlock.height !== baseSelectedLocation.block.height,
      }),
      selectedPointer,
      'reset layout',
    )
  }

  function resetSelectedText() {
    if (mode !== 'text') {
      return
    }

    const pointer = editing?.pointer ?? selectedPointer
    const location = pointer ? blockLocationFromPointer(deckValue, pointer) : null
    const baseLocation = location
      ? findBlockLocation(SAMPLE_DECK, location.slide.id, location.block.id)
      : null

    if (!pointer || !location || !baseLocation) {
      return
    }

    const element =
      editing?.pointer === pointer
        ? slideRef.current?.querySelector<HTMLElement>(
            PLAIN_TEXT_BLOCK_EDITOR_SELECTOR,
          )
        : null
    const blockElement = element?.closest<HTMLElement>('[data-block]')
    const liveText = normalizePlainTextBlockEditorText(
      element?.textContent ?? location.block.text,
    )
    const liveMinimumHeight =
      liveText.length === 0
        ? EMPTY_TEXT_BOX_HEIGHT
        : minimumHeightForBlock(
            location.block,
            baseLocation.block,
            getRect(location.block),
          )
    const liveRect = element && blockElement
      ? autoHeightRect(blockElement, getRect(location.block), liveMinimumHeight)
      : getRect(location.block)
    const resetRect = {
      ...liveRect,
      height: baseLocation.block.height,
    }

    if (element && liveText !== location.block.text) {
      commitTextPatch(pointer, liveText, liveRect)
    }

    setEditing(null)
    clearTransientState()

    const patch = [
      ...(liveText !== baseLocation.block.text ||
      location.block.text !== baseLocation.block.text
        ? setTextPatch(pointer, baseLocation.block.text)
        : []),
      ...(liveRect.height !== resetRect.height
        ? setLayoutPatch(pointer, resetRect)
        : []),
    ]

    commitPatch(patch, pointer, 'reset text')
  }

  function resetDeck() {
    if (mode === 'layout' || !hasDeckChanges) {
      return
    }

    commitActiveTextEdit()
    clearTransientState()
    selection?.empty?.()
    setActiveSlideId(SAMPLE_SLIDES[0].id)
    resetExportFeedback()
    stageRef.current?.scrollTo({ left: 0, top: 0 })
    commitDeckReset()
  }

  function resetCurrentTarget() {
    if (mode === 'layout') {
      resetSelectedLayout()
      return
    }

    if (selectedPointer) {
      resetSelectedText()
      return
    }

    resetDeck()
  }

  function undoDocumentChange() {
    commitActiveTextEdit()
    history.undo()
  }

  function redoDocumentChange() {
    commitActiveTextEdit()
    history.redo()
  }

  return {
    redoDocumentChange,
    resetCurrentTarget,
    undoDocumentChange,
  }
}
