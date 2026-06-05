import type { RefObject } from 'react'
import {
  applyPatch,
  type JSONPatchOperation,
  type Pointer,
  type SelectionSnap,
} from 'zod-crud'
import {
  SAMPLE_DECK,
  RetouchDeckSchema,
  blockLocationFromPointer,
  blockTextPointer,
  findBlockLocation,
  getRect,
  setLayoutPatch,
  setTextPatch,
  type Rect,
  type RetouchDeck,
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

export function useRetouchTextEditing({
  clearLayoutInteraction,
  commitPatch,
  deckValue,
  editing,
  mode,
  selectBlock,
  setEditing,
  slideRef,
}: {
  clearLayoutInteraction: () => void
  commitPatch: CommitPatch
  deckValue: RetouchDeck
  editing: EditingState | null
  mode: 'text' | 'layout'
  selectBlock: (pointer: Pointer) => void
  setEditing: (editing: EditingState | null) => void
  slideRef: RefObject<HTMLDivElement | null>
}) {
  function startTextEdit(pointer: Pointer, clientPoint?: Point) {
    if (mode !== 'text') {
      return
    }

    selectBlock(pointer)
    setEditing({ clientPoint, pointer })
  }

  function cancelTextEdit() {
    setEditing(null)
    clearLayoutInteraction()
  }

  function commitTextEdit(pointer: Pointer, text: string, rect: Rect) {
    setEditing(null)
    clearLayoutInteraction()
    commitTextPatch(pointer, text, rect)
  }

  function commitActiveTextEdit() {
    if (!editing) {
      return deckValue
    }

    const activeEditing = editing
    const location = blockLocationFromPointer(deckValue, activeEditing.pointer)
    const element = slideRef.current?.querySelector<HTMLElement>(
      PLAIN_TEXT_BLOCK_EDITOR_SELECTOR,
    )
    const blockElement = element?.closest<HTMLElement>('[data-block]')

    setEditing(null)
    clearLayoutInteraction()

    if (!location || !element || !blockElement) {
      return deckValue
    }

    const text = normalizePlainTextBlockEditorText(element.textContent ?? '')
    const baseBlock = findBlockLocation(
      SAMPLE_DECK,
      location.slide.id,
      location.block.id,
    )?.block
    const minimumHeight = minimumHeightForBlock(
      location.block,
      baseBlock,
      getRect(location.block),
    )
    const rect = autoHeightRect(blockElement, getRect(location.block), minimumHeight)

    return commitTextPatch(activeEditing.pointer, text, rect)
  }

  function commitTextPatch(pointer: Pointer, text: string, rect: Rect) {
    const location = blockLocationFromPointer(deckValue, pointer)

    if (!location) {
      return deckValue
    }

    const currentRect = getRect(location.block)
    const textChanged = text !== location.block.text
    const layoutChanged =
      rect.x !== currentRect.x ||
      rect.y !== currentRect.y ||
      rect.width !== currentRect.width ||
      (textChanged && rect.height !== currentRect.height)
    const patch = [
      ...(textChanged ? setTextPatch(pointer, text) : []),
      ...(layoutChanged ? setLayoutPatch(pointer, rect) : []),
    ]

    if (patch.length === 0) {
      return deckValue
    }

    commitPatch(
      patch,
      pointer,
      textChanged ? 'edit text' : 'resize text box',
      textChanged ? `text:${blockTextPointer(pointer)}` : undefined,
    )

    const applied = applyPatch(RetouchDeckSchema, deckValue, patch)

    return applied.result.ok ? applied.state : deckValue
  }

  return {
    cancelTextEdit,
    commitActiveTextEdit,
    commitTextEdit,
    commitTextPatch,
    startTextEdit,
  }
}
