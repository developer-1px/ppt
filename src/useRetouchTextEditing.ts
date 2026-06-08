import type { RefObject } from 'react'
import type { Point } from 'canvas/core'
import { applyPatch, type Pointer } from 'zod-crud'
import {
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
import { SAMPLE_DECK } from './sampleDeck'
import { autoHeightRect } from './editableTextDom'
import { minimumHeightForBlock } from './layoutInteraction'
import { normalizePlainTextBlockEditorText } from './plainTextBlockEditor'
import { readPlainTextBlockEditorElements } from './retouchSlideDom'
import type { RetouchSurfaceCommitPatch } from './retouchSurfaceContract'
import type { EditingState, RetouchMode } from './retouchViewState'

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
  commitPatch: RetouchSurfaceCommitPatch
  deckValue: RetouchDeck
  editing: EditingState | null
  mode: RetouchMode
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
    const editorElements = readPlainTextBlockEditorElements(slideRef.current)

    setEditing(null)
    clearLayoutInteraction()

    if (!location || !editorElements) {
      return deckValue
    }

    const text = normalizePlainTextBlockEditorText(
      editorElements.editorElement.textContent ?? '',
    )
    const block = location.block
    const blockRect = getRect(block)
    const baseBlock = findBlockLocation(
      SAMPLE_DECK,
      location.slide.id,
      block.id,
    )?.block
    const minimumHeight = minimumHeightForBlock(
      block,
      baseBlock,
      blockRect,
    )
    const rect = autoHeightRect(
      editorElements.blockElement,
      blockRect,
      minimumHeight,
    )

    return commitTextPatch(activeEditing.pointer, text, rect)
  }

  function commitTextPatch(pointer: Pointer, text: string, rect: Rect) {
    const location = blockLocationFromPointer(deckValue, pointer)

    if (!location) {
      return deckValue
    }

    const block = location.block
    const currentRect = getRect(block)
    const textChanged = text !== block.text
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
