import { useEffect } from 'react'
import type { JSONDocumentHistory, Pointer, SelectionState } from 'zod-crud'
import {
  findBlockLocation,
  type RetouchDeck,
} from './retouchModel'
import {
  arrowKeyDelta,
  historyShortcutAction,
  isControlTarget,
  isEditableTarget,
} from './editorKeyboard'
import type { Interaction } from './layoutInteraction'
import { createLayoutKeyboardNudgePatch } from './layoutKeyboardNudge'
import { closestSlideBlockElement } from './retouchSlideDom'
import type { RetouchSurfaceCommitPatch } from './retouchSurfaceContract'
import type { EditingState, RetouchMode } from './retouchViewState'

export function useRetouchKeyboardShortcuts({
  activeSlideId,
  commitPatch,
  deckValue,
  editing,
  history,
  interaction,
  mode,
  canPasteSelection,
  onCopySelection,
  onDeleteSelection,
  onDuplicateSelection,
  onPasteSelection,
  onSelectAllBlocks,
  selectedPointer,
  selectedPointers,
  selection,
  setEditing,
}: {
  activeSlideId: string
  commitPatch: RetouchSurfaceCommitPatch
  deckValue: RetouchDeck
  editing: EditingState | null
  history: JSONDocumentHistory
  interaction: Interaction | null
  mode: RetouchMode
  canPasteSelection: boolean
  onCopySelection: () => void
  onDeleteSelection: () => void
  onDuplicateSelection: () => void
  onPasteSelection: () => void
  onSelectAllBlocks: () => void
  selectedPointer: Pointer | null
  selectedPointers: Pointer[]
  selection: SelectionState | undefined
  setEditing: (editing: EditingState) => void
}) {
  useEffect(() => {
    function handleKeyDown(event: KeyboardEvent) {
      const historyAction = historyShortcutAction(event)

      if (
        event.defaultPrevented ||
        !historyAction ||
        isEditableTarget(event)
      ) {
        return
      }

      if (historyAction === 'redo' && history.canRedo) {
        event.preventDefault()
        history.redo()
        return
      }

      if (historyAction === 'undo' && history.canUndo) {
        event.preventDefault()
        history.undo()
      }
    }

    window.addEventListener('keydown', handleKeyDown)

    return () => {
      window.removeEventListener('keydown', handleKeyDown)
    }
  }, [history])

  useEffect(() => {
    function handleFocusedBlockEditKey(event: KeyboardEvent) {
      if (
        mode !== 'text' ||
        event.defaultPrevented ||
        isEditableTarget(event) ||
        (event.key !== 'Enter' && event.key !== 'F2')
      ) {
        return
      }

      const activeElement = document.activeElement

      if (!(activeElement instanceof HTMLElement)) {
        return
      }

      const blockElement = closestSlideBlockElement(activeElement)
      const blockId = blockElement?.dataset.block

      if (!blockId) {
        return
      }

      const location = findBlockLocation(deckValue, activeSlideId, blockId)

      if (!location) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      selection?.selectRanges?.([location.pointer])
      setEditing({ pointer: location.pointer })
    }

    window.addEventListener('keydown', handleFocusedBlockEditKey)

    return () => {
      window.removeEventListener('keydown', handleFocusedBlockEditKey)
    }
  }, [activeSlideId, deckValue, mode, selection, setEditing])

  useEffect(() => {
    function handleTextSelectionKey(event: KeyboardEvent) {
      if (
        mode !== 'text' ||
        event.defaultPrevented ||
        editing ||
        !selectedPointer ||
        isEditableTarget(event) ||
        isControlTarget(event) ||
        event.key !== 'Escape'
      ) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      selection?.empty?.()
    }

    window.addEventListener('keydown', handleTextSelectionKey)

    return () => {
      window.removeEventListener('keydown', handleTextSelectionKey)
    }
  }, [editing, mode, selectedPointer, selection])

  useEffect(() => {
    function handleLayoutCommandKey(event: KeyboardEvent) {
      if (
        mode !== 'layout' ||
        event.defaultPrevented ||
        editing ||
        interaction ||
        isEditableTarget(event) ||
        isControlTarget(event)
      ) {
        return
      }

      const key = event.key.toLowerCase()
      const commandKey = event.metaKey || event.ctrlKey

      if (commandKey && key === 'a') {
        event.preventDefault()
        event.stopPropagation()
        onSelectAllBlocks()
        return
      }

      if (commandKey && key === 'd' && selectedPointers.length > 0) {
        event.preventDefault()
        event.stopPropagation()
        onDuplicateSelection()
        return
      }

      if (commandKey && key === 'c' && selectedPointers.length > 0) {
        event.preventDefault()
        event.stopPropagation()
        onCopySelection()
        return
      }

      if (commandKey && key === 'v' && canPasteSelection) {
        event.preventDefault()
        event.stopPropagation()
        onPasteSelection()
        return
      }

      if (
        !commandKey &&
        (event.key === 'Delete' || event.key === 'Backspace') &&
        selectedPointers.length > 0
      ) {
        event.preventDefault()
        event.stopPropagation()
        onDeleteSelection()
      }
    }

    window.addEventListener('keydown', handleLayoutCommandKey)

    return () => {
      window.removeEventListener('keydown', handleLayoutCommandKey)
    }
  }, [
    canPasteSelection,
    editing,
    interaction,
    mode,
    onCopySelection,
    onDeleteSelection,
    onDuplicateSelection,
    onPasteSelection,
    onSelectAllBlocks,
    selectedPointers.length,
  ])

  useEffect(() => {
    function handleLayoutKey(event: KeyboardEvent) {
      if (
        mode !== 'layout' ||
        event.defaultPrevented ||
        editing ||
        interaction ||
        selectedPointers.length === 0 ||
        isEditableTarget(event) ||
        isControlTarget(event)
      ) {
        return
      }

      if (event.key === 'Escape') {
        event.preventDefault()
        event.stopPropagation()
        selection?.empty?.()
        return
      }

      const delta = arrowKeyDelta(event.key, event.shiftKey)

      if (!delta) {
        return
      }

      const nudgePatch = createLayoutKeyboardNudgePatch({
        activeSlideId,
        deckValue,
        delta,
        selectedPointer,
        selectedPointers,
      })

      if (!nudgePatch) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      commitPatch(
        nudgePatch.operations,
        nudgePatch.pointer,
        nudgePatch.label,
        nudgePatch.mergeKey,
        nudgePatch.selection,
      )
    }

    window.addEventListener('keydown', handleLayoutKey)

    return () => {
      window.removeEventListener('keydown', handleLayoutKey)
    }
  }, [
    activeSlideId,
    commitPatch,
    deckValue,
    editing,
    interaction,
    mode,
    selectedPointer,
    selectedPointers,
    selection,
  ])
}
