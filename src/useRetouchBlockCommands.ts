import { useState, type RefObject } from 'react'
import type { JSONDocument, JSONPatchOperation, Pointer, SelectionSnap } from 'zod-crud'
import type { RetouchCollection } from './retouchCollection'
import {
  blockLocationsFromPointers,
  blockPointer,
  rectEquals,
  setLayoutPatch,
  type Rect,
  type RetouchDeck,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'
import {
  createBlockInsertPatch,
  createTextBlock,
  duplicateBlocks,
} from './slideBlockOperations'
import {
  alignBlockLocations,
  distributeBlockLocations,
  type AlignSelectionAction,
  type DistributeSelectionAction,
  type SelectionLayoutTarget,
} from './selectionAlignment'
import {
  createLayerOrderPatch,
  type LayerOrderAction,
} from './selectionLayerOrder'
import {
  selectionSnapForPointers,
} from './layoutInteraction'
import {
  normalizeInspectorRect,
  type RectField,
} from './inspectorGeometry'
import type { RetouchSurfaceCommitPatch } from './retouchSurfaceContract'
import type { CanvasView, EditingState, RetouchMode } from './retouchViewState'

type CommitRetouchPatch = (
  patch: readonly JSONPatchOperation[],
  options:
    | string
    | { label: string; mergeKey?: string; selection?: SelectionSnap },
) => void

type UseRetouchBlockCommandsParams = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  clearLayoutInteraction: () => void
  commitActiveTextEdit: () => RetouchDeck
  commitPatch: RetouchSurfaceCommitPatch
  commitRetouchPatch: CommitRetouchPatch
  doc: JSONDocument<RetouchDeck>
  enterLayoutMode: () => void
  retouchCollection: RetouchCollection
  selectedBlock: SlideBlock | null
  selectedPointer: Pointer | null
  selectedPointers: Pointer[]
  selectedRect: Rect | null
  setCanvasView: (canvasView: CanvasView) => void
  setEditing: (editing: EditingState | null) => void
  setMode: (mode: RetouchMode) => void
  stageRef: RefObject<HTMLDivElement | null>
}

export function useRetouchBlockCommands({
  activeSlide,
  activeSlideIndex,
  clearLayoutInteraction,
  commitActiveTextEdit,
  commitPatch,
  commitRetouchPatch,
  doc,
  enterLayoutMode,
  retouchCollection,
  selectedBlock,
  selectedPointer,
  selectedPointers,
  selectedRect,
  setCanvasView,
  setEditing,
  setMode,
  stageRef,
}: UseRetouchBlockCommandsParams) {
  const [blockClipboard, setBlockClipboard] = useState<SlideBlock[]>([])

  function insertTextBlock() {
    commitActiveTextEdit()
    const nextBlock = createTextBlock(activeSlide)
    const blockIndex = activeSlide.blocks.length
    const insertPatch = createBlockInsertPatch({
      blocks: [nextBlock],
      insertIndex: blockIndex,
      slideIndex: activeSlideIndex,
    })
    const pointer = insertPatch.insertedPointers[0]

    if (!pointer) {
      return
    }

    commitRetouchPatch(insertPatch.operations, {
      label: 'add text block',
      selection: selectionSnapForPointers([pointer]),
    })
    setCanvasView('slide')
    setMode('text')
    clearLayoutInteraction()
    setEditing({ pointer })
    stageRef.current?.scrollTo({ left: 0, top: 0 })
  }

  function selectedActiveBlockLocations() {
    return blockLocationsFromPointers(doc.value, selectedPointers).filter(
      (location) => location.slide.id === activeSlide.id,
    )
  }

  function copySelectedBlocks() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    setBlockClipboard(locations.map((location) => ({ ...location.block })))
  }

  function pasteCopiedBlocks() {
    if (blockClipboard.length === 0) {
      return
    }

    commitActiveTextEdit()
    const pastedBlocks = duplicateBlocks(blockClipboard, activeSlide)
    const insertIndex = activeSlide.blocks.length
    const insertPatch = createBlockInsertPatch({
      blocks: pastedBlocks,
      insertIndex,
      slideIndex: activeSlideIndex,
    })

    commitRetouchPatch(insertPatch.operations, {
      label: 'paste blocks',
      selection: selectionSnapForPointers(insertPatch.insertedPointers),
    })
    setBlockClipboard(pastedBlocks)
    enterLayoutMode()
  }

  function duplicateSelectedBlock() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    commitActiveTextEdit()
    const duplicatedBlocks = duplicateBlocks(
      locations.map((location) => location.block),
      activeSlide,
    )
    const insertIndex = locations.at(-1)!.blockIndex + 1
    const insertPatch = createBlockInsertPatch({
      blocks: duplicatedBlocks,
      insertIndex,
      slideIndex: activeSlideIndex,
    })

    commitRetouchPatch(insertPatch.operations, {
      label: locations.length > 1 ? 'duplicate blocks' : 'duplicate block',
      selection: selectionSnapForPointers(insertPatch.insertedPointers),
    })
    enterLayoutMode()
  }

  function deleteSelectedBlock() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    commitActiveTextEdit()
    const nextSelectionIndex = Math.min(
      locations[0].blockIndex,
      activeSlide.blocks.length - locations.length - 1,
    )

    const deleted = retouchCollection.deleteBlocks(
      locations.map((location) => location.pointer),
    )
    if (!deleted.ok) {
      return
    }

    enterLayoutMode()

    if (nextSelectionIndex >= 0) {
      doc.selection?.selectRanges?.([blockPointer(activeSlideIndex, nextSelectionIndex)])
    } else {
      doc.selection?.empty()
    }
  }

  function commitSelectedLayoutTargets(
    targets: SelectionLayoutTarget[],
    label: string,
  ) {
    commitActiveTextEdit()
    enterLayoutMode()
    commitPatch(
      targets.flatMap((target) => setLayoutPatch(target.pointer, target.rect)),
      targets.at(-1)?.pointer ?? selectedPointer ?? targets[0].pointer,
      label,
      undefined,
      selectionSnapForPointers(
        targets.map((target) => target.pointer),
        selectedPointer ?? targets.at(-1)?.pointer,
      ),
    )
  }

  function alignSelectedBlocks(action: AlignSelectionAction) {
    const targets = alignBlockLocations(selectedActiveBlockLocations(), action)

    if (!targets) {
      return
    }

    commitSelectedLayoutTargets(targets, 'align selection')
  }

  function distributeSelectedBlocks(action: DistributeSelectionAction) {
    const targets = distributeBlockLocations(selectedActiveBlockLocations(), action)

    if (!targets) {
      return
    }

    commitSelectedLayoutTargets(targets, 'distribute selection')
  }

  function changeSelectedLayerOrder(action: LayerOrderAction) {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    const selectedIds = locations.map((location) => location.block.id)
    const layerOrderPatch = createLayerOrderPatch({
      action,
      activeSlideIndex,
      doc,
      selectedIds,
      selectedPointers: locations.map((location) => location.pointer),
    })

    if (!layerOrderPatch) {
      return
    }

    commitActiveTextEdit()
    commitRetouchPatch(
      layerOrderPatch.operations,
      {
        label: 'reorder layers',
        selection: selectionSnapForPointers(
          layerOrderPatch.nextSelectedPointers,
          layerOrderPatch.nextSelectedPointers.at(-1),
        ),
      },
    )
    enterLayoutMode()
  }

  function changeSelectedBlockRect(rect: Rect, changedField?: RectField) {
    if (!selectedPointer || !selectedBlock || !selectedRect) {
      return
    }

    const nextRect = normalizeInspectorRect(rect, selectedRect, changedField)

    if (rectEquals(nextRect, selectedRect)) {
      return
    }

    commitActiveTextEdit()
    enterLayoutMode()
    commitPatch(
      setLayoutPatch(selectedPointer, nextRect),
      selectedPointer,
      'edit block geometry',
      `layout:geometry:${selectedPointer}`,
    )
  }

  function selectAllBlocks() {
    if (activeSlide.blocks.length === 0) {
      doc.selection?.empty()
      return
    }

    enterLayoutMode()
    doc.selection?.selectRanges(
      activeSlide.blocks.map((_, blockIndex) =>
        blockPointer(activeSlideIndex, blockIndex),
      ),
    )
  }

  return {
    alignSelectedBlocks,
    canPasteSelection: blockClipboard.length > 0,
    changeSelectedBlockRect,
    changeSelectedLayerOrder,
    copySelectedBlocks,
    deleteSelectedBlock,
    distributeSelectedBlocks,
    duplicateSelectedBlock,
    insertTextBlock,
    pasteCopiedBlocks,
    selectAllBlocks,
  }
}
