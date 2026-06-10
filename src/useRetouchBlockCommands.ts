import { useState, type RefObject } from 'react'
import type { JSONDocument, Pointer } from 'zod-crud'
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
} from './retouchSelectionSnap'
import {
  normalizeInspectorRect,
  type RectField,
} from './inspectorGeometry'
import type {
  RetouchPatchCommit,
  RetouchSurfaceCommitPatch,
} from './retouchSurfaceContract'
import type { CanvasView, EditingState, RetouchMode } from './retouchViewState'

type UseRetouchBlockCommandsParams = {
  activeSlide: RetouchSlide
  activeSlideIndex: number
  clearLayoutInteraction: () => void
  commitActiveTextEdit: () => RetouchDeck
  commitPatch: RetouchSurfaceCommitPatch
  commitRetouchPatch: RetouchPatchCommit
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
  const activeBlocks = activeSlide.blocks
  const activeBlockCount = activeBlocks.length

  function commitInsertedBlocks(
    blocks: readonly SlideBlock[],
    insertIndex: number,
    label: string,
  ) {
    const insertPatch = createBlockInsertPatch({
      blocks,
      insertIndex,
      slideIndex: activeSlideIndex,
    })
    const primaryPointer = insertPatch.insertedPointers[0]

    if (primaryPointer === undefined) {
      return null
    }

    commitRetouchPatch(insertPatch.operations, {
      label,
      selection: selectionSnapForPointers(insertPatch.insertedPointers),
    })

    return primaryPointer
  }

  function insertTextBlock() {
    commitActiveTextEdit()
    const nextBlock = createTextBlock(activeSlide)
    const pointer = commitInsertedBlocks(
      [nextBlock],
      activeBlockCount,
      'add text block',
    )

    if (pointer === null) {
      return
    }
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

  function commitDuplicatedBlocks(
    blocks: readonly SlideBlock[],
    insertIndex: number,
    label: string,
  ) {
    commitActiveTextEdit()
    const duplicatedBlocks = duplicateBlocks(blocks, activeSlide)
    const primaryPointer = commitInsertedBlocks(duplicatedBlocks, insertIndex, label)

    return primaryPointer !== null ? duplicatedBlocks : null
  }

  function pasteCopiedBlocks() {
    if (blockClipboard.length === 0) {
      return
    }

    const pastedBlocks = commitDuplicatedBlocks(
      blockClipboard,
      activeBlockCount,
      'paste blocks',
    )

    if (pastedBlocks === null) {
      return
    }

    setBlockClipboard(pastedBlocks)
    enterLayoutMode()
  }

  function duplicateSelectedBlock() {
    const locations = selectedActiveBlockLocations()

    if (locations.length === 0) {
      return
    }

    const lastLocation = locations.at(-1)

    if (lastLocation === undefined) {
      return
    }

    const duplicatedBlocks = commitDuplicatedBlocks(
      locations.map((location) => location.block),
      lastLocation.blockIndex + 1,
      locations.length > 1 ? 'duplicate blocks' : 'duplicate block',
    )

    if (duplicatedBlocks === null) {
      return
    }

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
      activeBlockCount - locations.length - 1,
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
    const targetPointers = targets.map((target) => target.pointer)
    const lastTargetPointer = targetPointers.at(-1)

    commitActiveTextEdit()
    enterLayoutMode()
    commitPatch(
      targets.flatMap((target) => setLayoutPatch(target.pointer, target.rect)),
      lastTargetPointer ?? selectedPointer ?? targets[0].pointer,
      label,
      undefined,
      selectionSnapForPointers(
        targetPointers,
        selectedPointer ?? lastTargetPointer,
      ),
    )
  }

  function alignSelectedBlocks(action: AlignSelectionAction) {
    const targets = alignBlockLocations(selectedActiveBlockLocations(), action)

    if (targets === null) {
      return
    }

    commitSelectedLayoutTargets(targets, 'align selection')
  }

  function distributeSelectedBlocks(action: DistributeSelectionAction) {
    const targets = distributeBlockLocations(selectedActiveBlockLocations(), action)

    if (targets === null) {
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
    const selectedLocationPointers = locations.map((location) => location.pointer)
    const layerOrderPatch = createLayerOrderPatch({
      action,
      activeSlideIndex,
      doc,
      selectedIds,
      selectedPointers: selectedLocationPointers,
    })

    if (layerOrderPatch === null) {
      return
    }

    const nextPrimaryPointer = layerOrderPatch.nextSelectedPointers.at(-1)

    commitActiveTextEdit()
    commitRetouchPatch(
      layerOrderPatch.operations,
      {
        label: 'reorder layers',
        selection: selectionSnapForPointers(
          layerOrderPatch.nextSelectedPointers,
          nextPrimaryPointer,
        ),
      },
    )
    enterLayoutMode()
  }

  function changeSelectedBlockRect(rect: Rect, changedField?: RectField) {
    if (
      selectedPointer === null ||
      selectedBlock === null ||
      selectedRect === null
    ) {
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
    if (activeBlockCount === 0) {
      doc.selection?.empty()
      return
    }

    enterLayoutMode()
    doc.selection?.selectRanges(
      activeBlocks.map((_, blockIndex) =>
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
