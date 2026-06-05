import {
  canReorderLayers,
  type LayerOrderAction as SharedLayerOrderAction,
} from '@zod-crud/layer-order'
import {
  appendSegment,
  type JSONDocument,
  type JSONPatchOperation,
  type Pointer,
} from 'zod-crud'
import { slideBlocksPointer, type RetouchDeck, type SlideBlock } from './retouchModel'

export type LayerOrderAction = 'front' | 'forward' | 'backward' | 'back'

export type LayerOrderPatch = {
  operations: JSONPatchOperation[]
  nextSelectedPointers: Pointer[]
}

const LAYER_ORDER_ACTIONS: Record<LayerOrderAction, SharedLayerOrderAction> = {
  front: 'bringToFront',
  forward: 'bringForward',
  backward: 'sendBackward',
  back: 'sendToBack',
}

export function createLayerOrderPatch({
  action,
  activeSlideIndex,
  doc,
  selectedIds,
  selectedPointers,
}: {
  action: LayerOrderAction
  activeSlideIndex: number
  doc: JSONDocument<RetouchDeck>
  selectedIds: readonly string[]
  selectedPointers: readonly Pointer[]
}): LayerOrderPatch | null {
  const blocksPath = slideBlocksPointer(activeSlideIndex)
  const change = canReorderLayers(doc, selectedPointers, LAYER_ORDER_ACTIONS[action])

  if (!change.ok || change.parent !== blocksPath) {
    return null
  }

  const operation = change.operations[0]

  if (
    change.operations.length !== 1 ||
    !operation ||
    operation.op !== 'replace' ||
    operation.path !== blocksPath ||
    !Array.isArray(operation.value)
  ) {
    return null
  }

  const nextBlocks = operation.value as SlideBlock[]
  const selectedIdSet = new Set(selectedIds)

  const nextSelectedPointers = nextBlocks
    .map((block, blockIndex) =>
      selectedIdSet.has(block.id) ? appendSegment(blocksPath, blockIndex) : null,
    )
    .filter((pointer): pointer is Pointer => pointer !== null)

  return {
    operations: [...change.operations],
    nextSelectedPointers,
  }
}
