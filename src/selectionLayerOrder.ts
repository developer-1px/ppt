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
import {
  SlideBlockSchema,
  slideBlocksPointer,
  type RetouchDeck,
  type SlideBlock,
} from './retouchModel'

type LayerOrderPatch = {
  operations: JSONPatchOperation[]
  nextSelectedPointers: Pointer[]
}

const LAYER_ORDER_ACTIONS = {
  front: 'bringToFront',
  forward: 'bringForward',
  backward: 'sendBackward',
  back: 'sendToBack',
} satisfies Record<string, SharedLayerOrderAction>

export type LayerOrderAction = keyof typeof LAYER_ORDER_ACTIONS

const LayerOrderBlocksSchema = SlideBlockSchema.array()

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
    operation === undefined ||
    operation.op !== 'replace' ||
    operation.path !== blocksPath
  ) {
    return null
  }

  const nextBlocks = readLayerOrderBlocks(operation.value)

  if (nextBlocks === null) {
    return null
  }

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

function readLayerOrderBlocks(value: unknown): SlideBlock[] | null {
  const result = LayerOrderBlocksSchema.safeParse(value)

  return result.success ? result.data : null
}
