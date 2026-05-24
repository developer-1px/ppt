import type { SlideBlock } from './retouchModel'

export type LayerOrderAction = 'front' | 'forward' | 'backward' | 'back'

export function reorderBlocksByLayer(
  blocks: SlideBlock[],
  selectedIds: string[],
  action: LayerOrderAction,
) {
  const selectedIdSet = new Set(selectedIds)
  const selectedBlocks = blocks.filter((block) => selectedIdSet.has(block.id))

  if (selectedBlocks.length === 0) {
    return blocks
  }

  if (action === 'front') {
    return [
      ...blocks.filter((block) => !selectedIdSet.has(block.id)),
      ...selectedBlocks,
    ]
  }

  if (action === 'back') {
    return [
      ...selectedBlocks,
      ...blocks.filter((block) => !selectedIdSet.has(block.id)),
    ]
  }

  const nextBlocks = [...blocks]

  if (action === 'forward') {
    for (let index = nextBlocks.length - 2; index >= 0; index -= 1) {
      if (
        selectedIdSet.has(nextBlocks[index]!.id) &&
        !selectedIdSet.has(nextBlocks[index + 1]!.id)
      ) {
        const block = nextBlocks[index]!
        nextBlocks[index] = nextBlocks[index + 1]!
        nextBlocks[index + 1] = block
      }
    }
  }

  if (action === 'backward') {
    for (let index = 1; index < nextBlocks.length; index += 1) {
      if (
        selectedIdSet.has(nextBlocks[index]!.id) &&
        !selectedIdSet.has(nextBlocks[index - 1]!.id)
      ) {
        const block = nextBlocks[index]!
        nextBlocks[index] = nextBlocks[index - 1]!
        nextBlocks[index - 1] = block
      }
    }
  }

  return nextBlocks
}

export function blockOrderChanged(a: SlideBlock[], b: SlideBlock[]) {
  return a.length !== b.length || a.some((block, index) => block.id !== b[index]?.id)
}
