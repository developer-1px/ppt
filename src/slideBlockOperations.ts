import type { JSONPatchOperation, Pointer } from 'zod-crud'
import {
  SLIDE_HEIGHT,
  SLIDE_WIDTH,
  blockPointer,
  clamp,
  type RetouchSlide,
  type SlideBlock,
} from './retouchModel'
import {
  createRetouchBlockCopyId,
  createRetouchTextBlockId,
  nextRetouchTextBlockOrdinal,
} from './retouchIdResolver'

const DUPLICATE_OFFSET = 32

type BlockInsertInput = {
  blocks: readonly SlideBlock[]
  insertIndex: number
  slideIndex: number
}

type BlockInsertPatch = {
  insertedPointers: Pointer[]
  operations: JSONPatchOperation[]
}

export function createTextBlock(slide: RetouchSlide): SlideBlock {
  const ordinal = nextRetouchTextBlockOrdinal(slide)
  const id = createRetouchTextBlockId(slide)
  const offset = ((ordinal - 1) % 5) * 24

  return {
    id,
    role: 'body',
    tag: 'p',
    text: '',
    className: 'block-subtitle',
    x: 360 + offset,
    y: 312 + offset,
    width: 560,
    height: 56,
  }
}

function duplicateBlock(
  block: SlideBlock,
  slide: RetouchSlide,
): SlideBlock {
  return {
    ...block,
    id: createRetouchBlockCopyId(slide, block),
    x: clamp(block.x + DUPLICATE_OFFSET, 0, SLIDE_WIDTH - block.width),
    y: clamp(block.y + DUPLICATE_OFFSET, 0, SLIDE_HEIGHT - block.height),
  }
}

export function duplicateBlocks(
  blocks: readonly SlideBlock[],
  slide: RetouchSlide,
): SlideBlock[] {
  const duplicatedBlocks: SlideBlock[] = []

  for (const block of blocks) {
    duplicatedBlocks.push(
      duplicateBlock(block, {
        ...slide,
        blocks: [...slide.blocks, ...duplicatedBlocks],
      }),
    )
  }

  return duplicatedBlocks
}

export function createBlockInsertPatch({
  blocks,
  insertIndex,
  slideIndex,
}: BlockInsertInput): BlockInsertPatch {
  const inserts = blocks.map((block, offset) => ({
    block,
    pointer: blockPointer(slideIndex, insertIndex + offset),
  }))

  return {
    insertedPointers: inserts.map((insert) => insert.pointer),
    operations: inserts.map((insert) => ({
      op: 'add',
      path: insert.pointer,
      value: insert.block,
    })),
  }
}
